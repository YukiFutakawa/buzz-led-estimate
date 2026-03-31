"""LED選定エンジン

既存器具から最適なLED代替品を選定する。
選定優先度: 1.器具跡が残らない → 2.安い → 3.デザイン性
"""

from __future__ import annotations

import logging
import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import yaml

from models import (
    ExistingFixture,
    LEDProduct,
    MatchResult,
    BALLAST_TYPES,
)
from lineup_loader import LineupIndex
from size_parser import (
    FixtureDimensions,
    parse_fixture_size,
    is_size_compatible,
)

logger = logging.getLogger(__name__)


@dataclass
class FixtureClassification:
    """現調器具の分類結果"""
    lineup_sheet: str                # 検索先ラインナップシート
    watt_form: str = ""             # 20形, 40形 等
    has_emergency: bool = False     # 非常灯
    is_recessed: bool = False       # 埋込式
    is_waterproof: bool = False     # 防水（器具名/場所いずれか由来）
    wp_hard_filter: bool = False    # 場所由来の防水（ハードフィルタ対象）
    color_pref: str = ""            # 昼白色, 電球色
    bulb_count: int = 1             # 灯数
    fallback: bool = False          # フォールバック分類
    type_diameter: float = 0.0      # 器具種別名から抽出した直径(mm)
    type_dims: Optional[FixtureDimensions] = None  # 器具種別名から抽出した寸法
    location: str = ""              # ★追加: 設置場所（色選定等に使用）


class CategoryConfig:
    """category_mapping.yaml から読み込む設定"""

    def __init__(self, config_path: Optional[Path] = None):
        if config_path is None:
            config_path = (
                Path(__file__).parent.parent / "config" / "category_mapping.yaml"
            )
        with open(config_path, "r", encoding="utf-8") as f:
            self._config = yaml.safe_load(f)

        self.bulb_type_map: dict[str, str] = self._config.get(
            "bulb_type_to_watt_form", {}
        )
        self.fixture_keywords: list[dict] = self._config.get(
            "fixture_type_keywords", []
        )
        self.emergency_keywords: list[str] = self._config.get(
            "emergency_keywords", []
        )
        self.emergency_sheets: dict[str, str] = self._config.get(
            "emergency_sheets", {}
        )
        self.color_mapping: dict[str, str] = self._config.get(
            "color_mapping", {}
        )
        self.construction_prices: dict[str, int] = self._config.get(
            "construction_prices", {}
        )
        self.waterproof_location_keywords: list[str] = self._config.get(
            "waterproof_location_keywords", []
        )
        self.emergency_mfr_pref: dict = self._config.get(
            "emergency_manufacturer_preference", {}
        )


class LEDMatcher:
    """LED選定エンジン

    選定優先度:
    1. 器具跡が残らない (サイズ適合)
    2. 安い (仕入価格順)
    3. デザイン性 (同価格帯内で選定)
    """

    def __init__(
        self,
        lineup_index: LineupIndex,
        config: Optional[CategoryConfig] = None,
        feedback_rules: Optional[list[dict]] = None,
    ):
        self.index = lineup_index
        self.config = config or CategoryConfig()
        self._size_cache: dict[str, FixtureDimensions] = {}
        self._feedback_rules = feedback_rules or []

    def match_all(
        self, fixtures: list[ExistingFixture],
    ) -> list[MatchResult]:
        """全器具を一括マッチング

        同一カテゴリキーの器具は同じLED商品を再利用する。
        """
        results = []
        seen_keys: dict[str, MatchResult] = {}

        for fixture in fixtures:
            if fixture.is_excluded:
                # 除外器具はマッチング不要
                results.append(MatchResult(
                    fixture=fixture,
                    category_key="",
                    confidence=0.0,
                    match_notes=fixture.exclusion_reason or "除外",
                ))
                continue

            result = self.match_fixture(fixture)

            # 同一カテゴリキーの場合は同じLED商品を再利用
            if result.category_key and result.category_key in seen_keys:
                prev = seen_keys[result.category_key]
                result.led_product = prev.led_product
            elif result.category_key:
                seen_keys[result.category_key] = result

            results.append(result)

        return results

    def match_fixture(self, fixture: ExistingFixture) -> MatchResult:
        """1つの既存器具に対してLED代替品を選定"""

        # Step 1: 器具分類
        classification = self._classify_fixture(fixture)

        # Step 1.5: フィードバックルールによる上書きチェック
        override = self._check_feedback_override(fixture, classification)
        if override is not None:
            return override

        # Step 2: 候補商品検索
        candidates = self._search_candidates(classification, fixture)

        if not candidates:
            return MatchResult(
                fixture=fixture,
                category_key="",
                confidence=0.0,
                match_notes=(
                    f"該当商品なし（シート:{classification.lineup_sheet}）"
                ),
                needs_review=True,
            )

        # Step 3: サイズフィルタ（器具跡が残らない）
        # ★注: type_dimsはaffinity scoringのみに使用。サイズフィルタには
        # fixture_size（実測値）のみ使用。名前由来の寸法でフィルタすると
        # 有効な候補が除外される（例: 430Φ→Φ380リニューアルが除外される）
        existing_dims = self._parse_cached(fixture.fixture_size or "")
        size_ok, size_ng = self._filter_by_size(
            candidates, existing_dims, classification.is_recessed,
        )

        # サイズ適合商品がなければ全候補を使用（要レビュー）
        needs_size_review = False
        if not size_ok:
            size_ok = candidates
            needs_size_review = True

        # Step 4: 同形後継機優先スコアリング + 価格ソート
        # ★改善: まず全候補で親和度を計算し、高親和度の候補を残す
        all_scored = [
            (p, self._successor_affinity(fixture, p, classification))
            for p in size_ok
        ]

        if not all_scored:
            pool = size_ok
            _affinity_map: dict[int, float] = {}
        else:
            # 最高親和度グループを特定
            max_affinity = max(s for _, s in all_scored)
            threshold = max(max_affinity * 0.8, max_affinity - 2.0)
            top_tier_scored = [(p, s) for p, s in all_scored if s >= threshold]

            # affinityマップ（ソート用）
            _affinity_map = {id(p): s for p, s in top_tier_scored}

            # 高親和度グループ内で価格あり/なしを分ける
            priced_top = [p for p, s in top_tier_scored
                          if p.purchase_price_total and p.purchase_price_total > 0]
            unpriced_top = [p for p, s in top_tier_scored
                           if not p.purchase_price_total or p.purchase_price_total <= 0]

            if priced_top:
                pool = priced_top
            elif unpriced_top:
                pool = unpriced_top
                needs_size_review = True
            else:
                pool = [p for p, s in top_tier_scored]

        # ★改善: 価格昇順 → 同価格ならaffinity降順でソート
        pool.sort(key=lambda p: (
            p.purchase_price_total or 999999,
            -_affinity_map.get(id(p), 0.0),
        ))

        # ★改善: デザイン選定前に親和度で絞り込み
        # design_scoreが親和度を無視して低親和度商品を選ぶ問題を防ぐ
        if _affinity_map and len(pool) > 1:
            max_pool_aff = max(
                _affinity_map.get(id(p), 0.0) for p in pool
            )
            narrow_threshold = max_pool_aff - 1.0
            narrowed = [
                p for p in pool
                if _affinity_map.get(id(p), 0.0) >= narrow_threshold
            ]
            if narrowed:
                pool = narrowed

        # Step 5: デザイン選定（同親和度帯・同価格帯内）
        best = self._pick_best_design(pool, classification)

        # Step 6: 結果構築
        confidence = self._calc_confidence(
            fixture, best, classification, needs_size_review,
        )
        notes = self._build_notes(
            fixture, best, classification, existing_dims, needs_size_review,
        )
        construction_price = self._estimate_construction_price(
            fixture, best, classification,
        )

        return MatchResult(
            fixture=fixture,
            led_product=best,
            category_key=best.name or "",
            confidence=confidence,
            match_notes=notes,
            construction_unit_price=construction_price,
            needs_review=needs_size_review or classification.fallback,
        )

    # ===== 分類 =====

    def _classify_fixture(
        self, fixture: ExistingFixture,
    ) -> FixtureClassification:
        """現調器具をラインナップシートにマッピング"""

        ft = fixture.fixture_type or ""
        bt = fixture.bulb_type or ""
        ft_norm = _normalize(ft)
        bt_norm = _normalize(bt)

        # 非常灯判定
        has_emergency = any(
            kw in ft_norm for kw in self.config.emergency_keywords
        )

        # 灯数判定
        bulb_count = 1
        if "2灯" in ft_norm or "2灯" in bt_norm:
            bulb_count = 2
        # ★Fix23: "1灯のみ使用"→2灯器具でも1灯として扱う
        if "1灯のみ" in bt_norm:
            bulb_count = 1

        # ★改善: 防水判定を2段階に分離
        # wp_from_location: 場所由来 → ハードフィルタ（防水製品のみ検索）
        # wp_from_fixture: 器具名由来 → ソフト優先（親和度ボーナスのみ）
        # 理由: 「防水640×170逆富士蛍光灯」でも屋内通路なら非防水LEDが適切
        loc_norm = _normalize(fixture.location or "")
        wp_from_location = fixture.is_waterproof or any(
            kw in loc_norm for kw in self.config.waterproof_location_keywords
        )
        wp_from_fixture = any(
            kw in ft_norm for kw in ("防滴", "防水", "防雨", "防湿")
        )
        # 屋外キーワードはfixture_type名にあっても場所由来扱い
        if "屋外" in ft_norm:
            wp_from_location = True

        # ★Fix15: 吊り下げ/天吊/両笠器具 → 防水ハードフィルタ緩和
        # 吊り下げ式・両笠式は屋根下が前提 → 防水はソフト優先（親和度のみ）に切替
        # 例: 駐輪場の吊り下げトラフ蛍光灯 → 非防水の反射笠付形が適切
        # ★Fix22: 両笠もカバー付き→屋根下前提で防水ハード緩和
        if "吊り下げ" in ft_norm or "天吊" in ft_norm or "両笠" in ft_norm:
            wp_from_location = False

        is_waterproof = wp_from_location or wp_from_fixture

        # 埋込判定（ダウンライト、天井埋込等）
        # ★注: 逆富士/ベースライトは直付（surface-mount）なので埋込に含めない
        is_recessed = any(
            kw in ft_norm for kw in (
                "ダウンライト", "DL", "埋込", "埋め込み",
            )
        )

        # 色温度
        color_pref = ""
        if fixture.color_temp:
            ct = _normalize(fixture.color_temp)
            color_pref = self.config.color_mapping.get(ct, "")

        # ★Fix13a: 電球種別の末尾から色温度を補完
        # 例: "FHT16 L" → L=電球色, "FL20 N" → N=昼白色
        if not color_pref and bt_norm:
            bt_parts = bt_norm.strip().split()
            if len(bt_parts) >= 2:
                color_suffix = bt_parts[-1]
                color_pref = self.config.color_mapping.get(color_suffix, "")

        # ★改善: 器具種別名からサイズ（直径・寸法）を抽出
        type_diameter, type_dims = self._extract_size_from_type(ft_norm)

        # ★改善: ワット形の決定（電球種別 → 器具種別名 → 寸法推定の順）
        watt_form = self._extract_watt_form(bt_norm)
        if not watt_form:
            watt_form = self._extract_watt_from_type(ft_norm, type_dims)

        # ラインナップシートの決定
        sheet = self._determine_sheet(
            ft_norm, bt_norm, watt_form,
            has_emergency, is_waterproof,
        )

        # ★Fix6: 薄型ブラケット → 屋外ブラケットへ補正
        # 幅100mm未満のブラケットは薄型のストリップ照明（屋外通路用）。
        # 天井・壁面の丸形ブラケット(□250等)とは形状が異なる。
        # 例: 380×75天井ブラケット → 10形 屋外ブラケット〈W104〉が適切
        if (sheet == "天井・壁面"
                and type_dims and type_dims.width_mm
                and type_dims.width_mm < 100
                and "ブラケット" in ft_norm
                and type_dims.length_mm and type_dims.length_mm > 200):
            sheet = "屋外ﾌﾞﾗｹｯﾄ"

        # ★Fix21: 細長直管ブラケット → 屋外ブラケットへ補正
        # 幅≤120mm, 長さ>500mmの細長ブラケットは直管蛍光灯用
        # 例: 20形壁面ブラケット100×660, 740×100直管形ブラケット
        # → 屋外ブラケット（10形/20形）が適切
        if ("ブラケット" in ft_norm
                and type_dims and type_dims.width_mm
                and type_dims.width_mm <= 120
                and type_dims.length_mm and type_dims.length_mm > 500
                and sheet != "屋外ﾌﾞﾗｹｯﾄ"):
            sheet = "屋外ﾌﾞﾗｹｯﾄ"

        # ★Fix24: 幅ベースの2灯推定
        # 20形/40形で幅180mm超 → 2灯式の可能性が高い
        # 例: 非常灯兼用直付蛍光灯 675×205 FL20 → 2灯/W230が適切
        # fixture_sizeからも幅を補完
        inferred_width = None
        if type_dims and type_dims.width_mm:
            inferred_width = type_dims.width_mm
        if inferred_width is None:
            fs_norm = _normalize(fixture.fixture_size or "")
            fs_dims = self._parse_cached(fs_norm)
            if fs_dims and fs_dims.width_mm:
                inferred_width = fs_dims.width_mm
        # ★Fix40: 埋込器具は2灯推定から除外
        # 埋込40形W190/埋込20形W200 → 1灯（埋込は幅広でも1灯が多い）
        if (bulb_count == 1
                and inferred_width is not None
                and inferred_width > 180
                and watt_form in ("20形", "40形")
                and "ブラケット" not in ft_norm
                and not is_recessed):
            bulb_count = 2

        # ★Fix10: 壁面ブラケットΦ → 天井・壁面へ補正
        # 壁面ブラケットにΦ（丸形直径）がある場合、丸形ブラケット＝天井・壁面が適切
        # 例: 壁面ブラケットΦ235 → 天井・壁面（ポーチ・支柱は長方形壁掛け灯用）
        if (sheet == "ﾎﾟｰﾁ・支柱"
                and "壁面ブラケット" in ft_norm
                and type_diameter > 0):
            sheet = "天井・壁面"

        # ★Fix46b: 壁面ブラケット□200+ → 天井・壁面へ補正
        # □200以上の壁面ブラケットは大型角形で天井・壁面シートが適切
        # 例: 壁面ブラケット□230(駐輪場) → 天井・壁面ﾌﾞﾗｹｯﾄ〈□250〉
        if (sheet == "ﾎﾟｰﾁ・支柱"
                and "壁面ブラケット" in ft_norm
                and type_dims and type_dims.width_mm
                and type_dims.length_mm
                and type_dims.width_mm == type_dims.length_mm  # 正方形
                and type_dims.width_mm >= 200):
            sheet = "天井・壁面"

        # ★Fix47: 玄関灯(廊下) → 天井壁面
        # 廊下に設置された「玄関灯」は屋内壁面灯→天井・壁面が適切
        # 玄関灯Φ → 筒形ブラケット → 天井・壁面が適切
        if (sheet == "ﾎﾟｰﾁ・支柱"
                and "玄関灯" in ft_norm
                and ("廊下" in loc_norm or type_diameter > 0)):
            sheet = "天井・壁面"

        # ★Fix48: 40形長尺ブラケット → 40形ベースライト
        # 120×1280等の長尺ブラケットは40形蛍光灯→ベースライト交換が適切
        # Fix21の屋外ブラケット判定(l>500)より先に40形判定を行う
        if (sheet == "屋外ﾌﾞﾗｹｯﾄ"
                and "ブラケット" in ft_norm
                and type_dims and type_dims.length_mm
                and type_dims.length_mm > 800
                and watt_form == "40形"):
            sheet = "40形蛍光灯"

        # ★Fix54: 壁面ブラケット(階段) → 天井壁面
        # 階段に設置された壁面ブラケット（寸法指定なし）は天井・壁面が適切
        # 例: 壁面ブラケット(A:階段)→天井壁面ブラケット〈Φ252/L/W/FCL20〉
        # ※□110/□120等の小型正方形ブラケットはポーチに残す
        if (sheet == "ﾎﾟｰﾁ・支柱"
                and "壁面ブラケット" in ft_norm
                and "階段" in loc_norm
                and type_diameter == 0
                and not (type_dims and type_dims.width_mm)):
            sheet = "天井・壁面"

        # ★Fix60: 壁面ブラケット(fixture_size φ) → 天井壁面
        # fixture_type にΦがなくても fixture_size に220φ等がある場合
        # 丸形ブラケット→天井・壁面が適切
        # 例: 壁面ブラケット(220φ, 外通路) → 天井ブラケット〈Φ220〉
        if (sheet == "ﾎﾟｰﾁ・支柱"
                and "壁面ブラケット" in ft_norm
                and type_diameter == 0):
            fs_text_route = _normalize(fixture.fixture_size or "")
            m_fs_phi = re.search(r'(\d+)\s*[Φφ]', fs_text_route)
            if not m_fs_phi:
                m_fs_phi = re.search(r'[Φφ]\s*(\d+)', fs_text_route)
            if m_fs_phi:
                fs_phi_val = float(m_fs_phi.group(1))
                if fs_phi_val >= 200:
                    sheet = "天井・壁面"

        # ★Fix62: 寸法付き格子ブラケット → ポーチ
        # "280×140格子ブラケット"等、寸法付きの格子ブラケットは装飾壁面灯→ポーチ
        # 寸法なしの"格子ブラケット"は天井壁面に残す（一般ブラケットの可能性）
        if (sheet == "天井・壁面"
                and "格子" in ft_norm
                and "ブラケット" in ft_norm
                and type_dims and type_dims.width_mm):
            sheet = "ﾎﾟｰﾁ・支柱"

        # ★Fix63: ブラケット(外壁) → ポーチ
        # "外壁"に設置された"ブラケット"は屋外壁面灯→ポーチ系が適切
        # "ブラケット"→天井壁面にルーティングされるのを修正
        if (sheet == "天井・壁面"
                and "外壁" in loc_norm
                and "ブラケット" in ft_norm
                and "壁面" not in ft_norm):
            sheet = "ﾎﾟｰﾁ・支柱"

        # ★Fix55: ダウンライトΦ450+ → 丸・四角(大)
        # 非常に大きいΦのダウンライトはラウンドベースライト等の大型製品
        # 通常のダウンライトシートにはΦ450のような大型品がない
        if (sheet == "ﾀﾞｳﾝﾗｲﾄ"
                and type_diameter >= 400):
            sheet = "丸・四角(大)"

        # ★Fix59: 投光器(階段) → スポットライトシートへ
        # 階段の投光器は小型ビーム球スポットが適切
        # 投光器・高天井シートにはスポットライト製品がなく候補なしになる
        if (sheet == "投光器・高天井"
                and "階段" in loc_norm):
            sheet = "ｽﾎﾟｯﾄﾗｲﾄ"

        return FixtureClassification(
            lineup_sheet=sheet,
            watt_form=watt_form,
            has_emergency=has_emergency,
            is_recessed=is_recessed,
            is_waterproof=is_waterproof,
            wp_hard_filter=wp_from_location,
            color_pref=color_pref,
            bulb_count=bulb_count,
            fallback=sheet == "天井・壁面",
            type_diameter=type_diameter,
            type_dims=type_dims,
            location=_normalize(fixture.location or ""),
        )

    def _extract_watt_form(self, bulb_type: str) -> str:
        """電球種別からワット形を抽出"""
        # 完全一致を試す
        for key, form in self.config.bulb_type_map.items():
            if bulb_type.upper().startswith(key.upper()):
                return form

        # 数値抽出でフォールバック
        m = re.search(r'(\d+)\s*[Ww]', bulb_type)
        if m:
            watt = int(m.group(1))
            if watt <= 25:
                return "20形"
            elif watt <= 50:
                return "40形"

        return ""

    def _extract_watt_from_type(
        self, fixture_type: str,
        type_dims: Optional[FixtureDimensions] = None,
    ) -> str:
        """★改善: 器具種別名からワット形を推定

        例: "20形1灯用" → "20形"
            "非常灯付き20形1灯用" → "20形"
            "防水640×170逆富士蛍光灯" → "20形" (640mm≈FL20)
        """
        # パターン1: 明示的な "20形" / "40形" 表記
        m = re.search(r'(20|40)形', fixture_type)
        if m:
            return f"{m.group(1)}形"

        # ★Fix36: "20W" / "40W" 表記 → ワット形推定
        # 例: "20W直管形非常灯" → "20形"
        m_w = re.search(r'(20|40)\s*[Ww]', fixture_type)
        if m_w:
            return f"{m_w.group(1)}形"

        # パターン2: 長さから推定（蛍光灯/逆富士/トラフ/直管形/兼用照明）
        # ★Fix49: "兼用照明" を追加（2灯式630×200非常灯兼用照明→20形）
        if type_dims and type_dims.length_mm:
            length = type_dims.length_mm
            if any(kw in fixture_type for kw in
                   ("蛍光灯", "逆富士", "トラフ", "直管", "灯用", "兼用照明")):
                if length <= 750:   # FL20 ≈ 580mm + 器具 ≈ 600-740mm（ブラケット含む）
                    return "20形"
                elif length <= 1300:  # FL40 ≈ 1198mm + 器具 ≈ 1200-1250mm
                    return "40形"

        # パターン3: "1灯用" "2灯用" を含む場合、サイズから推定
        if "灯用" in fixture_type and type_dims and type_dims.length_mm:
            length = type_dims.length_mm
            if length <= 700:
                return "20形"
            elif length <= 1300:
                return "40形"

        return ""

    @staticmethod
    def _extract_size_from_type(
        fixture_type: str,
    ) -> tuple[float, Optional[FixtureDimensions]]:
        """★改善: 器具種別名から直径・寸法を抽出

        例: "300Φ天井ブラケット" → diameter=300
            "430Φ非常灯兼用ブラケット" → diameter=430
            "140×270壁面ブラケット" → 140×270
            "防水640×170逆富士蛍光灯" → 170×640
        """
        diameter = 0.0
        dims = None

        # パターン1: NNNΦまたはΦNNN（直径）
        m = re.search(r'(\d+)\s*[Φφ]', fixture_type)
        if m:
            diameter = float(m.group(1))
            dims = FixtureDimensions(diameter_mm=diameter, raw=fixture_type)
            return diameter, dims

        m = re.search(r'[Φφ]\s*(\d+)', fixture_type)
        if m:
            diameter = float(m.group(1))
            dims = FixtureDimensions(diameter_mm=diameter, raw=fixture_type)
            return diameter, dims

        # パターン2: □NNN / ■NNN（正方形寸法）
        # ★Fix46a: 壁面ブラケット□230, ダウンライト□125 等の正方形寸法
        m = re.search(r'[□■](\d+)', fixture_type)
        if m:
            sq = float(m.group(1))
            dims = FixtureDimensions(
                width_mm=sq, length_mm=sq, raw=fixture_type,
            )
            return 0.0, dims

        # パターン3: NNN×NNN（幅×長さ）
        m = re.search(r'(\d+)\s*[×xX]\s*(\d+)', fixture_type)
        if m:
            v1 = float(m.group(1))
            v2 = float(m.group(2))
            w = min(v1, v2)
            l = max(v1, v2)
            dims = FixtureDimensions(
                width_mm=w, length_mm=l, raw=fixture_type,
            )
            return 0.0, dims

        # パターン4: wNNN（幅のみ、例: トラフ蛍光灯w70）
        # ★Fix41a: "w70"/"W150"等の幅表記をパース
        # ※ "20W"等のワット表記を除外するため先頭が数字でない位置のみ
        m_w = re.search(r'(?<!\d)[wW](\d{2,3})(?!\d)', fixture_type)
        if m_w:
            w_val = float(m_w.group(1))
            if 30 <= w_val <= 500:  # 妥当な幅範囲
                dims = FixtureDimensions(
                    width_mm=w_val, raw=fixture_type,
                )
                return 0.0, dims

        return 0.0, None

    def _determine_sheet(
        self, fixture_type: str, bulb_type: str,
        watt_form: str, has_emergency: bool,
        is_waterproof: bool,
    ) -> str:
        """ラインナップシート名を決定"""

        # 非常灯は専用シートへ
        if has_emergency:
            return self.config.emergency_sheets.get(
                watt_form,
                self.config.emergency_sheets.get("default", "その他非常灯"),
            )

        # 照明種別キーワードマッチ
        for rule in self.config.fixture_keywords:
            keyword = rule["keyword"]
            if keyword in fixture_type:
                # 防水分岐がある場合
                if "sheet_by_waterproof" in rule:
                    wp_map = rule["sheet_by_waterproof"]
                    # YAMLはtrue/falseをboolean型で解析するため
                    # boolean/文字列の両方で検索する
                    result = wp_map.get(is_waterproof)
                    if result is None:
                        result = wp_map.get(
                            str(is_waterproof).lower(),
                            wp_map.get(False, wp_map.get("false", "天井・壁面")),
                        )
                    return result
                # 直接シート指定
                sheet = rule.get("sheet")
                if sheet:
                    # ★Fix9: 逆富士/トラフは検出済みwatt_formで動的にシート決定
                    # YAML定義は "20形蛍光灯" だが、器具名に "40形" があれば "40形蛍光灯" へ
                    # 例: "40形1灯用トラフ蛍光灯" → watt_form="40形" → "40形蛍光灯"
                    if keyword in ("逆富士", "トラフ") and watt_form in ("20形", "40形"):
                        return f"{watt_form}蛍光灯"
                    return sheet

        # 電球種別からのフォールバック
        if watt_form in ("20形", "40形"):
            return f"{watt_form}蛍光灯"
        if watt_form:
            return watt_form  # "ﾀﾞｳﾝﾗｲﾄ" 等

        # 最終フォールバック
        return "天井・壁面"

    # ===== 候補検索 =====

    # ===== フィードバックルール適用 =====

    def _check_feedback_override(
        self,
        fixture: ExistingFixture,
        classification: FixtureClassification,
    ) -> Optional[MatchResult]:
        """フィードバック学習ルールによるLED選定の上書き

        過去のフィードバックで修正実績が1回以上あるルールに合致する場合、
        通常の選定ロジックをスキップしてルール指定の製品を返す。
        """
        if not self._feedback_rules:
            return None

        ft_norm = _normalize(fixture.fixture_type or "").lower()
        if not ft_norm:
            return None

        for rule in self._feedback_rules:
            count = rule.get("count", 0)
            if count < 1:
                continue

            rule_ft = _normalize(rule.get("fixture_type", "")).lower()
            if not rule_ft:
                continue

            # 器具種別の照合（部分一致）
            if rule_ft not in ft_norm and ft_norm not in rule_ft:
                continue

            correct_name = rule.get("correct_selection", "")
            if not correct_name:
                continue

            # ラインナップから指定製品を検索
            candidates = self._search_candidates(classification, fixture)
            correct_norm = _normalize(correct_name).lower()

            matched_product = None
            for p in candidates:
                if p.name and _normalize(p.name).lower() == correct_norm:
                    matched_product = p
                    break

            if matched_product:
                construction_price = self._estimate_construction_price(
                    fixture, matched_product, classification,
                )
                logger.info(
                    f"フィードバックルール適用: {fixture.fixture_type} → "
                    f"{matched_product.name} (修正実績: {count}回)"
                )
                return MatchResult(
                    fixture=fixture,
                    led_product=matched_product,
                    category_key=matched_product.name or "",
                    confidence=0.95,
                    match_notes=(
                        f"フィードバックルール適用 (修正実績: {count}回)"
                    ),
                    construction_unit_price=construction_price,
                )

        return None

    def _search_candidates(
        self, cls: FixtureClassification,
        fixture: Optional[ExistingFixture] = None,
    ) -> list[LEDProduct]:
        """分類に基づいてラインナップから候補を検索"""
        # ★改善: 防水フィルタはwp_hard_filter（場所由来）のみ適用
        # 器具名由来の防水（防水逆富士 etc.）は親和度ボーナスのみ
        skip_waterproof = cls.lineup_sheet in (
            "ポール灯", "EEスイッチ他", "門柱灯", "外部・ﾊﾞｲﾊﾟｽ",
        )
        wp_filter = cls.wp_hard_filter if (cls.wp_hard_filter and not skip_waterproof) else None
        results = self.index.search(
            sheet_name=cls.lineup_sheet,
            waterproof=wp_filter,
            lighting_color=cls.color_pref if cls.color_pref else None,
        )

        # 灯数フィルタ
        if cls.bulb_count == 2:
            filtered = [p for p in results if p.name and "2灯" in p.name]
            if filtered:
                return filtered

        return results

    # ===== サイズフィルタ =====

    def _filter_by_size(
        self,
        candidates: list[LEDProduct],
        existing_dims: FixtureDimensions,
        is_recessed: bool,
    ) -> tuple[list[LEDProduct], list[LEDProduct]]:
        """サイズ適合フィルタ"""
        ok, ng = [], []

        if not existing_dims.has_dimensions:
            # 既存器具のサイズ不明 → 全て候補
            return candidates, []

        for p in candidates:
            led_dims = self._parse_cached(p.fixture_size or "")
            if not led_dims.has_dimensions:
                ok.append(p)  # LED側サイズ不明は候補に含める
                continue

            compatible, _ = is_size_compatible(
                existing_dims, led_dims, is_recessed=is_recessed,
            )
            if compatible:
                ok.append(p)
            else:
                ng.append(p)

        return ok, ng

    # ===== デザイン選定 =====

    def _pick_best_design(
        self, candidates: list[LEDProduct],
        cls: FixtureClassification,
    ) -> LEDProduct:
        """同価格帯内でデザイン性が最良の商品を選定"""
        if not candidates:
            raise ValueError("候補なし")

        if len(candidates) == 1:
            return candidates[0]

        # 最安値を基準に500円以内の商品を候補
        if candidates[0].purchase_price_total and candidates[0].purchase_price_total > 0:
            cheapest = candidates[0].purchase_price_total
            price_tier = [
                p for p in candidates
                if p.purchase_price_total
                and p.purchase_price_total <= cheapest + 500
            ]
            if price_tier:
                candidates = price_tier

        # デザインスコアでソート
        candidates.sort(key=lambda p: self._design_score(p, cls), reverse=True)
        return candidates[0]

    def _design_score(self, product: LEDProduct, cls: FixtureClassification) -> float:
        """デザインスコア（高い方が良い）"""
        score = 0.0

        # ★改善: メーカー優先度（正解データに基づく）
        # 三菱が蛍光灯系・非常灯ともに最安値かつ最もコンパクト
        mfr = product.manufacturer or ""
        if "三菱" in mfr:
            score += 3
        elif "東芝" in mfr or "TOSHIBA" in mfr:
            score += 2.5
        elif "パナソニック" in mfr or "Panasonic" in mfr:
            score += 2
        elif "コイズミ" in mfr:
            score += 1.5
        elif "遠藤" in mfr:
            score += 1
        elif "オーデリック" in mfr:
            score += 1

        # 非常灯は三菱優先（コンパクト + 安価）
        if cls.has_emergency:
            pref = self.config.emergency_mfr_pref
            preferred_mfr = pref.get("preferred", "")
            bonus = pref.get("bonus_score", 2)
            if preferred_mfr and preferred_mfr in mfr:
                score += bonus

        # ★改善: 器具色 — マンション共用部はダーク系が標準
        # 正解データ: K(ブラック), DS(ダークシルバー)が一貫して選ばれている
        # ★Fix5: エントランス/ロビー等の明るい場所はホワイトを優先
        fc = product.fixture_color or ""
        loc = cls.location or ""
        is_bright_area = any(
            kw in loc for kw in ("エントランス", "ロビー", "玄関")
        )
        if is_bright_area:
            # 明るいエリア → ホワイト/明るい色を優先
            if "ホワイト" in fc or "白" in fc:
                score += 2.0
            elif "シルバー" in fc or "グレー" in fc:
                score += 1.0
            elif "ブラック" in fc or "黒" in fc:
                score += 0.5
        else:
            # 通路等 → ダーク系を優先
            if "ブラック" in fc or "黒" in fc:
                score += 1.5
            elif "ダークシルバー" in fc or "ダーク" in fc:
                score += 1.0
            elif "シルバー" in fc or "グレー" in fc:
                score += 0.5

        # ★Fix18: 色温度一致ボーナス
        # 色温度が指定されている場合、一致する商品を優先
        # 例: FHT16 L → 電球色 → L商品を選択
        if cls.color_pref and product.lighting_color:
            if cls.color_pref in product.lighting_color:
                score += 3.0

        # ★改善: 非常灯リニューアル丸形 → 大径を強く優先
        # メーカー差(0.5)より径差の影響を大きくする
        prod_name = product.name or ""
        if cls.has_emergency and ("ﾘﾆｭｰｱﾙ" in prod_name or "リニューアル" in prod_name):
            m_phi = re.search(r'[Φφ](\d+)', prod_name)
            if m_phi:
                score += float(m_phi.group(1)) / 50.0  # Φ380→+7.6, Φ349→+6.98

        # 光束が高い方が良い
        if product.lumens:
            try:
                lm = float(product.lumens)
                score += min(lm / 1000, 3)  # 最大3点
            except (ValueError, TypeError):
                pass

        return score

    # ===== 同形後継機親和度 =====

    def _successor_affinity(
        self,
        fixture: ExistingFixture,
        product: LEDProduct,
        cls: FixtureClassification,
    ) -> float:
        """既存器具とLED商品の形状類似度（後継機親和度）を算出

        高い値ほど既存器具に近い形の後継機。
        同じスコアなら価格で決まるため、後継機 + 安価 が最優先される。

        Returns:
            0.0〜15.0 のスコア
        """
        score = 0.0
        ft_norm = _normalize(fixture.fixture_type or "")
        bt_norm = _normalize(fixture.bulb_type or "")
        prod_name = _normalize(product.name or "")
        prod_equiv = _normalize(product.watt_equivalent or "")

        # --- 1. ワット形一致（後継機の最重要指標） ---
        if cls.watt_form and cls.watt_form in prod_equiv:
            score += 4.0
        elif cls.watt_form:
            m = re.search(r'(\d+)', cls.watt_form)
            if m and m.group(1) in prod_equiv:
                score += 2.0

        # --- 2. 器具形状の一致 ---
        shape_keywords = [
            ("ブラケット", "ﾌﾞﾗｹｯﾄ"),
            ("ダウンライト", "ﾀﾞｳﾝﾗｲﾄ", "DL"),
            ("シーリング", "ｼｰﾘﾝｸﾞ"),
            ("ポーチ", "ﾎﾟｰﾁ"),
            ("直付", "直付"),
            ("埋込", "埋込"),
            ("逆富士", "逆富士"),
            ("ベースライト", "ﾍﾞｰｽﾗｲﾄ"),
            ("トラフ", "ﾄﾗﾌ"),
            ("投光", "投光"),
            ("スポット", "ｽﾎﾟｯﾄ"),
            ("足元", "ﾌｯﾄﾗｲﾄ"),
            ("門柱", "門柱"),
            ("階段", "階段"),
            ("コーンランプ", "ｺｰﾝﾗﾝﾌﾟ"),
        ]
        for kw_group in shape_keywords:
            existing_match = any(kw in ft_norm for kw in kw_group)
            product_match = any(kw in prod_name for kw in kw_group)
            if existing_match and product_match:
                score += 3.0
                break

        # --- 3. 電球種別の互換性 ---
        bulb_compat = {
            "FL": ["直管", "蛍光灯", "FL"],
            "FDL": ["コンパクト", "FDL", "ﾀﾞｳﾝﾗｲﾄ"],
            "FHT": ["コンパクト", "FHT"],
            "FPL": ["コンパクト", "FPL"],
            "FCL": ["丸形", "サークル", "FCL", "シーリング"],
            "白熱": ["白熱", "電球", "E26", "E17"],
        }
        for prefix, compat_kws in bulb_compat.items():
            if prefix in bt_norm.upper():
                if any(kw in prod_name for kw in compat_kws):
                    score += 2.0
                break

        # --- 4. ★改善: 防水属性の一致 ---
        # wp_hard_filter=True (場所由来) → 防水製品に強いボーナス
        # wp_hard_filter=False, is_waterproof=True (器具名由来) → ソフトボーナス
        # ★Fix42: 非防水器具 → 非防水製品をやや優先
        if cls.wp_hard_filter and product.is_waterproof:
            score += 1.0
        elif cls.is_waterproof and not cls.wp_hard_filter:
            # ★Fix25改善: 器具名由来防水のソフトアフィニティ（弱め）
            # 「防水逆富士」でも通路設置→非防水LEDが正解のケースが多い
            if product.is_waterproof:
                score += 0.5
            # 非防水も大きなペナルティなし（他のスコアで決まる）
        elif not cls.is_waterproof:
            # ★Fix42: 非防水器具は非防水LED製品をやや優先
            # ただし通路等で防水を選ぶケースもあるため控えめなペナルティ
            if not product.is_waterproof:
                score += 0.8
            else:
                score -= 0.5

        # --- 5. ★改善: 直径近似マッチング（連続スコア） ---
        # type_diameterまたはfixture_sizeから直径を取得
        existing_phi = cls.type_diameter
        if existing_phi == 0:
            # fixture_sizeからΦを補完（例: "320φ", "200φ"）
            fs = _normalize(fixture.fixture_size or "")
            m_fs = re.search(r'(\d+)\s*[Φφ]', fs)
            if not m_fs:
                m_fs = re.search(r'[Φφ]\s*(\d+)', fs)
            if m_fs:
                existing_phi = float(m_fs.group(1))

        if existing_phi > 0:
            prod_phi = self._extract_product_diameter(prod_name)
            if prod_phi > 0:
                diff = abs(existing_phi - prod_phi)
                if cls.has_emergency:
                    # 非常灯: サイズ近似より丸形リニューアル大型を優先
                    is_renewal = ("リニューアル" in prod_name
                                  or "ﾘﾆｭｰｱﾙ" in prod_name)
                    if is_renewal and "丸形" in prod_name:
                        score += 6.0
                        # 大きいΦを優先（380>330）
                        score += prod_phi / 100.0
                        # ★Fix12: 跡カバレッジをより強く重視
                        # Φ310→Φ380(cov=70)とΦ349(cov=39)の差を拡大
                        # 50mm以上のカバレッジは取付跡を確実に隠せる
                        if prod_phi >= existing_phi + 50:
                            score += 4.0  # 余裕あるカバレッジ
                        elif prod_phi >= existing_phi + 30:
                            score += 2.5  # 十分なカバレッジ
                        elif prod_phi >= existing_phi:
                            score += 1.0  # 最低限のカバレッジ
                        # ★Fix1: カバー不足ペナルティ（大型既存器具で跡が残る）
                        # 430Φ既存→Φ380(gap=50)vsΦ349(gap=81)の差を拡大
                        elif prod_phi < existing_phi:
                            gap = existing_phi - prod_phi
                            score -= gap * 0.04
                    else:
                        score += max(0.0, 2.0 - diff * 0.03)
                else:
                    # ★改善: 非非常灯: カバレッジ＋近接バランス型
                    # 10-30mmの適度なカバレッジが理想だが、
                    # ほぼ同サイズ(cov<10)も許容（FCL等で正当な選択）
                    if prod_phi >= existing_phi:
                        coverage = prod_phi - existing_phi
                        if coverage < 10:
                            score += 4.0  # ほぼ同サイズ
                        elif coverage <= 30:
                            score += 5.0  # 理想的カバレッジ
                        elif coverage <= 80:
                            score += 3.5  # やや大きい
                        else:
                            score += 1.5  # 大きすぎ
                    else:
                        # LED < 既存 → 跡残りリスク
                        shortfall = existing_phi - prod_phi
                        score += max(0.0, 2.0 - shortfall * 0.05)
                    # ★Fix20: 器具種別名に明示的Φ指定 → 完全一致に追加ボーナス
                    # 例: "壁面ブラケットΦ235" → Φ235製品をΦ264より強く優先
                    # type_diameter>0 = Φが器具名に含まれる（サイズ欄の数値とは区別）
                    # fixture_sizeの200φ等は柔軟なカバレッジが適切、
                    # 器具名のΦ235は専門家が明記した正確な寸法
                    if cls.type_diameter > 0 and diff < 5:
                        score += 3.0  # 明示Φの完全一致

        # --- 6. ★改善: リニューアル品優先（非非常灯の大型丸形） ---
        if (cls.type_diameter >= 300 and not cls.has_emergency):
            if "リニューアル" in prod_name or "ﾘﾆｭｰｱﾙ" in prod_name:
                score += 2.0

        # --- 7. ★改善+Fix38b: ポール灯ルーティング ---
        # ★Fix38b改善: 場所によりTランプ/コーンランプを使い分け
        # 駐車場/駐輪場の大型ポール灯 → コーンランプ（電池内蔵）
        # それ以外（建物前/建物周辺等）→ Tランプバイパス交換
        if "ポール灯" in ft_norm:
            is_parking = any(
                kw in cls.location
                for kw in ("駐車場", "駐輪場")
            )
            if is_parking:
                # 大型ポール灯(駐車場) → コーンランプ、電池内蔵優先
                if "コーンランプ" in prod_name or "ｺｰﾝﾗﾝﾌﾟ" in prod_name:
                    score += 5.0
                    if "電池内蔵" in prod_name or "電源内蔵" in prod_name:
                        score += 2.0
                elif "直管" in prod_name:
                    score -= 3.0
            else:
                # 建物前/周辺等 → 正解はTランプだがラインナップに未登録
                # ★注: ラインナップにTランプ商品がないため直管LEDバイパスが次善策
                # コーンランプは庭園/建物前ポール灯には大型すぎ
                if "直管" in prod_name and "LED" in prod_name:
                    score += 3.0
                elif ("バイパス" in prod_name or "ﾊﾞｲﾊﾟｽ" in prod_name):
                    score += 2.0
                # コーンランプは非駐車場ポール灯に不適
                if "コーンランプ" in prod_name or "ｺｰﾝﾗﾝﾌﾟ" in prod_name:
                    score -= 3.0
                if "ポール灯" in prod_name and "コーンランプ" in prod_name:
                    score -= 2.0

        # --- 8. ★改善: 器具種別名のキーワードと商品名の直接マッチング ---
        # 階段灯→非常用階段灯、EEスイッチ→EEスイッチ、逆富士→ベースライト
        type_product_affinity = [
            ("階段灯", ["階段灯"]),
            ("EEスイッチ", ["EEスイッチ", "EEスイッチ"]),
            ("トラフ", ["トラフ", "ﾄﾗﾌ"]),  # ★Fix41b: トラフ→トラフ製品直接マッチ
            ("逆富士", ["ベースライト", "ﾍﾞｰｽﾗｲﾄ"]),
            ("蛍光灯", ["ベースライト", "ﾍﾞｰｽﾗｲﾄ"]),
        ]
        for type_kw, prod_kws in type_product_affinity:
            if type_kw in ft_norm:
                if any(kw in prod_name for kw in prod_kws):
                    score += 3.0
                break

        # --- 9. ★改善: 壁面ブラケット → ポーチ/支柱灯商品を優先 ---
        if "壁面ブラケット" in ft_norm or "壁面ﾌﾞﾗｹｯﾄ" in ft_norm:
            # ★Fix8: 短い器具(length<200)には長方形は大きすぎ → ボーナスなし
            # 例: 140×100壁面ブラケットに長方形A(255mm)は大きすぎ、
            # 直管型B(140mm)が適切
            fixture_length = None
            if cls.type_dims and cls.type_dims.length_mm:
                fixture_length = cls.type_dims.length_mm
            if "長方形" in prod_name:
                if fixture_length and fixture_length < 200:
                    pass  # 短い器具には長方形ボーナスなし
                else:
                    score += 4.0
            # ★Fix8追加: 短い器具には直管型がコンパクトで適切
            # 140×100壁面→直管型B(140×110)が配管干渉を避ける
            if fixture_length and fixture_length < 200:
                if "直管型" in prod_name:
                    score += 3.0
            # ポーチ・支柱シートでは、支柱灯/ﾎﾟｰﾁﾗｲﾄ商品を優先
            if "支柱灯" in prod_name or "ﾎﾟｰﾁﾗｲﾄ" in prod_name:
                score += 3.0
            elif "ポーチライト" in prod_name:
                score += 3.0
            # ★Fix74: ポーチ・支柱シートで「屋外ブラケット」製品はペナルティ
            # ※屋外ブラケットシートの場合は除外（正当な屋外ブラケット選定）
            if (cls.lineup_sheet == "ﾎﾟｰﾁ・支柱"
                    and ("屋外ﾌﾞﾗｹｯﾄ" in prod_name
                         or "屋外ブラケット" in prod_name)):
                score -= 3.0

        # --- 9b. ★Fix68: ポーチ・支柱シートの全ブラケット→支柱灯・ポーチライト優先 ---
        # Rule 9は「壁面ブラケット」のみだが、格子ブラケット/ブラケット(外壁)/玄関灯等
        # もポーチ・支柱シートでは支柱灯・ポーチライト商品を優先
        # 「和風屋外ブラケット」等のブラケット名の商品よりポーチライト系を選ぶ
        if (cls.lineup_sheet == "ﾎﾟｰﾁ・支柱"
                and "壁面ブラケット" not in ft_norm):
            # 支柱灯・ポーチライト系をボーナス
            if "支柱灯" in prod_name or "ﾎﾟｰﾁﾗｲﾄ" in prod_name or "ポーチライト" in prod_name:
                score += 3.0
            # 「屋外ブラケット」「和風」名の製品はペナルティ
            if "屋外ﾌﾞﾗｹｯﾄ" in prod_name or "屋外ブラケット" in prod_name:
                score -= 2.0

        # --- 10. ★改善: 1灯/2灯の整合性 ---
        if cls.bulb_count == 1:
            if "2灯" in prod_name:
                score -= 3.0  # 1灯器具に2灯商品はペナルティ
        elif cls.bulb_count == 2:
            if "2灯" in prod_name:
                score += 2.0  # 2灯器具に2灯商品はボーナス

        # --- 11. ★改善: 丸形直径→FCL推定（300Φ→FCL30, 200Φ→FCL20）---
        # ★Fix2: 天井/シーリング/丸形のみ適用（壁面ブラケットはFCL使わない）
        # ブラケット200φに不適切にFCL20推定が適用されるのを防ぐ
        if existing_phi > 0 and not cls.has_emergency:
            is_ceiling_type = any(
                kw in ft_norm for kw in ("天井", "シーリング", "丸形")
            )
            has_fcl_bulb = "FCL" in bt_norm.upper()
            if is_ceiling_type or has_fcl_bulb:
                inferred_fcl = ""
                if 250 <= existing_phi <= 350:
                    inferred_fcl = "FCL30"
                elif 150 <= existing_phi <= 250:
                    inferred_fcl = "FCL20"
                if inferred_fcl and inferred_fcl in prod_name:
                    score += 3.0

        # --- 12. ★改善: 非常灯一体型 > 別置 ---
        # ★注: 直管形非常灯の一体型/別置は物件単位の判断のため
        #       システムでは一体型を標準として選定（より一般的）
        if cls.has_emergency:
            if "一体型" in prod_name:
                score += 3.0
            elif "別置" in prod_name:
                score -= 2.0

        # --- 13. ★改善: 非埋込器具に埋込製品はペナルティ ---
        # ★Fix64: 直付器具が埋込製品を選択するのを抑制（特に非常灯）
        if not cls.is_recessed:
            if "埋込" in prod_name:
                score -= 3.0

        # --- 14. ★改善: 幅・長さベース親和度 ---
        existing_width = None
        existing_length = None
        if cls.type_dims and cls.type_dims.width_mm:
            existing_width = cls.type_dims.width_mm
        if cls.type_dims and cls.type_dims.length_mm:
            existing_length = cls.type_dims.length_mm
        # fixture_sizeからの補完
        if existing_width is None:
            fs_dims = self._parse_cached(fixture.fixture_size or "")
            if fs_dims.width_mm:
                existing_width = fs_dims.width_mm
            if fs_dims.length_mm and existing_length is None:
                existing_length = fs_dims.length_mm

        if existing_width and existing_width > 0:
            # 狭い器具（幅100mm未満）→ トラフライト優先
            if existing_width < 100:
                if "トラフ" in prod_name or "ﾄﾗﾌ" in prod_name:
                    score += 4.0

            # LED商品の幅を取得（商品名のWパターン or fixture_sizeから）
            prod_width = None
            prod_length = None
            m_w = re.search(r'W(\d+)', prod_name)
            if m_w:
                prod_width = float(m_w.group(1))
            if prod_width is None:
                led_dims = self._parse_cached(product.fixture_size or "")
                if led_dims.width_mm:
                    prod_width = led_dims.width_mm
                if led_dims.length_mm:
                    prod_length = led_dims.length_mm

            if prod_width and prod_width > 0:
                if prod_width >= existing_width:
                    # LED幅が既存以上: 跡隠し可能
                    coverage = prod_width - existing_width
                    score += max(0.0, 3.0 - coverage * 0.03)
                    score += 1.0  # カバレッジボーナス
                else:
                    # LED幅が既存未満: 跡残りリスク
                    shortfall = existing_width - prod_width
                    score += max(0.0, 1.0 - shortfall * 0.05)

            # 長さの比較（利用可能な場合）
            if existing_length and prod_length:
                if prod_length >= existing_length:
                    coverage_l = prod_length - existing_length
                    score += max(0.0, 2.0 - coverage_l * 0.02)
                    score += 0.5
                else:
                    shortfall_l = existing_length - prod_length
                    score += max(0.0, 0.5 - shortfall_l * 0.03)

        # --- 15. ★改善: EEスイッチ JIS1L形 > 住宅用 ---
        if "EEスイッチ" in ft_norm:
            if "JIS" in prod_name:
                score += 2.0
            elif "住宅用" in prod_name:
                score -= 1.0

        # --- 16. ★Fix7: 屋外ブラケット形式名優先 ---
        # 屋外ブラケットでは "10形"/"20形" の標準形式商品を優先
        # コンパクト+ポリ台よりも直接交換型が適切
        if cls.lineup_sheet == "屋外ﾌﾞﾗｹｯﾄ":
            if re.match(r'(10|20|40)形', prod_name):
                score += 2.0

        # --- 17. ★Fix3: 天吊形ペナルティ ---
        # 天吊形（チェーン/パイプ吊り下げ）は特殊な取付形態。
        # 器具名に"天吊"がない場合、天吊形商品は不適切。
        # これにより通路2灯蛍光灯で防雨・防湿型が天吊形より優先される。
        if "天吊形" in prod_name and "天吊" not in ft_norm:
            score -= 2.0

        # --- 18. ★Fix11: 埋込/直付非常灯 → 非常専用照明優先 ---
        # 「埋込非常灯Φ100」「直付非常灯Φ150」等の小型専用非常灯は
        # 丸形リニューアルブラケットではなく非常専用照明が適切
        if cls.has_emergency:
            is_dedicated = (
                ("埋込" in ft_norm and "非常灯" in ft_norm)
                or ("直付" in ft_norm and "非常灯" in ft_norm)
            )
            # ただし「非常灯兼用天井ブラケット」等の兼用器具は除く
            is_combined = "兼用" in ft_norm or "付き" in ft_norm
            if is_dedicated and not is_combined:
                if "非常専用照明" in prod_name:
                    score += 12.0
                    # 取付方式（埋込/直付）一致でさらにボーナス
                    if "埋込" in ft_norm and "埋込" in prod_name:
                        score += 3.0
                    elif "直付" in ft_norm and "直付" in prod_name:
                        score += 3.0
                else:
                    # 非専用照明を大幅にペナルティ
                    score -= 5.0

        # --- 19. ★Fix13b: ダウンライトΦ完全一致 + 形式優先 ---
        # ダウンライトは天井穴径が固定 → Φ完全一致が理想
        # LED一体形が標準的な交換先、フラット形はリニューアル用
        if cls.is_recessed and existing_phi > 0:
            prod_phi_dl = self._extract_product_diameter(prod_name)
            if prod_phi_dl > 0:
                # Φ完全一致ボーナス
                if abs(existing_phi - prod_phi_dl) < 5:
                    score += 3.0
                # フラット形（穴径変換）は既存Φが直接マッチする場合不要
                if ("フラット形" in prod_name or "ﾌﾗｯﾄ形" in prod_name):
                    score -= 3.0
            # ★Fix17: LED一体形ダウンライトを優先
            # LED一体形が最もコスト効率が良く、標準的な交換先
            if "LED一体形" in prod_name or "LED一体形" in prod_name:
                score += 3.0
            elif "ランプ型" in prod_name or "ﾗﾝﾌﾟ型" in prod_name:
                score -= 1.0
            # ★Fix13c: FHT/FDLワット数での候補絞り込み
            # FHT16(16W) → 60w相当が適切、FHT42(42W) → 100w相当は過大
            if bt_norm:
                m_w = re.search(r'(\d+)', bt_norm)
                if m_w:
                    bulb_watt = int(m_w.group(1))
                    prod_watt_match = re.search(r'(\d+)w', prod_name)
                    if prod_watt_match:
                        prod_watt = int(prod_watt_match.group(1))
                        # 電球ワット数と商品のワット相当が近い方が良い
                        if bulb_watt <= 25 and prod_watt <= 60:
                            score += 2.0
                        elif bulb_watt <= 25 and prod_watt > 60:
                            score -= 2.0
                        elif bulb_watt > 25 and prod_watt > 60:
                            score += 2.0
            else:
                # ★Fix19: 電球種別未指定の場合、特定ランプ型製品にペナルティ
                # FHT42W等のランプ特定製品は交換先としてニッチ
                # 汎用的な60w/100w相当を優先すべき
                # 例: ダウンライトΦ150(電球不明) → 60w/Tが適切、FHT42W/Tではない
                # ★Fix65: ペナルティ強化 -1.0 → -3.0
                if re.search(r'FHT\d+', prod_name) or re.search(r'FDL\d+', prod_name):
                    score -= 3.0
                # ★Fix65b: 汎用60w相当を優先
                prod_w_match = re.search(r'(\d+)w', prod_name)
                if prod_w_match:
                    pw = int(prod_w_match.group(1))
                    if pw == 60:
                        score += 1.5  # 60w相当は最も汎用的

        # --- 20. ★Fix14: 吊り下げ/両笠 → 反射笠付形優先 ---
        # 吊り下げ式トラフ蛍光灯には反射笠付形が適切
        # 反射笠付形は光の指向性を確保し、吊り下げ環境での照度を維持
        if "吊り下げ" in ft_norm or "両笠" in ft_norm:
            if "反射笠付形" in prod_name or "反射笠付" in prod_name:
                score += 5.0

        # --- 21. ★Fix26: 誘導灯の取付方式・面数マッチング ---
        # 壁面/天井/片面/両面を器具名から解析し、適合製品にボーナス
        if "誘導灯" in ft_norm:
            if "壁面" in ft_norm:
                if "壁面" in prod_name and "天井" not in prod_name:
                    score += 5.0  # 壁面専用
                elif "壁面" in prod_name:
                    score += 2.0  # 天井・壁面兼用
            elif "天井" in ft_norm:
                if "天井" in prod_name and "壁面" not in prod_name:
                    score += 5.0  # 天井専用
                elif "天井" in prod_name:
                    score += 2.0  # 天井・壁面兼用
            if "両面" in ft_norm:
                if "両面" in prod_name:
                    score += 5.0
                elif "片面" in prod_name:
                    score -= 3.0
            elif "片面" in ft_norm:
                if "片面" in prod_name:
                    score += 2.0
            # ★Fix43: 誘導灯の防雨/メーカー選定
            # 屋内誘導灯 → 非防雨を優先、東芝が標準
            # ★Fix72: 防雨ペナルティ強化 -3.0→-5.0
            if not cls.wp_hard_filter:
                if product.is_waterproof:
                    score -= 5.0
                if "防雨" in prod_name or "防湿" in prod_name:
                    score -= 5.0
            # 誘導灯は東芝が正解データで最多
            prod_mfr = _normalize(product.manufacturer or "")
            if "東芝" in prod_mfr or "TOSHIBA" in prod_mfr:
                score += 3.0

        # --- 22. ★Fix27: ユニバーサルダウンライト → ユニバーサル製品優先 ---
        if "ユニバーサル" in ft_norm:
            if "ユニバーサル" in prod_name or "ﾕﾆﾊﾞｰｻﾙ" in prod_name:
                score += 5.0

        # --- 23. ★Fix28: スポットライト Φ→タイプ選定（場所依存） ---
        # 階段のスポット → ビームタイプ（高出力で広範囲照射）
        # 通路/壁面のスポット → 白熱タイプ（小型で壁面用）
        if "スポット" in ft_norm and "センサー" not in ft_norm:
            is_stairway = "階段" in cls.location
            if is_stairway:
                if "ビーム" in prod_name or "ﾋﾞｰﾑ" in prod_name:
                    score += 4.0
                elif "白熱" in prod_name:
                    score -= 2.0
                if "75" in prod_name:
                    score += 2.0
            else:
                # 通路等 → 白熱タイプ優先、60Wが適切
                if "白熱" in prod_name:
                    score += 2.0
                if "60" in prod_name:
                    score += 2.0
                elif "40" in prod_name:
                    score -= 0.5

        # --- 24. ★Fix29: 投光器（小型/階段用）→ スポットライト優先 ---
        if "投光器" in ft_norm:
            if "階段" in cls.location:
                # 階段の投光器は小型ビーム球スポットが適切
                if "ビーム" in prod_name or "ﾋﾞｰﾑ" in prod_name:
                    score += 5.0
                elif "HID" in prod_name:
                    score -= 3.0

        # --- 25a. ★Fix69: ポーチ縦長器具 → 長方形商品優先 ---
        # 510×200等の縦長器具は長方形商品がフィット、□形商品は不適
        if cls.lineup_sheet == "ﾎﾟｰﾁ・支柱":
            if (cls.type_dims and cls.type_dims.length_mm
                    and cls.type_dims.width_mm
                    and cls.type_dims.length_mm > cls.type_dims.width_mm * 2):
                # 長さが幅の2倍以上 → 明らかに縦長
                if "長方形" in prod_name:
                    score += 4.0
                elif re.search(r'[□■]\d+', prod_name):
                    score -= 2.0  # □型は縦長器具に不適

        # --- 25a2. ★Fix73: ポーチ ほぼ正方形器具(玄関灯) → □型商品優先 ---
        # 130×100玄関灯等のアスペクト比1.5以内 → □型商品がフィット
        # ※壁面ブラケットは除外（配管干渉で直管型が適切なケースがある）
        if (cls.lineup_sheet == "ﾎﾟｰﾁ・支柱"
                and "壁面ブラケット" not in ft_norm
                and "壁面ﾌﾞﾗｹｯﾄ" not in ft_norm):
            if (cls.type_dims and cls.type_dims.length_mm
                    and cls.type_dims.width_mm
                    and cls.type_dims.length_mm <= cls.type_dims.width_mm * 1.5):
                # ほぼ正方形
                if re.search(r'[□■]\d+', prod_name):
                    sq_pd = re.search(r'[□■](\d+)', prod_name)
                    if sq_pd:
                        pd_sq = float(sq_pd.group(1))
                        max_dim = max(cls.type_dims.length_mm, cls.type_dims.width_mm)
                        diff = abs(max_dim - pd_sq)
                        if diff < 20:
                            score += 5.0  # ほぼ同サイズ
                        elif diff < 50:
                            score += 3.0

        # --- 25. ★Fix30: ポーチ□寸法マッチング ---
        # □110→□120, □120→□120 等、正方形寸法の近接マッチ
        if cls.lineup_sheet == "ﾎﾟｰﾁ・支柱":
            sq_match = re.search(r'[□■](\d+)', ft_norm)
            if sq_match:
                fixture_sq = float(sq_match.group(1))
                prod_sq_match = re.search(r'[□■](\d+)', prod_name)
                if prod_sq_match:
                    prod_sq = float(prod_sq_match.group(1))
                    diff_sq = abs(fixture_sq - prod_sq)
                    if diff_sq < 15:
                        score += 5.0  # ほぼ同サイズ
                    elif diff_sq < 40:
                        score += 2.0  # 近いサイズ
                    elif diff_sq > 60:
                        score -= 1.0  # 大きすぎ
                    # ★Fix75: カバレッジボーナス
                    # 製品□≧器具□ → 既存取付跡をカバー
                    if prod_sq >= fixture_sq and diff_sq <= 50:
                        score += 2.0
            # ★Fix30b: ポーチ灯Φ→ランタン型/丸型マッチ
            if "ポーチ灯" in ft_norm or "ﾎﾟｰﾁ灯" in ft_norm:
                porch_phi = cls.type_diameter
                if porch_phi > 0 and porch_phi < 160:
                    if "ランタン" in prod_name or "ﾗﾝﾀﾝ" in prod_name:
                        score += 3.0

        # --- 26. ★Fix31: 庭園灯 → 直管LEDバイパス優先 ---
        # ★注: 正解データではTランプ(E26)だが、ラインナップにTランプ商品なし
        # 現状のラインナップで最適なのは直管LEDランプ（バイパス交換）
        # コーンランプは庭園灯には大きすぎるためペナルティ
        if "庭園灯" in ft_norm or "庭園" in ft_norm:
            if "直管" in prod_name and "LED" in prod_name:
                score += 3.0  # 直管LEDバイパスが庭園灯の最適解
            elif ("バイパス" in prod_name or "ﾊﾞｲﾊﾟｽ" in prod_name):
                score += 2.0
            # コーンランプは庭園灯には大型すぎ
            if "コーンランプ" in prod_name or "ｺｰﾝﾗﾝﾌﾟ" in prod_name:
                score -= 3.0
            if "ロングポール" in prod_name or "ﾛﾝｸﾞﾎﾟｰﾙ" in prod_name:
                score -= 2.0
            if "ポール灯" in prod_name:
                score -= 2.0

        # --- 27. ★Fix32: 階段灯サイズ→ワット/幅判定 ---
        # 120×640 → 20形 (長さ<750mm), 660×160 → 20形W127
        # ★Fix58: fixture_sizeからもサイズ取得（type_dimsがない場合）
        stair_dims = cls.type_dims
        if "階段灯" in ft_norm and not stair_dims:
            fs_str = _normalize(fixture.fixture_size or "")
            if fs_str:
                stair_dims = self._parse_cached(fs_str)
        if "階段灯" in ft_norm and stair_dims:
            stair_length = stair_dims.length_mm or 0
            stair_width = stair_dims.width_mm or 0
            if stair_length > 0 and stair_length < 750:
                if "20形" in prod_name:
                    score += 3.0
                elif "40形" in prod_name:
                    score -= 3.0
            # 幅<140mm → カバーなしW127が適切
            if stair_width > 0 and stair_width < 140:
                if "W127" in prod_name:
                    score += 3.0
                elif "カバー付" in prod_name or "ｶﾊﾞｰ付" in prod_name:
                    score -= 1.0

        # --- 28. ★Fix33: ダウンライト□→スクエアダウンライト ---
        if cls.is_recessed:
            sq_dl = re.search(r'[□■](\d+)', ft_norm)
            if sq_dl:
                # □形ダウンライトにはスクエアダウンライトが適切
                if "スクエア" in prod_name or "ｽｸｴｱ" in prod_name:
                    score += 5.0
                elif "ユニバーサル" in prod_name or "ﾕﾆﾊﾞｰｻﾙ" in prod_name:
                    score -= 2.0

        # --- 29. ★Fix34: 筒形ブラケット → 筒形製品優先 ---
        if "筒形" in ft_norm or "筒型" in ft_norm:
            if "筒形" in prod_name or "筒型" in prod_name:
                score += 5.0

        # --- 30. ★Fix35: 格子ブラケット → ポーチ系（楕円/長方形） ---
        if "格子" in ft_norm and "ブラケット" in ft_norm:
            if "楕円" in prod_name:
                score += 3.0
            elif "長方形" in prod_name:
                score += 2.0

        # --- 31. ★Fix47b: 玄関灯(天井壁面ルーティング時)→ブラケット優先 ---
        # Fix47で玄関灯(廊下/Φ)を天井壁面にルーティングした場合、
        # ブラケット製品を優先（シーリングダウンライトは不適切）
        if ("玄関灯" in ft_norm
                and cls.lineup_sheet == "天井・壁面"):
            if "ブラケット" in prod_name or "ﾌﾞﾗｹｯﾄ" in prod_name:
                score += 5.0
            elif "シーリング" in prod_name or "ｼｰﾘﾝｸﾞ" in prod_name:
                score -= 3.0
            elif "ダウンライト" in prod_name or "ﾀﾞｳﾝﾗｲﾄ" in prod_name:
                score -= 3.0

        # --- 32. ★Fix50/Fix66: 誘導灯 壁面専用/天井専用の判別 ---
        # ★Fix66改: ラインナップの壁面専用/天井専用は防雨版のみ。
        # 非防雨の「天井・壁面」兼用の方が適切なケースが多い。
        # → 天井・壁面をペナルティせず、専用マッチにはボーナスのみ付ける
        if "誘導灯" in ft_norm:
            if "壁面" in ft_norm and "天井" not in ft_norm:
                # 壁面のみ → 壁面専用（非防雨）にボーナス
                if ("壁面" in prod_name and "天井" not in prod_name
                        and "防雨" not in prod_name and "防湿" not in prod_name):
                    score += 3.0
            elif "天井" in ft_norm and "壁面" not in ft_norm:
                # 天井のみ → 天井専用（非防雨）にボーナス
                if ("天井" in prod_name and "壁面" not in prod_name
                        and "防雨" not in prod_name and "防湿" not in prod_name):
                    score += 3.0

        # --- 33. ★Fix51: 壁面ブラケット(通路/階段)→ 小径Φ≤150 ランタン型優先 ---
        # fixture_sizeに"120φ"等のΦ情報がある壁面ブラケット→ランタン型
        if ("壁面ブラケット" in ft_norm
                and cls.lineup_sheet == "ﾎﾟｰﾁ・支柱"):
            fs_text = _normalize(fixture.fixture_size or "")
            m_phi = re.search(r'(\d+)\s*[Φφ]', fs_text)
            if not m_phi:
                m_phi = re.search(r'[Φφ]\s*(\d+)', fs_text)
            if m_phi:
                fs_phi = float(m_phi.group(1))
                if fs_phi <= 150:
                    if "ランタン" in prod_name or "ﾗﾝﾀﾝ" in prod_name:
                        score += 6.0  # ★強化: 4.0→6.0
                    # 長方形は小型丸形器具に不適切
                    if "長方形" in prod_name:
                        score -= 2.0

        # --- 34. ★Fix52: 非防水器具 → 商品名「防雨」「防湿」ペナルティ ---
        # 器具が非防水かつ場所も非防水 → 防雨/防湿製品は不適切
        # 例: 非常灯兼用直付蛍光灯(屋内通路) → 非防雨ベースライトが正解
        if not cls.is_waterproof and not cls.wp_hard_filter:
            if "防雨" in prod_name or "防湿" in prod_name:
                score -= 2.0

        # --- 35. ★Fix53: 誘導灯 → C級/BL級判別、音声点滅ペナルティ ---
        # C級器具に音声点滅BL級を選択するのは不適切
        # 音声点滅は特殊仕様で高コスト
        # ★Fix67: fixture_sizeにも「C級」が入るケースに対応
        if "誘導灯" in ft_norm:
            fs_norm = _normalize(fixture.fixture_size or "")
            is_c_class = ("C級" in ft_norm or "c級" in ft_norm
                          or "C級" in fs_norm or "c級" in fs_norm)
            if is_c_class:
                if "BL級" in prod_name or "Bl級" in prod_name:
                    score -= 6.0  # ★Fix67: -4.0→-6.0 強化
                if "音声点滅" in prod_name:
                    score -= 6.0  # ★Fix67: -4.0→-6.0 強化
            # ★Fix53b: リニューアルプレート対応
            # 正解に「リニューアルプレート」が含まれる場合、
            # 既設器具(FBK品番等)からの置換にはリニューアル対応製品を優先
            if "リニューアル" in prod_name or "ﾘﾆｭｰｱﾙ" in prod_name:
                score += 1.0  # リニューアル対応を軽く優先

        # --- 36. ★Fix54b: 壁面ブラケット(階段→天井壁面)→ブラケット優先 ---
        # Fix54で壁面ブラケット(階段)を天井壁面にルーティングした場合
        if ("壁面ブラケット" in ft_norm
                and cls.lineup_sheet == "天井・壁面"
                and "階段" in cls.location):
            if "ブラケット" in prod_name or "ﾌﾞﾗｹｯﾄ" in prod_name:
                score += 5.0
            elif "シーリング" in prod_name or "ｼｰﾘﾝｸﾞ" in prod_name:
                score -= 3.0
            elif "ダウンライト" in prod_name or "ﾀﾞｳﾝﾗｲﾄ" in prod_name:
                score -= 3.0

        # --- 37. ★Fix55b: ダウンライトΦ450+ → ラウンドベースライト ---
        # 丸・四角(大)シートにルーティングされた大型ダウンライト
        if (cls.is_recessed
                and cls.lineup_sheet == "丸・四角(大)"
                and cls.type_diameter >= 400):
            if "ラウンドベースライト" in prod_name or "ﾗｳﾝﾄﾞﾍﾞｰｽﾗｲﾄ" in prod_name:
                score += 8.0
            if "埋込" in prod_name:
                score += 2.0

        # --- 38. ★Fix56: 埋込スクエアライト → □寸法マッチング ---
        # 丸・四角(大)シートでのスクエアライト選定
        if ("スクエアライト" in ft_norm
                and cls.lineup_sheet == "丸・四角(大)"):
            if "スクエアライト" in prod_name or "ｽｸｴｱﾗｲﾄ" in prod_name:
                score += 8.0
            # □寸法の近接マッチ（跡隠し重視: 大きめ > 完全一致）
            sq_ft = re.search(r'[□■](\d+)', ft_norm)
            sq_pd = re.search(r'[□■](\d+)', prod_name)
            if sq_ft and sq_pd:
                ft_sq = float(sq_ft.group(1))
                pd_sq = float(sq_pd.group(1))
                diff_sq = abs(ft_sq - pd_sq)
                if pd_sq >= ft_sq:
                    # 大きめ: 跡隠し可能
                    coverage = pd_sq - ft_sq
                    if coverage <= 100:
                        score += 5.0  # 適度なカバレッジ
                    else:
                        score += 2.0  # やや大きい
                    score += 2.0  # カバレッジボーナス
                elif diff_sq < 30:
                    score += 4.0  # ほぼ同サイズ
                elif diff_sq < 100:
                    score += 1.0  # 近いサイズだが小さい
            # 埋込指定 → 「直付・埋込」製品を優先
            if "埋込" in ft_norm:
                if "埋込" in prod_name:
                    score += 3.0
                if "直付" in prod_name and "埋込" not in prod_name:
                    score -= 2.0  # 直付のみは埋込には不適切

        # --- 39. ★Fix57: 投光器(階段)→ 投光器製品ペナルティ強化 ---
        # Fix29でビームスポットにボーナスを付けているが、
        # 投光器製品自体にもペナルティが必要
        if "投光器" in ft_norm and "階段" in cls.location:
            if "投光器" in prod_name:
                score -= 5.0
            elif "スポット" in prod_name or "ｽﾎﾟｯﾄ" in prod_name:
                score += 3.0

        # --- 40. ★Fix70: 非常灯兼用ブラケット → 丸形ブラケット製品を優先 ---
        # 「260×175非常灯兼用ブラケット」等 → 「非常用丸形ブラケット〈Φ380〉リニューアル」
        # 階段灯ではなくブラケット型の非常用製品が適切
        if (cls.has_emergency and "ブラケット" in ft_norm
                and cls.lineup_sheet == "その他非常灯"):
            if "丸形ﾌﾞﾗｹｯﾄ" in prod_name or "丸形ブラケット" in prod_name:
                score += 5.0
            elif "ブラケット" in prod_name or "ﾌﾞﾗｹｯﾄ" in prod_name:
                score += 3.0
            if "階段灯" in prod_name:
                score -= 3.0

        # --- 41. ★Fix71: 非常灯(エントランス/玄関) → 丸形ブラケット優先 ---
        # 「非常灯」(エントランス)で形状指定なし→丸形ブラケットが汎用
        # 蓄電池やベースライトは不適切
        if (cls.has_emergency
                and cls.lineup_sheet == "その他非常灯"
                and "ブラケット" not in ft_norm
                and "階段" not in ft_norm
                and "埋込" not in ft_norm
                and "直付" not in ft_norm):
            loc_lower = _normalize(fixture.location or "")
            if ("エントランス" in loc_lower or "玄関" in loc_lower
                    or "ホール" in loc_lower):
                if "丸形ﾌﾞﾗｹｯﾄ" in prod_name or "丸形ブラケット" in prod_name:
                    score += 4.0
                if "蓄電池" in prod_name:
                    score -= 5.0

        return score

    @staticmethod
    def _extract_product_diameter(prod_name: str) -> float:
        """LED商品名から直径(Φ)を抽出

        例: "天井・壁面ﾌﾞﾗｹｯﾄ〈Φ310-B/N/K/FCL30〉" → 310.0
            "非常用丸形ﾌﾞﾗｹｯﾄ〈Φ380-N〉ﾘﾆｭｰｱﾙ" → 380.0
        """
        m = re.search(r'[Φφ]\s*(\d+)', prod_name)
        if m:
            return float(m.group(1))
        return 0.0

    # ===== ユーティリティ =====

    def _parse_cached(self, size_str: str) -> FixtureDimensions:
        """サイズパース結果をキャッシュ"""
        if size_str not in self._size_cache:
            self._size_cache[size_str] = parse_fixture_size(size_str)
        return self._size_cache[size_str]

    def _calc_confidence(
        self,
        fixture: ExistingFixture,
        product: LEDProduct,
        cls: FixtureClassification,
        needs_size_review: bool,
    ) -> float:
        """マッチング信頼度を計算 (0.0-1.0)"""
        score = 0.5  # ベース

        if not needs_size_review:
            score += 0.2

        if not cls.fallback:
            score += 0.15

        if product.purchase_price_total and product.purchase_price_total > 0:
            score += 0.1

        if product.lumens:
            score += 0.05

        return min(score, 1.0)

    def _build_notes(
        self,
        fixture: ExistingFixture,
        product: LEDProduct,
        cls: FixtureClassification,
        existing_dims: FixtureDimensions,
        needs_size_review: bool,
    ) -> str:
        """選定メモを構築"""
        parts = []

        parts.append(f"シート:{cls.lineup_sheet}")

        if needs_size_review:
            parts.append("サイズ要確認")

        if cls.fallback:
            parts.append("フォールバック分類")

        if cls.has_emergency:
            parts.append("非常灯")

        if cls.is_waterproof:
            parts.append("防水")

        if product.manufacturer:
            parts.append(product.manufacturer)

        return " / ".join(parts)

    def _estimate_construction_price(
        self,
        fixture: ExistingFixture,
        product: LEDProduct,
        cls: FixtureClassification,
    ) -> int:
        """工事単価を推定"""
        prices = self.config.construction_prices

        # 交換方法から基本単価を決定
        method = product.replacement_method or ""
        ft_norm = _normalize(fixture.fixture_type or "")
        is_downlight = any(
            kw in ft_norm for kw in ("ダウンライト", "DL")
        )
        if "ランプ" in method or "球" in method:
            base = prices.get("ランプ交換", 1500)
        elif "バイパス" in method:
            base = prices.get("バイパス工事", 5000)
        elif cls.is_recessed and not is_downlight:
            # 蛍光灯埋込タイプは5000円（全ネジボルト作業等）
            # ※ダウンライトは埋込だが器具交換扱い(3000円)
            base = prices.get("埋込器具交換", 5000)
        else:
            base = prices.get("器具交換", 3000)

        # 加算
        if cls.is_waterproof:
            base += prices.get("waterproof_addon", 500)

        if cls.has_emergency:
            base += prices.get("emergency_addon", 1000)

        # 高所作業の判定（工事備考から）
        notes = fixture.construction_notes or ""
        if "高所" in notes or "足場" in notes:
            base += prices.get("high_location_addon", 2000)

        return base


def _normalize(text: str) -> str:
    """テキストを正規化"""
    return unicodedata.normalize('NFKC', text).strip()
