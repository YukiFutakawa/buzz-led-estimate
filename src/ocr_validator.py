"""OCR結果の検証・自動修正モジュール

AI OCRの出力JSONを検証し、以下の修正・警告を行う:
- 電球種別のファジーマッチ（誤読修正）
- 数値範囲チェック（異常値の警告）
- 幽霊行の除去（空データ行の削除）
- 行ラベル連続性チェック
- 色温度バリデーション
"""

from __future__ import annotations

import copy
import logging
import re
import unicodedata
from pathlib import Path
from typing import Optional

import yaml

logger = logging.getLogger(__name__)


# 既知の電球種別リスト（category_mapping.yaml + 追加）
KNOWN_BULB_TYPES = [
    "FL10", "FL15", "FL20", "FL30", "FL40",
    "FLR40", "FHF16", "FHF32",
    "FCL20", "FCL30", "FCL32", "FCL40",
    "FDL13", "FDL27",
    "FHT", "FPL", "FML",
    "LED",
]

# 既知の電球種別（ワット付き表記も含む）
KNOWN_BULB_PATTERNS = [
    r"FL\d+",
    r"FLR\d+",
    r"FHF\d+",
    r"FCL\d+",
    r"FDL\d+",
    r"FHT\d*",
    r"FPL\d*",
    r"FML\d*",
    r"白熱\d+W",
    r"ミニクリプトン\d*W?",
    r"ハロゲン\d*W?",
    r"LED",
]

# 有効な色温度値
VALID_COLOR_TEMPS = {"白", "黄", "N", "L", "昼白色", "電球色", "昼白", "電球", "温白色", "WW", ""}

# 行ラベルの正しい順序
ROW_LABEL_ORDER = list("ABCDEFGHIJKLMNOPQRST")


def _normalize_text(text: str) -> str:
    """全角→半角、空白除去"""
    return unicodedata.normalize("NFKC", text).strip()


def _levenshtein_distance(s1: str, s2: str) -> int:
    """レーベンシュタイン距離（編集距離）を計算"""
    if len(s1) < len(s2):
        return _levenshtein_distance(s2, s1)
    if len(s2) == 0:
        return len(s1)

    prev_row = range(len(s2) + 1)
    for i, c1 in enumerate(s1):
        curr_row = [i + 1]
        for j, c2 in enumerate(s2):
            insertions = prev_row[j + 1] + 1
            deletions = curr_row[j] + 1
            substitutions = prev_row[j] + (c1 != c2)
            curr_row.append(min(insertions, deletions, substitutions))
        prev_row = curr_row

    return prev_row[-1]


class OCRValidator:
    """OCR結果の検証・自動修正エンジン"""

    def __init__(self, config_path: Optional[Path] = None):
        self._known_bulb_types = list(KNOWN_BULB_TYPES)

        # category_mapping.yaml から追加の語彙を読み込み
        if config_path is None:
            config_path = (
                Path(__file__).parent.parent / "config" / "category_mapping.yaml"
            )
        self._load_config_vocabularies(config_path)

    def _load_config_vocabularies(self, config_path: Path) -> None:
        """category_mapping.yaml から既知語彙を読み込む"""
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                config = yaml.safe_load(f)

            # bulb_type_to_watt_form のキーを既知電球種別に追加
            bulb_map = config.get("bulb_type_to_watt_form", {})
            for key in bulb_map:
                if key not in self._known_bulb_types:
                    self._known_bulb_types.append(key)

        except Exception as e:
            logger.warning(f"設定ファイル読み込みスキップ: {e}")

    def validate_and_fix(self, ocr_result: dict) -> dict:
        """OCR結果を検証し、修正可能なエラーを修正

        Args:
            ocr_result: document_processor.ocr_survey_sheet() の出力

        Returns:
            検証・修正済みのOCR結果（_validation_warnings キー付き）
        """
        result = copy.deepcopy(ocr_result)
        warnings: list[str] = []

        fixtures = result.get("fixtures", [])

        # 1. 幽霊行の除去
        fixtures, ghost_warnings = self._remove_ghost_rows(fixtures)
        warnings.extend(ghost_warnings)

        # 2. 各行の検証・修正
        for fixture in fixtures:
            row_warnings = self._validate_fixture(fixture)
            warnings.extend(row_warnings)

        # 3. 行ラベルの連続性チェック
        label_warnings = self._check_row_label_continuity(fixtures)
        warnings.extend(label_warnings)

        result["fixtures"] = fixtures
        result["_validation_warnings"] = warnings

        if warnings:
            logger.info(f"OCRバリデーション: {len(warnings)}件の警告")
            for w in warnings:
                logger.debug(f"  - {w}")

        return result

    def _remove_ghost_rows(
        self, fixtures: list[dict],
    ) -> tuple[list[dict], list[str]]:
        """幽霊行（実質空データ）を除去"""
        valid = []
        warnings = []

        for fixture in fixtures:
            location = (fixture.get("location") or "").strip()
            ftype = (fixture.get("fixture_type") or "").strip()
            btype = (fixture.get("bulb_type") or "").strip()

            # 数量の合計
            floor_q = fixture.get("floor_quantities", {})
            total_qty = sum(
                int(v) for v in floor_q.values()
                if isinstance(v, (int, float)) and v > 0
            )

            # 全フィールドが空で数量も0なら幽霊行
            if not location and not ftype and not btype and total_qty == 0:
                label = fixture.get("row_label", "?")
                warnings.append(f"行{label}: 空データ行を除去")
                continue

            valid.append(fixture)

        return valid, warnings

    def _validate_fixture(self, fixture: dict) -> list[str]:
        """1行分の器具データを検証・修正"""
        warnings = []
        label = fixture.get("row_label", "?")

        # 電球種別のファジーマッチ
        bulb_type = fixture.get("bulb_type", "")
        if bulb_type:
            fixed_bulb = self._fix_bulb_type(bulb_type)
            if fixed_bulb != bulb_type:
                warnings.append(
                    f"行{label}: 電球種別修正 '{bulb_type}' → '{fixed_bulb}'"
                )
                fixture["bulb_type"] = fixed_bulb

        # 消費電力の範囲チェック
        power_w = fixture.get("power_w", 0)
        if power_w is not None and power_w != 0:
            try:
                pw = float(power_w)
                if pw < 0 or pw > 500:
                    warnings.append(
                        f"行{label}: 消費電力が範囲外 ({pw}W, 通常0-500W)"
                    )
            except (ValueError, TypeError):
                warnings.append(
                    f"行{label}: 消費電力が数値でない ({power_w})"
                )

        # 点灯時間の範囲チェック
        daily_hours = fixture.get("daily_hours", 0)
        if daily_hours is not None and daily_hours != 0:
            try:
                dh = float(daily_hours)
                if dh < 0 or dh > 24:
                    warnings.append(
                        f"行{label}: 点灯時間が範囲外 ({dh}h, 通常0-24h)"
                    )
            except (ValueError, TypeError):
                warnings.append(
                    f"行{label}: 点灯時間が数値でない ({daily_hours})"
                )

        # 階別数量の範囲チェック
        floor_q = fixture.get("floor_quantities", {})
        for floor_key, qty in floor_q.items():
            try:
                q = int(qty) if qty else 0
                if q < 0 or q > 99:
                    warnings.append(
                        f"行{label}: {floor_key}の数量が範囲外 ({q}, 通常0-99)"
                    )
            except (ValueError, TypeError):
                warnings.append(
                    f"行{label}: {floor_key}の数量が数値でない ({qty})"
                )

        # 色温度バリデーション
        color_temp = fixture.get("color_temp", "")
        if color_temp and color_temp not in VALID_COLOR_TEMPS:
            # 近いものに修正を試みる
            fixed_color = self._fix_color_temp(color_temp)
            if fixed_color != color_temp:
                warnings.append(
                    f"行{label}: 色温度修正 '{color_temp}' → '{fixed_color}'"
                )
                fixture["color_temp"] = fixed_color
            else:
                warnings.append(
                    f"行{label}: 不明な色温度 '{color_temp}'"
                )

        return warnings

    def _fix_bulb_type(self, raw: str) -> str:
        """電球種別のファジーマッチ修正

        よくある誤読パターン:
        - "FL2O" (ゼロ→オー) → "FL20"
        - "FDL 13" (余計空白) → "FDL13"
        - "FL2O" → "FL20"
        """
        # 正規化: 全角→半角、空白除去
        normalized = _normalize_text(raw)
        normalized = normalized.replace(" ", "")

        # よくある文字→数字の誤読修正
        # "O" → "0", "l" → "1", "I" → "1" (電球種別コードの後ろの数字部分)
        fixed = self._fix_alpha_digit_confusion(normalized)

        # 既知パターンにマッチするか確認
        for pattern in KNOWN_BULB_PATTERNS:
            if re.fullmatch(pattern, fixed, re.IGNORECASE):
                return fixed

        # ファジーマッチ: 既知の電球種別との編集距離が小さいものを探す
        best_match = None
        best_distance = 999

        for known in self._known_bulb_types:
            dist = _levenshtein_distance(fixed.upper(), known.upper())
            if dist < best_distance and dist <= 2:  # 編集距離2以内
                best_distance = dist
                best_match = known

        if best_match and best_distance > 0:
            return best_match

        return fixed

    def _fix_alpha_digit_confusion(self, text: str) -> str:
        """文字と数字の混同を修正

        電球種別は「アルファベット接頭辞 + 数字」の形式。
        数字部分に紛れ込んだ文字を修正する。
        """
        # パターン: FL, FDL, FCL, FHF, FLR, FHT, FPL, FML + 数字
        m = re.match(r"^(FL|FDL|FCL|FHF|FLR|FHT|FPL|FML|FMT)(.*)$", text, re.IGNORECASE)
        if not m:
            return text

        prefix = m.group(1).upper()
        suffix = m.group(2)

        # suffix内の文字→数字変換
        digit_fixes = {"O": "0", "o": "0", "l": "1", "I": "1", "S": "5", "B": "8"}
        fixed_suffix = ""
        for ch in suffix:
            if ch in digit_fixes:
                fixed_suffix += digit_fixes[ch]
            else:
                fixed_suffix += ch

        return prefix + fixed_suffix

    def _fix_color_temp(self, raw: str) -> str:
        """色温度の修正"""
        normalized = _normalize_text(raw)

        if "白" in normalized:
            return "白"
        if "黄" in normalized:
            return "黄"
        if normalized.upper() in ("N", "W"):
            return "白"
        if normalized.upper() in ("L",):
            return "黄"

        return raw

    def _check_row_label_continuity(
        self, fixtures: list[dict],
    ) -> list[str]:
        """行ラベルの連続性チェック"""
        warnings = []
        labels = [f.get("row_label", "") for f in fixtures]

        if not labels:
            return warnings

        # 重複チェック
        seen = set()
        for label in labels:
            if label in seen:
                warnings.append(f"行ラベル重複: '{label}'")
            seen.add(label)

        # 連続性チェック（A, B, C, ... の順か）
        expected_idx = 0
        for label in labels:
            if label not in ROW_LABEL_ORDER:
                continue

            actual_idx = ROW_LABEL_ORDER.index(label)
            if actual_idx > expected_idx + 1:
                # ギャップがある
                missing = ROW_LABEL_ORDER[expected_idx:actual_idx]
                # 最初の行がAでない場合は警告しない
                if expected_idx > 0:
                    warnings.append(
                        f"行ラベルギャップ: {missing} が欠落している可能性"
                    )
            expected_idx = actual_idx + 1

        return warnings
