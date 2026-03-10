# -*- coding: utf-8 -*-
"""正しい見積りパターン分析 & 現行マッチャー比較レポート

正しい見積りExcelからパターンを抽出し、
現行LEDマッチャーの選定結果と比較するレポートを生成する。

Usage:
    python analyze_patterns.py
"""

from __future__ import annotations

import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

import logging
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)

from correct_estimate_importer import (
    CorrectEstimateImporter,
    CorrectEstimate,
    CorrectFixtureMapping,
)
from models import ExistingFixture, FloorQuantities
from lineup_loader import LineupIndex
from led_matcher import LEDMatcher


# ============================================================
# 分析用データクラス
# ============================================================

@dataclass
class FixtureLEDPattern:
    """器具種別→LED選定パターン（集計単位）"""
    fixture_type: str           # 照明種別
    size_memo: str              # 現調備考/サイズ
    construction_memo: str      # 工事備考/電球種別
    led_selection: str          # 正しいLED選定名
    location: str               # 設置場所
    quantity: int               # 数量
    count: int = 1              # この組合せの出現回数
    source_files: list[str] = field(default_factory=list)


@dataclass
class ComparisonResult:
    """現行マッチャーと正解の比較結果"""
    fixture_type: str           # 照明種別
    size_memo: str              # サイズ情報
    construction_memo: str      # 電球種別
    location: str               # 設置場所
    correct_led: str            # ★正解LED選定
    our_led: str                # 現行の選定結果
    our_sheet: str              # 現行の分類先シート
    our_watt_form: str          # 現行のwatt form
    our_emergency: bool         # 現行の非常灯判定
    our_waterproof: bool        # 現行の防滴判定
    our_affinity: float         # 現行のaffinity
    match: bool                 # 一致したかどうか
    source_file: str            # ソースファイル名
    near_match: bool = False    # ニアマッチ（色/メーカーのみ差異）
    near_match_diff: str = ""   # ニアマッチの差異内容


# ============================================================
# 分析関数
# ============================================================

def _normalize(text: str) -> str:
    """NFKC正規化"""
    if not text:
        return ""
    return unicodedata.normalize('NFKC', str(text).strip())


def _check_near_match(correct: str, ours: str) -> tuple[bool, str]:
    """ニアマッチ判定: 色/メーカーのみの差異かチェック

    Returns:
        (is_near_match, diff_description)
    """
    c = _normalize(correct)
    o = _normalize(ours)
    if c == o:
        return False, ""  # 完全一致はニアマッチではない
    if not c or not o:
        return False, ""

    diffs = []

    # 製品カテゴリ（〈〉の前）が同じかチェック
    import re
    c_prefix = re.split(r'[〈<＜]', c)[0]
    o_prefix = re.split(r'[〈<＜]', o)[0]
    if c_prefix != o_prefix:
        return False, ""  # カテゴリが異なる → 非ニアマッチ

    # 〈〉内のパラメータを抽出
    c_params = re.search(r'[〈<＜](.+?)[〉>＞]', c)
    o_params = re.search(r'[〈<＜](.+?)[〉>＞]', o)
    if not c_params or not o_params:
        return False, ""

    c_parts = set(c_params.group(1).replace('-', '/').split('/'))
    o_parts = set(o_params.group(1).replace('-', '/').split('/'))

    # 差異を特定
    only_in_correct = c_parts - o_parts
    only_in_ours = o_parts - c_parts

    if not only_in_correct and not only_in_ours:
        # 〈〉の外側が異なる（例: リニューアル有無、3m表記等）
        return True, "接尾辞のみ差異"

    # 色/メーカーパターン
    color_codes = {"K", "W", "DS", "B", "S", "P"}
    lighting_codes = {"N", "L", "WW"}
    mfr_codes = {"M", "T", "P", "O", "K"}  # 三菱, 東芝, パナ, オーデリック, コイズミ

    for item in only_in_correct | only_in_ours:
        item_upper = item.strip().upper()
        # 器具色
        if item_upper in color_codes:
            if item in only_in_correct:
                diffs.append(f"器具色: 正解={item}")
            else:
                diffs.append(f"器具色: 現行={item}")
        # 照明色
        elif item_upper in lighting_codes:
            if item in only_in_correct:
                diffs.append(f"照明色: 正解={item}")
            else:
                diffs.append(f"照明色: 現行={item}")
        # メーカー
        elif item_upper in mfr_codes and len(item.strip()) == 1:
            if item in only_in_correct:
                diffs.append(f"メーカー: 正解={item}")
            else:
                diffs.append(f"メーカー: 現行={item}")
        # 近似Φ（5mm以内）
        elif re.match(r'Φ?\d+', item.strip()):
            # 対応するΦを相手から探す
            partner_set = only_in_ours if item in only_in_correct else only_in_correct
            partner_phi = None
            for p_item in partner_set:
                if re.match(r'Φ?\d+', p_item.strip()):
                    partner_phi = p_item
                    break
            if partner_phi:
                # Φ数値を抽出して差を確認
                c_val = re.search(r'(\d+)', item)
                p_val = re.search(r'(\d+)', partner_phi)
                if c_val and p_val:
                    diff_mm = abs(int(c_val.group(1)) - int(p_val.group(1)))
                    if diff_mm <= 5:
                        diffs.append(f"Φ近似({diff_mm}mm差)")
                    else:
                        return False, ""  # Φ差が大きい → 非ニアマッチ
            else:
                return False, ""
        # ワット形（60w vs FHT42W等）
        elif re.match(r'\d+w$', item.strip().lower()) or re.match(r'FHT\d+', item.strip()):
            diffs.append(f"ワット形: {item}")
        else:
            return False, ""  # 不明な差異 → 非ニアマッチ

    if diffs:
        return True, ", ".join(diffs)
    return False, ""


def tabulate_patterns(estimates: list[CorrectEstimate]) -> list[FixtureLEDPattern]:
    """全正しい見積りからパターンを集計"""
    patterns = []

    for est in estimates:
        for mapping in est.fixture_mappings:
            if not mapping.led_selection:
                continue

            patterns.append(FixtureLEDPattern(
                fixture_type=mapping.fixture_type,
                size_memo=mapping.size_memo,
                construction_memo=mapping.construction_memo,
                led_selection=mapping.led_selection,
                location=mapping.location,
                quantity=mapping.quantity,
                source_files=[est.file_path.name],
            ))

    return patterns


def compare_with_matcher(
    estimates: list[CorrectEstimate],
    lineup_index: LineupIndex,
) -> list[ComparisonResult]:
    """正解と現行マッチャーを比較"""
    matcher = LEDMatcher(lineup_index)
    results = []

    for est in estimates:
        for mapping in est.fixture_mappings:
            if not mapping.led_selection:
                continue

            # ExistingFixtureを再構築
            fixture = ExistingFixture(
                row_label=mapping.row_label,
                location=mapping.location,
                fixture_type=mapping.fixture_type,
                fixture_size=mapping.size_memo,
                bulb_type=mapping.construction_memo,
                quantities=FloorQuantities(floors=mapping.floor_quantities),
                power_consumption_w=mapping.power_w,
            )

            # 現行マッチャーで分類・選定
            try:
                cls = matcher._classify_fixture(fixture)
                match_result = matcher.match_fixture(fixture)

                our_led = ""
                our_affinity = 0.0
                if match_result.led_product:
                    our_led = match_result.led_product.name
                    our_affinity = matcher._successor_affinity(
                        fixture, match_result.led_product, cls
                    )

                is_match = _normalize(mapping.led_selection) == _normalize(our_led)
                near_match = False
                near_match_diff = ""
                if not is_match:
                    near_match, near_match_diff = _check_near_match(
                        mapping.led_selection, our_led
                    )

                results.append(ComparisonResult(
                    fixture_type=mapping.fixture_type,
                    size_memo=mapping.size_memo,
                    construction_memo=mapping.construction_memo,
                    location=mapping.location,
                    correct_led=mapping.led_selection,
                    our_led=our_led,
                    our_sheet=cls.lineup_sheet,
                    our_watt_form=cls.watt_form or "",
                    our_emergency=cls.has_emergency,
                    our_waterproof=cls.is_waterproof,
                    our_affinity=our_affinity,
                    match=is_match,
                    near_match=near_match,
                    near_match_diff=near_match_diff,
                    source_file=est.file_path.name,
                ))
            except Exception as e:
                logger.warning(f"比較エラー [{mapping.fixture_type}]: {e}")
                results.append(ComparisonResult(
                    fixture_type=mapping.fixture_type,
                    size_memo=mapping.size_memo,
                    construction_memo=mapping.construction_memo,
                    location=mapping.location,
                    correct_led=mapping.led_selection,
                    our_led=f"ERROR: {e}",
                    our_sheet="",
                    our_watt_form="",
                    our_emergency=False,
                    our_waterproof=False,
                    our_affinity=0,
                    match=False,
                    source_file=est.file_path.name,
                ))

    return results


def analyze_exclusions(estimates: list[CorrectEstimate]) -> dict:
    """除外パターンの分析"""
    reasons = {}
    details = []

    for est in estimates:
        for excl in est.excluded_fixtures:
            reason = excl.exclusion_reason
            reasons[reason] = reasons.get(reason, 0) + 1
            details.append({
                "property": est.property_name,
                "fixture_type": excl.fixture_type,
                "location": excl.location,
                "quantity": excl.quantity,
                "reason": reason,
            })

    return {
        "reason_frequency": reasons,
        "details": details,
        "total_excluded": len(details),
    }


# ============================================================
# レポート生成
# ============================================================

def generate_report(
    estimates: list[CorrectEstimate],
    patterns: list[FixtureLEDPattern],
    comparisons: list[ComparisonResult],
    exclusion_analysis: dict,
) -> str:
    """日本語レポートを生成"""
    lines = []

    lines.append("=" * 70)
    lines.append("  正しい見積りパターン分析レポート")
    lines.append("=" * 70)

    # --- 概要 ---
    lines.append(f"\n分析対象: {len(estimates)}ファイル")
    for est in estimates:
        lines.append(f"  - {est.file_path.name}")
        lines.append(f"    物件: {est.property_name} ({est.address})")
        lines.append(f"    器具: {len(est.fixture_mappings)}件, "
                     f"除外: {len(est.excluded_fixtures)}件, "
                     f"選定商品: {len(est.product_specs)}件")

    # --- Section 1: パターン一覧 ---
    lines.append(f"\n{'='*70}")
    lines.append("  1. 器具→LED選定パターン一覧")
    lines.append(f"{'='*70}")

    for i, p in enumerate(patterns, 1):
        lines.append(f"\n  [{i}] {p.fixture_type}")
        lines.append(f"      サイズ: {p.size_memo or '(なし)'}")
        lines.append(f"      電球:   {p.construction_memo or '(なし)'}")
        lines.append(f"      場所:   {p.location}")
        lines.append(f"      数量:   {p.quantity}")
        lines.append(f"      → LED: {p.led_selection}")
        lines.append(f"      出典:   {p.source_files[0]}")

    # --- Section 2: 現行マッチャーとの比較 ---
    lines.append(f"\n{'='*70}")
    lines.append("  2. 現行マッチャーとの比較")
    lines.append(f"{'='*70}")

    total = len(comparisons)
    matches = sum(1 for c in comparisons if c.match)
    near_matches = [c for c in comparisons if not c.match and c.near_match]
    true_mismatches = [c for c in comparisons if not c.match and not c.near_match]

    lines.append(f"\n  完全一致: {matches}/{total} ({matches/total*100:.1f}%)")
    lines.append(f"  ニアマッチ（色/メーカーのみ差異）: {len(near_matches)}件")
    lines.append(f"  実質一致率: {matches + len(near_matches)}/{total} "
                 f"({(matches + len(near_matches))/total*100:.1f}%)")
    lines.append(f"  真の不一致: {len(true_mismatches)}件")

    if matches > 0:
        lines.append(f"\n  --- 完全一致 ---")
        for c in comparisons:
            if c.match:
                lines.append(f"  [OK] {c.fixture_type}")
                lines.append(f"       → {c.correct_led}")

    if near_matches:
        lines.append(f"\n  --- ニアマッチ（色/メーカーのみ差異） ---")
        for c in near_matches:
            lines.append(f"\n  [≈] {c.fixture_type}")
            lines.append(f"       場所:   {c.location}")
            lines.append(f"       正解:   {c.correct_led}")
            lines.append(f"       現行:   {c.our_led}")
            lines.append(f"       差異:   {c.near_match_diff}")

    if true_mismatches:
        lines.append(f"\n  --- 真の不一致（要改善） ---")
        for c in true_mismatches:
            lines.append(f"\n  [NG] {c.fixture_type}")
            lines.append(f"       サイズ: {c.size_memo or '-'} | "
                        f"電球: {c.construction_memo or '-'} | "
                        f"場所: {c.location}")
            lines.append(f"       正解:   {c.correct_led}")
            lines.append(f"       現行:   {c.our_led or '(選定なし)'}")
            lines.append(f"       分類:   sheet={c.our_sheet}, "
                        f"watt={c.our_watt_form}, "
                        f"非常灯={c.our_emergency}, "
                        f"防滴={c.our_waterproof}, "
                        f"aff={c.our_affinity:.1f}")

            # 不一致の原因分析
            correct_norm = _normalize(c.correct_led)
            our_norm = _normalize(c.our_led)

            # カテゴリ（ベースライト vs ブラケット等）の違いを検出
            correct_category = _extract_category(correct_norm)
            our_category = _extract_category(our_norm)
            if correct_category != our_category:
                lines.append(f"       原因:   カテゴリ不一致 "
                            f"(正解={correct_category}, 現行={our_category})")

    # --- Section 3: 除外パターン ---
    lines.append(f"\n{'='*70}")
    lines.append("  3. 除外パターン分析")
    lines.append(f"{'='*70}")

    excl = exclusion_analysis
    lines.append(f"\n  除外総数: {excl['total_excluded']}件")

    if excl['reason_frequency']:
        lines.append(f"\n  --- 除外理由の頻度 ---")
        for reason, count in excl['reason_frequency'].items():
            lines.append(f"  [{count}回] {reason}")

        lines.append(f"\n  --- 除外器具の詳細 ---")
        for d in excl['details']:
            lines.append(f"  {d['property']} / {d['location']} / "
                        f"{d['fixture_type']} (×{d['quantity']})")
            lines.append(f"    → {d['reason']}")

    # --- Section 4: 改善提案 ---
    lines.append(f"\n{'='*70}")
    lines.append("  4. 改善提案（不一致パターンから自動生成）")
    lines.append(f"{'='*70}")

    if true_mismatches:
        lines.append(f"\n  以下のルール追加/修正が必要:")
        for i, c in enumerate(true_mismatches, 1):
            lines.append(f"\n  提案{i}: {c.fixture_type} の選定ルール")
            lines.append(f"    入力: type={c.fixture_type}, "
                        f"size={c.size_memo}, bulb={c.construction_memo}")
            lines.append(f"    期待: {c.correct_led}")
            lines.append(f"    現状: sheet={c.our_sheet} → {c.our_led}")

            # 具体的な修正提案
            correct_norm = _normalize(c.correct_led)
            if 'ベースライト' in correct_norm or 'ﾍﾞｰｽﾗｲﾄ' in correct_norm:
                if 'ブラケット' in _normalize(c.our_led) or 'ﾌﾞﾗｹｯﾄ' in _normalize(c.our_led):
                    lines.append(f"    修正案: _classify_fixture()で "
                                f"'{c.fixture_type}' → ベースライト系シートへルーティング")

            if 'トラフ' in correct_norm or 'ﾄﾗﾌ' in correct_norm:
                lines.append(f"    修正案: トラフライト系カテゴリの追加")

            if '階段灯' in correct_norm:
                lines.append(f"    修正案: 階段灯の専用カテゴリルーティング")

            if 'Tランプ' in correct_norm or 'ﾊﾞｲﾊﾟｽ' in correct_norm:
                lines.append(f"    修正案: ランプ交換/バイパスの判定ルール追加")

    if near_matches:
        lines.append(f"\n  --- ニアマッチの改善方向 ---")
        lines.append(f"  ニアマッチ{len(near_matches)}件は器具色/照明色/メーカーの差異のみ。")
        lines.append(f"  これらは物件単位の好み設定で対応可能:")
        lines.append(f"    - 器具色(K/W): 物件の既存器具色を入力項目に追加")
        lines.append(f"    - 照明色(N/L): 物件の好み色温度を入力項目に追加")
        lines.append(f"    - メーカー: 物件の指定メーカーを入力項目に追加")

    if not true_mismatches and not near_matches:
        lines.append(f"\n  全パターンが一致しています。改善不要。")

    return "\n".join(lines)


def _extract_category(product_name: str) -> str:
    """商品名から大カテゴリを抽出"""
    name = _normalize(product_name)
    categories = [
        ("ベースライト", "ベースライト"),
        ("ﾍﾞｰｽﾗｲﾄ", "ベースライト"),
        ("トラフ", "トラフライト"),
        ("ﾄﾗﾌ", "トラフライト"),
        ("ブラケット", "ブラケット"),
        ("ﾌﾞﾗｹｯﾄ", "ブラケット"),
        ("シーリング", "シーリング"),
        ("ｼｰﾘﾝｸﾞ", "シーリング"),
        ("ダウンライト", "ダウンライト"),
        ("ﾀﾞｳﾝﾗｲﾄ", "ダウンライト"),
        ("ポーチ", "ポーチ"),
        ("ﾎﾟｰﾁ", "ポーチ"),
        ("階段灯", "階段灯"),
        ("Tランプ", "Tランプ/バイパス"),
        ("ﾊﾞｲﾊﾟｽ", "Tランプ/バイパス"),
    ]
    for keyword, category in categories:
        if keyword in name:
            return category
    return "その他"


# ============================================================
# メイン実行
# ============================================================

def main():
    base = Path(__file__).parent.parent
    correct_dir = base / '正しい見積り'
    lineup_dir = base / 'ラインナップ表'

    print("=" * 70)
    print("  正しい見積りパターン分析ツール")
    print("=" * 70)

    # Step 1: 正解ファイル読み込み
    print(f"\n[Step 1] 正解ファイル読み込み: {correct_dir}")
    importer = CorrectEstimateImporter()
    estimates = importer.import_folder(correct_dir)

    if not estimates:
        print("  正解ファイルが見つかりません。")
        return

    # Step 2: ラインナップ読み込み
    print(f"\n[Step 2] ラインナップ表読み込み: {lineup_dir}")
    lineup_idx = LineupIndex()
    lineup_idx.load_all(lineup_dir)

    # Step 3: パターン集計
    print(f"\n[Step 3] パターン集計...")
    patterns = tabulate_patterns(estimates)

    # Step 4: 現行マッチャー比較
    print(f"\n[Step 4] 現行マッチャーとの比較...")
    comparisons = compare_with_matcher(estimates, lineup_idx)

    # Step 5: 除外パターン分析
    print(f"\n[Step 5] 除外パターン分析...")
    exclusion_analysis = analyze_exclusions(estimates)

    # Step 6: レポート生成
    print(f"\n[Step 6] レポート生成...")
    report = generate_report(estimates, patterns, comparisons, exclusion_analysis)

    # レポート出力
    print(f"\n{report}")

    # ファイルにも保存
    output_path = base / 'output' / 'pattern_analysis_report.txt'
    output_path.write_text(report, encoding='utf-8')
    print(f"\nレポート保存: {output_path}")


if __name__ == '__main__':
    main()
