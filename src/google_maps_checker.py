"""Googleマップによる物件照明チェッカー

物件住所からGoogleマップ/ストリートビューで物件を特定し、
照明の設置場所や抜け漏れを検証するスキーム。

使い方:
  1. run_maps_check(survey) → MapCheckResult を取得
  2. 結果にはGoogleマップURL + AI検証レポートが含まれる
  3. Excel出力に検証メモを追記（オプション）

APIキー不要の基本モード:
  - GoogleマップURL生成（ブラウザで開ける）
  - ストリートビューURL生成
  - 照明チェックリスト自動生成

APIキーありの拡張モード:
  - ストリートビュー画像取得 → AI分析
  - 建物外観から照明設置箇所を推定
  - 現調データとの差分レポート
"""

from __future__ import annotations

import json
import logging
import os
import re
import urllib.parse
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import yaml

logger = logging.getLogger(__name__)


@dataclass
class MapCheckResult:
    """Googleマップ検証結果"""
    address: str = ""
    maps_url: str = ""                  # Googleマップ検索URL
    streetview_url: str = ""            # ストリートビューURL
    checklist: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    missing_areas: list[str] = field(default_factory=list)
    confidence: str = "未検証"          # "未検証", "確認済", "要確認"


# ===== 建物タイプ別の照明チェックポイント =====

COMMON_LIGHTING_AREAS = [
    "共用廊下",
    "階段室",
    "エントランス/ロビー",
    "駐車場/駐輪場",
    "外部通路/アプローチ",
    "ゴミ置場",
    "集合ポスト周辺",
    "エレベーターホール",
]

BUILDING_TYPE_AREAS = {
    "マンション": [
        *COMMON_LIGHTING_AREAS,
        "非常灯（各階）",
        "屋上/塔屋",
        "管理人室",
        "防犯灯",
    ],
    "アパート": [
        "共用廊下",
        "階段室",
        "エントランス",
        "駐車場/駐輪場",
        "外部通路",
        "ゴミ置場",
        "ポスト周辺",
        "防犯灯",
    ],
    "事務所": [
        "エントランス/ロビー",
        "事務室",
        "会議室",
        "給湯室/トイレ",
        "廊下",
        "階段室",
        "非常灯",
        "外部看板灯",
        "駐車場",
    ],
    "店舗": [
        "店内照明",
        "外部看板灯",
        "ショーウィンドウ",
        "バックヤード",
        "トイレ",
        "駐車場",
    ],
    "default": COMMON_LIGHTING_AREAS,
}


def generate_maps_url(address: str) -> str:
    """GoogleマップのURL生成"""
    if not address:
        return ""
    encoded = urllib.parse.quote(address)
    return f"https://www.google.com/maps/search/?api=1&query={encoded}"


def generate_streetview_url(address: str) -> str:
    """GoogleストリートビューのURL生成"""
    if not address:
        return ""
    encoded = urllib.parse.quote(address)
    return f"https://www.google.com/maps/@?api=1&map_action=pano&query={encoded}"


def _detect_building_type(property_name: str, address: str) -> str:
    """物件名・住所から建物タイプを推定"""
    text = f"{property_name} {address}"

    type_keywords = {
        "マンション": ["マンション", "レジデンス", "タワー", "パレス", "ハイツ", "棟"],
        "アパート": ["アパート", "コーポ", "メゾネット", "ハイム", "荘", "コート"],
        "事務所": ["事務所", "オフィス", "ビル", "ビジネス"],
        "店舗": ["店舗", "ショップ", "モール", "テナント"],
    }

    for btype, keywords in type_keywords.items():
        if any(kw in text for kw in keywords):
            return btype

    return "default"


def _generate_checklist(
    building_type: str,
    existing_locations: list[str],
) -> tuple[list[str], list[str]]:
    """照明チェックリストを生成し、抜け漏れ候補を検出

    Args:
        building_type: 建物タイプ
        existing_locations: 現調データに含まれる設置場所一覧

    Returns:
        (checklist, missing_areas)
    """
    expected_areas = BUILDING_TYPE_AREAS.get(
        building_type,
        BUILDING_TYPE_AREAS["default"],
    )

    # 現調データの場所を正規化
    existing_norm = set()
    for loc in existing_locations:
        loc_norm = loc.replace(" ", "").replace("　", "")
        existing_norm.add(loc_norm)

    checklist = []
    missing = []

    for area in expected_areas:
        # 部分一致で確認
        area_keywords = area.replace("/", "").replace("（", "").replace("）", "")
        found = any(
            any(kw in exist for kw in _split_area_keywords(area))
            for exist in existing_norm
        )

        if found:
            checklist.append(f"✓ {area}")
        else:
            checklist.append(f"✗ {area} ← 未確認")
            missing.append(area)

    return checklist, missing


def _split_area_keywords(area: str) -> list[str]:
    """エリア名からマッチング用キーワードを生成"""
    # "共用廊下" → ["共用廊下", "廊下", "共用"]
    # "駐車場/駐輪場" → ["駐車場", "駐輪場", "駐車", "駐輪"]
    keywords = []
    for part in area.replace("（", "/").replace("）", "").split("/"):
        part = part.strip()
        if part:
            keywords.append(part)
            # 2文字以上なら先頭2文字もキーワードに
            if len(part) >= 3:
                keywords.append(part[:2])
    return keywords


def run_maps_check(
    survey_data,
    property_name: Optional[str] = None,
) -> MapCheckResult:
    """Googleマップ検証を実行

    Args:
        survey_data: SurveyData（models.py）
        property_name: 物件名（省略時はsurvey_dataから取得）

    Returns:
        MapCheckResult
    """
    info = survey_data.property_info
    address = info.address or ""
    name = property_name or info.name or ""

    if not address:
        return MapCheckResult(
            notes=["住所が未入力のため検証できません"],
            confidence="未検証",
        )

    # URL生成
    maps_url = generate_maps_url(address)
    streetview_url = generate_streetview_url(address)

    # 建物タイプ推定
    building_type = _detect_building_type(name, address)

    # 既存器具の設置場所を収集
    existing_locations = []
    for f in survey_data.fixtures:
        if f.location:
            existing_locations.append(f.location)
    for f in survey_data.excluded_fixtures:
        if f.location:
            existing_locations.append(f.location)

    # チェックリスト生成
    checklist, missing = _generate_checklist(
        building_type, existing_locations,
    )

    # 検証メモ
    notes = [
        f"建物タイプ推定: {building_type}",
        f"現調済み設置場所: {len(existing_locations)}箇所",
    ]
    if missing:
        notes.append(
            f"未確認エリア: {len(missing)}箇所 → Googleマップで要確認"
        )

    confidence = "要確認" if missing else "確認済"

    result = MapCheckResult(
        address=address,
        maps_url=maps_url,
        streetview_url=streetview_url,
        checklist=checklist,
        notes=notes,
        missing_areas=missing,
        confidence=confidence,
    )

    logger.info(
        f"マップ検証: {name} ({building_type}) "
        f"チェック={len(checklist)}項目 未確認={len(missing)}箇所"
    )

    return result


def format_check_report(result: MapCheckResult) -> str:
    """検証結果を人間可読なレポート文字列にフォーマット"""
    lines = []
    lines.append("=" * 50)
    lines.append("Googleマップ照明検証レポート")
    lines.append("=" * 50)
    lines.append(f"住所: {result.address}")
    lines.append(f"判定: {result.confidence}")
    lines.append("")

    if result.maps_url:
        lines.append(f"Googleマップ: {result.maps_url}")
    if result.streetview_url:
        lines.append(f"ストリートビュー: {result.streetview_url}")
    lines.append("")

    lines.append("--- チェックリスト ---")
    for item in result.checklist:
        lines.append(f"  {item}")
    lines.append("")

    if result.missing_areas:
        lines.append("--- 未確認エリア（要確認） ---")
        for area in result.missing_areas:
            lines.append(f"  ! {area}")
        lines.append("")

    for note in result.notes:
        lines.append(f"[memo] {note}")

    return "\n".join(lines)


# ===== AI異常検知 =====


def _build_fixture_text(survey_data, matches) -> str:
    """見積データを構造化テキストに変換"""
    lines = []
    match_map = {}
    for m in matches:
        match_map[m.fixture.row_label] = m

    for fix in survey_data.fixtures:
        qty_total = fix.quantities.total if fix.quantities else 0
        line = (
            f"行{fix.row_label}: "
            f"場所={fix.location}, "
            f"器具={fix.fixture_type}, "
            f"電球={fix.bulb_type}, "
            f"W={fix.power_consumption_w}, "
            f"点灯={fix.daily_hours}h/日, "
            f"色温度={fix.color_temp}, "
            f"数量={qty_total}個, "
            f"防水={'有' if fix.is_waterproof else '無'}"
        )
        m = match_map.get(fix.row_label)
        if m and m.led_product:
            led = m.led_product
            line += (
                f" → LED: {led.name}, "
                f"{led.power_w}W, "
                f"防水={'有' if led.is_waterproof else '無'}, "
                f"信頼度={m.confidence:.0%}"
            )
            if m.match_notes:
                line += f", メモ={m.match_notes}"
        lines.append(line)

    for fix in survey_data.excluded_fixtures:
        lines.append(
            f"行{fix.row_label}: "
            f"場所={fix.location}, "
            f"器具={fix.fixture_type} "
            f"【除外: {fix.exclusion_reason}】"
        )

    return "\n".join(lines)


def run_ai_anomaly_check(
    survey_data,
    matches: list,
    checklist_result: "MapCheckResult",
    api_key: Optional[str] = None,
) -> dict:
    """AIで見積データの違和感・異常を検出

    Args:
        survey_data: SurveyData
        matches: list[MatchResult]
        checklist_result: MapCheckResult（既存チェックリスト結果）
        api_key: Anthropic APIキー（省略時は環境変数）

    Returns:
        {"anomalies": [...], "summary": "..."}
    """
    if api_key is None:
        api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        logger.warning("APIキー未設定のためAI異常検知をスキップ")
        return {"anomalies": [], "summary": "APIキー未設定"}

    # プロンプト読み込み
    config_path = Path(__file__).parent.parent / "config" / "ai_prompts.yaml"
    with open(config_path, "r", encoding="utf-8") as f:
        prompts = yaml.safe_load(f)
    prompt_template = prompts.get("quotation_anomaly_check", "")
    if not prompt_template:
        logger.warning("quotation_anomaly_check プロンプトが未定義")
        return {"anomalies": [], "summary": "プロンプト未定義"}

    # データ組み立て
    info = survey_data.property_info
    building_type = _detect_building_type(info.name, info.address)
    property_info = (
        f"物件名: {info.name}\n"
        f"住所: {info.address}\n"
        f"建物タイプ推定: {building_type}"
    )
    fixture_data = _build_fixture_text(survey_data, matches)
    checklist_data = "\n".join(checklist_result.checklist) if checklist_result.checklist else "チェックリストなし"

    prompt = (
        prompt_template
        .replace("{property_info}", property_info)
        .replace("{fixture_data}", fixture_data)
        .replace("{checklist_data}", checklist_data)
    )

    # API呼び出し（テキストのみ）— claude_guard 経由
    try:
        import sys
        from pathlib import Path
        sys.path.insert(0, str(Path(__file__).parent.parent.parent))
        from claude_guard import get_guarded_client
        client = get_guarded_client(api_key=api_key)
        message = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=2048,
            messages=[{"role": "user", "content": prompt}],
        )
        response_text = message.content[0].text
    except Exception as e:
        logger.warning(f"AI異常検知APIエラー: {e}")
        return {"anomalies": [], "summary": f"APIエラー: {e}"}

    # JSON抽出
    result = _extract_anomaly_json(response_text)

    anomaly_count = len(result.get("anomalies", []))
    logger.info(f"AI異常検知完了: {anomaly_count}件の指摘")

    return result


def _extract_anomaly_json(response_text: str) -> dict:
    """レスポンスからJSONを抽出"""
    text = response_text.strip()

    # ```json ... ``` ブロック
    if "```json" in text:
        try:
            start = text.index("```json") + 7
            end = text.index("```", start)
            return json.loads(text[start:end].strip())
        except (ValueError, json.JSONDecodeError):
            pass

    # 最外の { ... }
    match = re.search(r'\{[\s\S]*\}', text)
    if match:
        try:
            return json.loads(match.group())
        except json.JSONDecodeError:
            pass

    return {"anomalies": [], "summary": "JSON解析失敗", "raw": text[:500]}


def format_anomaly_report(result: dict) -> str:
    """AI異常検知結果をレポート文字列にフォーマット"""
    lines = []
    lines.append("=" * 50)
    lines.append("AI異常検知レポート")
    lines.append("=" * 50)

    anomalies = result.get("anomalies", [])
    if not anomalies:
        lines.append("  指摘事項なし")
    else:
        for a in anomalies:
            severity = a.get("severity", "info")
            icon = "⚠" if severity == "warning" else "ℹ"
            lines.append(f"  {icon} [{a.get('category', '')}] {a.get('target', '')}")
            lines.append(f"    {a.get('message', '')}")
            if a.get("suggestion"):
                lines.append(f"    → {a['suggestion']}")
            lines.append("")

    summary = result.get("summary", "")
    if summary:
        lines.append(f"[総評] {summary}")

    return "\n".join(lines)
