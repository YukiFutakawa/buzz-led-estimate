"""OCR結果 → SurveyData 構造化データ変換

document_processor.py のOCR出力JSONを
models.py の ExistingFixture / PropertyInfo / SurveyData に変換する。
"""

from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Optional

from models import (
    ExistingFixture,
    FloorQuantities,
    PropertyInfo,
    SurveyData,
)

logger = logging.getLogger(__name__)


def parse_survey_ocr(
    ocr_result: dict,
    fixture_photos: Optional[dict[str, Path | list[Path]]] = None,
) -> SurveyData:
    """OCR結果をSurveyDataに変換

    Args:
        ocr_result: document_processor.ocr_survey_sheet() の出力
        fixture_photos: 行ラベル→写真パスのマッピング
            旧形式: {"A": Path("photos/A.jpg")} — 後方互換
            新形式: {"A": [Path("photos/A1.jpg"), Path("photos/A2.jpg")]}

    Returns:
        SurveyData
    """
    if fixture_photos is None:
        fixture_photos = {}

    # 旧形式(dict[str, Path])を新形式(dict[str, list[Path]])に統一
    normalized_photos: dict[str, list[Path]] = {}
    for label, val in fixture_photos.items():
        if isinstance(val, Path):
            normalized_photos[label] = [val]
        elif isinstance(val, list):
            normalized_photos[label] = val
        else:
            normalized_photos[label] = [Path(val)]
    fixture_photos = normalized_photos

    # ヘッダー情報
    header = ocr_result.get("header", {})
    if not header:
        # 写真直接解析モードの場合、property_infoキーを使用
        pi = ocr_result.get("property_info", {})
        if pi:
            header = {"property_name": pi.get("name", "")}
    property_info = _parse_header(header)

    # 器具データ（fixtures + excluded_fixtures を統合）
    raw_fixtures = ocr_result.get("fixtures", [])
    raw_excluded = ocr_result.get("excluded_fixtures", [])
    all_raw = raw_fixtures + raw_excluded
    fixtures = []
    excluded = []

    for raw in all_raw:
        fixture = _parse_fixture(raw, fixture_photos)
        if fixture.is_excluded:
            excluded.append(fixture)
        else:
            fixtures.append(fixture)

    # 備考
    special_notes = ocr_result.get("special_notes", "")
    if special_notes and not property_info.special_notes:
        property_info.special_notes = special_notes

    survey = SurveyData(
        property_info=property_info,
        fixtures=fixtures,
        excluded_fixtures=excluded,
    )

    logger.info(
        f"OCRパース完了: 物件={property_info.name}, "
        f"器具={len(fixtures)}件, 除外={len(excluded)}件"
    )
    return survey


def _parse_header(header: dict) -> PropertyInfo:
    """ヘッダー部分をPropertyInfoに変換"""
    return PropertyInfo(
        name=header.get("property_name", ""),
        address=header.get("address", ""),
        unlock_code=header.get("unlock_code", ""),
        distribution_board=header.get("distribution_board", ""),
        special_notes=header.get("special_notes", ""),
        survey_date=header.get("survey_date", ""),
        surveyor=header.get("surveyor", ""),
    )


def _parse_fixture(
    raw: dict,
    fixture_photos: dict[str, list[Path]],
) -> ExistingFixture:
    """1つの器具データをExistingFixtureに変換"""

    row_label = raw.get("row_label", "")

    # 階別数量の変換
    floor_quantities = _parse_floor_quantities(
        raw.get("floor_quantities", {})
    )

    # 電球種別から消費電力の補完
    bulb_type = raw.get("bulb_type", "")
    power_w = _safe_float(raw.get("power_w", 0))
    if power_w == 0 and bulb_type:
        power_w = _estimate_power_from_bulb(bulb_type)

    # 色温度の正規化
    color_temp = _normalize_color_temp(raw.get("color_temp", ""))

    # 器具サイズの正規化
    fixture_size = _normalize_size(raw.get("fixture_size", ""))

    # 防水判定
    fixture_type = raw.get("fixture_type", "")
    location = raw.get("location", "")
    is_waterproof = _detect_waterproof(fixture_type, location)

    # 除外判定
    is_excluded = raw.get("is_excluded", False)
    exclusion_reason = raw.get("exclusion_reason", "")
    if not is_excluded and "LED" in bulb_type.upper():
        is_excluded = True
        exclusion_reason = exclusion_reason or "LED済み"

    # 写真パス（新形式: dict[str, list[Path]]に対応）
    photo_paths = []
    if row_label in fixture_photos:
        photo_paths = list(fixture_photos[row_label])

    # OCR信頼度の伝搬
    ocr_confidence = raw.get("confidence", "high")
    if ocr_confidence not in ("high", "medium", "low"):
        ocr_confidence = "high"
    ocr_warnings = raw.get("_validation_warnings", [])

    return ExistingFixture(
        row_label=row_label,
        location=location,
        fixture_type=fixture_type,
        fixture_size=fixture_size,
        bulb_type=bulb_type,
        quantities=floor_quantities,
        power_consumption_w=power_w,
        daily_hours=_safe_float(raw.get("daily_hours", 0)),
        color_temp=color_temp,
        survey_notes=raw.get("notes", ""),
        construction_notes=raw.get("construction_notes", ""),
        photo_paths=photo_paths,
        is_waterproof=is_waterproof,
        is_excluded=is_excluded,
        exclusion_reason=exclusion_reason,
        ocr_confidence=ocr_confidence,
        ocr_warnings=ocr_warnings,
    )


def _parse_floor_quantities(raw: dict) -> FloorQuantities:
    """階別数量のパース

    入力例: {"1F": 5, "2F": 3} or {"1P": 5, "2P": 3}
    """
    floors = {}
    for key, val in raw.items():
        # "1F", "2F", "1P", "2P" 等から数字を抽出
        m = re.match(r"(\d+)", str(key))
        if m:
            floor_num = int(m.group(1))
            qty = _safe_int(val)
            if qty > 0:
                floors[floor_num] = qty
    return FloorQuantities(floors=floors)


def _estimate_power_from_bulb(bulb_type: str) -> float:
    """電球種別からの消費電力推定

    例: FL20 → 20W, FDL13 → 13W, 白熱60W → 60W
    """
    # 数値を抽出
    m = re.search(r"(\d+)\s*[Ww]?", bulb_type)
    if m:
        return float(m.group(1))
    return 0


def _normalize_color_temp(raw: str) -> str:
    """色温度の正規化"""
    raw = raw.strip()
    if not raw:
        return ""

    # チェックマーク位置での判定
    if raw in ("白", "N", "昼白色", "昼白"):
        return "白"
    if raw in ("黄", "L", "電球色", "電球"):
        return "黄"
    if "白" in raw:
        return "白"
    if "黄" in raw:
        return "黄"
    return raw


def _normalize_size(raw: str) -> str:
    """器具サイズの正規化

    手書きのバリエーションを統一:
    "W×150 D×90" → "W150×D90"
    "w 150 d 90" → "W150×D90"
    "φ100" → "Φ100"
    """
    if not raw:
        return ""

    s = raw.strip()

    # スペースの正規化
    s = re.sub(r'\s+', '', s)

    # W×150D×90 → W150×D90 パターン
    s = re.sub(r'[Ww]×?(\d+)', r'W\1', s)
    s = re.sub(r'[Dd]×?(\d+)', r'D\1', s)
    s = re.sub(r'[Hh]×?(\d+)', r'H\1', s)

    # φ → Φ
    s = s.replace('φ', 'Φ').replace('ф', 'Φ')

    # 区切りがない場合に×を追加
    # W150D90 → W150×D90
    s = re.sub(r'(\d)([DHWdhw])', r'\1×\2', s)

    return s


def _detect_waterproof(fixture_type: str, location: str) -> bool:
    """防水が必要な器具かどうかを推定"""
    text = f"{fixture_type} {location}"
    waterproof_keywords = (
        "防滴", "防水", "防雨", "防湿", "屋外",
        "外部", "ポーチ", "外壁", "バルコニー",
        "駐車場", "駐輪場", "通路",
    )
    return any(kw in text for kw in waterproof_keywords)


def _safe_float(val) -> float:
    """安全なfloat変換"""
    if val is None or val == "":
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        # 数値部分を抽出
        m = re.search(r'[\d.]+', str(val))
        if m:
            return float(m.group())
        return 0.0


def _safe_int(val) -> int:
    """安全なint変換"""
    if val is None or val == "":
        return 0
    try:
        return int(float(val))
    except (ValueError, TypeError):
        m = re.search(r'\d+', str(val))
        if m:
            return int(m.group())
        return 0


# ===== 写真ファイル紐付けユーティリティ =====

def match_photos_to_fixtures(
    photo_dir: Path,
    row_labels: Optional[list[str]] = None,
    exclude_paths: Optional[list[Path]] = None,
) -> dict[str, Path]:
    """写真ファイルを行ラベルに紐付け

    ファイル名の規則:
    - "A_玄関.jpg" → row_label "A"
    - "01_lobby.jpg" → 順番で割り当て
    - ファイル名ソート順で行ラベルA, B, C...に割り当て

    Args:
        photo_dir: 写真フォルダ
        row_labels: 割り当て先の行ラベル（省略時はA-T）
        exclude_paths: 除外する写真パス（チェックシート等）

    Returns:
        行ラベル→写真パスのdict
    """
    if row_labels is None:
        row_labels = list("ABCDEFGHIJKLMNOPQRST")

    if exclude_paths is None:
        exclude_paths = []
    exclude_set = {p.resolve() for p in exclude_paths}

    if not photo_dir.exists():
        logger.warning(f"写真フォルダが見つかりません: {photo_dir}")
        return {}

    # 画像ファイルを収集（チェックシート等を除外）
    image_exts = {".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".heic"}
    photos = sorted(
        f for f in photo_dir.iterdir()
        if f.suffix.lower() in image_exts
        and f.resolve() not in exclude_set
    )

    if not photos:
        return {}

    result: dict[str, Path] = {}

    for photo in photos:
        name = photo.stem.upper()

        # パターン1: ファイル名が行ラベルで始まる（"A_玄関.jpg"）
        matched = False
        for label in row_labels:
            if name.startswith(label + "_") or name == label:
                if label not in result:
                    result[label] = photo
                    matched = True
                    break
        if matched:
            continue

    # パターン2: 未割り当ての写真を順番にマッピング
    unassigned = [p for p in photos if p not in result.values()]
    unused_labels = [l for l in row_labels if l not in result]

    for label, photo in zip(unused_labels, unassigned):
        result[label] = photo

    logger.info(f"写真紐付け: {len(result)}枚 ({photo_dir.name})")
    return result
