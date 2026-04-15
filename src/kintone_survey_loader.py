"""Kintoneデータ → Streamlit session_state 変換ローダー (新方式用)"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Optional

from kintone_survey_client import (
    KintoneConfig,
    fetch_fixture_entries,
    download_photos_for_fixture,
)

logger = logging.getLogger(__name__)


# lighting_time → daily_hours 変換
_LIGHTING_TIME_TO_HOURS = {
    "12時間(明暗センサー)": 12.0,
    "12時間(タイマー)": 12.0,
    "24時間(常時点灯)": 24.0,
    "スイッチ": 0.0,
    "人感センサー(器具単体)": 0.0,
}


def _seq_label(idx: int) -> str:
    """0→A, 1→B, ..., 25→Z, 26→AA ..."""
    label = ""
    n = idx
    while True:
        label = chr(ord("A") + n % 26) + label
        n = n // 26 - 1
        if n < 0:
            break
    return label


def _format_fixture_size(h, w, d) -> str:
    """size_h, size_w, size_d から '120×200' / 'φ200' 形式の文字列を返す"""
    def num(v):
        try:
            return int(float(v)) if v not in (None, "") else 0
        except (TypeError, ValueError):
            return 0

    H, W, D = num(h), num(w), num(d)
    if D > 0 and H == 0 and W == 0:
        return f"φ{D}"
    if H > 0 and W > 0:
        return f"{H}×{W}"
    if H > 0:
        return str(H)
    if W > 0:
        return str(W)
    if D > 0:
        return f"φ{D}"
    return ""


def _lighting_time_to_hours(val: str) -> float:
    return _LIGHTING_TIME_TO_HOURS.get(val, 0.0)


def convert_fixture_to_dict(fx: dict, row_label: str) -> dict:
    """App68の1レコードを confirmed_fixtures 形式のdictに変換"""
    try:
        power_w = float(fx.get("wattage") or 0)
    except (TypeError, ValueError):
        power_w = 0.0

    return {
        "row_label": row_label,
        "location": fx.get("location_display") or fx.get("location_base") or "",
        "fixture_type": fx.get("fixture_display") or fx.get("fixture_kind") or "",
        "bulb_type": fx.get("bulb_type") or "",
        "fixture_size": _format_fixture_size(fx.get("size_h"), fx.get("size_w"), fx.get("size_d")),
        "floor_quantities": dict(fx.get("qty_by_floor") or {}),
        "power_w": power_w,
        "daily_hours": _lighting_time_to_hours(fx.get("lighting_time") or ""),
        "color_temp": fx.get("color_temp") or "",
        "is_excluded": False,
        "exclusion_reason": "",
        # 新方式固有情報
        "_source_record_id": fx.get("record_id"),
        "_source_waterproof": fx.get("waterproof"),
        "_source_lamp_count": fx.get("lamp_count"),
        "_source_fixture_shape": fx.get("fixture_shape"),
    }


def load_project_from_kintone(
    cfg: KintoneConfig,
    project: dict,
    tmpdir: Path,
) -> dict:
    """App63のproject dict を受け取り、App68の器具明細+写真をDLして整形済みデータを返す。

    Args:
        cfg: KintoneConfig
        project: fetch_eligible_projects() が返す要素
        tmpdir: 写真DL先の一時ディレクトリ

    Returns:
        dict: {
            "confirmed_fixtures": [...],
            "confirmed_photos": {row_label: [path_str, ...]},
            "photo_by_kind": {row_label: {"fixture": [...], "bulb": [...], "inside": [...], "other": [...]}},
            "property_summary": {...},  # 読取専用表示用
            "step_config": {...},
            "property_info": {...},
        }
    """
    record_id = project["record_id"]
    fixtures_raw = fetch_fixture_entries(cfg, record_id)
    if not fixtures_raw:
        raise ValueError(f"App68に器具明細が見つかりません (parent_record_id={record_id})")

    photos_dir = tmpdir / "kintone_photos"
    photos_dir.mkdir(parents=True, exist_ok=True)

    confirmed_fixtures = []
    confirmed_photos: dict[str, list[str]] = {}
    photo_by_kind: dict[str, dict[str, list[str]]] = {}
    # Excel反映対象チェック用の初期値（器具・電球はデフォON、内部・その他はOFF）
    photo_selection: dict[str, dict[str, list[bool]]] = {}

    for idx, fx in enumerate(fixtures_raw):
        label = _seq_label(idx)
        confirmed_fixtures.append(convert_fixture_to_dict(fx, label))

        # 写真DL
        dl = download_photos_for_fixture(cfg, fx, photos_dir / label)
        kind_paths = {k: [str(p) for p in v] for k, v in dl.items()}
        photo_by_kind[label] = kind_paths

        # 固定順: fixture → bulb → inside → other で confirmed_photos に格納
        ordered = []
        selection = {"fixture": [], "bulb": [], "inside": [], "other": []}
        for k in ("fixture", "bulb", "inside", "other"):
            for p in kind_paths[k]:
                ordered.append(p)
                selection[k].append(True if k in ("fixture", "bulb") else False)
        confirmed_photos[label] = ordered
        photo_selection[label] = selection

    # property_info（旧方式フォーマット: unlock_code / distribution_board / special_notes）
    panel_key = project.get("prop_panel_key") or ""
    special_work_content = project.get("prop_special_work_content") or ""
    property_info = {
        "unlock_code": "",
        "distribution_board": panel_key,
        "special_notes": special_work_content,
    }

    # 読取専用で画面表示する現調サマリー
    property_summary = {
        "property_name": project.get("property_name", ""),
        "address": project.get("address", ""),
        "submitted_at": project.get("submitted_at", ""),
        "parking_move": project.get("prop_parking_move") or "",
        "parking_move_no": project.get("prop_parking_move_no") or "",
        "panel_key": panel_key,
        "lm_apart_qty": project.get("prop_lm_apart_qty") or "",
        "lm_mansion_qty": project.get("prop_lm_mansion_qty") or "",
        "lm_timer_qty": project.get("prop_lm_timer_qty") or "",
        "special_work": project.get("prop_special_work") or "",
        "special_work_content": special_work_content,
        "special_work_amount": project.get("prop_special_work_amount") or "",
    }

    step_config = {
        "property_name": project.get("property_name", ""),
        "address": project.get("address", ""),
        "selected_company": "（指定なし）",
        "template_name": "田村基本形",
        "source": "kintone_new_mode",
        "source_record_id": record_id,
    }

    return {
        "confirmed_fixtures": confirmed_fixtures,
        "confirmed_excluded": [],
        "confirmed_photos": confirmed_photos,
        "photo_by_kind": photo_by_kind,
        "photo_selection": photo_selection,
        "property_summary": property_summary,
        "property_info": property_info,
        "step_config": step_config,
    }


def apply_photo_selection(session_state) -> dict:
    """photo_selection のチェック状態に基づき confirmed_photos を更新し、新しい辞書を返す"""
    photo_by_kind: dict = session_state.get("photo_by_kind") or {}
    selection: dict = session_state.get("photo_selection") or {}
    confirmed_photos = {}
    for label, kinds in photo_by_kind.items():
        sel = selection.get(label, {})
        ordered = []
        for k in ("fixture", "bulb", "inside", "other"):
            paths = kinds.get(k, [])
            flags = sel.get(k, [True] * len(paths))
            for p, checked in zip(paths, flags):
                if checked:
                    ordered.append(p)
        confirmed_photos[label] = ordered
    return confirmed_photos
