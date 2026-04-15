"""Kintone REST API クライアント (見積作成 新方式用)

App63(LED現調) / App68(現調器具明細) からデータ取得＋添付ファイルDL。
見積作成Streamlitの新方式フローで使用する。
"""

from __future__ import annotations

import os
import logging
import requests
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)


class KintoneConfig:
    """環境変数から設定を読み込む"""

    def __init__(self):
        self.domain = os.environ.get("KINTONE_DOMAIN", "wsd92wi2row6.cybozu.com")
        self.app63_id = int(os.environ.get("KINTONE_APP63_ID", "63"))
        self.app63_token = os.environ.get("KINTONE_APP63_TOKEN") or os.environ.get("KINTONE_SURVEY_TOKEN")
        self.fixture_app_id = int(os.environ.get("KINTONE_FIXTURE_DETAIL_APP_ID", "68"))
        self.fixture_token = os.environ.get("KINTONE_FIXTURE_DETAIL_TOKEN")
        # ファイルDLに両方のトークンが必要になる場合があるので複合ヘッダに使う

    @property
    def base_url(self) -> str:
        return f"https://{self.domain}/k/v1"

    def validate(self) -> Optional[str]:
        """設定チェック。不足があればエラーメッセージを返す"""
        missing = []
        if not self.app63_token:
            missing.append("KINTONE_APP63_TOKEN (または KINTONE_SURVEY_TOKEN)")
        if not self.fixture_token:
            missing.append("KINTONE_FIXTURE_DETAIL_TOKEN")
        if missing:
            return "環境変数が未設定です: " + ", ".join(missing)
        return None


def _request(method: str, url: str, token: str, **kwargs) -> dict:
    headers = kwargs.pop("headers", {}) or {}
    headers["X-Cybozu-API-Token"] = token
    resp = requests.request(method, url, headers=headers, timeout=30, **kwargs)
    if not resp.ok:
        logger.error(f"Kintone API失敗: {method} {url} status={resp.status_code} body={resp.text[:400]}")
        resp.raise_for_status()
    # ファイルDL時のレスポンスはJSONでないため呼び出し側で resp を直接扱う
    return resp


def fetch_eligible_projects(cfg: KintoneConfig, limit: int = 500) -> list[dict]:
    """見積進捗=02.数量入力完了 の App63レコードを取得。

    Returns:
        list[dict]: [{record_id, property_name, address, submitted_at, ...}, ...]
    """
    # ドロップダウン_5 = 見積進捗
    query = '"ドロップダウン_5" = "02.数量入力完了" order by mypage_submitted_at desc limit ' + str(limit)
    resp = _request(
        "GET",
        f"{cfg.base_url}/records.json",
        cfg.app63_token,
        params={
            "app": cfg.app63_id,
            "query": query,
        },
    )
    records = resp.json().get("records", [])
    result = []
    for r in records:
        result.append({
            "record_id": int(r["$id"]["value"]),
            "property_name": _val(r, "property_name"),
            "address": _val(r, "address"),
            "submitted_at": _val(r, "mypage_submitted_at"),
            # 物件総合データ
            "prop_parking_move": _val(r, "prop_parking_move"),
            "prop_parking_move_no": _val(r, "prop_parking_move_no"),
            "prop_panel_key": _val(r, "prop_panel_key"),
            "prop_lm_apart_qty": _val(r, "prop_lm_apart_qty"),
            "prop_lm_mansion_qty": _val(r, "prop_lm_mansion_qty"),
            "prop_lm_timer_qty": _val(r, "prop_lm_timer_qty"),
            "prop_special_work": _val(r, "prop_special_work"),
            "prop_special_work_content": _val(r, "prop_special_work_content"),
            "prop_special_work_amount": _val(r, "prop_special_work_amount"),
            # 物件写真(fileKey配列)
            "prop_nameplate_photo": _files(r, "prop_nameplate_photo"),
            "prop_lm_photo": _files(r, "prop_lm_photo"),
            "prop_special_work_photo": _files(r, "prop_special_work_photo"),
        })
    return result


def fetch_fixture_entries(cfg: KintoneConfig, parent_record_id: int) -> list[dict]:
    """App68 から parent_record_id の全器具明細レコードを取得"""
    query = f'parent_record_id = "{parent_record_id}" order by $id asc limit 500'
    resp = _request(
        "GET",
        f"{cfg.base_url}/records.json",
        cfg.fixture_token,
        params={
            "app": cfg.fixture_app_id,
            "query": query,
        },
    )
    records = resp.json().get("records", [])
    result = []
    for r in records:
        item = {
            "record_id": int(r["$id"]["value"]),
            "location_base": _val(r, "location_base"),
            "location_type": _val(r, "location_type"),
            "location_other": _val(r, "location_other"),
            "location_display": _val(r, "location_display"),
            "fixture_shape": _val(r, "fixture_shape"),
            "lamp_count": _val(r, "lamp_count"),
            "fixture_kind": _val(r, "fixture_kind"),
            "fixture_display": _val(r, "fixture_display"),
            "size_h": _val(r, "size_h"),
            "size_w": _val(r, "size_w"),
            "size_d": _val(r, "size_d"),
            "bulb_type": _val(r, "bulb_type"),
            "wattage": _val(r, "wattage"),
            "lighting_time": _val(r, "lighting_time"),
            "color_temp": _val(r, "color_temp"),
            "waterproof": _val(r, "waterproof"),
            "photo_fixture": _files(r, "photo_fixture"),
            "photo_bulb": _files(r, "photo_bulb"),
            "photo_inside": _files(r, "photo_inside"),
            "photo_other": _files(r, "photo_other"),
        }
        # 階別数量
        item["qty_by_floor"] = {}
        for f in range(1, 11):
            v = _val(r, f"qty_{f}f")
            try:
                n = int(v) if v else 0
            except (TypeError, ValueError):
                n = 0
            if n > 0:
                item["qty_by_floor"][f] = n
        result.append(item)
    return result


def download_file(cfg: KintoneConfig, file_key: str, dest: Path, token: str = None) -> Path:
    """Kintoneの添付ファイルを dest にDLする"""
    use_token = token or cfg.fixture_token or cfg.app63_token
    resp = _request(
        "GET",
        f"{cfg.base_url}/file.json",
        use_token,
        params={"fileKey": file_key},
        stream=True,
    )
    dest.parent.mkdir(parents=True, exist_ok=True)
    with open(dest, "wb") as f:
        for chunk in resp.iter_content(chunk_size=65536):
            if chunk:
                f.write(chunk)
    return dest


def download_photos_for_fixture(cfg: KintoneConfig, fixture: dict, dest_dir: Path) -> dict:
    """1つの器具レコードの全写真をDL。

    Returns:
        dict: {"fixture": [Path,...], "bulb": [Path,...], "inside": [...], "other": [...]}
    """
    dest_dir.mkdir(parents=True, exist_ok=True)
    result = {"fixture": [], "bulb": [], "inside": [], "other": []}
    for kind in ("fixture", "bulb", "inside", "other"):
        files = fixture.get(f"photo_{kind}") or []
        for idx, fobj in enumerate(files):
            ext = Path(fobj.get("name", "") or ".jpg").suffix or ".jpg"
            dest = dest_dir / f"rec{fixture['record_id']}_{kind}_{idx+1}{ext}"
            try:
                download_file(cfg, fobj["fileKey"], dest, token=cfg.fixture_token)
                result[kind].append(dest)
            except Exception as e:
                logger.warning(f"写真DL失敗 rec={fixture['record_id']} kind={kind}: {e}")
    return result


# ===== 内部ヘルパ =====

def _val(record: dict, code: str) -> str:
    field = record.get(code)
    if not field:
        return ""
    v = field.get("value", "")
    if isinstance(v, list):
        return v
    return v if v is not None else ""


def _files(record: dict, code: str) -> list[dict]:
    """添付フィールドから [{"fileKey":..., "name":..., "contentType":...}] を返す"""
    field = record.get(code)
    if not field:
        return []
    v = field.get("value", [])
    if not isinstance(v, list):
        return []
    return [{"fileKey": f.get("fileKey"), "name": f.get("name", ""), "contentType": f.get("contentType", "")} for f in v]
