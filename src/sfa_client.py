"""ネクストSFA API クライアント

案件データの取得・フィルタリングを行う。
"""

from __future__ import annotations

import json
import os
import re
import urllib.request
import urllib.error
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional


BASE_URL = "https://nsfa.next-cloud.jp"
ITEMS_PER_PAGE = 50  # APIが受け付ける最大値


@dataclass
class SFAProject:
    """SFA案件データ（LED見積に必要なフィールドのみ）"""

    id: str
    name: str  # 物件名
    address: str  # contents_7
    unlock_info: str  # contents_1 (解錠コード)
    management_company: str  # contents_4 (管理会社)
    memo: str  # 物件情報メモ
    phase_category: Optional[str]  # 案件決着済み / None
    phase: Optional[str]  # 見込み / 失注 / 入金済み etc.
    survey_date: Optional[str]  # contents_6 (現調日)
    construction_date: Optional[str]  # contents_9 (工事日)
    company_id: Optional[str] = None
    modified: Optional[str] = None

    # --- memo から抽出される情報 ---
    autolock: Optional[str] = None
    unlock_method: Optional[str] = None
    owner_permission: Optional[str] = None

    def __post_init__(self):
        self._parse_memo()

    def _parse_memo(self):
        """メモから物件情報を抽出"""
        if not self.memo:
            return
        patterns = {
            "autolock": r"オートロック有無[：:](.+)",
            "unlock_method": r"解錠方法[：:](.+)",
            "owner_permission": r"オーナー許可[：:](.+)",
        }
        for attr, pattern in patterns.items():
            m = re.search(pattern, self.memo)
            if m:
                setattr(self, attr, m.group(1).strip())

    @classmethod
    def from_api(cls, data: dict) -> SFAProject:
        return cls(
            id=data.get("id", ""),
            name=data.get("name", ""),
            address=data.get("contents_7", ""),
            unlock_info=data.get("contents_1", ""),
            management_company=data.get("contents_4", ""),
            memo=data.get("memo", ""),
            phase_category=data.get("phase_category"),
            phase=data.get("phase"),
            survey_date=data.get("contents_6"),
            construction_date=data.get("contents_9"),
            company_id=data.get("company_id"),
            modified=data.get("modified"),
        )


class SFAClient:
    """ネクストSFA REST API クライアント"""

    def __init__(
        self,
        api_key: Optional[str] = None,
        api_token: Optional[str] = None,
    ):
        self.api_key = api_key or os.environ.get("NEXTSFA_API_KEY", "")
        self.api_token = api_token or os.environ.get("NEXTSFA_API_TOKEN", "")
        if not self.api_key or not self.api_token:
            raise ValueError(
                "NEXTSFA_API_KEY と NEXTSFA_API_TOKEN が必要です。"
                ".env に設定するか引数で渡してください。"
            )

    # ------ 低レベル API ------

    def _get(self, path: str) -> dict:
        """GET リクエストを送信し JSON を返す"""
        url = f"{BASE_URL}{path}"
        req = urllib.request.Request(
            url,
            headers={
                "Authorization": f"Bearer {self.api_token}",
                "X-API-Key": self.api_key,
                "Accept": "application/json",
            },
        )
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read().decode("utf-8"))

    def _get_all_pages(self, path: str, max_pages: int = 600) -> list[dict]:
        """ページネーションで全件取得（limit=50で高速化）"""
        all_items: list[dict] = []
        for page in range(1, max_pages + 1):
            separator = "&" if "?" in path else "?"
            data = self._get(
                f"{path}{separator}page={page}&limit={ITEMS_PER_PAGE}"
            )
            items = data.get("data", [])
            if not items:
                break
            all_items.extend(items)
        return all_items

    def _patch(self, path: str, data: dict) -> dict:
        """PATCH リクエストを送信し JSON を返す"""
        url = f"{BASE_URL}{path}"
        body = json.dumps(data).encode("utf-8")
        req = urllib.request.Request(
            url,
            data=body,
            headers={
                "Authorization": f"Bearer {self.api_token}",
                "X-API-Key": self.api_key,
                "Content-Type": "application/json",
                "Accept": "application/json",
            },
            method="PATCH",
        )
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read().decode("utf-8"))

    # ------ 案件 (Projects) ------

    def get_projects(self, max_pages: int = 500) -> list[SFAProject]:
        """全案件を取得"""
        raw = self._get_all_pages("/v1/projects", max_pages=max_pages)
        return [SFAProject.from_api(r) for r in raw]

    def get_project(self, project_id: str) -> SFAProject:
        """案件1件を取得"""
        data = self._get(f"/v1/projects/{project_id}")
        return SFAProject.from_api(data["data"])

    def find_projects(
        self,
        *,
        phase: Optional[str] = None,
        phase_category: Optional[str] = None,
        name_contains: Optional[str] = None,
        address_contains: Optional[str] = None,
        has_survey: bool = False,
        max_pages: int = 500,
    ) -> list[SFAProject]:
        """条件に合う案件をフィルタリング

        Args:
            phase: フェーズ名で絞り込み（例: "見込み", "入金済み"）
            phase_category: フェーズカテゴリで絞り込み（例: "案件決着済み"）
            name_contains: 物件名に含まれる文字列
            address_contains: 住所に含まれる文字列
            has_survey: 現調日が設定されている案件のみ
            max_pages: 最大ページ数
        """
        projects = self.get_projects(max_pages=max_pages)
        results = []
        for p in projects:
            if phase and p.phase != phase:
                continue
            if phase_category and p.phase_category != phase_category:
                continue
            if name_contains and name_contains not in p.name:
                continue
            if address_contains and address_contains not in p.address:
                continue
            if has_survey and not p.survey_date:
                continue
            results.append(p)
        return results

    def update_project_phase(self, project_id: str, phase: str) -> SFAProject:
        """案件のフェーズを更新"""
        # user_attachments は PATCH 必須フィールド
        current = self._get(f"/v1/projects/{project_id}")
        ua = current["data"].get("user_attachments", [])
        if not ua:
            ua = [{"user_id": "3373"}]  # デフォルト: 二川悠記
        result = self._patch(
            f"/v1/projects/{project_id}",
            {"phase": phase, "user_attachments": ua},
        )
        data = result["data"]
        if isinstance(data, list):
            data = data[0] if data else current["data"]
        return SFAProject.from_api(data)

    # ------ 履歴 (Histories) ------

    def get_histories(
        self,
        project_id: str | None = None,
        max_pages: int = 600,
    ) -> list[dict]:
        """履歴情報を取得

        Args:
            project_id: 案件IDで絞り込み（None=全件）
            max_pages: 最大ページ数

        Returns:
            履歴エントリのリスト。各エントリに以下のキーを含む:
            - id: 履歴ID
            - target_original_id: 案件ID
            - memo: 自由テキスト（現調データが書かれている場合あり）
            - activity_name: 活動種別
            - start_date: 対応日時
            - user_attachments: 記入者
        """
        raw = self._get_all_pages("/v1/histories", max_pages=max_pages)
        if project_id:
            raw = [
                h for h in raw
                if str(h.get("target_original_id", "")) == str(project_id)
            ]
        return raw

    def get_project_histories(self, project_id: str) -> list[dict]:
        """特定案件の履歴を取得（高速版: 全件取得せずフィルタ）

        注意: REST APIにproject_idフィルタパラメータがないため、
        全件取得→ローカルフィルタ。大量呼び出しには get_histories() で
        一括取得してからローカルフィルタを推奨。
        """
        return self.get_histories(project_id=project_id)

    # ------ 企業 (Companies) ------

    def get_companies(self, max_pages: int = 100) -> list[dict]:
        """全企業を取得"""
        return self._get_all_pages("/v1/companies", max_pages=max_pages)

    # ------ 受注 (Orders) ------

    def get_orders(self, max_pages: int = 100) -> list[dict]:
        """全受注を取得"""
        return self._get_all_pages("/v1/orders", max_pages=max_pages)


# ------ CLI テスト用 ------

def _load_env():
    """プロジェクトの .env ファイルを読み込む"""
    env_path = Path(__file__).parent.parent / ".env"
    if env_path.exists():
        with open(env_path, "r") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    key, _, value = line.partition("=")
                    os.environ[key.strip()] = value.strip()


if __name__ == "__main__":
    _load_env()
    client = SFAClient()

    print("=== ネクストSFA 接続テスト ===")
    print()

    # 最初のページだけ取得
    projects = client.get_projects(max_pages=1)
    print(f"案件数（1ページ目）: {len(projects)} 件")
    print()

    for p in projects[:5]:
        print(f"  [{p.id}] {p.name}")
        print(f"       住所: {p.address}")
        print(f"       フェーズ: {p.phase_category} → {p.phase}")
        if p.survey_date:
            print(f"       現調日: {p.survey_date}")
        print()
