"""SFA案件キャッシュ

SFA APIはサーバー側フィルタリング・ソート非対応のため、
全案件(約23000件)の取得に数分かかる。
ローカルにキャッシュすることで2回目以降は即時読み込みにする。

使い方:
    cache = SFAProjectCache(sfa_client)
    projects = cache.get_projects()  # 初回: API全件取得+保存, 2回目: ファイル読込
    targets = cache.find_by_phase('数量入力待ち')  # 即時フィルタ
"""

from __future__ import annotations

import json
import logging
import time
from datetime import datetime
from pathlib import Path
from typing import Optional

from sfa_client import SFAClient, SFAProject

logger = logging.getLogger(__name__)

CACHE_DIR = Path(__file__).parent.parent / "cache"
CACHE_FILE = CACHE_DIR / "projects_cache.json"
HISTORY_CACHE_FILE = CACHE_DIR / "histories_cache.json"
DEFAULT_TTL_HOURS = 24


class SFAProjectCache:
    """SFA案件データのローカルキャッシュ"""

    def __init__(self, sfa_client: SFAClient, ttl_hours: float = DEFAULT_TTL_HOURS):
        self.sfa = sfa_client
        self.ttl_hours = ttl_hours
        self._projects: list[SFAProject] = []
        self._raw_data: list[dict] = []

    def get_projects(self, force_refresh: bool = False) -> list[SFAProject]:
        """全案件を取得（キャッシュ優先）

        Args:
            force_refresh: Trueなら強制的にAPI再取得
        """
        if not force_refresh and self._load_cache():
            return self._projects

        # キャッシュなし or 期限切れ → API全件取得
        logger.info("SFA全案件をAPI取得中（初回のみ数分かかります）...")
        self._raw_data = self._fetch_all_from_api()
        self._projects = [SFAProject.from_api(r) for r in self._raw_data]
        self._save_cache()
        logger.info(f"キャッシュ保存完了: {len(self._projects)} 件")
        return self._projects

    def find_by_phase(self, phase: str) -> list[SFAProject]:
        """フェーズ名で案件を検索（キャッシュから即時）"""
        if not self._projects:
            self.get_projects()
        return [p for p in self._projects if p.phase == phase]

    def find_projects(
        self,
        phase: Optional[str] = None,
        phase_category: Optional[str] = None,
        name_contains: Optional[str] = None,
    ) -> list[SFAProject]:
        """複合条件で案件を検索（キャッシュから即時）"""
        if not self._projects:
            self.get_projects()
        results = []
        for p in self._projects:
            if phase and p.phase != phase:
                continue
            if phase_category and p.phase_category != phase_category:
                continue
            if name_contains and name_contains not in p.name:
                continue
            results.append(p)
        return results

    def refresh(self) -> list[SFAProject]:
        """キャッシュを強制更新"""
        return self.get_projects(force_refresh=True)

    def cache_age_hours(self) -> Optional[float]:
        """キャッシュの経過時間（時間）。キャッシュなければNone"""
        if not CACHE_FILE.exists():
            return None
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                meta = json.load(f)
            cached_at = datetime.fromisoformat(meta["cached_at"])
            return (datetime.now() - cached_at).total_seconds() / 3600
        except Exception:
            return None

    # --- 内部メソッド ---

    def _load_cache(self) -> bool:
        """キャッシュファイルを読み込み。有効ならTrue"""
        if not CACHE_FILE.exists():
            logger.info("キャッシュなし → API取得が必要です")
            return False

        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                meta = json.load(f)
        except Exception as e:
            logger.warning(f"キャッシュ読み込みエラー: {e}")
            return False

        # TTLチェック
        cached_at = datetime.fromisoformat(meta["cached_at"])
        age_hours = (datetime.now() - cached_at).total_seconds() / 3600
        if age_hours > self.ttl_hours:
            logger.info(f"キャッシュ期限切れ（{age_hours:.1f}時間経過）→ API再取得")
            return False

        # データ復元
        self._raw_data = meta["data"]
        self._projects = [SFAProject.from_api(r) for r in self._raw_data]
        logger.info(
            f"キャッシュ読み込み: {len(self._projects)} 件 "
            f"({age_hours:.1f}時間前に取得)"
        )
        return True

    def _save_cache(self):
        """キャッシュファイルに保存"""
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        meta = {
            "cached_at": datetime.now().isoformat(),
            "count": len(self._raw_data),
            "data": self._raw_data,
        }
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False)
        logger.info(f"キャッシュ保存: {CACHE_FILE}")

    def _fetch_all_from_api(self) -> list[dict]:
        """APIから全ページ取得（進捗表示付き）"""
        all_items = []
        page = 1
        while True:
            try:
                data = self.sfa._get(
                    f"/v1/projects?page={page}&limit=50"
                )
            except Exception as e:
                logger.error(f"Page {page} エラー: {e}")
                break
            items = data.get("data", [])
            if not items:
                break
            all_items.extend(items)
            if page % 50 == 0:
                logger.info(f"  ... {len(all_items)} 件取得済み (page {page})")
            page += 1
        logger.info(f"API取得完了: {len(all_items)} 件 ({page - 1} ページ)")
        return all_items

    # === 履歴キャッシュ ===

    def get_histories(self, force_refresh: bool = False) -> list[dict]:
        """全履歴を取得（キャッシュ優先）"""
        if not force_refresh:
            cached = self._load_history_cache()
            if cached is not None:
                return cached

        logger.info("SFA全履歴をAPI取得中（初回のみ時間がかかります）...")
        histories = self._fetch_all_histories()
        self._save_history_cache(histories)
        logger.info(f"履歴キャッシュ保存完了: {len(histories)} 件")
        return histories

    def get_project_histories(self, project_id: str) -> list[dict]:
        """特定案件の履歴を取得（キャッシュから即時フィルタ）"""
        all_h = self.get_histories()
        return [
            h for h in all_h
            if str(h.get("target_original_id", "")) == str(project_id)
        ]

    def get_histories_by_project_ids(self, project_ids: set[str]) -> dict[str, list[dict]]:
        """複数案件の履歴をまとめて取得（キャッシュから）

        Returns:
            {project_id: [history_entries]}
        """
        all_h = self.get_histories()
        result: dict[str, list[dict]] = {pid: [] for pid in project_ids}
        for h in all_h:
            pid = str(h.get("target_original_id", ""))
            if pid in result:
                result[pid].append(h)
        return result

    def _load_history_cache(self) -> Optional[list[dict]]:
        """履歴キャッシュを読み込み"""
        if not HISTORY_CACHE_FILE.exists():
            return None
        try:
            with open(HISTORY_CACHE_FILE, "r", encoding="utf-8") as f:
                meta = json.load(f)
            cached_at = datetime.fromisoformat(meta["cached_at"])
            age_hours = (datetime.now() - cached_at).total_seconds() / 3600
            if age_hours > self.ttl_hours:
                logger.info(f"履歴キャッシュ期限切れ（{age_hours:.1f}時間経過）")
                return None
            data = meta["data"]
            logger.info(f"履歴キャッシュ読み込み: {len(data)} 件 ({age_hours:.1f}時間前)")
            return data
        except Exception as e:
            logger.warning(f"履歴キャッシュ読み込みエラー: {e}")
            return None

    def _save_history_cache(self, histories: list[dict]):
        """履歴キャッシュを保存"""
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        meta = {
            "cached_at": datetime.now().isoformat(),
            "count": len(histories),
            "data": histories,
        }
        with open(HISTORY_CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False)
        logger.info(f"履歴キャッシュ保存: {HISTORY_CACHE_FILE}")

    def _fetch_all_histories(self) -> list[dict]:
        """履歴APIから全ページ取得"""
        all_items = []
        page = 1
        while True:
            try:
                data = self.sfa._get(
                    f"/v1/histories?page={page}&limit=50"
                )
            except Exception as e:
                logger.error(f"History page {page} エラー: {e}")
                break
            items = data.get("data", [])
            if not items:
                break
            all_items.extend(items)
            if page % 100 == 0:
                logger.info(f"  ... 履歴 {len(all_items)} 件取得済み (page {page})")
            page += 1
        logger.info(f"履歴API取得完了: {len(all_items)} 件 ({page - 1} ページ)")
        return all_items

