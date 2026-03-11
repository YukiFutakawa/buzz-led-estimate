"""Google Apps Script ベースのフィードバック集約ストア

現場メンバーの Streamlit アプリからのフィードバックを
Google Apps Script Web App 経由で Google Sheets に一時保管し、
sync_feedback.py で Git リポジトリに同期する。

使い方:
    # Streamlit アプリから（送信）
    store = FeedbackStore.from_streamlit_secrets()
    store.submit_feedback(report_dict, comment_reading, comment_selection)

    # GitHub Actions から（取得・同期）
    store = FeedbackStore.from_env()
    records = store.get_unsynced_feedback()
    store.mark_synced([r["id"] for r in records])
"""

from __future__ import annotations

import json
import logging
from typing import Optional

import requests

logger = logging.getLogger(__name__)


class FeedbackStore:
    """Google Apps Script Web App バックエンドのフィードバックストア"""

    def __init__(self, webapp_url: str):
        self._url = webapp_url

    @classmethod
    def from_streamlit_secrets(cls) -> "FeedbackStore":
        """Streamlit Cloud の secrets から初期化"""
        import streamlit as st
        url = st.secrets["feedback"]["gas_webapp_url"]
        return cls(url)

    @classmethod
    def from_env(cls) -> "FeedbackStore":
        """環境変数から初期化（GitHub Actions 用）"""
        import os
        url = os.environ.get("GAS_WEBAPP_URL", "")
        if not url:
            raise ValueError("GAS_WEBAPP_URL を設定してください")
        return cls(url)

    # ---- フィードバック送信（Streamlit アプリから） ----

    def submit_feedback(
        self,
        report_dict: dict,
        comment_reading: str = "",
        comment_selection: str = "",
        submitter: str = "",
    ) -> str:
        """フィードバックを Google Sheets に送信

        Args:
            report_dict: FeedbackComparator が生成した比較結果 dict
            comment_reading: 現調情報の読み取りに関するコメント
            comment_selection: LED選定に関するコメント
            submitter: 送信者名（任意）

        Returns:
            feedback_id: 生成された ID (8文字)
        """
        summary = report_dict.get("summary", {})
        payload = {
            "action": "submit",
            "property_name": report_dict.get("property_name", ""),
            "summary": summary,
            "comment_reading": comment_reading,
            "comment_selection": comment_selection,
            "fixture_diffs": report_dict.get("fixture_diffs", []),
            "selection_diffs": report_dict.get("selection_diffs", []),
            "header_diffs": report_dict.get("header_diffs", []),
            "submitter": submitter,
        }

        resp = requests.post(self._url, json=payload, timeout=30)
        resp.raise_for_status()
        result = resp.json()

        if result.get("error"):
            raise RuntimeError(result["error"])

        return result.get("feedback_id", "")

    # ---- フィードバック取得（sync_feedback.py から） ----

    def get_unsynced_feedback(self) -> list[dict]:
        """未同期のフィードバックを取得"""
        resp = requests.get(
            self._url, params={"action": "get_unsynced"}, timeout=30,
        )
        resp.raise_for_status()
        result = resp.json()

        if result.get("error"):
            raise RuntimeError(result["error"])

        return result.get("records", [])

    def get_all_feedback(self) -> list[dict]:
        """全フィードバックを取得"""
        resp = requests.get(
            self._url, params={"action": "get_all"}, timeout=30,
        )
        resp.raise_for_status()
        return resp.json().get("records", [])

    def mark_synced(self, feedback_ids: list[str]) -> None:
        """指定した ID のフィードバックを同期済みにマーク"""
        if not feedback_ids:
            return

        payload = {
            "action": "mark_synced",
            "feedback_ids": feedback_ids,
        }
        resp = requests.post(self._url, json=payload, timeout=30)
        resp.raise_for_status()

    # ---- 統計（UI 表示用） ----

    def get_feedback_stats(self) -> dict:
        """フィードバック統計を取得（UI 表示用）"""
        try:
            resp = requests.get(
                self._url, params={"action": "get_stats"}, timeout=10,
            )
            resp.raise_for_status()
            return resp.json()
        except Exception as e:
            logger.warning(f"統計取得エラー: {e}")
            return {}
