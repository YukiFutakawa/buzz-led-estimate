"""拡張バッチ処理 - マルチソース対応

全データソース（ローカル写真・SFA ZIP・履歴テキスト）を
統合して案件を一括処理するオーケストレーター。
"""

from __future__ import annotations

import json
import logging
import os
import time
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Optional

from sfa_client import SFAClient, SFAProject
from sfa_cache import SFAProjectCache
from data_source_resolver import DataSourceResolver, DataSourceInfo

logger = logging.getLogger(__name__)

ROOT = Path(__file__).parent.parent
SURVEY_DIR = ROOT / "現調写真"
OUTPUT_DIR = ROOT / "output"
LINEUP_DIR = ROOT / "ラインナップ表"
TEMPLATE_DIR = ROOT / "見積りテンプレート"


@dataclass
class BatchResult:
    project_id: str
    project_name: str
    source_type: str
    status: str
    output_path: Optional[str] = None
    error: Optional[str] = None
    fixture_count: int = 0
    processing_time_sec: float = 0


class EnhancedBatchProcessor:

    def __init__(self, sfa_client=None, api_key=None):
        self.sfa = sfa_client or SFAClient()
        self.api_key = api_key or os.environ.get("ANTHROPIC_API_KEY", "")
        self.cache = SFAProjectCache(self.sfa)
        self.resolver = DataSourceResolver(SURVEY_DIR, self.sfa)
        self.results: list[BatchResult] = []

    def get_target_projects(self, phase="数量入力待ち", force_refresh=False):
        """キャッシュ経由で対象案件を取得（高速）"""
        if force_refresh:
            self.cache.refresh()
        return self.cache.find_by_phase(phase)

    def resolve_sources(self, projects):
        logger.info(f"=== Phase 1: データソース判定 ({len(projects)} 件) ===")
        return self.resolver.resolve_all(projects, preload_histories=True, history_cache=self.cache)

    def process_all(self, infos, template_name="田村基本形", dry_run=False):
        mode = "ドライラン" if dry_run else "本番"
        logger.info(f"=== Phase 3: 処理実行 ({mode}) ===")
        self.results = []
        for i, info in enumerate(infos, 1):
            logger.info(f"[{i}/{len(infos)}] {info.project_name} ({info.source_type})")
            if info.source_type == "none":
                self.results.append(BatchResult(
                    project_id=info.project_id, project_name=info.project_name,
                    source_type="none", status="no_data", error=info.notes))
                continue
            if dry_run:
                self.results.append(BatchResult(
                    project_id=info.project_id, project_name=info.project_name,
                    source_type=info.source_type, status="skipped", error="ドライラン"))
                continue
            start = time.time()
            try:
                result = self._process_one(info, template_name)
                result.processing_time_sec = time.time() - start
                self.results.append(result)
            except Exception as e:
                logger.error(f"処理エラー: {info.project_name}: {e}")
                self.results.append(BatchResult(
                    project_id=info.project_id, project_name=info.project_name,
                    source_type=info.source_type, status="error",
                    error=str(e), processing_time_sec=time.time() - start))
        return self.results

    def _process_one(self, info, template_name):
        if info.source_type == "local_photos":
            return self._process_local_photos(info, template_name)
        elif info.source_type == "history_text":
            return self._process_history_text(info, template_name)
        elif info.source_type == "hybrid":
            return self._process_hybrid(info, template_name)
        elif info.source_type == "sfa_zip":
            return self._process_sfa_zip(info, template_name)
        return BatchResult(project_id=info.project_id, project_name=info.project_name,
                           source_type=info.source_type, status="no_data", error="Unknown source type")

    def _process_local_photos(self, info, template_name):
        from pipeline import run_pipeline
        result_path = run_pipeline(
            survey_dir=info.local_photo_dir, lineup_dir=LINEUP_DIR,
            template_dir=TEMPLATE_DIR, template_name=template_name, api_key=self.api_key)
        return BatchResult(project_id=info.project_id, project_name=info.project_name,
                           source_type="local_photos", status="success", output_path=str(result_path))

    def _process_history_text(self, info, template_name):
        from history_text_parser import HistoryTextParser
        from pipeline import run_from_survey_data
        parser = HistoryTextParser(api_key=self.api_key)
        project = self.sfa.get_project(info.project_id)
        survey = parser.parse(info.history_memos, project)
        if not survey.fixtures:
            return BatchResult(project_id=info.project_id, project_name=info.project_name,
                               source_type="history_text", status="error",
                               error="テキストから器具データを抽出できませんでした")
        result_path = run_from_survey_data(
            survey=survey, lineup_dir=LINEUP_DIR,
            template_dir=TEMPLATE_DIR, template_name=template_name)
        return BatchResult(project_id=info.project_id, project_name=info.project_name,
                           source_type="history_text", status="success",
                           output_path=str(result_path), fixture_count=len(survey.fixtures))




    def _process_hybrid(self, info, template_name):
        """テキスト+写真のハイブリッド処理

        テキストと写真を同時にClaude APIに渡し、
        Claudeがphoto_refsで各器具に写真をマッピングする。
        """
        from history_text_parser import HistoryTextParser
        from pipeline import run_from_survey_data

        parser = HistoryTextParser(api_key=self.api_key)
        project = self.sfa.get_project(info.project_id)

        # 写真ファイル一覧を取得
        photo_files = []
        if info.history_photo_dir and info.history_photo_dir.exists():
            photo_files = sorted([
                f for f in info.history_photo_dir.rglob("*")
                if f.is_file() and f.suffix.lower() in {".jpg", ".jpeg", ".png", ".gif", ".webp"}
            ])

        # テキスト＋写真を一括解析（マルチモーダル）
        survey = parser.parse(
            info.history_memos, project,
            photo_paths=photo_files if photo_files else None,
        )
        if not survey.fixtures:
            return BatchResult(
                project_id=info.project_id, project_name=info.project_name,
                source_type="hybrid", status="error",
                error="テキストから器具データを抽出できませんでした")

        result_path = run_from_survey_data(
            survey=survey, lineup_dir=LINEUP_DIR,
            template_dir=TEMPLATE_DIR, template_name=template_name)

        return BatchResult(
            project_id=info.project_id, project_name=info.project_name,
            source_type="hybrid", status="success",
            output_path=str(result_path), fixture_count=len(survey.fixtures))

    def _process_sfa_zip(self, info, template_name):
        """SFA ZIPから写真パイプライン実行"""
        project_dir = SURVEY_DIR / info.project_name
        if project_dir.exists() and any(project_dir.rglob('*.jpg')):
            from pipeline import run_pipeline
            result_path = run_pipeline(
                survey_dir=project_dir, lineup_dir=LINEUP_DIR,
                template_dir=TEMPLATE_DIR, template_name=template_name, api_key=self.api_key)
            return BatchResult(project_id=info.project_id, project_name=info.project_name,
                               source_type='sfa_zip', status='success', output_path=str(result_path))
        return BatchResult(project_id=info.project_id, project_name=info.project_name,
                           source_type='sfa_zip', status='error',
                           error='ZIPのダウンロード・展開が必要です')

    def generate_report(self):
        total = len(self.results)
        by_status = {}
        by_source = {}
        total_time = 0.0
        for r in self.results:
            by_status[r.status] = by_status.get(r.status, 0) + 1
            by_source[r.source_type] = by_source.get(r.source_type, 0) + 1
            total_time += r.processing_time_sec
        lines = ["=" * 60, "  LED見積バッチ処理結果レポート", "=" * 60,
                 f"  合計: {total} 件  処理時間: {total_time:.1f}秒", "",
                 "  --- ステータス別 ---"]
        sl = {"success": "成功", "error": "エラー", "skipped": "スキップ", "no_data": "データなし"}
        for status, count in sorted(by_status.items()):
            lines.append(f"    {sl.get(status, status)}: {count} 件")
        lines.extend(["", "  --- データソース別 ---"])
        src_l = {"local_photos": "ローカル写真", "history_text": "履歴テキスト", "hybrid": "ハイブリッド", "sfa_zip": "SFA ZIP", "none": "データなし"}
        for source, count in sorted(by_source.items()):
            lines.append(f"    {src_l.get(source, source)}: {count} 件")
        successes = [r for r in self.results if r.status == "success"]
        if successes:
            lines.extend(["", "  --- 成功 ---"])
            for r in successes:
                lines.append(f"    [{r.project_id}] {r.project_name} ({src_l.get(r.source_type, r.source_type)})")
        errors = [r for r in self.results if r.status == "error"]
        if errors:
            lines.extend(["", "  --- エラー詳細 ---"])
            for r in errors:
                lines.append(f"    [{r.project_id}] {r.project_name}: {r.error}")
        no_data = [r for r in self.results if r.status == "no_data"]
        if no_data:
            lines.extend(["", "  --- データなし ---"])
            for r in no_data:
                lines.append(f"    [{r.project_id}] {r.project_name}")
        lines.append("=" * 60)
        return "\n".join(lines)

    def save_results(self, path=None):
        if path is None:
            path = OUTPUT_DIR / "batch_results.json"
        path.parent.mkdir(parents=True, exist_ok=True)
        data = [asdict(r) for r in self.results]
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        logger.info(f"結果保存: {path}")
        return path

    def send_report_email(self, recipient=None):
        """処理結果レポートをメール送信"""
        from report_mailer import send_report_email, DEFAULT_RECIPIENT
        report_text = self.generate_report()
        to = recipient or DEFAULT_RECIPIENT
        ok = send_report_email(
            report_text=report_text,
            recipient=to,
            results=self.results,
        )
        if ok:
            logger.info(f"レポートメール送信完了: {to}")
        else:
            logger.warning("レポートメール送信失敗（SMTP設定を確認してください）")
        return ok

