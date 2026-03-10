"""データソース判定モジュール

各SFA案件に対して、最適なデータソースを判定する。
判定: 履歴テキストあり＝処理対象
処理優先度: ローカル写真 > SFA ZIP > 履歴テキスト > なし
"""

from __future__ import annotations

import logging
import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from sfa_client import SFAClient, SFAProject

logger = logging.getLogger(__name__)

# 器具データを示すキーワード
FIXTURE_KEYWORDS = [
    "蛍光灯", "ダウンライト", "ブラケット", "シーリング", "ポーチ",
    "FL20", "FL40", "FL10", "FCL", "FDL", "FHF", "FPL",
    "蛍光", "白熱", "水銀灯", "ハロゲン",
    "灯", "台", "本", "個",
    "20W", "40W", "60W", "100W",
    "20形", "40形", "32形",
    "階段", "玄関", "廊下", "駐車場", "エントランス", "外部",
    "共用部", "通路", "ホール",
    "点灯", "消費電力",
    "LED済", "LED化済", "交換済",
    "防水", "非常灯", "誘導灯",
    "投光器", "スポット", "ポール灯",
]

SURVEY_FILE_KEYWORDS = ["施工前", "現調", "写真", "調査", "survey"]
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".heic"}


@dataclass
class SFAFileInfo:
    """SFAファイルタブの1ファイル"""
    file_key: str
    filename: str
    project_id: str
    project_name: str

    @property
    def ext(self) -> str:
        return Path(self.filename).suffix.lower()

    @property
    def is_zip(self) -> bool:
        return self.ext == ".zip"

    @property
    def is_survey_photos(self) -> bool:
        name_lower = self.filename.lower()
        if not self.is_zip:
            return False
        return any(kw in name_lower for kw in SURVEY_FILE_KEYWORDS)


@dataclass
class DataSourceInfo:
    """案件のデータソース解決結果"""
    project_id: str
    project_name: str
    source_type: str  # "local_photos" | "history_text" | "sfa_zip" | "hybrid" | "none"
    local_photo_dir: Optional[Path] = None
    history_photo_dir: Optional[Path] = None  # 履歴添付写真のDL先
    sfa_files: list[SFAFileInfo] = field(default_factory=list)
    history_memos: list[str] = field(default_factory=list)
    confidence: str = "high"
    notes: str = ""

    def upgrade_to_hybrid(self, photo_dir: Path) -> None:
        """履歴テキスト案件を写真DL後にhybridにアップグレード"""
        if self.source_type == "history_text":
            self.source_type = "hybrid"
            self.history_photo_dir = photo_dir
            n_photos = sum(1 for f in photo_dir.rglob("*") if f.is_file()) if photo_dir.exists() else 0
            self.notes = f"hybrid (テキスト+写真{n_photos}枚)"
            self.confidence = "high"


def _normalize(text: str) -> str:
    """テキストを正規化（全角→半角、空白除去、カタカナ統一）"""
    text = unicodedata.normalize("NFKC", text)
    text = re.sub(r"[\s　・\-－―–_＿※★☆◎○●▲△▼▽■□◆◇]", "", text)
    text = re.sub(r"[()（）\[\]【】「」『』]", "", text)
    return text.lower()


def _has_images(folder: Path) -> bool:
    if not folder.exists():
        return False
    return any(
        f.is_file() and f.suffix.lower() in IMAGE_EXTS
        for f in folder.rglob("*")
    )


def _count_images(folder: Path) -> int:
    if not folder.exists():
        return 0
    return sum(
        1 for f in folder.rglob("*")
        if f.is_file() and f.suffix.lower() in IMAGE_EXTS
    )


class DataSourceResolver:
    """各案件の最適データソースを判定"""

    def __init__(
        self,
        survey_photo_root: Path,
        sfa_client: SFAClient,
        all_histories: list[dict] | None = None,
    ):
        self.survey_root = survey_photo_root
        self.sfa = sfa_client
        self._local_folders = self._scan_local_folders()
        self._all_histories = all_histories

    def _scan_local_folders(self) -> dict[str, Path]:
        index: dict[str, Path] = {}
        if not self.survey_root.exists():
            return index
        for d in self.survey_root.iterdir():
            if d.is_dir():
                norm = _normalize(d.name)
                index[norm] = d
        logger.debug(f"ローカルフォルダ: {len(index)} 件")
        return index

    def _match_local_folder(self, project_name: str) -> Optional[Path]:
        """SFA案件名 → ローカルフォルダのマッチング"""
        norm = _normalize(project_name)
        if not norm:
            return None

        # 1. 完全一致
        if norm in self._local_folders:
            return self._local_folders[norm]

        # 2. 部分一致（双方向）
        for folder_norm, folder_path in self._local_folders.items():
            if norm in folder_norm or folder_norm in norm:
                return folder_path

        # 3. 先頭3文字一致
        if len(norm) >= 3:
            prefix = norm[:3]
            for folder_norm, folder_path in self._local_folders.items():
                if len(folder_norm) >= 3 and folder_norm[:3] == prefix:
                    return folder_path

        return None

    def _get_project_histories(self, project_id: str) -> list[dict]:
        if self._all_histories is not None:
            return [
                h for h in self._all_histories
                if str(h.get("target_original_id", "")) == str(project_id)
            ]
        return self.sfa.get_project_histories(project_id)

    @staticmethod
    def looks_like_fixture_data(memo: str) -> bool:
        """メモテキストが器具データを含むかのヒューリスティック判定"""
        if not memo or len(memo.strip()) < 10:
            return False
        hit_count = sum(1 for kw in FIXTURE_KEYWORDS if kw in memo)
        return hit_count >= 3

    def resolve(self, project: SFAProject) -> DataSourceInfo:
        """単一案件のデータソースを判定

        判定基準: 履歴テキストがある＝処理対象とする
        処理優先度: ローカル写真 > SFA ZIP > 履歴テキスト

        SFA ZIPは後からupdate_sfa_files()でアップグレード可能。
        """
        # まず履歴テキストを取得（処理対象かの判定に使用）
        histories = self._get_project_histories(project.id)
        fixture_memos = []
        for h in histories:
            memo = h.get("memo", "")
            if self.looks_like_fixture_data(memo):
                fixture_memos.append(memo)

        # 1. ローカル写真チェック（最優先の処理ソース）
        local_dir = self._match_local_folder(project.name)
        if local_dir and _has_images(local_dir):
            img_count = _count_images(local_dir)
            return DataSourceInfo(
                project_id=project.id,
                project_name=project.name,
                source_type="local_photos",
                local_photo_dir=local_dir,
                history_memos=fixture_memos,
                notes=f"ローカル写真 {img_count} 枚",
            )

        # 2. 履歴テキストがある＝処理対象
        #    source_typeは一旦 "history_text" にセットし、
        #    後からZIPが見つかれば "sfa_zip" にアップグレードされる
        if fixture_memos:
            return DataSourceInfo(
                project_id=project.id,
                project_name=project.name,
                source_type="history_text",
                history_memos=fixture_memos,
                confidence="medium",
                notes=f"履歴テキスト {len(fixture_memos)} 件",
            )

        # 3. データなし
        return DataSourceInfo(
            project_id=project.id,
            project_name=project.name,
            source_type="none",
            notes="データソースなし",
        )

    def resolve_all(
        self,
        projects: list[SFAProject],
        preload_histories: bool = True,
        history_cache=None,
    ) -> list[DataSourceInfo]:
        """複数案件を一括判定

        Args:
            history_cache: SFAProjectCache のインスタンス（キャッシュ経由で履歴取得）
        """
        if preload_histories and self._all_histories is None:
            if history_cache is not None:
                logger.info("SFA履歴をキャッシュから取得中...")
                self._all_histories = history_cache.get_histories()
            else:
                logger.info("SFA履歴をAPI直接取得中...")
                self._all_histories = self.sfa.get_histories()
            logger.info(f"履歴取得完了: {len(self._all_histories)} 件")

        results = []
        for p in projects:
            info = self.resolve(p)
            results.append(info)
            logger.debug(
                f"[{p.id}] {p.name} -> {info.source_type} ({info.notes})"
            )

        by_type: dict[str, int] = {}
        for r in results:
            by_type[r.source_type] = by_type.get(r.source_type, 0) + 1
        logger.info(
            "データソース判定完了: "
            + ", ".join(f"{k}={v}" for k, v in sorted(by_type.items()))
        )

        return results

    def update_sfa_files(
        self,
        infos: list[DataSourceInfo],
        sfa_files: dict[str, list[SFAFileInfo]],
    ) -> None:
        """SFAファイルスキャン結果を反映

        ブラウザスキャン後、ZIPが見つかった案件を "sfa_zip" にアップグレード。
        対象: "none" と "history_text" の案件
        （ZIPはテキストより処理優先度が高いため、テキスト判定済みでもZIPに切替）
        """
        for info in infos:
            # ローカル写真はそのまま（最優先）
            if info.source_type == "local_photos":
                continue
            files = sfa_files.get(info.project_id, [])
            survey_zips = [f for f in files if f.is_survey_photos]
            if survey_zips:
                info.source_type = "sfa_zip"
                info.sfa_files = survey_zips
                info.notes = f"SFA ZIP {len(survey_zips)} 件"
            elif files:
                info.sfa_files = files
                if info.source_type == "none":
                    info.notes = f"SFAファイル {len(files)} 件（非写真ZIP）"

