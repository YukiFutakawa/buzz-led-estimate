"""SFA履歴テキスト → SurveyData 変換

SFAの履歴情報メモに自由文で記録された現調データを
Claude APIで構造化し、SurveyDataオブジェクトに変換する。
"""

from __future__ import annotations

import base64
import json
import logging
import mimetypes
import os
import re
from pathlib import Path
from typing import Optional

import yaml

from models import PropertyInfo, SurveyData
from sfa_client import SFAProject
from survey_parser import parse_survey_ocr

logger = logging.getLogger(__name__)

DEFAULT_MODEL = "claude-sonnet-4-6"
DEFAULT_MAX_TOKENS = 8192


def _load_prompt(prompt_key: str) -> str:
    """ai_prompts.yaml からプロンプトを読み込み"""
    config_path = Path(__file__).parent.parent / "config" / "ai_prompts.yaml"
    with open(config_path, "r", encoding="utf-8") as f:
        prompts = yaml.safe_load(f)
    return prompts.get(prompt_key, "")


def _encode_image(image_path: Path) -> tuple[str, str]:
    """画像をbase64エンコードし、(base64_data, media_type) を返す"""
    mime_type, _ = mimetypes.guess_type(str(image_path))
    if mime_type is None:
        suffix = image_path.suffix.lower()
        mime_map = {
            ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
            ".png": "image/png", ".gif": "image/gif",
            ".webp": "image/webp", ".bmp": "image/bmp",
        }
        mime_type = mime_map.get(suffix, "image/jpeg")
    with open(image_path, "rb") as f:
        data = base64.standard_b64encode(f.read()).decode("utf-8")
    return data, mime_type


def _extract_json(response_text: str) -> dict:
    """API応答テキストからJSONを抽出"""
    text = response_text.strip()

    # 戦略1: ```json ... ``` ブロック
    m = re.search(r"```json\s*\n?(.*?)```", text, re.DOTALL)
    if m:
        try:
            return json.loads(m.group(1).strip())
        except json.JSONDecodeError:
            pass

    # 戦略2: 最外の { ... }
    start = text.find("{")
    end = text.rfind("}")
    if start != -1 and end > start:
        try:
            return json.loads(text[start : end + 1])
        except json.JSONDecodeError:
            pass

    # 戦略3: 全体
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        raise ValueError(f"JSONを抽出できません: {text[:200]}...")



# --- 点灯時間推定マッピング ---
_DAILY_HOURS_MAP = {
    # 24h: 非常灯・階段系
    "階段": 24, "階段室": 24, "非常灯": 24, "誘導灯": 24,
    # 12h: 外部・廊下系
    "外部": 12, "外部通路": 12, "軒下": 12, "玄関前": 12,
    "エントランス": 12, "入口": 12, "廊下": 12, "通路": 12,
    "共用廊下": 12, "内廊下": 12, "駐車場": 12, "駐輪場": 12,
    "車庫": 12, "ガレージ": 12, "外周": 12,
    # 8h: 短時間系
    "庭園": 8, "屋上": 8, "塔屋": 8, "ロビー": 8, "テナント": 8,
}

_DEFAULT_DAILY_HOURS = 12  # デフォルト点灯時間


def _estimate_daily_hours(location: str, fixture_type: str = "", notes: str = "") -> int:
    """場所名から点灯時間を推定

    Args:
        location: 設置場所
        fixture_type: 器具種別
        notes: 備考（非常灯内蔵等の情報）

    Returns:
        推定点灯時間 (h/日)
    """
    # 非常灯内蔵 → 24h
    combined = f"{fixture_type} {notes}".lower()
    if "非常" in combined or "誘導" in combined:
        return 24

    # 場所名でマッチング
    loc_lower = location.strip()
    for key, hours in _DAILY_HOURS_MAP.items():
        if key in loc_lower:
            return hours

    return _DEFAULT_DAILY_HOURS


def _post_process(ocr_result: dict, memo_text: str) -> dict:
    """Claude API結果の後処理バリデーション

    - 点灯時間の場所ベース補正
    - 数量0の器具を除去
    - bulb_count → notes への統合
    - 空の器具種別を検出・警告
    - 非常灯内蔵の daily_hours 強制補正

    Args:
        ocr_result: Claude APIの出力JSON
        memo_text: 元のメモテキスト（バリデーション用）

    Returns:
        補正済みの ocr_result
    """
    for key in ("fixtures", "excluded_fixtures"):
        fixtures = ocr_result.get(key, [])
        cleaned = []
        for fix in fixtures:
            # 数量チェック: 全ての階の合計が0なら除去
            fq = fix.get("floor_quantities", {})
            total_qty = sum(int(v) for v in fq.values() if str(v).isdigit())
            if total_qty == 0 and not fix.get("is_excluded", False):
                logger.debug(f"数量0の器具を除去: {fix.get('fixture_type', '?')}")
                continue

            # bulb_count 処理: notes に統合
            bulb_count = fix.get("bulb_count", 1)
            if bulb_count and int(bulb_count) > 1:
                existing_notes = fix.get("notes", "")
                count_note = f"{bulb_count}灯用"
                if count_note not in existing_notes:
                    fix["notes"] = f"{count_note}, {existing_notes}" if existing_notes else count_note

            # 点灯時間の補正
            location = fix.get("location", "")
            fixture_type = fix.get("fixture_type", "")
            notes = fix.get("notes", "")
            estimated_hours = _estimate_daily_hours(location, fixture_type, notes)

            current_hours = fix.get("daily_hours", 0)
            if not current_hours or current_hours == 0:
                fix["daily_hours"] = estimated_hours
            # 非常灯内蔵は必ず24h
            elif "非常" in f"{fixture_type} {notes}" and current_hours < 24:
                fix["daily_hours"] = 24

            # 空の器具種別を検出
            if not fixture_type and not fix.get("is_excluded", False):
                fix["_validation_warnings"] = fix.get("_validation_warnings", [])
                fix["_validation_warnings"].append("器具種別が空です")
                fix["confidence"] = "low"

            cleaned.append(fix)

        ocr_result[key] = cleaned

    # fixtures と excluded_fixtures の重複チェック
    fixture_keys = set()
    for fix in ocr_result.get("fixtures", []):
        key_str = f"{fix.get('fixture_type', '')}|{fix.get('location', '')}|{fix.get('bulb_type', '')}"
        fixture_keys.add(key_str)

    deduped_excluded = []
    for fix in ocr_result.get("excluded_fixtures", []):
        key_str = f"{fix.get('fixture_type', '')}|{fix.get('location', '')}|{fix.get('bulb_type', '')}"
        if key_str in fixture_keys:
            logger.warning(f"fixtures と excluded_fixtures に重複: {key_str}")
            # excluded が優先（LED済みの判定を尊重）
        deduped_excluded.append(fix)

    ocr_result["excluded_fixtures"] = deduped_excluded

    # ログ出力
    n_fix = len(ocr_result.get("fixtures", []))
    n_exc = len(ocr_result.get("excluded_fixtures", []))
    logger.info(f"後処理完了: fixtures={n_fix}, excluded={n_exc}")

    return ocr_result


class HistoryTextParser:
    """SFA履歴テキストからSurveyDataを生成"""

    def __init__(self, api_key: Optional[str] = None):
        self.api_key = api_key or os.environ.get("ANTHROPIC_API_KEY", "")
        self._client = None

    def _get_client(self):
        if self._client is None:
            try:
                import anthropic
                self._client = anthropic.Anthropic(api_key=self.api_key)
            except ImportError:
                raise RuntimeError("anthropic パッケージが必要です: pip install anthropic")
        return self._client

    def parse(
        self,
        memo_texts: list[str],
        project: SFAProject,
        photo_paths: list[Path] | None = None,
    ) -> SurveyData:
        """履歴メモテキストからSurveyDataを生成

        Args:
            memo_texts: 1つ以上の履歴メモテキスト
            project: SFA案件データ（PropertyInfo生成に使用）
            photo_paths: 現場写真パスのリスト（あればマルチモーダル解析）

        Returns:
            SurveyData (fixtures + property_info)
        """
        # 1. メモテキストを結合
        combined = "\n---\n".join(t.strip() for t in memo_texts if t.strip())
        if not combined:
            raise ValueError("有効なメモテキストがありません")

        logger.info(
            f"履歴テキスト解析: [{project.id}] {project.name} "
            f"({len(memo_texts)} テキスト, {len(combined)} 文字)"
        )

        # 2. プロンプト構築
        prompt_template = _load_prompt("history_text_parse")
        if not prompt_template:
            raise ValueError("history_text_parse プロンプトが見つかりません")

        prompt = prompt_template.replace("{property_name}", project.name)
        prompt = prompt.replace("{address}", project.address or "不明")
        prompt = prompt.replace("{memo_text}", combined)

        # 3. Claude API呼び出し（写真があればマルチモーダル）
        client = self._get_client()

        if photo_paths:
            # 写真ありの場合: テキスト+画像を同時送信
            photos_desc = f"以下に {len(photo_paths)} 枚の現場写真を添付します。photo_1, photo_2, ... の順に番号付きです。"
            prompt = prompt.replace("{photos_section}", photos_desc)

            content_blocks = [{"type": "text", "text": prompt}]
            for i, p in enumerate(photo_paths, 1):
                try:
                    b64, media = _encode_image(p)
                    content_blocks.append({"type": "text", "text": f"--- photo_{i}: {p.name} ---"})
                    content_blocks.append({
                        "type": "image",
                        "source": {"type": "base64", "media_type": media, "data": b64},
                    })
                except Exception as e:
                    logger.warning(f"写真エンコード失敗 ({p.name}): {e}")

            logger.info(f"マルチモーダル解析: テキスト + 写真{len(photo_paths)}枚")
            response = client.messages.create(
                model=DEFAULT_MODEL,
                max_tokens=DEFAULT_MAX_TOKENS,
                messages=[{"role": "user", "content": content_blocks}],
            )
        else:
            # テキストのみ
            prompt = prompt.replace(
                "{photos_section}",
                "添付写真はありません。photo_refs は全て [] にしてください。",
            )
            response = client.messages.create(
                model=DEFAULT_MODEL,
                max_tokens=DEFAULT_MAX_TOKENS,
                messages=[{"role": "user", "content": prompt}],
            )

        response_text = response.content[0].text
        logger.debug(f"API応答: {response_text[:200]}...")

        # 4. JSONパース
        ocr_result = _extract_json(response_text)

        # 4.5. 後処理バリデーション
        ocr_result = _post_process(ocr_result, combined)

        # 5. ヘッダー情報をSFAデータで補完
        if "header" not in ocr_result:
            ocr_result["header"] = {}
        header = ocr_result["header"]
        if not header.get("property_name"):
            header["property_name"] = project.name
        if not header.get("address"):
            header["address"] = project.address or ""
        if not header.get("survey_date") and project.survey_date:
            header["survey_date"] = project.survey_date

        # 6. photo_refs → 実ファイルパスのマッピングを構築
        fixture_photos = {}
        if photo_paths:
            all_fixes = ocr_result.get("fixtures", []) + ocr_result.get("excluded_fixtures", [])
            for fix in all_fixes:
                refs = fix.get("photo_refs", [])
                label = fix.get("row_label", "")
                if refs and label:
                    paths = [photo_paths[r - 1] for r in refs
                             if isinstance(r, int) and 0 < r <= len(photo_paths)]
                    if paths:
                        fixture_photos[label] = paths

        # 7. parse_survey_ocr() で SurveyData に変換（コード再利用）
        survey = parse_survey_ocr(ocr_result, fixture_photos=fixture_photos)

        # 物件名をSFAから確実に設定
        survey.property_info.name = project.name
        if project.address and not survey.property_info.address:
            survey.property_info.address = project.address

        # OCRテキスト（デバッグ用）を保持
        survey.raw_ocr_text = combined

        logger.info(
            f"テキスト解析完了: 器具={len(survey.fixtures)}件, "
            f"除外={len(survey.excluded_fixtures)}件"
        )

        return survey

    def parse_dry_run(
        self,
        memo_texts: list[str],
        project: SFAProject,
    ) -> dict:
        """ドライラン: APIは呼ばず、入力データのサマリーを返す

        バッチ実行前の確認用。
        """
        combined = "\n---\n".join(t.strip() for t in memo_texts if t.strip())
        return {
            "project_id": project.id,
            "project_name": project.name,
            "memo_count": len(memo_texts),
            "total_chars": len(combined),
            "preview": combined[:300],
        }
