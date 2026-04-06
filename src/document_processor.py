"""AI Vision による現調チェックシートOCR処理

チェックシートの写真画像からAI Vision APIで構造化データを抽出する。
対応API: Anthropic Claude (claude-sonnet-4-6)

改善版:
- 画像前処理（コントラスト強調・シャープ化・最適リサイズ）
- 2パス検証（フルOCR + 数値特化検証）
- OCR後バリデーション（ファジーマッチ・範囲チェック）
- JSON抽出の堅牢化
"""

from __future__ import annotations

import base64
import json
import logging
import mimetypes
import re
from pathlib import Path
from typing import Optional

import yaml

logger = logging.getLogger(__name__)

# デフォルトの設定
DEFAULT_MODEL = "claude-sonnet-4-6"
DEFAULT_MAX_TOKENS = 8192


def _load_prompt(prompt_key: str, config_path: Optional[Path] = None) -> str:
    """ai_prompts.yaml からプロンプトを読み込み"""
    if config_path is None:
        config_path = (
            Path(__file__).parent.parent / "config" / "ai_prompts.yaml"
        )
    with open(config_path, "r", encoding="utf-8") as f:
        prompts = yaml.safe_load(f)
    return prompts.get(prompt_key, "")


def _encode_image(image_path: Path) -> tuple[str, str]:
    """画像をbase64エンコードし、(base64_data, media_type) を返す"""
    mime_type, _ = mimetypes.guess_type(str(image_path))
    if mime_type is None:
        suffix = image_path.suffix.lower()
        mime_map = {
            ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg",
            ".png": "image/png",
            ".gif": "image/gif",
            ".webp": "image/webp",
            ".bmp": "image/bmp",
        }
        mime_type = mime_map.get(suffix, "image/jpeg")

    with open(image_path, "rb") as f:
        data = base64.standard_b64encode(f.read()).decode("utf-8")
    return data, mime_type


def _encode_image_preprocessed(image_path: Path) -> tuple[str, str]:
    """画像を前処理してからbase64エンコード"""
    from image_preprocessor import CheckSheetPreprocessor

    preprocessor = CheckSheetPreprocessor()
    img_bytes, media_type = preprocessor.preprocess(image_path)
    data = base64.standard_b64encode(img_bytes).decode("utf-8")
    return data, media_type


class DocumentProcessor:
    """AI Vision APIを使った文書処理エンジン

    使用方法:
        processor = DocumentProcessor(api_key="sk-ant-...")
        result = processor.ocr_survey_sheet("path/to/checksheet.jpg")
    """

    def __init__(
        self,
        api_key: Optional[str] = None,
        model: str = DEFAULT_MODEL,
        max_tokens: int = DEFAULT_MAX_TOKENS,
    ):
        self.model = model
        self.max_tokens = max_tokens
        self._client = None

        # APIキーの取得（引数 > 環境変数）
        if api_key is None:
            import os
            api_key = os.environ.get("ANTHROPIC_API_KEY")

        if api_key:
            self._init_client(api_key)
        else:
            logger.warning(
                "APIキーが未設定です。"
                "ANTHROPIC_API_KEY 環境変数を設定するか、"
                "api_key引数で渡してください。"
            )

    def _init_client(self, api_key: str) -> None:
        """Anthropic クライアントを初期化（claude_guard 経由）"""
        try:
            import sys
            from pathlib import Path
            # claude_guard.py は D:\Buzzarea\ にある
            # __file__ = src/document_processor.py → 4階層上
            buzzarea_dir = Path(__file__).resolve().parent.parent.parent.parent
            sys.path.insert(0, str(buzzarea_dir))
            from claude_guard import get_guarded_client
            self._client = get_guarded_client(api_key=api_key)
            logger.info(f"Anthropic API 初期化完了 (model={self.model}) [claude_guard]")
        except ImportError:
            logger.error(
                "anthropic または claude_guard が見つかりません。"
                "pip install anthropic を実行してください。"
            )
        except Exception as e:
            logger.error(f"Anthropic API 初期化エラー: {e}")

    @property
    def is_ready(self) -> bool:
        """APIが使用可能か"""
        return self._client is not None

    def ocr_survey_sheet(
        self,
        image_path: Path | str,
        custom_prompt: Optional[str] = None,
        verify: bool = True,
        preprocess: bool = True,
    ) -> dict:
        """現調チェックシートのOCR処理

        Args:
            image_path: チェックシート画像のパス
            custom_prompt: カスタムプロンプト（省略時はデフォルト使用）
            verify: 2パス検証を行うか（デフォルト: True）
            preprocess: 画像前処理を行うか（デフォルト: True）

        Returns:
            構造化されたOCR結果（dict）

        Raises:
            RuntimeError: APIが初期化されていない場合
            FileNotFoundError: 画像ファイルが見つからない場合
        """
        if not self.is_ready:
            raise RuntimeError(
                "API未初期化。APIキーを設定してください。"
            )

        image_path = Path(image_path)
        if not image_path.exists():
            raise FileNotFoundError(f"画像が見つかりません: {image_path}")

        # 画像エンコード（前処理あり/なし）
        if preprocess:
            try:
                b64_data, media_type = _encode_image_preprocessed(image_path)
                logger.info(f"画像前処理完了: {image_path.name}")
            except Exception as e:
                logger.warning(f"画像前処理失敗、生画像を使用: {e}")
                b64_data, media_type = _encode_image(image_path)
        else:
            b64_data, media_type = _encode_image(image_path)

        # Pass 1: フルOCR
        prompt = custom_prompt or _load_prompt("survey_sheet_ocr")
        logger.info(f"OCR処理開始 (Pass 1): {image_path.name}")
        response = self._call_vision_api(b64_data, media_type, prompt)
        result = self._extract_json(response)

        # OCR後バリデーション
        try:
            from ocr_validator import OCRValidator
            validator = OCRValidator()
            result = validator.validate_and_fix(result)
        except Exception as e:
            logger.warning(f"バリデーションスキップ: {e}")

        # Pass 2: 数値検証（オプション）
        if verify and result.get("fixtures"):
            try:
                result = self._verify_pass(
                    b64_data, media_type, result,
                )
            except Exception as e:
                logger.warning(f"検証パススキップ: {e}")

        fixture_count = len(result.get("fixtures", []))
        logger.info(f"OCR完了: {fixture_count}件の器具データ検出")
        return result

    def _verify_pass(
        self,
        b64_data: str,
        media_type: str,
        full_result: dict,
    ) -> dict:
        """2パス目: 数値に特化した検証OCR"""
        verify_prompt = _load_prompt("survey_sheet_verify")
        if not verify_prompt:
            return full_result

        logger.info("OCR検証 (Pass 2): 数値データの照合")
        response = self._call_vision_api(
            b64_data, media_type, verify_prompt,
        )
        verify_result = self._extract_json(response)

        return self._reconcile_passes(full_result, verify_result)

    def _reconcile_passes(
        self, full_result: dict, verify_result: dict,
    ) -> dict:
        """2回のOCR結果を照合し、矛盾があれば修正"""
        verify_rows = verify_result.get("rows", [])
        if not verify_rows:
            return full_result

        # 検証結果をrow_labelでルックアップ
        verify_map = {}
        for row in verify_rows:
            label = row.get("row_label", "")
            if label:
                verify_map[label] = row

        corrections = 0
        for fixture in full_result.get("fixtures", []):
            label = fixture.get("row_label", "")
            if label not in verify_map:
                continue

            v = verify_map[label]

            # 消費電力の照合
            f_power = fixture.get("power_w", 0)
            v_power = v.get("power_w", 0)
            if f_power != v_power and v_power != 0:
                logger.debug(
                    f"行{label} 消費電力修正: {f_power} → {v_power}"
                )
                fixture["power_w"] = v_power
                corrections += 1

            # 点灯時間の照合
            f_hours = fixture.get("daily_hours", 0)
            v_hours = v.get("daily_hours", 0)
            if f_hours != v_hours and v_hours != 0:
                logger.debug(
                    f"行{label} 点灯時間修正: {f_hours} → {v_hours}"
                )
                fixture["daily_hours"] = v_hours
                corrections += 1

            # 階別数量の照合
            f_qty = fixture.get("floor_quantities", {})
            v_qty = v.get("floor_quantities", {})
            if f_qty != v_qty and v_qty:
                logger.debug(
                    f"行{label} 数量修正: {f_qty} → {v_qty}"
                )
                fixture["floor_quantities"] = v_qty
                corrections += 1

            # 電球種別の照合
            f_bulb = fixture.get("bulb_type", "")
            v_bulb = v.get("bulb_type", "")
            if f_bulb != v_bulb and v_bulb:
                logger.debug(
                    f"行{label} 電球種別修正: {f_bulb} → {v_bulb}"
                )
                fixture["bulb_type"] = v_bulb
                corrections += 1

        # Pass 1にない行が検証で見つかった場合
        full_labels = {
            f.get("row_label") for f in full_result.get("fixtures", [])
        }
        missing = [
            label for label in verify_map
            if label not in full_labels
        ]
        if missing:
            full_result.setdefault("_missing_rows", []).extend(missing)
            logger.warning(
                f"検証で追加の行を検出: {missing}"
            )

        if corrections > 0:
            logger.info(f"2パス照合: {corrections}件のデータを修正")

        return full_result

    def ocr_survey_sheets(
        self, image_paths: list[Path | str],
        verify: bool = True,
        preprocess: bool = True,
    ) -> dict:
        """複数ページのチェックシートをまとめてOCR

        複数枚の場合は結果をマージする。

        Args:
            image_paths: チェックシート画像パスのリスト
            verify: 2パス検証を行うか
            preprocess: 画像前処理を行うか

        Returns:
            マージされたOCR結果
        """
        if not image_paths:
            return {"header": {}, "fixtures": [], "special_notes": ""}

        # 1枚目のOCR
        merged = self.ocr_survey_sheet(
            image_paths[0], verify=verify, preprocess=preprocess,
        )
        # ページ番号を付与（fixtures + excluded_fixtures 両方）
        for fix in merged.get("fixtures", []):
            fix["_page"] = 1
        for fix in merged.get("excluded_fixtures", []):
            fix["_page"] = 1

        # 2枚目以降があればマージ
        for page_idx, path in enumerate(image_paths[1:], start=2):
            try:
                additional = self.ocr_survey_sheet(
                    path, verify=verify, preprocess=preprocess,
                )
                # ページ番号を付与してから追加（fixtures + excluded_fixtures 両方）
                for fix in additional.get("fixtures", []):
                    fix["_page"] = page_idx
                for fix in additional.get("excluded_fixtures", []):
                    fix["_page"] = page_idx
                merged.setdefault("fixtures", []).extend(
                    additional.get("fixtures", [])
                )
                merged.setdefault("excluded_fixtures", []).extend(
                    additional.get("excluded_fixtures", [])
                )
                # 備考を追加
                extra_notes = additional.get("special_notes", "")
                if extra_notes:
                    existing_notes = merged.get("special_notes", "")
                    merged["special_notes"] = (
                        f"{existing_notes}\n{extra_notes}".strip()
                    )
            except Exception as e:
                logger.error(f"追加ページOCRエラー ({path}): {e}")

        return merged

    def analyze_fixture_photo(
        self, image_path: Path | str,
    ) -> dict:
        """器具写真の分析（種類、サイズ、メーカー等を推定）

        Args:
            image_path: 器具写真のパス

        Returns:
            器具情報のdict
        """
        if not self.is_ready:
            raise RuntimeError("API未初期化")

        image_path = Path(image_path)
        if not image_path.exists():
            raise FileNotFoundError(f"画像が見つかりません: {image_path}")

        prompt = _load_prompt("fixture_photo_analysis")
        # 器具写真は前処理不要（チェックシートとは異なる）
        b64_data, media_type = _encode_image(image_path)

        response = self._call_vision_api(b64_data, media_type, prompt)
        return self._extract_json(response)

    def _call_vision_api(
        self, b64_data: str, media_type: str, prompt: str,
        model_override: Optional[str] = None,
    ) -> str:
        """Claude Vision API 呼び出し"""
        model = model_override or self.model
        message = self._client.messages.create(
            model=model,
            max_tokens=self.max_tokens,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": media_type,
                                "data": b64_data,
                            },
                        },
                        {
                            "type": "text",
                            "text": prompt,
                        },
                    ],
                },
            ],
        )

        # テキスト応答を抽出
        text_parts = [
            block.text for block in message.content
            if hasattr(block, "text")
        ]
        return "\n".join(text_parts)

    def _extract_json(self, response_text: str) -> dict:
        """API応答テキストからJSONを抽出（堅牢版）

        複数の抽出戦略を試行し、JSONパースエラーへの耐性を高める。
        """
        text = response_text.strip()

        # 戦略1: ```json ... ``` ブロックの抽出
        json_block = self._extract_code_block(text)
        if json_block:
            try:
                return json.loads(json_block)
            except json.JSONDecodeError:
                # コードブロック内のJSONが壊れている場合、修復を試みる
                fixed = self._fix_json_errors(json_block)
                if fixed is not None:
                    return fixed

        # 戦略2: 最外の { ... } を抽出
        brace_json = self._extract_outermost_braces(text)
        if brace_json is not None:
            return brace_json

        # 戦略3: テキスト全体をパース
        try:
            return json.loads(text)
        except json.JSONDecodeError:
            pass

        # 戦略4: JSON修復を試みる
        fixed = self._fix_json_errors(text)
        if fixed is not None:
            return fixed

        # 全戦略失敗
        logger.error("JSON抽出失敗: 全戦略が失敗")
        logger.debug(f"Response text (先頭500文字): {text[:500]}")
        return {"raw_text": response_text, "parse_error": "全抽出戦略が失敗"}

    def _extract_code_block(self, text: str) -> Optional[str]:
        """```json ... ``` または ``` ... ``` からJSONテキストを抽出"""
        # ```json ブロック
        if "```json" in text:
            try:
                start = text.index("```json") + 7
                end = text.index("```", start)
                return text[start:end].strip()
            except ValueError:
                pass

        # ``` ブロック（言語指定なし）
        if "```" in text:
            try:
                start = text.index("```") + 3
                end = text.index("```", start)
                return text[start:end].strip()
            except ValueError:
                pass

        return None

    def _extract_outermost_braces(self, text: str) -> Optional[dict]:
        """最外の { ... } ペアを見つけてパース"""
        brace_start = text.find("{")
        if brace_start < 0:
            return None

        # 対応する閉じ括弧を見つける（ネスト対応）
        depth = 0
        in_string = False
        escape_next = False

        for i in range(brace_start, len(text)):
            ch = text[i]

            if escape_next:
                escape_next = False
                continue

            if ch == "\\":
                escape_next = True
                continue

            if ch == '"' and not escape_next:
                in_string = not in_string
                continue

            if in_string:
                continue

            if ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    json_text = text[brace_start:i + 1]
                    try:
                        return json.loads(json_text)
                    except json.JSONDecodeError:
                        # 修復を試みる
                        fixed = self._fix_json_errors(json_text)
                        if fixed is not None:
                            return fixed
                    return None

        return None

    # ===== 写真自動分類・マッチング =====

    def _call_vision_api_multi(
        self,
        images: list[tuple[str, str]],
        prompt: str,
        model_override: Optional[str] = None,
    ) -> str:
        """複数画像を含むメッセージを一括送信

        Claude APIはcontent配列に複数のimageブロックを含められる。
        画像N枚 + テキスト1個のメッセージを1回のAPI呼出しで送信する。

        Args:
            images: [(base64_data, media_type), ...] のリスト
            prompt: プロンプトテキスト
            model_override: モデル名の上書き

        Returns:
            APIレスポンステキスト
        """
        model = model_override or self.model
        content = []

        for b64_data, media_type in images:
            content.append({
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": media_type,
                    "data": b64_data,
                },
            })

        content.append({
            "type": "text",
            "text": prompt,
        })

        message = self._client.messages.create(
            model=model,
            max_tokens=self.max_tokens,
            messages=[
                {
                    "role": "user",
                    "content": content,
                },
            ],
        )

        text_parts = [
            block.text for block in message.content
            if hasattr(block, "text")
        ]
        return "\n".join(text_parts)

    def match_photos_to_rows(
        self,
        fixture_photos: list[Path],
        ocr_fixtures: list[dict],
    ) -> dict[str, list[Path]]:
        """器具写真をOCR行ラベルに自動紐付け

        サムネイル(768px)を生成し、OCR結果の器具リスト情報とともに
        Claude Visionに送信して、各写真と行ラベルの対応をAIが判定する。

        Args:
            fixture_photos: 器具写真パスのリスト
            ocr_fixtures: OCR結果のfixtures配列
                [{"row_label": "A", "location": "玄関", "fixture_type": "天井ブラケット", ...}, ...]

        Returns:
            行ラベル→写真パスリストのdict
            {"A": [Path("photo1.jpg")], "B": [Path("photo2.jpg"), Path("photo3.jpg")]}
        """
        if not self.is_ready:
            raise RuntimeError("API未初期化")

        if not fixture_photos or not ocr_fixtures:
            return {}

        from image_preprocessor import create_thumbnail

        logger.info(
            f"写真マッチング開始: {len(fixture_photos)}枚 → "
            f"{len(ocr_fixtures)}行"
        )

        # サムネイル生成（マッチングは少し大きめ768px）
        thumbnails = []
        for path in fixture_photos:
            try:
                b64_bytes, media_type = create_thumbnail(
                    path, max_long_edge=768, jpeg_quality=75,
                )
                b64_data = base64.standard_b64encode(b64_bytes).decode("utf-8")
                thumbnails.append((b64_data, media_type))
            except Exception as e:
                logger.warning(f"サムネイル生成失敗: {path.name}: {e}")
                b64_data, media_type = _encode_image(path)
                thumbnails.append((b64_data, media_type))

        # 器具リストをテキスト化（全バッチ共通）
        fixture_list_text = self._format_fixture_list(ocr_fixtures)
        prompt_template = _load_prompt("photo_row_matching")
        valid_labels = {f.get("row_label", "") for f in ocr_fixtures}

        # バッチ分割（20枚ずつ）
        BATCH_SIZE = 20
        photo_map: dict[str, list[Path]] = {}

        num_batches = (len(thumbnails) + BATCH_SIZE - 1) // BATCH_SIZE
        logger.info(f"写真マッチング: {len(thumbnails)}枚を{num_batches}バッチで処理")

        for batch_idx in range(num_batches):
            start = batch_idx * BATCH_SIZE
            end = min(start + BATCH_SIZE, len(thumbnails))
            batch_thumbs = thumbnails[start:end]
            batch_photos = fixture_photos[start:end]

            logger.info(
                f"  バッチ {batch_idx+1}/{num_batches}: "
                f"写真{start}〜{end-1} ({len(batch_thumbs)}枚)"
            )

            # ファイル名リスト（バッチ内のインデックスで生成）
            filename_list = "\n".join(
                f"  写真{i}: {batch_photos[i].name}"
                for i in range(len(batch_photos))
            )

            # プロンプト構築
            prompt = prompt_template.replace("{fixture_list}", fixture_list_text)
            if "{filename_list}" in prompt:
                prompt = prompt.replace("{filename_list}", filename_list)
            else:
                prompt += (
                    f"\n\n【参考】各写真のファイル名:\n{filename_list}\n"
                    "ファイル名に場所や器具の情報が含まれる場合は、マッチングの参考にしてください。"
                )

            # API呼出し（バッチ単位）
            try:
                response = self._call_vision_api_multi(batch_thumbs, prompt)
                result = self._extract_json_array(response)
            except Exception as e:
                logger.warning(f"  バッチ{batch_idx+1}のAPI呼び出し失敗: {e}")
                continue

            # 結果を統合（photo_indexはバッチ内インデックス→全体インデックスに変換）
            for item in result:
                photo_idx_in_batch = item.get("photo_index", -1)
                row_label = item.get("row_label", "")
                confidence = item.get("confidence", "low")

                if photo_idx_in_batch < 0 or photo_idx_in_batch >= len(batch_photos):
                    continue
                if row_label == "unmatched" or row_label not in valid_labels:
                    logger.info(
                        f"  未マッチ: {batch_photos[photo_idx_in_batch].name} "
                        f"(row={row_label}, conf={confidence})"
                    )
                    continue

                photo_map.setdefault(row_label, []).append(
                    batch_photos[photo_idx_in_batch]
                )
                logger.info(
                    f"  マッチ: {batch_photos[photo_idx_in_batch].name} → "
                    f"行{row_label} ({confidence})"
                )

        logger.info(
            f"写真マッチング完了: {sum(len(v) for v in photo_map.values())}枚 → "
            f"{len(photo_map)}行"
        )
        return photo_map

    @staticmethod
    def _format_fixture_list(ocr_fixtures: list[dict]) -> str:
        """OCR器具リストを人間が読めるテキスト形式に変換"""
        lines = []
        for f in ocr_fixtures:
            label = f.get("row_label", "?")
            location = f.get("location", "")
            f_type = f.get("fixture_type", "")
            f_size = f.get("fixture_size", "")
            bulb = f.get("bulb_type", "")

            # 数量の合計
            qty = sum(
                v for v in f.get("floor_quantities", {}).values()
                if isinstance(v, (int, float))
            )

            parts = [f"行{label}:"]
            if location:
                parts.append(f"場所={location}")
            if f_type:
                parts.append(f"種別={f_type}")
            if f_size:
                parts.append(f"サイズ={f_size}")
            if bulb:
                parts.append(f"電球={bulb}")
            if qty > 0:
                parts.append(f"数量={qty}")

            lines.append(" | ".join(parts))

        return "\n".join(lines)

    def _extract_json_array(self, response_text: str) -> list[dict]:
        """API応答テキストからJSON配列を抽出

        classify_images()とmatch_photos_to_rows()はJSON配列を返すため、
        既存の_extract_json()（dictのみ対応）とは別に配列対応版を用意。
        """
        text = response_text.strip()

        # 戦略1: ```json ... ``` ブロックの抽出
        code_block = self._extract_code_block(text)
        if code_block:
            try:
                result = json.loads(code_block)
                if isinstance(result, list):
                    return result
                if isinstance(result, dict):
                    return [result]
            except json.JSONDecodeError:
                pass

        # 戦略2: 最外の [ ... ] を抽出
        bracket_start = text.find("[")
        bracket_end = text.rfind("]")
        if bracket_start >= 0 and bracket_end > bracket_start:
            json_text = text[bracket_start:bracket_end + 1]
            # 末尾カンマ修復
            json_text = re.sub(r",\s*([}\]])", r"\1", json_text)
            try:
                result = json.loads(json_text)
                if isinstance(result, list):
                    return result
            except json.JSONDecodeError:
                pass

        # 戦略3: dict応答の場合（配列がネストされている可能性）
        dict_result = self._extract_json(text)
        if isinstance(dict_result, dict):
            # "results" や "classifications" 等のキーに配列があるか探す
            for key in ("results", "classifications", "matches", "items"):
                if key in dict_result and isinstance(dict_result[key], list):
                    return dict_result[key]

        logger.error("JSON配列の抽出に失敗")
        logger.debug(f"Response (先頭500文字): {text[:500]}")
        return []

    def _fix_json_errors(self, text: str) -> Optional[dict]:
        """よくあるJSONエラーの修復を試行"""
        # { ... } 部分を抽出
        start = text.find("{")
        end = text.rfind("}")
        if start < 0 or end <= start:
            return None

        json_text = text[start:end + 1]

        # 修復1: 末尾カンマの除去（ ,} → }, ,] → ]）
        json_text = re.sub(r",\s*([}\]])", r"\1", json_text)

        # 修復2: 制御文字の除去（JSONの文字列値内の改行など）
        # ただし \n, \t のエスケープシーケンスは保持

        try:
            return json.loads(json_text)
        except json.JSONDecodeError as e:
            logger.debug(f"JSON修復失敗: {e}")
            return None
