"""LED見積作成パイプライン

現調資料フォルダから見積Excelを自動生成する統合モジュール。

フロー:
  1. 現調フォルダ内の写真を収集
  1.5 AI画像分類（チェックシート / 器具写真 / 建物外観を自動判定）
  2. チェックシートをOCR → 構造化データ
  3. 器具写真をAIで行ラベルに自動紐付け → photo_paths
  4. LED選定
  5. Excel出力（写真付き）
"""

from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)

# 画像ファイルの拡張子
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".heic"}


def find_survey_images(survey_dir: Path) -> list[Path]:
    """現調フォルダから画像ファイルを番号順にソートして返す

    サブフォルダがある場合は最も画像が多いサブフォルダを使用する。
    """
    # 直下の画像を探す
    images = _collect_images(survey_dir)

    # 直下になければサブフォルダを探す
    if not images:
        best_dir = None
        best_count = 0
        for d in survey_dir.iterdir():
            if d.is_dir():
                imgs = _collect_images(d)
                if len(imgs) > best_count:
                    best_count = len(imgs)
                    best_dir = d
                    images = imgs

        if best_dir:
            logger.info(f"サブフォルダを使用: {best_dir.name} ({best_count}枚)")

    return images


def _collect_images(directory: Path) -> list[Path]:
    """ディレクトリ内の画像を番号順にソート"""
    import re

    images = [
        f for f in directory.iterdir()
        if f.is_file() and f.suffix.lower() in IMAGE_EXTS
    ]

    def _sort_key(p: Path) -> int:
        m = re.search(r'_(\d+)\.', p.name)
        if m:
            return int(m.group(1))
        return 0

    return sorted(images, key=_sort_key)


def separate_checksheet_and_photos(
    images: list[Path],
    checksheet_indices: Optional[list[int]] = None,
) -> tuple[list[Path], list[Path]]:
    """画像リストをチェックシートと器具写真に分離（従来方式・後方互換）

    Args:
        images: 全画像パスリスト（番号順）
        checksheet_indices: チェックシートの画像インデックス（0-based）。
            Noneの場合はインデックス0（_1.jpg）をチェックシートとする。

    Returns:
        (checksheet_paths, fixture_photo_paths)
    """
    if checksheet_indices is None:
        checksheet_indices = [0]

    cs_set = set(checksheet_indices)
    checksheets = [images[i] for i in cs_set if i < len(images)]
    fixture_photos = [
        img for i, img in enumerate(images) if i not in cs_set
    ]

    logger.info(
        f"チェックシート: {len(checksheets)}枚, "
        f"器具写真: {len(fixture_photos)}枚"
    )
    return checksheets, fixture_photos


def run_pipeline(
    survey_dir: Path,
    lineup_dir: Path,
    template_dir: Path,
    template_name: str = "田村基本形",
    output_path: Optional[Path] = None,
    checksheet_indices: Optional[list[int]] = None,
    api_key: Optional[str] = None,
    property_name: Optional[str] = None,
) -> Path:
    """現調フォルダからLED見積Excelを生成

    Args:
        survey_dir: 現調資料フォルダ（写真が入ったフォルダ）
        lineup_dir: ラインナップ表フォルダ
        template_dir: 見積テンプレートフォルダ
        template_name: テンプレート名
        output_path: 出力パス（省略時は自動生成）
        checksheet_indices: チェックシートのインデックス。
            指定時は従来の手動方式。省略(None)時はAI自動分類。
        api_key: Anthropic APIキー（省略時は環境変数）

    Returns:
        出力Excelファイルパス
    """
    from document_processor import DocumentProcessor
    from survey_parser import parse_survey_ocr, match_photos_to_fixtures
    from lineup_loader import LineupIndex
    from image_handler import LineupImageIndex
    from led_matcher import LEDMatcher
    from excel_writer import ExcelWriter
    from models import QuotationJob, MatchResult

    # --- Step 1: 画像の収集 ---
    logger.info("=" * 50)
    logger.info("Step 1: 画像の収集")
    all_images = find_survey_images(survey_dir)
    if not all_images:
        raise FileNotFoundError(f"画像が見つかりません: {survey_dir}")
    logger.info(f"発見した画像: {len(all_images)}枚")

    # DocumentProcessorの初期化（AI分類・OCRの両方で使用）
    processor = DocumentProcessor(api_key=api_key)

    # --- Step 1.5: 画像の分類 ---
    building_photos = []

    if checksheet_indices is not None:
        # 従来方式: インデックス指定による手動分類
        logger.info("Step 1.5: 画像の分類（手動指定モード）")
        checksheets, fixture_photos = separate_checksheet_and_photos(
            all_images, checksheet_indices,
        )
    else:
        # AI自動分類モード
        logger.info("Step 1.5: 画像の分類（AI自動分類モード）")
        try:
            classification = processor.classify_images(all_images)

            checksheets = [
                all_images[c["index"]]
                for c in classification
                if c["type"] == "checksheet"
            ]
            fixture_photos = [
                all_images[c["index"]]
                for c in classification
                if c["type"] == "fixture"
            ]
            building_photos = [
                all_images[c["index"]]
                for c in classification
                if c["type"] == "building"
            ]

            if not checksheets:
                # チェックシートなし → 写真直接解析モードに切替
                logger.info(
                    "チェックシートなし → 写真直接解析モードに切替"
                )
                if len(all_images) <= 4:
                    raise ValueError(
                        f"写真が{len(all_images)}枚しかありません。"
                        "自動処理には5枚以上必要です。"
                    )
                direct_result = processor.analyze_fixtures_from_photos(
                    all_images,
                    property_name=survey_dir.name,
                )
                # 写真マッピングを構築
                photo_map = {}
                for fix in direct_result.get("fixtures", []):
                    label = fix.get("row_label", "")
                    indices = fix.get("photo_indices", [])
                    if label and indices:
                        photo_map[label] = [
                            all_images[i]
                            for i in indices
                            if i < len(all_images)
                        ]
                # 建物外観写真
                b_indices = direct_result.get(
                    "building_photo_indices", []
                )
                building_photos = [
                    all_images[i]
                    for i in b_indices
                    if i < len(all_images)
                ]
                # Step 4: パース（OCR/写真紐付けをスキップ）
                survey = parse_survey_ocr(
                    direct_result, fixture_photos=photo_map,
                )
                # 物件名はフォルダ名（SFA管理名）を優先
                survey.property_info.name = survey_dir.name
                if building_photos:
                    survey.building_photo_path = building_photos[0]
                # Step 5以降に合流（LED選定へジャンプ）
                return run_from_survey_data(
                    survey=survey,
                    lineup_dir=lineup_dir,
                    template_dir=template_dir,
                    template_name=template_name,
                    output_path=output_path,
                )

            logger.info(
                f"AI分類結果: チェックシート={len(checksheets)}枚, "
                f"器具写真={len(fixture_photos)}枚, "
                f"建物写真={len(building_photos)}枚"
            )

        except Exception as e:
            # AI分類失敗時のフォールバック
            if "API" in str(e) or "初期化" in str(e):
                logger.warning(
                    f"AI分類に失敗しました: {e}\n"
                    "従来方式（1枚目=チェックシート）にフォールバックします。"
                )
                checksheets, fixture_photos = separate_checksheet_and_photos(
                    all_images, [0],
                )
            else:
                raise

    # --- Step 2: OCR ---
    logger.info("Step 2: チェックシートOCR")
    ocr_result = processor.ocr_survey_sheets(
        [str(p) for p in checksheets],
    )

    # --- Step 3: 写真紐付け ---
    logger.info("Step 3: 器具写真の紐付け")

    if checksheet_indices is not None:
        # 従来方式: ファイル名パターン or 順番で紐付け
        photo_dir = fixture_photos[0].parent if fixture_photos else survey_dir
        photo_map = match_photos_to_fixtures(
            photo_dir,
            exclude_paths=checksheets,
        )
    else:
        # AI自動マッチングモード
        if fixture_photos and ocr_result.get("fixtures"):
            try:
                photo_map = processor.match_photos_to_rows(
                    fixture_photos,
                    ocr_result.get("fixtures", []),
                )
            except Exception as e:
                logger.warning(
                    f"AI写真マッチングに失敗: {e}\n"
                    "従来方式にフォールバックします。"
                )
                photo_dir = (
                    fixture_photos[0].parent
                    if fixture_photos
                    else survey_dir
                )
                photo_map = match_photos_to_fixtures(
                    photo_dir,
                    exclude_paths=checksheets,
                )
        else:
            photo_map = {}

    # --- Step 4: パース ---
    logger.info("Step 4: OCR結果パース")
    survey = parse_survey_ocr(ocr_result, fixture_photos=photo_map)

    # 物件名: 引数 > OCR結果 > フォルダ名の優先順
    if property_name:
        survey.property_info.name = property_name
    elif not survey.property_info.name:
        survey.property_info.name = survey_dir.name

    # 建物外観写真をセット
    if building_photos:
        survey.building_photo_path = building_photos[0]

    # --- Step 5〜8: 共通処理 ---
    return run_from_survey_data(
        survey=survey,
        lineup_dir=lineup_dir,
        template_dir=template_dir,
        template_name=template_name,
        output_path=output_path,
    )


def run_from_survey_data(
    survey,
    lineup_dir: Path,
    template_dir: Path,
    template_name: str = "田村基本形",
    output_path: Optional[Path] = None,
) -> Path:
    """Step 5〜8: LED選定 → Excel出力 → 検証 → フィードバック

    SurveyData から直接パイプラインを実行する公開API。
    写真パイプライン・テキストパーサー・外部モジュールから呼び出し可能。

    Args:
        survey: SurveyData オブジェクト
        lineup_dir: ラインナップ表ディレクトリ
        template_dir: テンプレートディレクトリ
        template_name: テンプレート名
        output_path: 出力先パス（Noneなら自動生成）

    Returns:
        生成されたExcelファイルパス
    """
    from lineup_loader import LineupIndex
    from image_handler import LineupImageIndex
    from led_matcher import LEDMatcher
    from excel_writer import ExcelWriter
    from models import QuotationJob, MatchResult

    # --- Step 5: ラインナップ読み込み & LED選定 ---
    logger.info("Step 5: LED選定")
    lineup_idx = LineupIndex()
    lineup_idx.load_all(lineup_dir)

    img_idx = LineupImageIndex()
    img_idx.load_all(lineup_dir)

    feedback_rules = _load_feedback_rules()
    matcher = LEDMatcher(lineup_idx, feedback_rules=feedback_rules)
    matches = matcher.match_all(survey.fixtures)

    # 除外器具のMatchResultも生成
    for excl in survey.excluded_fixtures:
        matches.append(MatchResult(
            fixture=excl,
            category_key="",
            confidence=0.0,
            match_notes=excl.exclusion_reason or "除外",
        ))

    # 画像プリロード
    led_products = [m.led_product for m in matches if m.led_product]
    img_idx.preload_images(led_products)

    # --- Step 6: Excel出力 ---
    logger.info("Step 6: Excel出力")
    job = QuotationJob(
        survey=survey,
        matches=[m for m in matches if not m.fixture.is_excluded],
        template_name=template_name,
        output_path=output_path,
    )

    writer = ExcelWriter(template_dir, image_index=img_idx)
    result_path = writer.write_quotation(job)

    # --- Step 7: Googleマップ検証 + AI異常検知 ---
    logger.info("Step 7: Googleマップ検証 + AI異常検知")
    from google_maps_checker import (
        run_maps_check, format_check_report,
        run_ai_anomaly_check, format_anomaly_report,
    )
    map_result = run_maps_check(survey)
    if map_result.checklist:
        report = format_check_report(map_result)
        logger.info("\n" + report)

    # AI異常検知
    try:
        anomaly_result = run_ai_anomaly_check(
            survey, matches, map_result,
        )
        if anomaly_result.get("anomalies"):
            logger.warning(
                "\n" + format_anomaly_report(anomaly_result)
            )
        else:
            logger.info("AI異常検知: 指摘事項なし")
    except Exception as e:
        logger.warning(f"AI異常検知エラー（スキップ）: {e}")

    # --- Step 8: フィードバック自動比較 ---
    _auto_feedback(result_path)

    logger.info(f"見積作成完了: {result_path}")
    return result_path


# 後方互換エイリアス
_run_from_step5 = run_from_survey_data


def _load_feedback_rules() -> list[dict]:
    """フィードバック学習ルールをファイルから読み込み

    feedback/led_selection_rules.json が存在すれば読み込む。
    ファイルがなければ空リスト（従来動作にフォールバック）。
    """
    rules_path = Path(__file__).parent.parent / "feedback" / "led_selection_rules.json"
    if not rules_path.exists():
        return []

    try:
        rules = json.loads(rules_path.read_text(encoding="utf-8"))
        if rules:
            applicable = [r for r in rules if r.get("count", 0) >= 2]
            logger.info(
                f"フィードバックルール読み込み: {len(rules)}件 "
                f"(適用対象: {len(applicable)}件)"
            )
        return rules
    except Exception as e:
        logger.warning(f"フィードバックルール読み込みエラー: {e}")
        return []


def _auto_feedback(ai_output_path: Path) -> None:
    """正解ファイルが存在する場合、自動で差分比較を実行"""
    try:
        from feedback_comparator import FeedbackComparator

        correct_dir = ai_output_path.parent.parent / "正しい見積り"
        feedback_dir = ai_output_path.parent.parent / "feedback"

        if not correct_dir.exists():
            return

        comparator = FeedbackComparator()
        ai_name = comparator._extract_property_name(ai_output_path.stem)

        # 対応する正解ファイルを検索
        for correct_file in correct_dir.glob("*.xlsx"):
            if correct_file.name.startswith("~$"):
                continue
            correct_name = comparator._extract_property_name(correct_file.stem)
            if ai_name and correct_name and (
                ai_name in correct_name or correct_name in ai_name
            ):
                logger.info(f"Step 8: フィードバック比較 (正解: {correct_file.name})")
                report = comparator.compare(ai_output_path, correct_file)

                json_name = f"feedback_{report.property_name}.json"
                report.save_json(feedback_dir / json_name)

                logger.info(
                    f"  差分{report.total_diffs}件, "
                    f"LED選定一致率{report.led_selection_match_rate:.0%}"
                )
                return

        logger.info("Step 8: 対応する正解ファイルなし（フィードバックスキップ）")
    except Exception as e:
        logger.warning(f"フィードバック自動比較エラー: {e}")
