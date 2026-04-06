"""LED見積作成パイプライン

現調資料フォルダから見積Excelを自動生成する統合モジュール。

3ステップフロー（UI用）:
  Step 1: run_step1_ocr()        → チェックシートOCR + テキスト入力 → パース
  Step 2: run_step2_photo_suggest() → AI写真マッチング推定
  Step 3a: run_step3_preview()    → LED選定プレビュー
  Step 3b: run_step3_generate()   → LED選定 → Excel出力
"""

from __future__ import annotations

import json
import logging
from dataclasses import dataclass, field
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


# ============================================================
# 3ステップ分割API（UI用）
# ============================================================

@dataclass
class Step1Result:
    """Step 1（OCR + パース）の結果。UIで編集してStep 2/3に渡す。"""

    # OCR結果（生データ）— UI編集フォームの初期値に使う
    ocr_result: dict = field(default_factory=dict)

    # パース済みデータ — 編集後に確定版として使う
    survey: "SurveyData" = None  # type: ignore[assignment]

    # 画像情報
    all_images: list[Path] = field(default_factory=list)
    checksheet_paths: list[Path] = field(default_factory=list)


def run_step1_ocr(
    survey_dir: Optional[Path] = None,
    api_key: Optional[str] = None,
    property_name: Optional[str] = None,
    text_input: Optional[str] = None,
) -> Step1Result:
    """Step 1: チェックシートOCR + テキスト入力 → パース

    アップロードされたファイルは全てチェックシートとして扱う。
    テキスト入力がある場合はOCRスキップしてテキストからパースのみ行う。
    両方ある場合はOCR結果とテキストを結合してパースする。

    Args:
        survey_dir: 現調資料フォルダ（画像がある場合）
        api_key: Anthropic APIキー
        property_name: 物件名（省略時はフォルダ名）
        text_input: テキスト入力（手入力の器具情報など）

    Returns:
        Step1Result: OCR結果・パース済みSurveyData
    """
    from document_processor import DocumentProcessor
    from survey_parser import parse_survey_ocr

    result = Step1Result()
    ocr_text = ""

    # --- 画像からOCR ---
    if survey_dir is not None:
        logger.info("Step 1: 画像の収集")
        all_images = find_survey_images(survey_dir)
        if not all_images and text_input is None:
            raise FileNotFoundError(f"画像が見つかりません: {survey_dir}")
        result.all_images = all_images
        logger.info(f"発見した画像: {len(all_images)}枚")

        if all_images:
            # アップされたファイルは全てチェックシートとして扱う
            result.checksheet_paths = list(all_images)

            processor = DocumentProcessor(api_key=api_key)
            logger.info("Step 1: チェックシートOCR")
            result.ocr_result = processor.ocr_survey_sheets(
                [str(p) for p in result.checksheet_paths],
            )
            ocr_text = result.ocr_result.get("raw_text", "")

    # --- テキスト入力の処理 ---
    if text_input and ocr_text:
        # 両方ある場合: OCR結果とテキストを結合
        logger.info("Step 1: OCR結果とテキスト入力を結合してパース")
        combined_text = ocr_text + "\n" + text_input
        result.ocr_result["raw_text"] = combined_text
        survey = parse_survey_ocr(result.ocr_result, fixture_photos={})
    elif text_input:
        # テキストのみ: OCRスキップ
        logger.info("Step 1: テキスト入力からパース（OCRスキップ）")
        result.ocr_result = {"raw_text": text_input}
        survey = parse_survey_ocr(result.ocr_result, fixture_photos={})
    else:
        # OCRのみ
        logger.info("Step 1: OCR結果パース")
        survey = parse_survey_ocr(result.ocr_result, fixture_photos={})

    # 物件名の設定
    if property_name:
        survey.property_info.name = property_name
    elif survey_dir and not survey.property_info.name:
        survey.property_info.name = survey_dir.name

    result.survey = survey
    return result


def run_step2_photo_suggest(
    fixture_photos: list[Path],
    ocr_fixtures: list[dict],
    api_key: Optional[str] = None,
) -> list[dict]:
    """Step 2: AI写真マッチング推定（デフォルト選択用）

    器具写真をOCR結果の器具行にAIが自動マッチングする。
    結果はあくまで「推定」で、UIでユーザーが修正する前提。

    Args:
        fixture_photos: 器具写真パスのリスト
        ocr_fixtures: OCR結果のfixtures配列（row_label, location等を含む）
        api_key: Anthropic APIキー

    Returns:
        AIの推定結果リスト。各要素:
        {
            "photo_index": int,      # fixture_photos内のインデックス
            "photo_path": str,       # 写真ファイルパス
            "row_label": str,        # 推定された行ラベル（"A"〜"T"）
            "confidence": str,       # "high" / "medium" / "low"
            "reason": str,           # 推定理由
        }
    """
    from document_processor import DocumentProcessor

    if not fixture_photos or not ocr_fixtures:
        return []

    processor = DocumentProcessor(api_key=api_key)

    try:
        # AIマッチング実行
        photo_map = processor.match_photos_to_rows(
            fixture_photos, ocr_fixtures,
        )
        # photo_mapは {row_label: [Path, ...]} 形式
        # UIで使いやすい形式に変換
        suggestions = []
        for label, paths in photo_map.items():
            for path in paths:
                idx = None
                for i, fp in enumerate(fixture_photos):
                    if fp == path or str(fp) == str(path):
                        idx = i
                        break
                suggestions.append({
                    "photo_index": idx,
                    "photo_path": str(path),
                    "row_label": label,
                    "confidence": "medium",  # 個別信頼度は元APIにない
                    "reason": "AI自動推定",
                })
        return suggestions

    except Exception as e:
        logger.warning(f"AI写真マッチングに失敗: {e}")
        return []


def run_step3_preview(
    fixtures: list,
    lineup_dir: Path,
) -> tuple:
    """Step 3a: LED選定プレビュー

    各器具に対してAI選定結果と代替候補リストを返す。
    ユーザーが結果を確認・変更してからExcel生成に進むためのステップ。

    Returns:
        (matches, candidates_map)
        - matches: list[MatchResult] — 各器具のAI選定結果
        - candidates_map: dict[str, list[LEDProduct]] — {行ラベル: 候補リスト}
    """
    from lineup_loader import LineupIndex
    from led_matcher import LEDMatcher

    logger.info("Step 3a: LED選定プレビュー")
    lineup_idx = LineupIndex()
    lineup_idx.load_all(lineup_dir)

    feedback_rules = _load_feedback_rules()
    matcher = LEDMatcher(lineup_idx, feedback_rules=feedback_rules)
    matches = matcher.match_all(fixtures)

    # 各器具の代替候補を取得
    candidates_map = {}
    for fixture in fixtures:
        if fixture.is_excluded:
            continue
        candidates = matcher.get_top_candidates(fixture, max_count=5)
        candidates_map[fixture.row_label] = candidates

    logger.info(f"プレビュー完了: {len(matches)}件選定, {len(candidates_map)}件の候補リスト")
    return matches, candidates_map


def run_step3_generate(
    survey: "SurveyData",
    lineup_dir: Path,
    template_dir: Path,
    template_name: str = "田村基本形",
    output_path: Optional[Path] = None,
    user_led_selections: Optional[dict] = None,
) -> Path:
    """Step 3: LED選定 → Excel出力

    ユーザーが確定した器具情報・写真紐付けからExcelを生成する。
    user_led_selectionsが指定された場合、AIの自動選定を上書きする。

    Args:
        survey: 確定済みのSurveyData（Step 1で編集、Step 2で写真紐付け済み）
        lineup_dir: ラインナップ表ディレクトリ
        template_dir: テンプレートディレクトリ
        template_name: テンプレート名
        output_path: 出力先パス
        user_led_selections: ユーザーが手動選択したLED商品
            {行ラベル: LEDProduct or None} 形式。指定された行はAI選定を上書き。

    Returns:
        生成されたExcelファイルパス
    """
    from lineup_loader import LineupIndex
    from image_handler import LineupImageIndex
    from led_matcher import LEDMatcher
    from excel_writer import ExcelWriter
    from models import QuotationJob, MatchResult

    # --- ラインナップ読み込み & LED選定 ---
    logger.info("Step 3: LED選定")
    lineup_idx = LineupIndex()
    lineup_idx.load_all(lineup_dir)

    img_idx = LineupImageIndex()
    img_idx.load_all(lineup_dir)

    feedback_rules = _load_feedback_rules()
    matcher = LEDMatcher(lineup_idx, feedback_rules=feedback_rules)
    matches = matcher.match_all(survey.fixtures)

    # ユーザーの手動選択で上書き
    if user_led_selections:
        for match in matches:
            label = match.fixture.row_label
            if label in user_led_selections:
                user_product = user_led_selections[label]
                if user_product is not None:
                    match.led_product = user_product
                    match.match_notes = "ユーザー手動選択"
                    match.confidence = 1.0
                    match.needs_review = False

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

    # --- Excel出力 ---
    logger.info("Step 3: Excel出力")
    job = QuotationJob(
        survey=survey,
        matches=[m for m in matches if not m.fixture.is_excluded],
        template_name=template_name,
        output_path=output_path,
    )

    writer = ExcelWriter(template_dir, image_index=img_idx)
    result_path = writer.write_quotation(job)

    logger.info(f"見積作成完了: {result_path}")
    return result_path


def save_feedback_rule(
    fixture_type: str,
    wrong_selection: str,
    correct_selection: str,
) -> None:
    """LED選定の修正をルールファイルに即時反映

    同じfixture_typeとwrong_selectionの組み合わせが既にあればcountを増加、
    なければ新規追加。
    """
    rules_path = Path(__file__).parent.parent / "feedback" / "led_selection_rules.json"
    rules_path.parent.mkdir(parents=True, exist_ok=True)

    rules = []
    if rules_path.exists():
        try:
            rules = json.loads(rules_path.read_text(encoding="utf-8"))
        except Exception:
            rules = []

    # 既存ルールを検索
    found = False
    for rule in rules:
        if (rule.get("fixture_type") == fixture_type and
                rule.get("wrong_selection") == wrong_selection):
            rule["correct_selection"] = correct_selection
            rule["count"] = rule.get("count", 0) + 1
            found = True
            break

    if not found:
        rules.append({
            "fixture_type": fixture_type,
            "wrong_selection": wrong_selection,
            "correct_selection": correct_selection,
            "count": 1,
        })

    rules_path.write_text(
        json.dumps(rules, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    logger.info(f"フィードバックルール更新: {fixture_type} → {correct_selection}")


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
            applicable = [r for r in rules if r.get("count", 0) >= 1]
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
