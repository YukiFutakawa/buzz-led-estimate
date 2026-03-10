"""見積テンプレートExcelへのデータ書き込み"""

from __future__ import annotations

import logging
import shutil
from pathlib import Path
from typing import Optional

import openpyxl
from openpyxl.utils import get_column_letter

from models import (
    ExistingFixture,
    FloorQuantities,
    LEDProduct,
    MatchResult,
    PropertyInfo,
    QuotationJob,
    SurveyData,
    TEMPLATE_SHEET_MAP,
    INPUT_SHEET,
    SELECTION_SHEET,
    BREAKDOWN_SHEET,
    EXCLUSION_SHEET,
)
from image_handler import (
    LineupImageIndex,
    resize_for_cell,
    prepare_fixture_photo,
    insert_image_to_cell,
    BREAKDOWN_PHOTO_W, BREAKDOWN_PHOTO_H,
    EXCLUSION_PHOTO_W, EXCLUSION_PHOTO_H,
    SELECTION_PHOTO1_W, SELECTION_PHOTO2_W, SELECTION_PHOTO_H,
)

logger = logging.getLogger(__name__)

# ☆入力シートの行ラベル（A〜AD）と対応する行番号
ROW_LABELS = [
    "A", "B", "C", "D", "E", "F", "G", "H", "I", "J",
    "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T",
    "U", "V", "W", "Z", "Y", "Z", "AA", "AB", "AC", "AD",
]


def _get_sheet_names(template_name: str) -> dict[str, str]:
    """テンプレート名からシート名マッピングを取得"""
    if template_name in TEMPLATE_SHEET_MAP:
        return TEMPLATE_SHEET_MAP[template_name]
    return TEMPLATE_SHEET_MAP["default"]


def _col_to_idx(col_letter: str) -> int:
    """列文字をopenpyxlの列番号（1始まり）に変換"""
    result = 0
    for c in col_letter.upper():
        result = result * 26 + (ord(c) - ord('A') + 1)
    return result


def _safe_write(ws, row: int, col: int, value) -> None:
    """結合セルを考慮した安全な書き込み

    結合セルに書き込もうとした場合、その結合範囲の左上セルに書き込む。
    """
    cell = ws.cell(row=row, column=col)
    if not isinstance(cell, openpyxl.cell.cell.MergedCell):
        cell.value = value
        return

    # MergedCellの場合: 結合範囲の左上セルを探す
    for merged_range in ws.merged_cells.ranges:
        if (merged_range.min_row <= row <= merged_range.max_row and
                merged_range.min_col <= col <= merged_range.max_col):
            ws.cell(row=merged_range.min_row,
                    column=merged_range.min_col).value = value
            return

    # フォールバック: そのまま書き込む
    cell.value = value


class ExcelWriter:
    """見積テンプレートへの書き込みエンジン"""

    def __init__(
        self,
        template_dir: Path,
        image_index: Optional[LineupImageIndex] = None,
    ):
        self.template_dir = template_dir
        self.image_index = image_index

    def write_quotation(self, job: QuotationJob) -> Path:
        """QuotationJobの内容をテンプレートに書き込み、出力パスを返す"""

        # テンプレートファイルを特定
        template_path = self._find_template(job.template_name)
        if template_path is None:
            raise FileNotFoundError(
                f"テンプレートが見つかりません: {job.template_name}"
            )

        # 出力先にテンプレートをコピー
        if job.output_path is None:
            output_dir = self.template_dir.parent / "output"
            output_dir.mkdir(exist_ok=True)
            safe_name = job.survey.property_info.name or "unnamed"
            job.output_path = output_dir / f"【LED導入ｼﾐｭﾚｰｼｮﾝ】{safe_name}.xlsx"

        shutil.copy2(template_path, job.output_path)

        # Excelファイルを開いて書き込み
        wb = openpyxl.load_workbook(job.output_path)
        sheet_names = _get_sheet_names(job.template_name)

        # ☆入力シートへの書き込み
        ws_input = wb[sheet_names["input"]]
        self._write_property_info(ws_input, job.survey.property_info)
        self._write_fixture_rows(ws_input, job.matches)
        self._write_excluded_rows(ws_input, job.survey.excluded_fixtures)

        # 選定シートへの書き込み
        ws_selection = wb[sheet_names["selection"]]
        self._write_selection_sheet(ws_selection, job.matches)

        # 写真挿入（画像インデックスがある場合のみ）
        if self.image_index:
            # 選定シートにLED商品写真を挿入
            self._write_selection_photos(ws_selection, job.matches)

            # ⑩内訳シートにLED商品写真を挿入
            breakdown_name = sheet_names.get("breakdown")
            if breakdown_name and breakdown_name in wb.sheetnames:
                ws_breakdown = wb[breakdown_name]
                self._write_breakdown_photos(
                    ws_breakdown, job.matches, job.survey,
                )

            # ⑪除外シートに除外器具写真を挿入
            exclusion_name = sheet_names.get("exclusion")
            if exclusion_name and exclusion_name in wb.sheetnames:
                ws_exclusion = wb[exclusion_name]
                self._write_exclusion_photos(
                    ws_exclusion, job.survey.excluded_fixtures,
                )

        wb.save(job.output_path)
        wb.close()

        logger.info(f"見積ファイル出力完了: {job.output_path}")
        return job.output_path

    def _find_template(self, template_name: str) -> Optional[Path]:
        """テンプレート名からファイルパスを検索"""
        for f in self.template_dir.iterdir():
            if f.suffix == ".xlsx" and template_name in f.name:
                return f
        return None

    def _write_property_info(self, ws, info: PropertyInfo) -> None:
        """物件情報をヘッダーエリアに書き込み"""
        # ☆入力シートのヘッダー部分
        # A5:B5=ラベル結合, C5:F5=データ結合 の構造
        if info.name:
            _safe_write(ws, 5, _col_to_idx("C"), info.name)
        if info.address:
            _safe_write(ws, 6, _col_to_idx("C"), info.address)
        if info.unlock_code:
            _safe_write(ws, 7, _col_to_idx("C"), info.unlock_code)
        if info.distribution_board:
            _safe_write(ws, 8, _col_to_idx("C"), info.distribution_board)
        if info.special_notes:
            _safe_write(ws, 9, _col_to_idx("C"), info.special_notes)

        logger.info(f"物件情報書き込み: {info.name} / {info.address}")

    def _write_fixture_rows(self, ws, matches: list[MatchResult]) -> None:
        """☆入力シートのデータ行（Row 16-45）に器具データを書き込み"""
        start_row = INPUT_SHEET["data_start_row"]  # 16

        for i, match in enumerate(matches):
            if i >= 30:  # 最大30行（A-AD）
                logger.warning("器具種別が30を超えました。超過分はスキップします。")
                break

            row = start_row + i
            fixture = match.fixture

            # C列: 試算エリア（設置場所）
            _safe_write(ws, row, _col_to_idx("C"), fixture.location)

            # D列: 照明種別
            _safe_write(ws, row, _col_to_idx("D"), fixture.fixture_type)

            # E列: 現調備考
            if fixture.survey_notes:
                _safe_write(ws, row, _col_to_idx("E"), fixture.survey_notes)

            # F列: 工事備考
            if fixture.construction_notes:
                _safe_write(ws, row, _col_to_idx("F"), fixture.construction_notes)

            # G列: 器具分類②（選定シートへのリンクキー）
            if match.category_key:
                _safe_write(ws, row, _col_to_idx("G"), match.category_key)

            # I列: 一日点灯時間
            if fixture.daily_hours > 0:
                _safe_write(ws, row, _col_to_idx("I"), fixture.daily_hours)

            # K列: 消費電力（安定器補正済み）
            if fixture.adjusted_power_w > 0:
                _safe_write(ws, row, _col_to_idx("K"), fixture.adjusted_power_w)

            # L列: 電球数（合計）
            if fixture.quantities.total > 0:
                _safe_write(ws, row, _col_to_idx("L"), fixture.quantities.total)

            # M-V列: 各階数量（1F〜10F）
            floor_start_col = _col_to_idx("M")  # M=13
            for floor_idx, qty in enumerate(
                fixture.quantities.to_list(10)
            ):
                if qty > 0:
                    _safe_write(ws, row, floor_start_col + floor_idx, qty)

            # AE列: 工事単価
            if match.construction_unit_price > 0:
                _safe_write(ws, row, _col_to_idx("AE"),
                            match.construction_unit_price)

        logger.info(f"器具データ {len(matches)}行を書き込みました")

    def _write_excluded_rows(self, ws,
                             excluded: list[ExistingFixture]) -> None:
        """☆入力シートのLED済みセクション（Row 49+）に除外データを書き込み"""
        start_row = INPUT_SHEET["excluded_start_row"]  # 49

        for i, fixture in enumerate(excluded):
            if i >= 10:  # 最大10行
                break

            row = start_row + i

            # C列: 試算エリア
            _safe_write(ws, row, _col_to_idx("C"), fixture.location)

            # D列: 照明種別
            _safe_write(ws, row, _col_to_idx("D"), fixture.fixture_type)

            # E列: 現調備考
            if fixture.survey_notes:
                _safe_write(ws, row, _col_to_idx("E"), fixture.survey_notes)

            # L列: 電球数
            if fixture.quantities.total > 0:
                _safe_write(ws, row, _col_to_idx("L"),
                            fixture.quantities.total)

            # W列: 除外理由
            if fixture.exclusion_reason:
                _safe_write(ws, row, _col_to_idx("W"),
                            fixture.exclusion_reason)

            # AC列: アドバイス
            if fixture.exclusion_advice:
                _safe_write(ws, row, _col_to_idx("AC"),
                            fixture.exclusion_advice)

        if excluded:
            logger.info(f"除外データ {len(excluded)}行を書き込みました")

    def _write_selection_sheet(self, ws,
                               matches: list[MatchResult]) -> None:
        """選定シートにLED商品仕様を書き込み"""
        start_row = SELECTION_SHEET["data_start_row"]  # 3

        # 同じcategory_keyのマッチを重複除去（選定シートは1カテゴリ1行）
        seen_keys: set[str] = set()
        unique_matches: list[MatchResult] = []
        for match in matches:
            if match.category_key and match.category_key not in seen_keys:
                seen_keys.add(match.category_key)
                unique_matches.append(match)

        for i, match in enumerate(unique_matches):
            if i >= 30:
                break
            if match.led_product is None:
                continue

            row = start_row + i
            led = match.led_product

            # C列: リンクキー（☆入力!G列と一致させる）
            _safe_write(ws, row, _col_to_idx("C"), match.category_key)

            # D列: 照明色
            _safe_write(ws, row, _col_to_idx("D"), led.lighting_color)

            # E列: 器具色
            _safe_write(ws, row, _col_to_idx("E"), led.fixture_color)

            # F列: 器具サイズ
            _safe_write(ws, row, _col_to_idx("F"), led.fixture_size)

            # G列: 消費電力
            if led.power_w > 0:
                _safe_write(ws, row, _col_to_idx("G"), led.power_w)

            # H列: 全光束
            _safe_write(ws, row, _col_to_idx("H"), led.lumens)

            # I列: 合算定価
            if led.list_price_total > 0:
                _safe_write(ws, row, _col_to_idx("I"), led.list_price_total)

            # J列: 合算仕入
            if led.purchase_price_total > 0:
                _safe_write(ws, row, _col_to_idx("J"),
                            led.purchase_price_total)

            # K列: 防滴
            _safe_write(ws, row, _col_to_idx("K"),
                        "〇" if led.is_waterproof else "✕")

            # L列: 電球種別
            _safe_write(ws, row, _col_to_idx("L"), led.bulb_type)

            # M列: メーカー
            _safe_write(ws, row, _col_to_idx("M"), led.manufacturer)

            # N列: W相当
            _safe_write(ws, row, _col_to_idx("N"), led.watt_equivalent)

            # O列: 器具型番
            _safe_write(ws, row, _col_to_idx("O"), led.model_number)

            # P列: 定価
            if led.model_price > 0:
                _safe_write(ws, row, _col_to_idx("P"), led.model_price)

            # Q列: 仕入れ
            if led.model_purchase > 0:
                _safe_write(ws, row, _col_to_idx("Q"), led.model_purchase)

            # R-Z列: 追加型番・価格
            if led.model_number_2:
                _safe_write(ws, row, _col_to_idx("R"), led.model_number_2)
            if led.model_price_2:
                _safe_write(ws, row, _col_to_idx("S"), led.model_price_2)
            if led.model_purchase_2:
                _safe_write(ws, row, _col_to_idx("T"), led.model_purchase_2)
            if led.model_number_3:
                _safe_write(ws, row, _col_to_idx("U"), led.model_number_3)
            if led.model_price_3:
                _safe_write(ws, row, _col_to_idx("V"), led.model_price_3)
            if led.model_purchase_3:
                _safe_write(ws, row, _col_to_idx("W"), led.model_purchase_3)
            if led.model_number_4:
                _safe_write(ws, row, _col_to_idx("X"), led.model_number_4)
            if led.model_price_4:
                _safe_write(ws, row, _col_to_idx("Y"), led.model_price_4)
            if led.model_purchase_4:
                _safe_write(ws, row, _col_to_idx("Z"), led.model_purchase_4)

            # AA列: 消費電力（詳細）
            if led.power_detail > 0:
                _safe_write(ws, row, _col_to_idx("AA"), led.power_detail)

            # AB列: 全光束（詳細）
            _safe_write(ws, row, _col_to_idx("AB"), led.lumens_detail)

            # AC列: 器具素材
            _safe_write(ws, row, _col_to_idx("AC"), led.material)

            # AD列: 器具色選択肢
            _safe_write(ws, row, _col_to_idx("AD"), led.color_options)

            # AE列: 照明色選択肢
            _safe_write(ws, row, _col_to_idx("AE"),
                        led.lighting_color_options)

            # AF列: 定格寿命
            _safe_write(ws, row, _col_to_idx("AF"), led.lifespan)

            # AG列: 交換方法
            _safe_write(ws, row, _col_to_idx("AG"), led.replacement_method)

            # AH列: 口金
            _safe_write(ws, row, _col_to_idx("AH"), led.socket)

        logger.info(
            f"選定データ {len(unique_matches)}カテゴリを書き込みました"
        )

    # ===== 写真挿入メソッド =====

    def _write_selection_photos(
        self, ws, matches: list[MatchResult],
    ) -> None:
        """選定シートにLED商品写真を挿入

        A列=写真①、B列=写真② (各行は1カテゴリ)
        """
        start_row = SELECTION_SHEET["data_start_row"]  # 3

        # 重複除去（選定シートは1カテゴリ1行）
        seen_keys: set[str] = set()
        unique_matches: list[MatchResult] = []
        for match in matches:
            if match.category_key and match.category_key not in seen_keys:
                seen_keys.add(match.category_key)
                unique_matches.append(match)

        photo_count = 0
        for i, match in enumerate(unique_matches):
            if i >= 30:
                break
            if match.led_product is None:
                continue

            row = start_row + i

            # 写真①（A列）
            img1_data = self.image_index.get_product_image(
                match.led_product, photo_num=1,
            )
            if img1_data:
                try:
                    resized = resize_for_cell(
                        img1_data, SELECTION_PHOTO1_W, SELECTION_PHOTO_H,
                    )
                    insert_image_to_cell(
                        ws, resized, row, 1,  # A=1
                        SELECTION_PHOTO1_W, SELECTION_PHOTO_H,
                    )
                    photo_count += 1
                except Exception as e:
                    logger.warning(f"選定写真①挿入エラー (row={row}): {e}")

            # 写真②（B列）
            img2_data = self.image_index.get_product_image(
                match.led_product, photo_num=2,
            )
            if img2_data:
                try:
                    resized = resize_for_cell(
                        img2_data, SELECTION_PHOTO2_W, SELECTION_PHOTO_H,
                    )
                    insert_image_to_cell(
                        ws, resized, row, 2,  # B=2
                        SELECTION_PHOTO2_W, SELECTION_PHOTO_H,
                    )
                    photo_count += 1
                except Exception as e:
                    logger.warning(f"選定写真②挿入エラー (row={row}): {e}")

        logger.info(f"選定シート写真挿入: {photo_count}枚")

    def _write_breakdown_photos(
        self, ws, matches: list[MatchResult],
        survey: SurveyData,
    ) -> None:
        """⑩内訳シートに既存器具写真とLED商品写真を挿入

        Row 7: 既存器具写真（現調写真）→ B-U列
        Row 14: LED商品写真（ラインナップ表から）→ B-U列
        各列はRow 16-35（☆入力の器具行A-T）に対応
        """
        existing_row = BREAKDOWN_SHEET["existing_photo_row"]  # 7
        led_row = BREAKDOWN_SHEET["led_photo_row"]            # 14
        photo_w = BREAKDOWN_PHOTO_W                           # 100
        photo_h = BREAKDOWN_PHOTO_H                           # 90

        photo_count = 0
        for i, match in enumerate(matches):
            if i >= 20:  # B-U列 = 最大20列
                break

            col = _col_to_idx("B") + i  # B=2, C=3, ...

            # Row 7: 既存器具写真（現調写真を中央トリミング→照明アップ）
            fixture = match.fixture
            if fixture.photo_paths:
                photo_path = fixture.photo_paths[0]
                if photo_path.exists():
                    try:
                        resized = prepare_fixture_photo(
                            photo_path, photo_w, photo_h,
                        )
                        insert_image_to_cell(
                            ws, resized, existing_row, col,
                            photo_w, photo_h,
                        )
                        photo_count += 1
                    except Exception as e:
                        logger.warning(
                            f"内訳 既存写真挿入エラー (col={col}): {e}"
                        )

            # Row 14: LED商品写真（ラインナップ表から）
            if match.led_product:
                img_data = self.image_index.get_product_image(
                    match.led_product, photo_num=1,
                )
                if img_data:
                    try:
                        resized = resize_for_cell(
                            img_data, photo_w, photo_h,
                        )
                        insert_image_to_cell(
                            ws, resized, led_row, col,
                            photo_w, photo_h,
                        )
                        photo_count += 1
                    except Exception as e:
                        logger.warning(
                            f"内訳 LED写真挿入エラー (col={col}): {e}"
                        )

        logger.info(f"⑩内訳シート写真挿入: {photo_count}枚")

    def _write_exclusion_photos(
        self, ws, excluded: list[ExistingFixture],
    ) -> None:
        """⑪除外シートに除外器具の現調写真を挿入

        12ブロック: 左6(B列, Row 4/8/12/16/20/24) + 右6(I列)
        """
        blocks = EXCLUSION_SHEET["blocks"]
        photo_w = EXCLUSION_PHOTO_W  # 78
        photo_h = EXCLUSION_PHOTO_H  # 80

        photo_count = 0
        for i, fixture in enumerate(excluded):
            if i >= len(blocks):
                break

            block = blocks[i]
            col = _col_to_idx(block["photo_col"])
            row = block["photo_row"]

            # 除外器具の現調写真（中央トリミング+高画質）
            if fixture.photo_paths:
                photo_path = fixture.photo_paths[0]
                if photo_path.exists():
                    try:
                        resized = prepare_fixture_photo(
                            photo_path, photo_w, photo_h,
                        )
                        insert_image_to_cell(
                            ws, resized, row, col,
                            photo_w, photo_h,
                        )
                        photo_count += 1
                    except Exception as e:
                        logger.warning(
                            f"除外写真挿入エラー (block={i}): {e}"
                        )

        if photo_count > 0:
            logger.info(f"⑪除外シート写真挿入: {photo_count}枚")
