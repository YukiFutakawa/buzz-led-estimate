"""LED見積作成自動化システム - データモデル定義"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Optional


class FixtureCategory(str, Enum):
    """ラインナップ表のシートに対応する器具カテゴリ"""

    # 蛍光灯ラインナップ
    FLUORESCENT_40W = "40形蛍光灯"
    FLUORESCENT_20W = "20形蛍光灯"
    LED_TUBE = "直管形LED器具"
    EMERGENCY_40W = "非常灯40形"
    EMERGENCY_20W = "非常灯20形"
    EMERGENCY_OTHER = "その他非常灯"
    OUTDOOR_BRACKET = "屋外ﾌﾞﾗｹｯﾄ"

    # その他ラインナップ
    CEILING_WALL = "天井・壁面"
    PORCH_PILLAR = "ﾎﾟｰﾁ・支柱"
    DOWNLIGHT = "ﾀﾞｳﾝﾗｲﾄ"
    DOWNLIGHT_HIGH = "DL※高出力"
    SPOTLIGHT = "ｽﾎﾟｯﾄﾗｲﾄ"
    LED_BULB = "LED球"
    EXTERIOR_BYPASS = "外部・ﾊﾞｲﾊﾟｽ"
    EE_SWITCH = "EEスイッチ他"
    BYPASS_COWELL = "バイパスコーウェル"

    # マニアックラインナップ
    GUIDE_LIGHT = "誘導灯 各社"
    FOOT_LIGHT = "ﾌｯﾄﾗｲﾄ"
    GARDEN_LIGHT = "庭園灯"
    NAMEPLATE_PORCH = "表札・ﾎﾟｰﾁ（250lm以下）"
    ACCENT_LIGHT = "ｱｸｾﾝﾄﾗｲﾄ"
    MOTION_SENSOR = "人感ｾﾝｻ"
    PORCH_LIGHT = "ﾎﾟｰﾁﾗｲﾄ"
    CYLINDER_BRACKET = "筒型ブラ"
    INDOOR = "屋内"
    GATE_PILLAR = "門柱灯"
    SECURITY_LIGHT = "防犯灯"
    POLE_LIGHT = "ポール灯"
    FLOOD_LIGHT = "投光器・高天井"
    ARM_SPOT = "ｱｰﾑｽﾎﾟｯﾄ"


class LineupFile(str, Enum):
    """ラインナップ表ファイル"""
    FLUORESCENT = "新【蛍光灯】ラインナップ.xlsx"
    OTHER = "新【その他】ラインナップ.xlsx"
    MANIAC = "新【マニアック】ラインナップ表.xlsx"


class ReplacementMethod(str, Enum):
    """交換方法"""
    FIXTURE_REPLACE = "器具交換"
    BULB_REPLACE = "ランプ交換"
    BYPASS = "バイパス工事"
    WIRING_DIRECT = "配線直結"


# 安定器のある照明タイプ（消費電力+2W）
BALLAST_TYPES = frozenset([
    "FL",   # 直管蛍光灯
    "FCL",  # サークライン
    "FDL",  # コンパクト蛍光灯
    "FHT",
    "FPL",
    "FML",
])


@dataclass
class PropertyInfo:
    """物件情報 → ☆入力シートのヘッダー部に転記"""
    name: str = ""                   # 物件名 → ☆入力!A5エリア
    address: str = ""                # 物件住所 → ☆入力!A6エリア
    unlock_code: str = ""            # 解錠番号 → ☆入力!A7エリア
    distribution_board: str = ""     # 分電盤場所 → ☆入力!A8エリア
    special_notes: str = ""          # 工事依頼特記事項 → ☆入力!A9エリア
    client_name: str = ""            # クライアント名（テンプレート選択用）
    survey_date: str = ""            # 現地調査日
    surveyor: str = ""               # 現地調査担当


@dataclass
class FloorQuantities:
    """階別数量（1F〜10F）"""
    floors: dict[int, int] = field(default_factory=dict)

    @property
    def total(self) -> int:
        return sum(self.floors.values())

    def to_list(self, max_floors: int = 10) -> list[int]:
        """1F〜max_floorsの数量をリストで返す"""
        return [self.floors.get(i, 0) for i in range(1, max_floors + 1)]


@dataclass
class ExistingFixture:
    """現地調査チェックシートの1行 = 1つの照明種別

    ☆入力シートのRow 16-45（ラベルA〜AD）に対応
    """
    row_label: str                   # A〜AD（☆入力の行ラベル）
    location: str = ""               # 設置場所 → ☆入力!C列（試算エリア）
    fixture_type: str = ""           # 器具種別 → ☆入力!D列（照明種別）
    fixture_size: str = ""           # 器具サイズ（例: W150×D90）
    bulb_type: str = ""              # 電球種別（例: FL15W）
    quantities: FloorQuantities = field(default_factory=FloorQuantities)
    power_consumption_w: float = 0   # 消費電力(W) → ☆入力!K列
    daily_hours: float = 0           # 一日点灯時間 → ☆入力!I列
    operating_days: int = 30         # 稼働日数 → ☆入力!J列
    color_temp: str = ""             # 色温度（白=昼白色, 黄=電球色）
    survey_notes: str = ""           # 現調備考 → ☆入力!E列
    construction_notes: str = ""     # 工事備考 → ☆入力!F列
    photo_paths: list[Path] = field(default_factory=list)
    is_waterproof: bool = False      # 防滴要否
    is_high_location: bool = False   # 高所作業要否

    # 除外関連
    is_excluded: bool = False        # LED済み/除外フラグ
    exclusion_reason: str = ""       # 除外理由
    exclusion_advice: str = ""       # アドバイス

    # OCR信頼度トラッキング
    ocr_confidence: str = "high"     # "high", "medium", "low"
    ocr_warnings: list[str] = field(default_factory=list)

    @property
    def has_ballast(self) -> bool:
        """安定器ありの照明かどうか（+2W補正が必要）"""
        for bt in BALLAST_TYPES:
            if bt in self.bulb_type.upper():
                return True
        return False

    @property
    def adjusted_power_w(self) -> float:
        """安定器補正後の消費電力"""
        if self.has_ballast:
            return self.power_consumption_w + 2
        return self.power_consumption_w


@dataclass
class LEDProduct:
    """ラインナップ表の1行 = 1つのLED商品"""
    source_file: str = ""            # ラインナップファイル名
    source_sheet: str = ""           # シート名
    source_row: int = 0              # 行番号

    # 基本情報（選定シートA-K列に対応）
    name: str = ""                   # C列: 名称/電球種別
    lighting_color: str = ""         # D列: 照明色
    fixture_color: str = ""          # E列: 器具色
    fixture_size: str = ""           # F列: 器具サイズ（設置面）
    power_w: float = 0               # G列: 消費電力
    lumens: str = ""                 # H列: 全光束（文字列：「ー」の場合もある）
    list_price_total: int = 0        # I列: 合算定価
    purchase_price_total: int = 0    # J列: 合算仕入
    is_waterproof: bool = False      # K列: 防滴

    # 詳細情報（選定シートL-AJ列に対応）
    bulb_type: str = ""              # L列: 電球種別
    manufacturer: str = ""           # M列: メーカー
    watt_equivalent: str = ""        # N列: W相当
    model_number: str = ""           # O列: 器具型番
    model_price: int = 0             # P列: 定価
    model_purchase: int = 0          # Q列: 仕入れ
    model_number_2: str = ""         # R列
    model_price_2: int = 0           # S列
    model_purchase_2: int = 0        # T列
    model_number_3: str = ""         # U列
    model_price_3: int = 0           # V列
    model_purchase_3: int = 0        # W列
    model_number_4: str = ""         # X列
    model_price_4: int = 0           # Y列
    model_purchase_4: int = 0        # Z列
    power_detail: float = 0          # AA列: 消費電力（詳細）
    lumens_detail: str = ""          # AB列: 全光束（詳細）
    material: str = ""               # AC列: 器具素材
    color_options: str = ""          # AD列: 器具色選択肢
    lighting_color_options: str = "" # AE列: 照明色選択肢
    lifespan: str = ""               # AF列: 定格寿命
    replacement_method: str = ""     # AG列: 交換方法
    socket: str = ""                 # AH列: 口金
    hp_link: str = ""                # AI列: HP
    notes: str = ""                  # AJ列: 備考


@dataclass
class MatchResult:
    """既存器具とLED商品のマッチング結果"""
    fixture: ExistingFixture
    led_product: Optional[LEDProduct] = None
    category_key: str = ""           # ☆入力!G列 = 選定!C列のリンクキー
    confidence: float = 0.0          # マッチング信頼度 (0-1)
    match_notes: str = ""            # 選定理由
    construction_unit_price: int = 3000  # 工事単価 → ☆入力!AE列
    needs_review: bool = False       # 人間レビュー必要フラグ


@dataclass
class SurveyData:
    """文書処理の出力 = 1案件分の全構造化データ"""
    property_info: PropertyInfo = field(default_factory=PropertyInfo)
    fixtures: list[ExistingFixture] = field(default_factory=list)
    excluded_fixtures: list[ExistingFixture] = field(default_factory=list)
    building_photo_path: Optional[Path] = None
    raw_ocr_text: str = ""

    @property
    def all_fixtures(self) -> list[ExistingFixture]:
        return self.fixtures + self.excluded_fixtures

    @property
    def total_bulbs(self) -> int:
        return sum(f.quantities.total for f in self.fixtures)


@dataclass
class QuotationJob:
    """見積作成ジョブの全情報"""
    survey: SurveyData
    matches: list[MatchResult] = field(default_factory=list)
    template_name: str = "田村基本形"
    output_path: Optional[Path] = None
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)


# テンプレートのシート名マッピング
TEMPLATE_SHEET_MAP: dict[str, dict[str, str]] = {
    "default": {
        "input": "☆入力",
        "selection": "選定",
        "breakdown": "⑩内訳",
        "exclusion": "⑪除外",
        "check": "見積チェック",
        "cover": "①表紙",
        "quotation": "⑧見積書",
    },
    "高知ハウス": {
        "input": "☆入力",
        "selection": "選定",
        "breakdown": "⑨内訳",
        "exclusion": "⑩除外",
        "check": "見積チェック",
        "cover": "①表紙",
        "quotation": "⑧見積書",
    },
}

# ☆入力シートのセルマッピング定数
INPUT_SHEET = {
    "property_name_cell": "A5",
    "address_cell": "A6",
    "unlock_code_cell": "A7",
    "distribution_board_cell": "A8",
    "special_notes_cell": "A9",
    "data_start_row": 16,          # A行のデータ開始行
    "data_end_row": 45,            # AD行のデータ終了行
    "excluded_start_row": 49,      # LED済みセクション開始行
    "col_property_name": "B",      # 物件名
    "col_trial_area": "C",         # 試算エリア
    "col_lighting_type": "D",      # 照明種別
    "col_survey_notes": "E",       # 現調備考
    "col_construction_notes": "F", # 工事備考
    "col_fixture_category": "G",   # 器具分類②（選定リンクキー）
    "col_monthly_hours": "H",      # 月間点灯
    "col_daily_hours": "I",        # 一日点灯
    "col_operating_days": "J",     # 稼働日数
    "col_power": "K",              # 消費電力
    "col_total_quantity": "L",     # 電球数（合計）
    "col_floor_start": "M",       # 1階（M列〜V列=10階）
    "col_electricity_rate": "X",   # 電気単価
    "col_construction_price": "AE", # 工事単価
}

# ⑩内訳シート（高知ハウスは⑨内訳）のレイアウト定数
BREAKDOWN_SHEET = {
    "existing_photo_row": 7,       # 既存器具写真の行
    "led_photo_row": 14,           # LED商品写真の行
    "photo_start_col": "B",        # 写真開始列
    "photo_end_col": "U",          # 写真終了列（最大20列）
    "photo_width_px": 100,         # 写真幅 (px)
    "photo_height_px": 90,         # 写真高さ (px)
    # ☆入力のRow 16-35が⑩内訳のB-U列に対応
    # B列=Row16(A), C列=Row17(B), ..., U列=Row35(T)
    "input_row_offset": 14,        # ☆入力行番号 - ⑩内訳列番号オフセット
}

# ⑪除外シート（高知ハウスは⑩除外）のレイアウト定数
EXCLUSION_SHEET = {
    # 12ブロック: 左6つ(B列) + 右6つ(I列)
    # 各ブロックは4行間隔: Row 4, 8, 12, 16, 20, 24
    "blocks": [
        # (写真列, 写真行, 理由列, 理由行, アドバイス列, アドバイス行)
        # 左側ブロック（B列系）
        {"photo_col": "B", "photo_row": 4, "block_height": 4},
        {"photo_col": "B", "photo_row": 8, "block_height": 4},
        {"photo_col": "B", "photo_row": 12, "block_height": 4},
        {"photo_col": "B", "photo_row": 16, "block_height": 4},
        {"photo_col": "B", "photo_row": 20, "block_height": 4},
        {"photo_col": "B", "photo_row": 24, "block_height": 4},
        # 右側ブロック（I列系）
        {"photo_col": "I", "photo_row": 4, "block_height": 4},
        {"photo_col": "I", "photo_row": 8, "block_height": 4},
        {"photo_col": "I", "photo_row": 12, "block_height": 4},
        {"photo_col": "I", "photo_row": 16, "block_height": 4},
        {"photo_col": "I", "photo_row": 20, "block_height": 4},
        {"photo_col": "I", "photo_row": 24, "block_height": 4},
    ],
    "photo_width_px": 78,          # 除外器具写真幅 (px)
    "photo_height_px": 80,         # 除外器具写真高さ (px)
}

# 選定シートの列マッピング
SELECTION_SHEET = {
    "data_start_row": 3,
    "col_photo_1": "A",
    "col_photo_2": "B",
    "col_key": "C",               # 電球種別（リンクキー）
    "col_lighting_color": "D",
    "col_fixture_color": "E",
    "col_fixture_size": "F",
    "col_power": "G",
    "col_lumens": "H",
    "col_list_price": "I",
    "col_purchase_price": "J",
    "col_waterproof": "K",
    "col_bulb_type": "L",
    "col_manufacturer": "M",
    "col_watt_equiv": "N",
    "col_model": "O",
    "col_model_price": "P",
    "col_model_purchase": "Q",
    # R-Z: 追加型番・価格
    "col_power_detail": "AA",
    "col_lumens_detail": "AB",
    "col_material": "AC",
    "col_color_options": "AD",
    "col_lighting_color_options": "AE",
    "col_lifespan": "AF",
    "col_replacement_method": "AG",
    "col_socket": "AH",
    "col_hp": "AI",
    "col_notes": "AJ",
}
