"""LED見積シミュレーション作成 — Streamlit アプリ

現調データ（テキスト/写真/PDF）を入力 → LED選定 → Excel見積を自動生成。
起動: streamlit run app.py
"""

import os
import sys
import json
import tempfile
import zipfile
import base64
from datetime import datetime
from pathlib import Path

import streamlit as st

# src/ をインポートパスに追加
ROOT_DIR = Path(__file__).parent
sys.path.insert(0, str(ROOT_DIR / "src"))

# Streamlit上ではclaude_guardを本番モードに（承認プロンプトなし）
os.environ["CLAUDE_ENV"] = "production"

# APIキー: Streamlit Cloud secrets → .env → 環境変数 の順で取得
try:
    from dotenv import load_dotenv
    load_dotenv(ROOT_DIR / ".env", override=True)
except ImportError:
    pass

# Streamlit secrets があればそちらを優先（Cloud環境用）
try:
    if "ANTHROPIC_API_KEY" in st.secrets:
        os.environ["ANTHROPIC_API_KEY"] = st.secrets["ANTHROPIC_API_KEY"]
except Exception:
    pass

# ディレクトリ定数
LINEUP_DIR = ROOT_DIR / "ラインナップ表"
TEMPLATE_DIR = ROOT_DIR / "見積りテンプレート"
OUTPUT_DIR = ROOT_DIR / "output"
CONFIG_DIR = ROOT_DIR / "config"

# テンプレート一覧
TEMPLATES = [
    "田村基本形",
    "ウスイホーム",
    "オクスト",
    "クラスコ",
    "タカラ",
    "ニッショー",
    "マンション経営保障",
    "丸八アセットマネジメント",
    "スマイルサポート",
    "ライフサポート",
    "高知ハウス",
]


# ============================================================
# 管理会社ルール読み込み
# ============================================================
@st.cache_data
def load_management_company_rules():
    """config/management_company_rules.json から管理会社ルールを読み込む"""
    rules_path = CONFIG_DIR / "management_company_rules.json"
    if not rules_path.exists():
        return {}
    with open(rules_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    # name → ルール辞書のマッピング
    return {c["name"]: c for c in data.get("companies", [])}

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp"}
PDF_EXTENSIONS = {".pdf"}


# ============================================================
# セッション状態の初期化
# ============================================================

def init_session():
    """セッション状態の初期化"""
    if "ai_excel_history" not in st.session_state:
        st.session_state.ai_excel_history = []  # [{name, bytes, timestamp, property_name}, ...]

    # 3ステップUI用の状態
    if "current_step" not in st.session_state:
        st.session_state.current_step = 1  # 1=入力+器具編集, 2=写真紐付け, 3=LED選定+出力
    if "step1_result" not in st.session_state:
        st.session_state.step1_result = None  # Step1Result
    if "confirmed_fixtures" not in st.session_state:
        st.session_state.confirmed_fixtures = []  # 編集済み器具データ (list[dict])
    if "confirmed_excluded" not in st.session_state:
        st.session_state.confirmed_excluded = []  # 除外器具データ
    if "confirmed_photos" not in st.session_state:
        st.session_state.confirmed_photos = {}  # {行ラベル: [写真パス]}
    if "photo_suggestions" not in st.session_state:
        st.session_state.photo_suggestions = []  # AI推定結果
    if "step_config" not in st.session_state:
        st.session_state.step_config = {}  # テンプレート名等の設定


def save_ai_excel_to_session(excel_bytes: bytes, filename: str, property_name: str):
    """AI生成Excelをセッション履歴に保存（最新10件）"""
    st.session_state.ai_excel_history.insert(0, {
        "name": filename,
        "bytes": excel_bytes,
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "property_name": property_name,
    })
    # 最新10件のみ保持
    st.session_state.ai_excel_history = st.session_state.ai_excel_history[:10]


# ============================================================
# ユーティリティ関数
# ============================================================

def save_uploaded_files(uploaded_files, dest_dir: Path) -> list[Path]:
    """アップロードファイルを一時ディレクトリに保存（同名ファイルは連番で回避）"""
    saved = []
    dest_dir.mkdir(parents=True, exist_ok=True)
    for i, uf in enumerate(uploaded_files):
        stem = Path(uf.name).stem
        suffix = Path(uf.name).suffix
        path = dest_dir / f"{i+1:02d}_{stem}{suffix}"
        path.write_bytes(uf.getbuffer())
        saved.append(path)
    return saved


def extract_zips_and_files(files: list[Path], dest_dir: Path) -> tuple[list[Path], list[Path]]:
    """ZIPを解凍し、画像パスとPDFパスを分けて返す"""
    images = []
    pdfs = []
    for f in files:
        suffix = f.suffix.lower()
        if suffix == ".zip":
            extract_to = dest_dir / f.stem
            extract_to.mkdir(parents=True, exist_ok=True)
            with zipfile.ZipFile(f, "r") as zf:
                zf.extractall(extract_to)
            for item in sorted(extract_to.rglob("*")):
                if item.is_file():
                    if item.suffix.lower() in IMAGE_EXTENSIONS:
                        images.append(item)
                    elif item.suffix.lower() in PDF_EXTENSIONS:
                        pdfs.append(item)
        elif suffix in IMAGE_EXTENSIONS:
            images.append(f)
        elif suffix in PDF_EXTENSIONS:
            pdfs.append(f)
    return images, pdfs


def extract_text_from_file_ai(file_path: Path) -> str:
    """画像やPDFからAI（Claude Vision）でテキスト情報を抽出"""
    sys.path.insert(0, str(Path(__file__).parent.parent))
    from claude_guard import get_guarded_client

    client = get_guarded_client()
    suffix = file_path.suffix.lower()
    file_bytes = file_path.read_bytes()
    b64_data = base64.standard_b64encode(file_bytes).decode("utf-8")

    if suffix in PDF_EXTENSIONS:
        content_block = {
            "type": "document",
            "source": {
                "type": "base64",
                "media_type": "application/pdf",
                "data": b64_data,
            },
        }
    else:
        mime_map = {
            ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
            ".png": "image/png", ".gif": "image/gif",
            ".webp": "image/webp", ".bmp": "image/bmp",
        }
        media_type = mime_map.get(suffix, "image/jpeg")
        content_block = {
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": media_type,
                "data": b64_data,
            },
        }

    prompt_text = (
        "この画像/文書に記載されている照明器具の設置状況を、"
        "そのまま正確にテキストとして書き起こしてください。\n"
        "フロア名、器具の種類、ワット数、台数、状態（LED更新済み等）"
        "などの情報をできるだけ忠実に抽出してください。\n"
        "余計な説明は不要です。抽出したデータのみを出力してください。"
    )

    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4096,
        messages=[{
            "role": "user",
            "content": [
                content_block,
                {"type": "text", "text": prompt_text},
            ],
        }],
    )
    return response.content[0].text


# ============================================================
# 見積作成の処理関数
# ============================================================

# ============================================================
# 見積作成（3ステップUI）
# ============================================================

def _reset_steps():
    """ステップをリセットして最初に戻す"""
    # 一時フォルダのクリーンアップ
    old_tmpdir = st.session_state.get("_tmpdir_path")
    if old_tmpdir:
        import shutil
        shutil.rmtree(old_tmpdir, ignore_errors=True)
        st.session_state["_tmpdir_path"] = None

    st.session_state.current_step = 1
    st.session_state.step1_result = None
    st.session_state.confirmed_fixtures = []
    st.session_state.confirmed_excluded = []
    st.session_state.confirmed_photos = {}
    st.session_state.photo_suggestions = []
    st.session_state.step_config = {}
    # LED選定・編集・AI紐付け関連
    for k in ["led_match_results", "led_candidates", "user_led_overrides",
              "fixtures_edit_done", "_photo_ai_done"]:
        st.session_state.pop(k, None)
    # photo_idx_* のstate もクリア
    for k in list(st.session_state.keys()):
        if k.startswith("photo_idx_") or k.startswith("pcheck_"):
            del st.session_state[k]


def _show_step_indicator():
    """現在のステップをプログレス表示"""
    step = st.session_state.current_step
    steps = {1: "現調データ入力", 2: "写真の紐付け", 3: "LED選定・出力"}
    cols = st.columns(len(steps))
    for col, (num, label) in zip(cols, steps.items()):
        if num == step:
            col.markdown(f"**:blue[Step {num}. {label}]**")
        elif num < step:
            col.markdown(f"~~Step {num}. {label}~~")
        else:
            col.markdown(f"Step {num}. {label}")
    st.divider()


def tab_estimate():
    """見積作成タブ（3ステップUI）"""
    step = st.session_state.current_step

    _show_step_indicator()
    if step > 1:
        if st.button("最初からやり直す", type="secondary"):
            _reset_steps()
            st.rerun()

    if step == 1:
        _step1_input()
    elif step == 2:
        _step2_photo_assignment()
    elif step == 3:
        _step3_led_selection()


# ============================================================
# Step 1: 現調データ入力 + 器具情報の確認・編集
# ============================================================

def _step1_input():
    """現調データ入力 → OCR → 表形式で器具編集"""
    import pandas as pd

    st.subheader("Step 1: 現調データ入力")

    # --- 管理会社ルール読み込み ---
    company_rules = load_management_company_rules()
    company_names = ["（指定なし）"] + sorted(company_rules.keys())

    # --- 管理会社選択 ---
    selected_company = st.selectbox("管理会社", company_names)
    rule = company_rules.get(selected_company)

    # --- テンプレート選択 ---
    if rule:
        auto_template = rule.get("テンプレート", "田村基本形")
        template_idx = TEMPLATES.index(auto_template) if auto_template in TEMPLATES else 0
        template_name = st.selectbox("テンプレート", TEMPLATES, index=template_idx)
    else:
        template_name = st.selectbox("テンプレート", TEMPLATES)

    # --- 管理会社ルール表示 ---
    if rule:
        cols = st.columns(5)
        cols[0].metric("表紙", rule.get("表紙", "—"))
        cols[1].metric("紹介料", f"{rule['紹介料']}%" if rule.get("紹介料") else "なし")
        cols[2].metric("交換費", f"{rule['交換費']:,}円" if rule.get("交換費") is not None else "—")
        cols[3].metric("幹旋料", f"{rule['幹旋料']}%" if rule.get("幹旋料") else "なし")
        cols[4].metric("電気単価", f"{rule['電気単価']}円" if rule.get("電気単価") else "—")
        if rule.get("共有方法"):
            st.caption(f"共有方法: {rule['共有方法']} / データ形式: {rule.get('共有データ', '—')}")
        if rule.get("備考"):
            st.warning(f"注意: {rule['備考']}")

    # --- 物件情報 ---
    col1, col2 = st.columns(2)
    with col1:
        property_name = st.text_input("物件名", placeholder="例: リラハイツ")
    with col2:
        address = st.text_input("住所（省略可）", placeholder="例: 東京都渋谷区...")

    st.divider()

    # --- 現調シートのアップロード ---
    st.subheader("現調シート")
    uploaded_files = st.file_uploader(
        "現調シートをアップロード（複数可。画像/PDF/Excel対応）",
        type=["jpg", "jpeg", "png", "gif", "webp", "bmp", "zip", "pdf", "xlsx"],
        accept_multiple_files=True,
        key="survey_files",
    )
    if uploaded_files:
        parts = []
        for ext_label, exts in [("画像", IMAGE_EXTENSIONS | {".zip"}), ("PDF", {".pdf"}), ("Excel", {".xlsx"})]:
            count = sum(1 for f in uploaded_files if Path(f.name).suffix.lower() in exts)
            if count:
                parts.append(f"{ext_label}: {count}件")
        st.caption(" / ".join(parts))

    # --- テキスト入力（併用可） ---
    text_input = st.text_area(
        "テキスト入力（現調シートと併用可。器具情報をテキストで入力）",
        height=120,
        placeholder="例:\nA 階段 逆富士2灯 FHF32 各階2台\nB 廊下 ダウンライト 白熱60W 1階4台 2階4台",
        key="text_input_area",
    )

    st.divider()

    # --- バリデーション ---
    has_input = bool(uploaded_files) or bool(text_input.strip())
    can_run = bool(property_name) and has_input
    if not property_name:
        st.info("物件名を入力してください")
    elif not has_input:
        st.info("現調シートをアップロードするか、テキストで器具情報を入力してください")

    # --- 読み取り開始ボタン ---
    if st.button("読み取り開始", type="primary", disabled=not can_run, use_container_width=True):
        # 前回の一時フォルダがあれば削除
        old_tmpdir = st.session_state.get("_tmpdir_path")
        if old_tmpdir:
            import shutil
            shutil.rmtree(old_tmpdir, ignore_errors=True)

        photo_dir = None
        if uploaded_files:
            # 一時フォルダ作成
            tmpdir = Path(tempfile.mkdtemp(prefix="led_estimate_"))
            st.session_state["_tmpdir_path"] = str(tmpdir)

            saved = save_uploaded_files(uploaded_files, tmpdir / "uploads")
            images, pdfs = extract_zips_and_files(saved, tmpdir / "extracted")

            if images:
                import shutil
                photo_dir = tmpdir / "survey_sheets"
                photo_dir.mkdir(exist_ok=True)
                for i, p in enumerate(images):
                    # ファイル名重複を防止するため連番を付与
                    dest = photo_dir / f"{i+1:02d}_{p.name}"
                    shutil.copy2(p, dest)

        try:
            with st.status("読み取り中...", expanded=True) as status:
                from pipeline import run_step1_ocr
                api_key = os.environ.get("ANTHROPIC_API_KEY")
                result = run_step1_ocr(
                    survey_dir=photo_dir,
                    property_name=property_name,
                    api_key=api_key,
                    text_input=text_input.strip() if text_input.strip() else None,
                )
                status.update(label="読み取り完了", state="complete")

            # 結果をセッションに保存
            st.session_state.step1_result = result
            st.session_state.step_config = {
                "template_name": template_name,
                "property_name": property_name,
                "address": address,
                "selected_company": selected_company,
            }

            # OCR raw結果から直接器具データを構築（パーサーの再分類によるズレを回避）
            # raw結果には _page が正確に付与されている
            raw_all = (
                result.ocr_result.get("fixtures", [])
                + result.ocr_result.get("excluded_fixtures", [])
            )
            # 元の順序を記録（ページ内の原本順を保持するため）
            for idx, raw in enumerate(raw_all):
                raw["_order"] = idx

            # パーサーと同じ再分類ロジック: bulb_typeにLEDが含まれると除外扱い
            fixtures_data = []
            excluded_data = []
            for raw in raw_all:
                # 階別数量のパース
                raw_fq = raw.get("floor_quantities", {})
                fq = {}
                for k, v in raw_fq.items():
                    # "1F"→1, "2P"→2, "1"→1 等の変換
                    import re as _re
                    m = _re.search(r'(\d+)', str(k))
                    if m:
                        try:
                            fq[int(m.group(1))] = int(v)
                        except (ValueError, TypeError):
                            fq[int(m.group(1))] = 0

                bulb = raw.get("bulb_type", "")
                is_excl = raw.get("is_excluded", False)
                excl_reason = raw.get("exclusion_reason", "")
                if not is_excl and "LED" in bulb.upper():
                    is_excl = True
                    excl_reason = excl_reason or "LED済み"

                fix_dict = {
                    "row_label": raw.get("row_label", ""),
                    "location": raw.get("location", ""),
                    "fixture_type": raw.get("fixture_type", ""),
                    "bulb_type": bulb,
                    "floor_quantities": fq,
                    "power_w": float(raw.get("power_w", 0) or 0),
                    "daily_hours": float(raw.get("daily_hours", 0) or 0),
                    "color_temp": raw.get("color_temp", "白") or "白",
                    "is_excluded": is_excl,
                    "confidence": raw.get("confidence", "high"),
                    "survey_notes": raw.get("notes", ""),
                    "construction_notes": raw.get("construction_notes", ""),
                    "fixture_size": raw.get("fixture_size", ""),
                    "_page": raw.get("_page", 1),
                    "_order": raw.get("_order", 0),
                }
                if is_excl:
                    fix_dict["exclusion_reason"] = excl_reason
                    excluded_data.append(fix_dict)
                else:
                    fixtures_data.append(fix_dict)

            st.session_state.confirmed_fixtures = fixtures_data
            st.session_state.confirmed_excluded = excluded_data
            st.rerun()

        except Exception as e:
            st.error(f"読み取りエラー: {e}")
            import traceback
            st.code(traceback.format_exc())
            return

    # ========== 器具情報の表形式編集（OCR結果がある場合に表示） ==========

    fixtures = st.session_state.confirmed_fixtures
    excluded = st.session_state.confirmed_excluded

    if not fixtures and not excluded:
        return

    st.divider()
    st.subheader("読み取り結果の確認")
    st.caption("現調シートの読み取り結果です。編集はStep 3（写真紐付け後）で行えます。")

    # 全器具を統合し、ページ順→原本順（現調シートの記載順）でソート
    all_fixtures = sorted(
        list(fixtures) + list(excluded),
        key=lambda x: (x.get("_page", 1), x.get("_order", 0)),
    )

    # 「〃」を直上の値で展開（ソート後＝現調シートと同じ順序で処理）
    ditto_marks = {"〃", "//", "''", "″", '"', "〃"}
    for i, fix in enumerate(all_fixtures):
        if i == 0:
            continue
        prev = all_fixtures[i - 1]
        for key in ["location", "fixture_type", "bulb_type", "fixture_size"]:
            val = str(fix.get(key, "")).strip()
            if val in ditto_marks:
                fix[key] = prev.get(key, "")

    # 最大階数を算出
    max_floor = 1
    for fix in all_fixtures:
        fq = fix.get("floor_quantities", {})
        if fq:
            max_floor = max(max_floor, max(int(k) for k in fq.keys()))
    max_floor = max(max_floor, 3)

    # DataFrame構築（読み取り専用）
    rows = []
    for fix in all_fixtures:
        fq = fix.get("floor_quantities", {})
        row = {
            "P": fix.get("_page", 1),
            "行": fix.get("row_label", ""),
            "設置場所": fix.get("location", ""),
            "器具種別": fix.get("fixture_type", ""),
            "電球種別": fix.get("bulb_type", ""),
            "器具サイズ": fix.get("fixture_size", ""),
        }
        for f_num in range(1, max_floor + 1):
            row[f"{f_num}階"] = int(fq.get(f_num, fq.get(str(f_num), 0)))
        row["除外"] = "除外" if fix.get("is_excluded") else ""
        rows.append(row)

    df = pd.DataFrame(rows)
    st.dataframe(df, use_container_width=True, hide_index=True)

    st.info(f"読み取り結果: {len(fixtures)}件の器具 + {len(excluded)}件の除外器具")

    # 確定時は〃展開済みのデータをsession_stateに反映
    if st.button("次へ（写真の紐付けへ）", type="primary", use_container_width=True):
        # 〃展開済みのデータでsession_stateを更新
        new_fixtures = [f for f in all_fixtures if not f.get("is_excluded")]
        new_excluded = [f for f in all_fixtures if f.get("is_excluded")]
        st.session_state.confirmed_fixtures = new_fixtures
        st.session_state.confirmed_excluded = new_excluded
        st.session_state.current_step = 2
        st.rerun()


# ============================================================
# Step 2: 写真の紐付け
# ============================================================

def _step2_photo_assignment():
    """器具写真アップ → 各行への紐付け"""
    st.subheader("Step 2: 写真の紐付け")

    fixtures = st.session_state.confirmed_fixtures
    excluded = st.session_state.confirmed_excluded
    all_fixtures = list(fixtures) + list(excluded)

    if not all_fixtures:
        st.warning("器具情報がありません。Step 1で器具を確認してください。")
        if st.button("Step 1に戻る", use_container_width=True):
            st.session_state.current_step = 1
            st.rerun()
        return

    # --- 写真アップロード ---
    st.caption("器具写真をアップロードしてください（複数可）")
    uploaded_photos = st.file_uploader(
        "器具写真をアップロード",
        type=["jpg", "jpeg", "png", "gif", "webp", "bmp", "zip"],
        accept_multiple_files=True,
        key="fixture_photo_files",
    )

    # アップロードされた写真を一時フォルダに保存
    fixture_photos = []
    if uploaded_photos:
        tmpdir_str = st.session_state.get("_tmpdir_path")
        if not tmpdir_str:
            tmpdir = Path(tempfile.mkdtemp(prefix="led_estimate_"))
            st.session_state["_tmpdir_path"] = str(tmpdir)
            tmpdir_str = str(tmpdir)

        photo_dir = Path(tmpdir_str) / "fixture_photos"
        photo_dir.mkdir(exist_ok=True)

        for uf in uploaded_photos:
            dest = photo_dir / uf.name
            dest.write_bytes(uf.getbuffer())
            if dest.suffix.lower() in IMAGE_EXTENSIONS:
                fixture_photos.append(dest)
            elif dest.suffix.lower() == ".zip":
                # ZIPを解凍
                import zipfile as zf
                try:
                    with zf.ZipFile(dest) as z:
                        z.extractall(photo_dir / "zip_extracted")
                    for p in (photo_dir / "zip_extracted").rglob("*"):
                        if p.suffix.lower() in IMAGE_EXTENSIONS:
                            fixture_photos.append(p)
                except Exception:
                    pass

    if not fixture_photos:
        st.info("器具写真をアップロードしてください。写真なしではStep 3に進めません。")
        # 戻るボタン
        if st.button("戻る", use_container_width=True, key="back_step2_empty"):
            st.session_state.current_step = 1
            st.rerun()
        return

    st.caption(f"{len(fixture_photos)}枚の写真がアップロードされました")

    # --- AI写真マッチングの自動実行 ---
    suggestions = st.session_state.get("photo_suggestions", [])
    ai_run_key = "_photo_ai_done"

    if not st.session_state.get(ai_run_key):
        try:
            with st.status("写真をAIで自動紐付け中...", expanded=True) as status:
                from pipeline import run_step2_photo_suggest
                result = st.session_state.step1_result
                ocr_fixtures = result.ocr_result.get("fixtures", []) if result else []
                api_key = os.environ.get("ANTHROPIC_API_KEY")
                suggestions = run_step2_photo_suggest(
                    fixture_photos, ocr_fixtures, api_key=api_key,
                )
                st.session_state.photo_suggestions = suggestions
                st.session_state[ai_run_key] = True
                status.update(label="AI紐付け完了（下記で確認・修正してください）", state="complete")
            st.rerun()
        except Exception as e:
            st.warning(f"AI自動紐付けに失敗しました。手動で紐付けてください: {e}")
            st.session_state[ai_run_key] = True

    # AI推定結果をデフォルト選択に変換
    default_selections = {}
    for sug in suggestions:
        label = sug.get("row_label", "")
        idx = sug.get("photo_index")
        if label and idx is not None and idx < len(fixture_photos):
            default_selections.setdefault(label, set()).add(idx)

    if suggestions:
        matched = sum(1 for f in all_fixtures if f["row_label"] in default_selections)
        st.caption(f"AI紐付け: {len(suggestions)}件の推定 / {matched}/{len(all_fixtures)}行に適用")
    else:
        st.caption("AI紐付け: 結果なし（手動で紐付けてください）")

    st.divider()

    # --- 表形式 + 写真列 ---
    photo_map = {}

    # ヘッダー行
    hdr = st.columns([0.3, 0.3, 1, 1, 0.8, 0.4, 0.4, 0.4, 1.5])
    for col, txt in zip(hdr, ["P", "行", "設置場所", "器具種別", "電球種別", "1階", "2階", "3階", "写真"]):
        col.markdown(f"**{txt}**")
    st.divider()

    for fix_idx, fix in enumerate(all_fixtures):
        label = fix["row_label"]
        page = fix.get("_page", 1)
        fq = fix.get("floor_quantities", {})

        # 現在選択中の写真インデックス
        state_key = f"photo_idx_{fix_idx}"
        if state_key not in st.session_state:
            st.session_state[state_key] = list(default_selections.get(label, set()))
        selected_indices = st.session_state[state_key]

        # 行表示
        row_cols = st.columns([0.3, 0.3, 1, 1, 0.8, 0.4, 0.4, 0.4, 1.5])
        row_cols[0].write(page)
        row_cols[1].write(label)
        row_cols[2].write(fix.get("location", ""))
        row_cols[3].write(fix.get("fixture_type", ""))
        row_cols[4].write(fix.get("bulb_type", ""))
        row_cols[5].write(fq.get(1, fq.get("1", 0)))
        row_cols[6].write(fq.get(2, fq.get("2", 0)))
        row_cols[7].write(fq.get(3, fq.get("3", 0)))

        # 写真列: サムネイル表示
        with row_cols[8]:
            if selected_indices:
                img_cols = st.columns(min(len(selected_indices), 3))
                for j, pi in enumerate(selected_indices):
                    if pi < len(fixture_photos):
                        with img_cols[j % len(img_cols)]:
                            st.image(str(fixture_photos[pi]), width=60)
            else:
                st.caption("—")

        # 写真変更用expander（行の下に）
        with st.expander(f"写真を変更 ({label})", expanded=False):
            num_cols = min(len(fixture_photos), 6)
            cols = st.columns(num_cols)
            for pi, photo in enumerate(fixture_photos):
                with cols[pi % num_cols]:
                    is_selected = pi in selected_indices
                    check = st.checkbox(
                        f"#{pi+1}",
                        value=is_selected,
                        key=f"pcheck_{fix_idx}_{pi}",
                    )
                    st.image(str(photo), width=60)
                    if check and pi not in selected_indices:
                        selected_indices.append(pi)
                    elif not check and pi in selected_indices:
                        selected_indices.remove(pi)
            st.session_state[state_key] = selected_indices

        photo_map[label] = [fixture_photos[i] for i in selected_indices if i < len(fixture_photos)]

    # --- 確定ボタン ---
    st.divider()
    col1, col2 = st.columns([3, 1])
    with col1:
        if st.button("写真の紐付けを確定して次へ", type="primary", use_container_width=True):
            st.session_state.confirmed_photos = {
                label: [str(p) for p in paths]
                for label, paths in photo_map.items()
            }
            st.session_state.current_step = 3
            st.rerun()
    with col2:
        if st.button("戻る", use_container_width=True, key="back_step2"):
            st.session_state.photo_suggestions = []
            st.session_state.pop("_photo_ai_done", None)
            for k in list(st.session_state.keys()):
                if k.startswith("photo_idx_") or k.startswith("pcheck_"):
                    del st.session_state[k]
            st.session_state.current_step = 1
            st.rerun()


# ============================================================
# Step 3: LED選定 + Excel出力
# ============================================================

def _build_survey_from_session():
    """session_stateの確定データからSurveyDataを構築（プレビュー・Excel生成共通）"""
    from models import SurveyData, PropertyInfo, ExistingFixture as EF, FloorQuantities

    fixtures = st.session_state.confirmed_fixtures
    excluded = st.session_state.confirmed_excluded
    photo_map_raw = st.session_state.confirmed_photos
    config = st.session_state.step_config
    result = st.session_state.step1_result

    property_info = PropertyInfo(
        name=config.get("property_name", ""),
        address=config.get("address", ""),
    )
    if result and result.survey:
        orig_info = result.survey.property_info
        property_info.unlock_code = orig_info.unlock_code
        property_info.distribution_board = orig_info.distribution_board
        property_info.special_notes = orig_info.special_notes

    survey_fixtures = []
    for fix in fixtures:
        photo_paths_str = photo_map_raw.get(fix["row_label"], [])
        photo_paths = [Path(p) for p in photo_paths_str]
        ef = EF(
            row_label=fix["row_label"],
            location=fix["location"],
            fixture_type=fix["fixture_type"],
            fixture_size=fix.get("fixture_size", ""),
            bulb_type=fix["bulb_type"],
            quantities=FloorQuantities(floors={
                int(k): int(v) for k, v in fix.get("floor_quantities", {}).items()
            }),
            power_consumption_w=float(fix.get("power_w", 0)),
            daily_hours=float(fix.get("daily_hours", 0)),
            color_temp=fix.get("color_temp", ""),
            survey_notes=fix.get("survey_notes", ""),
            construction_notes=fix.get("construction_notes", ""),
            photo_paths=photo_paths,
        )
        survey_fixtures.append(ef)

    excl_fixtures = []
    for ex in excluded:
        ef = EF(
            row_label=ex.get("row_label", ""),
            location=ex.get("location", ""),
            fixture_type=ex.get("fixture_type", ""),
            bulb_type=ex.get("bulb_type", ""),
            is_excluded=True,
            exclusion_reason=ex.get("exclusion_reason", "除外"),
        )
        excl_fixtures.append(ef)

    survey = SurveyData(
        property_info=property_info,
        fixtures=survey_fixtures,
        excluded_fixtures=excl_fixtures,
    )

    return survey


def _step3_led_selection():
    """LED選定結果の確認・変更 & Excel生成"""
    st.subheader("Step 3: LED選定・Excel出力")

    fixtures = st.session_state.confirmed_fixtures
    config = st.session_state.step_config

    if not fixtures and not st.session_state.confirmed_excluded:
        st.warning("器具情報がありません")
        return

    # ========== 器具編集フェーズ（写真付き） ==========

    fixtures_confirmed = st.session_state.get("fixtures_edit_done", False)

    if not fixtures_confirmed:
        import pandas as pd

        st.caption("写真を確認しながら器具情報を編集してください。除外チェックでLED済み器具を除外できます。")

        photo_map_raw = st.session_state.confirmed_photos
        all_fix = sorted(
            list(fixtures) + list(st.session_state.confirmed_excluded),
            key=lambda x: (x.get("_page", 1), x.get("_order", 0)),
        )

        # --- 上部: 編集可能な表 ---
        max_floor = 1
        for fix in all_fix:
            fq = fix.get("floor_quantities", {})
            if fq:
                max_floor = max(max_floor, max(int(k) for k in fq.keys()))
        max_floor = max(max_floor, 3)

        rows = []
        for fix in all_fix:
            fq = fix.get("floor_quantities", {})
            row = {
                "P": fix.get("_page", 1),
                "行": fix.get("row_label", ""),
                "設置場所": fix.get("location", ""),
                "器具種別": fix.get("fixture_type", ""),
                "電球種別": fix.get("bulb_type", ""),
                "器具サイズ": fix.get("fixture_size", ""),
            }
            for f_num in range(1, max_floor + 1):
                row[f"{f_num}階"] = int(fq.get(f_num, fq.get(str(f_num), 0)))
            row["除外"] = fix.get("is_excluded", False)
            rows.append(row)

        df = pd.DataFrame(rows)
        column_config = {
            "P": st.column_config.NumberColumn("P", width="small"),
            "行": st.column_config.TextColumn("行", width="small"),
            "設置場所": st.column_config.TextColumn("設置場所", width="medium"),
            "器具種別": st.column_config.TextColumn("器具種別", width="medium"),
            "電球種別": st.column_config.TextColumn("電球種別", width="medium"),
            "器具サイズ": st.column_config.TextColumn("サイズ", width="small"),
            "除外": st.column_config.CheckboxColumn("除外", width="small"),
        }
        for f_num in range(1, max_floor + 1):
            column_config[f"{f_num}階"] = st.column_config.NumberColumn(
                f"{f_num}階", min_value=0, step=1, width="small",
            )

        edited_df = st.data_editor(
            df, column_config=column_config,
            num_rows="dynamic", use_container_width=True,
            key="fixture_editor_step3",
        )

        # --- 下部: 写真一覧（行ごと） ---
        st.divider()
        st.subheader("紐付け済み写真")
        for fix in all_fix:
            label = fix.get("row_label", "")
            photos = photo_map_raw.get(label, [])
            if not photos:
                continue
            st.markdown(f"**{label}: {fix.get('location', '')} — {fix.get('fixture_type', '')}**")
            cols = st.columns(min(len(photos), 4))
            for j, p in enumerate(photos):
                with cols[j % len(cols)]:
                    st.image(p, width=150)

        # --- 確定ボタン ---
        st.divider()
        if st.button("器具情報を確定してLED選定へ", type="primary", use_container_width=True):
            new_fixtures = []
            new_excluded = []

            def _seq_label(n):
                if n < 26:
                    return chr(65 + n)
                return chr(64 + n // 26) + chr(65 + n % 26)

            seq_idx = 0
            for _, row in edited_df.iterrows():
                orig_label = str(row.get("行", "")).strip()
                if not orig_label:
                    continue

                label = _seq_label(seq_idx)
                seq_idx += 1

                floor_q = {}
                for f_num in range(1, max_floor + 1):
                    val = int(row.get(f"{f_num}階", 0))
                    if val > 0:
                        floor_q[f_num] = val

                fix_dict = {
                    "row_label": label,
                    "location": str(row.get("設置場所", "")),
                    "fixture_type": str(row.get("器具種別", "")),
                    "bulb_type": str(row.get("電球種別", "")),
                    "fixture_size": str(row.get("器具サイズ", "")),
                    "floor_quantities": floor_q,
                    "power_w": 0,
                    "daily_hours": 0,
                    "color_temp": "白",
                    "is_excluded": bool(row.get("除外", False)),
                }

                if fix_dict["is_excluded"]:
                    fix_dict["exclusion_reason"] = "除外"
                    new_excluded.append(fix_dict)
                else:
                    new_fixtures.append(fix_dict)

            st.session_state.confirmed_fixtures = new_fixtures
            st.session_state.confirmed_excluded = new_excluded
            st.session_state.fixtures_edit_done = True
            st.rerun()

        # 戻るボタン
        st.divider()
        if st.button("戻る", use_container_width=True, key="back_step3_edit"):
            st.session_state.current_step = 2
            st.rerun()
        return  # 編集未完了ではここで終了

    # ========== LED選定フェーズ（編集完了後） ==========

    # --- 器具情報サマリー ---
    with st.expander(f"確定済み器具: {len(fixtures)}件", expanded=False):
        for fix in fixtures:
            floor_q = fix.get("floor_quantities", {})
            total = sum(floor_q.values())
            st.caption(
                f"{fix['row_label']}: {fix['location']} — {fix['fixture_type']} "
                f"({fix['bulb_type']}, {total}台)"
            )

    # ========== Phase A: LED選定プレビュー ==========

    has_preview = "led_match_results" in st.session_state

    if not has_preview:
        st.caption("AIがラインナップ表から最適なLED商品を選定します。結果を確認してから変更も可能です。")
        if st.button("LED選定プレビュー", type="primary", use_container_width=True):
            try:
                with st.status("LED選定中...", expanded=True) as status:
                    survey = _build_survey_from_session()
                    from pipeline import run_step3_preview
                    matches, candidates_map = run_step3_preview(
                        fixtures=survey.fixtures,
                        lineup_dir=LINEUP_DIR,
                    )
                    st.session_state.led_match_results = matches
                    st.session_state.led_candidates = candidates_map
                    status.update(label="LED選定完了", state="complete")
                st.rerun()
            except Exception as e:
                st.error(f"LED選定エラー: {e}")
                import traceback
                st.code(traceback.format_exc())

    # ========== Phase B: 選定結果表示 + 手動変更 ==========

    if has_preview:
        matches = st.session_state.led_match_results
        candidates_map = st.session_state.get("led_candidates", {})

        st.caption("AI選定結果を確認してください。変更したい行はプルダウンから別の商品を選べます。")

        # 行ラベル → MatchResult のマップ
        match_by_label = {}
        for m in matches:
            if not m.fixture.is_excluded:
                match_by_label[m.fixture.row_label] = m

        for fix in fixtures:
            label = fix["row_label"]
            m = match_by_label.get(label)
            if not m:
                continue

            led = m.led_product
            confidence_pct = f"{m.confidence:.0%}" if m.confidence else "---"

            # 候補リスト構築（AI選定 + 代替候補）
            candidates = candidates_map.get(label, [])

            # AI選定商品の表示名
            if led:
                ai_display = (
                    f"{led.name} ({led.purchase_price_total:,}円)"
                    if led.purchase_price_total
                    else led.name
                )
            else:
                ai_display = "該当なし"

            # selectbox用の選択肢を構築
            options_display = [f"AI選定: {ai_display}"]
            options_products = [led]  # index 0 = AI選定

            for c in candidates:
                if led and c.name == led.name:
                    continue  # AI選定と同じものはスキップ
                c_display = (
                    f"{c.name} ({c.purchase_price_total:,}円)"
                    if c.purchase_price_total
                    else c.name
                )
                options_display.append(c_display)
                options_products.append(c)

            # UI: 各器具行
            with st.container():
                col1, col2 = st.columns([1, 2])
                with col1:
                    review_mark = " *要確認" if m.needs_review else ""
                    st.markdown(
                        f"**{label}**: {fix['location']} — {fix['fixture_type']}"
                        f"  \n信頼度: {confidence_pct}{review_mark}"
                    )
                with col2:
                    selected_idx = st.selectbox(
                        f"LED商品 ({label})",
                        range(len(options_display)),
                        format_func=lambda i, od=options_display: od[i],
                        key=f"led_select_{label}",
                        label_visibility="collapsed",
                    )
                    # ユーザーがAI選定以外を選んだ場合、overridesに記録
                    if selected_idx > 0:
                        if "user_led_overrides" not in st.session_state:
                            st.session_state.user_led_overrides = {}
                        st.session_state.user_led_overrides[label] = options_products[selected_idx]
                    else:
                        # AI選定に戻した場合、overridesから削除
                        if "user_led_overrides" in st.session_state and label in st.session_state.user_led_overrides:
                            del st.session_state.user_led_overrides[label]

                if m.match_notes:
                    st.caption(f"  {m.match_notes}")
                st.divider()

        # ユーザー変更のサマリー
        overrides = st.session_state.get("user_led_overrides", {})
        if overrides:
            st.info(f"{len(overrides)}件の手動変更があります")

        # ========== Phase C: Excel生成 ==========

        col_gen, col_reset = st.columns([3, 1])
        with col_gen:
            if st.button("Excel生成", type="primary", use_container_width=True):
                try:
                    with st.status("Excel生成中...", expanded=True) as status:
                        survey = _build_survey_from_session()

                        # ユーザーの手動変更を反映
                        user_selections = st.session_state.get("user_led_overrides", None)
                        if user_selections and not user_selections:
                            user_selections = None

                        from pipeline import run_step3_generate
                        result_path = run_step3_generate(
                            survey=survey,
                            lineup_dir=LINEUP_DIR,
                            template_dir=TEMPLATE_DIR,
                            template_name=config.get("template_name", "田村基本形"),
                            user_led_selections=user_selections,
                        )
                        status.update(label="Excel生成完了", state="complete")

                    with open(result_path, "rb") as f:
                        excel_bytes = f.read()

                    save_ai_excel_to_session(
                        excel_bytes, result_path.name, config.get("property_name", ""),
                    )

                    st.download_button(
                        label=f"Excelダウンロード ({result_path.name})",
                        data=excel_bytes, file_name=result_path.name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary", use_container_width=True,
                    )

                except Exception as e:
                    st.error(f"エラー: {e}")
                    import traceback
                    st.code(traceback.format_exc())

        with col_reset:
            if st.button("再選定", use_container_width=True, key="reset_preview"):
                for k in ["led_match_results", "led_candidates", "user_led_overrides"]:
                    st.session_state.pop(k, None)
                st.rerun()

    # --- 戻るボタン ---
    st.divider()
    if st.button("戻る", use_container_width=True, key="back_step3"):
        # プレビュー結果もクリアして戻る
        for k in ["led_match_results", "led_candidates", "user_led_overrides"]:
            st.session_state.pop(k, None)
        st.session_state.current_step = 2
        st.rerun()

    # --- インラインフィードバック ---
    _show_inline_feedback()


# ============================================================
# インラインフィードバック（見積作成画面の下部）
# ============================================================

def _show_inline_feedback():
    """見積作成後のインラインフィードバック機能"""

    history = st.session_state.get("ai_excel_history", [])
    if not history:
        return

    st.divider()

    with st.expander("修正フィードバック（AI精度改善に協力する）"):
        st.caption(
            "修正済みExcelをアップロードすると、AI生成結果と自動比較して改善に役立てます。"
        )

        # AI生成Excel選択
        if len(history) == 1:
            ai_excel_bytes = history[0]["bytes"]
            ai_excel_name = history[0]["name"]
            st.caption(f"比較対象: {history[0]['property_name']} - {ai_excel_name}")
        else:
            options = [
                f"{h['property_name']} - {h['name']}（{h['timestamp']}）"
                for h in history
            ]
            selected_idx = st.selectbox(
                "比較対象のAI生成結果",
                range(len(options)),
                format_func=lambda i: options[i],
                key="fb_ai_select",
            )
            ai_excel_bytes = history[selected_idx]["bytes"]
            ai_excel_name = history[selected_idx]["name"]

        # 修正済みExcelアップロード
        correct_file = st.file_uploader(
            "修正済みExcel",
            type=["xlsx"],
            key="fb_correct_file",
        )

        if not correct_file:
            return

        # コメント入力
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            comment_reading = st.text_area(
                "読み取りコメント（任意）",
                height=80,
                placeholder="例: 3階の逆富士が40wではなく20wだった",
                key="fb_comment_reading",
            )
        with col_c2:
            comment_selection = st.text_area(
                "LED選定コメント（任意）",
                height=80,
                placeholder="例: FHF32はmyシリーズが適切",
                key="fb_comment_selection",
            )

        # 比較実行 → 結果をsession_stateに保存
        if st.button("比較実行", key="fb_compare", use_container_width=True):
            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir = Path(tmpdir)

                ai_path = tmpdir / (ai_excel_name or "ai_output.xlsx")
                ai_path.write_bytes(ai_excel_bytes)
                correct_path = tmpdir / correct_file.name
                correct_path.write_bytes(correct_file.getbuffer())

                try:
                    with st.status("比較中...", expanded=True) as status:
                        from feedback_comparator import FeedbackComparator
                        comparator = FeedbackComparator()
                        report = comparator.compare(ai_path, correct_path)
                        status.update(label="比較完了", state="complete")

                    # 比較結果をsession_stateに保存（送信ボタン用）
                    report_dict = {
                        "ai_file": report.ai_file,
                        "correct_file": report.correct_file,
                        "timestamp": datetime.now().isoformat(),
                        "property_name": report.property_name,
                        "comment": {
                            "reading": comment_reading,
                            "selection": comment_selection,
                        },
                        "summary": {
                            "total_diffs": report.total_diffs,
                            "fixture_match_rate": round(report.fixture_match_rate, 3),
                            "led_selection_match_rate": round(report.led_selection_match_rate, 3),
                            "fixtures": {"ai": report.total_fixtures_ai, "correct": report.total_fixtures_correct},
                            "excluded": {"ai": report.total_excluded_ai, "correct": report.total_excluded_correct},
                            "products": {"ai": report.total_products_ai, "correct": report.total_products_correct},
                        },
                        "header_diffs": [
                            {"cell": d.cell, "field_name": d.field_name, "severity": d.severity,
                             "ai_value": d.ai_value, "correct_value": d.correct_value}
                            for d in report.header_diffs
                        ],
                        "fixture_diffs": [
                            {"row": fd.row_label, "status": fd.status,
                             "fixture_type_ai": fd.fixture_type_ai, "fixture_type_correct": fd.fixture_type_correct,
                             "diffs": [{"field_name": d.field_name, "severity": d.severity,
                                        "ai_value": d.ai_value, "correct_value": d.correct_value} for d in fd.diffs]}
                            for fd in report.fixture_diffs if fd.status != "match"
                        ],
                        "selection_diffs": [
                            {"row": sd.row_number, "status": sd.status,
                             "diffs": [{"field_name": d.field_name, "ai_value": d.ai_value, "correct_value": d.correct_value}
                                       for d in sd.diffs]}
                            for sd in report.selection_diffs if sd.status != "match"
                        ],
                    }
                    st.session_state["fb_report_dict"] = report_dict
                    st.session_state["fb_report_display"] = {
                        "total_diffs": report.total_diffs,
                        "fixture_match_rate": report.fixture_match_rate,
                        "led_selection_match_rate": report.led_selection_match_rate,
                        "header_diffs": [
                            {"field_name": d.field_name, "cell": d.cell, "severity": d.severity,
                             "ai_value": d.ai_value, "correct_value": d.correct_value}
                            for d in report.header_diffs
                        ],
                        "fixture_issues": [
                            {"row_label": fd.row_label, "status": fd.status,
                             "fixture_type_ai": fd.fixture_type_ai, "fixture_type_correct": fd.fixture_type_correct,
                             "diffs": [{"field_name": d.field_name, "severity": d.severity,
                                        "ai_value": d.ai_value, "correct_value": d.correct_value} for d in fd.diffs]}
                            for fd in report.fixture_diffs if fd.status != "match"
                        ],
                        "selection_issues": [
                            {"row_number": sd.row_number,
                             "diffs": [{"field_name": d.field_name, "ai_value": d.ai_value, "correct_value": d.correct_value}
                                       for d in sd.diffs]}
                            for sd in report.selection_diffs if sd.status != "match"
                        ],
                    }
                    st.rerun()

                except Exception as e:
                    st.error(f"比較エラー: {e}")
                    import traceback
                    st.code(traceback.format_exc())

        # --- 送信完了メッセージ ---
        if "fb_submitted_id" in st.session_state:
            st.success(f"フィードバック送信完了しました (ID: {st.session_state['fb_submitted_id']})")
            del st.session_state["fb_submitted_id"]

        # --- 比較結果表示 & 送信（session_stateから読み出し） ---
        if "fb_report_display" in st.session_state:
            disp = st.session_state["fb_report_display"]

            # メトリクス
            m1, m2, m3 = st.columns(3)
            m1.metric("差分セル数", disp["total_diffs"])
            m2.metric("器具行一致率", f"{disp['fixture_match_rate']:.0%}")
            m3.metric("LED選定一致率", f"{disp['led_selection_match_rate']:.0%}")

            # ヘッダー差分
            if disp["header_diffs"]:
                with st.expander(f"ヘッダー差分（{len(disp['header_diffs'])}件）", expanded=True):
                    for d in disp["header_diffs"]:
                        severity_icon = {"critical": "\U0001f534", "major": "\U0001f7e0", "minor": "\U0001f7e2"}.get(d["severity"], "\u2b1c")
                        st.markdown(f"{severity_icon} **{d['field_name']}** ({d['cell']})")
                        c1, c2 = st.columns(2)
                        c1.code(f"AI: {d['ai_value']}", language=None)
                        c2.code(f"正解: {d['correct_value']}", language=None)

            # 器具行差分
            if disp["fixture_issues"]:
                with st.expander(f"器具行の差分（{len(disp['fixture_issues'])}件）", expanded=True):
                    for fd in disp["fixture_issues"]:
                        status_label = {
                            "modified": "\u270f\ufe0f 修正",
                            "missing_in_ai": "\u2795 AI欠落",
                            "extra_in_ai": "\u2796 AI余分",
                        }.get(fd["status"], fd["status"])
                        st.markdown(f"**行{fd['row_label']}** ({status_label})")
                        if fd["fixture_type_ai"] != fd["fixture_type_correct"]:
                            st.markdown(f"  器具: `{fd['fixture_type_ai']}` \u2192 `{fd['fixture_type_correct']}`")
                        for d in fd["diffs"]:
                            severity_icon = {"critical": "\U0001f534", "major": "\U0001f7e0", "minor": "\U0001f7e2"}.get(d["severity"], "\u2b1c")
                            st.markdown(f"  {severity_icon} {d['field_name']}: `{d['ai_value']}` \u2192 `{d['correct_value']}`")
                        st.markdown("---")

            # 選定シート差分
            if disp["selection_issues"]:
                with st.expander(f"選定シートの差分（{len(disp['selection_issues'])}件）"):
                    for sd in disp["selection_issues"]:
                        st.markdown(f"**行{sd['row_number']}**")
                        for d in sd["diffs"]:
                            st.markdown(f"  {d['field_name']}: `{d['ai_value']}` \u2192 `{d['correct_value']}`")

            if disp["total_diffs"] == 0:
                st.success("差分なし！AI出力と修正済みは完全一致です。")

            st.divider()

            # フィードバック送信（session_stateから読み出すので安定）
            report_dict = st.session_state["fb_report_dict"]
            feedback_json = json.dumps(report_dict, ensure_ascii=False, indent=2)
            prop_name = report_dict.get("property_name") or "unknown"
            filename = f"feedback_{prop_name}.json"

            # Google Sheets送信
            _sheets_available = False
            _sheets_error = ""
            try:
                from feedback_store import FeedbackStore
                _store = FeedbackStore.from_streamlit_secrets()
                _sheets_available = True
            except KeyError:
                _sheets_error = "Streamlit secrets に [feedback] gas_webapp_url が未設定です"
            except Exception as e:
                _sheets_error = f"フィードバック接続エラー: {e}"

            if _sheets_error:
                st.warning(f"Google Sheets 連携が無効です: {_sheets_error}")

            if _sheets_available:
                if st.button("フィードバック送信", key="fb_submit",
                             type="primary", use_container_width=True):
                    try:
                        fid = _store.submit_feedback(
                            report_dict=report_dict,
                            comment_reading=report_dict.get("comment", {}).get("reading", ""),
                            comment_selection=report_dict.get("comment", {}).get("selection", ""),
                        )
                        # LED選定差分からルールを自動記録（次回から自動適用される）
                        from pipeline import save_feedback_rule
                        for fd in report_dict.get("fixture_diffs", []):
                            for d in fd.get("diffs", []):
                                if ("LED" in d.get("field_name", "") and
                                        d.get("ai_value") and d.get("correct_value")):
                                    save_feedback_rule(
                                        fixture_type=(
                                            fd.get("fixture_type_correct")
                                            or fd.get("fixture_type_ai", "")
                                        ),
                                        wrong_selection=d["ai_value"],
                                        correct_selection=d["correct_value"],
                                    )

                        # 送信成功フラグを保存してから比較結果をクリア
                        st.session_state["fb_submitted_id"] = fid
                        del st.session_state["fb_report_dict"]
                        del st.session_state["fb_report_display"]
                        st.rerun()
                    except Exception as e:
                        st.error(f"送信エラー: {e}")

                try:
                    stats = _store.get_feedback_stats()
                    if stats.get("total_feedback"):
                        st.caption(
                            f"累計: {stats['total_feedback']}件 / "
                            f"LED選定平均一致率: {stats['avg_led_match_rate']:.0%}"
                        )
                except Exception:
                    pass

            st.download_button(
                label="フィードバックJSONダウンロード",
                data=feedback_json,
                file_name=filename,
                mime="application/json",
                use_container_width=True,
                key="fb_json_dl",
            )


# ============================================================
# メイン
# ============================================================

def main():
    st.set_page_config(page_title="LED見積作成", page_icon="\U0001f4a1", layout="centered")
    init_session()
    st.title("LED見積シミュレーション作成")

    tab_estimate()


if __name__ == "__main__":
    main()
