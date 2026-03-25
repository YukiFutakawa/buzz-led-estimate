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
    """AI生成Excel履歴をセッションに保持"""
    if "ai_excel_history" not in st.session_state:
        st.session_state.ai_excel_history = []  # [{name, bytes, timestamp, property_name}, ...]


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
    """アップロードファイルを一時ディレクトリに保存"""
    saved = []
    dest_dir.mkdir(parents=True, exist_ok=True)
    for uf in uploaded_files:
        path = dest_dir / uf.name
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

def process_text_route(text, photo_paths, property_name, address, template_name, status):
    """テキスト（+写真）→ Excel"""
    from history_text_parser import HistoryTextParser
    from pipeline import run_from_survey_data
    from sfa_client import SFAProject

    project = SFAProject(
        id="manual", name=property_name, address=address,
        unlock_info="", management_company="", memo="",
        phase_category="", phase="", survey_date="", construction_date="",
    )

    label = "AI解析中（テキスト＋写真）..." if photo_paths else "AI解析中（テキスト）..."
    status.update(label=label)
    parser = HistoryTextParser()
    survey = parser.parse(
        [text], project,
        photo_paths=photo_paths if photo_paths else None,
    )

    n_fix = len(survey.fixtures)
    n_exc = len(survey.excluded_fixtures)
    status.update(label=f"器具解析完了: {n_fix}件 (除外: {n_exc}件). LED選定中...")

    result_path = run_from_survey_data(
        survey=survey, lineup_dir=LINEUP_DIR,
        template_dir=TEMPLATE_DIR, template_name=template_name,
    )
    return result_path


def process_photo_route(photo_dir, property_name, template_name, status):
    """写真のみ → Excel"""
    from pipeline import run_pipeline

    status.update(label="AI解析中（写真）...")
    result_path = run_pipeline(
        survey_dir=photo_dir, lineup_dir=LINEUP_DIR,
        template_dir=TEMPLATE_DIR, template_name=template_name,
    )
    return result_path


# ============================================================
# タブ1: 見積作成
# ============================================================

def tab_estimate():
    """見積作成タブ"""

    # --- 管理会社ルール読み込み ---
    company_rules = load_management_company_rules()
    company_names = ["（指定なし）"] + sorted(company_rules.keys())

    # --- 管理会社選択 ---
    selected_company = st.selectbox("管理会社", company_names)
    rule = company_rules.get(selected_company)

    # --- テンプレート選択（管理会社選択時は自動設定） ---
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

    # --- ① 照明設置状況 ---
    st.subheader("① 照明設置状況")
    text_input = st.text_area(
        "テキスト入力", height=200,
        placeholder="例:\n4階\n20w非常内蔵逆富士 3台\n誘導灯 2台LED更新済み\n\n3階\n20w非常内蔵逆富士 3台",
    )

    survey_files = st.file_uploader(
        "または ファイル添付（テキスト / 画像 / PDF）",
        type=["txt", "jpg", "jpeg", "png", "gif", "webp", "bmp", "pdf"],
        accept_multiple_files=True,
        key="survey_files",
    )

    # アップロードファイルの分類表示
    ai_extract_files = []
    if survey_files:
        for sf in survey_files:
            suffix = Path(sf.name).suffix.lower()
            if suffix == ".txt":
                file_text = sf.read().decode("utf-8")
                text_input = (text_input + "\n" if text_input else "") + file_text
                st.text_area(f"\U0001f4c4 {sf.name} の内容", file_text, height=100, disabled=True)
            elif suffix in IMAGE_EXTENSIONS | PDF_EXTENSIONS:
                ai_extract_files.append(sf)
                icon = "\U0001f5bc\ufe0f" if suffix in IMAGE_EXTENSIONS else "\U0001f4ce"
                st.caption(f"{icon} {sf.name} \u2192 AI解析対象")

    st.divider()

    # --- ② 現調写真・資料 ---
    st.subheader("② 現調写真・資料")
    uploaded_photos = st.file_uploader(
        "ファイルをアップロード（複数選択可、ZIP対応）",
        type=["jpg", "jpeg", "png", "gif", "webp", "bmp", "zip", "pdf"],
        accept_multiple_files=True,
        key="photo_files",
    )
    if uploaded_photos:
        img_count = sum(1 for f in uploaded_photos if Path(f.name).suffix.lower() in IMAGE_EXTENSIONS | {".zip"})
        pdf_count = sum(1 for f in uploaded_photos if Path(f.name).suffix.lower() == ".pdf")
        parts = []
        if img_count:
            parts.append(f"\U0001f5bc\ufe0f 画像/ZIP: {img_count}件")
        if pdf_count:
            parts.append(f"\U0001f4ce PDF: {pdf_count}件")
        st.caption(" / ".join(parts))

    st.divider()

    # --- バリデーション ---
    has_text = bool(text_input and text_input.strip())
    has_ai_files = bool(ai_extract_files)
    has_photos = bool(uploaded_photos)
    can_run = bool(property_name) and (has_text or has_ai_files or has_photos)

    if not property_name:
        st.info("物件名を入力してください")
    elif not has_text and not has_ai_files and not has_photos:
        st.info("テキスト、画像、PDF、または写真を入力してください")

    # --- 実行ボタン ---
    if st.button("見積作成", type="primary", disabled=not can_run, use_container_width=True):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)

            # ①のAI解析ファイル処理（画像・PDF → テキスト抽出）
            if has_ai_files:
                ai_dir = tmpdir / "ai_files"
                ai_dir.mkdir(parents=True, exist_ok=True)
                with st.status("① ファイルからAI解析中...", expanded=True) as ai_status:
                    extracted_texts = []
                    for i, sf in enumerate(ai_extract_files):
                        ai_status.update(label=f"AI解析中: {sf.name} ({i+1}/{len(ai_extract_files)})...")
                        file_path = ai_dir / sf.name
                        file_path.write_bytes(sf.getbuffer())
                        try:
                            extracted = extract_text_from_file_ai(file_path)
                            extracted_texts.append(f"--- {sf.name} ---\n{extracted}")
                            st.write(f"\u2705 {sf.name}: 抽出完了")
                        except Exception as e:
                            st.warning(f"\u26a0\ufe0f {sf.name}: 抽出失敗 ({e})")
                    if extracted_texts:
                        ai_text = "\n\n".join(extracted_texts)
                        text_input = (text_input + "\n\n" if text_input else "") + ai_text
                        st.text_area("AI抽出結果", ai_text, height=200, disabled=True)
                    ai_status.update(label="① ファイル解析完了", state="complete")

            has_text = bool(text_input and text_input.strip())

            # ②の写真・資料処理
            photo_paths = None
            photo_dir = None
            if has_photos:
                saved = save_uploaded_files(uploaded_photos, tmpdir / "uploads")
                images, pdfs = extract_zips_and_files(saved, tmpdir / "extracted")
                photo_paths = images if images else None

                # PDFからもテキスト抽出
                if pdfs:
                    with st.status("② PDFからAI解析中...", expanded=True) as pdf_status:
                        pdf_texts = []
                        for i, pdf_path in enumerate(pdfs):
                            pdf_status.update(label=f"PDF解析中: {pdf_path.name} ({i+1}/{len(pdfs)})...")
                            try:
                                extracted = extract_text_from_file_ai(pdf_path)
                                pdf_texts.append(f"--- {pdf_path.name} ---\n{extracted}")
                                st.write(f"\u2705 {pdf_path.name}: 抽出完了")
                            except Exception as e:
                                st.warning(f"\u26a0\ufe0f {pdf_path.name}: 抽出失敗 ({e})")
                        if pdf_texts:
                            pdf_extra_text = "\n\n".join(pdf_texts)
                            text_input = (text_input + "\n\n" if text_input else "") + pdf_extra_text
                        pdf_status.update(label="② PDF解析完了", state="complete")
                    has_text = bool(text_input and text_input.strip())

                if photo_paths and not has_text:
                    import shutil
                    photo_dir = tmpdir / "photo_dir"
                    photo_dir.mkdir(exist_ok=True)
                    for p in photo_paths:
                        dest = photo_dir / p.name
                        if not dest.exists():
                            shutil.copy2(p, dest)

            try:
                with st.status("見積作成中...", expanded=True) as status:
                    if has_text:
                        result_path = process_text_route(
                            text=text_input, photo_paths=photo_paths,
                            property_name=property_name, address=address,
                            template_name=template_name, status=status,
                        )
                    elif photo_paths:
                        result_path = process_photo_route(
                            photo_dir=photo_dir, property_name=property_name,
                            template_name=template_name, status=status,
                        )
                    else:
                        st.error("処理可能なデータがありません")
                        return
                    status.update(label="見積作成完了", state="complete")

                with open(result_path, "rb") as f:
                    excel_bytes = f.read()

                # セッションに自動保存
                save_ai_excel_to_session(excel_bytes, result_path.name, property_name)

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
                        st.success(f"フィードバック送信完了 (ID: {fid})")
                        # 送信後にクリア
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
