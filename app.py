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
    import anthropic

    client = anthropic.Anthropic()
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
        property_name=property_name,
    )
    return result_path


# ============================================================
# タブ1: 見積作成
# ============================================================

def tab_estimate():
    """見積作成タブ"""

    # --- テンプレート選択 ---
    template_name = st.selectbox("テンプレート", TEMPLATES)

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

                st.caption("この結果はフィードバックタブで自動的に使用できます。")

            except Exception as e:
                st.error(f"エラー: {e}")
                import traceback
                st.code(traceback.format_exc())


# ============================================================
# タブ2: フィードバック
# ============================================================

def tab_feedback():
    """フィードバックタブ — AI出力 vs 修正済みExcelを比較"""

    st.markdown(
        "修正済みExcelをアップロードすると、"
        "AI生成結果とセルレベルで自動比較し差分レポートを作成します。"
    )

    st.divider()

    # --- AI生成Excel（システム保持） ---
    history = st.session_state.get("ai_excel_history", [])

    ai_excel_bytes = None
    ai_excel_name = None

    if history:
        st.markdown("**AI生成Excel（自動保持）**")
        # 選択肢を作成
        options = [
            f"{h['property_name']} - {h['name']}（{h['timestamp']}）"
            for h in history
        ]
        selected_idx = st.selectbox(
            "比較対象のAI生成結果を選択",
            range(len(options)),
            format_func=lambda i: options[i],
            key="fb_ai_select",
        )
        ai_excel_bytes = history[selected_idx]["bytes"]
        ai_excel_name = history[selected_idx]["name"]
        st.caption(f"\u2705 {ai_excel_name} を使用")
    else:
        st.markdown("**AI生成Excel（元）**")
        st.info("まだ見積を作成していません。見積作成タブで作成するか、手動でアップロードしてください。")
        ai_file_manual = st.file_uploader(
            "AI出力ファイル（手動アップロード）",
            type=["xlsx"],
            key="fb_ai_file_manual",
        )
        if ai_file_manual:
            ai_excel_bytes = ai_file_manual.getbuffer()
            ai_excel_name = ai_file_manual.name

    st.divider()

    # --- 修正済みExcel ---
    st.markdown("**修正済みExcel（正）**")
    correct_file = st.file_uploader(
        "人間が修正したファイル",
        type=["xlsx"],
        key="fb_correct_file",
    )

    st.divider()

    # --- コメント入力（2カテゴリ） ---
    st.markdown("**コメント（任意）**")

    comment_reading = st.text_area(
        "① 現調情報の読み取りについて",
        height=80,
        placeholder="例: 3階の逆富士が40wではなく20wだった。非常灯の数量が1台多くカウントされていた。",
        key="fb_comment_reading",
    )

    comment_selection = st.text_area(
        "② LED選定について",
        height=80,
        placeholder="例: FHF32はiDシリーズではなくmyシリーズが適切。非常灯は一体型を選定すべき。",
        key="fb_comment_selection",
    )

    # --- バリデーション ---
    can_compare = bool(ai_excel_bytes and correct_file)
    if not can_compare:
        if ai_excel_bytes and not correct_file:
            st.info("修正済みExcelをアップロードしてください")
        elif not ai_excel_bytes:
            st.info("AI生成Excelがありません。見積作成タブで作成してください")

    # --- 比較実行 ---
    if st.button("比較実行", type="primary", disabled=not can_compare, use_container_width=True):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)

            # ファイル保存
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

                # --- サマリー表示 ---
                st.subheader("比較結果")

                # メトリクス
                m1, m2, m3 = st.columns(3)
                m1.metric("差分セル数", report.total_diffs)
                m2.metric("器具行一致率", f"{report.fixture_match_rate:.0%}")
                m3.metric("LED選定一致率", f"{report.led_selection_match_rate:.0%}")

                m4, m5, m6 = st.columns(3)
                m4.metric("器具数（AI / 正解）",
                          f"{report.total_fixtures_ai} / {report.total_fixtures_correct}")
                m5.metric("除外数（AI / 正解）",
                          f"{report.total_excluded_ai} / {report.total_excluded_correct}")
                m6.metric("商品数（AI / 正解）",
                          f"{report.total_products_ai} / {report.total_products_correct}")

                # --- ヘッダー差分 ---
                if report.header_diffs:
                    with st.expander(f"ヘッダー差分（{len(report.header_diffs)}件）", expanded=True):
                        for d in report.header_diffs:
                            severity_icon = {"critical": "\U0001f534", "major": "\U0001f7e0", "minor": "\U0001f7e2"}.get(d.severity, "\u2b1c")
                            st.markdown(f"{severity_icon} **{d.field_name}** ({d.cell})")
                            c1, c2 = st.columns(2)
                            c1.code(f"AI: {d.ai_value}", language=None)
                            c2.code(f"正解: {d.correct_value}", language=None)

                # --- 器具行差分 ---
                fixture_issues = [fd for fd in report.fixture_diffs if fd.status != "match"]
                if fixture_issues:
                    with st.expander(f"器具行の差分（{len(fixture_issues)}件）", expanded=True):
                        for fd in fixture_issues:
                            status_label = {
                                "modified": "\u270f\ufe0f 修正",
                                "missing_in_ai": "\u2795 AI欠落",
                                "extra_in_ai": "\u2796 AI余分",
                            }.get(fd.status, fd.status)
                            st.markdown(f"**行{fd.row_label}** ({status_label})")
                            if fd.fixture_type_ai != fd.fixture_type_correct:
                                st.markdown(f"  器具: `{fd.fixture_type_ai}` \u2192 `{fd.fixture_type_correct}`")
                            for d in fd.diffs:
                                severity_icon = {"critical": "\U0001f534", "major": "\U0001f7e0", "minor": "\U0001f7e2"}.get(d.severity, "\u2b1c")
                                st.markdown(f"  {severity_icon} {d.field_name}: `{d.ai_value}` \u2192 `{d.correct_value}`")
                            st.markdown("---")

                # --- 選定シート差分 ---
                selection_issues = [sd for sd in report.selection_diffs if sd.status != "match"]
                if selection_issues:
                    with st.expander(f"選定シートの差分（{len(selection_issues)}件）"):
                        for sd in selection_issues:
                            st.markdown(f"**行{sd.row_number}**")
                            for d in sd.diffs:
                                st.markdown(f"  {d.field_name}: `{d.ai_value}` \u2192 `{d.correct_value}`")

                # 差分なしの場合
                if report.total_diffs == 0:
                    st.success("差分なし！AI出力と修正済みは完全一致です。")

                st.divider()

                # --- フィードバックJSON生成・ダウンロード ---
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
                        "fixtures": {
                            "ai": report.total_fixtures_ai,
                            "correct": report.total_fixtures_correct,
                        },
                        "excluded": {
                            "ai": report.total_excluded_ai,
                            "correct": report.total_excluded_correct,
                        },
                        "products": {
                            "ai": report.total_products_ai,
                            "correct": report.total_products_correct,
                        },
                    },
                    "header_diffs": [
                        {"cell": d.cell, "field": d.field_name, "severity": d.severity,
                         "ai": d.ai_value, "correct": d.correct_value}
                        for d in report.header_diffs
                    ],
                    "fixture_diffs": [
                        {
                            "row": fd.row_label, "status": fd.status,
                            "type_ai": fd.fixture_type_ai, "type_correct": fd.fixture_type_correct,
                            "diffs": [
                                {"field": d.field_name, "severity": d.severity,
                                 "ai": d.ai_value, "correct": d.correct_value}
                                for d in fd.diffs
                            ],
                        }
                        for fd in report.fixture_diffs if fd.status != "match"
                    ],
                    "selection_diffs": [
                        {
                            "row": sd.row_number, "status": sd.status,
                            "diffs": [
                                {"field": d.field_name, "ai": d.ai_value, "correct": d.correct_value}
                                for d in sd.diffs
                            ],
                        }
                        for sd in report.selection_diffs if sd.status != "match"
                    ],
                }

                feedback_json = json.dumps(report_dict, ensure_ascii=False, indent=2)
                timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                prop_name = report.property_name or "unknown"
                filename = f"feedback_{prop_name}_{timestamp_str}.json"

                # --- フィードバック送信（Google Sheets） ---
                _sheets_available = False
                try:
                    from feedback_store import FeedbackStore
                    _store = FeedbackStore.from_streamlit_secrets()
                    _sheets_available = True
                except Exception:
                    pass

                if _sheets_available:
                    if st.button("フィードバック送信",
                                 type="primary", use_container_width=True):
                        try:
                            fid = _store.submit_feedback(
                                report_dict=report_dict,
                                comment_reading=comment_reading,
                                comment_selection=comment_selection,
                            )
                            st.success(f"フィードバック送信完了 (ID: {fid})")
                        except Exception as e:
                            st.error(f"送信エラー: {e}")

                    # 累計統計
                    try:
                        stats = _store.get_feedback_stats()
                        if stats.get("total_feedback"):
                            st.caption(
                                f"累計フィードバック: {stats['total_feedback']}件 / "
                                f"LED選定平均一致率: {stats['avg_led_match_rate']:.0%}"
                            )
                    except Exception:
                        pass

                    st.divider()

                # --- JSON ダウンロード（バックアップ） ---
                st.download_button(
                    label="フィードバックJSONダウンロード",
                    data=feedback_json,
                    file_name=filename,
                    mime="application/json",
                    use_container_width=True,
                )

                if _sheets_available:
                    st.caption("送信済みデータは日次で自動的にシステム改善に反映されます。")
                else:
                    st.caption(
                        "ダウンロードしたJSONを保管しておくと、"
                        "蓄積データからAIロジックの改善に活用できます。"
                    )

            except Exception as e:
                st.error(f"比較エラー: {e}")
                import traceback
                st.code(traceback.format_exc())


# ============================================================
# メイン
# ============================================================

def main():
    st.set_page_config(page_title="LED見積作成", page_icon="\U0001f4a1", layout="centered")
    init_session()
    st.title("LED見積シミュレーション作成")

    tab1, tab2 = st.tabs(["\U0001f4cb 見積作成", "\U0001f4dd フィードバック"])

    with tab1:
        tab_estimate()

    with tab2:
        tab_feedback()


if __name__ == "__main__":
    main()
