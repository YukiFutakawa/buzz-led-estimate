"""LED見積シミュレーション作成 — Streamlit アプリ

現調データ（テキスト/写真/PDF）を入力 → LED選定 → Excel見積を自動生成。
起動: streamlit run app.py
"""

import os
import sys
import tempfile
import zipfile
import base64
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


def main():
    st.set_page_config(page_title="LED見積作成", page_icon="\U0001f4a1", layout="centered")
    st.title("LED見積シミュレーション作成")

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
                st.caption(f"{icon} {sf.name} → AI解析対象")

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


if __name__ == "__main__":
    main()
