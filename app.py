"""LED見積シミュレーション作成 — Streamlit アプリ

現調データ（テキスト/写真）を入力 → LED選定 → Excel見積を自動生成。
起動: streamlit run app.py
"""

import os
import sys
import tempfile
import zipfile
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


def save_uploaded_files(uploaded_files, dest_dir: Path) -> list[Path]:
    """アップロードファイルを一時ディレクトリに保存"""
    saved = []
    dest_dir.mkdir(parents=True, exist_ok=True)
    for uf in uploaded_files:
        path = dest_dir / uf.name
        path.write_bytes(uf.getbuffer())
        saved.append(path)
    return saved


def extract_zips(files: list[Path], dest_dir: Path) -> list[Path]:
    """ZIPを解凍し画像パスを返す。非ZIPはそのまま。"""
    result = []
    for f in files:
        if f.suffix.lower() == ".zip":
            extract_to = dest_dir / f.stem
            extract_to.mkdir(parents=True, exist_ok=True)
            with zipfile.ZipFile(f, "r") as zf:
                zf.extractall(extract_to)
            for img in sorted(extract_to.rglob("*")):
                if img.is_file() and img.suffix.lower() in IMAGE_EXTENSIONS:
                    result.append(img)
        elif f.suffix.lower() in IMAGE_EXTENSIONS:
            result.append(f)
    return result


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
    st.set_page_config(page_title="LED見積作成", page_icon="💡", layout="centered")
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

    # --- 照明設置状況（テキスト）---
    st.subheader("① 照明設置状況")
    text_input = st.text_area(
        "テキスト入力", height=200,
        placeholder="例:\n4階\n20w非常内蔵逆富士 3台\n誘導灯 2台LED更新済み\n\n3階\n20w非常内蔵逆富士 3台",
    )
    txt_file = st.file_uploader("または .txt ファイル", type=["txt"])
    if txt_file:
        text_input = txt_file.read().decode("utf-8")
        st.text_area("読み込み内容", text_input, height=150, disabled=True)

    st.divider()

    # --- 現調写真 ---
    st.subheader("② 現調写真")
    uploaded_photos = st.file_uploader(
        "画像ファイルをアップロード（複数選択可、ZIP対応）",
        type=["jpg", "jpeg", "png", "gif", "webp", "zip"],
        accept_multiple_files=True,
    )
    if uploaded_photos:
        st.caption(f"{len(uploaded_photos)} ファイル選択済み")

    st.divider()

    # --- バリデーション ---
    has_text = bool(text_input and text_input.strip())
    has_photos = bool(uploaded_photos)
    can_run = bool(property_name) and (has_text or has_photos)

    if not property_name:
        st.info("物件名を入力してください")
    elif not has_text and not has_photos:
        st.info("テキストまたは写真を入力してください")

    # --- 実行ボタン ---
    if st.button("見積作成", type="primary", disabled=not can_run, use_container_width=True):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir = Path(tmpdir)

            photo_paths = None
            photo_dir = None
            if has_photos:
                saved = save_uploaded_files(uploaded_photos, tmpdir / "uploads")
                photo_paths = extract_zips(saved, tmpdir / "extracted")
                if photo_paths and not has_text:
                    import shutil
                    photo_dir = tmpdir / "photo_dir"
                    photo_dir.mkdir(exist_ok=True)
                    for p in photo_paths:
                        dest = photo_dir / p.name
                        if not dest.exists():
                            shutil.copy2(p, dest)

            try:
                with st.status("処理中...", expanded=True) as status:
                    if has_text:
                        result_path = process_text_route(
                            text=text_input, photo_paths=photo_paths,
                            property_name=property_name, address=address,
                            template_name=template_name, status=status,
                        )
                    else:
                        result_path = process_photo_route(
                            photo_dir=photo_dir, property_name=property_name,
                            template_name=template_name, status=status,
                        )
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
