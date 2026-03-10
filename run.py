"""LED導入シミュレーション — 簡単実行スクリプト

使い方:
  1.「現調写真」フォルダに物件フォルダを入れる
  2. このファイルをダブルクリック（または python run.py）
  3. モードを選ぶ:
     [1] SFAから案件を選んで見積作成（写真は現調写真フォルダに配置済み前提）
     [2] ローカル写真フォルダから見積作成（従来モード）
     [3] フィードバック比較（AI出力 vs 正解Excelの差分分析）
  4. 完了後、outputフォルダにExcelが生成される

初回のみ:
  APIキーの入力が必要です（.envファイルに保存されます）
"""

import sys
import os
from pathlib import Path

# プロジェクトルート
ROOT = Path(__file__).parent
SRC_DIR = ROOT / "src"
ENV_FILE = ROOT / ".env"

# srcをPythonパスに追加
sys.path.insert(0, str(SRC_DIR))
os.chdir(SRC_DIR)


def load_env():
    """`.env` ファイルからAPIキーを読み込み、環境変数にセット"""
    if ENV_FILE.exists():
        with open(ENV_FILE, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    key, _, value = line.partition("=")
                    os.environ[key.strip()] = value.strip()


def ensure_api_key() -> str:
    """APIキーが設定されているか確認。なければ入力を求める"""
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")

    if api_key:
        print(f"  Anthropic APIキー: ...{api_key[-8:]} (設定済み)")
        return api_key

    print("\n" + "-" * 50)
    print("  Anthropic APIキーが未設定です。")
    print("  初回のみ入力が必要です（.envファイルに保存されます）")
    print("-" * 50)
    print()
    print("  APIキーの取得方法:")
    print("    1. https://console.anthropic.com/ にアクセス")
    print("    2. API Keys → Create Key")
    print("    3. 「sk-ant-...」で始まるキーをコピー")
    print()

    api_key = input("  APIキーを貼り付け: ").strip()

    if not api_key:
        print("  キャンセルしました。")
        return ""

    if not api_key.startswith("sk-ant-"):
        print("  警告: APIキーは通常「sk-ant-」で始まります。")
        confirm = input("  このまま保存しますか？ (y/n): ").strip().lower()
        if confirm != "y":
            return ""

    # .envファイルに追記
    existing = ""
    if ENV_FILE.exists():
        existing = ENV_FILE.read_text(encoding="utf-8")
    with open(ENV_FILE, "w", encoding="utf-8") as f:
        f.write(f"ANTHROPIC_API_KEY={api_key}\n")
        # 既存の他のキーも保持
        for line in existing.splitlines():
            if line.strip() and not line.startswith("ANTHROPIC_API_KEY"):
                f.write(line + "\n")
    print(f"  保存しました: {ENV_FILE}")

    os.environ["ANTHROPIC_API_KEY"] = api_key
    return api_key


def find_properties():
    """現調写真フォルダ内の物件フォルダを探す"""
    photo_dir = ROOT / "現調写真"
    if not photo_dir.exists():
        photo_dir.mkdir()
        print(f"「現調写真」フォルダを作成しました: {photo_dir}")
        return []

    properties = []
    for item in sorted(photo_dir.iterdir()):
        if item.is_dir():
            image_exts = {".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".heic"}
            images = [f for f in item.iterdir() if f.suffix.lower() in image_exts]
            if images:
                properties.append((item, len(images)))
        elif item.suffix.lower() == ".zip":
            properties.append((item, -1))

    return properties


def extract_zip(zip_path: Path) -> Path:
    """ZIPを解凍して物件フォルダを返す"""
    import zipfile

    extract_dir = zip_path.parent / zip_path.stem
    if extract_dir.exists():
        print(f"  既に解凍済み: {extract_dir.name}")
        return extract_dir

    print(f"  ZIP解凍中: {zip_path.name}")
    with zipfile.ZipFile(zip_path, "r") as zf:
        zf.extractall(extract_dir)
    print(f"  解凍完了: {extract_dir.name}")
    return extract_dir


def select_template():
    """テンプレートを選択"""
    templates = {
        "1": "田村基本形",
        "2": "ウスイホーム",
        "3": "オクスト",
        "4": "クラスコ",
        "5": "タカラ",
        "6": "ニッショー",
        "7": "マンション経営保障",
        "8": "丸八アセットマネジメント",
        "9": "スマイルサポート",
        "10": "ライフサポート",
        "11": "高知ハウス",
    }

    print("\n--- テンプレート選択 ---")
    for key, name in templates.items():
        default_mark = " (デフォルト)" if key == "1" else ""
        print(f"  {key}. {name}{default_mark}")

    choice = input("\n番号を入力 (Enterでデフォルト): ").strip()
    return templates.get(choice, "田村基本形")


# ==================================================
# SFA連携モード
# ==================================================

def sfa_mode(api_key: str):
    """ネクストSFAから案件を選んで見積作成"""
    from sfa_client import SFAClient, SFAProject

    sfa_key = os.environ.get("NEXTSFA_API_KEY", "")
    sfa_token = os.environ.get("NEXTSFA_API_TOKEN", "")

    if not sfa_key or not sfa_token:
        print("\n  SFA APIキーが未設定です。")
        print("  .envファイルに以下を追加してください:")
        print("    NEXTSFA_API_KEY=...")
        print("    NEXTSFA_API_TOKEN=...")
        input("\nEnterで終了...")
        return

    print(f"  SFA APIキー: ...{sfa_key[-8:]} (設定済み)")
    print("\n  ネクストSFAから案件を取得中...")

    try:
        client = SFAClient(api_key=sfa_key, api_token=sfa_token)
    except Exception as e:
        print(f"  SFA接続エラー: {e}")
        input("\nEnterで終了...")
        return

    # 案件を取得（最初の数ページ）
    print("  案件一覧を読み込み中...")
    projects = client.get_projects(max_pages=5)
    print(f"  {len(projects)} 件の案件を取得しました。")

    # フィルタリング: 現調写真フォルダに対応するフォルダがある案件を優先表示
    photo_dir = ROOT / "現調写真"
    local_folders = set()
    if photo_dir.exists():
        for item in photo_dir.iterdir():
            if item.is_dir():
                local_folders.add(item.name)

    # 案件を表示
    print("\n--- SFA案件一覧 ---")
    print("  [番号] 物件名 | フェーズ | 住所")
    print("  " + "-" * 60)

    matched = []  # ローカル写真あり
    unmatched = []  # ローカル写真なし

    for p in projects:
        has_local = any(
            p.name in folder or folder in p.name for folder in local_folders
        )
        if has_local:
            matched.append(p)
        else:
            unmatched.append(p)

    # 写真ありを先に表示
    display_list: list[SFAProject] = []
    if matched:
        print("  --- 写真フォルダあり ---")
        for p in matched:
            idx = len(display_list) + 1
            display_list.append(p)
            phase = p.phase or "未設定"
            addr = p.address[:20] if p.address else ""
            print(f"  {idx:3d}. * {p.name[:35]:<35s} | {phase:<8s} | {addr}")

    print("  --- その他の案件 ---")
    for p in unmatched[:30]:  # 最大30件表示
        idx = len(display_list) + 1
        display_list.append(p)
        phase = p.phase or "未設定"
        addr = p.address[:20] if p.address else ""
        print(f"  {idx:3d}.   {p.name[:35]:<35s} | {phase:<8s} | {addr}")

    if len(unmatched) > 30:
        print(f"  ... 他 {len(unmatched) - 30} 件")

    # 案件選択
    choice = input("\n案件番号を入力: ").strip()
    try:
        selected = display_list[int(choice) - 1]
    except (ValueError, IndexError):
        print("無効な番号です。")
        input("\nEnterで終了...")
        return

    print(f"\n  選択: {selected.name}")
    print(f"  住所: {selected.address}")
    if selected.unlock_info:
        print(f"  解錠: {selected.unlock_info}")
    if selected.management_company:
        print(f"  管理: {selected.management_company}")

    # 写真フォルダを特定
    survey_dir = None
    for folder in local_folders:
        if selected.name in folder or folder in selected.name:
            survey_dir = photo_dir / folder
            break

    if not survey_dir:
        # 名前の部分一致を試す
        for folder in local_folders:
            # 物件名の最初の数文字で照合
            if len(selected.name) >= 3 and selected.name[:3] in folder:
                survey_dir = photo_dir / folder
                break

    if not survey_dir or not survey_dir.exists():
        print(f"\n  この案件の写真フォルダが見つかりません。")
        print(f"  「現調写真」フォルダに「{selected.name}」フォルダを作成し、")
        print(f"  現調写真を配置してください。")
        input("\nEnterで終了...")
        return

    image_exts = {".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".heic"}
    images = [f for f in survey_dir.iterdir() if f.suffix.lower() in image_exts]
    print(f"\n  写真フォルダ: {survey_dir.name} ({len(images)}枚)")

    # テンプレート選択
    template_name = select_template()

    # 実行
    run_pipeline_with_sfa(api_key, survey_dir, template_name, selected)


def run_pipeline_with_sfa(api_key, survey_dir, template_name, sfa_project):
    """SFA案件情報付きでパイプラインを実行"""
    print(f"\n--- 実行内容 ---")
    print(f"  物件: {sfa_project.name}")
    print(f"  写真: {survey_dir.name}")
    print(f"  テンプレート: {template_name}")
    confirm = input("\n実行しますか？ (y/Enter): ").strip().lower()
    if confirm not in ("", "y", "yes"):
        print("キャンセルしました。")
        input("\nEnterで終了...")
        return

    print("\n" + "=" * 50)
    print("  処理開始...")
    print("=" * 50 + "\n")

    import logging

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )

    try:
        from pipeline import run_pipeline

        result_path = run_pipeline(
            survey_dir=survey_dir,
            lineup_dir=ROOT / "ラインナップ表",
            template_dir=ROOT / "見積りテンプレート",
            template_name=template_name,
            api_key=api_key,
        )

        print("\n" + "=" * 50)
        print("  完了！")
        print("=" * 50)
        print(f"\n出力ファイル: {result_path}")

        open_file = input("\nExcelを開きますか？ (y/Enter): ").strip().lower()
        if open_file in ("", "y", "yes"):
            os.startfile(result_path)

    except Exception as e:
        print(f"\nエラーが発生しました: {e}")
        import traceback

        traceback.print_exc()
        input("\nEnterで終了...")
        return

    input("\nEnterで終了...")


# ==================================================
# ローカルモード（従来）
# ==================================================

def local_mode(api_key: str):
    """ローカル写真フォルダから見積作成（従来モード）"""
    properties = find_properties()

    if not properties:
        print("\n「現調写真」フォルダに物件フォルダが見つかりません。")
        print(f"  場所: {ROOT / '現調写真'}")
        print("  → 現地調査の写真フォルダを入れてください。")
        input("\nEnterで終了...")
        return

    print("\n--- 物件一覧 ---")
    for i, (path, count) in enumerate(properties, 1):
        if count == -1:
            print(f"  {i}. {path.name}  [ZIP・未解凍]")
        else:
            print(f"  {i}. {path.name}  [{count}枚]")

    choice = input("\n番号を入力: ").strip()
    try:
        idx = int(choice) - 1
        selected_path, count = properties[idx]
    except (ValueError, IndexError):
        print("無効な番号です。")
        input("\nEnterで終了...")
        return

    if selected_path.suffix.lower() == ".zip":
        selected_path = extract_zip(selected_path)

    template_name = select_template()

    print(f"\n--- 実行内容 ---")
    print(f"  物件: {selected_path.name}")
    print(f"  テンプレート: {template_name}")
    confirm = input("\n実行しますか？ (y/Enter): ").strip().lower()
    if confirm not in ("", "y", "yes"):
        print("キャンセルしました。")
        input("\nEnterで終了...")
        return

    print("\n" + "=" * 50)
    print("  処理開始...")
    print("=" * 50 + "\n")

    import logging

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )

    try:
        from pipeline import run_pipeline

        result_path = run_pipeline(
            survey_dir=selected_path,
            lineup_dir=ROOT / "ラインナップ表",
            template_dir=ROOT / "見積りテンプレート",
            template_name=template_name,
            api_key=api_key,
        )

        print("\n" + "=" * 50)
        print("  完了！")
        print("=" * 50)
        print(f"\n出力ファイル: {result_path}")

        open_file = input("\nExcelを開きますか？ (y/Enter): ").strip().lower()
        if open_file in ("", "y", "yes"):
            os.startfile(result_path)

    except Exception as e:
        print(f"\nエラーが発生しました: {e}")
        import traceback

        traceback.print_exc()
        input("\nEnterで終了...")
        return

    input("\nEnterで終了...")


# ==================================================
# フィードバック比較モード
# ==================================================

def feedback_mode():
    """AI出力 vs 正解Excelの差分分析"""
    from feedback_comparator import FeedbackComparator
    from feedback_accumulator import FeedbackAccumulator

    output_dir = ROOT / "output"
    correct_dir = ROOT / "正しい見積り"
    feedback_dir = ROOT / "feedback"

    print("\n--- フィードバック比較モード ---")
    print(f"  AI出力フォルダ:   {output_dir}")
    print(f"  正解フォルダ:     {correct_dir}")
    print(f"  フィードバック保存: {feedback_dir}")

    print("\n  サブモード選択:")
    print("  1. 個別比較（AI出力と正解ファイルを選択して比較）")
    print("  2. 一括比較（output/ と 正しい見積り/ を自動マッチング）")
    print("  3. 蓄積レポート（過去のフィードバックを集約分析）")

    sub = input("\n番号を入力 (Enterで2): ").strip()

    if sub == "1":
        _feedback_single(output_dir, correct_dir, feedback_dir)
    elif sub == "3":
        _feedback_report(feedback_dir)
    else:
        _feedback_batch(output_dir, correct_dir, feedback_dir)

    input("\nEnterで終了...")


def _feedback_single(output_dir: Path, correct_dir: Path, feedback_dir: Path):
    """個別ファイル比較"""
    from feedback_comparator import FeedbackComparator

    # AI出力ファイル一覧
    ai_files = sorted(
        f for f in output_dir.glob("*.xlsx")
        if not f.name.startswith("~$")
    )
    if not ai_files:
        print("  AI出力ファイルが見つかりません。")
        return

    print("\n--- AI出力ファイル ---")
    for i, f in enumerate(ai_files, 1):
        print(f"  {i}. {f.name}")

    choice = input("\n番号を入力: ").strip()
    try:
        ai_file = ai_files[int(choice) - 1]
    except (ValueError, IndexError):
        print("無効な番号です。")
        return

    # 正解ファイル一覧
    correct_files = sorted(
        f for f in correct_dir.glob("*.xlsx")
        if not f.name.startswith("~$")
    )
    if not correct_files:
        print("  正解ファイルが見つかりません。")
        return

    print("\n--- 正解ファイル ---")
    for i, f in enumerate(correct_files, 1):
        print(f"  {i}. {f.name}")

    choice = input("\n番号を入力: ").strip()
    try:
        correct_file = correct_files[int(choice) - 1]
    except (ValueError, IndexError):
        print("無効な番号です。")
        return

    # 比較実行
    comparator = FeedbackComparator()
    try:
        report = comparator.compare(ai_file, correct_file)
        print(report.print_summary())

        # JSON保存
        json_name = f"feedback_{report.property_name}.json"
        report.save_json(feedback_dir / json_name)
        print(f"\n  フィードバック保存: {feedback_dir / json_name}")
    except Exception as e:
        print(f"\n比較エラー: {e}")
        import traceback
        traceback.print_exc()


def _feedback_batch(output_dir: Path, correct_dir: Path, feedback_dir: Path):
    """一括比較"""
    from feedback_comparator import FeedbackComparator

    comparator = FeedbackComparator()
    reports = comparator.compare_folder(output_dir, correct_dir)

    if not reports:
        print("\n  比較対象のファイルペアが見つかりませんでした。")
        print("  output/ 内のファイル名と 正しい見積り/ 内の物件名が一致する必要があります。")
        return

    print(f"\n  {len(reports)}件のファイルペアを比較しました。\n")

    for report in reports:
        print(report.print_summary())
        print()

        # JSON保存
        json_name = f"feedback_{report.property_name}.json"
        report.save_json(feedback_dir / json_name)

    # 蓄積レポートも自動生成
    from feedback_accumulator import FeedbackAccumulator
    acc = FeedbackAccumulator(feedback_dir)
    acc.load_all()
    print(acc.generate_improvement_report())

    # LED選定ルール出力
    rules_path = feedback_dir / "led_selection_rules.json"
    acc.export_led_rules_json(rules_path)


def _feedback_report(feedback_dir: Path):
    """蓄積レポート表示"""
    from feedback_accumulator import FeedbackAccumulator

    acc = FeedbackAccumulator(feedback_dir)
    count = acc.load_all()

    if count == 0:
        print("\n  フィードバックデータがありません。")
        print("  先に「一括比較」または「個別比較」を実行してください。")
        return

    print(acc.generate_improvement_report())


# ==================================================
# メイン
# ==================================================

def enhanced_batch_mode(api_key: str):
    """拡張バッチモード: 全データソースから一括処理"""
    import logging
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )

    from sfa_client import SFAClient
    from enhanced_batch import EnhancedBatchProcessor

    client = SFAClient()
    processor = EnhancedBatchProcessor(sfa_client=client, api_key=api_key)

    # 数量入力待ち案件を取得（キャッシュ経由で高速）
    print("
  数量入力待ち案件を取得中...")
    age = processor.cache.cache_age_hours()
    if age is not None:
        print(f"  キャッシュあり（{age:.1f}時間前）")
    else:
        print("  初回取得: 全案件スキャン中（数分かかります）...")
    projects = processor.get_target_projects(phase="数量入力待ち")
    print(f"  {len(projects)} 件の案件を検出")

    # Phase 1: データソース判定
    infos = processor.resolve_sources(projects)

    # 判定結果を表示
    by_type = {}
    for info in infos:
        by_type[info.source_type] = by_type.get(info.source_type, 0) + 1
    print("
  --- データソース判定結果 ---")
    labels = {
        "local_photos": "ローカル写真",
        "sfa_zip": "SFA ZIP",
        "history_text": "履歴テキスト",
        "none": "データなし",
    }
    for src, count in sorted(by_type.items()):
        print(f"    {labels.get(src, src)}: {count} 件")

    # 処理オプション
    print("
  処理オプション:")
    print("  1. ドライラン（データソース確認のみ）")
    print("  2. 本実行（API呼出しあり）")
    choice = input("
  番号を入力 (Enterで1): ").strip()
    dry_run = choice != "2"

    # 処理実行
    template_name = select_template()
    processor.process_all(infos, template_name=template_name, dry_run=dry_run)

    # レポート表示 & 保存
    report = processor.generate_report()
    print("
" + report)
    result_path = processor.save_results()
    print(f"
  結果保存: {result_path}")

    # レポートメール送信
    print("
  レポートメールを送信中...")
    if processor.send_report_email():
        print("  メール送信完了")
    else:
        print("  メール送信スキップ（SMTP設定が必要です）")
        print("  .env に SMTP_HOST, SMTP_USER, SMTP_PASS を追加してください")

    input("
Enterで終了...")


def main():
    print("=" * 50)
    print("  LED導入シミュレーション 自動作成システム")
    print("=" * 50)

    load_env()

    # モード選択（フィードバックモードはAPIキー不要）
    sfa_available = bool(
        os.environ.get("NEXTSFA_API_KEY") and os.environ.get("NEXTSFA_API_TOKEN")
    )

    print("\n--- モード選択 ---")
    print("  1. SFA連携モード（ネクストSFAから案件を選択）")
    if not sfa_available:
        print("     ※ SFA APIキーが未設定のため利用不可")
    print("  2. ローカルモード（現調写真フォルダから選択）")
    print("  3. フィードバック比較（AI出力 vs 正解の差分分析）")
    print("  4. 拡張バッチ（全ソース自動判定・一括処理）")

    mode = input("\n番号を入力 (Enterで1): ").strip()

    if mode == "3":
        feedback_mode()
        return

    if mode == "4":
        api_key = ensure_api_key()
        if api_key:
            enhanced_batch_mode(api_key)
        return

    # モード1,2はAPIキーが必要
    api_key = ensure_api_key()
    if not api_key:
        input("\nEnterで終了...")
        return

    if mode in ("", "1") and sfa_available:
        sfa_mode(api_key)
    elif mode == "1" and not sfa_available:
        print("  SFA APIキーが設定されていません。")
        print("  .envファイルに NEXTSFA_API_KEY と NEXTSFA_API_TOKEN を追加してください。")
        input("\nEnterで終了...")
    else:
        local_mode(api_key)


if __name__ == "__main__":
    main()
