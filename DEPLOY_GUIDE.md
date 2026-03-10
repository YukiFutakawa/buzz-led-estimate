# LED見積アプリ — デプロイ手順書

## 概要

このStreamlitアプリを社内メンバーと共有するため、Streamlit Community Cloud にデプロイします。
メンバーはブラウザでURLにアクセスするだけで使えるようになります。

---

## 事前準備

以下のアカウントが必要です（全て無料）:

| サービス | URL | 用途 |
|---------|-----|------|
| GitHub | https://github.com | コードをホスト |
| Streamlit Community Cloud | https://share.streamlit.io | アプリをホスト |

---

## 手順1: GitHubリポジトリを作成

### 1-1. GitHubにログイン

https://github.com にアクセスしてログイン。

### 1-2. 新しいリポジトリを作成

1. 右上の「+」→「New repository」をクリック
2. 以下を入力:
   - Repository name: `buzz-led-estimate`
   - Visibility: **Private**（社内限定のため）
   - 他はデフォルトのままでOK
3. 「Create repository」をクリック

---

## 手順2: Git LFSのインストール

ラインナップ表のExcelファイルが100MBを超えるため、Git LFS（大容量ファイル管理）が必要です。

### Windows（PowerShell を管理者として実行）:
```
winget install GitHub.GitLFS
```

### または:
https://git-lfs.github.com からダウンロード・インストール

### インストール確認:
```
git lfs version
```
→ `git-lfs/3.x.x` と表示されればOK

---

## 手順3: ローカルでGitリポジトリを初期化

コマンドプロンプト（またはPowerShell）で以下を実行:

```
cd C:\Users\yukin\OneDrive\Documents\Buzzarea\buzz_led\見積作成

git init
git lfs install
git lfs track "ラインナップ表/*.xlsx"
```

---

## 手順4: ファイルをコミット

```
git add .gitattributes
git add .gitignore
git add requirements.txt
git add app.py
git add src/
git add config/
git add "見積りテンプレート/"
git add "ラインナップ表/新【その他】ラインナップ.xlsx"
git add "ラインナップ表/新【マニアック】ラインナップ表.xlsx"
git add "ラインナップ表/新【蛍光灯】ラインナップ.xlsx"

git commit -m "LED見積シミュレーション Streamlitアプリ"
```

※ バックアップファイル（_backup_*.xlsx）、.env、現調写真、outputは
  .gitignoreで除外されるため含まれません。

---

## 手順5: GitHubにプッシュ

```
git remote add origin https://github.com/あなたのユーザー名/buzz-led-estimate.git
git branch -M main
git push -u origin main
```

※ 初回プッシュ時、Git LFSのファイルアップロードに時間がかかります（数分）。

---

## 手順6: Streamlit Community Cloudにデプロイ

### 6-1. Streamlit Cloudにログイン

1. https://share.streamlit.io にアクセス
2. 「Continue with GitHub」でGitHubアカウントでログイン

### 6-2. アプリをデプロイ

1. 「New app」をクリック
2. 以下を設定:
   - Repository: `あなたのユーザー名/buzz-led-estimate`
   - Branch: `main`
   - Main file path: `app.py`
3. 「Advanced settings」をクリック
4. **Secrets** に以下を入力:
   ```
   ANTHROPIC_API_KEY = "sk-ant-xxxxx（あなたのAPIキー）"
   ```
5. Python version: `3.11` を選択
6. 「Deploy!」をクリック

### 6-3. デプロイ完了

数分でデプロイが完了し、URLが発行されます:
```
https://あなたのアプリ名.streamlit.app
```

---

## 手順7: メンバーに共有

### プライベートリポジトリの場合:

Streamlit Cloudの設定で「Viewer access」を設定:
1. アプリの右上「...」→「Settings」→「Sharing」
2. 「This app is viewable to:」→ メンバーのメールアドレスを追加

### または:

リポジトリをパブリックにすれば、URLを知っている人は誰でもアクセス可能。

---

## トラブルシューティング

### デプロイ時にエラーが出る場合

1. Streamlit Cloudの「Manage app」→「Logs」でエラー内容を確認
2. よくある原因:
   - **Secretsの設定忘れ** → ANTHROPIC_API_KEY を設定しているか確認
   - **パッケージ不足** → requirements.txt に足りないパッケージがないか確認
   - **Git LFSのファイルが取得できない** → GitHub LFSの容量制限を確認

### APIキーを変更したい場合

1. Streamlit Cloudのアプリ画面 → 右上「...」→「Settings」→「Secrets」
2. `ANTHROPIC_API_KEY` の値を変更
3. 「Save」→ アプリが自動で再起動

### ラインナップ表を更新したい場合

1. ローカルでファイルを更新
2. `git add "ラインナップ表/更新したファイル.xlsx"`
3. `git commit -m "ラインナップ表更新"`
4. `git push`
5. Streamlit Cloudが自動で再デプロイ

---

## 費用

| サービス | プラン | 費用 |
|---------|-------|------|
| GitHub | Free | 無料（Git LFS: 1GB容量、1GB/月帯域） |
| Streamlit Community Cloud | Free | 無料 |
| Anthropic API | 従量課金 | テキスト解析1件あたり約1-5円 |

※ Git LFS の無料枠 (1GB) を超える場合は GitHub の Data Pack ($5/月) が必要。
  現在のラインナップ表は合計約320MBなので無料枠内に収まります。
