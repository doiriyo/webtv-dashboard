# asology インサイトダッシュボード

Instagram Graph APIから自動収集したasologyのインサイトデータを可視化するWebアプリ。

## セットアップ手順

### 1. GitHubリポジトリを作成

1. [GitHub](https://github.com) にログイン
2. 右上の「+」→「New repository」
3. リポジトリ名: `asology-dashboard`
4. 「Create repository」をクリック

### 2. コードをアップロード

**方法A: GitHubのWebUIから直接アップロード**

1. 作成したリポジトリページで「uploading an existing file」をクリック
2. このフォルダ内の全ファイル・フォルダをドラッグ＆ドロップ
3. 「Commit changes」をクリック

**方法B: Gitコマンド（PCにGitがインストール済みの場合）**

```bash
cd asology-dashboard
git init
git add .
git commit -m "initial commit"
git branch -M main
git remote add origin https://github.com/（あなたのユーザー名）/asology-dashboard.git
git push -u origin main
```

### 3. 自動デプロイを確認

1. コードをpushすると `.github/workflows/deploy.yml` が自動実行され、Xserverへ自動デプロイされる
2. 「Actions」タブで進捗を確認

## API URLの変更

GASのWebアプリURLが変わった場合は、`src/App.jsx` の冒頭にある `API_URL` を更新してください。

```javascript
const API_URL = 'https://script.google.com/macros/s/xxxxx/exec'
```

## ローカル開発

```bash
npm install
npm run dev
```

`http://localhost:5173` でプレビューできます。
