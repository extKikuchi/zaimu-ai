# AI Excel集計システム (Claude API版)

このアプリケーションは、Claude AIを活用してExcelファイルから財務データを自動抽出・集計するStreamlitアプリです。

## 🚀 特徴

- **高精度AI抽出**: Claude-3-Haiku APIによる柔軟なデータ認識
- **日本語対応**: 日本語の財務用語を深く理解
- **項目名のゆらぎ対応**: 表記の違いを自動で認識
- **自動単位変換**: 数値の単位を自動で正規化
- **リアルタイム処理**: AWS Lambdaによる高速処理

## 📋 対応項目

- 売上高、売上原価、売上総利益
- 販売費及び一般管理費、営業利益
- 経常利益、当期純利益
- その他のPL項目

## 🛠️ Streamlit Cloudでの設定方法

### 1. リポジトリの準備
このコードをGitHubリポジトリにプッシュしてください。

### 2. Streamlit Cloudでのデプロイ
1. [Streamlit Cloud](https://share.streamlit.io/)にアクセス
2. GitHubアカウントでログイン
3. "New app" をクリック
4. リポジトリ、ブランチ、メインファイル（app.py）を選択
5. "Deploy" をクリック

### 3. Secrets設定
1. デプロイ後、アプリのダッシュボードで設定（歯車アイコン）をクリック
2. "Secrets" タブを選択
3. 以下の形式でAWS認証情報を追加:

```toml
AWS_ACCESS_KEY_ID = "your_access_key_here"
AWS_SECRET_ACCESS_KEY = "your_secret_key_here"
```

4. "Save" をクリック
5. アプリを再起動

## 🔧 前提条件

### AWS Lambda関数
以下のLambda関数が事前にデプロイされている必要があります：
- `excel-claude-aggregator` (Claude API版 - 推奨)
- `excel-data-aggregator` (従来版)

### S3バケット
- バケット名: `excel-ai-aggregator-6142`
- リージョン: `ap-northeast-1` (東京)

### 必要な権限
AWS IAMユーザーには以下の権限が必要です：
- Lambda関数の実行権限
- S3バケットの読み書き権限

## 📖 使用方法

1. **ファイルアップロード**: サイドバーからExcelファイルを選択
2. **シート選択**: 「受注ベース収支計画」などの適切なシートを選択
3. **処理実行**: 「Claude AI集計実行」ボタンをクリック
4. **結果確認**: 抽出されたデータと生成されたファイルを確認
5. **ダウンロード**: 処理済みファイルをダウンロード

## 🏗️ システム構成

```
Streamlit Cloud (Frontend)
    ↓
AWS Lambda (Processing)
    ↓
Claude API (AI Analysis)
    ↓
S3 (File Storage)
```

## ⚠️ 注意事項

- Claude APIキーがAWS Secrets Managerに設定されている必要があります
- ファイルサイズは10MB以下を推奨
- 処理時間は通常1-3分程度です

## 🆘 トラブルシューティング

### AWS接続エラー
- Secrets設定を確認
- IAM権限を確認
- リージョン設定を確認

### Lambda関数エラー
- 関数がデプロイされているか確認
- タイムアウト設定を確認
- Claude APIキーの設定を確認

## 📞 サポート

問題が発生した場合は、エラーメッセージとともにお問い合わせください。
