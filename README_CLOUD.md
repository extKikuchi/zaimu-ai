# Excel データ集計システム

AWS Lambda連携のExcel集計システムをStreamlit Cloudで公開

## 🚀 機能

- 事業計画ExcelファイルからPLデータを自動抽出
- input用ファイルのPL計画シートに統一フォーマットで入力
- AWS Lambda + S3でのクラウド処理
- Streamlit WebUIでの直感的な操作

## 📋 必要な設定

### AWS認証情報
Streamlit Cloudのシークレット設定で以下を設定してください：

```
AWS_ACCESS_KEY_ID = "your_access_key"
AWS_SECRET_ACCESS_KEY = "your_secret_key"
AWS_DEFAULT_REGION = "ap-northeast-1"
```

### AWS設定
- S3バケット: `excel-aggregator-20250714-095126`
- Lambda関数: `excel-data-aggregator`
- リージョン: `ap-northeast-1`

## 🎯 使用方法

1. アプリケーションにアクセス
2. input用テンプレートファイルをアップロード
3. 事業計画ファイルをアップロード
4. "データ集計実行"をクリック
5. 処理結果をダウンロード

## 🔧 ローカル開発

```bash
pip install -r requirements.txt
streamlit run streamlit_cloud_app.py
```

## 📞 サポート

システムに関するご質問や問題がございましたら、お気軽にお問い合わせください。
