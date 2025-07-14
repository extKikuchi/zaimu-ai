# 📊 Excel データ集計システム - Streamlit Cloud デプロイガイド

## 🚨 Streamlit Cloud エラー対処法

### pandas ImportError の解決方法

Streamlit Cloudで `pandas` インポートエラーが発生した場合の対処方法：

#### 1. requirements.txt の最適化

以下の固定バージョンを使用してください：

```txt
# Streamlit Cloud対応 requirements.txt
streamlit==1.32.0
boto3==1.34.69
pandas==2.2.0
numpy==1.26.0
openpyxl==3.1.2
python-dateutil==2.8.2
```

#### 2. Streamlit Cloud での設定

1. **GitHub リポジトリを更新**
   - 修正した `requirements.txt` をコミット・プッシュ
   
2. **Streamlit Cloud でアプリを再起動**
   - "Manage app" → "Reboot app"
   - または新しいデプロイメントを作成

3. **Python バージョンの確認**
   - Streamlit Cloud は Python 3.9 を使用
   - ローカルテストも Python 3.9 で実行推奨

#### 3. 代替手順（エラーが続く場合）

1. **requirements_streamlit_cloud.txt を使用**
   ```bash
   # リポジトリルートに配置
   cp requirements_streamlit_cloud.txt requirements.txt
   ```

2. **段階的なライブラリ追加**
   ```txt
   # 最小構成から開始
   streamlit==1.32.0
   # 動作確認後、他のライブラリを追加
   ```

3. **キャッシュクリア**
   - Streamlit Cloud の設定画面でキャッシュをクリア

## 🔧 AWS 設定（Streamlit Cloud）

### 必要な Secrets 設定

Settings → Secrets に以下を設定：

```toml
AWS_ACCESS_KEY_ID = "your_access_key_id"
AWS_SECRET_ACCESS_KEY = "your_secret_access_key"
AWS_DEFAULT_REGION = "ap-northeast-1"
S3_BUCKET_NAME = "excel-aggregator-20250714-095126"
LAMBDA_FUNCTION_NAME = "excel-data-aggregator"
```

### 自動設定ファイル読み込み

アプリは以下の順序で設定を読み込みます：

1. ローカルの `aws_config.json`（開発環境）
2. Streamlit Cloud Secrets（本番環境）

## 🐛 トラブルシューティング

### エラー: "pandas インポートエラー"

**原因**: ライブラリバージョンの競合

**解決策**:
1. `requirements.txt` を固定バージョンに更新
2. Streamlit Cloud でアプリを再起動
3. サイドバーのデバッグ情報でバージョン確認

### エラー: "AWS接続に失敗しました"

**原因**: AWS認証情報の不備

**解決策**:
1. Streamlit Cloud Secrets の設定確認
2. IAM権限の確認（S3 + Lambda）
3. サイドバーの認証情報デバッグで確認

### エラー: "Lambda関数の実行に失敗しました"

**原因**: Lambda関数の設定問題

**解決策**:
1. Lambda関数が正しくデプロイされているか確認
2. 関数名が `excel-data-aggregator` で一致しているか確認
3. Lambda関数のタイムアウト設定（推奨: 300秒）

## 📈 パフォーマンス最適化

### Streamlit Cloud での推奨設定

- **Python**: 3.9
- **Memory**: 自動（最大 1GB）
- **CPU**: 共有（無料プラン）

### 大容量ファイル処理

- **制限**: 50MB / ファイル
- **推奨**: 複数ファイルは ZIP で圧縮
- **最適化**: Lambda のメモリサイズ 512MB 以上

## 🔗 関連リンク

- [Streamlit Cloud ドキュメント](https://docs.streamlit.io/streamlit-cloud)
- [AWS Lambda ドキュメント](https://docs.aws.amazon.com/lambda/)
- [pandas トラブルシューティング](https://pandas.pydata.org/docs/getting_started/install.html)

---

**更新日**: 2025-07-14  
**対応バージョン**: Streamlit 1.32.0, pandas 2.2.0
