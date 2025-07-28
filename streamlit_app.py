import streamlit as st
import boto3
import json
import pandas as pd
from datetime import datetime
import tempfile
import os
from pathlib import Path

# Streamlit設定
st.set_page_config(
    page_title="AI Excel集計システム (Claude API版)",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# セッション状態の初期化
def init_session_state():
    """セッション状態を初期化"""
    if 'file_configs' not in st.session_state:
        st.session_state.file_configs = []
    if 'processing' not in st.session_state:
        st.session_state.processing = False

# AWSクライアント設定
@st.cache_resource
def get_aws_clients():
    """AWSクライアントを取得"""
    try:
        # Streamlit CloudのSecretsから認証情報を取得
        aws_access_key_id = st.secrets["AWS_ACCESS_KEY_ID"]
        aws_secret_access_key = st.secrets["AWS_SECRET_ACCESS_KEY"]
        
        lambda_client = boto3.client(
            'lambda',
            region_name='ap-northeast-1',
            aws_access_key_id=aws_access_key_id,
            aws_secret_access_key=aws_secret_access_key
        )
        s3_client = boto3.client(
            's3',
            region_name='ap-northeast-1',
            aws_access_key_id=aws_access_key_id,
            aws_secret_access_key=aws_secret_access_key
        )
        return lambda_client, s3_client
    except Exception as e:
        st.error(f"AWS接続エラー: {e}")
        st.error("Streamlit Cloudのsecretsにaws_access_key_idとaws_secret_access_keyが設定されているか確認してください。")
        return None, None

# 設定値
BUCKET_NAME = "excel-ai-aggregator-6142"
REGION = "ap-northeast-1"

# Lambda関数の選択肢
LAMBDA_FUNCTIONS = {
    "Claude API版 (推奨)": "excel-claude-aggregator",
    "従来版": "excel-data-aggregator"
}

def get_excel_sheets(file_content):
    """Excelファイルからシート名のリストを取得"""
    try:
        import tempfile
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            tmp.write(file_content)
            tmp.flush()
            all_sheets = pd.read_excel(tmp.name, sheet_name=None, nrows=0)
            os.unlink(tmp.name)
            return list(all_sheets.keys())
    except Exception as e:
        st.warning(f"シート名を自動取得できませんでした。")
        return None

def main():
    """メイン関数"""
    init_session_state()
    
    st.title("📊 AI Excel集計システム (Claude API版)")
    st.markdown("---")
    
    # サイドバー
    st.sidebar.header("⚙️ システム設定")
    
    # AWS接続確認
    lambda_client, s3_client = get_aws_clients()
    if not lambda_client or not s3_client:
        st.error("AWS接続に失敗しました。認証情報を確認してください。")
        with st.expander("🔧 設定方法"):
            st.markdown("""
            ### Streamlit Cloudでの設定方法
            
            1. Streamlit Cloudのダッシュボードでアプリを選択
            2. 右上の歯車アイコン → Settings をクリック
            3. Secrets タブを選択
            4. 以下の形式でAWS認証情報を追加:
            
            ```toml
            AWS_ACCESS_KEY_ID = "your_access_key_here"
            AWS_SECRET_ACCESS_KEY = "your_secret_key_here"
            ```
            
            5. Save をクリック
            6. アプリを再起動
            """)
        return
    
    # Lambda関数選択
    try:
        selected_lambda = st.sidebar.selectbox(
            "使用するLambda関数",
            options=list(LAMBDA_FUNCTIONS.keys()),
            index=0,
            help="Claude API版は精度が高く推奨です"
        )
        
        lambda_function_name = LAMBDA_FUNCTIONS[selected_lambda]
        
        # Lambda関数の状態確認
        try:
            lambda_client.get_function(FunctionName=lambda_function_name)
            st.sidebar.success(f"✅ {selected_lambda} 利用可能")
        except Exception as e:
            st.sidebar.error(f"❌ {selected_lambda} が見つかりません")
            st.error(f"Lambda関数 '{lambda_function_name}' が見つかりません。先にデプロイを実行してください。")
            st.error(f"エラー詳細: {str(e)}")
            return
    except Exception as e:
        st.error(f"Lambda関数選択エラー: {e}")
        return
    
    # 会社名入力
    company_name = st.sidebar.text_input("会社名", value="株式会社テスト企業")
    
    st.sidebar.markdown("---")
    
    # ファイルアップロード
    st.sidebar.subheader("📁 ファイルアップロード")
    uploaded_files = st.sidebar.file_uploader(
        "集計対象のExcelファイルを選択",
        type=['xlsx', 'xls'],
        accept_multiple_files=True,
        key="file_uploader"
    )
    
    # メインエリア
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.subheader("📋 処理設定")
        
        if uploaded_files:
            st.success(f"✅ {len(uploaded_files)}個のファイルが選択されました")
            
            # ファイル別設定
            source_files_config = []
            
            for i, file in enumerate(uploaded_files):
                with st.container():
                    st.markdown(f"### 📄 {file.name}")
                    
                    # シート名設定
                    col_sheet, col_method = st.columns([2, 1])
                    
                    with col_method:
                        # 入力方法選択
                        method_key = f"method_{i}_{file.name}"
                        sheet_input_method = st.radio(
                            "入力方法",
                            ["手動入力", "自動検出"],
                            index=0,
                            key=method_key,
                            help="手動入力を推奨"
                        )
                    
                    with col_sheet:
                        if sheet_input_method == "手動入力":
                            # 一般的なシート名の選択肢
                            common_sheets = [
                                "受注ベース収支計画",
                                "Sheet1",
                                "【事業計画】PL推移", 
                                "PL - サマリー(年度)",
                                "損益計算書",
                                "PL",
                                "収支計画",
                                "事業計画"
                            ]
                            
                            sheet_key = f"sheet_{i}_{file.name}"
                            selected_sheet = st.selectbox(
                                "シート名を選択",
                                options=common_sheets,
                                index=0,  # デフォルトは「受注ベース収支計画」
                                key=sheet_key
                            )
                        else:
                            # 自動検出
                            try:
                                file_content = file.read()
                                file.seek(0)  # ファイルポインタをリセット
                                
                                sheet_names = get_excel_sheets(file_content)
                                
                                if sheet_names:
                                    auto_sheet_key = f"auto_sheet_{i}_{file.name}"
                                    selected_sheet = st.selectbox(
                                        "検出されたシート",
                                        options=sheet_names,
                                        index=0,
                                        key=auto_sheet_key
                                    )
                                else:
                                    selected_sheet = "Sheet1"
                                    st.warning("シート検出に失敗。Sheet1を使用します。")
                            except Exception as e:
                                selected_sheet = "Sheet1"
                                st.error(f"ファイル読み込みエラー: {e}")
                    
                    # 詳細設定
                    with st.expander(f"詳細設定 - {file.name}", expanded=False):
                        data_range_key = f"range_{i}_{file.name}"
                        data_range = st.text_input(
                            "データ範囲（例: A1:Z100）",
                            value="",
                            key=data_range_key,
                            help="空欄の場合は全データを対象とします"
                        )
                        
                        # Claude API版の特徴表示
                        if "claude" in lambda_function_name.lower():
                            st.info("""
                            🤖 **Claude API版の特徴**
                            - AI（Claude-4-Sonnet）がExcelデータを理解
                            - 項目名のゆらぎに強く、柔軟な認識が可能
                            - 数値の単位変換も自動で処理
                            - 抽出精度が大幅に向上
                            """)
                    
                    # 設定を保存
                    source_files_config.append({
                        "file_name": file.name,
                        "file_object": file,
                        "sheet_name": selected_sheet,
                        "data_range": data_range
                    })
            
            # 処理実行ボタン
            button_text = "🤖 Claude AI集計実行" if "claude" in lambda_function_name.lower() else "🚀 AI集計実行"
            
            if not st.session_state.processing:
                if st.button(button_text, type="primary", use_container_width=True, key="execute_button"):
                    st.session_state.processing = True
                    process_files(source_files_config, company_name, lambda_client, s3_client, lambda_function_name, selected_lambda)
                    st.session_state.processing = False
            else:
                st.info("処理中です...")
                
        else:
            st.info("📤 左のサイドバーからExcelファイルをアップロードしてください。")
            
            # 使用方法
            with st.expander("📖 使用方法", expanded=True):
                st.markdown("""
                ### 使い方
                1. **Lambda関数選択**: Claude API版（推奨）を選択
                2. **ファイルアップロード**: サイドバーから集計したいExcelファイルを選択
                3. **シート選択**: 「受注ベース収支計画」を選択（推奨）
                4. **実行**: 「Claude AI集計実行」ボタンをクリック
                
                ### Claude API版の優位性
                - **高精度**: AIがデータの意味を理解して抽出
                - **柔軟性**: 項目名のゆらぎに対応
                - **自動処理**: 単位変換や数値正規化を自動実行
                - **日本語対応**: 日本語の財務用語を深く理解
                
                ### 対応している項目
                - 売上高、売上原価、売上総利益
                - 販売費及び一般管理費、営業利益
                - 経常利益、当期純利益
                - その他のPL項目
                """)
    
    with col2:
        st.subheader("ℹ️ システム情報")
        
        # 使用中のLambda関数情報
        try:
            if "claude" in lambda_function_name.lower():
                status_color = "🤖"
                extraction_method = "Claude-4-Sonnet API"
                features = ["AI理解ベース", "高精度抽出", "柔軟認識"]
            else:
                status_color = "⚡"  
                extraction_method = "パターンマッチング"
                features = ["高速処理", "安定動作", "軽量"]
            
            st.info(f"""
            {status_color} **使用中: {selected_lambda}**
            
            **Lambda関数**: {lambda_function_name}
            **抽出方式**: {extraction_method}
            **S3バケット**: {BUCKET_NAME}
            **リージョン**: {REGION}
            **データ単位**: 百万円
            
            **特徴**: {" / ".join(features)}
            """)
            
            # S3バケット内容表示
            if st.button("📂 S3バケット確認", key="s3_check_button"):
                show_s3_contents(s3_client)
                
            # テンプレートダウンロード
            if st.button("📥 テンプレートダウンロード", key="template_download_button"):
                download_template(s3_client)
                
        except Exception as e:
            st.error(f"システム情報表示エラー: {e}")

def process_files(source_files_config, company_name, lambda_client, s3_client, lambda_function_name, selected_lambda):
    """ファイル処理を実行"""
    
    with st.spinner("📤 ファイルをS3にアップロード中..."):
        # S3にファイルをアップロード
        source_files = []
        
        for config in source_files_config:
            file_obj = config["file_object"]
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            s3_key = f"source_files/{timestamp}_{file_obj.name}"
            
            try:
                # ファイルをS3にアップロード
                file_obj.seek(0)  # ファイルポインタをリセット
                s3_client.put_object(
                    Bucket=BUCKET_NAME,
                    Key=s3_key,
                    Body=file_obj.read(),
                    ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
                
                # Lambda用の設定
                source_files.append({
                    "file_key": s3_key,
                    "sheet_name": config["sheet_name"],
                    "data_range": config.get("data_range", "")
                })
                
                st.success(f"✅ {file_obj.name} アップロード完了")
                
            except Exception as e:
                st.error(f"❌ {file_obj.name} アップロード失敗: {e}")
                return
    
    # Lambda実行
    processing_message = "🤖 Claude AIで処理中..." if "claude" in lambda_function_name.lower() else "🚀 AI集計処理中..."
    
    with st.spinner(processing_message):
        try:
            # Lambda関数呼び出し
            lambda_event = {
                "bucket": BUCKET_NAME,
                "input_template_key": "templates/input用.xlsx",
                "output_prefix": f"outputs/{datetime.now().strftime('%Y%m%d_%H%M%S')}/",
                "source_files": source_files,
                "company_name": company_name
            }
            
            response = lambda_client.invoke(
                FunctionName=lambda_function_name,
                Payload=json.dumps(lambda_event)
            )
            
            # レスポンス処理
            response_payload = json.loads(response['Payload'].read().decode('utf-8'))
            
            if response_payload.get('statusCode') == 200:
                body = json.loads(response_payload['body'])
                
                success_message = f"✅ {selected_lambda}での処理が完了しました！"
                st.success(success_message)
                
                # 結果表示
                st.subheader("📊 処理結果")
                
                results = body.get('results', [])
                
                # Claude版の場合の抽出データ表示
                if "claude" in lambda_function_name.lower():
                    st.subheader("🤖 Claude AIによる抽出データ")
                    
                    for result in results:
                        if result.get('extracted_data'):
                            st.markdown(f"#### 📄 {Path(result['source_file']).name}")
                            
                            extracted_data = result['extracted_data']
                            if extracted_data:
                                df_extracted = pd.DataFrame([
                                    {"項目": k, "抽出値": f"{v:,.0f}" if isinstance(v, (int, float)) else str(v)}
                                    for k, v in extracted_data.items()
                                ])
                                
                                st.dataframe(df_extracted, use_container_width=True)
                                st.success(f"✨ {len(extracted_data)}個の項目をClaude AIが自動認識・抽出しました")
                            else:
                                st.warning("データが抽出できませんでした")
                        else:
                            st.warning(f"📄 {Path(result['source_file']).name}: データ抽出に失敗")
                
                # 出力ファイルのダウンロードリンク表示
                processed_files = body.get('processed_files', [])
                if processed_files:
                    st.subheader("📥 生成されたファイル")
                    
                    for file_key in processed_files:
                        try:
                            # 署名付きURLを生成
                            download_url = s3_client.generate_presigned_url(
                                'get_object',
                                Params={'Bucket': BUCKET_NAME, 'Key': file_key},
                                ExpiresIn=3600
                            )
                            file_name = Path(file_key).name
                            
                            # ダウンロードボタンまたはリンク
                            col_download, col_info = st.columns([3, 1])
                            
                            with col_download:
                                st.markdown(f"📄 **{file_name}**")
                                st.markdown(f"[📥 ダウンロード]({download_url})")
                            
                            with col_info:
                                # ファイル情報表示
                                try:
                                    obj_info = s3_client.head_object(Bucket=BUCKET_NAME, Key=file_key)
                                    file_size = obj_info['ContentLength']
                                    st.text(f"サイズ: {file_size:,} bytes")
                                    st.text(f"更新: {obj_info['LastModified'].strftime('%H:%M:%S')}")
                                except:
                                    st.text("情報取得中...")
                            
                        except Exception as e:
                            st.error(f"ダウンロードリンク生成エラー: {e}")
                
                # 更新されたセル情報の表示
                st.subheader("🔄 テンプレート更新詳細")
                
                for result in results:
                    if result.get('updated_cells'):
                        st.markdown(f"#### 📄 {Path(result['source_file']).name}")
                        
                        # 更新されたセルの表示
                        updated_cells = result['updated_cells']
                        if updated_cells:
                            update_df = pd.DataFrame(updated_cells)
                            
                            # 見やすい形式に変換
                            display_df = pd.DataFrame([
                                {
                                    "セル": cell['cell'],
                                    "項目": cell['item'], 
                                    "旧値": cell['old_value'],
                                    "新値": f"{cell['new_value']:,.0f}" if isinstance(cell['new_value'], (int, float)) else str(cell['new_value']),
                                    "ソース": cell['source']
                                }
                                for cell in updated_cells
                            ])
                            
                            st.dataframe(display_df, use_container_width=True)
                            st.success(f"✅ {len(updated_cells)}個のセルを更新しました")
                        
                        # 使用期間情報
                        if result.get('used_period'):
                            st.info(f"🎯 使用期間: {result['used_period']}")
                
                # パフォーマンス情報
                total_items = body.get('total_extracted_items', 0)
                extraction_method = body.get('extraction_method', selected_lambda)
                
                st.info(f"""
                📈 **処理サマリー**
                - 使用システム: {extraction_method}
                - 処理ファイル数: {len(source_files)}
                - 抽出項目総数: {total_items}項目
                - 生成ファイル数: {len(processed_files)}個
                """)
                
                # 詳細情報
                with st.expander("🔍 詳細情報"):
                    st.json(body)
                    
            else:
                st.error(f"❌ {selected_lambda}での処理でエラーが発生しました")
                error_body = response_payload.get('body', '')
                st.error(error_body)
                
                # Claude版の場合のトラブルシューティング
                if "claude" in lambda_function_name.lower():
                    st.warning("""
                    🔧 **Claude API版でエラーが発生した場合:**
                    1. Claude APIキーが正しく設定されているか確認
                    2. AWS Secrets Managerの'claude-api-key'を確認
                    3. Lambda関数のタイムアウト設定を確認
                    4. ネットワーク接続とAPI制限を確認
                    """)
                
        except Exception as e:
            st.error(f"❌ Lambda実行エラー: {e}")
            
            # エラー詳細
            with st.expander("エラー詳細とサポート情報"):
                st.code(str(e))

def show_s3_contents(s3_client):
    """S3バケットの内容を表示"""
    try:
        response = s3_client.list_objects_v2(Bucket=BUCKET_NAME)
        
        if 'Contents' in response:
            objects = []
            for obj in response['Contents']:
                objects.append({
                    'ファイル名': obj['Key'],
                    'サイズ': f"{obj['Size']:,} bytes",
                    '最終更新': obj['LastModified'].strftime('%Y-%m-%d %H:%M:%S')
                })
            
            if objects:
                df = pd.DataFrame(objects)
                st.dataframe(df, use_container_width=True)
            else:
                st.info("バケットは空です")
        else:
            st.info("バケットは空です")
            
    except Exception as e:
        st.error(f"S3アクセスエラー: {e}")

def download_template(s3_client):
    """テンプレートファイルをダウンロード"""
    try:
        url = s3_client.generate_presigned_url(
            'get_object',
            Params={'Bucket': BUCKET_NAME, 'Key': 'templates/input用.xlsx'},
            ExpiresIn=3600
        )
        st.markdown(f"📄 [input用.xlsx テンプレートをダウンロード]({url})")
    except Exception as e:
        st.error(f"テンプレートダウンロードエラー: {e}")

if __name__ == "__main__":
    main()
