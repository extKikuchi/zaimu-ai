#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Streamlit Cloud対応版 Excel データ集計システム
AWS Lambda連携（smart_aggregator依存なし）
"""

import streamlit as st
import pandas as pd
import boto3
import json
import tempfile
import zipfile
from pathlib import Path
import io
import time
from datetime import datetime
import uuid
import os

# ページ設定
st.set_page_config(
    page_title="📊 Excel データ集計システム",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# AWS設定
@st.cache_data
def get_aws_config():
    """AWS設定を取得"""
    return {
        'region': st.secrets.get('AWS_DEFAULT_REGION', 'ap-northeast-1'),
        'lambda_function_name': st.secrets.get('LAMBDA_FUNCTION_NAME', 'excel-data-aggregator'),
        'bucket_name': st.secrets.get('S3_BUCKET_NAME', 'excel-aggregator-20250714-095126')
    }

@st.cache_resource
def init_aws_clients():
    """AWS クライアントの初期化"""
    try:
        aws_config = get_aws_config()
        
        # Streamlit Cloudのシークレットから認証情報を取得
        aws_access_key_id = st.secrets.get('AWS_ACCESS_KEY_ID')
        aws_secret_access_key = st.secrets.get('AWS_SECRET_ACCESS_KEY')
        
        if aws_access_key_id and aws_secret_access_key:
            session = boto3.Session(
                aws_access_key_id=aws_access_key_id,
                aws_secret_access_key=aws_secret_access_key,
                region_name=aws_config['region']
            )
        else:
            session = boto3.Session(region_name=aws_config['region'])
        
        s3_client = session.client('s3')
        lambda_client = session.client('lambda')
        
        # 接続テスト
        s3_client.head_bucket(Bucket=aws_config['bucket_name'])
        lambda_client.get_function(FunctionName=aws_config['lambda_function_name'])
        
        return s3_client, lambda_client, True, None
    except Exception as e:
        return None, None, False, str(e)

def upload_file_to_s3(file_obj, s3_client, key):
    """ファイルをS3にアップロード"""
    try:
        aws_config = get_aws_config()
        s3_client.upload_fileobj(file_obj, aws_config['bucket_name'], key)
        return True
    except Exception as e:
        st.error(f"S3アップロードエラー: {e}")
        return False

def invoke_lambda_function(lambda_client, payload):
    """Lambda関数を実行"""
    try:
        aws_config = get_aws_config()
        response = lambda_client.invoke(
            FunctionName=aws_config['lambda_function_name'],
            Payload=json.dumps(payload)
        )
        
        result = json.loads(response['Payload'].read().decode())
        return result
    except Exception as e:
        st.error(f"Lambda実行エラー: {e}")
        return None

def download_file_from_s3(s3_client, key):
    """S3からファイルをダウンロード"""
    try:
        aws_config = get_aws_config()
        response = s3_client.get_object(Bucket=aws_config['bucket_name'], Key=key)
        return response['Body'].read()
    except Exception as e:
        st.error(f"S3ダウンロードエラー: {e}")
        return None

def setup_sidebar():
    """サイドバーの設定"""
    st.sidebar.header("📋 システム情報")
    
    st.sidebar.markdown("""
    ### 🎯 Excel集計システム
    - **バージョン**: 1.0.0
    - **最終更新**: 2025-07-14
    - **環境**: Streamlit Cloud
    """)
    
    st.sidebar.markdown("""
    ### 📊 対応項目
    - 売上高
    - 売上原価  
    - 売上総利益
    - 営業利益
    - 経常利益
    - 当期純利益
    - EBITDA
    - EBIT
    """)
    
    st.sidebar.markdown("""
    ### 💡 ヘルプ
    - 最大ファイルサイズ: 50MB
    - 対応形式: .xlsx, .xls
    - 複数ファイル同時処理対応
    """)

def show_system_stats(s3_client):
    """システム統計の表示"""
    try:
        aws_config = get_aws_config()
        
        # S3バケットの内容を確認
        response = s3_client.list_objects_v2(Bucket=aws_config['bucket_name'])
        
        if 'Contents' in response:
            total_files = len(response['Contents'])
            total_size = sum(obj['Size'] for obj in response['Contents'])
            
            col1, col2 = st.columns(2)
            with col1:
                st.metric("📁 総ファイル数", total_files)
            with col2:
                st.metric("💾 総サイズ", f"{total_size / 1024 / 1024:.1f} MB")
            
            # 最近の処理状況
            output_files = [obj for obj in response['Contents'] if obj['Key'].startswith('outputs/')]
            if output_files:
                latest_file = max(output_files, key=lambda x: x['LastModified'])
                st.metric("📅 最新処理", latest_file['LastModified'].strftime('%m/%d %H:%M'))
        else:
            st.info("📭 処理履歴なし")
    except Exception as e:
        st.warning(f"⚠️ 統計取得エラー: {e}")

def download_files(s3_client, processed_files, zip_results):
    """ファイルダウンロード処理"""
    
    if len(processed_files) == 1:
        # 単一ファイルの場合
        file_key = processed_files[0]
        file_data = download_file_from_s3(s3_client, file_key)
        
        if file_data:
            file_name = Path(file_key).name
            st.download_button(
                label=f"📄 {file_name} をダウンロード",
                data=file_data,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    elif zip_results:
        # 複数ファイルをZIPで圧縮
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for file_key in processed_files:
                file_data = download_file_from_s3(s3_client, file_key)
                if file_data:
                    file_name = Path(file_key).name
                    zip_file.writestr(file_name, file_data)
        
        zip_buffer.seek(0)
        
        st.download_button(
            label=f"📦 全結果ファイルをZIPでダウンロード ({len(processed_files)}件)",
            data=zip_buffer.getvalue(),
            file_name=f"excel_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
            mime="application/zip"
        )
    
    else:
        # 個別ダウンロード
        for i, file_key in enumerate(processed_files):
            file_data = download_file_from_s3(s3_client, file_key)
            if file_data:
                file_name = Path(file_key).name
                st.download_button(
                    label=f"📄 {file_name}",
                    data=file_data,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"download_{i}"
                )

def cleanup_files(s3_client, template_key, source_keys):
    """一時ファイルのクリーンアップ"""
    try:
        aws_config = get_aws_config()
        
        # テンプレートファイルとソースファイルを削除
        keys_to_delete = [template_key] + source_keys
        
        for key in keys_to_delete:
            s3_client.delete_object(Bucket=aws_config['bucket_name'], Key=key)
        
        st.info(f"🧹 {len(keys_to_delete)} 個の一時ファイルをクリーンアップしました")
    except Exception as e:
        st.warning(f"⚠️ クリーンアップエラー: {e}")

def process_files(s3_client, lambda_client, input_template_file, source_files, 
                 show_progress, auto_download, zip_results, keep_files):
    """ファイル処理のメイン関数"""
    
    # 処理ID生成
    process_id = str(uuid.uuid4())[:8]
    
    # プログレスバーとステータス
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        # ステップ1: ファイルをS3にアップロード
        status_text.text("📤 ファイルをS3にアップロード中...")
        progress_bar.progress(10)
        
        # テンプレートファイルをアップロード
        template_key = f"templates/{process_id}_input.xlsx"
        input_template_file.seek(0)  # ファイルポインタをリセット
        if not upload_file_to_s3(input_template_file, s3_client, template_key):
            st.error("❌ テンプレートファイルのアップロードに失敗しました")
            return
        
        # ソースファイルをアップロード
        source_keys = []
        for i, source_file in enumerate(source_files):
            source_key = f"source-files/{process_id}_{i}_{source_file.name}"
            source_file.seek(0)  # ファイルポインタをリセット
            if upload_file_to_s3(source_file, s3_client, source_key):
                source_keys.append(source_key)
                progress_bar.progress(10 + (i + 1) * 25 / len(source_files))
            else:
                st.error(f"❌ {source_file.name} のアップロードに失敗しました")
                return
        
        # ステップ2: Lambda関数を実行
        status_text.text("⚡ Lambda関数を実行中...")
        progress_bar.progress(40)
        
        aws_config = get_aws_config()
        
        # Lambda実行用のペイロード
        payload = {
            "bucket": aws_config['bucket_name'],
            "input_template_key": template_key,
            "source_files": source_keys,
            "output_prefix": f"outputs/{process_id}_"
        }
        
        if show_progress:
            with st.expander("🔍 実行詳細"):
                st.json(payload)
        
        # Lambda関数実行
        lambda_result = invoke_lambda_function(lambda_client, payload)
        
        if not lambda_result:
            st.error("❌ Lambda関数の実行に失敗しました")
            return
        
        progress_bar.progress(60)
        
        # ステップ3: 結果の処理
        status_text.text("📊 結果を処理中...")
        
        if lambda_result.get('statusCode') == 200:
            body = json.loads(lambda_result['body'])
            results = body.get('results', [])
            processed_files = body.get('processed_files', [])
            
            progress_bar.progress(80)
            
            # 結果表示
            st.header("📋 処理結果")
            
            # 結果テーブル
            if results:
                results_df = pd.DataFrame(results)
                st.dataframe(results_df, use_container_width=True)
                
                # 成功/失敗の統計
                success_count = len([r for r in results if r['status'] == 'success'])
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric("✅ 成功", success_count)
                with col2:
                    st.metric("❌ 失敗", len(results) - success_count)
                with col3:
                    st.metric("📊 総抽出項目", sum(r.get('extracted_items', 0) for r in results))
            
            # ステップ4: ファイルダウンロード
            if processed_files:
                status_text.text("📥 ダウンロード準備中...")
                progress_bar.progress(90)
                
                st.header("📥 ダウンロード")
                
                download_files(s3_client, processed_files, zip_results)
                
                # ファイルクリーンアップ
                if not keep_files:
                    cleanup_files(s3_client, template_key, source_keys)
            
            progress_bar.progress(100)
            status_text.text("✅ 処理完了")
            
            st.success(f"🎉 処理完了: {len(processed_files)} ファイルが正常に処理されました")
            st.balloons()
        else:
            st.error(f"❌ Lambda実行エラー: {lambda_result}")
            
    except Exception as e:
        st.error(f"❌ 処理中にエラーが発生しました: {e}")
        with st.expander("🔍 エラー詳細"):
            st.exception(e)

def show_aws_setup_guide():
    """AWS設定ガイドの表示"""
    st.error("### ❌ AWS設定が必要です")
    
    with st.expander("🔧 AWS設定ガイド", expanded=True):
        st.markdown("""
        ### Streamlit Cloud でのAWS設定手順
        
        1. **Streamlit Cloud アプリの設定画面**にアクセス
        2. **⚙️ Settings** → **🔐 Secrets** タブを選択
        3. 以下の内容をコピー＆ペーストして **Save** をクリック：
        
        ```toml
        AWS_ACCESS_KEY_ID = "your_access_key_id"
        AWS_SECRET_ACCESS_KEY = "your_secret_access_key"
        AWS_DEFAULT_REGION = "ap-northeast-1"
        S3_BUCKET_NAME = "excel-aggregator-20250714-095126"
        LAMBDA_FUNCTION_NAME = "excel-data-aggregator"
        ```
        
        ### ⚠️ 重要な注意事項
        - AWS認証情報は**絶対に**GitHubリポジトリにコミットしないでください
        - Streamlit CloudのSecretsは暗号化されて安全に保存されます
        - 設定後、アプリを再起動してください
        
        ### 🔑 必要な AWS 権限
        - **S3**: GetObject, PutObject, DeleteObject, ListBucket
        - **Lambda**: InvokeFunction
        """)

def main():
    """メイン関数"""
    
    # ヘッダー
    st.title("📊 Excel データ集計システム")
    st.markdown("### 🚀 AWS Lambda連携 事業計画データ自動抽出システム")
    st.markdown("---")
    
    # サイドバー情報
    setup_sidebar()
    
    # AWS接続状態の確認
    with st.spinner("🔍 AWS接続を確認中..."):
        s3_client, lambda_client, aws_connected, error_message = init_aws_clients()
    
    if not aws_connected:
        st.error("❌ AWS接続に失敗しました")
        if error_message:
            st.code(f"エラー詳細: {error_message}")
        show_aws_setup_guide()
        return
    
    st.success("✅ AWS接続成功")
    
    # メインコンテンツ
    col1, col2 = st.columns([3, 1])
    
    with col1:
        st.header("📁 ファイルアップロード")
        
        # input用テンプレートファイル
        st.subheader("🎯 input用テンプレートファイル")
        input_template_file = st.file_uploader(
            "input用.xlsxファイルをアップロードしてください",
            type=['xlsx'],
            key="input_template",
            help="集計結果を入力するためのテンプレートファイル"
        )
        
        if input_template_file:
            st.success(f"✅ {input_template_file.name} が選択されました")
        
        # 事業計画ファイル
        st.subheader("📈 事業計画ファイル")
        source_files = st.file_uploader(
            "事業計画Excelファイルをアップロードしてください（複数可）",
            type=['xlsx', 'xls'],
            accept_multiple_files=True,
            key="source_files",
            help="データを抽出する元のExcelファイル"
        )
        
        if source_files:
            st.success(f"✅ {len(source_files)} ファイルが選択されました")
            for i, file in enumerate(source_files):
                st.write(f"  {i+1}. 📄 {file.name}")
        
        # 処理オプション
        st.subheader("⚙️ 処理オプション")
        col3, col4 = st.columns(2)
        
        with col3:
            show_progress = st.checkbox("処理状況を表示", value=True)
            zip_results = st.checkbox("結果をZIPで圧縮", value=len(source_files) > 1 if source_files else False)
        
        with col4:
            auto_download = st.checkbox("自動ダウンロード", value=True)
            keep_files = st.checkbox("S3にファイルを保持", value=False)
    
    with col2:
        st.header("📊 システム状態")
        
        # AWS設定情報
        aws_config = get_aws_config()
        st.info(f"🪣 **S3バケット**\n{aws_config['bucket_name']}")
        st.info(f"⚡ **Lambda関数**\n{aws_config['lambda_function_name']}")
        st.info(f"🌍 **リージョン**\n{aws_config['region']}")
        
        # システム統計
        show_system_stats(s3_client)
        
        # ヘルプ
        with st.expander("💡 使用方法"):
            st.markdown("""
            **基本的な使用方法:**
            1. input用テンプレートファイルをアップロード
            2. 事業計画ファイルをアップロード
            3. "データ集計実行"をクリック
            4. 処理結果をダウンロード
            
            **対応ファイル:**
            - Excel (.xlsx, .xls)
            - 最大50MB
            
            **抽出される項目:**
            - 売上高、売上原価、営業利益
            - 経常利益、当期純利益 など
            """)
    
    # 処理実行ボタン
    st.markdown("---")
    if st.button("🚀 データ集計実行", type="primary", use_container_width=True):
        if not input_template_file:
            st.error("❌ input用テンプレートファイルをアップロードしてください")
        elif not source_files:
            st.error("❌ 事業計画ファイルを少なくとも1つアップロードしてください")
        else:
            process_files(
                s3_client,
                lambda_client,
                input_template_file,
                source_files,
                show_progress,
                auto_download,
                zip_results,
                keep_files
            )

if __name__ == "__main__":
    main()
