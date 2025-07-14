#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Streamlit Cloud対応版 Excel データ集計システム
AWS Lambda連携（S3バケット自動作成対応）
"""

import streamlit as st
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
    
    # デフォルトバケット名を動的に生成
    default_bucket = st.secrets.get('S3_BUCKET_NAME', 'excel-aggregator-backup-20250714')
    
    config = {
        'region': st.secrets.get('AWS_DEFAULT_REGION', 'ap-northeast-1'),
        'lambda_function_name': st.secrets.get('LAMBDA_FUNCTION_NAME', 'excel-data-aggregator'),
        'bucket_name': default_bucket
    }
    
    return config

def get_or_create_bucket(s3_client):
    """S3バケットを取得または作成"""
    aws_config = get_aws_config()
    bucket_name = aws_config['bucket_name']
    
    try:
        # 既存バケットの確認
        s3_client.head_bucket(Bucket=bucket_name)
        st.success(f"✅ 既存のS3バケットを使用: {bucket_name}")
        return bucket_name
    except Exception as e:
        if "403" in str(e) or "Forbidden" in str(e):
            st.warning(f"⚠️ バケット {bucket_name} へのアクセス権限がありません")
            
            # 利用可能なバケットを確認
            try:
                response = s3_client.list_buckets()
                available_buckets = [bucket['Name'] for bucket in response['Buckets']]
                
                if available_buckets:
                    st.info("📋 利用可能なS3バケット:")
                    for bucket in available_buckets[:5]:  # 最初の5個を表示
                        st.write(f"   • {bucket}")
                    
                    # 最初の利用可能なバケットを使用
                    selected_bucket = available_buckets[0]
                    st.success(f"✅ 利用可能なバケットを使用: {selected_bucket}")
                    return selected_bucket
                else:
                    # 新しいバケットを作成
                    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
                    new_bucket_name = f"excel-aggregator-{timestamp}"
                    
                    try:
                        if aws_config['region'] == 'us-east-1':
                            s3_client.create_bucket(Bucket=new_bucket_name)
                        else:
                            s3_client.create_bucket(
                                Bucket=new_bucket_name,
                                CreateBucketConfiguration={'LocationConstraint': aws_config['region']}
                            )
                        
                        st.success(f"✅ 新しいS3バケットを作成: {new_bucket_name}")
                        return new_bucket_name
                    except Exception as create_error:
                        st.error(f"❌ バケット作成エラー: {create_error}")
                        return None
            except Exception as list_error:
                st.error(f"❌ バケット一覧取得エラー: {list_error}")
                return None
        else:
            st.error(f"❌ バケット確認エラー: {e}")
            return None

@st.cache_resource
def init_aws_clients():
    """AWS クライアントの初期化"""
    try:
        aws_config = get_aws_config()
        
        # Streamlit Cloudのシークレットから認証情報を取得
        aws_access_key_id = st.secrets.get('AWS_ACCESS_KEY_ID')
        aws_secret_access_key = st.secrets.get('AWS_SECRET_ACCESS_KEY')
        
        # デバッグ情報を表示
        st.sidebar.write("🔍 認証情報デバッグ:")
        st.sidebar.write(f"Access Key: {aws_access_key_id[:10]}..." if aws_access_key_id else "Access Key: 未設定")
        st.sidebar.write(f"Secret Key: {'***設定済み***' if aws_secret_access_key else '未設定'}")
        st.sidebar.write(f"Region: {aws_config['region']}")
        
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
        
        # 認証情報のテスト
        try:
            caller_identity = session.client('sts').get_caller_identity()
            st.sidebar.write(f"👤 AWSユーザー: {caller_identity.get('Arn', 'Unknown')}")
        except Exception as auth_error:
            st.sidebar.error(f"認証エラー: {auth_error}")
            return None, None, False, f"認証エラー: {auth_error}"
        
        # S3バケットの確認・作成
        bucket_name = get_or_create_bucket(s3_client)
        if not bucket_name:
            return None, None, False, "S3バケットの設定に失敗しました"
        
        # Lambda関数の確認
        try:
            lambda_client.get_function(FunctionName=aws_config['lambda_function_name'])
        except Exception as lambda_error:
            st.sidebar.warning(f"Lambda関数エラー: {lambda_error}")
            return s3_client, lambda_client, True, bucket_name  # S3は成功しているので続行
        
        return s3_client, lambda_client, True, bucket_name
        
    except Exception as e:
        st.sidebar.error(f"初期化エラー: {e}")
        return None, None, False, str(e)

def upload_file_to_s3(file_obj, s3_client, key, bucket_name):
    """ファイルをS3にアップロード"""
    try:
        s3_client.upload_fileobj(file_obj, bucket_name, key)
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

def download_file_from_s3(s3_client, key, bucket_name):
    """S3からファイルをダウンロード"""
    try:
        response = s3_client.get_object(Bucket=bucket_name, Key=key)
        return response['Body'].read()
    except Exception as e:
        st.error(f"S3ダウンロードエラー: {e}")
        return None

def setup_sidebar():
    """サイドバーの設定"""
    st.sidebar.header("📋 システム情報")
    
    st.sidebar.markdown("""
    ### 🎯 Excel集計システム
    - **バージョン**: 1.0.1
    - **最終更新**: 2025-07-14
    - **環境**: Streamlit Cloud
    """)
    
    st.sidebar.markdown("""
    ### 📊 対応項目
    - 売上高、売上原価、売上総利益
    - 営業利益、経常利益、当期純利益
    - EBITDA、EBIT
    """)

def show_system_stats(s3_client, bucket_name):
    """システム統計の表示"""
    try:
        # S3バケットの内容を確認
        response = s3_client.list_objects_v2(Bucket=bucket_name)
        
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
        s3_client, lambda_client, aws_connected, result = init_aws_clients()
    
    if not aws_connected:
        st.error("❌ AWS接続に失敗しました")
        st.code(f"エラー詳細: {result}")
        show_aws_setup_guide()
        return
    
    # 接続成功時はresultはbucket_name
    bucket_name = result
    st.success("✅ AWS接続成功")
    st.info(f"🪣 使用中のS3バケット: {bucket_name}")
    
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
        st.info(f"⚡ **Lambda関数**\n{aws_config['lambda_function_name']}")
        st.info(f"🌍 **リージョン**\n{aws_config['region']}")
        
        # システム統計
        show_system_stats(s3_client, bucket_name)
        
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
                bucket_name,
                input_template_file,
                source_files,
                show_progress,
                auto_download,
                zip_results,
                keep_files
            )

def process_files(s3_client, lambda_client, bucket_name, input_template_file, source_files, 
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
        input_template_file.seek(0)
        if not upload_file_to_s3(input_template_file, s3_client, template_key, bucket_name):
            st.error("❌ テンプレートファイルのアップロードに失敗しました")
            return
        
        # ソースファイルをアップロード
        source_keys = []
        for i, source_file in enumerate(source_files):
            source_key = f"source-files/{process_id}_{i}_{source_file.name}"
            source_file.seek(0)
            if upload_file_to_s3(source_file, s3_client, source_key, bucket_name):
                source_keys.append(source_key)
                progress_bar.progress(10 + (i + 1) * 25 / len(source_files))
            else:
                st.error(f"❌ {source_file.name} のアップロードに失敗しました")
                return
        
        progress_bar.progress(100)
        status_text.text("✅ ファイルアップロード完了")
        
        st.success(f"🎉 {len(source_files) + 1} ファイルのアップロードが完了しました")
        st.info("💡 Lambda関数との連携は準備中です。現在はファイルアップロード機能のみ動作します。")
        
    except Exception as e:
        st.error(f"❌ 処理中にエラーが発生しました: {e}")
        with st.expander("🔍 エラー詳細"):
            st.exception(e)

def show_aws_setup_guide():
    """AWS設定ガイドの表示"""
    st.error("### ❌ AWS設定が必要です")
    
    with st.expander("🔧 AWS設定ガイド", expanded=True):
        st.markdown("""
        ### S3バケットアクセス権限の設定
        
        1. **AWS Console** → **IAM** → **Users** → あなたのユーザー
        2. **Permissions** タブ → **Add permissions** → **Attach policies directly**
        3. 以下のポリシーを追加:
           - `AmazonS3FullAccess`
           - `AWSLambdaInvokeFunction`
        
        ### または新しいS3バケットを作成
        
        ```bash
        aws s3 mb s3://your-new-bucket-name --region ap-northeast-1
        ```
        
        ### Streamlit Cloud Secrets設定
        
        ```toml
        AWS_ACCESS_KEY_ID = "your_access_key"
        AWS_SECRET_ACCESS_KEY = "your_secret_key"
        AWS_DEFAULT_REGION = "ap-northeast-1"
        S3_BUCKET_NAME = "your-new-bucket-name"
        LAMBDA_FUNCTION_NAME = "excel-data-aggregator"
        ```
        """)

if __name__ == "__main__":
    main()
