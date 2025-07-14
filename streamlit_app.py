#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Streamlit デモアプリケーション
Excel データ集計システムのUI
"""

import streamlit as st
import pandas as pd
import os
import tempfile
import zipfile
from pathlib import Path
import io
import sys

# カレントディレクトリを追加
sys.path.append(str(Path(__file__).parent))

try:
    from smart_aggregator import UniversalAggregator
    from data_aggregator import InputFileUpdater
except ImportError as e:
    st.error(f"モジュールのインポートエラー: {e}")
    st.stop()

def main():
    st.set_page_config(
        page_title="Excel データ集計システム",
        page_icon="📊",
        layout="wide"
    )
    
    st.title("📊 Excel データ集計システム")
    st.markdown("---")
    
    st.markdown("""
    このシステムは、事業計画ExcelファイルからPLデータを抽出し、
    input用ファイルのPL計画シートのC列に自動入力します。
    """)
    
    # サイドバー - 設定
    st.sidebar.header("⚙️ 設定")
    
    # ファイルアップロード
    st.header("📁 ファイルアップロード")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("input用テンプレートファイル")
        input_template_file = st.file_uploader(
            "input用.xlsxファイルをアップロードしてください",
            type=['xlsx'],
            key="input_template"
        )
    
    with col2:
        st.subheader("事業計画ファイル")
        source_files = st.file_uploader(
            "事業計画Excelファイルをアップロードしてください（複数可）",
            type=['xlsx'],
            accept_multiple_files=True,
            key="source_files"
        )
    
    # 処理実行
    if st.button("🚀 データ集計実行", type="primary", use_container_width=True):
        if input_template_file is None:
            st.error("input用テンプレートファイルをアップロードしてください")
        elif not source_files:
            st.error("事業計画ファイルを少なくとも1つアップロードしてください")
        else:
            process_files(input_template_file, source_files)

def process_files(input_template_file, source_files):
    """ファイル処理実行"""
    
    with st.spinner("データ集計を実行中..."):
        try:
            # 一時ディレクトリの作成
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_path = Path(temp_dir)
                
                # input用テンプレートファイルを保存
                template_path = temp_path / "input_template.xlsx"
                with open(template_path, "wb") as f:
                    f.write(input_template_file.read())
                
                # 集計システムの初期化
                aggregator = UniversalAggregator()
                
                results = []
                output_files = []
                
                # 各ソースファイルを処理
                for i, source_file in enumerate(source_files):
                    st.info(f"処理中: {source_file.name}")
                    
                    # ソースファイルを保存
                    source_path = temp_path / f"source_{i}.xlsx"
                    with open(source_path, "wb") as f:
                        f.write(source_file.read())
                    
                    # 出力ファイルパス
                    output_path = temp_path / f"output_{i}_{Path(source_file.name).stem}.xlsx"
                    
                    # データ抽出と更新
                    success = aggregator.process_any_file(
                        str(source_path),
                        str(template_path),
                        str(output_path)
                    )
                    
                    if success and output_path.exists():
                        results.append({
                            'ファイル名': source_file.name,
                            'ステータス': '成功 ✅',
                            '出力ファイル': output_path.name
                        })
                        output_files.append(output_path)
                    else:
                        results.append({
                            'ファイル名': source_file.name,
                            'ステータス': '失敗 ❌',
                            '出力ファイル': 'N/A'
                        })
                
                # 結果表示
                st.header("📊 処理結果")
                results_df = pd.DataFrame(results)
                st.dataframe(results_df, use_container_width=True)
                
                # 成功したファイルのダウンロード
                successful_files = [f for f in output_files if f.exists()]
                
                if successful_files:
                    st.header("📥 ダウンロード")
                    
                    if len(successful_files) == 1:
                        # 単一ファイルの場合
                        output_file = successful_files[0]
                        with open(output_file, "rb") as f:
                            st.download_button(
                                label=f"📄 {output_file.name} をダウンロード",
                                data=f.read(),
                                file_name=output_file.name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    else:
                        # 複数ファイルの場合はZIP
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                            for output_file in successful_files:
                                zip_file.write(output_file, output_file.name)
                        
                        st.download_button(
                            label=f"📦 すべてのファイルをZIPでダウンロード ({len(successful_files)}件)",
                            data=zip_buffer.getvalue(),
                            file_name="excel_aggregation_results.zip",
                            mime="application/zip"
                        )
                
                st.success(f"処理完了: {len(successful_files)}/{len(source_files)} ファイルが正常に処理されました")
                
        except Exception as e:
            st.error(f"処理中にエラーが発生しました: {e}")
            st.exception(e)

# デモデータ表示
def show_demo_info():
    """デモ情報の表示"""
    with st.expander("💡 使用方法とサンプル"):
        st.markdown("""
        ### 使用方法
        1. **input用テンプレートファイル**: `input用.xlsx`をアップロード
        2. **事業計画ファイル**: 処理したいExcelファイルを1つ以上アップロード
        3. **実行**: "データ集計実行"ボタンをクリック
        4. **ダウンロード**: 処理されたファイルをダウンロード
        
        ### 対応ファイル形式
        - 事業計画PL推移シート
        - 受注ベース収支計画シート  
        - PL - サマリー(四半期）シート
        - その他のPL関連シート（自動検出）
        
        ### 抽出される項目
        - 売上高
        - 売上原価
        - 売上総利益
        - 販売費及び一般管理費
        - 営業利益（損失）
        - 経常利益（損失）
        - 当期純利益
        - EBITDA、EBIT など
        """)

if __name__ == "__main__":
    main()
    show_demo_info()
