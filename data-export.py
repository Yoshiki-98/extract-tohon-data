import streamlit as st
import pdfplumber
import pandas as pd
import tempfile
import os
from pathlib import Path
import time
from datetime import datetime

def clean_address_line(line):
    """
    住所行の文字列をクリーンアップする関数
    """
    # 区切り線を含む行は無視
    if any(c in '┠┨┝┥━┿╂╋├┼┤─' for c in line):
        return ''

    # 余分な文字を削除
    cleaned = line.replace('┃', '').replace('┨', '').replace('│', '').strip()
    # 全角スペースを半角に統一
    cleaned = cleaned.replace('　', ' ')
    # 連続する空白を1つに
    cleaned = ' '.join(cleaned.split())
    return cleaned

def process_multiple_pdfs(pdf_dir):
    """
    指定されたディレクトリ内の全PDFファイルを処理する
    """
    pdf_files = list(Path(pdf_dir).glob('*.pdf')) + list(Path(pdf_dir).glob('*.PDF'))
    if not pdf_files:
        raise Exception("PDFファイルが見つかりません")

    all_data = []
    for pdf_path in pdf_files:
        print(f"\n=== {pdf_path.name} の処理を開始 ===")
        df = extract_info_from_pdf(pdf_path)
        if not df.empty:
            all_data.append(df)
            print(f"{pdf_path.name}: {len(df)}件のデータを抽出")

    if all_data:
        combined_df = pd.concat(all_data, ignore_index=True)
        print(f"\n合計: {len(combined_df)}件のデータを抽出")
        return combined_df
    else:
        raise Exception("データを抽出できませんでした")

def validate_extracted_data(df):
    """
    抽出されたデータの検証を行う
    """
    validation_results = {
        'total_records': len(df),
        'unique_addresses': df['住所'].nunique(),
        'unique_names': df['氏名'].nunique(),
        'duplicates': df.duplicated().sum()
    }

    print("\n=== データ検証結果 ===")
    print(f"総レコード数: {validation_results['total_records']}")
    print(f"ユニーク住所数: {validation_results['unique_addresses']}")
    print(f"ユニーク氏名数: {validation_results['unique_names']}")
    print(f"重複レコード数: {validation_results['duplicates']}")

    return validation_results

def extract_info_from_pdf(pdf_path):
    data = []
    location = None
    current_address = []
    current_name = None

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                text = decode_text(text)
                print(f"\n処理中のファイル: {pdf_path.name}")
                print("抽出されたテキスト（一部）:", text[:200])

                if location is None:
                    location = extract_location_from_header(text)
                    print(f"所在地: {location}")

                try:
                    lines = text.split('\n')
                    for i, line in enumerate(lines):
                        print(f"\n処理行: {line}")
                        # ヘッダー行または区切り線をスキップ
                        if any(header in line for header in ['住 所', '持 分', '氏 名', '所 有 者', '共 有 者']) or \
                            any(c in '┠┨┝┥━┿╂╋├┼┤─' for c in line):
                            print("→ ヘッダーまたは区切り線をスキップ")
                            continue

                        # 空行をスキップ
                        if not line.strip():
                            continue

                        if '│' in line:
                            parts = line.split('│')
                            # 新しい人の行（氏名を含む行）の場合
                            if len(parts) >= 2 and parts[-1].strip() and not parts[-1].strip() == '┃':
                                # 前の人のデータがあれば保存
                                if current_name and current_address:
                                    complete_address = ' '.join(current_address)
                                    if complete_address:
                                        data.append({
                                            '氏名': current_name,
                                            '郵便番号': '',
                                            '住所': complete_address,
                                            '所在地': location if location else ''
                                        })
                                        print(f"前のデータを保存: 住所={complete_address}, 氏名={current_name}")
                                
                                # 新しい人のデータを開始
                                current_address = [clean_address_line(parts[0])]
                                current_name = clean_address_line(parts[-1])
                                print(f"新しい人のデータ開始: 住所={current_address}, 氏名={current_name}")
                            else:
                                # 住所の続きの行
                                cleaned_line = clean_address_line(parts[0])
                                if cleaned_line:
                                    current_address.append(cleaned_line)
                                    print(f"住所に追加: {cleaned_line}")
                except Exception as e:
                    print(f"エラー発生: {str(e)}")
                    print(f"問題の行: {line}")

                # ページ終了時に残っているデータを保存
                if current_name and current_address:
                    complete_address = ' '.join(current_address)
                    if complete_address:
                        data.append({
                            '氏名': current_name,
                            '郵便番号': '',
                            '住所': complete_address,
                            '所在地': location if location else ''
                        })
                        print(f"最終データを保存: 住所={complete_address}, 氏名={current_name}")
                    current_name = None
                    current_address = []

    return pd.DataFrame(data)

def decode_text(text):
    """
    文字コードの問題に対応するための関数
    """
    encodings = ['utf-8', 'cp932', 'shift_jis', 'euc-jp']
    for encoding in encodings:
        try:
            return text.encode(encoding, 'ignore').decode(encoding)
        except:
            continue
    return text

def extract_location_from_header(text):
    """
    ヘッダー部分から所在地を抽出する関数
    """
    lines = text.split('\n')
    for i, line in enumerate(lines):
        if '現在の情報です。' in line and i + 1 < len(lines):
            location_line = lines[i + 1]
            location = location_line.split('所有者一覧表')[0].strip()
            return location
    return None

def save_to_excel(df, output_path):
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='抽出データ', index=False)
            worksheet = writer.sheets['抽出データ']

            # 列幅の自動調整
            for idx, col in enumerate(['氏名', '郵便番号', '住所', '所在地']):
                max_length = max(
                    df[col].astype(str).apply(len).max(),
                    len(col)
                ) + 2
                worksheet.column_dimensions[chr(65 + idx)].width = max_length

            # 郵便番号列にWEBSERVICE関数を設定
            for row in range(2, len(df) + 2):
                cell = worksheet.cell(row=row, column=2)
                cell.value = f'=WEBSERVICE("http://api.excelapi.org/post/zipcode?address="&ENCODEURL(C{row}))'

    except Exception as e:
        print(f"Excel保存中にエラーが発生: {str(e)}")

def create_streamlit_app():
    st.title("謄本データ抽出アプリ")

    # ファイルアップロード
    uploaded_files = st.file_uploader("PDFファイルを選択してください", 
                                    type=['pdf'], 
                                    accept_multiple_files=True)

    if uploaded_files:
        st.write(f"{len(uploaded_files)}個のファイルがアップロードされました")

        # 処理開始ボタン
        if st.button("処理開始"):
            # 進捗バーの表示
            progress_bar = st.progress(0)
            status_text = st.empty()

            # 一時ディレクトリの作成
            with tempfile.TemporaryDirectory() as temp_dir:
                # ファイルの保存
                for uploaded_file in uploaded_files:
                    file_path = os.path.join(temp_dir, uploaded_file.name)
                    with open(file_path, 'wb') as f:
                        f.write(uploaded_file.getvalue())

                try:
                    # データ処理
                    status_text.text("PDFからデータを抽出中...")
                    progress_bar.progress(30)

                    combined_df = process_multiple_pdfs(temp_dir)
                    progress_bar.progress(60)

                    if not combined_df.empty:
                        # 検証結果の表示
                        validation_results = validate_extracted_data(combined_df)
                        st.write("### データ検証結果")
                        st.write(f"- 総レコード数: {validation_results['total_records']}")
                        st.write(f"- ユニーク住所数: {validation_results['unique_addresses']}")
                        st.write(f"- ユニーク氏名数: {validation_results['unique_names']}")
                        st.write(f"- 重複レコード数: {validation_results['duplicates']}")
                        
                        progress_bar.progress(80)

                        # Excelファイルの作成
                        status_text.text("Excelファイルを作成中...")
                        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                        output_filename = f"抽出データ_統合_{timestamp}.xlsx"
                        output_path = os.path.join(temp_dir, output_filename)
                        save_to_excel(combined_df, output_path)
                        
                        # ダウンロードボタンの作成
                        with open(output_path, 'rb') as f:
                            excel_data = f.read()
                        
                        progress_bar.progress(100)
                        status_text.text("処理完了！")
                        
                        st.download_button(
                            label="Excelファイルをダウンロード",
                            data=excel_data,
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

                        # データプレビューの表示
                        st.write("### データプレビュー")
                        st.markdown("""
                            <style>
                                .dataframe {
                                    font-size: 12px !important;
                                    white-space: nowrap !important;
                                    width: 100% !important;
                                }
                                .dataframe td, .dataframe th {
                                    padding: 4px !important;
                                    max-width: 200px !important;
                                    overflow: hidden !important;
                                    text-overflow: ellipsis !important;
                                }
                            </style>
                        """, unsafe_allow_html=True)

                        # DataFrameの表示設定
                        pd.set_option('display.max_columns', None)
                        pd.set_option('display.max_colwidth', None)
                        pd.set_option('display.width', None)

                        # データプレビューの表示（スクロール可能な形式で）
                        st.dataframe(
                            combined_df,
                            height=300,  # 高さを固定
                            use_container_width=True  # 幅をコンテナに合わせる
                        )

                except Exception as e:
                    st.error(f"エラーが発生しました: {str(e)}")
                    progress_bar.progress(100)
                    status_text.text("エラーが発生しました")

    # 使い方の説明
    with st.expander("使い方"):
        st.write("""
        1. 「PDFファイルを選択してください」から処理したいPDFファイルを選択します（複数選択可）
        2. 「処理開始」ボタンをクリックします
        3. 処理が完了すると、Excelファイルがダウンロード可能になります
        
        注意事項:
        - アップロード可能なファイル形式は PDF のみです
        - 特殊文字を含む氏名は Excel ファイル内で □ で出力されます
        - 処理には時間がかかる場合があります
        """)
    
    # フッター
    st.markdown("---")
    st.markdown("©2025 ARKNEXT株式会社 All Rights Reserved.")

if __name__ == "__main__":
    create_streamlit_app()
