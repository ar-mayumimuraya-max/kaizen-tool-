import streamlit as st
import pandas as pd
import io

# ページ設定
st.set_page_config(page_title="業務データ加工ツール（マルチクライアント）", layout="wide")
st.title("📊 業務データ加工ツール")

# ▼▼▼ 1. パターン選択ラジオボタン ▼▼▼
client_option = st.radio(
    "作成するデータのパターンを選択してください",
    ("バリューブックス用 (Excel出力)", "リンクシェア用 (テキスト出力)"),
    horizontal=True
)

st.markdown("---")

# 2. ファイルアップロード（共通）
uploaded_files = st.file_uploader(
    "CSVファイルをドラッグ＆ドロップしてください（複数可）", 
    type=['csv'], 
    accept_multiple_files=True
)

if uploaded_files:
    df_list = []
    
    # ファイル読み込み（共通）
    for file in uploaded_files:
        try:
            try:
                df_temp = pd.read_csv(file, encoding='utf-8')
            except UnicodeDecodeError:
                file.seek(0)
                df_temp = pd.read_csv(file, encoding='cp932')
            df_list.append(df_temp)
        except Exception as e:
            st.error(f"ファイル {file.name} の読み込みに失敗しました: {e}")

    if len(df_list) > 0:
        try:
            # データの結合（共通）
            df = pd.concat(df_list, ignore_index=True)

            # ---------------------------------------------------------
            # パターンA：バリューブックス用（いつものExcel加工）
            # ---------------------------------------------------------
            if client_option == "バリューブックス用 (Excel出力)":
                
                # 項目名の変更
                rename_map = {
                    'issue_id': 'issue_id',
                    'issue_name': '件名',
                    'deliver': '配信数',
                    'sent_date': '日付',
                    'send_purpose': '配信対象',
                    'open_unique': '開封',
                    'open_rate': '開封率',
                    'click_total': 'CT'
                }
                df = df.rename(columns=rename_map)

                # 並び替え
                if 'issue_id' in df.columns:
                    df = df.sort_values('issue_id', ascending=False)

                # 文字の置き換え
                if '配信対象' in df.columns:
                    df['配信対象'] = df['配信対象'].replace('Advertising (external)', 'PC')

                # 日付・曜日の処理
                if '日付' in df.columns:
                    df['日付'] = pd.to_datetime(df['日付'], errors='coerce')
                    day_map = {0: '月', 1: '火', 2: '水', 3: '木', 4: '金', 5: '土', 6: '日'}
                    df['曜日'] = df['日付'].dt.dayofweek.map(day_map)
                    df['日付'] = df['日付'].dt.strftime('%Y/%m/%d')
                
                # 開封率の％化
                if '開封率' in df.columns:
                    df['開封率'] = pd.to_numeric(df['開封率'], errors='coerce')
                    df['開封率'] = df['開封率'].apply(lambda x: f"{x:.1f}%" if pd.notnull(x) else "")

                # 列整理
                df['対象'] = ""
                target_order = ['件名', '対象', '配信対象', '日付', '曜日', '配信数', '開封', '開封率', 'CT']
                final_cols = [c for c in target_order if c in df.columns]
                df_final = df[final_cols]

                # 結果表示とExcelダウンロード
                st.success(f"✅ バリューブックス用データ作成完了！ ({len(df_final)}件)")
                st.dataframe(df_final)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='Sheet1')
                
                st.download_button(
                    label="📥 Excelファイルをダウンロード",
                    data=output.getvalue(),
                    file_name='valuebooks_data.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                )

            # ---------------------------------------------------------
            # パターンB：リンクシェア用（テキスト出力）
            # ---------------------------------------------------------
            elif client_option == "リンクシェア用 (テキスト出力)":
                
                # 並び替え（新しい順）
                if 'issue_id' in df.columns:
                    df = df.sort_values('issue_id', ascending=False)
                
                # 日付型変換
                if 'sent_date' in df.columns:
                    df['sent_date'] = pd.to_datetime(df['sent_date'], errors='coerce')

                # テキスト生成処理
                output_text_list = []
                
                for index, row in df.iterrows():
                    # データの取得
                    issue_id = row.get('issue_id', '')
                    deliver = row.get('deliver', 0)
                    click = row.get('click_total', 0) # CSV上の列名は click_total
                    date_val = row.get('sent_date', pd.NaT)

                    # フォーマット作成
                    date_str = f"{date_val.year}/{date_val.month}/{date_val.day}" if pd.notnull(date_val) else "日付不明"
                    deliver_str = f"{int(deliver):,}" if pd.notnull(deliver) else "0"
                    click_str = f"{int(click):,}" if pd.notnull(click) else "0"

                    # テキストブロックの作成（ここを修正しました）
                    text_block = (
                        f"{date_str}配信\n\n"
                        f"【issueID】{issue_id}\n"  # ここを \n 1つにしました
                        f"【Deliver】{deliver_str}\n"
                        f"【Click】{click_str}\n"
                        "--------------------------------------------------"
                    )
                    output_text_list.append(text_block)

                # 全データを結合
                final_text = "\n\n".join(output_text_list)

                # 結果表示
                st.success(f"✅ リンクシェア用テキスト作成完了！ ({len(df)}件)")
                
                # コピー用のテキストエリア
                st.text_area("以下のテキストをコピーして使用してください", final_text, height=400)
                
                # テキストファイルとしてダウンロード
                st.download_button(
                    label="📥 テキストファイル(.txt)をダウンロード",
                    data=final_text,
                    file_name='linkshare_data.txt',
                    mime='text/plain',
                )

        except Exception as e:
            st.error(f"加工中にエラーが発生しました: {e}")
