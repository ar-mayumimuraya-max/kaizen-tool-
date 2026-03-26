import streamlit as st
import pandas as pd
import io

# ページ設定
st.set_page_config(page_title="業務データ加工ツール（マルチクライアント）", layout="wide")
st.title("📊 業務データ加工ツール")

# ▼▼▼ 1. パターン選択ラジオボタン ▼▼▼
client_option = st.radio(
    "作成するデータのパターンを選択してください",
    (
        "バリューブックス用 (Excel出力)", 
        "リンクシェア用 (テキスト出力)",
        "【メール部用】開封率＆メルマガ費レポート (Excel出力)"
    ),
    horizontal=False
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
            # パターンA：バリューブックス用
            # ---------------------------------------------------------
            if client_option == "バリューブックス用 (Excel出力)":
                rename_map = {
                    'issue_id': 'issue_id', 'issue_name': '件名', 'deliver': '配信数',
                    'sent_date': '日付', 'send_purpose': '配信対象', 'open_unique': '開封',
                    'open_rate': '開封率', 'click_total': 'CT'
                }
                df = df.rename(columns=rename_map)
                
                if 'issue_id' in df.columns:
                    df = df.sort_values('issue_id', ascending=True)

                if '配信対象' in df.columns:
                    df['配信対象'] = df['配信対象'].replace('Advertising (external)', 'PC')

                if '日付' in df.columns:
                    df['日付'] = pd.to_datetime(df['日付'], errors='coerce')
                    day_map = {0: '月', 1: '火', 2: '水', 3: '木', 4: '金', 5: '土', 6: '日'}
                    df['曜日'] = df['日付'].dt.dayofweek.map(day_map)
                    df['日付'] = df['日付'].dt.strftime('%Y/%m/%d')
                
                if '開封率' in df.columns:
                    df['開封率'] = pd.to_numeric(df['開封率'], errors='coerce')
                    df['開封率'] = df['開封率'].apply(lambda x: f"{x:.1f}%" if pd.notnull(x) else "")

                df['対象'] = ""
                target_order = ['件名', '対象', '配信対象', '日付', '曜日', '配信数', '開封', '開封率', 'CT']
                final_cols = [c for c in target_order if c in df.columns]
                df_final = df[final_cols]

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
            # パターンB：リンクシェア用
            # ---------------------------------------------------------
            elif client_option == "リンクシェア用 (テキスト出力)":
                if 'issue_id' in df.columns:
                    df = df.sort_values('issue_id', ascending=True)
                
                if 'sent_date' in df.columns:
                    df['sent_date'] = pd.to_datetime(df['sent_date'], errors='coerce')

                output_text_list = []
                for index, row in df.iterrows():
                    issue_id = row.get('issue_id', '')
                    deliver = row.get('deliver', 0)
                    click = row.get('click_total', 0)
                    date_val = row.get('sent_date', pd.NaT)

                    date_str = f"{date_val.year}/{date_val.month}/{date_val.day}" if pd.notnull(date_val) else "日付不明"
                    deliver_str = f"{int(deliver):,}" if pd.notnull(deliver) else "0"
                    click_str = f"{int(click):,}" if pd.notnull(click) else "0"

                    text_block = (
                        f"{date_str}配信\n\n"
                        f"【issueID】{issue_id}\n"
                        f"【Deliver】{deliver_str}\n"
                        f"【Click】{click_str}\n"
                        "--------------------------------------------------"
                    )
                    output_text_list.append(text_block)

                final_text = "\n\n".join(output_text_list)
                st.success(f"✅ リンクシェア用テキスト作成完了！ ({len(df)}件)")
                st.text_area("以下のテキストをコピーして使用してください", final_text, height=400)
                st.download_button(
                    label="📥 テキストファイル(.txt)をダウンロード",
                    data=final_text,
                    file_name='linkshare_data.txt',
                    mime='text/plain',
                )

            # ---------------------------------------------------------
            # パターンC：【メール部用】開封率＆メルマガ費レポート
            # ---------------------------------------------------------
            elif client_option == "【メール部用】開封率＆メルマガ費レポート (Excel出力)":
                
                # 共通の抽出項目
                target_cols = ['issue_id', 'issue_name', 'send_purpose', 'deliver', 'open_total', 'open_unique', 'open_rate', 'click_total']
                
                # 古い順に並び替え
                if 'issue_id' in df.columns:
                    df = df.sort_values('issue_id', ascending=True)

                # ==========================================
                # シート1：開封率 (internalのみ)
                # ==========================================
                df_1 = df[df['send_purpose'] == 'Advertising (internal)'].copy()
                df_1 = df_1[target_cols]

                # スタイル関数（シート1）
                def style_sheet_1(x):
                    yellow = 'background-color: #FFFF00'
                    red = 'background-color: #FF9999'
                    df_style = pd.DataFrame('', index=x.index, columns=x.columns)
                    
                    # 指定列だけ色付け
                    for col in ['send_purpose', 'deliver', 'open_total']:
                        if col in df_style.columns:
                            df_style[col] = yellow
                    if 'open_rate' in df_style.columns:
                        df_style['open_rate'] = red
                    return df_style

                # ==========================================
                # シート2：AD費 (external かつ 「号外◆」を含む)
                # ==========================================
                # 1. Advertising (external) で絞る
                mask_external = df['send_purpose'] == 'Advertising (external)'
                # 2. 件名に「号外◆」が含まれるもので絞る（na=Falseは空欄を除外するため）
                mask_gogai = df['issue_name'].str.contains('号外◆', na=False)
                
                # 両方の条件を満たすデータを抽出
                df_2 = df[mask_external & mask_gogai].copy()
                
                # 必要な列だけにする
                df_2 = df_2[target_cols]

                # AD費の計算
                df_2['AD費'] = (df_2['deliver'] * 1.1).round(-5)
                
                # 列の並び順
                cols_order = ['issue_id', 'issue_name', 'send_purpose', 'deliver', 'AD費', 'open_total', 'open_unique', 'open_rate', 'click_total']
                df_2 = df_2[cols_order]

                # スタイル関数（シート2）
                def style_sheet_2(x):
                    yellow = 'background-color: #FFFF00'
                    red = 'background-color: #FF9999'
                    df_style = pd.DataFrame('', index=x.index, columns=x.columns)
                    
                    # 指定列だけ色付け
                    for col in ['send_purpose', 'open_total']:
                        if col in df_style.columns:
                            df_style[col] = yellow
                    for col in ['deliver', 'AD費']:
                        if col in df_style.columns:
                            df_style[col] = red
                    return df_style

                # ==========================================
                # Excel出力処理
                # ==========================================
                st.success(f"✅ メール部用レポート作成完了！ (Internal:{len(df_1)}件 / 号外AD:{len(df_2)}件)")
                
                st.write("▼ シート1：開封率（プレビュー）")
                st.dataframe(df_1.head())
                st.write("▼ シート2：AD費（プレビュー：号外◆のみ）")
                st.dataframe(df_2.head())

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    # シート1書き込み
                    df_1.style.apply(style_sheet_1, axis=None).to_excel(writer, index=False, sheet_name='開封率')
                    # シート2書き込み
                    df_2.style.apply(style_sheet_2, axis=None).to_excel(writer, index=False, sheet_name='AD費')
                
                st.download_button(
                    label="📥 レポート(Excel)をダウンロード",
                    data=output.getvalue(),
                    file_name='mail_report_gogai.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                )

        except Exception as e:
            st.error(f"加工中にエラーが発生しました: {e}")
