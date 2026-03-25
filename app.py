import streamlit as st
import pandas as pd
import io

# ページ設定
st.set_page_config(page_title="効果測定レポート加工ツール（最終完成版）", layout="wide")
st.title("📊 効果測定レポート加工ツール")
st.markdown("CSVをアップロードすると、指定の順序・変換ルールで加工し、**Excelファイル**として出力します。")

# ファイルアップロード（複数対応）
uploaded_files = st.file_uploader(
    "CSVファイルをドラッグ＆ドロップしてください（複数可）", 
    type=['csv'], 
    accept_multiple_files=True
)

if uploaded_files:
    df_list = []
    
    for file in uploaded_files:
        try:
            # CSV読み込み
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
            # (A) データの結合
            df = pd.concat(df_list, ignore_index=True)

            # ---------------------------------------------------------
            # (B) データ加工処理
            # ---------------------------------------------------------
            
            # 1. 項目名の変更マップ
            # ここを変更しました：open_total → open_unique を「開封」にする
            rename_map = {
                'issue_id': 'issue_id',
                'issue_name': '件名',
                'deliver': '配信数',
                'sent_date': '日付',
                'send_purpose': '配信対象',
                'open_unique': '開封',   # ← ここを open_unique に変更しました
                'open_rate': '開封率',
                'click_total': 'CT'
            }
            
            # 列名変更を実行
            df = df.rename(columns=rename_map)

            # 2. 並び替え（issue_id の降順）
            if 'issue_id' in df.columns:
                df = df.sort_values('issue_id', ascending=False)

            # 3. 文字の置き換え（Advertising (external) → PC）
            if '配信対象' in df.columns:
                df['配信対象'] = df['配信対象'].replace('Advertising (external)', 'PC')

            # 4. 日付・曜日の処理
            if '日付' in df.columns:
                df['日付'] = pd.to_datetime(df['日付'], errors='coerce')
                day_map = {0: '月', 1: '火', 2: '水', 3: '木', 4: '金', 5: '土', 6: '日'}
                df['曜日'] = df['日付'].dt.dayofweek.map(day_map)
                df['日付'] = df['日付'].dt.strftime('%Y/%m/%d')
            
            # 5. 「開封率」を％表記に変換
            if '開封率' in df.columns:
                df['開封率'] = pd.to_numeric(df['開封率'], errors='coerce')
                # 数字はそのままで、小数第1位まで丸めて％をつける
                df['開封率'] = df['開封率'].apply(lambda x: f"{x:.1f}%" if pd.notnull(x) else "")

            # 6. 「対象」列追加
            df['対象'] = ""

            # 7. 列の並び替えと抽出
            target_order = [
                '件名',
                '対象',
                '配信対象',
                '日付',
                '曜日',
                '配信数',
                '開封',
                '開封率',
                'CT'
            ]
            
            # 存在する列だけ残す
            final_cols = [c for c in target_order if c in df.columns]
            df_final = df[final_cols]

            # ---------------------------------------------------------
            # (C) Excelファイル生成とダウンロード
            # ---------------------------------------------------------
            
            st.success(f"✅ 加工完了！ {len(df_final)} 件のデータを処理しました。")
            st.dataframe(df_final)

            # Excelファイルをメモリ上に作成
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Sheet1')
            
            # ダウンロードボタン
            st.download_button(
                label="📥 Excelファイルをダウンロード",
                data=output.getvalue(),
                file_name='processed_data.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )

        except Exception as e:
            st.error(f"加工中にエラーが発生しました: {e}")
