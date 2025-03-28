import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

def app1():
    st.title('窓口売上インポート')

    st.markdown("""
        ### 使い方
        1. 全部門の日計月計をアップロード(一括)
        2. 該当月を選択
        3.「処理開始」にチェック
        4. 発生日を選択(該当月の末日を選択)
        5.「OK」にチェック
        6. ダウンロードボタンをクリック
        """)

    # ファイルのアップロード
    uploaded_files = st.file_uploader("日計報告を全てアップロードしてください", accept_multiple_files=True, type=['xlsm'])
    
    # 現在の日付から1ヶ月前の月を取得
    one_month_ago = (pd.to_datetime("today") - pd.DateOffset(months=1)).month
    
    # 月のリスト
    months = ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"]

    # ユーザーから月を選択してもらう
    selected_month = st.selectbox("読み込むシート名を選択してください:(初期値＝作業日の前月)", months, index=one_month_ago - 1)

    # 発生日の設定 (末日を選択)
    if selected_month:
        # 発生日は選択した月の末日を設定
        year = pd.to_datetime("today").year
        last_day = pd.to_datetime(f"{year}-{selected_month[:len(selected_month)-1]}-01") + pd.offsets.MonthEnd(0)
        selected_date = last_day.date()
        st.write(f"発生日（選択された月の末日）: {selected_date}")

    if st.checkbox('処理開始'):
        if uploaded_files and selected_month:
            dataframes = {}

            for uploaded_file in uploaded_files:
                df = pd.read_excel(uploaded_file, sheet_name=selected_month, header=7)
                df.columns = [col.replace('\n', '') for col in df.columns]
                dataframes[uploaded_file.name] = df

            columns = ["つくば", "羽生", "王子", "三鷹", "仙台", "川口", "船橋", "南森町", "高田馬場", "横浜関内", "福岡天神", "大宮"]
            index = ["自費", "社保", "国保", "販売品", "過不足金", "保険返金", "その他/保険証忘れ","振込入金", "自費返金", "JACCS入金", "口座振替"]
            
            final_df = pd.DataFrame(index=index, columns=columns)
            for filename, data in dataframes.items():
                for column in columns:
                    if column in filename:
                        for idx in index:
                            if idx == '口座振替':
                                key_column = '過不足金'
                                if key_column in data.columns and len(data[key_column]) > 35:
                                    value = data[key_column].iloc[35]
                                else:
                                    value = None
                            else:
                                clean_index = idx.replace(' ', '')
                                if clean_index in data.columns:
                                    value = data[clean_index].iloc[31] if len(data[clean_index]) > 31 else None
                            
                            final_df.at[idx, column] = value

            # ここでデータフレームを表示
            st.dataframe(final_df)

            st.success('修正がある場合は元データを修正するか、直接修正してください')

            if st.checkbox("仕訳作成"):
                output_columns = ['収支区分', '発生日', '取引先', '税区分', '勘定科目', '品目', '部門', '金額']
                output_df = pd.DataFrame(columns=output_columns)

                for col in final_df.columns:
                    for idx in final_df.index:
                        value = final_df.at[idx, col]
                        if pd.isna(value):
                            value = None
                        else:
                            try:
                                value = float(value)
                            except ValueError:
                                continue

                        if value is not None and value != 0:
                            new_row = pd.DataFrame({
                                '品目': [idx],
                                '部門': [col],
                                '金額': [value]
                            }, columns=output_columns)
                            output_df = pd.concat([output_df, new_row], ignore_index=True)

                def assign_tax_and_account(item):
                    if item in ['国保', '社保', '過不足金', '保険返金', 'その他/保険証忘れ']:
                        return '非課売上', '保険診療収入（窓口）'
                    elif item in ['自費', '振込入金', '自費返金', 'JACCS入金', '口座振替']:
                        return '課税売上10%', '自費収入'
                    elif item == '販売品':
                        return '課税売上10%', '雑収入'
                    else:
                        return None, None

                for index, row in output_df.iterrows():
                    tax_category, account = assign_tax_and_account(row['品目'])
                    output_df.at[index, '税区分'] = tax_category
                    output_df.at[index, '勘定科目'] = account
                    
                output_df['収支区分'] = '収入'
                output_df['発生日'] = selected_date

                # 品目の整形
                output_df['品目'] = output_df['品目'].replace({
                    'その他/保険証忘れ': 'その他',
                    '口座振替': 'JACCS',
                    'JACCS入金': 'JACCS'
                })

                st.write(output_df)

                def convert_df_to_excel(df):
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df.to_excel(writer, index=False, sheet_name='Sheet1')
                    processed_data = output.getvalue()
                    return processed_data

                st.success('処理が完了しました')
                
                excel_data = convert_df_to_excel(output_df)
                st.download_button(
                    label="ここからダウンロードしてください",
                    data=excel_data,
                    file_name='import_data(窓口).xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )

