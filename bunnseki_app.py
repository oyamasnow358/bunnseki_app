import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import RadarChart, Reference
import datetime

# 項目と発達段階の設定
categories = ["言語理解", "表出言語", "視覚記憶", "聴覚記憶", "読字", "書字", "数", "運動"]
age_ranges = ["0～3ヶ月", "3～6ヶ月", "6～9ヶ月", "9～12ヶ月", "12～18ヶ月", "18～24ヶ月", "2～3歳", "3～4歳", "4～5歳", "5～6歳", "6～7歳", "7歳以降"]

# 入力フォーム
st.title("発達段階の入力フォーム")
user_data = {}

for category in categories:
    user_data[category] = st.selectbox(f"{category}の発達段階を選択してください:", age_ranges)

# データの保存
if st.button("データを保存"):
    df = pd.DataFrame(user_data, index=[0])  # 入力データをDataFrameに変換
    excel_file_name = "development_data.xlsx"
    df.to_excel(excel_file_name, index=False)  # Excelファイルに保存
    st.success(f"{excel_file_name} にデータが保存されました！")

    # レーダーチャートの作成
    wb = Workbook()
    ws = wb.active
    ws.title = "発達データ"

    # ヘッダーの設定
    header = ["項目", "スコア"]
    ws.append(header)

    # 入力データを数値化してExcelに書き込み
    scores = {
        "0～3ヶ月": 1,
        "3～6ヶ月": 2,
        "6～9ヶ月": 3,
        "9～12ヶ月": 4,
        "12～18ヶ月": 5,
        "18～24ヶ月": 6,
        "2～3歳": 7,
        "3～4歳": 8,
        "4～5歳": 9,
        "5～6歳": 10,
        "6～7歳": 11,
        "7歳以降": 12,
    }

    for category in categories:
        row = [category, scores.get(user_data[category], 0)]
        ws.append(row)

    # 最大値を固定するためのダミーデータを追加
    ws.append(["最大値", 12])  # 最大値の行
    ws.append(["最小値", 0])   # 最小値の行

    # レーダーチャートの設定
    chart = RadarChart()
    chart.title = "発達段階の六角形グラフ"
    chart.style = 26

    # データ範囲を指定
    data_ref = Reference(ws, min_col=2, min_row=2, max_row=len(categories) + 3)  # ダミーデータを含める
    categories_ref = Reference(ws, min_col=1, min_row=2, max_row=len(categories) + 1)  # 本来のカテゴリだけ

    chart.add_data(data_ref, titles_from_data=False)
    chart.set_categories(categories_ref)

    # ワークシートにチャートを追加
    ws.add_chart(chart, "J2")

    # ファイル名にタイムスタンプを追加して保存
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    chart_file_name = f"development_chart_{timestamp}.xlsx"
    wb.save(chart_file_name)
    st.success(f"{chart_file_name} が作成されました！")
