import streamlit as st
import pandas as pd
import twstock
from datetime import datetime, timedelta, time
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.properties import WorksheetProperties, PageSetupProperties

st.title("蘇大哥專用工具（Excel）")

from twstock import codes

stock_options = [
    f"{code} {codes[code].name}"
    for code in sorted(codes.keys())
    if hasattr(codes[code], "name") and codes[code].name and 4 <= len(code) <= 6
]

selected = st.selectbox("選擇股票代碼", stock_options)
stock_id = selected.split()[0]

min_day = datetime(2015, 1, 1)
max_day = datetime(2035, 12, 31)

# 📅 日期區間
start_date = datetime.combine(
    st.date_input("起始日期", datetime.today() - timedelta(days=90), min_value=min_day, max_value=max_day),
    time.min
)
end_date = datetime.combine(
    st.date_input("結束日期", datetime.today(), min_value=min_day, max_value=max_day),
    time.max
)

# 📊 統計區間選擇
stat_mode = st.radio("統計區間", ["日", "週", "月"], index=0)

if start_date >= end_date:
    st.warning("⚠️ 結束日期必須晚於起始日期")
    st.stop()

if st.button("產出報表"):
    stock = twstock.Stock(stock_id)
    raw_data = stock.fetch_from(start_date.year, start_date.month)
    filtered = [d for d in raw_data if start_date <= d.date <= end_date]
    if not filtered:
        st.error("查無資料，請檢查股票代碼與時間範圍。")
        st.stop()

    # 用 datetime 保留
    df = pd.DataFrame([{
        '日期': d.date,
        '最高價': d.high,
        '最低價': d.low,
        '收盤價': d.close,
        '成交量': d.capacity
    } for d in filtered])

    # 週、月分組前先排序
    df = df.sort_values('日期').reset_index(drop=True)

    # -------- 根據統計方式彙總 --------
    if stat_mode == "日":
        report_df = df.copy()
        headers = ["日", "週", "高", "低", "漲幅", "量"]
        # 計算週幾與漲幅
        report_df["日"] = report_df["日期"].dt.day
        report_df["週"] = report_df["日期"].dt.dayofweek.map(lambda x: "一二三四五六日"[x])
        report_df["漲幅"] = report_df["最高價"] - report_df["最低價"]
        report_df["量"] = report_df["成交量"]
        # 排序欄位
        report_df = report_df[["日", "週", "最高價", "最低價", "漲幅", "量"]]
    elif stat_mode == "週":
        df['年'] = df['日期'].apply(lambda x: x.isocalendar()[0])
        df['週次'] = df['日期'].apply(lambda x: x.isocalendar()[1])
        report_df = df.groupby(['年', '週次']).agg({
            '日期': ['min', 'max'],
            '最高價': 'max',
            '最低價': 'min',
            '收盤價': 'last',
            '成交量': 'sum'
        }).reset_index()
        report_df.columns = ['年', '週次', '起始日', '結束日', '最高價', '最低價', '收盤價', '總成交量']
        headers = ["年", "週", "起始日", "結束日", "高", "低", "收盤", "量"]
    elif stat_mode == "月":
        df['年'] = df['日期'].dt.year
        df['月'] = df['日期'].dt.month
        report_df = df.groupby(['年', '月']).agg({
            '日期': ['min', 'max'],
            '最高價': 'max',
            '最低價': 'min',
            '收盤價': 'last',
            '成交量': 'sum'
        }).reset_index()
        report_df.columns = ['年', '月', '起始日', '結束日', '最高價', '最低價', '收盤價', '總成交量']
        headers = ["年", "月", "起始日", "結束日", "高", "低", "收盤", "量"]

    # -------- Excel 報表產出 --------
    wb = Workbook()
    ws = wb.active
    ws.title = "股價報表"

    # 標題
    ws.insert_rows(1)
    ws.insert_rows(2)
    title = f"{selected} {start_date.strftime('%Y-%m-%d')}～{end_date.strftime('%Y-%m-%d')}（{stat_mode}）"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    title_cell = ws.cell(row=1, column=1, value=title)
    title_cell.font = Font(bold=True, size=14)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    # 標題列
    for i, h in enumerate(headers):
        cell = ws.cell(row=2, column=i + 1, value=h)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    # 資料內容
    for r, row in enumerate(report_df.itertuples(index=False), start=3):
        for c, v in enumerate(row, start=1):
            ws.cell(row=r, column=c, value=v).alignment = Alignment(horizontal="center")

    # 自動欄寬
    for col_cells in ws.iter_cols(min_row=2, max_col=ws.max_column, max_row=ws.max_row):
        col_letter = get_column_letter(col_cells[0].column)
        max_len = max(len(str(c.value)) if c.value else 0 for c in col_cells)
        ws.column_dimensions[col_letter].width = max(2, min(max_len + 2, 16))

    # 列印設定
    ws.freeze_panes = "A3"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.scale = None
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = 9
    ws.sheet_properties = WorksheetProperties(
        pageSetUpPr=PageSetupProperties(fitToPage=True)
    )
    ws.page_setup.horizontalCentered = True
    ws.page_setup.verticalCentered = True
    ws.page_margins = PageMargins(
        left=0.3, right=0.3,
        top=0.5, bottom=0.5,
        header=0.2, footer=0.2
    )

    # 下載
    buffer = BytesIO()
    wb.save(buffer)
    st.success("✅ 報表產出成功")
    st.download_button("下載 Excel", data=buffer.getvalue(), file_name=f"{title}.xlsx")
