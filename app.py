import streamlit as st
import pandas as pd
import twstock
from datetime import datetime, timedelta
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.properties import WorksheetProperties, PageSetupProperties
import math
import calendar

st.set_page_config(page_title="股價報表產生器", layout="centered")

# ===== 股票選單 =====
from twstock import codes
stock_options = [
    f"{code} {codes[code].name}"
    for code in sorted(codes.keys())
    if hasattr(codes[code], "name") and codes[code].name and 4 <= len(code) <= 6
]

st.title("台股區間報表產生器")

selected = st.selectbox("選擇股票代碼", stock_options)
stock_id = selected.split()[0]
stock_name = selected.split()[1]
interval = st.radio("選擇統計區間", ["日", "週", "月"], horizontal=True)

min_day = datetime(2015, 1, 1)
max_day = datetime(2035, 12, 31)

col1, col2 = st.columns(2)
with col1:
    start_date = st.date_input("開始日期", value=datetime(2024, 1, 1), min_value=min_day, max_value=max_day)
with col2:
    end_date = st.date_input("結束日期", value=datetime.today(), min_value=min_day, max_value=max_day)

# 按下按鈕才執行
if st.button("產生報表"):
    start_date = datetime.combine(start_date, datetime.min.time())
    end_date = datetime.combine(end_date, datetime.min.time())

    stock = twstock.Stock(stock_id)
    today = datetime.today()
    # 區間對齊
    if interval == "週":
        start_date = start_date - timedelta(days=start_date.weekday())
        end_date = end_date + timedelta(days=6 - end_date.weekday())
    elif interval == "月":
        start_date = start_date.replace(day=1)
        last_day = calendar.monthrange(end_date.year, end_date.month)[1]
        end_date = end_date.replace(day=last_day)

    if end_date > today:
        end_date = today

    # 取得比對基準
    if interval == "週":
        prev_start = start_date - timedelta(days=7)
        prev_end = start_date - timedelta(days=1)
        raw_prev = stock.fetch_from(prev_start.year, prev_start.month)
        prev_filtered = [d for d in raw_prev if prev_start <= d.date <= prev_end]
        if prev_filtered:
            prev_high = max(d.high for d in prev_filtered)
            prev_low = min(d.low for d in prev_filtered)
            prev_volume = sum(d.capacity for d in prev_filtered)
            prev_diff = prev_high - prev_low
        else:
            prev_high, prev_low, prev_volume, prev_diff = None, None, None, None
    elif interval == "月":
        prev_month_end = start_date - timedelta(days=1)
        prev_month_start = prev_month_end.replace(day=1)
        raw_prev = stock.fetch_from(prev_month_start.year, prev_month_start.month)
        prev_filtered = [d for d in raw_prev if prev_month_start <= d.date <= prev_month_end]
        if prev_filtered:
            prev_high = max(d.high for d in prev_filtered)
            prev_low = min(d.low for d in prev_filtered)
            prev_volume = sum(d.capacity for d in prev_filtered)
            prev_diff = prev_high - prev_low
        else:
            prev_high, prev_low, prev_volume, prev_diff = None, None, None, None
    else:
        extra_date = start_date - timedelta(days=14)
        raw_prev = stock.fetch_from(extra_date.year, extra_date.month)
        prev_filtered = [d for d in raw_prev if d.date < start_date]
        if prev_filtered:
            d = max(prev_filtered, key=lambda x: x.date)
            prev_high = d.high
            prev_low = d.low
            prev_volume = d.capacity
            prev_diff = prev_high - prev_low
        else:
            prev_high, prev_low, prev_volume, prev_diff = None, None, None, None

    # 取得主資料
    raw_data = stock.fetch_from(start_date.year, start_date.month)
    filtered = [d for d in raw_data if start_date <= d.date <= end_date]

    if not filtered:
        st.error("查無資料，請檢查股票代碼與時間範圍。")
        st.stop()

    # 轉 DataFrame
    df = pd.DataFrame([{
        '日期': d.date,
        '最高價': d.high,
        '最低價': d.low,
        '成交量': d.capacity
    } for d in filtered])

    # 統計區間
    if interval == "日":
        agg_df = df.copy()
    elif interval == "週":
        df["iso_year"] = df["日期"].apply(lambda x: x.isocalendar()[0])
        df["iso_week"] = df["日期"].apply(lambda x: x.isocalendar()[1])
        grouped = df.groupby(["iso_year", "iso_week"])
        agg_df = grouped.agg({
            "日期": "last",
            "最高價": "max",
            "最低價": "min",
            "成交量": "sum"
        }).reset_index(drop=True)
    else:  # 月
        df["month"] = df["日期"].apply(lambda x: x.month)
        df["year"] = df["日期"].apply(lambda x: x.year)
        grouped = df.groupby(["year", "month"])
        agg_df = grouped.agg({
            "日期": "last",
            "最高價": "max",
            "最低價": "min",
            "成交量": "sum"
        }).reset_index(drop=True)

    # 顏色計算
    agg_df["差色"] = ""
    agg_df["高色"] = ""
    agg_df["低色"] = ""
    agg_df["成交符"] = ""
    agg_df["符色"] = ""

    for i, row in agg_df.iterrows():
        if i == 0 and prev_high is not None:
            cmp_high = prev_high
            cmp_low = prev_low
            cmp_volume = prev_volume
            cmp_diff = prev_diff
        else:
            prev_row = agg_df.iloc[i-1]
            cmp_high = prev_row["最高價"]
            cmp_low = prev_row["最低價"]
            cmp_volume = prev_row["成交量"]
            cmp_diff = prev_row["最高價"] - prev_row["最低價"]

        diff = row["最高價"] - row["最低價"]
        agg_df.loc[i, "高色"] = "FFCC3333" if row["最高價"] >= cmp_high else "FF3366CC"
        agg_df.loc[i, "低色"] = "FFCC3333" if row["最低價"] >= cmp_low else "FF3366CC"
        agg_df.loc[i, "成交符"] = "■"
        agg_df.loc[i, "符色"] = "FFCC3333" if row["成交量"] >= cmp_volume else "FF3366CC"
        agg_df.loc[i, "差色"] = "FFCC3333" if diff >= cmp_diff else "FF3366CC"

    # 分頁
    chunk_size = 43
    chunks = []
    for i in range(0, len(agg_df), chunk_size):
        chunks.append(agg_df.iloc[i:i+chunk_size].reset_index(drop=True))

    blocks_per_sheet = 3
    total_blocks = len(chunks)
    total_pages = math.ceil(total_blocks / blocks_per_sheet)

    wb = Workbook()
    ws = wb.active
    ws.title = f"{interval}報表"

    title = f"{stock_name} {start_date.strftime('%Y-%m-%d')}～{end_date.strftime('%Y-%m-%d')}（{interval}）"
    sheet_count = 0
    headers = ["日期", "", "高", "低", "漲幅", "量", ""] * 3
    starts = [1, 8, 15]
    prev_month = None
    prev_year = None

    for block, data in enumerate(chunks):
        if block % blocks_per_sheet == 0:
            sheet_count = block // blocks_per_sheet
            ws = wb.active if block == 0 else wb.create_sheet(title=f"{interval}報表{sheet_count+1}")
            ws.insert_rows(1)
            ws.insert_rows(2)
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=20)
            page_info = f"（第 {sheet_count+1}/{total_pages} 頁）" if total_pages > 1 else ""
            title_cell = ws.cell(row=1, column=1, value=f"{title}{page_info}")
            title_cell.font = Font(bold=True, size=14)
            title_cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[1].height = 30
            for i, h in enumerate(headers):
                cell = ws.cell(row=2, column=i + 1, value=h)
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center")
            for i in starts:
                ws.merge_cells(start_row=2, start_column=i, end_row=2, end_column=i+1)
            sheet_count += 1

        col = starts[block % 3]
        row_index = 3

        for i, row in data.iterrows():
            dt = row["日期"]
            if interval == "日":
                if (prev_month != None) and (dt.month != prev_month):
                    day_str = f"{dt.month}/{dt.day}"
                else:
                    day_str = f"{dt.day}"
                week_str = ["一", "二", "三", "四", "五", "六", "日"][dt.weekday()]
                prev_month = dt.month
                ws.cell(row=row_index, column=col, value=day_str).alignment = Alignment(horizontal="center")
                ws.cell(row=row_index, column=col+1, value=week_str).alignment = Alignment(horizontal="center")
            elif interval == "週":
                if (prev_year != None) and (dt.year != prev_year):
                    day_str = f"{dt.year}/{dt.month}/{dt.day}"
                else:
                    day_str = f"{dt.month}/{dt.day}"
                prev_year = dt.year
                ws.merge_cells(start_row=row_index, start_column=col, end_row=row_index, end_column=col+1)
                ws.cell(row=row_index, column=col, value=day_str).alignment = Alignment(horizontal="center")
            else:
                if (prev_year != None) and (dt.year != prev_year):
                    day_str = f"{dt.year}/{dt.month}"
                else:
                    day_str = f"{dt.month}"
                prev_year = dt.year
                ws.merge_cells(start_row=row_index, start_column=col, end_row=row_index, end_column=col+1)
                ws.cell(row=row_index, column=col, value=day_str).alignment = Alignment(horizontal="center")
            h = ws.cell(row=row_index, column=col+2, value=row["最高價"])
            h.font = Font(color=row["高色"])
            h.alignment = Alignment(horizontal="center")

            l = ws.cell(row=row_index, column=col+3, value=row["最低價"])
            l.font = Font(color=row["低色"])
            l.alignment = Alignment(horizontal="center")

            d_value = round(row["最高價"] - row["最低價"], 2)
            d = ws.cell(row=row_index, column=col+4, value=d_value)
            d.font = Font(color=row["差色"])
            d.alignment = Alignment(horizontal="center")

            v = ws.cell(row=row_index, column=col+5, value=row["成交符"])
            v.font = Font(color=row["符色"], size=10)  # 方形空心、可調大小
            v.alignment = Alignment(horizontal="center")
            row_index += 1

        # 欄寬自動
        for col_cells in ws.iter_cols(min_row=3, max_col=ws.max_column, max_row=ws.max_row):
            col_letter = get_column_letter(col_cells[0].column)
            max_len = max(len(str(c.value)) if c.value else 0 for c in col_cells)
            ws.column_dimensions[col_letter].width = max(2, min(max_len + 2, 16))
        for i in range(8, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(i)].width = ws.column_dimensions[get_column_letter(i-7)].width

    ws.freeze_panes = "A3"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.scale = None
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = 11
    ws.sheet_properties = WorksheetProperties(pageSetUpPr=PageSetupProperties(fitToPage=True))
    ws.page_setup.horizontalCentered = True
    ws.page_setup.verticalCentered = True
    ws.page_margins = PageMargins(
        left=0.3, right=0.3,
        top=0.2, bottom=0.2,
        header=0.2, footer=0.2
    )

    buffer = BytesIO()
    wb.save(buffer)
    st.success("✅ 報表產出成功")
    st.download_button("下載 Excel", data=buffer.getvalue(), file_name=f"{title}.xlsx")
