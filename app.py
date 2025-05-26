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

start_date = datetime.combine(
    st.date_input("起始日期", datetime.today() - timedelta(days=90), min_value=min_day, max_value=max_day),
    time.min
)
end_date = datetime.combine(
    st.date_input("結束日期", datetime.today(), min_value=min_day, max_value=max_day),
    time.max
)

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

    df = pd.DataFrame([{
        '日期': d.date.strftime("%Y-%m-%d"),
        '最高價': d.high,
        '最低價': d.low,
        '成交量': d.capacity
    } for d in filtered])

    # 補一筆資料作為比對用
    extra_date = start_date - timedelta(days=5)
    extra_data = stock.fetch_from(extra_date.year, extra_date.month)
    extra_point = next((d for d in reversed(extra_data) if d.date < start_date), None)
    if extra_point:
        df = pd.concat([pd.DataFrame([{
            '日期': extra_point.date.strftime("%Y-%m-%d"),
            '最高價': extra_point.high,
            '最低價': extra_point.low,
            '成交量': extra_point.capacity
        }]), df], ignore_index=True)

    # 計算欄位
    df["差色"] = ""
    for i in range(len(df)):
        if i == 0:
            df.loc[i, ["成交符", "符色", "差色"]] = "-", "000000", "000000"
        else:
            prev = df.iloc[i - 1]
            now = df.iloc[i]
            diff = now["最高價"] - now["最低價"]
            prev_diff = prev["最高價"] - prev["最低價"]
            df.loc[i, "高色"] = "FF0000" if now["最高價"] >= prev["最高價"] else "0000FF"
            df.loc[i, "低色"] = "FF0000" if now["最低價"] >= prev["最低價"] else "0000FF"
            df.loc[i, "成交符"] = "🔴" if now["成交量"] >= prev["成交量"] else "🔵"
            df.loc[i, "符色"] = "FF0000" if df.loc[i, "成交符"] == "🔴" else "0000FF"
            df.loc[i, "差色"] = "FF0000" if diff >= prev_diff else "0000FF"

    df = df.iloc[1:].reset_index(drop=True)

    # 分成三區塊
    base = len(df) // 3
    remainder = len(df) % 3
    sizes = [base + (1 if i < remainder else 0) for i in range(3)]
    chunks = []
    s = 0
    for size in sizes:
        chunks.append(df.iloc[s:s+size].reset_index(drop=True))
        s += size

    wb = Workbook()
    ws = wb.active
    ws.title = "股價報表"

    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    bottom_border = Border(bottom=Side(style="thin"))

    # 插入標題兩列
    ws.insert_rows(1)
    ws.insert_rows(2)
    title = f"{selected} {start_date.strftime('%Y-%m-%d')}～{end_date.strftime('%Y-%m-%d')}（日）"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=20)  # 固定到第18欄
    title_cell = ws.cell(row=1, column=1, value=title)
    title_cell.font = Font(bold=True, size=14)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    # 設定標題列
    headers = ["日", "週", "高", "低", "漲幅", "量", ""] * 3
    for i, h in enumerate(headers):
        cell = ws.cell(row=2, column=i + 1, value=h)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    starts = [1, 8, 15]
    for block, data in enumerate(chunks):
        col = starts[block]
        row_index = 3
        prev_month = None
        for i, row in data.iterrows():
            full_date = datetime.strptime(row["日期"], "%Y-%m-%d")
            current_month = full_date.month
            if i == 0 or current_month != prev_month:
                day_str = f"{full_date.month}/{full_date.day}"
            else:
                day_str = str(full_date.day)
            prev_month = current_month
        
            weekday_str = ["一", "二", "三", "四", "五", "六", "日"][full_date.weekday()]
            ws.cell(row=row_index, column=col, value=day_str).alignment = Alignment(horizontal="center")
            ws.cell(row=row_index, column=col+1, value=weekday_str).alignment = Alignment(horizontal="center")
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
            v.font = Font(color=row["符色"])
            v.alignment = Alignment(horizontal="center")
            # 每週結尾加底線
            current_week = full_date.isocalendar()[1]
            next_week = current_week
            # 只要不是區塊最後一筆，就取下一筆的 week
            if i < len(data) - 1:
                next_date = datetime.strptime(data.iloc[i + 1]["日期"], "%Y-%m-%d")
                next_week = next_date.isocalendar()[1]
            
            # 只有週有變才加底線（就算是區塊結束，也不自動加底線）
            if next_week != current_week:
                for offset in range(6):
                    ws.cell(row=row_index, column=col + offset).border = bottom_border
            row_index += 1

    # 欄寬自動
    for col_cells in ws.iter_cols(min_row=3, max_col=ws.max_column, max_row=ws.max_row):
        col_letter = get_column_letter(col_cells[0].column)
        max_len = max(len(str(c.value)) if c.value else 0 for c in col_cells)
        ws.column_dimensions[col_letter].width = max(2, min(max_len + 2, 16))

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

    buffer = BytesIO()
    wb.save(buffer)
    st.success("✅ 報表產出成功")
    st.download_button("下載 Excel", data=buffer.getvalue(), file_name=f"{title}.xlsx")
