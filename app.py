import streamlit as st
import pandas as pd
import twstock
from datetime import datetime, timedelta, time
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import WorksheetProperties, PageSetupProperties

st.title("蘇大哥專用工具（Excel）")

# ✅ 使用 twstock.codes，代碼 ➜ 中文名稱
from twstock import codes

stock_options = [
    f"{code} {codes[code].name}"
    for code in sorted(codes.keys())
    if hasattr(codes[code], "name") and codes[code].name and 4 <= len(code) <= 6
]

selected = st.selectbox("選擇股票代碼", stock_options)
stock_id = selected.split()[0]



# 📅 日期選擇
start_date = datetime.combine(
    st.date_input("起始日期", datetime.today() - timedelta(days=90)),
    time.min
)
end_date = datetime.combine(
    st.date_input("結束日期", datetime.today()),
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

    # 加入紅藍標記與成交符
    df["高色"], df["低色"], df["成交符"], df["符色"] = "", "", "", ""
    for i in range(len(df)):
        if i == 0:
            df.loc[i, ["成交符", "符色"]] = "-", "000000"
        else:
            prev = df.iloc[i - 1]
            now = df.iloc[i]
            df.loc[i, "高色"] = "FF0000" if now["最高價"] >= prev["最高價"] else "0000FF"
            df.loc[i, "低色"] = "FF0000" if now["最低價"] >= prev["最低價"] else "0000FF"
            df.loc[i, "成交符"] = "🔴" if now["成交量"] >= prev["成交量"] else "🔵"
            df.loc[i, "符色"] = "FF0000" if df.loc[i, "成交符"] == "🔴" else "0000FF"

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

    # 產出 Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "股價報表"
    thin = Border(left=Side(style='thin'), right=Side(style='thin'),
                  top=Side(style='thin'), bottom=Side(style='thin'))

    ws.insert_rows(1)
    ws.insert_rows(2)
    title = f"{selected} {start_date.strftime('%Y-%m-%d')}～{end_date.strftime('%Y-%m-%d')}（日）"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=12)
    title_cell = ws.cell(row=1, column=1, value=title)
    title_cell.font = Font(bold=True, size=14)
    title_cell.alignment = Alignment(horizontal="center")
    # ✅ 新增這一行：加高標題列的列高（單位：points）
    ws.row_dimensions[1].height = 30

    headers = ["日期", "最高價", "最低價", ""] * 3
    for i, h in enumerate(headers):
        cell = ws.cell(row=2, column=i + 1, value=h)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
        # cell.border = thin

    starts = [1, 5, 9]
    for block, data in enumerate(chunks):
        col = starts[block]
        for i, row in data.iterrows():
            r = i + 3
            ws.cell(row=r, column=col, value=row["日期"]).alignment = Alignment(horizontal="center")
            h = ws.cell(row=r, column=col+1, value=row["最高價"])
            h.font = Font(color=row["高色"])
            h.alignment = Alignment(horizontal="center")
            l = ws.cell(row=r, column=col+2, value=row["最低價"])
            l.font = Font(color=row["低色"])
            l.alignment = Alignment(horizontal="center")
            v = ws.cell(row=r, column=col+3, value=row["成交符"])
            v.font = Font(color=row["符色"])
            v.alignment = Alignment(horizontal="center")
            # for j in range(4):
            #     ws.cell(row=r, column=col+j).border = thin

    for col_cells in ws.iter_cols(min_row=3, max_col=ws.max_column, max_row=ws.max_row):
        col_letter = get_column_letter(col_cells[0].column)
        max_len = max(len(str(c.value)) if c.value else 0 for c in col_cells)
        ws.column_dimensions[col_letter].width = max(6, min(max_len + 2, 16))

    # 列印設定
    ws.freeze_panes = "A4"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = 9
    ws.sheet_properties = WorksheetProperties(
        pageSetUpPr=PageSetupProperties(fitToPage=True)
    )
    # ws.sheet_view.showGridLines = False  # ✅ 取消格線顯示

    # 提供下載
    buffer = BytesIO()
    wb.save(buffer)
    st.success("✅ 報表產出成功")
    st.download_button("下載 Excel", data=buffer.getvalue(), file_name=f"{title}.xlsx")
