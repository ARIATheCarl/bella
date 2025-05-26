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

interval = st.radio("選擇統計區間", ["日", "週", "月"], horizontal=True)

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

def weekday_str(dt):
    return ["一", "二", "三", "四", "五", "六", "日"][dt.weekday()]

def week_of_month(dt):
    # 計算當月第幾週
    first_day = dt.replace(day=1)
    dom = dt.day
    adjusted_dom = dom + first_day.weekday()
    return int((adjusted_dom - 1) / 7 + 1)

if st.button("產出報表"):
    stock = twstock.Stock(stock_id)
    # 先抓出所有區間的原始資料
    raw_data = stock.fetch_from(start_date.year, start_date.month)
    filtered = [d for d in raw_data if start_date <= d.date <= end_date]

    if not filtered:
        st.error("查無資料，請檢查股票代碼與時間範圍。")
        st.stop()

    # 建立 DataFrame
    df = pd.DataFrame([{
        '日期': d.date,
        '最高價': d.high,
        '最低價': d.low,
        '成交量': d.capacity
    } for d in filtered])

    # 若要比對第一筆資料用，補一筆上一日
    extra_date = start_date - timedelta(days=5)
    extra_data = stock.fetch_from(extra_date.year, extra_date.month)
    extra_point = next((d for d in reversed(extra_data) if d.date < start_date), None)
    if extra_point:
        df = pd.concat([pd.DataFrame([{
            '日期': extra_point.date,
            '最高價': extra_point.high,
            '最低價': extra_point.low,
            '成交量': extra_point.capacity
        }]), df], ignore_index=True)

    # 統計區間分類
    if interval == "日":
        group_keys = df.index[1:]  # 每天就是一個 key（排除比對用的首筆）
        agg_df = df.iloc[1:].copy()
    elif interval == "週":
        df["week"] = df["日期"].apply(lambda x: x.isocalendar()[1])
        df["year"] = df["日期"].apply(lambda x: x.year)
        # 以 year, week 分群
        grouped = df.iloc[1:].groupby(["year", "week"])
        agg_df = grouped.agg({
            "日期": "last",
            "最高價": "max",
            "最低價": "min",
            "成交量": "sum"
        }).reset_index(drop=True)
    else:  # interval == "月"
        df["month"] = df["日期"].apply(lambda x: x.month)
        df["year"] = df["日期"].apply(lambda x: x.year)
        grouped = df.iloc[1:].groupby(["year", "month"])
        agg_df = grouped.agg({
            "日期": "last",
            "最高價": "max",
            "最低價": "min",
            "成交量": "sum"
        }).reset_index(drop=True)

    # 差異色、成交符號、其他欄位計算
    agg_df["差色"] = ""
    agg_df["高色"] = ""
    agg_df["低色"] = ""
    agg_df["成交符"] = ""
    agg_df["符色"] = ""
    prev = df.iloc[0]  # 第一筆拿來比對
    for i, row in agg_df.iterrows():
        diff = row["最高價"] - row["最低價"]
        prev_diff = prev["最高價"] - prev["最低價"]
        agg_df.loc[i, "高色"] = "FF0000" if row["最高價"] >= prev["最高價"] else "0000FF"
        agg_df.loc[i, "低色"] = "FF0000" if row["最低價"] >= prev["最低價"] else "0000FF"
        agg_df.loc[i, "成交符"] = "🔴" if row["成交量"] >= prev["成交量"] else "🔵"
        agg_df.loc[i, "符色"] = "FF0000" if agg_df.loc[i, "成交符"] == "🔴" else "0000FF"
        agg_df.loc[i, "差色"] = "FF0000" if diff >= prev_diff else "0000FF"
        prev = row

    # 切成3區塊
    base = len(agg_df) // 3
    remainder = len(agg_df) % 3
    sizes = [base + (1 if i < remainder else 0) for i in range(3)]
    chunks = []
    s = 0
    for size in sizes:
        chunks.append(agg_df.iloc[s:s+size].reset_index(drop=True))
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
    title = f"{selected} {start_date.strftime('%Y-%m-%d')}～{end_date.strftime('%Y-%m-%d')}（{interval}）"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=20)
    title_cell = ws.cell(row=1, column=1, value=title)
    title_cell.font = Font(bold=True, size=14)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    # 設定標題列，第一二欄合併為「日期」
    headers = ["日期", "", "高", "低", "漲幅", "量", ""] * 3
    # 合併「日期」欄
    for block in range(3):
        col = block * 7 + 1
        ws.merge_cells(start_row=2, start_column=col, end_row=2, end_column=col+1)
        cell = ws.cell(row=2, column=col, value="日期")
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
        # 其他欄
        for idx, h in enumerate(headers[2:7], 3):
            cell2 = ws.cell(row=2, column=col+idx-1, value=h)
            cell2.font = Font(bold=True)
            cell2.alignment = Alignment(horizontal="center")

    # 寫入內容
    starts = [1, 8, 15]
    for block, data in enumerate(chunks):
        col = starts[block]
        row_index = 3
        prev_month = prev_year = None
        for i, row in data.iterrows():
            dt = row["日期"]

            # ========== 日期/週/月分流 ==============
            if interval == "日":
                # 換月首日顯示 M/D
                if i == 0 or dt.month != prev_month:
                    day_str = f"{dt.month}/{dt.day}"
                else:
                    day_str = f"{dt.day}"
                # 星期
                week_str = weekday_str(dt)
                prev_month = dt.month
                # 填入資料
                ws.cell(row=row_index, column=col, value=day_str).alignment = Alignment(horizontal="center")
                ws.cell(row=row_index, column=col+1, value=week_str).alignment = Alignment(horizontal="center")
            elif interval == "週":
                # 換年顯示 YYYY/M/D
                if i == 0 or dt.year != prev_year:
                    day_str = f"{dt.year}/{dt.month}/{dt.day}"
                else:
                    day_str = f"{dt.month}/{dt.day}"
                # 計算該月第幾週
                w_of_m = week_of_month(dt)
                prev_year = dt.year
                ws.cell(row=row_index, column=col, value=day_str).alignment = Alignment(horizontal="center")
                ws.cell(row=row_index, column=col+1, value=w_of_m).alignment = Alignment(horizontal="center")
            else:  # interval == "月"
                # 換年顯示 YYYY/M
                if i == 0 or dt.year != prev_year:
                    day_str = f"{dt.year}/{dt.month}"
                else:
                    day_str = f"{dt.month}"
                m_of_y = dt.month
                prev_year = dt.year
                ws.cell(row=row_index, column=col, value=day_str).alignment = Alignment(horizontal="center")
                ws.cell(row=row_index, column=col+1, value=m_of_y).alignment = Alignment(horizontal="center")

            # 其餘欄位
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
