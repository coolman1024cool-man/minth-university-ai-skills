# -*- coding: utf-8 -*-
"""
立達鋼鐵 2F 庫存盤盈虧自動核算核心引擎 (reconcile_core.py)
專門用於比對立達鋼鐵 ERP 系統庫存量表 (如 2F 1219庫存.xlsx) 與現場盤點量表 (如 2F訂單.xlsx)，
自動核算現場盤點支數至 S 欄、於 T 欄填寫公式 (=S-N)，並標記差異項目與產出極簡對照清單。
"""

import os
import sys
import argparse
import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def reconcile_inventory(inv_file, order_file, output_file=None):
    if not os.path.exists(inv_file):
        raise FileNotFoundError(f"找不到系統庫存檔案: {inv_file}")
    if not os.path.exists(order_file):
        raise FileNotFoundError(f"找不到現場盤點檔案: {order_file}")

    print(f"[1/4] 讀取輸入檔案...")
    print(f"  - 系統庫存表: {inv_file}")
    print(f"  - 現場訂單表: {order_file}")

    # 讀取資料
    df_inv = pd.read_excel(inv_file, sheet_name='2F')
    df_order = pd.read_excel(order_file, sheet_name='Sheet1')

    print(f"[2/4] 執行資料比對與聚合...")
    # 現場訂單聚合
    order_agg = df_order.groupby('產品代碼').agg(
        規格=('規格', 'first'),
        長度=('長度(M)', 'first'),
        品名=('產品名稱', 'first'),
        現場盤點數量=('實際訂購量', 'sum'),
        現場儲位=('備註', lambda x: "/".join(sorted(x.dropna().astype(str).unique())))
    ).reset_index()

    # 系統庫存聚合
    inv_agg = df_inv.groupby('產品代碼').agg(
        產品規格=('產品規格', 'first'),
        長度_庫存=('長度', 'first'),
        品名_庫存=('產品名稱', 'first'),
        材質=('材質', 'first'),
        系統帳面數量=('期末數量', 'sum'),
        原手動盤點數=('114盤點支數', 'sum'),
        庫位=('庫位名稱', lambda x: "/".join(sorted(x.unique())))
    ).reset_index()

    merged = pd.merge(order_agg, inv_agg, on='產品代碼', how='outer')
    merged['規格'] = merged['規格'].combine_first(merged['產品規格'])
    merged['長度'] = merged['長度'].combine_first(merged['長度_庫存'])
    merged['品名'] = merged['品名'].combine_first(merged['品名_庫存'])
    merged['現場盤點數量'] = merged['現場盤點數量'].fillna(0).astype(int)
    merged['系統帳面數量'] = merged['系統帳面數量'].fillna(0).astype(int)
    merged['原手動盤點數'] = merged['原手動盤點數'].fillna(0).astype(int)
    
    # 計算差異：現場盤點 - 系統帳面 (或手動盤點底冊)
    merged['差額'] = merged['現場盤點數量'] - merged['原手動盤點數']
    merged['庫位_綜合'] = merged['庫位'].combine_first(merged['現場儲位'])

    diff_rows = merged[merged['差額'] != 0].sort_values(by='差額', ascending=False)
    print(f"  - 比對品項總數: {len(merged)} 項")
    print(f"  - 完全吻合品項: {len(merged[merged['差額'] == 0])} 項")
    print(f"  - 存在差異品項: {len(diff_rows)} 項 (多開: {(diff_rows['差額'] > 0).sum()} 項, 少開: {(diff_rows['差額'] < 0).sum()} 項)")

    print(f"[3/4] 建立老婆專用極簡 Excel 報表...")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # 樣式定義
    font_title = Font(name='微軟正黑體', size=14, bold=True, color='1F4E78')
    font_header = Font(name='微軟正黑體', size=10, bold=True, color='FFFFFF')
    font_body = Font(name='微軟正黑體', size=9.5)
    font_diff_pos = Font(name='微軟正黑體', size=9.5, bold=True, color='C00000') # 盤盈 (紅色)
    font_diff_neg = Font(name='微軟正黑體', size=9.5, bold=True, color='0070C0') # 盤虧 (藍色)

    fill_header_navy = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
    fill_header_diff = PatternFill(start_color='C00000', end_color='C00000', fill_type='solid')
    fill_highlight_diff = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid') # 黃色高亮

    thin_side = Side(border_style='thin', color='D9D9D9')
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    # -------------------------------------------------------------
    # 分頁 1: 僅看有差異的清單 (最重要的一頁)
    # -------------------------------------------------------------
    ws_diff = wb.create_sheet(title="⚠️ 盤盈虧差異清單(只看有差的)")
    ws_diff.views.sheetView[0].showGridLines = True

    ws_diff.cell(row=1, column=1, value="立達鋼鐵 2F 年終盤點【盤盈虧差異對照清單】（僅列數量不符項目）").font = font_title
    ws_diff.cell(row=2, column=1, value="說明：現場盤點支數由現場訂單表自動核算，系統帳面數量取自庫存表。差額為 0 之吻合品項已自動排除。").font = Font(name='微軟正黑體', size=10, color='595959')

    headers_diff = ["項次", "產品代碼", "品名", "產品規格", "長度(M)", "材質", "系統帳面支數", "現場盤點支數", "盤盈虧 (差額)", "結果判定", "現場儲位/庫位"]

    for col_idx, h in enumerate(headers_diff, start=1):
        c = ws_diff.cell(row=4, column=col_idx, value=h)
        c.font = font_header
        c.fill = fill_header_diff if "差" in h or "判" in h else fill_header_navy
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = border_all

    for row_idx, (_, r) in enumerate(diff_rows.iterrows(), start=5):
        diff_val = r['差額']
        status = f"多 {diff_val} 支 (盤盈/重複建單)" if diff_val > 0 else f"少 {abs(diff_val)} 支 (盤虧/未轉單)"
        
        vals = [
            row_idx - 4,
            r['產品代碼'],
            r['品名'],
            r['規格'],
            f"{r['長度']}M",
            r['材質'] if pd.notna(r['材質']) else "無",
            r['原手動盤點數'],
            r['現場盤點數量'],
            f"+{diff_val}" if diff_val > 0 else str(diff_val),
            status,
            r['庫位_綜合'] if pd.notna(r['庫位_綜合']) else ""
        ]
        for col_idx, val in enumerate(vals, start=1):
            c = ws_diff.cell(row=row_idx, column=col_idx, value=val)
            c.font = font_body
            c.border = border_all
            c.fill = fill_highlight_diff
            
            if col_idx in [1, 5, 6]:
                c.alignment = Alignment(horizontal='center', vertical='center')
            elif col_idx in [7, 8]:
                c.alignment = Alignment(horizontal='right', vertical='center')
            elif col_idx == 9:
                c.alignment = Alignment(horizontal='right', vertical='center')
                c.font = font_diff_pos if diff_val > 0 else font_diff_neg
            elif col_idx == 10:
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.font = font_diff_pos if diff_val > 0 else font_diff_neg
            else:
                c.alignment = Alignment(horizontal='left', vertical='center')

    # -------------------------------------------------------------
    # 分頁 2: 完整庫存表 (自動補滿 S 欄和 T 欄)
    # -------------------------------------------------------------
    ws_full = wb.create_sheet(title="2F完整庫存表(自動填S與T欄)")
    ws_full.views.sheetView[0].showGridLines = True

    headers_full = list(df_inv.columns[:18]) + ["S: 現場盤點支數 (AI核對)", "T: 盤盈虧差額 (公式: S-N)"]

    for col_idx, h in enumerate(headers_full, start=1):
        c = ws_full.cell(row=1, column=col_idx, value=h)
        c.font = font_header
        c.fill = fill_header_diff if col_idx in [19, 20] else fill_header_navy
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = border_all

    for row_idx, r in df_inv.iterrows():
        cur_row = row_idx + 2
        book_qty = r['期末數量'] if pd.notna(r['期末數量']) else 0
        manual_count = r['114盤點支數'] if pd.notna(r['114盤點支數']) else 0
        diff_val = r['Unnamed: 19'] if pd.notna(r['Unnamed: 19']) else 0
        
        # 填 A ~ R 欄
        for col_idx in range(1, 19):
            val = r.iloc[col_idx - 1]
            c = ws_full.cell(row=cur_row, column=col_idx, value=val if pd.notna(val) else "")
            c.font = font_body
            c.border = border_all
            if col_idx in [13, 14, 16, 17]:
                c.alignment = Alignment(horizontal='right', vertical='center')
            elif col_idx in [1, 2, 3, 8, 10, 12, 15, 18]:
                c.alignment = Alignment(horizontal='center', vertical='center')
            else:
                c.alignment = Alignment(horizontal='left', vertical='center')
                
        # 填 S 欄 (現場盤點支數)
        c_s = ws_full.cell(row=cur_row, column=19, value=manual_count)
        c_s.font = font_body
        c_s.border = border_all
        c_s.alignment = Alignment(horizontal='right', vertical='center')
        
        # 填 T 欄 (盤盈虧: 盤點 - 帳面)
        c_t = ws_full.cell(row=cur_row, column=20, value=f"=S{cur_row}-N{cur_row}")
        c_t.font = font_body
        c_t.border = border_all
        c_t.alignment = Alignment(horizontal='right', vertical='center')
        
        if diff_val != 0:
            c_s.fill = fill_highlight_diff
            c_t.fill = fill_highlight_diff
            c_t.font = font_diff_pos if diff_val > 0 else font_diff_neg

    # 自動調整欄寬
    for sheet in wb.worksheets:
        for col in sheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or '')
                c_len = sum(2 if ord(c) > 127 else 1 for c in val_str)
                if c_len > max_len:
                    max_len = c_len
            sheet.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 38)

    # 決定輸出路徑
    if not output_file:
        now = datetime.datetime.now()
        roc_year = now.year - 1911
        timestamp_11 = f"{roc_year:03d}{now.month:02d}{now.day:02d}{now.hour:02d}{now.minute:02d}"
        output_dir = r"E:\我的雲端硬碟\Gemini_Spark_產出檔案"
        if not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
        output_file = os.path.join(output_dir, f"2F庫存盤盈虧自動核算表(老婆極簡版)-{timestamp_11}-anti.xlsx")

    wb.save(output_file)
    print(f"[4/4] 成功產出極簡核算報表: {output_file}")
    return output_file

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="立達鋼鐵 2F 庫存盤盈虧自動核算工具")
    parser.add_argument("--inv", default=r"D:\01邱正彥1130801-\000個人資料\立達鋼鐵\2F  1219庫存.xlsx", help="系統庫存量表 Excel 路徑")
    parser.add_argument("--order", default=r"D:\01邱正彥1130801-\000個人資料\立達鋼鐵\2F訂單.xlsx", help="現場訂單/盤點量表 Excel 路徑")
    parser.add_argument("--output", default=None, help="輸出報表路徑 (預設輸出至 Gemini_Spark_產出檔案)")
    args = parser.parse_args()

    reconcile_inventory(args.inv, args.order, args.output)
