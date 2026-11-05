# -*- coding: utf-8 -*-
"""Step 1：產生含 RANDBETWEEN 等公式的 Excel 假資料工作檔（尚未計算）。"""
import re
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

from spec import DESIGN_CHANGES, LISTS, TABLES

ACCESS_TYPE_ZH = {
    "TEXT": "簡短文字", "BYTE": "數字(位元組)", "INT": "數字(整數)", "LONG": "數字(長整數)",
    "AUTONUMBER": "自動編號", "MONEY": "貨幣", "DOUBLE": "數字(雙精準數)",
    "DATE": "日期/時間", "BOOLEAN": "是/否",
}
EXCEL_FMT = {"DATE": "yyyy/mm/dd", "MONEY": "#,##0", "DOUBLE": "0%"}

HEAD_FILL = PatternFill("solid", fgColor="1F4E78")
HEAD_FONT = Font(bold=True, color="FFFFFF")
SUB_FILL = PatternFill("solid", fgColor="DDEBF7")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def col_letter(table, field):
    t = next(t for t in TABLES if t["name"] == table)
    idx = [f[0] for f in t["fields"]].index(field)
    return get_column_letter(idx + 1)


def render(formula, table, row):
    def same_row(m):
        return f"{col_letter(table, m.group(1))}{row}"

    def other(m):
        t, f = m.group(1).split(".")
        c = col_letter(t, f)
        return f"{t}!${c}:${c}"

    formula = re.sub(r"\{F:([^}]+)\}", same_row, formula)
    formula = re.sub(r"\{R:([^}]+)\}", other, formula)
    return formula


def header(ws, values, row=1):
    for i, v in enumerate(values, 1):
        c = ws.cell(row=row, column=i, value=v)
        c.fill, c.font, c.border = HEAD_FILL, HEAD_FONT, BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def build(path):
    wb = Workbook()

    # ---------------- 說明
    ws = wb.active
    ws.title = "說明"
    lines = [
        ("HW2 建置資料庫與匯入測試資料 — 線上書店訂購管理系統", "title"),
        ("學號：________　　姓名：________", ""),
        ("", ""),
        ("一、工作檔結構", "h"),
        ("說明　　：本頁，作業流程與設計調整紀錄", ""),
        ("欄位規格：每張資料表的欄位、Access 資料類型、大小、驗證規則、範圍，以及產生該欄位的 Excel 公式", ""),
        ("參數　　：亂數取值用的清單（姓氏、名字、縣市、書名字庫…），皆已定義為「名稱」供公式使用", ""),
        ("類別 / 出版社 / 書籍 / 會員 / 訂單 / 訂單明細：六張資料表的假資料，每格都是公式", ""),
        ("檢核　　：主鍵重複、外鍵對不到、數值超出範圍的自動檢查，全部應為 0", ""),
        ("", ""),
        ("二、ETL 流程", "h"),
        ("E (Extract)：在本工作檔用 RANDBETWEEN / INDEX / MATCH / TEXT / CHOOSE 等公式產生假資料（按 F9 可重新亂數）", ""),
        ("T (Transform)：確認「檢核」頁全部為 0 → 全選各表 → 複製 → 選擇性貼上「值」到匯入檔（固定亂數結果、保留文字型態）", ""),
        ("L (Load)：Access「外部資料 → 新增資料來源 → 從檔案 → Excel」，選匯入檔，依序把六張工作表「附加至」已建好的資料表", ""),
        ("　　　　　匯入順序要先父表後子表：類別 → 出版社 → 書籍 → 會員 → 訂單 → 訂單明細（否則參考完整性會擋下）", ""),
        ("", ""),
        ("三、資料筆數", "h"),
    ]
    for t in TABLES:
        lines.append((f"{t['name']}：{t['rows']} 筆　（{t['desc']}）", ""))
    lines += [("", ""), ("四、逐表設計時所做的調整與修正", "h")]
    for i, (what, why) in enumerate(DESIGN_CHANGES, 1):
        lines.append((f"{i}. {what}", "b"))
        lines.append((f"　 → {why}", ""))
    for r, (text, style) in enumerate(lines, 1):
        c = ws.cell(row=r, column=1, value=text)
        if style == "title":
            c.font = Font(bold=True, size=16, color="1F4E78")
        elif style == "h":
            c.font = Font(bold=True, size=12, color="FFFFFF")
            c.fill = HEAD_FILL
        elif style == "b":
            c.font = Font(bold=True)
    ws.column_dimensions["A"].width = 120

    # ---------------- 欄位規格
    spec = wb.create_sheet("欄位規格")
    cols = ["資料表", "欄位名稱", "Access 資料類型", "欄位大小", "必填", "鍵 / 索引",
            "驗證規則", "驗證文字", "格式", "說明 / 數據範圍", "第一筆範例", "Excel 產製公式（第 2 列）"]
    header(spec, cols)
    r = 2
    for t in TABLES:
        for i, f in enumerate(t["fields"]):
            name, typ, size, req, key, rule, rtext, fmt, desc, formula = f
            if typ == "TEXT":
                size_s = str(size)
            elif typ in ("BYTE", "INT", "LONG", "DOUBLE", "AUTONUMBER"):
                size_s = {"BYTE": "位元組", "INT": "整數", "LONG": "長整數",
                          "DOUBLE": "雙精準數", "AUTONUMBER": "長整數"}[typ]
            else:
                size_s = "—"
            key_s = {"PK": "主索引鍵", "UQ": "索引(不可重複)"}.get(key, key.replace("FK:", "外鍵→") if key else "")
            letter = get_column_letter(i + 1)
            vals = [t["name"], name, ACCESS_TYPE_ZH[typ], size_s, "是" if req else "否", key_s,
                    rule, rtext, fmt, desc, f"='{t['name']}'!{letter}2", render(formula, t["name"], 2)]
            for j, v in enumerate(vals, 1):
                c = spec.cell(row=r, column=j, value=v)
                c.border = BOX
                c.alignment = Alignment(vertical="top", wrap_text=(j in (10, 12)))
            # 公式欄要顯示「公式文字」而非計算結果，所以強制存成字串
            spec.cell(row=r, column=12).data_type = "s"
            fmt_x = EXCEL_FMT.get(typ)
            if fmt_x:
                spec.cell(row=r, column=11).number_format = fmt_x
            if t["name"] in ("出版社", "會員", "訂單明細"):
                for j in range(1, len(cols) + 1):
                    spec.cell(row=r, column=j).fill = SUB_FILL  # 隔表底色，方便閱讀
            r += 1
    for j, w in enumerate([10, 12, 15, 9, 6, 20, 32, 28, 12, 34, 16, 90], 1):
        spec.column_dimensions[get_column_letter(j)].width = w
    spec.freeze_panes = "C2"

    # ---------------- 參數
    ps = wb.create_sheet("參數")
    for j, (lname, items) in enumerate(LISTS.items(), 1):
        c = ps.cell(row=1, column=j, value=lname)
        c.fill, c.font, c.border = HEAD_FILL, HEAD_FONT, BOX
        for i, v in enumerate(items, 2):
            ps.cell(row=i, column=j, value=v).border = BOX
        L = get_column_letter(j)
        dn = DefinedName(lname, attr_text=f"參數!${L}$2:${L}${len(items) + 1}")
        wb.defined_names[lname] = dn
        ps.column_dimensions[L].width = 13
    ps.freeze_panes = "A2"

    # ---------------- 六張資料表
    for t in TABLES:
        ws = wb.create_sheet(t["name"])
        header(ws, [f[0] for f in t["fields"]])
        for r in range(2, t["rows"] + 2):
            for j, f in enumerate(t["fields"], 1):
                c = ws.cell(row=r, column=j, value=render(f[9], t["name"], r))
                fmt_x = EXCEL_FMT.get(f[1])
                if fmt_x:
                    c.number_format = fmt_x
        for j, f in enumerate(t["fields"], 1):
            w = 12 if f[1] != "TEXT" else max(10, min(30, (f[2] or 10) + 4))
            ws.column_dimensions[get_column_letter(j)].width = w
        ws.freeze_panes = "B2"

    # ---------------- 檢核
    ck = wb.create_sheet("檢核")
    header(ck, ["檢核項目", "錯誤筆數（應為 0）"])
    checks = []

    def rng(table, field):
        n = next(t for t in TABLES if t["name"] == table)["rows"]
        L = col_letter(table, field)
        return f"{table}!${L}$2:${L}${n + 1}"

    for t in TABLES:
        n = t["rows"]
        checks.append((f"{t['name']}：資料筆數不足 30", f'=IF(COUNTA({rng(t["name"], t["fields"][0][0])})>=30,0,1)'))
        for f in t["fields"]:
            name, typ, size, req, key = f[:5]
            R = rng(t["name"], name)
            if key in ("PK", "UQ"):
                checks.append((f"{t['name']}.{name}：值重複", f"=SUMPRODUCT(--(COUNTIF({R},{R})>1))"))
            if key.startswith("FK:"):
                pt, pf = key[3:].split(".")
                checks.append((f"{t['name']}.{name}：外鍵對不到 {pt}",
                               f"=SUMPRODUCT(--(COUNTIF({rng(pt, pf)},{R})=0))"))
            if typ == "TEXT":
                checks.append((f"{t['name']}.{name}：超過欄位大小 {size}", f"=SUMPRODUCT(--(LEN({R})>{size}))"))
            if req:
                checks.append((f"{t['name']}.{name}：必填卻空白", f'=COUNTBLANK({R})'))
    checks.append(("訂單：訂購日期早於會員加入日期",
                   "=SUMPRODUCT(--(" + rng("訂單", "訂購日期") + "<SUMIF(" + rng("會員", "會員編號") + ","
                   + rng("訂單", "會員編號") + "," + rng("會員", "加入日期") + ")))"))
    checks.append(("訂單：沒有任何明細的訂單",
                   f"=SUMPRODUCT(--(COUNTIF({rng('訂單明細', '訂單編號')},{rng('訂單', '訂單編號')})=0))"))
    checks.append(("書籍：庫存 0 卻上架",
                   f"=SUMPRODUCT(({rng('書籍', '庫存量')}=0)*({rng('書籍', '是否上架')}=TRUE))"))
    for i, (label, formula) in enumerate(checks, 2):
        ck.cell(row=i, column=1, value=label).border = BOX
        ck.cell(row=i, column=2, value=formula).border = BOX
    total_row = len(checks) + 2
    ck.cell(row=total_row, column=1, value="合計（全部為 0 才可匯入）").font = Font(bold=True)
    ck.cell(row=total_row, column=2, value=f"=SUM(B2:B{total_row - 1})").font = Font(bold=True, color="C00000")
    ck.column_dimensions["A"].width = 46
    ck.column_dimensions["B"].width = 20

    wb.save(path)
    print(f"saved {path}: {len(checks)} checks")


if __name__ == "__main__":
    build(sys.argv[1])
