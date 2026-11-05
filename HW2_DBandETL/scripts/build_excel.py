# -*- coding: utf-8 -*-
"""Step 1：產生含 RANDBETWEEN 等公式的 Excel 假資料工作檔（尚未計算）。"""
import datetime as dt
import re
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

from spec import BASE_DATE, DESIGN_CHANGES, LISTS, SEMESTER, STUDENT_ID, TABLES

ACCESS_TYPE_ZH = {
    "TEXT": "簡短文字", "MEMO": "長文字", "BYTE": "數字", "INT": "數字", "LONG": "數字",
    "AUTONUMBER": "自動編號", "MONEY": "貨幣", "DATE": "日期/時間", "DATETIME": "日期/時間",
}
SIZE_ZH = {"BYTE": "位元組", "INT": "整數", "LONG": "長整數", "AUTONUMBER": "長整數",
           "DATE": "簡短日期", "DATETIME": "日期時間", "MEMO": "—", "MONEY": "—"}
EXCEL_FMT = {"DATE": "yyyy/mm/dd", "DATETIME": "yyyy/mm/dd hh:mm", "MONEY": "#,##0"}

HEAD_FILL = PatternFill("solid", fgColor="1F4E78")
HELP_FILL = PatternFill("solid", fgColor="808080")
HEAD_FONT = Font(bold=True, color="FFFFFF")
SUB_FILL = PatternFill("solid", fgColor="DDEBF7")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def table(name):
    return next(t for t in TABLES if t["name"] == name)


def columns(t):
    """資料欄位 + 輔助欄位（輔助欄放在最右邊，不匯入 Access）。"""
    return [f["name"] for f in t["fields"]] + [h[0] for h in t.get("helpers", [])]


def col_letter(tname, field):
    return get_column_letter(columns(table(tname)).index(field) + 1)


def render(formula, tname, row):
    formula = re.sub(r"\{F:([^}]+)\}", lambda m: f"{col_letter(tname, m.group(1))}{row}", formula)
    formula = re.sub(r"\{C:([^}]+)\}", lambda m: col_letter(tname, m.group(1)), formula)
    formula = formula.replace("{P}", str(row - 1))

    def other(m):
        t, f = m.group(1).split(".")
        c = col_letter(t, f)
        return f"{t}!${c}:${c}"

    return re.sub(r"\{R:([^}]+)\}", other, formula)


def rng(tname, field):
    """固定筆數的範圍，例如 StudentTable!$A$2:$A$61（檢核用）。"""
    L = col_letter(tname, field)
    return f"{tname}!${L}$2:${L}${table(tname)['rows'] + 1}"


def header_cell(ws, row, col, value, fill=HEAD_FILL):
    c = ws.cell(row=row, column=col, value=value)
    c.fill, c.font, c.border = fill, HEAD_FONT, BOX
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    return c


def sheet_readme(wb):
    ws = wb.active
    ws.title = "說明"
    lines = [
        ("HW2 建置資料庫與匯入測試資料 — 校園社團活動管理系統", "title"),
        (f"學號：{STUDENT_ID}　　姓名：________", ""),
        (f"依據：HW1_TableSchemaSketch_{STUDENT_ID}.xlsx（學生、社團、活動、活動報名 四張資料表）", ""),
        (f"資料情境：{SEMESTER}，系統基準日（今天）= {BASE_DATE:%Y/%m/%d}（定義在「參數」頁，名稱：基準日）", ""),
        ("", ""),
        ("一、工作檔結構", "h"),
        ("說明　　：本頁，ETL 流程與設計調整紀錄", ""),
        ("欄位規格：每個欄位的 Access 資料類型、大小、必填、索引、驗證規則、查閱、範圍，以及產生該欄位的 Excel 公式", ""),
        ("參數　　：亂數取值用的清單（姓氏、名字、系所、社團、活動類型、地點…），皆已定義為 Excel「名稱」", ""),
        ("StudentTable / ClubTable / ActivityTable / RegistrationTable：四張表的假資料，每格都是公式", ""),
        ("　　　　　灰色標題的「輔助_」欄是計算用的中間值，不匯入 Access", ""),
        ("檢核　　：主鍵重複、外鍵對不到、超出長度、必填空白、跨表邏輯（超收、報名晚於活動…）自動檢查，合計應為 0", ""),
        ("", ""),
        ("二、ETL 流程", "h"),
        ("E (Extract)：在本工作檔用 RANDBETWEEN / RAND / INDEX / MATCH / LARGE / COUNTIF / TEXT / IF 等公式產生假資料（按 F9 可重新亂數）", ""),
        ("T (Transform)：確認「檢核」頁合計為 0 → 各表選取資料欄（不含輔助欄）→ 複製 → 選擇性貼上「值」到匯入檔", ""),
        ("L (Load)：Access「外部資料 → 新增資料來源 → 從檔案 → Excel」，選匯入檔，把四張工作表「附加至」同名資料表", ""),
        ("　　　　　匯入順序要先父表後子表：StudentTable → ClubTable → ActivityTable → RegistrationTable", ""),
        ("", ""),
        ("三、資料筆數", "h"),
    ]
    for t in TABLES:
        lines.append((f"{t['name']}（{t['zh']}）：{t['rows']} 筆　— {t['desc']}", ""))
    lines += [("", ""), ("四、逐表設計時，對 HW1 草圖所做的調整與修正", "h")]
    for i, (what, why) in enumerate(DESIGN_CHANGES, 1):
        lines.append((f"{i}. {what}", "b"))
        lines.append((f"　 → {why}", ""))
    for r, (text, style) in enumerate(lines, 1):
        c = ws.cell(row=r, column=1, value=text)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if style == "title":
            c.font = Font(bold=True, size=16, color="1F4E78")
        elif style == "h":
            c.font, c.fill = Font(bold=True, size=12, color="FFFFFF"), HEAD_FILL
        elif style == "b":
            c.font = Font(bold=True)
    ws.column_dimensions["A"].width = 130


def sheet_spec(wb):
    ws = wb.create_sheet("欄位規格")
    cols = ["資料表", "欄位名稱", "欄位標題", "Access 資料類型", "欄位大小 / 格式", "必填", "鍵 / 索引",
            "查閱", "驗證規則", "驗證文字", "說明 / 數據範圍", "第一筆範例", "Excel 產製公式（第 2 列）"]
    for j, v in enumerate(cols, 1):
        header_cell(ws, 1, j, v)
    r = 2
    for ti, t in enumerate(TABLES):
        for i, f in enumerate(t["fields"]):
            typ = f["type"]
            size = str(f["size"]) if typ == "TEXT" else SIZE_ZH[typ]
            if f["fmt"]:
                size += f"／{f['fmt']}"
            keys = []
            if f["key"] == "PK":
                keys.append("主索引鍵")
            elif f["key"] == "UQ":
                keys.append("索引(不可重複)")
            elif f["key"].startswith("FK:"):
                keys.append("外鍵→" + f["key"][3:])
            if f["index"]:
                keys.append("索引(可重複)")
            for u in t.get("uniques", []):
                if f["name"] in u:
                    keys.append("複合唯一(" + "+".join(u) + ")" if len(u) > 1 else "索引(不可重複)")
            lk = f["lookup"]
            lookup = "" if not lk else ("值清單：" + "、".join(lk[1]) if lk[0] == "values" else "資料表：" + lk[1])
            vals = [t["name"], f["name"], f["caption"], ACCESS_TYPE_ZH[typ], size, "是" if f["req"] else "否",
                    "；".join(keys), lookup, f["rule"], f["rtext"], f["desc"],
                    f"='{t['name']}'!{get_column_letter(i + 1)}2", render(f["formula"], t["name"], 2)]
            for j, v in enumerate(vals, 1):
                c = ws.cell(row=r, column=j, value=v)
                c.border = BOX
                c.alignment = Alignment(vertical="top", wrap_text=j in (8, 9, 11, 13))
                if ti % 2:
                    c.fill = SUB_FILL  # 隔表底色，方便閱讀
            ws.cell(row=r, column=13).data_type = "s"  # 顯示公式文字，不讓它被計算
            if typ in EXCEL_FMT:
                ws.cell(row=r, column=12).number_format = EXCEL_FMT[typ]
            r += 1
        for h, formula in t.get("helpers", []):
            vals = [t["name"], h, "（輔助欄，不匯入）", "", "", "", "", "", "", "", "產生其他欄位用的中間值", "",
                    render(formula, t["name"], 2)]
            for j, v in enumerate(vals, 1):
                c = ws.cell(row=r, column=j, value=v)
                c.border, c.font = BOX, Font(color="808080", italic=True)
                c.alignment = Alignment(vertical="top", wrap_text=j == 13)
                if ti % 2:
                    c.fill = SUB_FILL
            ws.cell(row=r, column=13).data_type = "s"
            r += 1
    for j, w in enumerate([17, 19, 10, 12, 16, 6, 24, 28, 30, 26, 34, 16, 90], 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "C2"


def sheet_params(wb):
    ws = wb.create_sheet("參數")
    for j, (lname, items) in enumerate(LISTS.items(), 1):
        header_cell(ws, 1, j, lname)
        for i, v in enumerate(items, 2):
            c = ws.cell(row=i, column=j, value=v)
            c.border = BOX
            if isinstance(v, dt.datetime):
                c.number_format = "yyyy/mm/dd"
        L = get_column_letter(j)
        ref = f"參數!${L}$2" if len(items) == 1 else f"參數!${L}$2:${L}${len(items) + 1}"
        wb.defined_names[lname] = DefinedName(lname, attr_text=ref)
        ws.column_dimensions[L].width = 14 if lname not in ("地點清單", "簡介片語一", "簡介片語二") else 22
    ws.freeze_panes = "A2"


def sheet_tables(wb):
    for t in TABLES:
        ws = wb.create_sheet(t["name"])
        allcols = t["fields"] + [dict(name=h[0], formula=h[1], type="HELP", size=None)
                                 for h in t.get("helpers", [])]
        for j, f in enumerate(allcols, 1):
            header_cell(ws, 1, j, f["name"], HELP_FILL if f["type"] == "HELP" else HEAD_FILL)
            w = {"TEXT": max(10, min(28, (f["size"] or 10) + 2)), "MEMO": 50, "DATETIME": 17,
                 "HELP": 15}.get(f["type"], 13)
            ws.column_dimensions[get_column_letter(j)].width = w
        for r in range(2, t["rows"] + 2):
            for j, f in enumerate(allcols, 1):
                c = ws.cell(row=r, column=j, value=render(f["formula"], t["name"], r))
                if f["type"] in EXCEL_FMT:
                    c.number_format = EXCEL_FMT[f["type"]]
                if f["type"] == "HELP":
                    c.font = Font(color="808080")
        ws.freeze_panes = "B2"


def sheet_checks(wb):
    ws = wb.create_sheet("檢核")
    header_cell(ws, 1, 1, "檢核項目")
    header_cell(ws, 1, 2, "錯誤筆數（應為 0）")
    checks = []
    for t in TABLES:
        tn = t["name"]
        checks.append((f"{tn}：資料筆數少於 30", f"=IF(COUNTA({rng(tn, t['fields'][0]['name'])})>=30,0,1)"))
        for f in t["fields"]:
            R = rng(tn, f["name"])
            if f["key"] in ("PK", "UQ"):
                checks.append((f"{tn}.{f['name']}：值重複", f"=SUMPRODUCT(--(COUNTIF({R},{R})>1))"))
            if f["key"].startswith("FK:"):
                pt, pf = f["key"][3:].split(".")
                checks.append((f"{tn}.{f['name']}：外鍵在 {pt} 找不到",
                               f"=SUMPRODUCT(--(COUNTIF({rng(pt, pf)},{R})=0))"))
            if f["type"] == "TEXT":
                checks.append((f"{tn}.{f['name']}：超過欄位大小 {f['size']}", f"=SUMPRODUCT(--(LEN({R})>{f['size']}))"))
            if f["req"]:
                checks.append((f"{tn}.{f['name']}：必填卻空白", f'=SUMPRODUCT(--(LEN({R})=0))'))
        for u in t.get("uniques", []):
            if len(u) > 1:
                a, b = (rng(tn, x) for x in u)
                checks.append((f"{tn}：({'+'.join(u)}) 重複報名", f"=SUMPRODUCT(--(COUNTIFS({a},{a},{b},{b})>1))"))
            else:
                R = rng(tn, u[0])
                checks.append((f"{tn}.{u[0]}：值重複", f"=SUMPRODUCT(--(COUNTIF({R},{R})>1))"))

    A, Rg, S = "ActivityTable", "RegistrationTable", "StudentTable"
    act_id, reg_act = rng(A, "ActivityID"), rng(Rg, "ActivityID")

    def act_lookup(field):  # 每筆報名對應到的活動欄位（ActivityID 唯一，所以 SUMIF = 查值）
        return f"SUMIF({act_id},{reg_act},{rng(A, field)})"

    checks += [
        ("StudentTable.Grade：超出 1～4", f"=SUMPRODUCT(({rng(S, 'Grade')}<1)+({rng(S, 'Grade')}>4))"),
        ("StudentTable：學號入學年與年級不一致",
         f"=SUMPRODUCT(--(VALUE(MID({rng(S, 'StudentID')},2,2))<>116-{rng(S, 'Grade')}-100))"),
        ("ClubTable：社長不是 2 年級以上",
         f"=SUMPRODUCT(--(SUMIF({rng(S, 'StudentID')},{rng('ClubTable', 'PresidentStudentID')},{rng(S, 'Grade')})<2))"),
        ("ActivityTable：報名人數超過最大人數",
         f"=SUMPRODUCT(--(COUNTIF({reg_act},{act_id})>{rng(A, 'MaxParticipants')}))"),
        ("ActivityTable：沒有任何報名的活動", f"=SUMPRODUCT(--(COUNTIF({reg_act},{act_id})=0))"),
        ("ActivityTable：已過基準日卻不是已結束/取消",
         f'=SUMPRODUCT(({rng(A, "ActivityDate")}<基準日)*({rng(A, "ActivityStatus")}<>"已結束")*({rng(A, "ActivityStatus")}<>"取消"))'),
        ("RegistrationTable：報名時間不早於活動日期",
         f"=SUMPRODUCT(--({rng(Rg, 'RegisterTime')}>={act_lookup('ActivityDate')}))"),
        ("RegistrationTable：報名時間晚於基準日", f"=SUMPRODUCT(--({rng(Rg, 'RegisterTime')}>=基準日+1))"),
        ("RegistrationTable：未舉辦的活動卻已出席/缺席",
         f'=SUMPRODUCT(({act_lookup("ActivityDate")}>=基準日)*({rng(Rg, "AttendanceStatus")}<>"未簽到"))'),
        ("RegistrationTable：免繳與活動費用不一致",
         f'=SUMPRODUCT(--(({rng(Rg, "PaymentStatus")}="免繳")<>({act_lookup("Fee")}=0)))'),
    ]
    for i, (label, formula) in enumerate(checks, 2):
        ws.cell(row=i, column=1, value=label).border = BOX
        ws.cell(row=i, column=2, value=formula).border = BOX
    total = len(checks) + 2
    ws.cell(row=total, column=1, value="合計（全部為 0 才可匯入）").font = Font(bold=True)
    ws.cell(row=total, column=2, value=f"=SUM(B2:B{total - 1})").font = Font(bold=True, color="C00000")
    ws.column_dimensions["A"].width = 52
    ws.column_dimensions["B"].width = 20
    return len(checks)


def build(path):
    wb = Workbook()
    sheet_readme(wb)
    sheet_spec(wb)
    sheet_params(wb)
    sheet_tables(wb)
    n = sheet_checks(wb)
    wb.save(path)
    print(f"saved {path}: {n} checks")


if __name__ == "__main__":
    build(sys.argv[1])
