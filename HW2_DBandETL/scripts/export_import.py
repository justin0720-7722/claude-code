# -*- coding: utf-8 -*-
"""
Step 3：把已計算好的工作檔「貼上值」成匯入檔（.xlsx，只含資料欄、不含輔助欄），
並輸出給 Access 建置程式用的 TSV。
用法：python export_import.py <已計算工作檔.xlsx> <匯入檔.xlsx> <tsv輸出資料夾>
"""
import datetime as dt
import os
import re
import sys

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from spec import BASE_DATE, TABLES

EXCEL_FMT = {"DATE": "yyyy/mm/dd", "DATETIME": "yyyy/mm/dd hh:mm", "MONEY": "#,##0"}


def to_dt(v):
    if isinstance(v, dt.datetime):
        return v
    return dt.datetime(1899, 12, 30) + dt.timedelta(days=float(v))


def convert(v, typ):
    if v is None or v == "":
        return None
    if typ in ("TEXT", "MEMO"):
        return str(v)
    if typ == "DATE":
        return to_dt(v).replace(hour=0, minute=0, second=0, microsecond=0)
    if typ == "DATETIME":
        d = to_dt(v)
        return d.replace(second=0, microsecond=0) + dt.timedelta(minutes=round(d.second / 60))
    return int(round(float(v)))


def like(pattern, s):
    out, i = "", 0
    while i < len(pattern):
        ch = pattern[i]
        if ch == "[":
            j = pattern.index("]", i)
            out += pattern[i:j + 1]
            i = j + 1
            continue
        out += {"#": r"\d", "*": ".*", "?": "."}.get(ch, re.escape(ch))
        i += 1
    return re.fullmatch(out, s) is not None


def check_rule(rule, v):
    """支援本規格用到的 Access 驗證規則寫法。"""
    if v is None:
        return True  # 空值由「必填」另外檢查
    rule = rule.replace("Is Null Or ", "")
    if m := re.fullmatch(r'Like "(.+)"', rule):
        return like(m.group(1), v)
    if m := re.fullmatch(r"In \((.+)\)", rule):
        return str(v) in {x.strip().strip('"') for x in m.group(1).split(",")}
    if m := re.fullmatch(r"Between (\d+) And (\d+)", rule):
        return int(m.group(1)) <= v <= int(m.group(2))
    if m := re.fullmatch(r"Between (\d+) And Year\(Date\(\)\)", rule):
        return int(m.group(1)) <= v <= dt.date.today().year
    if m := re.fullmatch(r"Between #(\d+)/(\d+)/(\d+)# And #(\d+)/(\d+)/(\d+)#", rule):
        a = dt.datetime(int(m.group(3)), int(m.group(1)), int(m.group(2)))
        b = dt.datetime(int(m.group(6)), int(m.group(4)), int(m.group(5)))
        return a <= v <= b
    if m := re.fullmatch(r">=#(\d+)/(\d+)/(\d+)#", rule):
        return v >= dt.datetime(int(m.group(3)), int(m.group(1)), int(m.group(2)))
    if rule == ">0":
        return v > 0
    if rule == ">=0":
        return v >= 0
    raise ValueError(f"未支援的驗證規則：{rule}")


def validate(data):
    errors = []
    idx = {t["name"]: {f["name"]: i for i, f in enumerate(t["fields"])} for t in TABLES}

    def col(tn, fn):
        return [r[idx[tn][fn]] for r in data[tn]]

    for t in TABLES:
        tn = t["name"]
        if len(data[tn]) < 30:
            errors.append(f"{tn} 只有 {len(data[tn])} 筆")
        for f in t["fields"]:
            vals = col(tn, f["name"])
            nn = [v for v in vals if v is not None]
            if f["req"] and len(nn) != len(vals):
                errors.append(f"{tn}.{f['name']} 必填卻有空值")
            if f["type"] == "TEXT" and any(len(v) > f["size"] for v in nn):
                errors.append(f"{tn}.{f['name']} 超過長度 {f['size']}")
            if f["key"] in ("PK", "UQ") and len(set(nn)) != len(nn):
                errors.append(f"{tn}.{f['name']} 有重複值")
            if f["key"].startswith("FK:"):
                pt, pf = f["key"][3:].split(".")
                bad = set(nn) - set(col(pt, pf))
                if bad:
                    errors.append(f"{tn}.{f['name']} 外鍵對不到：{sorted(bad)[:5]}")
            if f["rule"]:
                bad = [v for v in vals if not check_rule(f["rule"], v)]
                if bad:
                    errors.append(f"{tn}.{f['name']} 不符合 {f['rule']}：{bad[:3]}")
            if f["lookup"] and f["lookup"][0] == "values" and not set(nn) <= set(f["lookup"][1]):
                errors.append(f"{tn}.{f['name']} 有不在查閱清單內的值")
        for u in t.get("uniques", []):
            keys = list(zip(*(col(tn, x) for x in u)))
            if len(set(keys)) != len(keys):
                errors.append(f"{tn} ({'+'.join(u)}) 有重複")

    # ---- 跨表規則
    stu = {r[0]: r for r in data["StudentTable"]}
    act = {r[0]: r for r in data["ActivityTable"]}
    for s in data["StudentTable"]:
        if int(s[0][1:3]) != 116 - s[3] - 100:
            errors.append(f"學號 {s[0]} 入學年與年級 {s[3]} 不一致")
    for c in data["ClubTable"]:
        if stu[c[3]][3] < 2:
            errors.append(f"社團 {c[1]} 的社長不是 2 年級以上")
    count = {}
    for r in data["RegistrationTable"]:
        count[r[1]] = count.get(r[1], 0) + 1
        a = act[r[1]]
        if not (r[3] < a[3] and r[3] < BASE_DATE + dt.timedelta(days=1)):
            errors.append(f"報名 {r[0]} 報名時間不合理")
        if a[3] >= BASE_DATE and r[4] != "未簽到":
            errors.append(f"報名 {r[0]} 活動未舉辦卻有出席紀錄")
        if (r[5] == "免繳") != (a[6] == 0):
            errors.append(f"報名 {r[0]} 繳費狀態與費用不一致")
    for a in data["ActivityTable"]:
        n = count.get(a[0], 0)
        if n == 0 or n > a[5]:
            errors.append(f"活動 {a[0]} 報名人數 {n} / 上限 {a[5]} 不合理")
        if a[3] < BASE_DATE and a[7] not in ("已結束", "取消"):
            errors.append(f"活動 {a[0]} 已過期卻是 {a[7]}")
    return errors


def main(src, dst, tsv_dir):
    wb = load_workbook(src, data_only=True)
    data = {}
    for t in TABLES:
        ws = wb[t["name"]]
        names = [f["name"] for f in t["fields"]]
        assert [c.value for c in ws[1]][:len(names)] == names, t["name"]
        data[t["name"]] = [
            [convert(v, f["type"]) for v, f in zip(r, t["fields"])]
            for r in ws.iter_rows(min_row=2, max_row=t["rows"] + 1, max_col=len(names), values_only=True)
        ]

    errors = validate(data)
    if errors:
        print("驗證失敗：", *errors[:30], sep="\n  ")
        sys.exit(1)
    print("驗證通過：", {k: len(v) for k, v in data.items()})

    # 匯入檔：純值、每張表一個工作表；工作表名 = Access 資料表名，欄名 = Access 欄位名
    out = Workbook()
    out.remove(out.active)
    for t in TABLES:
        ws = out.create_sheet(t["name"])
        ws.append([f["name"] for f in t["fields"]])
        for c in ws[1]:
            c.font, c.fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F4E78")
        for row in data[t["name"]]:
            ws.append(row)
        for j, f in enumerate(t["fields"], 1):
            L = get_column_letter(j)
            fmt = "@" if f["type"] in ("TEXT", "MEMO") else EXCEL_FMT.get(f["type"])
            if fmt:
                for c in ws[L][1:]:
                    c.number_format = fmt
            ws.column_dimensions[L].width = {"TEXT": max(10, min(28, (f["size"] or 10) + 2)), "MEMO": 50,
                                             "DATETIME": 17}.get(f["type"], 13)
        ws.freeze_panes = "A2"
    out.save(dst)
    print("saved", dst)

    # 給 Java (Jackcess) 的 TSV：schema.tsv、uniques.tsv、每表資料
    os.makedirs(tsv_dir, exist_ok=True)
    with open(os.path.join(tsv_dir, "schema.tsv"), "w", encoding="utf-8") as fh:
        for t in TABLES:
            for f in t["fields"]:
                lk = f["lookup"]
                lk_s = ["", "", "", ""] if not lk else (
                    ["values", ";".join(f'"{x}"' for x in lk[1]), "1", ""] if lk[0] == "values"
                    else ["query", lk[1], str(lk[2]), lk[3]])
                fh.write("\t".join([t["name"], f["name"], f["caption"], f["type"], str(f["size"] or 0),
                                    "1" if f["req"] else "0", f["key"], "1" if f["index"] else "0",
                                    f["rule"], f["rtext"], f["fmt"], f["desc"], *lk_s]) + "\n")
    with open(os.path.join(tsv_dir, "uniques.tsv"), "w", encoding="utf-8") as fh:
        for t in TABLES:
            for u in t.get("uniques", []):
                fh.write("\t".join([t["name"], *u]) + "\n")
    for t in TABLES:
        with open(os.path.join(tsv_dir, t["name"] + ".tsv"), "w", encoding="utf-8") as fh:
            for row in data[t["name"]]:
                fh.write("\t".join("" if v is None else v.strftime("%Y-%m-%dT%H:%M") if isinstance(v, dt.datetime)
                                   else str(v) for v in row) + "\n")


if __name__ == "__main__":
    main(*sys.argv[1:4])
