# -*- coding: utf-8 -*-
"""
Step 3：把已計算好的工作檔「貼上值」成匯入檔（.xlsx），並輸出給 Access 建置程式用的 TSV。
用法：python export_import.py <已計算工作檔.xlsx> <匯入檔.xlsx> <tsv輸出資料夾>
"""
import datetime as dt
import os
import re
import sys

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from spec import TABLES

EXCEL_FMT = {"DATE": "yyyy/mm/dd", "MONEY": "#,##0", "DOUBLE": "0.00"}


def to_date(v):
    if isinstance(v, dt.datetime):
        return v
    return dt.datetime(1899, 12, 30) + dt.timedelta(days=float(v))


def convert(v, typ):
    if typ == "TEXT":
        return str(v)
    if typ == "DATE":
        return to_date(v)
    if typ == "BOOLEAN":
        return bool(v)
    if typ in ("BYTE", "INT", "LONG", "AUTONUMBER"):
        return int(v)
    return float(v) if typ == "DOUBLE" else int(round(float(v)))


def like(pattern, s):
    rx = "".join({"#": r"\d", "*": ".*", "?": "."}.get(ch, re.escape(ch)) for ch in pattern)
    rx = rx.replace(r"\[0\-9\]", r"\d")
    return re.fullmatch(rx, s) is not None


def validate(data):
    errors = []
    for t in TABLES:
        rows = data[t["name"]]
        if len(rows) < 30:
            errors.append(f"{t['name']} 只有 {len(rows)} 筆")
        for i, f in enumerate(t["fields"]):
            name, typ, size, req, key, rule = f[:6]
            col = [r[i] for r in rows]
            if req and any(v in (None, "") for v in col):
                errors.append(f"{t['name']}.{name} 有空值")
            if typ == "TEXT" and any(len(v) > size for v in col):
                errors.append(f"{t['name']}.{name} 超過長度 {size}")
            if key in ("PK", "UQ") and len(set(col)) != len(col):
                errors.append(f"{t['name']}.{name} 有重複值")
            if key.startswith("FK:"):
                pt, pf = key[3:].split(".")
                ptab = next(x for x in TABLES if x["name"] == pt)
                pidx = [x[0] for x in ptab["fields"]].index(pf)
                parents = {r[pidx] for r in data[pt]}
                bad = [v for v in col if v not in parents]
                if bad:
                    errors.append(f"{t['name']}.{name} 外鍵對不到：{bad[:5]}")
            m = re.fullmatch(r'Like "(.+)"', rule or "")
            if m and not all(like(m.group(1), v) for v in col):
                errors.append(f"{t['name']}.{name} 不符合 {rule}")
            m = re.fullmatch(r"In \((.+)\)", rule or "")
            if m:
                allowed = {x.strip().strip('"') for x in m.group(1).split(",")}
                if not all(str(v) in allowed for v in col):
                    errors.append(f"{t['name']}.{name} 不符合 {rule}")
            m = re.fullmatch(r"Between ([\d.]+) And ([\d.]+)", rule or "")
            if m and not all(float(m.group(1)) <= v <= float(m.group(2)) for v in col):
                errors.append(f"{t['name']}.{name} 不符合 {rule}")
            if rule in (">=0", ">0"):
                lo_ok = (lambda v: v >= 0) if rule == ">=0" else (lambda v: v > 0)
                if not all(lo_ok(v) for v in col):
                    errors.append(f"{t['name']}.{name} 不符合 {rule}")
    # 跨表規則
    join = {r[0]: r[7] for r in data["會員"]}
    if any(o[2] < join[o[1]] for o in data["訂單"]):
        errors.append("有訂單早於會員加入日期")
    if {o[0] for o in data["訂單"]} - {d[1] for d in data["訂單明細"]}:
        errors.append("有訂單沒有明細")
    return errors


def main(src, dst, tsv_dir):
    wb = load_workbook(src, data_only=True)
    data = {}
    for t in TABLES:
        ws = wb[t["name"]]
        hdr = [c.value for c in ws[1]]
        assert hdr == [f[0] for f in t["fields"]], (t["name"], hdr)
        rows = []
        for r in ws.iter_rows(min_row=2, max_row=t["rows"] + 1, values_only=True):
            rows.append([convert(v, f[1]) for v, f in zip(r, t["fields"])])
        data[t["name"]] = rows

    errors = validate(data)
    if errors:
        print("驗證失敗：", *errors, sep="\n  ")
        sys.exit(1)
    print("驗證通過：", {k: len(v) for k, v in data.items()})

    # 匯入檔：純值、每張表一個工作表，工作表名 = Access 資料表名，欄名 = Access 欄位名
    out = Workbook()
    out.remove(out.active)
    for t in TABLES:
        ws = out.create_sheet(t["name"])
        ws.append([f[0] for f in t["fields"]])
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F4E78")
        for row in data[t["name"]]:
            ws.append(row)
        for j, f in enumerate(t["fields"], 1):
            L = get_column_letter(j)
            fmt = "@" if f[1] == "TEXT" else EXCEL_FMT.get(f[1])
            if fmt:
                for c in ws[L][1:]:
                    c.number_format = fmt
            ws.column_dimensions[L].width = 12 if f[1] != "TEXT" else max(10, min(30, f[2] + 4))
        ws.freeze_panes = "A2"
    out.save(dst)
    print("saved", dst)

    # 給 Java (Jackcess) 的 TSV：schema.tsv + 每表資料
    os.makedirs(tsv_dir, exist_ok=True)
    with open(os.path.join(tsv_dir, "schema.tsv"), "w", encoding="utf-8") as fh:
        for t in TABLES:
            for f in t["fields"]:
                name, typ, size, req, key, rule, rtext, fmt, desc = f[:9]
                fh.write("\t".join([t["name"], name, typ, str(size or 0), "1" if req else "0",
                                    key, rule, rtext, fmt, desc]) + "\n")
    for t in TABLES:
        with open(os.path.join(tsv_dir, t["name"] + ".tsv"), "w", encoding="utf-8") as fh:
            for row in data[t["name"]]:
                cells = []
                for v in row:
                    if isinstance(v, dt.datetime):
                        cells.append(v.strftime("%Y-%m-%d"))
                    elif isinstance(v, bool):
                        cells.append("true" if v else "false")
                    else:
                        cells.append(str(v))
                fh.write("\t".join(cells) + "\n")


if __name__ == "__main__":
    main(*sys.argv[1:4])
