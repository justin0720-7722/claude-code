#!/usr/bin/env bash
# 一鍵重建 HW2 全部檔案：工作檔(公式) → LibreOffice 重新計算 → 匯入檔(值) → Access .accdb
# 需要：python3 + openpyxl、LibreOffice、Java 17+、Maven（下載 Jackcess）
set -euo pipefail
export LANG=C.UTF-8 LC_ALL=C.UTF-8   # Java 需要 UTF-8 locale 才能處理中文檔名
HERE="$(cd "$(dirname "$0")" && pwd)"
OUT="$HERE/../submission"
WORK="$(mktemp -d)"
mkdir -p "$OUT"

python3 "$HERE/build_excel.py" "$WORK/HW2_假資料工作檔.xlsx"
soffice --headless --calc --convert-to xlsx --outdir "$WORK/calc" "$WORK/HW2_假資料工作檔.xlsx" >/dev/null 2>&1
cp "$WORK/calc/HW2_假資料工作檔.xlsx" "$OUT/HW2_假資料工作檔.xlsx"
python3 "$HERE/export_import.py" "$OUT/HW2_假資料工作檔.xlsx" "$OUT/HW2_匯入資料.xlsx" "$WORK/tsv"

LIB="$HERE/.lib"
if [ ! -f "$LIB/jackcess-4.0.5.jar" ]; then
  mkdir -p "$LIB"
  printf '%s' '<project xmlns="http://maven.apache.org/POM/4.0.0"><modelVersion>4.0.0</modelVersion><groupId>hw</groupId><artifactId>hw2</artifactId><version>1</version><dependencies><dependency><groupId>com.healthmarketscience.jackcess</groupId><artifactId>jackcess</artifactId><version>4.0.5</version></dependency></dependencies></project>' > "$WORK/pom.xml"
  mvn -q -f "$WORK/pom.xml" dependency:copy-dependencies -DoutputDirectory="$LIB"
fi
javac -encoding UTF-8 -cp "$LIB/*" -d "$WORK/cls" "$HERE/BuildAccdb.java"
java -Dfile.encoding=UTF-8 -cp "$LIB/*:$WORK/cls" BuildAccdb "$WORK/tsv" "$OUT/HW2_線上書店資料庫.accdb"
rm -rf "$WORK"

# 附上 AI 生成的程式，並打包成繳交用 zip（請把「學號」改成自己的學號）
mkdir -p "$OUT/AI生成內容"
cp "$HERE"/spec.py "$HERE"/build_excel.py "$HERE"/export_import.py "$HERE"/BuildAccdb.java "$HERE"/build.sh "$OUT/AI生成內容/"
ZIP="$HERE/../HW2_DBandETL_${STUDENT_ID:-學號}.zip"
python3 - "$OUT" "$ZIP" <<'PY'
import os, sys, zipfile
src, dst = sys.argv[1], sys.argv[2]
root = os.path.splitext(os.path.basename(dst))[0]
with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
    for d, _, files in os.walk(src):
        for f in sorted(files):
            p = os.path.join(d, f)
            z.write(p, os.path.join(root, os.path.relpath(p, src)))
print("saved", dst)
PY
