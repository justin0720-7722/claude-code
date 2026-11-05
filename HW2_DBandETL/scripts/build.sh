#!/usr/bin/env bash
# 一鍵重建 HW2 全部檔案：工作檔(公式) → LibreOffice 重新計算 → 匯入檔(值) → Access .accdb → zip
# 需要：python3 + openpyxl、LibreOffice、Java 17+、Maven（第一次會下載 Jackcess）
set -euo pipefail
export LANG=C.UTF-8 LC_ALL=C.UTF-8   # Java 需要 UTF-8 locale 才能處理中文檔名
HERE="$(cd "$(dirname "$0")" && pwd)"
SID="$(python3 -c "import sys; sys.path.insert(0, '$HERE'); from spec import STUDENT_ID; print(STUDENT_ID)")"
OUT="$HERE/../submission"
WORK="$(mktemp -d)"
rm -rf "$OUT" && mkdir -p "$OUT"

WORKBOOK="HW2_假資料工作檔_${SID}.xlsx"
python3 "$HERE/build_excel.py" "$WORK/$WORKBOOK"
soffice --headless --calc --convert-to xlsx --outdir "$WORK/calc" "$WORK/$WORKBOOK" >/dev/null 2>&1
cp "$WORK/calc/$WORKBOOK" "$OUT/$WORKBOOK"
python3 "$HERE/export_import.py" "$OUT/$WORKBOOK" "$OUT/HW2_匯入資料_${SID}.xlsx" "$WORK/tsv"

LIB="$HERE/.lib"
if [ ! -f "$LIB/jackcess-4.0.5.jar" ]; then
  mkdir -p "$LIB"
  printf '%s' '<project xmlns="http://maven.apache.org/POM/4.0.0"><modelVersion>4.0.0</modelVersion><groupId>hw</groupId><artifactId>hw2</artifactId><version>1</version><dependencies><dependency><groupId>com.healthmarketscience.jackcess</groupId><artifactId>jackcess</artifactId><version>4.0.5</version></dependency></dependencies></project>' > "$WORK/pom.xml"
  mvn -q -f "$WORK/pom.xml" dependency:copy-dependencies -DoutputDirectory="$LIB"
fi
javac -encoding UTF-8 -cp "$LIB/*" -d "$WORK/cls" "$HERE/BuildAccdb.java"
java -Dfile.encoding=UTF-8 -cp "$LIB/*:$WORK/cls" BuildAccdb "$WORK/tsv" "$OUT/HW2_社團活動資料庫_${SID}.accdb"

# 附上 AI 生成說明 / Prompt / 程式，並打包成繳交用 zip
cp "$HERE/AI生成說明與Prompt.md" "$OUT/"
mkdir -p "$OUT/AI生成內容"
cp "$HERE"/spec.py "$HERE"/build_excel.py "$HERE"/export_import.py "$HERE"/BuildAccdb.java "$HERE"/build.sh "$OUT/AI生成內容/"
rm -f "$HERE"/../HW2_DBandETL_*.zip
ZIP="$HERE/../HW2_DBandETL_${SID}.zip"
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
rm -rf "$WORK"
