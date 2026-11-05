# HW2 建置資料庫與匯入測試資料 — 校園社團活動管理系統（41411017E）

依 `HW1_TableSchemaSketch_41411017E.xlsx` 的四張表設計。

## 資料夾內容

```
HW2_DBandETL/
├── HW2_DBandETL_41411017E.zip                 ← 繳交用
├── submission/                                ← zip 的內容
│   ├── HW2_社團活動資料庫_41411017E.accdb      ① Access 資料庫（已建表、已匯入）
│   ├── HW2_假資料工作檔_41411017E.xlsx         ② 亂數生成工作檔（每格都是公式）
│   ├── HW2_匯入資料_41411017E.xlsx             ③ 匯入用資料檔（純值）
│   ├── AI生成說明與Prompt.md
│   └── AI生成內容/                             產生以上檔案的程式
└── scripts/                                   同上程式（build.sh 一鍵重建）
```

## 資料庫設計

```
StudentTable 1 ─┬─ ∞ ClubTable.PresidentStudentID（唯一：一人只當一個社長）
                │        ClubTable 1 ─ ∞ ActivityTable
                │                         ActivityTable 1 ─┐
                └──────────────────────────────────────────┴─ ∞ RegistrationTable
                                                              （ActivityID + StudentID 複合唯一）
```

| 資料表 | 筆數 |
|---|---|
| StudentTable | 60 |
| ClubTable | 30 |
| ActivityTable | 60 |
| RegistrationTable | 240 |

完整欄位規格見工作檔的「欄位規格」頁；對 HW1 草圖做的 14 項調整見「說明」頁第四節。

## 自己在 Access 裡做一次匯入（建議練習）

1. 開啟 `.accdb`，依「RegistrationTable → ActivityTable → ClubTable → StudentTable」順序清空資料（先子後父）
2. 外部資料 → 新增資料來源 → 從檔案 → Excel → 選 `HW2_匯入資料_41411017E.xlsx`
3. 選「在資料表中附加記錄的複本」→ 選同名資料表 → 選同名工作表 →「第一列是欄名」→ 完成
4. 依「StudentTable → ClubTable → ActivityTable → RegistrationTable」順序匯入（先父後子）

## 重新產生亂數資料

- Excel：開啟工作檔按 F9 → 確認「檢核」頁合計為 0 → 各表資料欄（不含灰色輔助欄）複製 → 選擇性貼上「值」到匯入檔
- 或執行 `scripts/build.sh`（需 python3 + openpyxl、LibreOffice、Java 17+、Maven）
