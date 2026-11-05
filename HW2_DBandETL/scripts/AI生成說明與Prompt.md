# AI 生成說明與 Prompt 提示語

學號：41411017E

作業規定：「如果是由 AI 產生的，請提供 AI 生成的內容檔案，以及所產製的 Prompt 提示語」。

## 1. 使用的 AI
- 工具：Claude Code（Anthropic）
- 用途：依照我的 HW1 草圖（`HW1_TableSchemaSketch_41411017E.xlsx`）規劃 HW2 的欄位規格與資料範圍、撰寫 Excel 亂數公式、產生工作檔與匯入檔、建立 Access 資料庫

## 2. Prompt 提示語（依對話順序）

**Prompt 1：貼上 HW2 作業說明全文**

> HW2_建置資料庫與匯入測試資料
> 將你的夢想實現吧! 請依據你HW1的規畫草圖，來完成你的Access資料庫建置吧
> 並在逐個設計資料表及資料欄位中，去調整與修正你的設計，使它更便捷、更符合操作需求
> 並依據你HW1的規畫草圖，設計你要準備的資料，包含他們大概要長什麼樣子
> 需要什麼格式、長度、數據狀態、範圍區間，並透過Excel先將資料做成我們預期的樣子
> 在把整理好的資料透過匯入的工具轉入MS Access中
> 本次作業包含兩部分：
> 1. 依據你的資料庫欄位規劃，用Excel打造一批假資料(測試驗證資料)
> 2. 用這包假資料匯入成果的Access資料庫
> （……每個須轉入的資料表，必須產生至少30筆的資料筆數，再將這批資料透過匯入方法轉入MS Access中）
> 繳交標的：1. MS Access(.accdb)　2. Excel檔(.xlsx) 亂數生成工作檔（各欄位包含產製語法指令）
> 3. 整理後要匯入的資料檔（.xlsx / .xls / .csv / .xml）

**Prompt 2：**

> 有需要我操作的地方嗎

**Prompt 3：上傳 HW1 草圖**

> （附檔）HW1_TableSchemaSketch_41411017E_.zip

## 3. AI 生成的內容

| 檔案 | 內容 | 如何產生 |
|---|---|---|
| `HW2_假資料工作檔_41411017E.xlsx` | 四張表的每一格都是公式（RANDBETWEEN、RAND、INDEX、MATCH、LARGE、COUNTIF、TEXT、IF、CEILING…），另有「說明」「欄位規格」「參數」「檢核」頁 | AI 撰寫公式後由程式寫入，再用 LibreOffice 計算一次 |
| `HW2_匯入資料_41411017E.xlsx` | 工作檔的「值」（等同 複製 → 選擇性貼上值），只含資料欄，一張表一個工作表 | 由工作檔產製 |
| `HW2_社團活動資料庫_41411017E.accdb` | 4 張資料表：欄位型別、大小、標題、必填、驗證規則、格式、查閱下拉選單；主鍵、索引、複合唯一索引；4 條關聯（強制參考完整性）；已匯入資料 | 由匯入檔匯入 |
| `AI生成內容/` | 產生上面三個檔案的程式：`spec.py`（欄位規格）、`build_excel.py`、`export_import.py`、`BuildAccdb.java`、`build.sh` | AI 撰寫 |

## 4. 資料筆數

| 資料表 | 筆數 | 主鍵 |
|---|---|---|
| StudentTable（學生資料） | 60 | StudentID，例如 41411017E 格式 |
| ClubTable（社團資料） | 30 | ClubID 1001～1030（自動編號） |
| ActivityTable（活動資料） | 60 | ActivityID 2001～2060（自動編號） |
| RegistrationTable（活動報名） | 240 | RegistrationID 3001～3240（自動編號） |

學生姓名、電話、Email 皆為隨機組合的虛構資料（Email 使用保留網域 example.edu.tw），與真實人物無關。
