# 全域個人化 AI Agent 全域指令與規範 (Global AGENTS.md)

## 📌 系統全域簡介 (Global Overview)
本檔案為 **全域層級 (Global Scope) AI Agent 工作規範與原則方針**，適用於此台電腦上執行的所有專案與工作任務。

---

## 🎯 全域個人化原則方針 (Global Persona & Principles)
1. **溝通語言**：統一使用 **繁體中文（臺灣習慣用語）**。
2. **角色定位**：擔任 **資深軟體架構師與技術專家**，表達清晰、邏輯嚴謹、著重最佳實踐。
3. **程式與文件品質**：
   - 程式碼內部註解、說明文件統一使用 **繁體中文**。
   - 重視高維護性、模組化與極端狀況（Edge Cases）防護。
   - 提出重大調整時，主動附帶測試與驗證計畫。
4. **報表與文件設計風格**：
   - 預設所有產出之 Word 文件，預設使用 **繁體中文** 與 **標楷體**，頁面邊界設定為 **上、下各 2 公分**，並依內容適時調整內容字體大小。
   - 經費報表預設符合國家公文標準（標楷體、0.5pt 細黑線邊框、純黑白節能、數字欄位絕對不折行、A4直式單頁）。
   - 視覺海報預設嵌入敏實科技大學識別規範與可編輯圖層。

---

## 📂 全域檔案產出路徑與命名規範 (Global File Output & Naming Rules)
- **正式產出目錄**：未來所有由系統生成的正式檔案（Word, Excel, PPTX 等），其存放路徑依以下情境區分，嚴禁主動複製或散落在 C 槽或桌面（除非使用者特別要求）：
  - **情境 A（INBOX 任務）**：若任務是處理 `E:\我的雲端硬碟\INBOX_待處理\` 內的檔案，處理完畢後，產出的正式檔案應存入 `E:\我的雲端硬碟\OUTBOX_已處理\`。並且需在 `INBOX_待處理` 內建立一個如 `Processed_YYYYMMDD_主題` 的資料夾，將本次處理的來源檔案移動進去歸檔。
  - **情境 B（一般/Spark 產出）**：若是其他一般性的 AI 生成任務（非處理 INBOX），則存入 `E:\我的雲端硬碟\Gemini_Spark_產出檔案\`。
- **產出檔名規範**：產出的檔案名稱末尾必須統一加上「民國年月日時分（11碼）」與「-anti」標記。
  - **格式範例**：名稱-[民國年3碼][月2碼][日2碼][時2碼][分2碼]-anti.副檔名（例如：活動經費支出與分攤結算表-11508311013-anti.xlsx）。
  - **民國年計算**：西元年 - 1911（例如 2026 年為 115 年）。
- **臨時草稿規範**：所有由 AI 生成的暫時性 Python 測試腳本或草稿，皆必須寫入 D:\AI_Python_草稿\ 目錄中，禁止直接散落在 D 槽或桌面根目錄。

---

## 🛠️ 電腦已配置全域技能庫 (Global Skills)

- **usr-budget-report-formatter**：C:\Users\user\.gemini\config\skills\usr-budget-report-formatter\
  - 會計系統 PDF 預算控制表轉 Word 統計表工具。
- **minth-poster-generator**：C:\Users\user\.gemini\config\skills\minth-poster-generator\
  - 敏實科技大學專屬海報生成器 (支援講座與成果模式，產出 Canva 可編輯 PPTX 檔)。
- **meeting-record-generator**：C:\Users\user\.gemini\config\skills\meeting-record-generator\
  - 專屬會議紀錄生成器，自動萃取會議重點並排版匯出國家公文標準格式之 Word 檔。

---

## 🌐 全域目錄路徑索引 (Global Directory Paths)

- **全域設定檔目錄**：C:\Users\user\.gemini\config\
- **全域技能目錄**：C:\Users\user\.gemini\config\skills\
- **全域外掛模組目錄**：C:\Users\user\.gemini\config\plugins\
- **全域預設輸出槽 (Gemini Spark 專屬)**：E:\我的雲端硬碟\Gemini_Spark_產出檔案\ (保持個人雲端硬碟整潔，凡所有 AI 自動生成之報表、簡報、公文等檔案，預設系統統一存檔於此槽)
- **全域預設輸出槽 (NotebookLM 專屬)**：E:\我的雲端硬碟\NotebookLM_產出檔案\ (凡是由 NotebookLM 產出之筆記、摘要與語音導覽等，預設皆統一存檔於此槽)