---
name: meeting-record-generator
description: 自動從對話截圖、會議白板照片、錄音檔或文字大綱中萃取重點，產出結構化會議紀錄，並匯出符合國家公文標準 (標楷體、上下2cm邊界) 的 Word 檔。
---

# 會議紀錄自動產製與排版 (Meeting Record Generator)

## 功能說明
這個技能用於將非結構化的會議資訊（如 LINE 對話截圖、白板照片、手寫筆記或錄音逐字稿）轉換為格式嚴謹、帶有具體 Action Items 的專業會議紀錄，並最終匯出為 Word 檔 (.docx)。

## 觸發時機
當使用者提供圖片或文字，並要求「產出會議紀錄」、「幫我寫會議紀錄」時，應自動啟用此技能。

## 執行步驟

### 第一步：資訊萃取與結構化 (生成 Markdown)
1. 分析使用者提供的素材，萃取以下核心要素：
   - 會議日期、地點、形式、參與人員。
   - 會議主旨 (Topic)。
   - 核心活動時程與規劃 (內容細節)。
   - 跨單位協調與待辦事項 (Action Items，需標註負責人)。
2. 產出一份 Markdown 格式的會議紀錄 (例如 `meeting_record.md`)。
3. Markdown 中待辦事項必須使用 `- [ ]` 格式。

### 第二步：自動排版與匯出 (呼叫腳本)
1. 在生成 Markdown 後，必須無縫自動執行本技能隨附的 Python 腳本 (`scripts/export_to_word.py`)，將該 Markdown 轉換為 Word 檔。
2. 預設匯出路徑：`E:\我的雲端硬碟\04_公文與重要表單\會議紀錄_[會議主旨].docx` (除非使用者另有指定)。
3. **強制格式限制** (已內建於腳本中)：
   - 全文強制使用 **繁體中文** 與 **標楷體** (DFKai-SB)。
   - 頁面邊界：**上、下各 2 公分**。
   - 標題字級適當放大，內文預設 12pt。

### 第三步：提交成果
1. 將輸出的 Word 檔案路徑與連結呈現給使用者。
2. 詢問使用者是否有遺漏或需要補充的細節。

---

## 工具與資源
* `scripts/export_to_word.py`：此 Python 腳本負責讀取 Markdown 並以 `python-docx` 產出完美排版的 Word 檔。執行指令範例：
  `python C:\Users\user\.gemini\config\skills\meeting-record-generator\scripts\export_to_word.py <md_path> <output_docx_path>`
