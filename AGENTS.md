# 敏實科技大學 USR 專案與智慧自動化工作規範 (AGENTS.md)

## 📌 專案簡介 (Project Overview)
本專案為 **敏實科技大學 (MINTH University of Science and Technology) USR 計畫與智慧自動化專屬 AI Agent 工作室**。旨在透過 AI 輔助，自動化處理 USR 計畫經費核銷彙整、結案報告重點萃取、公文與照片格式排版以及各項知識庫維護。

---

## 🛠️ 專案已配置個人化規則與技能庫 (Configured Customizations & Skills)

### 1. 個人化規則 (Persona Rules)
- **溝通語言**：繁體中文（臺灣習慣用語）。
- **角色定位**：資深軟體架構師、技術專家與 USR 計畫推動顧問。
- **排版標準**：
  - Word 公文與報表一律使用 **標楷體**，上下邊界各 2cm。
  - 經費報表符合國家公文規範（細黑線邊框、純黑白、數字欄位絕對不折行、A4 直式）。

### 2. 專案專屬技能庫 (Skills)
- **usr-web-curator**：敏實科大 USR 成果網站與多媒體數位策展機器人（跨資料夾檢索分類、排除簽到單/食材箱、720p 影音輕量轉碼、方案 A 播放清單、公文檢視器與 Firebase 熱部署一條龍）。
- **usr-budget-report-formatter**：會計系統 PDF 預算控制表轉 Word 統計表工具。
- **minth-poster-generator**：敏實科技大學專屬海報生成器 (支援成果與講座模式，產出 Canva 可編輯 PPTX 檔)。
- **meeting-record-generator**：專屬會議紀錄生成器，自動萃取會議重點並排版匯出國家公文標準格式 Word 檔。
- **activity-record-generator**：敏實科大官方標準「課程教學及活動記錄表」Word 檔生成工具（含校徽、裝訂線、成效紀要、剪影照片矩陣與附錄）。
- **work-study-data-compiler**：處理工讀金相關資料，自動補齊出生年月日及性別，產出最新 Excel 統計報表。
- **alltop-expense-bot**：自動解析收據發票照片，萃取核銷資訊寫入 ALLTOP 專用 Excel。
- **alltop-purchasing-bot**：自動解析廠商報價單，生成 ALLTOP RPA 上機表與 A2 規格表。
- **smart-calendar-assistant**：智慧行事曆小幫手，自動判讀截圖/文字萃取人事時地物並寫入 Google Calendar。
- **global-context-strategist**：全球局勢與地緣政治戰略顧問。
- **religious-studies-researcher**：歷史與宗教文物深度研究考究。
- **energy-expert**：全球能源發展與地質環境分析。

---

## 📂 目錄架構與外部知識庫 (Workspace Structure & Knowledge Base)

- 0000-antigravuty-0811/：主要 AI Agent 工作專案目錄
  - Obsidian_Vault/：**本機長效知識庫**（存放 Markdown 筆記、會議紀錄、論文摘要、評鑑分辦表）
  - .agents/rules/：個人化規則目錄
  - NotebookLM_Sources/：個人化知識庫來源目錄
- D:\01邱正彥1130801-\：主要業務資料庫目錄
  - 01經費/07預算核銷/：歷年經費與預算 PDF 檔案
  - 98計畫申請書/：USR 第三/四期計畫書研發檔案
  - 101生圖生成/：AI 生成海報與簡報資源

---

## 🔄 開工與收工標準作業程序 (Workflow SOP)

### 🌅 開工標準流程 (Start Session)
1. 檢視 git status 與近期 commit 記錄。
2. 檢查 Obsidian_Vault/ 知識庫狀態與待辦清單。
3. 確認本次 Session 的單一聚焦任務，確保 Context 保持清晰高效。

### 🌇 收工標準流程 (Finish Session)
1. **成果沉澱**：重要結論、會議紀錄或分析大綱寫入 Obsidian_Vault/。
2. **安全防護**：檢查敏感資料（API key、token、個資），確認 .gitignore 正常生效。
3. **版本控制**：檢查 git status + diff，只 stage 本次相關檔案，自動 commit 並 push 至 GitHub。
4. **Context 重置提醒**：若本次對話較長，主動提醒使用者於下次開工時建立新 Session，以確保最佳 AI 效能。

---

## 📋 全域規範提醒
- **臨時草稿規範**：所有暫時性 Python 測試腳本或草稿，皆必須寫入 D:\AI_Python_草稿\ 目錄中，禁止直接散落在 D 槽或桌面。
- **檔案產出路徑與命名規範**：
  - INBOX 任務：產出存入 E:\我的雲端硬碟\OUTBOX_已處理\，來源移入歸檔資料夾。
  - 一般產出：存入 E:\我的雲端硬碟\Gemini_Spark_產出檔案\。
  - 檔名規範：統一加上民國年月日時分（11碼）與 -anti 標記（如：名稱-11508311013-anti.xlsx）。