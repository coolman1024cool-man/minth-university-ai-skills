---
name: alltop-purchasing-bot
description: 自動解析廠商報價單，生成 ALLTOP 系統專用 RPA 上機表，並透過 Selenium 一條龍全自動完成請購單與 A2 規格表。
---

# ALLTOP 採購全自動化機器人 (ALLTOP Purchasing RPA Standard)

本技能為敏實科技大學專屬之 **ALLTOP 請購單 (02015) 一條龍全自動化架構標準**。
旨在將廠商估價單（PDF 或圖片）自動解析轉化為採購資料庫，並透過專屬 Chrome 遠端除錯接管 (Selenium Port: 9222)，全自動完成主單建立、預算綁定與全部明細之輸入與儲存。

---

## 🛠️ 標準工作目錄架構 (Standard Workspace Structure)

目錄路徑：`D:\01邱正彥1130801-\01經費\06未結案\ALLTOP_採購自動化工具\`

```
ALLTOP_採購自動化工具/
├── 1_啟動專屬瀏覽器.bat          # [步驟1] 啟動遠端除錯接管模式之 Chrome (Port: 9222)
├── 2_執行採購全自動機器人.bat     # [步驟2] 一條龍全流程執行請購單自動填報 (Selenium)
├── 3_點我自動生成A2規格表.bat     # [步驟3] 需要 A2 規格表時一鍵自動排版產出 Word
├── run_purchasing_bot_selenium.py # 核心自動化主程式 (真實 DOM 深度整合)
├── 一鍵生成A2表單.py              # A2 規格表自動排版生成程式 (標楷體、2cm邊界)
├── 採購自動化專用表.xlsx          # 採購明細來源資料庫 (Excel)
├── A2規格表_Template.docx        # 國家公文標準 A2 樣板檔
└── A2產出結果/                   # 生成之 A2 規格表正式存放目錄
```

---

## 📋 ALLTOP 02015 系統真實 DOM 規格標準 (Ground Truth Standard)

經即時探測驗證，ALLTOP 02015 請購系統之底層標準結構如下，爾後所有採購自動化程式皆以此為唯一規範：

### 1. 主單畫面 (`F1115_formBuyData_1.php`)
- **主單用途說明**：`textarea[name='memo']` 或 `textarea`
  - 命名標準：一律簡潔有力，禁止過度冗長列舉全部食材。
  - 範例：`115年USR個案-子計畫4(何東洋)國際生在地食農體驗活動(11/21杭菊節)相關費用`
- **主單確認按鈕**：`button.save.btn-success`

### 2. 預算明細頁籤 (`F1115_formBuyData_2.php`)
- **頁籤切換**：`a[href*='F1115_formBuyData_2.php']`
- **新增明細按鈕**：`button.new.btn-info`（文字：`新增`，觸發 `F1115_formBuyData_21.php?cmd=new`）

### 3. 明細輸入表單 (`F1115_formBuyData_21.php`)
- **預算選擇按鈕**：`button.lightblue.btn-info`（觸發彈窗 `iframe_windows` 加載 `SelectBudget.php`）
  - 彈窗選取方式：切換至 `iframe_windows`，尋找包含指定預算代碼（如 `114-168-115C01-07`）之 `tr`，點擊該行之 `button.green`。
- **真實欄位 ID 與 Name**：
  - **品名**：`input#remark`（注意：絕非 goods_name）
  - **規格**：`textarea#standards`（若無特殊規範填入「無」）
  - **用途及說明**：`textarea#notes`
  - **計量單位**：`input#unit`
  - **數量**：`input#amount`（注意：絕非 qty）
  - **單價**：`input#price`
- **明細存檔確認按鈕**：`button.save.btn-success`（文字：`確認`）

---

## 🔄 標準作業流程 (SOP)

1. **資料萃取與專用表生成**：
   - 解析廠商單據，提取品名、數量、單位、單價與金額。
   - 規格無者填「無」，用途說明包含計畫名稱、主持人與用途。
   - 寫入 `採購自動化專用表.xlsx`，並在案件資料夾備份。

2. **全流程自動化執行**：
   - 執行 `1_啟動專屬瀏覽器.bat` 進入 02015 主單或明細頁面。
   - 執行 `2_執行採購全自動機器人.bat`，系統全自動完成主單案由、頁籤切換、預算選取及全部明細之輸入與儲存。

3. **A2 規格表（選配）**：
   - 若案情需要 A2 表，點擊 `3_點我自動生成A2規格表.bat` 自動產出符合標準之 Word 檔。
