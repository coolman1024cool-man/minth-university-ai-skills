---
title: AI Agent 基本功 EP02：學習 Agent 必懂的核心觀念與初始化設定
speaker: 三師爸 Sense Bar
url: https://www.youtube.com/watch?v=8nwjYouFJoE
date: 2026-09-09
tags:
  - AIAgent
  - 三師爸
  - 學習筆記
  - 初始化設定
  - Context管理
  - AntiGravity
---

# 🤖 AI Agent 基本功 EP02：學習 Agent 必懂的核心觀念與初始化設定

> **影片來源**：[三師爸 Sense Bar YouTube 頻道](https://www.youtube.com/watch?v=8nwjYouFJoE)
> **相關專案**：[[三師爸教我學AGENT]] ｜ [[AntiGravity]]

---

## 📺 影片播放

<iframe width="100%" height="400" src="https://www.youtube.com/embed/8nwjYouFJoE" title="AI Agent 基本功EP02" frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share" allowfullscreen></iframe>

---

## 💡 本集核心觀念總覽

本集主要解決初學者在剛接觸 AI Agent 時最常遇到的**「卡關」與「越用越笨」**的根本原因。無論是使用 [[Claude Code]]、[[AntiGravity]]、Codex 或 OpenCode，底層邏輯皆通用。

`mermaid
mindmap
  root((AI Agent 必備基本功))
    權限設定
      半自動模式
      Bypass 模式（老手）
      邊界與安全防護
    記憶與初始化
      全域規範 Global AGENTS.md
      專案規範 Project AGENTS.md
      開工與收工規範 (Git / .gitignore)
    上下文管理 Context
      破解「越聊越笨」迷思
      Context 記憶體滿載原理
      對話壓縮與多 Session 分工
    用量與模型分級
      5小時與週額度管理
      S/A/B 級模型選用策略
      性價比 CP 值最佳化
`

---

## 🔑 四大核心模組深入解析

### 一、權限設定（Permissions & Security）
* **半自動 vs Bypass 授權**：
  * **新手建議**：先從「詢問授權（Approval）」開始，熟悉 Agent 在背後呼叫哪些工具（如終端機指令、檔案修改）。
  * **老手進階**：設定受信任路徑與 Bypass 權限，讓 Agent 能流暢自動化執行 multi-step 任務。
* **防護原則**：透過環境隔離與 .gitignore 避免敏感金鑰（API Keys、密碼）被 Agent 讀取或誤傳。

---

### 二、記憶與專案初始化（Memory & Initialization）
* **雙層規範架構**：
  1. [[全域規範 (Global AGENTS.md)]]：設定個人習慣（如：繁體中文、架構師角色、檔案命名時戳與輸出目錄）。
  2. [[專案規範 (Project AGENTS.md)]]：針對單一專案定義目標、檔案結構、技能庫（Skills）與工作里程碑。
* **開工與收工標準作業程序（SOP）**：
  * **開工（Start）**：檢查 git status、讀取規則、盤點現狀。
  * **收工（Finish）**：排除敏感資料、更新專案筆記、驗證與乾淨 commit。

---

### 三、上下文管理（Context Management）⭐ 最關鍵核心！
* ❓ **為什麼 AI 常常「越聊越笨」？**
  * **真相**：不是模型變笨，而是 **Context Window（上下文記憶體）被塞滿**！
  * 當一個對話進行太久，累積了過多的工具輸出、除錯歷程與無效文字，導致模型在龐大雜訊中「注意力被稀釋（Attention Drift）」。
* 🛠️ **最佳化解決方案**：
  1. **監控用量**：隨時留意對話長度與 Token 消耗。
  2. **主動重置 / 初始化對話**：任務階段告一段落時，及時新開對話（New Session），讓 AI 隨時保持在最聰明的清醒狀態。
  3. **分工原則（Single Responsibility）**：一個對話專心處理一個子任務（例如：研究歸研究、寫扣歸寫扣）。

---

### 四、技能（Skills）與模型分級策略（Model Tiering）
* **額度與頻率管理**：認識 5 小時額度與每週限制，避免在簡單工作上過度消耗高級模型額度。
* **S / A / B 級模型合理搭配**：
  * **S 級（高智慧/深度思考）**：架構設計、複雜邏輯除錯、重大決策。
  * **A / B 級（輕量/極速模型）**：資料抓取、檔案格式轉換、快速搜尋、簡單文字整理。

---

## 🎯 實踐行動檢核清單（Action Items）

- [ ] 在專案中落實 AGENTS.md 規則檔案，明確定義專案背景與輸出規範。
- [ ] 養成定期開啟新 Session 的習慣，避免單一對話過長導致 Context 溢滿。
- [ ] 在 D:\01邱正彥1130801-\0000-antigravuty-0811\ 專案中實踐開工與收工 SOP。
- [ ] 配合 [[Obsidian]] 作為本機長期記憶庫，將每次對話沉澱的成果結構化歸檔。