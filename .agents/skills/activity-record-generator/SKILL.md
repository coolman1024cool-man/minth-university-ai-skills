---
name: activity-record-generator
description: >-
  敏實科技大學 USR 計畫專屬「課程教學及活動記錄表」自動生成技能。
  嚴格遵循學校官方排版規範（頁首校徽 Logo、酒紅色分隔線、左側裝訂線、標楷體、A4直式版面）。
  具備智慧防呆與照片過濾機制（嚴格排除存摺、收據、發票、飲料便當等核銷雜物，活動剪影僅收錄現場活動照片），
  並確保附錄「簽到表」為專屬簽到單掃描檔（絕不隨意拿活動照片濫竽充數）。
  支援自動填寫子計畫核選、基本資訊、量化與質化執行成效、結構化活動紀要、檢討建議，
  並將 2~6 張真實活動照片等比例縮放居中排入活動剪影矩陣。
---

# 敏實科技大學 USR 活動記錄表生成技能 (Activity Record Generator)

## 📌 簡介 (Overview)
本技能用於自動化產出符合 **敏實科技大學 (MINTH University) 大學社會責任實踐計畫中心** 官方規範之 **「課程教學及活動記錄表」Word (.docx)** 檔案。

技能嚴格遵循以下官方標準與防呆原則：
1. **官方識別與版面結構**：內建官方頁首（敏實科技大學校徽高解析 Logo、中心抬頭、課程教學及活動記錄表）、酒紅色裝飾實線（Maroon `#800000`）以及左側邊界垂直「裝 訂」標記。
2. **標準邊界與字型**：A4 直式規格，頁面邊界上 1.5cm、下 1.5cm、左 2.0cm、右 2.0cm。全文統一採用標準 **標楷體**（英數字為 Times New Roman / Calibri）。
3. **活動剪影智慧過濾（排除核銷雜圖）**：
   - **嚴格聚焦現場照片**：僅收錄活動過程當下的照片（如：講師講授、同仁聽講互動、分組實作討論、頒獎或合影）。
   - **自動防呆排除**：自動過濾存摺封面/帳號、發票、收據、轉帳截圖、便當、飲料杯、點心等核銷佐證雜圖，絕不讓雜物混入活動剪影。
   - **等比例居中縮放與 EXIF 旋轉校正**：防拉伸變形，圖說統一置於照片正下方並以 `說明：` 前綴呈現。
4. **專用簽到簿掃描檔防呆**：
   - 附錄頁之「簽到表」必須為真實之 **活動簽到簿/簽到單掃描檔**（含有學員/出席人員手寫簽名與簽到欄位）。
   - **嚴禁隨意拿活動照片替代簽到表**。若未提供簽到表掃描檔，系統將主動發出提醒並省略該附錄頁，絕不誤植無關照片。
5. **兩大核心表格結構**：
   - **表格一（基本資訊與執行紀要）**：自動依子計畫代碼核選子計畫方塊、填入主持人與工讀生、活動名稱、時間地點、講者與會人、量化參與人次與問卷統計、質化指標、四段式結構化活動紀要、檢討與建議。
   - **表格二（活動剪影與佐證附件）**：2～6 張照片矩陣排版，末列自動勾選議程、簽到單、海報、簡報等附件。
6. **全域存檔與命名標準**：檔名尾端自動附加「民國年月日時分（11碼）」與 `-anti.docx`，預設自動存入 `E:\我的雲端硬碟\Gemini_Spark_產出檔案\`。

---

## 🛠️ 快速使用 (Quick Start)

### 方式 A：透過 CLI 指令生成
備妥活動資料 JSON 檔案後，在命令列執行：

```bash
python .agents/skills/activity-record-generator/scripts/generate_record.py --data "路徑/至/活動資料.json"
```

或使用全域路徑：
```bash
python C:\Users\user\.gemini\config\skills\activity-record-generator\scripts\generate_record.py --data "路徑/至/活動資料.json"
```

---

### 方式 B：Python 程式模組呼叫

```python
from generate_record import generate_activity_record

data = {
    "year": "115",
    "subproject_id": "5",
    "host_name": "黃瓊華",
    "assistant_name": "林芷暄",
    "activity_name": "國家品質獎：全面卓越類參獎準備重點_卓越經營 × TQM × 大學治理",
    "date_time": "115/3/10 13:30~16:30(星期二)",
    "location": "綜一館1樓電腦教室",
    "speaker_attendees": "吳英志 委員 / 子計畫5黃瓊華老師及本校教職員生共21員",
    "quantitative": {
        "inner_teacher": 21,
        "inner_student": 0,
        "inner_assistant": 0,
        "cross_teacher": 0,
        "cross_student": 0,
        "cross_assistant": 0,
        "field_partner": 0,
        "central_gov": 0,
        "local_gov": 0,
        "ngo_npo": 0,
        "survey_sent": 21,
        "survey_valid": 21
    },
    "qualitative_summary": "協助學校了解國家品質獎全面卓越類評審重點與參獎策略，找出組織改善機會點。",
    "activity_summary": "一、問題意識與目標\n...\n二、執行重點與方法\n...",
    "review_suggestions": "後續將積極推動校內各行政與教學單位導入自我評估機制，持續完善參獎準備工作。",
    # 活動剪影：僅放真實活動過程照片
    "photos": [
        {"path": "D:/photos/01_講師講授.jpg", "caption": "吳委員介紹全面卓越類參獎重點"},
        {"path": "D:/photos/02_同仁參與.jpg", "caption": "全校各處室派員參與品質獎講座"}
    ],
    "attachments": ["簽到單", "宣傳海報", "簡報"],
    # 專用簽到簿掃描檔 (嚴禁拿活動照片替代)
    "signin_sheet": "D:/photos/活動簽到表_掃描檔.png"
}

# 生成檔案並自動存入 Gemini_Spark_產出檔案
output_docx = generate_activity_record(data)
```

---

## 📋 資料欄位規範 (Data Specification)

| 欄位名稱 | 型態 | 說明與規範 |
| :--- | :--- | :--- |
| `year` | string | 計畫年度（如 `"115"`） |
| `subproject_id` | string / int | 子計畫編號：`"1"`, `"2"`, `"3"`, `"4"`, `"5"` 或 `"hub"` |
| `host_name` | string | 子計畫主持人姓名 |
| `assistant_name` | string | 工讀生 / 助理姓名 |
| `activity_name` | string | 活動名稱（含主副標題） |
| `date_time` | string | 日期與時段（如 `"115/3/10 13:30~16:30(星期二)"`） |
| `location` | string | 活動舉辦地點（如 `"綜一館1樓電腦教室"`） |
| `speaker_attendees` | string | 主講者與參與人員說明 |
| `quantitative` | dict | 量化指標（教師、學生、助理、跨校、場域夥伴、政府/NGO人次、問卷數） |
| `qualitative_summary`| string | 質化指標說明（成效影響與回饋） |
| `activity_summary` | string | 活動紀要（建議依問題意識、執行方法、成果亮點、社會效益撰寫） |
| `review_suggestions`| string | 檢討與建議（精進策略與後續行動） |
| `photos` | list[dict] | **活動過程現場照片陣列**（2～6張，每項含 `path` 與 `caption`，自動過濾存摺/發票/點心雜物） |
| `attachments` | list[str] | 勾選附件清單（可選：`"議程"`, `"簽到單"`, `"宣傳海報"`, `"簡報"` 或其他） |
| `signin_sheet` | string | **活動專屬簽到簿掃描檔路徑**（必為簽到名冊圖檔，嚴禁拿活動照片替代） |
| `annex_files` | list[dict] | 其他全頁掃描附件（如海報、簡報，每項含 `title` 與 `path`） |

---

## 📂 檔案目錄結構

```text
activity-record-generator/
├── SKILL.md                          # 技能說明與使用規範
├── templates/
│   └── template_activity_record.docx # 官方標準 Word 範本 (含 Logo、裝訂線、樣式)
├── scripts/
│   └── generate_record.py            # 核心生成 Python 腳本 (含智慧過濾與簽到簿防呆)
└── examples/
    ├── sample_input.json             # 完整測試資料範例
    └── sample_photos/                # 真實活動現場照片與簽到簿範例
```
