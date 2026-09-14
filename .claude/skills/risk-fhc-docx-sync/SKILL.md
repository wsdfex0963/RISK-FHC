---
name: risk-fhc-docx-sync
description: 將「金控單位風險管理考核項目」docx 轉換為 markdown，或依據 Excel 追蹤表更新 docx 表格內容。
user_invocable: true
---

# 金控風險管理考核項目 docx ↔ md 同步工具

## 用途

處理「金控單位風險管理考核項目_數位科技處」相關文件的格式轉換與內容更新：

1. **docx → md**：讀取 docx 並產出乾淨的 markdown 表格
2. **Excel → docx**：依據 Excel 追蹤表的狀態（新增/已完成/已廢止/進行中），更新 docx 中的兩個表格

## 文件結構

docx 內含兩個表格：

| 表格 | 欄位 | 說明 |
|------|------|------|
| 表格一：本次新增 | NO / 規章名稱 / 風險控管機制 | 當年度全新訂定的規章 |
| 表格二：不定期修訂 | NO / 規章名稱 / 風險控管機制 | 既有持續維護的規章 |

## 操作流程

### 1. docx → md 轉換

```
步驟：
1. 用 pandoc 讀取 docx 內容
2. 整理成 markdown 表格格式（兩個表格分開）
3. 移除 pandoc 殘留的 {.underline} 標記
4. 輸出為 金控單位風險管理考核項目_{年份}_數位科技處.md
```

### 2. Excel 追蹤表 → docx 更新

根據 Excel 中每筆規章的「狀態」欄位決定操作：

| Excel 狀態 | 操作 |
|------------|------|
| 新增（不在原 docx 中的規章） | 加入表格一「本次新增」 |
| 已廢止 | 從表格二移除 |
| 已完成 / 進行中（既有規章修訂） | 保留在表格二「不定期修訂」 |

**判斷「新增」的邏輯：**
- Excel 中狀態明確標示「新增」→ 新增
- Excel 中出現，但 docx 兩個表格都沒有的規章名稱 → 新增
- Excel 中有，docx 表格二也有 → 維持在表格二（只是修訂）

### 3. docx 編輯技術步驟

使用 `/docx` skill 的 unzip → edit XML → zip 流程：

```bash
# 1. 解壓
unzip -q input.docx -d unpacked/
find unpacked -type l -delete
python3 scripts/merge_runs.py unpacked/

# 2. 用 Python ElementTree 操作 word/document.xml
#    - 找到兩個 w:tbl 元素（表格一 = 本次新增, 表格二 = 不定期修訂）
#    - 新增/移動/刪除 w:tr 列
#    - 重新編號各列的第一個 cell

# 3. 重新打包
(cd unpacked && zip -Xr ../output.docx . -q)

# 4. 驗證
python3 scripts/office/validate.py output.docx --original input.docx

# 5. 轉 PDF 目視確認
soffice --headless --convert-to pdf output.docx
pdftoppm -jpeg -r 150 output.pdf page
# 逐頁檢視 page-*.jpg
```

### 4. 新增列的 XML 格式規範

新增至表格的列必須遵循以下格式：

- **字型**：標楷體（eastAsia）、Century Gothic（ascii/hAnsi）
- **字色**：藍色 `#0000FF`（本次新增表格）、黑色（不定期修訂表格）
- **底線**：管理事項關鍵字使用 `<w:u w:val="single"/>`
- **對齊**：NO 欄置中、規章名稱置中、風險控管機制左右對齊（both）
- **欄寬**：NO=596dxa、規章名稱=2381dxa、風險控管機制=11643dxa
- **paraId**：隨機產生，必須 < 0x80000000
- **tcBorders 順序**：top → left → bottom → right（schema 強制）

### 5. 風險控管機制的撰寫風格

新增規章若無現成說明文字，依以下格式撰寫：

```
為[目的說明]，[依據法規（如有）]，就[管理事項1]、[管理事項2]、...及[管理事項N]等事宜訂定本[準則/要點/細則/辦法/政策]，以資遵循。
```

- 管理事項加底線標示
- 參考同類型既有規章的措辭風格
- 明確標註為「AI 草擬，請視需要調整」
