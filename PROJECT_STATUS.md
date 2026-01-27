# PROJECT_STATUS.md

**Generated Date:** 2026-01-27  
**Document Type:** Project Status & Handoff Guide  
**Purpose:** AI / New Developer Handoff & Safe Modification Guide

---

## 【1】專案性質與目的（Project Nature & Purpose）

### 專案類型
**PDF 試題解析與試卷生成工具（Web Application）**

這是一個純前端的 Web 應用程式，主要功能是：
- 從 PDF 檔案中解析選擇題（Multiple Choice Questions）
- 支援多檔案上傳（每個 PDF 對應一個章節）
- 為每個 PDF 設定要選擇的題目數量
- 隨機抽題並生成 Word 格式的試卷（題目卷與答案卷）

### 判斷依據
- **檔案證據：** `index.html` 標題為 "Managerial Accounting Test Generator"
- **程式證據：** `app.js` 包含 `PDFParser`、`QuestionGenerator`、`WordGenerator` 類別
- **UI 證據：** 介面包含「上傳 PDF」、「設定出題規則」、「解析 PDF」、「生成試卷」等步驟

### 解決的問題
- 從 PDF 格式的題庫中自動提取選擇題
- 支援按章節（檔案）分配題目數量
- 自動生成 Word 格式的試卷與答案卷，便於列印與分發

### 預期使用者
**教師或教學人員**（根據介面語言為繁體中文，以及「考卷名稱」、「生成試卷」等用語判斷）

---

## 【2】實際使用流程（Actual User Flow）

### 步驟 1：選擇科目
- 使用者從下拉選單選擇科目（Managerial Accounting 或 Financial Accounting）
- 選擇後，頁面標題與考卷名稱預設值會自動更新

### 步驟 2：上傳 PDF 檔案
- 點擊上傳區域或拖放 PDF 檔案
- 可一次上傳多個 PDF（每個 PDF 對應一個章節）
- 上傳後，系統會自動開始解析（延遲 500ms）

### 步驟 3：設定出題規則
- 輸入考卷名稱（預設為目前科目的預設名稱）
- 在檔案列表中，為每個 PDF 設定要選擇的題目數量
- 系統會顯示每個 PDF 可用的題目總數

### 步驟 4：解析 PDF（自動執行）
- 上傳檔案後自動觸發解析
- 解析狀態會顯示在「步驟 3：解析 PDF」區塊
- 解析完成後顯示：
  - 總題數
  - 檔案數量
  - 每個檔案的題目數量

### 步驟 5：生成試卷
- 點擊「Generate & Download DOCX」按鈕
- 系統會：
  1. 從每個 PDF 中隨機選擇指定數量的題目
  2. 打亂所有選中題目的順序
  3. 重新編號（1, 2, 3...）
  4. 生成兩個 Word 檔案：
     - `[考卷名稱] - Questions.docx`（題目卷，學生用）
     - `[考卷名稱] - Answers.docx`（答案卷，教師用）

### 流程限制與注意事項
- 解析失敗時，會顯示錯誤訊息，但不會自動重試
- 若未為任何檔案設定題目數（或全部為 0），無法生成試卷
- 題目數量不能超過該 PDF 的可用題目數
- 檔案名稱會自動清理 Windows 不允許的字元（`< > : " / \ | ? *`）

---

## 【3】整體架構總覽（High-level Architecture）

### 架構類型
**純前端應用程式（Client-side Only）**
- 無後端伺服器
- 無 API 呼叫
- 無雲端服務
- 所有處理都在瀏覽器中完成

### 主要技術與 Library

#### 第三方套件（CDN 載入）
| 套件名稱 | 版本 | 載入位置 | 用途 |
|---------|------|---------|------|
| pdf.js | 3.11.174 | `index.html:58` | PDF 解析 |
| pdf.js worker | 3.11.174 | `index.html:61` | PDF.js 背景處理 |
| docx | 8.5.0 | `index.html:63` | Word 文檔生成 |
| FileSaver.js | 2.0.5 | `index.html:64` | 檔案下載 |

#### 核心技術
- **HTML5**：檔案上傳、拖放 API
- **JavaScript (ES6+)**：類別、async/await、箭頭函數
- **CSS3**：漸層背景、動畫、響應式設計

### 資料流動方式

```
使用者上傳 PDF
    ↓
PDFParser.parsePDFs()
    ↓
解析每個 PDF → 提取題目 → 儲存到 parsedQuestionsByFile[]
    ↓
使用者設定題目數量（每個 PDF）
    ↓
QuestionGenerator.shuffle() → 從每個 PDF 隨機選擇指定數量
    ↓
WordGenerator.generateQuestionSheet() → 生成題目卷
WordGenerator.generateAnswerSheet() → 生成答案卷
    ↓
FileSaver.js → 下載兩個 .docx 檔案
```

### 資料儲存
- **記憶體儲存**：所有資料（PDF 檔案、解析結果）都儲存在 JavaScript 變數中
- **無持久化**：重新整理頁面會遺失所有資料
- **無本地儲存**：未使用 localStorage 或 sessionStorage

---

## 【4】檔案結構與責任說明（File Responsibility Map）

| 檔案路徑 | 檔案角色 / 責任 | 是否為入口檔 | 是否為核心檔 | 修改風險備註 |
|---------|----------------|-------------|-------------|--------------|
| `index.html` | HTML 結構、UI 元素定義、第三方套件載入 | ✅ 是 | ⚠️ 部分 | **高風險**：修改 DOM 結構可能影響 JavaScript 選取器 |
| `app.js` | 核心業務邏輯：PDF 解析、題目生成、Word 生成、事件處理 | ❌ 否 | ✅ 是 | **極高風險**：包含所有核心功能，修改需謹慎 |
| `style.css` | 樣式定義、UI 美化 | ❌ 否 | ❌ 否 | **低風險**：主要影響視覺呈現 |
| `.gitattributes` | Git 屬性設定（文字檔案換行符號） | ❌ 否 | ❌ 否 | **無風險**：版本控制設定 |

### 入口檔案
- **`index.html`**：唯一入口檔案，瀏覽器直接開啟此檔案即可使用

### 核心模組（位於 `app.js`）

#### 1. 科目設定模組（第 6-20 行）
- `SUBJECT_CONFIG`：科目配置物件
- `currentSubject`：目前選擇的科目
- **修改風險：** 新增科目需同時修改 HTML 選單與解析邏輯

#### 2. PDF 解析器類別（第 29-435 行）
- `PDFParser`：解析 PDF 中的選擇題
  - `parsePDFs()`：解析多個 PDF
  - `parseSinglePDF()`：解析單個 PDF（依科目分派）
  - `parseFinancialByPage()`：Financial Accounting 專用解析
  - `parseFinancialOnePage()`：Financial Accounting 單頁解析
  - `extractQuestionsFromText()`：Managerial Accounting 專用解析
  - `parseQuestion()`：解析單個題目
  - `reconstructLines()`：重建 PDF 文字行
- **修改風險：** **極高**，解析邏輯與 PDF 格式緊密耦合

#### 3. 題目生成器類別（第 437-518 行）
- `QuestionGenerator`：隨機抽題與打亂順序
  - `generateExam()`：生成試卷（目前未使用章節配比功能）
  - `randomSelect()`：隨機選擇
  - `shuffle()`：打亂陣列
- **修改風險：** **中**，修改可能影響題目選擇邏輯

#### 4. Word 文檔生成器類別（第 520-937 行）
- `WordGenerator`：生成 Word 文檔
  - `generateQuestionSheet()`：生成題目卷（學生用）
  - `generateAnswerSheet()`：生成答案卷（教師用）
  - `downloadFile()`：下載檔案
- **修改風險：** **高**，修改可能影響 Word 文檔格式

#### 5. 主應用程式邏輯（第 939-1312 行）
- DOM 元素選取
- 事件監聽器設定
- 檔案上傳處理
- UI 狀態更新
- **修改風險：** **高**，修改可能影響使用者互動流程

### 第三方套件載入點
- **`index.html:58-64`**：所有第三方套件都在此處載入
- 修改此處需注意版本相容性

---

## 【5】核心資料結構（Key Data Structures）

### 1. 科目配置物件（SUBJECT_CONFIG）
```javascript
{
    managerial: {
        label: 'Managerial Accounting',
        pageTitle: 'Managerial Accounting Test Generator',
        defaultExamName: 'Managerial Accounting'
    },
    financial: {
        label: 'Financial Accounting',
        pageTitle: 'Financial Accounting Test Generator',
        defaultExamName: 'Financial Accounting'
    }
}
```

### 2. 題目物件（Question Object）
```javascript
{
    originalId: "MC.1.1" 或 "filename-p1",  // 原始題目 ID
    questionText: "題目文字內容...",          // 題目文字
    options: [                               // 選項陣列
        "a. 選項 A",
        "b. 選項 B ✔",                      // 正確答案可能包含 ✔
        "c. 選項 C",
        ...
    ],
    correctOption: "b",                      // 正確答案字母（小寫）
    hasCheckmark: true,                      // 是否有 checkmark 標記
    feedbackText: "Feedback 文字..."         // 回饋文字（可能為空字串）
}
```

### 3. 按檔案分組的題目陣列（parsedQuestionsByFile）
```javascript
[
    {
        file: File物件,                     // 原始 PDF 檔案物件
        fileName: "chapter1.pdf",           // 檔案名稱
        questions: [題目物件1, 題目物件2, ...] // 該檔案的題目陣列
    },
    {
        file: File物件,
        fileName: "chapter2.pdf",
        questions: [題目物件1, 題目物件2, ...]
    },
    ...
]
```

### 4. 解析結果物件（parsePDFs 返回值）
```javascript
{
    allQuestions: [題目物件1, 題目物件2, ...], // 所有題目的扁平陣列
    byFile: [                                 // 按檔案分組的陣列
        {
            file: File物件,
            fileName: "chapter1.pdf",
            questions: [題目物件1, ...]
        },
        ...
    ]
}
```

### 5. 題目數量設定物件（questionCounts，生成試卷時使用）
```javascript
[
    {
        fileIndex: 0,                        // 檔案索引
        fileName: "chapter1.pdf",            // 檔案名稱
        requestedCount: 5,                    // 請求的題目數
        availableCount: 10,                  // 可用的題目數
        questions: [題目物件1, ...]          // 該檔案的題目陣列
    },
    ...
]
```

---

## 【6】目前已完成的功能（What Works Now）

### ✅ 已穩定運作的功能

1. **科目選擇**
   - 可在 Managerial Accounting 與 Financial Accounting 之間切換
   - 切換時自動更新頁面標題與考卷名稱預設值

2. **PDF 上傳**
   - 支援點擊上傳
   - 支援拖放上傳
   - 支援多檔案上傳
   - 檔案列表顯示與移除功能

3. **PDF 解析**
   - **Managerial Accounting：** 解析格式為 `MC.xx.xx` 或 `MC.xx.xx.ALGO` 的題目
   - **Financial Accounting：** 解析 McGraw-Hill Connect Print View 格式的題目
   - 自動提取題目文字、選項、正確答案、Feedback
   - 按檔案分組顯示解析結果

4. **題目數量設定**
   - 為每個 PDF 設定要選擇的題目數量
   - 顯示每個 PDF 的可用題目數
   - 驗證題目數量（不能為負數、不能超過可用數量）

5. **隨機抽題**
   - 從每個 PDF 中隨機選擇指定數量的題目
   - 打亂所有選中題目的順序
   - 重新編號（1, 2, 3...）

6. **Word 文檔生成**
   - **題目卷（Questions.docx）：**
     - 包含考卷名稱、總題數
     - 作答檢查表（表格格式，每行 5 題）
     - 題目內容（移除標記與原始 ID）
     - 選項（移除 ✔ 和 ✓ 標記）
   - **答案卷（Answers.docx）：**
     - 包含考卷名稱與 "Answer Sheet" 標題
     - 答案摘要表格（Question No. / Answer）
     - 詳細答案（包含原始 ID、所有選項、正確答案標記、Feedback）

7. **檔案下載**
   - 自動下載兩個 Word 檔案
   - 檔案名稱清理（移除 Windows 不允許的字元）

### ⚠️ 部分完成或未使用的功能

1. **章節配比功能（未使用）**
   - `QuestionGenerator.parseChapterRatio()` 與 `generateExam()` 中的章節配比邏輯已實作
   - 但實際生成試卷時未使用此功能（見 `app.js:1263-1272`）
   - 目前是直接從每個 PDF 隨機選擇，而非使用章節配比文字

2. **科目專用解析規則（部分實作）**
   - `SUBJECT_CONFIG` 中有 `parserRules` 註解，但未實作
   - 目前透過 `currentSubject` 變數在 `parseSinglePDF()` 中分派解析邏輯

---

## 【7】目前限制與技術債（Known Limitations）

### 已知限制

1. **PDF 格式依賴性**
   - **Managerial Accounting：** 必須包含 `MC.xx.xx` 格式的題目標題
   - **Financial Accounting：** 必須是 McGraw-Hill Connect Print View 格式，且包含 `Award:` 行與特殊符號（`\uEA56`、`\uEA57`）
   - 不符合格式的 PDF 可能無法正確解析

2. **瀏覽器相容性**
   - 依賴現代瀏覽器的 File API、Blob API
   - 使用 CDN 載入第三方套件，需要網路連線

3. **記憶體限制**
   - 所有資料儲存在記憶體中，大型 PDF 或多檔案可能導致瀏覽器記憶體不足
   - 無資料持久化，重新整理頁面會遺失所有資料

4. **錯誤處理**
   - 解析失敗時會顯示錯誤訊息，但不會自動重試
   - 部分錯誤可能只顯示在 console，使用者看不到

5. **題目數量驗證**
   - 只驗證單個檔案的題目數量，不驗證總題數是否合理
   - 若所有檔案的題目數總和為 0，會顯示錯誤，但不會阻止使用者繼續操作

6. **Word 文檔格式**
   - 使用 docx.js 生成，格式固定，無法自訂
   - 題目卷會移除 `(Appendix...)` 和 `(Algorithmic)` 標籤，但答案卷不會

### 技術債

1. **程式碼註解**
   - 部分關鍵邏輯缺少註解（例如 `reconstructLines()` 的座標系統處理）
   - Financial Accounting 解析邏輯使用 `var` 而非 `const/let`

2. **程式碼重複**
   - `parseFinancialOnePage()` 中重複定義 `CIRCLE` 和 `ARROW` 常數（第 173-174 行），與類別屬性重複（第 31-34 行）

3. **未使用的功能**
   - `QuestionGenerator.generateExam()` 中的章節配比功能已實作但未使用
   - `QuestionGenerator.parseChapterRatio()` 已實作但從未被呼叫

4. **全域變數**
   - `pdfFiles`、`parsedQuestions`、`parsedQuestionsByFile` 等使用全域變數，可能造成命名衝突

5. **硬編碼值**
   - 作答檢查表每行題數固定為 5（`app.js:557`）
   - 檔案下載延遲固定為 300ms（`app.js:1298`）

---

## 【8】未來修改指引（VERY IMPORTANT）

### 若要「調整既有行為」，應從哪裡開始看

#### 修改 PDF 解析邏輯
1. **Managerial Accounting：**
   - 檢查 `PDFParser.extractQuestionsFromText()`（第 284-313 行）
   - 檢查 `PDFParser.parseQuestion()`（第 316-434 行）
   - 修改正則表達式（`mcHeaderPattern`、`optionStartMatch` 等）

2. **Financial Accounting：**
   - 檢查 `PDFParser.parseFinancialByPage()`（第 142-170 行）
   - 檢查 `PDFParser.parseFinancialOnePage()`（第 172-281 行）
   - 修改 `awardPattern`、`CIRCLE`、`ARROW` 符號處理

#### 修改 Word 文檔格式
- **題目卷：** `WordGenerator.generateQuestionSheet()`（第 523-684 行）
- **答案卷：** `WordGenerator.generateAnswerSheet()`（第 687-931 行）
- 修改 docx.js 的段落、表格、樣式設定

#### 修改抽題邏輯
- `QuestionGenerator.shuffle()`（第 510-517 行）
- 生成試卷時的抽題邏輯（第 1261-1280 行）

#### 修改 UI 流程
- 事件監聽器設定（第 986-1011 行）
- UI 狀態更新函數（`updateFileList()`、`updateExportButton()`）

### 哪些檔案【不應隨意修改】

1. **`app.js` 中的解析邏輯（第 29-435 行）**
   - **原因：** 與 PDF 格式緊密耦合，修改可能導致解析失敗
   - **建議：** 修改前先確認 PDF 格式，並進行充分測試

2. **`app.js` 中的 Word 生成邏輯（第 520-937 行）**
   - **原因：** 格式固定，修改可能影響文檔結構
   - **建議：** 修改前先了解 docx.js API，並測試生成的 Word 檔案

3. **`index.html` 中的 DOM 結構（第 10-56 行）**
   - **原因：** JavaScript 使用 `getElementById` 選取元素，修改 ID 會導致功能失效
   - **建議：** 修改 HTML 時，同步檢查 `app.js` 中的選取器（第 941-952 行）

4. **第三方套件版本（`index.html:58-64`）**
   - **原因：** 版本變更可能導致 API 不相容
   - **建議：** 升級前先閱讀 changelog，並進行完整測試

### 哪些區域較適合擴充

1. **新增科目**
   - 在 `SUBJECT_CONFIG` 中新增科目（第 6-19 行）
   - 在 `index.html` 的 `<select>` 中新增選項（第 13-16 行）
   - 在 `PDFParser.parseSinglePDF()` 中新增解析邏輯（第 119-140 行）

2. **新增題目類型**
   - 擴充 `parseQuestion()` 或新增解析方法
   - 擴充題目物件結構（需同步修改 Word 生成邏輯）

3. **改善錯誤處理**
   - 在 `parsePDFs()` 中新增更詳細的錯誤訊息（第 1136-1184 行）
   - 在生成試卷時新增驗證邏輯（第 1187-1312 行）

4. **新增 UI 功能**
   - 在 `updateFileList()` 中新增功能（第 1032-1084 行）
   - 新增設定選項（例如：作答檢查表每行題數）

### 修改前應先向使用者確認的關鍵問題

1. **PDF 格式變更：**
   - 「PDF 格式是否有變更？是否有範例檔案可提供？」
   - 「新格式與現有格式的差異為何？」

2. **功能需求：**
   - 「是否需要保留現有功能，還是可以替換？」
   - 「新功能的優先順序為何？」

3. **Word 文檔格式：**
   - 「是否需要調整 Word 文檔的格式或樣式？」
   - 「是否需要新增或移除某些內容？」

4. **測試資料：**
   - 「是否有測試用的 PDF 檔案？」
   - 「預期的解析結果為何？」

---

## 【9】給未來 AI 的交接備註（AI Handoff Notes）

### 若未來收到新增/修改需求，第一步應檢查哪些檔案

1. **PDF 解析相關需求：**
   - 先檢查 `app.js` 中的 `PDFParser` 類別（第 29-435 行）
   - 確認目前支援的 PDF 格式（見第 7 節「目前限制」）
   - 檢查 `currentSubject` 變數，確認是否需要新增科目

2. **Word 文檔格式相關需求：**
   - 先檢查 `WordGenerator` 類別（第 520-937 行）
   - 確認目前的文檔結構（題目卷 vs 答案卷）
   - 檢查 docx.js 的 API 文件

3. **UI 流程相關需求：**
   - 先檢查 `index.html` 的 DOM 結構
   - 檢查 `app.js` 中的事件監聽器（第 986-1012 行）
   - 檢查 UI 更新函數（`updateFileList()`、`updateExportButton()`）

4. **抽題邏輯相關需求：**
   - 先檢查 `QuestionGenerator` 類別（第 437-518 行）
   - 檢查生成試卷時的抽題邏輯（第 1261-1280 行）

### 在不確定需求影響範圍時，應先詢問使用者哪些問題

1. **PDF 格式：**
   - 「PDF 檔案的來源為何？（例如：McGraw-Hill Connect、自製題庫）」
   - 「PDF 格式是否與現有格式相同？是否有範例檔案？」

2. **功能範圍：**
   - 「此需求是否會影響現有功能？」
   - 「是否需要同時支援多種格式或科目？」

3. **預期行為：**
   - 「預期的輸出結果為何？（例如：Word 文檔格式、題目順序）」
   - 「是否有特殊需求？（例如：題目數量限制、格式要求）」

4. **測試資料：**
   - 「是否有測試用的 PDF 檔案？」
   - 「預期的解析結果為何？」

### 任何你認為「未來 AI 很容易誤判」的地方

1. **科目切換邏輯：**
   - `currentSubject` 是全域變數，切換科目時會影響 `parseSinglePDF()` 的解析邏輯
   - **誤判風險：** 可能以為解析邏輯是統一的，實際上依科目不同而不同

2. **題目物件結構：**
   - `options` 陣列中的正確答案可能包含 `✔` 標記，但 `correctOption` 是字母
   - **誤判風險：** 可能以為 `correctOption` 是選項文字，實際上是字母（a-e）

3. **檔案分組邏輯：**
   - `parsedQuestionsByFile` 是按檔案分組的陣列，但生成試卷時是從每個檔案分別抽題
   - **誤判風險：** 可能以為所有題目會混合在一起抽題，實際上是按檔案分別抽題後再打亂

4. **Word 文檔生成：**
   - 題目卷會移除 `(Appendix...)` 和 `(Algorithmic)` 標籤，但答案卷不會
   - **誤判風險：** 可能以為兩個文檔的題目文字完全相同，實際上題目卷有清理過

5. **章節配比功能：**
   - `QuestionGenerator.generateExam()` 中有章節配比邏輯，但實際生成試卷時未使用
   - **誤判風險：** 可能以為章節配比功能已啟用，實際上被忽略

6. **自動解析：**
   - 上傳檔案後會自動觸發解析（延遲 500ms），無需手動點擊按鈕
   - **誤判風險：** 可能以為需要手動觸發解析，實際上已自動化

7. **Financial Accounting 符號：**
   - 使用 Unicode 私有使用區符號（`\uEA56`、`\uEA57`）來識別選項和正確答案
   - **誤判風險：** 可能以為是標準符號，實際上是特定 PDF 格式的特殊符號

---

## 【附錄】檔案檢查清單

### 已檢查的檔案
- ✅ `index.html`（完整閱讀）
- ✅ `app.js`（完整閱讀）
- ✅ `style.css`（完整閱讀）
- ✅ `.gitattributes`（完整閱讀）

### 未找到的檔案類型
- ❌ `package.json`（無）
- ❌ `README.md`（無）
- ❌ `*.config.js`（無）
- ❌ 測試檔案（無）
- ❌ 範例 PDF 檔案（無）

### 關鍵字搜尋結果
- ✅ 搜尋「upload」：找到檔案上傳相關邏輯（`app.js:986-1029`）
- ✅ 搜尋「parse」：找到 PDF 解析相關邏輯（`app.js:29-435`）
- ✅ 搜尋「generate」：找到試卷生成相關邏輯（`app.js:437-937`）
- ✅ 搜尋「download」：找到檔案下載相關邏輯（`app.js:1294-1305`）

---

**文件結束**
