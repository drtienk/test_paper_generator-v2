# PROJECT_STATUS.md

**專案名稱：** Managerial Accounting Test Generator（會計試卷生成器）  
**本次文件製作日期：** 2026-01-30  
**文件用途：** This document is intended for future AI / developer handoff. 供完全無背景知識的接手者（含下一位 AI）在不問任何問題的情況下，快速理解並繼續開發此專案。

---

## 1. Project Overview（給完全不懂的人看）

### 這個 WebApp 是在做什麼的
這是一個**純前端的試題解析與試卷生成工具**，主要用途是：
- 從 **PDF 題庫**中解析選擇題（MC）與非選擇題（EX/PR）
- 支援**多個 PDF**（每個 PDF 對應一個章節），並為每個檔案設定要抽的 MC / EX 題數
- **Managerial Accounting** 另支援：上傳 **Word (.docx) 非選擇題題庫**、**考卷名稱組裝**（年/學期/考試類型）、**Exam Points 表格**（羅馬數字 I/II/III… 與分數）
- 依設定**隨機抽題**後，產生 **Word 格式**的題目卷與答案卷，供下載列印

### 使用者是誰
**教師或教學人員**（介面為繁體中文，用語為考卷名稱、生成試卷、解析 PDF 等）。

### 解決什麼問題
- 從 McGraw-Hill 等教材匯出的 PDF 題庫中，自動擷取選擇題與非選擇題
- 按章節（檔案）分配題數，避免手動複製貼上
- 一鍵產生符合格式的題目卷與答案卷（含封面、作答格、分數表、詳細答案與 Feedback）

### 目前已經可以做到哪些事情
- **科目切換**：Managerial Accounting / Financial Accounting，切換時標題與預設考卷名稱會更新
- **PDF 上傳**：點擊或拖放、多檔上傳；上傳後約 500ms 自動觸發解析
- **Word 上傳（僅 Managerial）**：上傳 .docx 作為非選擇題題庫，格式為「題號 + 題幹 + ANSWER: + 答案」
- **設定出題規則**：考卷名稱（可手輸或用 Builder）、Exam Points（動態列 I~X、總分）、每個 PDF 的 MC/EX 題數、Word 要抽幾題
- **解析 PDF**：Managerial 解析 MC.xx.xx / EX.xx.xx / PR.xx.xx；Financial 解析 McGraw-Hill Connect Print View（Award:、特殊符號）
- **生成試卷**：依各檔案設定的題數隨機抽題、打亂順序、重新編號，產出兩份 .docx（Questions / Answers），並觸發下載

---

## 2. Tech Stack & Execution Model

### 是否為純前端
**是。** 僅使用 HTML、CSS、JavaScript，無後端、無 API、無建置工具（無 Node/npm 必須）。

### 執行環境
- **Local-only 或靜態託管**：直接開 `index.html` 或放上 GitHub Pages / 任一靜態主機即可
- 需**網路**：第三方套件經 CDN 載入（pdf.js、docx、FileSaver、mammoth）

### 登入、權限、儲存
- **無登入、無權限**
- **無 localStorage / sessionStorage**：所有狀態僅在記憶體，重整即清空
- **無雲端同步**：無 Supabase、Firebase 等

### 專案類型
**表單導向 + 文件生成**：使用者透過表單上傳檔案、設定題數與考卷名稱，系統解析後生成 Word 試卷並下載。

---

## 3. High-Level Architecture（非常重要）

### 啟動流程
1. 瀏覽器載入 **`index.html`**
2. 依序載入：`style.css` → pdf.js（含 worker）→ docx → FileSaver → mammoth → **`app.js`**
3. **`app.js` 為單一入口腳本**：無獨立 init 檔案，腳本載入後立即執行
   - 定義 `SUBJECT_CONFIG`、全域變數、`PDFParser` / `QuestionGenerator` / `WordGenerator` 類別與 Word 題庫解析函數
   - 取得 DOM 元素（約 1936–1960 行）、綁定事件、呼叫 `applySubjectUI(null)`、`renderPointsConfigUI()`、`updateTotalDisplay()`、`updateExportButton()`，即完成初始化

### 核心模組分層概念
- **UI 層**：`index.html` 的區塊（科目、上傳、考卷名稱、Points、檔案列表、解析狀態、生成按鈕）；`app.js` 內的事件監聽與 `updateFileList` / `updateExportButton` / `updateWordParseUI` 等
- **State 層**：全域變數 `pdfFiles`、`parsedQuestions`、`parsedQuestionsByFile`、`exRequestedCountsByFileIndex`、`wordFile`、`wordNonMcQuestions`、`pointsRows`、`currentSubject`、`selectedYear/Term/ExamType`
- **Logic 層**：`PDFParser`（解析 PDF）、`parseWordFile` / `parseWordQuestions`（解析 Word 題庫）、`QuestionGenerator`（shuffle、randomSelect）、抽題與組題邏輯（在 Generate 按鈕的 handler 內）
- **Builder / Generator 層**：`WordGenerator`（封面、Points 表、答案格、題目卷、答案卷、EX/Word 非選擇題區塊）、檔名清理與下載

### 「不要改」的穩定核心
- **PDF 解析**：`PDFParser.reconstructLines`、`extractQuestionsFromText`、`parseQuestion`（Managerial MC）；`parseFinancialByPage`、`parseFinancialOnePage`（Financial）；`extractExerciseQuestionsFromText`、`parseExerciseQuestion`（EX）。與特定 PDF 版面與符號強耦合，改動易導致解析全掛。
- **題目物件結構**：MC 題的 `originalId`、`questionText`、`options`、`correctOption`、`hasCheckmark`、`feedbackText`；EX 題的 `originalId`、`type`、`promptText`、`requiredText`、`answerTextOrTokens`、`rawBlockText`。Word 生成與答案卷都依賴這些欄位。
- **`index.html` 的 id**：`app.js` 用 `getElementById` 抓元素（如 `subjectSelect`、`examName`、`questionCount_${index}`、`exQuestionCount_${index}`）。改 id 必須同步改 `app.js`。

### 「經常被修改」的功能模組通常在哪
- **UI 文案與步驟**：`index.html` 的 section 標題、提示文字
- **考卷名稱與 Points**：Exam Name Builder 的按鈕值（`index.html` 的 `data-value`）、`pointsRows` 預設與 `ROMAN_NUMERALS`、`renderPointsConfigUI` / `addPointsRow`
- **Word 版面**：`WordGenerator._buildManagerialCoverPage`、`generateQuestionSheet`、`generateAnswerSheet` 內的段落與表格（字型、間距、EX/Word 非選擇題區塊）
- **抽題與驗證**：Generate 按鈕的 click handler（題數驗證、MC/EX/Word 抽題、呼叫 `generateQuestionSheet` / `generateAnswerSheet` 的參數）

---

## 4. File & Folder Responsibility Map（交接重點）

### 重要檔案與用途、修改時機

| 檔案 | 負責什麼 | 什麼情況下需要改 | 新功能從這裡加還是只被呼叫 |
|------|----------|------------------|----------------------------|
| **index.html** | 唯一 HTML 入口：版面結構、科目選單、上傳區、考卷名稱與 Builder、Points 區、檔案列表容器、解析/生成區塊、所有 CDN script | 新增/刪除步驟、改 id、改文案、新增表單欄位 | 新 UI 區塊在這裡加；邏輯在 app.js |
| **app.js** | 全部業務邏輯：科目設定、PDF/Word 解析、抽題、Word 生成、DOM 取得、事件綁定、狀態更新、解析與生成流程 | 任何行為或流程變更、新增科目、改解析規則、改 Word 版面、改抽題規則 | 新功能多數在此加；僅樣式改 style.css |
| **style.css** | 全域與元件樣式（字體、背景、container、step、upload-area、按鈕、檔案列表、status 等） | 改版型、顏色、間距、響應式 | 只改樣式時改此檔；不改行為 |
| **.gitattributes** | Git 換行符號等屬性 | 通常不需改 | 僅版控設定 |

### 補充說明
- **無** `package.json`、無 `README.md`、無測試檔、無範例 PDF/Word；專案根目錄即上述幾檔。
- **app.js 行數約 2700+**：找功能時可搜 `parsePDFs`、`generateQuestionSheet`、`generateAnswerSheet`、`updateFileList`、`addFiles`、`parseSinglePDF`、`parseFinancialOnePage`、`extractExerciseQuestionsFromText`、`parseWordFile` 等關鍵字。

---

## 5. Core Data Structures & Concepts

### 專案中最重要的資料概念
- **SUBJECT_CONFIG / currentSubject**：科目設定（managerial / financial）；決定標題、是否顯示 Exam Name Builder / Word 上傳、解析與 Word 輸出版本。
- **pdfFiles**：使用者上傳的 PDF `File` 陣列。
- **parsedQuestions**：所有 MC 題的一維陣列（題目物件）。
- **parsedQuestionsByFile**：依檔案分組的解析結果，每項為 `{ file, fileName, questions, mcQuestions, exQuestions }`；`questions` 與 `mcQuestions` 相同，皆為該檔 MC 題陣列。
- **exRequestedCountsByFileIndex**：每個 PDF 要抽的 EX 題數（依檔案索引對應）。
- **wordFile / wordNonMcQuestions / wordParseState**：Word 題庫檔案、解析出的題目陣列、解析狀態（idle / parsing / parsed / error）。
- **pointsRows**：Exam Points 每一列的 `{ label, value }`（label 為羅馬數字 I、II…），用於 Managerial 封面表格與總分。
- **題目物件（MC）**：`originalId`、`questionText`、`options`（字串陣列，正確項可能含 ✔）、`correctOption`（a–e）、`hasCheckmark`、`feedbackText`。
- **題目物件（EX）**：`originalId`、`type`（EX/PR）、`promptText`、`requiredText`、`answerTextOrTokens`、`rawBlockText`。
- **Word 題目物件**：`originalId`、`questionLines`、`answerLines`、`hasAnswerSection`、`rawLines`；格式為題號開頭 + 題幹 + `ANSWER:` + 答案。

### 這些資料通常存在哪
- **記憶體**：上述全部；重整即消失。
- **不寫入**：無 localStorage、無後端、無上傳到雲端。
- **輸出**：僅透過 `WordGenerator` 產生 Blob，再以 FileSaver 下載為 .docx。

### 新接手的人「一定要先理解」的名詞
- **MC**：選擇題（Multiple Choice）。
- **EX / PR**：PDF 內的非選擇題（Exercise / Problem），僅 Managerial 解析並可納入試卷。
- **byFile / parsedQuestionsByFile**：按「每個 PDF」分開的題目，抽題時是「每個檔案各自抽 N 題」再合併打亂，不是全混在一起抽。
- **correctOption**：單一字母（a–e），不是選項全文。
- **Managerial 專用**：Exam Name Builder、Points 表、Word 上傳、EX 題、封面與答案格版型，都只在 `currentSubject === 'managerial'` 時使用或顯示。
- **Financial**：一 PDF 一頁一題、依 `Award:` 與 Unicode 符號（如 `\uEA56`、`\uEA57`）辨識選項與正確答案。

---

## 6. Current Features (As-Is)

### 已完成的功能列表
- 科目選擇（Managerial / Financial）與介面/標題/預設考卷名連動
- PDF 上傳（多檔、拖放）、上傳後自動解析
- 解析結果按檔案顯示 MC/EX 題數；每個檔案可設 MC 題數與 EX 題數
- 考卷名稱：手輸或（Managerial）Builder（年/學期/考試類型）
- Exam Points：動態列（I–X）、每列分數、總分、新增/刪除列（至少 2 列、最多 10 列）
- Word 非選擇題上傳（僅 Managerial）：解析 .docx、顯示題數、設定要抽題數、清除
- 解析 PDF：Managerial 的 MC + EX；Financial 的 MC（Connect Print View）
- 生成試卷：依各檔 MC/EX 題數抽題、Word 題數抽題（Managerial）、打亂、編號，產出題目卷與答案卷
- 題目卷：Managerial 含封面（單位、考卷名、Section/Name、說明、Points 表、答案格、I. MULTIPLE CHOICE）、題目列表、EX 區、Word 非選擇題區；Financial 為標題 + 總題數 + 作答檢查表 + 題目
- 答案卷：Answer Summary 表、Detailed Answers（題幹、Original ID、選項與 ✔、Correct Answer、Feedback）；Managerial 另有 EX 與 Word 非選擇題答案區

### 功能是否穩定
- **穩定**：科目切換、PDF 上傳與解析（已知格式）、MC 抽題與 Word 題目卷/答案卷基本流程、檔名清理與下載。
- **需注意**：解析高度依賴 PDF 版面與符號；PDF 格式一變就可能需改 `PDFParser`。Word 題庫需符合「題號 + ANSWER:」格式。

### 哪些功能是最近才加、風險較高
- **Exam Points 動態列**（新增/刪除列、羅馬數字重編）：與 `_buildManagerialCoverPage` 的 `ptsRows` 綁在一起，改預設或欄位時要兩邊一起看。
- **Word 非選擇題**（mammoth 解析、`parseWordQuestions`、題目卷/答案卷中的 Word 區塊）：若 .docx 結構與假設不同，解析可能錯位。
- **EX 題目卷的 Required 區塊清理**（`stripAnswerBlocksFromRequiredLines`）：邏輯較細，改題庫格式時易出錯。
- **章節配比**：`QuestionGenerator.parseChapterRatio` / `generateExam` 有實作，但**目前生成流程未使用**，實際是依各檔題數直接抽題。

---

## 7. Known Constraints & "Do NOT Break These"

### 明確不能改的行為
- **解析結果的 byFile 結構**：`parsedQuestionsByFile[i].questions` 必須是該檔的 **MC 題陣列**；EX 題在 `exQuestions`。生成時是「每檔取 MC 題數 + 每檔取 EX 題數」，若改成單一扁平陣列且不按檔分組，會破壞現有抽題邏輯。
- **MC 題目物件的欄位**：`originalId`、`questionText`、`options`、`correctOption`、`feedbackText` 被 Word 生成與答案卷大量使用；刪欄位或改名會導致版面或答案錯亂。
- **Financial 符號**：`PDFParser.FINANCIAL_SYMBOLS` 與 `parseFinancialOnePage` 內的 `CIRCLE` / `ARROW`（Unicode 私有區）用於辨識選項與正確答案；改掉會導致 Financial 解析失敗。
- **科目分派**：`parseSinglePDF` 依 `currentSubject === 'financial'` 走 Financial 解析，其餘走 Managerial；若合併或反轉條件，兩科解析都會亂。

### 哪些流程重構會整個壞掉
- **解析 → 存 byFile → 檔案列表顯示題數輸入 → 生成時讀各檔 MC/EX 題數**：任一環節改成「只存扁平 list」或「不按檔抽題」，都會與 UI 與 Word 輸出預期不符。
- **Generate 按鈕內**：先組 `questionCounts`（每檔 requestedCount/availableCount/questions）、再抽 MC、再抽 EX、再抽 Word、再呼叫 `generateQuestionSheet(examName, examQuestions, examPoints, exSelectedAll, wordNonMcSelected)` 與 `generateAnswerSheet(...)`；參數順序與內容被寫死，改簽名或順序會炸。

### 刻意寫得保守的地方
- **錯誤處理**：解析失敗會丟錯或寫 console，不一定每個分支都有使用者可見訊息；部分驗證只做基本檢查（如題數 ≥0、≤ 可用數）。
- **Word 題庫格式**：假設題目以數字題號開頭、答案區以 `ANSWER:` 標記；其餘格式未支援。
- **EX 題目卷**：Required 區塊會過濾「答案型」行（$、數字、✔ 等），避免把答案印給學生；規則若與題庫不符，可能多刪或少刪。

---

## 8. How to Add / Modify Features（給下一個 AI 的路線圖）

| 需求類型 | 從哪一層/哪個檔案開始 | 建議步驟 |
|----------|------------------------|----------|
| **新分頁/新步驟** | `index.html` 新增 section → `app.js` 用 `getElementById` 取元素（若需）→ 在適當時機顯示/隱藏（可參考 `parseSection`、`generateSection`、`wordUploadSection`） | 先加 HTML 與 id，再在 app.js 補變數與顯示邏輯 |
| **新欄位（表單）** | `index.html` 加 input/select → `app.js` 取元素、在 submit/生成時讀值並參與邏輯 | 注意 id 與現有命名不衝突 |
| **新表單（例如新上傳類型）** | 同上；若需解析，在 app.js 加解析函數與狀態變數，再在生成流程中讀取並傳給 WordGenerator | 參考 Word 上傳：`wordInput`、`parseWordFile`、`wordNonMcQuestions`、`wordNonMcSelected` |
| **新匯出格式** | 在 `app.js` 生成流程中，在現有 `WordGenerator` 呼叫之後或取代，組新 Blob 並用 `saveAs` 下載 | 若仍要 DOCX，可複用 `WordGenerator` 或擴充其方法；若是 PDF/其他格式，需新依賴或新類別 |
| **新邏輯規則（例如抽題權重）** | `app.js` 中 Generate 按鈕的 handler：在組 `questionCounts` 與 `selectedQuestions` / `exSelectedAll` / `wordNonMcSelected` 的段落改 | 不建議改 `QuestionGenerator` 的 shuffle/randomSelect 簽名，除非全專案一併改呼叫處 |
| **新科目** | `SUBJECT_CONFIG` 加一組 → `index.html` 的 `subjectSelect` 加 option → `parseSinglePDF` 加分支（或共用現有 parser）→ 若該科有專用 UI（如 Builder、Points），在 `applySubjectUI` 與生成流程中加條件 | 解析與 Word 版面若與現有科目不同，需在 `WordGenerator` 內用 `currentSubject` 分岐 |
| **改 PDF 解析** | `app.js` 的 `PDFParser`：Managerial MC 改 `extractQuestionsFromText` / `parseQuestion`；Managerial EX 改 `extractExerciseQuestionsFromText` / `parseExerciseQuestion`；Financial 改 `parseFinancialByPage` / `parseFinancialOnePage` | 先確認新 PDF 樣本與現有正則/符號，再小範圍改並用真實檔測試 |
| **改 Word 版面** | `app.js` 的 `WordGenerator`：封面 `_buildManagerialCoverPage`、題目卷 `generateQuestionSheet`、答案卷 `generateAnswerSheet` | 改 docx 的 Paragraph/Table 參數；注意 Managerial 與 Financial 分支 |

---

## 9. Current Development Status

### 專案階段
**功能擴充期、局部穩定**：核心流程（上傳 → 解析 → 設題數 → 生成 DOCX）已可穩定使用；Managerial 的 Exam Name Builder、Points、EX、Word 題庫為後續疊加，程式碼集中在一份 `app.js`，尚未模組化。

### 最近常改的是哪一塊
- **Managerial 專用**：封面（Points、答案格）、EX 題目卷/答案卷、Word 非選擇題上傳與輸出、Exam Name Builder。
- **抽題與生成**：Generate 按鈕內的 MC/EX/Word 抽題與參數傳遞。

### 技術債最多的地方
- **單檔 app.js 過大**：所有邏輯與 UI 綁定都在同一檔案，不利搜尋與重構。
- **全域變數多**：`pdfFiles`、`parsedQuestions`、`parsedQuestionsByFile`、`exRequestedCountsByFileIndex`、`wordFile`、`wordNonMcQuestions`、`pointsRows` 等，易與未來擴充衝突。
- **解析與格式強耦合**：PDF/Word 格式一變就要改解析程式，且註解不足，接手者需對照實際檔案閱讀。
- **未使用的程式**：`QuestionGenerator.generateExam` / `parseChapterRatio`（章節配比）目前未被呼叫。
- **重複常數**：例如 Financial 的 `CIRCLE`/`ARROW` 在類別屬性與 `parseFinancialOnePage` 內都有定義。

---

## 10. Recommended First Steps for the Next AI / Developer

### 接手後「第一天」該做什麼
1. **本地跑起來**：用瀏覽器直接開 `index.html`（需網路以載入 CDN）；若有 PDF/Word 題庫樣本，走一遍上傳 → 設題數 → 生成下載，確認行為與本文件一致。
2. **讀本文件**：尤其 §3 架構、§4 檔案責任、§5 資料結構、§7 不能改的點。
3. **對照程式**：在 `app.js` 搜 `parsePDFs`、`generateBtn.addEventListener`、`WordGenerator`、`parsedQuestionsByFile`，順一次「上傳 → 解析 → 存 byFile → 生成時讀題數 → 抽題 → 產生 DOCX」流程。
4. **確認需求再動刀**：任何改解析、改題目結構、改生成參數，先確認是否有對應的 PDF/Word 樣本與預期結果，避免盲目重構。

### 建議閱讀順序
1. **PROJECT_STATUS.md**（本文件）§1–§5、§7  
2. **index.html**：從頭到尾看結構與 id  
3. **app.js**：先看開頭（SUBJECT_CONFIG、全域變數、PDFParser/QuestionGenerator/WordGenerator 類別名與主要方法名），再看「主應用程式邏輯」區（DOM、事件、`updateFileList`、`parsePDFs`、Generate 的 handler）  
4. **style.css**：有需要改版再細看  

### 建議不要一開始就碰的地方
- **PDFParser 內部的正則與符號**（尤其是 Financial 的 `parseFinancialOnePage`、Managerial 的 `extractQuestionsFromText` / `parseQuestion`）：除非要支援新 PDF 格式，否則先別重構。
- **WordGenerator 內 docx 的細部參數**（字型、間距、表格欄寬）：先改文案或明顯區塊，再動版型。
- **`parsedQuestionsByFile` 的結構與「按檔抽題」的迴圈**：這是整份試卷邏輯的支點，不確定前勿改成「全題目混在一起」抽題。

---

**文件結束**
