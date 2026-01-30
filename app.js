// ============================================
// Managerial Accounting Test Generator - MVP
// ============================================

// ========== 科目設定（Subject / Course） ==========
const SUBJECT_CONFIG = {
    managerial: {
        label: 'Managerial Accounting',
        pageTitle: 'Managerial Accounting Test Generator',
        defaultExamName: 'Managerial Accounting',
        // parserRules: 暫時指向同一套
    },
    financial: {
        label: 'Financial Accounting',
        pageTitle: 'Financial Accounting Test Generator',
        defaultExamName: 'Financial Accounting',
        // parserRules: 暫時指向同一套（先一樣）
    }
};
let currentSubject = 'managerial';

// Managerial Accounting Exam Name Builder 設定
const MANAGERIAL_EXAMNAME_PREFIX = "ACCT 201 Managerial Accounting";
let selectedYear = '2026';
let selectedTerm = 'Spring';
let selectedExamType = 'Exam 1';

// ========== 全域變數 ==========
let pdfFiles = [];
let parsedQuestions = []; // 所有題目的陣列
let parsedQuestionsByFile = []; // 按檔案分組的題目 [{file, questions}, ...]
let parsedExerciseQuestions = []; // 所有 EX 題目的陣列（暫不進入匯出流程）
let exRequestedCountsByFileIndex = []; // 使用者對每檔案設定的 EX 題數（state）
let parser = null;
let generator = null;

// Word 非選擇題題庫（僅 Managerial）
let wordFile = null;
let wordNonMcQuestions = [];
let wordParseState = 'idle'; // 'idle' | 'parsing' | 'parsed' | 'error'

// ========== PDF 解析器類別 ==========
class PDFParser {
    static FINANCIAL_SYMBOLS = {
        CIRCLE: '\uEA56',
        ARROW: '\uEA57',
    };

    constructor() {
        this.questions = [];
    }

    // 解析多個 PDF 檔案（返回按檔案分組的結果）
    async parsePDFs(files) {
        this.questions = [];
        const allExQuestions = [];
        const resultsByFile = [];
        
        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            try {
                const parsed = await this.parseSinglePDF(file);
                const mcQuestions = (parsed && parsed.mcQuestions) ? parsed.mcQuestions : [];
                const exQuestions = (parsed && parsed.exQuestions) ? parsed.exQuestions : [];

                // 注意：維持既有行為 —— allQuestions / questions 仍代表 MC（含 Financial）題目集合
                this.questions = this.questions.concat(mcQuestions);
                allExQuestions.push(...exQuestions);
                resultsByFile.push({
                    file: file,
                    fileName: file.name,
                    // 為了不破壞既有流程：questions 仍等於 MC 題目陣列
                    questions: mcQuestions,
                    mcQuestions: mcQuestions,
                    exQuestions: exQuestions
                });
            } catch (error) {
                console.error(`解析 ${file.name} 失敗:`, error);
                throw new Error(`無法解析 ${file.name}: ${error.message}`);
            }
        }
        
        return {
            allQuestions: this.questions,
            allExQuestions: allExQuestions,
            byFile: resultsByFile
        };
    }

    // 從 textContent.items 重建行（依 Y 座標分組）
    reconstructLines(items) {
        if (!items || items.length === 0) return [];

        // 提取每個 item 的位置和文字
        const positioned = items.map(item => {
            const x = item.transform ? item.transform[4] : 0;
            const y = item.transform ? item.transform[5] : 0;
            return { x, y, str: item.str };
        });

        // 依 Y 座標降序排序（PDF 座標系 Y 從下往上），再依 X 升序
        positioned.sort((a, b) => {
            const yDiff = Math.abs(a.y - b.y);
            if (yDiff > 3) {
                return b.y - a.y; // Y 降序（上面的先）
            }
            return a.x - b.x; // X 升序（左邊的先）
        });

        // 依 Y 座標分組成行（容差 3 單位）
        const lines = [];
        let currentLine = [];
        let currentY = null;

        for (const item of positioned) {
            if (currentY === null || Math.abs(item.y - currentY) <= 3) {
                currentLine.push(item);
                if (currentY === null) currentY = item.y;
            } else {
                // 新的一行
                if (currentLine.length > 0) {
                    currentLine.sort((a, b) => a.x - b.x);
                    const lineText = currentLine.map(i => i.str).join(' ').trim();
                    if (lineText) lines.push(lineText);
                }
                currentLine = [item];
                currentY = item.y;
            }
        }

        // 處理最後一行
        if (currentLine.length > 0) {
            currentLine.sort((a, b) => a.x - b.x);
            const lineText = currentLine.map(i => i.str).join(' ').trim();
            if (lineText) lines.push(lineText);
        }

        return lines;
    }

    // 解析單個 PDF（依 currentSubject 分派 Financial / Managerial）
    async parseSinglePDF(file) {
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

        if (currentSubject === 'financial') {
            const mcQuestions = await this.parseFinancialByPage(pdf, file);
            return { mcQuestions, exQuestions: [] };
        }

        // 以下是原有 Managerial 邏輯，完全不要動
        let allLines = [];
        for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
            const page = await pdf.getPage(pageNum);
            const textContent = await page.getTextContent();
            const pageLines = this.reconstructLines(textContent.items);
            allLines = allLines.concat(pageLines);
        }
        const fullText = allLines.join('\n');
        const mcQuestions = this.extractQuestionsFromText(fullText);
        const exQuestions = this.extractExerciseQuestionsFromText(fullText);
        return { mcQuestions, exQuestions };
    }

    async parseFinancialByPage(pdfDoc, fileMeta) {
        const questions = [];
        const fileBaseName = (fileMeta && fileMeta.name ? fileMeta.name : 'FA').replace(/\.pdf$/i, '');

        console.log('[Financial] 開始解析: ' + fileBaseName + ', 共 ' + pdfDoc.numPages + ' 頁');

        for (let pageNum = 1; pageNum <= pdfDoc.numPages; pageNum++) {
            try {
                const page = await pdfDoc.getPage(pageNum);
                const textContent = await page.getTextContent();
                const rawLines = this.reconstructLines(textContent.items);
                const lines = rawLines.map(function(l) { return (l || '').trim(); }).filter(function(l) { return l.length > 0; });

                const q = this.parseFinancialOnePage(lines, pageNum, fileBaseName);

                if (q) {
                    questions.push(q);
                    console.log('[Financial][' + fileBaseName + '][p' + pageNum + '] ✓ opts=' + q.options.length + ', answer=' + q.correctOption);
                } else {
                    console.warn('[Financial][' + fileBaseName + '][p' + pageNum + '] ✗ 未解析出題目');
                }
            } catch (e) {
                console.error('[Financial][' + fileBaseName + '][p' + pageNum + '] 異常:', e.message);
            }
        }

        console.log('[Financial] ' + fileBaseName + ': 總共 ' + questions.length + ' 題');
        return questions;
    }

    parseFinancialOnePage(lines, pageNum, fileBaseName) {
        var CIRCLE = '\uEA56';
        var ARROW = '\uEA57';

        var awardPattern = /^\d+\.\s*Award:\s*\d+(\.\d+)?\s*point(s)?/i;
        var awardIdx = -1;
        for (var i = 0; i < lines.length; i++) {
            if (awardPattern.test(lines[i])) {
                awardIdx = i;
                break;
            }
        }

        if (awardIdx < 0) {
            console.warn('[Financial][' + fileBaseName + '][p' + pageNum + '] 未找到 Award 行');
            return null;
        }

        var content = lines.slice(awardIdx + 1);

        var endIdx = content.length;
        for (var i = 0; i < content.length; i++) {
            if (/^References$/i.test(content[i]) || /^Multiple Choice/i.test(content[i])) {
                endIdx = i;
                break;
            }
        }

        var questionContent = content.slice(0, endIdx);

        var questionLines = [];
        var optionLines = [];
        var foundFirstOption = false;

        for (var i = 0; i < questionContent.length; i++) {
            var line = questionContent[i];
            if (line.indexOf(CIRCLE) !== -1) {
                foundFirstOption = true;
                if (optionLines.length < 4) {
                    optionLines.push(line);
                }
            } else if (!foundFirstOption) {
                questionLines.push(line);
            }
        }

        var options = [];
        var answerIndex = -1;

        for (var i = 0; i < optionLines.length; i++) {
            var optText = optionLines[i];

            if (optText.indexOf(ARROW) !== -1) {
                answerIndex = i;
            }

            optText = optText.split(CIRCLE).join('');
            optText = optText.split(ARROW).join('');
            optText = optText.trim();

            var letter = String.fromCharCode(97 + i);
            options.push(letter + '. ' + optText);
        }

        if (options.length < 3) {
            console.warn('[Financial][' + fileBaseName + '][p' + pageNum + '] 選項不足: ' + options.length);
            return null;
        }

        if (answerIndex < 0) {
            console.warn('[Financial][' + fileBaseName + '][p' + pageNum + '] 未找到正確答案標記（箭頭）');
            return null;
        }

        var lastOptionIdx = -1;
        for (var i = 0; i < questionContent.length; i++) {
            if (questionContent[i].indexOf(CIRCLE) !== -1) {
                lastOptionIdx = i;
            }
        }

        var feedbackLines = [];
        if (lastOptionIdx >= 0) {
            for (var i = lastOptionIdx + 1; i < questionContent.length; i++) {
                var line = questionContent[i].trim();
                if (line && line.indexOf(CIRCLE) === -1 && line.indexOf(ARROW) === -1) {
                    feedbackLines.push(line);
                }
            }
        }
        var feedbackText = feedbackLines.join(' ').trim();

        var questionText = questionLines.join(' ').replace(/\s+/g, ' ').trim();

        if (!questionText) {
            console.warn('[Financial][' + fileBaseName + '][p' + pageNum + '] 題幹為空');
            return null;
        }

        var correctLetter = String.fromCharCode(97 + answerIndex);

        return {
            originalId: fileBaseName + '-p' + pageNum,
            questionText: questionText,
            options: options,
            correctOption: correctLetter,
            hasCheckmark: true,
            feedbackText: feedbackText
        };
    }

    // 從文字中提取題目
    extractQuestionsFromText(text) {
        const questions = [];
        
        // 尋找 MC 題目的標題行
        // 格式：MC.xx.xx 或 MC.xx.xx.ALGO
        // 允許數字和 MC 之間有任意空白（包括換行）
        const mcHeaderPattern = /\b(\d+\.\s*)?(MC\.\d+\.\d+(?:\.ALGO)?)\b/gi;
        const matches = [...text.matchAll(mcHeaderPattern)];
        
        for (let i = 0; i < matches.length; i++) {
            const match = matches[i];
            const originalId = match[2]; // MC.xx.xx 或 MC.xx.xx.ALGO
            const startPos = match.index;
            
            // 找到下一個題目的開始位置，或文字結尾
            let endPos = text.length;
            if (i < matches.length - 1) {
                endPos = matches[i + 1].index;
            }
            
            const questionBlock = text.substring(startPos, endPos);
            
            const question = this.parseQuestion(questionBlock, originalId);
            if (question) {
                questions.push(question);
            }
        }
        
        return questions;
    }

    // 從文字中提取 EX（非選擇題）
    // 目標：可靠抓到 EX 題目的分段與原始文字（不影響既有 MC 解析）
    extractExerciseQuestionsFromText(text) {
        const exQuestions = [];

        // EX/PR header：EX.03.37、PR.04.12.ALGO 等（非選擇題）
        // 允許前面有題號「12. 」等；並允許 EX/PR 與數字之間有空白或換行
        const exHeaderPattern = /\b(\d+\.\s*)?((?:EX|PR)\.\d+\.\d+(?:\.[A-Z0-9]+)*)\b/gi;

        // 邊界：下一個 MC 或 EX/PR header（用來切分區塊）
        // 注意：MC 的 suffix 只允許 .ALGO（沿用既有規則）
        const boundaryPattern = /\b(\d+\.\s*)?((?:MC\.\d+\.\d+(?:\.ALGO)?)|(?:EX|PR)\.\d+\.\d+(?:\.[A-Z0-9]+)*)\b/gi;
        const boundaries = [...text.matchAll(boundaryPattern)].map(m => ({
            index: m.index,
            id: m[2]
        })).sort((a, b) => a.index - b.index);

        const exMatches = [...text.matchAll(exHeaderPattern)];
        for (let i = 0; i < exMatches.length; i++) {
            const match = exMatches[i];
            const originalId = match[2];
            const startPos = match.index;

            // 找下一個邊界（MC/EX 皆可），或文字結尾
            let endPos = text.length;
            for (let b = 0; b < boundaries.length; b++) {
                if (boundaries[b].index > startPos) {
                    endPos = boundaries[b].index;
                    break;
                }
            }

            const rawBlockText = text.substring(startPos, endPos);
            const q = this.parseExerciseQuestion(rawBlockText, originalId);
            if (q) exQuestions.push(q);
        }

        return exQuestions;
    }

    // 解析單個 EX 題：以 "Required:" 作為題幹/要求分界（若不存在則保留整段）
    parseExerciseQuestion(rawBlockText, originalId) {
        try {
            const escapeRegExp = (s) => String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

            // 移除 header（含可能的前綴題號）
            const headerPatternWithNum = new RegExp(`\\b\\d+\\.\\s*${escapeRegExp(originalId)}\\b\\s*`, 'i');
            const headerPatternNoNum = new RegExp(`\\b${escapeRegExp(originalId)}\\b\\s*`, 'i');

            let body = (rawBlockText || '').replace(headerPatternWithNum, '').trim();
            body = body.replace(headerPatternNoNum, '').trim();

            // 在 EX 區塊內，常見後段包含 Solution/Feedback/Check My Work；先把這些後段切掉（保守）
            // 只做「第一次命中」的截斷，避免破壞題幹內容
            const stopMatch = body.match(/(?:\n|\r|\s)(Solution|Feedback|Check My Work|Post-Submission)\b/i);
            if (stopMatch && stopMatch.index != null && stopMatch.index > 0) {
                body = body.substring(0, stopMatch.index).trim();
            }

            // Required: 分段（題幹通常在 Required: 之後）
            let promptText = '';
            let requiredText = '';
            const requiredMatch = body.match(/\bRequired\s*:/i);
            if (requiredMatch && requiredMatch.index != null) {
                promptText = body.substring(0, requiredMatch.index).trim();
                requiredText = body.substring(requiredMatch.index).replace(/\bRequired\s*:/i, '').trim();
            } else {
                // 若找不到 Required:，先把整段放在 requiredText（較符合「敘述/要求」在下方的實務）
                requiredText = body.trim();
            }

            // 嘗試抓答案線索（不追求完美；先保留含 ✔/✓ 或底線的行作為 token）
            const tokens = [];
            const lines = (body || '').split('\n').map(l => (l || '').trim()).filter(Boolean);
            for (const line of lines) {
                if (/[✔✓]/.test(line) || /_{3,}/.test(line)) {
                    tokens.push(line);
                }
            }

            return {
                originalId,
                type: (originalId || '').toUpperCase().startsWith('PR') ? 'PR' : 'EX',
                promptText: promptText,
                requiredText: requiredText,
                answerTextOrTokens: tokens.join('\n').trim(),
                rawBlockText: rawBlockText
            };
        } catch (e) {
            console.error('解析 EX 題目失敗:', e);
            return null;
        }
    }

    // 解析單個題目
    parseQuestion(text, originalId) {
        try {
            // 移除標題行（MC.xx.xx 或 MC.xx.xx.ALGO）
            const headerPattern = /\b\d+\.\s*MC\.\d+\.\d+(?:\.ALGO)?\b\s*/i;
            let questionText = text.replace(headerPattern, '').trim();
            
            // 移除標題行（如果沒有數字前綴）
            questionText = questionText.replace(new RegExp(`\\b${originalId}\\b\\s*`, 'i'), '').trim();
            
            // 找到選項開始的位置
            // 選項可能是 "a." 或前面有 checkmark "✔ a." 或 "✓ a."
            const optionStartMatch = questionText.match(/(?:^|\n)\s*(?:[✔✓]\s*)?a\.\s/im);
            if (!optionStartMatch) {
                return null; // 沒有找到選項
            }
            const optionStartIndex = optionStartMatch.index;
            
            // 提取題目文字（從標題後到選項 "a." 之前）
            let questionTextOnly = questionText.substring(0, optionStartIndex).trim();
            
            // 清理題目文字：移除多餘空白和換行
            questionTextOnly = questionTextOnly.replace(/\s+/g, ' ').trim();
            
            // 提取選項部分（從 "a." 開始）
            const optionsText = questionText.substring(optionStartIndex);
            
            // 提取選項（a. 到 e.）
            const options = [];
            let correctOption = null;
            let hasCheckmark = false;
            
            // 逐行解析選項
            const lines = optionsText.split('\n');
            let currentOption = null;
            let currentText = '';
            let currentHasCheck = false;
            
            for (const line of lines) {
                const trimmed = line.trim();
                if (!trimmed) continue;
                
                // 檢查是否是新選項行（a. 到 e.）
                const newOptionMatch = trimmed.match(/^([✔✓])?\s*([a-e])\.\s*(.*)$/i);
                if (newOptionMatch) {
                    // 保存之前的選項
                    if (currentOption !== null) {
                        const optText = currentText.replace(/\s+/g, ' ').trim();
                        if (currentHasCheck) {
                            options.push(`${currentOption}. ${optText} ✔`);
                            correctOption = currentOption;
                            hasCheckmark = true;
                        } else {
                            options.push(`${currentOption}. ${optText}`);
                        }
                    }
                    
                    // 開始新選項
                    currentHasCheck = !!newOptionMatch[1];
                    currentOption = newOptionMatch[2].toLowerCase();
                    currentText = newOptionMatch[3] || '';
                } else if (currentOption !== null) {
                    // 檢查是否到了結束標記
                    if (/^(Feedback|Check My Work|Post-Submission|Solution)/i.test(trimmed)) {
                        break;
                    }
                    // 續行
                    currentText += ' ' + trimmed;
                }
            }
            
            // 保存最後一個選項
            if (currentOption !== null) {
                const optText = currentText.replace(/\s+/g, ' ').trim();
                if (currentHasCheck) {
                    options.push(`${currentOption}. ${optText} ✔`);
                    correctOption = currentOption;
                    hasCheckmark = true;
                } else {
                    options.push(`${currentOption}. ${optText}`);
                }
            }
            
            // 驗證必須有至少 2 個選項（有些題目可能只有 a-d）
            if (options.length < 2) {
                return null;
            }
            
            // 如果沒有找到 checkmark，嘗試從 Solution 行提取（備用）
            if (!correctOption) {
                const solutionMatch = text.match(/Solution\s+([a-e])/i);
                if (solutionMatch) {
                    correctOption = solutionMatch[1].toLowerCase();
                } else {
                    return null; // 沒有找到正確答案
                }
            }
            
            // 提取 Feedback（在 Feedback 標題之後，但排除 Post-Submission）
            let feedbackText = '';
            const feedbackMatch = text.match(/Feedback\s+(.+?)(?=Post-Submission|Check My Work|Solution|$)/is);
            if (feedbackMatch) {
                feedbackText = feedbackMatch[1].trim();
                // 移除 Post-Submission 如果存在
                feedbackText = feedbackText.replace(/Post-Submission.*$/is, '').trim();
            }
            
            return {
                originalId: originalId,
                questionText: questionTextOnly,
                options: options,
                correctOption: correctOption,
                hasCheckmark: hasCheckmark,
                feedbackText: feedbackText
            };
        } catch (error) {
            console.error('解析題目失敗:', error);
            return null;
        }
    }
}

// ========== 隨機抽題生成器類別 ==========
class QuestionGenerator {
    constructor(questions) {
        this.allQuestions = questions;
    }

    // 解析章節配比設定
    parseChapterRatio(ratioText) {
        const ratios = {};
        const lines = ratioText.trim().split('\n');
        
        for (const line of lines) {
            const trimmed = line.trim();
            if (!trimmed) continue;
            
            const match = trimmed.match(/^(.+?):\s*(\d+)$/);
            if (match) {
                const chapter = match[1].trim();
                const count = parseInt(match[2], 10);
                ratios[chapter] = count;
            }
        }
        
        return ratios;
    }

    // 隨機抽題
    generateExam(totalQuestions, chapterRatioText) {
        let selectedQuestions = [];
        
        // 如果有指定章節配比
        if (chapterRatioText && chapterRatioText.trim()) {
            const ratios = this.parseChapterRatio(chapterRatioText);
            const ratioTotal = Object.values(ratios).reduce((sum, val) => sum + val, 0);
            
            // 如果配比總數與總題數不一致，按比例調整
            if (ratioTotal !== totalQuestions && ratioTotal > 0) {
                const scale = totalQuestions / ratioTotal;
                for (const chapter in ratios) {
                    ratios[chapter] = Math.round(ratios[chapter] * scale);
                }
            }
            
            // 從各章節抽題（這裡簡化處理，因為我們沒有章節資訊）
            // 如果沒有章節資訊，就完全隨機
            selectedQuestions = this.randomSelect(this.allQuestions, totalQuestions);
        } else {
            // 沒有指定配比，完全隨機
            selectedQuestions = this.randomSelect(this.allQuestions, totalQuestions);
        }
        
        // 打亂順序
        selectedQuestions = this.shuffle(selectedQuestions);
        
        // 重新編號（1, 2, 3...）
        selectedQuestions.forEach((q, index) => {
            q.examNumber = index + 1;
        });
        
        return selectedQuestions;
    }

    // 從陣列中隨機選擇 n 個元素
    randomSelect(array, n) {
        if (n >= array.length) {
            return this.shuffle([...array]);
        }
        
        const shuffled = this.shuffle([...array]);
        return shuffled.slice(0, n);
    }

    // 打亂陣列順序
    shuffle(array) {
        const result = [...array];
        for (let i = result.length - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));
            [result[i], result[j]] = [result[j], result[i]];
        }
        return result;
    }
}

// ========== Word 非選擇題題庫解析 ==========
const WORD_Q_START_REGEX = /^\s*(\d{1,4})\.\s+/;
const WORD_ANSWER_MARKER = 'ANSWER:';

function parseWordQuestions(lines) {
    const questions = [];
    const starts = [];
    for (let i = 0; i < lines.length; i++) {
        const m = String(lines[i] || '').match(WORD_Q_START_REGEX);
        if (m) starts.push({ index: i, id: m[1] });
    }
    for (let s = 0; s < starts.length; s++) {
        const from = starts[s].index;
        const to = s < starts.length - 1 ? starts[s + 1].index : lines.length;
        const rawLines = lines.slice(from, to).map(l => (l == null ? '' : String(l)));
        const originalId = starts[s].id;
        let questionLines = [];
        let answerLines = [];
        let hasAnswerSection = false;
        for (let i = 0; i < rawLines.length; i++) {
            const idx = rawLines[i].indexOf(WORD_ANSWER_MARKER);
            if (idx !== -1) {
                hasAnswerSection = true;
                const before = rawLines[i].substring(0, idx).replace(/\s+$/, '');
                if (before) questionLines.push(before);
                const after = rawLines[i].substring(idx + WORD_ANSWER_MARKER.length).replace(/^\s+/, '');
                if (after) answerLines.push(after);
                for (let j = i + 1; j < rawLines.length; j++) {
                    answerLines.push(rawLines[j]);
                }
                break;
            }
            questionLines.push(rawLines[i]);
        }
        if (!hasAnswerSection) {
            questionLines = rawLines;
        }
        questions.push({
            source: 'word',
            originalId,
            questionLines,
            answerLines,
            hasAnswerSection,
            rawLines
        });
    }
    return questions;
}

async function parseWordFile(file) {
    if (!file || !file.name.toLowerCase().endsWith('.docx')) return;
    wordParseState = 'parsing';
    wordFile = file;
    wordNonMcQuestions = [];
    updateWordParseUI();
    try {
        const arrayBuffer = await file.arrayBuffer();
        const result = await mammoth.convertToHtml({ arrayBuffer });
        const html = result.value || '';
        const div = document.createElement('div');
        div.innerHTML = html;
        const paragraphs = div.querySelectorAll('p');
        const lines = Array.from(paragraphs).map(p => (p.textContent || '').replace(/\s+$/g, ''));
        wordNonMcQuestions = parseWordQuestions(lines);
        wordParseState = 'parsed';
    } catch (e) {
        wordParseState = 'error';
        console.error('Word parse error:', e);
    }
    updateWordParseUI();
}

function updateWordParseUI() {
    if (!wordParseStatus) return;
    if (wordParseState === 'idle') {
        wordParseStatus.textContent = '';
        if (wordConfig) wordConfig.style.display = 'none';
        return;
    }
    if (wordParseState === 'parsing') {
        wordParseStatus.textContent = 'Parsing...';
        wordParseStatus.style.color = '#667eea';
        if (wordConfig) wordConfig.style.display = 'none';
        return;
    }
    if (wordParseState === 'error') {
        wordParseStatus.textContent = 'Error';
        wordParseStatus.style.color = '#c00';
        if (wordConfig) wordConfig.style.display = 'none';
        return;
    }
    wordParseStatus.textContent = 'Parsed';
    wordParseStatus.style.color = '#0a0';
    if (wordFileNameSpan) wordFileNameSpan.textContent = wordFile ? wordFile.name : '';
    if (wordAvailableSpan) wordAvailableSpan.textContent = String(wordNonMcQuestions.length);
    if (wordConfig) wordConfig.style.display = 'block';
    if (wordRequestedInput) {
        const n = wordNonMcQuestions.length;
        wordRequestedInput.max = n;
        let v = parseInt(wordRequestedInput.value, 10);
        if (isNaN(v) || v < 0) v = 0;
        if (v > n) v = n;
        wordRequestedInput.value = String(v);
    }
}

// ========== Word 文檔生成器類別 ==========
class WordGenerator {
    // Helper：建立封面頁表格的置中且字體放大的 TableCell
    _buildCoverTableCell(text, isBold = false) {
        const fontSize = 26; // 封面頁表格字體大小（明顯放大）
        return new docx.TableCell({
            verticalAlignment: docx.VerticalAlign.CENTER,
            children: [
                new docx.Paragraph({
                    alignment: docx.AlignmentType.CENTER,
                    children: [
                        new docx.TextRun({
                            text: text || '',
                            size: fontSize,
                            bold: isBold
                        })
                    ]
                })
            ]
        });
    }

    // Helper：建立封面頁表格的空白置中 TableCell（用於 Score 欄或答案格）
    _buildCoverEmptyCell() {
        return new docx.TableCell({
            verticalAlignment: docx.VerticalAlign.CENTER,
            children: [
                new docx.Paragraph({
                    alignment: docx.AlignmentType.CENTER,
                    children: []
                })
            ]
        });
    }

    // Managerial 專用：產生封面頁元素（含答案格），傳入題目總數以決定格數
    _buildManagerialCoverPage(questionCount, examName, points) {
        // 安全 fallback：如果沒有傳入 points 或 rows 不存在，使用預設值
        const defaultRows = [
            { label: "I", value: 150 },
            { label: "II", value: 25 },
            { label: "III", value: 25 }
        ];
        const defaultTotal = 200;
        
        // 優先使用 points.rows，否則 fallback 到預設
        let ptsRows = (points && points.rows && Array.isArray(points.rows) && points.rows.length > 0) 
            ? points.rows 
            : defaultRows;
        const ptsTotal = (points && points.total !== undefined && points.total !== null) 
            ? points.total 
            : defaultTotal;
        
        // 封面頁表格行高設定（單位：twips，1 twip = 1/20 point）
        const COVER_POINTS_ROW_HEIGHT = 500; // Points 表格行高（約 25 points，明顯變高方便手寫）
        const ANSWER_GRID_ROW_HEIGHT = 550;  // Answer Grid 行高（比 Points 稍高，確保手寫空間）
        
        const out = [];
        const borderOption = (typeof docx.BorderStyle !== 'undefined')
            ? { style: docx.BorderStyle.SINGLE, size: 4 }
            : { size: 4 };

        out.push(
            new docx.Paragraph({
                children: [new docx.TextRun({ text: 'Department of Accounting and ISA', size: 22 })],
                alignment: docx.AlignmentType.CENTER,
                spacing: { after: 80 }
            }),
            new docx.Paragraph({
                children: [new docx.TextRun({ text: 'Shippensburg University', size: 22 })],
                alignment: docx.AlignmentType.CENTER,
                spacing: { after: 80 }
            }),
            new docx.Paragraph({
                children: [new docx.TextRun({ text: examName || 'Spring 2025 – Exam 2', size: 22 })],
                alignment: docx.AlignmentType.CENTER,
                spacing: { after: 320 }
            }),
            new docx.Paragraph({
                children: [new docx.TextRun({ text: 'Section: ___________        Name: __________________________', size: 22 })],
                spacing: { after: 320 }
            }),
            new docx.Paragraph({
                children: [new docx.TextRun({
                    text: '• For open-ended questions, you must show all supporting calculations in an ORGANIZED way to be eligible to receive total credit. If not, you\'ll not get partial credit.',
                    size: 20
                })],
                spacing: { after: 120 }
            }),
            new docx.Paragraph({
                children: [new docx.TextRun({
                    text: '• For open-ended question, you will receive minimum or zero points if you only show answers without any supporting calculations.',
                    size: 20
                })],
                spacing: { after: 120 }
            }),
            new docx.Paragraph({
                children: [new docx.TextRun({
                    text: '• You must transfer your answers for the multiple choice questions to this coversheet below. If not, you\'ll lose 10 points.',
                    size: 20
                })],
                spacing: { after: 240 }
            })
        );

        // 動態生成 Points 表格列
        const pointsTableRows = [
            // 表頭
            new docx.TableRow({ 
                height: { value: COVER_POINTS_ROW_HEIGHT, rule: docx.HeightRule.EXACT },
                children: [
                    this._buildCoverTableCell('Parts', true),
                    this._buildCoverTableCell('Points', true),
                    this._buildCoverTableCell('Score', true)
                ] 
            })
        ];
        
        // 動態生成每一列（依 ptsRows）
        ptsRows.forEach(row => {
            const label = row.label || '';
            const value = (row.value !== undefined && row.value !== null) ? row.value : 0;
            pointsTableRows.push(
                new docx.TableRow({ 
                    height: { value: COVER_POINTS_ROW_HEIGHT, rule: docx.HeightRule.EXACT },
                    children: [
                        this._buildCoverTableCell(label + '.'),
                        this._buildCoverTableCell(String(value)),
                        this._buildCoverEmptyCell()
                    ] 
                })
            );
        });
        
        // Total Points 列
        pointsTableRows.push(
            new docx.TableRow({ 
                height: { value: COVER_POINTS_ROW_HEIGHT, rule: docx.HeightRule.EXACT },
                children: [
                    this._buildCoverTableCell('Total Points'),
                    this._buildCoverTableCell(String(ptsTotal)),
                    this._buildCoverEmptyCell()
                ] 
            })
        );
        out.push(
            new docx.Table({
                rows: pointsTableRows,
                width: { size: 40, type: docx.WidthType.PERCENTAGE }
            })
        );

        out.push(
            new docx.Paragraph({
                children: [new docx.TextRun({
                    text: 'An important aspect of being a professional is that of honor and professional conduct.',
                    size: 20
                })],
                spacing: { before: 280, after: 120 }
            }),
            new docx.Paragraph({
                children: [new docx.TextRun({
                    text: 'I pledge that I have neither given nor received aid in the completion of this examination.',
                    size: 20
                })],
                spacing: { after: 120 }
            }),
            new docx.Paragraph({
                children: [new docx.TextRun({ text: '____________________________________________  LEGIBLE SIGNATURE', size: 20 })],
                spacing: { after: 320 }
            }),
            new docx.Paragraph({
                children: [new docx.TextRun({ text: 'I. MULTIPLE CHOICE', bold: true, size: 22 })],
                spacing: { after: 120 }
            }),
            new docx.Paragraph({
                children: [new docx.TextRun({ text: 'Write the letter that best answers each question.', size: 20 })],
                spacing: { after: 200 }
            })
        );

        const cols = 10;
        const blockCount = Math.ceil(questionCount / cols);
        const gridRows = [];
        const coverGridFontSize = 26; // Answer Grid 字體大小（與 Points 表格一致）
        const questionRowShading = { fill: "EDEDED" }; // 題號列淺灰色背景
        for (let b = 0; b < blockCount; b++) {
            // Q. row（題號列）：加上淺灰背景和粗體文字
            const qCells = [new docx.TableCell({
                verticalAlignment: docx.VerticalAlign.CENTER,
                shading: questionRowShading,
                children: [
                    new docx.Paragraph({
                        alignment: docx.AlignmentType.CENTER,
                        children: [new docx.TextRun({ text: 'Q.', size: coverGridFontSize, bold: true })]
                    })
                ],
                width: { size: 8, type: docx.WidthType.PERCENTAGE }
            })];
            // A. row（答案列）：維持原樣，不加背景色或粗體
            const aCells = [new docx.TableCell({
                verticalAlignment: docx.VerticalAlign.CENTER,
                children: [
                    new docx.Paragraph({
                        alignment: docx.AlignmentType.CENTER,
                        children: [new docx.TextRun({ text: 'A.', size: coverGridFontSize })]
                    })
                ],
                width: { size: 8, type: docx.WidthType.PERCENTAGE }
            })];
            for (let c = 0; c < cols; c++) {
                const num = b * cols + c + 1;
                // Q. row 的題號 cell：加上淺灰背景和粗體文字
                qCells.push(new docx.TableCell({
                    verticalAlignment: docx.VerticalAlign.CENTER,
                    shading: questionRowShading,
                    children: [
                        new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                                new docx.TextRun({
                                    text: num <= questionCount ? String(num) : '',
                                    size: coverGridFontSize,
                                    bold: true
                                })
                            ]
                        })
                    ],
                    width: { size: (92 / cols), type: docx.WidthType.PERCENTAGE }
                }));
                // A. row 的答案格：維持原樣，不加背景色
                aCells.push(new docx.TableCell({
                    verticalAlignment: docx.VerticalAlign.CENTER,
                    children: [
                        new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: []
                        })
                    ],
                    width: { size: (92 / cols), type: docx.WidthType.PERCENTAGE }
                }));
            }
            gridRows.push(
                new docx.TableRow({ 
                    height: { value: ANSWER_GRID_ROW_HEIGHT, rule: docx.HeightRule.EXACT },
                    children: qCells 
                }),
                new docx.TableRow({ 
                    height: { value: ANSWER_GRID_ROW_HEIGHT, rule: docx.HeightRule.EXACT },
                    children: aCells 
                })
            );
        }
        out.push(
            new docx.Table({
                rows: gridRows,
                width: { size: 100, type: docx.WidthType.PERCENTAGE },
                borders: {
                    top: borderOption,
                    bottom: borderOption,
                    left: borderOption,
                    right: borderOption,
                    insideHorizontal: borderOption,
                    insideVertical: borderOption
                }
            })
        );
        out.push(new docx.Paragraph({ children: [new docx.PageBreak()], spacing: { after: 0 } }));
        return out;
    }

    // Answer Sheet 專用：由左到右、每列 perRow 題的答案摘要格（與題目卷答案格同版型）
    _buildHorizontalAnswerGrid(questions, perRow) {
        perRow = perRow || 10;
        const n = questions.length;
        const blockCount = Math.ceil(n / perRow);
        const gridRows = [];
        for (let b = 0; b < blockCount; b++) {
            const qCells = [new docx.TableCell({
                children: [new docx.Paragraph({ children: [new docx.TextRun({ text: 'Q.', size: 18 })] })],
                width: { size: 8, type: docx.WidthType.PERCENTAGE }
            })];
            const aCells = [new docx.TableCell({
                children: [new docx.Paragraph({ children: [new docx.TextRun({ text: 'A.', size: 18 })] })],
                width: { size: 8, type: docx.WidthType.PERCENTAGE }
            })];
            for (let c = 0; c < perRow; c++) {
                const idx = b * perRow + c;
                const num = idx + 1;
                const label = num <= n ? String(num) : '';
                qCells.push(new docx.TableCell({
                    children: [new docx.Paragraph({
                        children: [new docx.TextRun({ text: label, size: 18 })],
                        alignment: docx.AlignmentType.CENTER
                    })],
                    width: { size: (92 / perRow), type: docx.WidthType.PERCENTAGE }
                }));
                const ans = (idx < n && questions[idx].correctOption) ? questions[idx].correctOption.toUpperCase() : '';
                aCells.push(new docx.TableCell({
                    children: [new docx.Paragraph({
                        children: [new docx.TextRun({ text: ans, size: 18 })],
                        alignment: docx.AlignmentType.CENTER
                    })],
                    width: { size: (92 / perRow), type: docx.WidthType.PERCENTAGE }
                }));
            }
            gridRows.push(
                new docx.TableRow({ children: qCells }),
                new docx.TableRow({ children: aCells })
            );
        }
        return gridRows;
    }

    // 生成題目卷（學生用）
    async generateQuestionSheet(examName, questions, points, exSelectedAll = [], wordNonMcSelected = []) {
        // 所有內容將添加到同一個 section，確保連續流動
        const allChildren = [];

        // 將多行文字（含 \n）轉成逐行 Paragraph（保留空白行）
        // 僅供題目卷 EX 題排版使用，避免影響既有 MC / Financial 行為
        const toParagraphsByLine = (text, paragraphOptions) => {
            const opts = paragraphOptions || {};
            const normalized = (text == null ? '' : String(text)).replace(/\r\n/g, '\n').replace(/\r/g, '\n');
            const lines = normalized.split('\n');
            const out = [];
            for (let i = 0; i < lines.length; i++) {
                const line = lines[i];
                if (line === '') {
                    out.push(new docx.Paragraph({ children: [], ...opts }));
                } else {
                    out.push(new docx.Paragraph({
                        children: [new docx.TextRun({ text: line.replace(/\s+$/g, ''), size: 20 })],
                        ...opts
                    }));
                }
            }
            return out;
        };

        if (currentSubject === 'managerial') {
            allChildren.push(...this._buildManagerialCoverPage(questions.length, examName, points));
        }
        
        // 1. 標題
        allChildren.push(
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: examName,
                        bold: true,
                        size: 32
                    })
                ],
                alignment: docx.AlignmentType.CENTER,
                spacing: { after: 400 }
            })
        );

        // 總題數／作答檢查表／勾選格表：僅 Financial 保留；Managerial 已有封面作答格，題目頁不再重複
        if (currentSubject !== 'managerial') {
            allChildren.push(
                new docx.Paragraph({
                    children: [new docx.TextRun({ text: `總題數：${questions.length} 題` })],
                    alignment: docx.AlignmentType.CENTER,
                    spacing: { after: 600 }
                })
            );
            const cols = 5;
            const rows = Math.ceil(questions.length / cols);
            const tableRows = [];
            for (let row = 0; row < rows; row++) {
                const cells = [];
                for (let col = 0; col < cols; col++) {
                    const index = row * cols + col;
                    if (index < questions.length) {
                        cells.push(
                            new docx.TableCell({
                                children: [
                                    new docx.Paragraph({
                                        children: [new docx.TextRun({ text: `${index + 1}`, size: 18 })],
                                        alignment: docx.AlignmentType.CENTER
                                    }),
                                    new docx.Paragraph({
                                        children: [new docx.TextRun({ text: '⬜', size: 20 })],
                                        alignment: docx.AlignmentType.CENTER
                                    })
                                ],
                                width: { size: 20, type: docx.WidthType.PERCENTAGE }
                            })
                        );
                    } else {
                        cells.push(new docx.TableCell({ children: [], width: { size: 20, type: docx.WidthType.PERCENTAGE } }));
                    }
                }
                tableRows.push(new docx.TableRow({ children: cells }));
            }
            allChildren.push(
                new docx.Paragraph({
                    children: [new docx.TextRun({ text: '作答檢查表', bold: true, size: 24 })],
                    spacing: { after: 200 }
                })
            );
            allChildren.push(new docx.Table({ rows: tableRows, width: { size: 100, type: docx.WidthType.PERCENTAGE } }));
            allChildren.push(new docx.Paragraph({ children: [], spacing: { after: 400 } }));
        }

        // 題目內容（移除所有標記和原始 ID）
        questions.forEach((q, index) => {
            // 清理題目文字：移除附錄標籤和 (Algorithmic)（僅在題目卷中）
            // 移除 (Appendix...) 格式的標籤，包括 (Appendix 4B), (Appendix A) 等
            let cleanedQuestionText = q.questionText.replace(/\(Appendix[^)]*\)/gi, '');
            // 移除 (Algorithmic) 標籤
            cleanedQuestionText = cleanedQuestionText.replace(/\(Algorithmic\)/gi, '');
            // 清理可能留下的多餘空格
            cleanedQuestionText = cleanedQuestionText.trim().replace(/\s+/g, ' ');
            
            // 題目編號和文字（格式：1. 題目文字）
            allChildren.push(
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: `${index + 1}. ${cleanedQuestionText}`,
                            size: 22
                        })
                    ],
                    spacing: { before: index === 0 ? 0 : 300, after: 200 }
                })
            );

            // 選項（移除 ✔ 和 ✓ 標記）
            q.options.forEach(function(option) {
                var cleanOption = option.replace(/[✔✓]/g, '').trim();
                allChildren.push(
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({
                                text: cleanOption,
                                size: 20
                            })
                        ],
                        spacing: { after: 100 },
                        indent: { left: 400 }
                    })
                );
            });
        });

        // EX 區塊（僅 Managerial Accounting，且要有 EX 題目）
        if (currentSubject === 'managerial' && exSelectedAll && exSelectedAll.length > 0) {
            // EX 區塊標題
            allChildren.push(
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: 'II. EXERCISES',
                            bold: true,
                            size: 22
                        })
                    ],
                    spacing: { before: 400, after: 200 }
                })
            );

            // EX 題目清理 Helper 函數（僅用於題目卷 EX 輸出）
            const sanitizeLine = (line) => {
                if (!line) return '';
                let cleaned = String(line);
                // normalize spaces（保留必要縮排，只 trimEnd）
                cleaned = cleaned.replace(/\s+$/g, '');
                // 移除 "(Algorithmic)"（大小寫不敏感）
                cleaned = cleaned.replace(/\(Algorithmic\)/gi, '');
                // 移除 check marks ✔/✓
                cleaned = cleaned.replace(/[✔✓]/g, '');
                return cleaned;
            };

            const isUnderlineLine = (line) => {
                if (!line) return false;
                return /^\s*_+\s*$/.test(String(line));
            };

            const stripAnswerBlocksFromRequiredLines = (requiredLines, originalLinesForCheck) => {
                const outputLines = [];
                const originalLines = originalLinesForCheck || [];
                
                // 答案區塊判斷函數
                const isDollarLine = (line) => /^\s*\$\s*(.*)?$/.test(line);
                const isPureNumber = (line) => /^\s*[-(]?\s*[\d,]+(\.\d+)?\s*\)?\s*$/.test(line);
                const isUnitLine = (line) => /^\s*(per unit|%|units?|DLH|hours?)\s*$/i.test(line);
                const isAnswerLabel = (line) => /^\s*(Direct labor|Direct materials|Answer|Solution|Feedback|Post-Submission)\s*$/i.test(line);
                const hasCheckOriginal = (idx) => {
                    if (idx < 0 || idx >= originalLines.length) return false;
                    return /[✔✓]/.test(String(originalLines[idx]));
                };

                let i = 0;
                while (i < requiredLines.length) {
                    const line = String(requiredLines[i] || '').trim();
                    const cleanLine = sanitizeLine(line);
                    
                    // 1. 先檢查是否為底線行，直接跳過（不輸出）
                    if (isUnderlineLine(line)) {
                        i++;
                        continue;
                    }

                    // 2. 檢查是否進入答案區塊
                    let isAnswerBlock = false;
                    
                    // 觸發條件 a: isDollarLine
                    if (isDollarLine(cleanLine)) {
                        isAnswerBlock = true;
                    }
                    // 觸發條件 b: hasCheckOriginal
                    else if (hasCheckOriginal(i)) {
                        isAnswerBlock = true;
                    }
                    // 觸發條件 c: isAnswerLabel 且接下來 1~3 行內出現答案特徵
                    else if (isAnswerLabel(cleanLine)) {
                        for (let lookahead = 1; lookahead <= 3 && (i + lookahead) < requiredLines.length; lookahead++) {
                            const nextLine = String(requiredLines[i + lookahead] || '').trim();
                            const nextCleanLine = sanitizeLine(nextLine);
                            if (isDollarLine(nextCleanLine) || isPureNumber(nextCleanLine) || hasCheckOriginal(i + lookahead)) {
                                isAnswerBlock = true;
                                break;
                            }
                        }
                    }

                    // 3. 如果進入答案區塊，連續吃掉答案行直到遇到題目文字
                    if (isAnswerBlock) {
                        while (i < requiredLines.length) {
                            const currentLine = String(requiredLines[i] || '').trim();
                            const currentCleanLine = sanitizeLine(currentLine);
                            
                            // 繼續吃答案行的條件
                            if (isDollarLine(currentCleanLine) ||
                                hasCheckOriginal(i) ||
                                isPureNumber(currentCleanLine) ||
                                isUnitLine(currentCleanLine) ||
                                isAnswerLabel(currentCleanLine) ||
                                isUnderlineLine(currentLine) ||
                                currentLine === '') {
                                i++;
                                continue;
                            }
                            
                            // 遇到題目文字，結束答案區塊
                            break;
                        }
                        continue;
                    }

                    // 4. 非答案區塊的行，輸出（但先清理）
                    if (cleanLine) {
                        outputLines.push(cleanLine);
                    } else if (line === '') {
                        // 保留空白行（但不在答案區塊中）
                        outputLines.push('');
                    }
                    i++;
                }

                return outputLines;
            };

            // 輸出每個 EX 題目
            exSelectedAll.forEach((ex, index) => {
                // 題號（延續 MC 題號）
                const exNumber = questions.length + index + 1;

                // 逐行輸出：prompt/stem + Required: + requiredText（保留 PDF 的換行）
                const cleanExText = (t) => {
                    let s = (t == null ? '' : String(t));
                    // 1) 移除 ✔/✓
                    s = s.replace(/[✔✓]/g, '');
                    // 2) 移除 Solution/Feedback/Check My Work/Post-Submission 等後段內容（保守，避免把整段切掉太多）
                    s = s.replace(/Solution\b[\s\S]*$/i, '');
                    s = s.replace(/Feedback\b[\s\S]*$/i, '');
                    s = s.replace(/Check My Work\b[\s\S]*$/i, '');
                    s = s.replace(/Post-Submission\b[\s\S]*$/i, '');
                    // 3) 移除原始 EX 編號
                    s = s.replace(/\bEX\.\d+\.\d+(?:\.[A-Z0-9]+)*\b/gi, '');
                    return s.trim();
                };

                // 處理 promptText：移除 "(Algorithmic)" 並逐行清理
                let promptText = ex.promptText || '';
                const promptLines = promptText ? promptText.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n') : [];
                const cleanedPromptLines = promptLines.map(line => sanitizeLine(line));
                const firstPromptLine = cleanedPromptLines.length > 0 ? cleanedPromptLines[0] : '';
                const remainingPromptLines = cleanedPromptLines.length > 1 ? cleanedPromptLines.slice(1) : [];

                // 先印出 EX 題號（用自己的連續題號；不印原始 EX 編號）
                allChildren.push(
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({
                                text: firstPromptLine ? `${exNumber}. ${firstPromptLine}` : `${exNumber}.`,
                                size: 22
                            })
                        ],
                        spacing: { before: index === 0 ? 0 : 300, after: 80 }
                    })
                );

                // prompt/stem（續行逐行輸出，已清理 "(Algorithmic)"）
                if (remainingPromptLines.length > 0) {
                    remainingPromptLines.forEach(line => {
                        if (line === '') {
                            allChildren.push(new docx.Paragraph({ children: [], spacing: { after: 80 }, indent: { left: 400 } }));
                        } else {
                            allChildren.push(
                                new docx.Paragraph({
                                    children: [new docx.TextRun({ text: line, size: 20 })],
                                    spacing: { after: 80 },
                                    indent: { left: 400 }
                                })
                            );
                        }
                    });
                }

                // Required:（單獨一行）
                if (ex.requiredText) {
                    // 準備原始行（用於檢查 ✔）
                    const originalRequiredLines = (ex.requiredText || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
                    const requiredLines = (ex.requiredText || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
                    
                    // 使用 stripAnswerBlocksFromRequiredLines 清理 Required 區塊
                    const cleanedRequiredLines = stripAnswerBlocksFromRequiredLines(requiredLines, originalRequiredLines);

                    // 只輸出清理後的行（過濾掉空行，除非是必要的段落分隔）
                    if (cleanedRequiredLines.length > 0) {
                        allChildren.push(
                            new docx.Paragraph({
                                children: [
                                    new docx.TextRun({
                                        text: 'Required:',
                                        bold: true,
                                        size: 20
                                    })
                                ],
                                spacing: { after: 50 },
                                indent: { left: 400 }
                            })
                        );

                        // requiredText（逐行輸出，已清理答案和底線）
                        cleanedRequiredLines.forEach(line => {
                            if (line === '') {
                                allChildren.push(new docx.Paragraph({ children: [], spacing: { after: 80 }, indent: { left: 400 } }));
                            } else {
                                allChildren.push(
                                    new docx.Paragraph({
                                        children: [new docx.TextRun({ text: line, size: 20 })],
                                        spacing: { after: 80 },
                                        indent: { left: 400 }
                                    })
                                );
                            }
                        });
                    }
                }
                
                // 作答區：用底線取代答案位置
                // 如果有 answerTextOrTokens，表示原本有答案位置，用底線取代
                // 否則在題目最後加 2-4 行底線當作答空間
                const answerLines = 3; // 預設 3 行作答空間
                for (let i = 0; i < answerLines; i++) {
                    allChildren.push(
                        new docx.Paragraph({
                            children: [
                                new docx.TextRun({
                                    text: '__________',
                                    size: 20
                                })
                            ],
                            spacing: { after: 100 },
                            indent: { left: 400 }
                        })
                    );
                }
            });
        }

        // Word 非選擇題區塊（僅 Managerial，放在 MC/EX 之後）
        if (currentSubject === 'managerial' && wordNonMcSelected && wordNonMcSelected.length > 0) {
            allChildren.push(
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: 'II. NON-MULTIPLE-CHOICE (Word)',
                            bold: true,
                            size: 22
                        })
                    ],
                    spacing: { before: 400, after: 200 }
                })
            );
            wordNonMcSelected.forEach((w, idx) => {
                if (idx > 0) {
                    allChildren.push(new docx.Paragraph({ children: [], spacing: { before: 300, after: 0 } }));
                }
                const qLines = w.questionLines || [];
                const lineOpts = { spacing: { after: 80 } };
                if (qLines.length > 0) {
                    const text = qLines.map(l => (l == null ? '' : String(l))).join('\n');
                    allChildren.push(...toParagraphsByLine(text, lineOpts));
                } else {
                    allChildren.push(
                        new docx.Paragraph({
                            children: [new docx.TextRun({ text: `${w.originalId}.`, size: 22 })],
                            ...lineOpts
                        })
                    );
                }
                const underlineCount = 3;
                for (let i = 0; i < underlineCount; i++) {
                    allChildren.push(
                        new docx.Paragraph({
                            children: [new docx.TextRun({ text: '__________', size: 20 })],
                            spacing: { after: 100 }
                        })
                    );
                }
            });
        }

        // 創建單一 section，包含所有內容（標題、表格、題目、EX）
        const doc = new docx.Document({
            sections: [{
                properties: {},
                children: allChildren
            }]
        });

        const blob = await docx.Packer.toBlob(doc);
        return blob;
    }

    // 生成答案卷（教師用）
    async generateAnswerSheet(examName, questions, exSelectedAll = [], wordNonMcSelected = []) {
        // 所有內容將添加到同一個 section，確保連續流動
        const allChildren = [];
        
        // 標題
        allChildren.push(
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: `${examName} - Answer Sheet`,
                        bold: true,
                        size: 32
                    })
                ],
                alignment: docx.AlignmentType.CENTER,
                spacing: { after: 400 }
            })
        );

        // 答案摘要表格（在詳細題目之前）：由左到右、每列 10 題，與題目卷答案格同版型
        const summaryTableRows = this._buildHorizontalAnswerGrid(questions, 10);

        // 答案列表（順序必須與題目卷一致）
        // 所有答案添加到同一個 section，讓 Word 自然處理分頁
        // 使用相同的問題物件，確保與題目卷完全一致
        const answerChildren = [];

        // 將多行文字（含 \n）轉成逐行 Paragraph（保留空白行）
        // 僅供答案卷 EX 題排版使用，避免影響既有 MC / Financial 行為
        const makeParagraphsFromLines = (text, paragraphOptions) => {
            const opts = paragraphOptions || {};
            const normalized = (text == null ? '' : String(text)).replace(/\r\n/g, '\n').replace(/\r/g, '\n');
            const lines = normalized.split('\n');
            const out = [];
            for (let i = 0; i < lines.length; i++) {
                const line = lines[i];
                if (line === '') {
                    out.push(new docx.Paragraph({ children: [], ...opts }));
                } else {
                    out.push(new docx.Paragraph({
                        children: [new docx.TextRun({ text: line.replace(/\s+$/g, ''), size: 20 })],
                        ...opts
                    }));
                }
            }
            return out;
        };
        
        // 添加摘要表格標題
        answerChildren.push(
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: 'Answer Summary',
                        bold: true,
                        size: 24
                    })
                ],
                spacing: { after: 200 }
            })
        );
        
        // 添加摘要表格（與題目卷答案格同邊框與寬度）
        const summaryBorder = (typeof docx.BorderStyle !== 'undefined')
            ? { style: docx.BorderStyle.SINGLE, size: 4 }
            : { size: 4 };
        answerChildren.push(
            new docx.Table({
                rows: summaryTableRows,
                width: { size: 100, type: docx.WidthType.PERCENTAGE },
                borders: {
                    top: summaryBorder,
                    bottom: summaryBorder,
                    left: summaryBorder,
                    right: summaryBorder,
                    insideHorizontal: summaryBorder,
                    insideVertical: summaryBorder
                }
            })
        );
        
        // 添加分隔
        answerChildren.push(
            new docx.Paragraph({
                children: [],
                spacing: { after: 400 }
            })
        );
        
        // 添加詳細答案標題
        answerChildren.push(
            new docx.Paragraph({
                children: [
                    new docx.TextRun({
                        text: 'Detailed Answers',
                        bold: true,
                        size: 24
                    })
                ],
                spacing: { after: 200 }
            })
        );
        
        questions.forEach((q, index) => {
            // 1. 題目編號和文字（格式：1. 題目文字，與題目卷相同）
            answerChildren.push(
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: `${index + 1}. ${q.questionText}`,
                            size: 22
                        })
                    ],
                    spacing: { before: index === 0 ? 0 : 300, after: 100 }
                })
            );
            
            // 2. 原始 ID（在題目文字下方）
            answerChildren.push(
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: `(Original: ${q.originalId})`,
                            size: 18,
                            italics: true
                        })
                    ],
                    spacing: { after: 200 }
                })
            );

            // 3. 所有選項（正確答案顯示 ✔）
            q.options.forEach(function(option, optIndex) {
                var letter = String.fromCharCode(97 + optIndex);
                var isCorrect = (letter === q.correctOption);
                var displayOption = isCorrect ? option + ' ✔' : option;
                answerChildren.push(
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({
                                text: displayOption,
                                size: 20
                            })
                        ],
                        spacing: { after: 100 },
                        indent: { left: 400 }
                    })
                );
            });

            // 4. 正確答案標記
            answerChildren.push(
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: `Correct Answer: ${q.correctOption.toUpperCase()}${q.hasCheckmark ? ' ✔' : ''}`,
                            bold: true,
                            size: 22
                        })
                    ],
                    spacing: { before: 200, after: 200 }
                })
            );

            // 5. Feedback（如果存在，只顯示在答案卷）
            if (q.feedbackText) {
                answerChildren.push(
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({
                                text: 'Feedback:',
                                bold: true,
                                size: 22
                            })
                        ],
                        spacing: { before: 0, after: 100 }
                    }),
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({
                                text: q.feedbackText,
                                size: 20
                            })
                        ],
                        spacing: { after: 300 },
                        indent: { left: 400 }
                    })
                );
            }
        });

        // EX 區塊（僅 Managerial Accounting，且要有 EX 題目）
        if (currentSubject === 'managerial' && exSelectedAll && exSelectedAll.length > 0) {
            // EX 區塊標題
            answerChildren.push(
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: 'II. EXERCISES (Selected)',
                            bold: true,
                            size: 22
                        })
                    ],
                    spacing: { before: 400, after: 200 }
                })
            );

            // 輸出每個 EX 題目的完整內容
            exSelectedAll.forEach((ex, index) => {
                // 1. 完整原始編號標題
                answerChildren.push(
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({
                                text: ex.originalId,
                                bold: true,
                                size: 22
                            })
                        ],
                        spacing: { before: index === 0 ? 0 : 300, after: 100 }
                    })
                );

                // 2. 題目敘述（promptText）- 逐行輸出
                if (ex.promptText && ex.promptText.trim()) {
                    answerChildren.push(...makeParagraphsFromLines(ex.promptText, {
                        spacing: { after: 80 }
                    }));
                }

                // 3. Required 區塊 - 逐行輸出
                if (ex.requiredText && ex.requiredText.trim()) {
                    answerChildren.push(
                        new docx.Paragraph({
                            children: [
                                new docx.TextRun({
                                    text: 'Required:',
                                    bold: true,
                                    size: 20
                                })
                            ],
                            spacing: { after: 50 }
                        })
                    );
                    answerChildren.push(...makeParagraphsFromLines(ex.requiredText, {
                        spacing: { after: 80 },
                        indent: { left: 400 }
                    }));
                }

                // 4. 答案內容（answerTextOrTokens）- 逐行輸出
                if (ex.answerTextOrTokens && ex.answerTextOrTokens.trim()) {
                    answerChildren.push(
                        new docx.Paragraph({
                            children: [
                                new docx.TextRun({
                                    text: 'Answer:',
                                    bold: true,
                                    size: 20
                                })
                            ],
                            spacing: { after: 50 }
                        })
                    );
                    answerChildren.push(...makeParagraphsFromLines(ex.answerTextOrTokens, {
                        spacing: { after: 80 },
                        indent: { left: 400 }
                    }));
                }

                // 5. 完整原始文字區塊（包含 Feedback/Solution/Check My Work/Post-Submission 等）
                // 從 rawBlockText 中提取 Feedback 等後段內容
                if (ex.rawBlockText && ex.rawBlockText.trim()) {
                    const rawText = ex.rawBlockText.trim();
                    
                    // 提取 Feedback 區塊 - 逐行輸出
                    const feedbackMatch = rawText.match(/Feedback\s+(.+?)(?=Post-Submission|Check My Work|Solution|$)/is);
                    if (feedbackMatch && feedbackMatch[1]) {
                        const feedbackContent = feedbackMatch[1].trim();
                        if (feedbackContent) {
                            answerChildren.push(
                                new docx.Paragraph({
                                    children: [
                                        new docx.TextRun({
                                            text: 'Feedback:',
                                            bold: true,
                                            size: 20
                                        })
                                    ],
                                    spacing: { after: 50 }
                                })
                            );
                            answerChildren.push(...makeParagraphsFromLines(feedbackContent, {
                                spacing: { after: 80 },
                                indent: { left: 400 }
                            }));
                        }
                    }

                    // 提取 Solution 區塊 - 逐行輸出
                    const solutionMatch = rawText.match(/Solution\s+(.+?)(?=Feedback|Post-Submission|Check My Work|$)/is);
                    if (solutionMatch && solutionMatch[1]) {
                        const solutionContent = solutionMatch[1].trim();
                        if (solutionContent) {
                            answerChildren.push(
                                new docx.Paragraph({
                                    children: [
                                        new docx.TextRun({
                                            text: 'Solution:',
                                            bold: true,
                                            size: 20
                                        })
                                    ],
                                    spacing: { after: 50 }
                                })
                            );
                            answerChildren.push(...makeParagraphsFromLines(solutionContent, {
                                spacing: { after: 80 },
                                indent: { left: 400 }
                            }));
                        }
                    }

                    // 提取 Check My Work 區塊 - 逐行輸出
                    const checkMatch = rawText.match(/Check My Work\s+(.+?)(?=Feedback|Post-Submission|Solution|$)/is);
                    if (checkMatch && checkMatch[1]) {
                        const checkContent = checkMatch[1].trim();
                        if (checkContent) {
                            answerChildren.push(
                                new docx.Paragraph({
                                    children: [
                                        new docx.TextRun({
                                            text: 'Check My Work:',
                                            bold: true,
                                            size: 20
                                        })
                                    ],
                                    spacing: { after: 50 }
                                })
                            );
                            answerChildren.push(...makeParagraphsFromLines(checkContent, {
                                spacing: { after: 80 },
                                indent: { left: 400 }
                            }));
                        }
                    }

                    // 提取 Post-Submission 區塊 - 逐行輸出
                    const postMatch = rawText.match(/Post-Submission\s+(.+?)(?=Feedback|Check My Work|Solution|$)/is);
                    if (postMatch && postMatch[1]) {
                        const postContent = postMatch[1].trim();
                        if (postContent) {
                            answerChildren.push(
                                new docx.Paragraph({
                                    children: [
                                        new docx.TextRun({
                                            text: 'Post-Submission:',
                                            bold: true,
                                            size: 20
                                        })
                                    ],
                                    spacing: { after: 50 }
                                })
                            );
                            answerChildren.push(...makeParagraphsFromLines(postContent, {
                                spacing: { after: 80 },
                                indent: { left: 400 }
                            }));
                        }
                    }
                }

                // 題目之間的分隔
                answerChildren.push(
                    new docx.Paragraph({
                        children: [],
                        spacing: { after: 200 }
                    })
                );
            });
        }

        // Word 非選擇題區塊（僅 Managerial，放在 EX 之後）
        if (currentSubject === 'managerial' && wordNonMcSelected && wordNonMcSelected.length > 0) {
            answerChildren.push(
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: 'II. NON-MULTIPLE-CHOICE (Word) - Answers',
                            bold: true,
                            size: 22
                        })
                    ],
                    spacing: { before: 400, after: 200 }
                })
            );
            wordNonMcSelected.forEach((w, idx) => {
                if (idx > 0) {
                    answerChildren.push(new docx.Paragraph({ children: [], spacing: { before: 300, after: 0 } }));
                }
                answerChildren.push(
                    new docx.Paragraph({
                        children: [new docx.TextRun({ text: w.originalId + '.', bold: true, size: 22 })],
                        spacing: { after: 100 }
                    })
                );
                const qLines = w.questionLines || [];
                if (qLines.length > 0) {
                    const qText = qLines.map(l => (l == null ? '' : String(l))).join('\n');
                    answerChildren.push(...makeParagraphsFromLines(qText, { spacing: { after: 80 } }));
                }
                if (w.hasAnswerSection && (w.answerLines || []).length > 0) {
                    answerChildren.push(
                        new docx.Paragraph({
                            children: [new docx.TextRun({ text: 'ANSWER:', bold: true, size: 20 })],
                            spacing: { after: 50 }
                        })
                    );
                    const aText = w.answerLines.map(l => (l == null ? '' : String(l))).join('\n');
                    answerChildren.push(...makeParagraphsFromLines(aText, { spacing: { after: 80 }, indent: { left: 400 } }));
                }
                answerChildren.push(
                    new docx.Paragraph({ children: [], spacing: { after: 200 } })
                );
            });
        }

        // 將摘要表格和詳細題目都添加到同一個陣列
        allChildren.push(...answerChildren);
        
        // 創建單一 section，包含所有內容（標題、摘要表格、詳細題目、EX 題目）
        const doc = new docx.Document({
            sections: [{
                properties: {},
                children: allChildren
            }]
        });

        const blob = await docx.Packer.toBlob(doc);
        return blob;
    }

    // 下載文件
    downloadFile(blob, filename) {
        saveAs(blob, filename);
    }
}

// ========== 主應用程式邏輯 ==========
// DOM 元素
const subjectSelect = document.getElementById('subjectSelect');
const mainTitle = document.getElementById('mainTitle');
const examNameInput = document.getElementById('examName');
const examNameBuilder = document.getElementById('examNameBuilder');
const pointsConfig = document.getElementById('pointsConfig');
const addPointsRowBtn = document.getElementById('addPointsRowBtn');
const pointsTotalDisplay = document.getElementById('pointsTotalDisplay');
const uploadArea = document.getElementById('uploadArea');
const pdfInput = document.getElementById('pdfInput');
const fileList = document.getElementById('fileList');
const parseSection = document.getElementById('parseSection');
const parseStatus = document.getElementById('parseStatus');
const parsedQuestionsDiv = document.getElementById('parsedQuestions');
const generateSection = document.getElementById('generateSection');
const generateBtn = document.getElementById('generateBtn');
const generateStatus = document.getElementById('generateStatus');
const wordUploadSection = document.getElementById('wordUploadSection');
const wordUploadArea = document.getElementById('wordUploadArea');
const wordDropZone = document.getElementById('wordDropZone');
const wordInput = document.getElementById('wordInput');
const wordParseStatus = document.getElementById('wordParseStatus');
const wordConfig = document.getElementById('wordConfig');
const wordFileNameSpan = document.getElementById('wordFileNameSpan');
const wordAvailableSpan = document.getElementById('wordAvailableSpan');
const wordRequestedInput = document.getElementById('wordRequestedInput');
const wordClearBtn = document.getElementById('wordClearBtn');

// Points Rows 資料結構（羅馬數字 I-X）
let pointsRows = [
    { label: "I", value: 150 },
    { label: "II", value: 25 },
    { label: "III", value: 25 }
];

// 羅馬數字對應表（I 到 X）
const ROMAN_NUMERALS = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X"];

// 計算總分
function computeTotalPoints() {
    return pointsRows.reduce((sum, row) => {
        const val = parseInt(row.value, 10) || 0;
        return sum + (val >= 0 ? val : 0);
    }, 0);
}

// 渲染 Points 設定 UI
function renderPointsConfigUI() {
    if (!pointsConfig) return;
    
    // 重新編號所有列，確保 label 連續（I, II, III...）
    pointsRows.forEach((row, index) => {
        row.label = ROMAN_NUMERALS[index];
    });
    
    pointsConfig.innerHTML = '';
    
    const container = document.createElement('div');
    container.style.display = 'flex';
    container.style.gap = '15px';
    container.style.alignItems = 'center';
    container.style.flexWrap = 'wrap';
    
    pointsRows.forEach((row, index) => {
        const rowDiv = document.createElement('div');
        rowDiv.style.display = 'flex';
        rowDiv.style.alignItems = 'center';
        rowDiv.style.gap = '5px';
        
        const label = document.createElement('label');
        label.textContent = row.label + '：';
        label.style.fontSize = '14px';
        label.style.color = '#555';
        
        const input = document.createElement('input');
        input.type = 'number';
        input.min = '0';
        input.step = '1';
        input.value = row.value;
        input.style.width = '80px';
        input.style.padding = '5px';
        input.style.border = '2px solid #ddd';
        input.style.borderRadius = '4px';
        input.style.textAlign = 'center';
        input.id = `pointsInput_${index}`;
        
        // 輸入變更時更新資料並重新計算總分
        input.addEventListener('input', () => {
            const val = parseInt(input.value, 10) || 0;
            pointsRows[index].value = val >= 0 ? val : 0;
            updateTotalDisplay();
        });
        input.addEventListener('change', () => {
            const val = parseInt(input.value, 10) || 0;
            pointsRows[index].value = val >= 0 ? val : 0;
            updateTotalDisplay();
        });
        
        const removeBtn = document.createElement('button');
        removeBtn.textContent = 'X';
        removeBtn.style.padding = '3px 8px';
        removeBtn.style.background = '#ff4757';
        removeBtn.style.color = 'white';
        removeBtn.style.border = 'none';
        removeBtn.style.borderRadius = '3px';
        removeBtn.style.cursor = 'pointer';
        removeBtn.style.fontSize = '12px';
        
        // 刪除列（至少保留 2 列）
        removeBtn.addEventListener('click', () => {
            if (pointsRows.length > 2) {
                pointsRows.splice(index, 1);
                renderPointsConfigUI();
                updateTotalDisplay();
            }
        });
        
        // 當列數 <= 2 時禁用刪除按鈕
        if (pointsRows.length <= 2) {
            removeBtn.disabled = true;
            removeBtn.style.opacity = '0.5';
            removeBtn.style.cursor = 'not-allowed';
        }
        
        rowDiv.appendChild(label);
        rowDiv.appendChild(input);
        rowDiv.appendChild(removeBtn);
        container.appendChild(rowDiv);
    });
    
    pointsConfig.appendChild(container);
    
    // 更新 Add Row 按鈕狀態（最多 10 列）
    if (addPointsRowBtn) {
        addPointsRowBtn.disabled = pointsRows.length >= 10;
        if (pointsRows.length >= 10) {
            addPointsRowBtn.style.opacity = '0.5';
            addPointsRowBtn.style.cursor = 'not-allowed';
        } else {
            addPointsRowBtn.style.opacity = '1';
            addPointsRowBtn.style.cursor = 'pointer';
        }
    }
}

// 更新總分顯示
function updateTotalDisplay() {
    if (pointsTotalDisplay) {
        const total = computeTotalPoints();
        pointsTotalDisplay.value = total.toString();
    }
}

// 新增 Points 列
function addPointsRow() {
    if (pointsRows.length >= 10) return;
    
    // 新增列時，label 會在 renderPointsConfigUI() 中自動重新編號
    pointsRows.push({ label: "", value: 0 });
    renderPointsConfigUI();
    updateTotalDisplay();
}

// 讀取 Exam Points 從 UI（並更新 total 顯示）
function readExamPointsFromUI() {
    const total = computeTotalPoints();
    
    // 更新 total 顯示
    updateTotalDisplay();
    
    // 返回 pointsRows 陣列和 total（用於 DOCX 生成）
    return {
        rows: pointsRows.map(row => ({ label: row.label, value: row.value })),
        total: total
    };
}

// 建立 Managerial Accounting Exam Name
function buildManagerialExamName() {
    return `${MANAGERIAL_EXAMNAME_PREFIX} ${selectedYear} ${selectedTerm} ${selectedExamType}`;
}

// 設定 Builder 按鈕的 active 狀態
function setActiveButton(group, value) {
    if (!examNameBuilder) return;
    const groupElement = examNameBuilder.querySelector(`[data-group="${group}"]`);
    if (!groupElement) return;
    
    const buttons = groupElement.querySelectorAll('.builder-btn');
    buttons.forEach(btn => {
        if (btn.getAttribute('data-value') === value) {
            btn.classList.add('active');
        } else {
            btn.classList.remove('active');
        }
    });
}

// 依 currentSubject 更新標題、<title>、考卷名稱預設（僅在未自訂或等於舊預設時更新）
function applySubjectUI(prevSubject) {
    const cfg = SUBJECT_CONFIG[currentSubject];
    document.title = cfg.pageTitle;
    mainTitle.textContent = cfg.pageTitle;
    examNameInput.placeholder = '例如：' + cfg.defaultExamName;

    if (currentSubject === 'managerial') {
        // 顯示 Exam Name Builder
        if (examNameBuilder) {
            examNameBuilder.style.display = 'block';
        }
        if (wordUploadSection) wordUploadSection.style.display = 'block';
        
        // 設定預設值（如果剛切換到 managerial）
        if (prevSubject !== 'managerial') {
            selectedYear = '2026';
            selectedTerm = 'Spring';
            selectedExamType = 'Exam 1';
        }
        
        // 更新按鈕 active 狀態
        setActiveButton('year', selectedYear);
        setActiveButton('term', selectedTerm);
        setActiveButton('examType', selectedExamType);
        
        // 更新 Exam Name
        examNameInput.value = buildManagerialExamName();
    } else {
        // Financial Accounting：隱藏 Builder 與 Word 上傳區
        if (examNameBuilder) {
            examNameBuilder.style.display = 'none';
        }
        if (wordUploadSection) wordUploadSection.style.display = 'none';
        
        // Financial Accounting 的既有邏輯
        const currentVal = (examNameInput.value || '').trim();
        const shouldUpdateExamName = prevSubject == null ||
            !currentVal ||
            currentVal === (SUBJECT_CONFIG[prevSubject] && SUBJECT_CONFIG[prevSubject].defaultExamName);
        if (shouldUpdateExamName) {
            examNameInput.value = cfg.defaultExamName;
        }
    }
}

// 科目切換
subjectSelect.value = currentSubject;
applySubjectUI(null);
subjectSelect.addEventListener('change', () => {
    const prev = currentSubject;
    currentSubject = subjectSelect.value;
    applySubjectUI(prev);
});

// Add Row 按鈕事件
if (addPointsRowBtn) {
    addPointsRowBtn.addEventListener('click', addPointsRow);
}

// Exam Name Builder 按鈕事件（事件委派）
if (examNameBuilder) {
    examNameBuilder.addEventListener('click', (e) => {
        if (e.target.classList.contains('builder-btn')) {
            const btn = e.target;
            const value = btn.getAttribute('data-value');
            const groupElement = btn.closest('[data-group]');
            
            if (groupElement) {
                const group = groupElement.getAttribute('data-group');
                
                // 更新狀態
                if (group === 'year') {
                    selectedYear = value;
                } else if (group === 'term') {
                    selectedTerm = value;
                } else if (group === 'examType') {
                    selectedExamType = value;
                }
                
                // 更新按鈕 active 狀態
                setActiveButton(group, value);
                
                // 更新 Exam Name（只在 managerial 時）
                if (currentSubject === 'managerial') {
                    examNameInput.value = buildManagerialExamName();
                }
            }
        }
    });
}

// 初始化 Points UI
renderPointsConfigUI();
updateTotalDisplay();

// 初始化解析器
parser = new PDFParser();

// 初始化：確保按鈕狀態正確
updateExportButton();

// 上傳區域點擊事件
uploadArea.addEventListener('click', () => {
    pdfInput.click();
});

// 拖放功能
uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.style.background = '#d0d8ff';
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.style.background = '#f0f4ff';
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.style.background = '#f0f4ff';
    const files = Array.from(e.dataTransfer.files).filter(f => f.type === 'application/pdf');
    addFiles(files);
});

// 檔案選擇事件
pdfInput.addEventListener('change', (e) => {
    const files = Array.from(e.target.files);
    addFiles(files);
});

// Word 上傳區（僅 Managerial 顯示）：點擊選檔 + 拖曳上傳
function handleWordDrop(files) {
    if (currentSubject !== 'managerial') return;
    if (!files || files.length === 0) return;
    if (files.length > 1) {
        wordParseState = 'error';
        if (wordParseStatus) {
            wordParseStatus.textContent = '只接受單一 Word 檔，請一次拖曳一個 .docx 檔案';
            wordParseStatus.style.color = '#c00';
        }
        if (wordConfig) wordConfig.style.display = 'none';
        return;
    }
    const file = files[0];
    const isDocx = (file.name && file.name.toLowerCase().endsWith('.docx')) ||
        (file.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    if (!isDocx) {
        wordParseState = 'error';
        if (wordParseStatus) {
            wordParseStatus.textContent = '請上傳 .docx 檔案（非選擇題題庫）';
            wordParseStatus.style.color = '#c00';
        }
        if (wordConfig) wordConfig.style.display = 'none';
        return;
    }
    parseWordFile(file);
}

function bindWordDropEvents() {
    if (!wordDropZone) return;
    wordDropZone.addEventListener('click', () => { if (wordInput) wordInput.click(); });
    wordDropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        wordDropZone.classList.add('dragover');
    });
    wordDropZone.addEventListener('dragleave', () => {
        wordDropZone.classList.remove('dragover');
    });
    wordDropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        wordDropZone.classList.remove('dragover');
        const files = Array.from(e.dataTransfer.files || []);
        handleWordDrop(files);
    });
}

bindWordDropEvents();

if (wordInput) {
    wordInput.addEventListener('change', (e) => {
        const f = e.target.files && e.target.files[0];
        if (f) parseWordFile(f);
    });
}
if (wordClearBtn) {
    wordClearBtn.addEventListener('click', () => {
        wordFile = null;
        wordNonMcQuestions = [];
        wordParseState = 'idle';
        if (wordInput) wordInput.value = '';
        if (wordRequestedInput) wordRequestedInput.value = '0';
        updateWordParseUI();
    });
}
function clampWordRequested() {
    if (!wordRequestedInput || !wordAvailableSpan) return;
    const max = parseInt(wordAvailableSpan.textContent, 10) || 0;
    let v = parseInt(wordRequestedInput.value, 10);
    if (isNaN(v)) v = 0;
    if (v < 0) v = 0;
    if (v > max) v = max;
    wordRequestedInput.value = String(v);
    wordRequestedInput.max = max;
}
if (wordRequestedInput) {
    wordRequestedInput.addEventListener('input', clampWordRequested);
    wordRequestedInput.addEventListener('change', clampWordRequested);
}

// 添加檔案
function addFiles(files) {
    console.log('PDF upload handler fired');
    files.forEach(file => {
        if (!pdfFiles.find(f => f.name === file.name && f.size === file.size)) {
            pdfFiles.push(file);
        }
    });
    updateFileList();
    
    // 自動開始解析
    if (pdfFiles.length > 0) {
        setTimeout(() => {
            parsePDFs();
        }, 500);
    }
}

// 更新檔案列表顯示（包含題目數量輸入框）
function updateFileList() {
    console.log('Per-file UI rendered');
    fileList.innerHTML = '';
    
    // 如果有解析結果，顯示題目數量選擇
    if (parsedQuestionsByFile.length > 0) {
        // 添加標題
        const titleDiv = document.createElement('div');
        titleDiv.style.marginBottom = '15px';
        titleDiv.style.fontWeight = 'bold';
        titleDiv.style.color = '#555';
        titleDiv.textContent = '選擇每個檔案的題目數量：';
        fileList.appendChild(titleDiv);
        
        parsedQuestionsByFile.forEach((item, index) => {
            const mcAvailable = item.questions.length; // 維持既有：questions === MC
            const exAvailable = (item.exQuestions && item.exQuestions.length) ? item.exQuestions.length : 0;
            const exRequested = (parseInt(exRequestedCountsByFileIndex[index], 10) || 0);
            const fileItem = document.createElement('div');
            fileItem.className = 'file-item';
            fileItem.innerHTML = `
                <div style="flex: 1;">
                    <span class="file-name">${item.fileName}</span>
                    <span style="color: #888; font-size: 12px; margin-left: 10px;">（MC 可用：${mcAvailable} 題 / EX 可用：${exAvailable} 題）</span>
                </div>
                <div style="display: flex; align-items: center; gap: 10px;">
                    <label style="font-size: 14px; color: #555;">MC：</label>
                    <input type="number" 
                           id="questionCount_${index}" 
                           class="question-count-input" 
                           min="0" 
                           max="${mcAvailable}"
                           step="1"
                           value="0" 
                           style="width: 60px; padding: 5px; border: 2px solid #ddd; border-radius: 4px; text-align: center;">
                    <label style="font-size: 14px; color: #555;">EX：</label>
                    <input type="number" 
                           id="exQuestionCount_${index}" 
                           class="question-count-input" 
                           min="0" 
                           max="${exAvailable}"
                           step="1"
                           value="${exRequested}" 
                           style="width: 60px; padding: 5px; border: 2px solid #ddd; border-radius: 4px; text-align: center;">
                    <button class="file-remove" onclick="removeFile(${index})">移除</button>
                </div>
            `;
            fileList.appendChild(fileItem);

            // EX Requested Count：基本驗證 + 存 state（不影響既有匯出）
            const exInput = document.getElementById(`exQuestionCount_${index}`);
            if (exInput) {
                const clampAndStore = () => {
                    let v = parseInt(exInput.value, 10);
                    if (isNaN(v)) v = 0;
                    if (v < 0) v = 0;
                    if (v > exAvailable) v = exAvailable;
                    exInput.value = String(v);
                    exRequestedCountsByFileIndex[index] = v;
                };
                exInput.addEventListener('input', clampAndStore);
                exInput.addEventListener('change', clampAndStore);
            }
        });
    } else {
        // 如果還沒有解析結果，只顯示檔案名稱
        pdfFiles.forEach((file, index) => {
            const fileItem = document.createElement('div');
            fileItem.className = 'file-item';
            fileItem.innerHTML = `
                <span class="file-name">${file.name}</span>
                <button class="file-remove" onclick="removeFile(${index})">移除</button>
            `;
            fileList.appendChild(fileItem);
        });
    }
    
    // 更新生成按鈕狀態
    updateExportButton();
}

// 移除檔案
window.removeFile = function(index) {
    pdfFiles.splice(index, 1);
    if (parsedQuestionsByFile.length > index) {
        parsedQuestionsByFile.splice(index, 1);
    }
    if (exRequestedCountsByFileIndex.length > index) {
        exRequestedCountsByFileIndex.splice(index, 1);
    }
    updateFileList();
    parsedQuestions = [];
    parsedExerciseQuestions = [];
    if (pdfFiles.length === 0) {
        parsedQuestionsByFile = [];
        exRequestedCountsByFileIndex = [];
    }
    parseSection.style.display = 'none';
    generateSection.style.display = 'none';
    
    // 如果還有檔案，重新解析
    if (pdfFiles.length > 0) {
        setTimeout(() => {
            parsePDFs();
        }, 500);
    } else {
        updateExportButton();
    }
};

// 更新生成按鈕的顯示狀態
function updateExportButton() {
    console.log('Export button rendered', {
        hasFiles: pdfFiles.length > 0,
        hasParsedData: parsedQuestionsByFile.length > 0
    });
    
    const hasFiles = pdfFiles.length > 0;
    const hasParsedData = parsedQuestionsByFile.length > 0;
    
    if (hasFiles && hasParsedData) {
        generateSection.style.display = 'block';
        generateBtn.disabled = false;
        generateBtn.textContent = 'Generate & Download DOCX';
    } else if (hasFiles) {
        // 有檔案但還沒解析完成
        generateSection.style.display = 'block';
        generateBtn.disabled = true;
        generateBtn.textContent = 'Generating & Download DOCX (解析中...)';
    } else {
        // 沒有檔案
        generateSection.style.display = 'none';
    }
}

// 解析 PDF
async function parsePDFs() {
    if (pdfFiles.length === 0) {
        parseStatus.innerHTML = '<div class="status error">請先上傳 PDF 檔案</div>';
        parseSection.style.display = 'block';
        return;
    }

    parseStatus.innerHTML = '<div class="status info"><span class="loading"></span>正在解析 PDF 檔案，請稍候...</div>';
    parseSection.style.display = 'block';
    generateSection.style.display = 'none';

    try {
        const parseResult = await parser.parsePDFs(pdfFiles);
        parsedQuestions = parseResult.allQuestions;
        parsedQuestionsByFile = parseResult.byFile;
        parsedExerciseQuestions = parseResult.allExQuestions || [];
        // 對齊每檔案的 EX requested state（預設 0）
        exRequestedCountsByFileIndex = parsedQuestionsByFile.map((_, i) => (parseInt(exRequestedCountsByFileIndex[i], 10) || 0));
        
        if (parsedQuestions.length === 0) {
            var msg = currentSubject === 'financial'
                ? '未能從 PDF 中提取到任何 Financial 題目（請確認為 McGraw-Hill Connect Print View 格式）。'
                : '未能從 PDF 中提取到任何 MC 題目。請確認 PDF 格式正確。';
            parseStatus.innerHTML = '<div class="status error">' + msg + '</div>';
            updateExportButton();
            return;
        }

        // 顯示解析結果（按檔案分組）
        const totalMC = parsedQuestions.length;
        const totalEX = parsedQuestionsByFile.reduce((sum, it) => sum + ((it.exQuestions && it.exQuestions.length) ? it.exQuestions.length : 0), 0);
        let infoHTML = `<h3>解析完成！</h3><ul>`;
        infoHTML += `<li>總共找到 <strong>${totalMC}</strong> 題 MC 題目</li>`;
        infoHTML += `<li>總共找到 <strong>${totalEX}</strong> 題 EX（非選擇題）</li>`;
        infoHTML += `<li>檔案數量：<strong>${parsedQuestionsByFile.length}</strong> 個</li>`;
        parsedQuestionsByFile.forEach((item, index) => {
            const mcCount = (item.mcQuestions ? item.mcQuestions.length : item.questions.length);
            const exCount = (item.exQuestions ? item.exQuestions.length : 0);
            infoHTML += `<li>${item.fileName}: MC ${mcCount} 題 / EX ${exCount} 題</li>`;
        });
        infoHTML += `</ul>`;
        parsedQuestionsDiv.innerHTML = infoHTML;
        
        // 更新檔案列表，顯示題目數量輸入框
        updateFileList();
        
        parseStatus.innerHTML = '<div class="status success">✓ PDF 解析成功！請在上方為每個檔案設定要選擇的題目數。</div>';
        
        // 確保生成按鈕可見且啟用
        updateExportButton();
        
    } catch (error) {
        parseStatus.innerHTML = `<div class="status error">解析失敗：${error.message}</div>`;
        console.error(error);
        updateExportButton();
    }
}

// 生成試卷
generateBtn.addEventListener('click', async () => {
    console.log('Export button clicked');
    
    if (parsedQuestionsByFile.length === 0) {
        generateStatus.innerHTML = '<div class="status error">請先上傳並解析 PDF 檔案</div>';
        return;
    }

    // 獲取考卷名稱（使用者輸入優先，否則用目前科目的 defaultExamName）
    let examName = examNameInput.value.trim();
    if (!examName) {
        examName = SUBJECT_CONFIG[currentSubject].defaultExamName;
    }
    
    // 清理檔案名稱中的無效字元（保留空格）
    const sanitizeFileName = (name) => {
        // 移除 Windows 檔案系統不允許的字元：< > : " / \ | ? *
        return name.replace(/[<>:"/\\|?*]/g, '');
    };
    
    const safeExamName = sanitizeFileName(examName);
    const questionFileName = `${safeExamName} - Questions.docx`;
    const answerFileName = `${safeExamName} - Answers.docx`;
    
    // 從每個 PDF 獲取要選擇的題目數量
    const questionCounts = [];
    let totalSelected = 0;
    let hasError = false;
    let errorMessage = '';

    for (let i = 0; i < parsedQuestionsByFile.length; i++) {
        const countInput = document.getElementById(`questionCount_${i}`);
        const requestedCount = parseInt(countInput ? countInput.value : 0, 10) || 0;
        const availableCount = parsedQuestionsByFile[i].questions.length;
        
        // EX Requested Count：先做基本驗證並存 state（本輪不納入匯出）
        const exInput = document.getElementById(`exQuestionCount_${i}`);
        const exRequested = parseInt(exInput ? exInput.value : 0, 10) || 0;
        const exAvailable = (parsedQuestionsByFile[i].exQuestions && parsedQuestionsByFile[i].exQuestions.length) ? parsedQuestionsByFile[i].exQuestions.length : 0;
        if (exRequested < 0) {
            hasError = true;
            errorMessage = `${parsedQuestionsByFile[i].fileName}: EX 題目數不能為負數`;
            break;
        }
        
        if (exRequested > exAvailable) {
            hasError = true;
            errorMessage = `${parsedQuestionsByFile[i].fileName}: EX 請求 ${exRequested} 題，但只有 ${exAvailable} 題可用`;
            break;
        }
        
        exRequestedCountsByFileIndex[i] = exRequested;
        
        if (requestedCount < 0) {
            hasError = true;
            errorMessage = `${parsedQuestionsByFile[i].fileName}: 題目數不能為負數`;
            break;
        }
        
        if (requestedCount > availableCount) {
            hasError = true;
            errorMessage = `${parsedQuestionsByFile[i].fileName}: 請求 ${requestedCount} 題，但只有 ${availableCount} 題可用`;
            break;
        }
        
        questionCounts.push({
            fileIndex: i,
            fileName: parsedQuestionsByFile[i].fileName,
            requestedCount: requestedCount,
            availableCount: availableCount,
            questions: parsedQuestionsByFile[i].questions
        });
        
        totalSelected += requestedCount;
    }

    // Word 非選擇題驗證（僅 Managerial，且已上傳 Word 時）
    if (!hasError && currentSubject === 'managerial' && wordParseState === 'parsed' && wordNonMcQuestions.length > 0) {
        clampWordRequested();
        const wReq = parseInt(wordRequestedInput && wordRequestedInput.value ? wordRequestedInput.value : 0, 10) || 0;
        const wAvail = wordNonMcQuestions.length;
        if (wReq < 0) {
            hasError = true;
            errorMessage = 'Word 要抽題數不能為負數';
        } else if (wReq > wAvail) {
            hasError = true;
            errorMessage = `Word 請求 ${wReq} 題，但只有 ${wAvail} 題可用`;
        }
    }

    if (hasError) {
        generateStatus.innerHTML = `<div class="status error">${errorMessage}</div>`;
        return;
    }

    if (totalSelected === 0) {
        generateStatus.innerHTML = '<div class="status error">請至少為一個檔案設定大於 0 的題目數</div>';
        return;
    }

    generateStatus.innerHTML = '<div class="status info"><span class="loading"></span>正在生成試卷，請稍候...</div>';
    generateBtn.disabled = true;

    try {
        console.log('Export started');
        
        // 從每個 PDF 中分別隨機選擇指定數量的題目
        const selectedQuestions = [];
        const generator = new QuestionGenerator([]); // 用於打亂功能
        
        questionCounts.forEach((item, index) => {
            if (item.requestedCount > 0) {
                // 從該檔案的題目中隨機選擇
                const shuffled = generator.shuffle([...item.questions]);
                const selected = shuffled.slice(0, item.requestedCount);
                selectedQuestions.push(...selected);
            }
        });
        
        // 打亂所有選中的題目順序
        const examQuestions = generator.shuffle(selectedQuestions);
        
        // 重新編號
        examQuestions.forEach((q, index) => {
            q.examNumber = index + 1;
        });

        // EX 抽題邏輯（僅 Managerial Accounting）
        let exSelectedAll = [];
        if (currentSubject === 'managerial') {
            const exSelectedByFile = [];
            for (let i = 0; i < parsedQuestionsByFile.length; i++) {
                const exRequested = exRequestedCountsByFileIndex[i] || 0;
                const exAvailable = (parsedQuestionsByFile[i].exQuestions && parsedQuestionsByFile[i].exQuestions.length) 
                    ? parsedQuestionsByFile[i].exQuestions 
                    : [];
                
                if (exRequested > 0 && exAvailable.length > 0) {
                    // 從該檔案的 exQuestions 隨機抽取
                    const shuffled = generator.shuffle([...exAvailable]);
                    const selected = shuffled.slice(0, exRequested);
                    exSelectedByFile.push(...selected);
                }
            }
            
            // 合併所有檔案抽到的 EX，可選擇是否再 shuffle 一次
            if (exSelectedByFile.length > 0) {
                exSelectedAll = generator.shuffle(exSelectedByFile);
            }
        }

        // Word 非選擇題抽題（僅 Managerial，且已上傳 Word 時；同一次 Generate 用同一組）
        let wordNonMcSelected = [];
        if (currentSubject === 'managerial' && wordParseState === 'parsed' && wordNonMcQuestions.length > 0 && wordRequestedInput) {
            const wReq = parseInt(wordRequestedInput.value, 10) || 0;
            if (wReq > 0) {
                wordNonMcSelected = generator.randomSelect(wordNonMcQuestions, wReq);
            }
        }

        // 讀取 Exam Points（僅 Managerial Accounting 需要）
        const examPoints = (currentSubject === 'managerial') ? readExamPointsFromUI() : null;
        
        // 生成 Word 文檔
        const wordGen = new WordGenerator();
        
        console.log('Generating Questions doc...');
        const questionBlob = await wordGen.generateQuestionSheet(examName, examQuestions, examPoints, exSelectedAll, wordNonMcSelected);
        console.log('Questions doc generated');
        
        console.log('Generating Answers doc...');
        const answerBlob = await wordGen.generateAnswerSheet(examName, examQuestions, exSelectedAll, wordNonMcSelected);
        console.log('Answers doc generated');
        
        // 下載兩個檔案（分別下載，確保不會覆蓋）
        console.log('Questions download triggered');
        wordGen.downloadFile(questionBlob, questionFileName);
        
        // 稍等一下再下載答案卷，避免瀏覽器同時下載衝突
        setTimeout(() => {
            console.log('Answers download triggered');
            wordGen.downloadFile(answerBlob, answerFileName);
            console.log('Export completed');
            
            generateStatus.innerHTML = '<div class="status success">✓ 試卷生成成功！已下載題目卷和答案卷。</div>';
            generateBtn.disabled = false;
        }, 300);

    } catch (error) {
        console.error('Export failed:', error);
        generateStatus.innerHTML = `<div class="status error">生成失敗：${error.message}</div>`;
        generateBtn.disabled = false;
    }
});
