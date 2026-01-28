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

// ========== 全域變數 ==========
let pdfFiles = [];
let parsedQuestions = []; // 所有題目的陣列
let parsedQuestionsByFile = []; // 按檔案分組的題目 [{file, questions}, ...]
let parser = null;
let generator = null;

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
        const resultsByFile = [];
        
        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            try {
                const questions = await this.parseSinglePDF(file);
                this.questions = this.questions.concat(questions);
                resultsByFile.push({
                    file: file,
                    fileName: file.name,
                    questions: questions
                });
            } catch (error) {
                console.error(`解析 ${file.name} 失敗:`, error);
                throw new Error(`無法解析 ${file.name}: ${error.message}`);
            }
        }
        
        return {
            allQuestions: this.questions,
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
            return await this.parseFinancialByPage(pdf, file);
        }

        // 以下是原有 Managerial 邏輯，完全不要動
        const questions = [];
        let allLines = [];
        for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
            const page = await pdf.getPage(pageNum);
            const textContent = await page.getTextContent();
            const pageLines = this.reconstructLines(textContent.items);
            allLines = allLines.concat(pageLines);
        }
        const fullText = allLines.join('\n');
        const extractedQuestions = this.extractQuestionsFromText(fullText);
        questions.push(...extractedQuestions);
        return questions;
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
        for (let b = 0; b < blockCount; b++) {
            const qCells = [new docx.TableCell({
                verticalAlignment: docx.VerticalAlign.CENTER,
                children: [
                    new docx.Paragraph({
                        alignment: docx.AlignmentType.CENTER,
                        children: [new docx.TextRun({ text: 'Q.', size: coverGridFontSize })]
                    })
                ],
                width: { size: 8, type: docx.WidthType.PERCENTAGE }
            })];
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
                qCells.push(new docx.TableCell({
                    verticalAlignment: docx.VerticalAlign.CENTER,
                    children: [
                        new docx.Paragraph({
                            alignment: docx.AlignmentType.CENTER,
                            children: [
                                new docx.TextRun({
                                    text: num <= questionCount ? String(num) : '',
                                    size: coverGridFontSize
                                })
                            ]
                        })
                    ],
                    width: { size: (92 / cols), type: docx.WidthType.PERCENTAGE }
                }));
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
    async generateQuestionSheet(examName, questions, points) {
        // 所有內容將添加到同一個 section，確保連續流動
        const allChildren = [];

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

        // 創建單一 section，包含所有內容（標題、表格、題目）
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
    async generateAnswerSheet(examName, questions) {
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

        // 將摘要表格和詳細題目都添加到同一個陣列
        allChildren.push(...answerChildren);
        
        // 創建單一 section，包含所有內容（標題、摘要表格、詳細題目）
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

// 依 currentSubject 更新標題、<title>、考卷名稱預設（僅在未自訂或等於舊預設時更新）
function applySubjectUI(prevSubject) {
    const cfg = SUBJECT_CONFIG[currentSubject];
    document.title = cfg.pageTitle;
    mainTitle.textContent = cfg.pageTitle;
    examNameInput.placeholder = '例如：' + cfg.defaultExamName;

    const currentVal = (examNameInput.value || '').trim();
    const shouldUpdateExamName = prevSubject == null ||
        !currentVal ||
        currentVal === (SUBJECT_CONFIG[prevSubject] && SUBJECT_CONFIG[prevSubject].defaultExamName);
    if (shouldUpdateExamName) {
        examNameInput.value = cfg.defaultExamName;
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
            const fileItem = document.createElement('div');
            fileItem.className = 'file-item';
            fileItem.innerHTML = `
                <div style="flex: 1;">
                    <span class="file-name">${item.fileName}</span>
                    <span style="color: #888; font-size: 12px; margin-left: 10px;">（可用：${item.questions.length} 題）</span>
                </div>
                <div style="display: flex; align-items: center; gap: 10px;">
                    <label style="font-size: 14px; color: #555;">題目數：</label>
                    <input type="number" 
                           id="questionCount_${index}" 
                           class="question-count-input" 
                           min="0" 
                           max="${item.questions.length}"
                           step="1"
                           value="0" 
                           style="width: 60px; padding: 5px; border: 2px solid #ddd; border-radius: 4px; text-align: center;">
                    <button class="file-remove" onclick="removeFile(${index})">移除</button>
                </div>
            `;
            fileList.appendChild(fileItem);
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
    updateFileList();
    parsedQuestions = [];
    if (pdfFiles.length === 0) {
        parsedQuestionsByFile = [];
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
        
        if (parsedQuestions.length === 0) {
            var msg = currentSubject === 'financial'
                ? '未能從 PDF 中提取到任何 Financial 題目（請確認為 McGraw-Hill Connect Print View 格式）。'
                : '未能從 PDF 中提取到任何 MC 題目。請確認 PDF 格式正確。';
            parseStatus.innerHTML = '<div class="status error">' + msg + '</div>';
            updateExportButton();
            return;
        }

        // 顯示解析結果（按檔案分組）
        let infoHTML = `<h3>解析完成！</h3><ul>`;
        infoHTML += `<li>總共找到 <strong>${parsedQuestions.length}</strong> 題 MC 題目</li>`;
        infoHTML += `<li>檔案數量：<strong>${parsedQuestionsByFile.length}</strong> 個</li>`;
        parsedQuestionsByFile.forEach((item, index) => {
            infoHTML += `<li>${item.fileName}: ${item.questions.length} 題</li>`;
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

        // 讀取 Exam Points（僅 Managerial Accounting 需要）
        const examPoints = (currentSubject === 'managerial') ? readExamPointsFromUI() : null;
        
        // 生成 Word 文檔
        const wordGen = new WordGenerator();
        
        console.log('Generating Questions doc...');
        const questionBlob = await wordGen.generateQuestionSheet(examName, examQuestions, examPoints);
        console.log('Questions doc generated');
        
        console.log('Generating Answers doc...');
        const answerBlob = await wordGen.generateAnswerSheet(examName, examQuestions);
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
