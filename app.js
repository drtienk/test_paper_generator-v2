// ============================================
// Cengage 隨機選擇題考卷產生器 - MVP
// ============================================

// ========== 全域變數 ==========
let pdfFiles = [];
let parsedQuestions = [];
let parser = null;
let generator = null;

// ========== PDF 解析器類別 ==========
class PDFParser {
    constructor() {
        this.questions = [];
    }

    // 解析多個 PDF 檔案
    async parsePDFs(files) {
        this.questions = [];
        
        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            try {
                const questions = await this.parseSinglePDF(file);
                this.questions = this.questions.concat(questions);
            } catch (error) {
                console.error(`解析 ${file.name} 失敗:`, error);
                throw new Error(`無法解析 ${file.name}: ${error.message}`);
            }
        }
        
        return this.questions;
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

    // 解析單個 PDF
    async parseSinglePDF(file) {
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        const questions = [];
        
        let allLines = [];

        // 讀取所有頁面
        for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
            const page = await pdf.getPage(pageNum);
            const textContent = await page.getTextContent();
            
            // 使用行重建方法
            const pageLines = this.reconstructLines(textContent.items);
            allLines = allLines.concat(pageLines);
        }
        
        // 將所有行合併成文字進行解析
        const fullText = allLines.join('\n');
        
        // 從文字中提取題目
        const extractedQuestions = this.extractQuestionsFromText(fullText);
        questions.push(...extractedQuestions);
        
        return questions;
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
    // 生成題目卷（學生用）
    async generateQuestionSheet(examName, questions) {
        const doc = new docx.Document({
            sections: [{
                properties: {},
                children: [
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
                    }),
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({
                                text: `總題數：${questions.length} 題`
                            })
                        ],
                        alignment: docx.AlignmentType.CENTER,
                        spacing: { after: 600 }
                    })
                ]
            }]
        });

        // 作答表格（第一頁）
        const tableRows = [];
        const cols = 5; // 每行 5 題
        const rows = Math.ceil(questions.length / cols);
        
        for (let row = 0; row < rows; row++) {
            const cells = [];
            for (let col = 0; col < cols; col++) {
                const index = row * cols + col;
                if (index < questions.length) {
                    cells.push(
                        new docx.TableCell({
                            children: [
                                new docx.Paragraph({
                                    children: [
                                        new docx.TextRun({
                                            text: `${index + 1}`,
                                            size: 18
                                        })
                                    ],
                                    alignment: docx.AlignmentType.CENTER
                                }),
                                new docx.Paragraph({
                                    children: [
                                        new docx.TextRun({
                                            text: '⬜',
                                            size: 20
                                        })
                                    ],
                                    alignment: docx.AlignmentType.CENTER
                                })
                            ],
                            width: { size: 20, type: docx.WidthType.PERCENTAGE }
                        })
                    );
                } else {
                    cells.push(
                        new docx.TableCell({
                            children: [],
                            width: { size: 20, type: docx.WidthType.PERCENTAGE }
                        })
                    );
                }
            }
            tableRows.push(
                new docx.TableRow({
                    children: cells
                })
            );
        }

        doc.addSection({
            properties: {},
            children: [
                new docx.Paragraph({
                    children: [
                        new docx.TextRun({
                            text: '作答檢查表',
                            bold: true,
                            size: 24
                        })
                    ],
                    spacing: { after: 200 }
                }),
                new docx.Table({
                    rows: tableRows,
                    width: { size: 100, type: docx.WidthType.PERCENTAGE }
                }),
                new docx.Paragraph({
                    children: [],
                    spacing: { after: 400 }
                })
            ]
        });

        // 題目內容（移除所有標記和原始 ID）
        // 所有題目添加到同一個 section，讓 Word 自然處理分頁
        const questionChildren = [];
        
        questions.forEach((q, index) => {
            // 清理題目文字：移除附錄標籤和 (Algorithmic)（僅在題目卷中）
            // 移除 (Appendix...) 格式的標籤，包括 (Appendix 4B), (Appendix A) 等
            let cleanedQuestionText = q.questionText.replace(/\(Appendix[^)]*\)/gi, '');
            // 移除 (Algorithmic) 標籤
            cleanedQuestionText = cleanedQuestionText.replace(/\(Algorithmic\)/gi, '');
            // 清理可能留下的多餘空格
            cleanedQuestionText = cleanedQuestionText.trim().replace(/\s+/g, ' ');
            
            // 題目編號和文字（格式：1. 題目文字）
            questionChildren.push(
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
            q.options.forEach(option => {
                const cleanOption = option.replace(/[✔✓]/g, '').trim();
                questionChildren.push(
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

        // 將所有題目添加到同一個 section
        doc.addSection({
            properties: {},
            children: questionChildren
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

        // 答案摘要表格（在詳細題目之前）
        const summaryTableRows = [];
        
        // 表頭
        summaryTableRows.push(
            new docx.TableRow({
                children: [
                    new docx.TableCell({
                        children: [
                            new docx.Paragraph({
                                children: [
                                    new docx.TextRun({
                                        text: 'Question No.',
                                        bold: true
                                    })
                                ],
                                alignment: docx.AlignmentType.CENTER
                            })
                        ],
                        shading: { fill: 'D3D3D3' }
                    }),
                    new docx.TableCell({
                        children: [
                            new docx.Paragraph({
                                children: [
                                    new docx.TextRun({
                                        text: 'Answer',
                                        bold: true
                                    })
                                ],
                                alignment: docx.AlignmentType.CENTER
                            })
                        ],
                        shading: { fill: 'D3D3D3' }
                    })
                ]
            })
        );
        
        // 表格內容（使用相同的問題順序）
        questions.forEach((q, index) => {
            summaryTableRows.push(
                new docx.TableRow({
                    children: [
                        new docx.TableCell({
                            children: [
                                new docx.Paragraph({
                                    children: [
                                        new docx.TextRun({
                                            text: `${index + 1}`,
                                            size: 20
                                        })
                                    ],
                                    alignment: docx.AlignmentType.CENTER
                                })
                            ]
                        }),
                        new docx.TableCell({
                            children: [
                                new docx.Paragraph({
                                    children: [
                                        new docx.TextRun({
                                            text: q.correctOption.toUpperCase(),
                                            size: 20
                                        })
                                    ],
                                    alignment: docx.AlignmentType.CENTER
                                })
                            ]
                        })
                    ]
                })
            );
        });

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
        
        // 添加摘要表格
        answerChildren.push(
            new docx.Table({
                rows: summaryTableRows,
                width: { size: 50, type: docx.WidthType.PERCENTAGE }
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

            // 3. 所有選項（保留原始 ✔/✓ 標記，與題目卷相同的順序和文字）
            q.options.forEach(option => {
                answerChildren.push(
                    new docx.Paragraph({
                        children: [
                            new docx.TextRun({
                                text: option,
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
const uploadArea = document.getElementById('uploadArea');
const pdfInput = document.getElementById('pdfInput');
const fileList = document.getElementById('fileList');
const parseSection = document.getElementById('parseSection');
const parseStatus = document.getElementById('parseStatus');
const parsedQuestionsDiv = document.getElementById('parsedQuestions');
const generateSection = document.getElementById('generateSection');
const generateBtn = document.getElementById('generateBtn');
const generateStatus = document.getElementById('generateStatus');

// 初始化解析器
parser = new PDFParser();

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

// 更新檔案列表顯示
function updateFileList() {
    fileList.innerHTML = '';
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

// 移除檔案
window.removeFile = function(index) {
    pdfFiles.splice(index, 1);
    updateFileList();
    parsedQuestions = [];
    parseSection.style.display = 'none';
    generateSection.style.display = 'none';
    
    // 如果還有檔案，重新解析
    if (pdfFiles.length > 0) {
        setTimeout(() => {
            parsePDFs();
        }, 500);
    }
};

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
        parsedQuestions = await parser.parsePDFs(pdfFiles);
        
        if (parsedQuestions.length === 0) {
            parseStatus.innerHTML = '<div class="status error">未能從 PDF 中提取到任何 MC 題目。請確認 PDF 格式正確。</div>';
            return;
        }

        // 顯示解析結果
        let infoHTML = `<h3>解析完成！</h3><ul>`;
        infoHTML += `<li>總共找到 <strong>${parsedQuestions.length}</strong> 題 MC 題目</li>`;
        infoHTML += `</ul>`;
        parsedQuestionsDiv.innerHTML = infoHTML;
        
        parseStatus.innerHTML = '<div class="status success">✓ PDF 解析成功！</div>';
        generateSection.style.display = 'block';
        
    } catch (error) {
        parseStatus.innerHTML = `<div class="status error">解析失敗：${error.message}</div>`;
        console.error(error);
    }
}

// 生成試卷
generateBtn.addEventListener('click', async () => {
    if (parsedQuestions.length === 0) {
        generateStatus.innerHTML = '<div class="status error">請先上傳並解析 PDF 檔案</div>';
        return;
    }

    const examName = document.getElementById('examName').value || '會計學測驗';
    const totalQuestions = parseInt(document.getElementById('totalQuestions').value, 10);
    const chapterRatio = document.getElementById('chapterRatio').value;

    if (totalQuestions <= 0) {
        generateStatus.innerHTML = '<div class="status error">總題數必須大於 0</div>';
        return;
    }

    if (totalQuestions > parsedQuestions.length) {
        generateStatus.innerHTML = `<div class="status error">總題數 (${totalQuestions}) 超過可用題目數 (${parsedQuestions.length})</div>`;
        return;
    }

    generateStatus.innerHTML = '<div class="status info"><span class="loading"></span>正在生成試卷，請稍候...</div>';
    generateBtn.disabled = true;

    try {
        // 生成題目
        generator = new QuestionGenerator(parsedQuestions);
        const examQuestions = generator.generateExam(totalQuestions, chapterRatio);

        // 生成 Word 文檔
        const wordGen = new WordGenerator();
        
        const questionBlob = await wordGen.generateQuestionSheet(examName, examQuestions);
        wordGen.downloadFile(questionBlob, 'Exam_Questions.docx');

        // 稍等一下再下載答案卷
        setTimeout(async () => {
            const answerBlob = await wordGen.generateAnswerSheet(examName, examQuestions);
            wordGen.downloadFile(answerBlob, 'Exam_Answers.docx');
            
            generateStatus.innerHTML = '<div class="status success">✓ 試卷生成成功！已下載題目卷和答案卷。</div>';
            generateBtn.disabled = false;
        }, 500);

    } catch (error) {
        generateStatus.innerHTML = `<div class="status error">生成失敗：${error.message}</div>`;
        generateBtn.disabled = false;
        console.error(error);
    }
});
