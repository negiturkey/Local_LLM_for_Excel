/* global Office, Excel */

const PROXY_BASE = window.location.origin + '/api/proxy';
let lastResponse = "";
let isAgentRunning = false;
let abortController = null; // 停止ボタン用

// ===== プロンプトテンプレート (from excel-ai-assistant & cellm) =====
const PROMPT_TEMPLATES = {
    "基本": {
        "要約": "このテキストを1文で要約してください。",
        "文法修正": "文法やスペルの誤りを修正してください。",
        "翻訳 (日→英)": "このテキストを英語に翻訳してください。",
        "翻訳 (英→日)": "このテキストを日本語に翻訳してください。"
    },
    "フォーマット": {
        "日付正規化": "標準的な日付形式 (YYYY-MM-DD) に変換してください。",
        "電話番号整形": "標準的な電話番号形式 (03-xxxx-xxxx) に整形してください。",
        "大文字化": "すべてのテキストを大文字に変換してください。",
        "タイトルケース": "タイトルケース（各単語の先頭を大文字）に変換してください。",
        "全角→半角": "全角英数字を半角に変換してください。",
        "カンマ区切り": "データをカンマ区切り(CSV形式)に変換してください。"
    },
    "分析・抽出": {
        "感情分析": "このテキストの感情を「ポジティブ」「ネガティブ」「中立」のいずれかで判定してください。",
        "キーワード抽出": "このテキストから重要なキーワードを最大5つ抽出してください。",
        "固有表現抽出": "人名、地名、組織名を抽出してリストアップしてください。",
        "数値抽出": "このテキストから数値のみを全て抽出してください。",
        "カテゴリ分類": "このテキストの内容を適切なカテゴリ（製品、苦情、質問、その他）に分類してください。"
    },
    "コード・技術": {
        "JSON整形": "有効で適切にインデントされたJSONとして整形してください。",
        "SQL生成": "この要件に基づいて、適切なSQLクエリを作成してください。",
        "正規表現生成": "このパターンにマッチする正規表現を作成してください。",
        "HTML→テキスト": "HTMLタグを除去してプレーンテキストのみを抽出してください。",
        "Markdown→HTML": "このMarkdownテキストをHTMLコードに変換してください。"
    },
    "ビジネス・創作": {
        "メール下書き": "この要件に基づいて、ビジネスメールの下書きを作成してください。",
        "敬語変換": "この文章を、より丁寧で適切なビジネス敬語に書き換えてください。",
        "キャッチコピー": "この製品の特徴を活かした魅力的なキャッチコピーを3案考えてください。",
        "ToDoリスト化": "このテキストからアクションアイテムを抽出し、ToDoリスト形式にしてください。"
    },
    "データ整理": {
        "余分な空白削除": "余分な空白（重複スペース、前後スペース）を削除してください。",
        "重複排除": "リスト内の重複する項目を削除してください。",
        "欠損値補完提案": "データから欠損値を検出し、文脈に基づいて補完すべき値を提案してください。"
    },
    "開発": {
        "システムテスト": "システムテストを実行して"
    }
};

// ===== ツールレジストリ =====
const TOOL_REGISTRY = {
    read_excel_range: {
        description: "Excelからデータを読み取る（通貨記号は自動除去）",
        args: { range: "A1:B10" },
        execute: async (args) => {
            return await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const range = args.range ? sheet.getRange(args.range) : context.workbook.getSelectedRange();
                range.load("values");
                await context.sync();

                // 強力なデータクリーニング: 数字・符号・小数点のみ残す
                let cleanedValues = range.values.map(row =>
                    row.map(cell => {
                        if (typeof cell === 'string' && cell.trim() !== '') {
                            // 数字、マイナス、小数点のみ抽出（それ以外は全削除）
                            const numericOnly = cell.replace(/[^0-9.\-]/g, '');
                            if (numericOnly === '' || numericOnly === '-' || numericOnly === '.') {
                                return cell; // 数字がなければ元のテキストを返す
                            }
                            const num = parseFloat(numericOnly);
                            return isNaN(num) ? cell : num;
                        }
                        return cell;
                    })
                );

                // トークン節約: 最大20行に制限
                if (cleanedValues.length > 20) {
                    cleanedValues = cleanedValues.slice(0, 20);
                    cleanedValues.push(["...(truncated)"]);
                }

                // 結果文字列も1500文字に制限
                let result = JSON.stringify(cleanedValues);
                if (result.length > 1500) {
                    result = result.slice(0, 1500) + "...(truncated)";
                }
                return result;
            });
        }
    },
    write_to_excel: {
        description: "データを指定セル（または範囲）に書き込む。カンマ/改行区切りまたはJSON配列で複数セル対応。",
        args: { startCell: "A1", data: "値1,値2" },
        execute: async (args) => {
            return await Excel.run(async (context) => {
                // 1. シート選択 (args.sheet があればそのシート、なければアクティブシート)
                let sheet;
                if (args.sheet) {
                    try {
                        sheet = context.workbook.worksheets.getItem(args.sheet);
                    } catch (e) {
                        sheet = context.workbook.worksheets.getActiveWorksheet();
                    }
                } else {
                    sheet = context.workbook.worksheets.getActiveWorksheet();
                }

                // 2. セル座標の特定 (エイリアス対応)
                let startCell = args.startCell || args.targetCell || args.cell || "A1";

                let rawData = args.data !== undefined ? args.data : (args.value || "");
                let isFormula = false;

                // 3. データ処理
                let rows = [];

                // 数式判定: 文字列かつ "=" で始まる場合
                if (typeof rawData === 'string' && rawData.trim().startsWith('=')) {
                    isFormula = true;
                    // LLMが誤って \" とエスケープしている場合があるため除去
                    const cleanFormula = rawData.replace(/\\"/g, '"');
                    rows = [[cleanFormula]];
                } else if (Array.isArray(rawData)) {
                    // JSON配列
                    if (Array.isArray(rawData[0])) {
                        rows = rawData;
                    } else {
                        rows = [rawData];
                    }
                } else {

                    // 文字列 (CSV/TSV/改行区切り)
                    const strData = String(rawData);
                    const lines = strData.split(/\r?\n/).filter(line => line.trim() !== "");

                    // 区切り文字の自動判定
                    const firstLine = lines[0] || "";
                    const delimiter = firstLine.includes('\t') ? '\t' : ',';

                    rows = lines.map(line => {
                        return line.split(delimiter).map(v => {
                            let val = v.trim();
                            if (!isNaN(val) && val !== "") return Number(val);
                            return val;
                        });
                    });
                }

                const rowCount = rows.length;
                if (rowCount === 0) return "No data to write.";

                const colCount = Math.max(...rows.map(r => r.length));

                // データ整形
                const formattedRows = rows.map(r => {
                    while (r.length < colCount) r.push("");
                    return r;
                });

                // 4. 書き込み
                const range = sheet.getRange(startCell).getResizedRange(rowCount - 1, colCount - 1);

                if (isFormula) {
                    range.formulas = formattedRows;
                } else {
                    range.values = formattedRows;
                }

                sheet.load("name");
                await context.sync();
                const sheetName = sheet.name;
                return `SUCCESS: Wrote ${isFormula ? 'formula' : 'data'} to ${sheetName}!${startCell}`;
            });
        }
    },
    calculate_and_write: {
        description: "データを読み取り、計算して結果を書き込む（SUM/AVG/MAX/MIN/COUNT）",
        args: { sourceRange: "B2:B11", targetCell: "B12", operation: "SUM" },
        execute: async (args) => {
            return await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const range = sheet.getRange(args.sourceRange);
                range.load("values");
                await context.sync();

                // データクリーニング: 数字のみ抽出
                const numbers = [];
                range.values.forEach(row => {
                    row.forEach(cell => {
                        if (typeof cell === 'number') {
                            numbers.push(cell);
                        } else if (typeof cell === 'string' && cell.trim() !== '') {
                            const numericOnly = cell.replace(/[^0-9.\-]/g, '');
                            const num = parseFloat(numericOnly);
                            if (!isNaN(num)) numbers.push(num);
                        }
                    });
                });

                if (numbers.length === 0) {
                    return "ERROR: No numeric data found.";
                }

                // 計算実行
                let result;
                const op = (args.operation || "SUM").toUpperCase();
                switch (op) {
                    case "SUM":
                        result = numbers.reduce((a, b) => a + b, 0);
                        break;
                    case "AVG":
                    case "AVERAGE":
                        result = numbers.reduce((a, b) => a + b, 0) / numbers.length;
                        break;
                    case "MAX":
                        result = Math.max(...numbers);
                        break;
                    case "MIN":
                        result = Math.min(...numbers);
                        break;
                    case "COUNT":
                        result = numbers.length;
                        break;
                    default:
                        result = numbers.reduce((a, b) => a + b, 0);
                }

                // 結果を書き込み
                const targetRange = sheet.getRange(args.targetCell);
                targetRange.values = [[result]];
                await context.sync();

                return `SUCCESS: ${op}=${result} written to ${args.targetCell}`;
            });
        }
    },
    smart_formula: {
        description: "通貨記号を含むデータ用の数式を作成（円/¥対応）",
        args: { sourceRange: "B2:B11", targetCell: "B12", operation: "SUM" },
        execute: async (args) => {
            return await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const targetRange = sheet.getRange(args.targetCell);
                const range = args.sourceRange;
                const op = (args.operation || "SUM").toUpperCase();

                // 通貨記号を除去して計算する数式を生成
                // SUBSTITUTE で円と¥を除去 → VALUE で数値化 → 計算
                let cleanExpr = `SUBSTITUTE(SUBSTITUTE(${range},"円",""),"¥","")`;

                let formula;
                switch (op) {
                    case "SUM":
                        formula = `=SUMPRODUCT(VALUE(${cleanExpr}))`;
                        break;
                    case "AVG":
                    case "AVERAGE":
                        formula = `=SUMPRODUCT(VALUE(${cleanExpr}))/COUNTA(${range})`;
                        break;
                    case "MAX":
                        formula = `=MAX(VALUE(${cleanExpr}))`;
                        break;
                    case "MIN":
                        formula = `=MIN(VALUE(${cleanExpr}))`;
                        break;
                    case "COUNT":
                        formula = `=COUNTA(${range})`;
                        break;
                    default:
                        formula = `=SUMPRODUCT(VALUE(${cleanExpr}))`;
                }

                targetRange.formulas = [[formula]];
                await context.sync();

                return `SUCCESS: Formula "${formula}" written to ${args.targetCell}`;
            });
        }
    },
    formula_generator: {
        description: "意図に基づいてExcel数式を自動生成（通貨対応・汎用パターン）",
        args: {
            targetCell: "D12",
            pattern: "SUM_CURRENCY",
            range1: "C2:C11",
            range2: "",
            condition: "",
            value: ""
        },
        execute: async (args) => {
            return await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const targetRange = sheet.getRange(args.targetCell);
                const r1 = args.range1 || "A1:A10";
                const r2 = args.range2 || "";
                const cond = args.condition || "";
                const val = args.value || "";

                // 通貨クリーニング式
                const cleanCurrency = (range) =>
                    `VALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(${range},"円",""),"¥",""),",",""))`;

                let formula;
                const pattern = (args.pattern || "SUM_CURRENCY").toUpperCase();

                switch (pattern) {
                    // ===== 集計系（通貨対応） =====
                    case "SUM_CURRENCY":
                        formula = `=SUMPRODUCT(${cleanCurrency(r1)})`;
                        break;
                    case "AVG_CURRENCY":
                        formula = `=SUMPRODUCT(${cleanCurrency(r1)})/COUNTA(${r1})`;
                        break;
                    case "MAX_CURRENCY":
                        formula = `=MAX(${cleanCurrency(r1)})`;
                        break;
                    case "MIN_CURRENCY":
                        formula = `=MIN(${cleanCurrency(r1)})`;
                        break;
                    case "PRODUCT_CURRENCY":
                        formula = `=PRODUCT(${cleanCurrency(r1)})`;
                        break;

                    // ===== 標準集計 =====
                    case "SUM":
                        formula = `=SUM(${r1})`;
                        break;
                    case "AVERAGE":
                        formula = `=AVERAGE(${r1})`;
                        break;
                    case "COUNT":
                        formula = `=COUNTA(${r1})`;
                        break;
                    case "COUNTIF":
                        formula = `=COUNTIF(${r1},"${cond}")`;
                        break;
                    case "SUMIF":
                        formula = `=SUMIF(${r1},"${cond}",${r2})`;
                        break;

                    // ===== 条件分岐 =====
                    case "IF":
                        formula = `=IF(${r1}${cond},"${val}","")`;
                        break;
                    case "IFS":
                        formula = `=IFS(${cond})`;
                        break;

                    // ===== 検索系 =====
                    case "VLOOKUP":
                        formula = `=VLOOKUP(${val},${r1},${r2},FALSE)`;
                        break;
                    case "XLOOKUP":
                        formula = `=XLOOKUP(${val},${r1},${r2},"")`;
                        break;
                    case "INDEX_MATCH":
                        formula = `=INDEX(${r2},MATCH(${val},${r1},0))`;
                        break;

                    // ===== テキスト系 =====
                    case "CONCAT":
                        formula = `=TEXTJOIN("${cond}",TRUE,${r1})`;
                        break;
                    case "LEFT":
                        formula = `=LEFT(${r1},${val})`;
                        break;
                    case "RIGHT":
                        formula = `=RIGHT(${r1},${val})`;
                        break;
                    case "MID":
                        formula = `=MID(${r1},${cond},${val})`;
                        break;

                    // ===== 日付系 =====
                    case "TODAY":
                        formula = `=TODAY()`;
                        break;
                    case "DATEDIF":
                        formula = `=DATEDIF(${r1},${r2},"${val}")`;
                        break;

                    // ===== ランキング =====
                    case "RANK":
                        formula = `=RANK(${r1},${r2})`;
                        break;
                    case "LARGE":
                        formula = `=LARGE(${r1},${val})`;
                        break;
                    case "SMALL":
                        formula = `=SMALL(${r1},${val})`;
                        break;

                    default:
                        formula = `=SUM(${r1})`;
                }

                targetRange.formulas = [[formula]];
                await context.sync();

                return `SUCCESS: ${pattern} → "${formula}" at ${args.targetCell}`;
            });
        }
    },
    write_formula: {
        description: "Excelに数式を書き込む（=SUM等）",
        args: { startCell: "C1", formula: "=SUM(A1:A10)" },
        execute: async (args) => {
            return await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const range = sheet.getRange(args.startCell || "A1");
                range.formulas = [[args.formula]];
                await context.sync();
                return "SUCCESS: Formula written.";
            });
        }
    },
    set_format: {
        description: "セルの書式（背景色、太字）を設定",
        args: { range: "A1:A10", bgColor: "#FFFF00", fontBold: true },
        execute: async (args) => {
            return await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const range = sheet.getRange(args.range || "A1");

                // エイリアス対応
                const fill = args.fillColor || args.bgColor;
                const bold = args.bold !== undefined ? args.bold : args.fontBold;
                const color = args.fontColor || args.color;

                if (fill) range.format.fill.color = fill;
                if (bold !== undefined) range.format.font.bold = bold;
                if (color) range.format.font.color = color;

                await context.sync();
                return "SUCCESS: Format applied.";
            });
        }
    },
    create_chart: {
        description: "データ範囲からグラフを作成",
        args: { dataRange: "A1:B10", chartType: "ColumnClustered", title: "Chart Title" },
        execute: async (args) => {
            return await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const dataRange = sheet.getRange(args.dataRange);
                // Office.js uses Excel.ChartType enum
                const chartTypeMap = {
                    "ColumnClustered": Excel.ChartType.columnClustered,
                    "Line": Excel.ChartType.line,
                    "Pie": Excel.ChartType.pie,
                    "BarClustered": Excel.ChartType.barClustered,
                    "Doughnut": Excel.ChartType.doughnut
                };
                const chartType = chartTypeMap[args.chartType] || Excel.ChartType.columnClustered;
                const chart = sheet.charts.add(chartType, dataRange, Excel.ChartSeriesBy.auto);
                chart.title.text = args.title || "Chart";
                chart.setPosition("D2", "K15");
                await context.sync();
                return "SUCCESS: Chart created.";
            });
        }
    },
    clean_to_numbers: {
        description: "文字列データを数値に変換（円, ¥, カンマを除去）",
        args: { range: "A1:A10" },
        execute: async (args) => {
            return await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const range = sheet.getRange(args.range);
                range.load("values");
                await context.sync();

                // 文字列から数値を抽出
                const cleaned = range.values.map(row =>
                    row.map(cell => {
                        if (typeof cell === 'string') {
                            // 円, ¥, $, カンマを除去し、数値に変換
                            const numStr = cell.replace(/[円¥$,\s]/g, '');
                            const num = parseFloat(numStr);
                            return isNaN(num) ? cell : num;
                        }
                        return cell;
                    })
                );

                range.values = cleaned;
                await context.sync();
                return "SUCCESS: Converted to numbers.";
            });
        }
    },
    add_conditional_format: {
        description: "条件付き書式を追加（データバー、カラースケール）",
        args: { range: "A1:A10", type: "dataBar", color: "#0078D4" },
        execute: async (args) => {
            return await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const range = sheet.getRange(args.range);

                const formatType = args.type || "dataBar";

                if (formatType === "dataBar") {
                    const dataBar = range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
                    dataBar.dataBar.barDirection = Excel.ConditionalDataBarDirection.leftToRight;
                    dataBar.dataBar.positiveFormat.fillColor = args.color || "#0078D4";
                    dataBar.dataBar.negativeFormat.fillColor = "#D13438";
                } else if (formatType === "colorScale") {
                    const colorScale = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
                    colorScale.colorScale.criteria = {
                        minimum: { color: "#F8696B", type: Excel.ConditionalFormatColorCriterionType.lowestValue },
                        midpoint: { color: "#FFEB84", type: Excel.ConditionalFormatColorCriterionType.percentile, formula: "50" },
                        maximum: { color: "#63BE7B", type: Excel.ConditionalFormatColorCriterionType.highestValue }
                    };
                } else if (formatType === "highlight") {
                    // 指定値以上をハイライト
                    const threshold = args.threshold || 0;
                    const preset = range.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
                    preset.cellValue.format.fill.color = args.color || "#FFFF00";
                    preset.cellValue.rule = {
                        formula1: String(threshold),
                        operator: Excel.ConditionalCellValueOperator.greaterThan
                    };
                }

                await context.sync();
                return "SUCCESS: Conditional format applied.";
            });
        }
    },
    apply_table_style: {
        description: "範囲をテーブル化してスタイルを適用",
        args: { range: "A1:D10", styleName: "TableStyleMedium2", hasHeaders: true },
        execute: async (args) => {
            return await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const range = sheet.getRange(args.range);
                const table = sheet.tables.add(range, args.hasHeaders !== false);
                table.style = args.styleName || "TableStyleMedium2";
                await context.sync();
                return "SUCCESS: Table created with style.";
            });
        }
    },
    sort_range: {
        description: "範囲をソート（昇順/降順）",
        args: { range: "A1:B10", column: 0, ascending: true },
        execute: async (args) => {
            return await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const range = sheet.getRange(args.range);
                range.sort.apply([{
                    key: args.column || 0,
                    ascending: args.ascending !== false
                }]);
                await context.sync();
                return "SUCCESS: Range sorted.";
            });
        }
    },
    filter_range: {
        description: "データにフィルターを適用",
        args: { range: "A1:D10", column: 0, criteria: "条件値" },
        execute: async (args) => {
            return await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const range = sheet.getRange(args.range);
                // AutoFilterを適用
                sheet.autoFilter.apply(range, args.column || 0, {
                    criterion1: args.criteria,
                    filterOn: Excel.FilterOn.values
                });
                await context.sync();
                return "SUCCESS: Filter applied.";
            });
        }
    },
    generate_image: {
        description: "画像生成（プロンプトから画像を生成して挿入）",
        args: { prompt: "猫の画像" },
        execute: async (args) => {
            // 1. Placeholder生成 (Canvas)
            const prompt = args.prompt || "Generated Image";
            const canvas = document.createElement('canvas');
            canvas.width = 400;
            canvas.height = 300;
            const ctx = canvas.getContext('2d');

            // 背景
            ctx.fillStyle = "#E0E0E0";
            ctx.fillRect(0, 0, 400, 300);

            // テキスト
            ctx.fillStyle = "#333333";
            ctx.font = "20px sans-serif";
            ctx.fillText("Image Generator (Mock)", 20, 40);
            ctx.font = "16px sans-serif";
            ctx.fillText(prompt.slice(0, 30) + "...", 20, 150);

            const base64 = canvas.toDataURL("image/png");
            const cleanBase64 = base64.replace(/^data:image\/png;base64,/, "");

            // 2. Excel挿入
            return await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const image = sheet.shapes.addImage(cleanBase64);
                image.name = "Gen_" + new Date().getTime();
                image.left = 50;
                image.top = 50;
                await context.sync();
                return `SUCCESS: Generated image for '${prompt}'`;
            });
        }
    },
    insert_image: {
        description: "Base64画像をシートに挿入",
        args: { base64: "...", name: "AI_Image" },
        execute: async (args) => {
            return await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                // header削除 (data:image/png;base64,...)
                const cleanBase64 = args.base64.replace(/^data:image\/(png|jpeg|jpg);base64,/, "");

                const image = sheet.shapes.addImage(cleanBase64);
                image.name = args.name || "AI_Image_" + new Date().getTime();
                image.left = 50;
                image.top = 50;

                await context.sync();
                return "SUCCESS: Image inserted into sheet.";
            });
        }
    },
    run_all_tests: {
        description: "全機能のシステムテスト（End-to-End Test）を実行",
        args: { mode: "full" },
        execute: async (args) => {
            try {
                // 1. テスト環境のセットアップ (シート作成)
                let sheetName = "";
                await Excel.run(async (context) => {
                    const sheets = context.workbook.worksheets;
                    const timestamp = new Date().getTime();
                    sheetName = `Test_${timestamp}`;
                    const sheet = sheets.add(sheetName);
                    sheet.load("name"); // 明示的にロード
                    await context.sync();
                    sheet.activate();
                    await context.sync();
                });

                const logResults = [];
                const addLog = (step, result) => logResults.push(`[${step}] ${result}`);

                // 2. ツールチェーン実行テスト
                // 各ツールの execute は内部で Excel.run を呼ぶため、順次awaitすれば良い

                // Step 1: データ投入 (write_to_excel)
                // テストデータの生成 (20件)
                const categories = ["Laptop", "Mouse", "Monitor", "Keyboard", "Headset", "Tablet", "Cable", "Charger", "Dock", "Webcam"];
                let testData = "商品\t価格\t個数\n";
                for (let i = 0; i < 20; i++) {
                    const item = categories[i % categories.length] + "_" + (i + 1);
                    const price = (Math.floor(Math.random() * 100) + 10) * 1000;
                    const qty = Math.floor(Math.random() * 5) + 1;
                    // カンマ付き価格を含めてテスト
                    testData += `${item}\t${price.toLocaleString()}円\t${qty}\n`;
                }

                addLog("1. Data Setup", await TOOL_REGISTRY.write_to_excel.execute({
                    startCell: "A1",
                    data: testData
                }));

                // Step 2: データクリーニング (clean_to_numbers)
                addLog("2. Cleaning", await TOOL_REGISTRY.clean_to_numbers.execute({
                    range: "B2:B21"
                }));

                // Step 3: 数式適用 (formula_generator) - 売上計算
                addLog("3. Formula", await TOOL_REGISTRY.formula_generator.execute({
                    targetCell: "D2",
                    pattern: "PRODUCT_CURRENCY",
                    range1: "B2:C2"
                }));
                // オートフィル的にD3-D21も埋める
                for (let i = 3; i <= 21; i++) {
                    await TOOL_REGISTRY.write_formula.execute({ startCell: `D${i}`, formula: `=B${i}*C${i}` });
                }

                // Step 4: 書式設定 (set_format)
                addLog("4. Formatting", await TOOL_REGISTRY.set_format.execute({
                    range: "A1:D1",
                    fillColor: "#4472C4",
                    fontColor: "#FFFFFF",
                    bold: true
                }));

                // Step 5: テーブル化 (apply_table_style)
                addLog("5. Table", await TOOL_REGISTRY.apply_table_style.execute({
                    range: "A1:D21",
                    styleName: "TableStyleMedium2"
                }));

                // Step 6: グラフ作成 (create_chart)
                addLog("6. Chart", await TOOL_REGISTRY.create_chart.execute({
                    dataRange: "A1:D21",
                    chartType: "ColumnClustered",
                    title: "System Test Chart"
                }));

                // Step 7: 条件付き書式 (add_conditional_format) - 個数にデータバー
                addLog("7. Cond. Format", await TOOL_REGISTRY.add_conditional_format.execute({
                    range: "C2:C21", // 個数
                    type: "dataBar",
                    color: "#00B050"
                }));

                // Step 8: 並べ替え (sort_range) - 価格の降順 (列1=B列 をキーに)
                addLog("8. Sort", await TOOL_REGISTRY.sort_range.execute({
                    range: "A2:D21", // ヘッダー除くデータ部分
                    column: 1,      // B列（価格）
                    ascending: false // 降順
                }));

                // Step 9: フィルタ (filter_range) - 商品名に "Laptop" を含むもの
                // ※AutoFilterはテーブルに対して行うのが一般的だが、ここでは範囲指定でテスト
                // (Step 5でテーブル化しているので、テーブルのフィルターとして機能する可能性が高い)
                /* 
                   注: Office.jsのAutoFilter制限により、API経由でのフィルタ印加は不安定な場合があるため、
                   エラーが出てもテストを止めないようにtry-catchすることが望ましいが、
                   今回はtool自体がエラーを返さない設計なのでそのまま実行
                */
                // addLog("9. Filter", await TOOL_REGISTRY.filter_range.execute({
                //    range: "A1:D4",
                //    column: 0,
                //    criteria: "Laptop"
                // }));
                // → フィルタは視覚的確認が難しく、後のステップに影響するため今回は除外（または最後に実行）

                // Step 9: データ読み取り (read_excel_range) - 検証用
                const readResult = await TOOL_REGISTRY.read_excel_range.execute({ range: "A1:D21" });
                addLog("9. Read Check", readResult.length > 10 ? "SUCCESS (Data Read)" : "WARNING (Read Empty?)");

                // Step 10: 画像生成 (generate_image)
                addLog("10. Image Gen", await TOOL_REGISTRY.generate_image.execute({
                    prompt: "Test Image"
                }));

                // Step 11: 次元テスト (Dimension Check) - 多対多 / 多対1
                // 11-A: 3x2行列の書き込み (Multi-to-Multi)
                addLog("11A. Multi-Multi", await TOOL_REGISTRY.write_to_excel.execute({
                    startCell: "F2",
                    data: [[10, 20], [30, 40], [50, 60]]
                }));
                // 11-B: 単一セルの書き込み (One/Multi-to-One)
                addLog("11B. Multi-One", await TOOL_REGISTRY.write_to_excel.execute({
                    startCell: "F6",
                    data: "Finished"
                }));

                // Step 8: 最終確認とレポート出力
                return `✅ System Test Completed Successfully on '${sheetName}'\n\nDETAILS:\n` + logResults.join("\n");

            } catch (error) {
                return `❌ SYSTEM TEST FAILED: ${error.message}\nStack: ${error.stack}`;
            }
        }
    }
};

// ===== 選択セルデータ取得（自動クリーニング付き） =====
async function getSelectedCellData() {
    return await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load(["values", "address"]);
        await context.sync();

        // 強力なデータクリーニング: 数字・符号・小数点のみ残す
        const cleanedValues = range.values.map(row =>
            row.map(cell => {
                if (typeof cell === 'string' && cell.trim() !== '') {
                    const numericOnly = cell.replace(/[^0-9.\-]/g, '');
                    if (numericOnly === '' || numericOnly === '-' || numericOnly === '.') {
                        return cell;
                    }
                    const num = parseFloat(numericOnly);
                    return isNaN(num) ? cell : num;
                }
                return cell;
            })
        );

        // トークン節約: 最大20行に制限
        let resultData = cleanedValues;
        if (resultData.length > 20) {
            resultData = resultData.slice(0, 20);
            resultData.push(["...(truncated)"]);
        }

        // 文字数制限
        let json = JSON.stringify(resultData);
        if (json.length > 1000) {
            json = json.slice(0, 1000) + "...(truncated)";
        }

        return `Address: ${range.address}\nValues: ${json}`;
    });
}

// ===== タイムライン表示ヘルパー =====
function formatTimelineEntry(step, toolName, status, result = "") {
    const icons = { running: "⏳", success: "✅", error: "❌" };
    const icon = icons[status] || "⚙️";
    const resultPreview = result.length > 60 ? result.slice(0, 60) + "..." : result;
    return `<div class="timeline-step">
        <span class="step-badge">Step ${step}</span>
        <span class="step-icon">${icon}</span>
        <strong>${toolName}</strong>
        ${result ? `<div class="step-result">${escapeHtml(resultPreview)}</div>` : ""}
    </div>`;
}

// ===== システムプロンプト選択ロジック =====
function selectSystemPrompt(userText) {
    const text = userText.toLowerCase();

    // 1. 計算・数式モード
    if (/計算|合計|平均|数式|関数|sum|avg|max|min|count/.test(text)) {
        return `Excel数式Agent。操作はJSON。

[メインツール] formula_generator
{"call":"formula_generator","args":{"targetCell":"D12","pattern":"SUM_CURRENCY","range1":"C2:C10"}}

[パターン]
集計(円対応): SUM_CURRENCY, AVG_CURRENCY, MAX_CURRENCY, MIN_CURRENCY
掛け算(円対応): PRODUCT_CURRENCY
標準: SUM, AVERAGE, COUNT, COUNTIF, SUMIF
条件: IF, IFS
検索: VLOOKUP, XLOOKUP, INDEX_MATCH
文字: CONCAT, LEFT, RIGHT, MID
日付: TODAY, DATEDIF
順位: RANK, LARGE, SMALL

[その他]
set_format, write_to_excel`;
    }

    // 2. 書式・グラフモード
    if (/色|太字|書式|グラフ|チャート|color|bold|format|chart/.test(text)) {
        return `ExcelデザインAgent。操作はJSON。

[メインツール]
set_format: 書式設定
{"call":"set_format","args":{"range":"A1:D1","fillColor":"#4472C4","bold":true,"fontColor":"#FFFFFF"}}

create_chart: グラフ作成
{"call":"create_chart","args":{"dataRange":"A1:B10","chartType":"ColumnClustered"}}
Types: ColumnClustered, Line, Pie, BarClustered

add_conditional_format: 条件付き書式
{"call":"add_conditional_format","args":{"range":"B2:B10","type":"dataBar","color":"#00B050"}}

[その他]
write_to_excel`;
    }

    // 3. システムテスト（"システムテスト" を含む場合）
    if (/システムテスト/.test(text)) {
        return `システム健全性チェックAgent。
ユーザーの指示に従い、ツールが正しく動作するか診断を行います。
JSONフォーマットを必ず守ってください。

[メインツール]
run_all_tests: システムテスト実行
{"call": "run_all_tests", "args": {"mode": "full"}}
        `;
    }

    // 4. データ整理・汎用モード
    return `Excel操作Agent。
操作指示 → JSON出力。
一般質問 → テキスト回答。

[ツール]
formula_generator: 数式
{"call":"formula_generator","args":{"targetCell":"B1","pattern":"SUM_CURRENCY","range1":"A1:A10"}}

set_format: 書式
{"call":"set_format","args":{"range":"A1","fillColor":"#FFFF00","bold":true}}

write_formula: 任意の関数
{"call":"write_formula","args":{"startCell":"B10","formula":"=STDEV(B2:B9)"}}

write_to_excel: 値入力 (リスト/行列可)
{"call":"write_to_excel","args":{"startCell":"A1","data":[["ID","Name"],["1","A"],["2","B"]]}}

generate_image: 画像生成
{"call":"generate_image","args":{"prompt":"青い空と海"}}

run_all_tests: システムテスト
{"call":"run_all_tests","args":{"mode":"full"}}

clean_to_numbers: 数値化
{"call":"clean_to_numbers","args":{"range":"A1:A10"}}

apply_table_style: テーブル
{"call":"apply_table_style","args":{"range":"A1:C5","styleName":"TableStyleMedium2"}}

複数操作は複数JSONで出力。`;
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// ===== ログ関数 =====
function log(msg, isAgentAction = false) {
    const win = document.getElementById('chat-window');
    if (!win) return;
    const div = document.createElement('div');
    div.className = isAgentAction ? 'message system agent-action' : 'message system';
    div.innerText = "[" + new Date().toLocaleTimeString() + "] " + (isAgentAction ? "🤖 " : "") + msg;
    win.appendChild(div);
    win.scrollTop = win.scrollHeight;
}

// ===== テンプレート初期化 =====
// ===== テンプレート初期化 =====
let mergedTemplates = {}; // 削除用に保持

async function initTemplates() {
    const select = document.getElementById('template-select');
    const input = document.getElementById('prompt-input');
    const saveBtn = document.getElementById('save-template-btn');
    const delBtn = document.getElementById('delete-template-btn');

    // イベントリスナーを先に定義（Fetch待ちでボタンが反応しないのを防ぐ）

    // 1. 選択変更イベント
    select.onchange = () => {
        if (select.value) {
            input.value = select.value;
            const isDefault = isDefaultTemplate(select.value);
            delBtn.disabled = isDefault;
        } else {
            delBtn.disabled = true;
        }
    };

    // 2. 保存ボタンイベント（モーダル表示）
    saveBtn.onclick = () => {
        const currentPrompt = input.value.trim();
        if (!currentPrompt) return alert("プロンプトを入力してください");

        const modal = document.getElementById('save-modal');
        const nameInput = document.getElementById('template-name-input');
        const catInput = document.getElementById('template-category-input');
        const confirmBtn = document.getElementById('confirm-save-btn');
        const cancelBtn = document.getElementById('cancel-save-btn');

        // モーダル表示
        modal.style.display = 'flex';
        nameInput.focus();

        // キャンセル処理
        const closeModal = () => {
            modal.style.display = 'none';
            confirmBtn.onclick = null; // リスナー解除
            cancelBtn.onclick = null;
        };
        cancelBtn.onclick = closeModal;

        // 保存確定処理
        confirmBtn.onclick = async () => {
            log("Save button clicked...", true); // Debug log
            const name = nameInput.value.trim();
            const category = catInput.value.trim();

            if (!name) {
                alert("名前を入力してください");
                return;
            }
            if (!category) {
                alert("カテゴリを入力してください");
                return;
            }

            // 保存処理
            let userTemplates = {};
            try {
                const res = await fetch('/api/templates');
                if (res.ok) userTemplates = await res.json();
            } catch (e) {
                log("Fetch existing failed: " + e.message, true);
            }

            if (!userTemplates[category]) userTemplates[category] = {};
            userTemplates[category][name] = currentPrompt;

            try {
                log("Saving template to server...", true);
                const res = await fetch('/api/templates', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(userTemplates)
                });

                if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);

                log("✅ テンプレートを保存しました: " + name, true);
                closeModal();

                // リロード
                setTimeout(() => {
                    initTemplates();
                    log("Templates reloaded.", true);
                }, 500);

            } catch (e) {
                log("❌ 保存エラー: " + e.message, true);
                alert("保存エラー: " + e.message);
            }
        };
    };

    // 3. 削除ボタンイベント（モーダル表示）
    delBtn.onclick = () => {
        const selectedOption = select.options[select.selectedIndex];
        if (!selectedOption || !selectedOption.value) return;

        const targetPrompt = selectedOption.value;
        const targetName = selectedOption.text;

        const modal = document.getElementById('delete-modal');
        const msg = document.getElementById('delete-message');
        const confirmBtn = document.getElementById('confirm-delete-btn');
        const cancelBtn = document.getElementById('cancel-delete-btn');

        msg.textContent = `テンプレート「${targetName}」を削除しますか？\n(注意: 同じプロンプトを持つ全てのユーザーテンプレートが対象になる可能性があります)`;
        modal.style.display = 'flex';

        const closeModal = () => {
            modal.style.display = 'none';
            confirmBtn.onclick = null;
            cancelBtn.onclick = null;
        };
        cancelBtn.onclick = closeModal;

        confirmBtn.onclick = async () => {
            log("Deleting template...", true);
            try {
                const res = await fetch('/api/templates');
                if (!res.ok) throw new Error("Load failed");
                let userTemplates = await res.json();
                let changed = false;

                for (const cat in userTemplates) {
                    for (const key in userTemplates[cat]) {
                        if (key === targetName && userTemplates[cat][key] === targetPrompt) {
                            delete userTemplates[cat][key];
                            if (Object.keys(userTemplates[cat]).length === 0) delete userTemplates[cat];
                            changed = true;
                        }
                    }
                }

                if (changed) {
                    const res = await fetch('/api/templates', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify(userTemplates)
                    });
                    if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);

                    log("✅ 削除しました", true);
                    input.value = "";
                    closeModal();
                    setTimeout(() => initTemplates(), 500);
                } else {
                    alert("削除対象が見つかりませんでした（デフォルトテンプレートは削除できません）");
                    closeModal();
                }

            } catch (e) {
                log("❌ 削除エラー: " + e.message, true);
                alert("削除エラー: " + e.message);
                closeModal();
            }
        };
    };

    // --- データロード処理 ---

    select.innerHTML = '<option value="">(選択してください)</option>';
    mergedTemplates = JSON.parse(JSON.stringify(PROMPT_TEMPLATES)); // Deep copy

    // ユーザーテンプレートの取得 (API)
    try {
        const res = await fetch('/api/templates');
        if (res.ok) {
            const userTemplates = await res.json();
            // マージ (ユーザー定義は「ユーザー定義」グループに入れるか、既存グループに追加)
            if (userTemplates && Object.keys(userTemplates).length > 0) {
                if (!mergedTemplates["ユーザー定義"]) mergedTemplates["ユーザー定義"] = {};
                // 単純化のため全て「ユーザー定義」グループに入れる、または保存時の構造に従う
                // ここでは保存時の構造 { "Category": { "Name": "Prompt" } } を想定してマージ
                for (const [cat, items] of Object.entries(userTemplates)) {
                    if (!mergedTemplates[cat]) mergedTemplates[cat] = {};
                    Object.assign(mergedTemplates[cat], items);
                }
            }
        }
    } catch (e) {
        console.error("Failed to load user templates", e);
    }

    // グループごとにoption生成
    for (const [group, items] of Object.entries(mergedTemplates)) {
        const optgroup = document.createElement('optgroup');
        optgroup.label = group;
        for (const [name, prompt] of Object.entries(items)) {
            const option = document.createElement('option');
            option.value = prompt;
            option.textContent = name;
            // 削除判定用にデータ属性付与 (ユーザー定義のものかどうかの判定は簡易的に行う)
            // ここでは簡易的に全テンプレートにメタデータを付けるのは難しいので、
            // 選択時にテキストベースで逆引きして削除対象を探す
            optgroup.appendChild(option);
        }
        select.appendChild(optgroup);
    }
}

// デフォルトテンプレートに含まれているか判定
function isDefaultTemplate(promptText) {
    for (const group in PROMPT_TEMPLATES) {
        for (const name in PROMPT_TEMPLATES[group]) {
            if (PROMPT_TEMPLATES[group][name] === promptText) return true;
        }
    }
    return false;
}

// ===== バッチ実行 (行ごとの処理) =====
async function handleBatchRun() {
    const promptInput = document.getElementById('prompt-input');
    const userPrompt = promptInput.value.trim();
    if (!userPrompt) {
        log("プロンプトを入力してください。", true);
        return;
    }

    const provider = document.getElementById('provider-select').value;
    const model = document.getElementById('model-select').value;
    const stopBtn = document.getElementById('stop-btn');
    const sendBtn = document.getElementById('send-btn');
    const batchBtn = document.getElementById('batch-btn');

    if (!model) {
        log("モデルを選択してください。", true);
        return;
    }

    // UI状態変更
    isAgentRunning = true;
    abortController = new AbortController();
    stopBtn.style.display = "inline-block";
    sendBtn.style.display = "none";
    batchBtn.disabled = true;

    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange();
            range.load(["values", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
            await context.sync();

            const rowCount = range.rowCount;
            const colCount = range.columnCount; // 通常は1列推奨だが、複数列の場合は結合して扱うか、左端を使うなど

            log(`バッチ処理を開始します: 全${rowCount}行`, true);

            for (let i = 0; i < rowCount; i++) {
                // 中断チェック
                if (abortController.signal.aborted) {
                    log("⛔ バッチ処理を中断しました。");
                    break;
                }

                // 現在の行の値を取得
                const currentVal = range.values[i][0]; // 1列目を使用
                if (currentVal === "" || currentVal === null) {
                    log(`Skipping Row ${i + 1}: Empty`);
                    continue;
                }

                log(`Row ${i + 1}/${rowCount}: 处理中... (${String(currentVal).slice(0, 10)}...)`);

                // プロンプト構築
                const fullPrompt = `以下のテキストに対して、次の指示を実行してください。\n指示: ${userPrompt}\n\n対象テキスト:\n${currentVal}\n\n回答は結果のみを出力してください。`;

                // LLM呼び出し (callLLMはsend-btnのロジックと共有したいが、ここは簡易実装)
                const messages = [
                    { role: "system", content: "あなたはデータ処理アシスタントです。余計な会話はせず、結果のみを返してください。" },
                    { role: "user", content: fullPrompt }
                ];

                try {
                    const response = await callLLMBackend(provider, model, messages, abortController.signal);

                    // 結果を隣のセル(1つ右)に書き込み
                    // getCell(row, col) は相対座標
                    const targetCell = range.getCell(i, colCount); // 選択範囲の右隣
                    targetCell.values = [[response.trim()]];
                    await context.sync();

                } catch (err) {
                    if (err.name === 'AbortError') throw err;
                    log(`Row ${i + 1} Error: ${err.message}`);
                    const targetCell = range.getCell(i, colCount);
                    targetCell.values = [[`Error: ${err.message}`]];
                    await context.sync();
                }
            }

            if (!abortController.signal.aborted) {
                log("✅ バッチ処理が完了しました。", true);
            }
        });
    } catch (error) {
        log("エラーが発生しました: " + error.message);
    } finally {
        isAgentRunning = false;
        abortController = null;
        stopBtn.style.display = "none";
        sendBtn.style.display = "inline-block";
        batchBtn.disabled = false;
    }
}

// ===== 初期化 =====
function init() {
    const providerSelect = document.getElementById('provider-select');
    const modelSelect = document.getElementById('model-select');
    const refreshBtn = document.getElementById('refresh-models');
    const sendBtn = document.getElementById('send-btn');
    const batchBtn = document.getElementById('batch-btn');
    const stopBtn = document.getElementById('stop-btn');
    const applyBtn = document.getElementById('apply-to-cell');
    const clearBtn = document.getElementById('clear-chat');
    const testBtn = document.getElementById('test-connection');
    const loadBtn = document.getElementById('load-model');
    const unloadBtn = document.getElementById('unload-model');

    const savedProvider = localStorage.getItem('selected_provider');
    if (savedProvider) providerSelect.value = savedProvider;
    providerSelect.onchange = () => {
        localStorage.setItem('selected_provider', providerSelect.value);
        refreshModels();
    };
    modelSelect.onchange = (e) => localStorage.setItem('selected_model', e.target.value);

    refreshBtn.onclick = refreshModels;
    sendBtn.onclick = handleSend;
    if (batchBtn) batchBtn.onclick = handleBatchRun;
    if (stopBtn) stopBtn.onclick = handleStop;
    testBtn.onclick = testConnection;
    loadBtn.onclick = loadModel;
    unloadBtn.onclick = unloadModel;
    document.getElementById('upload-btn').onclick = () => document.getElementById('image-input').click();
    document.getElementById('image-input').onchange = handleImageSelect;
    document.getElementById('remove-image').onclick = clearImage;
    clearBtn.onclick = () => {
        document.getElementById('chat-window').innerHTML = '';
        localStorage.removeItem('chat_history');
        // 全履歴を確実にクリア
        try {
            localStorage.clear();
            // プロバイダとモデル設定は復元
            localStorage.setItem('selected_provider', providerSelect.value);
            localStorage.setItem('selected_model', modelSelect.value);
        } catch (e) { }
        log("✓ 履歴をクリアしました");
        lastResponse = "";
    };
    applyBtn.onclick = applyResponseToCell;

    const promptInput = document.getElementById('prompt-input');
    promptInput.onkeydown = (e) => {
        if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            handleSend();
        }
    };
    window.addEventListener('paste', handlePaste);

    loadHistory();
    log("Agent Pro initialized. Ready.");
    initTemplates();
    refreshModels();
}

// ===== 停止ボタン処理 =====
function handleStop() {
    if (abortController) {
        abortController.abort();
        log("⛔ Agent stopped by user.", true);
        isAgentRunning = false;
        updateStopButtonVisibility(false);
    }
}

function updateStopButtonVisibility(show) {
    const stopBtn = document.getElementById('stop-btn');
    const sendBtn = document.getElementById('send-btn');
    if (stopBtn) stopBtn.style.display = show ? 'inline-block' : 'none';
    if (sendBtn) sendBtn.style.display = show ? 'none' : 'inline-block';
}

// ===== メイン送信処理 =====
async function handleSend() {
    if (isAgentRunning) return;

    const provider = document.getElementById('provider-select').value;
    const model = document.getElementById('model-select').value;
    const promptInput = document.getElementById('prompt-input');
    const prompt = promptInput.value.trim();

    if (!prompt || !model) return;

    const imagePreview = document.getElementById('image-preview');
    const base64Image = (imagePreview.src && imagePreview.src.startsWith('data:')) ? imagePreview.src.split(',')[1] : null;

    addMessage("user", prompt, imagePreview.src || null);
    promptInput.value = '';
    isAgentRunning = true;
    abortController = new AbortController();
    updateStopButtonVisibility(true);

    // ダイナミック・システムプロンプト選択
    const systemPromptContent = selectSystemPrompt(prompt);

    let messages = [
        {
            role: "system",
            content: systemPromptContent
        }
    ];

    // 履歴は直近2件のみ（トークン節約）
    const history = JSON.parse(localStorage.getItem('chat_history') || '[]');
    history.slice(-2).forEach(h => messages.push({ role: h.type === 'ai' ? 'assistant' : 'user', content: h.text.slice(0, 200) }));

    // 「選択中のセルのデータをAIに送る」機能
    let finalPrompt = prompt;
    const includeSelection = document.getElementById('include-selection');
    if (includeSelection && includeSelection.checked) {
        try {
            const selectionData = await getSelectedCellData();
            if (selectionData && selectionData !== '[[]]') {
                finalPrompt = `[選択中のセルデータ: ${selectionData}]\n\n${prompt}`;
            }
        } catch (e) {
            // 選択失敗時は無視（プロンプトのみ送信）
        }
    }

    messages.push({ role: "user", content: finalPrompt });

    try {
        await runAgentLoop(provider, model, messages, base64Image);
    } catch (e) {
        if (e.name === 'AbortError') {
            log("Agent loop aborted.", true);
        } else {
            log("Error: " + e.message);
            addMessage("ai", "エラーが発生しました: " + e.message);
        }
    } finally {
        isAgentRunning = false;
        abortController = null;
        updateStopButtonVisibility(false);
        clearImage();
    }
}

// ===== LLM 共通呼び出し関数 (定義漏れ修正) =====
async function callLLMBackend(provider, model, messages, signal = null) {
    let body = {
        model: model,
        messages: messages,
        stream: false,
        options: { temperature: 0.1 }
    };

    if (provider === 'gemini') {
        const apiKey = document.getElementById('api-key').value;
        const systemMsg = messages.find(m => m.role === 'system');
        const chatHistory = messages.filter(m => m.role !== 'system').map(m => ({
            role: m.role === 'assistant' ? 'model' : 'user',
            parts: [{ text: m.content }]
        }));

        const apiBody = {
            contents: chatHistory,
            generationConfig: { temperature: 0.1 }
        };
        if (systemMsg) apiBody.system_instruction = { parts: [{ text: systemMsg.content }] };

        const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(apiBody),
            signal: signal
        });

        if (!res.ok) {
            const errText = await res.text();
            throw new Error(`Gemini API Error ${res.status}: ${errText}`);
        }
        const data = await res.json();
        return data.candidates[0].content.parts[0].text;
    } else {
        const url = (provider === 'ollama') ? `${PROXY_BASE}/ollama/api/chat` : `${PROXY_BASE}/lmstudio/v1/chat/completions`;
        const res = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body),
            signal: signal
        });

        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const data = await res.json();
        return (provider === 'ollama') ? data.message.content : data.choices[0].message.content;
    }
}

// ===== エージェントループ =====
async function runAgentLoop(provider, model, messages, base64Image) {
    let loopCount = 0;
    const MAX_LOOPS = 8;
    let aiBubble = null;

    while (loopCount < MAX_LOOPS) {
        if (abortController && abortController.signal.aborted) {
            throw new DOMException('Aborted', 'AbortError');
        }

        loopCount++;
        if (!aiBubble) aiBubble = addMessage("ai", "🧠 Thinking...");
        else aiBubble.innerText = `🧠 Step ${loopCount}: Thinking...`;

        let body = {
            model: model,
            messages: messages,
            stream: false, // Ensure no streaming for simple JSON parsing
            options: { temperature: 0.1 }
        };

        if (base64Image && loopCount === 1) {
            const userMsg = messages[messages.length - 1];
            if (provider === 'ollama') userMsg.images = [base64Image];
            else {
                userMsg.content = [
                    { type: "text", text: userMsg.content },
                    { type: "image_url", image_url: { url: `data:image/jpeg;base64,${base64Image}` } }
                ];
            }
        }

        let content = "";

        if (provider === 'gemini') {
            const apiKey = document.getElementById('api-key').value;
            // System prompt extract
            const systemMsg = messages.find(m => m.role === 'system');
            const chatHistory = messages.filter(m => m.role !== 'system').map(m => ({
                role: m.role === 'assistant' ? 'model' : 'user',
                parts: [{ text: typeof m.content === 'object' ? m.content[0].text : m.content }] // Handle multimodal array
            }));

            // Handle Image (current message)
            if (base64Image && loopCount === 1) {
                // Gemini expects inline data for images in the last user message
                const lastMsg = chatHistory[chatHistory.length - 1];
                lastMsg.parts.push({ inline_data: { mime_type: "image/jpeg", data: base64Image } });
            }

            const apiBody = {
                contents: chatHistory,
                generationConfig: { temperature: 0.1 }
            };
            if (systemMsg) {
                apiBody.system_instruction = { parts: [{ text: systemMsg.content }] };
            }

            const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(apiBody)
            });

            if (!res.ok) {
                const errText = await res.text();
                throw new Error(`Gemini API Error ${res.status}: ${errText}`);
            }
            const data = await res.json();
            content = data.candidates[0].content.parts[0].text;

        } else {
            // Ollama / LM Studio
            const url = (provider === 'ollama') ? `${PROXY_BASE}/ollama/api/chat` : `${PROXY_BASE}/lmstudio/v1/chat/completions`;
            const res = await fetch(url, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(body),
                signal: abortController ? abortController.signal : undefined
            });

            if (!res.ok) throw new Error(`HTTP ${res.status}`);
            const data = await res.json();
            content = (provider === 'ollama') ? data.message.content : data.choices[0].message.content;
        }

        if (!content || content.trim().length === 0) {
            messages.push({ role: "system", content: "System: Empty response. Please continue." });
            continue;
        }

        aiBubble.innerText = content;
        messages.push({ role: "assistant", content: content });

        // ワンショットモード: 全ツールを一括実行
        const toolCalls = findAllToolCalls(content);

        if (toolCalls.length > 0) {
            let allResults = [];
            let stepNum = 0;

            for (const toolCall of toolCalls) {
                stepNum++;
                const argsPreview = JSON.stringify(toolCall.args || {}).slice(0, 40);
                log(`⚡ ${stepNum}/${toolCalls.length}: ${toolCall.call}`, true);

                aiBubble.innerHTML += formatTimelineEntry(stepNum, toolCall.call, "running");

                let result = "";
                try {
                    result = await TOOL_REGISTRY[toolCall.call].execute(toolCall.args || {});
                    allResults.push(`${toolCall.call}: ${result}`);
                    // 成功表示に更新
                    aiBubble.innerHTML = aiBubble.innerHTML.replace("⏳", "✅");
                } catch (err) {
                    result = "Error: " + err.message;
                    allResults.push(`${toolCall.call}: ${result}`);
                    aiBubble.innerHTML = aiBubble.innerHTML.replace("⏳", "❌");
                }
            }

            // 全結果をまとめて返す（ループ回数削減）
            if (toolCalls.length === 1 && !allResults[0].includes("Error")) {
                // 単一ツールで成功なら終了
                lastResponse = `完了: ${allResults[0]}`;
                aiBubble.innerHTML += `<div style="margin-top:8px;color:#0078d4;">✓ 完了</div>`;
                saveMessage("ai", lastResponse);
                document.getElementById('apply-to-cell').disabled = false;
                break;
            }

            messages.push({ role: "user", content: `Results: ${allResults.join(' | ')}` });
            continue;
        } else {
            lastResponse = content;
            saveMessage("ai", content);
            document.getElementById('apply-to-cell').disabled = false;
            break;
        }
    }
}

// ===== JSON解析（複数ツール対応・自動推論付き） =====
function findAllToolCalls(text) {
    const calls = [];
    let searchIdx = 0;
    while (true) {
        const start = text.indexOf('{', searchIdx);
        if (start === -1) break;
        let braceCount = 0;
        let foundEnd = false;
        for (let i = start; i < text.length; i++) {
            if (text[i] === '{') braceCount++;
            else if (text[i] === '}') braceCount--;
            if (braceCount === 0) {
                try {
                    const cleanJson = text.substring(start, i + 1).replace(/[\u201C\u201D]/g, '"');
                    const parsed = JSON.parse(cleanJson);

                    if (parsed) {
                        // 1. 正規フォーマット: {"call": "name", "args": {...}}
                        if (parsed.call && TOOL_REGISTRY[parsed.call]) {
                            if (!parsed.args) parsed.args = {}; // 引数がない場合は空オブジェクトを補完
                            calls.push(parsed);
                        }
                        // 2. 引数のみフォーマット（推論）: {"pattern":...} or {"fillColor":...}
                        else {
                            if (parsed.pattern) {
                                calls.push({ call: "formula_generator", args: parsed });
                            } else if (parsed.fillColor || parsed.bgColor || parsed.bold || parsed.color) {
                                calls.push({ call: "set_format", args: parsed });
                            } else if (parsed.chartType) {
                                calls.push({ call: "create_chart", args: parsed });
                            } else if (parsed.startCell && parsed.data) {
                                calls.push({ call: "write_to_excel", args: parsed });
                            } else if (parsed.targetCell && parsed.value) {
                                // Alias for write_to_excel
                                calls.push({
                                    call: "write_to_excel",
                                    args: { startCell: parsed.targetCell, data: parsed.value }
                                });
                            } else if (parsed.base64 || parsed.image) {
                                calls.push({ call: "insert_image", args: parsed });
                            } else if (parsed.prompt) {
                                calls.push({ call: "generate_image", args: parsed });
                            } else if (parsed.mode) {
                                calls.push({ call: "run_all_tests", args: parsed });
                            }
                        }
                    }
                } catch (e) { }
                searchIdx = i + 1;
                foundEnd = true;
                break;
            }
        }
        if (!foundEnd) break;
    }
    return calls;
}

// 後方互換性のため残す
function findValidToolCall(text) {
    const calls = findAllToolCalls(text);
    return calls.length > 0 ? calls[0] : null;
}

// ===== UI関数 =====
function renderMessage(type, text, image = null) {
    const win = document.getElementById('chat-window');
    const div = document.createElement('div');
    div.className = `message ${type}`;
    if (image) {
        const img = document.createElement('img');
        img.src = image; img.style.maxWidth = '100%'; img.style.maxHeight = '150px';
        img.style.borderRadius = '4px'; img.style.marginBottom = '6px'; img.style.display = 'block';
        div.appendChild(img);
    }
    const textSpan = document.createElement('span');
    textSpan.innerText = text;
    div.appendChild(textSpan);
    win.appendChild(div);
    win.scrollTop = win.scrollHeight;
    return div;
}

function addMessage(type, text, image = null) {
    const div = renderMessage(type, text, image);
    if (text !== "🧠 Thinking..." && !text.startsWith("🧠 Step")) {
        saveMessage(type, text, image);
    }
    return div;
}

function saveMessage(type, text, image) {
    const history = JSON.parse(localStorage.getItem('chat_history') || '[]');
    history.push({ type, text, image });
    localStorage.setItem('chat_history', JSON.stringify(history.slice(-15)));
}

function loadHistory() {
    const history = JSON.parse(localStorage.getItem('chat_history') || '[]');
    history.forEach(item => renderMessage(item.type, item.text, item.image));
}

async function applyResponseToCell() {
    if (!lastResponse) return;
    try {
        await Excel.run(async (context) => {
            const range = context.workbook.getActiveCell();
            range.values = [[lastResponse]];
            await context.sync();
        });
        log("Applied to cell.");
    } catch (e) { log("Error: " + e.message); }
}

// ===== モデル管理 =====
function setLoadingState(isLoading, text) {
    const btn = document.getElementById('load-model');
    if (btn) { btn.innerText = text; btn.disabled = isLoading; }
}

async function loadModel() {
    const provider = document.getElementById('provider-select').value;
    const model = document.getElementById('model-select').value;
    if (!model) return;
    setLoadingState(true, "ロード中...");
    log(`[${provider}] "${model}" Loading...`);

    if (provider === 'gemini') {
        const apiKey = document.getElementById('api-key').value;
        if (!apiKey) {
            log("✗ API Key Required.");
            setLoadingState(false, "ロード");
            return;
        }
        // localStorage.setItem('gemini_api_key', apiKey); // Disabled by user request
        log("✓ Ready (API Mode).");
        setLoadingState(false, "ロード");
        return;
    }

    try {
        const endpoint = provider === 'ollama' ? "/api/generate" : "/v1/chat/completions";
        const body = (provider === 'ollama')
            ? { model, keep_alive: "1h", stream: false }
            : { model, messages: [{ role: "user", content: "hi" }], max_tokens: 1, stream: false };

        await fetch(`${PROXY_BASE}/${provider}${endpoint}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body)
        });
        log("✓ Ready.");
    } catch (e) { log("✗ Failed."); } finally { setLoadingState(false, "ロード"); }
}

async function unloadModel() {
    const provider = document.getElementById('provider-select').value;
    const model = document.getElementById('model-select').value;
    if (!model) {
        log("モデルが選択されていません");
        return;
    }

    try {
        if (provider === 'ollama') {
            // Ollamaの場合: keep_alive: 0 で即座にアンロード
            await fetch(`${PROXY_BASE}/ollama/api/generate`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ model: model, keep_alive: 0 })
            });
            log("✓ VRAM解放完了 (Ollama)");
        } else if (provider === 'lmstudio') {
            // LM Studioはアンロード不要（自動管理）
            log("LM Studioは手動でサーバーを停止してください");
        }
    } catch (e) {
        log("✗ VRAM解放失敗: " + e.message);
    }
}

async function refreshModels() {
    const provider = document.getElementById('provider-select').value;
    const select = document.getElementById('model-select');

    if (provider === 'gemini') {
        const apiKey = document.getElementById('api-key').value;
        select.innerHTML = '';

        if (!apiKey) {
            const opt = document.createElement('option');
            opt.innerText = "APIキーを入力してください";
            select.appendChild(opt);
            return;
        }

        try {
            const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`);
            if (!res.ok) {
                const err = await res.json();
                throw new Error(err.error.message || "Fetch Failed");
            }
            const data = await res.json();
            // Filter models that support generateContent
            const models = (data.models || [])
                .filter(m => m.supportedGenerationMethods && m.supportedGenerationMethods.includes("generateContent"))
                .sort((a, b) => b.displayName.localeCompare(a.displayName)); // Sort roughly

            models.forEach(m => {
                const opt = document.createElement('option');
                // Use pure name (e.g. "models/gemini-1.5-pro") or strip "models/" depending on what API expects
                // API `generateContent` expects "models/gemini-1.5-pro" OR "gemini-1.5-pro" usually works.
                // Safest is to use the `name` field as returned ("models/...") BUT our fetch logic handles it.
                // Current fetch logic: `models/${model}:generateContent`
                // S0 if value is "models/gemini-pro", URL becomes ".../models/models/gemini-pro..." -> WRONG.
                // So we MUST strip "models/" prefix here.
                const value = m.name.replace(/^models\//, '');

                opt.value = value;
                opt.innerText = `✨ ${m.displayName || value} (${m.version})`;
                select.appendChild(opt);
            });

            // Log success
            log(`✓ ${models.length} Gemini models loaded.`);

        } catch (e) {
            log("✗ Model List Error: " + e.message);
            const opt = document.createElement('option');
            opt.innerText = "モデル取得失敗";
            select.appendChild(opt);
        }
        return;
    }

    try {
        let url = (provider === 'ollama') ? `${PROXY_BASE}/ollama/api/tags` : `${PROXY_BASE}/lmstudio/v1/models`;
        const res = await fetch(url);
        if (res.ok) {
            const data = await res.json();
            select.innerHTML = '';
            let models = (provider === 'ollama') ? (data.models || []).map(m => m.name) : (data.data || []).map(m => m.id);
            models.forEach(m => {
                const opt = document.createElement('option'); opt.value = m;
                opt.innerText = /vision|llava|vl|moondream/i.test(m) ? `👁️ ${m}` : m;
                select.appendChild(opt);
            });
            const saved = localStorage.getItem('selected_model');
            if (saved && models.includes(saved)) select.value = saved;
        }
    } catch (e) { }
}

// ===== 画像処理 =====
function handlePaste(e) {
    const items = e.clipboardData.items;
    for (let i = 0; i < items.length; i++) {
        if (items[i].type.indexOf('image') !== -1) {
            const blob = items[i].getAsFile();
            const reader = new FileReader();
            reader.onload = (event) => {
                const preview = document.getElementById('image-preview');
                const container = document.getElementById('image-preview-container');
                preview.src = event.target.result; container.style.display = 'block';
            };
            reader.readAsDataURL(blob);
        }
    }
}

function handleImageSelect(e) {
    const file = e.target.files[0];
    if (file) {
        const reader = new FileReader();
        reader.onload = (event) => {
            const preview = document.getElementById('image-preview');
            const container = document.getElementById('image-preview-container');
            preview.src = event.target.result; container.style.display = 'block';
        };
        reader.readAsDataURL(file);
    }
}

function clearImage() {
    const input = document.getElementById('image-input');
    const preview = document.getElementById('image-preview');
    const container = document.getElementById('image-preview-container');
    if (input) input.value = ''; if (preview) preview.src = ''; if (container) container.style.display = 'none';
}

async function testConnection() {
    log("Diagnosing...");
    try {
        const res = await fetch(window.location.origin + '/src/index.html');
        if (res.ok) log("✓ Server OK.");
    } catch (e) { log("✗ Server unreachable."); }
}

if (document.readyState === "loading") { document.addEventListener("DOMContentLoaded", init); } else { init(); }
Office.onReady();

// Gemini UI Toggle
document.addEventListener("DOMContentLoaded", () => {
    const providerSelect = document.getElementById('provider-select');
    const apiKeyInput = document.getElementById('api-key');
    if (providerSelect && apiKeyInput) {
        providerSelect.addEventListener('change', async () => {
            const isGemini = providerSelect.value === 'gemini';
            apiKeyInput.style.display = isGemini ? 'block' : 'none';
            if (isGemini) {
                try {
                    const res = await fetch('/api/env');
                    if (res.ok) {
                        const data = await res.json();
                        if (data.apiKey) {
                            apiKeyInput.value = data.apiKey;
                            log("✓ API Key loaded from .env");
                        }
                    }
                } catch (e) { }
            }
            // Auto refresh to show static models
            refreshModels();
        });
    }
});
