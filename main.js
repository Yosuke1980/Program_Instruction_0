/**
 * ラジオ番組データ取得システム - メインファイル
 * dist/99_main.jsから必要な関数のみを抽出して統合
 */

// =============================================================================
// 設定取得
// =============================================================================

function getConfig() {
    try {
        return CONFIG;
    } catch (error) {
        console.error('config.jsファイルが見つからないか、CONFIG定数が定義されていません。');
        throw new Error('設定ファイルが見つかりません');
    }
}

// =============================================================================
// WebAppエントリーポイント
// =============================================================================

function doGet() {
    return HtmlService.createTemplateFromFile('index')
        .evaluate()
        .setTitle('ラジオ番組データ表示')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =============================================================================
// 日付・週関連ユーティリティ
// =============================================================================

function getMondayOfWeek(date) {
    const dayOfWeek = date.getDay() === 0 ? 7 : date.getDay();
    const monday = new Date(date.getTime() - (dayOfWeek - 1) * 24 * 60 * 60 * 1000);
    monday.setHours(0, 0, 0, 0);
    return monday;
}

function parseSheetDate(sheetName) {
    try {
        const match = sheetName.match(/^(\d{2})\.(\d{1,2})\.(\d{1,2})-/);
        if (!match) return null;

        const year = 2000 + parseInt(match[1]);
        const month = parseInt(match[2]) - 1;
        const day = parseInt(match[3]);

        return new Date(year, month, day);
    } catch (error) {
        console.error(`シート名解析エラー: ${sheetName}`, error);
        return null;
    }
}

function generateWeekSheetName(mondayDate) {
    const year = mondayDate.getFullYear().toString().slice(-2);
    const month = mondayDate.getMonth() + 1;
    const day = mondayDate.getDate();

    const sunday = new Date(mondayDate.getTime() + 6 * 24 * 60 * 60 * 1000);
    const sundayMonth = sunday.getMonth() + 1;
    const sundayDay = sunday.getDate();

    return `${year}.${month}.${day.toString().padStart(2, '0')}-${sundayMonth}.${sundayDay.toString().padStart(2, '0')}`;
}

function findCurrentWeekIndex(spreadsheet) {
    console.log('[FIND-WEEK] 今週のシート検索開始');

    const today = new Date();
    const todayMonday = getMondayOfWeek(today);
    const expectedSheetName = generateWeekSheetName(todayMonday);

    console.log(`[FIND-WEEK] 期待するシート名: ${expectedSheetName}`);

    const allSheets = spreadsheet.getSheets();
    const weekSheets = allSheets.filter(sheet => {
        const name = sheet.getName();
        return name.match(/^\d{2}\.\d{1,2}\.\d{2}-/);
    });

    weekSheets.sort((a, b) => {
        const dateA = parseSheetDate(a.getName());
        const dateB = parseSheetDate(b.getName());
        if (!dateA || !dateB) return 0;
        return dateA.getTime() - dateB.getTime();
    });

    for (let i = 0; i < weekSheets.length; i++) {
        const sheetName = weekSheets[i].getName();
        const sheetDate = parseSheetDate(sheetName);

        if (sheetDate) {
            const sheetMonday = getMondayOfWeek(sheetDate);
            if (sheetMonday.getTime() === todayMonday.getTime()) {
                console.log(`[FIND-WEEK] 今週のシート発見: ${sheetName} (index: ${i})`);
                return i;
            }
        }
    }

    console.warn('[FIND-WEEK] 今週のシートが見つかりません。最も近いシートを選択します');
    let closestIndex = 0;
    let closestDiff = Infinity;

    for (let i = 0; i < weekSheets.length; i++) {
        const sheetDate = parseSheetDate(weekSheets[i].getName());
        if (sheetDate) {
            const sheetMonday = getMondayOfWeek(sheetDate);
            const diff = Math.abs(sheetMonday.getTime() - todayMonday.getTime());
            if (diff < closestDiff) {
                closestDiff = diff;
                closestIndex = i;
            }
        }
    }

    console.warn(`[FIND-WEEK] 最も近いシート: ${weekSheets[closestIndex].getName()} (index: ${closestIndex})`);
    return closestIndex;
}

function getWeekMapping(spreadsheet) {
    const thisWeekIndex = findCurrentWeekIndex(spreadsheet);

    return {
        thisWeek: thisWeekIndex + 1,
        nextWeek: thisWeekIndex + 2,
        nextWeek2: thisWeekIndex + 3,
        nextWeek3: thisWeekIndex + 4,
        nextWeek4: thisWeekIndex + 5
    };
}

// =============================================================================
// スプレッドシートデータ取得
// =============================================================================

function getSheetByWeekNumber(weekNumber) {
    console.log(`[GET-SHEET] 週番号 ${weekNumber} のシート取得開始`);

    const config = getConfig();
    const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
    const allSheets = spreadsheet.getSheets();

    const weekSheets = allSheets.filter(sheet => {
        const name = sheet.getName();
        return name.match(/^\d{2}\.\d{1,2}\.\d{2}-/);
    });

    weekSheets.sort((a, b) => {
        const dateA = parseSheetDate(a.getName());
        const dateB = parseSheetDate(b.getName());
        if (!dateA || !dateB) return 0;
        return dateA.getTime() - dateB.getTime();
    });

    const sheetIndex = weekNumber - 1;
    if (sheetIndex < 0 || sheetIndex >= weekSheets.length) {
        throw new Error(`無効な週番号: ${weekNumber}`);
    }

    const sheet = weekSheets[sheetIndex];
    console.log(`[GET-SHEET] シート取得: ${sheet.getName()}`);

    return sheet;
}

// =============================================================================
// データソート・整形ユーティリティ
// =============================================================================

/**
 * 番組データを曜日順にソート
 * monday, tuesday, wednesday, thursday, friday, saturday, sunday の順に並べる
 * その他のキー（収録予定など）は最後に配置
 */
function sortProgramDataByDays(programData) {
    const dayOrder = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday'];
    const sortedData = {};

    // 1. 曜日データを順番に追加
    dayOrder.forEach(day => {
        if (programData[day]) {
            sortedData[day] = programData[day];
        }
    });

    // 2. その他のキー（収録予定など）を追加
    Object.keys(programData).forEach(key => {
        if (!dayOrder.includes(key)) {
            sortedData[key] = programData[key];
        }
    });

    return sortedData;
}

// =============================================================================
// コアデータ抽出ロジック（dist/99_main.jsから抽出）
// 注意: この部分は CORE_DATA_EXTRACTION_FUNCTIONS.md に詳細が記載されています
// =============================================================================

/**
 * 番組データをJSON形式で生成
 * extractStructuredWeekDataを使用
 */
function generateProgramJSON(programName, weekType) {
    console.log(`[GENERATE-JSON] 番組: ${programName}, 週: ${weekType}`);

    try {
        const config = getConfig();
        const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);

        // 週タイプから週番号に変換
        const weekMapping = getWeekMapping(spreadsheet);
        const weekNumber = weekMapping[weekType];

        if (!weekNumber) {
            throw new Error(`無効な週タイプ: ${weekType}`);
        }

        console.log(`[GENERATE-JSON] 週番号: ${weekNumber}`);

        // シート取得
        const sheet = getSheetByWeekNumber(weekNumber);

        // ★★★ ここで dist/99_main.js の extractStructuredWeekData を呼び出す ★★★
        // この関数は dist/99_main.js に定義されており、
        // スプレッドシートから構造化されたデータを抽出します
        const allProgramData = extractStructuredWeekData(sheet);

        console.log(`[GENERATE-JSON] 全番組データ取得完了`);
        console.log(`[GENERATE-JSON] 番組一覧: ${Object.keys(allProgramData).join(', ')}`);

        // 指定された番組のデータを抽出
        if (!allProgramData[programName]) {
            throw new Error(`番組データが見つかりません: ${programName}`);
        }

        const programData = allProgramData[programName];

        // デバッグ: ソート前のキー順序を確認
        console.log(`[GENERATE-JSON] ソート前のキー順序: ${Object.keys(programData).join(', ')}`);

        // 曜日順にソートされたデータを作成
        const sortedProgramData = sortProgramDataByDays(programData);

        // デバッグ: ソート後のキー順序を確認
        console.log(`[GENERATE-JSON] ソート後のキー順序: ${Object.keys(sortedProgramData).join(', ')}`);

        // 結果を返す
        const result = {
            success: true,
            programName: programName,
            weekType: weekType,
            weekNumber: weekNumber,
            sheetName: sheet.getName(),
            timestamp: new Date().toISOString(),
            data: {
                programName: programName,
                weekData: sortedProgramData
            }
        };

        console.log(`[GENERATE-JSON] JSON生成完了`);
        return result;

    } catch (error) {
        console.error(`[GENERATE-JSON] エラー: ${error.message}`);
        console.error(`[GENERATE-JSON] スタック:`, error.stack);
        return {
            success: false,
            error: error.message,
            programName: programName,
            weekType: weekType,
            timestamp: new Date().toISOString()
        };
    }
}

// =============================================================================
// 公開関数（HTMLから呼び出される）
// =============================================================================

function getProgramData(programName, weekType) {
    return generateProgramJSON(programName, weekType);
}

function getAvailablePrograms() {
    const config = getConfig();
    return Object.keys(config.PROGRAM_STRUCTURE_KEYS);
}

// =============================================================================
// ダミー関数（debugOutputJSON）
// extractStructuredWeekData が依存している場合に備えて
// =============================================================================

function debugOutputJSON(stage, data, context) {
    // シンプルなコンソール出力のみ
    console.log(`[DEBUG-${stage}] ${context}`);
    // 詳細なデバッグが必要な場合はコメントアウトを解除
    // console.log(`[DEBUG-${stage}] Data:`, JSON.stringify(data, null, 2));
}

// =============================================================================
// メール設定管理機能（PropertiesServiceを使用）
// 全番組の設定を1つのJSONとして保存することで容量を効率化
// =============================================================================

const EMAIL_SETTINGS_KEY = 'ALL_EMAIL_SETTINGS';

/**
 * 全番組のメール設定を取得（内部用）
 */
function loadAllEmailSettings() {
    try {
        const props = PropertiesService.getScriptProperties();
        const settingsJson = props.getProperty(EMAIL_SETTINGS_KEY);

        if (settingsJson) {
            return JSON.parse(settingsJson);
        }

        return {};
    } catch (error) {
        console.error('[EMAIL-SETTINGS] 設定読み込みエラー:', error);
        return {};
    }
}

/**
 * 全番組のメール設定を保存（内部用）
 */
function saveAllEmailSettings(allSettings) {
    try {
        const props = PropertiesService.getScriptProperties();
        props.setProperty(EMAIL_SETTINGS_KEY, JSON.stringify(allSettings));
        return true;
    } catch (error) {
        console.error('[EMAIL-SETTINGS] 設定書き込みエラー:', error);
        throw error;
    }
}

/**
 * 番組のメール設定を取得
 */
function getEmailSettings(programName) {
    try {
        const allSettings = loadAllEmailSettings();

        if (allSettings[programName]) {
            return allSettings[programName];
        }

        // デフォルト設定を返す（config.jsの設定があれば使用）
        const config = getConfig();
        const defaultSchedule = config.EMAIL_SETTINGS.schedules[programName];

        if (defaultSchedule) {
            return {
                enabled: false,
                recipients: config.EMAIL_SETTINGS.recipients || '',
                day: defaultSchedule.day,
                hour: defaultSchedule.hour,
                weekType: defaultSchedule.weekType
            };
        }

        // 完全なデフォルト設定
        return {
            enabled: false,
            recipients: '',
            day: 5, // 金曜日
            hour: 9,
            weekType: 'nextWeek'
        };
    } catch (error) {
        console.error(`[EMAIL-SETTINGS] 設定取得エラー: ${programName}`, error);
        return {
            enabled: false,
            recipients: '',
            day: 5,
            hour: 9,
            weekType: 'nextWeek'
        };
    }
}

/**
 * 番組のメール設定を保存
 */
function saveEmailSettings(programName, settings) {
    try {
        // 設定を検証
        if (!settings.recipients || settings.recipients.trim() === '') {
            throw new Error('送信先メールアドレスが指定されていません');
        }

        // 全設定を読み込み
        const allSettings = loadAllEmailSettings();

        // 該当番組の設定を更新
        allSettings[programName] = settings;

        // 保存
        saveAllEmailSettings(allSettings);

        console.log(`[EMAIL-SETTINGS] 設定保存成功: ${programName}`);
        return {
            success: true,
            message: '設定を保存しました'
        };
    } catch (error) {
        console.error(`[EMAIL-SETTINGS] 設定保存エラー: ${programName}`, error);
        return {
            success: false,
            message: 'エラー: ' + error.message
        };
    }
}

/**
 * 全番組のメール設定を取得
 */
function getAllEmailSettings() {
    try {
        const programs = getAvailablePrograms();
        const savedSettings = loadAllEmailSettings();
        const allSettings = {};

        programs.forEach(programName => {
            if (savedSettings[programName]) {
                allSettings[programName] = savedSettings[programName];
            } else {
                allSettings[programName] = getEmailSettings(programName);
            }
        });

        return {
            success: true,
            settings: allSettings
        };
    } catch (error) {
        console.error('[EMAIL-SETTINGS] 全設定取得エラー:', error);
        return {
            success: false,
            error: error.message
        };
    }
}

/**
 * 古い個別プロパティをクリーンアップ（マイグレーション用）
 * 一度だけ実行すれば良い
 */
function cleanupOldEmailSettings() {
    try {
        const props = PropertiesService.getScriptProperties();
        const programs = getAvailablePrograms();
        let deletedCount = 0;

        // すべてのプロパティを取得してログ出力
        const allProps = props.getProperties();
        console.log('[CLEANUP] 現在のプロパティ一覧:', Object.keys(allProps));

        programs.forEach(programName => {
            const oldKey = 'EMAIL_SETTINGS_' + programName;
            if (props.getProperty(oldKey)) {
                props.deleteProperty(oldKey);
                deletedCount++;
                console.log(`[CLEANUP] 削除: ${oldKey}`);
            }
        });

        // その他の不要なプロパティも削除
        const keysToDelete = Object.keys(allProps).filter(key =>
            key.startsWith('EMAIL_SETTINGS_') && key !== 'ALL_EMAIL_SETTINGS'
        );

        keysToDelete.forEach(key => {
            props.deleteProperty(key);
            deletedCount++;
            console.log(`[CLEANUP] 追加削除: ${key}`);
        });

        console.log(`[CLEANUP] クリーンアップ完了: ${deletedCount}件削除`);
        return {
            success: true,
            message: `古い設定を${deletedCount}件削除しました`,
            count: deletedCount
        };
    } catch (error) {
        console.error('[CLEANUP] クリーンアップエラー:', error);
        return {
            success: false,
            message: 'エラー: ' + error.message
        };
    }
}

/**
 * デバッグプロパティを全削除（容量削減用）
 * 重要: 実行前に必要なデータがないか確認してください
 */
function cleanupAllDebugProperties() {
    try {
        const props = PropertiesService.getScriptProperties();
        const allProps = props.getProperties();
        let deletedCount = 0;

        console.log('[CLEANUP-DEBUG] デバッグプロパティ削除開始...');

        // debug_ で始まる全てのプロパティを削除
        Object.keys(allProps).forEach(key => {
            if (key.startsWith('debug_')) {
                props.deleteProperty(key);
                deletedCount++;
                if (deletedCount % 10 === 0) {
                    console.log(`[CLEANUP-DEBUG] ${deletedCount}件削除...`);
                }
            }
        });

        console.log(`[CLEANUP-DEBUG] クリーンアップ完了: ${deletedCount}件のデバッグプロパティを削除`);
        return {
            success: true,
            message: `${deletedCount}件のデバッグプロパティを削除しました`,
            count: deletedCount
        };
    } catch (error) {
        console.error('[CLEANUP-DEBUG] エラー:', error);
        return {
            success: false,
            message: 'エラー: ' + error.message
        };
    }
}

/**
 * 古いデバッグプロパティを自動削除（予防策）
 * 指定日数より古いdebugプロパティを削除
 */
function autoCleanupOldDebugProperties(daysOld = 7) {
    try {
        const props = PropertiesService.getScriptProperties();
        const allProps = props.getProperties();
        let deletedCount = 0;
        const cutoffDate = new Date();
        cutoffDate.setDate(cutoffDate.getDate() - daysOld);

        console.log(`[AUTO-CLEANUP] ${daysOld}日前 (${cutoffDate.toISOString()}) より古いdebugプロパティを削除します`);

        Object.keys(allProps).forEach(key => {
            if (key.startsWith('debug_')) {
                // タイムスタンプを抽出して日付比較
                const timestampMatch = key.match(/(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z)/);
                if (timestampMatch) {
                    const propDate = new Date(timestampMatch[1]);
                    if (propDate < cutoffDate) {
                        props.deleteProperty(key);
                        deletedCount++;
                    }
                } else {
                    // タイムスタンプが抽出できない古い形式は削除
                    props.deleteProperty(key);
                    deletedCount++;
                }
            }
        });

        console.log(`[AUTO-CLEANUP] ${deletedCount}件の古いdebugプロパティを削除しました`);
        return {
            success: true,
            message: `${deletedCount}件の古いdebugプロパティを削除しました`,
            count: deletedCount
        };
    } catch (error) {
        console.error('[AUTO-CLEANUP] エラー:', error);
        return {
            success: false,
            message: 'エラー: ' + error.message
        };
    }
}

/**
 * PropertiesServiceの全プロパティを確認（デバッグ用）
 */
function listAllProperties() {
    try {
        const props = PropertiesService.getScriptProperties();
        const allProps = props.getProperties();

        console.log('[DEBUG] プロパティ一覧:');
        let totalSize = 0;

        Object.keys(allProps).forEach(key => {
            const value = allProps[key];
            // 文字列の長さ（バイト数の概算）
            const size = value.length * 2; // UTF-16で概算
            totalSize += size;
            console.log(`  - ${key}: 約${size} bytes`);
        });

        console.log(`[DEBUG] 合計サイズ: 約${totalSize} bytes (上限: 524,288 bytes)`);
        console.log(`[DEBUG] 使用率: ${(totalSize / 524288 * 100).toFixed(2)}%`);

        return {
            success: true,
            properties: Object.keys(allProps),
            totalSize: totalSize,
            limit: 524288,
            usage: (totalSize / 524288 * 100).toFixed(2) + '%'
        };
    } catch (error) {
        console.error('[DEBUG] エラー:', error);
        return {
            success: false,
            message: error.message
        };
    }
}

/**
 * テストメール送信（WEBアプリから呼び出し）
 */
function sendTestEmail(programName, weekType) {
    try {
        console.log(`[EMAIL-TEST] テストメール送信: ${programName} (${weekType})`);

        // 番組のメール設定を取得
        const settings = getEmailSettings(programName);
        console.log(`[EMAIL-TEST] 取得した設定:`, JSON.stringify(settings));

        if (!settings.recipients || settings.recipients.trim() === '') {
            console.log('[EMAIL-TEST] メールアドレスが未設定');
            return {
                success: false,
                message: 'メールアドレスが設定されていません。メール設定管理タブで送信先を設定してください。'
            };
        }

        // メール送信関数を呼び出し（email_functions.jsの関数）
        console.log(`[EMAIL-TEST] メール送信先: ${settings.recipients}`);
        const result = sendProgramScheduleEmail(programName, weekType, settings.recipients);

        return result;
    } catch (error) {
        console.error('[EMAIL-TEST] テストメール送信エラー:', error);
        return {
            success: false,
            message: 'エラー: ' + error.message
        };
    }
}

// =============================================================================
// PDF生成機能
// =============================================================================

/**
 * 権限承認用のテスト関数
 * Apps Scriptエディタから実行して、DriveとDocumentsの権限を承認してください
 */
function testPDFPermissions() {
    try {
        Logger.log('[TEST] 権限テスト開始');

        // DriveAppの権限テスト
        const files = DriveApp.getFiles();
        Logger.log('[TEST] DriveApp権限: OK');

        // DocumentAppの権限テスト
        const doc = DocumentApp.create('PDF権限テスト');
        const docId = doc.getId();
        Logger.log('[TEST] DocumentApp権限: OK');

        // テストドキュメントを削除
        DriveApp.getFileById(docId).setTrashed(true);
        Logger.log('[TEST] テストドキュメント削除: OK');

        Logger.log('[TEST] 全ての権限テストが成功しました！');
        return '権限テスト成功！PDF機能が使用可能です。';

    } catch (error) {
        Logger.log('[TEST] エラー: ' + error.message);
        throw error;
    }
}

// =============================================================================
// PDF生成機能
// =============================================================================

/**
 * 番組データをPDF化してダウンロード用URLを返す
 * @param {string} programName - 番組名
 * @param {string} weekType - 週タイプ
 * @return {Object} { success: boolean, pdfUrl: string, message: string }
 */
function generateProgramPDF(programName, weekType) {
    try {
        console.log(`[PDF] PDF生成開始: ${programName} (${weekType})`);

        // 番組データを取得
        const result = getProgramData(programName, weekType);

        if (!result.success) {
            console.error('[PDF] データ取得エラー:', result.error);
            return {
                success: false,
                message: 'データ取得エラー: ' + result.error
            };
        }

        // 放送日（月曜日）を8桁形式で取得
        const weekData = result.data.weekData;
        let broadcastDate = '';

        if (weekData.monday && weekData.monday['日付']) {
            const dateStr = weekData.monday['日付'][0]; // 例: "11/11"
            const match = dateStr.match(/(\d{1,2})\/(\d{1,2})/);
            if (match) {
                const month = match[1].padStart(2, '0');
                const day = match[2].padStart(2, '0');
                const year = new Date().getFullYear();
                broadcastDate = `${year}${month}${day}`;
            }
        }

        // タイトル生成: 【連絡票】番組名_YYYYMMDD
        const safeFileName = programName.replace(/\s+/g, '');
        const title = `【連絡票】${safeFileName}_${broadcastDate}`;

        console.log(`[PDF] タイトル: ${title}`);

        // 一時的なGoogle Docを作成
        const doc = DocumentApp.create(title);
        const docId = doc.getId();
        const body = doc.getBody();

        // タイトルを追加
        body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING1);

        // 番組データをドキュメントに追加
        addProgramDataToDoc(body, result);

        // ドキュメントを保存
        doc.saveAndClose();

        // PDFとしてエクスポート
        const driveFile = DriveApp.getFileById(docId);
        const pdfBlob = driveFile.getAs('application/pdf');

        // PDFファイルを作成
        const pdfFile = DriveApp.createFile(pdfBlob);
        pdfFile.setName(title + '.pdf');

        // 共有リンクを取得
        pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        const pdfUrl = pdfFile.getUrl();

        // 一時的なGoogle Docを削除
        DriveApp.getFileById(docId).setTrashed(true);

        console.log(`[PDF] PDF生成成功: ${pdfUrl}`);

        return {
            success: true,
            pdfUrl: pdfUrl,
            message: 'PDF生成成功',
            title: title
        };

    } catch (error) {
        console.error('[PDF] PDF生成エラー:', error);
        return {
            success: false,
            message: 'PDF生成エラー: ' + error.message
        };
    }
}

/**
 * Google Docのボディに番組データを追加
 */
function addProgramDataToDoc(body, result) {
    const weekData = result.data.weekData;
    const dayOrder = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday'];
    const dayLabels = {
        'monday': '月曜日',
        'tuesday': '火曜日',
        'wednesday': '水曜日',
        'thursday': '木曜日',
        'friday': '金曜日',
        'saturday': '土曜日',
        'sunday': '日曜日'
    };

    dayOrder.forEach(day => {
        if (weekData[day]) {
            const dayData = weekData[day];
            const dateStr = dayData['日付'] ? dayData['日付'][0] : '';
            const dayLabel = dayLabels[day] || day;

            // 曜日見出し
            body.appendParagraph(dayLabel + ' ' + dateStr)
                .setHeading(DocumentApp.ParagraphHeading.HEADING2);

            Object.keys(dayData).forEach(key => {
                if (key === '日付') return;

                const value = dayData[key];
                if (!value || value.length === 0) return;

                // カテゴリ見出し
                body.appendParagraph(key)
                    .setHeading(DocumentApp.ParagraphHeading.HEADING3);

                if (Array.isArray(value)) {
                    if (value.length === 1 && value[0] === 'ー') return;

                    if ((key === '楽曲' || key === '指定曲') && typeof value[0] === 'object') {
                        value.forEach((music) => {
                            if (music.曲名) {
                                let text = music.曲名;
                                if (music.URL) {
                                    text += '\n' + music.URL;
                                }
                                if (music.付帯情報) {
                                    text += '\n' + music.付帯情報;
                                }
                                body.appendParagraph(text);
                            }
                        });
                    } else {
                        value.forEach((item) => {
                            body.appendParagraph(item);
                        });
                    }
                } else {
                    body.appendParagraph(String(value));
                }
            });

            // 区切り
            body.appendHorizontalRule();
        }
    });
}
