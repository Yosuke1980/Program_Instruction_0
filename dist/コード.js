// Google Apps Script用にCONFIGを直接定義
const CONFIG = {
    SPREADSHEET_ID: '1r6GLEvsZiqb3vkXZmrZ7XCpvW4RcecC7Bfp9YzhnQUE',
    CALENDAR_ID: 'c_8e4bd1f0e76b0a56a42c74b17b439c68e7da6ad78f13bab24461a0ac56df2d5c@group.calendar.google.com',
    MUSIC_SPREADSHEET_ID: '1F6iTNBENB9vZV5sCF4K2DbK0XUW3UaaKpjic3UEmWhE',
    MUSIC_SHEET_NAME: '指定曲リスト',
    EMAIL_ADDRESS: 'watanabe@fmyokohama.co.jp',
    // Googleドキュメントテンプレート設定
    DOCUMENT_TEMPLATES: {
        'ちょうどいいラジオ': '1rIW5gT20G974PgBC4IXyqjwW0OltItJpEmg2UpL8Ods',
        'PRIME TIME': '1z0iEKbsIvxYC0OcnOHgebB6p1nZj5TY6aDFxSSN0OhA',
        'FLAG': '12gqWH7aIG34vmFJZ5LnEPc21Qg_fFYI4VTmaEnCGAxc',
        'God Bless Saturday': '1UaWzLXif5NgiE079m7D51GegBYry9s-ge6P_n175kXY',
        'Route 847': '1GtlZEYnY6RTqv8qtLRnSFMJXHj4OunjYIRHLhA_uU-g'
    },
    DOCUMENT_OUTPUT_FOLDER_ID: '1-YJXe7YJzevvUUrcp-VlEyKO3YHLH4qf',
    // PORTSIDE専用カレンダー
    PORTSIDE_CALENDAR_ID: 'c_a25ffe304a25338178cd461a7b159103bd46ea62c64e773fdf4c61b4b36ab2fc@group.calendar.google.com'
};
// Google Apps Scriptのconsoleの型を修正（module modeでない場合はコメントアウト）
// declare global {
//   interface Console {
//     log(...args: any[]): void;
//     error(...args: any[]): void;
//   }
// }
/**
 * WebApp用のGAS追加コード
 * 既存のCode.gsファイルに以下の関数を追加してください
 */
/**
 * WebApp用のdoGet関数 - HTMLページを表示
 */
function doGet() {
    return HtmlService.createTemplateFromFile('index')
        .evaluate()
        .setTitle('ラジオ番組ドキュメント自動生成')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
/**
 * HTMLファイルにCSSやJSを含める関数
 */
function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
/**
 * WebApp用のエラーハンドリング付きラッパー関数群
 */
/**
 * ちょうどいいラジオ - WebApp用
 */
function webAppAutoGenerateChoudo() {
    try {
        console.log('WebApp: ちょうどいいラジオ実行開始');
        const result = autoGenerateChoudoDocument();
        console.log('WebApp: ちょうどいいラジオ実行完了', result);
        return result;
    }
    catch (error) {
        console.error('WebApp: ちょうどいいラジオ実行エラー', error);
        return {
            success: false,
            error: error.toString(),
            timestamp: new Date().toISOString()
        };
    }
}
/**
 * PRIME TIME - WebApp用
 */
function webAppAutoGeneratePrimeTime() {
    try {
        console.log('WebApp: PRIME TIME実行開始');
        const result = autoGeneratePrimeTimeDocument();
        console.log('WebApp: PRIME TIME実行完了', result);
        return result;
    }
    catch (error) {
        console.error('WebApp: PRIME TIME実行エラー', error);
        return {
            success: false,
            error: error.toString(),
            timestamp: new Date().toISOString()
        };
    }
}
/**
 * FLAG - WebApp用
 */
function webAppAutoGenerateFlag() {
    try {
        console.log('WebApp: FLAG実行開始');
        const result = autoGenerateFlagDocument();
        console.log('WebApp: FLAG実行完了', result);
        return result;
    }
    catch (error) {
        console.error('WebApp: FLAG実行エラー', error);
        return {
            success: false,
            error: error.toString(),
            timestamp: new Date().toISOString()
        };
    }
}
/**
 * Route 847 - WebApp用
 */
function webAppAutoGenerateRoute847() {
    try {
        console.log('WebApp: Route 847実行開始');
        const result = autoGenerateRoute847Document();
        console.log('WebApp: Route 847実行完了', result);
        return result;
    }
    catch (error) {
        console.error('WebApp: Route 847実行エラー', error);
        return {
            success: false,
            error: error.toString(),
            timestamp: new Date().toISOString()
        };
    }
}
/**
 * God Bless Saturday - WebApp用
 */
function webAppAutoGenerateGodBless() {
    try {
        console.log('WebApp: God Bless Saturday実行開始');
        const result = autoGenerateGodBlessDocument();
        console.log('WebApp: God Bless Saturday実行完了', result);
        return result;
    }
    catch (error) {
        console.error('WebApp: God Bless Saturday実行エラー', error);
        return {
            success: false,
            error: error.toString(),
            timestamp: new Date().toISOString()
        };
    }
}
/**
 * テスト実行 - WebApp用
 */
function webAppTestAllGeneration() {
    try {
        console.log('WebApp: テスト実行開始');
        const results = testAllAutoGeneration();
        console.log('WebApp: テスト実行完了', results);
        return results;
    }
    catch (error) {
        console.error('WebApp: テスト実行エラー', error);
        return [{
                program: 'テスト実行',
                success: false,
                error: error.toString(),
                timestamp: new Date().toISOString()
            }];
    }
}
/**
 * 実行状況確認 - WebApp用
 */
function webAppGetExecutionStatus() {
    try {
        const config = getConfig();
        // 基本的な設定確認
        const status = {
            configValid: !!config,
            spreadsheetAccess: false,
            emailConfigured: !!config.EMAIL_ADDRESS,
            templatesConfigured: !!config.DOCUMENT_TEMPLATES,
            timestamp: new Date().toISOString()
        };
        // スプレッドシートアクセス確認
        try {
            const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
            status.spreadsheetAccess = !!spreadsheet;
            status.spreadsheetName = spreadsheet.getName();
        }
        catch (error) {
            status.spreadsheetError = error.toString();
        }
        return status;
    }
    catch (error) {
        console.error('WebApp: ステータス取得エラー', error);
        return {
            success: false,
            error: error.toString(),
            timestamp: new Date().toISOString()
        };
    }
}
/**
 * 設定情報取得 - WebApp用（機密情報は除外）
 */
function webAppGetConfigInfo() {
    try {
        const config = getConfig();
        // 機密情報を除外した設定情報
        const publicConfig = {
            hasSpreadsheetId: !!config.SPREADSHEET_ID,
            hasCalendarId: !!config.CALENDAR_ID,
            hasMusicSpreadsheetId: !!config.MUSIC_SPREADSHEET_ID,
            musicSheetName: config.MUSIC_SHEET_NAME,
            hasEmailAddress: !!config.EMAIL_ADDRESS,
            hasDocumentTemplates: !!config.DOCUMENT_TEMPLATES,
            templateCount: config.DOCUMENT_TEMPLATES ? Object.keys(config.DOCUMENT_TEMPLATES).length : 0,
            hasOutputFolder: !!config.DOCUMENT_OUTPUT_FOLDER_ID,
            hasPortsideCalendar: !!config.PORTSIDE_CALENDAR_ID,
            timestamp: new Date().toISOString()
        };
        return publicConfig;
    }
    catch (error) {
        console.error('WebApp: 設定情報取得エラー', error);
        return {
            success: false,
            error: error.toString(),
            timestamp: new Date().toISOString()
        };
    }
}
/**
 * ログ記録用関数
 */
function webAppLog(action, data) {
    try {
        const logData = {
            timestamp: new Date().toISOString(),
            action: action,
            data: data,
            user: Session.getActiveUser().getEmail()
        };
        console.log('WebApp Log:', JSON.stringify(logData));
        // 必要に応じてスプレッドシートやCloud Loggingに記録
        // 現在はコンソールログのみ
        return { success: true };
    }
    catch (error) {
        console.error('WebApp: ログ記録エラー', error);
        return { success: false, error: error.toString(), docxSent: false };
    }
}
/**
 * ラジオ番組スケジュール表から指定番組の内容を構造化して抽出するGASコード
 *
 * 必要ファイル：
 * 1. このメインファイル（Code.gs）
 * 2. config.gs - 設定情報
 *
 * config.gsの内容：
 * const CONFIG = {
 *   SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID_HERE',
 *   CALENDAR_ID: 'YOUR_CALENDAR_ID_HERE',
 *   MUSIC_SPREADSHEET_ID: 'YOUR_MUSIC_SPREADSHEET_ID_HERE',
 *   MUSIC_SHEET_NAME: 'シート1',
 *   EMAIL_ADDRESS: 'your-email@example.com',
 *   // Googleドキュメントテンプレート設定
 *   DOCUMENT_TEMPLATES: {
 *     'ちょうどいいラジオ': 'TEMPLATE_DOC_ID_FOR_CHOUDO',
 *     'PRIME TIME': 'TEMPLATE_DOC_ID_FOR_PRIMETIME',
 *     'FLAG': 'TEMPLATE_DOC_ID_FOR_FLAG',
 *     'God Bless Saturday': 'TEMPLATE_DOC_ID_FOR_GODBLESS',
 *     'Route 847': 'TEMPLATE_DOC_ID_FOR_ROUTE847'
 *   },
 *   DOCUMENT_OUTPUT_FOLDER_ID: 'YOUR_OUTPUT_FOLDER_ID_HERE',
 *   // PORTSIDE専用カレンダー
 *   PORTSIDE_CALENDAR_ID: 'YOUR_PORTSIDE_CALENDAR_ID_HERE'
 * };
 */
/**
 * 設定情報を取得
 */
function getConfig() {
    try {
        return CONFIG;
    }
    catch (error) {
        console.error('config.gsファイルが見つからないか、CONFIG定数が定義されていません。');
        console.error('config.gsファイルを作成し、以下の内容を記載してください：');
        console.error('const CONFIG = {');
        console.error('  SPREADSHEET_ID: "YOUR_SPREADSHEET_ID_HERE",');
        console.error('  CALENDAR_ID: "YOUR_CALENDAR_ID_HERE",');
        console.error('  MUSIC_SPREADSHEET_ID: "YOUR_MUSIC_SPREADSHEET_ID_HERE",');
        console.error('  MUSIC_SHEET_NAME: "シート1",');
        console.error('  EMAIL_ADDRESS: "your-email@example.com"');
        console.error('};');
        throw new Error('設定ファイルが見つかりません');
    }
}
/**
 * 今週のみを抽出（週別表示）
 */
function extractThisWeek() {
    const config = getConfig();
    const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
    const thisWeekSheet = getSheetByWeek(spreadsheet, 0);
    if (!thisWeekSheet) {
        console.log('今週のシートが見つかりません');
        return {};
    }
    console.log('Processing this week:', thisWeekSheet.getName());
    const results = { '今週': extractStructuredWeekData(thisWeekSheet) };
    logStructuredResults(results);
    return results;
}
/**
 * 今週のみを抽出（番組別表示）
 */
function extractThisWeekByProgram() {
    const config = getConfig();
    const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
    const thisWeekSheet = getSheetByWeek(spreadsheet, 0);
    if (!thisWeekSheet) {
        console.log('今週のシートが見つかりません');
        return {};
    }
    console.log('Processing this week:', thisWeekSheet.getName());
    const results = { '今週': extractStructuredWeekData(thisWeekSheet) };
    logResultsByProgram(results);
    return results;
}
/**
 * 今週のみを抽出（番組別表示）してメール送信
 */
function extractThisWeekByProgramAndSendEmail() {
    const config = getConfig();
    const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
    const thisWeekSheet = getSheetByWeek(spreadsheet, 0);
    if (!thisWeekSheet) {
        console.log('今週のシートが見つかりません');
        // エラーメールを送信
        try {
            if (config.EMAIL_ADDRESS) {
                GmailApp.sendEmail(config.EMAIL_ADDRESS, '今週の番組スケジュール（番組別） - エラー', '今週のシートが見つかりませんでした。スプレッドシートの構成を確認してください。');
            }
        }
        catch (error) {
            console.error('エラーメール送信失敗:', error);
        }
        return {};
    }
    try {
        console.log('Processing this week:', thisWeekSheet.getName());
        const weekData = extractStructuredWeekData(thisWeekSheet);
        if (!weekData || typeof weekData !== 'object') {
            console.error('週データの抽出に失敗しました');
            return {};
        }
        const results = { '今週': weekData };
        logResultsByProgram(results);
        sendProgramEmail(results, '今週の番組スケジュール（番組別）');
        return results;
    }
    catch (error) {
        console.error('データ抽出エラー:', error);
        // エラーメールを送信
        try {
            if (config.EMAIL_ADDRESS) {
                GmailApp.sendEmail(config.EMAIL_ADDRESS, '今週の番組スケジュール（番組別） - エラー', `データ抽出中にエラーが発生しました。\nエラー: ${error.toString()}`);
            }
        }
        catch (emailError) {
            console.error('エラーメール送信失敗:', emailError);
        }
        return {};
    }
}
/**
 * 来週のみを抽出（番組別表示）
 */
function extractNextWeekByProgram() {
    const config = getConfig();
    const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
    const nextWeekSheet = getSheetByWeek(spreadsheet, 1);
    if (!nextWeekSheet) {
        console.log('来週のシートが見つかりません');
        return {};
    }
    console.log('Processing next week:', nextWeekSheet.getName());
    const results = { '来週': extractStructuredWeekData(nextWeekSheet) };
    logResultsByProgram(results);
    return results;
}
/**
 * 来週のみを抽出（番組別表示）してメール送信
 */
function extractNextWeekByProgramAndSendEmail() {
    const config = getConfig();
    const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
    const nextWeekSheet = getSheetByWeek(spreadsheet, 1);
    if (!nextWeekSheet) {
        console.log('来週のシートが見つかりません');
        // エラーメールを送信
        try {
            if (config.EMAIL_ADDRESS) {
                GmailApp.sendEmail(config.EMAIL_ADDRESS, '来週の番組スケジュール（番組別） - エラー', '来週のシートが見つかりませんでした。スプレッドシートの構成を確認してください。');
            }
        }
        catch (error) {
            console.error('エラーメール送信失敗:', error);
        }
        return {};
    }
    try {
        console.log('Processing next week:', nextWeekSheet.getName());
        const weekData = extractStructuredWeekData(nextWeekSheet);
        if (!weekData || typeof weekData !== 'object') {
            console.error('週データの抽出に失敗しました');
            return {};
        }
        const results = { '来週': weekData };
        logResultsByProgram(results);
        sendProgramEmail(results, '来週の番組スケジュール（番組別）');
        return results;
    }
    catch (error) {
        console.error('データ抽出エラー:', error);
        // エラーメールを送信
        try {
            if (config.EMAIL_ADDRESS) {
                GmailApp.sendEmail(config.EMAIL_ADDRESS, '来週の番組スケジュール（番組別） - エラー', `データ抽出中にエラーが発生しました。\nエラー: ${error.toString()}`);
            }
        }
        catch (emailError) {
            console.error('エラーメール送信失敗:', emailError);
        }
        return {};
    }
}
/**
 * 先週のみを抽出（番組別表示）
 */
function extractLastWeekByProgram() {
    const config = getConfig();
    const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
    const lastWeekSheet = getSheetByWeek(spreadsheet, -1);
    if (!lastWeekSheet) {
        console.log('先週のシートが見つかりません');
        return {};
    }
    console.log('Processing last week:', lastWeekSheet.getName());
    const results = { '先週': extractStructuredWeekData(lastWeekSheet) };
    logResultsByProgram(results);
    return results;
}
/**
 * 先週のみを抽出（番組別表示）してメール送信
 */
function extractLastWeekByProgramAndSendEmail() {
    const config = getConfig();
    const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
    const lastWeekSheet = getSheetByWeek(spreadsheet, -1);
    if (!lastWeekSheet) {
        console.log('先週のシートが見つかりません');
        // エラーメールを送信
        try {
            if (config.EMAIL_ADDRESS) {
                GmailApp.sendEmail(config.EMAIL_ADDRESS, '先週の番組スケジュール（番組別） - エラー', '先週のシートが見つかりませんでした。スプレッドシートの構成を確認してください。');
            }
        }
        catch (error) {
            console.error('エラーメール送信失敗:', error);
        }
        return {};
    }
    try {
        console.log('Processing last week:', lastWeekSheet.getName());
        const weekData = extractStructuredWeekData(lastWeekSheet);
        if (!weekData || typeof weekData !== 'object') {
            console.error('週データの抽出に失敗しました');
            return {};
        }
        const results = { '先週': weekData };
        logResultsByProgram(results);
        sendProgramEmail(results, '先週の番組スケジュール（番組別）');
        return results;
    }
    catch (error) {
        console.error('データ抽出エラー:', error);
        // エラーメールを送信
        try {
            if (config.EMAIL_ADDRESS) {
                GmailApp.sendEmail(config.EMAIL_ADDRESS, '先週の番組スケジュール（番組別） - エラー', `データ抽出中にエラーが発生しました。\nエラー: ${error.toString()}`);
            }
        }
        catch (emailError) {
            console.error('エラーメール送信失敗:', emailError);
        }
        return {};
    }
}
/**
 * 今週のみを抽出してメール送信＋ドキュメント作成
 */
function extractThisWeekAndSendEmailAndCreateDocs() {
    const config = getConfig();
    const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
    const thisWeekSheet = getSheetByWeek(spreadsheet, 0);
    if (!thisWeekSheet) {
        console.log('今週のシートが見つかりません');
        // エラーメールを送信
        try {
            if (config.EMAIL_ADDRESS) {
                GmailApp.sendEmail(config.EMAIL_ADDRESS, '今週の番組スケジュール - エラー', '今週のシートが見つかりませんでした。スプレッドシートの構成を確認してください。');
            }
        }
        catch (error) {
            console.error('エラーメール送信失敗:', error);
        }
        return {};
    }
    try {
        console.log('Processing this week:', thisWeekSheet.getName());
        const weekData = extractStructuredWeekData(thisWeekSheet);
        if (!weekData || typeof weekData !== 'object') {
            console.error('週データの抽出に失敗しました');
            return {};
        }
        const results = { '今週': weekData };
        logStructuredResults(results);
        sendProgramEmail(results, '今週の番組スケジュール');
        // Googleドキュメントも作成
        createProgramDocuments(results, '今週');
        return results;
    }
    catch (error) {
        console.error('データ抽出エラー:', error);
        // エラーメールを送信
        try {
            if (config.EMAIL_ADDRESS) {
                GmailApp.sendEmail(config.EMAIL_ADDRESS, '今週の番組スケジュール - エラー', `データ抽出中にエラーが発生しました。\nエラー: ${error.toString()}`);
            }
        }
        catch (emailError) {
            console.error('エラーメール送信失敗:', emailError);
        }
        return {};
    }
}
/**
 * 来週のみを抽出（番組別表示）してメール送信＋ドキュメント作成
 */
function extractNextWeekByProgramAndSendEmailAndCreateDocs() {
    const config = getConfig();
    const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
    const nextWeekSheet = getSheetByWeek(spreadsheet, 1);
    if (!nextWeekSheet) {
        console.log('来週のシートが見つかりません');
        // エラーメールを送信
        try {
            if (config.EMAIL_ADDRESS) {
                GmailApp.sendEmail(config.EMAIL_ADDRESS, '来週の番組スケジュール（番組別） - エラー', '来週のシートが見つかりませんでした。スプレッドシートの構成を確認してください。');
            }
        }
        catch (error) {
            console.error('エラーメール送信失敗:', error);
        }
        return {};
    }
    try {
        console.log('Processing next week:', nextWeekSheet.getName());
        const weekData = extractStructuredWeekData(nextWeekSheet);
        if (!weekData || typeof weekData !== 'object') {
            console.error('週データの抽出に失敗しました');
            return {};
        }
        const results = { '来週': weekData };
        logResultsByProgram(results);
        sendProgramEmail(results, '来週の番組スケジュール（番組別）');
        // Googleドキュメントも作成
        createProgramDocuments(results, '来週');
        return results;
    }
    catch (error) {
        console.error('データ抽出エラー:', error);
        // エラーメールを送信
        try {
            if (config.EMAIL_ADDRESS) {
                GmailApp.sendEmail(config.EMAIL_ADDRESS, '来週の番組スケジュール（番組別） - エラー', `データ抽出中にエラーが発生しました。\nエラー: ${error.toString()}`);
            }
        }
        catch (emailError) {
            console.error('エラーメール送信失敗:', emailError);
        }
        return {};
    }
}
/**
 * 番組ごとのGoogleドキュメントを作成
 */
function createProgramDocuments(allResults, weekLabel) {
    const config = getConfig();
    if (!config.DOCUMENT_TEMPLATES || !config.DOCUMENT_OUTPUT_FOLDER_ID) {
        console.error('ドキュメントテンプレート設定が不完全です。config.gsを確認してください。');
        return;
    }
    try {
        const outputFolder = DriveApp.getFolderById(config.DOCUMENT_OUTPUT_FOLDER_ID);
        const createdDocuments = [];
        Object.keys(allResults).forEach(weekName => {
            const weekResults = allResults[weekName];
            if (weekResults && typeof weekResults === 'object') {
                // 週名からシート名を取得し、月曜日の日付を計算
                let mondayDate = null;
                if (weekName.match(/^\d{2}\.\d{1,2}\.\d{2}-/)) {
                    // weekNameがシート名形式の場合
                    mondayDate = getStartDateFromSheetName(weekName);
                }
                else {
                    // weekNameが「今週」「来週」などの場合、現在時刻から計算
                    const today = new Date();
                    let offset = 0;
                    if (weekName === '来週')
                        offset = 1;
                    else if (weekName === '先週')
                        offset = -1;
                    const targetDate = new Date(today.getTime() + (offset * 7 * 24 * 60 * 60 * 1000));
                    const dayOfWeek = targetDate.getDay() === 0 ? 7 : targetDate.getDay();
                    mondayDate = new Date(targetDate.getTime() - (dayOfWeek - 1) * 24 * 60 * 60 * 1000);
                }
                Object.keys(weekResults).forEach(programName => {
                    if (config.DOCUMENT_TEMPLATES[programName]) {
                        const programData = weekResults[programName];
                        try {
                            const docId = createSingleProgramDocument(programName, programData, config.DOCUMENT_TEMPLATES[programName], outputFolder, weekLabel, mondayDate);
                            if (docId) {
                                createdDocuments.push({
                                    program: programName,
                                    docId: docId,
                                    url: `https://docs.google.com/document/d/${docId}/edit`
                                });
                                console.log(`${programName}のドキュメントを作成しました: ${docId}`);
                            }
                        }
                        catch (error) {
                            console.error(`${programName}のドキュメント作成エラー:`, error);
                        }
                    }
                    else {
                        console.log(`${programName}のテンプレートが設定されていません`);
                    }
                });
            }
        });
        // 作成されたドキュメントの情報をメールで通知
        if (createdDocuments.length > 0 && config.EMAIL_ADDRESS) {
            const docListText = createdDocuments.map(doc => `${doc.program}: ${doc.url}`).join('\n');
            GmailApp.sendEmail(config.EMAIL_ADDRESS, `${weekLabel}の番組ドキュメント作成完了`, `以下のドキュメントが作成されました：\n\n${docListText}`);
        }
    }
    catch (error) {
        console.error('ドキュメント作成処理エラー:', error);
        if (config.EMAIL_ADDRESS) {
            GmailApp.sendEmail(config.EMAIL_ADDRESS, 'ドキュメント作成エラー', `ドキュメント作成中にエラーが発生しました。\nエラー: ${error.toString()}`);
        }
    }
}
/**
 * 単一番組のドキュメントを作成
 */
function createSingleProgramDocument(programName, programData, templateId, outputFolder, weekLabel, mondayDate) {
    try {
        // テンプレートをコピー
        const templateFile = DriveApp.getFileById(templateId);
        // 【修正】月曜日の日付をyyyymmdd形式でフォーマット
        let dateStr = '';
        if (mondayDate && mondayDate instanceof Date) {
            dateStr = `${mondayDate.getFullYear()}${(mondayDate.getMonth() + 1).toString().padStart(2, '0')}${mondayDate.getDate().toString().padStart(2, '0')}`;
        }
        else {
            // フォールバック: 現在日付を使用
            const now = new Date();
            dateStr = `${now.getFullYear()}${(now.getMonth() + 1).toString().padStart(2, '0')}${now.getDate().toString().padStart(2, '0')}`;
        }
        // 【修正】番組名からスペースを削除
        const cleanProgramName = programName.replace(/\s+/g, '');
        // 【修正】番組の放送日数に応じてタイトル形式を決定
        let docName = '';
        if (programName === 'ちょうどいいラジオ' || programName === 'PRIME TIME') {
            // 月～木の放送がある番組：【連絡票】タイトル_yyyymmdd週
            docName = `【連絡票】${cleanProgramName}_${dateStr}週`;
        }
        else {
            // 週1日しか放送がない番組：【連絡票】タイトル_yyyymmdd
            docName = `【連絡票】${cleanProgramName}_${dateStr}`;
        }
        const copiedFile = templateFile.makeCopy(docName, outputFolder);
        const copiedDoc = DocumentApp.openById(copiedFile.getId());
        const body = copiedDoc.getBody();
        // プレースホルダーを置換
        replacePlaceholders(body, programName, programData);
        // ドキュメントを保存して閉じる
        copiedDoc.saveAndClose();
        console.log(`ドキュメント作成: ${docName} (月曜日: ${mondayDate ? mondayDate.toDateString() : '不明'})`);
        return copiedFile.getId();
    }
    catch (error) {
        console.error(`${programName}のドキュメント作成エラー:`, error);
        return null;
    }
}
/**
 * プレースホルダーを実際のデータに置換
 */
function replacePlaceholders(body, programName, programData) {
    // 番組名を置換
    body.replaceText('{{番組名}}', programName);
    body.replaceText('{{PROGRAM_NAME}}', programName);
    // 各曜日・カテゴリのデータを置換
    Object.keys(programData).forEach(day => {
        const dayData = programData[day];
        if (dayData && typeof dayData === 'object') {
            // 日付情報
            if (dayData['日付'] && Array.isArray(dayData['日付'])) {
                body.replaceText(`{{${day}_日付}}`, dayData['日付'][0] || 'ー');
                body.replaceText(`{{${day.toUpperCase()}_DATE}}`, dayData['日付'][0] || 'ー');
            }
            // 各カテゴリのデータ
            Object.keys(dayData).forEach(category => {
                const items = dayData[category];
                let content = '';
                if (Array.isArray(items)) {
                    if (category === '楽曲') {
                        // 楽曲の特別処理 - 複数フォーマット対応
                        content = formatMusicList(items); // デフォルト（番号付き）
                        // 楽曲の異なるフォーマット用プレースホルダー
                        const musicSimple = formatMusicListSimple(items); // 番号なし
                        const musicBullet = formatMusicListBullet(items); // 箇条書き
                        const musicTable = formatMusicListTable(items); // テーブル形式
                        const musicOneLine = formatMusicListOneLine(items); // 1行ずつ
                        // 楽曲の様々なフォーマットプレースホルダーを置換
                        body.replaceText(`{{${day}_楽曲_シンプル}}`, musicSimple);
                        body.replaceText(`{{${day}_楽曲_箇条書き}}`, musicBullet);
                        body.replaceText(`{{${day}_楽曲_テーブル}}`, musicTable);
                        body.replaceText(`{{${day}_楽曲_一行}}`, musicOneLine);
                        body.replaceText(`{{${day.toUpperCase()}_MUSIC_SIMPLE}}`, musicSimple);
                        body.replaceText(`{{${day.toUpperCase()}_MUSIC_BULLET}}`, musicBullet);
                        body.replaceText(`{{${day.toUpperCase()}_MUSIC_TABLE}}`, musicTable);
                        body.replaceText(`{{${day.toUpperCase()}_MUSIC_ONELINE}}`, musicOneLine);
                    }
                    else {
                        content = items.filter(item => item !== 'ー').join('\n');
                    }
                }
                else {
                    content = items !== 'ー' ? items.toString() : '';
                }
                // プレースホルダーを置換
                const placeholder1 = `{{${day}_${category}}}`;
                const placeholder2 = `{{${day.toUpperCase()}_${category.toUpperCase()}}}`;
                body.replaceText(placeholder1, content || 'ー');
                body.replaceText(placeholder2, content || 'ー');
                // 特殊なカテゴリ名の短縮版プレースホルダー
                const shortCategory = getShortCategoryName(category);
                if (shortCategory) {
                    body.replaceText(`{{${day}_${shortCategory}}}`, content || 'ー');
                    body.replaceText(`{{${day.toUpperCase()}_${shortCategory.toUpperCase()}}}`, content || 'ー');
                }
                // 英語版プレースホルダー（カテゴリ名を英語に変換）
                const englishCategory = convertCategoryToEnglish(category);
                if (englishCategory) {
                    const englishPlaceholder = `{{${day.toUpperCase()}_${englishCategory}}}`;
                    body.replaceText(englishPlaceholder, content || 'ー');
                }
            });
        }
    });
    // 収録予定など、曜日に依存しないデータを処理
    Object.keys(programData).forEach(key => {
        if (key.includes('収録予定')) {
            const scheduleData = programData[key];
            let scheduleContent = '';
            if (Array.isArray(scheduleData)) {
                scheduleContent = scheduleData.filter(item => item !== 'ー').join('\n');
            }
            else {
                scheduleContent = scheduleData !== 'ー' ? scheduleData.toString() : '';
            }
            // 収録予定用プレースホルダー
            body.replaceText(`{{${key}}}`, scheduleContent || 'ー');
            // 短縮版プレースホルダー
            const shortScheduleName = getShortScheduleName(key);
            if (shortScheduleName) {
                body.replaceText(`{{${shortScheduleName}}}`, scheduleContent || 'ー');
            }
        }
    });
    // 生成日時
    const now = new Date();
    const generatedTime = `${now.getFullYear()}/${(now.getMonth() + 1).toString().padStart(2, '0')}/${now.getDate().toString().padStart(2, '0')} ${now.getHours().toString().padStart(2, '0')}:${now.getMinutes().toString().padStart(2, '0')}`;
    body.replaceText('{{生成日時}}', generatedTime);
    body.replaceText('{{GENERATED_TIME}}', generatedTime);
}
/**
 * カテゴリ名の短縮版を取得
 */
function getShortCategoryName(category) {
    const shortNames = {
        '7:28パブ告知': '728パブ',
        '時間指定なし告知': '告知',
        'YOKOHAMA PORTSIDE INFORMATION': 'ポートサイド',
        '先行予約': '予約',
        'ラジオショッピング': 'ラジショ',
        'はぴねすくらぶ': 'はぴねす',
        'ヨコアリくん': 'ヨコアリ',
        '放送後': '放送後',
        '19:41Traffic': '1941Traffic',
        '営業コーナー': '営業',
        '指定曲': '指定曲',
        '時間指定なしパブ': 'パブ',
        'ラジショピ': 'ラジショ',
        '先行予約・限定予約': '予約',
        '12:40 電話パブ': '1240パブ',
        '13:29 パブリシティ': '1329パブ',
        '13:40 パブリシティ': '1340パブ',
        '12:15 リポート案件': '1215リポート',
        '14:29 リポート案件': '1429リポート',
        '14:41パブ': '1441パブ',
        'リポート 16:47': '1647リポート',
        '営業パブ 17:41': '1741パブ'
    };
    return shortNames[category] || null;
}
/**
 * 収録予定名の短縮版を取得
 */
function getShortScheduleName(scheduleName) {
    const shortNames = {
        'ちょうどいい暮らし収録予定': '暮らし収録',
        'ここが知りたい不動産収録予定': '不動産収録',
        'ちょうどいい歯ッピー収録予定': '歯ッピー収録',
        'ちょうどいいおカネの話収録予定': 'おカネ収録',
        'ちょうどいいごりごり隊収録予定': 'ごりごり収録',
        'ビジネスアイ収録予定': 'ビジネスアイ収録'
    };
    return shortNames[scheduleName] || null;
}
/**
 * 楽曲リストをフォーマット
 */
function formatMusicList(musicItems) {
    if (!Array.isArray(musicItems) || musicItems.length === 0) {
        return 'ー';
    }
    return musicItems.map((item, index) => {
        if (typeof item === 'object' && item.曲名) {
            let formatted = `${index + 1}. ${item.曲名}`;
            if (item.URL) {
                formatted += `\n${item.URL}`;
            }
            return formatted;
        }
        else {
            return `${index + 1}. ${item}`;
        }
    }).join('\n\n');
}
/**
 * 楽曲リストをシンプルにフォーマット（番号なし）
 */
function formatMusicListSimple(musicItems) {
    if (!Array.isArray(musicItems) || musicItems.length === 0) {
        return 'ー';
    }
    return musicItems.map(item => {
        if (typeof item === 'object' && item.曲名) {
            let formatted = `${item.曲名}`;
            if (item.URL) {
                formatted += `\n${item.URL}`;
            }
            return formatted;
        }
        else {
            return item;
        }
    }).join('\n\n');
}
/**
 * 楽曲リストを箇条書きでフォーマット
 */
function formatMusicListBullet(musicItems) {
    if (!Array.isArray(musicItems) || musicItems.length === 0) {
        return 'ー';
    }
    return musicItems.map(item => {
        if (typeof item === 'object' && item.曲名) {
            let formatted = `• ${item.曲名}`;
            if (item.URL) {
                formatted += `\n  ${item.URL}`;
            }
            return formatted;
        }
        else {
            return `• ${item}`;
        }
    }).join('\n\n');
}
/**
 * 楽曲リストをテーブル形式でフォーマット
 */
function formatMusicListTable(musicItems) {
    if (!Array.isArray(musicItems) || musicItems.length === 0) {
        return 'ー';
    }
    let result = '楽曲リスト\n';
    result += '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n';
    musicItems.forEach((item, index) => {
        if (typeof item === 'object' && item.曲名) {
            result += `${index + 1}. ${item.曲名}\n`;
            if (item.URL) {
                result += `   ${item.URL}\n`;
            }
        }
        else {
            result += `${index + 1}. ${item}\n`;
        }
        if (index < musicItems.length - 1) {
            result += '─────────────────────────────────────────────────────────\n';
        }
    });
    result += '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━';
    return result;
}
/**
 * 楽曲リストを1行ずつフォーマット
 */
function formatMusicListOneLine(musicItems) {
    if (!Array.isArray(musicItems) || musicItems.length === 0) {
        return 'ー';
    }
    return musicItems.map((item, index) => {
        if (typeof item === 'object' && item.曲名) {
            if (item.URL) {
                return `${index + 1}. ${item.曲名}\n${item.URL}`;
            }
            else {
                return `${index + 1}. ${item.曲名}`;
            }
        }
        else {
            return `${index + 1}. ${item}`;
        }
    }).join('\n\n');
}
/**
 * カテゴリ名を英語に変換
 */
function convertCategoryToEnglish(category) {
    const categoryMap = {
        '楽曲': 'MUSIC',
        '日付': 'DATE',
        'ゲスト': 'GUEST',
        '先行予約': 'ADVANCE_BOOKING',
        'ラジオショッピング': 'RADIO_SHOPPING',
        'はぴねすくらぶ': 'HAPPINESS_CLUB',
        '時間指定なし告知': 'GENERAL_ANNOUNCEMENT',
        'ヨコアリくん': 'YOKOARI_KUN',
        '放送後': 'AFTER_BROADCAST',
        '営業コーナー': 'BUSINESS_CORNER',
        '指定曲': 'REQUEST_SONG',
        'ラジショピ': 'RADIO_SHOPPING',
        '先行予約・限定予約': 'ADVANCE_LIMITED_BOOKING'
    };
    return categoryMap[category] || null;
}
/**
 * 3週間分を抽出（週別表示）
 */
function extractRadioSchedule() {
    const config = getConfig();
    const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
    const lastWeekSheet = getSheetByWeek(spreadsheet, -1);
    const thisWeekSheet = getSheetByWeek(spreadsheet, 0);
    const nextWeekSheet = getSheetByWeek(spreadsheet, 1);
    const allResults = {};
    if (lastWeekSheet) {
        console.log('Processing last week:', lastWeekSheet.getName());
        allResults['先週'] = extractStructuredWeekData(lastWeekSheet);
    }
    if (thisWeekSheet) {
        console.log('Processing this week:', thisWeekSheet.getName());
        allResults['今週'] = extractStructuredWeekData(thisWeekSheet);
    }
    if (nextWeekSheet) {
        console.log('Processing next week:', nextWeekSheet.getName());
        allResults['来週'] = extractStructuredWeekData(nextWeekSheet);
    }
    console.log('All extraction results:', JSON.stringify(allResults, null, 2));
    logStructuredResults(allResults);
    return allResults;
}
/**
 * 3週間分を抽出（番組別表示）
 */
function extractRadioScheduleByProgram() {
    const config = getConfig();
    const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
    const lastWeekSheet = getSheetByWeek(spreadsheet, -1);
    const thisWeekSheet = getSheetByWeek(spreadsheet, 0);
    const nextWeekSheet = getSheetByWeek(spreadsheet, 1);
    const allResults = {};
    if (lastWeekSheet) {
        console.log('Processing last week:', lastWeekSheet.getName());
        allResults['先週'] = extractStructuredWeekData(lastWeekSheet);
    }
    if (thisWeekSheet) {
        console.log('Processing this week:', thisWeekSheet.getName());
        allResults['今週'] = extractStructuredWeekData(thisWeekSheet);
    }
    if (nextWeekSheet) {
        console.log('Processing next week:', nextWeekSheet.getName());
        allResults['来週'] = extractStructuredWeekData(nextWeekSheet);
    }
    console.log('All extraction results:', JSON.stringify(allResults, null, 2));
    logResultsByProgram(allResults);
    return allResults;
}
/**
 * 3週間分を抽出してメール送信
 */
function extractRadioScheduleAndSendEmail() {
    const config = getConfig();
    try {
        const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
        const lastWeekSheet = getSheetByWeek(spreadsheet, -1);
        const thisWeekSheet = getSheetByWeek(spreadsheet, 0);
        const nextWeekSheet = getSheetByWeek(spreadsheet, 1);
        const allResults = {};
        if (lastWeekSheet) {
            console.log('Processing last week:', lastWeekSheet.getName());
            const lastWeekData = extractStructuredWeekData(lastWeekSheet);
            if (lastWeekData && typeof lastWeekData === 'object') {
                allResults['先週'] = lastWeekData;
            }
        }
        if (thisWeekSheet) {
            console.log('Processing this week:', thisWeekSheet.getName());
            const thisWeekData = extractStructuredWeekData(thisWeekSheet);
            if (thisWeekData && typeof thisWeekData === 'object') {
                allResults['今週'] = thisWeekData;
            }
        }
        if (nextWeekSheet) {
            console.log('Processing next week:', nextWeekSheet.getName());
            const nextWeekData = extractStructuredWeekData(nextWeekSheet);
            if (nextWeekData && typeof nextWeekData === 'object') {
                allResults['来週'] = nextWeekData;
            }
        }
        if (Object.keys(allResults).length === 0) {
            console.log('抽出できたデータがありません');
            // エラーメールを送信
            if (config.EMAIL_ADDRESS) {
                GmailApp.sendEmail(config.EMAIL_ADDRESS, '3週間分の番組スケジュール - エラー', '抽出できたデータがありませんでした。スプレッドシートの構成を確認してください。');
            }
            return {};
        }
        console.log('All extraction results:', JSON.stringify(allResults, null, 2));
        logStructuredResults(allResults);
        sendProgramEmail(allResults, '3週間分の番組スケジュール');
        return allResults;
    }
    catch (error) {
        console.error('データ抽出エラー:', error);
        // エラーメールを送信
        try {
            if (config.EMAIL_ADDRESS) {
                GmailApp.sendEmail(config.EMAIL_ADDRESS, '3週間分の番組スケジュール - エラー', `データ抽出中にエラーが発生しました。\nエラー: ${error.toString()}`);
            }
        }
        catch (emailError) {
            console.error('エラーメール送信失敗:', emailError);
        }
        return {};
    }
}
/**
 * 利用可能なシート名を一覧表示
 */
function showAvailableSheets() {
    const config = getConfig();
    const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
    const allSheets = spreadsheet.getSheets();
    const weekSheets = [];
    allSheets.forEach(sheet => {
        const sheetName = sheet.getName();
        if (sheetName.match(/^\d{2}\.\d{1,2}\.\d{2}-/)) {
            weekSheets.push(sheetName);
        }
    });
    console.log('=== 利用可能な週のシート一覧 ===');
    weekSheets.sort().forEach((sheetName, index) => {
        console.log(`${index + 1}. ${sheetName}`);
    });
    console.log('\n=== 簡単な使い方 ===');
    console.log('extractWeekByNumber(1) - 1番目の週');
    console.log('extractWeekByNumber(2) - 2番目の週');
    console.log('extractLatestWeek() - 最新の週');
    console.log('extractOldestWeek() - 最古の週');
    return weekSheets;
}
/**
 * 番号で週を指定して抽出（番組別表示）
 */
function extractWeekByNumber(number) {
    const config = getConfig();
    const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
    const allSheets = spreadsheet.getSheets();
    const weekSheets = [];
    allSheets.forEach(sheet => {
        const sheetName = sheet.getName();
        if (sheetName.match(/^\d{2}\.\d{1,2}\.\d{2}-/)) {
            weekSheets.push(sheetName);
        }
    });
    weekSheets.sort();
    if (number < 1 || number > weekSheets.length) {
        console.log(`番号は1から${weekSheets.length}の間で指定してください`);
        console.log('利用可能な週:');
        weekSheets.forEach((sheetName, index) => {
            console.log(`${index + 1}. ${sheetName}`);
        });
        return {};
    }
    const targetSheet = weekSheets[number - 1];
    console.log(`${number}番目の週を抽出: ${targetSheet}`);
    return extractSpecificWeekByProgram(targetSheet);
}
/**
 * 番号で週を指定してメール送信＋ドキュメント作成
 */
function extractWeekByNumberAndSendEmailAndCreateDocs(number) {
    const config = getConfig();
    const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
    const allSheets = spreadsheet.getSheets();
    const weekSheets = [];
    allSheets.forEach(sheet => {
        const sheetName = sheet.getName();
        if (sheetName.match(/^\d{2}\.\d{1,2}\.\d{2}-/)) {
            weekSheets.push(sheetName);
        }
    });
    weekSheets.sort();
    if (number < 1 || number > weekSheets.length) {
        console.log(`番号は1から${weekSheets.length}の間で指定してください`);
        console.log('利用可能な週:');
        weekSheets.forEach((sheetName, index) => {
            console.log(`${index + 1}. ${sheetName}`);
        });
        return {};
    }
    const targetSheet = weekSheets[number - 1];
    console.log(`${number}番目の週を抽出: ${targetSheet}`);
    return extractSpecificWeekByProgramAndSendEmailAndCreateDocs(targetSheet);
}
/**
 * 最新の週を抽出（番組別表示）
 */
function extractLatestWeek() {
    const config = getConfig();
    const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
    const allSheets = spreadsheet.getSheets();
    const weekSheets = [];
    allSheets.forEach(sheet => {
        const sheetName = sheet.getName();
        if (sheetName.match(/^\d{2}\.\d{1,2}\.\d{2}-/)) {
            weekSheets.push(sheetName);
        }
    });
    if (weekSheets.length === 0) {
        console.log('週のシートが見つかりません');
        return {};
    }
    weekSheets.sort();
    const latestSheet = weekSheets[weekSheets.length - 1];
    console.log(`最新の週を抽出: ${latestSheet}`);
    return extractSpecificWeekByProgram(latestSheet);
}
/**
 * 最古の週を抽出（番組別表示）
 */
function extractOldestWeek() {
    const config = getConfig();
    const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
    const allSheets = spreadsheet.getSheets();
    const weekSheets = [];
    allSheets.forEach(sheet => {
        const sheetName = sheet.getName();
        if (sheetName.match(/^\d{2}\.\d{1,2}\.\d{2}-/)) {
            weekSheets.push(sheetName);
        }
    });
    if (weekSheets.length === 0) {
        console.log('週のシートが見つかりません');
        return {};
    }
    weekSheets.sort();
    const oldestSheet = weekSheets[0];
    console.log(`最古の週を抽出: ${oldestSheet}`);
    return extractSpecificWeekByProgram(oldestSheet);
}
/**
 * 日付から週のシート名を生成
 */
function generateSheetNameFromDate(date) {
    const targetDate = new Date(date);
    // 月曜日を起点とした週の開始日を計算
    const dayOfWeek = targetDate.getDay() === 0 ? 7 : targetDate.getDay();
    const monday = new Date(targetDate.getTime() - (dayOfWeek - 1) * 24 * 60 * 60 * 1000);
    const sunday = new Date(monday.getTime() + 6 * 24 * 60 * 60 * 1000);
    const mondayYear = monday.getFullYear().toString().slice(-2);
    const mondayMonth = (monday.getMonth() + 1).toString();
    const mondayDay = monday.getDate().toString().padStart(2, '0');
    const sundayMonth = (sunday.getMonth() + 1).toString();
    const sundayDay = sunday.getDate().toString().padStart(2, '0');
    return `${mondayYear}.${mondayMonth}.${mondayDay}-${sundayMonth}.${sundayDay}`;
}
/**
 * 日付を指定して週を抽出（番組別表示）
 */
function extractWeekByDate(year, month, day) {
    const targetDate = new Date(year, month - 1, day);
    const sheetName = generateSheetNameFromDate(targetDate);
    console.log(`${year}年${month}月${day}日の週を抽出: ${sheetName}`);
    return extractSpecificWeekByProgram(sheetName);
}
/**
 * 日付を指定して週を抽出してメール送信＋ドキュメント作成
 */
function extractWeekByDateAndSendEmailAndCreateDocs(year, month, day) {
    const targetDate = new Date(year, month - 1, day);
    const sheetName = generateSheetNameFromDate(targetDate);
    console.log(`${year}年${month}月${day}日の週を抽出: ${sheetName}`);
    return extractSpecificWeekByProgramAndSendEmailAndCreateDocs(sheetName);
}
/**
 * 相対的な週を指定して抽出（0=今週、1=来週、-1=先週）
 */
function extractRelativeWeek(weekOffset) {
    const config = getConfig();
    const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
    const targetSheet = getSheetByWeek(spreadsheet, weekOffset);
    if (!targetSheet) {
        const offsetLabel = weekOffset === 0 ? '今週' : weekOffset > 0 ? `${weekOffset}週後` : `${Math.abs(weekOffset)}週前`;
        console.log(`${offsetLabel}のシートが見つかりません`);
        return {};
    }
    const sheetName = targetSheet.getName();
    const offsetLabel = weekOffset === 0 ? '今週' : weekOffset > 0 ? `${weekOffset}週後` : `${Math.abs(weekOffset)}週前`;
    console.log(`${offsetLabel}を抽出: ${sheetName}`);
    return extractSpecificWeekByProgram(sheetName);
}
/**
 * 相対的な週を指定してメール送信＋ドキュメント作成（0=今週、1=来週、-1=先週）
 */
function extractRelativeWeekAndSendEmailAndCreateDocs(weekOffset) {
    const config = getConfig();
    const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
    const targetSheet = getSheetByWeek(spreadsheet, weekOffset);
    if (!targetSheet) {
        const offsetLabel = weekOffset === 0 ? '今週' : weekOffset > 0 ? `${weekOffset}週後` : `${Math.abs(weekOffset)}週前`;
        console.log(`${offsetLabel}のシートが見つかりません`);
        return {};
    }
    const sheetName = targetSheet.getName();
    const offsetLabel = weekOffset === 0 ? '今週' : weekOffset > 0 ? `${weekOffset}週後` : `${Math.abs(weekOffset)}週前`;
    console.log(`${offsetLabel}を抽出: ${sheetName}`);
    return extractSpecificWeekByProgramAndSendEmailAndCreateDocs(sheetName);
}
/**
 * 指定した週を抽出（番組別表示）
 */
function extractSpecificWeekByProgram(sheetName) {
    const config = getConfig();
    try {
        const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
        const sheet = spreadsheet.getSheetByName(sheetName);
        if (!sheet) {
            console.log(`指定されたシート「${sheetName}」が見つかりません`);
            const allSheets = spreadsheet.getSheets();
            const availableSheets = [];
            allSheets.forEach(s => {
                if (s.getName().match(/^\d{2}\.\d{1,2}\.\d{2}-/)) {
                    availableSheets.push(s.getName());
                }
            });
            console.log('利用可能なシート名:', availableSheets);
            return {};
        }
        console.log(`Processing specified week: ${sheetName}`);
        const weekData = extractStructuredWeekData(sheet);
        if (!weekData || typeof weekData !== 'object') {
            console.error('週データの抽出に失敗しました');
            return {};
        }
        const results = { [sheetName]: weekData };
        logResultsByProgram(results);
        return results;
    }
    catch (error) {
        console.error('データ抽出エラー:', error);
        return {};
    }
}
/**
 * 指定した週を抽出（番組別表示）してメール送信＋ドキュメント作成
 */
function extractSpecificWeekByProgramAndSendEmailAndCreateDocs(sheetName) {
    const config = getConfig();
    try {
        const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
        const sheet = spreadsheet.getSheetByName(sheetName);
        if (!sheet) {
            console.log(`指定されたシート「${sheetName}」が見つかりません`);
            const allSheets = spreadsheet.getSheets();
            const availableSheets = [];
            allSheets.forEach(s => {
                if (s.getName().match(/^\d{2}\.\d{1,2}\.\d{2}-/)) {
                    availableSheets.push(s.getName());
                }
            });
            console.log('利用可能なシート名:', availableSheets);
            // エラーメールを送信
            if (config.EMAIL_ADDRESS) {
                GmailApp.sendEmail(config.EMAIL_ADDRESS, `${sheetName} の番組スケジュール（番組別） - エラー`, `指定されたシート「${sheetName}」が見つかりませんでした。\n\n利用可能なシート名:\n${availableSheets.join('\n')}`);
            }
            return {};
        }
        console.log(`Processing specified week: ${sheetName}`);
        const weekData = extractStructuredWeekData(sheet);
        if (!weekData || typeof weekData !== 'object') {
            console.error('週データの抽出に失敗しました');
            return {};
        }
        const results = { [sheetName]: weekData };
        logResultsByProgram(results);
        sendProgramEmail(results, `${sheetName} の番組スケジュール（番組別）`);
        // Googleドキュメントも作成
        createProgramDocuments(results, sheetName);
        return results;
    }
    catch (error) {
        console.error('データ抽出エラー:', error);
        // エラーメールを送信
        try {
            if (config.EMAIL_ADDRESS) {
                GmailApp.sendEmail(config.EMAIL_ADDRESS, `${sheetName} の番組スケジュール（番組別） - エラー`, `データ抽出中にエラーが発生しました。\nエラー: ${error.toString()}`);
            }
        }
        catch (emailError) {
            console.error('エラーメール送信失敗:', emailError);
        }
        return {};
    }
}
/**
 * 番組別ドキュメント作成のみ実行
 */
function createDocumentsOnly(sheetName) {
    const config = getConfig();
    try {
        const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
        const sheet = spreadsheet.getSheetByName(sheetName);
        if (!sheet) {
            console.log(`指定されたシート「${sheetName}」が見つかりません`);
            return {};
        }
        console.log(`Processing specified week: ${sheetName}`);
        const weekData = extractStructuredWeekData(sheet);
        if (!weekData || typeof weekData !== 'object') {
            console.error('週データの抽出に失敗しました');
            return {};
        }
        const results = { [sheetName]: weekData };
        // ドキュメントのみ作成（メール送信なし）
        createProgramDocuments(results, sheetName);
        return results;
    }
    catch (error) {
        console.error('ドキュメント作成エラー:', error);
        return {};
    }
}
/**
 * 指定した週を抽出してメール送信
 */
function extractSpecificWeekAndSendEmail(sheetName) {
    const config = getConfig();
    try {
        const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
        const sheet = spreadsheet.getSheetByName(sheetName);
        if (!sheet) {
            console.log(`指定されたシート「${sheetName}」が見つかりません`);
            const allSheets = spreadsheet.getSheets();
            const availableSheets = [];
            allSheets.forEach(s => {
                if (s.getName().match(/^\d{2}\.\d{1,2}\.\d{2}-/)) {
                    availableSheets.push(s.getName());
                }
            });
            console.log('利用可能なシート名:', availableSheets);
            // エラーメールを送信
            if (config.EMAIL_ADDRESS) {
                GmailApp.sendEmail(config.EMAIL_ADDRESS, `${sheetName} の番組スケジュール - エラー`, `指定されたシート「${sheetName}」が見つかりませんでした。\n\n利用可能なシート名:\n${availableSheets.join('\n')}`);
            }
            return {};
        }
        console.log(`Processing specified week: ${sheetName}`);
        const weekData = extractStructuredWeekData(sheet);
        if (!weekData || typeof weekData !== 'object') {
            console.error('週データの抽出に失敗しました');
            return {};
        }
        const results = { [sheetName]: weekData };
        logStructuredResults(results);
        sendProgramEmail(results, `${sheetName} の番組スケジュール`);
        return results;
    }
    catch (error) {
        console.error('データ抽出エラー:', error);
        // エラーメールを送信
        try {
            if (config.EMAIL_ADDRESS) {
                GmailApp.sendEmail(config.EMAIL_ADDRESS, `${sheetName} の番組スケジュール - エラー`, `データ抽出中にエラーが発生しました。\nエラー: ${error.toString()}`);
            }
        }
        catch (emailError) {
            console.error('エラーメール送信失敗:', emailError);
        }
        return {};
    }
}
/**
 * 番組データをメール形式に整形して送信
 */
function sendProgramEmail(allResults, subject) {
    const config = getConfig();
    if (!config.EMAIL_ADDRESS) {
        console.error('メールアドレスが設定されていません。config.gsでEMAIL_ADDRESSを設定してください。');
        return;
    }
    // allResultsの妥当性をチェック
    if (!allResults || typeof allResults !== 'object') {
        console.error('送信するデータが無効です:', allResults);
        return;
    }
    try {
        // 番組ごとに整理された本文を作成
        const emailBody = formatProgramDataForEmail(allResults);
        if (!emailBody || emailBody.trim() === '') {
            console.error('メール本文が空です');
            return;
        }
        // メール送信
        GmailApp.sendEmail(config.EMAIL_ADDRESS, subject, emailBody, {
            htmlBody: emailBody.replace(/\n/g, '<br>')
        });
        console.log(`メールを送信しました: ${config.EMAIL_ADDRESS}`);
        console.log(`件名: ${subject}`);
    }
    catch (error) {
        console.error('メール送信エラー:', error);
        console.error('エラーの詳細:', error instanceof Error ? error.toString() : String(error));
        console.error('スタックトレース:', error instanceof Error ? error.stack : 'No stack trace available');
    }
}
/**
 * 番組データをメール用テキストに整形
 */
function formatProgramDataForEmail(allResults) {
    const lines = [];
    // allResultsの妥当性をチェック
    if (!allResults || typeof allResults !== 'object') {
        console.error('formatProgramDataForEmail: 無効なデータが渡されました', allResults);
        return '番組データの取得に失敗しました。';
    }
    const resultKeys = Object.keys(allResults);
    if (resultKeys.length === 0) {
        console.warn('formatProgramDataForEmail: 番組データが空です');
        return '番組データが見つかりませんでした。';
    }
    lines.push('===============================');
    lines.push('    ラジオ番組スケジュール');
    lines.push('===============================');
    lines.push('');
    try {
        // 番組ごとにひとまとめにして表示
        const allPrograms = new Set();
        resultKeys.forEach(weekName => {
            const weekResults = allResults[weekName];
            if (weekResults && typeof weekResults === 'object') {
                Object.keys(weekResults).forEach(programName => {
                    allPrograms.add(programName);
                });
            }
        });
        if (allPrograms.size === 0) {
            return '対象番組が見つかりませんでした。';
        }
        allPrograms.forEach(programName => {
            lines.push(`◆◆◆ ${programName} ◆◆◆`);
            lines.push('='.repeat(60));
            lines.push('');
            resultKeys.forEach(weekName => {
                const weekResults = allResults[weekName];
                if (weekResults && typeof weekResults === 'object' && weekResults[programName]) {
                    lines.push(`--- ${weekName} ---`);
                    lines.push('');
                    const programData = weekResults[programName];
                    if (programData && typeof programData === 'object') {
                        Object.keys(programData).forEach(day => {
                            const dayData = programData[day];
                            if (dayData && typeof dayData === 'object') {
                                lines.push(`【${day}】`);
                                Object.keys(dayData).forEach(category => {
                                    const items = dayData[category];
                                    lines.push(`  ▼ ${category}:`);
                                    if (Array.isArray(items)) {
                                        if (items.length === 0) {
                                            lines.push(`    - データなし`);
                                        }
                                        else {
                                            items.forEach(item => {
                                                if (typeof item === 'object' && item !== null && item.曲名 !== undefined) {
                                                    // 楽曲オブジェクトの場合
                                                    lines.push(`    - 曲名: ${item.曲名}`);
                                                    if (item.URL) {
                                                        lines.push(`      URL: ${item.URL}`);
                                                    }
                                                }
                                                else {
                                                    lines.push(`    - ${item || 'ー'}`);
                                                }
                                            });
                                        }
                                    }
                                    else {
                                        lines.push(`    - ${items || 'ー'}`);
                                    }
                                });
                                lines.push(''); // 曜日間の空行
                            }
                        });
                    }
                    lines.push(''); // 週間の空行
                }
            });
            lines.push('');
            lines.push('='.repeat(60));
            lines.push('');
        });
    }
    catch (error) {
        console.error('formatProgramDataForEmail内でエラー:', error);
        lines.push('データの整形中にエラーが発生しました。');
        lines.push(`エラー: ${error.toString()}`);
    }
    // 生成日時を追加
    const now = new Date();
    lines.push('');
    lines.push(`生成日時: ${now.getFullYear()}/${(now.getMonth() + 1).toString().padStart(2, '0')}/${now.getDate().toString().padStart(2, '0')} ${now.getHours().toString().padStart(2, '0')}:${now.getMinutes().toString().padStart(2, '0')}`);
    return lines.join('\n');
}
/**
 * 週を指定してシートを取得
 */
function getSheetByWeek(spreadsheet, weekOffset) {
    const today = new Date();
    const targetDate = new Date(today.getTime() + (weekOffset * 7 * 24 * 60 * 60 * 1000));
    const dayOfWeek = targetDate.getDay() === 0 ? 7 : targetDate.getDay();
    const monday = new Date(targetDate.getTime() - (dayOfWeek - 1) * 24 * 60 * 60 * 1000);
    const sunday = new Date(monday.getTime() + 6 * 24 * 60 * 60 * 1000);
    const mondayYear = monday.getFullYear().toString().slice(-2);
    const mondayMonth = (monday.getMonth() + 1).toString();
    const mondayDay = monday.getDate().toString().padStart(2, '0');
    const sundayMonth = (sunday.getMonth() + 1).toString();
    const sundayDay = sunday.getDate().toString().padStart(2, '0');
    const sheetName = `${mondayYear}.${mondayMonth}.${mondayDay}-${sundayMonth}.${sundayDay}`;
    console.log(`Looking for sheet (offset ${weekOffset}): ${sheetName}`);
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
        console.warn(`Sheet not found: ${sheetName}`);
    }
    return sheet;
}
/**
 * 1つの週のデータを構造化して抽出
 */
function extractStructuredWeekData(sheet) {
    const markers = findMarkerRows(sheet);
    console.log('Markers found for', sheet.getName(), ':', markers);
    const dateRanges = getDateRanges(markers);
    console.log('Date ranges for', sheet.getName(), ':', dateRanges);
    const results = extractAndStructurePrograms(sheet, dateRanges, markers);
    return results;
}
/**
 * 区切り行（マーカー）を特定する
 */
function findMarkerRows(sheet) {
    const data = sheet.getDataRange().getValues();
    const rsRows = [];
    let newFridayRow = -1;
    let theBurnRow = -1;
    let mantenRow = -1;
    let chuuiRow = -1;
    let remarksCol = -1;
    for (let j = 0; j < data[0].length; j++) {
        if (data[0][j] && data[0][j].toString().includes('備考')) {
            remarksCol = j;
            break;
        }
    }
    for (let i = 0; i < data.length; i++) {
        if (remarksCol >= 0 && data[i][remarksCol] && data[i][remarksCol].toString().includes('RS')) {
            rsRows.push(i);
        }
        if (data[i].some(cell => cell && cell.toString().includes('New!Friday'))) {
            newFridayRow = i;
        }
        if (data[i].some(cell => cell && cell.toString().includes('THE BURN'))) {
            theBurnRow = i;
        }
        if (data[i].some(cell => cell && cell.toString().includes('まんてん'))) {
            mantenRow = i;
        }
        if (data[i].some(cell => cell && cell.toString().includes('注意：'))) {
            chuuiRow = i;
            break;
        }
    }
    return { rsRows, newFridayRow, theBurnRow, mantenRow, chuuiRow, remarksCol };
}
/**
 * 各曜日のデータ範囲を決定
 */
function getDateRanges(markers) {
    const { rsRows, newFridayRow, theBurnRow, mantenRow, chuuiRow } = markers;
    if (rsRows.length < 4) {
        throw new Error(`RS行が4つ見つからない。見つかった数: ${rsRows.length}`);
    }
    return {
        monday: { start: rsRows[0], end: rsRows[1] - 1 },
        tuesday: { start: rsRows[1], end: rsRows[2] - 1 },
        wednesday: { start: rsRows[2], end: rsRows[3] - 1 },
        thursday: { start: rsRows[3], end: rsRows[4] ? rsRows[4] - 2 : newFridayRow - 1 },
        friday: { start: newFridayRow + 1, end: theBurnRow - 1 },
        saturday: { start: theBurnRow + 1, end: mantenRow - 1 },
        sunday: { start: mantenRow + 1, end: chuuiRow - 1 }
    };
}
/**
 * シート名から開始日（月曜日）のDateオブジェクトを取得
 */
function getStartDateFromSheetName(sheetName) {
    const dateMatch = sheetName.match(/^(\d{2})\.(\d{1,2})\.(\d{2})-/);
    if (!dateMatch) {
        console.warn('シート名から日付を抽出できません:', sheetName);
        return new Date();
    }
    const year = parseInt('20' + dateMatch[1]);
    const month = parseInt(dateMatch[2]);
    const day = parseInt(dateMatch[3]);
    return new Date(year, month - 1, day);
}
/**
 * シート名から各曜日の日付を計算
 */
function calculateDayDates(sheetName) {
    const dateMatch = sheetName.match(/^(\d{2})\.(\d{1,2})\.(\d{2})-/);
    if (!dateMatch) {
        console.warn('シート名から日付を抽出できません:', sheetName);
        return {
            monday: '不明', tuesday: '不明', wednesday: '不明', thursday: '不明',
            friday: '不明', saturday: '不明', sunday: '不明'
        };
    }
    const year = parseInt('20' + dateMatch[1]);
    const month = parseInt(dateMatch[2]);
    const day = parseInt(dateMatch[3]);
    const mondayDate = new Date(year, month - 1, day);
    const dayDates = {
        monday: '', tuesday: '', wednesday: '', thursday: '',
        friday: '', saturday: '', sunday: ''
    };
    const dayNames = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday'];
    dayNames.forEach((dayName, index) => {
        const currentDate = new Date(mondayDate);
        currentDate.setDate(mondayDate.getDate() + index);
        const monthStr = (currentDate.getMonth() + 1).toString();
        const dayStr = currentDate.getDate().toString();
        dayDates[dayName] = `${monthStr}/${dayStr}`;
    });
    return dayDates;
}
/**
 * 対象番組の特定と構造化（修正版）
 */
function extractAndStructurePrograms(sheet, dateRanges, markers) {
    const data = sheet.getDataRange().getValues();
    const headerRow = data[0];
    const { rsRows, theBurnRow, remarksCol } = markers;
    const results = {};
    const dayDates = calculateDayDates(sheet.getName());
    const startDate = getStartDateFromSheetName(sheet.getName());
    // 楽曲データベースを1回だけ取得して使い回す
    console.log('楽曲データベースを読み込み中...');
    const musicDatabase = getMusicData();
    // 収録予定を事前に取得
    const recordingSchedules = extractRecordingSchedules(startDate);
    ['monday', 'tuesday', 'wednesday', 'thursday'].forEach(day => {
        headerRow.forEach((program, colIndex) => {
            if (program && program.toString().includes('ちょうどいいラジオ')) {
                const rawContent = extractColumnData(data, colIndex, dateRanges[day]);
                const remarksData = extractRemarksData(data, remarksCol, dateRanges[day]);
                if (!results['ちょうどいいラジオ'])
                    results['ちょうどいいラジオ'] = {};
                // 【修正】currentSheetパラメータを追加
                let dayStructure = structureChoudo(rawContent, dayDates[day], remarksData, musicDatabase, startDate, sheet // ← この行が足りなかった！
                );
                // 放送後情報を追加
                const broadcastAfterInfo = generateBroadcastAfterInfo(startDate, day, recordingSchedules);
                dayStructure['放送後'] = broadcastAfterInfo;
                results['ちょうどいいラジオ'][day] = dayStructure;
            }
            if (program && program.toString().includes('PRIME TIME')) {
                const rawContent = extractColumnData(data, colIndex, dateRanges[day]);
                if (!results['PRIME TIME'])
                    results['PRIME TIME'] = {};
                results['PRIME TIME'][day] = structurePrimeTime(rawContent, dayDates[day], musicDatabase);
            }
        });
    });
    if (results['ちょうどいいラジオ']) {
        results['ちょうどいいラジオ']['ちょうどいい暮らし収録予定'] = recordingSchedules['ちょうどいい暮らし'] || ['ー'];
        results['ちょうどいいラジオ']['ここが知りたい不動産収録予定'] = recordingSchedules['ここが知りたい不動産'] || ['ー'];
        results['ちょうどいいラジオ']['ちょうどいい歯ッピー収録予定'] = recordingSchedules['ちょうどいい歯っぴー'] || ['ー'];
        results['ちょうどいいラジオ']['ちょうどいいおカネの話収録予定'] = recordingSchedules['ちょうどいいおカネの話'] || ['ー'];
        results['ちょうどいいラジオ']['ちょうどいいごりごり隊収録予定'] = recordingSchedules['ちょうどいいごりごり隊'] || ['ー'];
        results['ちょうどいいラジオ']['ビジネスアイ収録予定'] = recordingSchedules['ビジネスアイ'] || ['ー'];
    }
    // 【修正】金曜日の処理："New!Friday"行からヘッダーを取得
    if (markers.newFridayRow >= 0) {
        const fridayHeaderRow = data[markers.newFridayRow];
        fridayHeaderRow.forEach((program, colIndex) => {
            if (program && program.toString().includes('FLAG')) {
                const rawContent = extractColumnData(data, colIndex, dateRanges.friday);
                if (!results['FLAG'])
                    results['FLAG'] = {};
                results['FLAG']['friday'] = structureFlag(rawContent, dayDates.friday, musicDatabase);
            }
        });
    }
    if (theBurnRow >= 0) {
        const saturdayHeaderRow = data[theBurnRow];
        saturdayHeaderRow.forEach((program, colIndex) => {
            if (program && program.toString().includes('God Bless Saturday')) {
                const rawContent = extractColumnData(data, colIndex, dateRanges.saturday);
                if (!results['God Bless Saturday'])
                    results['God Bless Saturday'] = {};
                results['God Bless Saturday']['saturday'] = structureGodBless(rawContent, dayDates.saturday, musicDatabase);
            }
            if (program && program.toString().includes('Route 847')) {
                const rawContent = extractColumnData(data, colIndex, dateRanges.saturday);
                if (!results['Route 847'])
                    results['Route 847'] = {};
                results['Route 847']['saturday'] = structureRoute847(rawContent, dayDates.saturday, musicDatabase);
            }
        });
    }
    return results;
}
/**
 * 指定された列と行範囲からデータを抽出
 */
function extractColumnData(data, colIndex, range) {
    const content = [];
    for (let row = range.start; row <= range.end; row++) {
        if (data[row] && data[row][colIndex]) {
            const cellValue = data[row][colIndex].toString().trim();
            if (cellValue !== '') {
                content.push(cellValue);
            }
        }
    }
    return content;
}
/**
 * 備考列からRS・HC情報を抽出
 */
function extractRemarksData(data, remarksCol, range) {
    const remarksData = {
        radioShopping: [],
        hapinessClub: []
    };
    if (remarksCol < 0) {
        return remarksData;
    }
    for (let row = range.start; row <= range.end; row++) {
        if (data[row] && data[row][remarksCol]) {
            const cellValue = data[row][remarksCol].toString().trim();
            if (cellValue.startsWith('RS:') || cellValue.startsWith('RS：')) {
                const content = cellValue.replace(/^RS[：:]/, '').trim();
                if (content) {
                    remarksData.radioShopping.push(content);
                }
            }
            if (cellValue.startsWith('HC:') || cellValue.startsWith('HC：')) {
                const content = cellValue.replace(/^HC[：:]/, '').trim();
                if (content) {
                    remarksData.hapinessClub.push(content);
                }
            }
        }
    }
    return remarksData;
}
/**
 * ヨコアリくん（火曜のみ）の判定
 */
function checkYokoAriKun(startDate) {
    const config = getConfig();
    try {
        const calendar = CalendarApp.getCalendarById(config.CALENDAR_ID);
        if (!calendar) {
            console.error('指定されたカレンダーが見つかりません:', config.CALENDAR_ID);
            return 'なし';
        }
        // 火曜日の日付を計算
        const monday = new Date(startDate);
        const tuesday = new Date(monday);
        tuesday.setDate(monday.getDate() + 1);
        // 火曜日の1日分のイベントを検索
        const nextDay = new Date(tuesday);
        nextDay.setDate(tuesday.getDate() + 1);
        console.log(`ヨコアリくん検索範囲（火曜日）: ${tuesday.toDateString()}`);
        const events = calendar.getEvents(tuesday, nextDay);
        // 「横浜アリーナスポットオンエア」が含まれるイベントを検索
        const yokoAriEvent = events.find(event => event.getTitle().includes('横浜アリーナスポットオンエア'));
        if (yokoAriEvent) {
            console.log(`ヨコアリくん: あり (${yokoAriEvent.getTitle()})`);
            return 'あり';
        }
        else {
            console.log('ヨコアリくん: なし');
            return 'なし';
        }
    }
    catch (error) {
        console.error('ヨコアリくん判定エラー:', error);
        return 'なし';
    }
}
/**
 * 放送後情報を生成（日付マッチング修正版）
 */
function generateBroadcastAfterInfo(startDate, dayName, schedules) {
    const dayMapping = {
        'monday': 1,
        'tuesday': 2,
        'wednesday': 3,
        'thursday': 4,
        'friday': 5,
        'saturday': 6,
        'sunday': 0
    };
    const targetDayIndex = dayMapping[dayName];
    if (targetDayIndex === undefined) {
        return ['ー'];
    }
    // 対象日の日付を計算
    const monday = new Date(startDate);
    const targetDate = new Date(monday);
    if (targetDayIndex === 0) { // 日曜日の場合
        targetDate.setDate(monday.getDate() + 6);
    }
    else {
        targetDate.setDate(monday.getDate() + (targetDayIndex - 1));
    }
    const targetDateStr = `${targetDate.getMonth() + 1}/${targetDate.getDate()}`;
    console.log(`放送後チェック対象日: ${dayName} (${targetDateStr})`);
    const broadcastAfterItems = [];
    // 各収録予定をチェック
    Object.keys(schedules).forEach(scheduleName => {
        const scheduleItems = schedules[scheduleName];
        if (Array.isArray(scheduleItems)) {
            scheduleItems.forEach(item => {
                if (item && isExactDateMatch(item, targetDateStr)) {
                    // 収録予定の番組名を抽出
                    let programName = scheduleName;
                    if (item.includes('ちょうどいい暮らし')) {
                        programName = 'ちょうどいい暮らし';
                    }
                    else if (item.includes('ここが知りたい不動産')) {
                        programName = 'ここが知りたい不動産';
                    }
                    else if (item.includes('ちょうどいい歯')) {
                        programName = 'ちょうどいい歯ッピー';
                    }
                    else if (item.includes('おカネの話')) {
                        programName = 'ちょうどいいおカネの話';
                    }
                    else if (item.includes('ごりごり隊')) {
                        programName = 'ちょうどいいごりごり隊';
                    }
                    else if (item.includes('ビジネスアイ')) {
                        programName = 'ビジネスアイ';
                    }
                    broadcastAfterItems.push(`【収録】${programName}`);
                    console.log(`放送後: ${targetDateStr} - 【収録】${programName}`);
                }
            });
        }
    });
    return broadcastAfterItems.length > 0 ? broadcastAfterItems : ['ー'];
}
/**
 * 日付の厳密なマッチングを行う関数
 */
function isExactDateMatch(text, targetDateStr) {
    if (!text || !targetDateStr) {
        return false;
    }
    // 正規表現で厳密な日付マッチングを行う
    // パターン: 先頭、スペース、または区切り文字の後に「月/日」があり、その後に区切り文字、スペース、または末尾がくる
    const escapedDate = targetDateStr.replace('/', '\\/');
    const datePattern = new RegExp(`(^|\\s|[^0-9])${escapedDate}(\\s|[^0-9]|$)`);
    const isMatch = datePattern.test(text);
    // デバッグ用ログ
    console.log(`日付マッチング: "${text}" vs "${targetDateStr}" = ${isMatch}`);
    return isMatch;
}
/**
 * 日付マッチングのテスト関数（デバッグ用）
 */
function testDateMatching() {
    console.log('=== 日付マッチングテスト ===');
    const testCases = [
        { text: '6/2 ちょうどいいおカネの話', target: '6/2', expected: true },
        { text: '6/24 ちょうどいいおカネの話', target: '6/2', expected: false },
        { text: '6/20 ちょうどいいおカネの話', target: '6/2', expected: false },
        { text: '2024/6/2 10:00 ちょうどいいおカネの話', target: '6/2', expected: true },
        { text: 'ちょうどいいおカネの話 6/2', target: '6/2', expected: true },
        { text: '収録: 6/2 ちょうどいいおカネの話', target: '6/2', expected: true },
        { text: '6/2ちょうどいいおカネの話', target: '6/2', expected: false }, // 区切り文字なし
        { text: '16/2 ちょうどいいおカネの話', target: '6/2', expected: false }, // 16/2は6/2とは別
    ];
    testCases.forEach(test => {
        const result = isExactDateMatch(test.text, test.target);
        const status = result === test.expected ? '✓' : '✗';
        console.log(`${status} "${test.text}" vs "${test.target}" = ${result} (期待値: ${test.expected})`);
    });
}
/**
 * 現在のシートのA列から先行予約情報を抽出（改良版）
 */
function getAdvanceBookingFromCurrentSheet(sheet) {
    try {
        const data = sheet.getDataRange().getValues();
        console.log(`現在のシート（A列のみ）から先行予約情報を抽出中: ${sheet.getName()}`);
        const bookingsByDay = {
            monday: [],
            tuesday: [],
            wednesday: [],
            thursday: [],
            friday: [],
            saturday: [],
            sunday: []
        };
        let currentDay = null;
        let foundBookings = 0;
        // A列のみをチェックして先行予約情報を探す
        data.forEach((row, rowIndex) => {
            const cellValue = row[0]; // A列の値のみ
            if (!cellValue) {
                return; // 空のセルはスキップ
            }
            const text = cellValue.toString().trim();
            if (!text) {
                return; // 空の文字列はスキップ
            }
            // 日付形式を判定（〇／〇（〇）のパターン）
            const datePattern = /^\s*[0-9０-９]+\s*[\/／]\s*[0-9０-９]+\s*[（\(]\s*[月火水木金土日]\s*[）\)]\s*$/;
            if (datePattern.test(text)) {
                // 日付行の場合、曜日を判定
                if (text.includes('月'))
                    currentDay = 'monday';
                else if (text.includes('火'))
                    currentDay = 'tuesday';
                else if (text.includes('水'))
                    currentDay = 'wednesday';
                else if (text.includes('木'))
                    currentDay = 'thursday';
                else if (text.includes('金'))
                    currentDay = 'friday';
                else if (text.includes('土'))
                    currentDay = 'saturday';
                else if (text.includes('日'))
                    currentDay = 'sunday';
                console.log(`先行予約: 日付行検出 - ${text} (${currentDay}) [行${rowIndex + 1}]`);
            }
            else {
                // 【簡素化】曜日が設定されていて、3文字以上のテキストなら全て採用
                const isBookingInfo = currentDay && text.length > 2;
                if (isBookingInfo && bookingsByDay[currentDay]) {
                    // 重複チェック
                    if (!bookingsByDay[currentDay].includes(text)) {
                        bookingsByDay[currentDay].push(text);
                        foundBookings++;
                        console.log(`先行予約: ${currentDay} - ${text} [行${rowIndex + 1}]`);
                    }
                }
            }
        });
        console.log(`先行予約情報抽出完了: 合計${foundBookings}件`);
        return bookingsByDay;
    }
    catch (error) {
        console.error('現在のシートからの先行予約情報抽出エラー:', error);
        return null;
    }
}
/**
 * ちょうどいいラジオの構造化（統一版）
 */
function structureChoudo(content, date, remarksData, musicDatabase, startDate, currentSheet) {
    const structure = {
        '日付': [date],
        '7:28パブ告知': [],
        '時間指定なし告知': [],
        'YOKOHAMA PORTSIDE INFORMATION': [],
        '楽曲': [],
        '先行予約': [],
        'ゲスト': [],
        'ラジオショッピング': remarksData ? remarksData.radioShopping : [],
        'はぴねすくらぶ': remarksData ? remarksData.hapinessClub : [],
        'ヨコアリくん': [],
        '放送後': []
    };
    // 先行予約情報を一時的に保存する配列
    const allAdvanceBookings = [];
    // PORTSIDE情報をカレンダーから取得
    let portsideFromCalendar = null;
    if (startDate) {
        portsideFromCalendar = getPortsideInformationFromCalendar(startDate);
    }
    content.forEach(item => {
        if (item.includes('♪')) {
            // 第1段階: ♪マーク付き楽曲データを収集
            structure['楽曲'].push(item);
        }
        else if (item.includes('指定曲')) {
            // 第2段階: 「指定曲」テキストを含む項目を収集
            structure['楽曲'].push(item);
        }
        else if (item.includes('７：２８') || item.includes('7:28')) {
            structure['7:28パブ告知'].push(item);
        }
        else if (item.includes('YOKOHAMA PORTSIDE')) {
            structure['YOKOHAMA PORTSIDE INFORMATION'].push(item);
        }
        else if (item.includes('先行予約')) {
            // 番組表からの先行予約情報を一時配列に追加
            allAdvanceBookings.push(item);
        }
        else if (item.includes('ゲスト')) {
            structure['ゲスト'].push(item);
        }
        else {
            structure['時間指定なし告知'].push(item);
        }
    });
    // カレンダーからのPORTSIDE情報を追加
    if (portsideFromCalendar) {
        const currentDay = getCurrentDayFromDate(date);
        console.log(`現在処理中の曜日: ${currentDay} (${date})`);
        if (currentDay && portsideFromCalendar[currentDay]) {
            const calendarPortsideInfo = portsideFromCalendar[currentDay];
            if (calendarPortsideInfo.length > 0) {
                console.log(`カレンダーからPORTSIDE情報を取得: ${calendarPortsideInfo.length}件`);
                calendarPortsideInfo.forEach(info => {
                    structure['YOKOHAMA PORTSIDE INFORMATION'].push(`YOKOHAMA PORTSIDE INFORMATION [${info}]`);
                });
            }
        }
    }
    // 楽曲処理
    if (structure['楽曲'].length > 0) {
        structure['楽曲'] = splitMusicData(structure['楽曲'], musicDatabase);
    }
    // 【修正】現在のシートから先行予約情報を取得して一時配列に追加
    if (currentSheet) {
        console.log(`${date}: currentSheetから先行予約情報を取得中...`);
        const advanceBookings = getAdvanceBookingFromCurrentSheet(currentSheet);
        if (advanceBookings) {
            const currentDay = getCurrentDayFromDate(date);
            console.log(`${date}: 判定された曜日 = ${currentDay}`);
            if (currentDay && advanceBookings[currentDay] && advanceBookings[currentDay].length > 0) {
                // A列からの先行予約情報を一時配列に追加
                allAdvanceBookings.push(...advanceBookings[currentDay]);
                console.log(`${date}: 現在のシートから先行予約情報を追加: ${advanceBookings[currentDay].length}件`);
                advanceBookings[currentDay].forEach(booking => {
                    console.log(`  - ${booking}`);
                });
            }
            else {
                console.log(`${date}: ${currentDay}の先行予約情報が見つかりませんでした`);
            }
        }
        else {
            console.log(`${date}: advanceBookingsの取得に失敗しました`);
        }
    }
    else {
        console.log(`${date}: currentSheetがnullです`);
    }
    // ヨコアリくん判定（火曜日のみ）
    const currentDay = getCurrentDayFromDate(date);
    if (currentDay === 'tuesday' && startDate) {
        const yokoAriStatus = checkYokoAriKun(startDate);
        structure['ヨコアリくん'] = [yokoAriStatus];
    }
    else {
        structure['ヨコアリくん'] = ['ー'];
    }
    // 空の項目を「ー」で埋める（ヨコアリくんと放送後は除く）
    Object.keys(structure).forEach(key => {
        if (key !== '日付' && key !== 'ヨコアリくん' && key !== '放送後' && key !== '先行予約' && structure[key].length === 0) {
            structure[key] = ['ー'];
        }
    });
    // 【修正】先行予約情報を最終的に一つのテキストに結合
    if (allAdvanceBookings.length > 0) {
        const combinedText = allAdvanceBookings.join('\n');
        structure['先行予約'] = [combinedText];
        console.log(`${date}: 先行予約情報を結合: "${combinedText}"`);
    }
    else {
        structure['先行予約'] = ['ー'];
    }
    return structure;
}
/**
 * 日付文字列から曜日名を取得（monday, tuesday, etc.）
 */
function getCurrentDayFromDate(dateString) {
    try {
        // dateStringは "5/31" のような形式
        const currentYear = new Date().getFullYear();
        const [month, day] = dateString.split('/');
        const date = new Date(currentYear, parseInt(month) - 1, parseInt(day));
        const dayNames = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
        return dayNames[date.getDay()];
    }
    catch (error) {
        console.error('日付文字列の解析エラー:', error);
        return null;
    }
}
/**
 * YOKOHAMA PORTSIDE INFORMATION をカレンダーから取得
 */
function getPortsideInformationFromCalendar(startDate) {
    const config = getConfig();
    const portsideInfo = {
        monday: [],
        tuesday: [],
        wednesday: [],
        thursday: [],
        friday: [],
        saturday: [],
        sunday: []
    };
    if (!config.PORTSIDE_CALENDAR_ID) {
        console.warn('PORTSIDE_CALENDAR_IDが設定されていません。番組表からの情報を使用します。');
        return null;
    }
    try {
        const calendar = CalendarApp.getCalendarById(config.PORTSIDE_CALENDAR_ID);
        if (!calendar) {
            console.error('PORTSIDE専用カレンダーが見つかりません:', config.PORTSIDE_CALENDAR_ID);
            return null;
        }
        // 起点の月曜日から1週間分の期間を設定
        const monday = new Date(startDate);
        const endDate = new Date(monday);
        endDate.setDate(monday.getDate() + 7); // 月曜から次の月曜まで
        console.log(`PORTSIDE情報検索範囲: ${monday.toDateString()} - ${endDate.toDateString()}`);
        // 1週間分のイベントを取得
        const events = calendar.getEvents(monday, endDate);
        if (events.length === 0) {
            console.log('PORTSIDE専用カレンダーにイベントが見つかりませんでした');
            return portsideInfo;
        }
        console.log(`PORTSIDE専用カレンダーから${events.length}件のイベントを取得`);
        // 各イベントを曜日別に分類
        events.forEach(event => {
            const eventDate = event.getStartTime();
            const eventTitle = event.getTitle().trim();
            if (!eventTitle) {
                return; // タイトルが空の場合はスキップ
            }
            // 曜日を計算（月曜日を起点とした週）
            const daysDiff = Math.floor((eventDate.getTime() - monday.getTime()) / (24 * 60 * 60 * 1000));
            if (daysDiff >= 0 && daysDiff < 7) {
                const dayNames = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday'];
                const dayName = dayNames[daysDiff];
                if (dayName && portsideInfo[dayName]) {
                    portsideInfo[dayName].push({ 曲名: eventTitle, URL: '' });
                    console.log(`${dayName}: ${eventTitle}`);
                }
            }
        });
        // 各曜日でデータがない場合は空配列のまま
        Object.keys(portsideInfo).forEach(day => {
            if (portsideInfo[day].length === 0) {
                console.log(`${day}: PORTSIDE情報なし`);
            }
        });
        return portsideInfo;
    }
    catch (error) {
        console.error('PORTSIDE専用カレンダーアクセスエラー:', error);
        return null;
    }
}
/**
 * Googleカレンダーから収録予定を抽出
 */
function extractRecordingSchedules(startDate) {
    const config = getConfig();
    const schedules = {};
    const searchKeywords = {
        'ちょうどいい暮らし': 'ちょうどいい暮らし',
        'ここが知りたい不動産': 'ここが知りたい不動産',
        'ちょうどいい歯っぴー': 'ちょうどいい歯っぴー',
        'ちょうどいいおカネの話': 'ちょうどいいおカネの話',
        'ちょうどいいごりごり隊': 'ちょうどいいごりごり隊',
        'ビジネスアイ': 'ビジネスアイ'
    };
    try {
        const calendar = CalendarApp.getCalendarById(config.CALENDAR_ID);
        if (!calendar) {
            console.error('指定されたカレンダーが見つかりません:', config.CALENDAR_ID);
            Object.keys(searchKeywords).forEach(key => {
                schedules[searchKeywords[key]] = ['カレンダーアクセスエラー'];
            });
            return schedules;
        }
        const endDate = new Date(startDate);
        endDate.setDate(startDate.getDate() + 90);
        console.log(`カレンダー検索範囲: ${startDate.toDateString()} - ${endDate.toDateString()}`);
        Object.keys(searchKeywords).forEach(programKey => {
            const keyword = searchKeywords[programKey];
            console.log(`「${keyword}」の収録予定を検索中...`);
            const events = calendar.getEvents(startDate, endDate);
            const matchingEvents = events.filter(event => {
                return event.getTitle().includes(keyword);
            });
            matchingEvents.sort((a, b) => {
                const diffA = Math.abs(a.getStartTime().getTime() - startDate.getTime());
                const diffB = Math.abs(b.getStartTime().getTime() - startDate.getTime());
                return diffA - diffB;
            });
            const upcomingEvents = matchingEvents.slice(0, 2);
            if (upcomingEvents.length > 0) {
                schedules[keyword] = upcomingEvents.map(event => {
                    const eventDate = event.getStartTime();
                    const dateStr = `${eventDate.getMonth() + 1}/${eventDate.getDate()}`;
                    const timeStr = `${eventDate.getHours().toString().padStart(2, '0')}:${eventDate.getMinutes().toString().padStart(2, '0')}`;
                    return `${dateStr} ${timeStr} ${event.getTitle()}`;
                });
                console.log(`「${keyword}」: ${schedules[keyword].length}件の予定を取得`);
            }
            else {
                schedules[keyword] = ['予定なし'];
                console.log(`「${keyword}」: 予定が見つかりませんでした`);
            }
        });
    }
    catch (error) {
        console.error('カレンダーアクセスエラー:', error);
        Object.keys(searchKeywords).forEach(key => {
            schedules[searchKeywords[key]] = [`エラー: ${error instanceof Error ? error.message : String(error)}`];
        });
    }
    return schedules;
}
/**
 * 楽曲スプレッドシートから楽曲データを取得
 */
function getMusicData() {
    const config = getConfig();
    try {
        const musicSpreadsheet = SpreadsheetApp.openById(config.MUSIC_SPREADSHEET_ID);
        const musicSheet = musicSpreadsheet.getSheetByName(config.MUSIC_SHEET_NAME);
        if (!musicSheet) {
            console.error(`楽曲シート「${config.MUSIC_SHEET_NAME}」が見つかりません`);
            return [];
        }
        const data = musicSheet.getDataRange().getValues();
        if (data.length === 0) {
            console.error('楽曲データが空です');
            return [];
        }
        const headers = data[0];
        const songTitleColIndex = findColumnIndex(headers, ['曲名', 'タイトル', 'title', 'song']);
        const artistColIndex = findColumnIndex(headers, ['アーティスト', 'artist', 'singer']);
        const urlColIndex = findColumnIndex(headers, ['URL', 'url', '音源', '音源データ', 'link']);
        const metadataColIndex = findColumnIndex(headers, ['付帯情報', 'metadata', 'meta']);
        if (songTitleColIndex === -1 || artistColIndex === -1) {
            console.error('必要な列（曲名、アーティスト）が見つかりません');
            console.error('見つかった列:', headers);
            return [];
        }
        console.log(`楽曲データベース: ${data.length - 1}件の楽曲を読み込み`);
        console.log(`曲名列: ${headers[songTitleColIndex]}, アーティスト列: ${headers[artistColIndex]}, URL列: ${urlColIndex >= 0 ? headers[urlColIndex] : 'なし'}, 付帯情報列: ${metadataColIndex >= 0 ? headers[metadataColIndex] : 'なし'}`);
        return data.slice(1).map(row => ({
            title: row[songTitleColIndex] ? row[songTitleColIndex].toString().trim() : '',
            artist: row[artistColIndex] ? row[artistColIndex].toString().trim() : '',
            url: urlColIndex >= 0 && row[urlColIndex] ? row[urlColIndex].toString().trim() : '',
            metadata: metadataColIndex >= 0 && row[metadataColIndex] ? row[metadataColIndex].toString().trim() : ''
        })).filter(song => song.title || song.artist);
    }
    catch (error) {
        console.error('楽曲スプレッドシートアクセスエラー:', error);
        return [];
    }
}
/**
 * 列インデックスを検索（複数の候補から）
 */
function findColumnIndex(headers, candidates) {
    for (let candidate of candidates) {
        const index = headers.findIndex(header => header && header.toString().toLowerCase().includes(candidate.toLowerCase()));
        if (index >= 0)
            return index;
    }
    return -1;
}
/**
 * ソーススプレッドシートから楽曲付帯情報を抽出
 */
function extractMusicMetadata() {
    const SOURCE_SHEET_ID = '1r6GLEvsZiqb3vkXZmrZ7XCpvW4RcecC7Bfp9YzhnQUE';
    try {
        const sourceSpreadsheet = SpreadsheetApp.openById(SOURCE_SHEET_ID);
        const sourceSheet = sourceSpreadsheet.getSheets()[0]; // 最初のシート
        if (!sourceSheet) {
            console.error('ソーススプレッドシートが見つかりません');
            return [];
        }
        const data = sourceSheet.getDataRange().getValues();
        if (data.length === 0) {
            console.error('ソーススプレッドシートが空です');
            return [];
        }
        const musicMetadata = [];
        let consecutiveMusicLines = 0;
        let currentMusicBlock = [];
        // 2列目（インデックス1）をスキャン
        for (let i = 0; i < data.length; i++) {
            const cellValue = data[i][1] ? data[i][1].toString().trim() : '';
            if (cellValue.startsWith('♪')) {
                consecutiveMusicLines++;
                currentMusicBlock.push({ row: i, text: cellValue });
            }
            else {
                // ♪で始まらない行に遭遇
                if (consecutiveMusicLines >= 2) {
                    // 2行以上続いた♪のブロックを処理
                    currentMusicBlock.forEach(block => {
                        const parsed = parseMusicLine(block.text);
                        if (parsed) {
                            musicMetadata.push(parsed);
                        }
                    });
                }
                consecutiveMusicLines = 0;
                currentMusicBlock = [];
            }
        }
        // 最後のブロックも処理
        if (consecutiveMusicLines >= 2) {
            currentMusicBlock.forEach(block => {
                const parsed = parseMusicLine(block.text);
                if (parsed) {
                    musicMetadata.push(parsed);
                }
            });
        }
        console.log(`楽曲付帯情報: ${musicMetadata.length}件を抽出`);
        return musicMetadata;
    }
    catch (error) {
        console.error('楽曲付帯情報の抽出エラー:', error);
        return [];
    }
}
/**
 * 楽曲行をパースして楽曲名、アーティスト名、付帯情報を抽出
 */
function parseMusicLine(text) {
    if (!text || !text.startsWith('♪')) {
        return null;
    }
    // ♪を除去
    let content = text.substring(1).trim();
    // 付帯情報（※または＜で始まる）を抽出
    const metadataMatch = content.match(/(※[^※＜]*|＜[^※＜]*)/g);
    const metadata = metadataMatch ? metadataMatch.join(' ') : '';
    // 付帯情報を除去してアーティスト名と楽曲名を抽出
    content = content.replace(/(※[^※＜]*|＜[^※＜]*)/g, '').trim();
    // アーティスト名と楽曲名を分離（／で区切られていると仮定）
    const parts = content.split('／');
    if (parts.length >= 2) {
        return {
            title: parts[0].trim(),
            artist: parts[1].trim(),
            metadata: metadata
        };
    }
    else {
        // ／がない場合はすべてを楽曲名として扱う
        return {
            title: content,
            artist: '',
            metadata: metadata
        };
    }
}
/**
 * 楽曲付帯情報を既存の楽曲DBに統合
 */
function integrateMusicMetadata() {
    const config = getConfig();
    const MUSIC_SHEET_ID = config.MUSIC_SPREADSHEET_ID;
    try {
        // ソーススプレッドシートから付帯情報を抽出
        const metadataList = extractMusicMetadata();
        if (metadataList.length === 0) {
            console.log('統合する付帯情報がありません');
            return;
        }
        // 既存の楽曲DBを開く
        const musicSpreadsheet = SpreadsheetApp.openById(MUSIC_SHEET_ID);
        const musicSheet = musicSpreadsheet.getSheetByName(config.MUSIC_SHEET_NAME);
        if (!musicSheet) {
            console.error(`楽曲シート「${config.MUSIC_SHEET_NAME}」が見つかりません`);
            return;
        }
        const data = musicSheet.getDataRange().getValues();
        if (data.length === 0) {
            console.error('楽曲データが空です');
            return;
        }
        const headers = data[0];
        const songTitleColIndex = findColumnIndex(headers, ['曲名', 'タイトル', 'title', 'song']);
        const artistColIndex = findColumnIndex(headers, ['アーティスト', 'artist', 'singer']);
        let metadataColIndex = findColumnIndex(headers, ['付帯情報', 'metadata', 'meta']);
        // 付帯情報列が存在しない場合は追加
        if (metadataColIndex === -1) {
            metadataColIndex = headers.length;
            musicSheet.getRange(1, metadataColIndex + 1).setValue('付帯情報');
            console.log('付帯情報列を追加しました');
        }
        let updatedCount = 0;
        // 各行をチェックして付帯情報を統合
        for (let i = 1; i < data.length; i++) {
            const rowTitle = data[i][songTitleColIndex] ? data[i][songTitleColIndex].toString().trim() : '';
            const rowArtist = data[i][artistColIndex] ? data[i][artistColIndex].toString().trim() : '';
            // マッチする付帯情報を検索
            const matchingMetadata = metadataList.find(meta => {
                const titleMatch = meta.title && rowTitle &&
                    (meta.title.toLowerCase().includes(rowTitle.toLowerCase()) ||
                        rowTitle.toLowerCase().includes(meta.title.toLowerCase()));
                const artistMatch = meta.artist && rowArtist &&
                    (meta.artist.toLowerCase().includes(rowArtist.toLowerCase()) ||
                        rowArtist.toLowerCase().includes(meta.artist.toLowerCase()));
                return titleMatch || artistMatch;
            });
            if (matchingMetadata && matchingMetadata.metadata) {
                // 付帯情報を設定
                musicSheet.getRange(i + 1, metadataColIndex + 1).setValue(matchingMetadata.metadata);
                updatedCount++;
                console.log(`統合: ${rowTitle}／${rowArtist} <- ${matchingMetadata.metadata}`);
            }
        }
        console.log(`楽曲付帯情報の統合完了: ${updatedCount}件を更新`);
    }
    catch (error) {
        console.error('楽曲付帯情報統合エラー:', error);
    }
}
/**
 * 楽曲を検索してマッチしたデータを取得
 */
function searchMusicData(musicDatabase, searchText) {
    if (!musicDatabase || musicDatabase.length === 0) {
        return null;
    }
    const cleanedSearchText = searchText.toLowerCase().trim();
    if (!cleanedSearchText) {
        return null;
    }
    console.log(`楽曲検索: "${cleanedSearchText}"`);
    // 1. 番組表のテキストが楽曲DBの曲名に含まれているかチェック
    let match = musicDatabase.find(song => song.title && song.title.toLowerCase().includes(cleanedSearchText));
    if (match) {
        console.log(`曲名で検索ヒット: ${match.title} / ${match.artist}`);
        return match;
    }
    // 2. 番組表のテキストが楽曲DBのアーティスト名に含まれているかチェック
    match = musicDatabase.find(song => song.artist && song.artist.toLowerCase().includes(cleanedSearchText));
    if (match) {
        console.log(`アーティスト名で検索ヒット: ${match.title} / ${match.artist}`);
        return match;
    }
    // 3. 逆方向: 楽曲DBの曲名が番組表のテキストに含まれているかチェック
    match = musicDatabase.find(song => song.title && cleanedSearchText.includes(song.title.toLowerCase()));
    if (match) {
        console.log(`逆検索（曲名）でヒット: ${match.title} / ${match.artist}`);
        return match;
    }
    // 4. 逆方向: 楽曲DBのアーティスト名が番組表のテキストに含まれているかチェック
    match = musicDatabase.find(song => song.artist && cleanedSearchText.includes(song.artist.toLowerCase()));
    if (match) {
        console.log(`逆検索（アーティスト）でヒット: ${match.title} / ${match.artist}`);
        return match;
    }
    // 5. スペースで区切って部分検索
    const words = cleanedSearchText.split(/\s+/);
    if (words.length > 1) {
        match = musicDatabase.find(song => {
            const songText = `${song.title} ${song.artist}`.toLowerCase();
            return words.some(word => songText.includes(word) && word.length > 1); // 1文字は除外
        });
        if (match) {
            console.log(`部分検索でヒット: ${match.title} / ${match.artist}`);
            return match;
        }
    }
    console.log(`検索結果: 見つかりませんでした`);
    return null;
}
/**
 * 楽曲情報を分割・拡張する
 */
function splitMusicData(content, musicDatabase) {
    const musicList = [];
    let currentSong = '';
    // musicDatabaseがundefinedの場合はデフォルト値を返す
    if (!musicDatabase) {
        console.warn('楽曲データベースが見つかりません');
        return [{ 曲名: 'ー', URL: '' }];
    }
    content.forEach(item => {
        if (item.includes('♪')) {
            const parts = item.split('♪');
            if (parts[0] && currentSong) {
                currentSong += parts[0];
                musicList.push(processMusicText(currentSong.trim(), musicDatabase));
                currentSong = '';
            }
            for (let i = 1; i < parts.length; i++) {
                if (currentSong) {
                    musicList.push(processMusicText(currentSong.trim(), musicDatabase));
                }
                currentSong = parts[i];
            }
        }
        else {
            currentSong += item;
        }
    });
    if (currentSong) {
        musicList.push(processMusicText(currentSong.trim(), musicDatabase));
    }
    const filteredList = musicList.filter(song => song && (song.曲名 || song.URL));
    return filteredList.length > 0 ? filteredList : [{ 曲名: 'ー', URL: '' }];
}
/**
 * 楽曲テキストを処理して拡張情報を付加
 */
function processMusicText(text, musicDatabase) {
    if (!text)
        return { 曲名: '', URL: '', 付帯情報: '' };
    const cleanedText = cleanMusicText(text);
    const musicData = searchMusicData(musicDatabase, cleanedText);
    if (musicData && musicData.title && musicData.artist) {
        const songTitle = `♪${musicData.title}／${musicData.artist}`;
        const metadata = musicData.metadata || '';
        return {
            曲名: metadata ? `${songTitle} ${metadata}` : songTitle,
            URL: musicData.url || '',
            付帯情報: metadata
        };
    }
    else {
        return {
            曲名: cleanedText,
            URL: '',
            付帯情報: ''
        };
    }
}
/**
 * 楽曲付帯情報システムのテスト
 */
function testMusicMetadataSystem() {
    console.log('=== 楽曲付帯情報システムテスト ===');
    // 1. 楽曲付帯情報の抽出テスト
    console.log('\n--- 楽曲付帯情報抽出テスト ---');
    const metadataList = extractMusicMetadata();
    console.log(`抽出結果: ${metadataList.length}件`);
    metadataList.slice(0, 3).forEach((item, index) => {
        console.log(`${index + 1}. 楽曲: ${item.title}／${item.artist}`);
        console.log(`   付帯情報: ${item.metadata}`);
    });
    // 2. 楽曲データベースの統合テスト
    console.log('\n--- 楽曲DB統合テスト ---');
    integrateMusicMetadata();
    // 3. 更新された楽曲DBの読み込みテスト
    console.log('\n--- 更新された楽曲DB読み込みテスト ---');
    const musicDB = getMusicData();
    const metadataIncluded = musicDB.filter(song => song.metadata);
    console.log(`付帯情報付き楽曲: ${metadataIncluded.length}件`);
    metadataIncluded.slice(0, 3).forEach((item, index) => {
        console.log(`${index + 1}. ${item.title}／${item.artist}`);
        console.log(`   付帯情報: ${item.metadata}`);
    });
    console.log('\n=== テスト完了 ===');
}
/**
 * 楽曲テキストから数字と丸囲み数字を削除
 */
function cleanMusicText(text) {
    if (!text)
        return '';
    return text
        .replace(/[0-9０-９]/g, '')
        .replace(/[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]/g, '')
        .replace(/[⑴⑵⑶⑷⑸⑹⑺⑻⑼⑽⑾⑿⒀⒁⒂⒃⒄⒅⒆⒇]/g, '')
        .replace(/[㈠㈡㈢㈣㈤㈥㈦㈧㈨㈩]/g, '')
        .trim();
}
/**
 * PRIME TIMEの構造化（楽曲データ修正版）
 */
function structurePrimeTime(content, date, musicDatabase) {
    const structure = {
        '日付': [date],
        '19:41Traffic': [],
        '19:43': [],
        '20:51': [],
        '営業コーナー': [],
        '楽曲': [], // 「指定曲」から「楽曲」に変更
        'ゲスト': [],
        '時間指定なしパブ': [],
        'ラジショピ': [],
        '先行予約・限定予約': []
    };
    content.forEach(item => {
        if (item.includes('♪')) {
            // 第1段階: ♪マーク付き楽曲データを収集
            structure['楽曲'].push(item);
        }
        else if (item.includes('指定曲')) {
            // 第2段階: 「指定曲」テキストを含む項目を収集
            structure['楽曲'].push(item);
        }
        else if (item.includes('１９：４１') || item.includes('19:41')) {
            structure['19:41Traffic'].push(item);
        }
        else if (item.includes('１９：４３') || item.includes('19:43')) {
            structure['19:43'].push(item);
        }
        else if (item.includes('２０：５１') || item.includes('20:51')) {
            structure['20:51'].push(item);
        }
        else if (item.includes('営業')) {
            structure['営業コーナー'].push(item);
        }
        else if (item.includes('ゲスト')) {
            structure['ゲスト'].push(item);
        }
        else if (item.includes('ラジショ')) {
            structure['ラジショピ'].push(item);
        }
        else if (item.includes('先行予約') || item.includes('限定予約')) {
            structure['先行予約・限定予約'].push(item);
        }
        else {
            structure['時間指定なしパブ'].push(item);
        }
    });
    // 楽曲処理：musicDatabaseが渡されている場合は楽曲データを拡張
    if (structure['楽曲'].length > 0 && musicDatabase) {
        structure['楽曲'] = splitMusicData(structure['楽曲'], musicDatabase);
    }
    // 空の項目を「ー」で埋める
    Object.keys(structure).forEach(key => {
        if (key !== '日付' && structure[key].length === 0) {
            structure[key] = ['ー'];
        }
    });
    return structure;
}
/**
 * FLAGの構造化
 */
function structureFlag(content, date, musicDatabase) {
    const structure = {
        '日付': [date],
        '12:40 電話パブ': [],
        '13:29 パブリシティ': [],
        '13:40 パブリシティ': [],
        '12:15 リポート案件': [],
        '14:29 リポート案件': [],
        '時間指定なし告知': [],
        '楽曲': [],
        '先行予約': [],
        'ゲスト': []
    };
    content.forEach(item => {
        if (item.includes('♪')) {
            // 第1段階: ♪マーク付き楽曲データを収集
            structure['楽曲'].push(item);
        }
        else if (item.includes('指定曲')) {
            // 第2段階: 「指定曲」テキストを含む項目を収集
            structure['楽曲'].push(item);
        }
        else if (item.includes('１２：４０') || item.includes('12:40')) {
            structure['12:40 電話パブ'].push(item);
        }
        else if (item.includes('１３：２９') || item.includes('13:29')) {
            structure['13:29 パブリシティ'].push(item);
        }
        else if (item.includes('１３：４０') || item.includes('13:40')) {
            structure['13:40 パブリシティ'].push(item);
        }
        else if (item.includes('１２：１５') || item.includes('12:15')) {
            structure['12:15 リポート案件'].push(item);
        }
        else if (item.includes('１４：２９') || item.includes('14:29')) {
            structure['14:29 リポート案件'].push(item);
        }
        else if (item.includes('先行予約')) {
            structure['先行予約'].push(item);
        }
        else if (item.includes('ゲスト')) {
            structure['ゲスト'].push(item);
        }
        else {
            structure['時間指定なし告知'].push(item);
        }
    });
    if (structure['楽曲'].length > 0) {
        structure['楽曲'] = splitMusicData(structure['楽曲'], musicDatabase);
    }
    Object.keys(structure).forEach(key => {
        if (key !== '日付' && structure[key].length === 0) {
            structure[key] = ['ー'];
        }
    });
    return structure;
}
/**
 * God Bless Saturdayの構造化
 */
function structureGodBless(content, date, musicDatabase) {
    const structure = {
        '日付': [date],
        '楽曲': [],
        '14:41パブ': [],
        '時間指定なしパブ': []
    };
    content.forEach(item => {
        if (item.includes('♪')) {
            // 第1段階: ♪マーク付き楽曲データを収集
            structure['楽曲'].push(item);
        }
        else if (item.includes('指定曲')) {
            // 第2段階: 「指定曲」テキストを含む項目を収集
            structure['楽曲'].push(item);
        }
        else if (item.includes('１４：４１') || item.includes('14:41')) {
            structure['14:41パブ'].push(item);
        }
        else {
            structure['時間指定なしパブ'].push(item);
        }
    });
    if (structure['楽曲'].length > 0) {
        structure['楽曲'] = splitMusicData(structure['楽曲'], musicDatabase);
    }
    Object.keys(structure).forEach(key => {
        if (key !== '日付' && structure[key].length === 0) {
            structure[key] = ['ー'];
        }
    });
    return structure;
}
/**
 * Route 847の構造化
 */
function structureRoute847(content, date, musicDatabase) {
    const structure = {
        '日付': [date],
        'リポート 16:47': [],
        '営業パブ 17:41': [],
        '時間指定なし告知': [],
        '楽曲': []
    };
    content.forEach(item => {
        if (item.includes('♪')) {
            // 第1段階: ♪マーク付き楽曲データを収集
            structure['楽曲'].push(item);
        }
        else if (item.includes('指定曲')) {
            // 第2段階: 「指定曲」テキストを含む項目を収集
            structure['楽曲'].push(item);
        }
        else if (item.includes('１６：４７') || item.includes('16:47')) {
            structure['リポート 16:47'].push(item);
        }
        else if (item.includes('１７：４１') || item.includes('17:41')) {
            structure['営業パブ 17:41'].push(item);
        }
        else {
            structure['時間指定なし告知'].push(item);
        }
    });
    if (structure['楽曲'].length > 0) {
        structure['楽曲'] = splitMusicData(structure['楽曲'], musicDatabase);
    }
    Object.keys(structure).forEach(key => {
        if (key !== '日付' && structure[key].length === 0) {
            structure[key] = ['ー'];
        }
    });
    return structure;
}
/**
 * 構造化された結果を読みやすい形でログに出力（週別表示）
 */
function logStructuredResults(allResults) {
    console.log('\n===============================');
    console.log('    構造化された番組内容抽出結果');
    console.log('===============================\n');
    Object.keys(allResults).forEach(weekName => {
        const weekResults = allResults[weekName];
        console.log(`=== ${weekName} ===`);
        Object.keys(weekResults).forEach(programName => {
            const programData = weekResults[programName];
            console.log(`\n◆ ${programName}`);
            Object.keys(programData).forEach(day => {
                const dayData = programData[day];
                console.log(`\n  【${day}】`);
                Object.keys(dayData).forEach(category => {
                    const items = dayData[category];
                    console.log(`    ▼ ${category}:`);
                    if (Array.isArray(items)) {
                        items.forEach(item => {
                            if (typeof item === 'object' && item.曲名 !== undefined) {
                                // 楽曲オブジェクトの場合
                                console.log(`      - 曲名: ${item.曲名}`);
                                if (item.URL) {
                                    console.log(`        URL: ${item.URL}`);
                                }
                            }
                            else {
                                console.log(`      - ${item}`);
                            }
                        });
                    }
                    else {
                        console.log(`      - ${items}`);
                    }
                });
            });
        });
        console.log('\n' + '='.repeat(50) + '\n');
    });
}
/**
 * 番組ごとにひとまとめにしてログ表示
 */
function logResultsByProgram(allResults) {
    console.log('\n===============================');
    console.log('    番組別まとめ表示');
    console.log('===============================\n');
    const allPrograms = new Set();
    Object.keys(allResults).forEach(weekName => {
        const weekResults = allResults[weekName];
        Object.keys(weekResults).forEach(programName => {
            allPrograms.add(programName);
        });
    });
    allPrograms.forEach(programName => {
        console.log(`\n${'='.repeat(60)}`);
        console.log(`◆◆◆ ${programName} ◆◆◆`);
        console.log(`${'='.repeat(60)}\n`);
        Object.keys(allResults).forEach(weekName => {
            const weekResults = allResults[weekName];
            if (weekResults[programName]) {
                console.log(`--- ${weekName} ---\n`);
                const programData = weekResults[programName];
                Object.keys(programData).forEach(day => {
                    const dayData = programData[day];
                    console.log(`【${day}】`);
                    Object.keys(dayData).forEach(category => {
                        const items = dayData[category];
                        console.log(`  ▼ ${category}:`);
                        if (Array.isArray(items)) {
                            items.forEach(item => {
                                if (typeof item === 'object' && item.曲名 !== undefined) {
                                    // 楽曲オブジェクトの場合
                                    console.log(`    - 曲名: ${item.曲名}`);
                                    if (item.URL) {
                                        console.log(`      URL: ${item.URL}`);
                                    }
                                }
                                else {
                                    console.log(`    - ${item}`);
                                }
                            });
                        }
                        else {
                            console.log(`    - ${items}`);
                        }
                    });
                    console.log(''); // 曜日間の空行
                });
                console.log(''); // 週間の空行
            }
        });
    });
}
/**
 * 先行予約情報の詳細デバッグ用関数
 */
function debugAdvanceBooking(sheetName) {
    const config = getConfig();
    if (!sheetName) {
        // 最新の週のシートを使用
        const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
        const allSheets = spreadsheet.getSheets();
        const weekSheets = [];
        allSheets.forEach(sheet => {
            const name = sheet.getName();
            if (name.match(/^\d{2}\.\d{1,2}\.\d{2}-/)) {
                weekSheets.push(name);
            }
        });
        if (weekSheets.length === 0) {
            console.log('週のシートが見つかりません');
            return;
        }
        weekSheets.sort();
        sheetName = weekSheets[weekSheets.length - 1];
    }
    console.log(`=== ${sheetName} 詳細デバッグ ===`);
    try {
        const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
        const sheet = spreadsheet.getSheetByName(sheetName);
        if (!sheet) {
            console.log(`シート「${sheetName}」が見つかりません`);
            return;
        }
        const data = sheet.getDataRange().getValues();
        console.log(`シートの全行数: ${data.length}`);
        // A列の全てのデータを表示
        console.log('\n=== A列の全データ ===');
        data.forEach((row, rowIndex) => {
            const cellValue = row[0];
            if (cellValue && cellValue.toString().trim()) {
                const text = cellValue.toString().trim();
                console.log(`行${rowIndex + 1}: "${text}"`);
                // 日付パターンかどうか判定
                const datePattern = /^\s*[0-9０-９]+\s*[\/／]\s*[0-9０-９]+\s*[（\(]\s*[月火水木金土日]\s*[）\)]\s*$/;
                if (datePattern.test(text)) {
                    console.log(`  → 日付行として判定`);
                }
                else if (text.length > 5) {
                    console.log(`  → 候補として判定（${text.length}文字）`);
                }
            }
        });
        // 実際の抽出結果
        console.log('\n=== 抽出結果 ===');
        const bookings = getAdvanceBookingFromCurrentSheet(sheet);
        if (bookings) {
            const dayLabels = {
                monday: '月曜日',
                tuesday: '火曜日',
                wednesday: '水曜日',
                thursday: '木曜日',
                friday: '金曜日',
                saturday: '土曜日',
                sunday: '日曜日'
            };
            Object.keys(bookings).forEach(day => {
                console.log(`\n${dayLabels[day]}:`);
                if (bookings[day].length > 0) {
                    bookings[day].forEach(booking => {
                        console.log(`  - ${booking}`);
                    });
                }
                else {
                    console.log('  - なし');
                }
            });
        }
        else {
            console.log('先行予約情報の取得に失敗しました');
        }
        return bookings;
    }
    catch (error) {
        console.error('デバッグ実行エラー:', error);
    }
}
/**
 * 先行予約情報の詳細デバッグ用関数（改良版）
 */
function debugAdvanceBookingDetailed(sheetName) {
    const config = getConfig();
    if (!sheetName) {
        const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
        const thisWeekSheet = getSheetByWeek(spreadsheet, 0);
        if (thisWeekSheet) {
            sheetName = thisWeekSheet.getName();
        }
        else {
            console.log('今週のシートが見つかりません');
            return;
        }
    }
    console.log(`=== ${sheetName} 詳細デバッグ（A列全データ）===`);
    try {
        const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
        const sheet = spreadsheet.getSheetByName(sheetName);
        if (!sheet) {
            console.log(`シート「${sheetName}」が見つかりません`);
            return;
        }
        const data = sheet.getDataRange().getValues();
        console.log(`シートの全行数: ${data.length}`);
        let currentDay = null;
        // A列の全てのデータを表示
        console.log('\n=== A列の全データ（詳細判定）===');
        data.forEach((row, rowIndex) => {
            const cellValue = row[0];
            if (cellValue && cellValue.toString().trim()) {
                const text = cellValue.toString().trim();
                // 日付パターンかどうか判定
                const datePattern = /^\s*[0-9０-９]+\s*[\/／]\s*[0-9０-９]+\s*[（\(]\s*[月火水木金土日]\s*[）\)]\s*$/;
                if (datePattern.test(text)) {
                    // 日付行の場合
                    if (text.includes('月'))
                        currentDay = 'monday';
                    else if (text.includes('火'))
                        currentDay = 'tuesday';
                    else if (text.includes('水'))
                        currentDay = 'wednesday';
                    else if (text.includes('木'))
                        currentDay = 'thursday';
                    else if (text.includes('金'))
                        currentDay = 'friday';
                    else if (text.includes('土'))
                        currentDay = 'saturday';
                    else if (text.includes('日'))
                        currentDay = 'sunday';
                    console.log(`行${rowIndex + 1}: "${text}" → 【日付行】現在の曜日: ${currentDay}`);
                }
                else {
                    // データ行の場合
                    console.log(`行${rowIndex + 1}: "${text}" (${text.length}文字) → 現在の曜日: ${currentDay || 'なし'}`);
                    // 現在の判定条件をチェック
                    const hasKeyword = (text.includes('先行予約') ||
                        text.includes('予約') ||
                        text.includes('チケット') ||
                        text.includes('早期') ||
                        text.includes('限定') ||
                        text.includes('発売') ||
                        text.includes('受付') ||
                        text.includes('申込') ||
                        text.includes('抽選') ||
                        text.includes('販売') ||
                        text.includes('公演'));
                    const isLongText = text.length > 5;
                    const hasDaySet = currentDay !== null;
                    console.log(`  キーワード有り: ${hasKeyword}`);
                    console.log(`  5文字以上: ${isLongText}`);
                    console.log(`  曜日設定済み: ${hasDaySet}`);
                    const currentJudgment = hasKeyword || (isLongText && hasDaySet);
                    console.log(`  現在の判定: ${currentJudgment ? '✓ 採用' : '✗ 無視'}`);
                    // より緩い判定（曜日が設定されていれば採用）
                    const relaxedJudgment = hasDaySet && text.length > 2;
                    console.log(`  緩い判定: ${relaxedJudgment ? '✓ 採用' : '✗ 無視'}`);
                    console.log('');
                }
            }
        });
        console.log('\n=== 現在の抽出結果 ===');
        const bookings = getAdvanceBookingFromCurrentSheet(sheet);
        if (bookings) {
            const dayLabels = {
                monday: '月曜日', tuesday: '火曜日', wednesday: '水曜日', thursday: '木曜日',
                friday: '金曜日', saturday: '土曜日', sunday: '日曜日'
            };
            Object.keys(bookings).forEach(day => {
                console.log(`\n${dayLabels[day]}:`);
                if (bookings[day].length > 0) {
                    bookings[day].forEach(booking => {
                        console.log(`  - ${booking}`);
                    });
                }
                else {
                    console.log('  - なし');
                }
            });
        }
        return bookings;
    }
    catch (error) {
        console.error('デバッグ実行エラー:', error);
    }
}
/**
 * 現在のシートのA列から先行予約情報を抽出（緩い判定版）
 */
function getAdvanceBookingFromCurrentSheetRelaxed(sheet) {
    try {
        const data = sheet.getDataRange().getValues();
        console.log(`現在のシート（A列のみ）から先行予約情報を抽出中（緩い判定）: ${sheet.getName()}`);
        const bookingsByDay = {
            monday: [], tuesday: [], wednesday: [], thursday: [],
            friday: [], saturday: [], sunday: []
        };
        let currentDay = null;
        let foundBookings = 0;
        // A列のみをチェック
        data.forEach((row, rowIndex) => {
            const cellValue = row[0];
            if (!cellValue) {
                return;
            }
            const text = cellValue.toString().trim();
            if (!text) {
                return;
            }
            // 日付形式を判定
            const datePattern = /^\s*[0-9０-９]+\s*[\/／]\s*[0-9０-９]+\s*[（\(]\s*[月火水木金土日]\s*[）\)]\s*$/;
            if (datePattern.test(text)) {
                // 日付行の場合、曜日を判定
                if (text.includes('月'))
                    currentDay = 'monday';
                else if (text.includes('火'))
                    currentDay = 'tuesday';
                else if (text.includes('水'))
                    currentDay = 'wednesday';
                else if (text.includes('木'))
                    currentDay = 'thursday';
                else if (text.includes('金'))
                    currentDay = 'friday';
                else if (text.includes('土'))
                    currentDay = 'saturday';
                else if (text.includes('日'))
                    currentDay = 'sunday';
                console.log(`先行予約（緩い判定）: 日付行検出 - ${text} (${currentDay}) [行${rowIndex + 1}]`);
            }
            else {
                // 【緩い判定】曜日が設定されていて、3文字以上のテキストなら採用
                const isBookingInfo = currentDay && text.length > 2;
                if (isBookingInfo && bookingsByDay[currentDay]) {
                    // 重複チェック
                    if (!bookingsByDay[currentDay].includes(text)) {
                        bookingsByDay[currentDay].push(text);
                        foundBookings++;
                        console.log(`先行予約（緩い判定）: ${currentDay} - ${text} [行${rowIndex + 1}]`);
                    }
                }
            }
        });
        console.log(`先行予約情報抽出完了（緩い判定）: 合計${foundBookings}件`);
        return bookingsByDay;
    }
    catch (error) {
        console.error('先行予約情報抽出エラー（緩い判定）:', error);
        return null;
    }
}
/**
 * FLAG番組のみを抽出（RS行がない翌々週対応）
 */
function extractFlagOnly(sheetName) {
    const config = getConfig();
    try {
        const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
        const sheet = spreadsheet.getSheetByName(sheetName);
        if (!sheet) {
            console.error(`指定されたシート「${sheetName}」が見つかりません`);
            console.log('利用可能なシート一覧:', spreadsheet.getSheets().map(s => s.getName()));
            return {};
        }
        console.log(`FLAG専用抽出: ${sheetName}`);
        // 金曜日のマーカーを探す
        const fridayMarkers = findFridayMarkers(sheet);
        console.log('金曜日マーカー:', fridayMarkers);
        if (fridayMarkers.newFridayRow === -1 || fridayMarkers.theBurnRow === -1) {
            console.error(`金曜日のマーカーが見つかりません - シート: ${sheetName}`);
            console.log('- newFridayRow:', fridayMarkers.newFridayRow);
            console.log('- theBurnRow:', fridayMarkers.theBurnRow);
            // シートの構造を確認
            const data = sheet.getDataRange().getValues();
            console.log('シートの行数:', data.length);
            console.log('最初の10行:', data.slice(0, 10).map((row, index) => `${index}: ${row.slice(0, 5).join(', ')}`));
            return {};
        }
        // 金曜日の範囲を決定
        const fridayRange = {
            start: fridayMarkers.newFridayRow + 1,
            end: fridayMarkers.theBurnRow - 1
        };
        console.log('金曜日の範囲:', fridayRange);
        // FLAGデータを抽出
        const flagData = extractFlagData(sheet, fridayRange, fridayMarkers.newFridayRow);
        if (!flagData) {
            console.error(`FLAGデータが見つかりませんでした - シート: ${sheetName}`);
            console.log('- 金曜日の範囲:', fridayRange);
            console.log('- newFridayRow:', fridayMarkers.newFridayRow);
            // 範囲内のデータを確認
            const data = sheet.getDataRange().getValues();
            if (fridayRange.start < data.length && fridayRange.end < data.length) {
                console.log('金曜日範囲のデータサンプル:');
                for (let i = fridayRange.start; i <= Math.min(fridayRange.end, fridayRange.start + 5); i++) {
                    console.log(`行${i}:`, data[i].slice(0, 10).join(', '));
                }
            }
            return {};
        }
        const results = {
            [sheetName]: {
                'FLAG': flagData
            }
        };
        console.log('FLAG抽出結果:', results);
        return results;
    }
    catch (error) {
        console.error('FLAG抽出エラー:', error);
        return {};
    }
}
/**
 * 金曜日のマーカーを探す
 */
function findFridayMarkers(sheet) {
    const data = sheet.getDataRange().getValues();
    let newFridayRow = -1;
    let theBurnRow = -1;
    // マーカー検索のログ
    const foundMarkers = [];
    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        // New!Friday マーカーを探す
        if (row.some(cell => cell && cell.toString().includes('New!Friday'))) {
            newFridayRow = i;
            foundMarkers.push(`New!Friday found at row ${i}`);
        }
        // THE BURN マーカーを探す
        if (row.some(cell => cell && cell.toString().includes('THE BURN'))) {
            theBurnRow = i;
            foundMarkers.push(`THE BURN found at row ${i}`);
        }
        // 類似のマーカーもログに記録
        const rowText = row.join(' ').toLowerCase();
        if (rowText.includes('friday') || rowText.includes('burn')) {
            foundMarkers.push(`Similar marker at row ${i}: ${rowText.substring(0, 100)}`);
        }
    }
    console.log('マーカー検索結果:', foundMarkers);
    return { newFridayRow, theBurnRow };
}
/**
 * FLAGデータを抽出
 */
function extractFlagData(sheet, fridayRange, newFridayRow) {
    const data = sheet.getDataRange().getValues();
    // New!Friday行からヘッダーを取得
    const fridayHeaderRow = data[newFridayRow];
    console.log('金曜日ヘッダー行:', fridayHeaderRow.slice(0, 10));
    let flagColIndex = -1;
    const headerPrograms = [];
    fridayHeaderRow.forEach((program, colIndex) => {
        const programStr = program ? program.toString() : '';
        headerPrograms.push(`${colIndex}: "${programStr}"`);
        if (program && programStr.includes('FLAG')) {
            flagColIndex = colIndex;
            console.log(`FLAG列発見: ${colIndex}列目 ("${programStr}")`);
        }
    });
    if (flagColIndex === -1) {
        console.error('FLAG列が見つかりません');
        console.log('利用可能なプログラム列:', headerPrograms.slice(0, 20));
        return null;
    }
    console.log(`FLAG列確定: ${flagColIndex}列目`);
    // FLAGのデータを抽出
    const rawContent = extractColumnData(data, flagColIndex, fridayRange);
    console.log('FLAG生データ:', rawContent);
    // 日付を計算
    const dayDates = calculateDayDates(sheet.getName());
    // 楽曲データベースを取得
    const musicDatabase = getMusicData();
    // FLAGデータを構造化
    const flagStructure = {
        'friday': structureFlag(rawContent, dayDates.friday, musicDatabase)
    };
    return flagStructure;
}
/**
 * 指定週数後のFLAGを抽出
 */
function extractFlagWeeksLater(weeksOffset) {
    const config = getConfig();
    const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
    // 数値型であることを確実にする
    const numericWeeksOffset = Number(weeksOffset);
    // 指定週数後のシートを取得
    const today = new Date();
    const targetDate = new Date(today.getTime() + (numericWeeksOffset * 7 * 24 * 60 * 60 * 1000));
    const dayOfWeek = targetDate.getDay() === 0 ? 7 : targetDate.getDay();
    const monday = new Date(targetDate.getTime() - (dayOfWeek - 1) * 24 * 60 * 60 * 1000);
    const sunday = new Date(monday.getTime() + 6 * 24 * 60 * 60 * 1000);
    const mondayYear = monday.getFullYear().toString().slice(-2);
    const mondayMonth = (monday.getMonth() + 1).toString();
    const mondayDay = monday.getDate().toString().padStart(2, '0');
    const sundayMonth = (sunday.getMonth() + 1).toString();
    const sundayDay = sunday.getDate().toString().padStart(2, '0');
    const sheetName = `${mondayYear}.${mondayMonth}.${mondayDay}-${sundayMonth}.${sundayDay}`;
    // 厳密な条件分岐で週ラベルを決定
    let offsetLabel;
    if (numericWeeksOffset === 0) {
        offsetLabel = '今週';
    }
    else if (numericWeeksOffset === 1) {
        offsetLabel = '翌週';
    }
    else if (numericWeeksOffset === 2) {
        offsetLabel = '翌々週';
    }
    else if (numericWeeksOffset === 3) {
        offsetLabel = '翌翌々週';
    }
    else if (numericWeeksOffset === 4) {
        offsetLabel = '翌々翌々週';
    }
    else {
        offsetLabel = `${numericWeeksOffset}週後`;
    }
    console.log(`${offsetLabel}のシート名: ${sheetName} (weeksOffset: ${weeksOffset} → ${numericWeeksOffset})`);
    // シートの存在確認
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
        console.error(`${offsetLabel}: シート「${sheetName}」が見つかりません`);
        console.log('利用可能なシート一覧:', spreadsheet.getSheets().map(s => s.getName()));
        return {};
    }
    console.log(`${offsetLabel}: シート「${sheetName}」が見つかりました`);
    return extractFlagOnly(sheetName);
}
/**
 * FLAG 4週間分を抽出（指定週を含めて4週分）
 */
function extractFlag4Weeks(startWeekOffset = 0) {
    // 数値型であることを確実にする
    const numericStartWeekOffset = Number(startWeekOffset);
    console.log(`FLAG 4週間分抽出開始（${numericStartWeekOffset}週後から4週分）`);
    const allResults = {};
    for (let i = 0; i < 4; i++) {
        const weekOffset = numericStartWeekOffset + i;
        // 厳密な条件分岐で週ラベルを決定
        let offsetLabel;
        if (weekOffset === 0) {
            offsetLabel = '今週';
        }
        else if (weekOffset === 1) {
            offsetLabel = '翌週';
        }
        else if (weekOffset === 2) {
            offsetLabel = '翌々週';
        }
        else if (weekOffset === 3) {
            offsetLabel = '翌翌々週';
        }
        else if (weekOffset === 4) {
            offsetLabel = '翌々翌々週';
        }
        else {
            offsetLabel = `${weekOffset}週後`;
        }
        try {
            const weekResults = extractFlagWeeksLater(weekOffset);
            if (weekResults && Object.keys(weekResults).length > 0) {
                // 週ラベルを付けて結果をマージ
                Object.keys(weekResults).forEach(sheetName => {
                    const newKey = `${offsetLabel}(${sheetName})`;
                    allResults[newKey] = weekResults[sheetName];
                });
                console.log(`${offsetLabel}: 抽出成功 (weekOffset: ${weekOffset})`);
            }
            else {
                console.log(`${offsetLabel}: データなし (weekOffset: ${weekOffset})`);
            }
        }
        catch (error) {
            console.error(`${offsetLabel}: 抽出エラー - ${error instanceof Error ? error.message : String(error)} (weekOffset: ${weekOffset})`);
        }
    }
    console.log('FLAG 4週間分抽出完了');
    return allResults;
}
/**
 * FLAG 4週間分を抽出してメール送信
 */
function extractFlag4WeeksAndSendEmail(startWeekOffset = 0) {
    const config = getConfig();
    try {
        const results = extractFlag4Weeks(startWeekOffset);
        if (!results || Object.keys(results).length === 0) {
            console.log('FLAG 4週間分のデータが見つかりません');
            // エラーメールを送信
            if (config.EMAIL_ADDRESS) {
                const startLabel = startWeekOffset === 0 ? '今週' : `${startWeekOffset}週後`;
                GmailApp.sendEmail(config.EMAIL_ADDRESS, `FLAG 4週間分の番組スケジュール（${startLabel}から） - エラー`, 'FLAG 4週間分のデータが見つかりませんでした。シートの構成を確認してください。');
            }
            return {};
        }
        console.log('FLAG 4週間分抽出結果:', results);
        const startLabel = startWeekOffset === 0 ? '今週' : `${startWeekOffset}週後`;
        sendProgramEmail(results, `FLAG 4週間分の番組スケジュール（${startLabel}から4週分）`);
        return results;
    }
    catch (error) {
        console.error('FLAG 4週間分抽出エラー:', error);
        // エラーメールを送信
        try {
            if (config.EMAIL_ADDRESS) {
                const startLabel = startWeekOffset === 0 ? '今週' : `${startWeekOffset}週後`;
                GmailApp.sendEmail(config.EMAIL_ADDRESS, `FLAG 4週間分の番組スケジュール（${startLabel}から） - エラー`, `データ抽出中にエラーが発生しました。\nエラー: ${error.toString()}`);
            }
        }
        catch (emailError) {
            console.error('エラーメール送信失敗:', emailError);
        }
        return {};
    }
}
/**
 * 今週から4週間分のFLAGを抽出してメール送信（便利関数）
 */
function extractFlagThisWeek4WeeksAndSendEmail() {
    return extractFlag4WeeksAndSendEmail(0);
}
/**
 * 指定した週のFLAGのみを抽出してメール送信
 */
function extractSpecificFlagAndSendEmail(sheetName) {
    const config = getConfig();
    try {
        const results = extractFlagOnly(sheetName);
        if (!results || Object.keys(results).length === 0) {
            console.log(`${sheetName}のFLAGデータが見つかりません`);
            // エラーメールを送信
            if (config.EMAIL_ADDRESS) {
                GmailApp.sendEmail(config.EMAIL_ADDRESS, `${sheetName} のFLAG番組スケジュール - エラー`, `${sheetName}のFLAGデータが見つかりませんでした。シートの構成を確認してください。`);
            }
            return {};
        }
        console.log(`${sheetName}のFLAG抽出結果:`, results);
        sendProgramEmail(results, `${sheetName} のFLAG番組スケジュール`);
        return results;
    }
    catch (error) {
        console.error(`${sheetName}のFLAG抽出エラー:`, error);
        // エラーメールを送信
        try {
            if (config.EMAIL_ADDRESS) {
                GmailApp.sendEmail(config.EMAIL_ADDRESS, `${sheetName} のFLAG番組スケジュール - エラー`, `データ抽出中にエラーが発生しました。\nエラー: ${error.toString()}`);
            }
        }
        catch (emailError) {
            console.error('エラーメール送信失敗:', emailError);
        }
        return {};
    }
}
/**
 * FLAG 4週間分のドキュメントを作成
 */
function createFlag4WeeksDocuments(flagResults, documentTitle = 'FLAG 4週間分') {
    const config = getConfig();
    if (!config.DOCUMENT_TEMPLATES || !config.DOCUMENT_TEMPLATES['FLAG'] || !config.DOCUMENT_OUTPUT_FOLDER_ID) {
        console.error('FLAGドキュメントテンプレート設定が不完全です。config.gsを確認してください。');
        return;
    }
    try {
        const outputFolder = DriveApp.getFolderById(config.DOCUMENT_OUTPUT_FOLDER_ID);
        const templateId = config.DOCUMENT_TEMPLATES['FLAG'];
        // テンプレートをコピー
        const templateFile = DriveApp.getFileById(templateId);
        // ファイル名を生成（現在日時ベース）
        const now = new Date();
        const dateStr = `${now.getFullYear()}${(now.getMonth() + 1).toString().padStart(2, '0')}${now.getDate().toString().padStart(2, '0')}`;
        const docName = `【連絡票】FLAG_4週間分_${dateStr}`;
        const copiedFile = templateFile.makeCopy(docName, outputFolder);
        const copiedDoc = DocumentApp.openById(copiedFile.getId());
        const body = copiedDoc.getBody();
        // FLAG 4週間分のプレースホルダーを置換
        replaceFlagPlaceholders(body, flagResults);
        // ドキュメントを保存して閉じる
        copiedDoc.saveAndClose();
        console.log(`FLAG 4週間分ドキュメント作成: ${docName}`);
        // 作成完了メールを送信
        if (config.EMAIL_ADDRESS) {
            const docUrl = `https://docs.google.com/document/d/${copiedFile.getId()}/edit`;
            GmailApp.sendEmail(config.EMAIL_ADDRESS, 'FLAG 4週間分ドキュメント作成完了', `FLAG 4週間分のドキュメントが作成されました：\n\n${docName}\n${docUrl}`);
        }
        return copiedFile.getId();
    }
    catch (error) {
        console.error('FLAG 4週間分ドキュメント作成エラー:', error);
        if (config.EMAIL_ADDRESS) {
            GmailApp.sendEmail(config.EMAIL_ADDRESS, 'FLAG 4週間分ドキュメント作成エラー', `ドキュメント作成中にエラーが発生しました。\nエラー: ${error.toString()}`);
        }
        return null;
    }
}
/**
 * FLAG 4週間分のプレースホルダーを置換
 */
function replaceFlagPlaceholders(body, flagResults) {
    // 基本情報を置換
    body.replaceText('{{番組名}}', 'FLAG');
    body.replaceText('{{PROGRAM_NAME}}', 'FLAG');
    // 生成日時を置換
    const now = new Date();
    const generatedTime = `${now.getFullYear()}/${(now.getMonth() + 1).toString().padStart(2, '0')}/${now.getDate().toString().padStart(2, '0')} ${now.getHours().toString().padStart(2, '0')}:${now.getMinutes().toString().padStart(2, '0')}`;
    body.replaceText('{{生成日時}}', generatedTime);
    body.replaceText('{{GENERATED_TIME}}', generatedTime);
    // 週ラベルのマッピング（week0から開始）
    const weekLabelMap = {
        '今週': 'week0',
        '翌週': 'week1',
        '翌々週': 'week2',
        '翌翌々週': 'week3'
    };
    // 各週のデータを処理
    Object.keys(flagResults).forEach(weekKey => {
        const weekData = flagResults[weekKey];
        // 週ラベルを抽出（例：'今週(25.6.02-6.08)' → '今週'）
        const weekLabel = weekKey.split('(')[0];
        const weekShortName = weekLabelMap[weekLabel] || 'unknown';
        console.log(`処理中の週: ${weekKey} → ${weekShortName}`);
        console.log(`week2デバッグ - weekLabel: "${weekLabel}", weekShortName: "${weekShortName}", データ存在: ${!!weekData}`);
        if (weekLabel === '翌々週') {
            console.log('WEEK2 詳細デバッグ:');
            console.log('- weekKey:', weekKey);
            console.log('- weekData keys:', weekData ? Object.keys(weekData) : 'なし');
            console.log('- FLAG data exists:', !!(weekData && weekData['FLAG']));
            console.log('- friday data exists:', !!(weekData && weekData['FLAG'] && weekData['FLAG']['friday']));
            if (weekData && weekData['FLAG'] && weekData['FLAG']['friday']) {
                console.log('- friday categories:', Object.keys(weekData['FLAG']['friday']));
            }
        }
        if (weekData && weekData['FLAG'] && weekData['FLAG']['friday']) {
            const fridayData = weekData['FLAG']['friday'];
            // 各カテゴリのプレースホルダーを置換
            Object.keys(fridayData).forEach(category => {
                const items = fridayData[category];
                let content = '';
                if (Array.isArray(items)) {
                    if (category === '楽曲') {
                        // 楽曲の特別処理
                        content = formatMusicList(items);
                        // 楽曲の異なるフォーマット用プレースホルダー
                        const musicSimple = formatMusicListSimple(items);
                        const musicBullet = formatMusicListBullet(items);
                        const musicTable = formatMusicListTable(items);
                        const musicOneLine = formatMusicListOneLine(items);
                        // 週別楽曲プレースホルダーを置換
                        body.replaceText(`{{${weekShortName}_楽曲_シンプル}}`, musicSimple);
                        body.replaceText(`{{${weekShortName}_楽曲_箇条書き}}`, musicBullet);
                        body.replaceText(`{{${weekShortName}_楽曲_テーブル}}`, musicTable);
                        body.replaceText(`{{${weekShortName}_楽曲_一行}}`, musicOneLine);
                        body.replaceText(`{{${weekLabel}_楽曲_シンプル}}`, musicSimple);
                        body.replaceText(`{{${weekLabel}_楽曲_箇条書き}}`, musicBullet);
                        body.replaceText(`{{${weekLabel}_楽曲_テーブル}}`, musicTable);
                        body.replaceText(`{{${weekLabel}_楽曲_一行}}`, musicOneLine);
                    }
                    else {
                        content = items.filter(item => item !== 'ー').join('\n');
                    }
                }
                else {
                    content = items !== 'ー' ? items.toString() : '';
                }
                // 基本プレースホルダーを置換
                // 週番号形式（week0, week1, etc.）
                const weekShortPlaceholder = `{{${weekShortName}_${category}}}`;
                const weekLabelPlaceholder = `{{${weekLabel}_${category}}}`;
                if (weekLabel === '翌々週') {
                    console.log(`WEEK2 プレースホルダー置換:`, {
                        category,
                        weekShortPlaceholder,
                        weekLabelPlaceholder,
                        content: content || 'ー',
                        contentLength: (content || 'ー').length
                    });
                }
                body.replaceText(weekShortPlaceholder, content || 'ー');
                // 週ラベル形式（今週、翌週、etc.）
                body.replaceText(weekLabelPlaceholder, content || 'ー');
                // 英語版プレースホルダー
                const englishCategory = convertCategoryToEnglish(category);
                if (englishCategory) {
                    body.replaceText(`{{${weekShortName.toUpperCase()}_${englishCategory}}}`, content || 'ー');
                }
                // 短縮版プレースホルダー
                const shortCategory = getShortCategoryName(category);
                if (shortCategory) {
                    body.replaceText(`{{${weekShortName}_${shortCategory}}}`, content || 'ー');
                    body.replaceText(`{{${weekLabel}_${shortCategory}}}`, content || 'ー');
                }
            });
        }
    });
}
/**
 * FLAG 4週間分を抽出してドキュメント作成
 */
function extractFlag4WeeksAndCreateDocument(startWeekOffset = 0) {
    try {
        console.log('FLAG 4週間分抽出＋ドキュメント作成開始');
        // FLAG 4週間分を抽出
        const flagResults = extractFlag4Weeks(startWeekOffset);
        if (!flagResults || Object.keys(flagResults).length === 0) {
            console.log('FLAG 4週間分のデータが見つかりません');
            return null;
        }
        // ドキュメントを作成
        const docId = createFlag4WeeksDocuments(flagResults);
        console.log('FLAG 4週間分抽出＋ドキュメント作成完了');
        return docId;
    }
    catch (error) {
        console.error('FLAG 4週間分抽出＋ドキュメント作成エラー:', error);
        return null;
    }
}
/**
 * FLAG 4週間分を抽出してメール送信＋ドキュメント作成
 */
function extractFlag4WeeksAndSendEmailAndCreateDocument(startWeekOffset = 0) {
    try {
        console.log('FLAG 4週間分：抽出＋メール＋ドキュメント作成開始');
        // FLAG 4週間分を抽出
        const flagResults = extractFlag4Weeks(startWeekOffset);
        if (!flagResults || Object.keys(flagResults).length === 0) {
            console.log('FLAG 4週間分のデータが見つかりません');
            return null;
        }
        // メール送信
        const startLabel = startWeekOffset === 0 ? '今週' : `${startWeekOffset}週後`;
        sendProgramEmail(flagResults, `FLAG 4週間分の番組スケジュール（${startLabel}から4週分）`);
        // ドキュメントを作成
        const docId = createFlag4WeeksDocuments(flagResults);
        console.log('FLAG 4週間分：抽出＋メール＋ドキュメント作成完了');
        return docId;
    }
    catch (error) {
        console.error('FLAG 4週間分：抽出＋メール＋ドキュメント作成エラー:', error);
        return null;
    }
}
/**
 * FLAG用プレースホルダーのサンプルテンプレート
 */
function showFlagTemplateSample() {
    console.log('=== FLAG 4週間分テンプレートの例 ===');
    console.log('');
    console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    console.log('【{{番組名}}】4週間分スケジュール');
    console.log('生成日時: {{生成日時}}');
    console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    console.log('');
    console.log('■ 今週分');
    console.log('日付: {{今週_日付}} または {{week0_日付}}');
    console.log('');
    console.log('【楽曲】');
    console.log('{{今週_楽曲_テーブル}} または {{week0_楽曲_テーブル}}');
    console.log('');
    console.log('【12:40 電話パブ】');
    console.log('{{今週_12:40 電話パブ}} または {{week0_1240パブ}}');
    console.log('');
    console.log('【13:29 パブリシティ】');
    console.log('{{今週_13:29 パブリシティ}} または {{week0_1329パブ}}');
    console.log('');
    console.log('■ 翌週分');
    console.log('日付: {{翌週_日付}} または {{week1_日付}}');
    console.log('楽曲: {{翌週_楽曲}} または {{week1_楽曲}}');
    console.log('');
    console.log('■ 翌々週分');
    console.log('日付: {{翌々週_日付}} または {{week2_日付}}');
    console.log('楽曲: {{翌々週_楽曲}} または {{week2_楽曲}}');
    console.log('');
    console.log('■ 翌翌々週分');
    console.log('日付: {{翌翌々週_日付}} または {{week3_日付}}');
    console.log('楽曲: {{翌翌々週_楽曲}} または {{week3_楽曲}}');
    console.log('');
    console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    console.log('');
    console.log('=== 利用可能なプレースホルダー形式 ===');
    console.log('');
    console.log('◆ 週ラベル形式:');
    console.log('{{今週_カテゴリ名}}, {{翌週_カテゴリ名}}, {{翌々週_カテゴリ名}}, {{翌翌々週_カテゴリ名}}');
    console.log('');
    console.log('◆ 週番号形式:');
    console.log('{{week0_カテゴリ名}}, {{week1_カテゴリ名}}, {{week2_カテゴリ名}}, {{week3_カテゴリ名}}');
    console.log('');
    console.log('◆ 楽曲の特別フォーマット:');
    console.log('{{今週_楽曲_シンプル}} - 番号なし');
    console.log('{{今週_楽曲_箇条書き}} - 箇条書き');
    console.log('{{今週_楽曲_テーブル}} - テーブル形式');
    console.log('{{今週_楽曲_一行}} - 1行ずつ');
    console.log('');
    console.log('◆ FLAGのカテゴリ名:');
    console.log('日付, 楽曲, 12:40 電話パブ, 13:29 パブリシティ, 13:40 パブリシティ,');
    console.log('12:15 リポート案件, 14:29 リポート案件, 時間指定なし告知, 先行予約, ゲスト');
}
/**
 * GoogleドキュメントをDocxファイルに変換してメール送信
 */
function sendDocumentAsDocx(docId, fileName, subject, body) {
    const config = getConfig();
    if (!config.EMAIL_ADDRESS) {
        console.error('メールアドレスが設定されていません');
        return false;
    }
    try {
        // Google Drive APIのexport URLを使用してDocx形式で取得
        const url = `https://docs.google.com/document/d/${docId}/export?format=docx`;
        // アクセストークンを取得
        const token = ScriptApp.getOAuthToken();
        // Docxファイルを取得
        const response = UrlFetchApp.fetch(url, {
            method: 'get',
            headers: {
                'Authorization': `Bearer ${token}`
            }
        });
        if (response.getResponseCode() !== 200) {
            throw new Error(`ドキュメントの取得に失敗しました。レスポンスコード: ${response.getResponseCode()}`);
        }
        // Blobとして取得
        const docxBlob = response.getBlob();
        // ファイル名を設定（.docx拡張子を追加）
        const docxFileName = fileName.endsWith('.docx') ? fileName : `${fileName}.docx`;
        docxBlob.setName(docxFileName);
        // メール送信（Docx添付）
        GmailApp.sendEmail(config.EMAIL_ADDRESS, subject, body, {
            attachments: [docxBlob]
        });
        console.log(`Docxファイル送信成功: ${docxFileName}`);
        return true;
    }
    catch (error) {
        console.error('Docxファイル送信エラー:', error);
        // エラーメールを送信（添付なし）
        try {
            GmailApp.sendEmail(config.EMAIL_ADDRESS, `${subject} - Docx変換エラー`, `Docxファイルの変換・送信でエラーが発生しました。\n\nエラー: ${error.toString()}\n\nGoogleドキュメントURL: https://docs.google.com/document/d/${docId}/edit`);
        }
        catch (emailError) {
            console.error('エラーメール送信失敗:', emailError);
        }
        return false;
    }
}
/**
 * 番組の放送日に基づいた日付文字列を生成
 */
function generateBroadcastDateString(programName, weekOffset = 1) {
    const today = new Date();
    const targetDate = new Date(today.getTime() + (weekOffset * 7 * 24 * 60 * 60 * 1000));
    // 番組ごとの放送曜日を定義
    const broadcastDays = {
        'ちょうどいいラジオ': 1, // 月曜日基準（月～木放送なので週の概念）
        'PRIME TIME': 1, // 月曜日基準（月～木放送なので週の概念）
        'FLAG': 5, // 金曜日
        'Route 847': 6, // 土曜日
        'God Bless Saturday': 6 // 土曜日
    };
    let broadcastDay = broadcastDays[programName];
    if (broadcastDay === undefined) {
        // 不明な番組の場合は月曜日基準
        broadcastDay = 1;
    }
    // 対象週の放送日を計算
    const dayOfWeek = targetDate.getDay() === 0 ? 7 : targetDate.getDay();
    const monday = new Date(targetDate.getTime() - (dayOfWeek - 1) * 24 * 60 * 60 * 1000);
    let broadcastDate;
    if (broadcastDay === 1) {
        // 月曜日基準（週の概念）
        broadcastDate = monday;
    }
    else {
        // 具体的な曜日
        broadcastDate = new Date(monday.getTime() + (broadcastDay - 1) * 24 * 60 * 60 * 1000);
    }
    const year = broadcastDate.getFullYear();
    const month = (broadcastDate.getMonth() + 1).toString().padStart(2, '0');
    const day = broadcastDate.getDate().toString().padStart(2, '0');
    return `${year}${month}${day}`;
}
/**
 * 指定番組の今週分ドキュメントを生成（内部関数）
 */
function extractSpecificProgramThisWeek(programName) {
    try {
        const config = getConfig();
        const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
        const thisWeekSheet = getSheetByWeek(spreadsheet, 0);
        if (!thisWeekSheet) {
            throw new Error('今週のシートが見つかりません');
        }
        console.log(`${programName} 今週分処理: ${thisWeekSheet.getName()}`);
        // 今週のデータを抽出
        const weekData = extractStructuredWeekData(thisWeekSheet);
        if (!weekData || !weekData[programName]) {
            throw new Error(`${programName}のデータが見つかりません`);
        }
        const results = { '今週': weekData };
        // 指定番組のみのドキュメントを作成
        const docId = createSingleProgramDocument(programName, weekData[programName], config.DOCUMENT_TEMPLATES[programName], DriveApp.getFolderById(config.DOCUMENT_OUTPUT_FOLDER_ID), '今週', getStartDateFromSheetName(thisWeekSheet.getName()));
        if (!docId) {
            throw new Error('ドキュメント作成に失敗しました');
        }
        const docUrl = `https://docs.google.com/document/d/${docId}/edit`;
        return {
            success: true,
            docId: docId,
            url: docUrl,
            sheetName: thisWeekSheet.getName()
        };
    }
    catch (error) {
        console.error(`${programName} 今週分抽出エラー:`, error);
        return {
            success: false,
            error: error.toString()
        };
    }
}
/**
 * 指定番組の翌週分ドキュメントを生成（内部関数）
 */
function extractSpecificProgramNextWeek(programName) {
    try {
        const config = getConfig();
        const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
        const nextWeekSheet = getSheetByWeek(spreadsheet, 1);
        if (!nextWeekSheet) {
            throw new Error('翌週のシートが見つかりません');
        }
        console.log(`${programName} 翌週分処理: ${nextWeekSheet.getName()}`);
        // 翌週のデータを抽出
        const weekData = extractStructuredWeekData(nextWeekSheet);
        if (!weekData || !weekData[programName]) {
            throw new Error(`${programName}のデータが見つかりません`);
        }
        const results = { '翌週': weekData };
        // 指定番組のみのドキュメントを作成
        const docId = createSingleProgramDocument(programName, weekData[programName], config.DOCUMENT_TEMPLATES[programName], DriveApp.getFolderById(config.DOCUMENT_OUTPUT_FOLDER_ID), '翌週', getStartDateFromSheetName(nextWeekSheet.getName()));
        if (!docId) {
            throw new Error('ドキュメント作成に失敗しました');
        }
        const docUrl = `https://docs.google.com/document/d/${docId}/edit`;
        return {
            success: true,
            docId: docId,
            url: docUrl,
            sheetName: nextWeekSheet.getName()
        };
    }
    catch (error) {
        console.error(`${programName} 翌週分抽出エラー:`, error);
        return {
            success: false,
            error: error.toString()
        };
    }
}
/**
 * 【金曜朝用】ちょうどいいラジオの翌週分ドキュメントを自動生成＋Docx送信
 */
function autoGenerateChoudoDocument() {
    console.log('=== ちょうどいいラジオ 翌週分ドキュメント自動生成開始 ===');
    try {
        const config = getConfig();
        const result = extractSpecificProgramNextWeek('ちょうどいいラジオ');
        if (result && result.success) {
            console.log('ちょうどいいラジオ 翌週分ドキュメント生成成功');
            // Docxファイル名を生成
            const broadcastDateStr = generateBroadcastDateString('ちょうどいいラジオ', 1);
            const fileName = `【連絡票】ちょうどいいラジオ_${broadcastDateStr}週`;
            // Docxファイルとしてメール送信
            const docxSent = sendDocumentAsDocx(result.docId, fileName, '【自動生成完了】ちょうどいいラジオ 翌週分ドキュメント', `ちょうどいいラジオの翌週分ドキュメントが自動生成されました。\n\nシート: ${result.sheetName}\nGoogleドキュメントURL: ${result.url}\n\n生成日時: ${new Date().toLocaleString('ja-JP')}`);
            return Object.assign(Object.assign({}, result), { docxSent: docxSent });
        }
        else {
            throw new Error(result ? result.error : '不明なエラー');
        }
    }
    catch (error) {
        console.error('ちょうどいいラジオ 自動生成エラー:', error);
        // エラーメールを送信
        const config = getConfig();
        if (config.EMAIL_ADDRESS) {
            GmailApp.sendEmail(config.EMAIL_ADDRESS, '【自動生成エラー】ちょうどいいラジオ 翌週分ドキュメント', `ちょうどいいラジオの翌週分ドキュメント自動生成でエラーが発生しました。\n\nエラー: ${error.toString()}\n\n生成日時: ${new Date().toLocaleString('ja-JP')}`);
        }
        return { success: false, error: error.toString(), docxSent: false };
    }
}
/**
 * 【金曜朝用】PRIME TIMEの翌週分ドキュメントを自動生成＋Docx送信
 */
function autoGeneratePrimeTimeDocument() {
    console.log('=== PRIME TIME 翌週分ドキュメント自動生成開始 ===');
    try {
        const config = getConfig();
        const result = extractSpecificProgramNextWeek('PRIME TIME');
        if (result && result.success) {
            console.log('PRIME TIME 翌週分ドキュメント生成成功');
            // Docxファイル名を生成
            const broadcastDateStr = generateBroadcastDateString('PRIME TIME', 1);
            const fileName = `【連絡票】PRIMETIME_${broadcastDateStr}週`;
            // Docxファイルとしてメール送信
            const docxSent = sendDocumentAsDocx(result.docId, fileName, '【自動生成完了】PRIME TIME 翌週分ドキュメント', `PRIME TIMEの翌週分ドキュメントが自動生成されました。\n\nシート: ${result.sheetName}\nGoogleドキュメントURL: ${result.url}\n\n生成日時: ${new Date().toLocaleString('ja-JP')}`);
            return Object.assign(Object.assign({}, result), { docxSent: docxSent });
        }
        else {
            throw new Error(result ? result.error : '不明なエラー');
        }
    }
    catch (error) {
        console.error('PRIME TIME 自動生成エラー:', error);
        // エラーメールを送信
        const config = getConfig();
        if (config.EMAIL_ADDRESS) {
            GmailApp.sendEmail(config.EMAIL_ADDRESS, '【自動生成エラー】PRIME TIME 翌週分ドキュメント', `PRIME TIMEの翌週分ドキュメント自動生成でエラーが発生しました。\n\nエラー: ${error.toString()}\n\n生成日時: ${new Date().toLocaleString('ja-JP')}`);
        }
        return { success: false, error: error.toString(), docxSent: false };
    }
}
/**
 * 【水曜朝用】FLAGの今週からの4週分ドキュメントを自動生成＋Docx送信
 */
function autoGenerateFlagDocument() {
    console.log('=== FLAG 4週間分ドキュメント自動生成開始 ===');
    try {
        const config = getConfig();
        const docId = extractFlag4WeeksAndCreateDocument(0); // 今週から4週分
        if (docId) {
            console.log('FLAG 4週間分ドキュメント生成成功');
            // FLAGの放送日（今週の金曜日）の日付を使用
            const broadcastDateStr = generateBroadcastDateString('FLAG', 0);
            const fileName = `【連絡票】FLAG_4週間分_${broadcastDateStr}`;
            const docUrl = `https://docs.google.com/document/d/${docId}/edit`;
            // Docxファイルとしてメール送信
            const docxSent = sendDocumentAsDocx(docId, fileName, '【自動生成完了】FLAG 4週間分ドキュメント', `FLAGの4週間分ドキュメント（今週からの4週分）が自動生成されました。\n\nGoogleドキュメントURL: ${docUrl}\n\n生成日時: ${new Date().toLocaleString('ja-JP')}`);
            return { success: true, docId: docId, url: docUrl, docxSent: docxSent };
        }
        else {
            throw new Error('ドキュメント作成に失敗しました');
        }
    }
    catch (error) {
        console.error('FLAG 自動生成エラー:', error);
        // エラーメールを送信
        const config = getConfig();
        if (config.EMAIL_ADDRESS) {
            GmailApp.sendEmail(config.EMAIL_ADDRESS, '【自動生成エラー】FLAG 4週間分ドキュメント', `FLAGの4週間分ドキュメント自動生成でエラーが発生しました。\n\nエラー: ${error.toString()}\n\n生成日時: ${new Date().toLocaleString('ja-JP')}`);
        }
        return { success: false, error: error.toString(), docxSent: false };
    }
}
/**
 * 【水曜朝用】Route 847の今週分ドキュメントを自動生成＋Docx送信
 */
function autoGenerateRoute847Document() {
    console.log('=== Route 847 今週分ドキュメント自動生成開始 ===');
    try {
        const config = getConfig();
        const result = extractSpecificProgramThisWeek('Route 847');
        if (result && result.success) {
            console.log('Route 847 今週分ドキュメント生成成功');
            // Route 847の放送日（今週の土曜日）の日付を使用
            const broadcastDateStr = generateBroadcastDateString('Route 847', 0);
            const fileName = `【連絡票】Route847_${broadcastDateStr}`;
            // Docxファイルとしてメール送信
            const docxSent = sendDocumentAsDocx(result.docId, fileName, '【自動生成完了】Route 847 今週分ドキュメント', `Route 847の今週分ドキュメントが自動生成されました。\n\nシート: ${result.sheetName}\nGoogleドキュメントURL: ${result.url}\n\n生成日時: ${new Date().toLocaleString('ja-JP')}`);
            return Object.assign(Object.assign({}, result), { docxSent: docxSent });
        }
        else {
            throw new Error(result ? result.error : '不明なエラー');
        }
    }
    catch (error) {
        console.error('Route 847 自動生成エラー:', error);
        // エラーメールを送信
        const config = getConfig();
        if (config.EMAIL_ADDRESS) {
            GmailApp.sendEmail(config.EMAIL_ADDRESS, '【自動生成エラー】Route 847 今週分ドキュメント', `Route 847の今週分ドキュメント自動生成でエラーが発生しました。\n\nエラー: ${error.toString()}\n\n生成日時: ${new Date().toLocaleString('ja-JP')}`);
        }
        return { success: false, error: error.toString(), docxSent: false };
    }
}
/**
 * 【木曜朝用】God Bless Saturdayの今週分ドキュメントを自動生成＋Docx送信
 */
function autoGenerateGodBlessDocument() {
    console.log('=== God Bless Saturday 今週分ドキュメント自動生成開始 ===');
    try {
        const config = getConfig();
        const result = extractSpecificProgramThisWeek('God Bless Saturday');
        if (result && result.success) {
            console.log('God Bless Saturday 今週分ドキュメント生成成功');
            // God Bless Saturdayの放送日（今週の土曜日）の日付を使用
            const broadcastDateStr = generateBroadcastDateString('God Bless Saturday', 0);
            const fileName = `【連絡票】GodBlessSaturday_${broadcastDateStr}`;
            // Docxファイルとしてメール送信
            const docxSent = sendDocumentAsDocx(result.docId, fileName, '【自動生成完了】God Bless Saturday 今週分ドキュメント', `God Bless Saturdayの今週分ドキュメントが自動生成されました。\n\nシート: ${result.sheetName}\nGoogleドキュメントURL: ${result.url}\n\n生成日時: ${new Date().toLocaleString('ja-JP')}`);
            return Object.assign(Object.assign({}, result), { docxSent: docxSent });
        }
        else {
            throw new Error(result ? result.error : '不明なエラー');
        }
    }
    catch (error) {
        console.error('God Bless Saturday 自動生成エラー:', error);
        // エラーメールを送信
        const config = getConfig();
        if (config.EMAIL_ADDRESS) {
            GmailApp.sendEmail(config.EMAIL_ADDRESS, '【自動生成エラー】God Bless Saturday 今週分ドキュメント', `God Bless Saturdayの今週分ドキュメント自動生成でエラーが発生しました。\n\nエラー: ${error.toString()}\n\n生成日時: ${new Date().toLocaleString('ja-JP')}`);
        }
        return { success: false, error: error.toString(), docxSent: false };
    }
}
/**
 * 全自動生成関数の実行状況をテスト（Docx送信込み）
 */
function testAllAutoGeneration() {
    console.log('=== 全番組自動生成テスト（Docx送信込み）===');
    const results = [];
    console.log('\n1. ちょうどいいラジオテスト（翌週分）');
    try {
        const choudoResult = autoGenerateChoudoDocument();
        results.push({
            program: 'ちょうどいいラジオ',
            success: choudoResult.success,
            error: choudoResult.error,
            docxSent: choudoResult.docxSent
        });
    }
    catch (error) {
        results.push({ program: 'ちょうどいいラジオ', success: false, error: error.toString(), docxSent: false });
    }
    console.log('\n2. PRIME TIMEテスト（翌週分）');
    try {
        const primeResult = autoGeneratePrimeTimeDocument();
        results.push({
            program: 'PRIME TIME',
            success: primeResult.success,
            error: primeResult.error,
            docxSent: primeResult.docxSent
        });
    }
    catch (error) {
        results.push({ program: 'PRIME TIME', success: false, error: error.toString(), docxSent: false });
    }
    console.log('\n3. FLAGテスト（今週からの4週分）');
    try {
        const flagResult = autoGenerateFlagDocument();
        results.push({
            program: 'FLAG',
            success: flagResult.success,
            error: flagResult.error,
            docxSent: flagResult.docxSent
        });
    }
    catch (error) {
        results.push({ program: 'FLAG', success: false, error: error.toString(), docxSent: false });
    }
    console.log('\n4. Route 847テスト（今週分）');
    try {
        const routeResult = autoGenerateRoute847Document();
        results.push({
            program: 'Route 847',
            success: routeResult.success,
            error: routeResult.error,
            docxSent: routeResult.docxSent
        });
    }
    catch (error) {
        results.push({ program: 'Route 847', success: false, error: error.toString(), docxSent: false });
    }
    console.log('\n5. God Bless Saturdayテスト（今週分）');
    try {
        const godBlessResult = autoGenerateGodBlessDocument();
        results.push({
            program: 'God Bless Saturday',
            success: godBlessResult.success,
            error: godBlessResult.error,
            docxSent: godBlessResult.docxSent
        });
    }
    catch (error) {
        results.push({ program: 'God Bless Saturday', success: false, error: error.toString(), docxSent: false });
    }
    console.log('\n=== テスト結果 ===');
    results.forEach(result => {
        const status = result.success ? '✓' : '✗';
        const docxStatus = result.docxSent ? '✓Docx送信済' : '✗Docx送信失敗';
        console.log(`${status} ${result.program}: ${result.success ? '成功' : result.error} ${result.success ? docxStatus : ''}`);
    });
    return results;
}
/**
 * Docx送信機能のテスト（単体）
 */
function testDocxSending() {
    console.log('=== Docx送信機能テスト ===');
    // 既存のドキュメントIDでテスト（実際のドキュメントIDに置き換えてください）
    const testDocId = 'YOUR_TEST_DOCUMENT_ID_HERE';
    const testFileName = 'テスト用ドキュメント';
    const testSubject = 'Docx送信テスト';
    const testBody = 'これはDocx送信機能のテストです。';
    console.log(`テスト対象ドキュメントID: ${testDocId}`);
    const result = sendDocumentAsDocx(testDocId, testFileName, testSubject, testBody);
    console.log(`テスト結果: ${result ? '成功' : '失敗'}`);
    return result;
}
/**
 * 緩い判定でのテスト実行
 */
function testAdvanceBookingRelaxed() {
    console.log('=== 先行予約情報テスト（緩い判定版）===');
    const config = getConfig();
    const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
    const thisWeekSheet = getSheetByWeek(spreadsheet, 0);
    if (!thisWeekSheet) {
        console.log('今週のシートが見つかりません');
        return;
    }
    const bookings = getAdvanceBookingFromCurrentSheetRelaxed(thisWeekSheet);
    if (bookings) {
        const dayLabels = {
            monday: '月曜日', tuesday: '火曜日', wednesday: '水曜日', thursday: '木曜日',
            friday: '金曜日', saturday: '土曜日', sunday: '日曜日'
        };
        Object.keys(bookings).forEach(day => {
            console.log(`\n【${dayLabels[day]}】`);
            if (bookings[day].length > 0) {
                bookings[day].forEach(booking => {
                    console.log(`  - ${booking}`);
                });
            }
            else {
                console.log('  - 情報なし');
            }
        });
    }
    else {
        console.log('先行予約情報の取得に失敗しました');
    }
    return bookings;
}
/**
 * 設定確認用のテスト関数
 */
function testConfig() {
    try {
        const config = getConfig();
        console.log('設定読み込み成功:');
        console.log('SPREADSHEET_ID:', config.SPREADSHEET_ID);
        console.log('EMAIL_ADDRESS:', config.EMAIL_ADDRESS);
        console.log('MUSIC_SHEET_NAME:', config.MUSIC_SHEET_NAME);
        console.log('PORTSIDE_CALENDAR_ID:', config.PORTSIDE_CALENDAR_ID);
        return config;
    }
    catch (error) {
        console.error('設定読み込みエラー:', error);
    }
}
/**
 * PORTSIDE情報をテスト取得（デバッグ用）
 */
function testPortsideInformation() {
    const today = new Date();
    const targetDate = new Date(today.getTime());
    // 今週の月曜日を計算
    const dayOfWeek = targetDate.getDay() === 0 ? 7 : targetDate.getDay();
    const monday = new Date(targetDate.getTime() - (dayOfWeek - 1) * 24 * 60 * 60 * 1000);
    console.log('=== PORTSIDE情報テスト ===');
    console.log(`対象週の月曜日: ${monday.toDateString()}`);
    const portsideInfo = getPortsideInformationFromCalendar(monday);
    if (portsideInfo) {
        console.log('取得結果:');
        Object.keys(portsideInfo).forEach(day => {
            console.log(`${day}: ${portsideInfo[day].length}件`);
            portsideInfo[day].forEach(info => {
                console.log(`  - ${info}`);
            });
        });
    }
    else {
        console.log('PORTSIDE情報の取得に失敗しました');
    }
    return portsideInfo;
}
/**
 * 先行予約テスト関数
 */
function testAdvanceBooking() {
    console.log('=== 先行予約情報テスト ===');
    const config = getConfig();
    const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
    const thisWeekSheet = getSheetByWeek(spreadsheet, 0);
    if (!thisWeekSheet) {
        console.log('今週のシートが見つかりません');
        return;
    }
    const bookings = getAdvanceBookingFromCurrentSheet(thisWeekSheet);
    if (bookings) {
        const dayLabels = {
            monday: '月曜日',
            tuesday: '火曜日',
            wednesday: '水曜日',
            thursday: '木曜日',
            friday: '金曜日',
            saturday: '土曜日',
            sunday: '日曜日'
        };
        Object.keys(bookings).forEach(day => {
            console.log(`\n【${dayLabels[day]}】`);
            if (bookings[day].length > 0) {
                bookings[day].forEach(booking => {
                    console.log(`  - ${booking}`);
                });
            }
            else {
                console.log('  - 情報なし');
            }
        });
    }
    else {
        console.log('先行予約情報の取得に失敗しました');
    }
    return bookings;
}
/**
 * ドキュメントテンプレート作成用のサンプル関数
 * この関数を実行して、テンプレートの作り方を確認できます
 */
function createSampleTemplate() {
    console.log('=== Googleドキュメントテンプレートの作り方 ===');
    console.log('');
    console.log('1. 新しいGoogleドキュメントを作成');
    console.log('2. 以下のプレースホルダーを使用してテンプレートを作成:');
    console.log('');
    console.log('◆ 基本プレースホルダー:');
    console.log('{{番組名}} または {{PROGRAM_NAME}} - 番組名');
    console.log('{{生成日時}} または {{GENERATED_TIME}} - 生成日時');
    console.log('');
    console.log('◆ 曜日別プレースホルダー（例：月曜日の場合）:');
    console.log('{{monday_日付}} - 月曜日の日付');
    console.log('{{monday_楽曲}} - 月曜日の楽曲リスト');
    console.log('{{monday_ゲスト}} - 月曜日のゲスト情報');
    console.log('');
    console.log('◆ 【ちょうどいいラジオ】専用プレースホルダー:');
    console.log('');
    console.log('--- 曜日別コンテンツ ---');
    console.log('{{monday_7:28パブ告知}} または {{monday_728パブ}} - 7:28のパブリシティ');
    console.log('{{monday_時間指定なし告知}} または {{monday_告知}} - その他の告知');
    console.log('{{monday_YOKOHAMA PORTSIDE INFORMATION}} または {{monday_ポートサイド}} - ポートサイド情報');
    console.log('{{monday_先行予約}} または {{monday_予約}} - 先行予約情報');
    console.log('{{monday_ラジオショッピング}} または {{monday_ラジショ}} - ラジオショッピング');
    console.log('{{monday_はぴねすくらぶ}} または {{monday_はぴねす}} - はぴねすくらぶ');
    console.log('{{tuesday_ヨコアリくん}} または {{tuesday_ヨコアリ}} - ヨコアリくん（火曜のみ）');
    console.log('{{monday_放送後}} - 放送後情報（収録予定がある場合）');
    console.log('');
    console.log('--- 収録予定（番組全体で共通）---');
    console.log('{{ちょうどいい暮らし収録予定}} または {{暮らし収録}} - ちょうどいい暮らし');
    console.log('{{ここが知りたい不動産収録予定}} または {{不動産収録}} - 不動産コーナー');
    console.log('{{ちょうどいい歯ッピー収録予定}} または {{歯ッピー収録}} - 歯科コーナー');
    console.log('{{ちょうどいいおカネの話収録予定}} または {{おカネ収録}} - お金コーナー');
    console.log('{{ちょうどいいごりごり隊収録予定}} または {{ごりごり収録}} - ごりごり隊');
    console.log('{{ビジネスアイ収録予定}} または {{ビジネスアイ収録}} - ビジネスアイ');
    console.log('');
    console.log('◆ 【PRIME TIME】専用プレースホルダー:');
    console.log('{{monday_19:41Traffic}} または {{monday_1941Traffic}} - 19:41のトラフィック情報');
    console.log('{{monday_19:43}} - 19:43のコンテンツ');
    console.log('{{monday_20:51}} - 20:51のコンテンツ');
    console.log('{{monday_営業コーナー}} または {{monday_営業}} - 営業コーナー');
    console.log('{{monday_指定曲}} - 指定曲');
    console.log('{{monday_時間指定なしパブ}} または {{monday_パブ}} - パブリシティ');
    console.log('{{monday_ラジショピ}} または {{monday_ラジショ}} - ラジオショッピング');
    console.log('{{monday_先行予約・限定予約}} または {{monday_予約}} - 予約情報');
    console.log('');
    console.log('◆ 【FLAG】専用プレースホルダー:');
    console.log('{{friday_12:40 電話パブ}} または {{friday_1240パブ}} - 12:40電話パブ');
    console.log('{{friday_13:29 パブリシティ}} または {{friday_1329パブ}} - 13:29パブリシティ');
    console.log('{{friday_13:40 パブリシティ}} または {{friday_1340パブ}} - 13:40パブリシティ');
    console.log('{{friday_12:15 リポート案件}} または {{friday_1215リポート}} - 12:15リポート');
    console.log('{{friday_14:29 リポート案件}} または {{friday_1429リポート}} - 14:29リポート');
    console.log('');
    console.log('◆ 【God Bless Saturday】専用プレースホルダー:');
    console.log('{{saturday_14:41パブ}} または {{saturday_1441パブ}} - 14:41パブリシティ');
    console.log('{{saturday_時間指定なしパブ}} または {{saturday_パブ}} - その他パブ');
    console.log('');
    console.log('◆ 【Route 847】専用プレースホルダー:');
    console.log('{{saturday_リポート 16:47}} または {{saturday_1647リポート}} - 16:47リポート');
    console.log('{{saturday_営業パブ 17:41}} または {{saturday_1741パブ}} - 17:41営業パブ');
    console.log('');
    console.log('◆ 利用可能な曜日名:');
    console.log('monday, tuesday, wednesday, thursday, friday, saturday, sunday');
    console.log('');
    console.log('=== 実際のテンプレート例 ===');
    console.log('');
    console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    console.log('【{{番組名}}】週間スケジュール');
    console.log('生成日時: {{生成日時}}');
    console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    console.log('');
    console.log('■ 収録予定');
    console.log('・ちょうどいい暮らし: {{暮らし収録}}');
    console.log('・不動産コーナー: {{不動産収録}}');
    console.log('・歯ッピーコーナー: {{歯ッピー収録}}');
    console.log('・おカネの話: {{おカネ収録}}');
    console.log('・ごりごり隊: {{ごりごり収録}}');
    console.log('・ビジネスアイ: {{ビジネスアイ収録}}');
    console.log('');
    console.log('【月曜日 - {{monday_日付}}】');
    console.log('┌─────────────────────────────────────────────────────┐');
    console.log('│ 7:28 パブリシティ                              │');
    console.log('└─────────────────────────────────────────────────────┘');
    console.log('{{monday_728パブ}}');
    console.log('');
    console.log('┌─────────────────────────────────────────────────────┐');
    console.log('│ 楽曲リスト                                     │');
    console.log('└─────────────────────────────────────────────────────┘');
    console.log('{{monday_楽曲_テーブル}}');
    console.log('');
    console.log('┌─────────────────────────────────────────────────────┐');
    console.log('│ YOKOHAMA PORTSIDE INFORMATION                  │');
    console.log('└─────────────────────────────────────────────────────┘');
    console.log('{{monday_ポートサイド}}');
    console.log('');
    console.log('┌─────────────────────────────────────────────────────┐');
    console.log('│ 先行予約                                       │');
    console.log('└─────────────────────────────────────────────────────┘');
    console.log('{{monday_予約}}');
    console.log('');
    console.log('┌─────────────────────────────────────────────────────┐');
    console.log('│ 放送後                                         │');
    console.log('└─────────────────────────────────────────────────────┘');
    console.log('{{monday_放送後}}');
    console.log('');
    console.log('【火曜日 - {{tuesday_日付}}】');
    console.log('┌─────────────────────────────────────────────────────┐');
    console.log('│ ヨコアリくん                                   │');
    console.log('└─────────────────────────────────────────────────────┘');
    console.log('{{tuesday_ヨコアリ}}');
    console.log('');
    console.log('（以下同様に各曜日を設定）');
    console.log('');
    console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
}
/*
=== 【NEW!】ちょうどいいラジオ新機能 ===

1. ヨコアリくん機能
- 火曜日のみ対象
- CALENDAR_IDのカレンダーで「横浜アリーナスポットオンエア」があれば「あり」、なければ「なし」
- 火曜日以外は「ー」で表示

2. 放送後機能
- その日に収録予定がある場合、「【収録】ここが知りたい不動産」形式で表示
- 収録予定カレンダーと連動

3. 先行予約強化機能
- メインスプレッドシートのA列から自動取得
- A列の日付以外の文字列を抽出
- 曜日ごとに分類して先行予約情報に追加

=== YOKOHAMA PORTSIDE INFORMATION カレンダー連携機能 ===

【概要】
ちょうどいいラジオのYOKOHAMA PORTSIDE INFORMATIONを専用のGoogleカレンダーから自動取得します。

【設定方法】
1. config.gsに以下を追加：
   PORTSIDE_CALENDAR_ID: 'YOUR_PORTSIDE_CALENDAR_ID_HERE'

2. PORTSIDE専用カレンダーの準備：
   - YOKOHAMA PORTSIDE INFORMATION関連の情報のみを掲載
   - イベントタイトルがそのまま番組で使用されます
   - 日付ごとに必要な情報をカレンダーに登録

【動作仕様】
- 起点となる月曜日から1週間分のイベントを検索
- 各曜日のイベントタイトルを取得
- 番組表の既存情報と併用（カレンダー情報には専用フォーマット適用）
- カレンダーにアクセスできない場合は番組表の情報のみ使用

【テスト機能】
- testPortsideInformation() - 今週のPORTSIDE情報をテスト取得
- testAdvanceBooking() - 先行予約情報をテスト取得

【出力例】
ヨコアリくん：あり（火曜日のみ）
放送後：【収録】ここが知りたい不動産
PORTSIDE情報：YOKOHAMA PORTSIDE INFORMATION [横浜赤レンガ倉庫イベント情報]

=== Googleドキュメント作成機能の使い方 ===

1. config.gsの設定例：
const CONFIG = {
  SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID_HERE',
  CALENDAR_ID: 'YOUR_CALENDAR_ID_HERE',
  MUSIC_SPREADSHEET_ID: 'YOUR_MUSIC_SPREADSHEET_ID_HERE',
  MUSIC_SHEET_NAME: 'シート1',
  EMAIL_ADDRESS: 'your-email@example.com',
  
  // Googleドキュメントテンプレート設定
  DOCUMENT_TEMPLATES: {
    'ちょうどいいラジオ': '1BxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxBc',
    'PRIME TIME': '1DxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxEf',
    'FLAG': '1GxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxHi',
    'God Bless Saturday': '1JxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxKl',
    'Route 847': '1MxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxNo'
  },
  DOCUMENT_OUTPUT_FOLDER_ID: '1PxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxQr',
  
  // PORTSIDE専用カレンダー
  PORTSIDE_CALENDAR_ID: 'YOUR_PORTSIDE_CALENDAR_ID_HERE'
};

2. 新しいプレースホルダー：
- {{tuesday_ヨコアリくん}} または {{tuesday_ヨコアリ}} - ヨコアリくん（火曜のみ）
- {{monday_放送後}} - 放送後情報
- {{monday_予約}} - 先行予約（スプレッドシート連携）

3. 利用可能な関数：
- extractThisWeekAndSendEmailAndCreateDocs() - 今週分メール＋ドキュメント
- extractNextWeekByProgramAndSendEmailAndCreateDocs() - 来週分メール＋ドキュメント
- extractSpecificWeekByProgramAndSendEmailAndCreateDocs('25.6.02-6.08') - 指定週メール＋ドキュメント
- createDocumentsOnly('25.6.02-6.08') - ドキュメントのみ作成
- createSampleTemplate() - テンプレート作成方法を表示
- testPortsideInformation() - PORTSIDE情報のテスト取得
- testAdvanceBooking() - 先行予約情報のテスト取得
- debugAdvanceBooking() - 先行予約のデバッグ情報表示

4. 出力：
- 指定フォルダに番組別のGoogleドキュメントが作成されます
- ファイル名形式：
  - 月～木放送番組: 【連絡票】番組名_yyyymmdd週（例：【連絡票】ちょうどいいラジオ_20250526週）
  - 週1日放送番組: 【連絡票】番組名_yyyymmdd（例：【連絡票】FLAG_20250530）
  - スペースは自動削除されます（例：God Bless Saturday → GodBlessSaturday）
- 作成されたドキュメントのリンクがメールで通知されます

=== デバッグ・テスト機能 ===

1. 設定確認：
- testConfig() - config.gsの設定内容を表示

2. 先行予約デバッグ：
- debugAdvanceBooking() - A列の全データと抽出結果を表示
- debugAdvanceBookingDetailed() - A列の詳細分析（判定条件の確認）
- testAdvanceBooking() - 先行予約情報の抽出テスト（標準判定）
- testAdvanceBookingRelaxed() - 先行予約情報の抽出テスト（緩い判定）

3. PORTSIDE情報テスト：
- testPortsideInformation() - PORTSIDE専用カレンダーからの情報取得テスト

4. 基本的な抽出テスト：
- extractThisWeekByProgram() - 今週のデータ抽出（番組別表示）
- extractNextWeekByProgram() - 来週のデータ抽出（番組別表示）

=== 先行予約の判定条件について ===

【標準判定】（getAdvanceBookingFromCurrentSheet）：
- 曜日が設定されていて3文字以上のテキストなら全て採用
- キーワード判定は不要（シンプル判定）

【緩い判定】（getAdvanceBookingFromCurrentSheetRelaxed）：
- 標準判定と同じ（後方互換性のため残存）

【使い分け】：
1. debugAdvanceBookingDetailed() で詳細分析を確認
2. 基本的に標準判定で十分（3文字以上なら全て採用）
3. さらに細かい調整が必要な場合のみ手動でフィルタリング

=== トラブルシューティング ===

1. 先行予約が表示されない場合：
- debugAdvanceBooking()でA列のデータを確認
- 日付形式が「5/26（月）」のようになっているか確認
- 先行予約情報が日付行の下にあるか確認

2. PORTSIDE情報が表示されない場合：
- testPortsideInformation()でカレンダー連携を確認
- PORTSIDE_CALENDAR_IDが正しく設定されているか確認

3. 楽曲情報が拡張されない場合：
- MUSIC_SPREADSHEET_IDとMUSIC_SHEET_NAMEが正しく設定されているか確認
- 楽曲スプレッドシートに「曲名」「アーティスト」「URL」列があるか確認

4. ドキュメント作成に失敗する場合：
- DOCUMENT_TEMPLATESとDOCUMENT_OUTPUT_FOLDER_IDが正しく設定されているか確認
- テンプレートドキュメントとフォルダにアクセス権限があるか確認

5. データ範囲抽出エラーが発生する場合：
- RS行が4個以上あるか確認（月〜木用）
- "New!Friday"、"THE BURN"、"まんてん"、"注意："の各マーカーが存在するか確認
- 備考列に「備考」というヘッダーがあるか確認
*/
/**
 * FLAG Week2 問題の詳細テスト
 */
function testFlagWeek2Problem() {
    console.log('=== FLAG Week2 問題診断テスト ===');
    try {
        console.log('\n【ステップ1】全4週のデータ抽出テスト');
        const flagResults = extractFlag4Weeks(0);
        console.log('\n【ステップ2】各週のデータ存在確認');
        const weekLabels = ['今週', '翌週', '翌々週', '翌翌々週'];
        const weekData = {};
        Object.keys(flagResults).forEach(weekKey => {
            const weekLabel = weekKey.split('(')[0];
            weekData[weekLabel] = flagResults[weekKey];
            console.log(`\n--- ${weekLabel} ---`);
            console.log(`キー: ${weekKey}`);
            console.log(`データ存在: ${!!flagResults[weekKey]}`);
            if (flagResults[weekKey] && flagResults[weekKey]['FLAG']) {
                console.log(`FLAG データ存在: ${!!flagResults[weekKey]['FLAG']}`);
                if (flagResults[weekKey]['FLAG']['friday']) {
                    const categories = Object.keys(flagResults[weekKey]['FLAG']['friday']);
                    console.log(`カテゴリ数: ${categories.length}`);
                    console.log(`カテゴリ: ${categories.join(', ')}`);
                }
                else {
                    console.log('金曜日データなし');
                }
            }
            else {
                console.log('FLAGデータなし');
            }
        });
        console.log('\n【ステップ3】Week2 詳細診断');
        const week2Key = Object.keys(flagResults).find(key => key.startsWith('翌々週'));
        if (week2Key) {
            console.log(`Week2 キー発見: ${week2Key}`);
            const week2Data = flagResults[week2Key];
            if (week2Data && week2Data['FLAG'] && week2Data['FLAG']['friday']) {
                console.log('Week2 データ構造:');
                console.log(JSON.stringify(week2Data, null, 2));
                const categories = Object.keys(week2Data['FLAG']['friday']);
                console.log(`\nWeek2 カテゴリ詳細 (${categories.length}個):`);
                categories.forEach(category => {
                    const items = week2Data['FLAG']['friday'][category];
                    console.log(`- ${category}: ${Array.isArray(items) ? items.length + '件' : typeof items}`);
                    if (Array.isArray(items) && items.length > 0) {
                        console.log(`  内容: ${items.slice(0, 3).join(', ')}${items.length > 3 ? '...' : ''}`);
                    }
                });
            }
            else {
                console.log('Week2 データ構造に問題あり');
                console.log('week2Data:', week2Data);
            }
        }
        else {
            console.log('Week2 キーが見つかりません');
            console.log('利用可能なキー:', Object.keys(flagResults));
        }
        console.log('\n【ステップ4】プレースホルダー置換テスト');
        // テスト用のドキュメントボディを模擬
        const testBody = {
            replaceText: function (placeholder, replacement) {
                console.log(`置換: ${placeholder} → ${replacement || 'ー'}`);
                return this;
            }
        };
        console.log('プレースホルダー置換をシミュレーション:');
        replaceFlagPlaceholders(testBody, flagResults);
        console.log('\n【ステップ5】個別週抽出テスト');
        console.log('各週を個別に抽出してテスト:');
        for (let i = 0; i < 4; i++) {
            try {
                const weekResult = extractFlagWeeksLater(i);
                const weekLabel = ['今週', '翌週', '翌々週', '翌翌々週'][i];
                console.log(`${weekLabel} (offset:${i}): ${Object.keys(weekResult).length > 0 ? '成功' : '失敗'}`);
                if (i === 2 && Object.keys(weekResult).length === 0) {
                    console.log('★ Week2 (翌々週) で問題発生!');
                }
            }
            catch (error) {
                console.error(`週 ${i} の抽出でエラー:`, error);
            }
        }
        console.log('\n=== FLAG Week2 診断完了 ===');
        return {
            success: true,
            week2Found: !!week2Key,
            week2HasData: !!(week2Key && flagResults[week2Key] && flagResults[week2Key]['FLAG']),
            allWeeksData: weekData
        };
    }
    catch (error) {
        console.error('FLAG Week2 テストでエラー:', error);
        return {
            success: false,
            error: error instanceof Error ? error.message : String(error)
        };
    }
}
/**
 * Week2 シート名計算の詳細テスト
 */
function testWeek2SheetCalculation() {
    console.log('=== Week2 シート名計算テスト ===');
    const today = new Date();
    console.log('今日の日付:', today.toLocaleDateString('ja-JP'));
    for (let i = 0; i < 4; i++) {
        const targetDate = new Date(today.getTime() + (i * 7 * 24 * 60 * 60 * 1000));
        const dayOfWeek = targetDate.getDay() === 0 ? 7 : targetDate.getDay();
        const monday = new Date(targetDate.getTime() - (dayOfWeek - 1) * 24 * 60 * 60 * 1000);
        const sunday = new Date(monday.getTime() + 6 * 24 * 60 * 60 * 1000);
        const mondayYear = monday.getFullYear().toString().slice(-2);
        const mondayMonth = (monday.getMonth() + 1).toString();
        const mondayDay = monday.getDate().toString().padStart(2, '0');
        const sundayMonth = (sunday.getMonth() + 1).toString();
        const sundayDay = sunday.getDate().toString().padStart(2, '0');
        const sheetName = `${mondayYear}.${mondayMonth}.${mondayDay}-${sundayMonth}.${sundayDay}`;
        const weekLabel = ['今週', '翌週', '翌々週', '翌翌々週'][i];
        console.log(`${weekLabel} (offset:${i}):`);
        console.log(`  対象日: ${targetDate.toLocaleDateString('ja-JP')}`);
        console.log(`  月曜日: ${monday.toLocaleDateString('ja-JP')}`);
        console.log(`  日曜日: ${sunday.toLocaleDateString('ja-JP')}`);
        console.log(`  シート名: ${sheetName}`);
        if (i === 2) {
            console.log('★ これがWeek2のシート名です');
        }
    }
    console.log('\n=== シート名計算テスト完了 ===');
}
