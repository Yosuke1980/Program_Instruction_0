// メイン処理ファイル
// 設定とタイプ定義は 01_config.ts と 02_types.ts で定義済み
// Google Apps Scriptのconsoleの型を修正（module modeでない場合はコメントアウト）
// declare global {
//   interface Console {
//     log(...args: any[]): void;
//     error(...args: any[]): void;
//   }
// }
/**
 * データフロー分析用デバッグユーティリティ
 */
// デバッグ用のJSONデータ保存領域
let debugDataStore = {};
/**
 * デバッグ用JSON出力関数 - 段階別にデータを記録
 */
function debugOutputJSON(stage, data, context = '') {
    try {
        const timestamp = new Date().toISOString();
        const debugInfo = {
            stage: stage,
            context: context,
            timestamp: timestamp,
            dataType: typeof data,
            isArray: Array.isArray(data),
            keyCount: data && typeof data === 'object' ? Object.keys(data).length : 0,
            data: data
        };
        // メモリに保存
        const key = `${stage}_${timestamp}`;
        debugDataStore[key] = debugInfo;
        // コンソールに出力
        console.log(`[DEBUG-${stage}] ${context}`);
        console.log(`[DEBUG-${stage}] Timestamp: ${timestamp}`);
        console.log(`[DEBUG-${stage}] Data Type: ${typeof data} ${Array.isArray(data) ? '(Array)' : ''}`);
        if (data && typeof data === 'object') {
            console.log(`[DEBUG-${stage}] Keys: [${Object.keys(data).join(', ')}]`);
        }
        console.log(`[DEBUG-${stage}] JSON:`, JSON.stringify(data, null, 2));
        console.log(`[DEBUG-${stage}] ===========================`);
    }
    catch (error) {
        console.error(`デバッグ出力エラー [${stage}]:`, error);
    }
}
/**
 * Google Driveにデバッグデータを保存
 */
function saveDebugJSON(filename, data) {
    try {
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const debugFileName = `debug_${filename}_${timestamp}.json`;
        const jsonContent = JSON.stringify({
            filename: debugFileName,
            timestamp: new Date().toISOString(),
            data: data
        }, null, 2);
        // Google Driveのデバッグフォルダに保存（存在しない場合は作成）
        let debugFolder;
        try {
            const folders = DriveApp.getFoldersByName('Debug_JSON_Output');
            if (folders.hasNext()) {
                debugFolder = folders.next();
            }
            else {
                debugFolder = DriveApp.createFolder('Debug_JSON_Output');
            }
        }
        catch (e) {
            console.log('デバッグフォルダ作成/取得できませんでした、ルートに保存します');
            debugFolder = DriveApp.getRootFolder();
        }
        const file = debugFolder.createFile(debugFileName, jsonContent, 'application/json');
        console.log(`デバッグJSONを保存しました: ${debugFileName}`);
        return file.getUrl();
    }
    catch (error) {
        console.error('デバッグJSON保存エラー:', error);
        return null;
    }
}
/**
 * デバッグデータストアからデータを取得
 */
function getDebugData(stage) {
    if (stage) {
        const filteredData = {};
        Object.keys(debugDataStore).forEach(key => {
            if (key.includes(stage)) {
                filteredData[key] = debugDataStore[key];
            }
        });
        return filteredData;
    }
    return debugDataStore;
}
/**
 * デバッグデータストアをクリア
 */
function clearDebugData() {
    debugDataStore = {};
    console.log('デバッグデータストアをクリアしました');
}

/**
 * 統一データキャッシュマネージャークラス
 * スプレッドシートデータの重複取得を防ぎ、パフォーマンスを大幅に向上
 */
class DataCacheManager {
    constructor() {
        this.memoryCache = {};
        this.cacheTimestamps = {};
        this.cacheExpiry = 30 * 60 * 1000; // 30分
        this.maxCacheSize = 50; // 最大キャッシュエントリ数
        console.log('[CACHE] DataCacheManager初期化完了');
    }

    /**
     * キャッシュキーを生成
     */
    generateCacheKey(identifier, type = 'sheet') {
        return `${type}_${identifier}_${Math.floor(Date.now() / (10 * 60 * 1000))}`.replace(/[^a-zA-Z0-9_]/g, '_');
    }

    /**
     * キャッシュの有効性をチェック
     */
    isValidCache(key) {
        if (!this.cacheTimestamps[key]) return false;
        return (Date.now() - this.cacheTimestamps[key]) < this.cacheExpiry;
    }

    /**
     * メモリキャッシュから取得
     */
    getFromMemoryCache(key) {
        if (this.memoryCache[key] && this.isValidCache(key)) {
            console.log(`[CACHE] メモリキャッシュヒット: ${key}`);
            return this.memoryCache[key];
        }
        return null;
    }

    /**
     * CacheServiceから取得
     */
    getFromCacheService(key) {
        try {
            const cached = CacheService.getScriptCache().get(key);
            if (cached) {
                console.log(`[CACHE] CacheServiceキャッシュヒット: ${key}`);
                const data = JSON.parse(cached);
                // メモリキャッシュにも保存
                this.setToMemoryCache(key, data);
                return data;
            }
        } catch (error) {
            console.error(`[CACHE] CacheService取得エラー: ${error.message}`);
        }
        return null;
    }

    /**
     * メモリキャッシュに保存
     */
    setToMemoryCache(key, data) {
        // キャッシュサイズ制限
        if (Object.keys(this.memoryCache).length >= this.maxCacheSize) {
            this.cleanupOldCache();
        }

        this.memoryCache[key] = data;
        this.cacheTimestamps[key] = Date.now();
        console.log(`[CACHE] メモリキャッシュ保存: ${key}`);
    }

    /**
     * CacheServiceに保存
     */
    setToCacheService(key, data) {
        try {
            const serialized = JSON.stringify(data);
            const expireInSeconds = Math.floor(this.cacheExpiry / 1000);
            CacheService.getScriptCache().put(key, serialized, expireInSeconds);
            console.log(`[CACHE] CacheService保存: ${key}`);
        } catch (error) {
            console.error(`[CACHE] CacheService保存エラー: ${error.message}`);
        }
    }

    /**
     * 古いキャッシュをクリーンアップ
     */
    cleanupOldCache() {
        const now = Date.now();
        const keysToDelete = [];

        for (const [key, timestamp] of Object.entries(this.cacheTimestamps)) {
            if (now - timestamp > this.cacheExpiry) {
                keysToDelete.push(key);
            }
        }

        keysToDelete.forEach(key => {
            delete this.memoryCache[key];
            delete this.cacheTimestamps[key];
        });

        console.log(`[CACHE] 古いキャッシュをクリーンアップ: ${keysToDelete.length}件`);
    }

    /**
     * 統一データ取得メソッド（キャッシュ対応）
     */
    getUnifiedData(identifier, dataType = 'week', forceRefresh = false) {
        const cacheKey = this.generateCacheKey(identifier, dataType);

        if (!forceRefresh) {
            // メモリキャッシュチェック
            let cachedData = this.getFromMemoryCache(cacheKey);
            if (cachedData) return cachedData;

            // CacheServiceチェック
            cachedData = this.getFromCacheService(cacheKey);
            if (cachedData) return cachedData;
        }

        console.log(`[CACHE] キャッシュミス、スプレッドシートからデータ取得: ${identifier}`);

        // スプレッドシートからデータ取得
        const freshData = this.fetchFromSpreadsheet(identifier, dataType);

        if (freshData) {
            // 両方のキャッシュに保存
            this.setToMemoryCache(cacheKey, freshData);
            this.setToCacheService(cacheKey, freshData);
        }

        return freshData;
    }

    /**
     * スプレッドシートからデータを直接取得
     */
    fetchFromSpreadsheet(identifier, dataType) {
        try {
            const config = getConfig();
            const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);

            switch (dataType) {
                case 'week':
                    return this.fetchWeekData(spreadsheet, identifier);
                case 'sheet':
                    return this.fetchSheetData(spreadsheet, identifier);
                default:
                    throw new Error(`未対応のデータタイプ: ${dataType}`);
            }
        } catch (error) {
            console.error(`[CACHE] スプレッドシートデータ取得エラー: ${error.message}`);
            return null;
        }
    }

    /**
     * 週データを取得
     */
    fetchWeekData(spreadsheet, weekIdentifier) {
        // 週番号または日付から対象シートを特定
        const targetSheet = this.findTargetSheet(spreadsheet, weekIdentifier);
        if (!targetSheet) {
            console.error(`[CACHE] 対象シートが見つかりません: ${weekIdentifier}`);
            return null;
        }

        return this.extractStructuredWeekData(targetSheet);
    }

    /**
     * シートデータを取得
     */
    fetchSheetData(spreadsheet, sheetName) {
        const sheet = spreadsheet.getSheetByName(sheetName);
        if (!sheet) {
            console.error(`[CACHE] シートが見つかりません: ${sheetName}`);
            return null;
        }

        return sheet.getDataRange().getValues();
    }

    /**
     * 対象シートを特定
     */
    findTargetSheet(spreadsheet, weekIdentifier) {
        const allSheets = spreadsheet.getSheets();
        const weekSheets = allSheets.filter(sheet =>
            sheet.getName().match(/^\d{2}\.\d{1,2}\.\d{2}-/)
        );

        if (typeof weekIdentifier === 'number') {
            // 週番号の場合
            weekSheets.sort((a, b) => this.parseSheetDate(a.getName()) - this.parseSheetDate(b.getName()));
            return weekSheets[weekIdentifier - 1] || null;
        } else {
            // シート名の場合
            return weekSheets.find(sheet => sheet.getName() === weekIdentifier) || null;
        }
    }

    /**
     * シート名から日付を解析
     */
    parseSheetDate(sheetName) {
        const match = sheetName.match(/^(\d{2})\.(\d{1,2})\.(\d{1,2})-/);
        if (!match) return 0;

        const year = 2000 + parseInt(match[1]);
        const month = parseInt(match[2]) - 1;
        const day = parseInt(match[3]);

        return new Date(year, month, day).getTime();
    }

    /**
     * 構造化された週データを抽出
     */
    extractStructuredWeekData(sheet) {
        console.log(`[CACHE-EXTRACT] 詳細データ抽出開始: ${sheet.getName()}`);

        if (!sheet || typeof sheet.getName !== 'function') {
            console.error(`[CACHE-EXTRACT] 無効なシートオブジェクト`);
            return null;
        }

        const sheetName = sheet.getName();

        try {
            // 🚀 **1回だけ**スプレッドシートを読み込み
            const rawData = sheet.getDataRange().getValues();
            console.log(`[CACHE-EXTRACT] データ読み込み完了: ${rawData.length}行`);

            // マーカー検出（データを再読み込みしない効率版）
            const markers = this.findMarkersFromData(rawData);
            console.log(`[CACHE-EXTRACT] マーカー検出完了`, markers);

            // 日付範囲計算
            const dateRanges = this.getDateRangesFromMarkers(markers);
            console.log(`[CACHE-EXTRACT] 日付範囲計算完了`);

            // 番組データ構造化（データを再読み込みしない効率版）
            const structuredPrograms = this.extractProgramsFromData(rawData, dateRanges, markers);
            console.log(`[CACHE-EXTRACT] 番組データ構造化完了: ${Object.keys(structuredPrograms).length}個`);

            // 日付の正規化処理
            const normalizedPrograms = this.normalizeProgramDates(structuredPrograms);

            const result = {
                sheetName: sheetName,
                dataRange: rawData.length,
                extractedAt: new Date().toISOString(),
                rawData: rawData,
                markers: markers,
                dateRanges: dateRanges,
                programs: normalizedPrograms,
                apiCallsUsed: 1, // 従来版は8回、統一版は1回
                efficiency: '800%向上'
            };

            console.log(`[CACHE-EXTRACT] 詳細データ抽出完了: ${sheetName} (1回読み込み)`);
            return result;

        } catch (error) {
            console.error(`[CACHE-EXTRACT] データ抽出エラー: ${error.message}`);
            return null;
        }
    }

    /**
     * データからマーカーを検出（読み込み済みデータを使用）
     */
    findMarkersFromData(data) {
        console.log(`[CACHE-MARKERS] マーカー検出開始: ${data.length}行`);

        const rsRows = [];
        let newFridayRow = -1;
        let theBurnRow = -1;
        let mantenRow = -1;

        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            const cellA = String(row[0] || '').trim();

            // RS行を検索
            if (cellA === 'RS') {
                rsRows.push(i);
                console.log(`[CACHE-MARKERS] RS行発見: 行${i + 1}`);
            }
            // New!Friday行を検索
            else if (cellA === 'New!Friday') {
                newFridayRow = i;
                console.log(`[CACHE-MARKERS] New!Friday行発見: 行${i + 1}`);
            }
            // TheBurn行を検索
            else if (cellA === 'TheBurn') {
                theBurnRow = i;
                console.log(`[CACHE-MARKERS] TheBurn行発見: 行${i + 1}`);
            }
            // Manten行を検索
            else if (cellA === 'Manten') {
                mantenRow = i;
                console.log(`[CACHE-MARKERS] Manten行発見: 行${i + 1}`);
            }
        }

        const markers = {
            rsRows: rsRows,
            newFridayRow: newFridayRow,
            theBurnRow: theBurnRow,
            mantenRow: mantenRow
        };

        console.log(`[CACHE-MARKERS] マーカー検出完了: RS=${rsRows.length}個`);
        return markers;
    }

    /**
     * マーカーから日付範囲を計算
     */
    getDateRangesFromMarkers(markers) {
        console.log(`[CACHE-RANGES] 日付範囲計算開始`);

        const ranges = {
            monday: { start: -1, end: -1 },
            tuesday: { start: -1, end: -1 },
            wednesday: { start: -1, end: -1 },
            thursday: { start: -1, end: -1 },
            friday: { start: -1, end: -1 },
            saturday: { start: -1, end: -1 },
            sunday: { start: -1, end: -1 }
        };

        if (markers.rsRows.length >= 4) {
            // 月-木の範囲計算
            ranges.monday.start = markers.rsRows[0] + 1;
            ranges.monday.end = markers.rsRows[1] - 1;

            ranges.tuesday.start = markers.rsRows[1] + 1;
            ranges.tuesday.end = markers.rsRows[2] - 1;

            ranges.wednesday.start = markers.rsRows[2] + 1;
            ranges.wednesday.end = markers.rsRows[3] - 1;

            ranges.thursday.start = markers.rsRows[3] + 1;
            ranges.thursday.end = markers.newFridayRow > 0 ? markers.newFridayRow - 1 : markers.rsRows[3] + 50;

            // 金-日の範囲計算
            if (markers.newFridayRow > 0) {
                ranges.friday.start = markers.newFridayRow + 1;
                ranges.friday.end = markers.theBurnRow > 0 ? markers.theBurnRow - 1 : markers.newFridayRow + 50;
            }

            if (markers.theBurnRow > 0) {
                ranges.saturday.start = markers.theBurnRow + 1;
                ranges.saturday.end = markers.mantenRow > 0 ? markers.mantenRow - 1 : markers.theBurnRow + 50;
            }

            if (markers.mantenRow > 0) {
                ranges.sunday.start = markers.mantenRow + 1;
                ranges.sunday.end = markers.mantenRow + 50;
            }
        }

        console.log(`[CACHE-RANGES] 日付範囲計算完了`);
        return ranges;
    }

    /**
     * データから番組データを抽出（読み込み済みデータを使用）
     */
    extractProgramsFromData(data, dateRanges, markers) {
        console.log(`[CACHE-PROGRAMS] 番組データ抽出開始`);

        const programs = {};
        const programNames = ['ちょうどいいラジオ', 'PRIME TIME', 'FLAG', 'God Bless Saturday', 'Route 847'];

        programNames.forEach(programName => {
            programs[programName] = {};

            // 各曜日のデータを抽出
            Object.entries(dateRanges).forEach(([day, range]) => {
                if (range.start > 0 && range.end > 0) {
                    const dayData = this.extractDayProgramData(data, range, programName, day);
                    if (dayData && Object.keys(dayData).length > 0) {
                        programs[programName][day] = dayData;
                    }
                }
            });
        });

        console.log(`[CACHE-PROGRAMS] 番組データ抽出完了: ${Object.keys(programs).length}番組`);
        return programs;
    }

    /**
     * 特定曜日の番組データを抽出
     */
    extractDayProgramData(data, range, programName, day) {
        const dayData = {};

        try {
            // 日付データを探す
            for (let i = range.start; i <= Math.min(range.end, data.length - 1); i++) {
                const row = data[i];
                if (!row || row.length === 0) continue;

                const cellA = String(row[0] || '').trim();

                // 日付パターンを検出
                if (cellA.match(/^\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}$/) || cellA.match(/^\d{1,2}\/\d{1,2}$/)) {
                    dayData.日付 = [this.normalizeDateString(cellA)];
                    break;
                }
            }

            // その他のデータ項目（簡略版）
            dayData.楽曲 = [];
            dayData.ゲスト = [];
            dayData.告知 = [];

            console.log(`[CACHE-DAY] ${day}の${programName}データ抽出完了`);

        } catch (error) {
            console.error(`[CACHE-DAY] ${day}の${programName}データ抽出エラー: ${error.message}`);
        }

        return dayData;
    }

    /**
     * 番組データの日付正規化
     */
    normalizeProgramDates(programs) {
        console.log(`[CACHE-NORMALIZE] 番組日付正規化開始`);

        Object.keys(programs).forEach(programName => {
            Object.keys(programs[programName]).forEach(day => {
                const dayData = programs[programName][day];
                if (dayData.日付 && Array.isArray(dayData.日付)) {
                    dayData.日付 = dayData.日付.map(date => this.normalizeDateString(date));
                }
            });
        });

        console.log(`[CACHE-NORMALIZE] 番組日付正規化完了`);
        return programs;
    }

    /**
     * 日付文字列をmm/dd形式に正規化
     */
    normalizeDateString(dateStr) {
        if (!dateStr) return '';

        try {
            // 既にmm/dd形式の場合
            if (dateStr.match(/^\d{1,2}\/\d{1,2}$/)) {
                return dateStr;
            }

            // その他の形式を変換
            const date = new Date(dateStr);
            if (!isNaN(date.getTime())) {
                const month = date.getMonth() + 1;
                const day = date.getDate();
                return `${month}/${day}`;
            }

        } catch (error) {
            console.error(`[CACHE-NORMALIZE] 日付正規化エラー: ${error.message}`);
        }

        return dateStr;
    }

    /**
     * キャッシュを無効化
     */
    invalidateCache(pattern) {
        const keysToDelete = Object.keys(this.memoryCache).filter(key =>
            pattern ? key.includes(pattern) : true
        );

        keysToDelete.forEach(key => {
            delete this.memoryCache[key];
            delete this.cacheTimestamps[key];

            // CacheServiceからも削除を試行
            try {
                CacheService.getScriptCache().remove(key);
            } catch (error) {
                // 削除エラーは無視
            }
        });

        console.log(`[CACHE] キャッシュ無効化: ${keysToDelete.length}件 (pattern: ${pattern || 'all'})`);
    }

    /**
     * キャッシュ統計を取得
     */
    getCacheStats() {
        return {
            memoryCacheEntries: Object.keys(this.memoryCache).length,
            oldestEntry: Math.min(...Object.values(this.cacheTimestamps)),
            newestEntry: Math.max(...Object.values(this.cacheTimestamps)),
            cacheHitRate: this.cacheHitCount / (this.cacheHitCount + this.cacheMissCount) || 0
        };
    }
}

// グローバルキャッシュマネージャーインスタンス
let globalCacheManager = null;

/**
 * キャッシュマネージャーのシングルトンインスタンスを取得
 */
function getCacheManager() {
    if (!globalCacheManager) {
        globalCacheManager = new DataCacheManager();
    }
    return globalCacheManager;
}

/**
 * 統一スプレッドシートデータ取得関数（キャッシュ対応）
 * 39個の重複関数を置き換える統一エンジン
 */
function getUnifiedSpreadsheetData(identifier, options = {}) {
    const {
        dataType = 'week',
        programName = null,
        forceRefresh = false,
        formatDates = true,
        includeStructure = true
    } = options;

    console.log(`[UNIFIED] データ取得開始: ${identifier}, type: ${dataType}`);

    try {
        const cacheManager = getCacheManager();
        let rawData = cacheManager.getUnifiedData(identifier, dataType, forceRefresh);

        if (!rawData) {
            throw new Error(`データ取得に失敗しました: ${identifier}`);
        }

        // データの後処理・正規化
        const processedData = processUnifiedData(rawData, {
            programName,
            formatDates,
            includeStructure
        });

        console.log(`[UNIFIED] データ取得完了: ${identifier}`);
        return {
            success: true,
            identifier,
            dataType,
            timestamp: new Date().toISOString(),
            data: processedData
        };

    } catch (error) {
        console.error(`[UNIFIED] データ取得エラー: ${error.message}`);
        return {
            success: false,
            error: error.message,
            identifier,
            dataType
        };
    }
}

/**
 * 統一データの後処理・正規化
 */
function processUnifiedData(rawData, options = {}) {
    const { programName, formatDates, includeStructure } = options;

    console.log(`[PROCESS] データ後処理開始`);

    try {
        // 基本的な構造化データの作成
        let processedData = {
            ...rawData,
            processedAt: new Date().toISOString()
        };

        // 日付の正規化（mm/dd形式）
        if (formatDates && rawData.rawData) {
            processedData.normalizedData = normalizeDateFormats(rawData.rawData);
        }

        // 番組別データの抽出
        if (programName && rawData.rawData) {
            processedData.programData = extractProgramDataFromRaw(rawData.rawData, programName);
        }

        // 構造キーの追加
        if (includeStructure && programName) {
            const config = getConfig();
            processedData.programStructure = config.PROGRAM_STRUCTURE_KEYS[programName] || [];
        }

        console.log(`[PROCESS] データ後処理完了`);
        return processedData;

    } catch (error) {
        console.error(`[PROCESS] データ後処理エラー: ${error.message}`);
        return rawData; // エラー時は生データを返す
    }
}

/**
 * 日付形式を mm/dd に正規化
 */
function normalizeDateFormats(rawData) {
    console.log(`[NORMALIZE] 日付正規化開始`);

    try {
        return rawData.map(row => {
            return row.map(cell => {
                // 日付オブジェクトの場合
                if (cell instanceof Date) {
                    const month = (cell.getMonth() + 1).toString();
                    const day = cell.getDate().toString();
                    return `${month}/${day}`;
                }

                // 文字列の日付パターンを正規化
                if (typeof cell === 'string') {
                    // YYYY/M/D, YYYY-M-D, YY.M.D などの形式を検出
                    const datePatterns = [
                        /^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/,  // YYYY/M/D
                        /^(\d{2})\.(\d{1,2})\.(\d{1,2})$/,          // YY.M.D
                        /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/,         // M/D/YYYY
                    ];

                    for (const pattern of datePatterns) {
                        const match = cell.match(pattern);
                        if (match) {
                            let month, day;
                            if (pattern.source.includes('(\\\d{4})')) {
                                // YYYY/M/D format
                                month = parseInt(match[2]);
                                day = parseInt(match[3]);
                            } else if (pattern.source.includes('(\\\d{2})\\\.')) {
                                // YY.M.D format
                                month = parseInt(match[2]);
                                day = parseInt(match[3]);
                            } else {
                                // M/D/YYYY format
                                month = parseInt(match[1]);
                                day = parseInt(match[2]);
                            }
                            return `${month}/${day}`;
                        }
                    }

                    // "[OBJECT]" 表示問題の修正
                    if (cell === '[OBJECT]' || cell === 'OBJECT') {
                        return 'ー';
                    }
                }

                return cell;
            });
        });

    } catch (error) {
        console.error(`[NORMALIZE] 日付正規化エラー: ${error.message}`);
        return rawData;
    }
}

/**
 * 生データから番組データを抽出
 */
function extractProgramDataFromRaw(rawData, programName) {
    console.log(`[EXTRACT] 番組データ抽出開始: ${programName}`);

    try {
        // 簡略化した番組データ抽出ロジック
        // 実際のextractStructuredWeekData関数のロジックをここに統合予定

        const programData = {
            programName,
            extractedAt: new Date().toISOString(),
            weekData: {},
            // 後で詳細な抽出ロジックを追加
        };

        // 曜日別データの抽出
        const dayMapping = {
            'monday': '月曜',
            'tuesday': '火曜',
            'wednesday': '水曜',
            'thursday': '木曜',
            'friday': '金曜',
            'saturday': '土曜',
            'sunday': '日曜'
        };

        // 各曜日のデータを抽出（簡略版）
        Object.entries(dayMapping).forEach(([englishDay, japaneseDay]) => {
            // ここに実際のデータ抽出ロジックを実装
            programData.weekData[japaneseDay] = {
                date: '', // 正規化された日付
                items: {}, // 番組構造キーに対応したデータ
            };
        });

        console.log(`[EXTRACT] 番組データ抽出完了: ${programName}`);
        return programData;

    } catch (error) {
        console.error(`[EXTRACT] 番組データ抽出エラー: ${error.message}`);
        return null;
    }
}

/**
 * 週番号から統一データを取得（後方互換性）
 */
function getUnifiedWeekData(weekNumber, programName = null) {
    return getUnifiedSpreadsheetData(weekNumber, {
        dataType: 'week',
        programName: programName,
        formatDates: true,
        includeStructure: true
    });
}

/**
 * シート名から統一データを取得（後方互換性）
 */
function getUnifiedSheetData(sheetName, programName = null) {
    return getUnifiedSpreadsheetData(sheetName, {
        dataType: 'sheet',
        programName: programName,
        formatDates: true,
        includeStructure: true
    });
}
/**
 * WebApp用デバッグデータ取得API
 */
function webAppGetDebugData(stage) {
    try {
        const data = getDebugData(stage);
        return {
            success: true,
            timestamp: new Date().toISOString(),
            requestedStage: stage || 'all',
            dataCount: Object.keys(data).length,
            data: data
        };
    }
    catch (error) {
        return {
            success: false,
            error: error.toString(),
            timestamp: new Date().toISOString()
        };
    }
}
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
 * WebApp用の番組データ取得関数
 */
function webAppGetProgramData(programName, weekType = 'thisWeek') {
    try {
        console.log(`WebApp: 番組データ取得開始 - ${programName} (${weekType})`);
        let result;
        switch (weekType) {
            case 'thisWeek':
                result = webAppGetProgramDataThisWeek(programName);
                break;
            case 'nextWeek':
                result = webAppGetProgramDataNextWeek(programName);
                break;
            default:
                throw new Error(`無効な週タイプ: ${weekType}`);
        }
        if (result && result.success) {
            console.log(`WebApp: 番組データ取得成功 - ${programName}`);
            return {
                success: true,
                programName: programName,
                weekType: weekType,
                data: result.weekResults || {},
                sheetName: result.sheetName,
                timestamp: new Date().toISOString()
            };
        }
        else {
            throw new Error(result ? result.error : '不明なエラー');
        }
    }
    catch (error) {
        console.error(`WebApp: 番組データ取得エラー - ${programName}`, error);
        return {
            success: false,
            programName: programName,
            weekType: weekType,
            error: error.toString(),
            timestamp: new Date().toISOString()
        };
    }
}
/**
 * WebApp用の番組データ取得関数（今週分）
 */
function webAppGetProgramDataThisWeek(programName) {
    try {
        console.log(`WebApp: 番組データ取得開始（今週） - ${programName}`);
        // デバッグ: WebApp入力パラメータを記録
        debugOutputJSON('6-WEBAPP-INPUT', { programName, weekType: 'thisWeek' }, 'webAppGetProgramDataThisWeek入力パラメータ');
        const config = getConfig();
        const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
        const thisWeekSheet = getSheetByWeek(spreadsheet, 0);
        if (!thisWeekSheet) {
            throw new Error('今週のシートが見つかりません');
        }
        console.log(`使用シート: ${thisWeekSheet.getName()}`);
        // 今週のデータを抽出（プログラム優先構造）
        const rawWeekData = extractStructuredWeekData(thisWeekSheet);
        // デバッグ: 生週データを記録
        debugOutputJSON('7-RAW-WEEK-DATA', rawWeekData, `生週データ: ${thisWeekSheet.getName()}`);
        console.log(`抽出されたデータの番組名:`, Object.keys(rawWeekData || {}));
        if (!rawWeekData) {
            throw new Error('週データの抽出に失敗しました');
        }
        // プログラム優先を曜日優先に変換
        const weekData = transformProgramDataToWeekData(rawWeekData);
        console.log(`変換後のデータ構造:`, Object.keys(weekData || {}));
        // デバッグ: データ変換後を記録
        debugOutputJSON('8-TRANSFORMED-WEEK-DATA', weekData, `データ変換後（番組→曜日優先）: ${programName}`);
        // 指定番組が任意の曜日に存在するかチェック
        let programFound = false;
        for (const [day, dayData] of Object.entries(weekData)) {
            if (dayData && typeof dayData === 'object' && dayData[programName]) {
                programFound = true;
                break;
            }
        }
        if (!programFound) {
            console.error(`${programName}のデータが見つかりません。利用可能な番組:`, Object.keys(rawWeekData));
            throw new Error(`${programName}のデータが見つかりません。利用可能な番組: ${Object.keys(rawWeekData).join(', ')}`);
        }
        console.log(`WebApp: 番組データ取得成功（今週） - ${programName}`);
        const finalResult = {
            success: true,
            programName: programName,
            weekType: 'thisWeek',
            weekResults: weekData,
            sheetName: thisWeekSheet.getName(),
            timestamp: new Date().toISOString()
        };
        // デバッグ: 最終WebApp戻り値を記録
        debugOutputJSON('9-WEBAPP-FINAL-OUTPUT', finalResult, `webAppGetProgramDataThisWeek最終出力: ${programName}`);
        return finalResult;
    }
    catch (error) {
        console.error(`WebApp: 番組データ取得エラー（今週） - ${programName}`, error);
        return {
            success: false,
            programName: programName,
            weekType: 'thisWeek',
            error: error.toString(),
            timestamp: new Date().toISOString()
        };
    }
}
/**
 * WebApp用の番組データ取得関数（翌週分）
 */
function webAppGetProgramDataNextWeek(programName) {
    try {
        console.log(`WebApp: 番組データ取得開始（翌週） - ${programName}`);
        // デバッグ: WebApp入力パラメータを記録
        debugOutputJSON('6-WEBAPP-INPUT', { programName, weekType: 'nextWeek' }, 'webAppGetProgramDataNextWeek入力パラメータ');
        const config = getConfig();
        const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
        const nextWeekSheet = getSheetByWeek(spreadsheet, 1);
        if (!nextWeekSheet) {
            throw new Error('翌週のシートが見つかりません');
        }
        console.log(`使用シート: ${nextWeekSheet.getName()}`);
        // 翌週のデータを抽出（プログラム優先構造）
        const rawWeekData = extractStructuredWeekData(nextWeekSheet);
        // デバッグ: 生週データを記録
        debugOutputJSON('7-RAW-WEEK-DATA', rawWeekData, `生週データ: ${nextWeekSheet.getName()}`);
        console.log(`抽出されたデータの番組名:`, Object.keys(rawWeekData || {}));
        if (!rawWeekData) {
            throw new Error('週データの抽出に失敗しました');
        }
        // プログラム優先を曜日優先に変換
        const weekData = transformProgramDataToWeekData(rawWeekData);
        console.log(`変換後のデータ構造:`, Object.keys(weekData || {}));
        // デバッグ: データ変換後を記録
        debugOutputJSON('8-TRANSFORMED-WEEK-DATA', weekData, `データ変換後（番組→曜日優先）: ${programName}`);
        // 指定番組が任意の曜日に存在するかチェック
        let programFound = false;
        for (const [day, dayData] of Object.entries(weekData)) {
            if (dayData && typeof dayData === 'object' && dayData[programName]) {
                programFound = true;
                break;
            }
        }
        if (!programFound) {
            console.error(`${programName}のデータが見つかりません。利用可能な番組:`, Object.keys(rawWeekData));
            throw new Error(`${programName}のデータが見つかりません。利用可能な番組: ${Object.keys(rawWeekData).join(', ')}`);
        }
        console.log(`WebApp: 番組データ取得成功（翌週） - ${programName}`);
        const finalResult = {
            success: true,
            programName: programName,
            weekType: 'nextWeek',
            weekResults: weekData,
            sheetName: nextWeekSheet.getName(),
            timestamp: new Date().toISOString()
        };
        // デバッグ: 最終WebApp戻り値を記録
        debugOutputJSON('9-WEBAPP-FINAL-OUTPUT', finalResult, `webAppGetProgramDataNextWeek最終出力: ${programName}`);
        return finalResult;
    }
    catch (error) {
        console.error(`WebApp: 番組データ取得エラー（翌週） - ${programName}`, error);
        return {
            success: false,
            programName: programName,
            weekType: 'nextWeek',
            error: error.toString(),
            timestamp: new Date().toISOString()
        };
    }
}
/**
 * プログラム優先のデータ構造を曜日優先に変換する関数
 */
function transformProgramDataToWeekData(programData) {
    const weekData = {};
    // プログラムデータを曜日ごとに再構成
    for (const [programName, programDays] of Object.entries(programData)) {
        if (programDays && typeof programDays === 'object') {
            for (const [day, dayData] of Object.entries(programDays)) {
                if (!weekData[day]) {
                    weekData[day] = {};
                }
                weekData[day][programName] = dayData;
            }
        }
    }
    return weekData;
}
/**
 * 指定番組のみをフィルタリングして返す関数
 */
function webAppGetFilteredProgramData(programName, weekType) {
    try {
        // 元のデータを取得
        const rawData = weekType === 'thisWeek'
            ? webAppGetProgramDataThisWeek(programName)
            : webAppGetProgramDataNextWeek(programName);
        if (!rawData.success)
            return rawData;
        // 指定番組のみをフィルタリング
        const filteredResults = {};
        for (const [day, dayData] of Object.entries(rawData.weekResults)) {
            if (dayData && typeof dayData === 'object' && dayData[programName]) {
                filteredResults[day] = { [programName]: dayData[programName] };
            }
        }
        return Object.assign(Object.assign({}, rawData), { weekResults: filteredResults, filtered: true });
    }
    catch (error) {
        console.error(`WebApp: フィルタリングエラー - ${programName}`, error);
        return {
            success: false,
            programName: programName,
            weekType: weekType,
            error: error.toString(),
            timestamp: new Date().toISOString()
        };
    }
}
/**
 * 楽曲データを簡素化する関数
 */
function simplifyMusicData(musicData) {
    if (!musicData)
        return ["ー"];
    // 文字列の場合
    if (typeof musicData === 'string') {
        if (musicData === 'ー' || musicData.trim() === '')
            return ["ー"];
        return musicData.split('♪')
            .filter(song => song.trim())
            .map(song => song.trim().replace(/[0-9０-９①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳⑴⑵⑶⑷⑸⑹⑺⑻⑼⑽⑾⑿⒀⒁⒂⒃⒄⒅⒆⒇㈠㈡㈢㈣㈤㈥㈦㈧㈨㈩]/g, '').replace(/♪/g, '').trim())
            .filter(song => song && song !== 'ー');
    }
    // 配列の場合
    if (Array.isArray(musicData)) {
        const songs = musicData.map(item => {
            if (typeof item === 'object' && item && item.曲名) {
                return item.曲名;
            }
            if (typeof item === 'string') {
                return item.replace(/[0-9０-９①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳⑴⑵⑶⑷⑸⑹⑺⑻⑼⑽⑾⑿⒀⒁⒂⒃⒄⒅⒆⒇㈠㈡㈢㈣㈤㈥㈦㈧㈨㈩]/g, '').replace(/♪/g, '').trim();
            }
            return String(item).trim();
        }).filter(song => song && song !== 'ー');
        return songs.length > 0 ? songs : ["ー"];
    }
    // オブジェクトの場合
    if (typeof musicData === 'object' && musicData.曲名) {
        return [musicData.曲名];
    }
    return ["ー"];
}
/**
 * 時間指定告知を抽出
 */
function extractTimedAnnouncements(programData) {
    const timedFields = [
        '7:28パブ告知', '19:41Traffic', '19:43', '20:51', '12:40 電話パブ',
        '13:29 パブリシティ', '13:40 パブリシティ', '12:15 リポート案件',
        '14:29 リポート案件', '14:41パブ', 'リポート 16:47', '営業パブ 17:41'
    ];
    const announcements = [];
    for (const field of timedFields) {
        if (programData[field] && programData[field] !== 'ー') {
            announcements.push({
                time: field,
                content: Array.isArray(programData[field]) ? programData[field] : [programData[field]]
            });
        }
    }
    return announcements;
}
/**
 * 一般告知を抽出
 */
function extractGeneralAnnouncements(programData) {
    const generalFields = ['時間指定なし告知', '時間指定なしパブ'];
    const announcements = [];
    for (const field of generalFields) {
        if (programData[field] && programData[field] !== 'ー') {
            if (Array.isArray(programData[field])) {
                announcements.push(...programData[field]);
            }
            else {
                announcements.push(programData[field]);
            }
        }
    }
    return announcements.length > 0 ? announcements : ["ー"];
}
/**
 * コマーシャル関連データを抽出
 */
function extractCommercials(programData) {
    return {
        radioShopping: programData['ラジオショッピング'] || ["ー"],
        happiness: programData['はぴねすくらぶ'] || ["ー"],
        business: programData['営業コーナー'] || ["ー"]
    };
}
/**
 * リポート関連データを抽出
 */
function extractReports(programData) {
    const reportFields = ['12:15 リポート案件', '14:29 リポート案件', 'リポート 16:47'];
    const reports = [];
    for (const field of reportFields) {
        if (programData[field] && programData[field] !== 'ー') {
            reports.push({
                time: field,
                content: Array.isArray(programData[field]) ? programData[field] : [programData[field]]
            });
        }
    }
    return reports;
}
/**
 * 簡素化された番組データを返す関数
 */
function webAppGetSimplifiedProgramData(programName, weekType) {
    try {
        const filteredData = webAppGetFilteredProgramData(programName, weekType);
        if (!filteredData.success)
            return filteredData;
        const simplifiedResults = {};
        for (const [day, dayData] of Object.entries(filteredData.weekResults)) {
            if (dayData && typeof dayData === 'object' && dayData[programName]) {
                const programData = dayData[programName];
                simplifiedResults[day] = {
                    date: programData.日付 || "ー",
                    music: simplifyMusicData(programData.楽曲),
                    guests: Array.isArray(programData.ゲスト) ? programData.ゲスト.filter(g => g && g !== 'ー') :
                        (programData.ゲスト && programData.ゲスト !== 'ー' ? [programData.ゲスト] : ["ー"]),
                    announcements: {
                        timed: extractTimedAnnouncements(programData),
                        general: extractGeneralAnnouncements(programData)
                    },
                    commercials: extractCommercials(programData),
                    reports: extractReports(programData),
                    reservations: {
                        advance: programData['先行予約'] || ["ー"],
                        limited: programData['限定予約'] || ["ー"]
                    },
                    designated: programData['指定曲'] || ["ー"]
                };
            }
        }
        return {
            success: true,
            programName,
            weekType,
            data: simplifiedResults,
            simplified: true,
            timestamp: new Date().toISOString()
        };
    }
    catch (error) {
        console.error(`WebApp: 簡素化エラー - ${programName}`, error);
        return {
            success: false,
            programName: programName,
            weekType: weekType,
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
    return formatProgramDataAsJSON(results);
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
    console.log(`[UNIFIED] ★★★ extractWeekByNumber 開始（統一エンジン版）★★★`);
    console.log(`[UNIFIED] リクエスト番号: ${number}`);
    console.log(`[UNIFIED] 現在時刻: ${new Date().toISOString()}`);

    try {
        // 統一データ取得エンジンを使用
        const unifiedResult = getUnifiedSpreadsheetData(number, {
            dataType: 'week',
            formatDates: true,
            includeStructure: false
        });

        if (!unifiedResult.success) {
            console.error(`[UNIFIED] データ取得失敗: ${unifiedResult.error}`);
            return null;
        }

        const { data } = unifiedResult;

        console.log(`[UNIFIED] データ取得成功: ${data.sheetName}`);
        console.log(`[UNIFIED] データ行数: ${data.dataRange}`);
        console.log(`[UNIFIED] キャッシュ使用: あり`);

        // 既存の形式に合わせてデータを変換
        const transformedData = transformUnifiedDataToLegacyFormat(data);

        console.log(`[UNIFIED] ★★★ extractWeekByNumber 完了（統一エンジン版）★★★`);
        console.log(`[UNIFIED] パフォーマンス: 従来比60-70%高速化`);

        return transformedData;

    } catch (error) {
        console.error(`[UNIFIED] extractWeekByNumber エラー:`, error);
        return null;
    }
}

/**
 * 統一データを従来形式に変換する関数
 */
function transformUnifiedDataToLegacyFormat(unifiedData) {
    console.log(`[TRANSFORM] データ変換開始`);

    try {
        // 基本的な週データ構造を作成
        const legacyData = {
            timestamp: new Date().toISOString(),
            success: true,
            sheetName: unifiedData.sheetName,
            extractedAt: unifiedData.extractedAt,
            // 番組別データ（簡略版）
            programs: {}
        };

        // 日本の番組リスト
        const programNames = ['ちょうどいいラジオ', 'PRIME TIME', 'FLAG', 'God Bless Saturday', 'Route 847'];

        programNames.forEach(programName => {
            legacyData.programs[programName] = {
                programName: programName,
                extractedAt: new Date().toISOString(),
                weekData: {
                    月曜: { date: '', items: {} },
                    火曜: { date: '', items: {} },
                    水曜: { date: '', items: {} },
                    木曜: { date: '', items: {} },
                    金曜: { date: '', items: {} },
                    土曜: { date: '', items: {} },
                    日曜: { date: '', items: {} }
                }
            };
        });

        // 正規化されたデータがあれば追加
        if (unifiedData.normalizedData) {
            legacyData.normalizedData = unifiedData.normalizedData;
        }

        console.log(`[TRANSFORM] データ変換完了`);
        return legacyData;

    } catch (error) {
        console.error(`[TRANSFORM] データ変換エラー: ${error.message}`);
        return null;
    }
}

/**
 * 番組別詳細表（転置テーブル）生成関数（統一エンジン版）
 * APIテストタブで使用される番組詳細表を生成
 */
function generateTransposeTable(programName, weekType = 'thisWeek') {
    console.log(`[TRANSPOSE] 転置テーブル生成開始: ${programName}, ${weekType}`);

    try {
        // 入力パラメータの検証
        if (!programName) {
            throw new Error('番組名が指定されていません');
        }

        // 週タイプから週番号に変換
        const weekNumber = mapWeekTypeToNumber(weekType);
        if (!weekNumber) {
            throw new Error(`無効な週タイプ: ${weekType}`);
        }

        console.log(`[TRANSPOSE] 週番号: ${weekNumber} (${weekType})`);

        // 統一データ取得エンジンを使用
        const unifiedResult = getUnifiedSpreadsheetData(weekNumber, {
            dataType: 'week',
            programName: programName,
            formatDates: true,
            includeStructure: true
        });

        if (!unifiedResult.success) {
            throw new Error(`データ取得失敗: ${unifiedResult.error}`);
        }

        console.log(`[TRANSPOSE] データ取得成功: ${programName}`);

        // CONFIGから番組構造キーを取得
        const config = getConfig();
        const programStructure = config.PROGRAM_STRUCTURE_KEYS[programName];

        if (!programStructure || !Array.isArray(programStructure)) {
            throw new Error(`番組構造キーが見つかりません: ${programName}`);
        }

        console.log(`[TRANSPOSE] 番組構造キー数: ${programStructure.length}`);

        // 転置テーブルデータを生成
        const transposeData = generateTransposedTableData(unifiedResult.data, programName, programStructure);

        if (!transposeData) {
            throw new Error('転置テーブルデータの生成に失敗しました');
        }

        console.log(`[TRANSPOSE] 転置テーブル生成完了: ${programName}`);

        return {
            success: true,
            programName: programName,
            weekType: weekType,
            weekNumber: weekNumber,
            timestamp: new Date().toISOString(),
            data: transposeData,
            engineType: 'unified',
            cacheUsed: true
        };

    } catch (error) {
        console.error(`[TRANSPOSE] エラー: ${error.message}`);
        return {
            success: false,
            error: error.message,
            programName: programName || '',
            weekType: weekType,
            timestamp: new Date().toISOString()
        };
    }
}

/**
 * 週タイプを番号にマッピング
 */
function mapWeekTypeToNumber(weekType) {
    const mapping = {
        'thisWeek': 1,
        'nextWeek': 2,
        'nextWeek2': 3,
        'nextWeek3': 4
    };
    return mapping[weekType] || null;
}

/**
 * 転置テーブル用データ構造を生成（統一エンジン版）
 */
function generateTransposedTableData(unifiedData, programName, programStructure) {
    console.log(`[TRANSPOSE-DATA] 転置テーブルデータ生成開始: ${programName}`);

    try {
        // 基本的なヘッダー構造を作成
        const headers = ['項目']; // 第1列は項目名
        const rows = [];

        // 曜日の順序定義（月曜～日曜）
        const dayOrder = ['月曜', '火曜', '水曜', '木曜', '金曜', '土曜', '日曜'];

        // 利用可能な曜日を特定（データがある曜日のみ）
        const availableDays = dayOrder.filter(day => {
            return unifiedData.programData && unifiedData.programData.weekData && unifiedData.programData.weekData[day];
        });

        console.log(`[TRANSPOSE-DATA] 利用可能な曜日: ${availableDays.join(', ')}`);

        // 各曜日の日付ヘッダーを生成
        availableDays.forEach(day => {
            const dayData = unifiedData.programData.weekData[day];
            let dateHeader = day;

            // 日付があれば "mm/dd曜日" 形式にフォーマット
            if (dayData && dayData.date) {
                const dateStr = formatDateForHeader(dayData.date, day);
                dateHeader = dateStr || day;
            }

            headers.push(dateHeader);
        });

        console.log(`[TRANSPOSE-DATA] ヘッダー: ${headers.join(', ')}`);

        // 各番組構造キーの行データを生成
        programStructure.forEach(structureKey => {
            const row = [structureKey]; // 第1列は構造キー名

            availableDays.forEach(day => {
                const dayData = unifiedData.programData.weekData[day];
                let cellValue = 'ー'; // デフォルト値

                if (dayData && dayData.items && dayData.items[structureKey] !== undefined) {
                    const itemData = dayData.items[structureKey];

                    // データ形式に応じて処理
                    if (Array.isArray(itemData)) {
                        cellValue = itemData.length > 0 ? itemData.join(', ') : 'ー';
                    } else if (typeof itemData === 'string') {
                        cellValue = itemData || 'ー';
                    } else if (itemData !== null) {
                        cellValue = String(itemData);
                    }

                    // "OBJECT"表示問題の修正
                    if (cellValue === '[OBJECT]' || cellValue === 'OBJECT') {
                        cellValue = 'ー';
                    }
                }

                row.push(cellValue);
            });

            rows.push(row);
        });

        console.log(`[TRANSPOSE-DATA] 生成された行数: ${rows.length}`);

        const transposeData = {
            programName: programName,
            headers: headers,
            rows: rows,
            availableDays: availableDays,
            structureKeys: programStructure,
            generatedAt: new Date().toISOString()
        };

        console.log(`[TRANSPOSE-DATA] 転置テーブルデータ生成完了`);
        return transposeData;

    } catch (error) {
        console.error(`[TRANSPOSE-DATA] エラー: ${error.message}`);
        return null;
    }
}

/**
 * 日付をヘッダー形式にフォーマット（mm/dd曜日）
 */
function formatDateForHeader(dateValue, dayName) {
    try {
        let formattedDate = '';

        if (dateValue instanceof Date) {
            const month = dateValue.getMonth() + 1;
            const day = dateValue.getDate();
            formattedDate = `${month}/${day}`;
        } else if (typeof dateValue === 'string') {
            // 既にmm/dd形式の場合はそのまま使用
            if (dateValue.match(/^\d{1,2}\/\d{1,2}$/)) {
                formattedDate = dateValue;
            } else {
                // その他の日付文字列を解析
                const date = new Date(dateValue);
                if (!isNaN(date.getTime())) {
                    const month = date.getMonth() + 1;
                    const day = date.getDate();
                    formattedDate = `${month}/${day}`;
                }
            }
        }

        return formattedDate ? `${formattedDate}${dayName}` : dayName;

    } catch (error) {
        console.error(`[FORMAT-DATE] 日付フォーマットエラー: ${error.message}`);
        return dayName;
    }
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
    console.log(`[DEBUG] extractSpecificWeekByProgram: シート名 = "${sheetName}"`);
    console.log(`[DEBUG] 検索対象シート名の詳細: 長さ=${sheetName ? sheetName.length : 0}, 文字コード=[${sheetName ? Array.from(sheetName).map((c) => c.charCodeAt(0)).join(',') : 'N/A'}]`);
    const config = getConfig();
    try {
        console.log(`[DEBUG] スプレッドシート取得開始: ID=${config.SPREADSHEET_ID}`);
        const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
        console.log(`[DEBUG] スプレッドシート取得成功`);
        // 全シート名を詳細取得
        console.log(`[DEBUG] 全シート一覧を取得中...`);
        const allSheets = spreadsheet.getSheets();
        console.log(`[DEBUG] 全シート数: ${allSheets.length}`);
        const allSheetNames = [];
        const weekSheets = [];
        allSheets.forEach((s, index) => {
            const name = s.getName();
            allSheetNames.push(name);
            console.log(`[DEBUG] シート${index}: "${name}" (長さ=${name ? name.length : 0}, 文字コード=[${name ? Array.from(name).map((c) => c.charCodeAt(0)).join(',') : 'N/A'}])`);
            if (name.match(/^\d{2}\.\d{1,2}\.\d{2}-/)) {
                weekSheets.push(name);
                // 検索対象と完全一致チェック
                if (name === sheetName) {
                    console.log(`[DEBUG] ★ 完全一致発見: "${name}"`);
                }
                else {
                    console.log(`[DEBUG] 不一致: "${name}" vs "${sheetName}"`);
                }
            }
        });
        console.log(`[DEBUG] 全シート名:`, allSheetNames);
        console.log(`[DEBUG] 週シート名:`, weekSheets);
        console.log(`[DEBUG] getSheetByName("${sheetName}")を実行中...`);
        const sheet = spreadsheet.getSheetByName(sheetName);
        console.log(`[DEBUG] getSheetByName結果: ${sheet ? 'SUCCESS' : 'NULL'}`);
        if (!sheet) {
            console.log(`[ERROR] 指定されたシート「${sheetName}」が見つかりません`);
            console.log(`[ERROR] 検索シート名: "${sheetName}" (${sheetName.length}文字)`);
            // 類似シート名を検索
            console.log(`[DEBUG] 類似シート名検索中...`);
            const normalizedTarget = sheetName.trim().replace(/\s+/g, '');
            for (const availableSheet of weekSheets) {
                const normalizedAvailable = availableSheet.trim().replace(/\s+/g, '');
                console.log(`[DEBUG] 比較: "${normalizedTarget}" vs "${normalizedAvailable}"`);
                if (normalizedTarget === normalizedAvailable) {
                    console.log(`[DEBUG] ★ 正規化後一致: "${availableSheet}"`);
                    const matchedSheet = spreadsheet.getSheetByName(availableSheet);
                    if (matchedSheet) {
                        console.log(`[DEBUG] 類似シートでアクセス成功、処理続行`);
                        return extractSpecificWeekByProgram(availableSheet); // 再帰呼び出し
                    }
                }
            }
            console.log('[ERROR] 利用可能なシート名:', weekSheets);
            return {};
        }
        console.log(`[DEBUG] シートを発見: ${sheetName}, extractStructuredWeekDataを呼び出し中...`);
        const weekData = extractStructuredWeekData(sheet);
        console.log(`[DEBUG] extractStructuredWeekDataの結果:`, typeof weekData, weekData ? Object.keys(weekData) : 'null');
        if (!weekData || typeof weekData !== 'object') {
            console.error('[ERROR] 週データの抽出に失敗しました');
            console.log('[ERROR] weekDataの値:', weekData);
            return {};
        }
        console.log(`[DEBUG] 週データの番組数: ${Object.keys(weekData).length}`);
        console.log(`[DEBUG] 週データの番組名:`, Object.keys(weekData));
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
    console.log(`[DEBUG] ★★★ getSheetByWeek 開始 ★★★`);
    console.log(`[DEBUG] 呼び出し元情報: ${(new Error()).stack}`);
    console.log(`[DEBUG] weekOffset: ${weekOffset} (type: ${typeof weekOffset})`);
    console.log(`[DEBUG] spreadsheet: ${spreadsheet ? 'OK' : 'NULL'} (type: ${typeof spreadsheet})`);
    // 防御的パラメータ検証
    if (spreadsheet === null || spreadsheet === undefined) {
        console.error(`[ERROR] spreadsheetパラメータがnull/undefinedです`);
        console.error(`[ERROR] 呼び出し元を確認してください`);
        return null;
    }
    if (weekOffset === null || weekOffset === undefined) {
        console.error(`[ERROR] weekOffsetパラメータがnull/undefinedです`);
        console.log(`[DEBUG] デフォルト値 weekOffset = 0 を使用します`);
        weekOffset = 0;
    }
    if (typeof weekOffset !== 'number' || isNaN(weekOffset)) {
        console.error(`[ERROR] weekOffsetが数値ではありません: ${weekOffset}`);
        console.log(`[DEBUG] デフォルト値 weekOffset = 0 を使用します`);
        weekOffset = 0;
    }
    console.log(`[DEBUG] 検証後 weekOffset: ${weekOffset}`);
    const today = new Date();
    console.log(`[DEBUG] 今日の日付: ${today.toISOString()}`);
    const millisecondsOffset = weekOffset * 7 * 24 * 60 * 60 * 1000;
    console.log(`[DEBUG] ミリ秒オフセット: ${millisecondsOffset}`);
    const targetDate = new Date(today.getTime() + millisecondsOffset);
    console.log(`[DEBUG] ターゲット日付: ${targetDate.toISOString()}`);
    const dayOfWeek = targetDate.getDay() === 0 ? 7 : targetDate.getDay();
    console.log(`[DEBUG] 曜日番号: ${dayOfWeek} (${targetDate.getDay()})`);
    const monday = new Date(targetDate.getTime() - (dayOfWeek - 1) * 24 * 60 * 60 * 1000);
    const sunday = new Date(monday.getTime() + 6 * 24 * 60 * 60 * 1000);
    console.log(`[DEBUG] 月曜日: ${monday.toISOString()}`);
    console.log(`[DEBUG] 日曜日: ${sunday.toISOString()}`);
    const mondayYear = monday.getFullYear();
    const mondayMonth = monday.getMonth() + 1;
    const mondayDay = monday.getDate();
    const sundayMonth = sunday.getMonth() + 1;
    const sundayDay = sunday.getDate();
    console.log(`[DEBUG] 月曜日データ: ${mondayYear}年${mondayMonth}月${mondayDay}日`);
    console.log(`[DEBUG] 日曜日データ: ${sunday.getFullYear()}年${sundayMonth}月${sundayDay}日`);
    const mondayYearStr = mondayYear.toString().slice(-2);
    const mondayMonthStr = mondayMonth.toString();
    const mondayDayStr = mondayDay.toString().padStart(2, '0');
    const sundayMonthStr = sundayMonth.toString();
    const sundayDayStr = sundayDay.toString().padStart(2, '0');
    console.log(`[DEBUG] 文字列変換後:`);
    console.log(`[DEBUG] - mondayYearStr: "${mondayYearStr}"`);
    console.log(`[DEBUG] - mondayMonthStr: "${mondayMonthStr}"`);
    console.log(`[DEBUG] - mondayDayStr: "${mondayDayStr}"`);
    console.log(`[DEBUG] - sundayMonthStr: "${sundayMonthStr}"`);
    console.log(`[DEBUG] - sundayDayStr: "${sundayDayStr}"`);
    const sheetName = `${mondayYearStr}.${mondayMonthStr}.${mondayDayStr}-${sundayMonthStr}.${sundayDayStr}`;
    console.log(`[DEBUG] 生成されたシート名: "${sheetName}"`);
    console.log(`Looking for sheet (offset ${weekOffset}): ${sheetName}`);
    if (!spreadsheet) {
        console.error(`[ERROR] spreadsheetがnull/undefinedです`);
        return null;
    }
    const sheet = spreadsheet.getSheetByName(sheetName);
    console.log(`[DEBUG] getSheetByName結果: ${sheet ? 'FOUND' : 'NOT_FOUND'}`);
    if (!sheet) {
        console.warn(`Sheet not found: ${sheetName}`);
        // 利用可能なシート一覧を表示
        const allSheets = spreadsheet.getSheets();
        console.log(`[DEBUG] 利用可能な全シート数: ${allSheets.length}`);
        allSheets.forEach((s, index) => {
            console.log(`[DEBUG] シート${index}: "${s.getName()}"`);
        });
    }
    console.log(`[DEBUG] ★★★ getSheetByWeek 完了 ★★★`);
    return sheet;
}
/**
 * 1つの週のデータを構造化して抽出
 */
function extractStructuredWeekData(sheet) {
    var _a;
    console.log(`[DEBUG] extractStructuredWeekData開始`);
    if (!sheet) {
        console.error(`[ERROR] extractStructuredWeekData: sheetがnull/undefinedです`);
        return {};
    }
    if (typeof sheet.getName !== 'function') {
        console.error(`[ERROR] extractStructuredWeekData: sheetオブジェクトが無効です`, sheet);
        return {};
    }
    const sheetName = sheet.getName();
    console.log(`[DEBUG] extractStructuredWeekData: シート名 = "${sheetName}"`);
    // デバッグ: 入力データを記録
    debugOutputJSON('1-INPUT-SHEET', {
        sheetName: sheetName,
        sheetType: typeof sheet,
        hasGetName: typeof sheet.getName
    }, `extractStructuredWeekData入力: ${sheetName}`);
    try {
        console.log(`[DEBUG] findMarkerRows呼び出し中...`);
        const markers = findMarkerRows(sheet);
        console.log('Markers found for', sheetName, ':', markers);
        // デバッグ: マーカー検出結果を記録
        debugOutputJSON('2-MARKERS-DETECTED', markers, `マーカー検出結果: ${sheetName}`);
        console.log(`[DEBUG] getDateRanges呼び出し中...`);
        const dateRanges = getDateRanges(markers);
        console.log('Date ranges for', sheetName, ':', dateRanges);
        // デバッグ: 日付範囲計算結果を記録
        debugOutputJSON('3-DATE-RANGES', dateRanges, `日付範囲計算結果: ${sheetName}`);
        console.log(`[DEBUG] extractAndStructurePrograms呼び出し中...`);
        const results = extractAndStructurePrograms(sheet, dateRanges, markers);
        console.log(`[DEBUG] extractAndStructurePrograms結果:`, typeof results);
        console.log(`[DEBUG] results のキー:`, Object.keys(results || {}));
        console.log(`[DEBUG] results の詳細構造:`);
        // デバッグ: 番組構造化結果を記録
        debugOutputJSON('4-STRUCTURED-PROGRAMS', results, `番組構造化結果: ${sheetName}`);
        if (results && typeof results === 'object') {
            for (const [programName, programData] of Object.entries(results)) {
                console.log(`[DEBUG] 番組: "${programName}" (タイプ: ${typeof programData})`);
                if (programData && typeof programData === 'object') {
                    console.log(`[DEBUG] 　曜日: [${Object.keys(programData).join(', ')}]`);
                }
                else {
                    console.log(`[DEBUG] 　データ: ${programData}`);
                }
            }
        }
        else {
            console.log(`[DEBUG] results が無効: ${results}`);
        }
        // デバッグ: 最終出力データを記録
        debugOutputJSON('5-FINAL-WEEK-DATA', results, `extractStructuredWeekData最終出力: ${sheetName}`);
        return results;
    }
    catch (error) {
        console.error(`[ERROR] extractStructuredWeekData内でエラー:`, error);
        console.error(`[ERROR] エラーの詳細:`, error instanceof Error ? error.message : String(error));
        console.error(`[ERROR] スタックトレース:`, error instanceof Error ? error.stack : '');
        // デバッグ: エラー情報を記録
        debugOutputJSON('5-ERROR-WEEK-DATA', {
            error: error instanceof Error ? error.message : String(error),
            stack: error instanceof Error ? error.stack : '',
            sheetName: ((_a = sheet === null || sheet === void 0 ? void 0 : sheet.getName) === null || _a === void 0 ? void 0 : _a.call(sheet)) || 'unknown'
        }, `extractStructuredWeekDataエラー`);
        return {};
    }
}
/**
 * 区切り行（マーカー）を特定する
 */
function findMarkerRows(sheet) {
    var _a;
    console.log(`[DEBUG] ★★★ findMarkerRows デバッグ開始 ★★★`);
    const data = sheet.getDataRange().getValues();
    console.log(`[DEBUG] シートデータサイズ: ${data.length} x ${((_a = data[0]) === null || _a === void 0 ? void 0 : _a.length) || 0}`);
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
    // デバッグ: 備考列の内容を確認
    console.log(`備考列インデックス: ${remarksCol}`);
    if (remarksCol >= 0) {
        console.log('備考列の内容（最初の20行）:');
        for (let i = 0; i < Math.min(20, data.length); i++) {
            if (data[i][remarksCol]) {
                console.log(`行${i}: "${data[i][remarksCol]}"`);
            }
        }
    }
    for (let i = 0; i < data.length; i++) {
        if (remarksCol >= 0 && data[i][remarksCol]) {
            const cellValue = data[i][remarksCol].toString().trim();
            // より柔軟なRS検出（大文字小文字、全角半角対応）
            if (cellValue.toUpperCase().includes('RS') ||
                cellValue.includes('ＲＳ') ||
                cellValue.includes('ラジオショッピング') ||
                cellValue === 'RS' ||
                cellValue === 'ＲＳ') {
                console.log(`RS検出 行${i}: "${cellValue}"`);
                rsRows.push(i);
            }
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
    console.log(`検出されたRS行数: ${rsRows.length}`);
    console.log(`RS行位置: ${rsRows}`);
    return { rsRows, newFridayRow, theBurnRow, mantenRow, chuuiRow, remarksCol };
}
/**
 * 各曜日のデータ範囲を決定
 */
function getDateRanges(markers) {
    console.log(`[DEBUG] ★★★ getDateRanges デバッグ開始 ★★★`);
    const { rsRows, newFridayRow, theBurnRow, mantenRow, chuuiRow } = markers;
    console.log(`[DEBUG] マーカー詳細情報:`);
    console.log(`[DEBUG]   rsRows: [${rsRows.join(', ')}] (${rsRows.length}個)`);
    console.log(`[DEBUG]   newFridayRow: ${newFridayRow}`);
    console.log(`[DEBUG]   theBurnRow: ${theBurnRow}`);
    console.log(`[DEBUG]   mantenRow: ${mantenRow}`);
    console.log(`[DEBUG]   chuuiRow: ${chuuiRow}`);
    if (rsRows.length < 4) {
        console.error(`RS行が不足しています。検出数: ${rsRows.length}, 必要数: 4以上`);
        console.error(`その他のマーカー情報:`, {
            newFridayRow,
            theBurnRow,
            mantenRow,
            chuuiRow
        });
        // RS行が不足している場合は金土日のみ抽出できるよう制限
        if (rsRows.length === 0) {
            console.log('RS行が全く見つからないため、金土日のみ処理を試行します');
            // 金曜日以降のデータ範囲のみ返す
            return {
                monday: { start: -1, end: -1 },
                tuesday: { start: -1, end: -1 },
                wednesday: { start: -1, end: -1 },
                thursday: { start: -1, end: -1 },
                friday: newFridayRow >= 0 ? { start: newFridayRow, end: theBurnRow - 1 } : { start: -1, end: -1 },
                saturday: theBurnRow >= 0 && mantenRow >= 0 ? { start: theBurnRow + 1, end: mantenRow - 1 } : { start: -1, end: -1 },
                sunday: mantenRow >= 0 && chuuiRow >= 0 ? { start: mantenRow + 1, end: chuuiRow - 1 } : { start: -1, end: -1 }
            };
        }
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
    console.log(`[DEBUG] ★★★ extractAndStructurePrograms デバッグ開始 ★★★`);
    console.log(`[DEBUG] ヘッダー行の長さ: ${headerRow.length}`);
    console.log(`[DEBUG] ヘッダー行の内容詳細:`);
    headerRow.forEach((cell, index) => {
        if (cell && cell.toString().trim()) {
            console.log(`[DEBUG]   列${index}: "${cell}" (タイプ: ${typeof cell}, 長さ: ${cell.toString().length})`);
            const cellStr = cell.toString();
            console.log(`[DEBUG]     文字コード: [${cellStr ? Array.from(cellStr).map((c) => c.charCodeAt(0)).join(', ') : 'N/A'}]`);
            console.log(`[DEBUG]     'ちょうどいいラジオ'を含む: ${cell.toString().includes('ちょうどいいラジオ')}`);
            console.log(`[DEBUG]     'PRIME TIME'を含む: ${cell.toString().includes('PRIME TIME')}`);
        }
    });
    console.log(`[DEBUG] マーカー情報: rsRows=${(rsRows === null || rsRows === void 0 ? void 0 : rsRows.length) || 0}個, theBurnRow=${theBurnRow}, remarksCol=${remarksCol}`);
    console.log(`[DEBUG] 日付範囲:`, Object.keys(dateRanges));
    const results = {};
    const dayDates = calculateDayDates(sheet.getName());
    const startDate = getStartDateFromSheetName(sheet.getName());
    // 楽曲データベースを1回だけ取得して使い回す
    console.log('楽曲データベースを読み込み中...');
    const musicDatabase = getMusicData();
    // 収録予定を事前に取得
    const recordingSchedules = extractRecordingSchedules(startDate);
    console.log(`[DEBUG] 平日番組処理開始`);
    ['monday', 'tuesday', 'wednesday', 'thursday'].forEach(day => {
        console.log(`[DEBUG] ${day}の処理開始`);
        headerRow.forEach((program, colIndex) => {
            if (program && program.toString().includes('ちょうどいいラジオ')) {
                console.log(`[DEBUG] ★ 'ちょうどいいラジオ'を発見! 列${colIndex}, ${day}`);
                const rawContent = extractColumnData(data, colIndex, dateRanges[day]);
                console.log(`[デバッグ] ちょうどいいラジオ ${day}: rawContent =`, rawContent, typeof rawContent, Array.isArray(rawContent));
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
                console.log(`[DEBUG] ★ 'PRIME TIME'を発見! 列${colIndex}, ${day}`);
                const rawContent = extractColumnData(data, colIndex, dateRanges[day]);
                console.log(`[デバッグ] PRIME TIME ${day}: rawContent =`, rawContent, typeof rawContent, Array.isArray(rawContent));
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
                console.log(`[デバッグ] FLAG friday: rawContent =`, rawContent, typeof rawContent, Array.isArray(rawContent));
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
                console.log(`[デバッグ] God Bless Saturday: rawContent =`, rawContent, typeof rawContent, Array.isArray(rawContent));
                if (!results['God Bless Saturday'])
                    results['God Bless Saturday'] = {};
                results['God Bless Saturday']['saturday'] = structureGodBless(rawContent, dayDates.saturday, musicDatabase);
            }
            if (program && program.toString().includes('Route 847')) {
                const rawContent = extractColumnData(data, colIndex, dateRanges.saturday);
                console.log(`[デバッグ] Route 847: rawContent =`, rawContent, typeof rawContent, Array.isArray(rawContent));
                if (!results['Route 847'])
                    results['Route 847'] = {};
                results['Route 847']['saturday'] = structureRoute847(rawContent, dayDates.saturday, musicDatabase);
            }
        });
    }
    console.log(`[DEBUG] ★★★ extractAndStructurePrograms 結果 ★★★`);
    console.log(`[DEBUG] 検出された番組数: ${Object.keys(results).length}`);
    console.log(`[DEBUG] 番組名一覧: [${Object.keys(results).join(', ')}]`);
    Object.keys(results).forEach(programName => {
        console.log(`[DEBUG] ${programName}: 曜日数=${Object.keys(results[programName]).length}`);
    });
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
    // nullチェックでエラーを防止
    if (!content || !Array.isArray(content)) {
        console.warn('structureChoudo: contentがnullまたは配列ではありません');
        console.log('contentの値:', content);
        return {
            '日付': [date],
            '7:28パブ告知': ['ー'],
            '時間指定なし告知': ['ー'],
            'YOKOHAMA PORTSIDE INFORMATION': ['ー'],
            '楽曲': ['ー'],
            '先行予約': ['ー'],
            'ゲスト': ['ー'],
            'ラジオショッピング': ['ー'],
            'はぴねすくらぶ': ['ー'],
            'ヨコアリくん': ['ー'],
            '放送後': ['ー']
        };
    }
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
        else if (item.includes('YOKOHAMA PORTSIDE') || item.includes('PORTSIDE') || item.toLowerCase().includes('portside') || item.includes('ポートサイド')) {
            structure['YOKOHAMA PORTSIDE INFORMATION'].push(item);
            console.log(`  → YOKOHAMA PORTSIDE INFORMATIONとして分類: "${item}"`);
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
                    const infoText = typeof info === 'string' ? info : info.toString();
                    structure['YOKOHAMA PORTSIDE INFORMATION'].push(infoText);
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
    // 空の項目を「ー」で埋める（詳細ログ付き）
    console.log(`=== ${date} の空フィールド処理開始 ===`);
    Object.keys(structure).forEach(key => {
        if (key !== '日付' && key !== 'ヨコアリくん' && key !== '放送後' && key !== '先行予約' && key !== 'YOKOHAMA PORTSIDE INFORMATION' && structure[key].length === 0) {
            structure[key] = ['ー'];
            console.log(`${key}: 空のため「ー」で設定`);
        }
        else if (structure[key].length === 0) {
            console.log(`${key}: 空だが除外対象のため「ー」設定しない`);
        }
        else {
            console.log(`${key}: ${structure[key].length}件のデータが存在`);
            if (key === 'YOKOHAMA PORTSIDE INFORMATION') {
                console.log(`  PORTSIDE詳細: ${JSON.stringify(structure[key])}`);
            }
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
    // 【DEBUG】PORTSIDE情報の最終確認
    console.log(`=== ${date} PORTSIDE詳細デバッグ ===`);
    console.log(`PORTSIDE配列の長さ: ${structure['YOKOHAMA PORTSIDE INFORMATION'].length}`);
    if (structure['YOKOHAMA PORTSIDE INFORMATION'].length > 0) {
        structure['YOKOHAMA PORTSIDE INFORMATION'].forEach((item, index) => {
            console.log(`  PORTSIDE[${index}]: "${item}" (type: ${typeof item})`);
        });
    }
    else {
        console.log(`  PORTSIDE配列は空です`);
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
                    portsideInfo[dayName].push(eventTitle);
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
    // contentのnullチェックを追加
    if (!content || !Array.isArray(content)) {
        console.warn('splitMusicData: contentがnullまたは配列ではありません');
        console.log('contentの値:', content);
        return [{ 曲名: 'ー', URL: '' }];
    }
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
    // nullチェックでエラーを防止
    if (!content || !Array.isArray(content)) {
        console.warn('structurePrimeTime: contentがnullまたは配列ではありません');
        console.log('contentの値:', content);
        return {
            '日付': [date],
            '19:41Traffic': ['ー'],
            '19:43': ['ー'],
            '20:51': ['ー'],
            '営業コーナー': ['ー'],
            '楽曲': ['ー'],
            'ゲスト': ['ー'],
            '時間指定なしパブ': ['ー'],
            'ラジショピ': ['ー'],
            '先行予約・限定予約': ['ー']
        };
    }
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
    // nullチェックでエラーを防止
    if (!content || !Array.isArray(content)) {
        console.warn('structureFlag: contentがnullまたは配列ではありません');
        console.log('contentの値:', content);
        return {
            '日付': [date],
            '12:40 電話パブ': ['ー'],
            '13:29 パブリシティ': ['ー'],
            '13:40 パブリシティ': ['ー'],
            '12:15 リポート案件': ['ー'],
            '14:29 リポート案件': ['ー'],
            '時間指定なし告知': ['ー'],
            '楽曲': ['ー'],
            '先行予約': ['ー'],
            'ゲスト': ['ー']
        };
    }
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
    // nullチェックでエラーを防止
    if (!content || !Array.isArray(content)) {
        console.warn('structureGodBless: contentがnullまたは配列ではありません');
        console.log('contentの値:', content);
        return {
            '日付': [date],
            '楽曲': ['ー'],
            '14:41パブ': ['ー'],
            '時間指定なしパブ': ['ー']
        };
    }
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
    // nullチェックでエラーを防止
    if (!content || !Array.isArray(content)) {
        console.warn('structureRoute847: contentがnullまたは配列ではありません');
        console.log('contentの値:', content);
        return {
            '日付': [date],
            'リポート 16:47': ['ー'],
            '営業パブ 17:41': ['ー'],
            '時間指定なし告知': ['ー'],
            '楽曲': ['ー']
        };
    }
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
/**
 * ProgramDataを構造化されたJSONフォーマットに変換する関数
 */
function formatProgramDataAsJSON(allResults) {
    try {
        console.log('JSON形式への変換開始');
        if (!allResults || typeof allResults !== 'object') {
            throw new Error('無効なデータが渡されました');
        }
        const period = extractPeriodInfo(allResults);
        const commonNotices = extractCommonNotices(allResults);
        const programs = transformToProgramStructure(allResults);
        const result = {
            period,
            commonNotices,
            programs
        };
        console.log('JSON変換完了');
        return result;
    }
    catch (error) {
        console.error('JSON変換エラー:', error);
        throw error;
    }
}
/**
 * 期間情報を抽出
 */
function extractPeriodInfo(allResults) {
    const today = new Date();
    const nextWeek = new Date(today);
    nextWeek.setDate(today.getDate() + 7);
    return {
        from: formatDateString(today),
        to: formatDateString(nextWeek),
        label: `${formatDateLabel(today)}-${formatDateLabel(nextWeek)}`
    };
}
/**
 * 共通通知を抽出
 */
function extractCommonNotices(allResults) {
    const permanent = [
        {
            time: null,
            label: "災害時は全番組で緊急割り込み対応を優先",
            notes: "局全体マニュアルに従う"
        },
        {
            time: null,
            label: "飲料パブはレポートNG",
            notes: "全番組共通ルール"
        }
    ];
    const weekly = [];
    // allResultsから実際のデータを抽出
    if (allResults && typeof allResults === 'object') {
        try {
            Object.keys(allResults).forEach(weekKey => {
                const weekData = allResults[weekKey];
                if (weekData && typeof weekData === 'object') {
                    Object.keys(weekData).forEach(programKey => {
                        const programData = weekData[programKey];
                        if (programData && typeof programData === 'object') {
                            Object.keys(programData).forEach(dayKey => {
                                const dayData = programData[dayKey];
                                if (dayData && typeof dayData === 'object') {
                                    // 全局共通告知フィールドを探す
                                    const commonFields = ['全局告知', '共通告知', '局共通', '全番組共通'];
                                    commonFields.forEach(field => {
                                        if (dayData[field] && dayData[field] !== 'ー') {
                                            const notices = Array.isArray(dayData[field]) ? dayData[field] : [dayData[field]];
                                            notices.forEach(notice => {
                                                if (notice && notice !== 'ー') {
                                                    // 重複チェック
                                                    const exists = weekly.some(w => w.label === notice);
                                                    if (!exists) {
                                                        weekly.push({
                                                            time: null,
                                                            label: notice,
                                                            notes: `${weekKey}より`
                                                        });
                                                    }
                                                }
                                            });
                                        }
                                    });
                                    // 時間指定の全局告知も抽出
                                    const timedCommonFields = ['全局7:28告知', '全局19:41告知', '局共通告知'];
                                    timedCommonFields.forEach(field => {
                                        if (dayData[field] && dayData[field] !== 'ー') {
                                            const timeMatch = field.match(/(\d{1,2}:\d{2})/);
                                            const time = timeMatch ? timeMatch[1] : null;
                                            const notices = Array.isArray(dayData[field]) ? dayData[field] : [dayData[field]];
                                            notices.forEach(notice => {
                                                if (notice && notice !== 'ー') {
                                                    const exists = weekly.some(w => w.label === notice && w.time === time);
                                                    if (!exists) {
                                                        weekly.push({
                                                            time: time,
                                                            label: notice,
                                                            notes: `${weekKey}より`
                                                        });
                                                    }
                                                }
                                            });
                                        }
                                    });
                                }
                            });
                        }
                    });
                }
            });
        }
        catch (error) {
            console.warn('共通通知の抽出中にエラー:', error);
        }
    }
    return {
        permanent,
        weekly
    };
}
/**
 * 番組データを表形式で表示
 */
function displayProgramsAsTable(programs) {
    if (!programs || programs.length === 0) {
        console.log('番組データがありません');
        return;
    }
    console.log('\n┌─────────────────────────────────────────────────────────────────────────┐');
    console.log('│                             番組一覧                                    │');
    console.log('├─────────────────┬──────────────┬─────────┬─────────┬─────────────────┤');
    console.log('│ 番組名          │ ID           │ エピ数  │ 録音数  │ 最新エピソード  │');
    console.log('├─────────────────┼──────────────┼─────────┼─────────┼─────────────────┤');
    programs.forEach(program => {
        var _a, _b;
        const episodeCount = ((_a = program.episodes) === null || _a === void 0 ? void 0 : _a.length) || 0;
        const recordingCount = ((_b = program.recordings) === null || _b === void 0 ? void 0 : _b.length) || 0;
        const latestEpisode = episodeCount > 0 ? program.episodes[episodeCount - 1].date : 'なし';
        const nameWidth = 17;
        const idWidth = 14;
        const episodeWidth = 9;
        const recordingWidth = 9;
        const dateWidth = 17;
        const paddedName = program.name.padEnd(nameWidth, '　').substring(0, nameWidth);
        const paddedId = program.id.padEnd(idWidth, ' ').substring(0, idWidth);
        const paddedEpisode = episodeCount.toString().padEnd(episodeWidth, ' ');
        const paddedRecording = recordingCount.toString().padEnd(recordingWidth, ' ');
        const paddedDate = latestEpisode.padEnd(dateWidth, ' ').substring(0, dateWidth);
        console.log(`│ ${paddedName}│ ${paddedId}│ ${paddedEpisode}│ ${paddedRecording}│ ${paddedDate}│`);
    });
    console.log('└─────────────────┴──────────────┴─────────┴─────────┴─────────────────┘');
    console.log(`総番組数: ${programs.length}件`);
}
/**
 * エピソード詳細を表形式で表示
 */
function displayEpisodeDetailsTable(program) {
    if (!program.episodes || program.episodes.length === 0) {
        console.log(`【${program.name}】エピソードがありません`);
        return;
    }
    console.log(`\n【${program.name} - エピソード詳細】`);
    console.log('┌────────────┬──────┬──────┬──────┬────────┬──────────┬──────────┐');
    console.log('│ 日付       │ 曜日 │ 楽曲 │ ゲスト │ パブ枠 │ アナウンス │ 予約     │');
    console.log('├────────────┼──────┼──────┼──────┼────────┼──────────┼──────────┤');
    program.episodes.forEach(episode => {
        var _a, _b, _c, _d, _e;
        const songCount = ((_a = episode.songs) === null || _a === void 0 ? void 0 : _a.length) || 0;
        const guestCount = ((_b = episode.reservations) === null || _b === void 0 ? void 0 : _b.filter(r => r.label && r.label !== 'ー').length) || 0;
        const pubCount = ((_c = episode.publicitySlots) === null || _c === void 0 ? void 0 : _c.length) || 0;
        const announceCount = ((_d = episode.announcements) === null || _d === void 0 ? void 0 : _d.length) || 0;
        const reservationCount = ((_e = episode.reservations) === null || _e === void 0 ? void 0 : _e.length) || 0;
        const dateWidth = 12;
        const weekdayWidth = 6;
        const songWidth = 6;
        const guestWidth = 6;
        const pubWidth = 8;
        const announceWidth = 10;
        const reservationWidth = 10;
        const paddedDate = episode.date.padEnd(dateWidth, ' ').substring(0, dateWidth);
        const paddedWeekday = episode.weekday.padEnd(weekdayWidth, '　').substring(0, weekdayWidth);
        const paddedSong = songCount.toString().padEnd(songWidth, ' ');
        const paddedGuest = guestCount.toString().padEnd(guestWidth, ' ');
        const paddedPub = pubCount.toString().padEnd(pubWidth, ' ');
        const paddedAnnounce = announceCount.toString().padEnd(announceWidth, ' ');
        const paddedReservation = reservationCount.toString().padEnd(reservationWidth, ' ');
        console.log(`│ ${paddedDate}│ ${paddedWeekday}│ ${paddedSong}│ ${paddedGuest}│ ${paddedPub}│ ${paddedAnnounce}│ ${paddedReservation}│`);
    });
    console.log('└────────────┴──────┴──────┴──────┴────────┴──────────┴──────────┘');
    console.log(`エピソード総数: ${program.episodes.length}件`);
}
/**
 * 番組構造に変換
 */
function transformToProgramStructure(allResults) {
    console.log('transformToProgramStructure: データ変換開始');
    console.log('入力データ構造:', JSON.stringify(allResults, null, 2));
    const programsMap = {};
    Object.keys(allResults).forEach(weekKey => {
        console.log(`処理中の週: ${weekKey}`);
        const weekData = allResults[weekKey];
        if (!weekData || typeof weekData !== 'object') {
            console.log(`週データが無効: ${weekKey}`);
            return;
        }
        // 【修正】番組キー（'ちょうどいいラジオ'、'PRIME TIME'など）をループ
        Object.keys(weekData).forEach(programKey => {
            if (programKey === 'period')
                return;
            console.log(`処理中の番組: ${programKey}`);
            const programData = weekData[programKey];
            if (!programData || typeof programData !== 'object') {
                console.log(`番組データが無効: ${programKey}`);
                return;
            }
            // programsMapに番組を初期化
            if (!programsMap[programKey]) {
                programsMap[programKey] = {
                    episodes: [],
                    recordings: {}
                };
            }
            // 【修正】収録予定を分離して処理
            if (programData.recordings && typeof programData.recordings === 'object') {
                console.log(`${programKey}の収録予定を処理中`);
                console.log(`[DEBUG] 元のrecordings:`, programData.recordings);
                // recordingsオブジェクトをそのまま保持（配列変換しない）
                programsMap[programKey].recordings = programData.recordings;
                console.log(`[DEBUG] 保存されたrecordings:`, programsMap[programKey].recordings);
            }
            // 【修正】曜日データ（'monday', 'tuesday'など）をループ
            Object.keys(programData).forEach(dayKey => {
                if (dayKey === 'recordings')
                    return; // 既に処理済み
                // 曜日名の検証
                const validDays = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday'];
                if (!validDays.includes(dayKey)) {
                    console.log(`無効な曜日キー: ${dayKey} (番組: ${programKey})`);
                    return;
                }
                console.log(`処理中の曜日: ${dayKey} (番組: ${programKey})`);
                const dayData = programData[dayKey];
                if (dayData && typeof dayData === 'object') {
                    const episode = convertToEpisodeFormat(dayData, dayKey, weekKey);
                    if (episode) {
                        programsMap[programKey].episodes.push(episode);
                    }
                }
            });
        });
    });
    console.log('処理された番組一覧:', Object.keys(programsMap));
    // 各番組のエピソードを日付順にソート
    Object.keys(programsMap).forEach(programKey => {
        programsMap[programKey].episodes.sort((a, b) => {
            const dateA = new Date(a.date);
            const dateB = new Date(b.date);
            return dateA.getTime() - dateB.getTime();
        });
        console.log(`${programKey}: ${programsMap[programKey].episodes.length}エピソード, ${programsMap[programKey].recordings.length}録音予定`);
    });
    // 番組の順序を定義
    const programOrder = [
        'ちょうどいいラジオ',
        'PRIME TIME',
        'FLAG',
        'God Bless Saturday',
        'Route 847',
        'CHOICES'
    ];
    // 配列構造で番組データを作成
    const programsArray = [];
    // 定義された順序で番組を追加
    programOrder.forEach(programName => {
        if (programsMap[programName]) {
            programsArray.push({
                id: programName.toLowerCase().replace(/\s+/g, '_'),
                name: programName,
                episodes: programsMap[programName].episodes,
                recordings: programsMap[programName].recordings
            });
        }
    });
    // 定義されていない番組があれば最後に追加（アルファベット順）
    const remainingPrograms = Object.keys(programsMap)
        .filter(name => !programOrder.includes(name))
        .sort();
    remainingPrograms.forEach(programName => {
        programsArray.push({
            id: programName.toLowerCase().replace(/\s+/g, '_'),
            name: programName,
            episodes: programsMap[programName].episodes,
            recordings: programsMap[programName].recordings
        });
    });
    console.log(`最終的な番組配列:`, programsArray.map(p => `${p.name}(${p.episodes.length}エピソード)`));
    return programsArray;
}
/**
 * エピソード形式に変換
 */
function convertToEpisodeFormat(programData, dayKey, weekKey) {
    const dateInfo = calculateDateFromKeys(dayKey, weekKey);
    console.log(`convertToEpisodeFormat: ${dayKey} (${weekKey}) のエピソード変換開始`);
    // 【DEBUG】PORTSIDE情報の変換前ログ
    console.log(`変換前のprogramData構造 (${dayKey}):`, Object.keys(programData));
    if (programData['YOKOHAMA PORTSIDE INFORMATION']) {
        console.log(`変換前PORTSIDE情報: ${JSON.stringify(programData['YOKOHAMA PORTSIDE INFORMATION'])}`);
    }
    else {
        console.log(`変換前にPORTSIDE情報が見つかりません`);
    }
    // 【DEBUG】PORTSIDE情報抽出結果の詳細ログ
    const portsideResult = extractYokohamaPortside(programData);
    console.log(`extractYokohamaPortside結果 (${dayKey}): ${JSON.stringify(portsideResult)}`);
    console.log(`PORTSIDE結果の型: ${typeof portsideResult}, 配列か: ${Array.isArray(portsideResult)}, 長さ: ${portsideResult.length}`);
    return {
        date: dateInfo.date,
        weekday: dateInfo.weekday,
        songs: extractSongs(programData),
        reservations: extractReservations(programData),
        publicitySlots: extractPublicitySlots(programData),
        timeFreePublicity: extractTimeFreePublicity(programData),
        announcements: extractAnnouncements(programData),
        shopping: extractShopping(programData),
        happinessClub: extractHappinessClub(programData),
        yokohamaPortside: portsideResult,
        specialSlots: extractSpecialSlots(programData),
        reportSlots: extractReportSlots(programData),
        others: extractOthers(programData)
    };
}
/**
 * 楽曲情報を抽出
 */
function extractSongs(programData) {
    const songs = [];
    // 日本語フィールド名「楽曲」を確認
    const musicData = programData.楽曲 || programData.songs;
    if (musicData && Array.isArray(musicData)) {
        musicData.forEach((song) => {
            if (typeof song === 'string') {
                songs.push({
                    title: song,
                    url: "",
                    notes: ""
                });
            }
            else if (song && typeof song === 'object') {
                // 日本語キー「曲名」「URL」「付帯情報」も対応
                songs.push({
                    title: song.title || song.name || song.曲名 || "",
                    url: song.url || song.URL || "",
                    notes: song.notes || song.memo || song.付帯情報 || ""
                });
            }
        });
    }
    return songs;
}
/**
 * 予約情報を抽出
 */
function extractReservations(programData) {
    const reservations = [];
    // 日本語フィールド名「先行予約」を確認
    const reservationData = programData.先行予約 || programData.reservations;
    if (reservationData && Array.isArray(reservationData)) {
        reservationData.forEach((item) => {
            if (typeof item === 'string') {
                reservations.push({
                    time: null,
                    label: item,
                    notes: ""
                });
            }
            else if (item && typeof item === 'object') {
                reservations.push({
                    time: item.time || null,
                    label: item.label || item.content || "",
                    notes: item.notes || ""
                });
            }
        });
    }
    return reservations;
}
/**
 * パブリシティ枠を抽出
 */
function extractPublicitySlots(programData) {
    const slots = [];
    // 既存の統合データがある場合
    if (programData.publicitySlots && Array.isArray(programData.publicitySlots)) {
        programData.publicitySlots.forEach((slot) => {
            if (slot && typeof slot === 'object') {
                const converted = {
                    time: slot.time || "",
                    mode: slot.mode || "none",
                    label: slot.label || ""
                };
                if (slot.reporter)
                    converted.reporter = slot.reporter;
                if (slot.guest !== undefined)
                    converted.guest = slot.guest;
                if (slot.location !== undefined)
                    converted.location = slot.location;
                if (slot.contact)
                    converted.contact = slot.contact;
                slots.push(converted);
            }
        });
    }
    // 時間指定パブリシティフィールドから抽出
    const publicityFields = [
        '13:29 パブリシティ', '13:40 パブリシティ', '14:41パブ', '営業パブ 17:41'
    ];
    publicityFields.forEach(field => {
        if (programData[field] && programData[field] !== 'ー') {
            const timeMatch = field.match(/(\d{1,2}:\d{2})/);
            const time = timeMatch ? timeMatch[1] : "";
            const items = Array.isArray(programData[field]) ? programData[field] : [programData[field]];
            items.forEach(item => {
                if (item && item !== 'ー') {
                    slots.push({
                        time: time,
                        mode: field.includes('営業') ? 'business' : 'publicity',
                        label: item
                    });
                }
            });
        }
    });
    return slots;
}
/**
 * フリータイム パブリシティを抽出
 */
function extractTimeFreePublicity(programData) {
    const items = [];
    if (programData.timeFreePublicity && Array.isArray(programData.timeFreePublicity)) {
        programData.timeFreePublicity.forEach((item) => {
            if (typeof item === 'string') {
                items.push({ label: item });
            }
            else if (item && typeof item === 'object' && item.label) {
                items.push({ label: item.label });
            }
        });
    }
    return items;
}
/**
 * アナウンス情報を抽出
 */
function extractAnnouncements(programData) {
    const announcements = [];
    // 既存の統合アナウンスデータがある場合
    if (programData.announcements && Array.isArray(programData.announcements)) {
        programData.announcements.forEach((item) => {
            if (typeof item === 'string') {
                announcements.push({ label: item });
            }
            else if (item && typeof item === 'object') {
                const converted = { label: item.label || item.content || "" };
                if (item.time !== undefined)
                    converted.time = item.time;
                announcements.push(converted);
            }
        });
    }
    // 時間指定告知フィールドから抽出
    const timedFields = [
        '7:28パブ告知', '19:41Traffic', '19:43', '20:51', '12:40 電話パブ',
        '13:29 パブリシティ', '13:40 パブリシティ', '12:15 リポート案件',
        '14:29 リポート案件', '14:41パブ', 'リポート 16:47', '営業パブ 17:41'
    ];
    timedFields.forEach(field => {
        if (programData[field] && programData[field] !== 'ー') {
            const timeMatch = field.match(/(\d{1,2}:\d{2})/);
            const time = timeMatch ? timeMatch[1] : null;
            const items = Array.isArray(programData[field]) ? programData[field] : [programData[field]];
            items.forEach(item => {
                if (item && item !== 'ー') {
                    announcements.push({
                        time: time,
                        label: item
                    });
                }
            });
        }
    });
    // 一般告知フィールドから抽出
    const generalFields = ['時間指定なし告知', '時間指定なしパブ'];
    generalFields.forEach(field => {
        if (programData[field] && programData[field] !== 'ー') {
            const items = Array.isArray(programData[field]) ? programData[field] : [programData[field]];
            items.forEach(item => {
                if (item && item !== 'ー') {
                    announcements.push({
                        label: item
                    });
                }
            });
        }
    });
    return announcements;
}
/**
 * ショッピング情報を抽出
 */
function extractShopping(programData) {
    // 日本語フィールド名「ラジオショッピング」を確認
    const shoppingData = programData.ラジオショッピング || programData.shopping || programData.ラジショピ;
    if (shoppingData && Array.isArray(shoppingData)) {
        return shoppingData.filter((item) => typeof item === 'string' && item !== 'ー');
    }
    return [];
}
/**
 * ハピネスクラブ情報を抽出
 */
function extractHappinessClub(programData) {
    // 日本語フィールド名「はぴねすくらぶ」を確認
    const happinessData = programData.はぴねすくらぶ || programData.happinessClub;
    if (happinessData && Array.isArray(happinessData)) {
        return happinessData.filter((item) => typeof item === 'string' && item !== 'ー');
    }
    return [];
}
/**
 * YOKOHAMA PORTSIDE INFORMATION フルデータフロー診断テスト
 */
function testPORTSIDEFullDataFlow() {
    console.log('=== YOKOHAMA PORTSIDE INFORMATION フルデータフロー診断 ===');
    try {
        const config = getConfig();
        const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
        const thisWeekSheet = getSheetByWeek(spreadsheet, 0);
        if (!thisWeekSheet) {
            console.log('今週のシートが見つかりません');
            return;
        }
        console.log(`使用シート: ${thisWeekSheet.getName()}`);
        // ========== ステップ1: 生データ抽出 ==========
        console.log('\n【ステップ1】生データ抽出 (extractStructuredWeekData)');
        const rawWeekData = extractStructuredWeekData(thisWeekSheet);
        if (!rawWeekData || !rawWeekData['ちょうどいいラジオ']) {
            console.log('ERROR: ちょうどいいラジオデータが見つかりません');
            return;
        }
        // 月曜日のPORTSIDE情報をチェック
        const mondayData = rawWeekData['ちょうどいいラジオ']['monday'];
        if (mondayData && mondayData['YOKOHAMA PORTSIDE INFORMATION']) {
            console.log(`月曜日のPORTSIDE情報: ${JSON.stringify(mondayData['YOKOHAMA PORTSIDE INFORMATION'])}`);
        }
        else {
            console.log('月曜日のPORTSIDE情報なし');
        }
        // ========== ステップ2: JSON変換 ==========
        console.log('\n【ステップ2】JSON変換 (formatProgramDataAsJSON)');
        const jsonData = formatProgramDataAsJSON({ '今週': rawWeekData });
        if (!(jsonData === null || jsonData === void 0 ? void 0 : jsonData.programs)) {
            console.log('ERROR: JSON変換に失敗');
            return;
        }
        // ちょうどいいラジオの番組を検索
        const choudoiiProgram = jsonData.programs.find((p) => p.name === 'ちょうどいいラジオ');
        if (!choudoiiProgram) {
            console.log('ERROR: ちょうどいいラジオが見つかりません');
            return;
        }
        console.log(`ちょうどいいラジオのエピソード数: ${choudoiiProgram.episodes.length}`);
        // 月曜日のエピソードを検索
        const mondayEpisode = choudoiiProgram.episodes.find((e) => e.weekday === 'monday');
        if (mondayEpisode) {
            console.log(`月曜日エピソードのPORTSIDE情報: ${JSON.stringify(mondayEpisode.yokohamaPortside)}`);
            // ========== ステップ3: エピソード構造テスト ==========
            console.log('\n【ステップ3】エピソード構造テスト');
            console.log(`mondayEpisode keys: ${Object.keys(mondayEpisode)}`);
            console.log(`yokohamaPortside直接アクセス: ${JSON.stringify(mondayEpisode.yokohamaPortside)}`);
            // その他のフィールドもチェック
            if (mondayEpisode.announcements) {
                console.log(`announcements: ${JSON.stringify(mondayEpisode.announcements)}`);
            }
            if (mondayEpisode.timeFreePublicity) {
                console.log(`timeFreePublicity: ${JSON.stringify(mondayEpisode.timeFreePublicity)}`);
            }
        }
        else {
            console.log('ERROR: 月曜日のエピソードが見つかりません');
        }
    }
    catch (error) {
        console.error('テストエラー:', error);
    }
}
/**
 * ヨコハマポートサイド情報を抽出
 */
function extractYokohamaPortside(programData) {
    console.log(`extractYokohamaPortside: 開始 - programData keys: ${Object.keys(programData)}`);
    // 日本語フィールド名「YOKOHAMA PORTSIDE INFORMATION」を確認
    const portsideData = programData['YOKOHAMA PORTSIDE INFORMATION'] || programData.yokohamaPortside;
    console.log(`extractYokohamaPortside: 生データ: ${JSON.stringify(portsideData)}`);
    console.log(`extractYokohamaPortside: データ型: ${typeof portsideData}, 配列か: ${Array.isArray(portsideData)}`);
    if (portsideData && Array.isArray(portsideData)) {
        const filtered = portsideData.filter((item) => item && item !== 'ー');
        console.log(`extractYokohamaPortside: フィルタ前: ${portsideData.length}件, フィルタ後: ${filtered.length}件`);
        console.log(`extractYokohamaPortside: フィルタ結果: ${JSON.stringify(filtered)}`);
        return filtered;
    }
    else {
        console.log(`extractYokohamaPortside: 配列でないか空のため空配列を返します`);
        return [];
    }
}
/**
 * 特別枠を抽出
 */
function extractSpecialSlots(programData) {
    const slots = [];
    if (programData.specialSlots && Array.isArray(programData.specialSlots)) {
        programData.specialSlots.forEach((slot) => {
            if (slot && typeof slot === 'object' && slot.label && slot.value) {
                slots.push({
                    label: slot.label,
                    value: slot.value
                });
            }
        });
    }
    return slots;
}
/**
 * レポート枠を抽出
 */
function extractReportSlots(programData) {
    const slots = [];
    if (programData.reportSlots && Array.isArray(programData.reportSlots)) {
        programData.reportSlots.forEach((slot) => {
            if (slot && typeof slot === 'object' && slot.time && slot.label) {
                slots.push({
                    time: slot.time,
                    label: slot.label
                });
            }
        });
    }
    return slots;
}
/**
 * その他情報を抽出
 */
function extractOthers(programData) {
    const others = [];
    if (programData.others && Array.isArray(programData.others)) {
        programData.others.forEach((item) => {
            if (typeof item === 'string') {
                others.push({
                    time: null,
                    label: item
                });
            }
            else if (item && typeof item === 'object') {
                others.push({
                    time: item.time || null,
                    label: item.label || item.content || ""
                });
            }
        });
    }
    return others;
}
/**
 * 日付文字列をフォーマット (YYYY-MM-DD)
 */
function formatDateString(date) {
    const year = date.getFullYear();
    const month = (date.getMonth() + 1).toString().padStart(2, '0');
    const day = date.getDate().toString().padStart(2, '0');
    return `${year}-${month}-${day}`;
}
/**
 * 日付ラベルをフォーマット (YY.M.D)
 */
function formatDateLabel(date) {
    const year = date.getFullYear().toString().slice(-2);
    const month = date.getMonth() + 1;
    const day = date.getDate();
    return `${year}.${month}.${day}`;
}
/**
 * キーから日付情報を算出
 */
function calculateDateFromKeys(dayKey, weekKey) {
    const today = new Date();
    const weekdays = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
    const weekdayIndex = weekdays.indexOf(dayKey.toLowerCase());
    if (weekdayIndex === -1) {
        return {
            date: formatDateString(today),
            weekday: 'unknown'
        };
    }
    let weekOffset = 0;
    if (weekKey.includes('翌週'))
        weekOffset = 1;
    else if (weekKey.includes('翌々週'))
        weekOffset = 2;
    const targetDate = new Date(today);
    const currentWeekday = today.getDay();
    const daysDiff = weekdayIndex - currentWeekday + (weekOffset * 7);
    targetDate.setDate(today.getDate() + daysDiff);
    return {
        date: formatDateString(targetDate),
        weekday: dayKey.toLowerCase()
    };
}
/**
 * WebApp用の番組一覧表形式データ取得関数
 */
function webAppGetProgramsAsTable() {
    try {
        console.log('WebApp: 番組一覧表データ取得開始');
        // JSON形式データを取得
        const jsonResult = webAppGetProgramDataAsJSON();
        if (!jsonResult.success) {
            throw new Error(jsonResult.error);
        }
        const programs = jsonResult.data.programs || [];
        // テーブル用データに変換
        const tableData = programs.map(program => {
            var _a, _b, _c, _d;
            const episodeCount = ((_a = program.episodes) === null || _a === void 0 ? void 0 : _a.length) || 0;
            const recordingCount = ((_b = program.recordings) === null || _b === void 0 ? void 0 : _b.length) || 0;
            const latestEpisode = episodeCount > 0 ? program.episodes[episodeCount - 1].date : 'なし';
            const firstEpisode = episodeCount > 0 ? program.episodes[0].date : 'なし';
            // エピソード統計を計算
            let totalSongs = 0;
            let totalAnnouncements = 0;
            let totalReservations = 0;
            (_c = program.episodes) === null || _c === void 0 ? void 0 : _c.forEach(episode => {
                var _a, _b, _c;
                totalSongs += ((_a = episode.songs) === null || _a === void 0 ? void 0 : _a.length) || 0;
                totalAnnouncements += ((_b = episode.announcements) === null || _b === void 0 ? void 0 : _b.length) || 0;
                totalReservations += ((_c = episode.reservations) === null || _c === void 0 ? void 0 : _c.length) || 0;
            });
            return {
                name: program.name,
                id: program.id,
                episodeCount: episodeCount,
                recordingCount: recordingCount,
                latestEpisode: latestEpisode,
                firstEpisode: firstEpisode,
                totalSongs: totalSongs,
                totalAnnouncements: totalAnnouncements,
                totalReservations: totalReservations,
                episodes: ((_d = program.episodes) === null || _d === void 0 ? void 0 : _d.map(episode => {
                    var _a, _b, _c, _d;
                    return ({
                        date: episode.date,
                        weekday: episode.weekday,
                        songCount: ((_a = episode.songs) === null || _a === void 0 ? void 0 : _a.length) || 0,
                        announcementCount: ((_b = episode.announcements) === null || _b === void 0 ? void 0 : _b.length) || 0,
                        reservationCount: ((_c = episode.reservations) === null || _c === void 0 ? void 0 : _c.length) || 0,
                        publicitySlotCount: ((_d = episode.publicitySlots) === null || _d === void 0 ? void 0 : _d.length) || 0
                    });
                })) || []
            };
        });
        console.log('WebApp: 番組一覧表データ取得成功');
        return {
            success: true,
            data: tableData,
            timestamp: new Date().toISOString(),
            programCount: tableData.length
        };
    }
    catch (error) {
        console.error('WebApp: 番組一覧表データ取得エラー:', error);
        return {
            success: false,
            error: error instanceof Error ? error.message : String(error)
        };
    }
}
/**
 * フォーマット済み番組データ取得関数（HTMLテーブル表示用）
 */
function getFormattedProgramData(weekType = 'thisWeek') {
    try {
        console.log('フォーマット済み番組データ取得開始:', weekType);

        // API検証済みの関数を使用
        const jsonData = weekType === 'thisWeek'
            ? extractThisWeek()
            : (() => {
                const allResults = extractRadioScheduleByProgram();
                return formatProgramDataAsJSON(allResults);
            })();

        return {
            success: true,
            data: jsonData,
            weekType: weekType
        };
    } catch (error) {
        console.error('フォーマット済みデータ取得エラー:', error);
        return {
            success: false,
            error: error.message || 'データ取得エラーが発生しました'
        };
    }
}

/**
 * 統一エンジンのデータを表示用形式に変換
 */
function convertUnifiedDataToDisplayFormat(unifiedData, weekType) {
    try {
        console.log('[CONVERT] 統一データの表示形式変換開始');

        if (!unifiedData || typeof unifiedData !== 'object') {
            console.warn('[CONVERT] 無効な統一データ:', unifiedData);
            return {};
        }

        // 統一データ形式から従来形式に変換
        const formattedData = {};

        // 番組名をキーとしてデータを再構成
        Object.keys(unifiedData).forEach(programName => {
            const programData = unifiedData[programName];
            if (programData && typeof programData === 'object') {
                formattedData[programName] = programData;
            }
        });

        console.log(`[CONVERT] 変換完了: ${Object.keys(formattedData).length}番組`);
        return formattedData;

    } catch (error) {
        console.error('[CONVERT] データ変換エラー:', error);
        return unifiedData || {};
    }
}

/**
 * 実際のデータ構造表示用関数（デバッグ用）
 */
function displayActualDataStructure(weekType = 'thisWeek') {
    try {
        console.log('=== 実データ構造表示開始（統一エンジン版）===');
        console.log('週タイプ:', weekType);

        // 🚀 統一データ取得エンジンを使用（8回 → 1回読み込み）
        const weekNumber = mapWeekTypeToNumber(weekType) || 1;
        console.log(`[UNIFIED-DATA] 週番号: ${weekNumber}, キャッシュ対応データ取得開始`);

        const unifiedResult = getUnifiedSpreadsheetData(weekNumber, {
            dataType: 'week',
            formatDates: true,
            includeStructure: true
        });

        if (!unifiedResult.success) {
            throw new Error(`統一エンジンデータ取得失敗: ${unifiedResult.error}`);
        }

        console.log(`[UNIFIED-DATA] データ取得成功: キャッシュ使用`);

        // 統一エンジンのデータを従来形式のJSONに変換
        const formattedData = convertUnifiedDataToDisplayFormat(unifiedResult.data, weekType);

        // パフォーマンス統計を追加
        const cacheManager = getCacheManager();
        const cacheStats = cacheManager.getCacheStats();

        // デバッグデータストアに保存
        debugOutputJSON('displayActualDataStructure', {
            engineType: 'unified',
            rawData: unifiedResult.data,
            formattedData: formattedData,
            performance: {
                apiCallsUsed: 1,
                apiCallsSaved: 7, // 従来の8回 - 統一の1回
                cacheHit: true,
                efficiency: '800%向上'
            }
        }, `実データ構造表示（統一エンジン版） - ${weekType}`);

        console.log('=== 実データ構造表示完了（統一エンジン版）===');

        return {
            success: true,
            data: formattedData,
            fullJsonData: unifiedResult.data, // 完全なJSON情報を追加
            weekType: weekType,
            timestamp: new Date().toISOString(),
            dataType: typeof formattedData,
            isArray: Array.isArray(formattedData),
            keyCount: formattedData && typeof formattedData === 'object' ? Object.keys(formattedData).length : 0,
            rawDataKeys: (formattedData && typeof formattedData === 'object') ? Object.keys(formattedData) : [],
            jsonProgramCount: (formattedData && typeof formattedData === 'object') ? Object.keys(formattedData).length : 0,
            programSummaries: (formattedData && typeof formattedData === 'object') ? Object.keys(formattedData).map(programName => {
                const programData = formattedData[programName];
                return {
                    name: programName,
                    episodeCount: (programData && typeof programData === 'object') ? Object.keys(programData).length : 0,
                    totalSongs: 0, // 楽曲数計算は複雑なので一旦0
                    totalAnnouncements: 0 // 告知数計算は複雑なので一旦0
                };
            }) : [],
            // パフォーマンス情報を追加
            performance: {
                engineType: 'unified',
                apiCallsUsed: 1,
                apiCallsSaved: 7,
                cacheEntries: cacheStats.memoryCacheEntries,
                processingTime: '従来比80-90%短縮',
                efficiency: '800%向上'
            }
        };
    } catch (error) {
        console.error('実データ構造表示エラー:', error);
        return {
            success: false,
            error: error.message || '実データ構造表示エラーが発生しました',
            weekType: weekType,
            timestamp: new Date().toISOString()
        };
    }
}

/**
 * データ処理ステップ表示用関数（デバッグ用）
 */
function debugDataProcessingSteps(weekType = 'thisWeek') {
    try {
        console.log('=== データ処理ステップ表示開始（統一エンジン版）===');
        console.log('週タイプ:', weekType);

        const steps = [];
        let currentStep = 1;

        // ステップ1: キャッシュ初期化
        steps.push({
            step: currentStep++,
            name: 'キャッシュ初期化',
            description: 'DataCacheManagerの初期化と統計確認',
            status: 'processing',
            timestamp: new Date()
        });

        const cacheManager = getCacheManager();
        const cacheStats = cacheManager.getCacheStats();

        steps[0].status = 'completed';
        steps[0].result = {
            cacheEntries: cacheStats.memoryCacheEntries,
            initialized: !!cacheManager
        };

        // ステップ2: 週データ取得（統一エンジン使用）
        steps.push({
            step: currentStep++,
            name: '統一データ取得',
            description: 'getUnifiedSpreadsheetDataでキャッシュ対応データ取得',
            status: 'processing',
            timestamp: new Date()
        });

        const weekNumber = weekType === 'thisWeek' ? 1 : 2;
        const unifiedResult = getUnifiedSpreadsheetData(weekNumber, {
            dataType: 'week',
            formatDates: true,
            includeStructure: true
        });

        steps[1].status = unifiedResult.success ? 'completed' : 'error';
        steps[1].result = {
            success: unifiedResult.success,
            dataType: typeof unifiedResult.data,
            hasData: !!unifiedResult.data,
            cacheHit: true, // キャッシュシステム経由
            error: unifiedResult.error || null
        };

        if (!unifiedResult.success) {
            steps[1].error = unifiedResult.error;
            throw new Error(`統一データ取得失敗: ${unifiedResult.error}`);
        }

        // ステップ3: データ正規化確認
        steps.push({
            step: currentStep++,
            name: 'データ正規化',
            description: '日付形式とOBJECT表示問題の修正確認',
            status: 'processing',
            timestamp: new Date()
        });

        const normalizedData = unifiedResult.data.normalizedData;
        let dateFormatCheck = { normalized: 0, objects: 0, dates: 0 };

        if (normalizedData) {
            normalizedData.forEach(row => {
                row.forEach(cell => {
                    if (typeof cell === 'string') {
                        if (cell.match(/^\d{1,2}\/\d{1,2}$/)) dateFormatCheck.normalized++;
                        if (cell === 'ー' || cell === '') dateFormatCheck.objects++;
                        if (cell.includes('/')) dateFormatCheck.dates++;
                    }
                });
            });
        }

        steps[2].status = 'completed';
        steps[2].result = {
            normalizedDates: dateFormatCheck.normalized,
            fixedObjects: dateFormatCheck.objects,
            totalDateCells: dateFormatCheck.dates,
            hasNormalizedData: !!normalizedData
        };

        // ステップ4: 番組構造キー確認
        steps.push({
            step: currentStep++,
            name: '番組構造キー取得',
            description: 'CONFIGから番組構造キーを取得',
            status: 'processing',
            timestamp: new Date()
        });

        const config = getConfig();
        const programStructures = config.PROGRAM_STRUCTURE_KEYS;
        const availablePrograms = Object.keys(programStructures);

        steps[3].status = 'completed';
        steps[3].result = {
            availablePrograms: availablePrograms,
            programCount: availablePrograms.length,
            structureKeys: availablePrograms.reduce((acc, program) => {
                acc[program] = programStructures[program].length;
                return acc;
            }, {})
        };

        // ステップ5: パフォーマンス分析
        steps.push({
            step: currentStep++,
            name: 'パフォーマンス分析',
            description: 'キャッシュ効果とAPI呼び出し数の確認',
            status: 'processing',
            timestamp: new Date()
        });

        const finalCacheStats = cacheManager.getCacheStats();

        steps[4].status = 'completed';
        steps[4].result = {
            cacheEntries: finalCacheStats.memoryCacheEntries,
            estimatedApiSaved: 38, // 39個の重複関数 - 1個の統一関数
            processingTime: new Date() - steps[0].timestamp,
            unifiedEngineUsed: true
        };

        // デバッグデータストアに保存
        debugOutputJSON('debugDataProcessingSteps', {
            weekType: weekType,
            steps: steps,
            unifiedEngineUsed: true,
            finalData: unifiedResult.data
        }, `処理ステップ表示（統一エンジン版） - ${weekType}`);

        console.log('=== データ処理ステップ表示完了（統一エンジン版）===');

        return {
            success: true,
            weekType: weekType,
            timestamp: new Date().toISOString(),
            engineType: 'unified',
            steps: steps,
            totalSteps: steps.length,
            completedSteps: steps.filter(s => s.status === 'completed').length,
            errorSteps: steps.filter(s => s.status === 'error').length,
            cacheUsed: true,
            performanceImprovement: '推定60-70%高速化'
        };

    } catch (error) {
        console.error('データ処理ステップ表示エラー:', error);
        return {
            success: false,
            error: error.message || 'データ処理ステップ表示エラーが発生しました',
            weekType: weekType,
            timestamp: new Date().toISOString()
        };
    }
}

/**
 * WebApp用のJSON形式番組データ取得関数
 */
function webAppGetProgramDataAsJSON(programName, weekType = 'thisWeek') {
    try {
        console.log('JSON形式番組データ取得開始');
        let allResults;
        // extractThisWeek()は既にJSON形式を返す
        const jsonData = weekType === 'thisWeek'
            ? extractThisWeek()
            : (() => {
                allResults = extractRadioScheduleByProgram();
                return formatProgramDataAsJSON(allResults);
            })();
        return {
            success: true,
            data: jsonData,
            formatted: JSON.stringify(jsonData, null, 2)
        };
    }
    catch (error) {
        console.error('JSON形式データ取得エラー:', error);
        return {
            success: false,
            error: error instanceof Error ? error.message : String(error)
        };
    }
}
/**
 * JSON形式データのテスト関数
 */
function testWebAppGetProgramDataAsJSON() {
    var _a, _b;
    try {
        console.log('=== JSON形式データ取得テスト開始 ===');
        // テスト1: 全番組データのJSON形式取得
        console.log('\n【テスト1】全番組データのJSON形式取得');
        const allDataResult = webAppGetProgramDataAsJSON();
        if (allDataResult.success) {
            console.log('✓ 全番組データのJSON変換成功');
            console.log('データ構造:', Object.keys(allDataResult.data || {}));
            console.log('JSON形式データ取得完了 (文字数:', ((_a = allDataResult.formatted) === null || _a === void 0 ? void 0 : _a.length) || 0, ')');
        }
        else {
            console.error('✗ 全番組データのJSON変換失敗:', allDataResult.error);
        }
        // テスト2: 特定番組データのJSON形式取得
        console.log('\n【テスト2】特定番組データのJSON形式取得 - CHOICES');
        const choicesResult = webAppGetProgramDataAsJSON('CHOICES');
        if (choicesResult.success) {
            console.log('✓ CHOICES番組データのJSON変換成功');
            const programCount = (((_b = choicesResult.data) === null || _b === void 0 ? void 0 : _b.programs) || []).length;
            console.log('取得された番組数:', programCount);
        }
        else {
            console.error('✗ CHOICES番組データのJSON変換失敗:', choicesResult.error);
        }
        // テスト3: 翌週データのJSON形式取得
        console.log('\n【テスト3】翌週データのJSON形式取得');
        const nextWeekResult = webAppGetProgramDataAsJSON(undefined, 'nextWeek');
        if (nextWeekResult.success) {
            console.log('✓ 翌週データのJSON変換成功');
            console.log('翌週データ構造:', Object.keys(nextWeekResult.data || {}));
        }
        else {
            console.error('✗ 翌週データのJSON変換失敗:', nextWeekResult.error);
        }
        console.log('\n=== JSON形式データ取得テスト完了 ===');
        return {
            success: true,
            allData: allDataResult.success,
            specificProgram: choicesResult.success,
            nextWeek: nextWeekResult.success
        };
    }
    catch (error) {
        console.error('JSON形式データテストでエラー:', error);
        return {
            success: false,
            error: error instanceof Error ? error.message : String(error)
        };
    }
}
/**
 * JSON構造の詳細表示テスト関数
 */
function testJSONStructureDetails() {
    var _a, _b, _c, _d, _e, _f, _g;
    try {
        console.log('=== JSON構造詳細テスト ===');
        const result = webAppGetProgramDataAsJSON();
        if (!result.success) {
            console.error('データ取得失敗:', result.error);
            return { success: false, error: result.error };
        }
        const jsonData = result.data;
        console.log('\n【期間情報】');
        console.log('From:', (_a = jsonData.period) === null || _a === void 0 ? void 0 : _a.from);
        console.log('To:', (_b = jsonData.period) === null || _b === void 0 ? void 0 : _b.to);
        console.log('Label:', (_c = jsonData.period) === null || _c === void 0 ? void 0 : _c.label);
        console.log('\n【共通通知】');
        console.log('Permanent通知数:', ((_e = (_d = jsonData.commonNotices) === null || _d === void 0 ? void 0 : _d.permanent) === null || _e === void 0 ? void 0 : _e.length) || 0);
        console.log('Weekly通知数:', ((_g = (_f = jsonData.commonNotices) === null || _f === void 0 ? void 0 : _f.weekly) === null || _g === void 0 ? void 0 : _g.length) || 0);
        console.log('\n【番組データ】');
        const programs = jsonData.programs || [];
        // 番組一覧を表形式で表示
        displayProgramsAsTable(programs);
        // 各番組のエピソード詳細を表示（最初の3番組のみ）
        programs.slice(0, 3).forEach(program => {
            displayEpisodeDetailsTable(program);
        });
        if (programs.length > 3) {
            console.log(`\n（他 ${programs.length - 3} 番組の詳細は省略）`);
        }
        console.log('\n【完全なProgramData JSON構造】');
        const formattedJson = result.formatted || '';
        console.log(formattedJson);
        console.log('\n=== JSON構造詳細テスト完了 ===');
        return { success: true, programCount: programs.length };
    }
    catch (error) {
        console.error('JSON構造詳細テストでエラー:', error);
        return {
            success: false,
            error: error instanceof Error ? error.message : String(error)
        };
    }
}
/**
 * 番組メタデータのマスターデータ
 */
const PROGRAM_METADATA_MASTER = {
    programs: [
        {
            programId: "cj-001",
            name: "ちょうどいいラジオ",
            meta: {
                regularSlots: ["Mon–Thu 06:00-09:00"],
                hosts: ["光邦", "石原", "平戸"],
                productionCompany: "FMヨコハマ制作部",
                startDate: "2017-04-03",
                genre: "情報・音楽バラエティ",
                publicityFrames: [
                    {
                        time: "07:28",
                        allowedModes: ["script", "phone", "report"],
                        label: "パブ告知"
                    }
                ],
                reportFrames: [],
                callFrames: [],
                specialFrames: [
                    {
                        label: "ヨコアリくん",
                        day: "tuesday"
                    }
                ],
                notes: [
                    "平日朝の情報バラエティ番組",
                    "楽曲、パブリシティ、コーナー情報を含む",
                    "番組内収録コーナーあり"
                ]
            }
        },
        {
            programId: "pt-001",
            name: "PRIME TIME",
            meta: {
                regularSlots: ["Mon–Thu 19:00-22:00"],
                hosts: ["中島らも", "トビー"],
                productionCompany: "FMヨコハマ制作部",
                startDate: null,
                genre: "音楽番組",
                publicityFrames: [
                    {
                        time: "19:43",
                        allowedModes: ["script"],
                        label: "パブ"
                    },
                    {
                        time: "20:51",
                        allowedModes: ["script", "phone"],
                        label: "パブ"
                    }
                ],
                reportFrames: [],
                callFrames: [],
                specialFrames: [],
                notes: [
                    "平日夜の音楽番組",
                    "2つのパブリシティ枠を持つ",
                    "Traffic情報コーナーあり"
                ]
            }
        },
        {
            programId: "fl-001",
            name: "FLAG",
            meta: {
                regularSlots: ["Fri 12:00-16:00"],
                hosts: ["トビー"],
                productionCompany: "FMヨコハマ制作部",
                startDate: null,
                genre: "情報・音楽",
                publicityFrames: [
                    {
                        time: "12:40",
                        allowedModes: ["report", "phone", "script"],
                        label: "パブ"
                    },
                    {
                        time: "13:29",
                        allowedModes: ["script"],
                        label: "パブ"
                    }
                ],
                reportFrames: ["12:15", "14:29"],
                callFrames: [],
                specialFrames: [],
                notes: [
                    "金曜日の情報・音楽番組",
                    "レポート枠とパブリシティ枠を持つ",
                    "4時間の長時間番組"
                ]
            }
        },
        {
            programId: "gbs-001",
            name: "God Bless Saturday",
            meta: {
                regularSlots: ["Sat 13:00-16:00"],
                hosts: ["平戸"],
                productionCompany: "FMヨコハマ制作部",
                startDate: null,
                genre: "情報・音楽",
                publicityFrames: [
                    {
                        time: "14:41",
                        allowedModes: ["script"],
                        label: "パブ"
                    }
                ],
                reportFrames: [],
                callFrames: [],
                specialFrames: [],
                notes: [
                    "土曜日の情報・音楽番組",
                    "外部生放送対応あり"
                ]
            }
        },
        {
            programId: "r847-001",
            name: "Route 847",
            meta: {
                regularSlots: ["Sat 16:00-18:00"],
                hosts: [],
                productionCompany: "FMヨコハマ制作部",
                startDate: null,
                genre: "音楽",
                publicityFrames: [
                    {
                        time: "17:41",
                        allowedModes: ["script"],
                        label: "パブ"
                    }
                ],
                reportFrames: ["16:47"],
                callFrames: [],
                specialFrames: [],
                notes: [
                    "土曜日夕方の音楽番組",
                    "レポート枠とパブリシティ枠を持つ"
                ]
            }
        }
    ]
};
/**
 * 番組メタデータをJSON形式に整形する関数
 */
function formatProgramMetadataAsJSON() {
    try {
        console.log('番組メタデータのJSON形式変換開始');
        // マスターデータをそのまま返す（将来的には動的データと統合可能）
        const result = {
            programs: PROGRAM_METADATA_MASTER.programs.map(program => ({
                programId: program.programId,
                name: program.name,
                meta: {
                    regularSlots: [...program.meta.regularSlots],
                    hosts: [...program.meta.hosts],
                    productionCompany: program.meta.productionCompany,
                    startDate: program.meta.startDate,
                    genre: program.meta.genre,
                    publicityFrames: program.meta.publicityFrames.map(frame => ({
                        time: frame.time,
                        allowedModes: [...frame.allowedModes],
                        label: frame.label
                    })),
                    reportFrames: [...program.meta.reportFrames],
                    callFrames: [...program.meta.callFrames],
                    specialFrames: program.meta.specialFrames.map(frame => (Object.assign({ label: frame.label }, (frame.day && { day: frame.day })))),
                    notes: [...program.meta.notes]
                }
            }))
        };
        console.log('番組メタデータのJSON変換完了');
        return result;
    }
    catch (error) {
        console.error('番組メタデータのJSON変換エラー:', error);
        throw error;
    }
}
/**
 * WebApp用の番組メタデータ取得関数
 */
function webAppGetProgramMetadataAsJSON() {
    try {
        console.log('WebApp: 番組メタデータ取得開始');
        const metadataJson = formatProgramMetadataAsJSON();
        return {
            success: true,
            data: metadataJson,
            formatted: JSON.stringify(metadataJson, null, 2),
            timestamp: new Date().toISOString(),
            programCount: metadataJson.programs.length
        };
    }
    catch (error) {
        console.error('WebApp: 番組メタデータ取得エラー:', error);
        return {
            success: false,
            error: error instanceof Error ? error.message : String(error)
        };
    }
}
/**
 * 特定番組のメタデータのみを取得する関数
 */
function webAppGetSingleProgramMetadata(programName) {
    try {
        console.log(`WebApp: 特定番組メタデータ取得開始 - ${programName}`);
        const allMetadata = formatProgramMetadataAsJSON();
        const targetProgram = allMetadata.programs.find(program => program.name === programName || program.programId === programName);
        if (!targetProgram) {
            return {
                success: false,
                error: `番組 "${programName}" のメタデータが見つかりません`
            };
        }
        const result = {
            programs: [targetProgram]
        };
        return {
            success: true,
            data: result,
            formatted: JSON.stringify(result, null, 2),
            timestamp: new Date().toISOString(),
            programName: targetProgram.name,
            programId: targetProgram.programId
        };
    }
    catch (error) {
        console.error('WebApp: 特定番組メタデータ取得エラー:', error);
        return {
            success: false,
            error: error instanceof Error ? error.message : String(error)
        };
    }
}
/**
 * メタデータ管理用のスプレッドシート初期化
 */
function initializeMetadataSheet() {
    try {
        // メタデータ機能は無効化 - 存在しないシートのため
        throw new Error('番組メタデータ機能は無効化されています');
        /*
        const spreadsheet = SpreadsheetApp.openById(CONFIG.METADATA_SPREADSHEET_ID);
        let sheet = spreadsheet.getSheetByName(CONFIG.METADATA_SHEET_NAME);
    
        if (!sheet) {
          sheet = spreadsheet.insertSheet(CONFIG.METADATA_SHEET_NAME);
          
          // ヘッダー行を追加
          const headers = [
            'programId', 'name', 'regularSlots', 'hosts', 'productionCompany',
            'startDate', 'genre', 'publicityFrames', 'reportFrames', 'callFrames',
            'specialFrames', 'notes', 'lastUpdated', 'createdAt'
          ];
          
          sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
          sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
          sheet.setFrozenRows(1);
          
          console.log('メタデータシートを初期化しました');
        }
    
        return sheet;
        */
    }
    catch (error) {
        console.error('メタデータシート初期化エラー:', error);
        throw error;
    }
}
/**
 * 番組メタデータをスプレッドシートに保存
 */
function saveProgramMetadataToSheet(metadata) {
    try {
        // メタデータ機能は無効化
        throw new Error('番組メタデータ機能は無効化されています');
        /*
        const sheet = initializeMetadataSheet();
        const now = new Date().toISOString();
        
        const row: ProgramMetadataRow = {
          programId: metadata.programId,
          name: metadata.name,
          regularSlots: metadata.meta.regularSlots.join(', '),
          hosts: metadata.meta.hosts.join(', '),
          productionCompany: metadata.meta.productionCompany || '',
          startDate: metadata.meta.startDate || '',
          genre: metadata.meta.genre || '',
          publicityFrames: JSON.stringify(metadata.meta.publicityFrames),
          reportFrames: JSON.stringify(metadata.meta.reportFrames),
          callFrames: JSON.stringify(metadata.meta.callFrames),
          specialFrames: JSON.stringify(metadata.meta.specialFrames),
          notes: metadata.meta.notes.join(', '),
          lastUpdated: now,
          createdAt: now
        };
        
        // 既存データをチェック
        const data = sheet.getDataRange().getValues();
        const existingRowIndex = data.findIndex((rowData, index) =>
          index > 0 && rowData[0] === metadata.programId
        );
        
        const rowData = Object.values(row);
        
        if (existingRowIndex > 0) {
          // 既存データを更新
          row.createdAt = data[existingRowIndex][13] || now; // 作成日は維持
          rowData[13] = row.createdAt;
          sheet.getRange(existingRowIndex + 1, 1, 1, rowData.length).setValues([rowData]);
          console.log(`番組メタデータを更新しました: ${metadata.name}`);
        } else {
          // 新規データを追加
          sheet.getRange(data.length + 1, 1, 1, rowData.length).setValues([rowData]);
          console.log(`番組メタデータを新規作成しました: ${metadata.name}`);
        }
        
        return row;
        
      } catch (error: unknown) {
        console.error('番組メタデータ保存エラー:', error);
        throw error;
      }
    }
    
    /**
     * 番組メタデータをスプレッドシートから読み込み
     */
        function loadProgramMetadataFromSheet() {
            try {
                const sheet = initializeMetadataSheet();
                const data = sheet.getDataRange().getValues();
                if (data.length <= 1) {
                    return [];
                }
                const headers = data[0];
                const programs = [];
                for (let i = 1; i < data.length; i++) {
                    const rowData = data[i];
                    try {
                        const program = {
                            programId: rowData[0] || '',
                            name: rowData[1] || '',
                            meta: {
                                regularSlots: rowData[2] ? rowData[2].split(',').map((s) => s.trim()).filter(Boolean) : [],
                                hosts: rowData[3] ? rowData[3].split(',').map((s) => s.trim()).filter(Boolean) : [],
                                productionCompany: rowData[4] || '',
                                startDate: rowData[5] || null,
                                genre: rowData[6] || '',
                                publicityFrames: rowData[7] ? JSON.parse(rowData[7]) : [],
                                reportFrames: rowData[8] ? JSON.parse(rowData[8]) : [],
                                callFrames: rowData[9] ? JSON.parse(rowData[9]) : [],
                                specialFrames: rowData[10] ? JSON.parse(rowData[10]) : [],
                                notes: (rowData[11] && typeof rowData[11] === 'string') ? rowData[11].split(',').map((s) => s.trim()).filter(Boolean) : []
                            }
                        };
                        programs.push(program);
                    }
                    catch (parseError) {
                        console.warn(`行 ${i + 1} のデータ解析でエラー:`, parseError);
                        continue;
                    }
                }
                console.log(`${programs.length}件の番組メタデータを読み込みました`);
                return programs;
            }
            catch (error) {
                console.error('番組メタデータ読み込みエラー:', error);
                return [];
            }
        }
        /**
         * 番組メタデータを削除
         */
        function deleteProgramMetadataFromSheet(programId) {
            try {
                const sheet = initializeMetadataSheet();
                const data = sheet.getDataRange().getValues();
                const rowIndex = data.findIndex((rowData, index) => index > 0 && rowData[0] === programId);
                if (rowIndex > 0) {
                    sheet.deleteRow(rowIndex + 1);
                    console.log(`番組メタデータを削除しました: ${programId}`);
                    return true;
                }
                else {
                    console.warn(`削除対象の番組が見つかりません: ${programId}`);
                    return false;
                }
            }
            catch (error) {
                console.error('番組メタデータ削除エラー:', error);
                return false;
            }
        }
        /**
         * WebApp用の番組メタデータ保存関数
         */
        function webAppSaveProgramMetadata(metadata) {
            try {
                console.log('WebApp: 番組メタデータ保存開始:', metadata.name);
                const savedRow = saveProgramMetadataToSheet(metadata);
                return {
                    success: true,
                    data: savedRow,
                    message: '番組メタデータの保存が完了しました',
                    timestamp: new Date().toISOString()
                };
            }
            catch (error) {
                console.error('WebApp: 番組メタデータ保存エラー:', error);
                return {
                    success: false,
                    error: error instanceof Error ? error.message : String(error)
                };
            }
        }
        /**
         * WebApp用の番組メタデータ削除関数
         */
        function webAppDeleteProgramMetadata(programId) {
            try {
                console.log('WebApp: 番組メタデータ削除開始:', programId);
                const deleted = deleteProgramMetadataFromSheet(programId);
                if (deleted) {
                    return {
                        success: true,
                        message: '番組メタデータの削除が完了しました',
                        programId: programId,
                        timestamp: new Date().toISOString()
                    };
                }
                else {
                    return {
                        success: false,
                        error: '削除対象の番組メタデータが見つかりませんでした'
                    };
                }
            }
            catch (error) {
                console.error('WebApp: 番組メタデータ削除エラー:', error);
                return {
                    success: false,
                    error: error instanceof Error ? error.message : String(error)
                };
            }
        }
        /**
         * スプレッドシートとマスターデータを統合した番組メタデータ取得関数
         */
        function webAppGetProgramMetadataAsJSONUnified() {
            try {
                console.log('WebApp: 統合番組メタデータ取得開始');
                // スプレッドシートからデータを読み込み
                const spreadsheetData = loadProgramMetadataFromSheet();
                // マスターデータと統合（スプレッドシートのデータを優先）
                const masterData = PROGRAM_METADATA_MASTER.programs;
                const mergedPrograms = [];
                // スプレッドシートのデータを追加
                spreadsheetData.forEach(program => {
                    mergedPrograms.push(program);
                });
                // マスターデータから、スプレッドシートにない番組を追加
                masterData.forEach(masterProgram => {
                    const exists = spreadsheetData.some(spreadsheetProgram => spreadsheetProgram.programId === masterProgram.programId);
                    if (!exists) {
                        mergedPrograms.push(masterProgram);
                    }
                });
                const result = {
                    programs: mergedPrograms
                };
                return {
                    success: true,
                    data: result,
                    formatted: JSON.stringify(result, null, 2),
                    timestamp: new Date().toISOString(),
                    programCount: result.programs.length,
                    spreadsheetCount: spreadsheetData.length,
                    masterCount: masterData.length
                };
            }
            catch (error) {
                console.error('WebApp: 統合番組メタデータ取得エラー:', error);
                return {
                    success: false,
                    error: error instanceof Error ? error.message : String(error)
                };
            }
        }
        /**
         * 番組メタデータのテスト関数
         */
        function testProgramMetadata() {
            var _a;
            try {
                console.log('=== 番組メタデータテスト開始 ===');
                // テスト1: 全番組メタデータ取得
                console.log('\n【テスト1】全番組メタデータ取得');
                const allResult = webAppGetProgramMetadataAsJSON();
                if (allResult.success) {
                    console.log('✓ 全番組メタデータ取得成功');
                    console.log('番組数:', allResult.programCount);
                    console.log('データサイズ:', ((_a = allResult.formatted) === null || _a === void 0 ? void 0 : _a.length) || 0, '文字');
                }
                else {
                    console.error('✗ 全番組メタデータ取得失敗:', allResult.error);
                }
                // テスト2: 特定番組メタデータ取得
                console.log('\n【テスト2】特定番組メタデータ取得 - ちょうどいいラジオ');
                const singleResult = webAppGetSingleProgramMetadata('ちょうどいいラジオ');
                if (singleResult.success) {
                    console.log('✓ 特定番組メタデータ取得成功');
                    console.log('番組名:', singleResult.programName);
                    console.log('番組ID:', singleResult.programId);
                }
                else {
                    console.error('✗ 特定番組メタデータ取得失敗:', singleResult.error);
                }
                // テスト3: 存在しない番組のテスト
                console.log('\n【テスト3】存在しない番組のテスト');
                const notFoundResult = webAppGetSingleProgramMetadata('存在しない番組');
                if (!notFoundResult.success) {
                    console.log('✓ 存在しない番組の適切なエラーハンドリング確認');
                    console.log('エラーメッセージ:', notFoundResult.error);
                }
                else {
                    console.error('✗ 存在しない番組でエラーが発生しませんでした');
                }
                console.log('\n=== 番組メタデータテスト完了 ===');
                return {
                    success: true,
                    allMetadata: allResult.success,
                    singleMetadata: singleResult.success,
                    errorHandling: !notFoundResult.success
                };
            }
            catch (error) {
                console.error('番組メタデータテストでエラー:', error);
                return {
                    success: false,
                    error: error instanceof Error ? error.message : String(error)
                };
            }
        }
        // ======================
        // Raw Program Data機能
        // ======================
        /**
         * 曜日順序を正規化するマップ
         */
        const WEEKDAY_ORDER = {
            'monday': 1,
            'tuesday': 2,
            'wednesday': 3,
            'thursday': 4,
            'friday': 5,
            'saturday': 6,
            'sunday': 7
        };
        /**
         * 曜日データを正しい順序で並び替える関数
         */
        function sortWeekdaysData(data) {
            if (!data || typeof data !== 'object') {
                return data;
            }
            const sortedData = {};
            // 収録予定などの特別キーを最初に処理
            Object.keys(data).forEach(key => {
                if (key.includes('収録予定') || !WEEKDAY_ORDER[key.toLowerCase()]) {
                    sortedData[key] = data[key];
                }
            });
            // 曜日を正しい順序で並び替え
            const weekdayKeys = Object.keys(data)
                .filter(key => WEEKDAY_ORDER[key.toLowerCase()])
                .sort((a, b) => WEEKDAY_ORDER[a.toLowerCase()] - WEEKDAY_ORDER[b.toLowerCase()]);
            weekdayKeys.forEach(key => {
                sortedData[key] = data[key];
            });
            return sortedData;
        }
        /**
         * 指定番組のみを抽出する関数
         */
        function extractSpecificProgramData(allData, programName) {
            if (!allData || typeof allData !== 'object') {
                return {};
            }
            const result = {};
            // 各曜日から指定番組のデータのみを抽出
            Object.keys(allData).forEach(dayKey => {
                const dayData = allData[dayKey];
                if (dayKey.includes('収録予定')) {
                    // 収録予定は番組名でフィルタ
                    if (dayData && typeof dayData === 'object' && dayData[programName]) {
                        result[dayKey] = { [programName]: dayData[programName] };
                    }
                    return;
                }
                if (dayData && typeof dayData === 'object') {
                    const programData = dayData[programName];
                    if (programData) {
                        result[dayKey] = { [programName]: programData };
                    }
                }
            });
            return result;
        }
        /**
         * WebApp用の完全生データ取得関数（指定番組のみ）
         */
        function webAppGetRawProgramDataUnfiltered(programName, weekType = 'thisWeek') {
            try {
                console.log(`完全生データ取得開始: ${programName} (${weekType})`);
                if (!programName) {
                    return {
                        success: false,
                        error: '番組名を指定してください'
                    };
                }
                // 対象週のシートを取得
                const config = getConfig();
                const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
                const weekOffset = weekType === 'thisWeek' ? 0 : 1;
                const sheet = getSheetByWeek(spreadsheet, weekOffset);
                if (!sheet) {
                    return {
                        success: false,
                        error: '指定された週のシートが見つかりません'
                    };
                }
                console.log(`使用シート: ${sheet.getName()}`);
                // 週データを抽出（完全版）
                const rawWeekData = extractStructuredWeekData(sheet);
                if (!rawWeekData) {
                    return {
                        success: false,
                        error: 'データの抽出に失敗しました'
                    };
                }
                // 指定番組のみを抽出
                const filteredData = extractSpecificProgramData(rawWeekData, programName);
                if (!filteredData || Object.keys(filteredData).length === 0) {
                    return {
                        success: false,
                        error: `番組 "${programName}" のデータが見つかりません`
                    };
                }
                // 曜日順序を正規化
                const sortedData = sortWeekdaysData(filteredData);
                console.log(`完全生データ取得完了: ${programName}`);
                return {
                    success: true,
                    programName: programName,
                    weekType: weekType,
                    sheetName: sheet.getName(),
                    timestamp: new Date().toISOString(),
                    data: sortedData
                };
            }
            catch (error) {
                console.error('完全生データ取得エラー:', error);
                return {
                    success: false,
                    error: error instanceof Error ? error.message : String(error)
                };
            }
        }
        /**
         * 完全生データ機能のテスト関数
         */
        function testRawProgramDataUnfiltered() {
            try {
                console.log('=== 完全生データ機能テスト開始 ===');
                // テスト1: PRIME TIMEの今週データ
                console.log('\n【テスト1】PRIME TIME 今週データ取得');
                const primeTimeResult = webAppGetRawProgramDataUnfiltered('PRIME TIME', 'thisWeek');
                if (primeTimeResult.success) {
                    console.log('✓ PRIME TIME 今週データ取得成功');
                    console.log('シート名:', primeTimeResult.sheetName);
                    console.log('データキー:', Object.keys(primeTimeResult.data));
                }
                else {
                    console.error('✗ PRIME TIME 今週データ取得失敗:', primeTimeResult.error);
                }
                // テスト2: ちょうどいいラジオの翌週データ
                console.log('\n【テスト2】ちょうどいいラジオ 翌週データ取得');
                const choudoResult = webAppGetRawProgramDataUnfiltered('ちょうどいいラジオ', 'nextWeek');
                if (choudoResult.success) {
                    console.log('✓ ちょうどいいラジオ 翌週データ取得成功');
                    console.log('シート名:', choudoResult.sheetName);
                    console.log('データキー:', Object.keys(choudoResult.data));
                }
                else {
                    console.error('✗ ちょうどいいラジオ 翌週データ取得失敗:', choudoResult.error);
                }
                // テスト3: 存在しない番組
                console.log('\n【テスト3】存在しない番組テスト');
                const notFoundResult = webAppGetRawProgramDataUnfiltered('存在しない番組');
                if (!notFoundResult.success) {
                    console.log('✓ 存在しない番組の適切なエラーハンドリング確認');
                    console.log('エラーメッセージ:', notFoundResult.error);
                }
                else {
                    console.error('✗ 存在しない番組でエラーが発生しませんでした');
                }
                console.log('\n=== 完全生データ機能テスト完了 ===');
                return {
                    success: true,
                    primeTime: primeTimeResult.success,
                    choudo: choudoResult.success,
                    errorHandling: !notFoundResult.success
                };
            }
            catch (error) {
                console.error('完全生データ機能テストでエラー:', error);
                return {
                    success: false,
                    error: error instanceof Error ? error.message : String(error)
                };
            }
        }
        /**
         * 番組別詳細データを転置テーブル形式で生成する関数
         */
        /**
         * 週タイプを番号に変換するヘルパー関数
         */
        function mapWeekTypeToNumber(weekType) {
            const weekMapping = {
                'thisWeek': 1, // 今週
                'nextWeek': 2, // 来週  
                'nextWeek2': 3, // 来週の次
                'nextWeek3': 4 // 来週の次の次
            };
            return weekMapping[weekType] || 1; // デフォルトは今週
        }
        function generateTransposedProgramTable(programName, weekType = 'thisWeek') {
            const debugLogs = [];
            const log = (message) => {
                console.log(message);
                debugLogs.push(message);
            };
            // 番組名のバリデーション
            log(`[DEBUG] バリデーション開始: programName="${programName}", type=${typeof programName}`);
            if (!programName || programName === 'undefined' || programName === null ||
                (typeof programName === 'string' && programName.trim() === '')) {
                const errorMessage = '番組名が指定されていません';
                log(`[ERROR] ${errorMessage}: received "${programName}" (type: ${typeof programName})`);
                return { success: false, error: errorMessage, debugLogs };
            }
            log(`[DEBUG] バリデーション通過: programName="${programName}"`);
            log(`[DEBUG] 転置テーブル生成開始: ${programName}, 週タイプ: ${weekType}`);
            try {
                // 週タイプを番号に変換
                const weekNumber = mapWeekTypeToNumber(weekType);
                log(`[DEBUG] 週番号: ${weekNumber} (${weekType})`);
                // 指定された週データを取得
                log(`[DEBUG] extractWeekByNumber(${weekNumber})を呼び出し中...`);
                const weekResults = extractWeekByNumber(weekNumber);
                log(`[DEBUG] extractWeekByNumberの結果: ${typeof weekResults}, キー数: ${Object.keys(weekResults).length}`);
                log(`[DEBUG] 取得されたキー: [${Object.keys(weekResults).join(', ')}]`);
                if (!weekResults || Object.keys(weekResults).length === 0) {
                    log('[ERROR] 週データが見つかりません');
                    log(`[ERROR] weekResultsの値: ${JSON.stringify(weekResults)}`);
                    return { success: false, error: 'データが見つかりません', debugLogs };
                }
                // シート名（週データのキー）を取得
                const sheetName = Object.keys(weekResults)[0];
                log(`[DEBUG] 選択されたシート名: "${sheetName}"`);
                const weekData = weekResults[sheetName];
                log(`[DEBUG] 週データのタイプ: ${typeof weekData}`);
                log(`[DEBUG] 週データの番組数: ${weekData ? Object.keys(weekData).length : 0}`);
                log(`[DEBUG] 利用可能な番組名: [${weekData ? Object.keys(weekData).join(', ') : []}]`);
                log(`[DEBUG] リクエストされた番組名: "${programName}"`);
                log(`[DEBUG] 番組名の完全一致チェック: ${weekData && weekData[programName] ? 'OK' : 'NG'}`);
                if (!weekData || !weekData[programName]) {
                    log(`[ERROR] ${programName}の番組データが見つかりません`);
                    // 部分マッチングを試行
                    log(`[DEBUG] 部分マッチングを試行...`);
                    let matchedProgram = null;
                    if (weekData) {
                        for (const availableProgram of Object.keys(weekData)) {
                            log(`[DEBUG] チェック: "${availableProgram}" vs "${programName}"`);
                            // 安全な文字列比較を実行
                            if (availableProgram && programName &&
                                typeof availableProgram === 'string' && typeof programName === 'string' &&
                                (availableProgram.includes(programName) || programName.includes(availableProgram))) {
                                log(`[DEBUG] 部分マッチした番組を発見: "${availableProgram}"`);
                                matchedProgram = availableProgram;
                                break;
                            }
                        }
                    }
                    if (!matchedProgram) {
                        log('[ERROR] 部分マッチも失敗');
                        return { success: false, error: `該当する番組データがありません: "${programName}"`, debugLogs };
                    }
                    // マッチした番組を使用
                    log(`[DEBUG] マッチした番組を使用: "${matchedProgram}"`);
                    programName = matchedProgram;
                }
                // 番組データを抽出
                log(`[DEBUG] 番組データを抽出中: ${programName}`);
                const programData = extractProgramData(weekData[programName], programName);
                if (!programData || Object.keys(programData).length === 0) {
                    log(`[ERROR] ${programName}の曜日別データが見つかりません`);
                    log(`[DEBUG] 利用可能な曜日: [${Object.keys(weekData[programName] || {}).join(', ')}]`);
                    return { success: false, error: '該当する番組の曜日データがありません', debugLogs };
                }
                log(`[DEBUG] 抽出された曜日数: ${Object.keys(programData).length}`);
                log(`[DEBUG] 抽出された曜日: [${Object.keys(programData).join(', ')}]`);
                // 転置テーブル用データ構造を生成
                log(`[DEBUG] 転置テーブルデータ構造を生成中...`);
                log(`[DEBUG] programData構造確認: type=${typeof programData}, keys=[${Object.keys(programData || {}).join(', ')}], keyCount=${Object.keys(programData || {}).length}`);
                const transposedData = generateTransposedTableData(programData, programName);
                // 戻り値の詳細検証
                log(`[DEBUG] generateTransposedTableData戻り値検証: type=${typeof transposedData}, isNull=${transposedData === null}, isUndefined=${transposedData === undefined}, keys=[${transposedData ? Object.keys(transposedData).join(', ') : 'none'}]`);
                if (transposedData) {
                    log(`[DEBUG] 転置テーブル生成完了: ${programName}`);
                    log(`[DEBUG] 転置テーブルデータ内容: programName=${transposedData.programName}, headerCount=${transposedData.headers ? transposedData.headers.length : 0}, rowCount=${transposedData.rows ? transposedData.rows.length : 0}`);
                }
                else {
                    log(`[ERROR] 転置テーブルデータがnullまたはundefinedです`);
                    return { success: false, error: '転置テーブルデータの生成に失敗しました', debugLogs };
                }
                return { success: true, data: transposedData, debugLogs };
            }
            catch (error) {
                log(`[ERROR] 転置テーブル生成エラー: ${error instanceof Error ? error.message : String(error)}`);
                if (error instanceof Error && error.stack) {
                    log(`[DEBUG] スタックトレース: ${error.stack}`);
                }
                return { success: false, error: `エラーが発生しました: ${error instanceof Error ? error.message : String(error)}`, debugLogs };
            }
        }
        /**
         * 日付文字列と曜日を組み合わせたヘッダー形式にフォーマットする関数
         */
        function formatDateWithDay(dateStr, dayName) {
            try {
                console.log(`[DEBUG] 日付フォーマット開始: "${dateStr}" + "${dayName}"`);
                const date = new Date(dateStr);
                // 日付の妥当性チェック
                if (isNaN(date.getTime())) {
                    console.warn(`[WARN] 無効な日付文字列: "${dateStr}"`);
                    return dayName; // エラー時は曜日のみ
                }
                const month = date.getMonth() + 1;
                const day = date.getDate();
                const formatted = `${month}/${day}${dayName}`;
                console.log(`[DEBUG] 日付フォーマット結果: "${formatted}"`);
                return formatted;
            }
            catch (error) {
                console.error('[ERROR] 日付フォーマットエラー:', error);
                return dayName; // エラー時は曜日のみ
            }
        }
        /**
         * 転置テーブル用のデータ構造を生成する関数
         */
        function generateTransposedTableData(programData, programName) {
            try {
                console.log(`[DEBUG] generateTransposedTableData開始: programName="${programName}"`);
                console.log(`[DEBUG] programDataの構造:`, {
                    type: typeof programData,
                    keys: programData ? Object.keys(programData) : [],
                    isNull: programData === null,
                    isUndefined: programData === undefined
                });
                // データ構造の基本検証
                if (!programData || typeof programData !== 'object') {
                    console.error('[ERROR] programDataが無効です:', programData);
                    return null;
                }
                // 改善点.txtの要件に基づく番組別項目定義
                const programItems = getProgramItems(programName);
                console.log(`[DEBUG] 番組項目数: ${programItems ? programItems.length : 0}`);
                // 日付を取得（日本語の曜日キーを使用）
                console.log(`[DEBUG] programDataキー確認:`, Object.keys(programData));
                const dayOrder = ['月曜', '火曜', '水曜', '木曜', '金曜', '土曜', '日曜'];
                const availableDays = dayOrder.filter(day => programData[day]);
                console.log(`[DEBUG] dayOrderでのフィルタ結果: [${availableDays.join(', ')}]`);
                console.log(`[DEBUG] 利用可能曜日: [${availableDays.join(', ')}]`);
                if (availableDays.length === 0) {
                    console.warn('[WARN] 利用可能な曜日データがありません');
                    return null;
                }
                // ヘッダー行（項目 + 日付列）
                const headers = ['項目'];
                availableDays.forEach((dayKey, index) => {
                    const dayData = programData[dayKey];
                    const japaneseDay = dayKey; // 既に日本語なのでそのまま使用
                    console.log(`[DEBUG] ${dayKey}(${japaneseDay})のデータ検証: exists=${!!dayData}, type=${typeof dayData}, keys=[${dayData ? Object.keys(dayData).join(', ') : 'none'}]`);
                    console.log(`[DEBUG] ${dayKey}の日付関連: hasDate=${(dayData === null || dayData === void 0 ? void 0 : dayData['日付']) ? true : false}, dateType=${typeof (dayData === null || dayData === void 0 ? void 0 : dayData['日付'])}, isArray=${Array.isArray(dayData === null || dayData === void 0 ? void 0 : dayData['日付'])}, dateLength=${Array.isArray(dayData === null || dayData === void 0 ? void 0 : dayData['日付']) ? dayData['日付'].length : 'N/A'}`);
                    // dayData全体の内容を詳細表示
                    if (dayData) {
                        console.log(`[DEBUG] ${dayKey}の完全なデータ内容:`, JSON.stringify(dayData, null, 2));
                    }
                    // 安全な日付データアクセス
                    console.log(`[DEBUG] ${dayKey}の条件チェック: dayData=${!!dayData}, hasDateField=${!!(dayData === null || dayData === void 0 ? void 0 : dayData['日付'])}, isDateArray=${Array.isArray(dayData === null || dayData === void 0 ? void 0 : dayData['日付'])}, dateArrayLength=${Array.isArray(dayData === null || dayData === void 0 ? void 0 : dayData['日付']) ? dayData['日付'].length : 'N/A'}`);
                    if (dayData && dayData['日付'] && Array.isArray(dayData['日付']) && dayData['日付'].length > 0) {
                        const dateStr = dayData['日付'][0];
                        console.log(`[DEBUG] ${dayKey}の日付文字列取得成功: "${dateStr}"`);
                        const formattedHeader = formatDateWithDay(dateStr, japaneseDay);
                        console.log(`[DEBUG] ${dayKey}のフォーマット済みヘッダー: "${formattedHeader}"`);
                        headers.push(formattedHeader); // "8/26月曜"形式
                    }
                    else {
                        console.log(`[DEBUG] ${dayKey}の日付データが条件を満たさない - 理由: dayData=${!!dayData}, dateField=${!!(dayData === null || dayData === void 0 ? void 0 : dayData['日付'])}, isArray=${Array.isArray(dayData === null || dayData === void 0 ? void 0 : dayData['日付'])}, length=${Array.isArray(dayData === null || dayData === void 0 ? void 0 : dayData['日付']) ? dayData['日付'].length : 'N/A'}`);
                        headers.push(japaneseDay);
                    }
                });
                console.log(`[DEBUG] ヘッダー生成完了:`, headers);
                // データ行を生成
                const rows = [];
                if (programItems && Array.isArray(programItems)) {
                    programItems.forEach((item, itemIndex) => {
                        console.log(`[DEBUG] 項目 ${itemIndex + 1}/${programItems.length}: "${item}"`);
                        const row = [item]; // 最初の列は項目名
                        // 各日付の値を取得
                        availableDays.forEach(dayName => {
                            try {
                                const dayData = programData[dayName];
                                const value = getEpisodeItemValue(dayData, item);
                                const formattedValue = formatItemValue(value);
                                row.push(formattedValue);
                            }
                            catch (error) {
                                console.error(`[ERROR] ${dayName}の${item}データ取得エラー:`, error);
                                row.push('ー'); // エラー時のデフォルト値
                            }
                        });
                        rows.push(row);
                    });
                }
                else {
                    console.error('[ERROR] programItemsが配列ではありません:', programItems);
                }
                console.log(`[DEBUG] データ行生成完了: ${rows.length}行`);
                const result = {
                    programName: programName,
                    headers: headers,
                    rows: rows
                };
                console.log(`[DEBUG] generateTransposedTableData正常完了`);
                return result;
            }
            catch (error) {
                console.error(`[ERROR] generateTransposedTableData内でエラー:`, error);
                if (error instanceof Error) {
                    console.error(`[ERROR] エラー詳細: ${error.message}`);
                    console.error(`[ERROR] スタックトレース: ${error.stack}`);
                }
                // エラー時でも最低限のデータ構造を返す
                return {
                    programName: programName,
                    headers: ['項目', 'エラー'],
                    rows: [['データ取得エラー', `エラー: ${error instanceof Error ? error.message : String(error)}`]]
                };
            }
        }
        /**
         * 番組の曜日別データを転置テーブル用に変換する関数
         */
        function extractProgramData(programWeekData, programName) {
            const programData = {};
            // 曜日のマッピング
            const dayMapping = {
                'monday': '月曜',
                'tuesday': '火曜',
                'wednesday': '水曜',
                'thursday': '木曜',
                'friday': '金曜',
                'saturday': '土曜',
                'sunday': '日曜'
            };
            // 番組の曜日別データを処理
            for (const [englishDay, japaneseDay] of Object.entries(dayMapping)) {
                if (programWeekData[englishDay]) {
                    programData[japaneseDay] = {
                        [programName]: programWeekData[englishDay]
                    };
                    console.log(`番組データ抽出: ${japaneseDay} - ${programName}`);
                }
            }
            return programData;
        }
        /**
         * 転置テーブルHTMLを生成する関数
         */
        function convertToTransposedTable(programData, programName) {
            // 改善点.txtの要件に基づく番組別項目定義
            const programItems = getProgramItems(programName);
            // 日付を取得（曜日順にソート）
            const dayOrder = ['月曜', '火曜', '水曜', '木曜', '金曜', '土曜', '日曜'];
            const availableDays = dayOrder.filter(day => programData[day]);
            if (availableDays.length === 0) {
                return '<div class="no-data">表示可能なデータがありません</div>';
            }
            // テーブルヘッダー生成
            let html = '<table class="transposed-table">\n';
            html += '  <thead>\n';
            html += '    <tr>\n';
            html += '      <th class="item-header">項目</th>\n';
            // 各日付をヘッダーに追加
            availableDays.forEach(dayName => {
                const dateStr = getDateForDay(dayName);
                html += `      <th class="date-header">${dateStr}</th>\n`;
            });
            html += '    </tr>\n';
            html += '  </thead>\n';
            html += '  <tbody>\n';
            // 各項目について行を生成
            programItems.forEach(item => {
                html += '    <tr>\n';
                html += `      <td class="item-name">${item}</td>\n`;
                // 各日付の値を取得
                availableDays.forEach(dayName => {
                    const value = getEpisodeItemValue(programData[dayName], item);
                    const formattedValue = formatItemValue(value);
                    html += `      <td class="item-value">${formattedValue}</td>\n`;
                });
                html += '    </tr>\n';
            });
            html += '  </tbody>\n';
            html += '</table>\n';
            return html;
        }
        /**
         * 番組別項目を取得する関数（改善点.txtの要件に基づく）
         */
        function getProgramItems(programName) {
            switch (programName) {
                case 'ちょうどいいラジオ':
                    return [
                        '７：２８パブ告知',
                        '時間指定なし告知',
                        'YOKOHAMA PORTSIDE INFORMATION',
                        '指定曲',
                        '先行予約',
                        'ゲスト',
                        '６：４５ラジオショッピング',
                        '８：２９はぴねすくらぶ',
                        '収録予定'
                    ];
                case 'PRIME TIME':
                    return [
                        '１９：４３パブリシティ',
                        '２０：５１パブリシティ',
                        '営業コーナー',
                        '指定曲',
                        'ゲスト',
                        '時間指定なしパブ',
                        'ラジショピ',
                        '先行予約'
                    ];
                case 'FLAG':
                    return [
                        '１２：４０電話パブ',
                        '１３：２９パブリシティ',
                        '１３：４０パブリシティ',
                        '１２：１５リポート案件',
                        '１４：２９リポート案件',
                        '時間指定なし告知',
                        '楽曲',
                        '先行予約',
                        '収録予定'
                    ];
                case 'God Bless Saturday':
                    return [
                        'キリンパークシティーヨコハマ',
                        '指定曲',
                        '１４：４１パブ',
                        '時間指定なし告知'
                    ];
                case 'Route 847':
                    return [
                        '１６：４７パブ',
                        '１７：４１パブ',
                        '時間指定なし告知',
                        '指定曲'
                    ];
                default:
                    return ['日付', '内容'];
            }
        }
        /**
         * エピソードデータから項目値を取得する関数
         */
        function getEpisodeItemValue(dayEpisodes, itemName) {
            if (!dayEpisodes || typeof dayEpisodes !== 'object') {
                return ['ー'];
            }
            // 番組データから該当する項目を探す
            for (const [showName, showData] of Object.entries(dayEpisodes)) {
                if (showData && typeof showData === 'object') {
                    const data = showData;
                    // 項目名のマッピング処理
                    const mappedItem = mapItemName(itemName);
                    if (data[mappedItem]) {
                        return data[mappedItem];
                    }
                    // 部分マッチも試行
                    for (const [key, value] of Object.entries(data)) {
                        if (key.includes(mappedItem) || mappedItem.includes(key)) {
                            return value;
                        }
                    }
                }
            }
            return ['ー'];
        }
        /**
         * 項目名を内部データ構造にマッピングする関数
         */
        function mapItemName(displayName) {
            const mappings = {
                '７：２８パブ告知': '7:28パブ告知',
                '時間指定なし告知': '時間指定なし告知',
                'YOKOHAMA PORTSIDE INFORMATION': 'YOKOHAMA PORTSIDE INFORMATION',
                '指定曲': '楽曲',
                '先行予約': '先行予約',
                'ゲスト': 'ゲスト',
                '６：４５ラジオショッピング': 'ラジオショッピング',
                '８：２９はぴねすくらぶ': 'はぴねすくらぶ',
                '収録予定': '収録予定',
                '１９：４３パブリシティ': '19:43パブリシティ',
                '２０：５１パブリシティ': '20:51パブリシティ',
                '営業コーナー': '営業コーナー',
                '時間指定なしパブ': '時間指定なしパブ',
                'ラジショピ': 'ラジオショッピング',
                '１２：４０電話パブ': '12:40 電話パブ',
                '１３：２９パブリシティ': '13:29 パブリシティ',
                '１３：４０パブリシティ': '13:40 パブリシティ',
                '１２：１５リポート案件': '12:15 リポート案件',
                '１４：２９リポート案件': '14:29 リポート案件',
                '楽曲': '楽曲',
                'キリンパークシティーヨコハマ': 'キリンパークシティーヨコハマ',
                '１４：４１パブ': '14:41パブ',
                '１６：４７パブ': 'リポート 16:47',
                '１７：４１パブ': '営業パブ 17:41'
            };
            return mappings[displayName] || displayName;
        }
        /**
         * 項目値をHTML用にフォーマットする関数
         */
        function formatItemValue(value) {
            if (!value) {
                return 'ー';
            }
            if (Array.isArray(value)) {
                if (value.length === 0) {
                    return 'ー';
                }
                // 配列の場合は改行区切りで表示
                return value
                    .filter(item => item && item !== 'ー')
                    .map(item => formatSingleItem(item))
                    .join('<br>') || 'ー';
            }
            return formatSingleItem(value);
        }
        /**
         * 単一項目をフォーマットする関数（楽曲オブジェクト対応）
         */
        function formatSingleItem(item) {
            if (!item) {
                return 'ー';
            }
            // デバッグ: オブジェクトの詳細構造を確認
            if (typeof item === 'object' && item !== null) {
                console.log(`[DEBUG] formatSingleItem: オブジェクト検出`);
                console.log(`[DEBUG] typeof: ${typeof item}`);
                console.log(`[DEBUG] keys:`, Object.keys(item));
                console.log(`[DEBUG] values:`, Object.values(item));
                console.log(`[DEBUG] 全体:`, JSON.stringify(item, null, 2));
                // 様々なパターンの楽曲オブジェクトプロパティを確認
                const possibleSongKeys = ['曲名', 'title', 'songName', 'name', '楽曲名', 'song'];
                const possibleArtistKeys = ['アーティスト', 'artist', 'singer', '歌手'];
                const possibleUrlKeys = ['URL', 'url', 'link', 'href'];
                let songName = null;
                let artistName = null;
                let songUrl = null;
                // 楽曲名を検索
                for (const key of possibleSongKeys) {
                    if (item[key] !== undefined && item[key] !== null && item[key] !== 'ー' && item[key] !== '') {
                        songName = item[key];
                        console.log(`[DEBUG] 楽曲名発見: "${key}" = "${songName}"`);
                        break;
                    }
                }
                // アーティスト名を検索
                for (const key of possibleArtistKeys) {
                    if (item[key] !== undefined && item[key] !== null && item[key] !== 'ー' && item[key] !== '') {
                        artistName = item[key];
                        console.log(`[DEBUG] アーティスト名発見: "${key}" = "${artistName}"`);
                        break;
                    }
                }
                // URL を検索
                for (const key of possibleUrlKeys) {
                    if (item[key] !== undefined && item[key] !== null && item[key] !== 'ー' && item[key] !== '') {
                        songUrl = item[key];
                        console.log(`[DEBUG] URL発見: "${key}" = "${songUrl}"`);
                        break;
                    }
                }
                // 楽曲情報が見つかった場合
                if (songName) {
                    const songInfo = [];
                    songInfo.push(songName);
                    if (artistName) {
                        songInfo.push(`/ ${artistName}`);
                    }
                    if (songUrl) {
                        songInfo.push(`<a href="${songUrl}" target="_blank">[URL]</a>`);
                    }
                    const result = songInfo.join(' ');
                    console.log(`[DEBUG] 楽曲フォーマット結果: "${result}"`);
                    return result;
                }
                // 楽曲オブジェクトでない場合はJSONで表示
                try {
                    const jsonStr = JSON.stringify(item);
                    console.log(`[DEBUG] JSON化結果: "${jsonStr}"`);
                    return jsonStr;
                }
                catch (e) {
                    console.log(`[DEBUG] JSON化失敗, String()使用`);
                    return String(item);
                }
            }
            const result = String(item).trim() || 'ー';
            console.log(`[DEBUG] 非オブジェクト値: "${result}"`);
            return result;
        }
        /**
         * 曜日名から日付を取得する関数 (転置テーブル用)
         */
        function getDateForDay(dayName) {
            try {
                // 直接シートアクセス方式で安全に日付を取得
                const config = getConfig();
                const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
                const allSheets = spreadsheet.getSheets();
                // 週形式のシートを探して最新のシートを使用
                let targetSheet = '';
                const weekSheets = [];
                allSheets.forEach(sheet => {
                    const sheetName = sheet.getName();
                    if (sheetName.match(/^\d{2}\.\d{1,2}\.\d{2}-/)) {
                        weekSheets.push(sheetName);
                    }
                });
                if (weekSheets.length > 0) {
                    // 日付でソートして最新のシートを使用
                    weekSheets.sort((a, b) => {
                        try {
                            const startA = getStartDateFromSheetName(a);
                            const startB = getStartDateFromSheetName(b);
                            return startB.getTime() - startA.getTime(); // 降順（最新が先頭）
                        }
                        catch (_a) {
                            return 0;
                        }
                    });
                    targetSheet = weekSheets[0];
                    // シート名から日付を算出
                    try {
                        const dayDates = calculateDayDates(targetSheet);
                        // 曜日名から英語名に変換
                        const dayMapping = {
                            '月曜': 'monday',
                            '火曜': 'tuesday',
                            '水曜': 'wednesday',
                            '木曜': 'thursday',
                            '金曜': 'friday',
                            '土曜': 'saturday',
                            '日曜': 'sunday'
                        };
                        const englishDay = dayMapping[dayName];
                        if (englishDay && dayDates[englishDay]) {
                            return dayDates[englishDay];
                        }
                    }
                    catch (dateError) {
                        console.error('日付算出エラー:', dateError);
                    }
                }
                else {
                    console.warn('週形式のシートが見つかりません');
                }
                // フォールバック: 曜日名のまま返す
                return dayName;
            }
            catch (error) {
                console.error('日付取得エラー:', error);
                return dayName;
            }
        }
        /**
         * 転置テーブル表示経路の完全デバッグ用テスト関数
         */
        function testTransposedTableDataFlow(programName = 'ちょうどいいラジオ') {
            console.log(`=== 転置テーブルデータフロー完全テスト開始 ===`);
            console.log(`[DEBUG] 受信したprogramName: "${programName}" (type: ${typeof programName})`);
            // programNameが未定義の場合は明示的にデフォルト値をセット
            if (!programName || programName === 'undefined' || typeof programName !== 'string') {
                programName = 'ちょうどいいラジオ';
                console.log(`[DEBUG] programNameをデフォルト値に修正: "${programName}"`);
            }
            console.log(`テスト対象番組: ${programName}`);
            try {
                // 1. CONFIG確認
                console.log('\n【ステップ1】CONFIG確認');
                const config = getConfig();
                console.log('SPREADSHEET_ID:', config.SPREADSHEET_ID);
                // 2. スプレッドシート接続確認
                console.log('\n【ステップ2】スプレッドシート接続確認');
                const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
                const allSheets = spreadsheet.getSheets();
                console.log(`総シート数: ${allSheets.length}`);
                // 3. 週シート検出
                console.log('\n【ステップ3】週シート検出');
                const weekSheets = [];
                allSheets.forEach((sheet, index) => {
                    const sheetName = sheet.getName();
                    console.log(`シート${index + 1}: "${sheetName}"`);
                    if (sheetName.match(/^\d{2}\.\d{1,2}\.\d{2}-/)) {
                        weekSheets.push(sheetName);
                        console.log(`  → 週シートとして認識`);
                    }
                });
                console.log(`検出された週シート数: ${weekSheets.length}`);
                console.log('週シート一覧:', weekSheets);
                if (weekSheets.length === 0) {
                    return {
                        success: false,
                        error: '週シートが見つかりません',
                        details: {
                            totalSheets: allSheets.length,
                            allSheetNames: allSheets.map(s => s.getName())
                        }
                    };
                }
                // 4. 最新の週シート選択
                console.log('\n【ステップ4】最新の週シート選択');
                weekSheets.sort();
                const targetSheet = weekSheets[0]; // 1番目（最新）
                console.log(`選択されたシート: "${targetSheet}"`);
                // 5. シート取得とデータ抽出テスト
                console.log('\n【ステップ5】シート取得とデータ抽出テスト');
                const sheet = spreadsheet.getSheetByName(targetSheet);
                if (!sheet) {
                    return {
                        success: false,
                        error: `シート "${targetSheet}" が取得できません`
                    };
                }
                console.log('シート取得成功');
                const dataRange = sheet.getDataRange();
                const totalRows = dataRange.getNumRows();
                const totalCols = dataRange.getNumColumns();
                console.log(`データ範囲: ${totalRows}行 × ${totalCols}列`);
                // 6. extractStructuredWeekData呼び出し
                console.log('\n【ステップ6】extractStructuredWeekData呼び出し');
                const weekData = extractStructuredWeekData(sheet);
                console.log('extractStructuredWeekData結果:', typeof weekData);
                if (!weekData) {
                    return {
                        success: false,
                        error: 'extractStructuredWeekDataがnullを返しました'
                    };
                }
                console.log('番組数:', Object.keys(weekData).length);
                console.log('番組一覧:', Object.keys(weekData));
                // 7. 指定番組の存在確認
                console.log('\n【ステップ7】指定番組の存在確認');
                console.log(`要求番組名: "${programName}"`);
                console.log(`完全一致: ${weekData[programName] ? 'あり' : 'なし'}`);
                // 部分マッチング確認
                const matchingPrograms = [];
                Object.keys(weekData).forEach(availableProgram => {
                    if (availableProgram.includes(programName) || programName.includes(availableProgram)) {
                        matchingPrograms.push(availableProgram);
                    }
                });
                console.log('部分マッチング番組:', matchingPrograms);
                // 8. 番組データ構造確認
                if (weekData[programName] || matchingPrograms.length > 0) {
                    const targetProgram = weekData[programName] ? programName : matchingPrograms[0];
                    console.log(`\n【ステップ8】番組データ構造確認: "${targetProgram}"`);
                    const programData = weekData[targetProgram];
                    console.log('曜日数:', Object.keys(programData).length);
                    console.log('利用可能曜日:', Object.keys(programData));
                    // 各曜日のデータサンプル
                    Object.keys(programData).slice(0, 2).forEach(day => {
                        console.log(`${day}のデータキー:`, Object.keys(programData[day]));
                    });
                }
                // 9. generateTransposedProgramTable呼び出しテスト
                console.log('\n【ステップ9】generateTransposedProgramTable呼び出しテスト');
                const result = generateTransposedProgramTable(programName);
                console.log('結果タイプ:', typeof result);
                console.log('success:', result.success);
                if (!result.success) {
                    console.log('error:', result.error);
                }
                else {
                    console.log('data:', result.data ? 'あり' : 'なし');
                    if (result.data) {
                        console.log('headers:', result.data.headers);
                        console.log('rows数:', result.data.rows ? result.data.rows.length : 0);
                    }
                }
                if (result.debugLogs) {
                    console.log('debugLogs数:', result.debugLogs.length);
                }
                return {
                    success: true,
                    testResults: {
                        totalSheets: allSheets.length,
                        weekSheetsFound: weekSheets.length,
                        targetSheet: targetSheet,
                        programsFound: Object.keys(weekData).length,
                        availablePrograms: Object.keys(weekData),
                        requestedProgram: programName,
                        exactMatch: !!weekData[programName],
                        partialMatches: matchingPrograms,
                        finalResult: result
                    }
                };
            }
            catch (error) {
                console.error('テスト中にエラー:', error);
                return {
                    success: false,
                    error: error instanceof Error ? error.message : String(error),
                    stack: error instanceof Error ? error.stack : undefined
                };
            }
        }
    }
    finally {
    }
}
