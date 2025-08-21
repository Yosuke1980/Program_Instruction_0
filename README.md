# ラジオ番組ドキュメント自動生成システム

Google Apps Scriptを使用したラジオ番組スケジュール管理・ドキュメント自動生成システムです。

## 開発環境

このプロジェクトはTypeScriptで開発され、Google Apps Script用のJavaScriptにコンパイルされます。

### 必要なソフトウェア

- Node.js (v14以上)
- npm

### セットアップ

```bash
# 依存関係のインストール
npm install

# TypeScriptの型チェック
npm run lint

# Google Apps Script用JavaScriptの生成
npm run build:gas
```

### 開発フロー

1. `src/コード.ts`でTypeScriptコードを編集
2. `npm run lint`で型チェック
3. `npm run build:gas`でGoogle Apps Script用のJavaScriptを生成
4. `dist/コード.js`の内容をGoogle Apps Scriptエディタにコピー

### 利用可能なスクリプト

- `npm run build`: TypeScriptをコンパイル
- `npm run build:watch`: ファイル変更を監視してリアルタイムコンパイル
- `npm run build:gas`: Google Apps Script用JavaScriptを生成
- `npm run lint`: 型チェックのみ実行（コンパイルしない）

## ファイル構成

```
├── src/
│   ├── コード.ts          # メインのTypeScriptファイル
│   └── config.ts          # 設定ファイル（現在は使用せず、コード.ts内に統合）
├── dist/
│   └── コード.js          # コンパイル済みJavaScript
├── tsconfig.json          # TypeScript設定
├── package.json           # Node.js設定
└── README.md             # このファイル
```

## 機能

- ラジオ番組スケジュールの自動取得
- Googleドキュメントの自動生成
- メール送信機能
- WebApp UI
- 楽曲データベース連携

## 注意事項

- `dist/コード.js`がGoogle Apps Scriptで実行される最終的なファイルです
- TypeScriptの開発では`src/コード.ts`を編集してください
- モジュールシステムはGoogle Apps Scriptでサポートされていないため、すべて単一ファイルにまとめられます