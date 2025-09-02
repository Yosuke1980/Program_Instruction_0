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
  PORTSIDE_CALENDAR_ID: 'c_a25ffe304a25338178cd461a7b159103bd46ea62c64e773fdf4c61b4b36ab2fc@group.calendar.google.com',
  
  // 番組メタデータ管理用スプレッドシート
  METADATA_SPREADSHEET_ID: '1r6GLEvsZiqb3vkXZmrZ7XCpvW4RcecC7Bfp9YzhnQUE', // 同じスプレッドシートを使用
  METADATA_SHEET_NAME: '番組メタデータ'
};