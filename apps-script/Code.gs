// 課後練習與音檔庫：載入器（只要貼一次）
// 每次有人送成績或讀設定時，會自動從 GitHub 讀取最新的主程式 main.gs。
var SRC = 'https://raw.githubusercontent.com/qwe0975169776-dotcom/listening-practice/main/apps-script/main.gs';
function load_() {
  var cache = CacheService.getScriptCache();
  var code = cache.get('main');
  if (!code) { code = UrlFetchApp.fetch(SRC).getContentText(); cache.put('main', code, 300); }
  (0, eval)(code);
}
function doPost(e) { load_(); return handlePost(e); }
function doGet(e) { load_(); return handleGet(e); }
// 第一次設定時，在上方選單選「authorize」再按「執行」，允許程式連到 GitHub。
function authorize() { UrlFetchApp.fetch(SRC); SpreadsheetApp.getActiveSpreadsheet(); Logger.log('授權完成'); }
