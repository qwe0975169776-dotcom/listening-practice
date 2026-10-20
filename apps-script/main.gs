// 課後練習與音檔庫：Apps Script 主程式（由「載入器」自動從 GitHub 讀取，不需要手動貼）
var PW_SHA256 = 'ef797c8118f02dfb649607dd5d3f8c7623048c9c063d532cc95c5ed7a898a64f'; // 老師密碼 12345678 的雜湊
var RESULT_HEAD = ['時間', '姓名', '週次', '題組', '模式', '答對', '題數', '正確率', '用時（秒）', '答錯的題號', '作答編號'];
var DETAIL_HEAD = ['時間', '姓名', '週次', '題組', '模式', '題號', '學生的答案', '對錯', '正確答案（隨機題）', '作答編號', '題目序號'];

function sheet_(name, head) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  var first = sh.getLastRow() ? sh.getRange(1, 1, 1, head.length).getValues()[0] : [];
  if (first.join('|') !== head.join('|')) { sh.getRange(1, 1, 1, head.length).setValues([head]); sh.setFrozenRows(1); }
  return sh;
}
function json_(x) {
  return ContentService.createTextOutput(JSON.stringify(x)).setMimeType(ContentService.MimeType.JSON);
}
function okPw_(pw) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(pw || ''), Utilities.Charset.UTF_8);
  var hex = raw.map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('');
  return hex === PW_SHA256;
}
function handlePost(e) {
  if (!e || !e.postData) { Logger.log('這個函式只在網頁送資料時使用，不用在編輯器按「執行」。'); return ContentService.createTextOutput('no data'); }
  var d = JSON.parse(e.postData.contents);
  if (d.type === 'config') {
    if (!okPw_(d.pw)) return json_({ ok: false, error: 'password' });
    PropertiesService.getScriptProperties().setProperty('config', JSON.stringify({ hidden: d.hidden || [], updated: new Date().toISOString() }));
    return json_({ ok: true });
  }
  var now = new Date();
  var rate = d.total ? Math.round(d.score / d.total * 100) + '%' : '';
  sheet_('成績', RESULT_HEAD).appendRow([now, d.name, '第 ' + d.week + ' 週', d.set, d.mode, d.score, d.total, rate, d.seconds, d.wrong, d.id || '']);
  var rows = (d.details || []).map(function (x) {
    return [now, d.name, '第 ' + d.week + ' 週', d.set, d.mode, x.n, x.val, x.ok ? '○' : '✕', x.ans || '', d.id || '', x.idx == null ? '' : x.idx];
  });
  if (rows.length) {
    var sh = sheet_('作答明細', DETAIL_HEAD);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, DETAIL_HEAD.length).setValues(rows);
  }
  return json_({ ok: true });
}
function handleGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action === 'config') {
    var c = PropertiesService.getScriptProperties().getProperty('config');
    return json_(c ? JSON.parse(c) : { hidden: [] });
  }
  if (p.action === 'results') {
    if (!okPw_(p.pw)) return json_({ ok: false, error: 'password' });
    var out = { ok: true, results: [], details: [] };
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var r = ss.getSheetByName('成績');
    if (r && r.getLastRow() > 1) {
      var n = Math.min(300, r.getLastRow() - 1);
      out.results = r.getRange(r.getLastRow() - n + 1, 1, n, RESULT_HEAD.length).getValues();
    }
    var dsh = ss.getSheetByName('作答明細');
    if (dsh && dsh.getLastRow() > 1) {
      var m = Math.min(3000, dsh.getLastRow() - 1);
      out.details = dsh.getRange(dsh.getLastRow() - m + 1, 1, m, DETAIL_HEAD.length).getValues();
    }
    return json_(out);
  }
  return ContentService.createTextOutput('課後練習成績接收器正常運作中');
}
