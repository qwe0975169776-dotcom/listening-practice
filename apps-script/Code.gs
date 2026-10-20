// 課後練習與音檔庫：把學生的練習成績寫進這個試算表的「成績」工作表
function doPost(e) {
  var d = JSON.parse(e.postData.contents);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('成績') || ss.insertSheet('成績');
  if (sh.getLastRow() === 0) {
    sh.appendRow(['時間', '姓名', '週次', '題組', '模式', '答對', '題數', '正確率', '用時（秒）', '答錯的題號']);
    sh.setFrozenRows(1);
  }
  var rate = d.total ? Math.round(d.score / d.total * 100) + '%' : '';
  sh.appendRow([new Date(), d.name, '第 ' + d.week + ' 週', d.set, d.mode, d.score, d.total, rate, d.seconds, d.wrong]);
  return ContentService.createTextOutput('ok');
}
function doGet() {
  return ContentService.createTextOutput('課後練習成績接收器正常運作中');
}
