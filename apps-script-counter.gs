/**
 * tmun.ai 學員遊戲計數器（遊玩人次、曝光、點閱率）
 * 放在「表單回覆試算表」裡：擴充功能 → Apps Script → 貼上全部程式碼 → 儲存
 * 部署：部署 → 新增部署作業 → 類型選「網頁應用程式」
 *       執行身分：我　／　誰可以存取：所有人 → 部署 → 複製網址
 * 統計資料會存在試算表裡自動新增的「遊戲統計」工作表。
 */
const STATS_SHEET = '遊戲統計';

function doGet(e) {
  const p = (e && e.parameter) || {};
  const sh = statsSheet_();
  if (p.a === 'view' || p.a === 'play') {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const ids = String(p.ids || p.id || '').split(',')
        .map(function (s) { return s.replace(/[^\w-]/g, '').slice(0, 40); })
        .filter(String).slice(0, 300);
      const col = p.a === 'view' ? 2 : 3;
      const values = sh.getDataRange().getValues();
      const rowOf = {};
      for (var i = 1; i < values.length; i++) rowOf[values[i][0]] = i + 1;
      ids.forEach(function (id) {
        var row = rowOf[id];
        if (!row) { sh.appendRow([id, 0, 0, new Date()]); row = sh.getLastRow(); rowOf[id] = row; }
        var cell = sh.getRange(row, col);
        cell.setValue((Number(cell.getValue()) || 0) + 1);
        sh.getRange(row, 4).setValue(new Date());
      });
    } finally {
      lock.releaseLock();
    }
  }
  const out = {};
  sh.getDataRange().getValues().slice(1).forEach(function (r) {
    out[r[0]] = { view: Number(r[1]) || 0, play: Number(r[2]) || 0 };
  });
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function statsSheet_() {
  const ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(STATS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(STATS_SHEET, ss.getSheets().length);
    sh.appendRow(['遊戲ID', '曝光次數', '遊玩人次', '最後更新']);
  }
  return sh;
}
