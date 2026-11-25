/**
 * tmun.ai 學員遊戲計數器＋維護功能（遊玩人次、曝光、點閱率、刪除／恢復）
 * 放在「表單回覆試算表」裡：擴充功能 → Apps Script → 貼上全部程式碼 → 儲存
 * ★ 先把下面的 ADMIN_KEY 改成只有你知道的密碼 ★
 * 部署：部署 → 新增部署作業 → 類型選「網頁應用程式」
 *       執行身分：我　／　誰可以存取：所有人 → 部署 → 複製網址
 * 修改程式碼後：部署 → 管理部署作業 → 編輯（鉛筆）→ 版本選「新版本」→ 部署（網址不變）
 * 統計資料存在自動新增的「遊戲統計」工作表；「已隱藏」欄為 TRUE 的遊戲不會出現在遊戲牆。
 */
const ADMIN_KEY = '請改成你的密碼';
const STATS_SHEET = '遊戲統計';

function doGet(e) {
  const p = (e && e.parameter) || {};
  const sh = statsSheet_();
  const isAdmin = ADMIN_KEY && ADMIN_KEY !== '請改成你的密碼' && p.key === ADMIN_KEY;

  if (p.a === 'check') return json_({ ok: isAdmin });

  if (p.a === 'view' || p.a === 'play' || p.a === 'hide' || p.a === 'unhide') {
    if ((p.a === 'hide' || p.a === 'unhide') && !isAdmin) return json_({ error: 'key' });
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const ids = String(p.ids || p.id || '').split(',')
        .map(function (s) { return s.replace(/[^\w-]/g, '').slice(0, 40); })
        .filter(String).slice(0, 300);
      const values = sh.getDataRange().getValues();
      const rowOf = {};
      for (var i = 1; i < values.length; i++) rowOf[values[i][0]] = i + 1;
      ids.forEach(function (id) {
        var row = rowOf[id];
        if (!row) { sh.appendRow([id, 0, 0, new Date(), false]); row = sh.getLastRow(); rowOf[id] = row; }
        if (p.a === 'view' || p.a === 'play') {
          var cell = sh.getRange(row, p.a === 'view' ? 2 : 3);
          cell.setValue((Number(cell.getValue()) || 0) + 1);
        } else {
          sh.getRange(row, 5).setValue(p.a === 'hide');
        }
        sh.getRange(row, 4).setValue(new Date());
      });
    } finally {
      lock.releaseLock();
    }
  }

  const out = {};
  sh.getDataRange().getValues().slice(1).forEach(function (r) {
    out[r[0]] = { view: Number(r[1]) || 0, play: Number(r[2]) || 0, hidden: r[4] === true || String(r[4]).toUpperCase() === 'TRUE' };
  });
  return json_(out);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function statsSheet_() {
  const ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(STATS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(STATS_SHEET, ss.getSheets().length);
    sh.appendRow(['遊戲ID', '曝光次數', '遊玩人次', '最後更新', '已隱藏']);
  } else if (!sh.getRange(1, 5).getValue()) {
    sh.getRange(1, 5).setValue('已隱藏');
  }
  return sh;
}
