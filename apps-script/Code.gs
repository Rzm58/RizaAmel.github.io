/**
 * Penerima RSVP & Ucapan untuk undangan Riza & Amel.
 * Tempel seluruh isi file ini ke Apps Script yang dibuat dari spreadsheet RSVP.
 */
var SPREADSHEET_ID = '1TM-xrPJwZJd_wqzxWfCNH5xolwcWTq5xPSnWZvnw_iw';
var SHEET_NAME = 'RSVP';
var HEADERS = ['Waktu', 'Nama', 'Kehadiran', 'Jumlah Tamu', 'Ucapan'];
var MAX_LIST = 50;

function getSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

// Dipanggil form saat tamu menekan "KIRIM UCAPAN".
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var p = (e && e.parameter) || {};
    var name = String(p.name || '').trim().slice(0, 100);
    var msg = String(p.msg || '').trim().slice(0, 1000);
    if (!name || !msg) return json_({ ok: false, error: 'nama dan ucapan wajib diisi' });

    var attend = ['Hadir', 'Tidak Hadir', 'Masih Ragu'].indexOf(p.attend) >= 0 ? p.attend : 'Masih Ragu';
    var guests = attend === 'Hadir' ? Math.min(Math.max(parseInt(p.guests, 10) || 1, 1), 10) : 0;

    // Awali dengan apostrof agar teks yang diawali = + - @ tidak dibaca sebagai rumus.
    var safe = function (s) { return /^[=+\-@]/.test(s) ? "'" + s : s; };
    getSheet_().appendRow([new Date(), safe(name), attend, guests, safe(msg)]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Dipanggil halaman undangan untuk menampilkan ucapan terbaru dari semua tamu.
function doGet() {
  var sh = getSheet_();
  var last = sh.getLastRow();
  if (last < 2) return json_({ ok: true, wishes: [] });
  var first = Math.max(2, last - MAX_LIST + 1);
  var rows = sh.getRange(first, 1, last - first + 1, HEADERS.length).getValues();
  var wishes = rows.reverse().map(function (r) {
    return {
      time: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      name: r[1],
      attend: r[2],
      guests: String(r[3]),
      msg: r[4]
    };
  });
  return json_({ ok: true, wishes: wishes });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
