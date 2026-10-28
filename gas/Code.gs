// 蔵書スキャン — スプレッドシートに紐づける Apps Script。
// GET  ?isbn=XXXXXXXXXXXXX  → 書誌情報 + 既存登録 + 入力候補
// GET  (isbn なし)           → 入力候補のみ
// POST {isbn, location, title, author, publisher, genre, force}
//      → 重複があり force でなければ {status:'duplicate'}、それ以外は追記して {status:'ok'}

const SHEET_NAME = '蔵書';
const HEADERS = ['保管場所', '書籍名', '著者名', '出版社', 'ジャンル', 'ISBN', '登録日時'];
const COL = { location: 0, title: 1, author: 2, publisher: 3, genre: 4, isbn: 5, date: 6 };

function doGet(e) {
  const isbn = normalizeIsbn_((e.parameter || {}).isbn || '');
  const rows = readRows_();
  const res = { options: options_(rows) };
  if (isbn) {
    res.isbn = isbn;
    res.book = lookup_(isbn);
    res.duplicates = duplicatesOf_(rows, isbn);
  }
  return json_(res);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const isbn = normalizeIsbn_(body.isbn || '');
  if (!body.title) return json_({ status: 'error', message: '書籍名が空です' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const dups = isbn ? duplicatesOf_(readRows_(), isbn) : [];
    if (dups.length && !body.force) return json_({ status: 'duplicate', duplicates: dups });

    const row = [];
    row[COL.location] = cell_(body.location);
    row[COL.title] = cell_(body.title);
    row[COL.author] = cell_(body.author);
    row[COL.publisher] = cell_(body.publisher);
    row[COL.genre] = cell_(body.genre);
    row[COL.isbn] = isbn ? "'" + isbn : '';
    row[COL.date] = new Date();
    sheet_().appendRow(row);
    return json_({ status: 'ok' });
  } finally {
    lock.releaseLock();
  }
}

// ---- 書誌検索: openBD（和書に強い）→ Google Books の順 ----

function lookup_(isbn) {
  return fromOpenBd_(isbn) || fromGoogleBooks_(isbn);
}

function fromOpenBd_(isbn) {
  try {
    const res = UrlFetchApp.fetch('https://api.openbd.jp/v1/get?isbn=' + isbn, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return null;
    const hit = JSON.parse(res.getContentText())[0];
    if (!hit || !hit.summary || !hit.summary.title) return null;
    const s = hit.summary;
    return {
      title: [s.title, s.volume].filter(String).join(' '),
      author: cleanAuthor_(s.author),
      publisher: s.publisher || '',
      source: 'openBD',
    };
  } catch (err) {
    return null;
  }
}

function fromGoogleBooks_(isbn) {
  try {
    const key = PropertiesService.getScriptProperties().getProperty('GOOGLE_BOOKS_API_KEY');
    const url = 'https://www.googleapis.com/books/v1/volumes?q=isbn:' + isbn + (key ? '&key=' + key : '');
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return null;
    const item = (JSON.parse(res.getContentText()).items || [])[0];
    if (!item) return null;
    const v = item.volumeInfo;
    return {
      title: [v.title, v.subtitle].filter(Boolean).join(' '),
      author: (v.authors || []).join('、'),
      publisher: v.publisher || '',
      source: 'Google Books',
    };
  } catch (err) {
    return null;
  }
}

// openBD の「夏目,漱石,1867-1916」「村上春樹／著」形式を「夏目漱石」「村上春樹」に整える
function cleanAuthor_(s) {
  return (s || '')
    .replace(/,\s*\d{4}-(\d{4})?/g, '')
    .replace(/／\S+/g, '')
    .replace(/,/g, '')
    .trim();
}

// ---- シート ----

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function readRows_() {
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, HEADERS.length).getValues();
}

function duplicatesOf_(rows, isbn) {
  return rows
    .filter(r => String(r[COL.isbn]) === isbn)
    .map(r => ({
      location: r[COL.location],
      date: r[COL.date] instanceof Date
        ? Utilities.formatDate(r[COL.date], Session.getScriptTimeZone(), 'yyyy/MM/dd')
        : String(r[COL.date]),
    }));
}

function options_(rows) {
  const uniq = i => [...new Set(rows.map(r => String(r[i]).trim()).filter(Boolean))];
  return { locations: uniq(COL.location), genres: uniq(COL.genre) };
}

// ---- 小物 ----

function normalizeIsbn_(s) {
  const d = String(s).replace(/[^0-9Xx]/g, '').toUpperCase();
  if (d.length === 13) return d;
  if (d.length === 10) {
    const core = '978' + d.slice(0, 9);
    let sum = 0;
    for (let i = 0; i < 12; i++) sum += Number(core[i]) * (i % 2 ? 3 : 1);
    return core + ((10 - (sum % 10)) % 10);
  }
  return '';
}

// 先頭が = + - @ だと数式として解釈されるので文字列に固定する
function cell_(v) {
  const s = String(v == null ? '' : v).trim();
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
