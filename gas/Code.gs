/**
 * เจ้าคุณผจญภัย: หลังบ้านเก็บความคืบหน้าของผู้เล่นลง Google Sheet
 *
 * Script Properties ที่ต้องตั้ง:
 *   CLIENT_ID = OAuth Client ID ของปุ่ม "เข้าสู่ระบบด้วย Google" (ต้องตรงกับค่าในเกม)
 *   SHEET_ID  = ID ของ Google Sheet ที่ใช้เก็บข้อมูล
 *
 * Deploy เป็น Web app: Execute as = Me, Who has access = Anyone
 * เกมส่ง POST {action:'load'|'save', idToken, data} มา
 * ทุกคำขอต้องมี idToken ของ Google และถูกตรวจกับ Google ก่อนเสมอ ไม่เชื่ออีเมลที่ฝั่งเกมส่งมาเอง
 */
const TAB = 'players';
const HEAD = ['sub', 'email', 'google_name', 'kid_name', 'points', 'medals', 'stage', 'updated_at', 'data'];
const MAX_ROWS = 10000;   // เพดานจำนวนบัญชี กันชีตโตไม่จำกัด
const MAX_DATA = 45000;   // ช่องหนึ่งของชีตเก็บได้ 50,000 ตัวอักษร

function doGet() { return json_({ ok: true, app: 'chaokhun' }); }

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const user = verify_(req.idToken);
    out = req.action === 'save' ? save_(user, req.data) : load_(user);
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return json_(out);
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k); }

/** ตรวจ ID token กับ Google: ต้องออกให้แอปนี้ (aud) ออกโดย Google (iss) และยังไม่หมดอายุ */
function verify_(token) {
  if (!token) throw new Error('no_token');
  const r = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(token), { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('bad_token');
  const p = JSON.parse(r.getContentText());
  if (p.aud !== prop_('CLIENT_ID')) throw new Error('bad_audience');
  if (['accounts.google.com', 'https://accounts.google.com'].indexOf(p.iss) < 0) throw new Error('bad_issuer');
  if (Number(p.exp) * 1000 < Date.now()) throw new Error('expired');
  if (!p.sub) throw new Error('bad_token');
  return { sub: String(p.sub), email: p.email || '', name: p.name || '' };
}

function sheet_() {
  const ss = SpreadsheetApp.openById(prop_('SHEET_ID'));
  let sh = ss.getSheetByName(TAB);
  if (!sh) {
    sh = ss.insertSheet(TAB);
    sh.getRange(1, 1, 1, HEAD.length).setNumberFormat('@').setValues([HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function findRow_(sh, sub) {
  const n = sh.getLastRow();
  if (n < 2) return 0;
  const v = sh.getRange(2, 1, n - 1, 1).getDisplayValues();
  for (let i = 0; i < v.length; i++) if (v[i][0] === sub) return i + 2;
  return 0;
}

function load_(user) {
  const sh = sheet_(), r = findRow_(sh, user.sub);
  if (!r) return { ok: true, found: false, email: user.email };
  const row = sh.getRange(r, 1, 1, HEAD.length).getDisplayValues()[0];
  let data = null;
  try { data = JSON.parse(row[8]); } catch (e) {}
  return { ok: true, found: !!data, email: user.email, updated: row[7], data: data };
}

function save_(user, data) {
  if (!data || typeof data !== 'object') throw new Error('no_data');
  const js = JSON.stringify(data);
  if (js.length > MAX_DATA) throw new Error('too_big');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    let r = findRow_(sh, user.sub);
    if (!r) {
      if (sh.getLastRow() - 1 >= MAX_ROWS) throw new Error('full');
      r = sh.getLastRow() + 1;
    }
    const now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
    const P = data.P || {}, A = data.A || {};
    const row = [user.sub, user.email, user.name, String(P.name || ''), String(data.total || 0), String(A.medals || 0), String(A.adv || 0), now, js];
    const rg = sh.getRange(r, 1, 1, HEAD.length);
    rg.setNumberFormat('@');   // เก็บเป็นข้อความทุกคอลัมน์
    rg.setValues([row]);
    return { ok: true, updated: now };
  } finally {
    lock.releaseLock();
  }
}
