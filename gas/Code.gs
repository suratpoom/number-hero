/**
 * เจ้าคุณผจญภัย: หลังบ้านเก็บความคืบหน้าของผู้เล่นลง Google Sheet (ผูกกับชีตนี้)
 *
 * ไม่มีการล็อกอิน: ผู้เล่นพิมพ์ชื่อหรือเลือกชื่อที่มีอยู่ แล้วเกมบันทึกตามชื่อนั้น
 * ใครรู้ชื่อก็เปิดเล่นต่อและบันทึกทับได้ จึงเหมาะกับกลุ่มที่ไว้ใจกัน (ครอบครัว เพื่อน)
 *
 * Deploy เป็น Web app: Execute as = Me, Who has access = Anyone
 * เกมส่ง POST {action:'list'|'load'|'save'|'setpin', name, pin, data, newPin}
 * รหัส 4 หลัก (ไม่บังคับ): ชื่อที่ตั้งรหัสไว้ ต้องส่งรหัสตรงกันถึงจะโหลดหรือบันทึกได้ · เก็บเป็นข้อความธรรมดาในคอลัมน์ pin (กันเด็กเล่นทับกัน ไม่ใช่ระบบความปลอดภัยจริงจัง)
 */
const TAB = 'players';
const HEAD = ['key', 'name', 'points', 'medals', 'stage', 'updated_at', 'data', 'pin'];
const MAX_ROWS = 10000;   // เพดานจำนวนผู้เล่น กันชีตโตไม่จำกัด
const MAX_DATA = 45000;   // ช่องหนึ่งของชีตเก็บได้ 50,000 ตัวอักษร
const MAX_NAME = 12;

function doGet() { return json_({ ok: true, app: 'chaokhun' }); }

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    if (req.action === 'list') out = list_();
    else if (req.action === 'save') out = save_(req.name, req.data, req.pin);
    else if (req.action === 'setpin') out = setPin_(req.name, req.pin, req.newPin);
    else out = load_(req.name, req.pin);
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return json_(out);
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/** ชื่อที่แสดง: ตัดช่องว่างหัวท้าย ยุบช่องว่างซ้ำ ยาวไม่เกิน MAX_NAME · key = ชื่อตัวพิมพ์เล็ก ใช้เทียบว่าเป็นคนเดียวกัน */
function cleanName_(n) {
  const s = String(n == null ? '' : n).replace(/\s+/g, ' ').trim().slice(0, MAX_NAME);
  if (!s) throw new Error('no_name');
  return s;
}
function key_(n) { return cleanName_(n).toLowerCase(); }
/** รหัสต้องเป็นเลข 4 หลักพอดี ไม่งั้นถือว่าไม่มีรหัส */
function cleanPin_(p) { const s = String(p == null ? '' : p).trim(); return /^\d{4}$/.test(s) ? s : ''; }
function checkPin_(stored, pin) { if (stored && cleanPin_(pin) !== stored) throw new Error('bad_pin'); }

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(TAB);
  if (!sh) {
    sh = ss.insertSheet(TAB);
    sh.getRange(1, 1, 1, HEAD.length).setNumberFormat('@').setValues([HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  if (sh.getRange(1, HEAD.length).getDisplayValue() !== 'pin') sh.getRange(1, HEAD.length).setNumberFormat('@').setValue('pin').setFontWeight('bold');
  return sh;
}

function findRow_(sh, key) {
  const n = sh.getLastRow();
  if (n < 2) return 0;
  const v = sh.getRange(2, 1, n - 1, 1).getDisplayValues();
  for (let i = 0; i < v.length; i++) if (v[i][0] === key) return i + 2;
  return 0;
}

/** รายชื่อผู้เล่น 60 คนที่เล่นล่าสุด (ไม่ส่งข้อมูลเกมเต็ม) */
function list_() {
  const sh = sheet_(), n = sh.getLastRow();
  if (n < 2) return { ok: true, players: [] };
  const rows = sh.getRange(2, 1, n - 1, HEAD.length).getDisplayValues()
    .filter(function (r) { return r[0]; })
    .map(function (r) { return { name: r[1], points: r[2], medals: r[3], stage: r[4], updated: r[5], hasPin: !!r[7] }; });
  rows.sort(function (a, b) { return a.updated < b.updated ? 1 : -1; });
  return { ok: true, players: rows.slice(0, 60) };
}

function load_(name, pin) {
  const sh = sheet_(), r = findRow_(sh, key_(name));
  if (!r) return { ok: true, found: false };
  const row = sh.getRange(r, 1, 1, HEAD.length).getDisplayValues()[0];
  checkPin_(row[7], pin);
  let data = null;
  try { data = JSON.parse(row[6]); } catch (e) {}
  return { ok: true, found: !!data, name: row[1], updated: row[5], data: data, hasPin: !!row[7] };
}

function save_(name, data, pin) {
  if (!data || typeof data !== 'object') throw new Error('no_data');
  const js = JSON.stringify(data);
  if (js.length > MAX_DATA) throw new Error('too_big');
  const nm = cleanName_(name), key = nm.toLowerCase();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    let r = findRow_(sh, key), keep = cleanPin_(pin);
    if (!r) {
      if (sh.getLastRow() - 1 >= MAX_ROWS) throw new Error('full');
      r = sh.getLastRow() + 1;          // ชื่อใหม่: รหัสที่ส่งมา (ถ้ามี) กลายเป็นรหัสของชื่อนี้
    } else {
      keep = sh.getRange(r, HEAD.length).getDisplayValue();
      checkPin_(keep, pin);
    }
    const now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
    const A = data.A || {};
    const row = [key, nm, String(data.total || 0), String(A.medals || 0), String((A.adv || 0) + 1), now, js, keep];
    const rg = sh.getRange(r, 1, 1, HEAD.length);
    rg.setNumberFormat('@');   // เก็บเป็นข้อความทุกคอลัมน์
    rg.setValues([row]);
    return { ok: true, updated: now };
  } finally {
    lock.releaseLock();
  }
}

/** ตั้ง เปลี่ยน หรือลบรหัส (newPin ว่าง = ลบ) ต้องส่งรหัสเดิมให้ถูกก่อน */
function setPin_(name, pin, newPin) {
  const np = String(newPin == null ? '' : newPin).trim();
  if (np && !/^\d{4}$/.test(np)) throw new Error('pin_format');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_(), r = findRow_(sh, key_(name));
    if (!r) throw new Error('not_found');
    const cell = sh.getRange(r, HEAD.length);
    checkPin_(cell.getDisplayValue(), pin);
    cell.setNumberFormat('@').setValue(np);
    return { ok: true, hasPin: !!np };
  } finally {
    lock.releaseLock();
  }
}

/** รันครั้งเดียวจากตัวแก้ไขสคริปต์ เพื่ออนุญาตสิทธิ์และสร้างแท็บ players */
function setup() { sheet_(); return 'ok'; }
