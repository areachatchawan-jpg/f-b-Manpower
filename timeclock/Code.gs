/**
 * ระบบหลังบ้าน: ลงเวลานอกสถานที่ + OT  (Google Apps Script)
 * วางโค้ดนี้ใน Extensions > Apps Script ของ Google Sheet แล้ว Deploy เป็น Web app
 * ชีตที่สร้างอัตโนมัติ: Employees, Attendance, OT
 */
const TZ = 'Asia/Bangkok';
const PHOTO_FOLDER = 'หลักฐานลงเวลา';
const OT_RATE = 60; // ค่า OT บาทต่อชั่วโมง อัตราเดียวทุกคน
const DEFAULT_WAGE = { type: 'monthly', amount: 0, start: '08:00', end: '17:00' };

// พิกัดเริ่มต้นของจุดสแกน (ชีต Hotels จะถูกสร้างจากค่านี้ แล้ว HR แก้ในชีตได้เลย)
const H_HOTEL = ['id', 'ชื่อโรงแรม', 'lat', 'lng', 'รัศมี (ม.)', 'GPS แม่นยำอย่างน้อย (±ม.)'];
const DEFAULT_HOTELS = [
  ['baiyoke', 'Baiyoke Sky Hotel Bangkok', 13.754400, 100.540583, 50, 30],
  ['queensland', 'Queensland Hotel Bangkok', 13.756750, 100.538559, 50, 30]
];
const H_EMP = ['รหัส', 'ชื่อ', 'ประเภทค่าจ้าง (monthly/daily)', 'ค่าจ้าง', 'เข้างาน', 'เลิกงาน', 'credentialId', 'ลงทะเบียนเมื่อ', 'อุปกรณ์', 'สังกัดโรงแรม', 'สิทธิ์ (admin/staff)'];
// ผู้ที่ลงทะเบียนด้วยชื่อนี้เป็นคนแรกจะได้สิทธิ์แอดมินอัตโนมัติ (เปลี่ยนสิทธิ์ภายหลังได้ที่คอลัมน์ "สิทธิ์" ในชีต Employees)
const ADMIN_NAME = 'ชัชวาลย์';
const H_ATT = ['เวลาเซิร์ฟเวอร์', 'รหัส', 'ชื่อ', 'เข้า/ออก', 'lat', 'lng', 'ความแม่นยำ (ม.)', 'แผนที่', 'รูปหลักฐาน', 'สถานที่/หมายเหตุ', 'เวลาเครื่อง', 'ยืนยันลายนิ้วมือ', 'อุปกรณ์', 'โรงแรม', 'ระยะห่างจากโรงแรม (ม.)', 'พื้นที่'];
const H_OT = ['ยื่นเมื่อ', 'รหัส', 'ชื่อ', 'วันที่ทำ', 'เริ่ม', 'สิ้นสุด', 'พัก (นาที)', 'ชั่วโมง', 'ประเภท', 'อัตรา (เท่า)', 'บาท/ชม.', 'จำนวนเงิน (บาท)', 'งาน/เหตุผล', 'รูปหลักฐาน', 'แผนที่', 'ตรวจกับเวลาสแกนออก', 'สถานะ', 'โรงแรม'];

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      const fn = { register, clock, ot, list, profile, report, otStatus }[d.action];
      if (!fn) throw new Error('คำสั่งไม่ถูกต้อง');
      return out(Object.assign({ ok: true }, fn(d)));
    } finally { lock.releaseLock(); }
  } catch (err) {
    return out({ ok: false, error: String(err.message || err) });
  }
}
function doGet() { return out({ ok: true, msg: 'ระบบลงเวลาทำงานอยู่' }); }

/** กด Run ฟังก์ชันนี้ 1 ครั้งในหน้า Apps Script เพื่ออนุญาตสิทธิ์และสร้างชีตทั้งหมด */
function setup() {
  sheet('Employees', H_EMP); sheet('Attendance', H_ATT); sheet('OT', H_OT); hotels();
  const it = DriveApp.getFoldersByName(PHOTO_FOLDER); if (!it.hasNext()) DriveApp.createFolder(PHOTO_FOLDER);
  Logger.log('พร้อมใช้งาน');
}

// ---------- แอดมิน: ต้องเป็นเครื่องที่ลงทะเบียน + สิทธิ์ admin ในชีต Employees ----------
function checkAdmin(d) { const e = auth(d); if (roleOf(e) !== 'admin') throw new Error('บัญชีนี้ไม่มีสิทธิ์ดูรายงาน'); return e; }

function report(d) {
  checkAdmin(d);
  const m = String(d.month || '').slice(0, 7);
  const att = sheet('Attendance', H_ATT).getDataRange().getValues().slice(1)
    .map(r => ({ ts: txt(r[0]), id: String(r[1]), name: String(r[2]), type: r[3] === 'เข้า' ? 'in' : 'out', lat: r[4], lng: r[5], acc: r[6], photo: r[8], place: r[9], hotel: r[13], dist: r[14] }))
    .filter(a => a.ts.slice(0, 7) === m);
  const ot = sheet('OT', H_OT).getDataRange().getValues().slice(1)
    .map((r, i) => ({ row: i + 2, submitted: txt(r[0]), id: String(r[1]), name: String(r[2]), date: txt(r[3]).slice(0, 10), start: txt(r[4]).slice(-5), end: txt(r[5]).slice(-5), brk: r[6], hours: Number(r[7]) || 0, amount: Number(r[11]) || 0, reason: r[12], photo: r[13], check: r[15], status: r[16], hotel: r[17] }))
    .filter(o => o.date.slice(0, 7) === m);
  const emp = sheet('Employees', H_EMP).getDataRange().getValues().slice(1)
    .map(r => ({ id: String(r[0]), name: String(r[1]), hotel: r[9] || '', registered: !!r[6], role: r[10] || 'staff' }));
  return { att, ot, emp, rate: OT_RATE };
}

function otStatus(d) {
  checkAdmin(d);
  const s = sheet('OT', H_OT), row = Number(d.row);
  if (!(row >= 2 && row <= s.getLastRow())) throw new Error('ไม่พบรายการ OT');
  if (['อนุมัติ', 'ไม่อนุมัติ', 'รออนุมัติ'].indexOf(d.status) < 0) throw new Error('สถานะไม่ถูกต้อง');
  if (String(s.getRange(row, 2).getValue()) !== String(d.id)) throw new Error('ข้อมูลเปลี่ยนไปแล้ว กรุณารีเฟรช');
  s.getRange(row, 17).setValue(d.status);
  return {};
}
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// คอลัมน์วันที่/เวลา เก็บเป็นข้อความ เพื่อไม่ให้ Sheets แปลงเขตเวลาเอง
const TEXT_COLS = { Employees: ['A', 'E', 'F', 'H'], Attendance: ['A', 'B', 'K'], OT: ['A', 'B', 'D', 'E', 'F'] };
function sheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name); s.appendRow(headers); s.setFrozenRows(1);
    s.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    (TEXT_COLS[name] || []).forEach(c => s.getRange(c + ':' + c).setNumberFormat('@'));
  }
  return s;
}
const now = () => Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
const txt = v => v instanceof Date ? Utilities.formatDate(v, TZ, v.getFullYear() < 1900 ? 'HH:mm' : 'yyyy-MM-dd HH:mm:ss') : String(v == null ? '' : v);

function hotels() {
  const s = sheet('Hotels', H_HOTEL);
  if (s.getLastRow() < 2) s.getRange(2, 1, DEFAULT_HOTELS.length, H_HOTEL.length).setValues(DEFAULT_HOTELS);
  return s.getDataRange().getValues().slice(1).filter(r => r[0])
    .map(r => ({ id: String(r[0]), name: String(r[1]), lat: Number(r[2]), lng: Number(r[3]), radius: Number(r[4]) || 50, maxAcc: Number(r[5]) || 30 }));
}
function distM(a, b) {
  const r = Math.PI / 180, x = (b.lng - a.lng) * r * Math.cos((a.lat + b.lat) / 2 * r), y = (b.lat - a.lat) * r;
  return Math.round(Math.sqrt(x * x + y * y) * 6371000);
}
function findEmp(empId) {
  const s = sheet('Employees', H_EMP), v = s.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]).trim() === String(empId).trim()) return { s, row: i + 1, v: v[i] };
  return null;
}
function roleOf(e) { return String(e.v[10] || '').trim().toLowerCase() === 'admin' ? 'admin' : 'staff'; }
function roleFor(name) {
  const v = sheet('Employees', H_EMP).getDataRange().getValues().slice(1);
  const hasAdmin = v.some(r => String(r[10] || '').trim().toLowerCase() === 'admin');
  return !hasAdmin && String(name).indexOf(ADMIN_NAME) >= 0 ? 'admin' : 'staff';
}
function wageOf(e) {
  return { type: e.v[2] === 'daily' ? 'daily' : 'monthly', amount: Number(e.v[3]) || 0, start: txt(e.v[4]) || '08:00', end: txt(e.v[5]) || '17:00' };
}
/** ตรวจว่าเครื่องที่ส่งมาคือเครื่องที่ลงทะเบียนลายนิ้วมือไว้ */
function auth(d) {
  const e = findEmp(d.empId);
  if (!e) throw new Error('ไม่พบรหัสพนักงาน กรุณาลงทะเบียน');
  if (!e.v[6] || e.v[6] !== d.credId) throw new Error('อุปกรณ์นี้ไม่ใช่เครื่องที่ลงทะเบียนไว้ ติดต่อ HR');
  return e;
}

function register(d) {
  if (!d.empId || !d.name || !d.credId) throw new Error('ข้อมูลไม่ครบ');
  const e = findEmp(d.empId);
  if (e && e.v[6]) throw new Error('รหัสนี้ลงทะเบียนเครื่องอื่นไว้แล้ว ให้ HR ลบช่อง credentialId ก่อน');
  if (e) {
    e.s.getRange(e.row, 7, 1, 4).setValues([[d.credId, now(), d.device || '', d.hotel || e.v[9] || '']]);
    if (!e.v[10]) e.s.getRange(e.row, 11).setValue(roleFor(d.name));
    const e2 = findEmp(d.empId);
    return { profile: wageOf(e2), role: roleOf(e2), hotels: hotels() };
  }
  const role = roleFor(d.name);
  sheet('Employees', H_EMP).appendRow([d.empId, d.name, DEFAULT_WAGE.type, DEFAULT_WAGE.amount, DEFAULT_WAGE.start, DEFAULT_WAGE.end, d.credId, now(), d.device || '', d.hotel || '', role]);
  return { profile: DEFAULT_WAGE, role, hotels: hotels() };
}

function profile(d) { const e = auth(d); return { profile: wageOf(e), role: roleOf(e), hotels: hotels() }; }

function savePhoto(dataUrl, name) {
  if (!dataUrl || dataUrl.indexOf('base64,') < 0) return '';
  const it = DriveApp.getFoldersByName(PHOTO_FOLDER);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER);
  const blob = Utilities.newBlob(Utilities.base64Decode(dataUrl.split('base64,')[1]), 'image/jpeg', name + '.jpg');
  return folder.createFile(blob).getUrl();
}
const mapUrl = g => g && g.lat ? 'https://maps.google.com/?q=' + g.lat + ',' + g.lng : '';

function clock(d) {
  const e = auth(d), ts = now(), g = d.gps || {};
  if (!g.lat) throw new Error('ไม่มีพิกัด GPS');
  if (!d.photo) throw new Error('ไม่มีรูปหลักฐาน');
  // ตรวจพื้นที่ซ้ำที่เซิร์ฟเวอร์ (ไม่เชื่อค่าจากหน้าเว็บ)
  const h = hotels().filter(x => x.id === d.hotelId)[0];
  if (!h) throw new Error('ไม่พบโรงแรมที่เลือก');
  const dist = distM({ lat: Number(g.lat), lng: Number(g.lng) }, h);
  if (Number(g.acc) > h.maxAcc) throw new Error('สัญญาณ GPS ไม่แม่นพอ (±' + g.acc + ' ม.) ลองใหม่ใกล้หน้าต่างหรือทางเข้า');
  if (dist > h.radius) throw new Error('อยู่นอกพื้นที่ ' + h.name + ' (ห่าง ' + dist + ' ม.) ไม่สามารถสแกนได้');
  const url = savePhoto(d.photo, d.empId + '_' + d.type + '_' + ts.replace(/[: ]/g, '-'));
  sheet('Attendance', H_ATT).appendRow([ts, d.empId, e.v[1], d.type === 'in' ? 'เข้า' : 'ออก', g.lat, g.lng, g.acc, mapUrl(g), url, d.place || '', d.clientTime || '', 'ผ่าน', d.device || '', h.name, dist, 'ในพื้นที่']);
  return { ts, dist };
}

function calcOt(start, end, brk, type, w) {
  const toM = t => { const p = String(t).split(':'); return Number(p[0]) * 60 + Number(p[1]); };
  let mins = toM(end) - toM(start); if (mins <= 0) mins += 1440;
  mins = Math.max(0, mins - (Number(brk) || 0));
  const hours = Math.round(mins / 60 * 100) / 100;
  return { hours, rate: 1, hourly: OT_RATE, amount: Math.round(hours * OT_RATE * 100) / 100 };
}

function ot(d) {
  const e = auth(d);
  if (!d.date || !d.start || !d.end) throw new Error('กรอกวันที่และเวลาให้ครบ');
  const c = calcOt(d.start, d.end, d.brk, d.type, wageOf(e));
  // เทียบกับเวลาสแกนออกจริงของวันนั้น
  const att = sheet('Attendance', H_ATT).getDataRange().getValues().slice(1);
  const outs = att.filter(r => String(r[1]) === String(d.empId) && r[3] === 'ออก' && txt(r[0]).slice(0, 10) === d.date).map(r => txt(r[0]).slice(11, 16));
  const overnight = d.end <= d.start;
  const check = overnight ? 'OT ข้ามคืน ตรวจด้วยตนเอง' : outs.some(t => t >= d.end) ? 'ตรงกับเวลาสแกนออก' : 'ไม่พบสแกนออกที่ครอบคลุม';
  const label = 'OT';
  const url = savePhoto(d.photo, d.empId + '_OT_' + d.date);
  sheet('OT', H_OT).appendRow([now(), d.empId, e.v[1], d.date, d.start, d.end, d.brk || 0, c.hours, label, c.rate, c.hourly, c.amount, d.reason || '', url, mapUrl(d.gps), check, 'รออนุมัติ', d.hotel || '']);
  return Object.assign({ check }, c);
}

function list(d) {
  auth(d);
  const id = String(d.empId);
  const att = sheet('Attendance', H_ATT).getDataRange().getValues().slice(1).filter(r => String(r[1]) === id).slice(-120)
    .map(r => ({ ts: txt(r[0]), type: r[3] === 'เข้า' ? 'in' : 'out', lat: r[4], lng: r[5], acc: r[6], photo: r[8], place: r[9], hotel: r[13], dist: r[14], inArea: r[15] }));
  const typ = { 'OT': 'ot' };
  const ot = sheet('OT', H_OT).getDataRange().getValues().slice(1).filter(r => String(r[1]) === id).slice(-60)
    .map(r => ({ date: txt(r[3]).slice(0, 10), start: txt(r[4]).slice(-5), end: txt(r[5]).slice(-5), hours: r[7], type: typ[r[8]], rate: r[9], amount: r[11], reason: r[12], photo: r[13], check: r[15], status: r[16], hotel: r[17] }));
  return { att, ot };
}
