/**
 * ระบบหลังบ้าน: ลงเวลานอกสถานที่ + OT  (Google Apps Script)
 * วางโค้ดนี้ใน Extensions > Apps Script ของ Google Sheet แล้ว Deploy เป็น Web app
 * ชีตที่สร้างอัตโนมัติ: Employees, Attendance, OT
 */
const TZ = 'Asia/Bangkok';
const PHOTO_FOLDER = 'หลักฐานลงเวลา';
const OT_RATE = 60; // ค่า OT บาทต่อชั่วโมง อัตราเดียวทุกคน
const DEFAULT_WAGE = { type: 'monthly', amount: 0, start: '08:00', end: '17:00' };

// จุดสแกนเริ่มต้น (ชีต Hotels จะถูกสร้างจากค่านี้) · 1 แถว = 1 จุดสแกน ใส่ id ซ้ำได้ = โรงแรมนั้นมีหลายจุด
// อยู่ใกล้จุดไหนในรัศมีก็สแกนผ่าน · เพิ่มจุดใหม่ = เพิ่มแถว (ใช้รัศมี/ความแม่นยำจากแถวแรกของโรงแรมนั้น)
const H_HOTEL = ['id', 'ชื่อโรงแรม', 'lat', 'lng', 'รัศมี (ม.)', 'GPS แม่นยำอย่างน้อย (±ม.)', 'จุด / หมายเหตุ'];
const DEFAULT_HOTELS = [
  ['baiyoke', 'Baiyoke Sky Hotel Bangkok', 13.7544, 100.540583, 50, 65, 'กลางทาวเวอร์'],
  ['baiyoke', 'Baiyoke Sky Hotel Bangkok', 13.75469, 100.54004, 50, 65, 'วัดจริง 21 พ.ย.'],
  ['queensland', 'Queensland Hotel Bangkok', 13.75675, 100.538559, 50, 65, 'หน้าตึก'],
  ['queensland', 'Queensland Hotel Bangkok', 13.756764, 100.53798, 50, 65, 'วัดจริง 21 พ.ย.'],
  ['queensland', 'Queensland Hotel Bangkok', 13.757218, 100.537377, 50, 65, 'โซนจัดเลี้ยง วัดจริง 21 พ.ย.']
];
const H_EMP = ['รหัส', 'ชื่อ', 'ประเภทค่าจ้าง (monthly/daily)', 'ค่าจ้าง', 'เข้างาน', 'เลิกงาน', 'credentialId', 'ลงทะเบียนเมื่อ', 'อุปกรณ์', 'สังกัดโรงแรม'];
// แอดมิน: เก็บเฉพาะ hash ของรหัสผ่านเริ่มต้น เปลี่ยนรหัสได้ที่ชีต Admin (ช่อง B2)
const ADMIN_NAME = 'ชัชวาลย์ ต่อสุวรรณ';
const ADMIN_SALT = 'tc-baiyoke-queensland:';
const ADMIN_HASH = 'e8162daedddbbab76011664c0e7676be4de7135ff75784a56374eab865275732';
const H_ATT = ['เวลาเซิร์ฟเวอร์', 'รหัส', 'ชื่อ', 'เข้า/ออก', 'lat', 'lng', 'ความแม่นยำ (ม.)', 'แผนที่', 'รูปหลักฐาน', 'สถานที่/หมายเหตุ', 'เวลาเครื่อง', 'ยืนยันลายนิ้วมือ', 'อุปกรณ์', 'โรงแรม', 'ระยะห่างจากโรงแรม (ม.)', 'พื้นที่'];
const H_OT = ['คำนวณเมื่อ', 'รหัส', 'ชื่อ', 'วันที่ทำ', 'เลิกงานปกติ (เริ่ม OT)', 'สแกนออก', 'พัก (นาที)', 'ชั่วโมง OT', 'ประเภท', 'อัตรา (เท่า)', 'บาท/ชม.', 'จำนวนเงิน (บาท)', 'งาน/เหตุผล (พนักงานกรอก)', 'รูปหลักฐาน', 'แผนที่', 'ที่มา', 'สถานะ', 'โรงแรม', 'สแกนเข้า (อ้างอิง)'];
// ไม่มีกะตายตัว: ปัดเวลาสแกนเข้าเป็นชั่วโมงตรงที่ใกล้ที่สุด แล้วทำงานปกติ WORK_MIN นาที (รวมพัก) ส่วนที่เกินคือ OT
const WORK_MIN = 9 * 60; // OT นับเฉพาะชั่วโมงเต็ม (เศษนาทีไม่นับ)

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const fn = { register, clock, list, profile, report, otStatus, adminLogin, otNote }[d.action];
    if (!fn) throw new Error('คำสั่งไม่ถูกต้อง');
    // ล็อกเฉพาะคำสั่งที่เขียนข้อมูล คำสั่งอ่านทำงานพร้อมกันได้ ไม่ต้องรอคิว
    if (['list', 'profile', 'report', 'adminLogin'].indexOf(d.action) >= 0) return out(Object.assign({ ok: true }, fn(d)));
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try { return out(Object.assign({ ok: true }, fn(d))); } finally { lock.releaseLock(); }
  } catch (err) {
    return out({ ok: false, error: String(err.message || err) });
  }
}
function doGet() { return out({ ok: true, msg: 'ระบบลงเวลาทำงานอยู่' }); }

/** กด Run ฟังก์ชันนี้ 1 ครั้งในหน้า Apps Script เพื่ออนุญาตสิทธิ์และสร้างชีตทั้งหมด */
function setup() {
  sheet('Employees', H_EMP); sheet('Attendance', H_ATT); sheet('OT', H_OT); hotels(); adminSheet();
  const it = DriveApp.getFoldersByName(PHOTO_FOLDER); if (!it.hasNext()) DriveApp.createFolder(PHOTO_FOLDER);
  Logger.log('พร้อมใช้งาน');
}

// ---------- แอดมิน: ใช้รหัสผ่าน (ผิด 5 ครั้ง ล็อก 15 นาที) ----------
function adminSheet() {
  const s = sheet('Admin', ['ตั้งค่า', 'ค่า']);
  if (s.getLastRow() < 2) { s.appendRow(['ชื่อแอดมิน', ADMIN_NAME]); s.appendRow(['รหัสผ่านใหม่ (เว้นว่าง = ใช้รหัสเริ่มต้น)', '']); }
  if (s.getLastRow() < 5) s.getRange(5, 1, 1, 3).setValues([['แอดมินเพิ่มเติม (ชื่อ)', 'รหัสผ่าน', 'ตำแหน่ง']]).setFontWeight('bold');
  return s;
}
function hex(bytes) { return bytes.map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join(''); }
function checkAdmin(d) {
  const cache = CacheService.getScriptCache(), fails = Number(cache.get('admin_fail') || 0);
  if (fails >= 5) throw new Error('ใส่รหัสผิดหลายครั้ง ระบบล็อก 15 นาที');
  const s = adminSheet(), pass = String(d.adminPass || '');
  const custom = String(s.getRange(3, 2).getDisplayValue()).trim();
  let name = '';
  if (pass && (custom ? pass === custom : hex(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, ADMIN_SALT + pass, Utilities.Charset.UTF_8)) === ADMIN_HASH)) name = ADMIN_NAME;
  // แอดมินเพิ่มเติม: แถว 6 ลงไป คอลัมน์ A = ชื่อ, B = รหัสผ่าน (ลบแถว = ยกเลิกสิทธิ์)
  if (!name && pass && s.getLastRow() >= 6) {
    const rows = s.getRange(6, 1, s.getLastRow() - 5, 2).getDisplayValues();
    for (const r of rows) if (r[0].trim() && r[1].trim() && r[1].trim() === pass) { name = r[0].trim(); break; }
  }
  if (!name) { cache.put('admin_fail', String(fails + 1), 900); throw new Error('รหัสผ่านแอดมินไม่ถูกต้อง'); }
  cache.remove('admin_fail');
  return name;
}
function adminLogin(d) { return { name: checkAdmin(d) }; }

function report(d) {
  checkAdmin(d);
  const m = String(d.month || '').slice(0, 7);
  const att = sheet('Attendance', H_ATT).getDataRange().getValues().slice(1)
    .map(r => ({ ts: txt(r[0]), id: String(r[1]), name: String(r[2]), type: r[3] === 'เข้า' ? 'in' : 'out', lat: r[4], lng: r[5], acc: r[6], photo: r[8], place: r[9], hotel: r[13], dist: r[14] }))
    .filter(a => a.ts.slice(0, 7) === m);
  const ot = sheet('OT', H_OT).getDataRange().getValues().slice(1)
    .map((r, i) => ({ row: i + 2, submitted: txt(r[0]), id: String(r[1]), name: String(r[2]), date: txt(r[3]).slice(0, 10), start: txt(r[4]).slice(-5), end: txt(r[5]).slice(-5), brk: r[6], hours: Number(r[7]) || 0, amount: Number(r[11]) || 0, reason: r[12], photo: r[13], check: r[15], status: r[16], hotel: r[17], inTs: txt(r[18]) }))
    .filter(o => o.date.slice(0, 7) === m);
  const emp = sheet('Employees', H_EMP).getDataRange().getValues().slice(1)
    .map(r => ({ id: String(r[0]), name: String(r[1]), hotel: r[9] || '', registered: !!r[6] }));
  return { att, ot, emp, rate: OT_RATE };
}

function otStatus(d) {
  const who = checkAdmin(d);
  const s = sheet('OT', H_OT), row = Number(d.row);
  if (!(row >= 2 && row <= s.getLastRow())) throw new Error('ไม่พบรายการ OT');
  if (['อนุมัติ', 'ไม่อนุมัติ', 'รออนุมัติ'].indexOf(d.status) < 0) throw new Error('สถานะไม่ถูกต้อง');
  if (String(s.getRange(row, 2).getValue()) !== String(d.id)) throw new Error('ข้อมูลเปลี่ยนไปแล้ว กรุณารีเฟรช');
  s.getRange(row, 17).setValue(d.status);
  if (!s.getRange(1, 20).getValue()) s.getRange(1, 20).setValue('ผู้อนุมัติ');
  s.getRange(row, 20).setValue(d.status === 'รออนุมัติ' ? '' : who + ' · ' + now());
  return {};
}
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// คอลัมน์วันที่/เวลา เก็บเป็นข้อความ เพื่อไม่ให้ Sheets แปลงเขตเวลาเอง
const TEXT_COLS = { Employees: ['A', 'E', 'F', 'H'], Attendance: ['A', 'B', 'K'], OT: ['A', 'B', 'D', 'E', 'F', 'S'] };
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

// จำพิกัดโรงแรมไว้ 5 นาที (แก้ชีต Hotels แล้วมีผลภายใน 5 นาที)
function hotels() {
  const c = CacheService.getScriptCache(), hit = c.get('hotels');
  if (hit) return JSON.parse(hit);
  const v = loadHotels(); c.put('hotels', JSON.stringify(v), 300); return v;
}
function loadHotels() {
  const s = sheet('Hotels', H_HOTEL);
  if (s.getLastRow() < 2) s.getRange(2, 1, DEFAULT_HOTELS.length, H_HOTEL.length).setValues(DEFAULT_HOTELS);
  const by = {};
  s.getDataRange().getValues().slice(1).filter(r => r[0] && r[2] !== '' && r[3] !== '').forEach(r => {
    const id = String(r[0]).trim();
    if (!by[id]) by[id] = { id, name: String(r[1]), radius: Number(r[4]) || 50, maxAcc: Number(r[5]) || 65, pts: [] };
    by[id].pts.push({ lat: Number(r[2]), lng: Number(r[3]), note: String(r[6] || '') });
  });
  return Object.keys(by).map(k => Object.assign(by[k], { lat: by[k].pts[0].lat, lng: by[k].pts[0].lng }));
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
    return { profile: wageOf(findEmp(d.empId)), hotels: hotels() };
  }
  sheet('Employees', H_EMP).appendRow([d.empId, d.name, DEFAULT_WAGE.type, DEFAULT_WAGE.amount, DEFAULT_WAGE.start, DEFAULT_WAGE.end, d.credId, now(), d.device || '', d.hotel || '']);
  return { profile: DEFAULT_WAGE, hotels: hotels() };
}

function profile(d) { return { profile: wageOf(auth(d)), hotels: hotels() }; }

function savePhoto(dataUrl, name) {
  if (!dataUrl || dataUrl.indexOf('base64,') < 0) return '';
  const c = CacheService.getScriptCache(); let folder = null;
  try { const id = c.get('photo_folder'); if (id) folder = DriveApp.getFolderById(id); } catch (e) { folder = null; }
  if (!folder) { const it = DriveApp.getFoldersByName(PHOTO_FOLDER); folder = it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER); c.put('photo_folder', folder.getId(), 21600); }
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
  const dist = Math.min.apply(null, h.pts.map(p => distM({ lat: Number(g.lat), lng: Number(g.lng) }, p)));
  if (Number(g.acc) > h.maxAcc) throw new Error('สัญญาณ GPS ไม่แม่นพอ (±' + g.acc + ' ม.) ลองใหม่ใกล้หน้าต่างหรือทางเข้า');
  if (dist > h.radius) throw new Error('อยู่นอกพื้นที่ ' + h.name + ' (ห่าง ' + dist + ' ม.) ไม่สามารถสแกนได้');
  const url = savePhoto(d.photo, d.empId + '_' + d.type + '_' + ts.replace(/[: ]/g, '-'));
  const out = { ts, dist };
  if (d.type === 'out') out.ot = autoOt(d.empId, e.v[1], ts, h.name);
  sheet('Attendance', H_ATT).appendRow([ts, d.empId, e.v[1], d.type === 'in' ? 'เข้า' : 'ออก', g.lat, g.lng, g.acc, mapUrl(g), url, d.place || '', d.clientTime || '', 'ผ่าน', d.device || '', h.name, dist, 'ในพื้นที่']);
  return out;
}

const parseTs = t => new Date(String(t).replace(' ', 'T') + '+07:00');
const hm = dt => Utilities.formatDate(dt, TZ, 'HH:mm');

/** คำนวณ OT อัตโนมัติเมื่อสแกนออก: จับคู่กับการสแกนเข้าล่าสุดภายใน 24 ชม. */
function autoOt(empId, name, outTs, hotel) {
  const att = sheet('Attendance', H_ATT).getDataRange().getValues().slice(1), outT = parseTs(outTs);
  let inTs = null;
  for (let i = att.length - 1; i >= 0; i--) {
    const r = att[i]; if (String(r[1]) !== String(empId) || r[3] !== 'เข้า') continue;
    const t = parseTs(txt(r[0])); if (t < outT && outT - t <= 864e5) { inTs = txt(r[0]); break; }
  }
  if (!inTs) return { noIn: true };
  const start = new Date(Math.round(parseTs(inTs).getTime() / 36e5) * 36e5), normalEnd = new Date(start.getTime() + WORK_MIN * 6e4);
  const worked = Math.round((outT - parseTs(inTs)) / 6e4), otMin = Math.max(0, Math.floor((outT - normalEnd) / 6e4));
  const res = { inTs, start: hm(start), normalEnd: hm(normalEnd), worked, otMin, hours: Math.floor(otMin / 60), amount: Math.floor(otMin / 60) * OT_RATE }; // นับเฉพาะชั่วโมงเต็ม ไม่คิดเศษ
  const s = sheet('OT', H_OT), v = s.getDataRange().getValues();
  let row = -1; for (let i = 1; i < v.length; i++) if (String(v[i][1]) === String(empId) && txt(v[i][18]) === inTs) { row = i + 1; break; }
  if (res.hours < 1) { if (row > 0 && v[row - 1][16] === 'รออนุมัติ') s.getRange(row, 8, 1, 5).setValues([[0, 'OT อัตโนมัติ', 1, OT_RATE, 0]]); return res; }
  const src = 'สแกนเข้า ' + inTs.slice(11, 16) + ' (นับเริ่ม ' + res.start + ') → ออก ' + outTs.slice(11, 16) + ' · เลิกงานปกติ ' + res.normalEnd;
  const vals = [now(), String(empId), name, inTs.slice(0, 10), res.normalEnd, outTs.slice(11, 16), 0, res.hours, 'OT อัตโนมัติ', 1, OT_RATE, res.amount,
    row > 0 ? v[row - 1][12] : '', '', '', src, 'รออนุมัติ', hotel, inTs];
  if (row > 0) s.getRange(row, 1, 1, vals.length).setValues([vals]); else s.appendRow(vals);
  return res;
}

/** พนักงานใส่เหตุผล/งานที่ทำ ให้ OT ที่ยังรออนุมัติ */
function otNote(d) {
  auth(d);
  const s = sheet('OT', H_OT), v = s.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][1]) === String(d.empId) && txt(v[i][18]) === String(d.key)) {
    if (v[i][16] !== 'รออนุมัติ') throw new Error('รายการนี้แอดมินพิจารณาแล้ว แก้ไม่ได้');
    s.getRange(i + 1, 13).setValue("'" + String(d.reason || '').slice(0, 300)); return {};
  }
  throw new Error('ไม่พบรายการ OT');
}

function list(d) {
  auth(d);
  const id = String(d.empId);
  const att = sheet('Attendance', H_ATT).getDataRange().getValues().slice(1).filter(r => String(r[1]) === id).slice(-120)
    .map(r => ({ ts: txt(r[0]), type: r[3] === 'เข้า' ? 'in' : 'out', lat: r[4], lng: r[5], acc: r[6], photo: r[8], place: r[9], hotel: r[13], dist: r[14], inArea: r[15] }));
  const typ = { 'OT': 'ot', 'OT อัตโนมัติ': 'ot' };
  const ot = sheet('OT', H_OT).getDataRange().getValues().slice(1).filter(r => String(r[1]) === id).slice(-60)
    .map(r => ({ date: txt(r[3]).slice(0, 10), start: txt(r[4]).slice(-5), end: txt(r[5]).slice(-5), hours: r[7], type: typ[r[8]], rate: r[9], amount: r[11], reason: r[12], photo: r[13], check: r[15], status: r[16], hotel: r[17], inTs: txt(r[18]) }));
  return { att, ot };
}
