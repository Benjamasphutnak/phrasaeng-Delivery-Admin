const SPREADSHEET_ID = '1-DIW9DM_gSDsJvALssPX2sc-uUxZwo3_UgDqyHOJ_Ms';
const SHEET_NAME = ''; // ใส่ชื่อแผ่นงาน เช่น 'Orders' ถ้าเว้นว่างจะใช้แผ่นแรกสุด

// ลิงก์รูปโลโก้ .png สำหรับไอคอนแอพ (บนแท็บเบราว์เซอร์ / ตอนเพิ่มไปหน้าจอโฮม)
// เว้นว่างไว้ได้ แอพยังใช้งานปกติ แค่ไอคอนจะเป็นของ Google
const LOGO_URL = '';

// ============================================================
// 🔗 เชื่อมกับระบบสั่งของ "พระแสงเดลิเวอรี่" (คนละบัญชี) — ระบบสั่งของส่งงานที่ส่งสำเร็จเข้ามาทางนี้
// APP_SECRET ต้องตรงกับ ADMIN_APP_SECRET ใน Code.gs ของระบบสั่งของ · DELIVERY_URL = ลิงก์ /exec ของระบบสั่งของ
// ⚠️ เว็บแอปนี้ต้อง Deploy แบบ "ผู้ที่มีสิทธิ์เข้าถึง: ทุกคน" ถึงจะรับข้อมูลจากอีกระบบได้ (ป้องกันด้วย APP_SECRET)
// ============================================================
const APP_SECRET = '9V1KIvL4VyvphMMQYVSGw-Yn1MuSXnHC';
const DELIVERY_URL = 'https://script.google.com/macros/s/AKfycbyrqSO6l5CNmj_BuyKv7XCCOI4eNHoxNvcxzcSdA7FP3eu1pJ7QuHpWwEhW0kcIH6o0/exec';
const DEFAULT_ADMIN_FEE = 5; // ค่าจ้างแอดมินต่อออเดอร์ (ออเดอร์อัตโนมัติ)
const DEFAULT_SYS_FEE = 5;   // ค่าระบบต่อออเดอร์
// คอลัมน์เพิ่มเติมของออเดอร์อัตโนมัติ: N(14)=รายการสินค้า O(15)=เบอร์ไรเดอร์ P(16)=แหล่งที่มา "auto" Q(17)=เวลาส่งสำเร็จ

function jsonOut_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!b.secret || b.secret !== APP_SECRET) return jsonOut_({ success: false, message: 'unauthorized' });
    if (b.action === 'upsertOrder') return jsonOut_(upsertAutoOrder_(b.order || {}));
    if (b.action === 'setPaid') return jsonOut_(setPaidFromDelivery_(b.orderId, b.paid === true));
    return jsonOut_({ success: false, message: 'unknown action' });
  } catch (err) {
    return jsonOut_({ success: false, message: String(err) });
  }
}

// รับงานใหม่จากระบบสั่งของ: ถ้ามีเลขออเดอร์นี้อยู่แล้วจะไม่เขียนทับ (แอดมินอาจแก้ยอดไปแล้ว)
function upsertAutoOrder_(o) {
  const no = String(o.orderId || '').trim();
  if (!no) return { success: false, message: 'no orderId' };
  return withLock_(function () {
    const sheet = getSheet_();
    const last = sheet.getLastRow();
    if (last >= 2) {
      const ids = sheet.getRange(2, 1, last - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === no) return { success: true, existed: true };
      }
    }
    const goods = Number(o.goods) || 0, ship = Number(o.ship) || 0;
    const paid = Number(o.paid) || (goods + ship);
    const net = goods + ship - DEFAULT_ADMIN_FEE - DEFAULT_SYS_FEE;
    sheet.appendRow([
      no, o.date || '', "'" + (o.time || ''), o.shop || '', o.rider || '',
      goods, ship, paid, DEFAULT_ADMIN_FEE, DEFAULT_SYS_FEE, net,
      'โอนแล้ว', 'รอโอน',
      o.items || '', "'" + (o.riderPhone || ''), 'auto', "'" + (o.deliveredAt || '')
    ]);
    return { success: true, existed: false };
  });
}

// ระบบสั่งของแจ้งมาว่าแอดมินกดโอนในปุ่ม Telegram แล้ว → ติ๊กโอนแล้วที่นี่
function setPaidFromDelivery_(orderNo, paid) {
  return withLock_(function () {
    const sheet = getSheet_();
    const last = sheet.getLastRow();
    if (last < 2) return { success: false };
    const ids = sheet.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(orderNo)) {
        sheet.getRange(i + 2, 13).setValue(paid ? 'โอนแล้ว' : 'รอโอน');
        return { success: true };
      }
    }
    return { success: false, message: 'not found' };
  });
}

// แจ้งระบบสั่งของว่าโอนแล้ว/ยกเลิก (เฉพาะออเดอร์ที่มาจากระบบสั่งของ) — คืน true ถ้าสำเร็จ
function notifyDelivery_(orderNo, paid) {
  try {
    const res = UrlFetchApp.fetch(DELIVERY_URL, {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true, followRedirects: true,
      payload: JSON.stringify({ action: 'adminPayoutPaid', secret: APP_SECRET, orderId: String(orderNo), paid: paid === true })
    });
    return JSON.parse(res.getContentText()).success === true;
  } catch (err) {
    console.warn('แจ้งระบบสั่งของไม่สำเร็จ ' + orderNo + ': ' + err);
    return false;
  }
}

function doGet() {
  const out = HtmlService.createHtmlOutputFromFile('Phrasaeng')
      .setTitle('แอดมินพระแสงเดลิเวอรี่')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  if (LOGO_URL) {
    try { out.setFaviconUrl(LOGO_URL); } catch (e) { console.warn('ใช้ลิงก์โลโก้ไม่ได้: ' + e); }
  }
  return out;
}

/* ---------- ตัวช่วย ---------- */
function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getSheets()[0];
  if (!sh) throw new Error('ไม่พบแผ่นงานชื่อ ' + SHEET_NAME);
  return sh;
}

// กันแอดมินหลายคนกดบันทึกพร้อมกันแล้วข้อมูลทับกัน
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function toYmd_(v, tz) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  return String(v).slice(0, 10);
}

// ออกเลขออเดอร์ฝั่งเซิร์ฟเวอร์ ไม่ให้ซ้ำแม้มีการลบออเดอร์กลางวัน
function nextOrderNo_(rows, dateStr) {
  const prefix = 'PS-' + dateStr.replace(/-/g, '').slice(2) + '-';
  let max = 0;
  for (let i = 1; i < rows.length; i++) {
    const no = String(rows[i][0]);
    if (no.indexOf(prefix) === 0) {
      const n = parseInt(no.slice(prefix.length), 10);
      if (n > max) max = n;
    }
  }
  return prefix + String(max + 1).padStart(3, '0');
}

/* ---------- อ่านข้อมูล ---------- */
function getOrders() {
  const sheet = getSheet_();
  const tz = sheet.getParent().getSpreadsheetTimeZone();
  const range = sheet.getDataRange();
  const rows = range.getValues();
  const disp = range.getDisplayValues(); // ใช้ค่าที่แสดงสำหรับเวลา (ส่ง Date กลับไปหน้าเว็บไม่ได้)
  if (rows.length <= 1) return [];

  const data = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row[0]) continue;
    const no = String(row[0]);
    data.push({
      id: no,
      no: no,
      date: toYmd_(row[1], tz),
      time: disp[i][2],
      shop: String(row[3]),
      rider: String(row[4]),
      goods: Number(row[5]) || 0,
      ship: Number(row[6]) || 0,
      paid: Number(row[7]) || 0,
      admin: Number(row[8]) || 0,
      sys: Number(row[9]) || 0,
      net: Number(row[10]) || 0,
      st: row[11] === 'ยกเลิก' ? 'cancel' : (row[11] === 'รอโอน' ? 'wait' : 'paid'),
      rPaid: row[12] === true || row[12] === 'TRUE' || row[12] === 'โอนแล้ว',
      items: String(row[13] || ''),
      riderPhone: String(disp[i][14] || '').replace(/[^0-9]/g, ''),
      src: String(row[15] || ''),
      doneAt: String(disp[i][16] || '')
    });
  }
  return data.reverse();
}

/* ---------- บันทึก / แก้ไข ---------- */
function saveOrderToSheet(order) {
  return withLock_(function () {
    const sheet = getSheet_();
    const rows = sheet.getDataRange().getValues();

    let rowIndex = -1;
    if (order.editNo) {
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]) === String(order.editNo)) { rowIndex = i + 1; break; }
      }
    }
    const no = rowIndex > -1 ? String(order.editNo) : nextOrderNo_(rows, order.date);

    const netVal = Number(order.goods) + Number(order.ship) - Number(order.admin) - Number(order.sys);
    const rowData = [
      no,
      order.date,
      "'" + order.time, // เก็บเวลาเป็นข้อความ ไม่ให้ชีตแปลงเป็นวันที่
      order.shop,
      order.rider,
      Number(order.goods),
      Number(order.ship),
      Number(order.paid),
      Number(order.admin),
      Number(order.sys),
      netVal,
      order.st === 'cancel' ? 'ยกเลิก' : (order.st === 'wait' ? 'รอโอน' : 'โอนแล้ว'),
      order.rPaid ? 'โอนแล้ว' : 'รอโอน'
    ];

    if (rowIndex > -1) {
      sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
    return no;
  });
}

/* ---------- ลบ ---------- */
function deleteOrderFromSheet(id) {
  return withLock_(function () {
    const sheet = getSheet_();
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) { sheet.deleteRow(i + 1); break; }
    }
    return 'success';
  });
}

/* ---------- โอนให้ไรเดอร์ ----------
   ส่งรายการเลขออเดอร์ที่เห็นบนหน้าจอมาด้วย จะติ๊กเฉพาะออเดอร์เหล่านั้น
   (กันออเดอร์ที่เพิ่งเข้ามาระหว่างกดถูกติ๊กว่าโอนแล้วทั้งที่ยังไม่ได้โอน) */
function updateRiderPaymentInSheet(riderName, orderNos) {
  return withLock_(function () {
    const sheet = getSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { ok: true, autoNos: [] };
    const rows = sheet.getRange(2, 1, lastRow - 1, 16).getValues();
    const only = Array.isArray(orderNos) && orderNos.length ? new Set(orderNos.map(String)) : null;

    const autoNos = [];
    const col13 = rows.map(function (r) {
      const match = r[4] === riderName && r[11] !== 'ยกเลิก' && (!only || only.has(String(r[0])));
      if (match && r[12] !== 'โอนแล้ว' && String(r[15]) === 'auto') autoNos.push(String(r[0]));
      return [match ? 'โอนแล้ว' : r[12]];
    });
    sheet.getRange(2, 13, col13.length, 1).setValues(col13);
    return { ok: true, autoNos: autoNos };
  });
}

/* ---------- ล้างทั้งหมด ---------- */
function clearAllSheetData() {
  return withLock_(function () {
    const sheet = getSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
    return 'success';
  });
}
/* ---------- ติ๊กโอนให้ไรเดอร์ทีละออเดอร์ (paid=false คือยกเลิก กลับเป็นรอโอน) ---------- */
function setOrderRiderPaid(orderNo, paid) {
  return withLock_(function () {
    const sheet = getSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('ไม่พบออเดอร์ ' + orderNo);
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(orderNo)) {
        sheet.getRange(i + 2, 13).setValue(paid ? 'โอนแล้ว' : 'รอโอน');
        const isAuto = String(sheet.getRange(i + 2, 16).getValue()) === 'auto';
        return { ok: true, auto: isAuto, orderNo: String(orderNo) };
      }
    }
    throw new Error('ไม่พบออเดอร์ ' + orderNo);
  });
}

// แจ้งระบบสั่งของ (เรียกแยกจากการบันทึก เพื่อไม่ให้ติดล็อกชีตระหว่างรอเว็บอีกฝั่ง) — client เรียกต่อหลังติ๊กสำเร็จ
function syncPaidToDelivery(orderNo, paid) {
  return notifyDelivery_(orderNo, paid === true);
}

/* ---------- รายชื่อร้านค้า / ไรเดอร์ (เก็บในชีต ไม่หายเวลาอัปเดตโค้ด) ----------
   แผ่นงาน "ตั้งค่ารายชื่อ" : คอลัมน์ A = ร้านค้า, คอลัมน์ B = ไรเดอร์
   ระบบสร้างให้เองครั้งแรก และแก้ในชีตตรง ๆ ก็ได้ */
const LIST_SHEET = 'ตั้งค่ารายชื่อ';
const LIST_COL = { shop: 1, rider: 2 };
const DEFAULT_SHOPS = ['ปรุงยาเภสัช','ซุปเปอร์ทเวนตี้','Amazon','91โดนกะแฟ','ร้านกล้วยทอดหน้าโรงบาลพระแสง',
  'เหนียวหมูฟีนิกซ์','KFC','สลัดโรลใส่ใจ','ร้านหมูซิ่งหมูสด','แช่บ๊วย ด้วยรัก','บีพี ช็อป','ร้านน้ำคอชาใต้',
  'คอฟฟี่เฮ้าส์','ร้านน้ำติดนมอมมุก'];
const DEFAULT_RIDERS = ['ประเสริฐ พุฒนาค','เบญจมาศ พุฒนาค','วีระยุทธิ์ เพชรทอง'];

function getListSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(LIST_SHEET);
  if (!sh) {
    // ใส่ไว้ท้ายสุด จะได้ไม่ไปแทนที่แผ่นออเดอร์ (แผ่นแรก)
    sh = ss.insertSheet(LIST_SHEET, ss.getSheets().length);
    sh.getRange(1, 1, 1, 2).setValues([['ร้านค้า', 'ไรเดอร์']]).setFontWeight('bold');
    // ครั้งแรก: รายชื่อตั้งต้น + ชื่อร้าน/ไรเดอร์ทุกชื่อที่เคยมีในออเดอร์ (กู้ร้านที่เคยเพิ่มไว้กลับมา)
    const shops = DEFAULT_SHOPS.slice(), riders = DEFAULT_RIDERS.slice();
    getSheet_().getDataRange().getValues().slice(1).forEach(function (r) {
      const s = String(r[3] || '').trim(), d = String(r[4] || '').trim();
      if (s && shops.indexOf(s) < 0) shops.push(s);
      if (d && riders.indexOf(d) < 0) riders.push(d);
    });
    writeCol_(sh, 'shop', shops);
    writeCol_(sh, 'rider', riders);
  }
  return sh;
}
function readCol_(sh, type) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, LIST_COL[type], last - 1, 1).getValues()
    .map(function (r) { return String(r[0]).trim(); })
    .filter(function (s) { return s; });
}
function writeCol_(sh, type, names) {
  const col = LIST_COL[type];
  const last = Math.max(sh.getLastRow(), 2);
  sh.getRange(2, col, last - 1, 1).clearContent();
  if (names.length) sh.getRange(2, col, names.length, 1).setValues(names.map(function (n) { return [n]; }));
}

function getLists() {
  return withLock_(function () {
    const sh = getListSheet_();
    return { shops: readCol_(sh, 'shop'), riders: readCol_(sh, 'rider') };
  });
}

function addListItem(type, name) {
  if (!LIST_COL[type]) throw new Error('ประเภทไม่ถูกต้อง');
  name = String(name || '').trim();
  if (!name) throw new Error('ยังไม่ได้ใส่ชื่อ');
  return withLock_(function () {
    const sh = getListSheet_();
    const list = readCol_(sh, type);
    if (list.indexOf(name) < 0) { list.push(name); writeCol_(sh, type, list); }
    return list;
  });
}

// ลบออกจากรายชื่อเท่านั้น ออเดอร์เก่าของร้านนี้ยังอยู่ครบ
function removeListItem(type, name) {
  if (!LIST_COL[type]) throw new Error('ประเภทไม่ถูกต้อง');
  return withLock_(function () {
    const sh = getListSheet_();
    const list = readCol_(sh, type).filter(function (n) { return n !== String(name); });
    writeCol_(sh, type, list);
    return list;
  });
}
