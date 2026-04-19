// พร้อมเพย์ไม่ฝังในโค้ด — ตั้งค่าได้จากหน้าแอดมิน (บันทึกลงชีท Store_Settings)
const SHEET_URL = "https://docs.google.com/spreadsheets/d/1F-6G6ZoPNmH16r-Yk9Mzy0TV3qw6Ic2y9TRgcMG5fzU/edit"; // URL ของ Google Sheets ของคุณ

/**
 * IMAGE_FOLDER_ID = ID โฟลเดอร์ใน Google Drive สำหรับเก็บรูปภาพเมนู (ถ้าไม่ตั้งจะใช้ค่าเริ่มต้น)
 */

/**
 * แปลง error ตอนสเปรดชีตเต็มขีดจำกัด 10 ล้านเซลล์ (Google Sheets)
 */
function translateSpreadsheetQuotaError_(e) {
  var s = String(e && e.message !== undefined ? e.message : e);
  if (/10000000|10,000,000|above the limit of/i.test(s)) {
    return 'ไฟล์ Google Sheets ใกล้หรือเกินขีดจำกัด 10 ล้านเซลล์ — ลบแถว/คอลัมน์ว่างที่ไม่ใช้ในทุกชีท (เลือกแถวใต้ข้อมูลจริง แล้วคลิกขวา > ลบแถว N แถวด้านบน) หรือย้ายข้อมูลเก่าไปไฟล์ใหม่ จากนั้นรัน getSpreadsheetCellStats() ใน Apps Script แล้วดู View > Logs';
  }
  return '';
}

/**
 * รันใน Apps Script Editor (เลือกฟังก์ชันนี้แล้ว Run) — ดู View > Logs
 * แสดงแถว×คอลัมน์ต่อชีท เพื่อหาว่าชีทไหนกินพื้นที่มากเกินไป
 */
function getSpreadsheetCellStats() {
  const ss = SpreadsheetApp.openByUrl(SHEET_URL);
  const sheets = ss.getSheets();
  let total = 0;
  const lines = [];
  for (var i = 0; i < sheets.length; i++) {
    const sh = sheets[i];
    const r = sh.getMaxRows();
    const c = sh.getMaxColumns();
    const cells = r * c;
    total += cells;
    lines.push(sh.getName() + ': แถว ' + r + ' × คอลัมน์ ' + c + ' ≈ ' + cells.toLocaleString() + ' เซลล์');
  }
  const summary = 'รวมทั้งไฟล์ ≈ ' + total.toLocaleString() + ' เซลล์ (จำกัด 10,000,000)';
  Logger.log(summary);
  Logger.log(lines.join('\n'));
  return { totalCells: total, perSheet: lines, summary: summary };
}

// ==========================================
// ฟังก์ชันตรวจสอบรหัสผ่าน
// ==========================================
/** สถานะที่ล็อกอินได้: ว่าง(ข้อมูลเก่า), ใช้งาน, Active / active — นอกนั้น = ระงับ */
function isStaffAccountActive_(status) {
  var s = String(status != null ? status : '').trim();
  if (s === '') return true;
  if (s === 'ใช้งาน') return true;
  var l = s.toLowerCase();
  if (l === 'active') return true;
  return false;
}

function verifyAdminLogin(username, password) {
  try {
    const ss = SpreadsheetApp.openByUrl(SHEET_URL);
    let sheet = ss.getSheetByName('Staff');
    
    // หากยังไม่มีชีท 'Staff' ให้สร้างใหม่และใส่ค่าเริ่มต้น
    if (!sheet) {
      sheet = ss.insertSheet('Staff');
      sheet.appendRow(['Username', 'Password', 'ชื่อ-สกุล', 'สิทธิ์การใช้งาน', 'สถานะ']);
      sheet.getRange('A1:E1').setFontWeight('bold').setBackground('#f3f3f3');
      sheet.appendRow(['admin', 'admin123', 'ผู้ดูแลระบบ', 'Owner', 'Active']);
      sheet.setColumnWidth(1, 150); sheet.setColumnWidth(2, 150); sheet.setColumnWidth(3, 200);
    }

    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      let dbUser = String(data[i][0]).trim();
      let dbPass = String(data[i][1]).trim();
      let dbName = String(data[i][2]).trim();
      let dbRole = String(data[i][3]).trim();
      let dbStatus = String(data[i][4]).trim();

      if (dbUser === String(username).trim() && dbPass === String(password).trim()) {
        if (!isStaffAccountActive_(dbStatus)) {
          return { authenticated: false, message: 'บัญชีนี้ถูกระงับการใช้งาน กรุณาติดต่อผู้ดูแล' };
        }
        return { authenticated: true, user: { name: dbName, role: dbRole } };
      }
    }
    
    return { authenticated: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };

  } catch (e) {
    Logger.log("Login Error: " + e.toString());
    const props = PropertiesService.getScriptProperties();
    let fallbackPass = props.getProperty('ADMIN_PASS') || 'admin123';
    
    if (password === fallbackPass && (username === 'admin' || username === '')) {
      return { authenticated: true, warning: 'เข้าสู่ระบบด้วยโหมดสำรอง', user: { name: 'Owner', role: 'Owner' } };
    }
    return { authenticated: false, message: 'เกิดข้อผิดพลาดในการเชื่อมต่อฐานข้อมูล' };
  }
}

// ==========================================
// 0.1 ตั้งค่าร้าน / พร้อมเพย์ / โลโก้ / ข้อความท้ายใบเสร็จ (ชีท Store_Settings แถว 2, คอลัมน์ A–D)
// ==========================================
const DEFAULT_STORE_NAME = 'ตำแรด แซ่บนัว';
const DEFAULT_LOGO_URL = 'https://lh3.googleusercontent.com/d/1P7QyEmBDrd16omNL91JyE3FnKOwipBZ5';
const DEFAULT_RECEIPT_THANK_YOU = '*** ขอบคุณที่ใช้บริการ ***';

/** ชีทเก่า 5 คอลัมน์ (มี storeAddress ที่ D) → รวมเป็น 4 คอลัมน์ โดยคง receiptFooter */
function migrateStoreSettingsLayoutIfNeeded_(sheet) {
  var lc = sheet.getLastColumn();
  if (lc === 0) return;
  var h4 = String(sheet.getRange(1, 4).getValue() || '').trim();
  var h5 = lc >= 5 ? String(sheet.getRange(1, 5).getValue() || '').trim() : '';
  if (lc >= 5 && (h5 === 'receiptFooter' || h4 === 'storeAddress')) {
    var r = sheet.getRange(2, 1, 2, 5).getDisplayValues()[0];
    var footer = String(r[4] != null ? r[4] : '').trim() || DEFAULT_RECEIPT_THANK_YOU;
    sheet.getRange(1, 1, 1, 4).setValues([['storeName', 'promptPayId', 'logoUrl', 'receiptFooter']]);
    sheet.getRange(2, 1, 1, 4).setValues([[r[0], r[1], r[2], footer]]);
    sheet.deleteColumn(5);
    sheet.getRange(2, 2).setNumberFormat('@');
    sheet.getRange(2, 4).setNumberFormat('@');
    return;
  }
  if (lc === 4 && h4 === 'storeAddress') {
    sheet.getRange(1, 4).setValue('receiptFooter');
    sheet.getRange(2, 4).setValue(DEFAULT_RECEIPT_THANK_YOU);
    sheet.getRange(2, 4).setNumberFormat('@');
  }
}

/** ให้มีอย่างน้อย 4 คอลัมน์ + หัว receiptFooter (รองรับชีทเก่า) */
function ensureStoreSettingsFourColumns_(sheet) {
  if (sheet.getLastColumn() === 0) {
    sheet.insertColumnBefore(1);
  }
  while (sheet.getLastColumn() < 4) {
    sheet.insertColumnAfter(sheet.getLastColumn());
  }
  var h4 = sheet.getRange(1, 4).getValue();
  if (h4 === '' || h4 === null) {
    sheet.getRange(1, 4).setValue('receiptFooter');
  }
}

/**
 * กู้เลขพร้อมเพย์ที่เคยถูกชีทเก็บเป็นตัวเลขแล้ว 0 นำหน้าหาย (เช่น 092... → 927... 9 หลัก)
 */
function normalizePromptPayIdStored(s) {
  s = String(s != null ? s : '').trim();
  if (s === '') return '';
  if (/^\d{9}$/.test(s)) return '0' + s;
  return s;
}

function getStoreSettings() {
  try {
    const ss = SpreadsheetApp.openByUrl(SHEET_URL);
    const sh = ss.getSheetByName('Store_Settings');
    if (!sh || sh.getLastRow() < 2) {
      return { storeName: DEFAULT_STORE_NAME, promptPayId: '', logoUrl: DEFAULT_LOGO_URL, receiptFooter: DEFAULT_RECEIPT_THANK_YOU };
    }
    migrateStoreSettingsLayoutIfNeeded_(sh);
    ensureStoreSettingsFourColumns_(sh);
    // getDisplayValues = ข้อความตามที่แสดมบนชีท ไม่แปลงเลขทำให้ 0 นำหน้าหาย
    const r = sh.getRange(2, 1, 2, 4).getDisplayValues()[0];
    var footer = String(r[3] != null ? r[3] : '').trim();
    return {
      storeName: String(r[0] != null ? r[0] : '').trim() || DEFAULT_STORE_NAME,
      promptPayId: normalizePromptPayIdStored(r[1]),
      logoUrl: String(r[2] != null ? r[2] : '').trim() || DEFAULT_LOGO_URL,
      receiptFooter: footer || DEFAULT_RECEIPT_THANK_YOU
    };
  } catch (e) {
    Logger.log('getStoreSettings: ' + e);
    return { storeName: DEFAULT_STORE_NAME, promptPayId: '', logoUrl: DEFAULT_LOGO_URL, receiptFooter: DEFAULT_RECEIPT_THANK_YOU };
  }
}

function getEffectivePromptPayId() {
  return getStoreSettings().promptPayId;
}

function saveStoreSettings(item) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openByUrl(SHEET_URL);
    let sheet = ss.getSheetByName('Store_Settings');
    if (!sheet) {
      sheet = ss.insertSheet('Store_Settings');
      sheet.appendRow(['storeName', 'promptPayId', 'logoUrl', 'receiptFooter']);
      sheet.getRange('A1:D1').setFontWeight('bold').setBackground('#f3f3f3');
    }
    migrateStoreSettingsLayoutIfNeeded_(sheet);
    ensureStoreSettingsFourColumns_(sheet);
    if (sheet.getLastRow() < 2) {
      sheet.appendRow(['', '', '', DEFAULT_RECEIPT_THANK_YOU]);
    }
    // บังคับคอลัมน์พร้อมเพย์เป็นข้อความ — ไม่ให้ชีทแปลงเป็นตัวเลขแล้ว 0 นำหน้าหาย
    sheet.getRange(2, 2).setNumberFormat('@');

    let logoUrl = String(item.logoUrl != null ? item.logoUrl : '').trim();

    if (item.imageFile && item.imageFile.indexOf('base64,') !== -1) {
      try {
        const parts = item.imageFile.split('base64,');
        const base64Data = parts[1];
        const mimeType = item.imageMimeType || 'image/png';
        const fileName = item.imageFileName || ('store_logo_' + new Date().getTime() + '.png');
        const decodedData = Utilities.base64Decode(base64Data);
        const blob = Utilities.newBlob(decodedData, mimeType, fileName);
        const props = PropertiesService.getScriptProperties();
        const folderId = props.getProperty('IMAGE_FOLDER_ID') || '1zTztG1OY9ZwNrmIYdM5kf5V_P_tRB2_Q';
        const folder = DriveApp.getFolderById(folderId);
        const file = folder.createFile(blob);
        try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (shareErr) { }
        logoUrl = 'https://lh3.googleusercontent.com/d/' + file.getId();
      } catch (imgErr) {
        throw new Error('อัปโหลดโลโก้ไม่สำเร็จ กรุณาลองรูปขนาดเล็กลงหรือตรวจสอบสิทธิ์โฟลเดอร์');
      }
    }

    const storeName = String(item.storeName != null ? item.storeName : '').trim() || DEFAULT_STORE_NAME;
    const promptPayId = String(item.promptPayId != null ? item.promptPayId : '').trim();
    var receiptFooter = String(item.receiptFooter != null ? item.receiptFooter : '').trim();
    if (!receiptFooter) receiptFooter = DEFAULT_RECEIPT_THANK_YOU;
    if (!logoUrl) {
      const cur = sheet.getRange(2, 3).getValue();
      logoUrl = String(cur != null ? cur : '').trim() || DEFAULT_LOGO_URL;
    }

    sheet.getRange(2, 1, 1, 4).setValues([[storeName, promptPayId, logoUrl, receiptFooter]]);
    sheet.getRange(2, 2).setNumberFormat('@');
    sheet.getRange(2, 4).setNumberFormat('@');
    return { success: true };
  } catch (e) {
    throw new Error(translateSpreadsheetQuotaError_(e) || e.message || e.toString());
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 0.2 จัดการพนักงาน (ชีท Staff — ไม่ส่งรหัสผ่านกลับไปที่ client)
// ==========================================
function getStaffList() {
  try {
    const ss = SpreadsheetApp.openByUrl(SHEET_URL);
    const sheet = ss.getSheetByName('Staff');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    return data.slice(1).map(function (r, i) {
      return {
        row: i + 2,
        username: String(r[0] != null ? r[0] : '').trim(),
        displayName: String(r[2] != null ? r[2] : '').trim(),
        role: String(r[3] != null ? r[3] : '').trim(),
        status: String(r[4] != null ? r[4] : '').trim()
      };
    }).filter(function (x) { return x.username !== ''; });
  } catch (e) {
    Logger.log('getStaffList: ' + e);
    return [];
  }
}

function saveStaffMember(item) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openByUrl(SHEET_URL);
    let sheet = ss.getSheetByName('Staff');
    if (!sheet) {
      sheet = ss.insertSheet('Staff');
      sheet.appendRow(['Username', 'Password', 'ชื่อ-สกุล', 'สิทธิ์การใช้งาน', 'สถานะ']);
      sheet.getRange('A1:E1').setFontWeight('bold').setBackground('#f3f3f3');
    }

    const username = String(item.username != null ? item.username : '').trim();
    if (!username) throw new Error('กรุณาระบุชื่อผู้ใช้');

    const rowIndex = parseInt(item.row, 10);
    const allData = sheet.getDataRange().getValues();
    for (let i = 1; i < allData.length; i++) {
      var u = String(allData[i][0] != null ? allData[i][0] : '').trim();
      if (u === username && (isNaN(rowIndex) || rowIndex <= 0 || i + 1 !== rowIndex)) {
        throw new Error('ชื่อผู้ใช้นี้ถูกใช้แล้ว');
      }
    }

    var pass = String(item.password != null ? item.password : '').trim();
    var displayName = String(item.displayName != null ? item.displayName : '').trim() || username;
    var role = String(item.role != null ? item.role : '').trim() || 'Cashier';
    var status = String(item.status != null ? item.status : '').trim() || 'Active';

    if (!isNaN(rowIndex) && rowIndex > 0) {
      var oldPass = String(sheet.getRange(rowIndex, 2).getValue() != null ? sheet.getRange(rowIndex, 2).getValue() : '').trim();
      if (!pass) pass = oldPass;
      if (!pass) throw new Error('กรุณาระบุรหัสผ่าน');
      sheet.getRange(rowIndex, 1, 1, 5).setValues([[username, pass, displayName, role, status]]);
    } else {
      if (!pass) throw new Error('กรุณาระบุรหัสผ่านสำหรับบัญชีใหม่');
      sheet.appendRow([username, pass, displayName, role, status]);
    }
    return { success: true };
  } catch (e) {
    throw new Error(translateSpreadsheetQuotaError_(e) || e.message || e.toString());
  } finally {
    lock.releaseLock();
  }
}

function deleteStaffMember(row) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openByUrl(SHEET_URL);
    const sheet = ss.getSheetByName('Staff');
    if (!sheet) return { success: true };
    const rowIndex = parseInt(row, 10);
    if (!isNaN(rowIndex) && rowIndex > 0) sheet.deleteRow(rowIndex);
    return { success: true };
  } catch (e) {
    throw new Error(translateSpreadsheetQuotaError_(e) || e.message || e.toString());
  } finally {
    lock.releaseLock();
  }
}

/** URL Web App แบบ https://script.google.com/macros/s/.../exec — ไม่ใช่ iframe userCodeAppPanel */
function getDeployedWebAppUrl_() {
  try {
    return ScriptApp.getService().getUrl() || '';
  } catch (e) {
    return '';
  }
}

/** เรียกจากหน้าเว็บผ่าน google.script.run เมื่อต้องการ URL /exec จริง (แก้ redirect ผิดปลายทาง) */
function getWebAppExecUrl() {
  return getDeployedWebAppUrl_();
}

// ==========================================
// 1. ฟังก์ชันหลักสำหรับเรียกหน้าเว็บ
// ==========================================
function doGet(e) {
  /** ลูกค้าสั่งอาหาร (index) — แนะนำ URL: ?page=Customer&table=...&token=... (รองรับเดิมแค่ ?table=...) */
  if (e.parameter && e.parameter.table) {
    let tableNo = e.parameter.table;
    let token = e.parameter.token || '';
    /** โหมด Take Away: ?table=Take%20Away&takeaway=1 — ไม่ต้องมี token (บันทึกโต๊ะเป็น "Take Away") */
    let isTakeaway = e.parameter.takeaway === '1' || String(tableNo).trim() === 'Take Away';

    if (!isTakeaway && token) {
       try {
         const ss = SpreadsheetApp.openByUrl(SHEET_URL);
         const sheet = ss.getSheetByName('Tables');
         if (sheet) {
            const tableData = sheet.getDataRange().getValues();
            let isValid = false;
            
            for (let i = 1; i < tableData.length; i++) {
              if (tableData[i][0].toString() === tableNo) {
                 if (tableData[i][2].toString().includes(token) && tableData[i][2].toString() !== "EXPIRED") {
                    isValid = true;
                 }
                 break;
              }
            }
            
            if (!isValid) {
               return HtmlService.createHtmlOutput('<div style="text-align:center; padding: 50px; font-family: sans-serif; background: #f8fafc; height: 100vh;"><div style="background: white; padding: 30px; border-radius: 20px; box-shadow: 0 10px 25px rgba(0,0,0,0.05); display: inline-block; max-width: 90%;"><h1 style="font-size: 50px; margin: 0;">❌</h1><h2 style="color: #e11d48;">ลิงก์สั่งอาหารนี้หมดอายุแล้ว</h2><p style="color: #64748b;">รายการอาหารถูกชำระเงินเรียบร้อยแล้ว<br>กรุณาติดต่อพนักงานหากต้องการเปิดโต๊ะใหม่ครับ</p></div></div>')
                  .setTitle('ลิงก์หมดอายุ')
                  .addMetaTag('viewport', 'width=device-width, initial-scale=1');
            }
         }
       } catch(err) {}
    } else if (!isTakeaway) {
        return HtmlService.createHtmlOutput('<div style="text-align:center; padding: 50px; font-family: sans-serif; background: #f8fafc; height: 100vh;"><div style="background: white; padding: 30px; border-radius: 20px; box-shadow: 0 10px 25px rgba(0,0,0,0.05); display: inline-block; max-width: 90%;"><h1 style="font-size: 50px; margin: 0;">⚠️</h1><h2 style="color: #e11d48;">ลิงก์ไม่ถูกต้อง</h2><p style="color: #64748b;">กรุณาสแกน QR Code จากทางร้านเพื่อสั่งอาหารครับ</p></div></div>')
                  .setTitle('ลิงก์ไม่ถูกต้อง')
                  .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }

    let template = HtmlService.createTemplateFromFile('index');
    template.tableNo = isTakeaway ? 'Take Away' : tableNo;
    template.webAppBaseUrl = getDeployedWebAppUrl_();
    return template.evaluate()
      .setTitle(isTakeaway ? 'สั่งอาหาร - Take Away' : ('สั่งอาหาร - โต๊ะ ' + tableNo))
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  /** logout.html — ?page=Logout */
  if (e.parameter && e.parameter.page === 'Logout') {
    var tLogout = HtmlService.createTemplateFromFile('logout');
    tLogout.webAppBaseUrl = getDeployedWebAppUrl_();
    return tLogout.evaluate()
      .setTitle('ออกจากระบบ')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  /** admin.html — ?page=Admin */
  if (e.parameter && e.parameter.page === 'Admin') {
    var tAdmin = HtmlService.createTemplateFromFile('admin');
    tAdmin.webAppBaseUrl = getDeployedWebAppUrl_();
    return tAdmin.evaluate()
      .setTitle('ระบบจัดการหลังบ้าน - ตำแรด แซ่บนัว')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  /** login.html — ค่าเริ่มต้น /exec หรือ ?page=Login */
  var tLogin = HtmlService.createTemplateFromFile('login');
  tLogin.webAppBaseUrl = getDeployedWebAppUrl_();
  return tLogin.evaluate()
    .setTitle('เข้าสู่ระบบแอดมิน - ตำแรด แซ่บนัว')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================
// 2. การจัดการเมนูอาหาร
// ==========================================
function getMenu() {
  try {
    const ss = SpreadsheetApp.openByUrl(SHEET_URL);
    const sheet = ss.getSheetByName('Menu');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return []; 
    
    return data.slice(1).map((r, i) => ({ 
      row: i + 2, 
      id: r[0], 
      name: r[1], 
      price: r[2], 
      category: r[3] ? r[3].toString().trim() : 'เมนูอื่นๆ', 
      image: r[4] 
    })).filter(item => item.name && item.name.toString().trim() !== ""); 
  } catch (e) {
    return [];
  }
}

function updateMenuItem(item) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openByUrl(SHEET_URL);
    let sheet = ss.getSheetByName('Menu');
    if (!sheet) {
      sheet = ss.insertSheet('Menu');
      sheet.appendRow(["รหัสเมนู", "ชื่อเมนู", "ราคา", "หมวดหมู่", "รูปภาพ(URL/Base64)"]);
    }

    let finalImageUrl = item.image || "";

    if (item.imageFile && item.imageFile.includes('base64,')) {
      try {
        const parts = item.imageFile.split('base64,');
        const base64Data = parts[1];
        const mimeType = item.imageMimeType || "image/png";
        const fileName = item.imageFileName || ("menu_img_" + new Date().getTime() + ".png");
        const decodedData = Utilities.base64Decode(base64Data);
        const blob = Utilities.newBlob(decodedData, mimeType, fileName);
        
        // ใช้ Script Properties สำหรับ Folder ID ถ้ายไม่มีให้ใช้ค่าเริ่มต้น
        const props = PropertiesService.getScriptProperties();
        const folderId = props.getProperty('IMAGE_FOLDER_ID') || "1zTztG1OY9ZwNrmIYdM5kf5V_P_tRB2_Q"; 
        const folder = DriveApp.getFolderById(folderId);
        
        const file = folder.createFile(blob);
        try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(shareErr) { }
        finalImageUrl = "https://lh3.googleusercontent.com/d/" + file.getId();
      } catch (imgErr) {
        throw new Error("ล้มเหลวในการประมวลผลรูปภาพ กรุณาตรวจสอบสิทธิ์ของ Folder หรือใช้รูปที่มีขนาดเล็กลง");
      }
    }

    const rowIndex = parseInt(item.row, 10);

    if (!isNaN(rowIndex) && rowIndex > 0) {
      sheet.getRange(rowIndex, 1, 1, 5).setValues([[item.id || "", item.name || "", Number(item.price) || 0, item.category || "", finalImageUrl]]);
    } else {
      sheet.appendRow([item.id || "", item.name || "", Number(item.price) || 0, item.category || "", finalImageUrl]);
    }
    return { success: true };
  } catch (e) {
    throw new Error(translateSpreadsheetQuotaError_(e) || e.message || e.toString());
  } finally {
    lock.releaseLock();
  }
}

function deleteMenuItem(row) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openByUrl(SHEET_URL);
    const sheet = ss.getSheetByName('Menu');
    if (sheet) {
      const rowIndex = parseInt(row, 10);
      if (!isNaN(rowIndex) && rowIndex > 0) sheet.deleteRow(rowIndex);
    }
    return { success: true };
  } catch (e) {
    throw e;
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 3. การจัดการออเดอร์ (ฝั่งลูกค้า)
// ==========================================
// Order_Details: A Detail_ID, B Order_ID, C Menu_Name, D Detail (หมายเหตุ), E Quantity, F Subtotal
// รองรับแถวเก่า (ไม่มีคอลัมน์ D): C=ชื่อ, D=จำนวน, E=ยอดรวมบรรทัด
function parseOrderDetailRow_(d) {
  const name = String(d[2] != null ? d[2] : '').trim();
  if (!name) return null;
  const sub5 = d[5];
  const hasColF = sub5 !== undefined && sub5 !== '' && sub5 !== null && !isNaN(Number(sub5));
  if (hasColF) {
    const detail = d[3] != null ? String(d[3]) : '';
    const qty = Number(d[4]) || 1;
    const subtotal = Number(sub5) || 0;
    return { name: name, note: detail, qty: qty, price: qty ? subtotal / qty : 0, total: subtotal };
  }
  const qty = Number(d[3]) || 1;
  const subtotal = Number(d[4]) || 0;
  return { name: name, note: '', qty: qty, price: qty ? subtotal / qty : 0, total: subtotal };
}

function submitOrder(tableNo, cartItems, total) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // ป้องกันลูกค้ารายอื่นแทรกแซงการบันทึก
    const ss = SpreadsheetApp.openByUrl(SHEET_URL);
    const orderId = "ORD-" + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd-HHmmss");
    const ordersSheet = ss.getSheetByName('Orders');
    ordersSheet.appendRow([orderId, tableNo, new Date(), 'รอรับออเดอร์', total, 'ยังไม่ชำระ']);

    const detailsSheet = ss.getSheetByName('Order_Details');
    cartItems.forEach(item => {
      const detail = item.note != null ? String(item.note) : (item.detail != null ? String(item.detail) : '');
      detailsSheet.appendRow([orderId + "-D", orderId, item.name, detail, item.qty, item.price * item.qty]);
    });
    const orderRowIndex = ordersSheet.getLastRow();
    return {
      success: true,
      message: "สั่งอาหารสำเร็จ!",
      qrCode: getQRLink(total),
      orderId: orderId,
      row: orderRowIndex
    };
  } catch (e) {
    return { success: false, message: "ระบบเกิดการทำงานทับซ้อน กรุณาลองใหม่อีกครั้ง: " + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function getTableOrders(tableNo) {
  try {
    const ss = SpreadsheetApp.openByUrl(SHEET_URL);
    const ordersSheet = ss.getSheetByName('Orders');
    const detailsSheet = ss.getSheetByName('Order_Details');
    
    if (!ordersSheet || !detailsSheet) return [];
    
    const ordersData = ordersSheet.getDataRange().getValues();
    const detailsData = detailsSheet.getDataRange().getValues();
    
    const tableOrders = ordersData.slice(1)
      .filter(function (r) {
        if (r[1].toString() !== tableNo.toString()) return false;
        var st = String(r[3] != null ? r[3] : '').trim();
        var pay = String(r[5] != null ? r[5] : '').trim();
        if (pay === 'ชำระแล้ว') return false;
        if (st === 'ยกเลิก' || pay === 'ยกเลิก') return false;
        return true;
      });
      
    return tableOrders.map(r => {
      const orderId = r[0];
      const items = detailsData.slice(1)
        .filter(d => d[1] === orderId)
        .map(d => {
          const p = parseOrderDetailRow_(d);
          if (!p) return '';
          return p.note ? `${p.name} (${p.note}) x${p.qty}` : `${p.name} x${p.qty}`;
        })
        .filter(Boolean)
        .join(',');

      return {
        id: orderId,
        time: r[2] instanceof Date ? Utilities.formatDate(r[2], "GMT+7", "HH:mm") : "-",
        status: r[3],
        total: r[4],
        items: items
      };
    });
  } catch (e) {
    return [];
  }
}

// ==========================================
// 4. การจัดการออเดอร์ (ฝั่งแอดมิน)
// ==========================================
function getActiveOrders() {
  try {
    const ss = SpreadsheetApp.openByUrl(SHEET_URL);
    const ordersSheet = ss.getSheetByName('Orders');
    const detailsSheet = ss.getSheetByName('Order_Details');
    
    if (!ordersSheet || !detailsSheet) return [];
    
    const ordersData = ordersSheet.getDataRange().getValues();
    const detailsData = detailsSheet.getDataRange().getValues();
    
    if (ordersData.length <= 1) return [];

    const ppId = String(getEffectivePromptPayId() || '').trim();
    function qrForOrderAmount_(amt) {
      if (!ppId) return '';
      const a = Number(amt);
      if (!a || a <= 0) return 'https://promptpay.io/' + ppId + '.png';
      return 'https://promptpay.io/' + ppId + '/' + a + '.png';
    }

    const detailsMap = {};
    for (let i = 1; i < detailsData.length; i++) {
      let d = detailsData[i];
      let ordId = d[1];
      const parsed = parseOrderDetailRow_(d);
      if (!parsed) continue;
      if (!detailsMap[ordId]) detailsMap[ordId] = [];
      detailsMap[ordId].push({
        name: parsed.name,
        note: parsed.note,
        qty: parsed.qty,
        price: parsed.price,
        total: parsed.total
      });
    }

    return ordersData.slice(1).map((r, i) => {
      return {
        row: i + 2, 
        id: r[0], 
        table: r[1], 
        time: r[2] instanceof Date ? Utilities.formatDate(r[2], "GMT+7", "HH:mm") : "-", 
        status: r[3], 
        total: r[4],
        payment: r[5],
        qrCode: qrForOrderAmount_(r[4]),
        items: detailsMap[r[0]] || [] 
      };
    })
      .filter(o => o.payment !== 'ชำระแล้ว' && o.id !== "")
      .filter(o => String(o.status).trim() !== 'ยกเลิก' && String(o.payment).trim() !== 'ยกเลิก')
      .reverse();
  } catch (e) {
    return [];
  }
}

function updateOrderStatus(row, status, paymentMethod = 'ไม่ระบุ', verificationRef) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // จัดคิวเพื่อป้องกันการบันทึกเงินทับซ้อน
    const ss = SpreadsheetApp.openByUrl(SHEET_URL);
    const sheet = ss.getSheetByName('Orders');
    const rowIndex = parseInt(row, 10);
    const refTrim = verificationRef != null && String(verificationRef).trim() !== ''
      ? String(verificationRef).trim()
      : '';
    
    if (!isNaN(rowIndex) && rowIndex > 0) {
      const currentStatus = String(sheet.getRange(rowIndex, 4).getValue() || '').trim();
      const currentPayment = String(sheet.getRange(rowIndex, 6).getValue() || '').trim();

      // แอดมินกด ทำอาหาร / เสิร์ฟครบ — ไม่เปลี่ยนสถานะแถวที่ยกเลิกแล้ว (คอลัมน์สถานะหรือชำระเงิน)
      if (status === 'กำลังทำอาหาร' || status === 'เสิร์ฟแล้ว') {
        if (currentStatus === 'ยกเลิก' || currentPayment === 'ยกเลิก') {
          return 'อัปเดตสถานะสำเร็จ';
        }
      }

      sheet.getRange(rowIndex, 4).setValue(status);
      
      if (status === 'ชำระเงินแล้ว') {
        const currentPaymentStatus = sheet.getRange(rowIndex, 6).getValue();
        
        if (currentPaymentStatus !== 'ชำระแล้ว') {
            sheet.getRange(rowIndex, 6).setValue('ชำระแล้ว'); 
            
            let financeSheet = ss.getSheetByName('Finance');
            if (!financeSheet) {
               financeSheet = ss.insertSheet('Finance');
               financeSheet.appendRow(['Trans_ID', 'Date', 'Type', 'Amount', 'Note']);
               financeSheet.getRange("A1:E1").setFontWeight("bold");
            }
            
            const orderId = sheet.getRange(rowIndex, 1).getValue();
            const tableNo = "โต๊ะ " + sheet.getRange(rowIndex, 2).getValue();
            const amount = sheet.getRange(rowIndex, 5).getValue();
            let note = tableNo;
            if (refTrim) {
              note += ' | อ้างอิง:' + refTrim;
            }
            if (paymentMethod === 'โอนชำระ' || (String(paymentMethod).indexOf('โอน') !== -1)) {
              note += ' | ตรวจยอดเข้าบัญชีแล้ว';
            }
            
            financeSheet.appendRow([orderId, new Date(), paymentMethod, amount, note]);
        }
      } else if (status === 'ยกเลิก') {
        sheet.getRange(rowIndex, 6).setValue('ยกเลิก');
      }
    }
    return "อัปเดตสถานะสำเร็จ";
  } catch (e) {
    return "เกิดข้อผิดพลาด: " + e.toString();
  } finally {
    lock.releaseLock();
  }
}

function saveEditedOrder(orderId, orderRow, newTotal, updatedItems) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openByUrl(SHEET_URL);
    const ordersSheet = ss.getSheetByName('Orders');
    const detailsSheet = ss.getSheetByName('Order_Details');

    const rowIndex = parseInt(orderRow, 10);
    if (!isNaN(rowIndex) && rowIndex > 0) {
      ordersSheet.getRange(rowIndex, 5).setValue(newTotal);
    }

    let detailsData = detailsSheet.getDataRange().getValues();
    for (let i = detailsData.length - 1; i >= 1; i--) {
      if (detailsData[i][1] === orderId) {
        detailsSheet.deleteRow(i + 1);
      }
    }

    updatedItems.forEach(item => {
      if (item.qty > 0) {
        const detail = item.note != null ? String(item.note) : (item.detail != null ? String(item.detail) : '');
        detailsSheet.appendRow([orderId + "-D", orderId, item.name, detail, item.qty, item.price * item.qty]);
      }
    });

    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function getAdminHistory() {
  try {
    const ss = SpreadsheetApp.openByUrl(SHEET_URL);
    const ordersSheet = ss.getSheetByName('Orders');
    const detailsSheet = ss.getSheetByName('Order_Details');

    if (!ordersSheet || !detailsSheet) return [];

    const ordersData = ordersSheet.getDataRange().getValues();
    const detailsData = detailsSheet.getDataRange().getValues();

    const paidOrders = ordersData.slice(1)
      .map((r, i) => ({
        row: i + 2, id: r[0], table: r[1], dateObj: r[2], status: r[3], total: r[4], payment: r[5]
      }))
      .filter(o => o.payment === 'ชำระแล้ว' && o.id !== "")
      .reverse()
      .slice(0, 100); 

    return paidOrders.map(o => {
      let timeStr = "-";
      try {
        if (o.dateObj instanceof Date) {
          timeStr = Utilities.formatDate(o.dateObj, "GMT+7", "dd/MM/yyyy HH:mm");
        } else {
          let fallbackDate = new Date(o.dateObj);
          if(!isNaN(fallbackDate.getTime())) timeStr = Utilities.formatDate(fallbackDate, "GMT+7", "dd/MM/yyyy HH:mm");
        }
      } catch(err) {}

      const items = detailsData.slice(1)
        .filter(d => d[1] === o.id)
        .map(d => {
          const p = parseOrderDetailRow_(d);
          if (!p) return '';
          return p.note ? `${p.name} (${p.note}) x${p.qty}` : `${p.name} x${p.qty}`;
        })
        .filter(Boolean)
        .join(', ');

      return {
        id: o.id,
        table: o.table,
        time: timeStr,
        items: items,
        total: o.total
      };
    });
  } catch (e) {
    return [];
  }
}

// ==========================================
// 5. สรุปรายได้ (ชีท Finance) + Top 10 + รายชั่วโมง + รายเดือน + ยอดวันนี้
// ==========================================
function computeTopProducts_(ss, start, end) {
  try {
    const ordersSheet = ss.getSheetByName('Orders');
    const detailsSheet = ss.getSheetByName('Order_Details');
    if (!ordersSheet || !detailsSheet) return [];

    const ordersData = ordersSheet.getDataRange().getValues();
    const paidIds = {};
    for (let i = 1; i < ordersData.length; i++) {
      const pay = String(ordersData[i][5] != null ? ordersData[i][5] : '').trim();
      if (pay !== 'ชำระแล้ว') continue;
      const dt = ordersData[i][2];
      const d = dt instanceof Date ? dt : new Date(dt);
      if (isNaN(d.getTime())) continue;
      if (d >= start && d <= end) {
        paidIds[String(ordersData[i][0])] = true;
      }
    }

    const detailsData = detailsSheet.getDataRange().getValues();
    const agg = {};
    for (let j = 1; j < detailsData.length; j++) {
      const oid = detailsData[j][1];
      if (!paidIds[oid]) continue;
      const parsed = parseOrderDetailRow_(detailsData[j]);
      if (!parsed) continue;
      const name = parsed.name;
      const qty = parsed.qty;
      const lineAmt = parsed.total;
      if (!agg[name]) agg[name] = { name: name, qty: 0, amount: 0 };
      agg[name].qty += qty;
      agg[name].amount += lineAmt;
    }

    const arr = [];
    for (const k in agg) {
      if (Object.prototype.hasOwnProperty.call(agg, k)) arr.push(agg[k]);
    }
    arr.sort((a, b) => b.qty - a.qty);
    return arr.slice(0, 10);
  } catch (e) {
    Logger.log('computeTopProducts_: ' + e);
    return [];
  }
}

function monthKeyToThaiLabel_(mk) {
  const th = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
  const parts = String(mk).split('-');
  if (parts.length < 2) return mk;
  const m = parseInt(parts[1], 10);
  const y = parts[0];
  if (isNaN(m) || m < 1 || m > 12) return mk;
  return th[m - 1] + ' ' + y;
}

function getFinanceSummary(startDateStr, endDateStr) {
  const emptyRet = {
    totalSales: 0, totalOrders: 0, avgSales: 0, cashTotal: 0, transferTotal: 0,
    labels: [], data: [],
    salesToday: 0,
    topProducts: [],
    hourlyLabels: [],
    hourlyData: [],
    monthlyLabels: [],
    monthlyData: []
  };

  try {
    const ss = SpreadsheetApp.openByUrl(SHEET_URL);
    const sheet = ss.getSheetByName('Finance');

    let start = startDateStr ? new Date(startDateStr) : null;
    if (start) start.setHours(0, 0, 0, 0);

    let end = endDateStr ? new Date(endDateStr) : null;
    if (end) end.setHours(23, 59, 59, 999);

    if (!start && !end) {
      end = new Date();
      start = new Date();
      start.setDate(start.getDate() - 6);
      start.setHours(0, 0, 0, 0);
    } else if (start && !end) {
      end = new Date(start);
      end.setHours(23, 59, 59, 999);
    } else if (!start && end) {
      start = new Date(end);
      start.setHours(0, 0, 0, 0);
    }

    const todayStr = Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy');
    let salesToday = 0;
    const hourlyData = [];
    const monthlyMap = {};
    for (let h = 0; h < 24; h++) hourlyData.push(0);

    if (!sheet) {
      emptyRet.topProducts = computeTopProducts_(ss, start, end);
      const hl = [];
      for (let h = 0; h < 24; h++) hl.push(String(h).length < 2 ? '0' + h : String(h));
      emptyRet.hourlyLabels = hl;
      emptyRet.hourlyData = hourlyData;
      return emptyRet;
    }

    const data = sheet.getDataRange().getValues();

    let totalSales = 0;
    let totalOrders = 0;
    let cashTotal = 0;
    let transferTotal = 0;
    const labels = [];
    const salesDataMap = {};

    const diffDays = Math.round((end - start) / (1000 * 60 * 60 * 24));

    if (diffDays <= 60) {
      let currentDate = new Date(start);
      while (currentDate <= end) {
        const dStr = Utilities.formatDate(currentDate, 'GMT+7', 'dd/MM/yyyy');
        labels.push(dStr);
        salesDataMap[dStr] = 0;
        currentDate.setDate(currentDate.getDate() + 1);
      }
    }

    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      let orderDate = r[1] instanceof Date ? r[1] : new Date(r[1]);

      if (isNaN(orderDate.getTime())) continue;

      const orderDateStr = Utilities.formatDate(orderDate, 'GMT+7', 'dd/MM/yyyy');
      if (orderDateStr === todayStr) {
        const amountToday = Number(r[3]) || 0;
        salesToday += amountToday;
      }

      if (orderDate >= start && orderDate <= end) {
        const payType = r[2] ? r[2].toString().trim() : '';
        const amount = Number(r[3]) || 0;

        totalSales += amount;
        totalOrders++;

        if (payType === 'เงินสด') {
          cashTotal += amount;
        } else {
          transferTotal += amount;
        }

        if (diffDays > 60) {
          if (salesDataMap[orderDateStr] === undefined) {
            salesDataMap[orderDateStr] = 0;
            labels.push(orderDateStr);
          }
          salesDataMap[orderDateStr] += amount;
        } else {
          if (salesDataMap[orderDateStr] !== undefined) {
            salesDataMap[orderDateStr] += amount;
          }
        }

        const hour = parseInt(Utilities.formatDate(orderDate, 'GMT+7', 'H'), 10);
        if (!isNaN(hour) && hour >= 0 && hour <= 23) {
          hourlyData[hour] += amount;
        }

        const mk = Utilities.formatDate(orderDate, 'GMT+7', 'yyyy-MM');
        monthlyMap[mk] = (monthlyMap[mk] || 0) + amount;
      }
    }

    if (diffDays > 60) {
      labels.sort((a, b) => {
        const partA = a.split('/');
        const partB = b.split('/');
        return new Date(partA[2], partA[1] - 1, partA[0]) - new Date(partB[2], partB[1] - 1, partB[0]);
      });
    }

    const shortLabels = labels.map(l => l.substring(0, 5));
    const salesData = labels.map(l => salesDataMap[l] || 0);

    const monthKeys = Object.keys(monthlyMap).sort();
    const monthlyLabels = monthKeys.map(monthKeyToThaiLabel_);
    const monthlyData = monthKeys.map(k => monthlyMap[k]);

    const hourlyLabels = [];
    for (let hh = 0; hh < 24; hh++) {
      hourlyLabels.push(String(hh).length < 2 ? '0' + hh : String(hh));
    }

    const topProducts = computeTopProducts_(ss, start, end);

    return {
      totalSales: totalSales,
      totalOrders: totalOrders,
      avgSales: totalOrders > 0 ? (totalSales / totalOrders) : 0,
      cashTotal: cashTotal,
      transferTotal: transferTotal,
      labels: shortLabels,
      data: salesData,
      salesToday: salesToday,
      topProducts: topProducts,
      hourlyLabels: hourlyLabels,
      hourlyData: hourlyData,
      monthlyLabels: monthlyLabels,
      monthlyData: monthlyData
    };
  } catch (e) {
    Logger.log('Finance Error: ' + e);
    return emptyRet;
  }
}

// ==========================================
// 6. เครื่องมืออื่นๆ (Utils)
// ==========================================
function getQRLink(amount) {
  const id = String(getEffectivePromptPayId() || '').trim();
  if (!id) return '';
  if (!amount || amount <= 0) return `https://promptpay.io/${id}.png`;
  return `https://promptpay.io/${id}/${amount}.png`;
}

// ==========================================
// 7. จัดการลิงก์และ QR Code
// ==========================================
function saveQRCodeToSheet(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const ss = SpreadsheetApp.openByUrl(SHEET_URL);
    const sheet = ss.getSheetByName("Tables"); 
    if (!sheet) throw new Error("ไม่พบชีทชื่อ Tables");
    
    const tableData = sheet.getDataRange().getValues();
    let rowIndex = -1;
    
    for (let i = 1; i < tableData.length; i++) {
      if (tableData[i][0].toString() === data.tableNum.toString()) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex > -1) {
      sheet.getRange(rowIndex, 2).setValue("ใช้งาน");
      sheet.getRange(rowIndex, 3).setValue(data.finalUrl);
      return { success: true, duplicate: true }; 
    } else {
      sheet.appendRow([data.tableNum, "ใช้งาน", data.finalUrl]);
      return { success: true };
    }
  } catch (e) {
    throw e;
  } finally {
    lock.releaseLock();
  }
}

function invalidateQRLink(tableNo) {
  try {
    const ss = SpreadsheetApp.openByUrl(SHEET_URL);
    const sheet = ss.getSheetByName("Tables"); 
    if (!sheet) return;
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === tableNo.toString()) {
        sheet.getRange(i + 1, 2).setValue("หมดอายุ (เคลียร์บิลแล้ว)");
        sheet.getRange(i + 1, 3).setValue("EXPIRED"); 
        break;
      }
    }
  } catch(e) {}
}

// ==========================================
// 8. แจ้งเรียกเช็คบิล (ฝั่งลูกค้า)
// ==========================================
function requestBill(tableNo) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const ss = SpreadsheetApp.openByUrl(SHEET_URL);
    const sheet = ss.getSheetByName('Orders');
    const data = sheet.getDataRange().getValues();
    
    var updatedCount = 0;
    for (var i = 1; i < data.length; i++) {
      if (data[i][1].toString() !== tableNo.toString()) continue;
      var status = String(data[i][3] != null ? data[i][3] : '').trim();
      var payment = String(data[i][5] != null ? data[i][5] : '').trim();
      // ไม่แตะแถวที่ยกเลิกแล้ว (สถานะหรือคอลัมน์ชำระเงิน) หรือชำระครบแล้ว
      if (payment === 'ชำระแล้ว') continue;
      if (status === 'ยกเลิก' || payment === 'ยกเลิก') continue;
      sheet.getRange(i + 1, 4).setValue('เรียกเช็คบิล');
      updatedCount++;
    }
    
    return { success: updatedCount > 0, updatedCount: updatedCount };
  } catch (e) {
    throw e;
  } finally {
    lock.releaseLock();
  }
}