// ============================================================
//  GOOGLE APPS SCRIPT — REST API Backend
//  อัปเดต: เพิ่ม getAttendance + deleteAttendanceRow
//  Deploy > New deployment > Web App
//  Execute as: Me | Who has access: Anyone
// ============================================================

//==================== TELEGRAM ====================
const TELEGRAM_TOKEN   = "7443792854:AAH6mB8K5n4_Ny9rka3a-2RcMpECRjPVxeE";
const TELEGRAM_CHAT_ID = "-4577321374";
//==================================================

function doGet(e) {
  const action = e.parameter.action;
  let result;
  if      (action === 'getConfig')      result = getConfig();
  else if (action === 'getKnownFaces')  result = getKnownFaces();
  else if (action === 'getAttendance')  result = getAttendance();   // ← ใหม่
  else result = { error: 'Unknown action: ' + action };

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// สร้าง JSON response (ต้องคืนเป็น JSON เสมอ ไม่เช่นนั้น GAS จะส่งหน้า error
// ที่ไม่มี CORS header กลับมา ทำให้ฝั่งเว็บขึ้นว่า "Failed to fetch" แม้บันทึกสำเร็จแล้ว)
function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let data;
  try { data = JSON.parse(e.postData.contents); }
  catch (err) {
    return jsonOut({ success: false, error: 'Invalid JSON body' });
  }

  const action = data.action;
  try {
    let result;
    if      (action === 'registerUser')        result = registerUser(data.name, data.position, data.faceDescriptor);
    else if (action === 'logAttendance')       result = logAttendance(data.name, data.position, data.status, data.lat, data.lng);
    else if (action === 'saveConfig')          result = saveConfig(data.lat, data.lng, data.radius);
    else if (action === 'deleteAttendanceRow') result = deleteAttendanceRow(data.rowIndex);
    else if (action === 'clearAllAttendance')  result = clearAllAttendance();
    else result = { success: false, error: 'Unknown action: ' + action };
    return jsonOut(result);
  } catch (err) {
    return jsonOut({ success: false, error: err.message });
  }
}

// ============================================================
//  ลงทะเบียนใบหน้า
//  Users sheet: A=ชื่อ | B=ตำแหน่ง | C=Descriptor | D=วันที่
// ============================================================
function registerUser(name, position, faceDescriptor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Users');
  if (!sheet) {
    sheet = ss.insertSheet('Users');
    sheet.appendRow(['ชื่อ-นามสกุล', 'ตำแหน่ง', 'Face Descriptor', 'วันที่ลงทะเบียน']);
    sheet.getRange('1:1').setFontWeight('bold');
    sheet.getRange('1:1').setBackground('#4f46e5');
    sheet.getRange('1:1').setFontColor('#ffffff');
    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 140);
    sheet.setColumnWidth(3, 300);
    sheet.setColumnWidth(4, 160);
  }
  sheet.appendRow([name, position || '-', JSON.stringify(faceDescriptor), new Date()]);
  return { success: true, message: 'บันทึกข้อมูลใบหน้า "' + name + '" (' + position + ') เรียบร้อย' };
}

// ============================================================
//  ดึงข้อมูลใบหน้าทั้งหมด
//  Users sheet: A=ชื่อ | B=ตำแหน่ง | C=Descriptor | D=วันที่
// ============================================================
function getKnownFaces() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  let users = [];
  for (let i = 1; i < data.length; i++) {
    const name     = data[i][0];
    const position = data[i][1];
    const jsonStr  = data[i][2];
    if (name && jsonStr) {
      try {
        users.push({ label: name, position: position || '', descriptor: JSON.parse(jsonStr) });
      } catch (e) {}
    }
  }
  return users;
}

// ============================================================
//  บันทึกเวลาเข้า/ออกงาน
//  Attendance sheet: A=ชื่อ | B=ตำแหน่ง | C=สถานะ | D=เวลา | E=วันที่ | F=Lat | G=Lng | H=Map
// ============================================================
function logAttendance(name, position, status, lat, lng) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Attendance');
  if (!sheet) {
    sheet = ss.insertSheet('Attendance');
    sheet.appendRow(['ชื่อ-นามสกุล', 'ตำแหน่ง', 'สถานะ', 'เวลา', 'วันที่', 'Latitude', 'Longitude', 'Google Map Link']);
    sheet.getRange('1:1').setFontWeight('bold');
    sheet.getRange('1:1').setBackground('#059669');
    sheet.getRange('1:1').setFontColor('#ffffff');
    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 140);
    sheet.setColumnWidth(3, 160);
    sheet.setColumnWidth(8, 220);
  }

  const now     = new Date();
  const tz      = Session.getScriptTimeZone();
  const dateStr = Utilities.formatDate(now, tz, 'd/M/yyyy');
  const timeStr = Utilities.formatDate(now, tz, 'HH:mm:ss');
  const mapLink = (lat && lng) ? 'https://www.google.com/maps?q=' + lat + ',' + lng : '-';

  sheet.appendRow([
    name,
    position || '-',
    status   || 'ลงเวลามาทำงาน',
    timeStr,
    "'" + dateStr,
    lat  || '-',
    lng  || '-',
    mapLink
  ]);

  // แจ้งเตือน Telegram ล้มเหลวไม่ควรทำให้การบันทึกเวลาล้มเหลว
  try {
    sendTelegram(name, position, status, dateStr, timeStr, lat, lng);
  } catch (err) {
    console.warn('ส่ง Telegram ไม่สำเร็จ: ' + err.message);
  }
  return { success: true, message: 'บันทึกเวลาสำเร็จ' };
}

// ============================================================
//  ดึงข้อมูล Attendance ทั้งหมด → ส่งกลับเป็น JSON array
//  พร้อม rowIndex (เลขแถวจริงใน sheet) เพื่อใช้ลบ
// ============================================================
function getAttendance() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Attendance');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];   // มีแค่ header

  let rows = [];
  for (let i = 1; i < data.length; i++) {
    // คอลัมน์: 0=ชื่อ 1=ตำแหน่ง 2=สถานะ 3=เวลา 4=วันที่ 5=Lat 6=Lng 7=Map
    const dateRaw = String(data[i][4]).replace(/^'/, '');   // ตัด ' ออก
    rows.push({
      rowIndex : i + 1,           // แถวจริงใน sheet (1-based, +1 เพราะ header อยู่แถว 1)
      name     : data[i][0] || '',
      position : data[i][1] || '',
      status   : data[i][2] || '',
      time     : data[i][3] || '',
      date     : dateRaw,
      lat      : data[i][5] || '',
      lng      : data[i][6] || '',
      mapLink  : data[i][7] || ''
    });
  }
  // เรียงใหม่สุดก่อน
  rows.reverse();
  return rows;
}

// ============================================================
//  ลบแถวจาก Attendance sheet ตาม rowIndex (แถวจริงใน sheet)
// ============================================================
function deleteAttendanceRow(rowIndex) {
  if (!rowIndex || rowIndex < 2) return { success: false, error: 'rowIndex ไม่ถูกต้อง' };

  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Attendance');
    if (!sheet) return { success: false, error: 'ไม่พบชีต Attendance' };

    sheet.deleteRow(rowIndex);
    return { success: true, message: 'ลบแถว ' + rowIndex + ' สำเร็จ' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ============================================================
//  ลบข้อมูล Attendance ทั้งหมด (เก็บแถว header ไว้)
// ============================================================
function clearAllAttendance() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Attendance');
    if (!sheet) return { success: false, error: 'ไม่พบชีต Attendance' };

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, message: 'ไม่มีข้อมูลให้ลบ' };

    // ลบแถวข้อมูลทั้งหมด เก็บแค่แถวที่ 1 (header)
    sheet.deleteRows(2, lastRow - 1);

    return { success: true, message: 'ลบข้อมูลทั้งหมด ' + (lastRow - 1) + ' รายการเรียบร้อย' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ============================================================
//  จัดการ Config (GPS)
// ============================================================
function saveConfig(lat, lng, radius) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Config');
  if (!sheet) {
    sheet = ss.insertSheet('Config');
    sheet.getRange('A1:B1').setValues([['Parameter', 'Value']]);
    sheet.getRange('A2').setValue('Target Latitude');
    sheet.getRange('A3').setValue('Target Longitude');
    sheet.getRange('A4').setValue('Allowed Radius (KM)');
    sheet.setColumnWidth(1, 150);
  }
  sheet.getRange('B2').setValue(lat);
  sheet.getRange('B3').setValue(lng);
  sheet.getRange('B4').setValue(radius);
  return { success: true, message: 'บันทึกการตั้งค่าลง Google Sheets เรียบร้อย' };
}

function getConfig() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Config');
  let config  = { lat: 0, lng: 0, radius: 0.5 };
  if (sheet) {
    const latVal    = sheet.getRange('B2').getValue();
    const lngVal    = sheet.getRange('B3').getValue();
    const radiusVal = sheet.getRange('B4').getValue();
    if (latVal    !== '') config.lat    = parseFloat(latVal);
    if (lngVal    !== '') config.lng    = parseFloat(lngVal);
    if (radiusVal !== '') config.radius = parseFloat(radiusVal);
  }
  return config;
}

// ============================================================
//  Telegram notification
// ============================================================
function sendTelegram(name, position, status, dateStr, timeStr, lat, lng) {
  const mapLink = (lat && lng) ? 'https://maps.google.com/?q=' + lat + ',' + lng : '-';
  const emoji   = status === 'ลงเวลาเลิกงาน' ? '🔴' : '🟢';

  const text =
    '📌 แจ้งเตือนการลงเวลา\n\n' +
    '👤 ชื่อ      : ' + name                + '\n' +
    '🏷️ ตำแหน่ง  : ' + (position || '-')   + '\n\n' +
    emoji + ' สถานะ  : ' + (status || '-')  + '\n\n' +
    '📅 วันที่    : ' + dateStr             + '\n' +
    '⏰ เวลา      : ' + timeStr             + '\n\n' +
    '📍 Latitude  : ' + lat                 + '\n' +
    '📍 Longitude : ' + lng                 + '\n\n' +
    '🗺 แผนที่\n'    + mapLink;

  UrlFetchApp.fetch(
    'https://api.telegram.org/bot' + TELEGRAM_TOKEN + '/sendMessage',
    {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ chat_id: TELEGRAM_CHAT_ID, text: text })
    }
  );
}