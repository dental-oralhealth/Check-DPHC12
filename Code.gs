const SS_ID = '1IHAL4__mJkVwtUq1GfQrsHHZuhQgNdvhFG6L92U73cM';
const SHEET_NAME = 'การตอบแบบฟอร์ม 1';

/* ===============================
   CONFIG COLUMN
=================================*/

const COL = {
  STATUS: 4,
  PNAME: 5,
  FULLNAME: 6,
  POSITION: 7,
  LEVEL: 8,
  ORGANIZATION: 10,
  PROVINCE: 11,
  STATUS2: 21,
  ROOM_INN: 22,
  CHECK_IN: 24,
  JOIN_MEET: 25,
  FOOD: 26,
  PHONE: 27,        // AA
  CONFIRM_TIME: 29  // AC
};


/* ===============================
   WEB APP ENTRY POINT
   — รองรับทั้ง GitHub Pages (fetch)
     และ Apps Script (google.script.run)
=================================*/

function doGet(e) {

  // ถ้าไม่มี parameter → serve หน้า HTML เดิม (ไม่จำเป็นแล้ว แต่เผื่อไว้)
  if (!e || !e.parameter || !e.parameter.action) {
    return HtmlService
      .createHtmlOutput('<p>API is running.</p>')
      .setTitle('DPHC12 API');
  }

  const action = e.parameter.action;
  const phone  = e.parameter.phone || '';

  let result;

  if (action === 'search') {
    result = search(phone);
  } else if (action === 'confirm') {
    result = confirm(phone);
  } else {
    result = { error: 'Unknown action' };
  }

  // รองรับ JSONP (callback=xxx) เพื่อแก้ CORS จาก GitHub Pages
  const callback = e.parameter.callback;
  const json = JSON.stringify(result);
  const output = callback ? `${callback}(${json})` : json;
  const mime = callback
    ? ContentService.MimeType.JAVASCRIPT
    : ContentService.MimeType.JSON;

  return ContentService
    .createTextOutput(output)
    .setMimeType(mime);
}


/* ===============================
   GET SHEET
=================================*/

function getSheet() {

  return SpreadsheetApp
    .openById(SS_ID)
    .getSheetByName(SHEET_NAME);
}


/* ===============================
   CLEAN PHONE
=================================*/

function cleanPhone(phone) {

  if (!phone) return '';

  return phone
    .toString()
    .trim()
    .replace(/\D/g, '');
}


/* ===============================
   SEARCH
=================================*/

function search(phone) {

  const keyword = cleanPhone(phone);

  if (keyword.length < 4) {
    return [];
  }

  const sheet   = getSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return [];

  const lastCol = sheet.getLastColumn();

  const data = sheet
    .getRange(2, 1, lastRow - 1, lastCol)
    .getDisplayValues();

  return data
    .filter(row => {
      const sheetPhone = cleanPhone(row[COL.PHONE - 1]);
      return sheetPhone.includes(keyword);
    })
    .map((row, index) => ({
      rowNumber:    index + 2,
      status:       row[COL.STATUS       - 1],
      pname:        row[COL.PNAME        - 1],
      fullName:     row[COL.FULLNAME     - 1],
      position:     row[COL.POSITION     - 1],
      level:        row[COL.LEVEL        - 1],
      organization: row[COL.ORGANIZATION - 1],
      province:     row[COL.PROVINCE     - 1],
      status2:      row[COL.STATUS2      - 1],
      room_inn:     row[COL.ROOM_INN     - 1],
      check_in:     row[COL.CHECK_IN     - 1],
      join_meet:    row[COL.JOIN_MEET    - 1],
      food:         row[COL.FOOD         - 1],
      phone:        row[COL.PHONE        - 1],
      confirmTime:  row[COL.CONFIRM_TIME - 1]
    }));
}


/* ===============================
   CONFIRM
=================================*/

function confirm(phone) {

  const cleanInput = cleanPhone(phone);

  if (!cleanInput) {
    return { error: 'กรุณาระบุเบอร์โทร' };
  }

  const sheet   = getSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return { error: 'ไม่มีข้อมูล' };
  }

  const lastCol = sheet.getLastColumn();

  const data = sheet
    .getRange(2, 1, lastRow - 1, lastCol)
    .getDisplayValues();

  for (let i = 0; i < data.length; i++) {

    const row        = data[i];
    const sheetPhone = cleanPhone(row[COL.PHONE - 1]);

    if (sheetPhone === cleanInput) {

      const rowNumber    = i + 2;
      const confirmValue = row[COL.CONFIRM_TIME - 1];

      // ยืนยันแล้ว
      if (confirmValue) {
        return {
          error:       'ผู้ใช้นี้ Confirm แล้ว',
          confirmTime: confirmValue
        };
      }

      const now = new Date();

      sheet
        .getRange(rowNumber, COL.CONFIRM_TIME)
        .setValue(now);

      return {
        success:   true,
        rowNumber: rowNumber,
        fullName:  row[COL.FULLNAME - 1],
        time:      Utilities.formatDate(
                     now,
                     Session.getScriptTimeZone(),
                     'dd/MM/yyyy HH:mm:ss'
                   )
      };
    }
  }

  return { error: 'ไม่พบเบอร์โทร' };
}
