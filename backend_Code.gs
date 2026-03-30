// ==========================================================================
// GOOGLE APPS SCRIPT BACKEND V2.4-Diagnostic (Official Latest)
// พัฒนาโดย: Antigravity AI
// สำหรับ: ระบบสร้างเอกสาร PDF อัจฉริยะ (FM-สท.03-x)
// ==========================================================================

const CONFIG = {
  SPREADSHEET_ID: '1xAOgwRIoCLmhYYziPQuDTf_YtvDv-VDuCH3WjeIQou4',
  DEBUG_MODE: true
};

/**
 * [ WEB SERVER ]
 * สำหรับแสดงผลหน้าเว็บ HTML หลัก
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Smart PDF Generator | ระบบสร้างเอกสารอัจฉริยะ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * [ ROUTER ]
 * จัดการคำขอ POST จาก Frontend
 */
function doPost(e) {
  try {
    const request = JSON.parse(e.postData.contents);
    const action = request.action;

    if (action === 'getSchema') {
      return getSchema(request.formId);
    } else if (action === 'generate' || action === 'preview') {
      return handlePDFRequest(request);
    } else if (action === 'getRecentData') {
      return getRecentData(request.formId);
    }

    return createJsonResponse('error', null, 'Invalid action');
  } catch (err) {
    return createJsonResponse('error', null, err.toString());
  }
}

/**
 * ดึงโครงสร้างหัวตาราง (Headers) + ระบบ Diagnostic
 */
function getSchema(formId) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  
  // ลองหาด้วยชื่อเต็ม (เช่น 03-1.json) ก่อน ถ้าไม่เจอค่อยหาแบบตัดนามสกุล (03-1)
  let sheet = ss.getSheetByName(formId);
  if (!sheet) {
    sheet = ss.getSheetByName(formId.replace('.json', ''));
  }

  // [DIAGNOSTIC MODE] หากยังไม่เจอ ให้ลิสต์ชื่อ Sheet ทั้งหมดออกมาโชว์
  if (!sheet) {
    const allSheetNames = ss.getSheets().map(s => s.getName()).join(', ');
    return createJsonResponse('error', null, 
      'ไม่พบ Sheet ที่มีชื่อว่า: ' + formId + '\n\n' +
      '🔍 ตรวจพบคิวรีในไฟล์นี้คือ: [' + allSheetNames + ']\n' +
      'กรุณาเปลี่ยนชื่อหน้าใน Google Sheets ให้ตรงกับชื่อใดชื่อหนึ่งด้านบนครับ'
    );
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  return createJsonResponse('success', {
    version: "v2.2-AntiDuplicate",
    headers: headers,
    spreadsheetName: ss.getName()
  });
}

/**
 * ดึงข้อมูลล่าสุด 20 รายการเพื่อนำกลับมาแก้ไข/สั่งพิมพ์ใหม่
 */
function getRecentData(formId) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  
  // ลองหาด้วยชื่อเต็มก่อน
  let sheet = ss.getSheetByName(formId);
  if (!sheet) {
    sheet = ss.getSheetByName(formId.replace('.json', ''));
  }
  
  if (!sheet) return createJsonResponse('error', null, 'ไม่พบ Sheet สำหรับดึงข้อมูลล่าสุด');
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return createJsonResponse('success', { records: [] });
  
  const startRow = Math.max(2, lastRow - 19); // ดึง 20 แถวล่าสุด
  const numRows = lastRow - startRow + 1;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();
  
  const records = data.map((row, index) => {
    const obj = { _rowIndex: startRow + index };
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  }).reverse(); // เอาอันล่าสุดขึ้นก่อน
  
  return createJsonResponse('success', { records: records });
}

/**
 * จัดการการสร้าง PDF (หรือบันทึกข้อมูล)
 */
function handlePDFRequest(request) {
  const { action, formId, data, rowIndex } = request;
  
  // บันทึกข้อมูลลง Google Sheets
  const sheetName = formId.replace('.json', '');
  saveToSheet(sheetName, data, rowIndex);

  // จำลองการสร้างไฟล์ (ในที่นี้คือคืนค่าข้อมูลที่รับมาเพื่อรอระบบสร้าง PDF จริง)
  return createJsonResponse('success', {
    url: 'https://docs.google.com/spreadsheets/d/' + CONFIG.SPREADSHEET_ID,
    name: (data['subject'] || 'เอกสารใหม่') + '_' + (data['actDate'] || '')
  });
}

/**
 * บันทึกข้อมูลลงใน Sheet พร้อมป้องกันชื่อซ้ำ (Update)
 */
function saveToSheet(sheetName, rowData, rowIndex) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  
  // ยืดหยุ่นในการหาชื่อ Sheet ตอนบันทึกด้วย
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.getSheetByName(sheetName + '.json');
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const row = headers.map(h => {
    let val = rowData[h];
    if (val === undefined || val === null) return "";
    if (Array.isArray(val)) return val.join(', ');
    return val;
  });

  if (rowIndex && rowIndex > 1) {
    // โหมดแก้ไข (Update)
    sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  } else {
    // โหมดเพิ่มใหม่ (Append)
    sheet.appendRow(row);
  }
}

/**
 * ส่งคำตอบกลับเป็น JSON
 */
function createJsonResponse(status, data, message = "") {
  return ContentService.createTextOutput(JSON.stringify({
    status: status,
    data: data,
    message: message
  })).setMimeType(ContentService.MimeType.JSON);
}
