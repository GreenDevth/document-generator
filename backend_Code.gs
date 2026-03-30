// ==========================================================================
// GOOGLE APPS SCRIPT BACKEND V2.6-GodMode (Official Latest & PDF Restored)
// พัฒนาโดย: Antigravity AI
// สำหรับ: ระบบสร้างเอกสาร PDF อัจฉริยะ (FM-สท.03-x)
// ==========================================================================

const CONFIG = {
  SPREADSHEET_ID: '1xAOgwRIoCLmhYYziPQuDTf_YtvDv-VDuCH3WjeIQou4',
  DEBUG_MODE: true,
  FORMS: {
    '031': {
      templateId: '1zcM4L5gSnNRLF2vgzpiE5E5a7C3g-o8qHG_BDppAm-A',
      folderId: '1i2szOSdvsEnfYrGbGM-0wALzQ1QPqkbc'
    },
    '033': {
      templateId: '1B4o6jYsKBGRKg2dS5qdCDsRHbpJA42GANk55sG_HKlQ',
      folderId: '1BYkDyzPahByYGadyLaLc5es9zmpjd_2y'
    },
    '034': {
      templateId: '1CriGH1mgj1-MZcBvdNUfC5TYbMytf063U6Tau312ej0',
      folderId: '1WhGML56hu4IKawJyb1CCr-xM1BMTE1Vv'
    },
    '035': {
      templateId: '11HT5_VmKVB7msoJ6j8oVH2NnLm2VGeGKyCt03SQQj3Y',
      folderId: '1mWf-fALH6UPl_COPLtk7AmMhsC-sLEHr'
    }
  }
};

/**
 * [ WEB SERVER ]
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
 * [ DIAGNOSTIC HELPER ]
 */
function getTargetSheet(formId) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const cleanId = formId.replace('.json', '').replace('-', '');
  
  let sheet = ss.getSheetByName(cleanId);
  if (!sheet) sheet = ss.getSheetByName(formId);
  if (!sheet) sheet = ss.getSheetByName(formId.replace('.json', ''));
  
  return { sheet: sheet, ss: ss, cleanId: cleanId };
}

/**
 * ดึงโครงสร้างหัวตาราง (Headers)
 */
function getSchema(formId) {
  const { sheet, ss, cleanId } = getTargetSheet(formId);

  if (!sheet) {
    const allSheetNames = ss.getSheets().map(s => s.getName()).join(', ');
    return createJsonResponse('error', null, 
      'ไม่พบ Sheet: ' + formId + '\n\n' +
      '🔍 ตรวจพบคิวรีในไฟล์นี้คือ: [' + allSheetNames + ']\n' +
      'กรุณาเปลี่ยนชื่อหน้าใน Google Sheets ให้ตรงกับชื่อใดชื่อหนึ่งด้านบนครับ'
    );
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  return createJsonResponse('success', {
    version: "v2.6-GodMode",
    headers: headers,
    spreadsheetName: ss.getName()
  });
}

/**
 * ดึงข้อมูลล่าสุด 20 รายการ
 */
function getRecentData(formId) {
  const { sheet } = getTargetSheet(formId);
  if (!sheet) return createJsonResponse('error', null, 'ไม่พบ Sheet สำหรับดึงข้อมูลล่าสุด');
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return createJsonResponse('success', { records: [] });
  
  const startRow = Math.max(2, lastRow - 19);
  const numRows = lastRow - startRow + 1;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();
  
  const records = data.map((row, index) => {
    const obj = { _rowIndex: startRow + index };
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  }).reverse();
  
  return createJsonResponse('success', { records: records });
}

/**
 * [ CORE ] จัดการการสร้าง PDF จริง
 */
function handlePDFRequest(request) {
  const { action, formId, data, rowIndex } = request;
  const isPreview = (action === 'preview');
  
  const { sheet, ss, cleanId } = getTargetSheet(formId);
  const formConfig = CONFIG.FORMS[cleanId];
  
  if (!formConfig) throw new Error('ไม่พบการตั้งค่า Template/Folder สำหรับ: ' + cleanId);

  // --- RESTORED CORE PDF LOGIC ---
  const result = createPDF(formConfig, data, isPreview);

  if (!isPreview) {
    saveToSheet(sheet, data, rowIndex);
  }

  return createJsonResponse('success', result);
}

/**
 * สร้าง PDF จาก Google Slides Template
 */
function createPDF(formConfig, rowData, isPreview) {
  const templateFile = DriveApp.getFileById(formConfig.templateId);
  const parentFolder = DriveApp.getFolderById(formConfig.folderId);

  const timestamp = Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd_HHmmss");
  const randomId = Math.floor(Math.random() * 1000).toString().padStart(3, '0');
  const subject = rowData['subject'] || rowData['ชื่อเรื่อง / วิชา'] || 'DOC';
  const tempName = (isPreview ? 'PREVIEW_' : '') + '[' + subject + ']_' + timestamp + '_' + randomId;
  
  const tempFile = templateFile.makeCopy(tempName, parentFolder);
  const presentation = SlidesApp.openById(tempFile.getId());
  const slides = presentation.getSlides();

  slides.forEach(slide => {
    slide.getShapes().forEach(shape => replaceInText(shape.getText(), rowData));
    slide.getTables().forEach(table => {
      for (let r = 0; r < table.getNumRows(); r++) {
        for (let c = 0; c < table.getRow(r).getNumCells(); c++) {
          replaceInText(table.getCell(r, c).getText(), rowData);
        }
      }
    });
  });

  presentation.saveAndClose();

  const pdfBlob = tempFile.getAs(MimeType.PDF);
  const pdfFile = parentFolder.createFile(pdfBlob).setName(tempName + ".pdf");
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  tempFile.setTrashed(true);

  return {
    fileId: pdfFile.getId(),
    url: pdfFile.getUrl(), // ใช้ URL ปกติเพื่อให้เปิดหน้า Preview สำหรับพิมพ์
    name: pdfFile.getName()
  };
}

/**
 * แทนที่ข้อความ
 */
function replaceInText(textRange, data) {
  for (let [key, value] of Object.entries(data)) {
    let valStr = "";
    if (Array.isArray(value)) {
      valStr = value.join(', ');
    } else {
      valStr = (value === null || value === undefined) ? "" : value.toString();
    }

    let keysToTry = [key, key.toLowerCase(), key.toUpperCase(), key.replace('input_', '')];
    
    keysToTry.forEach(k => {
      textRange.replaceAllText(`{{${k}}}`, valStr, false);
      textRange.replaceAllText(`{{ ${k} }}`, valStr, false);
    });
  }
}

/**
 * บันทึกข้อมูล
 */
function saveToSheet(sheet, rowData, rowIndex) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => {
    let val = rowData[h.toString().trim()];
    if (val === undefined || val === null) return "";
    const valStr = val.toString();
    if (/[0-9๐-๙]/.test(valStr)) return "'" + valStr;
    return valStr;
  });

  if (rowIndex && rowIndex > 1) {
    sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  } else {
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
    message: message || (status === 'error' ? 'เกิดข้อผิดพลาด' : '')
  })).setMimeType(ContentService.MimeType.JSON);
}
