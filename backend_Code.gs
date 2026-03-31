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
  const { action, formId, data, rowIndex, tableData } = request;
  const isPreview = (action === 'preview');
  
  const { sheet, ss, cleanId } = getTargetSheet(formId);
  const formConfig = CONFIG.FORMS[cleanId];
  
  if (!formConfig) throw new Error('ไม่พบการตั้งค่า Template/Folder สำหรับ: ' + cleanId);

  // --- RESTORED CORE PDF LOGIC ---
  const result = createPDF(formConfig, data, isPreview, tableData);

  if (!isPreview) {
    const rIndex = parseInt(rowIndex) || 0;
    saveToSheet(sheet, data, rIndex, tableData);
  }

  const rowTotal = (tableData && Array.isArray(tableData)) ? tableData.length : 1;
  return createJsonResponse('success', result, `บันทึกข้อมูลเรียบร้อยแล้ว (ได้รับ ${rowTotal} รายการ)`);
}

/**
 * สร้าง PDF จาก Template (รองรับ Google Docs และ Google Slides)
 */
function createPDF(formConfig, rowData, isPreview, tableData) {
  const templateFile = DriveApp.getFileById(formConfig.templateId);
  const parentFolder = DriveApp.getFolderById(formConfig.folderId);
  const mimeType = templateFile.getMimeType();

  const timestamp = Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd_HHmmss");
  const randomId = Math.floor(Math.random() * 1000).toString().padStart(3, '0');
  const subject = rowData['subject'] || rowData['ชื่อเรื่อง / วิชา'] || 'DOC';
  const tempName = (isPreview ? 'PREVIEW_' : '') + '[' + subject + ']_' + timestamp + '_' + randomId;
  
  const tempFile = templateFile.makeCopy(tempName, parentFolder);
  
  // รองรับทั้ง Google Docs และ Google Slides
  if (mimeType === MimeType.GOOGLE_SLIDES) {
    const presentation = SlidesApp.openById(tempFile.getId());
    const slides = presentation.getSlides();

    slides.forEach(slide => {
      slide.getShapes().forEach(shape => replaceInTextSlides(shape.getText(), rowData));
      slide.getTables().forEach(table => {
        for (let r = 0; r < table.getNumRows(); r++) {
          for (let c = 0; c < table.getRow(r).getNumCells(); c++) {
            replaceInTextSlides(table.getCell(r, c).getText(), rowData);
          }
        }
      });
    });
    presentation.saveAndClose();
  } 
  else if (mimeType === MimeType.GOOGLE_DOCS) {
    const doc = DocumentApp.openById(tempFile.getId());
    const body = doc.getBody();
    const header = doc.getHeader();
    const footer = doc.getFooter();

    if (body) {
      if (tableData && Array.isArray(tableData)) {
        fillTableRows(body, tableData);
      }
      replaceInTextDocs(body, rowData);
    }
    if (header) replaceInTextDocs(header, rowData);
    if (footer) replaceInTextDocs(footer, rowData);
    
    doc.saveAndClose();
  } 
  else {
    throw new Error('ประเภทไฟล์เทมเพลตไม่รองรับ ต้องเป็น Google Docs หรือ Google Slides เท่านั้น (พบ: ' + mimeType + ')');
  }

  // พัก 3 วินาทีเพื่อให้เซิร์ฟเวอร์ Google อัพเดทข้อมูลที่เปลี่ยนแปลงลงไฟล์ ก่อนแปลงเป็น PDF
  Utilities.sleep(3000);

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
 * แทนที่ข้อความ (สำหรับ Google Slides)
 */
function replaceInTextSlides(textRange, data) {
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
 * แทนที่ข้อความ (สำหรับ Google Docs)
 */
function replaceInTextDocs(element, data) {
  for (let [key, value] of Object.entries(data)) {
    let valStr = "";
    if (Array.isArray(value)) {
      valStr = value.join(', ');
    } else {
      valStr = (value === null || value === undefined) ? "" : value.toString();
    }

    let keysToTry = [key, key.toLowerCase(), key.toUpperCase(), key.replace('input_', '')];
    
    keysToTry.forEach(k => {
      let findText1 = `\\{\\{${k}\\}\\}`;
      let findText2 = `\\{\\{ ${k} \\}\\}`;
      element.replaceText(findText1, valStr);
      element.replaceText(findText2, valStr);
    });
  }
}

/**
 * [ DYNAMIC TABLE REPLICATION ]
 * ค้นหาตารางที่มี Placeholder ในตัว และทำการ Duplicate แถวนั้นตามจำนวนข้อมูล
 */
function fillTableRows(element, tableData) {
  const tables = element.getTables();
  
  tables.forEach(table => {
    let templateRowIndex = -1;
    let templateRow = null;
    
    // ค้นหาแถวที่มี Placeholder สำหรับ EP หรือ Format (เป็นตัวบ่งบอกเทมเพลต)
    for (let r = 0; r < table.getNumRows(); r++) {
      let rowText = table.getRow(r).getText().toLowerCase();
      if (rowText.includes('{{ep}}') || rowText.includes('{{format}}') || rowText.includes('{{item}}') || rowText.includes('{{ตอน}}') || rowText.includes('{{รูปแบบสื่อ}}')) {
        templateRowIndex = r;
        templateRow = table.getRow(r);
        break;
      }
    }
    
    if (templateRowIndex !== -1) {
      // ลบแถววที่เหลืออยู่ (ที่เป็นจุดไข่ปลาหรือช่องว่าง) เพื่อให้ตารางสะอาด
      // เราจะลบตั้งแต่แถวถัดจาก Template ลงมาจนเกือบสุด
      let totalRows = table.getNumRows();
      for (let r = totalRows - 1; r > templateRowIndex; r--) {
        // ตรวจสอบว่าไม่ใช่แถว Footer ของตาราง (อย่างง่ายๆ คือมีข้อความอื่นที่ไม่ใช่แค่จุดหรือว่าง)
        let text = table.getRow(r).getText().trim();
        if (text.length === 0 || text === "...................." || /^[\s.]+$/.test(text)) {
           table.removeRow(r);
        }
      }

      // เริ่มสร้างแถวใหม่ตามข้อมูล
      tableData.forEach((rowData, index) => {
        let newRow = table.insertTableRow(templateRowIndex + index + 1, templateRow.copy());
        
        // แทนที่ข้อมูลในแถวใหม่นี้
        for (let [key, value] of Object.entries(rowData)) {
          let valStr = (value === null || value === undefined) ? "" : value.toString();
          // ใช้ Regex เพื่อหาและแทนที่แบบไม่สนตัวพิมพ์เล็ก-ใหญ่ และยืดหยุ่นเรื่องช่องว่าง
          let escapedKey = key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); 
          let regex = new RegExp(`\\\\{\\\\{\\\\s*${escapedKey}\\\\s*\\\\}\\\\}`, 'gi');
          newRow.replaceText(regex.source, valStr);
        }
      });
      
      // ลบแถวเทมเพลตต้นฉบับทิ้ง
      table.removeRow(templateRowIndex);
    }
  });
}

/**
 * บันทึกข้อมูล (รองรับการบันทึกหลายแถวรวดเดียวจากตารางทะเบียน)
 */
function saveToSheet(sheet, baseData, rowIndex, tableData) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // กรณีมีข้อมูลตาราง (Batch) ให้บันทึกแยกแถวกัน
  if (tableData && Array.isArray(tableData) && tableData.length > 0) {
    const rowsToAppend = tableData.map(itemData => {
      // รวมข้อมูลหลัก (subject, cb) เข้ากับข้อมูลรายบรรทัด (ep, dur)
      const mergedData = { ...baseData, ...itemData };
      return headers.map(h => {
        let key = h.toString().trim();
        let val = mergedData[key];
        
        // ค้นหาแบบไม่สนตัวพิมพ์ (Case-insensitive mapping)
        if (val === undefined) {
          const matchedKey = Object.keys(mergedData).find(k => k.toLowerCase() === key.toLowerCase());
          if (matchedKey) val = mergedData[matchedKey];
        }
        
        if (Array.isArray(val)) return val.join(', ');
        return (val === undefined || val === null) ? "" : val;
      });
    });

    // บังคับแปลงเป็นตัวเลข และจัดการ rowIndex ให้ถูกต้อง
    const startRow = parseInt(rowIndex) || 0;

    // ถ้าเป็นการบันทึกใหม่ (Append)
    if (startRow <= 0) {
      rowsToAppend.forEach(row => {
        sheet.appendRow(row);
        SpreadsheetApp.flush(); // บังคับเขียนข้อมูลลง Sheet ทันทีป้องกันการข้ามแถว
      });
    } else {
      // ถ้าเป็นการแก้ไข (Edit) - แทนที่แถวเดิม และแทรกที่เหลือ
      sheet.getRange(startRow, 1, 1, headers.length).setValues([rowsToAppend[0]]);
      if (rowsToAppend.length > 1) {
        const remainingRows = rowsToAppend.slice(1);
        sheet.insertRowsAfter(startRow, remainingRows.length);
        sheet.getRange(startRow + 1, 1, remainingRows.length, headers.length).setValues(remainingRows);
      }
      SpreadsheetApp.flush();
    }
  } 
  // กรณีข้อมูลปกติ (ไม่มีตาราง)
  else {
    const row = headers.map(h => {
      let key = h.toString().trim();
      let val = baseData[key];
      
      if (val === undefined) {
        const matchedKey = Object.keys(baseData).find(k => k.toLowerCase() === key.toLowerCase());
        if (matchedKey) val = baseData[matchedKey];
      }
      
      if (Array.isArray(val)) return val.join(', ');
      return (val === undefined || val === null) ? "" : val;
    });

    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
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
