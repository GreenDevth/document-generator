// ==========================================================================
// GOOGLE APPS SCRIPT BACKEND V2.7-GodMode (Total Repair)
// พัฒนาโดย: Antigravity AI
// สำหรับ: ระบบสร้างเอกสาร PDF อัจฉริยะ (FM-สท.03-x)
// ==========================================================================

const CONFIG = {
  SPREADSHEET_ID: '1xAOgwRIoCLmhYYziPQuDTf_YtvDv-VDuCH3WjeIQou4',
  DEBUG_MODE: true,
  FORMS: {
    '031': { templateId: '1zcM4L5gSnNRLF2vgzpiE5E5a7C3g-o8qHG_BDppAm-A', folderId: '1i2szOSdvsEnfYrGbGM-0wALzQ1QPqkbc' },
    '033': { templateId: '1B4o6jYsKBGRKg2dS5qdCDsRHbpJA42GANk55sG_HKlQ', folderId: '1BYkDyzPahByYGadyLaLc5es9zmpjd_2y' },
    '034': { templateId: '1CriGH1mgj1-MZcBvdNUfC5TYbMytf063U6Tau312ej0', folderId: '1WhGML56hu4IKawJyb1CCr-xM1BMTE1Vv' },
    '035': { templateId: '11HT5_VmKVB7msoJ6j8oVH2NnLm2VGeGKyCt03SQQj3Y', folderId: '1mWf-fALH6UPl_COPLtk7AmMhsC-sLEHr' }
  }
};

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Smart PDF Generator | ระบบสร้างเอกสารอัจฉริยะ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const request = JSON.parse(e.postData.contents);
    const action = request.action;
    if (action === 'getSchema') return getSchema(request.formId);
    if (action === 'generate' || action === 'preview') return handlePDFRequest(request);
    if (action === 'getRecentData') return getRecentData(request.formId);
    return createJsonResponse('error', null, 'Invalid action');
  } catch (err) {
    return createJsonResponse('error', null, err.toString());
  }
}

function getTargetSheet(formId) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const cleanId = formId.replace('.json', '').replace('-', '');
  let sheet = ss.getSheetByName(cleanId) || ss.getSheetByName(formId) || ss.getSheetByName(formId.replace('.json', ''));
  return { sheet: sheet, ss: ss, cleanId: cleanId };
}

function getSchema(formId) {
  const { sheet, ss } = getTargetSheet(formId);
  if (!sheet) return createJsonResponse('error', null, 'ไม่พบ Sheet: ' + formId);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return createJsonResponse('success', { version: "v2.7-GodMode", headers: headers });
}

function getRecentData(formId) {
  const { sheet } = getTargetSheet(formId);
  if (!sheet) return createJsonResponse('error', null, 'ไม่พบ Sheet');
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

function handlePDFRequest(request) {
  const { action, formId, data, rowIndex, tableData } = request;
  const isPreview = (action === 'preview');
  const { sheet, ss, cleanId } = getTargetSheet(formId);
  const formConfig = CONFIG.FORMS[cleanId];
  
  if (!formConfig) throw new Error('ไม่พบข้อมูล Template สำหรับ: ' + cleanId);

  // 1. สร้าง PDF
  const result = createPDF(formConfig, data, isPreview, tableData);

  // 2. บันทึกข้อมูลลง Sheet
  let rowsCount = 1;
  if (!isPreview) {
    const rIndex = parseInt(rowIndex) || 0;
    saveToSheet(sheet, data, rIndex, tableData);
    rowsCount = (tableData && Array.isArray(tableData)) ? tableData.length : 1;
  }

  return createJsonResponse('success', result, `บันทึกข้อมูลเรียบร้อยแล้ว (ได้รับ ${rowsCount} รายการ)`);
}

function createPDF(formConfig, rowData, isPreview, tableData) {
  const templateFile = DriveApp.getFileById(formConfig.templateId);
  const parentFolder = DriveApp.getFolderById(formConfig.folderId);
  const timestamp = Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd_HHmmss");
  const subject = rowData['subject'] || rowData['ชื่อเรื่อง / วิชา'] || 'DOC';
  const tempName = (isPreview ? 'PREVIEW_' : '') + '[' + subject + ']_' + timestamp;
  
  const tempFile = templateFile.makeCopy(tempName, parentFolder);
  const mimeType = templateFile.getMimeType();

  if (mimeType === MimeType.GOOGLE_DOCS) {
    const doc = DocumentApp.openById(tempFile.getId());
    const body = doc.getBody();
    if (body) {
      if (tableData && Array.isArray(tableData)) fillTableRows(body, tableData);
      replaceInText(body, rowData);
    }
    const header = doc.getHeader();
    if (header) replaceInText(header, rowData);
    doc.saveAndClose();
  } else if (mimeType === MimeType.GOOGLE_SLIDES) {
    const pres = SlidesApp.openById(tempFile.getId());
    pres.getSlides().forEach(slide => replaceInSlides(slide, rowData));
    pres.saveAndClose();
  }

  Utilities.sleep(2000);
  const pdfBlob = tempFile.getAs(MimeType.PDF);
  const pdfFile = parentFolder.createFile(pdfBlob).setName(tempName + ".pdf");
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  tempFile.setTrashed(true);

  return { fileId: pdfFile.getId(), url: pdfFile.getUrl(), name: pdfFile.getName() };
}

function replaceInText(element, data) {
  for (let [key, value] of Object.entries(data)) {
    const val = Array.isArray(value) ? value.join(', ') : (value || "");
    const k = key.toLowerCase();
    // ใช้การแทนที่แบบแม่นยำ (Double backslash สำหรับ Apps Script Regex)
    element.replaceText("\\{\\{" + k + "\\}\\}", val);
    element.replaceText("\\{\\{ " + k + " \\}\\}", val);
  }
}

function replaceInSlides(slide, data) {
  for (let [key, value] of Object.entries(data)) {
    const val = Array.isArray(value) ? value.join(', ') : (value || "");
    const k = key.toLowerCase();
    slide.replaceAllText("{{"+k+"}}", val);
    slide.replaceAllText("{{ "+k+" }}", val);
  }
}

function fillTableRows(body, tableData) {
  const tables = body.getTables();
  tables.forEach(table => {
    let templateRowIndex = -1;
    for (let r = 0; r < table.getNumRows(); r++) {
      let text = table.getRow(r).getText().toLowerCase();
      // ดิจิทัล - ค้นหาแถวที่มี Placeholder
      if (text.includes('{{ep}}') || text.includes('{{format}}') || text.includes('{{item}}')) {
        templateRowIndex = r;
        break;
      }
    }

    if (templateRowIndex !== -1) {
      const templateRow = table.getRow(templateRowIndex);
      tableData.forEach((rowData, i) => {
        const newRow = table.insertTableRow(templateRowIndex + i + 1, templateRow.copy());
        for (let [key, value] of Object.entries(rowData)) {
          const val = value || "";
          const k = key.toLowerCase();
          newRow.replaceText("\\{\\{" + k + "\\}\\}", val);
          newRow.replaceText("\\{\\{ " + k + " \\}\\}", val);
        }
      });
      table.removeRow(templateRowIndex);
    }
  });
}

function saveToSheet(sheet, baseData, rowIndex, tableData) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // กรณี Batch
  if (tableData && Array.isArray(tableData) && tableData.length > 0) {
    const rowsToAppend = tableData.map(item => {
      const merged = { ...baseData, ...item };
      return headers.map(h => {
        let val = merged[h.toString().trim()] || merged[h.toString().trim().toLowerCase()];
        if (val === undefined) {
           const match = Object.keys(merged).find(k => k.toLowerCase() === h.toString().toLowerCase().trim());
           val = match ? merged[match] : "";
        }
        return Array.isArray(val) ? val.join(', ') : (val || "");
      });
    });

    if (rowIndex <= 0) {
      rowsToAppend.forEach(r => { sheet.appendRow(r); SpreadsheetApp.flush(); });
    } else {
      sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowsToAppend[0]]);
      if (rowsToAppend.length > 1) {
        const extra = rowsToAppend.slice(1);
        sheet.insertRowsAfter(rowIndex, extra.length);
        sheet.getRange(rowIndex + 1, 1, extra.length, headers.length).setValues(extra);
      }
      SpreadsheetApp.flush();
    }
  } else {
    // กรณีปกติ
    const row = headers.map(h => {
      let val = baseData[h] || baseData[h.toString().toLowerCase().trim()];
      return Array.isArray(val) ? val.join(', ') : (val || "");
    });
    if (rowIndex > 0) sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
    else sheet.appendRow(row);
  }
}

function createJsonResponse(status, data, message = "") {
  return ContentService.createTextOutput(JSON.stringify({ status, data, message }))
    .setMimeType(ContentService.MimeType.JSON);
}
