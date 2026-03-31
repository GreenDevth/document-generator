// ==========================================================================
// GOOGLE APPS SCRIPT BACKEND V2.9.2-GODMODE (PDF TABLE REPAIR)
// ==========================================================================

const CONFIG = {
  SPREADSHEET_ID: '1xAOgwRIoCLmhYYziPQuDTf_YtvDv-VDuCH3WjeIQou4',
  DEBUG_MODE: true,
  FORMS: {
    '031': { templateId: '1zcM4L5gSnNRLF2vgzpiE5E5a7C3g-o8qHG_BDppAm-A', folderId: '1i2szOSdvsEnfYrGbGM-0wALzQ1QPqkbc' },
    '033': { templateId: '1B4o6jYsKBGRKg2dS5qdCDsRHbpJA42GANk55sG_HKlQ', folderId: '1B4o6jYsKBGRKg2dS5qdCDsRHbpJA42GANk55sG_HKlQ' }, // Fix Folder ID if needed
    '034': { templateId: '1CriGH1mgj1-MZcBvdNUfC5TYbMytf063U6Tau312ej0', folderId: '1WhGML56hu4IKawJyb1CCr-xM1BMTE1Vv' },
    '035': { templateId: '11HT5_VmKVB7msoJ6j8oVH2NnLm2VGeGKyCt03SQQj3Y', folderId: '1mWf-fALH6UPl_COPLtk7AmMhsC-sLEHr' }
  }
};

function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
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
  } catch (err) { return createJsonResponse('error', null, err.toString()); }
}

function getTargetSheet(formId) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const cleanId = formId.replace('.json', '').replace('-', '');
  let sheet = ss.getSheetByName(cleanId) || ss.getSheetByName(formId);
  return { sheet, ss, cleanId };
}

function getSchema(formId) {
  const { sheet } = getTargetSheet(formId);
  if (!sheet) return createJsonResponse('error', null, 'ไม่พบ Sheet');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return createJsonResponse('success', { version: "v2.9.2-GodMode", headers });
}

function handlePDFRequest(request) {
  const { action, formId, data, rowIndex, tableData } = request;
  const isPreview = (action === 'preview');
  const { sheet, cleanId } = getTargetSheet(formId);
  const formConfig = CONFIG.FORMS[cleanId];
  
  if (!formConfig) throw new Error('ไม่พบข้อมูล Template สำหรับ: ' + cleanId);

  // 1. สร้าง PDF
  const result = createPDF(formConfig, data, isPreview, tableData);

  // 2. บันทึกข้อมูลลง Sheet (ถ้าไม่ใช่ Preview)
  let rowsCount = (tableData && Array.isArray(tableData)) ? tableData.length : 1;
  if (!isPreview) {
    saveToSheet(sheet, data, parseInt(rowIndex) || 0, tableData);
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
  const doc = DocumentApp.openById(tempFile.getId());
  const body = doc.getBody();

  if (body) {
    // ลำดับสำคัญ: 1. เคลียร์และสร้างแถวในตารางก่อน
    if (tableData && Array.isArray(tableData)) {
      fillTableRowsSimple(body, tableData);
    }
    // 2. เก็บรายละเอียดอื่นๆ ทั่วทั่งเอกสาร (หัวกระดาษ/ท้ายกระดาษ)
    replaceInElement(body, rowData);
    const header = doc.getHeader();
    if (header) replaceInElement(header, rowData);
  }

  doc.saveAndClose();
  Utilities.sleep(2500); // พักรอกระบวนการหลังบ้าน Google

  const pdfBlob = tempFile.getAs(MimeType.PDF);
  const pdfFile = parentFolder.createFile(pdfBlob).setName(tempName + ".pdf");
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  tempFile.setTrashed(true);

  return { fileId: pdfFile.getId(), url: pdfFile.getUrl(), name: pdfFile.getName() };
}

function fillTableRowsSimple(body, tableData) {
  const tables = body.getTables();
  tables.forEach(table => {
    let templateRowIndex = -1;
    
    // ค้นหาแถวที่มี Placeholder
    for (let r = 0; r < table.getNumRows(); r++) {
      let text = table.getRow(r).getText().toLowerCase();
      if (text.includes('{{ep}}') || text.includes('{{ตอน}}') || text.includes('{{item}}')) {
        templateRowIndex = r;
        break;
      }
    }

    if (templateRowIndex !== -1) {
      const templateRow = table.getRow(templateRowIndex);
      
      // ขั้นที่ 1: ลบแถวขยะ (จุดไข่ปลา) ที่อยู่ถัดจากเทมเพลตลงมาให้หมด
      let lastRowIdx = table.getNumRows() - 1;
      for (let r = lastRowIdx; r > templateRowIndex; r--) {
        let rowText = table.getRow(r).getText().trim();
        // ถ้าเป็นแถวที่มีแค่จุดหรือว่าง ให้ลบออก
        if (rowText === "" || rowText.includes("....") || /^[\s.]+$/.test(rowText)) {
          table.removeRow(r);
        }
      }

      // ขั้นที่ 2: วนลูปสร้างแถวใหม่ตามข้อมูลที่มี
      tableData.forEach((item, i) => {
        let newRow = table.insertTableRow(templateRowIndex + i + 1, templateRow.copy());
        // แทนที่ข้อมูลรายบรรทัด (EP, รูปแบบสื่อ, ฯลฯ)
        for (let [key, val] of Object.entries(item)) {
          const k = key.toLowerCase();
          const v = val || "";
          newRow.replaceText("\\{\\{" + k + "\\}\\}", v);
          newRow.replaceText("\\{\\{ " + k + " \\}\\}", v);
        }
      });

      // ขั้นที่ 3: ลบแถวเทมเพลตต้นฉบับทิ้ง
      table.removeRow(templateRowIndex);
    }
  });
}

function replaceInElement(element, data) {
  for (let [key, value] of Object.entries(data)) {
    const val = Array.isArray(value) ? value.join(', ') : (value || "");
    const k = key.toLowerCase();
    element.replaceText("\\{\\{" + k + "\\}\\}", val);
    element.replaceText("\\{\\{ " + k + " \\}\\}", val);
  }
}

function saveToSheet(sheet, baseData, rowIndex, tableData) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (tableData && Array.isArray(tableData) && tableData.length > 0) {
    const rows = tableData.map(item => {
      const merged = { ...baseData, ...item };
      return headers.map(h => {
        const key = h.toString().trim();
        let val = merged[key] || merged[key.toLowerCase()];
        if (val === undefined) {
          const match = Object.keys(merged).find(k => k.toLowerCase() === key.toLowerCase());
          val = match ? merged[match] : "";
        }
        return Array.isArray(val) ? val.join(', ') : (val || "");
      });
    });

    if (rowIndex <= 0) {
      rows.forEach(r => { sheet.appendRow(r); SpreadsheetApp.flush(); });
    } else {
      // โหมดแก้ไข (รองรับแค่แถวเดียวสำหรับชุดเบื้องต้น)
      sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rows[0]]);
    }
  }
}

function createJsonResponse(status, data, message = "") {
  return ContentService.createTextOutput(JSON.stringify({ status, data, message }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getRecentData(formId) {
  const { sheet } = getTargetSheet(formId);
  if (!sheet) return createJsonResponse('error', null, 'ไม่พบ Sheet');
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return createJsonResponse('success', { records: [] });
  const startRow = Math.max(2, lastRow - 19);
  const numRows = lastRow - startRow + 1;
  const hs = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const d = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();
  const records = d.map((r, idx) => {
    const obj = { _rowIndex: startRow + idx };
    hs.forEach((h, i) => obj[h] = r[i]);
    return obj;
  }).reverse();
  return createJsonResponse('success', { records });
}
