// ==========================================================================
// GOOGLE APPS SCRIPT BACKEND V3.0-TOTAL-POWER (indexed columns FIX)
// ==========================================================================

const CONFIG = {
  SPREADSHEET_ID: '1xAOgwRIoCLmhYYziPQuDTf_YtvDv-VDuCH3WjeIQou4',
  DEBUG_MODE: true,
  FORMS: {
    '031': { templateId: '1zcM4L5gSnNRLF2vgzpiE5E5a7C3g-o8qHG_BDppAm-A', folderId: '1i2szOSdvsEnfYrGbGM-0wALzQ1QPqkbc' },
    '033': { templateId: '1B4o6jYsKBGRKg2dS5qdCDsRHbpJA42GANk55sG_HKlQ', folderId: '1B4o6jYsKBGRKg2dS5qdCDsRHbpJA42GANk55sG_HKlQ' }, 
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
    if (action === 'updateRow') return updateRow(request);
    if (action === 'deleteProject') return deleteProject(request);
    if (action === 'deleteData') return deleteData(request);
    return createJsonResponse('error', null, 'Invalid action');
  } catch (err) { return createJsonResponse('error', null, err.toString()); }
}

function getTargetSheet(formId) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const cleanId = formId.toString().replace('.json', '').replace('-', '').trim();
  let sheet = ss.getSheetByName(cleanId) || ss.getSheetByName(formId);
  return { sheet, ss, cleanId };
}

function getSchema(formId) {
  const { sheet } = getTargetSheet(formId);
  if (!sheet) return createJsonResponse('error', null, 'ไม่พบ Sheet');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return createJsonResponse('success', { version: "v3.0-TotalPower", headers });
}

function handlePDFRequest(request) {
  const { action, formId, data, rowIndex, tableData } = request;
  const isPreview = (action === 'preview');
  const { sheet, cleanId } = getTargetSheet(formId);
  const formConfig = CONFIG.FORMS[cleanId];
  
  if (!formConfig) throw new Error('ไม่พบข้อมูล Template สำหรับ: ' + cleanId);

  const result = createPDF(formConfig, data, isPreview, tableData, cleanId);

  let rowsCount = (tableData && Array.isArray(tableData)) ? tableData.length : 1;
  if (!isPreview) {
    saveToSheet(sheet, data, parseInt(rowIndex) || 0, tableData);
  }

  return createJsonResponse('success', result, `บันทึกข้อมูลเรียบร้อยแล้ว (ได้รับ ${rowsCount} รายการ)`);
}

function createPDF(formConfig, rowData, isPreview, tableData, cleanId) {
  const templateFile = DriveApp.getFileById(formConfig.templateId);
  const parentFolder = DriveApp.getFolderById(formConfig.folderId);
  const timestamp = Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd_HHmmss");
  const subject = rowData['subject'] || rowData['ชื่อเรื่อง / วิชา'] || 'DOC';
  const tempName = (isPreview ? 'PREVIEW_' : '') + '[' + subject + ']_' + timestamp;
  
  const tempFile = templateFile.makeCopy(tempName, parentFolder);
  const doc = DocumentApp.openById(tempFile.getId());
  const body = doc.getBody();

  if (body) {
    const finalData = { ...rowData };
    
    // หากเป็น 034 และมาจากการรวมกลุ่ม (Dash) ให้ดึงข้อมูลกระจายตัว
    if (cleanId === '034' && tableData && tableData.length > 0) {
      tableData.forEach((row, i) => {
        const idx = i + 1;
        finalData['format' + idx] = row.format || row['รูปแบบสื่อ'] || row['format' + idx] || finalData['format' + idx] || "";
        finalData['ep' + idx] = row.ep || row['ตอน'] || row['ep' + idx] || finalData['ep' + idx] || "";
        finalData['teach' + idx] = row.teach || row['วิทยากร/ผู้บรรยาย'] || row['วิทยากร'] || row['teach' + idx] || finalData['teach' + idx] || "";
        finalData['dur' + idx] = row.dur || row['ความยาว'] || row['dur' + idx] || finalData['dur' + idx] || "";
        finalData['dma' + idx] = row.dma || row['ผลิตแล้วเสร็จวันที่'] || row['วันผลิตแล้วเสร็จ'] || row['dma' + idx] || finalData['dma' + idx] || "";
      });
    }

    // --- 034 TOTAL POWER: ปรับปรุงการจัดการหน่วยเวลา (V3.2) ---
    for (let k in finalData) {
      const kl = k.toLowerCase();
      if ((kl.includes('dur') || kl.includes('ความยาว')) && finalData[k]) {
        let val = finalData[k].toString().split('ชม.')[0].trim(); // ล้างของเดิมออกก่อน
        if (val && !isNaN(parseFloat(val))) {
          finalData[k] = val + ' ชม.';
        }
      }
    }

    // แทนที่ Tag สแกนทั้งเอกสาร (Universal Scan)
    replaceInElement(body, finalData);
    
    // สำหรับฟอร์มอื่นๆ (031, 033, 035)
    if (cleanId !== '034' && tableData && Array.isArray(tableData)) {
      fillTableRowsSimple(body, tableData);
    }
    
    // จัดการ Checkbox
    const checkboxKeys = ['cb1', 'cb2', 'cb3', 'cb4', 'cb5', 'cb6'];
    checkboxKeys.forEach(k => {
      const val = finalData[k];
      const isChecked = (val === '✓' || val === true || val === 'true');
      const symbol = isChecked ? "☑" : "☐";
      body.replaceText("\\{\\{" + k + "\\}\\}", symbol);
      body.replaceText("\\{\\{ " + k + " \\}\\}", symbol);
    });

    const header = doc.getHeader();
    if (header) replaceInElement(header, finalData);
  }

  doc.saveAndClose();
  Utilities.sleep(1500); 

  const pdfBlob = tempFile.getAs(MimeType.PDF);
  const pdfFile = parentFolder.createFile(pdfBlob).setName(tempName + ".pdf");
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  tempFile.setTrashed(true);

  return { fileId: pdfFile.getId(), url: pdfFile.getUrl(), name: pdfFile.getName() };
}

function replaceInElement(element, data) {
  if (!element || !data) return;
  const masterMap = {};
  for (let [k, v] of Object.entries(data)) {
    if (k) masterMap[k.toString().toLowerCase().trim()] = v;
  }
  const text = element.getText();
  const tagPattern = /\{\{\s*([^}]+)\s*\}\}/g;
  let match;
  const tagsInDoc = [];
  while ((match = tagPattern.exec(text)) !== null) {
    tagsInDoc.push(match[0]);
  }
  const uniqueTags = [...new Set(tagsInDoc)];
  uniqueTags.forEach(fullTag => {
    const tagName = fullTag.replace(/[\{\}\s]/g, "").toLowerCase();
    if (masterMap.hasOwnProperty(tagName)) {
      const value = masterMap[tagName];
      const finalVal = Array.isArray(value) ? value.join(', ') : (value !== null && value !== undefined ? value.toString() : "");
      const escapedTag = fullTag.replace(/[\{\}\[\]\(\)\*\+\?\.\-\^\$\|]/g, "\\$&");
      try {
        element.replaceText(escapedTag, finalVal);
      } catch (e) {
        try { element.replaceText(fullTag.replace("{","\\{").replace("}","\\}"), finalVal); } catch(err) {}
      }
    }
  });
}

function fillTableRowsSimple(body, tableData) {
  const tables = body.getTables();
  tables.forEach(table => {
    let templateRowIndex = -1;
    for (let r = 0; r < table.getNumRows(); r++) {
      let text = table.getRow(r).getText().toLowerCase();
      if (text.includes('{{ep}}') || text.includes('{{ตอน}}') || text.includes('{{item}}')) {
        templateRowIndex = r; break;
      }
    }
    if (templateRowIndex !== -1) {
      const templateRow = table.getRow(templateRowIndex);
      let lastRowIdx = table.getNumRows() - 1;
      for (let r = lastRowIdx; r > templateRowIndex; r--) {
        let rowText = table.getRow(r).getText().trim();
        if (rowText === "" || rowText.includes("....") || /^[\s.]+$/.test(rowText)) table.removeRow(r);
      }
      tableData.forEach((item, i) => {
        let newRow = table.insertTableRow(templateRowIndex + i + 1, templateRow.copy());
        for (let [key, val] of Object.entries(item)) {
          newRow.replaceText("\\{\\{" + key.toLowerCase() + "\\}\\}", val || "");
          newRow.replaceText("\\{\\{ " + key.toLowerCase() + " \\}\\}", val || "");
        }
      });
      table.removeRow(templateRowIndex);
    }
  });
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
  const startRow = Math.max(2, lastRow - 499);
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

function updateRow(request) {
  const { formId, rowIndex, data } = request;
  const { sheet } = getTargetSheet(formId);
  if (!sheet) return createJsonResponse('error', null, 'ไม่พบ Sheet');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowOutput = headers.map(h => {
    const key = h.toString().trim();
    let val = (data[key] !== undefined) ? data[key] : data[key.toLowerCase()];
    return Array.isArray(val) ? val.join(', ') : (val || "").toString();
  });
  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowOutput]);
  return createJsonResponse('success', null, 'อัปเดตเรียบร้อย');
}

function deleteProject(request) {
  const { formId, subject } = request;
  const { sheet } = getTargetSheet(formId);
  if (!sheet) return createJsonResponse('error', null, 'ไม่พบ Sheet');
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return createJsonResponse('error', null, 'ไม่มีข้อมูล');
  const data = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
  const searchSubject = (subject || "").toString().toLowerCase().trim();
  let count = 0;
  for (let r = lastRow - 1; r >= 1; r--) {
    if ((data[r][3] || "").toString().toLowerCase().trim() === searchSubject) {
      sheet.deleteRow(r + 1); count++;
    }
  }
  return createJsonResponse('success', null, 'ลบแล้ว ' + count + ' แถว');
}

function deleteData(request) {
  const { rowIndex } = request;
  const { sheet } = getTargetSheet('001'); // Fallback
  if (!sheet) return createJsonResponse('error', null, 'ไม่พบ Sheet');
  sheet.deleteRow(rowIndex);
  return createJsonResponse('success', null, 'ลบข้อมูลสำเร็จ');
}
