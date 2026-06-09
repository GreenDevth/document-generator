// ==========================================================================
// GOOGLE APPS SCRIPT BACKEND V3.5-SLIDES-EDITION
// รองรับ Google Slides และ Smart Mapping ภาษาไทย
// ==========================================================================

const DEFAULT_CONFIG = {
  SPREADSHEET_ID: '1xAOgwRIoCLmhYYziPQuDTf_YtvDv-VDuCH3WjeIQou4',
  DEBUG_MODE: true,
  FORMS: {
    '031': { templateId: '1zcM4L5gSnNRLF2vgzpiE5E5a7C3g-o8qHG_BDppAm-A', folderId: '1i2szOSdvsEnfYrGbGM-0wALzQ1QPqkbc' },
    '033': { templateId: '1B4o6jYsKBGRKg2dS5qdCDsRHbpJA42GANk55sG_HKlQ', folderId: '1BYkDyzPahByYGadyLaLc5es9zmpjd_2y' }, 
    '034': { templateId: '1CriGH1mgj1-MZcBvdNUfC5TYbMytf063U6Tau312ej0', folderId: '1WhGML56hu4IKawJyb1CCr-xM1BMTE1Vv' },
    '035': { templateId: '11HT5_VmKVB7msoJ6j8oVH2NnLm2VGeGKyCt03SQQj3Y', folderId: '1mWf-fALH6UPl_COPLtk7AmMhsC-sLEHr' }
  }
};

// ดึงค่าคอนฟิกส่วนบุคคลจากระบบ Script Properties ของแต่ละคน
function getConfig() {
  const props = PropertiesService.getScriptProperties();
  const userConfig = props.getProperty('USER_CONFIG');
  if (userConfig) {
    try {
      return JSON.parse(userConfig);
    } catch (e) {
      // ดำเนินการใช้ค่าเริ่มต้นเมื่อเกิดข้อผิดพลาดในการแปลงข้อมูล
    }
  }
  return DEFAULT_CONFIG;
}

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
    const spreadsheetId = request.spreadsheetId; // ดึงไอดีชีตที่ส่งมาจากหน้าบ้าน
    
    if (action === 'verifyPassword') return verifyPassword(request.password);
    if (action === 'setupWorkspace') return setupWorkspace(request.password, request.targetFolderId);
    if (action === 'getSheets') return getSheets(spreadsheetId);
    if (action === 'getSchema') return getSchema(request.formId, spreadsheetId);
    if (action === 'generate' || action === 'preview') return handlePDFRequest(request);
    if (action === 'getRecentData') return getRecentData(request.formId, spreadsheetId);
    if (action === 'updateRow') return updateRow(request);
    if (action === 'deleteProject') return deleteProject(request);
    if (action === 'deleteData') return deleteData(request);
    if (action === 'getDriveFiles') return getDriveFiles(request.formId, spreadsheetId, request.folderId); // รองรับโฟลเดอร์ย่อยย่อย
    if (action === 'createSubFolder') return createSubFolder(request.parentFolderId, request.folderName); // สร้างโฟลเดอร์ย่อยใหม่
    if (action === 'moveFiles') return moveFiles(request.fileIds, request.targetFolderId); // ย้ายไฟล์ PDF ไปเก็บ
    if (action === 'getFileBytes') return getFileBytes(request.fileId);
    if (action === 'getConfig') return getBackendConfig(); // เพิ่มดึงค่าคอนฟิก
    if (action === 'saveConfig') return saveBackendConfig(request.configData); // เพิ่มบันทึกค่าคอนฟิก
    if (action === 'toggleAutoTrigger') return toggleAutoTrigger(request.enable, spreadsheetId); // เปิด-ปิดทริกเกอร์ออโต้
    if (action === 'getAutoTriggerStatus') return getAutoTriggerStatus(); // เช็คสถานะทริกเกอร์ออโต้
    return createJsonResponse('error', null, 'Invalid action');
  } catch (err) { return createJsonResponse('error', null, err.toString()); }
}

function getTargetSheet(formId, spreadsheetId) {
  const config = getConfig();
  const ssId = spreadsheetId || config.SPREADSHEET_ID;
  const ss = SpreadsheetApp.openById(ssId);
  const cleanId = formId.toString().replace('.json', '').replace('-', '').trim();
  let sheet = ss.getSheetByName(cleanId) || ss.getSheetByName(formId);
  return { sheet, ss, cleanId };
}

function getSchema(formId, spreadsheetId) {
  const { sheet } = getTargetSheet(formId, spreadsheetId);
  if (!sheet) return createJsonResponse('error', null, 'ไม่พบ Sheet');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return createJsonResponse('success', { version: "v3.5-Slides-Power", headers });
}

function handlePDFRequest(request) {
  const { action, formId, data, rowIndex, tableData, skipSave, spreadsheetId } = request;
  const isPreview = (action === 'preview') || (skipSave === true);
  const { sheet } = getTargetSheet(formId, spreadsheetId);
  if (!sheet) throw new Error('ไม่พบข้อมูล Sheet: ' + formId);
  
  const sheetName = sheet.getName();
  const formConfig = getFormConfig(sheetName);
  
  if (!formConfig) throw new Error('ไม่พบข้อมูล Template สำหรับแผ่นงาน: ' + sheetName);
  const cleanId = formConfig.baseId;

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
  const presentation = SlidesApp.openById(tempFile.getId());
  
  if (presentation) {
    const finalData = { ...rowData };
    
    // --- SMART MAPPING & TABLE MERGE ---
    const smartMap = {
      'ตอน': 'ep', 'ep': 'ตอน',
      'ความยาว': 'duration', 'dur': 'duration', 'duration': 'dur',
      'ความยาวรายการ': 'duration', 'ความยาวรายการ (นาที)': 'duration',
      'ชื่อรายการ': 'subject', 'subject': 'ชื่อรายการ', 'ชื่อเรื่อง': 'subject'
    };

    // 1. ดึงข้อมูลจากแถวแรกของตาราง (ถ้ามี) มาใส่ในข้อมูลหลัก
    if (tableData && Array.isArray(tableData) && tableData.length > 0) {
      const firstRow = tableData[0];
      for (let [tk, tv] of Object.entries(firstRow)) {
        if (finalData[tk] === undefined || finalData[tk] === "") {
          finalData[tk] = tv;
        }
      }
    }

    // 2. ทำการ Mapping ชื่อฟิลด์ตาม smartMap
    for (let [thai, eng] of Object.entries(smartMap)) {
      if (finalData[thai] !== undefined && finalData[eng] === undefined) {
        finalData[eng] = finalData[thai];
      }
    }

    // หากเป็น 034 (แบบกลุ่ม)
    if (cleanId === '034') {
      const firstRow = (tableData && tableData.length > 0) ? tableData[0] : {};
      
      for (let idx = 1; idx <= 11; idx++) {
        const row = (tableData && tableData[idx-1]) ? tableData[idx-1] : {};
        
        finalData['format' + idx] = row.format || row['รูปแบบสื่อ'] || firstRow['format' + idx] || firstRow['รูปแบบสื่อ' + idx] || "";
        finalData['ep' + idx] = row.ep || row['ตอน'] || firstRow['ep' + idx] || firstRow['ตอน' + idx] || "";
        finalData['teach' + idx] = row.teach || row['วิทยากร'] || firstRow['teach' + idx] || firstRow['วิทยากร' + idx] || "";
        finalData['dur' + idx] = row.dur || row['ความยาว'] || firstRow['dur' + idx] || firstRow['ความยาว' + idx] || "";
        
        let dmaVal = row.dma || row['ผลิตแล้วเสร็จวันที่'] || firstRow['dma' + idx] || firstRow['ผลิตแล้วเสร็จวันที่' + idx] || "";
        if (isDate(dmaVal)) {
          const dmaStr = Utilities.formatDate(dmaVal, "GMT+7", "dd/MM/yyyy");
          const p = dmaStr.split('/');
          const year = parseInt(p[2]);
          dmaVal = `${p[0]}/${p[1]}/${year < 2400 ? year + 543 : year}`;
        }
        finalData['dma' + idx] = dmaVal;
      }
    }

    // จัดการ Checkbox
    const checkboxKeys = ['cb1', 'cb2', 'cb3', 'cb4', 'cb5', 'cb6'];
    checkboxKeys.forEach(k => {
      const val = finalData[k];
      const isChecked = (val === '✓' || val === true || val === 'true');
      const symbol = isChecked ? "✓" : "";
      presentation.replaceAllText("{{" + k + "}}", symbol);
      presentation.replaceAllText("{{ " + k + " }}", symbol);
      presentation.replaceAllText("{{" + k.toUpperCase() + "}}", symbol);
    });

    // แทนที่ Tag ทั่วไป
    const sortedKeys = Object.keys(finalData).sort((a, b) => b.length - a.length);
    
    for (let k of sortedKeys) {
      const v = finalData[k];
      if (k && v !== undefined) {
        let valStr = "";
        
        if (isDate(v)) {
          valStr = (v.getFullYear() < 1905) ? Utilities.formatDate(v, "GMT+7", "H:mm") : Utilities.formatDate(v, "GMT+7", "dd/MM/yyyy");
        } else if (typeof v === 'string' && v.includes('T') && v.includes('Z') && v.length > 15) {
          try {
            const d = new Date(v);
            valStr = (d.getFullYear() < 1905) ? Utilities.formatDate(d, "GMT+7", "H:mm") : Utilities.formatDate(d, "GMT+7", "dd/MM/yyyy");
          } catch(e) { valStr = v.toString(); }
        } else {
          valStr = (v === null) ? "" : v.toString();
        }

        presentation.replaceAllText("{{" + k + "}}", valStr);
        presentation.replaceAllText("{{ " + k + " }}", valStr);
        presentation.replaceAllText("{{" + k.toLowerCase() + "}}", valStr);
        presentation.replaceAllText("{{ " + k.toLowerCase() + " }}", valStr);
        presentation.replaceAllText("{{" + k.toUpperCase() + "}}", valStr);
        presentation.replaceAllText("{{ " + k.toUpperCase() + " }}", valStr);
      }
    }
    
    // เติมตารางในสไลด์
    if (cleanId !== '034' && tableData && Array.isArray(tableData)) {
      fillSlidesTable(presentation, tableData);
    }
  }

  presentation.saveAndClose();
  Utilities.sleep(1500); 

  const pdfBlob = tempFile.getAs(MimeType.PDF);
  const pdfFile = parentFolder.createFile(pdfBlob).setName(tempName + ".pdf");
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  tempFile.setTrashed(true);

  return { fileId: pdfFile.getId(), url: pdfFile.getUrl(), name: pdfFile.getName() };
}

function fillSlidesTable(presentation, tableData) {
  const slides = presentation.getSlides();
  slides.forEach(slide => {
    const tables = slide.getTables();
    tables.forEach(table => {
      let templateRowIndex = -1;
      for (let r = 0; r < table.getNumRows(); r++) {
        for (let c = 0; c < table.getRow(r).getNumCells(); c++) {
          const text = table.getCell(r, c).getText().asString();
          if (text.includes('{{')) {
            templateRowIndex = r;
            break;
          }
        }
        if (templateRowIndex !== -1) break;
      }

      if (templateRowIndex !== -1) {
        tableData.forEach(rowData => {
          const newRow = table.appendRow();
          for (let c = 0; c < table.getRow(templateRowIndex).getNumCells(); c++) {
            const templateCell = table.getCell(templateRowIndex, c);
            const templateText = templateCell.getText().asString();
            let cellValue = templateText;
            
            for (let [k, v] of Object.entries(rowData)) {
              const valStr = (v === null || v === undefined) ? "" : v.toString();
              cellValue = cellValue.replace(new RegExp('{{\\s*' + k + '\\s*}}', 'gi'), valStr);
            }
            cellValue = cellValue.replace(/{{\s*[^}]+\s*}}/g, '');
            newRow.getCell(c).getText().setText(cellValue);
          }
        });
        table.removeRow(templateRowIndex);
      }
    });
  });
}

function saveToSheet(sheet, baseData, rowIndex, tableData) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  let rows = [];
  if (tableData && Array.isArray(tableData) && tableData.length > 0) {
    rows = tableData.map(item => {
      const merged = { ...baseData, ...item };
      return mapRowToHeaders(headers, merged);
    });
  } else {
    rows = [mapRowToHeaders(headers, baseData)];
  }

  if (rowIndex <= 0) {
    rows.forEach(r => { sheet.appendRow(r); SpreadsheetApp.flush(); });
  } else {
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rows[0]]);
  }
}

function mapRowToHeaders(headers, rowData) {
  return headers.map(h => {
    const key = h.toString().trim();
    let val = rowData[key] || rowData[key.toLowerCase()];
    if (val === undefined) {
      const match = Object.keys(rowData).find(k => k.toLowerCase() === key.toLowerCase());
      val = match ? rowData[match] : "";
    }
    
    if (typeof val === 'string' && val.includes('T') && val.includes('Z') && val.length > 15) {
      try {
        const d = new Date(val);
        if (!isNaN(d.getTime())) {
          val = (d.getFullYear() < 1905) ? Utilities.formatDate(d, "GMT+7", "H:mm") : Utilities.formatDate(d, "GMT+7", "dd/MM/yyyy");
        }
      } catch(e) {}
    }

    return Array.isArray(val) ? val.join(', ') : (val || "");
  });
}

function createJsonResponse(status, data, message = "") {
  return ContentService.createTextOutput(JSON.stringify({ status, data, message }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getRecentData(formId, spreadsheetId) {
  const { sheet } = getTargetSheet(formId, spreadsheetId);
  if (!sheet) return createJsonResponse('error', null, 'ไม่พบ Sheet');
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return createJsonResponse('success', { records: [], headers: [] });
  const startRow = Math.max(2, lastRow - 499);
  const numRows = lastRow - startRow + 1;
  const allHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const v = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();
  const d = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getDisplayValues();
  
  const records = d.map((r, idx) => {
    const obj = { _rowIndex: startRow + idx };
    allHeaders.forEach((h, i) => {
      const cleanHeader = h ? h.toString().trim() : "";
      if (!cleanHeader) return;
      
      let val = r[i];
      let raw = v[idx][i];

      if (isDate(raw)) {
        if (raw.getFullYear() < 1905) {
          const hh = raw.getHours();
          const mm = raw.getMinutes();
          const ss = raw.getSeconds();
          val = (hh > 0 ? hh + ":" : "") + (mm < 10 && hh > 0 ? "0" + mm : mm) + ":" + (ss < 10 ? "0" + ss : ss);
        } else {
          val = Utilities.formatDate(raw, "GMT+7", "dd/MM/yyyy");
          const p = val.split('/');
          const year = parseInt(p[2]);
          val = `${p[0]}/${p[1]}/${year < 2400 ? year + 543 : year}`;
        }
      }
      
      if (val === "0:00" || val === "00:00" || val === "0:00:00") val = "-";
      
      obj[cleanHeader] = val;
    });
    return obj;
  });
  const cleanHeaders = allHeaders.filter(h => h && h.toString().trim() !== "").map(h => h.toString().trim());
  return createJsonResponse('success', { records, headers: cleanHeaders });
}

function updateRow(request) {
  const { formId, rowIndex, data, spreadsheetId } = request;
  const { sheet } = getTargetSheet(formId, spreadsheetId);
  if (!sheet) return createJsonResponse('error', null, 'ไม่พบ Sheet');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowOutput = headers.map(h => {
    const key = h.toString().trim();
    let val = (data[key] !== undefined) ? data[key] : data[key.toLowerCase()];
    
    if (typeof val === 'string' && val.includes('T') && val.includes('Z') && val.length > 15) {
      try {
        const d = new Date(val);
        if (!isNaN(d.getTime())) {
          val = (d.getFullYear() < 1905) ? Utilities.formatDate(d, "GMT+7", "H:mm") : Utilities.formatDate(d, "GMT+7", "dd/MM/yyyy");
        }
      } catch(e) {}
    }

    return Array.isArray(val) ? val.join(', ') : (val || "").toString();
  });
  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowOutput]);
  return createJsonResponse('success', null, 'อัปเดตเรียบร้อย');
}

function deleteProject(request) {
  const { formId, subject, spreadsheetId } = request;
  const { sheet } = getTargetSheet(formId, spreadsheetId);
  if (!sheet) return createJsonResponse('error', null, 'ไม่พบ Sheet');
  const lastRow = sheet.getLastRow();
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
  const { rowIndex, spreadsheetId } = request;
  const { sheet } = getTargetSheet('031', spreadsheetId);
  if (!sheet) return createJsonResponse('error', null, 'ไม่พบ Sheet');
  sheet.deleteRow(rowIndex);
  return createJsonResponse('success', null, 'ลบข้อมูลลำดับนั้นแล้ว');
}

function getDriveFiles(formId, spreadsheetId, customFolderId = null) {
  try {
    const config = getFormConfig(formId);
    if (!config || !config.folderId) return createJsonResponse('error', null, 'ไม่พบโฟลเดอร์หลักสำหรับ: ' + formId);

    const rootFolderId = config.folderId;
    const targetFolderId = customFolderId || rootFolderId;
    const folder = DriveApp.getFolderById(targetFolderId);
    const folderName = folder.getName();
    
    // 1. ดึงโฟลเดอร์ย่อยทั้งหมดที่อยู่ภายใต้โฟลเดอร์ปัจจุบัน
    const subFolders = folder.getFolders();
    const folderList = [];
    while (subFolders.hasNext()) {
      const sub = subFolders.next();
      folderList.push({
        id: sub.getId(),
        name: sub.getName()
      });
    }

    // 2. ดึงไฟล์ PDF ทั้งหมดที่อยู่ในโฟลเดอร์ปัจจุบัน
    const files = folder.getFilesByType(MimeType.PDF);
    const fileList = [];
    while (files.hasNext() && fileList.length < 100) {
      const file = files.next();
      fileList.push({
        id: file.getId(),
        name: file.getName(),
        date: Utilities.formatDate(file.getDateCreated(), "GMT+7", "dd/MM/yyyy HH:mm")
      });
    }

    // ทำการเรียงลำดับโฟลเดอร์ย่อยและไฟล์ตามตัวอักษรภาษาไทย
    folderList.sort((a, b) => a.name.localeCompare(b.name, 'th'));
    fileList.sort((a, b) => a.name.localeCompare(b.name, 'th'));

    return createJsonResponse('success', {
      rootFolderId: rootFolderId,
      currentFolderId: targetFolderId,
      folderName: folderName,
      folders: folderList,
      files: fileList
    });
  } catch (e) { 
    return createJsonResponse('error', null, 'เกิดข้อผิดพลาดในการดึงรายการไฟล์: ' + e.toString()); 
  }
}

function getFileBytes(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const base64 = Utilities.base64Encode(file.getBlob().getBytes());
    return createJsonResponse('success', base64);
  } catch (e) { return createJsonResponse('error', null, e.toString()); }
}

function authorizeProject() {
  const pres = SlidesApp.create('Authorization_Test_Slides');
  SlidesApp.openById(pres.getId());
  DriveApp.getRootFolder();
  DriveApp.getFileById(pres.getId()).setTrashed(true);
  Logger.log('ยินดีด้วย! คุณอนุญาตสิทธิ์ระบบ Slides เรียบร้อยแล้วครับ');
}

function batchUpdateRowsInSheet(formId, updates, spreadsheetId) {
  if (!updates || !Array.isArray(updates)) return { status: 'error', message: 'ไม่มีข้อมูลอัปเดต' };
  const { sheet } = getTargetSheet(formId, spreadsheetId);
  if (!sheet) return { status: 'error', message: 'ไม่พบ Sheet' };
  
  const allHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  updates.forEach(u => {
    const rowIndex = u.rowIndex;
    const data = u.data;
    if (rowIndex > 0) {
      const rowOutput = allHeaders.map(h => {
        const key = h.toString().trim();
        let val = (data[key] !== undefined) ? data[key] : data[key.toLowerCase()];
        
        if (typeof val === 'string' && val.includes('T') && val.includes('Z') && val.length > 15) {
          try {
            const d = new Date(val);
            if (!isNaN(d.getTime())) {
              val = (d.getFullYear() < 1905) ? Utilities.formatDate(d, "GMT+7", "H:mm") : Utilities.formatDate(d, "GMT+7", "dd/MM/yyyy");
            }
          } catch(e) {}
        }
        return Array.isArray(val) ? val.join(', ') : (val || "").toString();
      });
      sheet.getRange(rowIndex, 1, 1, allHeaders.length).setValues([rowOutput]);
    }
  });
  return { status: 'success', message: 'บันทึกข้อมูลแบบกลุ่มเรียบร้อยแล้ว' };
}

function verifyPassword(password) {
  const isCorrect = (password === '1234');
  return createJsonResponse('success', { isCorrect });
}

function getFormConfig(sheetName) {
  const config = getConfig();
  const name = sheetName.toString().replace(/[-_\s]/g, '').trim();
  let baseId = '';
  if (name.includes('031')) baseId = '031';
  else if (name.includes('033')) baseId = '033';
  else if (name.includes('034')) baseId = '034';
  else if (name.includes('035')) baseId = '035';
  
  if (baseId && config.FORMS[baseId]) {
    return { ...config.FORMS[baseId], baseId: baseId };
  }
  return null;
}

function getSheets(spreadsheetId) {
  try {
    const config = getConfig();
    const ssId = spreadsheetId || config.SPREADSHEET_ID;
    const ss = SpreadsheetApp.openById(ssId);
    const sheets = ss.getSheets();
    const result = sheets.map(s => {
      const name = s.getName();
      const configItem = getFormConfig(name);
      let rowCount = 0;
      if (configItem) {
        rowCount = Math.max(0, s.getLastRow() - 1);
      }
      return {
        name: name,
        isValid: configItem !== null,
        baseFormId: configItem ? configItem.baseId : null,
        rowCount: rowCount
      };
    }).filter(item => item.isValid);
    return createJsonResponse('success', result);
  } catch(e) {
    return createJsonResponse('error', null, e.toString());
  }
}

function setupWorkspace(password, targetFolderId) {
  if (password !== '1234') {
    return createJsonResponse('error', null, 'รหัสผ่านสำหรับการติดตั้งไม่ถูกต้อง');
  }

  try {
    let mainFolder;
    
    // ตรวจสอบและดึงโฟลเดอร์ Google Drive ของผู้ใช้ตามที่ระบุ
    if (targetFolderId) {
      try {
        mainFolder = DriveApp.getFolderById(targetFolderId);
      } catch (err) {
        return createJsonResponse('error', null, 'ไม่สามารถเข้าถึงโฟลเดอร์ Google Drive ตาม ID ที่ป้อนมาได้ กรุณาตรวจสอบลิงก์โฟลเดอร์และสิทธิ์การเข้าถึง');
      }
    } else {
      const mainFolderName = "ระบบจัดการเอกสารอัจฉริยะ (Smart Document Generator)";
      const folders = DriveApp.getFoldersByName(mainFolderName);
      if (folders.hasNext()) {
        mainFolder = folders.next();
      } else {
        mainFolder = DriveApp.createFolder(mainFolderName);
      }
    }
    
    // สร้างโฟลเดอร์ย่อย PDF_Outputs สำหรับเก็บผลลัพธ์
    const pdfFolders = {};
    const formTypes = ['031', '033', '034', '035'];
    formTypes.forEach(id => {
      const folderName = "PDF_Outputs_" + id;
      const subFolders = mainFolder.getFoldersByName(folderName);
      if (subFolders.hasNext()) {
        pdfFolders[id] = subFolders.next();
      } else {
        pdfFolders[id] = mainFolder.createFolder(folderName);
      }
    });

    // ทำสำเนาฐานข้อมูล Google Sheets ไปยังโฟลเดอร์ปลายทาง
    const MASTER_SPREADSHEET_ID = '1xAOgwRIoCLmhYYziPQuDTf_YtvDv-VDuCH3WjeIQou4';
    const timestamp = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm");
    const masterSheetFile = DriveApp.getFileById(MASTER_SPREADSHEET_ID);
    const newSheetFile = masterSheetFile.makeCopy("[ฐานข้อมูล] _ ระบบจัดการเอกสาร _ " + timestamp, mainFolder);
    const newSpreadsheetId = newSheetFile.getId();
    const newSpreadsheetUrl = newSheetFile.getUrl();

    // ทำสำเนา Google Slides ไปยังโฟลเดอร์ปลายทาง
    const MASTER_SLIDES = {
      '031': '1zcM4L5gSnNRLF2vgzpiE5E5a7C3g-o8qHG_BDppAm-A',
      '033': '1B4o6jYsKBGRKg2dS5qdCDsRHbpJA42GANk55sG_HKlQ',
      '034': '1CriGH1mgj1-MZcBvdNUfC5TYbMytf063U6Tau312ej0',
      '035': '11HT5_VmKVB7msoJ6j8oVH2NnLm2VGeGKyCt03SQQj3Y'
    };

    const newSlides = {};
    for (let key in MASTER_SLIDES) {
      const slideFile = DriveApp.getFileById(MASTER_SLIDES[key]);
      const copiedSlide = slideFile.makeCopy("Template_" + key + "_" + timestamp, mainFolder);
      newSlides[key] = copiedSlide.getId();
    }

    // บันทึกคอนฟิกใหม่ลงใน Properties
    const newConfig = {
      SPREADSHEET_ID: newSpreadsheetId,
      DEBUG_MODE: true,
      FORMS: {
        '031': { templateId: newSlides['031'], folderId: pdfFolders['031'].getId() },
        '033': { templateId: newSlides['033'], folderId: pdfFolders['033'].getId() },
        '034': { templateId: newSlides['034'], folderId: pdfFolders['034'].getId() },
        '035': { templateId: newSlides['035'], folderId: pdfFolders['035'].getId() }
      }
    };

    const props = PropertiesService.getScriptProperties();
    props.setProperty('USER_CONFIG', JSON.stringify(newConfig));

    return createJsonResponse('success', {
      spreadsheetUrl: newSpreadsheetUrl,
      spreadsheetId: newSpreadsheetId,
      config: newConfig
    }, 'ติดตั้งระบบและคัดลอกเทมเพลตไปยังโฟลเดอร์ของคุณเรียบร้อยแล้ว!');

  } catch (error) {
    return createJsonResponse('error', null, 'เกิดข้อผิดพลาดระหว่างติดตั้ง: ' + error.toString());
  }
}

// ตรวจสอบความถูกต้องว่าค่าเป็นตัวแปรวันที่หรือไม่
function isDate(val) {
  return val instanceof Date && !isNaN(val.valueOf());
}

// ดึงค่าคอนฟิกปัจจุบันส่งไปหน้าบ้าน
function getBackendConfig() {
  try {
    const config = getConfig();
    return createJsonResponse('success', config);
  } catch (e) {
    return createJsonResponse('error', null, 'เกิดข้อผิดพลาดในการดึงคอนฟิก: ' + e.toString());
  }
}

// บันทึกการตั้งค่าโครงสร้างคอนฟิกใหม่ลงใน Properties
function saveBackendConfig(configData) {
  try {
    if (!configData) {
      return createJsonResponse('error', null, 'ไม่พบข้อมูลการตั้งค่าสำหรับบันทึก');
    }
    const props = PropertiesService.getScriptProperties();
    props.setProperty('USER_CONFIG', JSON.stringify(configData));
    return createJsonResponse('success', configData, 'บันทึกการตั้งค่าระบบเรียบร้อยแล้ว!');
  } catch (e) {
    return createJsonResponse('error', null, 'เกิดข้อผิดพลาดในการบันทึกคอนฟิก: ' + e.toString());
  }
}

// สร้างโฟลเดอร์ย่อยใหม่ในโฟลเดอร์ที่กำหนด
function createSubFolder(parentFolderId, folderName) {
  try {
    if (!parentFolderId || !folderName) {
      return createJsonResponse('error', null, 'ระบุตำแหน่งที่ตั้งหรือชื่อโฟลเดอร์ย่อยไม่ครบถ้วน');
    }
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const newFolder = parentFolder.createFolder(folderName);
    return createJsonResponse('success', {
      id: newFolder.getId(),
      name: newFolder.getName()
    }, 'สร้างโฟลเดอร์ย่อย "' + folderName + '" เรียบร้อยแล้ว!');
  } catch (e) {
    return createJsonResponse('error', null, 'เกิดข้อผิดพลาดในการสร้างโฟลเดอร์: ' + e.toString());
  }
}

// ย้ายไฟล์ PDF ทั้งหมดที่กำหนดไปยังโฟลเดอร์ปลายทาง
function moveFiles(fileIds, targetFolderId) {
  try {
    if (!fileIds || !Array.isArray(fileIds) || !targetFolderId) {
      return createJsonResponse('error', null, 'ระบุข้อมูลย้ายไฟล์ไม่ครบถ้วน');
    }
    const targetFolder = DriveApp.getFolderById(targetFolderId);
    let successCount = 0;
    
    fileIds.forEach(id => {
      try {
        const file = DriveApp.getFileById(id);
        file.moveTo(targetFolder);
        successCount++;
      } catch (err) {
        // ข้ามหากไม่พบบางไฟล์ หรือไม่มีสิทธิ์
      }
    });
    
    return createJsonResponse('success', null, 'ย้ายไฟล์ย่อยจำนวน ' + successCount + ' รายการเข้าโฟลเดอร์เป้าหมายสำเร็จแล้ว!');
  } catch (e) {
    return createJsonResponse('error', null, 'เกิดข้อผิดพลาดในการย้ายไฟล์: ' + e.toString());
  }
}

// ตรวจจับและจัดสร้าง PDF ในแถวข้อมูลใหม่โดยอัตโนมัติ
function autoGenerateNewRows() {
  const config = getConfig();
  const ssId = config.SPREADSHEET_ID;
  if (!ssId) return;
  
  try {
    const ss = SpreadsheetApp.openById(ssId);
    const formTypes = ['031', '033', '034', '035'];
    
    formTypes.forEach(formId => {
      try {
        const cleanId = formId.toString().replace(/[-_\s]/g, '');
        const sheet = ss.getSheetByName(cleanId) || ss.getSheetByName(formId);
        if (!sheet) return;
        
        const lastRow = sheet.getLastRow();
        if (lastRow <= 1) return; // ไม่มีข้อมูลนอกจากหัวตาราง
        
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        
        // ค้นหาดัชนีของคอลัมน์ PDF_Link หรือ ลิงก์ไฟล์ PDF
        let pdfLinkColIdx = -1;
        headers.forEach((h, colIdx) => {
          const name = h.toString().trim().toLowerCase();
          if (name === 'pdf_link' || name === 'ลิงก์ไฟล์ pdf' || name === 'pdf link' || name === 'ลิงก์เอกสาร pdf') {
            pdfLinkColIdx = colIdx;
          }
        });
        
        let targetColIdx = pdfLinkColIdx + 1;
        if (pdfLinkColIdx === -1) {
          // หากไม่มี ให้สร้างคอลัมน์ชื่อ PDF_Link ต่อท้ายตารางฐานข้อมูลโดยอัตโนมัติ
          targetColIdx = headers.length + 1;
          sheet.getRange(1, targetColIdx).setValue('PDF_Link');
          // อัปเดตหัวตาราง
          headers.push('PDF_Link');
        }
        
        // สแกนแถวข้อมูลย้อนหลังจากแถวสุดท้าย (สแกนสูงสุด 100 แถวเพื่อประสิทธิภาพ)
        const scanStartRow = Math.max(2, lastRow - 99);
        const numRows = lastRow - scanStartRow + 1;
        
        const dataRange = sheet.getRange(scanStartRow, 1, numRows, headers.length);
        const values = dataRange.getValues();
        
        for (let i = 0; i < values.length; i++) {
          const rowVal = values[i];
          const pdfLinkValue = rowVal[targetColIdx - 1];
          
          // หากพบคอลัมน์ PDF_Link ว่างอยู่ (แสดงว่ายังไม่ได้ถูกสร้างเอกสาร)
          if (!pdfLinkValue || pdfLinkValue.toString().trim() === "") {
            const rowIndexInSheet = scanStartRow + i;
            
            // แปลงแถวข้อมูลเป็น Object map กับหัวตาราง
            const rowData = {};
            headers.forEach((h, colIdx) => {
              const cleanHeader = h ? h.toString().trim() : "";
              if (cleanHeader) {
                rowData[cleanHeader] = rowVal[colIdx];
              }
            });
            
            // ดึง Config ของแบบฟอร์มและสร้าง PDF
            const formConfig = getFormConfig(formId);
            if (formConfig) {
              const isPreview = false;
              let tableData = null;
              
              // ดำเนินการสร้างไฟล์ PDF
              const result = createPDF(formConfig, rowData, isPreview, tableData, formConfig.baseId);
              
              // อัปเดตลิงก์ PDF ที่ได้กลับไปบันทึกใน Google Sheets แถวนั้น
              sheet.getRange(rowIndexInSheet, targetColIdx).setValue(result.url);
              SpreadsheetApp.flush();
            }
          }
        }
      } catch (sheetErr) {
        Logger.log('เกิดข้อผิดพลาดในการตรวจสอบชีต ' + formId + ': ' + sheetErr.toString());
      }
    });
  } catch (err) {
    Logger.log('เกิดข้อผิดพลาดในการทำงานของระบบสร้างอัตโนมัติ: ' + err.toString());
  }
}

// ตรวจสอบและดึงสถานะการทำงานของ Installable Trigger
function getAutoTriggerStatus() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const isActive = triggers.some(t => t.getHandlerFunction() === 'autoGenerateNewRows');
    return createJsonResponse('success', { isActive });
  } catch (e) {
    return createJsonResponse('error', null, 'ไม่สามารถดึงสถานะทริกเกอร์ได้: ' + e.toString());
  }
}

// เปิดหรือปิดการทำงานของระบบสร้าง PDF อัตโนมัติ (Trigger)
function toggleAutoTrigger(enable, spreadsheetId) {
  try {
    const config = getConfig();
    const ssId = spreadsheetId || config.SPREADSHEET_ID;
    
    // ค้นหาและลบทริกเกอร์เดิมที่ชื่อฟังก์ชัน autoGenerateNewRows ทั้งหมดก่อนเพื่อรีเซ็ต
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(t => {
      if (t.getHandlerFunction() === 'autoGenerateNewRows') {
        ScriptApp.deleteTrigger(t);
      }
    });
    
    if (enable) {
      // 1. สร้างทริกเกอร์ดักจับเมื่อพบข้อมูลเปลี่ยนแปลง (onChange)
      const ss = SpreadsheetApp.openById(ssId);
      ScriptApp.newTrigger('autoGenerateNewRows')
               .forSpreadsheet(ss)
               .onChange()
               .create();
               
      // 2. สร้างทริกเกอร์อิงเวลา (Time-driven) ทำงานทุก 15 นาที เพื่อเป็นระบบแบ็คอัพสำรอง
      ScriptApp.newTrigger('autoGenerateNewRows')
               .timeBased()
               .everyMinutes(15)
               .create();
    }
    
    return createJsonResponse('success', { isActive: enable }, 'ปรับปรุงทริกเกอร์ระบบสร้างเอกสารอัตโนมัติสำเร็จแล้ว!');
  } catch (e) {
    return createJsonResponse('error', null, 'ไม่สามารถเปิด/ปิดทริกเกอร์ได้: ' + e.toString());
  }
}
