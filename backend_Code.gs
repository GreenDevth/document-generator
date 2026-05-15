// ==========================================================================
// GOOGLE APPS SCRIPT BACKEND V3.5-SLIDES-EDITION
// รองรับ Google Slides และ Smart Mapping ภาษาไทย
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
    if (action === 'getDriveFiles') return getDriveFiles(request.formId);
    if (action === 'getFileBytes') return getFileBytes(request.fileId);
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
  return createJsonResponse('success', { version: "v3.5-Slides-Power", headers });
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

    // 1. ดึงข้อมูลจากแถวแรกของตาราง (ถ้ามี) มาใส่ในข้อมูลหลัก เพื่อให้ Tag ในเนื้อหาหลักทำงานได้
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
    if (cleanId === '034' && tableData && tableData.length > 0) {
      tableData.forEach((row, i) => {
        const idx = i + 1;
        // ปรับปรุง: เฉพาะกรณีที่มีค่าส่งมาจริงๆ เท่านั้นถึงจะทับ (ป้องกันการเอาค่าว่างไปทับรายการที่ 1 ที่มีอยู่ใน rowData อยู่แล้ว)
        if (row.format || row['รูปแบบสื่อ']) finalData['format' + idx] = row.format || row['รูปแบบสื่อ'];
        if (row.ep || row['ตอน']) finalData['ep' + idx] = row.ep || row['ตอน'];
        if (row.teach || row['วิทยากร']) finalData['teach' + idx] = row.teach || row['วิทยากร'];
        if (row.dur || row['ความยาว']) finalData['dur' + idx] = row.dur || row['ความยาว'];
        if (row.dma || row['ผลิตแล้วเสร็จวันที่']) finalData['dma' + idx] = row.dma || row['ผลิตแล้วเสร็จวันที่'];
      });
    }

    // จัดการ Checkbox (แบบเรียบง่ายตามรูปที่ 2)
    const checkboxKeys = ['cb1', 'cb2', 'cb3', 'cb4', 'cb5', 'cb6'];
    checkboxKeys.forEach(k => {
      const val = finalData[k];
      const isChecked = (val === '✓' || val === true || val === 'true');
      const symbol = isChecked ? "✓" : ""; // ใช้เครื่องหมายถูกธรรมดา หรือว่างไว้
      presentation.replaceAllText("{{" + k + "}}", symbol);
      presentation.replaceAllText("{{ " + k + " }}", symbol);
      presentation.replaceAllText("{{" + k.toUpperCase() + "}}", symbol); // เพิ่มตัวพิมพ์ใหญ่ด้วย
    });

    // แทนที่ Tag ทั่วไปแบบ Case-Insensitive 
    // --- ปรับปรุง V3.6: เรียงลำดับ Key ตามความยาวเพื่อป้องกันการทับซ้อน (เช่น {{ep1}} ทับ {{ep10}}) ---
    const sortedKeys = Object.keys(finalData).sort((a, b) => b.length - a.length);
    
    for (let k of sortedKeys) {
      const v = finalData[k];
      if (k && v !== undefined) {
        let valStr = "";
        
        if (v instanceof Date) {
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
    
    // เติมตารางในสไลด์ (ถ้ามี)
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

/** เติมข้อมูลลงในตารางของ Google Slides */
function fillSlidesTable(presentation, tableData) {
  const slides = presentation.getSlides();
  slides.forEach(slide => {
    const tables = slide.getTables();
    tables.forEach(table => {
      // ค้นหาแถวที่มีตัวแปรตัวอย่าง เช่น {{subject}} หรือ {{format}}
      let templateRowIndex = -1;
      for (let r = 0; r < table.getNumRows(); r++) {
        for (let c = 0; c < table.getRow(r).getNumCells(); r++) {
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
            
            // แทนที่ตัวแปรในเซลล์
            for (let [k, v] of Object.entries(rowData)) {
              const valStr = (v === null || v === undefined) ? "" : v.toString();
              cellValue = cellValue.replace(new RegExp('{{\\s*' + k + '\\s*}}', 'gi'), valStr);
            }
            // ล้าง Tag ที่เหลือ
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
        
        // --- ปรับปรุง: ตรวจสอบและแปลง ISO Date String กลับเป็นวันที่ปกติก่อนลง Sheet ---
        if (typeof val === 'string' && val.includes('T') && val.includes('Z') && val.length > 15) {
          try {
            const d = new Date(val);
            if (!isNaN(d.getTime())) {
              // ถ้าปี < 1905 มักจะเป็น Duration หรือเวลาเฉยๆ ให้คงไว้ หรือถ้าเป็นวันที่ให้แปลงเป็น dd/MM/yyyy
              val = (d.getFullYear() < 1905) ? Utilities.formatDate(d, "GMT+7", "H:mm") : Utilities.formatDate(d, "GMT+7", "dd/MM/yyyy");
            }
          } catch(e) {}
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
  if (lastRow <= 1) return createJsonResponse('success', { records: [], headers: [] });
  const startRow = Math.max(2, lastRow - 499);
  const numRows = lastRow - startRow + 1;
  const allHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // ดึงทั้งค่าจริง (Values) และค่าที่แสดงผล (DisplayValues)
  const v = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();
  const d = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getDisplayValues();
  
  const records = d.map((r, idx) => {
    const obj = { _rowIndex: startRow + idx };
    allHeaders.forEach((h, i) => {
      if (!h) return;
      let val = r[i]; // เริ่มต้นด้วยค่าที่เห็นในหน้า Sheet
      let raw = v[idx][i];

      // กรณีพิเศษ: ถ้าค่าดิบเป็น Date Object (Google มักจะแปลง Duration เป็น Date)
      if (raw instanceof Date && raw.getFullYear() < 1905) {
        // คำนวณเวลาใหม่เองเพื่อป้องกัน Timezone เพี้ยน
        const hh = raw.getHours();
        const mm = raw.getMinutes();
        const ss = raw.getSeconds();
        if (hh > 0) {
          // ถ้ามีชั่วโมง: โชว์ hh:mm (และ :ss ถ้ามีวินาทีที่ไม่ใช่ 0)
          val = hh + ":" + (mm < 10 ? "0" + mm : mm);
          if (ss > 0) val += ":" + (ss < 10 ? "0" + ss : ss);
        } else {
          // ถ้าไม่มีชั่วโมง: โชว์ m:ss
          val = mm + ":" + (ss < 10 ? "0" + ss : ss);
        }
      }
      
      // ถ้าเป็น 0:00 ให้โชว์เป็นขีด
      if (val === "0:00" || val === "00:00" || val === "0:00:00") val = "-";
      
      obj[h] = val;
    });
    return obj;
  });
  const cleanHeaders = allHeaders.filter(h => h && h.toString().trim() !== "");
  return createJsonResponse('success', { records, headers: cleanHeaders });
}


function updateRow(request) {
  const { formId, rowIndex, data } = request;
  const { sheet } = getTargetSheet(formId);
  if (!sheet) return createJsonResponse('error', null, 'ไม่พบ Sheet');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowOutput = headers.map(h => {
    const key = h.toString().trim();
    let val = (data[key] !== undefined) ? data[key] : data[key.toLowerCase()];
    
    // --- ปรับปรุง: ตรวจสอบและแปลง ISO Date String สำหรับการ Update แถวด้วย ---
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
  const { formId, subject } = request;
  const { sheet } = getTargetSheet(formId);
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
  const { rowIndex } = request;
  const { sheet } = getTargetSheet('031'); // Default
  if (!sheet) return createJsonResponse('error', null, 'ไม่พบ Sheet');
  sheet.deleteRow(rowIndex);
  return createJsonResponse('success', null, 'ลบข้อมูลลำดับนั้นแล้ว');
}

function getDriveFiles(formId) {
  try {
    const cleanId = formId.toString().replace('.json', '').replace('-', '').trim();
    const config = CONFIG.FORMS[cleanId];
    if (!config || !config.folderId) return createJsonResponse('error', null, 'ไม่พบโฟลเดอร์สำหรับ: ' + cleanId);

    const folder = DriveApp.getFolderById(config.folderId);
    const folderName = folder.getName();
    const files = folder.getFilesByType(MimeType.PDF);
    const result = [];
    
    while (files.hasNext() && result.length < 50) {
      const file = files.next();
      result.push({
        id: file.getId(),
        name: file.getName(),
        date: Utilities.formatDate(file.getDateCreated(), "GMT+7", "dd/MM/yyyy HH:mm")
      });
    }

    if (result.length === 0) {
      const allFiles = folder.getFiles();
      const sampleFiles = [];
      while (allFiles.hasNext() && sampleFiles.length < 5) {
        const f = allFiles.next();
        sampleFiles.push(`${f.getName()} (${f.getMimeType()})`);
      }
      const debugMsg = sampleFiles.length > 0 ? `พบไฟล์อื่นแต่ไม่ใช่ PDF: ${sampleFiles.join(', ')}` : 'โฟลเดอร์นี้ว่างเปล่าครับ';
      return createJsonResponse('error', null, `ไม่พบไฟล์ PDF ในโฟลเดอร์ "${folderName}"\n\n[สถานะ: ${debugMsg}]`);
    }
    return createJsonResponse('success', result);
  } catch (e) { return createJsonResponse('error', null, 'เกิดข้อผิดพลาด: ' + e.toString()); }
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
