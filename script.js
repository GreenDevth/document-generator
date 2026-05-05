// ==========================================================================
// SMART DOCUMENT GENERATOR FRONTEND V2.9.3-TOTAL-POWER
// ==========================================================================

const scriptUrlInput = document.getElementById('scriptUrl');
if (scriptUrlInput) {
    scriptUrlInput.value = 'https://script.google.com/macros/s/AKfycbz9oaurUPmKICdUsGV34G_ahDBWhGLK6K5zirAHd2kbTP7Zx2XtniWI2189FszuKaIy/exec';
}
const formTypeSelect = document.getElementById('formType');
const dynamicSection = document.getElementById('dynamicContent');
const fieldsContainer = document.getElementById('fieldsContainer');
const isRecurringCheck = document.getElementById('isRecurring');
const recurringControls = document.getElementById('recurringControls');
const roundsInput = document.getElementById('recurringCount');
const intervalDaysInput = document.getElementById('intervalDays');
const resultBox = document.getElementById('resultBox');
const linksContainer = document.getElementById('linksContainer');
const btnLoadData = document.getElementById('btnLoadData');

if (btnLoadData) {
    btnLoadData.onclick = () => fetchRecentData();
}

const thaiMonths = [
    'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
    'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
];

let currentHeaders = [];
let isProcessing = false;
let currentEditRowIndex = 0;
let dashboardData = [];
let currentDashFormId = '031'; // ค่าเริ่มต้นของ Dashboard

// --- GLOBAL ACTIONS ---
async function clearLocalStorage() {
    if (await showModal('🧹 กวาดล้างข้อมูล', 'คุณต้องการล้างข้อมูลในฟอร์มใช่หรือไม่?', true, '🗑️')) {
        window.location.reload();
    }
}

async function clearForm() {
    if (await showModal('📋 เคลียร์ฟอร์ม', 'คุณต้องการลบข้อมูลที่พิมพ์ไว้และเริ่มสร้างรายการใหม่ใช่หรือไม่?', true, '✨')) {
        const tbody = document.getElementById('batchTableBody');
        if (tbody) tbody.innerHTML = '';
        const inputs = fieldsContainer.querySelectorAll('input');
        inputs.forEach(inp => inp.value = '');
        saveAllDrafts();
    }
}

const labelMap = {
    'subject': 'ชื่อรายการที่ผลิต', 'owner': 'ผู้รับผิดชอบการผลิต', 'extdocno': 'เอกสารเลขที่ (ภายนอก)',
    'extdocdate': 'ลงวันที่เอกสาร', 'activity': 'ชื่อกิจกรรม / รายการ', 'location': 'สถานที่จัดกิจกรรม',
    'actdate': 'วันที่จัดกิจกรรม', 'actmouth': 'เดือน (จัดงาน)', 'actac': 'พ.ศ. (จัดงาน)',
    'acttime': 'เวลา (จัดงาน)', 'findate': 'วันที่กำหนดส่ง', 'finmouth': 'เดือน (กำหนดส่ง)',
    'finac': 'พ.ศ. (กำหนดส่ง)', 'producer': 'ผู้รับผิดชอบการผลิต', 'team': 'ทีมงานผลิต',
    'checkbox': 'หน่วยงานที่ได้รับมอบหมาย', 'cb1': 'หน่วยจัดและผลิตรายการวิทยุ',
    'cb2': 'หน่วยจัดและผลิตรายการโทรทัศน์', 'cb3': 'หน่วยผลิตและพัฒนาสื่อการศึกษา',
    'cb4': 'ต้นฉบับ ( ) CD', 'cb5': 'ต้นฉบับ ( ) DVD', 'cb6': 'ต้นฉบับ ( ) อื่นๆ',
    'ep': 'ตอน', 'format': 'รูปแบบสื่อ', 'duration': 'ความยาวรายการ (นาที)',
    'second': 'ความยาวรายการ (วินาที)', 'sucdate': 'ผลิตแล้วเสร็จวันที่',
    'sucmouth': 'เดือน (ที่ผลิตเสร็จ)', 'sucac': 'พ.ศ. (ที่ผลิตเสร็จ)',
    'usedate': 'กำหนดออกอากาศ/นำไปใช้วันที่', 'usemouth': 'เดือน (ที่ออกอากาศ)',
    'useac': 'พ.ศ. (ที่ออกอากาศ)', 'more': 'หมายเหตุ / รายละเอียดเพิ่มเติม',
    'extdocnoint': 'เอกสารเลขที่ (ภายใน)', 'extdocdateint': 'ลงวันที่ (ภายใน)',
    'date': 'วันที่', 'mouth': 'เดือน', 'ac': 'พ.ศ.',
    'item': 'รายการอุปกรณ์', 'teach': 'วิทยากร/ผู้บรรยาย', 'dur': 'ความยาว',
    'dma': 'วันผลิตแล้วเสร็จ', 'qty': 'จำนวน', 'unit': 'หน่วย'
};
const checkboxConfig = { 'checkbox': ['หน่วยจัดและผลิตรายการวิทยุ', 'หน่วยจัดและผลิตรายการโทรทัศน์', 'หน่วยผลิตและพัฒนาสื่อการศึกษา'] };

if (isRecurringCheck) {
    isRecurringCheck.onchange = () => {
        if (recurringControls) recurringControls.style.display = isRecurringCheck.checked ? 'grid' : 'none';
    };
}

function toThaiDigits(num) { return num.toString().replace(/[0-9]/g, digit => "๐๑๒๓๔๕๖๗๘๙"[digit]); }
function toArabicDigits(str) { return str.toString().replace(/[๐-๙]/g, digit => "๐๑๒๓๔๕๖๗๘๙".indexOf(digit)); }

function calcNextDate(day, monthThai, yearBE, daysToAdd) {
    const dayArabic = toArabicDigits(day.toString());
    const yearArabic = toArabicDigits(yearBE.toString());
    const mIdx = thaiMonths.indexOf(monthThai.trim());
    if (mIdx === -1) return { day, month: monthThai, year: yearBE };
    const yearAD = parseInt(yearArabic) - 543;
    const date = new Date(yearAD, mIdx, parseInt(dayArabic));
    date.setDate(date.getDate() + daysToAdd);
    return {
        day: (/[๐-๙]/.test(day.toString()) || /[๐-๙]/.test(yearBE.toString())) ? toThaiDigits(date.getDate()) : date.getDate(),
        month: thaiMonths[date.getMonth()],
        year: (/[๐-๙]/.test(day.toString()) || /[๐-๙]/.test(yearBE.toString())) ? toThaiDigits(date.getFullYear() + 543) : date.getFullYear() + 543
    };
}

formTypeSelect.addEventListener('change', async () => {
    const url = scriptUrlInput.value.trim();
    const formId = formTypeSelect.value;
    if (!url || !formId || formId === "-- เลือกแบบฟอร์ม --") return;
    showLoading('กำลังดึงโครงสร้างข้อมูล...');
    try {
        const response = await fetch(url, {
            method: 'POST',
            body: JSON.stringify({ action: 'getSchema', formId }),
            headers: { 'Content-Type': 'text/plain;charset=utf-8' }
        });
        const json = await response.json();
        if (json.status === 'success') {
            renderFields(json.data.headers, formId);
            dynamicSection.classList.add('active');
            hideLoading();
        } else { throw new Error(json.message); }
    } catch (err) { hideLoading(); showModal('❌ ผิดพลาด', err.message, false, null, '⚠️'); }
});

function renderFields(headers, formId) {
    currentHeaders = headers;
    fieldsContainer.innerHTML = '';
    const fid = formId.toString().trim();
    
    if (fid === '034') {
        renderGrid034(headers);
        return;
    }

    const tableKeywords = ['ep', 'format', 'teach', 'dur', 'dma', 'item', 'qty', 'ตอน', 'รูปแบบสื่อ', 'วิทยากร', 'ความยาว', 'ผลิตแล้วเสร็จ'];
    const currentTableHeaders = headers.filter(h => tableKeywords.some(k => h.toLowerCase().includes(k)));
    const basicHeaders = headers.filter(h => !currentTableHeaders.includes(h));

    const cbHeaders = basicHeaders.filter(h => h.toLowerCase().startsWith('cb') || checkboxConfig[h]);
    const regularHeaders = basicHeaders.filter(h => !cbHeaders.includes(h));

    if (cbHeaders.length > 0) {
        const cbGroup = document.createElement('div');
        cbGroup.className = 'form-group full-width';
        cbGroup.style.background = 'rgba(79, 70, 229, 0.05)';
        cbGroup.style.padding = '20px';
        cbGroup.style.borderRadius = '15px';
        cbGroup.style.border = '1px solid var(--primary-light)';
        const label = document.createElement('label');
        label.textContent = 'หน่วยงานที่ได้รับมอบหมาย / หัวข้อเลือก';
        cbGroup.appendChild(label);
        const optionsContainer = document.createElement('div');
        optionsContainer.style.display = 'grid';
        optionsContainer.style.gridTemplateColumns = 'repeat(auto-fit, minmax(280px, 1fr))';
        optionsContainer.style.gap = '10px';
        cbHeaders.forEach(header => {
            if (checkboxConfig[header]) {
                checkboxConfig[header].forEach(opt => renderSingleCheckbox(optionsContainer, header, opt, opt));
            } else {
                renderSingleCheckbox(optionsContainer, header, labelMap[header.toLowerCase()] || header, '✓');
            }
        });
        cbGroup.appendChild(optionsContainer);
        fieldsContainer.appendChild(cbGroup);
    }

    let i = 0;
    while (i < regularHeaders.length) {
        const header = regularHeaders[i];
        const low = header.toLowerCase();
        const isGroupStart = (low.startsWith('act') || low.startsWith('fin') || low.startsWith('suc') || low.startsWith('use') || (low === 'date' && regularHeaders[i+1]?.toLowerCase() === 'mouth'));
        if (isGroupStart) {
            const groupPrefix = low.substring(0, 3);
            const groupContainer = document.createElement('div');
            groupContainer.className = 'date-row';
            while (i < regularHeaders.length && (regularHeaders[i].toLowerCase().startsWith(groupPrefix) || (groupPrefix === 'dat' && ['date', 'mouth', 'ac'].includes(regularHeaders[i].toLowerCase())))) {
                renderInputGroup(groupContainer, regularHeaders[i]);
                i++;
            }
            fieldsContainer.appendChild(groupContainer);
        } else {
            const group = document.createElement('div');
            group.className = 'form-group';
            if (low.includes('detail') || low.includes('remark') || low.includes('note') || low.includes('team')) group.classList.add('full-width');
            renderInputGroup(group, header);
            fieldsContainer.appendChild(group);
            i++;
        }
    }

    if (currentTableHeaders.length > 0) renderTableRegistry(currentTableHeaders);
}

function renderGrid034(headers) {
    const basicHeaders = headers.filter(h => !['ep', 'teach', 'dur', 'dma', 'format'].some(k => h.toLowerCase().includes(k)));
    const cbHeaders = basicHeaders.filter(h => h.toLowerCase().startsWith('cb') || checkboxConfig[h]);
    const regularHeaders = basicHeaders.filter(h => !cbHeaders.includes(h));

    if (cbHeaders.length > 0) {
        const cbGroup = document.createElement('div');
        cbGroup.className = 'form-group full-width checkbox-group-area';
        cbGroup.innerHTML = `<label style="font-weight:700; color:var(--primary); margin-bottom:15px; display:block;">🔹 หน่วยงานที่เกี่ยวข้อง / รูปแบบการจัดส่ง</label>`;
        const optionsContainer = document.createElement('div');
        optionsContainer.className = 'cb-grid-container';
        cbHeaders.forEach(header => {
            const low = header.toLowerCase();
            if (checkboxConfig[header]) {
                checkboxConfig[header].forEach(opt => renderSingleCheckbox(optionsContainer, header, opt, opt));
            } else {
                const displayLabel = labelMap[low] || header;
                renderSingleCheckbox(optionsContainer, header, displayLabel, '✓');
            }
        });
        cbGroup.appendChild(optionsContainer);
        fieldsContainer.appendChild(cbGroup);
    }

    regularHeaders.forEach(h => {
        const group = document.createElement('div');
        group.className = 'form-group';
        if (h.toLowerCase().includes('subject')) group.classList.add('full-width');
        renderInputGroup(group, h);
        fieldsContainer.appendChild(group);
    });

    const gridContainer = document.createElement('div');
    gridContainer.className = 'table-batch-container grid-034-container';
    gridContainer.innerHTML = `
        <h3>📝 รายการที่ผลิต (11 แถว)</h3>
        <div class="batch-table-wrapper">
            <table class="batch-table grid-034">
                <thead>
                    <tr>
                        <th>รูปแบบสื่อ</th>
                        <th>ชื่อเรื่อง/ตอน</th>
                        <th>วิทยากร/ผู้บรรยาย</th>
                        <th>ความยาว</th>
                        <th>วันผลิตแล้วเสร็จ</th>
                    </tr>
                </thead>
                <tbody id="grid034Body"></tbody>
            </table>
        </div>
    `;
    fieldsContainer.appendChild(gridContainer);
    const tbody = gridContainer.querySelector('#grid034Body');
    
    for (let r = 1; r <= 11; r++) {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td><input type="text" id="input_format${r}"></td>
            <td><input type="text" id="input_ep${r}"></td>
            <td><input type="text" id="input_teach${r}"></td>
            <td><input type="text" id="input_dur${r}"></td>
            <td><input type="text" id="input_dma${r}"></td>
        `;
        tbody.appendChild(tr);
    }
}

function renderTableRegistry(headers) {
    const container = document.createElement('div');
    container.className = 'table-batch-container';
    container.innerHTML = `
        <div class="batch-header">
            <div class="batch-title">📜 รายการที่ต้องบันทึก (${headers.length} คอลัมน์)</div>
            <div style="display: flex; gap: 10px;">
                <button type="button" class="btn-add" id="btnAddRow">เพิ่มรายการใหม่</button>
            </div>
        </div>
        <div class="batch-table-wrapper"><table class="batch-table"><thead><tr>
            ${headers.map(h => `<th>${labelMap[h.toLowerCase()] || h}</th>`).join('')}
            <th>ลบ</th>
        </tr></thead><tbody id="batchTableBody"></tbody></table></div>
    `;
    fieldsContainer.appendChild(container);
    container.querySelector('#btnAddRow').onclick = () => addBatchRow(headers);
    addBatchRow(headers);
}

function addBatchRow(headers, existingData = null) {
    const tbody = document.getElementById('batchTableBody');
    const rowIndex = tbody.children.length;
    const tr = document.createElement('tr');
    tr.className = 'batch-row';
    headers.forEach(h => {
        const td = document.createElement('td');
        const low = h.toLowerCase();
        let input;
        if (low === 'dur' || low === 'ความยาว') {
            const grp = document.createElement('div'); grp.className = 'unit-input-group';
            input = document.createElement('input'); input.type = 'text'; input.dataset.key = h; input.placeholder = '0.00';
            const sel = document.createElement('select'); sel.className = 'unit-select'; sel.dataset.unitFor = h;
            ['ชม.', 'น.'].forEach(u => { const o = document.createElement('option'); o.value = u; o.textContent = u; sel.appendChild(o); });
            grp.appendChild(input); grp.appendChild(sel);
            if (existingData && existingData[h]) {
                const m = existingData[h].toString().match(/^([\d.]+)\s*(.*)$/);
                if (m) { input.value = m[1]; sel.value = m[2] || 'ชม.'; } else input.value = existingData[h];
            }
            td.appendChild(grp);
        } else {
            input = document.createElement('input'); input.type = 'text'; input.dataset.key = h;
            input.placeholder = labelMap[low] || h;
            if (existingData && existingData[h]) {
                input.value = existingData[h];
            } else if (low === 'ep' && !existingData) {
                input.value = `EP${rowIndex + 1}`;
            } else if (!existingData && rowIndex > 0) {
                const prevRow = tbody.children[rowIndex - 1];
                const prevInput = prevRow.querySelector(`input[data-key="${h}"]`);
                if (prevInput && (low.includes('format') || low.includes('teach') || low.includes('dma') || low.includes('สื่อ') || low.includes('วิทยากร'))) {
                    input.value = prevInput.value;
                }
            }
            td.appendChild(input);
        }
        tr.appendChild(td);
    });
    const delTd = document.createElement('td');
    delTd.innerHTML = '<button type="button" class="btn-remove">ลบ</button>';
    delTd.querySelector('button').onclick = () => { tr.remove(); saveAllDrafts(); };
    tr.appendChild(delTd);
    tbody.appendChild(tr);
    saveAllDrafts();
}

// ฟังก์ชันช่วยบันทึกข้อมูลร่างไว้ในเครื่อง (Autosave)
function saveAllDrafts() {
    try {
        const tbody = document.getElementById('batchTableBody');
        if (!tbody) return;
        const rows = Array.from(tbody.rows);
        const data = rows.map(row => {
            const inputs = Array.from(row.querySelectorAll('input'));
            const rowData = {};
            inputs.forEach(inp => {
                if (inp.dataset.key) rowData[inp.dataset.key] = inp.value;
            });
            return rowData;
        });
        localStorage.setItem('table_data_draft', JSON.stringify(data));
    } catch (e) { console.warn('Autosave failed:', e); }
}

function checkTableRescue(headers) {
    const tableDraft = localStorage.getItem('table_data_draft');
    if (tableDraft) {
        const data = JSON.parse(tableDraft);
        const rescueContainer = document.getElementById('rescueContainer');
        const btn = document.createElement('button');
        btn.type = 'button'; btn.className = 'btn-warning';
        btn.innerHTML = `<span>⚡ เรียกคืนตารางเดิม (${data.length} รายการ)</span>`;
        btn.onclick = () => { restoreBatchData(data, headers); btn.remove(); };
        rescueContainer.appendChild(btn);
    } else { addBatchRow(headers); }
}

function restoreBatchData(data, headers) {
    const tbody = document.getElementById('batchTableBody');
    tbody.innerHTML = '';
    data.forEach(d => addBatchRow(headers, d));
}

function collectTableData() {
    const fid = formTypeSelect.value.toString().trim();
    if (fid === '034') {
        const row = {};
        // ดึงค่าแถวแรกไว้เป็นแม่แบบ (Template) สำหรับ Autofill
        const tplFormat = document.getElementById(`input_format1`).value;
        const tplTeach = document.getElementById(`input_teach1`).value;
        let tplDur = document.getElementById(`input_dur1`).value;
        if (tplDur && !tplDur.includes('ชม.')) tplDur = `${tplDur} ชม.`;
        const tplDma = document.getElementById(`input_dma1`).value;

        for (let r = 1; r <= 11; r++) {
            let f = document.getElementById(`input_format${r}`).value;
            let e = document.getElementById(`input_ep${r}`).value;
            let t = document.getElementById(`input_teach${r}`).value;
            let d = document.getElementById(`input_dur${r}`).value;
            let a = document.getElementById(`input_dma${r}`).value;

            // Smart Autofill: ถ้าแถว 2-11 ว่าง ให้ดึงค่าจากแถว 1 มาใช้
            if (r > 1) {
                if (!f) f = tplFormat;
                if (!t) t = tplTeach;
                if (!d) d = tplDur; else if (d && !d.includes('ชม.')) d = `${d} ชม.`;
                if (!a) a = tplDma;
            } else {
                if (d && !d.includes('ชม.')) d = `${d} ชม.`;
            }

            row[`format${r}`] = f;
            row[`ep${r}`] = e;
            row[`teach${r}`] = t;
            row[`dur${r}`] = d;
            row[`dma${r}`] = a;

            // เพิ่มเติม: สำหรับแถวที่ 1 ให้ส่งแบบไม่มีเลขกำกับด้วย เพื่อรองรับ Template ทุกรุ่น
            if (r === 1) {
                row['format'] = f;
                row['ep'] = e;
                row['teach'] = t;
                row['dur'] = d;
                row['dma'] = a;
            }
        }
        return [row];
    }
    
    const rows = document.querySelectorAll('.batch-row');
    const data = [];
    rows.forEach(tr => {
        const obj = {};
        tr.querySelectorAll('input[data-key]').forEach(input => {
            let val = input.value;
            const sel = tr.querySelector(`select[data-unit-for="${input.dataset.key}"]`);
            if (sel && val) val = `${val} ${sel.value}`;
            obj[input.dataset.key] = val;
        });
        data.push(obj);
    });
    return data.length > 0 ? data : null;
}

async function fetchRecentData(targetFormId = null) {
    const url = scriptUrlInput.value.trim();
    const formId = targetFormId || currentDashFormId || formTypeSelect.value;
    
    if (!url || !formId || formId === "-- เลือกแบบฟอร์ม --") {
        showModal('⚠️ คำเตือน', 'กรุณาเลือกประเภทแบบฟอร์มก่อนครับ', false, null, '🚧');
        return;
    }
    
    currentDashFormId = formId; // จำไว้ว่าตอนนี้ดู Tab ไหนอยู่
    showLoading(`กำลังดึงข้อมูล ${formId}...`);
    
    try {
        const response = await fetch(url, { 
            method: 'POST', 
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({ action: 'getRecentData', formId, limit: 100 }) 
        });
        const json = await response.json();
        if (json.status === 'success') {
            dashboardData = json.data.records;
            displayHistoryModal(dashboardData, json.data.headers);
        } else { throw new Error(json.message); }
    } catch (e) { showModal('❌ ผิดพลาด', e.message, false, null, '⚠️'); }
    finally { hideLoading(); }
}

function switchDashTab(formId) {
    currentDashFormId = formId;
    // อัปเดต UI ปุ่ม Tab
    document.querySelectorAll('.dash-tab').forEach(btn => {
        btn.classList.remove('active');
        if (btn.innerText.includes(formId)) btn.classList.add('active');
    });
    fetchRecentData(formId);
}

function displayHistoryModal(records, headers = []) {
    const modal = document.getElementById('dashboardModal');
    const headerRow = document.getElementById('dashHeaderRow');
    const fid = currentDashFormId || formTypeSelect.value.toString().trim();
    
    // ระบบ Fallback: ถ้าหาหัวข้อไม่เจอ ให้เอา Key จากข้อมูลมาใช้แทน
    let displayHeaders = (headers && headers.length > 0) ? headers : [];
    if (displayHeaders.length === 0 && records.length > 0) {
        displayHeaders = Object.keys(records[0]).filter(k => k !== '_rowIndex');
    }
    
    // กรองเอาเฉพาะ 25 คอลัมน์แรก (เพื่อให้ครอบคลุมข้อมูลส่วนใหญ่) และตัดคอลัมน์ว่าง
    displayHeaders = displayHeaders.filter(h => h && h.toString().trim() !== "").slice(0, 25);
    
    headerRow.innerHTML = `<th>ลำดับ</th>`;
    displayHeaders.forEach(h => {
        headerRow.innerHTML += `<th>${h}</th>`;
    });
    headerRow.innerHTML += `<th style="width: 200px; position: sticky; right: 0; background: var(--primary-light);">การปฏิบัติการ</th>`;
    
    // เพิ่มปุ่ม Refresh ใน UI Search
    const searchWrap = document.querySelector('.dash-search-container');
    const existingRefBtn = document.getElementById('btnRefreshDash');
    if (existingRefBtn) existingRefBtn.remove();
    const refBtn = document.createElement('button');
    refBtn.id = 'btnRefreshDash'; refBtn.type = 'button'; refBtn.className = 'btn-refresh';
    refBtn.innerHTML = '🔄 อัปเดตจาก Sheet';
    refBtn.onclick = () => fetchRecentData(fid);
    searchWrap.appendChild(refBtn);

    renderDashTable(records, displayHeaders, fid);
    modal.classList.add('active');
}

function renderDashTable(records, headers, dashboardFormId = null) {
    const body = document.getElementById('dashBody');
    const formId = (dashboardFormId || formTypeSelect.value).toString().trim();
    body.innerHTML = '';
    
    if (records.length === 0) {
        body.innerHTML = `<tr><td colspan="${headers.length + 2}" style="text-align:center;padding:50px;">📭 ไม่พบประวัติการบันทึกข้อมูล</td></tr>`;
        return;
    }

    records.forEach((r, idx) => {
        const tr = document.createElement('tr');
        tr.dataset.rowIndex = r._rowIndex;
        
        // แสดงเลขแถวที่แท้จริงจากใน Sheet
        let rowHtml = `<td>${r._rowIndex}</td>`;
        
        // สร้าง Input ตามหัวตารางที่มีจริง (ระบบ Smart Auto-Fit)
        headers.forEach(h => {
            const val = r[h] || '';
            const headerStr = h.toString();
            const valStr = val.toString();
            
            // คำนวณจำนวนตัวอักษรที่มากที่สุด (ระหว่างหัวข้อกับเนื้อหา)
            const charCount = Math.max(headerStr.length, valStr.length);
            
            // คำนวณความกว้าง (ประมาณ 10px ต่อตัวอักษร + 30px สำหรับ Padding)
            let widthNum = (charCount * 11) + 30;
            
            // จำกัดขอบเขต: ขั้นต่ำ 60px, ขั้นสูง 450px
            widthNum = Math.min(Math.max(widthNum, 60), 450);
            
            const width = widthNum + 'px';
            const textAlign = charCount <= 3 ? 'center' : 'left';

            rowHtml += `<td><input type="text" class="dash-inline-input" data-field="${h}" value="${val}" style="width: ${width}; min-width: ${width}; text-align: ${textAlign};"></td>`;
        });

        rowHtml += `
            <td style="position: sticky; right: 0; background: #fff; box-shadow: -5px 0 10px rgba(0,0,0,0.05);">
                <div class="action-btns">
                    <button type="button" class="btn-action btn-save" onclick="saveRecordInline(${idx}, ${JSON.stringify(headers).replace(/"/g, '&quot;')})" title="บันทึกลง Sheet">💾</button>
                    <button type="button" class="btn-action btn-edit" onclick="editRecord(${idx})" title="ดึงข้อมูลลงฟอร์ม">✍️</button>
                    <button type="button" class="btn-action btn-pdf" onclick="generateRecordPDF(${idx})" title="สร้าง PDF">📄</button>
                    <button type="button" class="btn-action btn-del" onclick="deleteRecord(${idx})" title="ลบ">🗑️</button>
                </div>
            </td>
        `;
        tr.innerHTML = rowHtml;
        body.appendChild(tr);
    });

    // เพิ่มปุ่มพริ้นต์รวบยอด (Global PDF) สำหรับ 034
    const searchContainer = document.querySelector('.dash-search-container');
    const existingGlobalBtn = document.getElementById('btnGlobalPDF034');
    if (existingGlobalBtn) existingGlobalBtn.remove();

    if (formId === '034') {
        const globalBtn = document.createElement('button');
        globalBtn.type = 'button';
        globalBtn.id = 'btnGlobalPDF034';
        globalBtn.className = 'btn-pdf-global';
        globalBtn.innerHTML = `
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 6 2 18 2 18 9"></polyline><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"></path><rect x="6" y="14" width="12" height="8"></rect></svg>
            สร้างไฟล์ PDF รวบยอด (034)
        `;
        globalBtn.onclick = generateBatchPDF034; // ผูกฟังก์ชันโดยตรง
        searchContainer.appendChild(globalBtn);
    }
}

async function generateBatchPDF034() {
    alert('กำลังเตรียมสร้าง PDF รวบยอดจากตาราง Dashboard ครับ...');
    try {
        const url = scriptUrlInput.value.trim();
        const formId = formTypeSelect.value.toString().trim();
        const rows = document.querySelectorAll('#dashBody tr');
        const tableData = [];
        
        rows.forEach(tr => {
            const inputs = tr.querySelectorAll('.dash-inline-input');
            if (inputs.length === 0) return;
            const rowObj = {};
            inputs.forEach(inp => {
                let val = inp.value;
                if (inp.dataset.field === 'dur' && val && !val.includes('ชม.')) val = `${val} ชม.`;
                rowObj[inp.dataset.field] = val;
            });
            tableData.push(rowObj);
        });

        if (tableData.length === 0) {
            showModal('⚠️ คำเตือน', 'ไม่พบข้อมูลในตาราง Dashboard ครับ', false, null, '🚧');
            return;
        }

        showLoading(`กำลังสร้าง PDF สรุป 034 (${tableData.length} รายการ)...`);
        const response = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({ action: 'generate', formId, data: tableData[0], tableData, rowIndex: 0 })
        });
        const result = await response.json();
        if (result.status === 'success') {
            showModal('📜 สร้างสรุปสำเร็จ', `เสร็จเรียบร้อยครับ`, false, () => {
                if (result.data.url) window.open(result.data.url, '_blank');
            }, '📄');
        } else { throw new Error(result.message); }
    } catch (e) { showModal('❌ ผิดพลาด', e.message, false, null, '⚠️'); }
    finally { hideLoading(); }
}

async function saveRecordInline(idx) {
    const r = dashboardData[idx];
    const tr = document.querySelector(`#dashBody tr:nth-child(${idx + 1})`);
    const url = scriptUrlInput.value.trim();
    
    // รวบรวมข้อมูลที่แก้ไขจากช่อง Input
    const inputs = tr.querySelectorAll('.dash-inline-input');
    const updatedData = { ...r };
    inputs.forEach(input => {
        const field = input.dataset.field;
        updatedData[field] = input.value;
        // Mapping กลับสำหรับคีย์ภาษาไทย (ถ้ามี)
        if (field === 'subject') updatedData['ชื่อรายการที่ผลิต'] = input.value;
        if (field === 'ep') updatedData['ตอน'] = input.value;
        if (field === 'teach') updatedData['วิทยากร/ผู้บรรยาย'] = input.value;
        if (field === 'dma') updatedData['วันผลิตแล้วเสร็จ'] = input.value;
    });

    showLoading('กำลังบันทึกข้อมูลลง Sheet...');
    try {
        const response = await fetch(url, {
            method: 'POST',
            body: JSON.stringify({
                action: 'updateRow',
                formId: formTypeSelect.value,
                rowIndex: r._rowIndex,
                data: updatedData
            })
        });
        const json = await response.json();
        if (json.status === 'success') {
            dashboardData[idx] = updatedData; // Update Local State
            showModal('✅ บันทึกสำเร็จ', `อัปเดตข้อมูลแถวที่ ${idx + 1} เรียบร้อยแล้วครับ`, false, '✨');
        } else { throw new Error(json.message); }
    } catch (e) { showModal('❌ ผิดพลาด', e.message, false, null, '⚠️'); }
    finally { hideLoading(); }
}

function closeDashboardModal() {
    document.getElementById('dashboardModal').classList.remove('active');
}

function filterDashboard() {
    const query = document.getElementById('dashSearch').value.toLowerCase();
    const filtered = dashboardData.filter(r => {
        const title = (r.subject || r['ชื่อรายการที่ผลิต'] || r.activity || '').toLowerCase();
        const owner = (r.owner || r['ผู้รับผิดชอบการผลิต'] || '').toLowerCase();
        return title.includes(query) || owner.includes(query);
    });
    renderDashTable(filtered);
}

function editRecord(idx) {
    const r = dashboardData[idx];
    const fid = currentDashFormId || formTypeSelect.value.toString().trim();
    
    // ตั้งค่าประเภทฟอร์มหลักให้ตรงกับที่เลือกใน Dashboard
    if (formTypeSelect.value !== fid) {
        formTypeSelect.value = fid;
        formTypeSelect.dispatchEvent(new Event('change'));
    }

    currentEditRowIndex = r._rowIndex;
    
    // แสดง Banner แจ้งเตือนการแก้ไข
    const banner = document.getElementById('editModeBanner');
    const rowNumSpan = document.getElementById('editRowNumber');
    if (banner && rowNumSpan) {
        banner.style.display = 'flex';
        rowNumSpan.textContent = r._rowIndex;
    }

    restoreBaseData(r);
    
    if (fid === '034') {
        // --- 034 Special: ดึงทุกบรรทัดในประวัติที่มีชื่อเรื่องเดียวกันลง Grid ---
        const subject = (r.subject || r['ชื่อรายการที่ผลิต'] || "").toString().toLowerCase().trim();
        const related = dashboardData.filter(row => {
            const rowSub = (row.subject || row['ชื่อรายการที่ผลิต'] || "").toString().toLowerCase().trim();
            return rowSub === subject && subject !== "";
        }).sort((a,b) => {
             const epA = parseInt((a.ep || a['ตอน'] || "0").toString().replace(/\D/g, '')) || 0;
             const epB = parseInt((b.ep || b['ตอน'] || "0").toString().replace(/\D/g, '')) || 0;
             return epA - epB;
        });

        // เคลียร์และเติมข้อมูลลง Grid 11 ช่อง
        for (let r = 1; r <= 11; r++) {
            const data = related[r-1] || {};
            const f = document.getElementById(`input_format${r}`);
            const e = document.getElementById(`input_ep${r}`);
            const t = document.getElementById(`input_teach${r}`);
            const d = document.getElementById(`input_dur${r}`);
            const a = document.getElementById(`input_dma${r}`);
            if(f) f.value = data.format || data['รูปแบบสื่อ'] || '';
            if(e) e.value = data.ep || data['ตอน'] || '';
            if(t) t.value = data.teach || data['วิทยากร/ผู้บรรยาย'] || data['วิทยากร'] || '';
            if(d) d.value = (data.dur || data['ความยาว'] || data['ความยาวรายการ (นาที)'] || '').toString().replace(' ชม.', '');
            if(a) a.value = data.dma || data['ผลิตแล้วเสร็จวันที่'] || data['วันผลิตแล้วเสร็จ'] || '';
        }
    } else {
        // กู้คืนตารางโหนด (ฟอร์มอื่น)
        const tableKeywords = ['ep', 'format', 'teach', 'dur', 'dma', 'item', 'qty', 'ตอน', 'รูปแบบสื่อ', 'วิทยากร', 'ความยาว', 'ผลิตแล้วเสร็จ'];
        const tableHeaders = currentHeaders.filter(h => tableKeywords.some(k => h.toLowerCase().includes(k)));
        if (tableHeaders.length > 0) {
            const tbody = document.getElementById('batchTableBody');
            if (tbody) {
                tbody.innerHTML = '';
                if (r.tableData || r.tableBody || r._tableData) {
                    try {
                        const rows = JSON.parse(r.tableData || r.tableBody || r._tableData);
                        rows.forEach(rowData => addBatchRow(tableHeaders, rowData));
                    } catch(e) { addBatchRow(tableHeaders, r); }
                } else { addBatchRow(tableHeaders, r); }
            }
        }
    }
    
    closeDashboardModal();
    showModal('✅ โหลดข้อมูลสำเร็จ', `ดึงข้อมูลโครงการ "${r.subject || 'ไม่ระบุชื่อ'}" มาเตรียมแก้ไขเรียบร้อยครับ`, false, null, '✨');
}

async function deleteRecord(idx) {
    const r = dashboardData[idx];
    const url = scriptUrlInput.value.trim();
    if (await showModal('🗑️ ยืนยันการลบ', 'คุณต้องการลบรายการนี้ออกจาก Google Sheet ใช่หรือไม่?', true, '🧨')) {
        showLoading('กำลังลบข้อมูล...');
        try {
            const response = await fetch(url, { 
                method: 'POST', 
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify({ action: 'deleteData', rowIndex: r._rowIndex }) 
            });
            const json = await response.json();
            if (json.status === 'success') {
                await showModal('✅ สำเร็จ', 'ลบข้อมูลเรียบร้อยแล้วครับ', false, '✨');
                fetchRecentData(); // Refresh table
            } else { throw new Error(json.message); }
        } catch (e) { showModal('❌ ผิดพลาด', e.message, false, '⚠️'); }
        finally { hideLoading(); }
    }
}

async function generateRecordPDF(idx) {
    const r = dashboardData[idx];
    const url = scriptUrlInput.value.trim();
    const formId = formTypeSelect.value;
    
    let tableData = null;
    let mainData = { ...r };

    // --- SMART LOGIC: 034 รวบกลุ่ม / อื่นๆ แถวละหน้า ---
    if (formId === '034') {
        showLoading('กำลังรวบรวมข้อมูลโครงการ (034)...');
        const subjectToMatch = (r.subject || r['ชื่อรายการที่ผลิต'] || "").toString().toLowerCase().trim();
        
        // กวาดหาทุกบรรทัดใน Dashboard ที่มี Subject เดียวกัน
        tableData = dashboardData.filter(row => {
            const rowSubject = (row.subject || row['ชื่อรายการที่ผลิต'] || "").toString().toLowerCase().trim();
            return rowSubject === subjectToMatch && subjectToMatch !== "";
        });

        // เรียงลำดับตาม EP เพื่อความสวยงามใน PDF
        tableData.sort((a, b) => {
            const epA = parseInt((a.ep || a['ตอน'] || "0").toString().replace(/\D/g, '')) || 0;
            const epB = parseInt((b.ep || b['ตอน'] || "0").toString().replace(/\D/g, '')) || 0;
            return epA - epB;
        });

        if (tableData.length === 0) tableData = [r]; // Fallback
    }

    showLoading(`กำลังสร้างไฟล์ PDF (${formId === '034' ? 'แบบหน้าเดียวรวมกลุ่ม' : 'แบบรายบรรทัด'})...`);
    try {
        const response = await fetch(url, { 
            method: 'POST', 
            body: JSON.stringify({ 
                action: 'generate', 
                formId: formId, 
                data: mainData, 
                tableData: tableData, 
                rowIndex: r._rowIndex 
            }) 
        });
        const result = await response.json();
        if (result.status === 'success') {
            showModal('📜 สร้างไฟล์สำเร็จ', 
                `สร้าง PDF ฟอร์ม ${formId} เรียบร้อยแล้วครับ<br><br>` +
                `<a href="${result.data.url}" target="_blank" style="display:inline-block; padding:10px 20px; background:#4f46e5; color:white; border-radius:8px; text-decoration:none; font-weight:bold;">🌐 เปิดดูไฟล์ PDF</a>`, 
                false, '📄');
        } else { throw new Error(result.message); }
    } catch (e) { showModal('❌ ผิดพลาด', e.message, false, '⚠️'); }
    finally { hideLoading(); }
}

async function sendData(action) {
    if (isProcessing) return;
    const url = scriptUrlInput.value.trim();
    const formId = formTypeSelect.value;
    const isRecurring = isRecurringCheck.checked && action === 'generate';
    const rounds = isRecurring ? (parseInt(roundsInput.value) || 1) : 1;
    isProcessing = true; resultBox.style.display = 'none'; linksContainer.innerHTML = '';
    showLoading('กำลังส่งข้อมูล...');
    try {
        const baseData = {};
        currentHeaders.forEach(h => {
            const checkboxes = document.querySelectorAll(`input[name="${h}"]`);
            if (checkboxes.length > 0) {
                const selected = Array.from(checkboxes).filter(cb => cb.checked).map(cb => cb.value);
                baseData[h] = selected.length > 1 ? selected : (selected.length === 1 ? selected[0] : "");
            } else {
                const el = document.getElementById(`input_${h}`);
                if (el) baseData[h] = el.value;
            }
        });
        const tableData = collectTableData();
        const interval = parseInt(intervalDaysInput.value) || 7;
        
        for (let r = 0; r < rounds; r++) {
            const currentData = { ...baseData };
            // โลจิกคำนวณวันที่ใหม่ (Dynamic detection for all date groups: actdate, findate, sucdate, etc.)
            if (r > 0) {
                Object.keys(currentData).forEach(key => {
                    const lowKey = key.toLowerCase();
                    if (lowKey.endsWith('date')) {
                        const prefix = lowKey.substring(0, lowKey.length - 4);
                        // ค้นหาฟิลด์ mouth และ ac ที่มี prefix เดียวกัน
                        const mKey = Object.keys(currentData).find(k => k.toLowerCase() === prefix + 'mouth');
                        const yKey = Object.keys(currentData).find(k => k.toLowerCase() === prefix + 'ac');
                        
                        if (mKey && yKey && currentData[key] && currentData[mKey] && currentData[yKey]) {
                            const next = calcNextDate(currentData[key], currentData[mKey], currentData[yKey], r * interval);
                            currentData[key] = next.day;
                            currentData[mKey] = next.month;
                            currentData[yKey] = next.year;
                        }
                    }
                });
            }

            const response = await fetch(url, { 
                method: 'POST', 
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify({ action, formId, data: currentData, tableData, rowIndex: currentEditRowIndex }) 
            });
            const result = await response.json();
            if (result.status === 'success') {
                addResultLink(result.data.url, result.data.name);
                // เมื่อบันทึกสำเร็จ และไม่ใช่การทำซ้ำ ให้ล้างโหมดแก้ไข
                if (rounds === 1) {
                    currentEditRowIndex = 0;
                    const banner = document.getElementById('editModeBanner');
                    if (banner) banner.style.display = 'none';
                }
            } else { throw new Error(result.message); }
        }
        resultBox.style.display = 'block'; resultBox.scrollIntoView({ behavior: 'smooth' });
    } catch (e) { showModal('❌ ข้อผิดพลาด', e.message, false, '⚠️'); }
    finally { isProcessing = false; hideLoading(); }
}

function renderInputGroup(container, h) {
    const low = h.toLowerCase();
    const label = document.createElement('label'); label.textContent = labelMap[low] || h;
    const input = document.createElement('input'); input.id = `input_${h}`; input.placeholder = label.textContent;
    container.appendChild(label); container.appendChild(input);
}

function renderSingleCheckbox(container, name, labelText, value) {
    const wrapper = document.createElement('label'); wrapper.className = 'cb-wrapper';
    wrapper.style.display = 'flex'; wrapper.style.alignItems = 'center'; wrapper.style.gap = '10px';
    const cb = document.createElement('input'); cb.type = 'checkbox'; cb.name = name; cb.value = value;
    wrapper.appendChild(cb); wrapper.appendChild(document.createTextNode(labelText));
    container.appendChild(wrapper);
}

function addResultLink(url, name) {
    const div = document.createElement('div'); div.className = 'batch-link';
    div.innerHTML = `<a href="${url}" target="_blank">📄 เปิดดูไฟล์: ${name}</a>`;
    linksContainer.appendChild(div);
}

function showModal(title, message, isConfirm = false, icon = '🔔') {
    return new Promise((resolve) => {
        const overlay = document.getElementById('modalOverlay');
        document.getElementById('modalIcon').textContent = icon;
        document.getElementById('modalTitle').textContent = title;
        const desc = document.getElementById('modalDesc');
        if (typeof message === 'string') desc.innerHTML = message;
        else { desc.innerHTML = ''; desc.appendChild(message); }
        
        const btnCancel = document.getElementById('modalCancel');
        const btnConfirm = document.getElementById('modalConfirm');
        
        btnCancel.style.display = isConfirm ? 'block' : 'none';
        btnConfirm.textContent = isConfirm ? 'ยืนยัน' : 'ตกลง';
        overlay.classList.add('active');

        btnConfirm.onclick = () => { 
            overlay.classList.remove('active'); 
            resolve(true); 
        };
        btnCancel.onclick = () => { 
            overlay.classList.remove('active'); 
            resolve(false); 
        };
    });
}

function showLoading(m) { document.getElementById('loadingOverlay').classList.add('active'); document.getElementById('loadingText').textContent = m; }
function hideLoading() { document.getElementById('loadingOverlay').classList.remove('active'); }

document.getElementById('docForm').onsubmit = async (e) => {
    e.preventDefault();
    const rounds = isRecurringCheck.checked ? (parseInt(roundsInput.value) || 1) : 1;
    if (await showModal('💾 ยืนยันการบันทึก', `ต้องการสร้างไฟล์หรือไม่?`, true)) {
        sendData('generate');
    }
};
document.getElementById('btnPreview').onclick = () => sendData('preview');

async function clearForm() {
    const ok = await showModal('📋 เคลียร์ฟอร์ม', 'คุณต้องการลบข้อมูลที่พิมพ์ไว้และเริ่มสร้างรายการใหม่ใช่หรือไม่?', true, '✨');
    if (ok) {
        const tbody = document.getElementById('batchTableBody');
        if (tbody) tbody.innerHTML = '';
        const inputs = fieldsContainer.querySelectorAll('input');
        inputs.forEach(inp => inp.value = '');
    }
}

// --- GENERATE ALL PDFS FROM DASHBOARD ---
async function generateAllPDFs() {
    if (!dashboardData || dashboardData.length === 0) {
        showModal('⚠️ คำเตือน', 'ไม่มีข้อมูลให้สร้าง PDF ครับ', false, null, '🚧');
        return;
    }

    const confirmed = await showModal('📄 ยืนยันการสร้างทั้งหมด', 
        `คุณต้องการสร้าง PDF จากรายการทั้งหมด ${dashboardData.length} รายการในหน้านี้ใช่หรือไม่?\n(ระบบจะใช้เวลาครู่หนึ่งตามจำนวนรายการครับ)`, true, '🖨️');
    if (!confirmed) return;

    const url = scriptUrlInput.value.trim();
    const formId = currentDashFormId || formTypeSelect.value;
    const successList = [];
    const errorList = [];

    showLoading(`เริ่มกระบวนการสร้าง PDF ทั้งหมด (0/${dashboardData.length})...`);

    for (let i = 0; i < dashboardData.length; i++) {
        const r = dashboardData[i];
        updateLoadingText(`กำลังสร้าง PDF รายการที่ ${i + 1}/${dashboardData.length}\n"${r.subject || r['ชื่อรายการที่ผลิต'] || 'ไม่ระบุชื่อ'}"`);
        
        try {
            let tableData = [];
            if (r.tableData || r.tableBody || r._tableData) {
                try { tableData = JSON.parse(r.tableData || r.tableBody || r._tableData); } 
                catch(e) { tableData = [r]; }
            } else { tableData = [r]; }

            const response = await fetch(url, { 
                method: 'POST', 
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify({ 
                    action: 'generate', 
                    formId: formId, 
                    data: r, 
                    tableData: tableData, 
                    rowIndex: r._rowIndex 
                }) 
            });
            const result = await response.json();
            if (result.status === 'success') {
                successList.push(result.data);
                addResultLink(result.data.url, result.data.name);
            } else {
                errorList.push({ name: r.subject || `แถวที่ ${r._rowIndex}`, message: result.message });
            }
        } catch (err) {
            errorList.push({ name: r.subject || `แถวที่ ${r._rowIndex}`, message: err.message });
        }
    }

    hideLoading();
    
    let summaryHtml = `<div style="text-align:left;">`;
    summaryHtml += `<p style="font-weight:700; color:var(--success);">✅ สร้างสำเร็จ: ${successList.length} รายการ</p>`;
    
    // เพิ่มปุ่มรวมไฟล์ถ้ามีมากกว่า 1 ไฟล์
    if (successList.length > 1) {
        summaryHtml += `<button type="button" class="btn-pdf-all" style="width:100%; margin-bottom:15px; background:var(--primary);" onclick='mergeAllGeneratedPDFs(${JSON.stringify(successList.map(s => s.url))})'>📄 รวมไฟล์ทั้งหมดเป็นไฟล์เดียว (.pdf)</button>`;
    }

    if (errorList.length > 0) summaryHtml += `<p style="font-weight:700; color:var(--error);">❌ ผิดพลาด: ${errorList.length} รายการ</p>`;
    
    summaryHtml += `<div style="max-height:200px; overflow-y:auto; margin-top:15px; border:1px solid #f1f5f9; padding:15px; border-radius:12px; background:#f8fafc; font-size:0.85rem;">`;
    successList.forEach(s => {
        summaryHtml += `<div style="margin-bottom:8px; border-bottom:1px solid #eee; padding-bottom:4px;">✔️ <a href="${s.url}" target="_blank" style="color:var(--primary); text-decoration:none;">${s.name}</a></div>`;
    });
    errorList.forEach(e => {
        summaryHtml += `<div style="margin-bottom:8px; color:var(--error); border-bottom:1px solid #eee; padding-bottom:4px;">✖️ ${e.name}: ${e.message}</div>`;
    });
    summaryHtml += `</div></div>`;

    showModal('🏁 กระบวนการเสร็จสิ้น', summaryHtml, false, '✨');
}

// --- PDF MERGE SYSTEM ---
async function mergeAllGeneratedPDFs(pdfUrls) {
    showLoading('กำลังดาวน์โหลดและรวมไฟล์ PDF... (กรุณารอสักครู่)');
    try {
        const { PDFDocument } = PDFLib;
        const mergedPdf = await PDFDocument.create();
        
        for (let i = 0; i < pdfUrls.length; i++) {
            updateLoadingText(`กำลังรวมไฟล์ที่ ${i + 1}/${pdfUrls.length}...`);
            let bytes;
            if (pdfUrls[i].startsWith('http')) {
                bytes = await fetch(pdfUrls[i]).then(res => res.arrayBuffer());
            } else {
                // กรณีเป็น Drive File ID
                const base64 = await callBackend('getFileBytes', { fileId: pdfUrls[i] });
                bytes = Uint8Array.from(atob(base64), c => c.charCodeAt(0)).buffer;
            }
            const pdf = await PDFDocument.load(bytes);
            const copiedPages = await mergedPdf.copyPages(pdf, pdf.getPageIndices());
            copiedPages.forEach((page) => mergedPdf.addPage(page));
        }
        
        const mergedPdfBytes = await mergedPdf.save();
        const blob = new Blob([mergedPdfBytes], { type: 'application/pdf' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `Combined_Documents_${new Date().getTime()}.pdf`;
        link.click();
        
        showModal('✨ รวมไฟล์สำเร็จ', 'ดาวน์โหลดไฟล์รวมขนาดใหญ่เรียบร้อยแล้วครับ', false, '✅');
    } catch (e) {
        console.error(e);
        showModal('❌ ผิดพลาดในการรวมไฟล์', 'ไม่สามารถรวมไฟล์ได้: ' + e.message, false, '⚠️');
    } finally {
        hideLoading();
    }
}

function updateLoadingText(text) {
    const desc = document.getElementById('loadingText');
    if (desc) desc.innerText = text;
}

// --- DRIVE FILE PICKER ---
async function openDriveFilePicker() {
    const formId = currentDashFormId || formTypeSelect.value;
    showLoading('กำลังโหลดรายชื่อไฟล์จาก Drive...');
    try {
        const response = await fetch(scriptUrlInput.value.trim(), {
            method: 'POST',
            body: JSON.stringify({ action: 'getDriveFiles', formId: formId })
        });
        const json = await response.json();
        hideLoading();
        
        if (json.status === 'success' && json.data.length > 0) {
            let html = '<div class="drive-list-container">';
            // เพิ่มส่วนหัว Select All
            html += `
                <div class="select-all-header">
                    <input type="checkbox" id="driveSelectAll" onchange="toggleDriveSelectAll(this)">
                    <label for="driveSelectAll" style="cursor:pointer;">เลือกไฟล์ทั้งหมด (${json.data.length} ไฟล์)</label>
                </div>
            `;
            
            json.data.forEach(f => {
                html += `
                <label class="drive-file-item">
                    <input type="checkbox" class="drive-file-cb" value="${f.id}" data-name="${f.name}"> 
                    <div>
                        <div style="font-weight:600; color:#1e293b;">${f.name}</div>
                        <div style="font-size:0.75rem; color:#64748b;">สร้างเมื่อ: ${f.date}</div>
                    </div>
                </label>`;
            });
            html += '</div>';
            
            const confirmed = await showModal('🗂️ เลือกไฟล์จาก Drive', html, true, '📁');
            if (confirmed) {
                const selected = Array.from(document.querySelectorAll('.drive-file-cb:checked')).map(cb => ({
                    url: cb.value, 
                    name: cb.dataset.name
                }));
                if (selected.length > 0) {
                    mergeAllGeneratedPDFs(selected.map(s => s.url));
                }
            }
        } else {
            // แสดงข้อความจริงจาก Backend (รวมถึงข้อมูล Debug ที่เราใส่ไว้)
            showModal('📁 แจ้งเตือน', json.message || 'ไม่พบไฟล์ PDF ในโฟลเดอร์นี้ครับ', false, 'ℹ️');
        }
    } catch (e) {
        hideLoading();
        showModal('❌ ผิดพลาด', e.message, false, '⚠️');
    }
}

async function callBackend(action, params) {
    const response = await fetch(scriptUrlInput.value.trim(), {
        method: 'POST',
        body: JSON.stringify({ action, ...params })
    });
    const json = await response.json();
    if (json.status === 'success') return json.data;
    throw new Error(json.message);
}

function toggleDriveSelectAll(masterCb) {
    const cbs = document.querySelectorAll('.drive-file-cb');
    cbs.forEach(cb => cb.checked = masterCb.checked);
}

// ไม่ใช้ AutoReload อีกต่อไป
function autoReload() { }
if (document.readyState === 'loading') { document.addEventListener('DOMContentLoaded', autoReload); } else { autoReload(); }
