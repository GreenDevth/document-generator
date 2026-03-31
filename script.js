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

const thaiMonths = [
    'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
    'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
];

let currentHeaders = [];
let isProcessing = false;
let currentEditRowIndex = 0;
let dashboardData = [];

// --- GLOBAL ACTIONS ---
function clearLocalStorage() {
    showModal('🧹 กวาดล้างข้อมูล', 'คุณต้องการล้างข้อมูลในฟอร์มใช่หรือไม่?', true, () => {
        window.location.reload();
    }, '🗑️');
}

function clearForm() {
    showModal('📋 เคลียร์ฟอร์ม', 'คุณต้องการลบข้อมูลที่พิมพ์ไว้และเริ่มสร้างรายการใหม่ใช่หรือไม่?', true, () => {
        const tbody = document.getElementById('batchTableBody');
        if (tbody) tbody.innerHTML = '';
        const inputs = fieldsContainer.querySelectorAll('input');
        inputs.forEach(inp => inp.value = '');
    }, '✨');
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

async function fetchRecentData() {
    const url = scriptUrlInput.value.trim();
    const formId = formTypeSelect.value;
    if (!url || !formId || formId === "-- เลือกแบบฟอร์ม --") {
        showModal('⚠️ คำเตือน', 'กรุณาเลือกประเภทแบบฟอร์มก่อนกวาดดูประวัติครับ', false, null, '🚧');
        return;
    }
    showLoading('กำลังดึงข้อมูลทั้งหมด...');
    try {
        const response = await fetch(url, { 
            method: 'POST', 
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({ action: 'getRecentData', formId, limit: 100 }) 
        });
        const json = await response.json();
        if (json.status === 'success') {
            dashboardData = json.data.records;
            displayHistoryModal(dashboardData);
        } else { throw new Error(json.message); }
    } catch (e) { showModal('❌ ผิดพลาด', e.message, false, null, '⚠️'); }
    finally { hideLoading(); }
}

function displayHistoryModal(records) {
    const modal = document.getElementById('dashboardModal');
    const headerRow = document.getElementById('dashHeaderRow');
    const fid = formTypeSelect.value.toString().trim(); // กำหนดค่า fid ให้ถูกต้อง
    
    // ตั้งค่าหัวตารางแบบ Spreadsheet ครบถ้วน
    headerRow.innerHTML = `
        <th>ลำดับ</th>
        <th>ชื่อรายการ (Subject) ${fid === '034' ? '<small style="display:block;color:#94a3b8;font-weight:300;">(รวบพริ้นต์ตารางเดียว)</small>' : ''}</th>
        <th>ตอน (EP)</th>
        <th>วิทยากร (Teacher)</th>
        <th>ความยาว (Dur)</th>
        <th>ผลิตเสร็จ (DMA)</th>
        <th>ผู้รับผิดชอบ (Owner)</th>
        <th>วันที่บันทึก</th>
        <th style="width: 250px; position: sticky; right: 0; background: var(--primary-light);">การปฏิบัติการ</th>
    `;
    
    // เพิ่มปุ่ม Refresh ใน UI Search
    const searchWrap = document.querySelector('.dash-search-container');
    const existingRefBtn = document.getElementById('btnRefreshDash');
    if (existingRefBtn) existingRefBtn.remove();
    const refBtn = document.createElement('button');
    refBtn.id = 'btnRefreshDash'; refBtn.type = 'button'; refBtn.className = 'btn-refresh';
    refBtn.innerHTML = '🔄 อัปเดตจาก Sheet';
    refBtn.onclick = fetchRecentData;
    searchWrap.appendChild(refBtn);

    renderDashTable(records, fid); // ส่ง fid เข้าไปด้วยเพื่อความแม่นยำ
    modal.classList.add('active');
}

function renderDashTable(records, dashboardFormId = null) {
    const body = document.getElementById('dashBody');
    const formId = (dashboardFormId || formTypeSelect.value).toString().trim();
    body.innerHTML = '';
    
    if (records.length === 0) {
        body.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:50px;">📭 ไม่พบประวัติการบันทึกข้อมูล</td></tr>';
        return;
    }

    records.forEach((r, idx) => {
        const tr = document.createElement('tr');
        tr.dataset.rowIndex = r._rowIndex;
        
        // กวาดหาข้อมูลจากหลายๆ คีย์ที่อาจเกิดขึ้นได้ (Mapping)
        const subject = r.subject || r['ชื่อรายการที่ผลิต'] || r.activity || '';
        const ep = r.ep || r['ตอน'] || '';
        const teach = r.teach || r['วิทยากร/ผู้บรรยาย'] || r['วิทยากร'] || '';
        const dur = r.dur || r['ความยาว'] || r['ความยาวรายการ (นาที)'] || '';
        const dma = r.dma || r['ผลิตแล้วเสร็จวันที่'] || r['วันผลิตแล้วเสร็จ'] || '';
        const owner = r.owner || r['ผู้รับผิดชอบการผลิต'] || '';
        const date = r.timestamp || r.Date || '-';
        
        tr.innerHTML = `
            <td>${idx + 1}</td>
            <td><input type="text" class="dash-inline-input" data-field="subject" value="${subject}"></td>
            <td><input type="text" class="dash-inline-input" data-field="ep" value="${ep}" style="width: 80px;"></td>
            <td><input type="text" class="dash-inline-input" data-field="teach" value="${teach}"></td>
            <td><input type="text" class="dash-inline-input" data-field="dur" value="${dur}" style="width: 100px;"></td>
            <td><input type="text" class="dash-inline-input" data-field="dma" value="${dma}" style="width: 120px;"></td>
            <td><input type="text" class="dash-inline-input" data-field="owner" value="${owner}"></td>
            <td style="font-size:0.75rem; color:#64748b; white-space: nowrap;">${date}</td>
            <td style="position: sticky; right: 0; background: #fff; box-shadow: -5px 0 10px rgba(0,0,0,0.05);">
                <div class="action-btns">
                    <button type="button" class="btn-action btn-save" onclick="saveRecordInline(${idx})" title="บันทึกลง Sheet">💾</button>
                    <button type="button" class="btn-action btn-edit" onclick="editRecord(${idx})" title="ดึงข้อมูลลงฟอร์ม">✍️</button>
                    <button type="button" class="btn-action btn-pdf" onclick="generateRecordPDF(${idx})" title="สร้าง PDF ใหม่จากข้อมูลใน Row นี้">📄</button>
                    <button type="button" class="btn-action btn-del" onclick="deleteRecord(${idx})" title="ลบออกจาก Sheet">🗑️</button>
                </div>
            </td>
        `;
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
            showModal('✅ บันทึกสำเร็จ', `อัปเดตข้อมูลแถวที่ ${idx + 1} เรียบร้อยแล้วครับ`, false, null, '✨');
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
    const fid = formTypeSelect.value.toString().trim();
    currentEditRowIndex = r._rowIndex;
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
    showModal('🗑️ ยืนยันการลบ', 'คุณต้องการลบรายการนี้ออกจาก Google Sheet ใช่หรือไม่?', true, async () => {
        showLoading('กำลังลบข้อมูล...');
        try {
            const response = await fetch(url, { 
                method: 'POST', 
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify({ action: 'deleteData', rowIndex: r._rowIndex }) 
            });
            const json = await response.json();
            if (json.status === 'success') {
                showModal('✅ สำเร็จ', 'ลบข้อมูลเรียบร้อยแล้วครับ', false, () => {
                    fetchRecentData(); // Refresh table inside modal
                }, '✨');
            } else { throw new Error(json.message); }
        } catch (e) { showModal('❌ ผิดพลาด', e.message, false, null, '⚠️'); }
        finally { hideLoading(); }
    }, '🧨');
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
            showModal('📜 สร้างไฟล์สำเร็จ', `สร้าง PDF ฟอร์ม ${formId} เรียบร้อยแล้วครับ`, false, () => {
                window.open(result.data.url, '_blank');
            }, '📄');
        } else { throw new Error(result.message); }
    } catch (e) { showModal('❌ ผิดพลาด', e.message, false, null, '⚠️'); }
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
            // โลจิกคำนวณวันที่ใหม่ถ้าเป็นรอบที่ 2 เป็นต้นไป
            const currentData = { ...baseData };
            if (r > 0 && currentData.date && currentData.mouth && currentData.ac) {
                const next = calcNextDate(currentData.date, currentData.mouth, currentData.ac, r * interval);
                currentData.date = next.day;
                currentData.mouth = next.month;
                currentData.ac = next.year;
            }

            const response = await fetch(url, { 
                method: 'POST', 
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify({ action, formId, data: currentData, tableData, rowIndex: 0 }) 
            });
            const result = await response.json();
            if (result.status === 'success') {
                addResultLink(result.data.url, result.data.name);
                if (rounds === 1) {
                    // ไม่ต้องทำความสะอาด Draft อีกต่อไป
                }
            } else { throw new Error(result.message); }
        }
        resultBox.style.display = 'block'; resultBox.scrollIntoView({ behavior: 'smooth' });
    } catch (e) { showModal('❌ ข้อผิดพลาด', e.message, false, null, '⚠️'); }
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

function showModal(title, message, isConfirm = false, onConfirm = null, icon = '🔔') {
    const overlay = document.getElementById('modalOverlay');
    document.getElementById('modalIcon').textContent = icon;
    document.getElementById('modalTitle').textContent = title;
    const desc = document.getElementById('modalDesc');
    if (typeof message === 'string') desc.textContent = message;
    else { desc.innerHTML = ''; desc.appendChild(message); }
    const btnCancel = document.getElementById('modalCancel');
    const btnConfirm = document.getElementById('modalConfirm');
    btnCancel.style.display = isConfirm ? 'block' : 'none';
    btnConfirm.textContent = isConfirm ? 'ยืนยัน' : 'ตกลง';
    overlay.classList.add('active');
    btnConfirm.onclick = () => { overlay.classList.remove('active'); if (onConfirm) onConfirm(); };
    btnCancel.onclick = () => { overlay.classList.remove('active'); };
}

function showLoading(m) { document.getElementById('loadingOverlay').classList.add('active'); document.getElementById('loadingText').textContent = m; }
function hideLoading() { document.getElementById('loadingOverlay').classList.remove('active'); }

document.getElementById('docForm').onsubmit = (e) => {
    e.preventDefault();
    const rounds = isRecurringCheck.checked ? (parseInt(roundsInput.value) || 1) : 1;
    showModal('💾 ยืนยันการบันทึก', `ต้องการสร้างไฟล์หรือไม่?`, true, () => sendData('generate'));
};
document.getElementById('btnPreview').onclick = () => sendData('preview');

function clearForm() {
    showModal('📋 เคลียร์ฟอร์ม', 'คุณต้องการลบข้อมูลที่พิมพ์ไว้และเริ่มสร้างรายการใหม่ใช่หรือไม่?', true, () => {
        const tbody = document.getElementById('batchTableBody');
        if (tbody) tbody.innerHTML = '';
        const inputs = fieldsContainer.querySelectorAll('input');
        inputs.forEach(inp => inp.value = '');
    }, '✨');
}

// ไม่ใช้ AutoReload อีกต่อไป
function autoReload() { }
if (document.readyState === 'loading') { document.addEventListener('DOMContentLoaded', autoReload); } else { autoReload(); }
