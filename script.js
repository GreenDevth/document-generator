// ==========================================================================
// SMART DOCUMENT GENERATOR FRONTEND V2.9.1-ULTIMATE-PERSISTENCE
// ==========================================================================

const scriptUrlInput = document.getElementById('scriptUrl');
const formTypeSelect = document.getElementById('formType');
const dynamicSection = document.getElementById('dynamicContent');
const fieldsContainer = document.getElementById('fieldsContainer');
const resultBox = document.getElementById('resultBox');
const isRecurringCheck = document.getElementById('isRecurring');
const recurringControls = document.getElementById('recurringControls');
const roundsInput = document.getElementById('recurringCount');
const intervalInput = document.getElementById('intervalDays');
const linksContainer = document.getElementById('linksContainer');

const thaiMonths = [
    'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
    'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
];

let currentHeaders = [];
let isProcessing = false;
let currentEditRowIndex = 0;

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
    'dma': 'วันผลิตแล้วเสร็จ', 'qty': 'จำนวน'
};

const checkboxConfig = { 'checkbox': ['หน่วยจัดและผลิตรายการวิทยุ', 'หน่วยจัดและผลิตรายการโทรทัศน์', 'หน่วยผลิตและพัฒนาสื่อการศึกษา'] };

const defaultUrl = 'https://script.google.com/macros/s/AKfycbz9oaurUPmKICd.../exec';
scriptUrlInput.value = localStorage.getItem('gas_url') || defaultUrl;

function toThaiDigits(num) { return num.toString().replace(/[0-9]/g, digit => "๐๑๒๓๔๕๖๗๘๙"[digit]); }
function toArabicDigits(str) { return str.toString().replace(/[๐-๙]/g, digit => "๐๑๒๓๔๕๖๗๘๙".indexOf(digit)); }

// --- PERSISTENCE LOGIC ---
function saveAllDrafts() {
    const baseData = {};
    currentHeaders.forEach(h => {
        const checkboxes = document.querySelectorAll(`input[name="${h}"]`);
        if (checkboxes.length > 0) {
            baseData[h] = Array.from(checkboxes).filter(cb => cb.checked).map(cb => cb.value);
        } else {
            const el = document.getElementById(`input_${h}`);
            if (el) baseData[h] = el.value;
        }
    });
    localStorage.setItem('form_id_draft', formTypeSelect.value);
    localStorage.setItem('base_data_draft', JSON.stringify(baseData));
    const tableData = collectTableData();
    if (tableData) localStorage.setItem('table_data_draft', JSON.stringify(tableData));
    else localStorage.removeItem('table_data_draft');
    console.log("💾 Global Data Saved Automatically");
}

function restoreBaseData() {
    const baseDraft = localStorage.getItem('base_data_draft');
    if (baseDraft) {
        const data = JSON.parse(baseDraft);
        for (let [h, val] of Object.entries(data)) {
            const checkboxes = document.querySelectorAll(`input[name="${h}"]`);
            if (checkboxes.length > 0) {
                const vals = Array.isArray(val) ? val : [val];
                checkboxes.forEach(cb => cb.checked = vals.includes(cb.value));
            } else {
                const el = document.getElementById(`input_${h}`);
                if (el) el.value = val;
            }
        }
    }
}

function calcNextDate(day, monthThai, yearBE, daysToAdd) {
    const dayArabic = toArabicDigits(day.toString());
    const yearArabic = toArabicDigits(yearBE.toString());
    const mIdx = thaiMonths.indexOf(monthThai.trim());
    if (mIdx === -1) return { day, month: monthThai, year: yearBE };
    const yearAD = parseInt(yearArabic) - 543;
    const date = new Date(yearAD, mIdx, parseInt(dayArabic));
    date.setDate(date.getDate() + daysToAdd);
    const resultDay = date.getDate();
    const resultYear = date.getFullYear() + 543;
    const isThaiInput = /[๐-๙]/.test(day.toString()) || /[๐-๙]/.test(yearBE.toString());
    return {
        day: isThaiInput ? toThaiDigits(resultDay) : resultDay,
        month: thaiMonths[date.getMonth()],
        year: isThaiInput ? toThaiDigits(resultYear) : resultYear
    };
}

formTypeSelect.addEventListener('change', async () => {
    const url = scriptUrlInput.value.trim();
    const formId = formTypeSelect.value;
    if (!url || !formId || formId === "-- เลือกแบบฟอร์ม --") return;
    
    localStorage.setItem('gas_url', url);
    localStorage.setItem('form_id_draft', formId);
    
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
            restoreBaseData();
            hideLoading();
        } else { throw new Error(json.message); }
    } catch (err) { hideLoading(); showModal('❌ ผิดพลาด', err.message, false, null, '⚠️'); }
});

function renderFields(headers, formId) {
    currentHeaders = headers;
    fieldsContainer.innerHTML = '';
    const tableKeywords = ['ep', 'format', 'teach', 'dur', 'dma', 'item', 'qty', 'ตอน', 'รูปแบบสื่อ', 'วิทยากร', 'ความยาว', 'ผลิตแล้วเสร็จ'];
    const currentTableHeaders = headers.filter(h => tableKeywords.some(k => h.toLowerCase().includes(k)));
    const basicHeaders = headers.filter(h => !currentTableHeaders.includes(h));

    // 1. Checkboxes
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

    // 2. Regular Fields
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

function renderTableRegistry(headers) {
    const container = document.createElement('div');
    container.className = 'table-batch-container';
    container.innerHTML = `
        <div id="rescueContainer"></div>
        <div class="batch-header">
            <div class="batch-title">📜 รายการที่ต้องบันทึก (${headers.length} คอลัมน์)</div>
            <button type="button" class="btn-add" id="btnAddRow">เพิ่มรายการใหม่</button>
        </div>
        <div class="batch-table-wrapper"><table class="batch-table"><thead><tr>
            ${headers.map(h => `<th>${labelMap[h.toLowerCase()] || h}</th>`).join('')}
            <th>ลบ</th>
        </tr></thead><tbody id="batchTableBody"></tbody></table></div>
    `;
    fieldsContainer.appendChild(container);
    container.querySelector('#btnAddRow').onclick = () => addBatchRow(headers);
    checkTableRescue(headers);
}

function addBatchRow(headers, existingData = null) {
    const tbody = document.getElementById('batchTableBody');
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
            input.oninput = () => saveAllDrafts(); sel.onchange = () => saveAllDrafts();
            if (existingData && existingData[h]) {
                const m = existingData[h].toString().match(/^([\d.]+)\s*(.*)$/);
                if (m) { input.value = m[1]; sel.value = m[2] || 'ชม.'; } else input.value = existingData[h];
            }
            td.appendChild(grp);
        } else {
            input = document.createElement('input'); input.type = 'text'; input.dataset.key = h;
            input.placeholder = labelMap[low] || h;
            input.oninput = () => saveAllDrafts();
            if (existingData && existingData[h]) input.value = existingData[h];
            else if (low === 'ep' && !existingData) input.value = `EP${tbody.children.length + 1}`;
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
        btn.innerHTML = `<span>⚡ เรียกคือข้อมูลล่าสุด (${data.length} รายการ)</span>`;
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

async function sendData(action) {
    if (isProcessing) return;
    const url = scriptUrlInput.value.trim();
    const formId = formTypeSelect.value;
    const rounds = isRecurringCheck.checked ? (parseInt(roundsInput.value) || 1) : 1;
    const interval = parseInt(intervalInput.value) || 7;
    
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
        for (let r = 0; r < rounds; r++) {
            const currentRoundData = { ...baseData };
            if (tableData && tableData.length > 0) {
                const first = tableData[0];
                for (let key in first) { if (!currentRoundData[key]) currentRoundData[key] = first[key]; }
            }
            if (r > 0) {
                const dFields = [{d:'actDate',m:'actMouth',y:'actAC'},{d:'finDate',m:'finMouth',y:'finAC'},{d:'sucDate',m:'sucMouth',y:'sucAC'},{d:'useDate',m:'useMouth',y:'useAC'},{d:'Date',m:'Mouth',y:'AC'}];
                dFields.forEach(g => {
                    if (currentRoundData[g.d] && currentRoundData[g.m] && currentRoundData[g.y]) {
                        const n = calcNextDate(currentRoundData[g.d], currentRoundData[g.m], currentRoundData[g.y], r * interval);
                        currentRoundData[g.d]=n.day; currentRoundData[g.m]=n.month; currentRoundData[g.y]=n.year;
                    }
                });
            }
            showLoading(`กำลังสร้างไฟล์... (${r+1}/${rounds})`);
            const response = await fetch(url, {
                method: 'POST',
                body: JSON.stringify({ action, formId, data: currentRoundData, tableData, rowIndex: 0 })
            });
            const result = await response.json();
            if (result.status === 'success') {
                addResultLink(result.data.url, result.data.name);
                if (rounds === 1) {
                    document.getElementById('resultMsg').innerHTML = `<strong>${result.message}</strong>`;
                    localStorage.removeItem('base_data_draft'); localStorage.removeItem('table_data_draft');
                }
            } else { throw new Error(result.message); }
            if (rounds > 1) await new Promise(res => setTimeout(res, 500));
        }
        if (rounds > 1) document.getElementById('resultMsg').innerHTML = `<strong>สร้างไฟล์ทั้งหมด ${rounds} ชุดสำเร็จ</strong>`;
        resultBox.style.display = 'block'; resultBox.scrollIntoView({ behavior: 'smooth' });
    } catch (e) { showModal('❌ ข้อผิดพลาด', e.message, false, null, '⚠️'); }
    finally { isProcessing = false; hideLoading(); }
}

function renderInputGroup(container, h) {
    const low = h.toLowerCase();
    const label = document.createElement('label'); label.textContent = labelMap[low] || h;
    const input = document.createElement('input'); input.id = `input_${h}`; input.placeholder = label.textContent;
    input.oninput = () => saveAllDrafts();
    container.appendChild(label); container.appendChild(input);
}

function renderSingleCheckbox(container, name, labelText, value) {
    const wrapper = document.createElement('label');
    wrapper.style.display = 'flex'; wrapper.style.alignItems = 'center'; wrapper.style.gap = '10px';
    const cb = document.createElement('input'); cb.type = 'checkbox'; cb.name = name; cb.value = value;
    cb.onchange = () => saveAllDrafts();
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
    document.getElementById('modalDesc').textContent = message;
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
    showModal('💾 ยืนยันการทำงาน', `ต้องการสร้างไฟล์ ${rounds} ชุดหรือไม่?`, true, () => sendData('generate'));
};
document.getElementById('btnPreview').onclick = () => sendData('preview');

// --- AUTO RELOAD ENGINE (V2.9.1) ---
function autoReload() {
    const savedFormId = localStorage.getItem('form_id_draft');
    if (savedFormId && savedFormId !== "-- เลือกแบบฟอร์ม --") {
        console.log("🔄 Auto-Reloading Form:", savedFormId);
        formTypeSelect.value = savedFormId;
        // ใช้ setTimeout เล็กน้อยเพื่อให้แน่ใจว่าระบบ Listener พร้อมทำงาน
        setTimeout(() => {
            formTypeSelect.dispatchEvent(new Event('change'));
        }, 100);
    }
}

// เรียกใช้งานเมื่อ DOM พร้อม
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', autoReload);
} else {
    autoReload();
}
