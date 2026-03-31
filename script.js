// ==========================================================================
// SMART DOCUMENT GENERATOR FRONTEND V2.7.1-GodMode-RESTORED
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
let currentEditRowIndex = 0; // 0 = New, >0 = Edit

// 1. ตารางแปลชื่อ (Technical Name -> Thai Label)
const labelMap = {
    'subject': 'ชื่อรายการที่ผลิต',
    'owner': 'ผู้รับผิดชอบการผลิต',
    'extdocno': 'เอกสารเลขที่ (ภายนอก)',
    'extdocdate': 'ลงวันที่เอกสาร',
    'activity': 'ชื่อกิจกรรม / รายการ',
    'location': 'สถานที่จัดกิจกรรม',
    'actdate': 'วันที่จัดกิจกรรม',
    'actmouth': 'เดือน (จัดงาน)',
    'actac': 'พ.ศ. (จัดงาน)',
    'acttime': 'เวลา (จัดงาน)',
    'findate': 'วันที่กำหนดส่ง',
    'finmouth': 'เดือน (กำหนดส่ง)',
    'finac': 'พ.ศ. (กำหนดส่ง)',
    'producer': 'ผู้รับผิดชอบการผลิต',
    'team': 'ทีมงานผลิต',
    'checkbox': 'หน่วยงานที่ได้รับมอบหมาย',
    'cb1': 'หน่วยจัดและผลิตรายการวิทยุ',
    'cb2': 'หน่วยจัดและผลิตรายการโทรทัศน์',
    'cb3': 'หน่วยผลิตและพัฒนาสื่อการศึกษา',
    'cb4': 'ต้นฉบับ ( ) CD',
    'cb5': 'ต้นฉบับ ( ) DVD',
    'cb6': 'ต้นฉบับ ( ) อื่นๆ',
    'ep': 'ตอน',
    'format': 'รูปแบบสื่อ',
    'duration': 'ความยาวรายการ (นาที)',
    'second': 'ความยาวรายการ (วินาที)',
    'sucdate': 'ผลิตแล้วเสร็จวันที่',
    'sucmouth': 'เดือน (ที่ผลิตเสร็จ)',
    'sucac': 'พ.ศ. (ที่ผลิตเสร็จ)',
    'usedate': 'กำหนดออกอากาศ/นำไปใช้วันที่',
    'usemouth': 'เดือน (ที่ออกอากาศ)',
    'useac': 'พ.ศ. (ที่ออกอากาศ)',
    'more': 'หมายเหตุ / รายละเอียดเพิ่มเติม',
    'extdocnoint': 'เอกสารเลขที่ (ภายใน)',
    'extdocdateint': 'ลงวันที่ (ภายใน)',
    'date': 'วันที่',
    'mouth': 'เดือน',
    'ac': 'พ.ศ.',
    'item': 'รายการอุปกรณ์',
    'teach': 'วิทยากร/ผู้บรรยาย',
    'dur': 'ความยาว',
    'dma': 'วันผลิตแล้วเสร็จ',
    'qty': 'จำนวน'
};

const checkboxConfig = {
    'checkbox': [
        'หน่วยจัดและผลิตรายการวิทยุ',
        'หน่วยจัดและผลิตรายการโทรทัศน์',
        'หน่วยผลิตและพัฒนาสื่อการศึกษา'
    ]
};

const defaultUrl = 'https://script.google.com/macros/s/AKfycbz9oaurUPmKICd.../exec';
const savedUrl = localStorage.getItem('gas_url');
scriptUrlInput.value = savedUrl || defaultUrl;

function toThaiDigits(num) { return num.toString().replace(/[0-9]/g, d => "๐๑๒๓๔๕๖๗๘๙"[d]); }
function toArabicDigits(str) { return str.toString().replace(/[๐-๙]/g, d => "๐๑๒๓๔๕๖๗๘๙".indexOf(d)); }

// เมื่อเลือกประเภทฟอร์ม
formTypeSelect.addEventListener('change', async () => {
    const url = scriptUrlInput.value.trim();
    const formId = formTypeSelect.value;
    if (!url || !formId || formId === "-- เลือกแบบฟอร์ม --") return;
    localStorage.setItem('gas_url', url);
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
        }
    } catch (err) { hideLoading(); console.error(err); }
});

isRecurringCheck.addEventListener('change', () => {
    recurringControls.classList.toggle('active', isRecurringCheck.checked);
});

function renderFields(headers, formId) {
    currentHeaders = headers;
    fieldsContainer.innerHTML = '';

    const tableKeywords = ['ep', 'format', 'teach', 'dur', 'dma', 'item', 'qty', 'ตอน', 'รูปแบบสื่อ', 'วิทยากร', 'ความยาว', 'ผลิตแล้วเสร็จ'];
    const currentTableHeaders = headers.filter(h => tableKeywords.some(k => h.toLowerCase().includes(k)));
    const basicHeaders = headers.filter(h => !currentTableHeaders.includes(h));

    // 1. เรนเดอร์ Checkboxes (CB1-3)
    const cbHeaders = basicHeaders.filter(h => h.toLowerCase().startsWith('cb') || checkboxConfig[h]);
    const regularHeaders = basicHeaders.filter(h => !cbHeaders.includes(h));

    if (cbHeaders.length > 0) {
        const cbGroup = document.createElement('div');
        cbGroup.className = 'form-group full-width';
        cbGroup.style.background = 'rgba(79, 70, 229, 0.05)';
        cbGroup.style.padding = '20px';
        cbGroup.style.borderRadius = '15px';
        cbGroup.style.marginBottom = '20px';
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

    // 2. เรนเดอร์ Regular Fields (จัดกลุ่มวันที่)
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

    // 3. เรนเดอร์ Dynamic Table Registry
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
    checkRescueData(headers);
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
            input.oninput = () => saveTableDraft(); sel.onchange = () => saveTableDraft();
            if (existingData && existingData[h]) {
                const m = existingData[h].toString().match(/^([\d.]+)\s*(.*)$/);
                if (m) { input.value = m[1]; sel.value = m[2] || 'ชม.'; } else input.value = existingData[h];
            }
            td.appendChild(grp);
        } else {
            input = document.createElement('input'); input.type = 'text'; input.dataset.key = h;
            input.placeholder = labelMap[low] || h;
            input.oninput = () => saveTableDraft();
            if (existingData && existingData[h]) input.value = existingData[h];
            else if (low === 'ep' && !existingData) input.value = `EP${tbody.children.length + 1}`;
            td.appendChild(input);
        }
        tr.appendChild(td);
    });
    const delTd = document.createElement('td');
    delTd.innerHTML = '<button type="button" class="btn-remove">ลบ</button>';
    delTd.querySelector('button').onclick = () => { tr.remove(); saveTableDraft(); };
    tr.appendChild(delTd);
    tbody.appendChild(tr);
}

function saveTableDraft() {
    const data = collectTableData();
    if (data) localStorage.setItem('table_draft', JSON.stringify(data));
    else localStorage.removeItem('table_draft');
}

function checkRescueData(headers) {
    const rescueData = localStorage.getItem('rescue_batch') || localStorage.getItem('table_draft');
    if (rescueData) {
        const data = JSON.parse(rescueData);
        const rescueContainer = document.getElementById('rescueContainer');
        const btn = document.createElement('button');
        btn.type = 'button'; btn.className = 'btn-warning';
        btn.innerHTML = `<span>⚡ เรียกคืนข้อมูลร่างเดิม (${data.length} รายการ)</span>`;
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
    const tableData = collectTableData();
    
    showLoading('กำลังส่งข้อมูล...');
    isProcessing = true; resultBox.style.display = 'none'; linksContainer.innerHTML = '';

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

        // ดึงข้อมูลแถวแรกมาใส่ baseData
        if (tableData && tableData.length > 0) {
            const first = tableData[0];
            for (let key in first) { if (!baseData[key]) baseData[key] = first[key]; }
        }

        const response = await fetch(url, {
            method: 'POST',
            body: JSON.stringify({ action, formId, data: baseData, tableData, rowIndex: currentEditRowIndex })
        });
        const result = await response.json();
        if (result.status === 'success') {
            addResultLink(result.data.url, result.data.name);
            document.getElementById('resultMsg').innerHTML = `<strong>${result.message}</strong>`;
            resultBox.style.display = 'block';
            localStorage.removeItem('table_draft');
        } else { throw new Error(result.message); }
    } catch (e) { alert(e.message); }
    finally { isProcessing = false; hideLoading(); }
}

function renderInputGroup(container, h) {
    const low = h.toLowerCase();
    const label = document.createElement('label'); label.textContent = labelMap[low] || h;
    let input;
    if (container.classList.contains('date-row') || !labelMap[low]) {
        input = document.createElement('input'); 
    } else {
        input = document.createElement('input');
    }
    input.id = `input_${h}`; input.placeholder = label.textContent;
    container.appendChild(label); container.appendChild(input);
}

function renderSingleCheckbox(container, name, labelText, value) {
    const wrapper = document.createElement('label');
    wrapper.style.display = 'flex'; wrapper.style.alignItems = 'center'; wrapper.style.gap = '10px';
    wrapper.style.cursor = 'pointer'; wrapper.style.fontWeight = 'normal'; wrapper.style.fontSize = '0.9rem';
    const cb = document.createElement('input'); cb.type = 'checkbox'; cb.name = name; cb.value = value;
    cb.style.width = '20px'; cb.style.height = '20px';
    wrapper.appendChild(cb); wrapper.appendChild(document.createTextNode(labelText));
    container.appendChild(wrapper);
}

function addResultLink(url, name) {
    const div = document.createElement('div');
    div.className = 'batch-link';
    div.innerHTML = `<a href="${url}" target="_blank">📄 เปิดดูไฟล์: ${name}</a>`;
    linksContainer.appendChild(div);
}

function showLoading(m) { document.getElementById('loadingOverlay').classList.add('active'); document.getElementById('loadingText').textContent = m; }
function hideLoading() { document.getElementById('loadingOverlay').classList.remove('active'); }
function showModal(t, m) { alert(`${t}\n${m}`); }

document.getElementById('docForm').onsubmit = (e) => { e.preventDefault(); sendData('generate'); };
document.getElementById('btnPreview').onclick = () => sendData('preview');
