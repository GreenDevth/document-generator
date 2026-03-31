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

const labelMap = {
    'subject': 'ชื่อรายการที่ผลิต',
    'ep': 'ตอน',
    'format': 'รูปแบบสื่อ',
    'teach': 'วิทยากร/ผู้บรรยาย',
    'dur': 'ความยาว',
    'dma': 'วันผลิตแล้วเสร็จ'
};

const defaultUrl = 'https://script.google.com/macros/s/AKfycbwtcuDQ6gzAg3tibBAKL0aVVn5l3C_adShptRkgZaQNUYCzacuQKZv_ywlLYUw24c_l/exec';
const savedUrl = localStorage.getItem('gas_url');
scriptUrlInput.value = savedUrl || defaultUrl;

function toThaiDigits(num) { return num.toString().replace(/[0-9]/g, d => "๐๑๒๓๔๕๖๗๘๙"[d]); }
function toArabicDigits(str) { return str.toString().replace(/[๐-๙]/g, d => "๐๑๒๓๔๕๖๗๘๙".indexOf(d)); }

formTypeSelect.addEventListener('change', async () => {
    const url = scriptUrlInput.value.trim();
    const formId = formTypeSelect.value;
    if (!url || !formId || formId === "-- เลือกแบบฟอร์ม --") return;
    localStorage.setItem('gas_url', url);
    showLoading('กำลังดึงโครงสร้างข้อมูล...');
    try {
        const res = await fetch(url, { method: 'POST', body: JSON.stringify({ action: 'getSchema', formId }) });
        const json = await res.json();
        if (json.status === 'success') {
            if (json.data.version !== "v2.7-GodMode") {
                showModal('⚠️ เวอร์ชันไม่ตรงกัน!', `สคริปต์หลังบ้านเป็น ${json.data.version} กรุณา Deploy ใหม่เป็น v2.7 ครับ`, false, null, '🚧');
            }
            renderFields(json.data.headers, formId);
            dynamicSection.classList.add('active');
            hideLoading();
        }
    } catch (e) { hideLoading(); console.error(e); }
});

function renderFields(headers, formId) {
    currentHeaders = headers;
    fieldsContainer.innerHTML = '';
    const tableKeywords = ['ep', 'format', 'teach', 'dur', 'dma', 'item', 'qty', 'ตอน', 'รูปแบบสื่อ', 'วิทยากร', 'ความยาว', 'ผลิตแล้วเสร็จ'];
    const currentTableHeaders = headers.filter(h => tableKeywords.some(k => h.toLowerCase().includes(k)));
    const basicHeaders = headers.filter(h => !currentTableHeaders.includes(h));

    // Render Basic Fields
    basicHeaders.forEach(h => {
        const group = document.createElement('div');
        group.className = 'form-group';
        renderInputGroup(group, h);
        fieldsContainer.appendChild(group);
    });

    // Render Table
    if (currentTableHeaders.length > 0) renderTableRegistry(currentTableHeaders);
}

function saveTableDraft() {
    const data = collectTableData();
    if (data) localStorage.setItem('table_draft', JSON.stringify(data));
    else localStorage.removeItem('table_draft');
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

function addBatchRow(headers, existingData = null) {
    const tbody = document.getElementById('batchTableBody');
    const tr = document.createElement('tr');
    tr.className = 'batch-row';
    headers.forEach(h => {
        const td = document.createElement('td');
        const low = h.toLowerCase();
        let input;
        if (low === 'dur' || low === 'ความยาว') {
            const group = document.createElement('div'); group.className = 'unit-input-group';
            input = document.createElement('input'); input.type = 'text'; input.dataset.key = h; input.placeholder = '0.00';
            const sel = document.createElement('select'); sel.className = 'unit-select'; sel.dataset.unitFor = h;
            ['ชม.', 'น.'].forEach(u => { const o = document.createElement('option'); o.value = u; o.textContent = u; sel.appendChild(o); });
            group.appendChild(input); group.appendChild(sel);
            input.oninput = () => saveTableDraft(); sel.onchange = () => saveTableDraft();
            if (existingData && existingData[h]) {
                const m = existingData[h].toString().match(/^([\d.]+)\s*(.*)$/);
                if (m) { input.value = m[1]; sel.value = m[2] || 'ชม.'; } else input.value = existingData[h];
            }
            td.appendChild(group);
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
            const el = document.getElementById(`input_${h}`);
            if (el) baseData[h] = el.value;
        });

        const response = await fetch(url, {
            method: 'POST',
            body: JSON.stringify({ action, formId, data: baseData, tableData, rowIndex: 0 })
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

function addResultLink(url, name) {
    const div = document.createElement('div');
    div.className = 'batch-link';
    div.innerHTML = `<a href="${url}" target="_blank">📄 เปิดดูไฟล์: ${name}</a>`;
    linksContainer.appendChild(div);
}

function renderInputGroup(container, h) {
    const label = document.createElement('label'); label.textContent = labelMap[h.toLowerCase()] || h;
    const input = document.createElement('input'); input.id = `input_${h}`; input.placeholder = label.textContent;
    container.appendChild(label); container.appendChild(input);
}

function showModal(t, m) { alert(`${t}\n${m}`); }
function showLoading(m) { document.getElementById('loadingOverlay').classList.add('active'); document.getElementById('loadingText').textContent = m; }
function hideLoading() { document.getElementById('loadingOverlay').classList.remove('active'); }

document.getElementById('docForm').onsubmit = (e) => { e.preventDefault(); sendData('generate'); };
document.getElementById('btnPreview').onclick = () => sendData('preview');
