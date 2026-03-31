// ย้ายการประกาศตัวแปรหลักมาไว้ด้านบนสุดเพื่อให้มั่นใจว่าฟังก์ชันเรียกใช้ได้
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
let isProcessing = false; // ตัวแปรป้องกันการส่งซ้ำ

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
    'cb2': 'หน่วยผลิตและพัฒนาสื่อการศึกษา',
    'cb3': 'หน่วยจัดและผลิตรายการโทรทัศน์',
    'cb4': 'ต้นฉบับ ( ) CD',
    'cb5': 'ต้นฉบับ ( ) DVD',
    'cb6': 'ต้นฉบับ ( ) อื่นๆ',
    'ep': 'ตอน',
    'fomat': 'รูปแบบรายการ',
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
    'format': 'รูปแบบสื่อ',
    'teach': 'วิทยากร/ผู้บรรยาย',
    'dur': 'ความยาว',
    'dma': 'วันผลิตแล้วเสร็จ'
};

// 2. รายการตัวเลือกสำหรับ Checkbox แบบกลุ่ม (ถ้ามี)
const checkboxConfig = {
    'checkbox': [
        'หน่วยจัดและผลิตรายการวิทยุ',
        'หน่วยจัดและผลิตรายการโทรทัศน์',
        'หน่วยผลิตและพัฒนาสื่อการศึกษา'
    ]
};

// กำหนด URL ของ Web App เป็นค่าเริ่มต้น
const defaultUrl = 'https://script.google.com/macros/s/AKfycbwtcuDQ6gzAg3tibBAKL0aVVn5l3C_adShptRkgZaQNUYCzacuQKZv_ywlLYUw24c_l/exec';

// โหลด URL ที่เคยบันทึกไว้ หรือใช้ค่าเริ่มต้น
const savedUrl = localStorage.getItem('gas_url');
scriptUrlInput.value = savedUrl || defaultUrl;

// ฟังก์ชันแปลง อาราบิก -> ไทย
function toThaiDigits(num) {
    return num.toString().replace(/[0-9]/g, digit => "๐๑๒๓๔๕๖๗๘๙"[digit]);
}

// ฟังก์ชันแปลง ไทย -> อาราบิก
function toArabicDigits(str) {
    return str.toString().replace(/[๐-๙]/g, digit => "๐๑๒๓๔๕๖๗๘๙".indexOf(digit));
}

// เมื่อเลือกประเภทฟอร์ม
formTypeSelect.addEventListener('change', async () => {
    const url = scriptUrlInput.value.trim();
    const formId = formTypeSelect.value;

    if (!url || !formId || formId === "-- เลือกแบบฟอร์ม --") return;

    localStorage.setItem('gas_url', url);
    showLoading('กำลังดึงโครงสร้างข้อมูล (Schema)...');
    resultBox.style.display = 'none';

    try {
        const response = await fetch(url, {
            method: 'POST',
            body: JSON.stringify({
                action: 'getSchema',
                formId: formId
            }),
            headers: {
                'Content-Type': 'text/plain;charset=utf-8'
            }
        });

        const json = await response.json();
        if (json.status === 'success') {
            // แสดงปุ่มดึงข้อมูลเมื่อเลือกฟอร์มแล้ว
            btnLoadData.style.display = 'flex';

            // ตรวจสอบเวอร์ชันของสคริปต์ฝั่ง Server
            if (json.data.version !== "v2.6-GodMode") {
                showModal('⚠️ ตรวจพบสคริปต์เวอร์ชันเก่า!', 
                          'ระบบปัจจุบันยังเป็นเวอร์ชันเก่า (' + json.data.version + ') กรุณาเข้าไปที่ Apps Script แล้วกด Deploy (New Version) ให้เป็น v2.6 ครับ', 
                          false, null, '🚧');
            }
            
            renderFields(json.data.headers);
            dynamicSection.classList.add('active');
            // เมื่อโหลดเสร็จให้ซ่อน Loading
            hideLoading();
        } else {
            throw new Error(json.message);
        }
    } catch (err) {
        console.error(err);
        hideLoading();
        showModal('❌ ดึงข้อมูลไม่สำเร็จ', 'ไม่สามารถดึงโครงสร้างฟอร์มได้: ' + err.message, false, null, '⚠️');
        dynamicSection.classList.remove('active');
    }
});

// เปิด/ปิดแผงควบคุม Recurring
isRecurringCheck.addEventListener('change', () => {
    recurringControls.classList.toggle('active', isRecurringCheck.checked);
});

// ฟังก์ชันคำนวณวันถัดไป (รองรับ พ.ศ. และ เดือนไทย + เลขไทย)
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

// ฟังก์ชันสร้าง UI ฟิลด์กรอกข้อมูล
function renderFields(headers) {
    currentHeaders = headers;
    fieldsContainer.innerHTML = '';

    const cbHeaders = headers.filter(h => h.toLowerCase().startsWith('cb') || checkboxConfig[h]);
    const remainingHeaders = headers.filter(h => !cbHeaders.includes(h));

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
        label.style.color = 'var(--primary)';
        cbGroup.appendChild(label);

        const optionsContainer = document.createElement('div');
        optionsContainer.style.display = 'grid';
        optionsContainer.style.gridTemplateColumns = 'repeat(auto-fit, minmax(280px, 1fr))';
        optionsContainer.style.gap = '10px';
        optionsContainer.style.marginTop = '15px';

        cbHeaders.forEach(header => {
            const lowHeader = header.toLowerCase().trim();
            if (checkboxConfig[header]) {
                checkboxConfig[header].forEach(opt => {
                    renderSingleCheckbox(optionsContainer, header, opt, opt);
                });
            } else {
                renderSingleCheckbox(optionsContainer, header, labelMap[lowHeader] || header, '✓');
            }
        });

        cbGroup.appendChild(optionsContainer);
        fieldsContainer.appendChild(cbGroup);
    }

    let i = 0;
    while (i < remainingHeaders.length) {
        const header = remainingHeaders[i];
        const low = header.toLowerCase();

        const isGroupStart = (low.startsWith('act') || low.startsWith('fin') || low.startsWith('suc') || low.startsWith('use')) &&
            (low.includes('date') || low.includes('mouth') || low.includes('ac') || low.includes('time'));

        if (isGroupStart) {
            const groupPrefix = low.substring(0, 3);
            const groupContainer = document.createElement('div');
            groupContainer.className = 'date-row';

            while (i < remainingHeaders.length && remainingHeaders[i].toLowerCase().startsWith(groupPrefix)) {
                renderInputGroup(groupContainer, remainingHeaders[i]);
                i++;
            }
            fieldsContainer.appendChild(groupContainer);
        } else {
            const group = document.createElement('div');
            group.className = 'form-group';
            if (low.includes('detail') || low.includes('remark') || low.includes('note') || low.includes('address') || low.includes('team')) {
                group.classList.add('full-width');
            }
            renderInputGroup(group, header);
            fieldsContainer.appendChild(group);
            i++;
        }
    }
}

function renderInputGroup(container, header) {
    const low = header.toLowerCase().trim();
    const thaiLabel = labelMap[low] || header;

    const label = document.createElement('label');
    label.textContent = thaiLabel;

    let input;
    if (container.classList.contains('full-width')) {
        input = document.createElement('textarea');
        input.rows = 3;
    } else {
        input = document.createElement('input');
        if (low.includes('price') || low.includes('amount')) input.type = 'number';
        else input.type = 'text';
    }

    input.id = `input_${header}`;
    input.required = true;
    input.placeholder = thaiLabel;

    container.appendChild(label);
    container.appendChild(input);
}

function renderSingleCheckbox(container, name, labelText, value) {
    const wrapper = document.createElement('label');
    wrapper.style.display = 'flex';
    wrapper.style.alignItems = 'center';
    wrapper.style.gap = '10px';
    wrapper.style.cursor = 'pointer';
    wrapper.style.fontWeight = 'normal';
    wrapper.style.fontSize = '0.9rem';
    wrapper.style.textTransform = 'none';
    wrapper.style.marginBottom = '5px';

    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.name = name;
    cb.value = value;
    cb.style.width = '20px';
    cb.style.height = '20px';

    wrapper.appendChild(cb);
    wrapper.appendChild(document.createTextNode(labelText));
    container.appendChild(wrapper);
}

function showLoading(msg) {
    document.getElementById('loadingText').textContent = msg;
    document.getElementById('loadingOverlay').classList.add('active');
}

function hideLoading() {
    document.getElementById('loadingOverlay').classList.remove('active');
}

/**
 * ระบบ Premium Modal Center
 * @param {string} title หัวข้อ
 * @param {string} message ข้อความ
 * @param {boolean} isConfirm มีปุ่มยืนยัน/ยกเลิก หรือไม่
 * @param {function} onConfirm ฟังก์ชันคอลแบ็คเมื่อกดตกลง
 * @param {string} icon อีโมจิไอคอน
 */
function showModal(title, message, isConfirm = false, onConfirm = null, icon = '🔔') {
    const overlay = document.getElementById('modalOverlay');
    const iconEl = document.getElementById('modalIcon');
    const titleEl = document.getElementById('modalTitle');
    const descEl = document.getElementById('modalDesc');
    const btnCancel = document.getElementById('modalCancel');
    const btnConfirm = document.getElementById('modalConfirm');

    iconEl.textContent = icon;
    titleEl.textContent = title;
    descEl.textContent = message;
    btnCancel.style.display = isConfirm ? 'block' : 'none';
    btnConfirm.textContent = isConfirm ? 'ยืนยัน' : 'ตกลง';

    overlay.classList.add('active');

    const handleClose = (result) => {
        overlay.classList.remove('active');
        
        // ถอด Event ออกทันทีเพื่อป้องกันการทำงานซ้ำซ้อน
        btnConfirm.onclick = null;
        btnCancel.onclick = null;
        
        if (result && onConfirm) {
            // หน่วงเวลาเล็กน้อยเพื่อให้ Modal ตัวแรกปิดสนิทก่อนเปิดตัวถัดไป
            setTimeout(() => onConfirm(), 200);
        }
    };

    btnConfirm.onclick = () => handleClose(true);
    btnCancel.onclick = () => handleClose(false);
}

// --- ต่อสายไฟ (Event Listeners) กลับคืนเข้าระบบ ---
document.getElementById('docForm').addEventListener('submit', (e) => {
    e.preventDefault();
    e.stopImmediatePropagation();
    
    const isRecurring = isRecurringCheck.checked;
    if (isRecurring) {
        const rounds = parseInt(roundsInput.value) || 1;
        const interval = parseInt(intervalInput.value) || 7;
        showModal('🔄 ยืนยันการทำงานต่อเนื่อง', `ระบบกำลังจะสร้างเอกสารจำนวน ${rounds} ชุด (ห่างกันชุดละ ${interval} วัน) ใช่หรือไม่?`, true, () => {
            sendData('generate');
        }, '🔄');
    } else {
        // เพิ่มการยืนยันสำหรับโหมดปกติ 1 ชุด
        showModal('💾 ยืนยันการสร้างเอกสาร', 'คุณต้องการบันทึกข้อมูลและสร้างไฟล์ PDF 1 ชุด ใช่หรือไม่?', true, () => {
            sendData('generate');
        }, '📄');
    }
});

document.getElementById('btnPreview').addEventListener('click', () => {
    sendData('preview');
});

async function sendData(action) {
    if (isProcessing) return; // ถ้ากำลังทำงานอยู่ ให้หยุดทันที
    
    const url = scriptUrlInput.value.trim();
    const formId = formTypeSelect.value;
    
    const isRecurring = isRecurringCheck.checked && action === 'generate';
    const rounds = isRecurring ? parseInt(roundsInput.value) || 1 : 1;
    const interval = parseInt(intervalInput.value) || 7;

    isProcessing = true; // ล็อคสถานะ
    const btnSubmit = document.getElementById('btnSubmit');
    const btnPreview = document.getElementById('btnPreview');
    if (btnSubmit) btnSubmit.disabled = true;
    if (btnPreview) btnPreview.disabled = true;

    linksContainer.innerHTML = '';
    showLoading(action === 'preview' ? 'กำลังสร้างตัวอย่าง...' : `กำลังเตรียมข้อมูล (${rounds} รอบ)...`);
    resultBox.style.display = 'none';

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

        for (let r = 0; r < rounds; r++) {
            const currentRoundData = { ...baseData };
            
            if (r > 0) {
                const daysToAdd = r * interval;
                const dateFields = [
                    { d: 'actDate', m: 'actMouth', y: 'actAC' },
                    { d: 'finDate', m: 'finMouth', y: 'finAC' },
                    { d: 'sucDate', m: 'sucMouth', y: 'sucAC' },
                    { d: 'useDate', m: 'useMouth', y: 'useAC' }
                ];

                dateFields.forEach(group => {
                    if (currentRoundData[group.d] && currentRoundData[group.m] && currentRoundData[group.y]) {
                        const next = calcNextDate(currentRoundData[group.d], currentRoundData[group.m], currentRoundData[group.y], daysToAdd);
                        currentRoundData[group.d] = next.day;
                        currentRoundData[group.m] = next.month;
                        currentRoundData[group.y] = next.year;
                    }
                });
            }

            showLoading(`กำลังสร้างไฟล์รอบที่ ${r + 1}/${rounds}...`);
            console.log(`Processing Round ${r + 1}/${rounds}`, currentRoundData);

            const response = await fetch(url, {
                method: 'POST',
                body: JSON.stringify({
                    action: action,
                    formId: formId,
                    data: currentRoundData,
                    rowIndex: (rounds > 1) ? 0 : currentEditRowIndex // 0 = append, >1 = update
                }),
                headers: { 'Content-Type': 'text/plain;charset=utf-8' }
            });

            const result = await response.json();
            if (result.status === 'success') {
                addResultLink(result.data.url, result.data.name, r + 1, rounds);
            } else {
                throw new Error(`รอบที่ ${r + 1} ล้มเหลว: ` + result.message);
            }
            
            // หน่วงเวลาเล็กน้อยเพื่อให้ Backend ไม่ทำงานหนักเกินไป
            if (rounds > 1) await new Promise(res => setTimeout(res, 500));
        }

        document.getElementById('resultTitle').textContent = action === 'preview' ? 'สร้างตัวอย่างสำเร็จ' : 'สร้างไฟล์ทั้งหมดสำเร็จ! 🚀';
        document.getElementById('resultMsg').textContent = `สร้างเอกสารจำนวน ${rounds} ชุดเรียบร้อยแล้ว`;
        resultBox.style.display = 'block';
        resultBox.scrollIntoView({ behavior: 'smooth' });

    } catch (err) {
        console.error('Error in sendData:', err);
        alert('เกิดข้อผิดพลาด: ' + err.message);
    } finally {
        isProcessing = false; // ปลดล็อคสถานะ
        if (btnSubmit) btnSubmit.disabled = false;
        if (btnPreview) btnPreview.disabled = false;
        hideLoading();
    }
}

// เพิ่มลิงก์ลงในสถานะผลลัพธ์
function addResultLink(url, name, round, total) {
    const a = document.createElement('a');
    a.className = 'batch-link';
    a.href = url;
    a.target = '_blank';
    a.innerHTML = `
        <span>📄 ไฟล์รอบที่ ${round}: ${name}</span>
        <span style="color: var(--primary);">เปิดดู / พิมพ์ 📄</span>
    `;
    // จัดการซ่อนปุ่มดาวน์โหลดเดี่ยวของเดิม (ถ้ามี)
    if (round === 1) {
        const oldLink = document.getElementById('downloadLink');
        if (oldLink) oldLink.style.display = 'none';
    }
    linksContainer.appendChild(a);
}

// --- ระบบ Data Explorer (Reprint) ---
let currentEditRowIndex = 0; // 0 = สร้างใหม่, > 1 = แก้ไขแถวเดิม
const btnLoadData = document.getElementById('btnLoadData');
const dataOverlay = document.getElementById('dataOverlay');
const dataBody = document.getElementById('dataBody');
const editModeBanner = document.getElementById('editModeBanner');
const editRowNumber = document.getElementById('editRowNumber');

btnLoadData.addEventListener('click', fetchRecentData);

async function fetchRecentData() {
    const url = scriptUrlInput.value.trim();
    const formId = formTypeSelect.value;
    if (!url || !formId || formId === "-- เลือกแบบฟอร์ม --") {
        showModal('⚠️ คำเตือน', 'กรุณาเลือกประเภทฟอร์มก่อนดึงข้อมูลครับ', false, null, '💡');
        return;
    }

    showLoading('กำลังดึงข้อมูลล่าสุดจาก Sheet...');
    try {
        const response = await fetch(url, {
            method: 'POST',
            body: JSON.stringify({ action: 'getRecentData', formId: formId }),
            headers: { 'Content-Type': 'text/plain;charset=utf-8' }
        });
        const json = await response.json();
        if (json.status === 'success') {
            renderDataTable(json.data.records);
        } else {
            throw new Error(json.message);
        }
    } catch (err) {
        console.error('Error fetching data:', err);
        showModal('❌ ผิดพลาด', 'ไม่สามารถดึงข้อมูลได้: ' + err.message, false, null, '⚠️');
    } finally {
        hideLoading();
    }
}

function renderDataTable(records) {
    dataBody.innerHTML = '';
    if (!records || records.length === 0) {
        dataBody.innerHTML = '<tr><td colspan="4" style="text-align:center; padding: 40px;">ไม่พบข้อมูลใน Sheet นี้</td></tr>';
    } else {
        records.forEach((rec, idx) => {
            const tr = document.createElement('tr');
            // แมปคอลัมน์สำคัญ
            const date = rec['actDate'] || rec['วันที่จัดกิจกรรม'] || rec['actdate'] || '-';
            const subject = rec['subject'] || rec['ชื่อเรื่อง / วิชา'] || rec['subject'] || '-';
            const owner = rec['owner'] || rec['เจ้าของงาน / หน่วยงาน'] || rec['owner'] || '-';
            
            tr.innerHTML = `
                <td>${idx + 1}</td>
                <td style="white-space: nowrap;">${date}</td>
                <td><strong>${subject}</strong></td>
                <td>${owner}</td>
            `;
            tr.onclick = () => {
                applyDataToForm(rec);
                dataOverlay.classList.remove('active');
            };
            dataBody.appendChild(tr);
        });
    }
    dataOverlay.classList.add('active');
}

function applyDataToForm(data) {
    currentEditRowIndex = data._rowIndex || 0;
    
    // แสดง/ซ่อน Banner ตามสถานะ
    if (currentEditRowIndex > 1) {
        editRowNumber.textContent = currentEditRowIndex;
        editModeBanner.style.display = 'flex';
    } else {
        editModeBanner.style.display = 'none';
    }

    currentHeaders.forEach(h => {
        const val = data[h];
        const input = document.getElementById(`input_${h}`);
        if (input) {
            input.value = (val !== undefined && val !== null) ? val : "";
        } else {
            // จัดการ Checkbox
            const checkboxes = document.querySelectorAll(`input[name="${h}"]`);
            checkboxes.forEach(cb => {
                if (Array.isArray(val)) {
                    cb.checked = val.includes(cb.value);
                } else {
                    cb.checked = (val === cb.value || val === "✓");
                }
            });
        }
    });
    showModal('📥 โหลดข้อมูลสำเร็จ', 'ข้อมูลถูกนำเข้าสู่ฟอร์มแล้ว คุณสามารถแก้ไขและสั่งพิมพ์ใหม่ได้ทันทีครับ', false, null, '✅');
}

// ฟังก์ชันล้างฟอร์ม (สร้างรายการใหม่)
function clearForm() {
    showModal('🆕 เริ่มรายการใหม่', 'คุณต้องการล้างข้อมูลในฟอร์มเพื่อเริ่มกรอกรายการใหม่ ใช่หรือไม่?', true, () => {
        currentEditRowIndex = 0;
        editModeBanner.style.display = 'none';
        
        const inputs = document.querySelectorAll('#fieldsContainer input, #fieldsContainer textarea');
        inputs.forEach(input => {
            if (input.type === 'checkbox') input.checked = false;
            else input.value = '';
        });
        showModal('✨ เรียบร้อย', 'ล้างข้อมูลในฟอร์มเรียบร้อยแล้วครับ', false, null, '🧼');
    }, '💡');
}
