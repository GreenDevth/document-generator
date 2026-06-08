// ==========================================================================
// SMART DOCUMENT GENERATOR MAIN JAVASCRIPT (VITE + SPA EDITION)
// ควบคุมและจัดการกระบวนการหน้าบ้านทั้งหมด
// ==========================================================================

import './style.css';

// --- ตัวแปรและสถานะการทำงานหลัก ---
let scriptUrl = '';
let activeSheet = ''; // แผ่นงานที่ผู้ใช้กำลังเปิดทำงานอยู่ในปัจจุบัน
let currentHeaders = [];
let isProcessing = false;
let currentEditRowIndex = 0;
let dashboardData = [];

const thaiMonths = [
    'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
    'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
];

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

// --- ส่วนตรวจสอบการเข้าสู่ระบบและเริ่มใช้งานหน้าเว็บ ---
document.addEventListener('DOMContentLoaded', () => {
    const savedUrl = sessionStorage.getItem('scriptUrl');
    const isLoggedIn = sessionStorage.getItem('isLoggedIn');

    // ตั้งค่า URL เริ่มต้นในหน้าล็อกอิน
    const scriptUrlInput = document.getElementById('scriptUrlInput');
    if (scriptUrlInput) {
        scriptUrlInput.value = savedUrl || 'https://script.google.com/macros/s/AKfycbydeCn_wKywlK6l9aRbcUZcjEbLfV1LweCCt7cfdk0Uwpx-ytDoIwiD5BUD2j7pjYXZ/exec';
    }

    if (isLoggedIn === 'true' && savedUrl) {
        scriptUrl = savedUrl;
        showPage('portalPage');
        loadDynamicSheets();
    } else {
        showPage('loginPage');
    }

    setupEventListeners();
});

// --- การควบคุมและจัดการอีเวนต์หลักของแอป ---
function setupEventListeners() {
    // ฟอร์มเข้าสู่ระบบ
    const loginForm = document.getElementById('loginForm');
    if (loginForm) {
        loginForm.onsubmit = async (e) => {
            e.preventDefault();
            const password = document.getElementById('authPassword').value;
            const url = document.getElementById('scriptUrlInput').value.trim();

            if (!url) {
                showModal('⚠️ คำเตือน', 'กรุณาระบุ Apps Script Web App URL', false);
                return;
            }

            showLoading('กำลังเข้าสู่ระบบและตรวจสอบสิทธิ์...');
            try {
                const response = await fetch(url, {
                    method: 'POST',
                    body: JSON.stringify({ action: 'verifyPassword', password: password }),
                    headers: { 'Content-Type': 'text/plain;charset=utf-8' }
                });
                const result = await response.json();

                if (result.status === 'success' && result.data.isCorrect) {
                    scriptUrl = url;
                    sessionStorage.setItem('scriptUrl', url);
                    sessionStorage.setItem('isLoggedIn', 'true');
                    showPage('portalPage');
                    loadDynamicSheets();
                } else {
                    throw new Error('รหัสผ่านไม่ถูกต้อง กรุณาลองใหม่อีกครั้ง');
                }
            } catch (err) {
                showModal('❌ ข้อผิดพลาด', err.message, false);
            } finally {
                hideLoading();
            }
        };
    }

    // ปุ่มออกจากระบบ
    const btnLogoutPortal = document.getElementById('btnLogoutPortal');
    if (btnLogoutPortal) {
        btnLogoutPortal.onclick = () => {
            sessionStorage.removeItem('isLoggedIn');
            scriptUrl = '';
            showPage('loginPage');
            // รีเซ็ตรหัสผ่านในช่อง
            const pass = document.getElementById('authPassword');
            if (pass) pass.value = '';
        };
    }

    // ปุ่มย้อนกลับจาก Workspace ไป Portal
    const btnBackToPortal = document.getElementById('btnBackToPortal');
    if (btnBackToPortal) {
        btnBackToPortal.onclick = () => {
            showPage('portalPage');
            loadDynamicSheets(); // รีโหลดชีตใหม่เพื่ออัปเดตสถิติ
        };
    }

    // ปุ่มสลับแท็บภายใน Workspace
    const tabFormBtn = document.getElementById('tabFormBtn');
    const tabTableBtn = document.getElementById('tabTableBtn');

    if (tabFormBtn && tabTableBtn) {
        tabFormBtn.onclick = () => switchWorkspaceView('form');
        tabTableBtn.onclick = () => switchWorkspaceView('table');
    }

    // ปุ่มเคลียร์ฟอร์ม
    const btnClearFormBtn = document.getElementById('btnClearFormBtn');
    if (btnClearFormBtn) {
        btnClearFormBtn.onclick = () => clearForm();
    }

    // ปุ่มยกเลิกการแก้ไข (Edit Banner)
    const btnCancelEdit = document.getElementById('btnCancelEdit');
    if (btnCancelEdit) {
        btnCancelEdit.onclick = () => cancelEditMode();
    }

    // ปุ่มดู Preview หน้ากรอกข้อมูล
    const btnPreview = document.getElementById('btnPreview');
    if (btnPreview) {
        btnPreview.onclick = () => sendData('preview');
    }

    // ฟอร์มหลัก หน้ากรอกข้อมูล
    const docForm = document.getElementById('docForm');
    if (docForm) {
        docForm.onsubmit = async (e) => {
            e.preventDefault();
            if (await showModal('💾 ยืนยันการบันทึก', `คุณต้องการสร้างไฟล์เอกสาร PDF ใช่หรือไม่?`, true)) {
                sendData('generate');
            }
        };
    }

    // ปุ่มควบคุมหน้าตาราง Dashboard
    const btnRefreshWorkspace = document.getElementById('btnRefreshWorkspace');
    if (btnRefreshWorkspace) {
        btnRefreshWorkspace.onclick = () => fetchRecentData(activeSheet);
    }

    const btnDrivePickerWorkspace = document.getElementById('btnDrivePickerWorkspace');
    if (btnDrivePickerWorkspace) {
        btnDrivePickerWorkspace.onclick = () => openDriveFilePicker();
    }

    const btnPDFAllWorkspace = document.getElementById('btnPDFAllWorkspace');
    if (btnPDFAllWorkspace) {
        btnPDFAllWorkspace.onclick = () => generateAllPDFs();
    }

    // สวิตช์การส่งต่อเนื่อง (Recurring)
    const isRecurringCheck = document.getElementById('isRecurring');
    const recurringControls = document.getElementById('recurringControls');
    if (isRecurringCheck) {
        isRecurringCheck.onchange = () => {
            if (recurringControls) {
                if (isRecurringCheck.checked) {
                    recurringControls.classList.add('active');
                } else {
                    recurringControls.classList.remove('active');
                }
            }
        };
    }
}

// --- ฟังก์ชันสำหรับการจัดการสลับหน้าเพจ (SPA Navigation) ---
function showPage(pageId) {
    document.querySelectorAll('.page-section, .auth-container').forEach(el => {
        el.classList.remove('active');
        // ล้างสไตล์ inline ที่ค้างอยู่เพื่อให้การควบคุมสลับหน้า SPA เป็นของ CSS Class ทั้งหมด
        el.style.display = '';
    });
    const target = document.getElementById(pageId);
    if (target) {
        target.classList.add('active');
    }
}

function switchWorkspaceView(viewType) {
    document.querySelectorAll('.work-tab').forEach(tab => tab.classList.remove('active'));
    document.querySelectorAll('.workspace-view').forEach(view => view.classList.remove('active'));

    if (viewType === 'form') {
        document.getElementById('tabFormBtn').classList.add('active');
        document.getElementById('formWorkspaceView').classList.add('active');
    } else if (viewType === 'table') {
        document.getElementById('tabTableBtn').classList.add('active');
        document.getElementById('tableWorkspaceView').classList.add('active');
        fetchRecentData(activeSheet);
    }
}

// --- การจัดการและเรนเดอร์ข้อมูลแผ่นงานแบบ Dynamic ใน Portal ---
async function loadDynamicSheets() {
    const grid = document.getElementById('portalSheetsGrid');
    if (!grid) return;

    try {
        const response = await fetch(scriptUrl, {
            method: 'POST',
            body: JSON.stringify({ action: 'getSheets' }),
            headers: { 'Content-Type': 'text/plain;charset=utf-8' }
        });
        const json = await response.json();

        if (json.status === 'success' && Array.isArray(json.data)) {
            grid.innerHTML = '';

            // ดึงข้อมูลจำนวนแผ่นงานและสถิติรวมของระบบมาวาด
            let totalRows = 0;
            json.data.forEach(s => {
                totalRows += s.rowCount;

                const card = document.createElement('div');
                const cleanNum = s.baseFormId || '031';
                card.className = `sheet-card form-${cleanNum}`;

                card.innerHTML = `
                    <div class="sheet-card-info">
                        <h3>FM สท.03-${cleanNum.substring(2)}</h3>
                        <p>ชื่อแผ่นงาน: <strong>${s.name}</strong></p>
                        <div class="sheet-stat">📊 มีข้อมูลทั้งหมด: ${s.rowCount} รายการ</div>
                    </div>
                    <div class="sheet-card-actions">
                        <button type="button" class="btn btn-primary btn-goto-form">📝 ลงทะเบียนฟอร์ม</button>
                        <button type="button" class="btn btn-outline btn-goto-table">📊 ดูตารางข้อมูล</button>
                    </div>
                `;

                // จัดการเมื่อผู้ใช้กดที่การ์ดหรือปุ่มในการ์ด
                card.querySelector('.btn-goto-form').onclick = (e) => {
                    e.stopPropagation();
                    enterWorkspace(s.name, 'form');
                };
                card.querySelector('.btn-goto-table').onclick = (e) => {
                    e.stopPropagation();
                    enterWorkspace(s.name, 'table');
                };
                card.onclick = () => {
                    enterWorkspace(s.name, 'form');
                };

                grid.appendChild(card);
            });

            // อัปเดตสถิติด้านบนของ Portal (หากต้องการเพิ่มแสดงภาพรวม)
            updateStatWidgets(json.data.length, totalRows);

        } else {
            throw new Error(json.message || 'ไม่สามารถโหลดรายชื่อแผ่นงานได้');
        }
    } catch (err) {
        grid.innerHTML = `<div style="grid-column:span 3;color:var(--error);text-align:center;padding:20px;">⚠️ โหลดข้อมูลล้มเหลว: ${err.message}</div>`;
    }
}

// อัปเดตสถิติภาพรวมในหน้า Portal
function updateStatWidgets(sheetCount, totalRows) {
    const existingStat = document.querySelector('.portal-stat-grid');
    if (existingStat) existingStat.remove();

    const portalHeader = document.querySelector('#portalPage header');
    const statGrid = document.createElement('div');
    statGrid.className = 'portal-stat-grid';
    statGrid.style.display = 'grid';
    statGrid.style.gridTemplateColumns = 'repeat(auto-fit, minmax(240px, 1fr))';
    statGrid.style.gap = '20px';
    statGrid.style.margin = '20px 0 35px 0';

    statGrid.innerHTML = `
        <div class="stat-card">
            <div class="stat-icon">📂</div>
            <div class="stat-info">
                <h4>แผ่นงานที่พร้อมใช้งาน</h4>
                <p>${sheetCount} แผ่นงาน</p>
            </div>
        </div>
        <div class="stat-card">
            <div class="stat-icon" style="background:var(--primary-light);">📝</div>
            <div class="stat-info">
                <h4>บันทึกข้อมูลรวมสะสม</h4>
                <p>${totalRows} รายการ</p>
            </div>
        </div>
        <div class="stat-card">
            <div class="stat-icon" style="background:rgba(245, 158, 11, 0.1);">⚡</div>
            <div class="stat-info">
                <h4>ความยืดหยุ่นของระบบ</h4>
                <p>Dynamic 100%</p>
            </div>
        </div>
    `;
    portalHeader.parentNode.insertBefore(statGrid, portalHeader.nextSibling);
}

// --- ฟังก์ชันเข้าสู่ Workspace ทำงานของแผ่นงานที่เลือก ---
async function enterWorkspace(sheetName, defaultView = 'form') {
    activeSheet = sheetName;

    // ตั้งชื่อแผ่นงานในหน้าทำงาน
    const activeSheetNameEl = document.getElementById('activeSheetName');
    if (activeSheetNameEl) activeSheetNameEl.textContent = sheetName;

    showPage('appWorkspace');
    cancelEditMode(); // เคลียร์ฟอร์มการแก้ไข

    // โหลด Schema โครงสร้างคอลัมน์ของแผ่นงานนั้น
    const success = await loadFormSchema(sheetName);
    if (success) {
        switchWorkspaceView(defaultView);
    }
}

// --- ฟังก์ชันดึง Schema ของแผ่นงานแบบ Async เพื่อหลีกเลี่ยง Race Condition ---
async function loadFormSchema(formId) {
    if (!scriptUrl || !formId) return false;
    showLoading('กำลังวิเคราะห์โครงสร้างแผ่นงาน...');
    try {
        const response = await fetch(scriptUrl, {
            method: 'POST',
            body: JSON.stringify({ action: 'getSchema', formId }),
            headers: { 'Content-Type': 'text/plain;charset=utf-8' }
        });
        const json = await response.json();
        if (json.status === 'success') {
            renderFields(json.data.headers, formId);
            hideLoading();
            return true;
        } else {
            throw new Error(json.message);
        }
    } catch (err) {
        hideLoading();
        showModal('❌ ผิดพลาด', 'ไม่สามารถเชื่อมโยงแบบฟอร์มได้: ' + err.message, false);
        return false;
    }
}

// เรนเดอร์ฟิลด์กรอกข้อมูลแบบ Dynamic ตามโครงสร้างหัวตาราง (Headers)
function renderFields(headers, formId) {
    currentHeaders = headers;
    const fieldsContainer = document.getElementById('fieldsContainer');
    fieldsContainer.innerHTML = '';

    // แปลงชื่อชีตเพื่อระบุประเภทแบบฟอร์ม (031, 033, 034, 035) สำหรับใช้คัดแยกเลย์เอาต์พิเศษ
    const cleanId = formId.toString().replace(/[-_\s]/g, '');
    let baseFormId = '031';
    if (cleanId.includes('031')) baseFormId = '031';
    else if (cleanId.includes('033')) baseFormId = '033';
    else if (cleanId.includes('034')) baseFormId = '034';
    else if (cleanId.includes('035')) baseFormId = '035';

    if (baseFormId === '034') {
        renderGrid034(headers);
        return;
    }

    const tableKeywords = ['ep', 'format', 'teach', 'dur', 'dma', 'item', 'qty', 'ตอน', 'รูปแบบสื่อ', 'วิทยากร', 'ความยาว', 'ผลิตแล้วเสร็จ'];
    const currentTableHeaders = headers.filter(h => tableKeywords.some(k => h.toLowerCase().includes(k)));
    const basicHeaders = headers.filter(h => !currentTableHeaders.includes(h));

    const cbHeaders = basicHeaders.filter(h => h.toLowerCase().startsWith('cb') || checkboxConfig[h]);
    const regularHeaders = basicHeaders.filter(h => !cbHeaders.includes(h));

    // เรนเดอร์ Checkbox (ถ้ามี)
    if (cbHeaders.length > 0) {
        const cbGroup = document.createElement('div');
        cbGroup.className = 'form-group full-width';
        cbGroup.style.background = 'rgba(79, 70, 229, 0.04)';
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

    // เรนเดอร์ Text Fields ทั่วไป
    let i = 0;
    while (i < regularHeaders.length) {
        const header = regularHeaders[i];
        const low = header.toLowerCase();

        // จัดกลุ่มสำหรับวันที่ (วัน/เดือน/พ.ศ.) ให้มาอยู่แถวเดียวกัน
        const isGroupStart = (low.startsWith('act') || low.startsWith('fin') || low.startsWith('suc') || low.startsWith('use') || (low === 'date' && regularHeaders[i + 1]?.toLowerCase() === 'mouth'));

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
            if (low.includes('detail') || low.includes('remark') || low.includes('note') || low.includes('team')) {
                group.classList.add('full-width');
            }
            renderInputGroup(group, header);
            fieldsContainer.appendChild(group);
            i++;
        }
    }

    // เรนเดอร์ส่วนตารางข้อมูลย่อยด้านล่าง (ถ้ามีฟิลด์จำพวก EP/รูปแบบสื่อ/วิทยากร)
    if (currentTableHeaders.length > 0) {
        renderTableRegistry(currentTableHeaders);
    }
}

// เรนเดอร์ตารางย่อย 11 แถวแบบคงที่สำหรับฟอร์ม 03-4
function renderGrid034(headers) {
    const tableKeywords = ['ep', 'teach', 'dur', 'dma', 'format', 'ตอน', 'วิทยากร', 'ความยาว', 'ผลิตแล้วเสร็จ', 'รูปแบบสื่อ'];
    const basicHeaders = headers.filter(h => !tableKeywords.some(k => h.toLowerCase().includes(k)));
    const cbHeaders = basicHeaders.filter(h => h.toLowerCase().startsWith('cb') || checkboxConfig[h]);
    const regularHeaders = basicHeaders.filter(h => !cbHeaders.includes(h));
    const fieldsContainer = document.getElementById('fieldsContainer');

    if (cbHeaders.length > 0) {
        const cbGroup = document.createElement('div');
        cbGroup.className = 'form-group full-width';
        cbGroup.style.background = 'rgba(79, 70, 229, 0.04)';
        cbGroup.style.padding = '20px';
        cbGroup.style.borderRadius = '15px';
        cbGroup.style.border = '1px solid var(--primary-light)';
        cbGroup.innerHTML = `<label style="font-weight:700; color:var(--primary); margin-bottom:15px; display:block;">🔹 หน่วยงานที่เกี่ยวข้อง / รูปแบบการจัดส่ง</label>`;

        const optionsContainer = document.createElement('div');
        optionsContainer.style.display = 'grid';
        optionsContainer.style.gridTemplateColumns = 'repeat(auto-fit, minmax(280px, 1fr))';
        optionsContainer.style.gap = '10px';

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
    gridContainer.className = 'table-batch-container';
    gridContainer.innerHTML = `
        <h3 class="batch-title">📝 รายการที่ผลิต (11 แถวแบบฟอร์มสรุปรวมกลุ่ม)</h3>
        <div class="batch-table-wrapper" style="margin-top: 15px;">
            <table class="batch-table">
                <thead>
                    <tr>
                        <th>รูปแบบสื่อ</th>
                        <th>ชื่อเรื่อง/ตอน</th>
                        <th>วิทยากร/ผู้บรรยาย</th>
                        <th>ความยาว (เช่น 1:12)</th>
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

// เรนเดอร์ส่วนแผงจัดเก็บข้อมูลตารางย่อย (สำหรับแบบฟอร์มอื่นๆ)
function renderTableRegistry(headers) {
    const fieldsContainer = document.getElementById('fieldsContainer');
    const container = document.createElement('div');
    container.className = 'table-batch-container';
    container.innerHTML = `
        <div class="batch-header">
            <div class="batch-title">📜 รายการรายละเอียดเพิ่มเติม (${headers.length} คอลัมน์)</div>
            <button type="button" class="btn-add" id="btnAddRow">➕ เพิ่มรายการใหม่</button>
        </div>
        <div class="batch-table-wrapper">
            <table class="batch-table">
                <thead>
                    <tr>
                        ${headers.map(h => `<th>${labelMap[h.toLowerCase()] || h}</th>`).join('')}
                        <th style="width:70px;text-align:center;">ลบ</th>
                    </tr>
                </thead>
                <tbody id="batchTableBody"></tbody>
            </table>
        </div>
    `;
    fieldsContainer.appendChild(container);
    container.querySelector('#btnAddRow').onclick = () => addBatchRow(headers);
    addBatchRow(headers); // เพิ่มแถวแรกเป็นค่าเริ่มต้น
}

function addBatchRow(headers, existingData = null) {
    const tbody = document.getElementById('batchTableBody');
    if (!tbody) return;
    const rowIndex = tbody.children.length;
    const tr = document.createElement('tr');
    tr.className = 'batch-row';

    headers.forEach(h => {
        const td = document.createElement('td');
        const low = h.toLowerCase();
        let input;

        if (low === 'dur' || low === 'ความยาว') {
            const grp = document.createElement('div');
            grp.className = 'unit-input-group';
            input = document.createElement('input');
            input.type = 'text';
            input.dataset.key = h;
            input.placeholder = '0.00';

            const sel = document.createElement('select');
            sel.className = 'unit-select';
            sel.dataset.unitFor = h;

            ['ชม.', 'น.'].forEach(u => {
                const o = document.createElement('option');
                o.value = u;
                o.textContent = u;
                sel.appendChild(o);
            });

            grp.appendChild(input);
            grp.appendChild(sel);

            if (existingData && existingData[h]) {
                const m = existingData[h].toString().match(/^([\d.]+)\s*(.*)$/);
                if (m) {
                    input.value = m[1];
                    sel.value = m[2] || 'ชม.';
                } else {
                    input.value = existingData[h];
                }
            }
            td.appendChild(grp);
        } else {
            input = document.createElement('input');
            input.type = 'text';
            input.dataset.key = h;
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
    delTd.style.textAlign = 'center';
    delTd.innerHTML = '<button type="button" class="btn-remove">🗑️</button>';
    delTd.querySelector('button').onclick = () => tr.remove();
    tr.appendChild(delTd);
    tbody.appendChild(tr);
}

// --- ฟังก์ชันรวบรวมข้อมูลเพื่อส่งประมวลผล ---
function collectTableData() {
    const cleanId = activeSheet.toString().replace(/[-_\s]/g, '');
    let baseFormId = '031';
    if (cleanId.includes('031')) baseFormId = '031';
    else if (cleanId.includes('033')) baseFormId = '033';
    else if (cleanId.includes('034')) baseFormId = '034';
    else if (cleanId.includes('035')) baseFormId = '035';

    if (baseFormId === '034') {
        const row = {};
        for (let r = 1; r <= 11; r++) {
            let f = document.getElementById(`input_format${r}`).value;
            let e = document.getElementById(`input_ep${r}`).value;
            let t = document.getElementById(`input_teach${r}`).value;
            let d = document.getElementById(`input_dur${r}`).value;
            let a = document.getElementById(`input_dma${r}`).value;

            if (d && !d.includes('ชม.')) d = `${d} ชม.`;

            row[`format${r}`] = f;
            row[`ep${r}`] = e;
            row[`teach${r}`] = t;
            row[`dur${r}`] = d;
            row[`dma${r}`] = a;

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

// --- การเรียกส่งข้อมูลไปยัง Apps Script API ---
async function sendData(action) {
    if (isProcessing) return;
    const isRecurring = document.getElementById('isRecurring').checked && action === 'generate';
    const rounds = isRecurring ? (parseInt(document.getElementById('recurringCount').value) || 1) : 1;

    isProcessing = true;
    const resultBox = document.getElementById('resultBox');
    const linksContainer = document.getElementById('linksContainer');
    resultBox.style.display = 'none';
    linksContainer.innerHTML = '';

    showLoading('กำลังเริ่มประมวลผลข้อมูล...');

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
        const interval = parseInt(document.getElementById('intervalDays').value) || 7;
        const subjectName = baseData.subject || baseData['ชื่อรายการที่ผลิต'] || 'ไม่ระบุชื่อ';

        showToast('🚀 เริ่มส่งข้อมูล', `กำลังจัดทำ PDF สำหรับโครงการ: ${subjectName}`, 'info');
        setTimeout(() => hideLoading(), 800);

        for (let r = 0; r < rounds; r++) {
            const currentData = { ...baseData };
            if (r > 0) {
                Object.keys(currentData).forEach(key => {
                    const lowKey = key.toLowerCase();
                    if (lowKey.endsWith('date')) {
                        const prefix = lowKey.substring(0, lowKey.length - 4);
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

            // ส่งข้อมูลไปหลังบ้าน
            fetch(scriptUrl, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify({ action, formId: activeSheet, data: currentData, tableData, rowIndex: currentEditRowIndex })
            })
                .then(res => res.json())
                .then(result => {
                    if (result.status === 'success') {
                        addResultLink(result.data.url, result.data.name);
                        showToast('✅ สร้างสำเร็จ', `สร้าง PDF: ${result.data.name} เรียบร้อยแล้ว`, 'success');
                        resultBox.style.display = 'block';

                        if (rounds === 1) {
                            cancelEditMode();
                        }
                    } else {
                        showToast('❌ ผิดพลาด', result.message, 'error');
                    }
                })
                .catch(err => {
                    showToast('❌ เกิดข้อผิดพลาด', err.message, 'error');
                })
                .finally(() => {
                    if (r === rounds - 1) isProcessing = false;
                });
        }

        showToast('💡 ส่งคำขอแล้ว', 'คุณสามารถสลับหน้าจอหรือรอระบบแจ้งการสร้าง PDF สำเร็จได้ครับ', 'info', 8000);

    } catch (e) {
        hideLoading();
        isProcessing = false;
        showModal('❌ ข้อผิดพลาด', e.message, false);
    }
}

// --- ฟังก์ชันดึงประวัติข้อมูลล่าสุดมาแสดงผลในตาราง (Dashboard) ---
async function fetchRecentData(targetFormId) {
    if (!scriptUrl || !targetFormId) return;

    showLoading(`กำลังดึงประวัติข้อมูลแผ่นงาน ${targetFormId}...`);
    try {
        const response = await fetch(scriptUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({ action: 'getRecentData', formId: targetFormId, limit: 100 })
        });
        const json = await response.json();
        if (json.status === 'success') {
            dashboardData = json.data.records;
            renderDashTable(dashboardData, json.data.headers);
        } else {
            throw new Error(json.message);
        }
    } catch (e) {
        showModal('❌ ผิดพลาด', 'ไม่สามารถอ่านประวัติข้อมูลได้: ' + e.message, false);
    } finally {
        hideLoading();
    }
}

// เรนเดอร์ข้อมูลแถวลงในหน้าจอจัดการตาราง
function renderDashTable(records, headers = []) {
    const body = document.getElementById('dashBody');
    const headerRow = document.getElementById('dashHeaderRow');
    body.innerHTML = '';

    let displayHeaders = (headers && headers.length > 0) ? [...headers] : [];

    if (records.length > 0) {
        records.forEach(r => {
            const rawTable = r.tableData || r.tableBody || r._tableData;
            if (rawTable) {
                try {
                    const parsed = typeof rawTable === 'string' ? JSON.parse(rawTable) : rawTable;
                    if (Array.isArray(parsed) && parsed.length > 0) {
                        Object.keys(parsed[0]).forEach(k => {
                            if (!displayHeaders.includes(k) && k.toLowerCase() !== 'id') {
                                displayHeaders.push(k);
                            }
                        });
                    }
                } catch (e) { }
            }
        });
    }

    if (displayHeaders.length === 0 && records.length > 0) {
        displayHeaders = Object.keys(records[0]).filter(k => k !== '_rowIndex');
    }

    const exclude = ['tabledata', 'tablebody', '_tabledata', 'id', 'createdat'];
    displayHeaders = displayHeaders.filter(h => h && h.toString().trim() !== "" && !exclude.includes(h.toLowerCase())).slice(0, 100);

    headerRow.innerHTML = `<th>ลำดับ</th>`;
    displayHeaders.forEach(h => {
        headerRow.innerHTML += `<th>${h}</th>`;
    });
    headerRow.innerHTML += `<th style="width: 200px; position: sticky; right: 0; background: var(--primary-light); z-index: 10;">การทำงาน</th>`;

    if (records.length === 0) {
        body.innerHTML = `<tr><td colspan="${displayHeaders.length + 2}" style="text-align:center;padding:50px;">📭 ไม่พบรายการบันทึกข้อมูลย้อนหลังในชีตนี้</td></tr>`;
        return;
    }

    // เรนเดอร์ข้อมูลแบบแก้ไขแบบ Inline
    records.forEach((r, idx) => {
        const tr = document.createElement('tr');
        tr.dataset.rowIndex = r._rowIndex;

        let rowHtml = `<td>${idx + 1}</td>`;
        displayHeaders.forEach(h => {
            const headerStr = h.toString();
            let val = r[headerStr] || "";

            // แปลงค่าวันที่ ISO ให้เป็น พ.ศ. 
            if (typeof val === 'string' && val.includes('T') && val.endsWith('Z')) {
                try {
                    const d = new Date(val);
                    if (!isNaN(d.getTime())) {
                        if (d.getFullYear() < 1905) {
                            const hh = d.getHours();
                            const mm = d.getMinutes();
                            const ss = d.getSeconds();
                            val = hh + ":" + (mm < 10 ? "0" + mm : mm) + (ss > 0 ? ":" + (ss < 10 ? "0" + ss : ss) : "");
                        } else {
                            const day = d.getDate().toString().padStart(2, '0');
                            const month = (d.getMonth() + 1).toString().padStart(2, '0');
                            const year = d.getFullYear();
                            const beYear = year < 2400 ? year + 543 : year;
                            val = `${day}/${month}/${beYear}`;
                        }
                    }
                } catch (e) { }
            }

            const lowerH = headerStr.toLowerCase();
            if ((lowerH.includes('dur') || lowerH.includes('ความยาว')) && val.toString().startsWith('0:')) {
                val = val.toString().substring(2);
            }
            if (val === "0:00" || val === "00:00") val = "-";

            const charCount = Math.max(headerStr.length, val.toString().length);
            let widthNum = (charCount * 11) + 30;
            widthNum = Math.min(Math.max(widthNum, 60), 450);

            const width = widthNum + 'px';
            const textAlign = charCount <= 3 ? 'center' : 'left';

            rowHtml += `<td><input type="text" class="dash-inline-input" data-field="${h}" value="${val}" style="width: ${width}; min-width: ${width}; text-align: ${textAlign};"></td>`;
        });

        rowHtml += `
            <td style="position: sticky; right: 0; background: #fff; box-shadow: -5px 0 10px rgba(0,0,0,0.05);">
                <div class="action-btns">
                    <button type="button" class="btn-action btn-save" onclick="window.saveRecordInline(${idx})" title="บันทึกแก้ไขแถวนี้ลงชีต">💾</button>
                    <button type="button" class="btn-action btn-edit" onclick="window.editRecord(${idx})" title="ดึงข้อมูลกลับมาแก้ไขบนฟอร์ม">✍️</button>
                    <button type="button" class="btn-action btn-pdf" onclick="window.generateRecordPDF(${idx})" title="สร้าง PDF รายการเดียว">📄</button>
                    <button type="button" class="btn-action btn-del" onclick="window.deleteRecord(${idx})" title="ลบข้อมูล">🗑️</button>
                </div>
            </td>
        `;
        tr.innerHTML = rowHtml;
        body.appendChild(tr);
    });

    // อัปเดตผูกปุ่มสรุปและจัดการคอลัมน์ด้านบน
    updateTableControlPanel();
}

function updateTableControlPanel() {
    const btnTool = document.getElementById('dashToolBtns');
    if (!btnTool) return;

    ['btnSaveAllWorkspace', 'btnGlobalPDFWorkspace'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.remove();
    });

    // ปุ่มเซฟการเปลี่ยนแปลงอินไลน์ทั้งหมด
    const saveAllBtn = document.createElement('button');
    saveAllBtn.type = 'button';
    saveAllBtn.id = 'btnSaveAllWorkspace';
    saveAllBtn.className = 'btn-save-all';
    saveAllBtn.style.backgroundColor = '#8b5cf6';
    saveAllBtn.style.color = 'white';
    saveAllBtn.innerHTML = '💾 บันทึกทั้งหมด';
    saveAllBtn.onclick = () => saveAllRecordsInline();
    btnTool.appendChild(saveAllBtn);

    // ปุ่มสร้าง PDF รวบยอด (มีผลเฉพาะของ 03-4 หรือฟอร์มที่ระบุตารางย่อย)
    const cleanId = activeSheet.toString().replace(/[-_\s]/g, '');
    if (cleanId.includes('034')) {
        const globalBtn = document.createElement('button');
        globalBtn.type = 'button';
        globalBtn.id = 'btnGlobalPDFWorkspace';
        globalBtn.className = 'btn-pdf-all';
        globalBtn.style.backgroundColor = '#f59e0b';
        globalBtn.innerHTML = `📄 สร้าง PDF รวบยอด`;
        globalBtn.onclick = () => generateBatchPDF034();
        btnTool.appendChild(globalBtn);
    }
}

// --- ฟังก์ชันบันทึก แก้ไข และลบข้อมูลแถวผ่านตาราง ---
async function saveRecordInline(idx) {
    const r = dashboardData[idx];
    const tr = document.querySelector(`#dashBody tr:nth-child(${idx + 1})`);

    const inputs = tr.querySelectorAll('.dash-inline-input');
    const updatedData = { ...r };
    inputs.forEach(input => {
        const field = input.dataset.field;
        updatedData[field] = input.value;
        if (field === 'subject') updatedData['ชื่อรายการที่ผลิต'] = input.value;
        if (field === 'ep') updatedData['ตอน'] = input.value;
        if (field === 'teach') updatedData['วิทยากร/ผู้บรรยาย'] = input.value;
        if (field === 'dma') updatedData['วันผลิตแล้วเสร็จ'] = input.value;
    });

    showLoading('กำลังแก้ไขข้อมูลลงแผ่นงาน...');
    try {
        const response = await fetch(scriptUrl, {
            method: 'POST',
            body: JSON.stringify({
                action: 'updateRow',
                formId: activeSheet,
                rowIndex: r._rowIndex,
                data: updatedData
            })
        });
        const json = await response.json();
        if (json.status === 'success') {
            dashboardData[idx] = updatedData;
            showModal('✅ สำเร็จ', `ปรับปรุงข้อมูลแถวที่ ${idx + 1} เรียบร้อยแล้ว`, false);
        } else {
            throw new Error(json.message);
        }
    } catch (e) {
        showModal('❌ ผิดพลาด', e.message, false);
    } finally {
        hideLoading();
    }
}

async function saveAllRecordsInline() {
    const rows = document.querySelectorAll('#dashBody tr');
    const updates = [];

    rows.forEach((tr, idx) => {
        const inputs = tr.querySelectorAll('.dash-inline-input');
        if (inputs.length === 0) return;

        const originalData = dashboardData[idx];
        const updatedData = { ...originalData };
        let hasChanged = false;

        inputs.forEach(input => {
            const field = input.dataset.field;
            if (updatedData[field] !== input.value) {
                updatedData[field] = input.value;
                hasChanged = true;
                if (field === 'subject') updatedData['ชื่อรายการที่ผลิต'] = input.value;
                if (field === 'ep') updatedData['ตอน'] = input.value;
                if (field === 'teach') updatedData['วิทยากร/ผู้บรรยาย'] = input.value;
                if (field === 'dma') updatedData['วันผลิตแล้วเสร็จ'] = input.value;
            }
        });

        if (hasChanged) {
            updates.push({
                rowIndex: originalData._rowIndex,
                data: updatedData,
                localIndex: idx
            });
        }
    });

    if (updates.length === 0) {
        showModal('ℹ️ แจ้งเตือน', 'ไม่พบการเปลี่ยนแปลงใดๆ ในตาราง', false);
        return;
    }

    if (!await showModal('💾 บันทึกทั้งหมด', `คุณต้องการเซฟข้อมูลที่แก้ไขทั้ง ${updates.length} รายการใช่หรือไม่?`, true)) return;

    showLoading(`กำลังอัปเดตข้อมูล ${updates.length} รายการ...`);
    try {
        const response = await fetch(scriptUrl, {
            method: 'POST',
            body: JSON.stringify({
                action: 'batchUpdateRowsInSheet', // ตรงกับฟังก์ชันหลังบ้านใน Apps Script
                formId: activeSheet,
                updates: updates
            })
        });
        const json = await response.json();
        if (json.status === 'success') {
            updates.forEach(u => {
                dashboardData[u.localIndex] = u.data;
            });
            showModal('✅ บันทึกสำเร็จ', `ปรับปรุงข้อมูลครบถ้วน ${updates.length} แถวเรียบร้อยแล้ว`, false);
            fetchRecentData(activeSheet);
        } else {
            throw new Error(json.message);
        }
    } catch (e) {
        showModal('❌ ผิดพลาด', e.message, false);
    } finally {
        hideLoading();
    }
}

async function editRecord(idx) {
    const r = dashboardData[idx];

    // ตั้งค่าแถวการแก้ไขและเปิดแบนเนอร์แสดงสถานะ Edit
    currentEditRowIndex = r._rowIndex;
    const banner = document.getElementById('editModeBanner');
    const rowNumSpan = document.getElementById('editRowNumber');
    if (banner && rowNumSpan) {
        banner.style.display = 'flex';
        rowNumSpan.textContent = r._rowIndex;
    }

    const fmt = (v) => {
        if (!v) return '';
        if (typeof v === 'string' && v.includes('T') && v.includes('Z') && v.length > 15) {
            try {
                const d = new Date(v);
                if (!isNaN(d.getTime())) {
                    if (d.getFullYear() < 1905) {
                        const hh = d.getHours();
                        const mm = d.getMinutes();
                        const ss = d.getSeconds();
                        return hh + ":" + (mm < 10 ? "0" + mm : mm) + (ss > 0 ? ":" + (ss < 10 ? "0" + ss : ss) : "");
                    }
                    const day = d.getDate().toString().padStart(2, '0');
                    const month = (d.getMonth() + 1).toString().padStart(2, '0');
                    const year = d.getFullYear();
                    const beYear = year < 2400 ? year + 543 : year;
                    return `${day}/${month}/${beYear}`;
                }
            } catch (e) { }
        }
        return v;
    };

    restoreBaseData(r, fmt);

    const cleanId = activeSheet.toString().replace(/[-_\s]/g, '');
    let baseFormId = '031';
    if (cleanId.includes('031')) baseFormId = '031';
    else if (cleanId.includes('033')) baseFormId = '033';
    else if (cleanId.includes('034')) baseFormId = '034';
    else if (cleanId.includes('035')) baseFormId = '035';

    if (baseFormId === '034') {
        const subject = (r.subject || r['ชื่อรายการที่ผลิต'] || "").toString().toLowerCase().trim();
        const related = dashboardData.filter(row => {
            const rowSub = (row.subject || row['ชื่อรายการที่ผลิต'] || "").toString().toLowerCase().trim();
            return rowSub === subject && subject !== "";
        }).sort((a, b) => {
            const epA = parseInt((a.ep || a['ตอน'] || "0").toString().replace(/\D/g, '')) || 0;
            const epB = parseInt((b.ep || b['ตอน'] || "0").toString().replace(/\D/g, '')) || 0;
            return epA - epB;
        });

        for (let r = 1; r <= 11; r++) {
            const data = related[r - 1] || {};
            const f = document.getElementById(`input_format${r}`);
            const e = document.getElementById(`input_ep${r}`);
            const t = document.getElementById(`input_teach${r}`);
            const d = document.getElementById(`input_dur${r}`);
            const a = document.getElementById(`input_dma${r}`);

            if (f) f.value = fmt(data.format || data['รูปแบบสื่อ'] || r[`format${r}`] || r[`รูปแบบสื่อ${r}`]);
            if (e) e.value = fmt(data.ep || data['ตอน'] || r[`ep${r}`] || r[`ตอน${r}`]);
            if (t) t.value = fmt(data.teach || data['วิทยากร/ผู้บรรยาย'] || data['วิทยากร'] || r[`teach${r}`] || r[`วิทยากร${r}`]);
            if (d) d.value = fmt(data.dur || data['ความยาว'] || data['ความยาวรายการ (นาที)'] || r[`dur${r}`] || r[`ความยาว${r}`]).toString().replace(' ชม.', '');
            if (a) a.value = fmt(data.dma || data.sucdate || data['วันผลิตแล้วเสร็จ'] || data['ผลิตแล้วเสร็จวันที่'] || r[`dma${r}`] || r[`sucdate${r}`] || r[`วันผลิตแล้วเสร็จ${r}`] || r[`ผลิตแล้วเสร็จวันที่${r}`]);
        }
    } else {
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
                    } catch (e) { addBatchRow(tableHeaders, r); }
                } else { addBatchRow(tableHeaders, r); }
            }
        }
    }

    switchWorkspaceView('form');
    showToast('✏️ โหมดแก้ไข', `ดึงข้อมูลของโครงการ "${r.subject || 'ไม่ระบุชื่อ'}" เข้าสู่ฟอร์มเรียบร้อยแล้ว`, 'info');
}

function restoreBaseData(r, formatFn = null) {
    currentHeaders.forEach(h => {
        let val = r[h] || '';
        if (formatFn) val = formatFn(val);

        const checkboxes = document.querySelectorAll(`input[name="${h}"]`);
        if (checkboxes.length > 0) {
            checkboxes.forEach(cb => {
                cb.checked = Array.isArray(val) ? val.includes(cb.value) : (val === cb.value || val === '✓');
            });
        } else {
            const el = document.getElementById(`input_${h}`);
            if (el) el.value = val;
        }
    });
}

async function deleteRecord(idx) {
    const r = dashboardData[idx];
    if (await showModal('🗑️ ยืนยันการลบ', 'คุณต้องการลบรายการข้อมูลนี้ออกจากแผ่นงานใช่หรือไม่?', true)) {
        showLoading('กำลังลบข้อมูลออกจาก Sheet...');
        try {
            const response = await fetch(scriptUrl, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify({ action: 'deleteData', rowIndex: r._rowIndex })
            });
            const json = await response.json();
            if (json.status === 'success') {
                showToast('✅ ลบข้อมูลแล้ว', 'ลบแถวข้อมูลเรียบร้อยแล้วครับ', 'success');
                fetchRecentData(activeSheet);
            } else {
                throw new Error(json.message);
            }
        } catch (e) {
            showModal('❌ ผิดพลาด', e.message, false);
        } finally {
            hideLoading();
        }
    }
}

async function generateRecordPDF(idx) {
    const r = dashboardData[idx];
    let tableData = null;
    let mainData = { ...r };

    const cleanId = activeSheet.toString().replace(/[-_\s]/g, '');
    if (cleanId.includes('034')) {
        showLoading('กำลังดึงแถวรวบยอดข้อมูลโครงการ...');
        const subjectToMatch = (r.subject || r['ชื่อรายการที่ผลิต'] || "").toString().toLowerCase().trim();

        tableData = dashboardData.filter(row => {
            const rowSubject = (row.subject || row['ชื่อรายการที่ผลิต'] || "").toString().toLowerCase().trim();
            return rowSubject === subjectToMatch && subjectToMatch !== "";
        });

        tableData.sort((a, b) => {
            const epA = parseInt((a.ep || a['ตอน'] || "0").toString().replace(/\D/g, '')) || 0;
            const epB = parseInt((b.ep || b['ตอน'] || "0").toString().replace(/\D/g, '')) || 0;
            return epA - epB;
        });

        if (tableData.length === 0) tableData = [r];
    }

    showLoading(`กำลังประมวลผลการสร้างเอกสาร PDF...`);
    try {
        const response = await fetch(scriptUrl, {
            method: 'POST',
            body: JSON.stringify({
                action: 'generate',
                formId: activeSheet,
                data: mainData,
                tableData: tableData,
                rowIndex: r._rowIndex,
                skipSave: true
            })
        });
        const result = await response.json();
        if (result.status === 'success') {
            showModal('📜 สร้างไฟล์สำเร็จ',
                `สร้าง PDF แผ่นงาน ${activeSheet} เรียบร้อยแล้วครับ<br><br>` +
                `<a href="${result.data.url}" target="_blank" style="display:inline-block; padding:12px 24px; background:#4f46e5; color:white; border-radius:10px; text-decoration:none; font-weight:bold; box-shadow:0 4px 10px rgba(79,70,229,0.2);">🌐 เปิดดูไฟล์เอกสาร PDF</a>`,
                false);
        } else {
            throw new Error(result.message);
        }
    } catch (e) {
        showModal('❌ ผิดพลาด', e.message, false);
    } finally {
        hideLoading();
    }
}

async function generateBatchPDF034() {
    try {
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
            showModal('⚠️ คำเตือน', 'ไม่พบข้อมูลในตารางสำหรับรวมยอด PDF', false);
            return;
        }

        showLoading(`กำลังสร้าง PDF สรุปรวมแบบกลุ่ม (${tableData.length} รายการ)...`);
        const response = await fetch(scriptUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({
                action: 'generate',
                formId: activeSheet,
                data: tableData[0],
                tableData,
                rowIndex: 0,
                skipSave: true
            })
        });
        const result = await response.json();
        if (result.status === 'success') {
            showModal('📜 สร้างสรุปสำเร็จ', `เสร็จเรียบร้อยแล้ว`, false, '📄', () => {
                if (result.data.url) window.open(result.data.url, '_blank');
            });
            if (result.data.url) window.open(result.data.url, '_blank');
        } else {
            throw new Error(result.message);
        }
    } catch (e) {
        showModal('❌ ผิดพลาด', e.message, false);
    } finally {
        hideLoading();
    }
}

async function generateAllPDFs() {
    if (!dashboardData || dashboardData.length === 0) {
        showModal('⚠️ คำเตือน', 'ไม่มีข้อมูลให้สร้างไฟล์ PDF', false);
        return;
    }

    const confirmed = await showModal('📄 ยืนยันสร้างทั้งหมด',
        `คุณต้องการเริ่มคำสั่งสร้าง PDF จากรายการข้อมูลทั้งหมด ${dashboardData.length} รายการในหน้านี้ใช่หรือไม่? (ระบบจะประมวลผลเบื้องหลัง)`, true);
    if (!confirmed) return;

    showToast('🚀 เริ่มสร้าง PDF ทั้งหมด', `ระบบกำลังทยอยทำ PDF ทั้งหมด ${dashboardData.length} รายการ...`, 'info', 8000);

    // ย้ายไปทำงานแบบ Asynchronous ใน Background
    (async () => {
        let successCount = 0;
        let errorCount = 0;

        for (let i = 0; i < dashboardData.length; i++) {
            const r = dashboardData[i];
            try {
                let tableData = [];
                if (r.tableData || r.tableBody || r._tableData) {
                    try { tableData = JSON.parse(r.tableData || r.tableBody || r._tableData); }
                    catch (e) { tableData = [r]; }
                } else {
                    tableData = [r];
                }

                const response = await fetch(scriptUrl, {
                    method: 'POST',
                    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                    body: JSON.stringify({
                        action: 'generate',
                        formId: activeSheet,
                        data: r,
                        tableData,
                        rowIndex: r._rowIndex,
                        skipSave: true
                    })
                });
                const result = await response.json();
                if (result.status === 'success') {
                    successCount++;
                    if (successCount % 5 === 0 || i === dashboardData.length - 1) {
                        showToast('⏳ ความคืบหน้า', `บันทึกเสร็จแล้ว ${successCount}/${dashboardData.length} รายการ`, 'info');
                    }
                } else {
                    errorCount++;
                }
            } catch (err) {
                errorCount++;
            }
        }
        showToast('🏁 ดำเนินการสร้างครบแล้ว', `สร้างสำเร็จ ${successCount} ไฟล์ ${errorCount > 0 ? `(ผิดพลาด ${errorCount} รายการ)` : ''}`, successCount > 0 ? 'success' : 'error', 15000);
    })();
}

// --- ฟังก์ชันช่วยเหลือและคำนวณวันเวลา ---
function calcNextDate(day, monthThai, yearBE, daysToAdd) {
    const dayArabic = toArabicDigits(day.toString());
    const yearArabic = toArabicDigits(yearBE.toString());
    const mIdx = thaiMonths.indexOf(monthThai.trim());
    if (mIdx === -1) return { day, month: monthThai, year: yearBE };
    const yearAD = parseInt(yearArabic) - 543;
    const date = new Date(yearAD, mIdx, parseInt(dayArabic));
    date.setDate(date.getDate() + daysToAdd);

    const hasThaiDigits = (/[๐-๙]/.test(day.toString()) || /[๐-๙]/.test(yearBE.toString()));
    return {
        day: hasThaiDigits ? toThaiDigits(date.getDate()) : date.getDate(),
        month: thaiMonths[date.getMonth()],
        year: hasThaiDigits ? toThaiDigits(date.getFullYear() + 543) : date.getFullYear() + 543
    };
}

function toThaiDigits(num) { return num.toString().replace(/[0-9]/g, digit => "๐๑๒๓๔๕๖๗๘๙"[digit]); }
function toArabicDigits(str) { return str.toString().replace(/[๐-๙]/g, digit => "๐๑๒๓๔๕๖๗๘๙".indexOf(digit)); }

// --- จัดการปุ่มยกเลิกและเคลียร์ข้อมูลฟอร์ม ---
function cancelEditMode() {
    currentEditRowIndex = 0;
    const banner = document.getElementById('editModeBanner');
    if (banner) banner.style.display = 'none';
}

function clearForm() {
    const fieldsContainer = document.getElementById('fieldsContainer');

    // ล้าง Input Text และ Textarea
    const inputs = fieldsContainer.querySelectorAll('input[type="text"], input[type="number"], textarea');
    inputs.forEach(inp => inp.value = '');

    // ล้าง Checkbox
    const checkboxes = fieldsContainer.querySelectorAll('input[type="checkbox"]');
    checkboxes.forEach(cb => cb.checked = false);

    // ล้างตารางรายละเอียด (ถ้ามี)
    const tbody = document.getElementById('batchTableBody');
    if (tbody) tbody.innerHTML = '';

    // รีเซ็ตการแก้ไข
    cancelEditMode();

    // ซ่อนกล่องดาวน์โหลด PDF เก่า
    document.getElementById('resultBox').style.display = 'none';

    showToast('✨ ล้างข้อมูลฟอร์ม', 'ล้างข้อมูลหน้าแบบฟอร์มเดิมเรียบร้อยแล้ว', 'success');
}

// --- ฟังก์ชันค้นหาและฟิลเตอร์ในตาราง Dashboard ---
window.filterDashboard = function () {
    const query = document.getElementById('dashSearch').value.toLowerCase();
    const filtered = dashboardData.filter(r => {
        const title = (r.subject || r['ชื่อรายการที่ผลิต'] || r.activity || '').toLowerCase();
        const owner = (r.owner || r['ผู้รับผิดชอบการผลิต'] || '').toLowerCase();
        return title.includes(query) || owner.includes(query);
    });
    renderDashTable(filtered);
};

// --- ดึงรายการ PDF จาก Google Drive เพื่อนำมารวมไฟล์ (Drive Picker) ---
async function openDriveFilePicker() {
    showLoading('กำลังเชื่อมต่อเพื่ออ่านรายการไฟล์บน Drive...');
    try {
        const response = await fetch(scriptUrl, {
            method: 'POST',
            body: JSON.stringify({ action: 'getDriveFiles', formId: activeSheet })
        });
        const json = await response.json();
        hideLoading();

        if (json.status === 'success' && json.data.length > 0) {
            let html = '<div class="drive-list-container">';
            html += `
                <div class="select-all-header">
                    <input type="checkbox" id="driveSelectAll">
                    <label for="driveSelectAll" style="cursor:pointer;margin-left:8px;">เลือกทั้งหมด (${json.data.length} ไฟล์)</label>
                </div>
            `;

            json.data.forEach(f => {
                html += `
                <label class="drive-file-item">
                    <input type="checkbox" class="drive-file-cb" value="${f.id}" data-name="${f.name}"> 
                    <div style="margin-left:12px;">
                        <div style="font-weight:600; color:#1e293b;">${f.name}</div>
                        <div style="font-size:0.75rem; color:#64748b;">สร้างเมื่อ: ${f.date}</div>
                    </div>
                </label>`;
            });
            html += '</div>';

            // รอรับคำยืนยัน
            const confirmed = await showModal('🗂️ เลือกเอกสารบน Drive เพื่อนำมารวมไฟล์', html, true);

            // การจัดการ Select All ภายหลังการแสดง DOM
            const selectAll = document.getElementById('driveSelectAll');
            if (selectAll) {
                selectAll.onchange = () => {
                    document.querySelectorAll('.drive-file-cb').forEach(cb => cb.checked = selectAll.checked);
                };
            }

            if (confirmed) {
                const selected = Array.from(document.querySelectorAll('.drive-file-cb:checked')).map(cb => cb.value);
                if (selected.length > 0) {
                    mergeAllGeneratedPDFs(selected);
                } else {
                    showToast('⚠️ คำเตือน', 'ไม่ได้เลือกไฟล์ใด ๆ', 'info');
                }
            }
        } else {
            showModal('📁 แจ้งเตือน', json.message || 'ไม่พบไฟล์ PDF ในโฟลเดอร์สำหรับแผ่นงานนี้ครับ', false);
        }
    } catch (e) {
        hideLoading();
        showModal('❌ ผิดพลาด', e.message, false);
    }
}

// --- ฟังก์ชันการดาวน์โหลดและรวมไฟล์ PDF ผ่าน pdf-lib ---
async function mergeAllGeneratedPDFs(fileIds) {
    showLoading('กำลังเริ่มดาวน์โหลดและรวมไฟล์ PDF ทั้งหมด... (กรุณารอสักครู่)');
    try {
        const { PDFDocument } = window.PDFLib;
        const mergedPdf = await PDFDocument.create();

        for (let i = 0; i < fileIds.length; i++) {
            updateLoadingText(`กำลังรวมไฟล์รายการที่ ${i + 1}/${fileIds.length}...`);

            // ดึงไฟล์ base64 ผ่าน Apps Script
            const base64 = await callBackend('getFileBytes', { fileId: fileIds[i] });
            const bytes = Uint8Array.from(atob(base64), c => c.charCodeAt(0)).buffer;

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

        showModal('✨ รวมไฟล์สำเร็จ', 'ระบบได้ทำการดาวน์โหลดไฟล์ PDF รวมรายการของท่านเรียบร้อยแล้ว', false);
    } catch (e) {
        showModal('❌ เกิดข้อผิดพลาด', 'ล้มเหลวในการรวมไฟล์ PDF: ' + e.message, false);
    } finally {
        hideLoading();
    }
}

async function callBackend(action, params) {
    const response = await fetch(scriptUrl, {
        method: 'POST',
        body: JSON.stringify({ action, ...params })
    });
    const json = await response.json();
    if (json.status === 'success') return json.data;
    throw new Error(json.message);
}

function updateLoadingText(text) {
    const desc = document.getElementById('loadingText');
    if (desc) desc.innerText = text;
}

// --- ระบบกล่องยืนยันการอนุมัติและแจ้งเตือน (Custom Modal) ---
function showModal(title, message, isConfirm = false, icon = '🔔', onConfirm = null) {
    return new Promise((resolve) => {
        const overlay = document.getElementById('modalOverlay');
        document.getElementById('modalIcon').textContent = icon;
        document.getElementById('modalTitle').textContent = title;

        const desc = document.getElementById('modalDesc');
        if (typeof message === 'string') {
            desc.innerHTML = message;
        } else {
            desc.innerHTML = '';
            desc.appendChild(message);
        }

        const btnCancel = document.getElementById('modalCancel');
        const btnConfirm = document.getElementById('modalConfirm');

        btnCancel.style.display = isConfirm ? 'block' : 'none';
        btnConfirm.textContent = isConfirm ? 'ยืนยัน' : 'ตกลง';
        overlay.classList.add('active');

        btnConfirm.onclick = () => {
            overlay.classList.remove('active');
            if (onConfirm) onConfirm();
            resolve(true);
        };
        btnCancel.onclick = () => {
            overlay.classList.remove('active');
            resolve(false);
        };
    });
}

function showLoading(msg) {
    document.getElementById('loadingOverlay').classList.add('active');
    document.getElementById('loadingText').textContent = msg;
}

function hideLoading() {
    document.getElementById('loadingOverlay').classList.remove('active');
}

// Toast
function showToast(title, msg, type = 'info', duration = 5000) {
    const container = document.getElementById('toastContainer');
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;

    let icon = '🔔';
    if (type === 'success') icon = '✅';
    if (type === 'error') icon = '❌';
    if (type === 'info') icon = 'ℹ️';

    toast.innerHTML = `
        <div class="toast-icon" style="font-size: 1.3rem;">${icon}</div>
        <div class="toast-content" style="margin-left: 10px;">
            <div class="toast-title">${title}</div>
            <div class="toast-msg">${msg}</div>
        </div>
        <div class="toast-close" onclick="this.parentElement.remove()">×</div>
    `;

    container.appendChild(toast);

    setTimeout(() => {
        toast.classList.add('fade-out');
        setTimeout(() => toast.remove(), 400);
    }, duration);
}

function renderInputGroup(container, h) {
    const low = h.toLowerCase();
    const group = document.createElement('div');
    group.className = 'form-group';

    const label = document.createElement('label');
    label.textContent = labelMap[low] || h;

    const input = document.createElement('input');
    input.id = `input_${h}`;
    input.placeholder = label.textContent;

    container.appendChild(label);
    container.appendChild(input);
}

function renderSingleCheckbox(container, name, labelText, value) {
    const wrapper = document.createElement('label');
    wrapper.style.display = 'flex';
    wrapper.style.alignItems = 'center';
    wrapper.style.gap = '10px';
    wrapper.style.cursor = 'pointer';
    wrapper.style.padding = '8px 12px';
    wrapper.style.background = 'white';
    wrapper.style.borderRadius = '10px';
    wrapper.style.border = '1px solid #e2e8f0';

    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.name = name;
    cb.value = value;
    cb.style.width = '18px';
    cb.style.height = '18px';
    cb.style.accentColor = 'var(--primary)';

    wrapper.appendChild(cb);
    wrapper.appendChild(document.createTextNode(labelText));
    container.appendChild(wrapper);
}

function addResultLink(url, name) {
    const container = document.getElementById('linksContainer');
    const div = document.createElement('div');
    div.className = 'batch-link';
    div.innerHTML = `<a href="${url}" target="_blank" style="text-decoration:none;color:inherit;display:block;width:100%;">📄 เปิดดูไฟล์เอกสาร: ${name}</a>`;
    container.appendChild(div);
}

// ผูกฟังก์ชันเหล่านี้กับ window เพื่อให้ onClick บนปุ่มตารางสามารถเรียกใช้ได้
window.saveRecordInline = saveRecordInline;
window.editRecord = editRecord;
window.generateRecordPDF = generateRecordPDF;
window.deleteRecord = deleteRecord;
