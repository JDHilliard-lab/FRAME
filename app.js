// =========================================================================
// GLOBAL APP STATE & ICONS
// =========================================================================
let currentView = 'dashboard';
let dashUnit = 'in';
const emptyImgUrl = 'data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7';
let dashActiveImageObj = new Image(); 
dashActiveImageObj.src = emptyImgUrl;
let dashSelectedRowIndex = 0;
let dashTempHoverUrl = null;
let dashLocalLibrary = {}; 

// Minimalist SVGs
const svgMove = `<svg class="svg-icon" viewBox="0 0 24 24"><path d="M5 9l-3 3 3 3M9 5l3-3 3 3M19 9l3 3-3 3M9 19l3 3-3 3M2 12h20M12 2v20"/></svg>`;
const svgEdit = `<svg class="svg-icon" viewBox="0 0 24 24"><path d="M12 20h9M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"/></svg>`;
const svgDup = `<svg class="svg-icon" viewBox="0 0 24 24"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>`;
const svgTrash = `<svg class="svg-icon" viewBox="0 0 24 24"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>`;

const dashDefaultData = { 
    id: "GR.ART-01", imageCode: "TBD", level: "1", qty: 0, product: "Framed Art", location: "LOBBY", 
    bleed: 0.25, canvasDepth: "", canvasWrap: "",
    extW: 24, extH: 30, fType: "image", fW: 1.25, fColor: "#1a1a1a", fCode: "Upload or Sync Library", 
    swatchDataUrl: "", swatchName: "",
    m1A: true, m1T: 3, m1B: 3, m1L: 3, m1R: 3, m1Locked: false, mColor: "White Mat",
    m2A: true, m2: 0.25, 
    glass: "2mm Standard", hardware: "3-Point Security", mount: "Standard Mount", backing: "Foamcore", notes: "", prodNotes: "" 
};
let dashProjectData = [ JSON.parse(JSON.stringify(dashDefaultData)) ];

let warnedLinkedFrames = new Set(); 

let elevUnit = 'in';
let elevations = [{ name: "Elevation 1", frames: [], wallW: 185, wallH: 108, personPos: { x: -60 } }];
let currentElevIndex = 0;
let elevFrames = elevations[0].frames;
let elevPersonPos = elevations[0].personPos;
let elevScale = 1;
let elevZoomFactor = 1;

// =========================================================================
// INITIALIZATION & NAVIGATION
// =========================================================================
function initMasterApp() {
    document.getElementById('g_date').valueAsDate = new Date();
    renderNavTabs();
    selectDashRow(0); 
    populateDashPushSelector();
    
    document.addEventListener('click', function(event) {
        const container = document.getElementById('customSwatchContainer');
        const sList = document.getElementById('swatchDropdownList');
        if (container && !container.contains(event.target)) {
            if (sList && sList.style.display === 'block') { sList.style.display = 'none'; restoreDashThumbnail(); }
        }
        const bList = document.getElementById('bulkDropdownList');
        const bBtn = document.querySelector('[onclick="toggleBulkDropdown()"]');
        if (bList && bList.style.display === 'block' && event.target !== bBtn && !bBtn.contains(event.target) && !bList.contains(event.target)) {
            bList.style.display = 'none';
        }
    });
}

function toggleTheme() {
    document.body.classList.toggle('light-theme');
}

function renderNavTabs() {
    const container = document.getElementById('nav-tabs-container');
    let html = `<div class="nav-tab ${currentView==='dashboard'?'active':''}" onclick="switchView('dashboard')">Frame Dashboard</div><div class="tab-divider"></div>`;
    
    elevations.forEach((elev, idx) => {
        let isActive = (currentView === 'elevation' && currentElevIndex === idx) ? 'active' : '';
        html += `<div class="nav-tab ${isActive}" onclick="switchView('elevation', ${idx})">
                    <span>${elev.name}</span>
                    <span class="tab-close" onclick="deleteElevation(${idx}, event)" title="Delete Wall">×</span>
                 </div>`;
    });
    container.innerHTML = html;
}

function updateElevationNameFromInput(newName) {
    if (newName.trim() !== "") {
        elevations[currentElevIndex].name = newName;
        renderNavTabs();
        populateDashPushSelector();
    }
}

function deleteElevation(idx, e) {
    e.stopPropagation();
    if(confirm("Delete this entire elevation wall? This cannot be undone.")) {
        elevations.splice(idx, 1);
        if (elevations.length === 0) {
            let w = elevUnit === 'cm' ? parseFloat((185 * 2.54).toFixed(2)) : 185;
            let h = elevUnit === 'cm' ? parseFloat((108 * 2.54).toFixed(2)) : 108;
            let px = elevUnit === 'cm' ? parseFloat((-60 * 2.54).toFixed(2)) : -60;
            elevations.push({ name: "Elevation 1", frames: [], wallW: w, wallH: h, personPos: {x: px} });
            currentElevIndex = 0;
        } else if (currentElevIndex > idx) {
            currentElevIndex--;
        }
        if (currentElevIndex === idx || elevations.length === 1) switchView('dashboard');
        
        renderNavTabs();
        populateDashPushSelector();
        recalculateDashboardQuantities();
    }
}

function addNewElevationTab() {
    let newIndex = elevations.length;
    let w = elevUnit === 'cm' ? parseFloat((185 * 2.54).toFixed(2)) : 185;
    let h = elevUnit === 'cm' ? parseFloat((108 * 2.54).toFixed(2)) : 108;
    let px = elevUnit === 'cm' ? parseFloat((-60 * 2.54).toFixed(2)) : -60;
    
    elevations.push({ name: "Elevation " + (newIndex + 1), frames: [], wallW: w, wallH: h, personPos: {x: px} });
    renderNavTabs();
    populateDashPushSelector();
    switchView('elevation', newIndex);
}

function switchView(viewType, index = 0) {
    if (currentView === 'elevation' && elevations[currentElevIndex]) {
        elevations[currentElevIndex].wallW = parseFloat(document.getElementById('wallW').value) || 185;
        elevations[currentElevIndex].wallH = parseFloat(document.getElementById('wallH').value) || 108;
    }

    if (viewType === 'dashboard') {
        document.getElementById('view-dashboard').classList.add('active');
        document.getElementById('view-elevation').classList.remove('active');
        currentView = 'dashboard';
        recalculateDashboardQuantities(); 
    } else {
        document.getElementById('view-dashboard').classList.remove('active');
        document.getElementById('view-elevation').classList.add('active');
        currentView = 'elevation';
        currentElevIndex = index;
        
        let elev = elevations[currentElevIndex];
        document.getElementById('elev-title-input').value = elev.name;
        document.getElementById('wallW').value = elev.wallW;
        document.getElementById('wallH').value = elev.wallH;
        elevFrames = elev.frames;
        elevPersonPos = elev.personPos;
        
        populateElevBulkList();
        initElevControls();
        drawElevAll();
    }
    renderNavTabs();
}

function saveMasterProject() {
    if(currentView === 'elevation' && elevations[currentElevIndex]) {
        elevations[currentElevIndex].wallW = parseFloat(document.getElementById('wallW').value) || 185;
        elevations[currentElevIndex].wallH = parseFloat(document.getElementById('wallH').value) || 108;
    }
    const getStr = (id) => document.getElementById(id).value;
    const globalMeta = { projName: getStr('g_projName'), desc: getStr('g_desc'), date: getStr('g_date'), issued: getStr('g_issued'), client: getStr('g_client'), attn: getStr('g_attn'), delivery: getStr('g_delivery') };
    const masterData = { type: 'master-studio-v3', dashUnit: dashUnit, elevUnit: elevUnit, globalMeta: globalMeta, dashProjectData: dashProjectData, elevations: elevations };
    const blob = new Blob([JSON.stringify(masterData, null, 2)], { type: 'application/json' });
    const link = document.createElement('a'); link.href = URL.createObjectURL(blob); link.download = `Master_Studio_Project.json`; link.click();
}

function loadMasterProject(event) {
    const file = event.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = JSON.parse(e.target.result);
            if (data.type && data.type.startsWith('master-studio')) {
                dashUnit = data.dashUnit || 'in'; elevUnit = data.elevUnit || 'in';
                if (data.globalMeta) {
                    const setVal = (id, val) => { const el = document.getElementById(id); if(el) el.value = val; };
                    setVal('g_projName', data.globalMeta.projName); setVal('g_desc', data.globalMeta.desc); setVal('g_date', data.globalMeta.date);
                    setVal('g_issued', data.globalMeta.issued); setVal('g_client', data.globalMeta.client); setVal('g_attn', data.globalMeta.attn); setVal('g_delivery', data.globalMeta.delivery);
                }
                if (data.dashProjectData) dashProjectData = data.dashProjectData;
                if (data.elevations) elevations = data.elevations;
            } else { return alert("Invalid format. Please build a new project in Master Studio."); }

            document.getElementById('dashBtnInch').classList.toggle('active', dashUnit === 'in'); document.getElementById('dashBtnCm').classList.toggle('active', dashUnit === 'cm');
            document.getElementById('elevBtnInch').classList.toggle('active', elevUnit === 'in'); document.getElementById('elevBtnCm').classList.toggle('active', elevUnit === 'cm');
            
            recalculateDashboardQuantities(); selectDashRow(0); renderNavTabs(); switchView('dashboard'); 
        } catch (err) { alert("Invalid project file."); }
    };
    reader.readAsText(file); event.target.value = '';
}

// =========================================================================
// THE BRIDGE: DROPDOWN BULK CHECKBOXES & DASHBOARD PUSH
// =========================================================================
function recalculateDashboardQuantities() {
    let counts = {};
    elevations.forEach(elev => {
        elev.frames.forEach(f => { if (f.active && f.id) counts[f.id] = (counts[f.id] || 0) + 1; });
    });
    
    dashProjectData.forEach(d => { d.qty = counts[d.id] !== undefined ? counts[d.id] : 0; });
    
    if (currentView === 'dashboard') {
        loadDashDataIntoControls(dashProjectData[dashSelectedRowIndex]);
        renderDashTable();
    }
}

function toggleBulkDropdown() {
    const list = document.getElementById('bulkDropdownList');
    list.style.display = list.style.display === 'none' ? 'block' : 'none';
}

function populateElevBulkList() {
    const container = document.getElementById('bulkImportCheckboxes');
    if(!container) return;
    container.innerHTML = '';
    dashProjectData.forEach((f, idx) => {
        container.innerHTML += `
            <label style="display: flex; align-items: center; gap: 6px; padding: 6px; border-bottom: 1px solid var(--border-color); cursor: pointer; text-transform: none; color: var(--text-strong); font-size: 0.75rem; margin:0;">
                <input type="checkbox" class="bulk-import-cb" value="${idx}">
                <b>${f.id}</b> <span style="color:var(--text-muted);">(${f.product})</span>
            </label>
        `;
    });
}

function importSelectedFramesBulk() {
    const cbs = document.querySelectorAll('.bulk-import-cb:checked');
    if(cbs.length === 0) return alert("Select at least one frame!");
    
    let factor = (dashUnit === 'cm' && elevUnit === 'in') ? (1/2.54) : (dashUnit === 'in' && elevUnit === 'cm') ? 2.54 : 1;
    let startX = 10;
    if (elevFrames.length > 0) {
        let maxRight = 0;
        elevFrames.forEach(fr => { if (fr.x + fr.w > maxRight) maxRight = fr.x + fr.w; });
        startX = maxRight + 10;
    }

    cbs.forEach(cb => {
        const f = dashProjectData[parseInt(cb.value)];
        const newFrame = {
            id: f.id, letter: getElevLetter(elevFrames.length),
            w: (parseFloat(f.extW) || 24) * factor, h: (parseFloat(f.extH) || 30) * factor,
            fW: (parseFloat(f.fW) || 1.25) * factor, fType: f.fType || 'color', fColor: f.fColor || '#1a1a1a', fCode: f.fCode || '', swatchDataUrl: f.swatchDataUrl || '',
            m1T: (parseFloat(f.m1T) || 0) * factor, m1B: (parseFloat(f.m1B) || 0) * factor, m1L: (parseFloat(f.m1L) || 0) * factor, m1R: (parseFloat(f.m1R) || 0) * factor,
            m1Locked: f.m1Locked || false, m2: (parseFloat(f.m2) || 0) * factor, m2Active: f.m2A || false,
            x: startX, y: 10, isOpen: false, isGrouped: false, dimTo: [], active: true
        };
        elevFrames.push(newFrame);
        startX += (newFrame.w + 5);
        cb.checked = false; 
    });
    
    document.getElementById('bulkDropdownList').style.display = 'none';
    initElevControls(); drawElevAll(); recalculateDashboardQuantities();
}

function populateDashPushSelector() {
    const select = document.getElementById('dashPushSelector');
    if(!select) return;
    select.innerHTML = '<option value="">-- Push to Wall --</option>';
    elevations.forEach((e, idx) => {
        const opt = document.createElement('option'); opt.value = idx; opt.textContent = e.name; select.appendChild(opt);
    });
}

function pushFrameToElevation() {
    const select = document.getElementById('dashPushSelector');
    if (select.value === "") return alert("Select an elevation first!");
    const eIdx = parseInt(select.value);
    const targetElev = elevations[eIdx];
    const f = dashProjectData[dashSelectedRowIndex];
    
    let factor = (dashUnit === 'cm' && elevUnit === 'in') ? (1/2.54) : (dashUnit === 'in' && elevUnit === 'cm') ? 2.54 : 1;
    let startX = 10;
    if (targetElev.frames.length > 0) {
        let maxRight = 0;
        targetElev.frames.forEach(fr => { if (fr.x + fr.w > maxRight) maxRight = fr.x + fr.w; });
        startX = maxRight + 10;
    }

    targetElev.frames.push({
        id: f.id, letter: getElevLetter(targetElev.frames.length),
        w: (parseFloat(f.extW) || 24) * factor, h: (parseFloat(f.extH) || 30) * factor,
        fW: (parseFloat(f.fW) || 1.25) * factor, fType: f.fType || 'color', fColor: f.fColor || '#1a1a1a', fCode: f.fCode || '', swatchDataUrl: f.swatchDataUrl || '',
        m1T: (parseFloat(f.m1T) || 0) * factor, m1B: (parseFloat(f.m1B) || 0) * factor, m1L: (parseFloat(f.m1L) || 0) * factor, m1R: (parseFloat(f.m1R) || 0) * factor,
        m1Locked: f.m1Locked || false, m2: (parseFloat(f.m2) || 0) * factor, m2Active: f.m2A || false,
        x: startX, y: 10, isOpen: false, isGrouped: false, dimTo: [], active: true
    });
    
    recalculateDashboardQuantities();
    alert(`Pushed ${f.id} to ${targetElev.name}!`);
}

function jumpToDashboard(frameId) {
    const targetIdx = dashProjectData.findIndex(d => d.id === frameId);
    if (targetIdx !== -1) { switchView('dashboard'); selectDashRow(targetIdx); }
}

function checkGlobalEditingWarning(id) {
    if (warnedLinkedFrames.has(id)) return true;

    let count = 0;
    elevations.forEach(e => { e.frames.forEach(f => { if(f.id === id) count++; }); });
    
    const banner = document.getElementById('linkedWarningBanner');
    if (count > 0) {
        document.getElementById('linkedCount').innerText = count;
        banner.style.display = 'flex';
        warnedLinkedFrames.add(id); 
    } else {
        banner.style.display = 'none';
    }
    return true;
}

function pushUpdatesToElevations(dashIndex) {
    const d = dashProjectData[dashIndex];
    let factor = (dashUnit === 'cm' && elevUnit === 'in') ? (1/2.54) : (dashUnit === 'in' && elevUnit === 'cm') ? 2.54 : 1;

    elevations.forEach(elev => {
        elev.frames.forEach(f => {
            if (f.id === d.id) {
                f.w = (parseFloat(d.extW) || 24) * factor; f.h = (parseFloat(d.extH) || 30) * factor;
                f.fW = (parseFloat(d.fW) || 1.25) * factor; f.fType = d.fType; f.fColor = d.fColor; f.swatchDataUrl = d.swatchDataUrl;
                f.m1T = (parseFloat(d.m1T) || 0) * factor; f.m1B = (parseFloat(d.m1B) || 0) * factor; f.m1L = (parseFloat(d.m1L) || 0) * factor; f.m1R = (parseFloat(d.m1R) || 0) * factor;
                f.m2 = (parseFloat(d.m2) || 0) * factor; f.m2Active = d.m2A;
            }
        });
    });
}

// =========================================================================
// DASHBOARD LOGIC
// =========================================================================
const dashFmt = (num) => {
    let n = Number(num);
    if (num === "" || num === undefined || num === null || isNaN(n)) return 0; 
    return parseFloat(n.toFixed(3)); 
};

function toggleDashSection(id, btn) {
    const sec = document.getElementById(id);
    const span = btn.querySelector('span');
    if (sec.classList.contains('open')) {
        sec.classList.remove('open'); span.innerHTML = `<svg class="svg-icon" style="width:10px; height:10px; transform:rotate(-90deg);" viewBox="0 0 24 24"><polyline points="6 9 12 15 18 9"></polyline></svg>`;
    } else {
        sec.classList.add('open'); span.innerHTML = `<svg class="svg-icon" style="width:10px; height:10px;" viewBox="0 0 24 24"><polyline points="6 9 12 15 18 9"></polyline></svg>`;
    }
}

function setDashUnit(newUnit) {
    if(dashUnit === newUnit) return;
    const factor = newUnit === 'cm' ? 2.54 : (1/2.54);
    dashProjectData.forEach(row => {
        ['extW', 'extH', 'fW', 'bleed', 'canvasDepth', 'canvasWrap', 'm1T', 'm1B', 'm1L', 'm1R', 'm2'].forEach(prop => {
            if (row[prop] !== "" && row[prop] !== undefined && !isNaN(row[prop])) row[prop] = dashFmt(row[prop] * factor);
        });
    });
    dashUnit = newUnit;
    document.getElementById('dashBtnInch').classList.toggle('active', dashUnit === 'in');
    document.getElementById('dashBtnCm').classList.toggle('active', dashUnit === 'cm');
    loadDashDataIntoControls(dashProjectData[dashSelectedRowIndex]); 
}

function selectDashRow(index) {
    if (index >= dashProjectData.length) return; 
    dashSelectedRowIndex = index;
    loadDashDataIntoControls(dashProjectData[index]);
    document.querySelectorAll('#rfiBody tr').forEach((tr, i) => { tr.classList.toggle('selected', i === index); });
    checkGlobalEditingWarning(dashProjectData[index].id);
}

function addDashRow() {
    const newRow = JSON.parse(JSON.stringify(dashProjectData[dashSelectedRowIndex])); 
    newRow.id = newRow.id + "-NEW"; newRow.qty = 0; 
    dashProjectData.push(newRow);
    dashSelectedRowIndex = dashProjectData.length - 1;
    loadDashDataIntoControls(dashProjectData[dashSelectedRowIndex]);
    renderDashTable(); 
    checkGlobalEditingWarning(newRow.id);
}

function deleteDashRow() {
    if(dashProjectData.length <= 1) return alert("Cannot delete the last row.");
    const idToDelete = dashProjectData[dashSelectedRowIndex].id;
    dashProjectData.splice(dashSelectedRowIndex, 1);
    dashSelectedRowIndex = Math.max(0, dashSelectedRowIndex - 1);
    
    elevations.forEach(elev => { elev.frames = elev.frames.filter(f => f.id !== idToDelete); });
    
    loadDashDataIntoControls(dashProjectData[dashSelectedRowIndex]);
    renderDashTable();
    checkGlobalEditingWarning(dashProjectData[dashSelectedRowIndex].id);
}

function loadDashDataIntoControls(data) {
    const setVal = (id, val) => { const el = document.getElementById(id); if(el) el.value = val; };
    setVal('m_itemCode', data.id); setVal('m_imageCode', data.imageCode); setVal('m_level', data.level); setVal('m_qty', data.qty);
    setVal('m_product', data.product); setVal('m_location', data.location); setVal('m_bleed', dashFmt(data.bleed)); setVal('canvasDepth', dashFmt(data.canvasDepth)); setVal('canvasWrap', dashFmt(data.canvasWrap));
    setVal('extW', dashFmt(data.extW)); setVal('extH', dashFmt(data.extH)); setVal('fType', data.fType); setVal('fW', dashFmt(data.fW)); 
    setVal('fColor', data.fColor); setVal('m_fCode', data.fCode); setVal('m_color', data.mColor);
    setVal('m_glass', data.glass); setVal('m_hardware', data.hardware); setVal('m_mount', data.mount); setVal('m_backing', data.backing); setVal('m_notes', data.notes); setVal('m_prodNotes', data.prodNotes);

    document.getElementById('m1Toggle').classList.toggle('active', data.m1A); document.getElementById('m1Toggle').innerText = data.m1A ? 'ON' : 'OFF';
    document.querySelectorAll('.m1-input').forEach(el => el.disabled = !data.m1A);
    document.getElementById('m2Toggle').classList.toggle('active', data.m2A); document.getElementById('m2Toggle').innerText = data.m2A ? 'ON' : 'OFF';
    document.getElementById('m2').disabled = !data.m2A;
    document.getElementById('m1Lock').classList.toggle('active', data.m1Locked); document.getElementById('m1Lock').innerText = data.m1Locked ? 'LOCKED' : 'UNLOCKED';
    
    setVal('m1T', dashFmt(data.m1T)); setVal('m1B', dashFmt(data.m1B)); setVal('m1L', dashFmt(data.m1L)); setVal('m1R', dashFmt(data.m1R)); setVal('m2', dashFmt(data.m2));

    handleDashProductChange(false);
    document.getElementById('swatchSelectedDisplay').textContent = (data.fType === 'image' && data.swatchName) ? data.swatchName : 'Frame';

    if(data.swatchDataUrl) { dashActiveImageObj.src = data.swatchDataUrl; document.getElementById('swatchThumbPreview').style.backgroundImage = `url(${data.swatchDataUrl})`; } 
    else { dashActiveImageObj.src = emptyImgUrl; document.getElementById('swatchThumbPreview').style.backgroundImage = `none`; }
    updateDashVisualsFromDOM();
}

function dashHtIn(idx, field, val) {
    let row = dashProjectData[idx];
    if(['qty','extW','extH','fW','m1T','m1R','m1B','m1L','m2','m_bleed','canvasDepth','canvasWrap'].includes(field)) val = parseFloat(val) || 0;
    if (field === 'id') { const oldId = row.id; elevations.forEach(elev => { elev.frames.forEach(f => { if (f.id === oldId) f.id = val; }); }); }
    row[field] = val;

    if (idx === dashSelectedRowIndex) {
        const map = { 'id':'m_itemCode', 'imageCode':'m_imageCode', 'level':'m_level', 'qty':'m_qty', 'location':'m_location', 'extW':'extW', 'extH':'extH', 'fCode':'m_fCode', 'fW':'fW', 'canvasDepth':'canvasDepth', 'canvasWrap':'canvasWrap', 'mColor':'m_color', 'm1T':'m1T', 'm1R':'m1R', 'm1B':'m1B', 'm1L':'m1L', 'm2':'m2', 'glass':'m_glass', 'hardware':'m_hardware', 'backing':'m_backing', 'mount':'m_mount', 'notes':'m_notes', 'prodNotes':'m_prodNotes' };
        if(map[field] && document.getElementById(map[field])) document.getElementById(map[field]).value = dashFmt(row[field]);
        if(field === 'product') { document.getElementById('m_product').value = row.product; handleDashProductChange(false); }
        updateDashVisualsFromDOM();
        pushUpdatesToElevations(idx);
    }
    renderDashTable(); 
    if (field === 'id') recalculateDashboardQuantities();
}

function syncDashAndCalculate() {
    const getRaw = (id) => { const el = document.getElementById(id); return el ? el.value : ""; };
    const getVal = (id) => parseFloat(getRaw(id)) || 0;
    const getStr = (id) => getRaw(id);
    const row = dashProjectData[dashSelectedRowIndex];
    
    const oldId = row.id; const newId = getStr('m_itemCode');
    if (oldId !== newId) { elevations.forEach(elev => { elev.frames.forEach(f => { if (f.id === oldId) f.id = newId; }); }); }

    dashProjectData[dashSelectedRowIndex] = {
        id: newId, imageCode: getStr('m_imageCode'), level: getStr('m_level'), qty: getVal('m_qty'), product: getStr('m_product'), location: getStr('m_location'),
        bleed: getVal('m_bleed'), canvasDepth: getRaw('canvasDepth'), canvasWrap: getRaw('canvasWrap'),
        extW: getVal('extW'), extH: getVal('extH'), fType: getStr('fType'), fW: getVal('fW'), fColor: getStr('fColor'), fCode: getStr('m_fCode'),
        swatchDataUrl: row.swatchDataUrl, swatchName: row.swatchName,
        m1A: document.getElementById('m1Toggle').classList.contains('active'), 
        m1T: getVal('m1T'), m1B: getVal('m1B'), m1L: getVal('m1L'), m1R: getVal('m1R'), m1Locked: document.getElementById('m1Lock').classList.contains('active'), mColor: getStr('m_color'),
        m2A: document.getElementById('m2Toggle').classList.contains('active'), m2: getVal('m2'),
        glass: getStr('m_glass'), hardware: getStr('m_hardware'), mount: getStr('m_mount'), backing: getStr('m_backing'), notes: getStr('m_notes'), prodNotes: getStr('m_prodNotes')
    };
    
    updateDashVisualsFromDOM(); renderDashTable(); pushUpdatesToElevations(dashSelectedRowIndex);
    if (oldId !== newId) recalculateDashboardQuantities();
}

function updateDashVisualsFromDOM() {
    const data = dashProjectData[dashSelectedRowIndex];
    const fVis = document.getElementById('dash-frame-visual');
    const viewObj = document.getElementById('view-dashboard');
    
    if (data.swatchDataUrl && data.fType === 'image') { viewObj.style.setProperty('--frame-bg', `url(${data.swatchDataUrl})`); } 
    else { viewObj.style.setProperty('--frame-bg', `none`); }
    if (data.fType === 'color') { document.getElementById('colorControls').style.display = 'flex'; document.getElementById('imageControls').style.display = 'none'; fVis.className = 'frame-vis frame-vis-solid'; } 
    else { document.getElementById('colorControls').style.display = 'none'; document.getElementById('imageControls').style.display = 'flex'; fVis.className = 'frame-vis frame-vis-image'; }

    const isCanvas = (data.product === "Framed Canvas (Floater)");
    const m1T = (data.m1A && !isCanvas) ? data.m1T : 0; const m1B = (data.m1A && !isCanvas) ? data.m1B : 0; const m1L = (data.m1A && !isCanvas) ? data.m1L : 0; const m1R = (data.m1A && !isCanvas) ? data.m1R : 0;
    const m2 = (data.m2A && !isCanvas) ? data.m2 : 0;

    const finalW = data.extW - (data.fW * 2) - m1L - m1R - (m2 * 2); const finalH = data.extH - (data.fW * 2) - m1T - m1B - (m2 * 2);
    
    document.getElementById('disp_openW').innerText = dashFmt(Math.max(0, finalW));
    document.getElementById('disp_openH').innerText = dashFmt(Math.max(0, finalH));
    document.getElementById('printFileDisplay').innerText = `${dashFmt(Math.max(0, finalW) + (data.bleed * 2))} x ${dashFmt(Math.max(0, finalH) + (data.bleed * 2))}`;

    const ratio = 300 / Math.max(data.extW, data.extH); 
    const m1Vis = document.getElementById('dash-mat1-visual'); const m2Vis = document.getElementById('dash-mat2-visual'); const artVis = document.getElementById('dash-art-visual');

    viewObj.style.setProperty('--fW', (data.fW * ratio) + 'px'); viewObj.style.setProperty('--frame-W', (data.extW * ratio) + 'px'); viewObj.style.setProperty('--frame-color', data.fColor);
    fVis.style.width = (data.extW * ratio) + "px"; fVis.style.height = (data.extH * ratio) + "px";

    if(data.m1A && !isCanvas) { 
        m1Vis.style.display = 'block'; m1Vis.style.top = (data.fW * ratio) + "px"; m1Vis.style.left = (data.fW * ratio) + "px"; 
        m1Vis.style.width = ((data.extW - (data.fW * 2)) * ratio) + "px"; m1Vis.style.height = ((data.extH - (data.fW * 2)) * ratio) + "px"; 
        m1Vis.style.borderTopWidth = (data.m1T * ratio) + 'px'; m1Vis.style.borderBottomWidth = (data.m1B * ratio) + 'px';
        m1Vis.style.borderLeftWidth = (data.m1L * ratio) + 'px'; m1Vis.style.borderRightWidth = (data.m1R * ratio) + 'px';
        m1Vis.style.borderColor = '#ffffff'; 
    } else { m1Vis.style.display = 'none'; }
    
    if(data.m2A && !isCanvas) { 
        m2Vis.style.display = 'block'; m2Vis.style.top = ((data.fW + m1T) * ratio) + "px"; m2Vis.style.left = ((data.fW + m1L) * ratio) + "px"; 
        m2Vis.style.width = ((data.extW - (data.fW * 2) - m1L - m1R) * ratio) + "px"; m2Vis.style.height = ((data.extH - (data.fW * 2) - m1T - m1B) * ratio) + "px"; 
        m2Vis.style.borderWidth = (data.m2 * ratio) + 'px'; m2Vis.style.borderColor = '#fdfdfd'; 
    } else { m2Vis.style.display = 'none'; }
    
    artVis.style.border = isCanvas ? "4px solid #111" : "1px solid #aaa";
    artVis.style.top = ((data.fW + m1T + m2) * ratio) + "px"; artVis.style.left = ((data.fW + m1L + m2) * ratio) + "px";
    artVis.style.width = (Math.max(0, finalW) * ratio) + "px"; artVis.style.height = (Math.max(0, finalH) * ratio) + "px";
    artVis.innerText = `${dashFmt(Math.max(0, finalW))}${dashUnit === 'in' ? '"' : ' cm'} × ${dashFmt(Math.max(0, finalH))}${dashUnit === 'in' ? '"' : ' cm'}`;
}

// RESTORED SYNC & DB FUNCTIONS
function handleDashProductChange(shouldSync = true) {
    const val = document.getElementById('m_product').value;
    const matWrapper = document.getElementById('matWrapper');
    const bleedSettings = document.getElementById('bleedSettings');
    const canvasSettings = document.getElementById('canvasSettings');

    if(val === "Framed Canvas (Floater)") {
        matWrapper.classList.add('disabled-section'); bleedSettings.style.display = 'none'; canvasSettings.style.display = 'grid';
    } else {
        matWrapper.classList.remove('disabled-section'); bleedSettings.style.display = 'grid'; canvasSettings.style.display = 'none';
    }
    if(shouldSync) syncDashAndCalculate();
}

function toggleDashMat(id) {
    const b = document.getElementById(id + 'Toggle');
    b.classList.toggle('active');
    b.innerText = b.classList.contains('active') ? 'ON' : 'OFF';
    if(id === 'm1') document.querySelectorAll('.m1-input').forEach(e => e.disabled = !b.classList.contains('active'));
    else document.getElementById('m2').disabled = !b.classList.contains('active');
    syncDashAndCalculate();
}

function toggleDashLock() {
    const b = document.getElementById('m1Lock');
    b.classList.toggle('active');
    b.innerText = b.classList.contains('active') ? 'LOCKED' : 'UNLOCKED';
    handleDashMatSync('m1T');
}

function handleDashMatSync(id) {
    if (document.getElementById('m1Lock').classList.contains('active')) {
        const v = document.getElementById(id).value;
        ['m1T','m1B','m1L','m1R'].forEach(x => document.getElementById(x).value = v);
    }
    syncDashAndCalculate();
}

function restoreDashThumbnail() {
    const d = dashProjectData[dashSelectedRowIndex];
    const t = document.getElementById('swatchThumbPreview');
    if (d.fType === 'image' && d.swatchDataUrl) t.style.backgroundImage = `url(${d.swatchDataUrl})`;
    else t.style.backgroundImage = 'none';
    if(dashTempHoverUrl) { URL.revokeObjectURL(dashTempHoverUrl); dashTempHoverUrl = null; }
}

// ROBUST FOLDER SYNC LOGIC
function syncDashLibraryFolder(e) {
    const f = e.target.files;
    if(!f || f.length === 0) return;
    dashLocalLibrary = {}; let c = 0;
    
    for(let file of f) {
        const ext = file.name.split('.').pop().toLowerCase();
        // Fallback for strict OS's that drop mime types on local folder reads
        const isImage = file.type.startsWith('image/') || ['jpg', 'jpeg', 'png', 'webp', 'svg'].includes(ext);
        if(!isImage) continue;
        
        const parts = file.webkitRelativePath.split('/');
        const filename = parts[parts.length - 1]; 
        const vendor = parts.length > 2 ? parts[parts.length - 3] : (parts.length > 1 ? parts[parts.length - 2] : parts[0]);
        const collection = parts.length > 2 ? parts[parts.length - 2] : "General";
        
        const baseName = filename.substring(0, filename.lastIndexOf('.'));
        const nP = baseName.split('_');
        let w = 1.25;
        let code = baseName;
        
        if(nP.length > 1 && !isNaN(parseFloat(nP[nP.length - 1]))) {
            w = parseFloat(nP[nP.length - 1]);
            code = nP.slice(0, -1).join(' ');
        }
        
        if(!dashLocalLibrary[vendor]) dashLocalLibrary[vendor] = {};
        if(!dashLocalLibrary[vendor][collection]) dashLocalLibrary[vendor][collection] = [];
        dashLocalLibrary[vendor][collection].push({ code: code, width: w, file: file });
        c++;
    }
    populateDashVendorDropdown(); alert(`Successfully synced ${c} swatches from Frame library!`);
}

function populateDashVendorDropdown() {
    const v = document.getElementById('libVendor'); v.innerHTML = '<option value="">Vendor</option>';
    Object.keys(dashLocalLibrary).sort().forEach(k => v.innerHTML += `<option value="${k}">${k}</option>`);
}

function updateDashCollectionDropdown() {
    const v = document.getElementById('libVendor').value; const c = document.getElementById('libCollection');
    c.innerHTML = '<option value="">Collection</option>';
    if(!v || !dashLocalLibrary[v]) return;
    Object.keys(dashLocalLibrary[v]).sort().forEach(k => c.innerHTML += `<option value="${k}">${k}</option>`);
    if(c.options.length === 2) { c.selectedIndex = 1; updateDashCustomSwatchDropdown(); }
}

function updateDashCustomSwatchDropdown() {
    const v = document.getElementById('libVendor').value; const c = document.getElementById('libCollection').value;
    const s = document.getElementById('swatchDropdownList'); s.innerHTML = '';
    if(!v || !c || !dashLocalLibrary[v][c]) return;
    dashLocalLibrary[v][c].forEach((i, idx) => {
        const li = document.createElement('li'); li.textContent = `${i.code} (${i.width}")`;
        li.onmouseenter = () => {
            if(dashTempHoverUrl) URL.revokeObjectURL(dashTempHoverUrl);
            dashTempHoverUrl = URL.createObjectURL(i.file);
            document.getElementById('swatchThumbPreview').style.backgroundImage = `url(${dashTempHoverUrl})`;
        };
        li.onclick = () => { document.getElementById('swatchSelectedDisplay').textContent = li.textContent; s.style.display = 'none'; loadDashFromCustomLibrary(idx); };
        s.appendChild(li);
    });
    s.onmouseleave = restoreDashThumbnail;
}

function loadDashFromCustomLibrary(idx) {
    const v = document.getElementById('libVendor').value; const c = document.getElementById('libCollection').value;
    if(!v || !c || idx === undefined) return;
    const item = dashLocalLibrary[v][c][idx]; const r = new FileReader();
    r.onload = e => {
        const u = e.target.result; const w = dashUnit === 'cm' ? dashFmt(item.width*2.54) : item.width;
        document.getElementById('fW').value = w; document.getElementById('m_fCode').value = item.code;
        dashProjectData[dashSelectedRowIndex].fW = w; dashProjectData[dashSelectedRowIndex].fCode = item.code;
        dashProjectData[dashSelectedRowIndex].swatchDataUrl = u; dashProjectData[dashSelectedRowIndex].swatchName = item.code;
        document.getElementById('view-dashboard').style.setProperty('--frame-bg', `url(${u})`);
        dashActiveImageObj.src = u; dashActiveImageObj.onload = () => syncDashAndCalculate();
    };
    r.readAsDataURL(item.file);
}

function loadDashCustomSwatch(e) {
    const f = e.target.files[0]; if(!f) return;
    const n = f.name.split('.')[0]; const r = new FileReader();
    r.onload = e => {
        document.getElementById('view-dashboard').style.setProperty('--frame-bg', `url(${e.target.result})`);
        dashProjectData[dashSelectedRowIndex].swatchDataUrl = e.target.result; dashProjectData[dashSelectedRowIndex].swatchName = n;
        document.getElementById('m_fCode').value = n; document.getElementById('swatchSelectedDisplay').textContent = 'Frame';
        document.getElementById('swatchThumbPreview').style.backgroundImage = `url(${e.target.result})`;
        dashActiveImageObj.src = e.target.result; dashActiveImageObj.onload = () => syncDashAndCalculate();
    };
    r.readAsDataURL(f);
}

function saveFramePreset() {
    const d = { type: 'frame-preset', unit: dashUnit, frame: dashProjectData[dashSelectedRowIndex] };
    const b = new Blob([JSON.stringify(d, null, 2)], { type: 'application/json' });
    const a = document.createElement('a'); a.href = URL.createObjectURL(b); a.download = `Preset_${d.frame.fCode ? d.frame.fCode.replace(/[^a-z0-9]/gi, '_') : 'Frame'}.json`; a.click();
}

function loadFramePreset(e) {
    const f = e.target.files[0]; if(!f) return;
    const r = new FileReader();
    r.onload = e => {
        try {
            const d = JSON.parse(e.target.result);
            if(d.type === 'frame-preset' && d.frame) {
                const c = dashProjectData[dashSelectedRowIndex];
                d.frame.id = c.id; d.frame.location = c.location; d.frame.level = c.level; d.frame.qty = c.qty; d.frame.imageCode = c.imageCode;
                if(d.unit && d.unit !== dashUnit) {
                    const factor = (dashUnit === 'in') ? (1/2.54) : 2.54;
                    ['extW','extH','fW','bleed','canvasDepth','canvasWrap','m1T','m1B','m1L','m1R','m2'].forEach(p => { if(d.frame[p] !== undefined && !isNaN(d.frame[p])) d.frame[p] = dashFmt(d.frame[p]*factor); });
                }
                dashProjectData[dashSelectedRowIndex] = d.frame; loadDashDataIntoControls(d.frame); renderDashTable(); pushUpdatesToElevations(dashSelectedRowIndex);
            } else alert("Invalid preset.");
        } catch(err) { alert("Error loading file."); }
    };
    r.readAsText(f); e.target.value = '';
}

function exportDashNativePNG() {
    const d = dashProjectData[dashSelectedRowIndex]; const dpi = 72; const w = d.extW*dpi; const h = d.extH*dpi; const fw = d.fW*dpi; const p = 40;
    const c = document.createElement('canvas'); c.width = w + (p*2); c.height = h + (p*2);
    const x = c.getContext('2d'); x.translate(p, p);
    x.shadowColor = 'rgba(0,0,0,0.4)'; x.shadowBlur = 25; x.shadowOffsetY = 15; x.fillStyle = '#000'; x.fillRect(0,0,w,h); x.shadowColor = 'transparent';
    
    function fR(img, rw, rh, sx, sy) {
        if(!img || img.src === emptyImgUrl || d.fType === 'color') { x.fillStyle = d.fColor; x.fillRect(sx,sy,rw,rh); return; }
        const s = rw/img.width; const pt = x.createPattern(img,'repeat'); const m = new DOMMatrix().translate(sx,sy).scale(s,s);
        pt.setTransform(m); x.fillStyle = pt; x.fillRect(sx,sy,rw,rh);
    }
    x.save(); x.beginPath(); x.moveTo(0,0); x.lineTo(fw,fw); x.lineTo(fw,h-fw); x.lineTo(0,h); x.closePath(); x.clip(); fR(dashActiveImageObj,fw,h,0,0); x.restore();
    x.save(); x.beginPath(); x.moveTo(w,0); x.lineTo(w-fw,fw); x.lineTo(w-fw,h-fw); x.lineTo(w,h); x.closePath(); x.clip(); x.translate(w,0); x.scale(-1,1); fR(dashActiveImageObj,fw,h,0,0); x.restore();
    x.save(); x.beginPath(); x.moveTo(0,0); x.lineTo(w,0); x.lineTo(w-fw,fw); x.lineTo(fw,fw); x.closePath(); x.clip(); x.translate(w/2,fw/2); x.rotate(90*Math.PI/180); x.translate(-fw/2,-w/2); fR(dashActiveImageObj,fw,w,0,0); x.restore();
    x.save(); x.beginPath(); x.moveTo(0,h); x.lineTo(w,h); x.lineTo(w-fw,h-fw); x.lineTo(fw,h-fw); x.closePath(); x.clip(); x.translate(w/2,h-fw/2); x.rotate(-90*Math.PI/180); x.translate(-fw/2,-w/2); fR(dashActiveImageObj,fw,w,0,0); x.restore();
    
    function dIS(bx, by, bw, bh, bl, os, op) {
        x.save(); x.beginPath(); x.rect(bx,by,bw,bh); x.clip(); x.shadowColor = `rgba(0,0,0,${op})`; x.shadowBlur = bl; x.shadowOffsetY = os;
        x.lineWidth = 20; x.strokeStyle = '#000'; x.strokeRect(bx-10,by-10,bw+20,bh+20); x.restore();
    }
    const iX = fw; const iY = fw; const iW = w-(fw*2); const iH = h-(fw*2); const isC = (d.product === "Framed Canvas (Floater)");
    const mT = (d.m1A && !isC) ? d.m1T : 0; const mB = (d.m1A && !isC) ? d.m1B : 0; const mL = (d.m1A && !isC) ? d.m1L : 0; const mR = (d.m1A && !isC) ? d.m1R : 0; const m2 = (d.m2A && !isC) ? d.m2 : 0;
    
    if (d.m1A && !isC) { x.fillStyle = "#ffffff"; x.fillRect(iX,iY,iW,iH); dIS(iX,iY,iW,iH,25,10,0.45); x.strokeStyle = "#cccccc"; x.lineWidth = 1; x.strokeRect(iX,iY,iW,iH); }
    const m2X = iX+(mL*dpi); const m2Y = iY+(mT*dpi); const m2W = iW-((mL+mR)*dpi); const m2H = iH-((mT+mB)*dpi);
    if (d.m2A && !isC) { x.fillStyle = "#fdfdfd"; x.fillRect(m2X,m2Y,m2W,m2H); dIS(m2X,m2Y,m2W,m2H,15,6,0.35); x.strokeStyle = "#cccccc"; x.lineWidth = 1; x.strokeRect(m2X,m2Y,m2W,m2H); }
    const aX = m2X+(m2*dpi); const aY = m2Y+(m2*dpi); const aW = m2W-(m2*2*dpi); const aH = m2H-(m2*2*dpi); x.clearRect(aX,aY,aW,aH); 
    
    if (isC) { x.strokeStyle = "#111111"; x.lineWidth = 10; x.strokeRect(aX+5,aY+5,aW-10,aH-10); dIS(aX,aY,aW,aH,20,5,0.8); } 
    else { dIS(aX,aY,aW,aH,8,3,0.25); x.strokeStyle = "#aaaaaa"; x.lineWidth = 1; x.strokeRect(aX,aY,aW,aH); }
    const a = document.createElement('a'); a.download = `${d.id||'Frame'}.png`; a.href = c.toDataURL("image/png"); a.click();
}

function exportDashCSV() {
    const g = (id) => document.getElementById(id).value; const u = ` (${dashUnit})`;
    let csv = `,RFI,PROJECT NAME,${g('g_projName')},,,,,,,,,,,,,,,,,,,,,,,,,\n,,DESCRIPTION,${g('g_desc')},,,,,,,,,,,,,,,,,,,,,,,,,\n,,DATE,${g('g_date')},,,,,,,,,,,,,,,,,,,,,,,,,\n,,ISSUED BY,${g('g_issued')},,,,,,,,,,,,,,,,,,,,,,,,,\n,,CLIENT NAME,${g('g_client')},,,,,,,,,,,,,,,,,,,,,,,,,\n,,Attn:,${g('g_attn')},,,,,,,,,,,,,,,,,,,,,,,,,\n,,Delivery,${g('g_delivery')},,,,,,,,,,,,,,,,,,,,,,,,,\n\nLEVEL,Qty,Item Code,PRODUCT,LOCATION,Image code,Overall Width${u},Overall Height${u},Art Size W${u},Art Size H${u},Print File W${u},Print File H${u},Canvas Stretcher Depth${u},Canvas Image Wrap${u},Mat Code/ Color,Mat Top,Mat Right,Mat Bottom,Mat Left,Mat 2 Reveal,Glass,Frame Code,Frame (Width)${u},Security Hardware,Backing/Substrate,Mount,Notes,Production Notes,Image_Filename\n`;
    dashProjectData.forEach(r => {
        const iC = (r.product === "Framed Canvas (Floater)");
        const mT = (r.m1A && !iC) ? r.m1T : 0; const mB = (r.m1A && !iC) ? r.m1B : 0; const mL = (r.m1A && !iC) ? r.m1L : 0; const mR = (r.m1A && !iC) ? r.m1R : 0; const m2 = (r.m2A && !iC) ? r.m2 : 0;
        const fW = r.extW - (r.fW*2) - mL - mR - (m2*2); const fH = r.extH - (r.fW*2) - mT - mB - (m2*2);
        const iW = iC ? "N/A" : dashFmt(Math.max(0, fW) + (r.bleed*2)); const iH = iC ? "N/A" : dashFmt(Math.max(0, fH) + (r.bleed*2));
        const d = [r.level, r.qty, r.id, r.product, r.location, r.imageCode, dashFmt(r.extW), dashFmt(r.extH), dashFmt(Math.max(0,fW)), dashFmt(Math.max(0,fH)), iW, iH, r.canvasDepth ? dashFmt(r.canvasDepth) : "N/A", r.canvasWrap ? dashFmt(r.canvasWrap) : "N/A", (r.m1A && !iC) ? r.mColor : "None", (r.m1A && !iC) ? dashFmt(r.m1T) : "None", (r.m1A && !iC) ? dashFmt(r.m1R) : "None", (r.m1A && !iC) ? dashFmt(r.m1B) : "None", (r.m1A && !iC) ? dashFmt(r.m1L) : "None", (r.m2A && !iC) ? dashFmt(r.m2) : "None", r.glass, r.fCode, dashFmt(r.fW), r.hardware, r.backing, r.mount, r.notes, r.prodNotes, `${r.id}.png`];
        csv += d.map(s => `"${String(s).replace(/"/g, '""')}"`).join(',') + '\n';
    });
    const b = new Blob([csv], { type: 'text/csv;charset=utf-8;' }); const a = document.createElement("a"); a.href = URL.createObjectURL(b); a.download = `RFI_Project_Tracker.csv`; a.click();
}

function renderDashTable() {
    const tbody = document.getElementById('rfiBody'); tbody.innerHTML = '';
    dashProjectData.forEach((row, index) => {
        const isCanvas = (row.product === "Framed Canvas (Floater)");
        const m1T = (row.m1A && !isCanvas) ? row.m1T : 0; const m1B = (row.m1A && !isCanvas) ? row.m1B : 0; const m1L = (row.m1A && !isCanvas) ? row.m1L : 0; const m1R = (row.m1A && !isCanvas) ? row.m1R : 0;
        const m2 = (row.m2A && !isCanvas) ? row.m2 : 0;

        const finalW = row.extW - (row.fW * 2) - m1L - m1R - (m2 * 2); const finalH = row.extH - (row.fW * 2) - m1T - m1B - (m2 * 2);
        const imgW = isCanvas ? "N/A" : dashFmt(Math.max(0, finalW) + (row.bleed * 2)); const imgH = isCanvas ? "N/A" : dashFmt(Math.max(0, finalH) + (row.bleed * 2));
        
        const tr = document.createElement('tr');
        if (index === dashSelectedRowIndex) tr.className = 'selected';
        tr.addEventListener('click', () => { if (dashSelectedRowIndex !== index) selectDashRow(index); });
        
        tr.innerHTML = `
            <td><input class="tbl-in" type="text" value="${row.level}" oninput="dashHtIn(${index}, 'level', this.value)" style="width:30px;"></td>
            <td><input class="tbl-in" type="number" value="${row.qty}" disabled style="width:30px; opacity:0.6; background:transparent;"></td>
            <td style="font-weight:bold;"><input class="tbl-in" type="text" value="${row.id}" oninput="dashHtIn(${index}, 'id', this.value)" style="width:80px; font-weight:bold;"></td>
            <td><select class="tbl-in no-arrow" onchange="dashHtIn(${index}, 'product', this.value)"><option ${row.product === 'Framed Art' ? 'selected' : ''}>Framed Art</option><option ${row.product === 'Framed Canvas (Floater)' ? 'selected' : ''}>Framed Canvas (Floater)</option><option ${row.product === 'Sourced Object' ? 'selected' : ''}>Sourced Object</option></select></td>
            <td><input class="tbl-in" type="text" value="${row.location}" oninput="dashHtIn(${index}, 'location', this.value)" style="width:90px;"></td>
            <td><input class="tbl-in" type="text" value="${row.imageCode}" oninput="dashHtIn(${index}, 'imageCode', this.value)" style="width:200px;"></td>
            <td><input class="tbl-in" type="number" step="0.125" value="${dashFmt(row.extW)}" oninput="dashHtIn(${index}, 'extW', this.value)" style="width:45px;"></td>
            <td><input class="tbl-in" type="number" step="0.125" value="${dashFmt(row.extH)}" oninput="dashHtIn(${index}, 'extH', this.value)" style="width:45px;"></td>
            <td id="calc-openW-${index}" style="padding: 4px 8px;">${dashFmt(Math.max(0, finalW))}</td><td id="calc-openH-${index}" style="padding: 4px 8px;">${dashFmt(Math.max(0, finalH))}</td><td id="calc-printW-${index}" style="color:var(--accent); font-weight:bold; padding: 4px 8px;">${imgW}</td><td id="calc-printH-${index}" style="color:var(--accent); font-weight:bold; padding: 4px 8px;">${imgH}</td>
            <td><input class="tbl-in" type="number" step="0.125" value="${row.canvasDepth || ''}" oninput="dashHtIn(${index}, 'canvasDepth', this.value)" style="width:45px;"></td>
            <td><input class="tbl-in" type="number" step="0.125" value="${row.canvasWrap || ''}" oninput="dashHtIn(${index}, 'canvasWrap', this.value)" style="width:45px;"></td>
            <td><input class="tbl-in" type="text" value="${row.mColor}" oninput="dashHtIn(${index}, 'mColor', this.value)" style="width:80px;"></td>
            <td><input class="tbl-in" type="number" step="0.125" value="${dashFmt(row.m1T)}" oninput="dashHtIn(${index}, 'm1T', this.value)" style="width:40px;"></td><td><input class="tbl-in" type="number" step="0.125" value="${dashFmt(row.m1R)}" oninput="dashHtIn(${index}, 'm1R', this.value)" style="width:40px;"></td><td><input class="tbl-in" type="number" step="0.125" value="${dashFmt(row.m1B)}" oninput="dashHtIn(${index}, 'm1B', this.value)" style="width:40px;"></td><td><input class="tbl-in" type="number" step="0.125" value="${dashFmt(row.m1L)}" oninput="dashHtIn(${index}, 'm1L', this.value)" style="width:40px;"></td>
            <td><input class="tbl-in" type="number" step="0.125" value="${dashFmt(row.m2)}" oninput="dashHtIn(${index}, 'm2', this.value)" style="width:40px;"></td>
            <td><input class="tbl-in" type="text" value="${row.glass}" oninput="dashHtIn(${index}, 'glass', this.value)" style="width:80px;"></td><td><input class="tbl-in" type="text" value="${row.fCode}" oninput="dashHtIn(${index}, 'fCode', this.value)" style="width:80px;"></td><td><input class="tbl-in" type="number" step="0.125" value="${dashFmt(row.fW)}" oninput="dashHtIn(${index}, 'fW', this.value)" style="width:45px;"></td>
            <td><input class="tbl-in" type="text" value="${row.hardware}" oninput="dashHtIn(${index}, 'hardware', this.value)" style="width:80px;"></td><td><input class="tbl-in" type="text" value="${row.backing}" oninput="dashHtIn(${index}, 'backing', this.value)" style="width:80px;"></td><td><input class="tbl-in" type="text" value="${row.mount}" oninput="dashHtIn(${index}, 'mount', this.value)" style="width:80px;"></td><td><input class="tbl-in" type="text" value="${row.notes}" oninput="dashHtIn(${index}, 'notes', this.value)" style="width:90px;"></td><td><input class="tbl-in" type="text" value="${row.prodNotes}" oninput="dashHtIn(${index}, 'prodNotes', this.value)" style="width:90px;"></td>
        `;
        tbody.appendChild(tr);
    });
}

// =========================================================================
// VIEW 2: ELEVATION LOGIC
// =========================================================================
function getElevLetter(index) { let res = ""; let curr = index; while (curr >= 0) { res = String.fromCharCode((curr % 26) + 65) + res; curr = Math.floor(curr / 26) - 1; } return res; }
function updateElevZoom(val) { elevZoomFactor = parseFloat(val); drawElevAll(); }
function fitElevZoom() { elevZoomFactor = 1; document.getElementById('zoomSlider').value = 1; drawElevAll(); }
function updateElevWall() {
    elevations[currentElevIndex].wallW = parseFloat(document.getElementById('wallW').value) || 185;
    elevations[currentElevIndex].wallH = parseFloat(document.getElementById('wallH').value) || 108;
    drawElevAll();
}

function autoElevRelabel() {
    let sortedFrames = [...elevFrames].sort((a, b) => a.x - b.x);
    let letterMap = {};
    sortedFrames.forEach((f, i) => { letterMap[f.letter] = getElevLetter(i); f.letter = getElevLetter(i); });
    sortedFrames.forEach(f => { if (f.dimTo && f.dimTo.length > 0) f.dimTo = f.dimTo.map(t => letterMap[t] || t); });
    elevFrames = sortedFrames; initElevControls(); drawElevAll();
}

function initElevControls() {
    const container = document.getElementById('frame-controls');
    let html = ``;
    elevFrames.forEach((f, idx) => {
        const activeNeighbors = elevFrames.filter(n => n.letter !== f.letter && n.active);
        const targetButtons = activeNeighbors.map(n => `<button class="toggle-status ${f.dimTo.includes(n.letter)?'active':''}" style="padding:1px 3px; font-size:8px; border-radius:2px;" onclick="toggleElevDimTarget(${idx}, '${n.letter}', event)">${n.letter}</button>`).join('');

        html += `
            <div class="compact-frame-item">
                <div style="flex:1; min-width:0; display:flex; flex-direction:column;">
                    <span style="font-weight:bold; font-size:0.75rem; color:var(--text-strong);">${f.letter} <span style="font-weight:normal; font-size:0.65rem; color:var(--text-muted);">(${f.id})</span></span>
                    <div style="display:flex; gap:2px; margin-top:2px; flex-wrap:wrap;">${targetButtons}</div>
                </div>
                <div class="frame-item-icons">
                    <div style="width:38px; display:flex; justify-content:center;">
                        <button class="toggle-status ${f.active?'active':''}" style="font-size:0.5rem; padding:2px 5px;" onclick="toggleElevActive(${idx}, event)">${f.active?'ON':'OFF'}</button>
                    </div>
                    <div style="width:28px; display:flex; justify-content:center;">
                        <button class="icon-btn ${f.isGrouped ? 'grouped' : ''}" title="Move/Group" onclick="toggleElevGroup(${idx}, event)">${svgMove}</button>
                    </div>
                    <div style="width:26px; display:flex; justify-content:center;">
                        <button class="icon-btn" title="Edit Master" onclick="jumpToDashboard('${f.id}')">${svgEdit}</button>
                    </div>
                    <div style="width:26px; display:flex; justify-content:center;">
                        <button class="icon-btn" title="Duplicate" onclick="duplicateElevFrame(${idx}, event)">${svgDup}</button>
                    </div>
                    <div style="width:26px; display:flex; justify-content:center;">
                        <button class="icon-btn" title="Remove" onclick="removeElevFrame(${idx}, event)">${svgTrash}</button>
                    </div>
                </div>
            </div>`;
    });
    container.innerHTML = html;
}

function toggleElevDimTarget(idx, targetLetter, e) {
    e.stopPropagation(); const arr = elevFrames[idx].dimTo || [];
    if (arr.includes(targetLetter)) elevFrames[idx].dimTo = arr.filter(l => l !== targetLetter);
    else elevFrames[idx].dimTo.push(targetLetter);
    initElevControls(); drawElevAll();
}

function toggleElevGroup(idx, e) { e.stopPropagation(); elevFrames[idx].isGrouped = !elevFrames[idx].isGrouped; initElevControls(); }
function removeElevFrame(idx, e) { e.stopPropagation(); elevFrames.splice(idx, 1); elevFrames.forEach((f, i) => f.letter = getElevLetter(i)); initElevControls(); drawElevAll(); recalculateDashboardQuantities(); }
function toggleElevActive(idx, e) { e.stopPropagation(); elevFrames[idx].active = !elevFrames[idx].active; initElevControls(); drawElevAll(); recalculateDashboardQuantities(); }

function duplicateElevFrame(idx, e) { 
    e.stopPropagation(); const temp = elevFrames[idx]; const nF = JSON.parse(JSON.stringify(temp));
    nF.letter = getElevLetter(elevFrames.length); nF.x = temp.x + 10; nF.y = temp.y - 10; elevFrames.push(nF); 
    initElevControls(); drawElevAll(); recalculateDashboardQuantities(); 
}

function toggleElevLayer(id, btn) {
    const layer = document.getElementById(id);
    const isHidden = (layer.style.display === 'none');
    layer.style.display = isHidden ? 'block' : 'none';
    btn.classList.toggle('active', isHidden); 
}

function selectAllElevFrames() {
    elevFrames.forEach(f => f.active = true);
    initElevControls(); drawElevAll(); recalculateDashboardQuantities();
}

function deselectAllElevFrames() {
    elevFrames.forEach(f => f.active = false);
    initElevControls(); drawElevAll(); recalculateDashboardQuantities();
}

function toggleGroupAllElevFrames(e) {
    e.stopPropagation();
    const btn = document.getElementById('groupAllBtn');
    const anyGrouped = elevFrames.some(f => f.isGrouped);
    // If any are grouped, ungroup all; else group all
    elevFrames.forEach(f => f.isGrouped = !anyGrouped);
    btn.classList.toggle('grouped', !anyGrouped);
    initElevControls();
}

async function batchDownloadAllFrames() {
    if (dashProjectData.length === 0) return alert("No frames to download.");
    const origIndex = dashSelectedRowIndex;
    for (let i = 0; i < dashProjectData.length; i++) {
        selectDashRow(i);
        await new Promise(resolve => setTimeout(resolve, 80));
        exportDashNativePNG();
        await new Promise(resolve => setTimeout(resolve, 150));
    }
    selectDashRow(origIndex);
    alert(`Batch download complete! ${dashProjectData.length} frame(s) saved.`);
}

function elevFmt(val) { return elevUnit === 'in' ? Math.round(val).toString() : parseFloat(val).toFixed(1); }

function drawElevAll() {
    const wallW = parseFloat(document.getElementById('wallW').value) || 1; const wallH = parseFloat(document.getElementById('wallH').value) || 1;
    const workspace = document.querySelector('#view-elevation .workspace');
    
    let baseScale = Math.min((workspace.clientWidth - 160)/wallW, (workspace.clientHeight - 160)/wallH);
    elevScale = baseScale * elevZoomFactor;
    
    const wall = document.getElementById('wall');
    wall.style.width = (wallW * elevScale) + 'px'; wall.style.height = (wallH * elevScale) + 'px';

    const gridLayer = document.getElementById('grid-layer');
    const gridCellSize = (elevUnit === 'in' ? 1 : 2.54) * elevScale;
    gridLayer.style.backgroundSize = gridCellSize + 'px ' + gridCellSize + 'px';
    gridLayer.style.backgroundImage = 'linear-gradient(to right, rgba(0,0,0,0.06) 1px, transparent 1px), linear-gradient(to top, rgba(0,0,0,0.06) 1px, transparent 1px)';

    const personHeightIn = 72; // always 72 inches = 6 feet
    const personHeight = elevUnit === 'in' ? personHeightIn : parseFloat((personHeightIn * 2.54).toFixed(2)); // 182.88 cm
    const pWrap = document.getElementById('person-wrap');
    const pMeasure = document.getElementById('person-measure');
    document.getElementById('person').style.height = (personHeight * elevScale) + 'px';
    pMeasure.style.height = (personHeight * elevScale) + 'px';
    pMeasure.querySelector('.arch-label').innerText = elevUnit === 'in' ? `72"` : `182.9 cm`;
    pWrap.style.left = (elevPersonPos.x * elevScale) + 'px';

    const frameLayer = document.getElementById('frame-layer'); frameLayer.innerHTML = '';
    const labelLayer = document.getElementById('label-layer'); labelLayer.innerHTML = '';
    const odLayer = document.getElementById('od-layer'); odLayer.innerHTML = '';
    const centerLayer = document.getElementById('frame-center-layer'); centerLayer.innerHTML = '';
    
    elevFrames.forEach((f, idx) => {
        if(!f.active) return;
        
        const el = document.createElement('div'); el.className = 'draggable frame-vis';
        el.style.cssText = `width:${f.w*elevScale}px; height:${f.h*elevScale}px; left:${f.x*elevScale}px; bottom:${f.y*elevScale}px;`;
        el.style.setProperty('--fW', (f.fW * elevScale) + 'px');
        el.style.setProperty('--frame-W', (f.w * elevScale) + 'px');
        el.style.setProperty('--frame-color', f.fColor || '#1a1a1a');
        
        if (f.swatchDataUrl) {
            el.classList.add('frame-vis-image');
            el.style.setProperty('--frame-bg', `url(${f.swatchDataUrl})`);
        } else {
            el.classList.add('frame-vis-solid');
            el.style.setProperty('--frame-bg', `none`);
        }

        const rails = ['top', 'bottom', 'left', 'right'];
        rails.forEach(pos => {
            const rail = document.createElement('div');
            rail.className = `frame-rail rail-${pos}`;
            rail.innerHTML = `<div class="rail-bg"></div>`;
            el.appendChild(rail);
        });

        const m1 = document.createElement('div'); m1.className = 'mat-visual';
        m1.style.cssText = `top:${f.fW*elevScale}px; left:${f.fW*elevScale}px; width:${(f.w - f.fW*2)*elevScale}px; height:${(f.h - f.fW*2)*elevScale}px; border-top-width:${f.m1T*elevScale}px; border-bottom-width:${f.m1B*elevScale}px; border-left-width:${f.m1L*elevScale}px; border-right-width:${f.m1R*elevScale}px; border-color:#ffffff;`;
        el.appendChild(m1);
        
        let m2Val = f.m2Active ? f.m2 : 0;
        if (f.m2Active) {
            const m2 = document.createElement('div'); m2.className = 'mat2-visual';
            m2.style.cssText = `top:${(f.fW + f.m1T)*elevScale}px; left:${(f.fW + f.m1L)*elevScale}px; width:${(f.w - f.fW*2 - f.m1L - f.m1R)*elevScale}px; height:${(f.h - f.fW*2 - f.m1T - f.m1B)*elevScale}px; border-width:${m2Val*elevScale}px; border-color:#fdfdfd;`;
            el.appendChild(m2);
        }
        
        const artW = f.w - f.fW*2 - f.m1L - f.m1R - m2Val*2; const artH = f.h - f.fW*2 - f.m1T - f.m1B - m2Val*2;
        const art = document.createElement('div'); art.className = 'art-visual';
        const artFontSize = Math.max(6, Math.min(11, (artW * elevScale) / 6));
        art.style.cssText = `top:${(f.fW + f.m1T + m2Val)*elevScale}px; left:${(f.fW + f.m1L + m2Val)*elevScale}px; width:${artW*elevScale}px; height:${artH*elevScale}px; font-size:${artFontSize}px;`;
        const unitSuffix = elevUnit === 'in' ? '"' : ' cm';
        art.innerText = (artW > 0) ? `${artW.toFixed(1)}${unitSuffix}\nx\n${artH.toFixed(1)}${unitSuffix}` : "";
        el.appendChild(art);
        
        const labelTag = document.createElement('div'); labelTag.className = 'frame-id-tag';
        labelTag.style.left = (f.x * elevScale) + 'px'; labelTag.style.bottom = ((f.y + f.h) * elevScale) + 'px';
        labelTag.innerText = f.letter; labelLayer.appendChild(labelTag);

        const odTag = document.createElement('div'); odTag.className = 'od-id-tag';
        odTag.style.left = ((f.x + f.w) * elevScale) + 'px'; odTag.style.bottom = ((f.y + f.h) * elevScale) + 'px';
        const odSuffix = elevUnit === 'in' ? '"' : 'cm';
        odTag.innerText = `OD: ${f.w.toFixed(1)}x${f.h.toFixed(1)}${odSuffix}`; odLayer.appendChild(odTag);

        const crossH = document.createElement('div'); crossH.className = 'crosshair-h';
        crossH.style.width = ((f.w + 6) * elevScale) + 'px'; crossH.style.left = ((f.x - 3) * elevScale) + 'px'; crossH.style.bottom = ((f.y + f.h/2) * elevScale) + 'px';
        const crossV = document.createElement('div'); crossV.className = 'crosshair-v';
        crossV.style.height = ((f.h + 6) * elevScale) + 'px'; crossV.style.left = ((f.x + f.w/2) * elevScale) + 'px'; crossV.style.bottom = ((f.y - 3) * elevScale) + 'px';
        centerLayer.appendChild(crossH); centerLayer.appendChild(crossV);
        
        makeElevDraggable(el, idx); frameLayer.appendChild(el);
    });
    drawElevTargetedSpacing(); drawElevGuides(wallW, wallH);
}

function makeElevDraggable(el, idx) {
    el.onmousedown = function(e) {
        if (e.target.tagName === 'INPUT' || e.target.tagName === 'BUTTON') return;
        e.preventDefault(); let sx = e.clientX, sy = e.clientY;
        document.onmousemove = function(e) {
            let dx = (sx - e.clientX)/elevScale, dy = (sy - e.clientY)/elevScale; sx = e.clientX; sy = e.clientY;
            const snap = elevUnit === 'in' ? 1 : 2.54;
            if(idx === 'person') { elevPersonPos.x -= dx; } 
            else { 
                let frame = elevFrames[idx]; let prevX = frame.x; let prevY = frame.y;
                frame.x = Math.round((frame.x - dx)/snap)*snap; frame.y = Math.round((frame.y + dy)/snap)*snap; 
                let actualDx = frame.x - prevX; let actualDy = frame.y - prevY;
                if(frame.isGrouped) { elevFrames.forEach((f, i) => { if(i !== idx && f.active && f.isGrouped) { f.x += actualDx; f.y += actualDy; } }); }
            }
            drawElevAll();
        };
        document.onmouseup = () => document.onmousemove = null;
    };
}

function drawElevGuides(wallW, wallH) {
    const guideLayer = document.getElementById('guide-layer'); guideLayer.innerHTML = '';
    const archLayer = document.getElementById('arch-dim-layer'); archLayer.innerHTML = '';
    
    const cl = document.createElement('div'); cl.className = 'center-guide';
    cl.style.left = ((wallW / 2) * elevScale) + 'px'; cl.style.bottom = '0px';
    cl.innerHTML = `<span class="center-label">WALL CENTER</span>`;
    guideLayer.appendChild(cl);

    const hangVal = elevUnit === 'in' ? 57 : 144.78;
    if(hangVal < wallH) {
        const hl = document.createElement('div'); hl.className = 'hang-guide';
        hl.style.bottom = (hangVal * elevScale) + 'px';
        hl.innerHTML = `<span class="hang-label">HANG HEIGHT: ${elevFmt(hangVal)}${elevUnit==='in'?'"':''}</span>`;
        guideLayer.appendChild(hl);
    }
    
    createElevArchDim(0, wallH + 6, wallW, wallH + 6, 'h', `${elevFmt(wallW)}${elevUnit === 'in' ? '"' : ' cm'}`, archLayer, true);
    createElevArchDim(-6, 0, -6, wallH, 'v', `${elevFmt(wallH)}${elevUnit === 'in' ? '"' : ' cm'}`, archLayer, true);
}

function drawElevTargetedSpacing() {
    const layer = document.getElementById('dim-layer'); layer.innerHTML = '';
    let drawnPairs = new Set();
    elevFrames.forEach(f1 => {
        if (!f1.active || !f1.dimTo) return;
        f1.dimTo.forEach(targetLetter => {
            let f2 = elevFrames.find(tf => tf.letter === targetLetter && tf.active);
            if(f2) {
                let pairId = [f1.letter, f2.letter].sort().join('-');
                if(!drawnPairs.has(pairId)) {
                    let leftF = f1.x < f2.x ? f1 : f2; let rightF = f1.x < f2.x ? f2 : f1;
                    if (rightF.x >= leftF.x + leftF.w) {
                        let gapX = rightF.x - (leftF.x + leftF.w);
                        let oTop = Math.max(leftF.y, rightF.y); let oBot = Math.min(leftF.y + leftF.h, rightF.y + rightF.h);
                        let anchorY = oBot > oTop ? oTop + (oBot - oTop)/2 : (leftF.y + leftF.h/2 + rightF.y + rightF.h/2)/2;
                        createElevArchSpacing(leftF.x + leftF.w, anchorY, rightF.x, anchorY, 'h', layer, elevFmt(gapX));
                    }
                    let botF = f1.y < f2.y ? f1 : f2; let topF = f1.y < f2.y ? f2 : f1;
                    if (topF.y >= botF.y + botF.h) {
                        let gapY = topF.y - (botF.y + botF.h);
                        let oLeft = Math.max(botF.x, topF.x); let oRight = Math.min(botF.x + botF.w, topF.x + topF.w);
                        let anchorX = oRight > oLeft ? oLeft + (oRight - oLeft)/2 : (botF.x + botF.w/2 + topF.x + topF.w/2)/2;
                        createElevArchSpacing(anchorX, botF.y + botF.h, anchorX, topF.y, 'v', layer, elevFmt(gapY));
                    }
                    drawnPairs.add(pairId);
                }
            }
        });
    });
}

function createElevArchDim(x1, y1, x2, y2, type, label, container, isWallOuter) {
    const dim = document.createElement('div');
    dim.className = 'arch-dim ' + (type === 'h' ? 'arch-dim-h' : 'arch-dim-v');
    
    const left = Math.min(x1, x2) * elevScale;
    const bottom = Math.min(y1, y2) * elevScale;
    dim.style.left = left + 'px';
    dim.style.bottom = bottom + 'px';

    const offset = 6 * elevScale; 

    if(type === 'h') {
        const width = Math.abs(x2 - x1) * elevScale;
        dim.style.width = width + 'px';
        dim.innerHTML = `
            ${isWallOuter ? `<div style="position:absolute; left:0; top:0; width:1px; height:${offset}px; border-left:1.5px dashed var(--guide-color);"></div><div style="position:absolute; right:0; top:0; width:1px; height:${offset}px; border-left:1.5px dashed var(--guide-color);"></div>` : ''}
            <div class="dim-line-segment"></div>
            <span class="arch-label-new">${label}</span>
            <div class="dim-line-segment"></div>
        `;
    } else {
        const height = Math.abs(y2 - y1) * elevScale;
        dim.style.height = height + 'px';
        dim.innerHTML = `
            ${isWallOuter ? `<div style="position:absolute; left:0; bottom:0; height:1px; width:${offset}px; border-top:1.5px dashed var(--guide-color);"></div><div style="position:absolute; left:0; top:0; height:1px; width:${offset}px; border-top:1.5px dashed var(--guide-color);"></div>` : ''}
            <div class="dim-line-segment-v"></div>
            <span class="arch-label-new arch-label-v">${label}</span>
            <div class="dim-line-segment-v"></div>
        `;
    }
    container.appendChild(dim);
}

function createElevArchSpacing(x1, y1, x2, y2, type, container, label) {
    const dim = document.createElement('div'); 
    dim.className = 'arch-dim ' + (type === 'h' ? 'arch-dim-h' : 'arch-dim-v');
    
    if(type === 'h') {
        const width = Math.abs(x2 - x1) * elevScale; const left = Math.min(x1, x2) * elevScale; const bottom = y1 * elevScale;
        dim.style.cssText = `width:${width}px; height:1.2px; left:${left}px; bottom:${bottom}px;`;
        dim.innerHTML = `<div class="dim-line-segment"></div><span class="arch-label-new">${label}</span><div class="dim-line-segment"></div>`;
    } else {
        const height = Math.abs(y2 - y1) * elevScale; const left = x1 * elevScale; const bottom = Math.min(y1, y2) * elevScale;
        dim.style.cssText = `height:${height}px; width:1.2px; left:${left}px; bottom:${bottom}px;`;
        dim.innerHTML = `<div class="dim-line-segment-v"></div><span class="arch-label-new arch-label-v">${label}</span><div class="dim-line-segment-v"></div>`;
    }
    container.appendChild(dim);
}

function setElevUnit(u) {
    if(elevUnit === u) return;
    if (elevations[currentElevIndex]) {
        elevations[currentElevIndex].wallW = parseFloat(document.getElementById('wallW').value) || elevations[currentElevIndex].wallW;
        elevations[currentElevIndex].wallH = parseFloat(document.getElementById('wallH').value) || elevations[currentElevIndex].wallH;
    }
    const f = u === 'cm' ? 2.54 : (1/2.54);
    elevations.forEach(elev => {
        elev.wallW = parseFloat((parseFloat(elev.wallW) * f).toFixed(2));
        elev.wallH = parseFloat((parseFloat(elev.wallH) * f).toFixed(2));
        elev.frames.forEach(fr => {
            ['w','h','fW','m1T','m1B','m1L','m1R','m2','x','y'].forEach(p => {
                fr[p] = parseFloat((parseFloat(fr[p] || 0) * f).toFixed(4));
            });
        });
        elev.personPos.x = parseFloat((parseFloat(elev.personPos.x || 0) * f).toFixed(2));
    });
    
    elevUnit = u;
    document.getElementById('wallW').value = elevations[currentElevIndex].wallW;
    document.getElementById('wallH').value = elevations[currentElevIndex].wallH;
    document.getElementById('elevBtnInch').classList.toggle('active', elevUnit==='in'); 
    document.getElementById('elevBtnCm').classList.toggle('active', elevUnit==='cm');
    initElevControls(); drawElevAll();
}

async function exportElevPNG() {
    const ws = document.querySelector('#view-elevation .workspace');
    const wrap = document.getElementById('export-wrap');
    const wall = document.getElementById('wall');
    
    const oldOverflow = ws.style.overflow; ws.style.overflow = 'visible';
    
    const arts = document.querySelectorAll('#view-elevation .art-visual');
    const originalBackgrounds = []; const originalColors = [];
    arts.forEach(art => {
        originalBackgrounds.push(art.style.background); originalColors.push(art.style.color);
        art.style.background = 'transparent'; art.style.color = 'transparent';
    });

    const oldWallBg = wall.style.background; wall.style.background = 'transparent'; 
    
    const canvas = await html2canvas(wrap, { backgroundColor: null, scale: 2 });
    
    ws.style.overflow = oldOverflow;
    wall.style.background = oldWallBg; 
    arts.forEach((art, i) => { art.style.background = originalBackgrounds[i]; art.style.color = originalColors[i]; }); 

    const a = document.createElement('a'); a.download = `${elevations[currentElevIndex].name.replace(/[^a-z0-9]/gi, '_')}.png`; a.href = canvas.toDataURL("image/png"); a.click();
}

// BOOT UP THE ENGINE
initMasterApp();