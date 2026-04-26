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
const svgMove = `<svg class="svg-icon" viewBox="0 0 24 24"><path d="M5 9l-3 3 3 3M9 5l3-3 3 3M19 9l3 3-3 3M9 19l3 3 3-3M2 12h20M12 2v20"/></svg>`;
const svgEdit = `<svg class="svg-icon" viewBox="0 0 24 24"><path d="M12 20h9M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"/></svg>`;
const svgDup = `<svg class="svg-icon" viewBox="0 0 24 24"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>`;
const svgTrash = `<svg class="svg-icon" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/><path d="M10 11v6M14 11v6"/><path d="M9 6V4a2 2 0 0 1 2-2h2a2 2 0 0 1 2 2v2"/></svg>`;

const dashDefaultData = { 
    id: "ART.001", imageCode: "TBD", level: "1", qty: 0, product: "Framed Art", location: "LOBBY", 
    bleed: 0.25, canvasDepth: "", canvasWrap: "",
    extW: 24, extH: 24, fType: "color", fW: 0.75, fColor: "#000000", fCode: "Standard Black", 
    swatchDataUrl: "", swatchName: "",
    m1A: true, m1T: 3, m1B: 3, m1L: 3, m1R: 3, m1Locked: false, m1ColorName: "B 97 White", m1ColorHex: "#ffffff",
    m2A: false, m2: 0.25, m2ColorName: "B 97 White", m2ColorHex: "#ffffff", matsLinked: true,
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

let pendingDuplicateIndex = null;

// =========================================================================
// INITIALIZATION & NAVIGATION
// =========================================================================
function initMasterApp() {
    document.getElementById('g_date').valueAsDate = new Date();
    renderNavTabs();
    selectDashRow(0); 
    populateDashPushSelector();
    updateDimFontSize();
    
    document.addEventListener('click', function(event) {
        const container = document.getElementById('customSwatchContainer');
        const sList = document.getElementById('swatchDropdownList');
        if (container && !container.contains(event.target)) {
            if (sList && sList.style.display === 'block') { sList.style.display = 'none'; restoreDashThumbnail(); }
        }
        const bList = document.getElementById('bulkDropdownList');
        const bBtn = document.getElementById('bulkImportBtn');
        if (bList && bList.style.display === 'block') {
            if (bBtn && !bBtn.contains(event.target) && !bList.contains(event.target)) {
                bList.style.display = 'none';
            }
        }
    });
}

function toggleTheme() { document.body.classList.toggle('light-theme'); }

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
    if (newName.trim() !== "") { elevations[currentElevIndex].name = newName; renderNavTabs(); populateDashPushSelector(); }
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
        } else if (currentElevIndex > idx) { currentElevIndex--; }
        if (currentElevIndex === idx || elevations.length === 1) switchView('dashboard');
        
        renderNavTabs(); populateDashPushSelector(); recalculateDashboardQuantities();
    }
}

function addNewElevationTab() {
    let newIndex = elevations.length;
    let w = elevUnit === 'cm' ? parseFloat((185 * 2.54).toFixed(2)) : 185;
    let h = elevUnit === 'cm' ? parseFloat((108 * 2.54).toFixed(2)) : 108;
    let px = elevUnit === 'cm' ? parseFloat((-60 * 2.54).toFixed(2)) : -60;
    
    elevations.push({ name: "Elevation " + (newIndex + 1), frames: [], wallW: w, wallH: h, personPos: {x: px} });
    renderNavTabs(); populateDashPushSelector(); switchView('elevation', newIndex);
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
        currentView = 'elevation'; currentElevIndex = index;
        
        let elev = elevations[currentElevIndex];
        document.getElementById('elev-title-input').value = elev.name;
        document.getElementById('wallW').value = elev.wallW;
        document.getElementById('wallH').value = elev.wallH;
        elevFrames = elev.frames; elevPersonPos = elev.personPos;
        
        populateElevBulkList(); initElevControls(); drawElevAll();
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
    const masterData = { type: 'master-studio-v6', dashUnit: dashUnit, elevUnit: elevUnit, globalMeta: globalMeta, dashProjectData: dashProjectData, elevations: elevations };
    const blob = new Blob([JSON.stringify(masterData, null, 2)], { type: 'application/json' });
    const link = document.createElement('a'); link.href = URL.createObjectURL(blob); link.download = `Master_Studio_Project.json`; link.click();
}

function loadMasterProject(event) {
    const file = event.target.files[0]; if (!file) return;
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
    elevations.forEach(elev => { elev.frames.forEach(f => { if (f.active && f.id) counts[f.id] = (counts[f.id] || 0) + 1; }); });
    dashProjectData.forEach(d => { d.qty = counts[d.id] !== undefined ? counts[d.id] : 0; });
    if (currentView === 'dashboard') { loadDashDataIntoControls(dashProjectData[dashSelectedRowIndex]); renderDashTable(); }
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
    if(cbs.length === 0) return alert("Select at least one frame from the list!");
    
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
            m1A: f.m1A !== false, m1Locked: f.m1Locked || false, m1ColorHex: f.m1ColorHex || '#ffffff',
            m2: (parseFloat(f.m2) || 0) * factor, m2A: f.m2A || false, m2ColorHex: f.m2ColorHex || '#ffffff',
            x: startX, y: 10, isOpen: false, isGrouped: false, dimTo: [], active: true
        };
        elevFrames.push(newFrame); startX += (newFrame.w + 5); cb.checked = false; 
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
        m1A: f.m1A !== false, m1Locked: f.m1Locked || false, m1ColorHex: f.m1ColorHex || '#ffffff',
        m2: (parseFloat(f.m2) || 0) * factor, m2A: f.m2A || false, m2ColorHex: f.m2ColorHex || '#ffffff',
        x: startX, y: 10, isOpen: false, isGrouped: false, dimTo: [], active: true
    });
    
    recalculateDashboardQuantities(); alert(`Pushed ${f.id} to ${targetElev.name}!`);
}

function jumpToDashboard(frameId) {
    const targetIdx = dashProjectData.findIndex(d => d.id === frameId);
    if (targetIdx !== -1) { switchView('dashboard'); selectDashRow(targetIdx); }
}

function checkGlobalEditingWarning(id) {
    // Always recompute and update the banner based on the currently selected row.
    // (Previously this returned early if the row was 'already warned', which left the
    //  banner stuck open when the user switched to a row with zero wall references.)
    let count = 0;
    elevations.forEach(e => { e.frames.forEach(f => { if(f.id === id) count++; }); });
    
    const banner = document.getElementById('linkedWarningBanner');
    if (count > 1) {
        // Only show the banner when the spec is genuinely shared across multiple instances.
        // A single instance on one wall is the normal case and doesn't need a warning.
        document.getElementById('linkedCount').innerText = count;
        banner.style.display = 'flex';
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
                f.m1A = d.m1A !== false; f.m1ColorHex = d.m1ColorHex;
                f.m2 = (parseFloat(d.m2) || 0) * factor; f.m2A = d.m2A; f.m2ColorHex = d.m2ColorHex;
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

function generateNextItemCode() {
    let max = 0;
    dashProjectData.forEach(d => {
        if (d.id.startsWith("ART.")) {
            let num = parseInt(d.id.replace("ART.", ""), 10);
            if (!isNaN(num) && num > max) max = num;
        }
    });
    return "ART." + String(max + 1).padStart(3, '0');
}

function toggleDashSection(id, btn) {
    const sec = document.getElementById(id);
    const span = btn.querySelector('span');
    if (sec.classList.contains('open')) {
        sec.classList.remove('open'); span.innerHTML = `<svg class="svg-icon" style="width:10px; height:10px; transform:rotate(-90deg);" viewBox="0 0 24 24"><polyline points="6 9 12 15 18 9"></polyline></svg>`;
    } else {
        sec.classList.add('open'); span.innerHTML = `<svg class="svg-icon" style="width:10px; height:10px;" viewBox="0 0 24 24"><polyline points="6 9 12 15 18 9"></polyline></svg>`;
    }
}

function toggleDashSwatchDropdown() {
    const s = document.getElementById('swatchDropdownList');
    s.style.display = s.style.display === 'block' ? 'none' : 'block';
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
    const newRow = JSON.parse(JSON.stringify(dashDefaultData)); 
    newRow.id = generateNextItemCode();
    newRow.extW = dashUnit === 'cm' ? dashFmt(24*2.54) : 24; 
    newRow.extH = dashUnit === 'cm' ? dashFmt(24*2.54) : 24; 
    newRow.fW = dashUnit === 'cm' ? dashFmt(0.75*2.54) : 0.75;
    newRow.fType = "color"; newRow.fColor = "#000000"; newRow.fCode = "Standard Black";
    
    dashProjectData.push(newRow);
    dashSelectedRowIndex = dashProjectData.length - 1;
    loadDashDataIntoControls(dashProjectData[dashSelectedRowIndex]);
    renderDashTable(); checkGlobalEditingWarning(newRow.id);
}

function duplicateDashRow() {
    const newRow = JSON.parse(JSON.stringify(dashProjectData[dashSelectedRowIndex])); 
    newRow.id = generateNextItemCode(); 
    newRow.qty = 0; 
    dashProjectData.push(newRow);
    dashSelectedRowIndex = dashProjectData.length - 1;
    loadDashDataIntoControls(dashProjectData[dashSelectedRowIndex]);
    renderDashTable(); checkGlobalEditingWarning(newRow.id);
}

function detachDashRow() {
    // duplicateDashRow already creates a new spec with a fresh item code and qty=0,
    // and selects it. Existing wall references stay on the original spec.
    duplicateDashRow();
    const newId = dashProjectData[dashSelectedRowIndex].id;
    // Force the banner to re-evaluate against the new (un-linked) spec.
    checkGlobalEditingWarning(newId);
    alert(`Detached. Settings copied to a new independent item code (${newId}). The walls still reference the original; you can now edit this new spec without affecting them, or push it to new walls.`);
}

function deleteDashRow() {
    if(dashProjectData.length <= 1) return alert("Cannot delete the last row.");
    const idToDelete = dashProjectData[dashSelectedRowIndex].id;
    dashProjectData.splice(dashSelectedRowIndex, 1);
    dashSelectedRowIndex = Math.max(0, dashSelectedRowIndex - 1);
    
    elevations.forEach(elev => { elev.frames = elev.frames.filter(f => f.id !== idToDelete); });
    
    loadDashDataIntoControls(dashProjectData[dashSelectedRowIndex]);
    renderDashTable(); checkGlobalEditingWarning(dashProjectData[dashSelectedRowIndex].id);
}

function loadDashDataIntoControls(data) {
    const setVal = (id, val) => { const el = document.getElementById(id); if(el) el.value = val; };
    setVal('m_itemCode', data.id); setVal('m_imageCode', data.imageCode); setVal('m_level', data.level); setVal('m_qty', data.qty);
    setVal('m_product', data.product); setVal('m_location', data.location); setVal('m_bleed', dashFmt(data.bleed)); setVal('canvasDepth', dashFmt(data.canvasDepth)); setVal('canvasWrap', dashFmt(data.canvasWrap));
    setVal('extW', dashFmt(data.extW)); setVal('extH', dashFmt(data.extH)); setVal('fType', data.fType); setVal('fW', dashFmt(data.fW)); 
    setVal('fColor', data.fColor); setVal('m_fCode', data.fCode); 
    // Reflect frame style on the toggle buttons (hidden #fType is the source of truth)
    document.getElementById('fTypeBtnLibrary').classList.toggle('active', data.fType === 'image');
    document.getElementById('fTypeBtnSolid').classList.toggle('active', data.fType === 'color');
    applyFrameStyleDimming(data.fType);
    setVal('m1_color', data.m1ColorName); setVal('m1_colorHex', data.m1ColorHex || '#ffffff');
    setVal('m2_color', data.m2ColorName); setVal('m2_colorHex', data.m2ColorHex || '#ffffff');
    setVal('m_glass', data.glass); setVal('m_hardware', data.hardware); setVal('m_mount', data.mount); setVal('m_backing', data.backing); setVal('m_notes', data.notes); setVal('m_prodNotes', data.prodNotes);

    const m1On = data.m1A !== false;
    document.getElementById('m1Toggle').classList.toggle('active', m1On); document.getElementById('m1Toggle').innerText = m1On ? 'ON' : 'OFF';
    document.querySelectorAll('.m1-input').forEach(el => el.disabled = !m1On);
    
    // M2 follows M1: it can only be active when M1 is active. The toggle is disabled when M1 is off.
    const m2EffectivelyOn = m1On && data.m2A;
    const m2Btn = document.getElementById('m2Toggle');
    m2Btn.classList.toggle('active', m2EffectivelyOn); m2Btn.innerText = m2EffectivelyOn ? 'ON' : 'OFF';
    m2Btn.disabled = !m1On;
    m2Btn.style.opacity = m1On ? '1' : '0.4';
    m2Btn.style.cursor = m1On ? 'pointer' : 'not-allowed';
    document.getElementById('m2').disabled = !m2EffectivelyOn;
    
    document.getElementById('m1Lock').classList.toggle('active', data.m1Locked); document.getElementById('m1Lock').innerText = data.m1Locked ? 'LOCKED' : 'UNLOCKED';
    
    const linkBtn = document.getElementById('matLinkBtn');
    linkBtn.classList.toggle('active', data.matsLinked !== false);
    linkBtn.style.color = linkBtn.classList.contains('active') ? 'var(--accent)' : 'var(--text-muted)';
    
    setVal('m1T', dashFmt(data.m1T)); setVal('m1B', dashFmt(data.m1B)); setVal('m1L', dashFmt(data.m1L)); setVal('m1R', dashFmt(data.m1R)); setVal('m2', dashFmt(data.m2));

    handleDashProductChange(false);
    document.getElementById('swatchSelectedDisplay').textContent = (data.fType === 'image' && data.swatchName) ? data.swatchName : 'Frame';

    if(data.swatchDataUrl && data.fType === 'image') { dashActiveImageObj.src = data.swatchDataUrl; document.getElementById('swatchThumbPreview').style.backgroundImage = `url(${data.swatchDataUrl})`; } 
    else { dashActiveImageObj.src = emptyImgUrl; document.getElementById('swatchThumbPreview').style.backgroundImage = `none`; }
    updateDashVisualsFromDOM();
}

function dashHtIn(idx, field, val) {
    let row = dashProjectData[idx];
    if(['qty','extW','extH','fW','m1T','m1R','m1B','m1L','m2','m_bleed','canvasDepth','canvasWrap'].includes(field)) val = parseFloat(val) || 0;
    if (field === 'id') { const oldId = row.id; elevations.forEach(elev => { elev.frames.forEach(f => { if (f.id === oldId) f.id = val; }); }); }
    row[field] = val;

    if (idx === dashSelectedRowIndex) {
        const map = { 'id':'m_itemCode', 'imageCode':'m_imageCode', 'level':'m_level', 'qty':'m_qty', 'location':'m_location', 'extW':'extW', 'extH':'extH', 'fCode':'m_fCode', 'fW':'fW', 'canvasDepth':'canvasDepth', 'canvasWrap':'canvasWrap', 'm1ColorName':'m1_color', 'm1ColorHex':'m1_colorHex', 'm2ColorName':'m2_color', 'm2ColorHex':'m2_colorHex', 'm1T':'m1T', 'm1R':'m1R', 'm1B':'m1B', 'm1L':'m1L', 'm2':'m2', 'glass':'m_glass', 'hardware':'m_hardware', 'backing':'m_backing', 'mount':'m_mount', 'notes':'m_notes', 'prodNotes':'m_prodNotes' };
        if(map[field] && document.getElementById(map[field])) document.getElementById(map[field]).value = field.includes('Color') ? val : dashFmt(row[field]);
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

    const isColor = getStr('fType') === 'color';
    const isLinked = document.getElementById('matLinkBtn').classList.contains('active');
    
    let m1Name = getStr('m1_color'); let m1Hex = getStr('m1_colorHex');
    let m2Name = isLinked ? m1Name : getStr('m2_color'); let m2Hex = isLinked ? m1Hex : getStr('m2_colorHex');
    
    if (isLinked) { document.getElementById('m2_color').value = m2Name; document.getElementById('m2_colorHex').value = m2Hex; }

    const m1Active = document.getElementById('m1Toggle').classList.contains('active');
    // M2 can never be active if M1 is off (M2 sits inside M1).
    const m2Active = m1Active && document.getElementById('m2Toggle').classList.contains('active');

    dashProjectData[dashSelectedRowIndex] = {
        id: newId, imageCode: getStr('m_imageCode'), level: getStr('m_level'), qty: getVal('m_qty'), product: getStr('m_product'), location: getStr('m_location'),
        bleed: getVal('m_bleed'), canvasDepth: getRaw('canvasDepth'), canvasWrap: getRaw('canvasWrap'),
        extW: getVal('extW'), extH: getVal('extH'), fType: getStr('fType'), fW: getVal('fW'), fColor: getStr('fColor'), fCode: getStr('m_fCode'),
        swatchDataUrl: isColor ? "" : row.swatchDataUrl, swatchName: isColor ? "" : row.swatchName,
        m1A: m1Active, 
        m1T: getVal('m1T'), m1B: getVal('m1B'), m1L: getVal('m1L'), m1R: getVal('m1R'), m1Locked: document.getElementById('m1Lock').classList.contains('active'), 
        m1ColorName: m1Name, m1ColorHex: m1Hex,
        m2A: m2Active, m2: getVal('m2'),
        m2ColorName: m2Name, m2ColorHex: m2Hex, matsLinked: isLinked,
        glass: getStr('m_glass'), hardware: getStr('m_hardware'), mount: getStr('m_mount'), backing: getStr('m_backing'), notes: getStr('m_notes'), prodNotes: getStr('m_prodNotes')
    };
    
    updateDashVisualsFromDOM(); renderDashTable(); pushUpdatesToElevations(dashSelectedRowIndex);
    if (oldId !== newId) recalculateDashboardQuantities();
}

function updateDashVisualsFromDOM() {
    const data = dashProjectData[dashSelectedRowIndex];
    const fVis = document.getElementById('dash-frame-visual');
    const viewObj = document.getElementById('view-dashboard');
    
    // Both Library and Color rows stay visible; the inactive side dims.
    // (Previously imageControls was display:none in Color mode, which felt jarring.)
    document.getElementById('imageControls').style.display = 'flex';
    applyFrameStyleDimming(data.fType);
    if (data.fType === 'color') {
        viewObj.style.setProperty('--frame-bg', `none`);
        document.getElementById('swatchSelectedDisplay').textContent = "Frame";
        document.getElementById('swatchThumbPreview').style.backgroundImage = `none`;
    } else {
        viewObj.style.setProperty('--frame-bg', `url(${data.swatchDataUrl})`);
    }

    const isCanvas = (data.product === "Framed Canvas (Floater)");
    
    // Strict Boolean Enforcment to clear dead visual layers
    const effM1A = (data.m1A !== false && !isCanvas);
    const effM2A = (data.m2A === true && !isCanvas);
    
    const effM1T = effM1A ? data.m1T : 0; const effM1B = effM1A ? data.m1B : 0; const effM1L = effM1A ? data.m1L : 0; const effM1R = effM1A ? data.m1R : 0;
    const effM2 = effM2A ? data.m2 : 0;
    
    const mat1Color = data.m1ColorHex || '#ffffff';
    const mat2Color = data.m2ColorHex || '#ffffff';

    const finalW = data.extW - (data.fW * 2) - effM1L - effM1R - (effM2 * 2); const finalH = data.extH - (data.fW * 2) - effM1T - effM1B - (effM2 * 2);
    
    document.getElementById('disp_openW').innerText = dashFmt(Math.max(0, finalW));
    document.getElementById('disp_openH').innerText = dashFmt(Math.max(0, finalH));
    document.getElementById('printFileDisplay').innerText = `${dashFmt(Math.max(0, finalW) + (data.bleed * 2))} x ${dashFmt(Math.max(0, finalH) + (data.bleed * 2))}`;

    const ratio = 300 / Math.max(data.extW, data.extH); 
    
    fVis.innerHTML = ''; 
    fVis.style.width = (data.extW * ratio) + "px"; fVis.style.height = (data.extH * ratio) + "px";

    if (data.fType === 'color') {
        fVis.className = 'frame-vis frame-vis-solid';
        fVis.style.border = `${data.fW * ratio}px solid ${data.fColor}`;
        viewObj.style.setProperty('--frame-color', data.fColor);
    } else {
        fVis.className = 'frame-vis frame-vis-image';
        fVis.style.border = 'none';
        viewObj.style.setProperty('--fW', (data.fW * ratio) + 'px'); 
        viewObj.style.setProperty('--frame-W', (data.extW * ratio) + 'px');
        const rails = ['top', 'bottom', 'left', 'right'];
        rails.forEach(pos => {
            const rail = document.createElement('div'); rail.className = `frame-rail rail-${pos}`; rail.innerHTML = `<div class="rail-bg"></div>`; fVis.appendChild(rail);
        });
    }

    let offsetW = (data.fType === 'color') ? 0 : (data.fW * ratio);

    if(effM1A) { 
        const m1Vis = document.createElement('div'); m1Vis.className = 'mat-visual'; m1Vis.id = 'dash-mat1-visual';
        m1Vis.style.top = offsetW + "px"; m1Vis.style.left = offsetW + "px"; 
        m1Vis.style.width = ((data.extW - (data.fW * 2)) * ratio) + "px"; m1Vis.style.height = ((data.extH - (data.fW * 2)) * ratio) + "px"; 
        m1Vis.style.borderTopWidth = (data.m1T * ratio) + 'px'; m1Vis.style.borderBottomWidth = (data.m1B * ratio) + 'px';
        m1Vis.style.borderLeftWidth = (data.m1L * ratio) + 'px'; m1Vis.style.borderRightWidth = (data.m1R * ratio) + 'px';
        m1Vis.style.borderColor = mat1Color; viewObj.style.setProperty('--m1-color', mat1Color); fVis.appendChild(m1Vis);
    }
    
    if(effM2A) { 
        const m2Vis = document.createElement('div'); m2Vis.className = 'mat2-visual'; m2Vis.id = 'dash-mat2-visual';
        let m2TopOffset = (data.fType === 'color') ? (data.m1T * ratio) : ((data.fW + data.m1T) * ratio);
        let m2LeftOffset = (data.fType === 'color') ? (data.m1L * ratio) : ((data.fW + data.m1L) * ratio);
        m2Vis.style.top = m2TopOffset + "px"; m2Vis.style.left = m2LeftOffset + "px"; 
        m2Vis.style.width = ((data.extW - (data.fW * 2) - effM1L - effM1R) * ratio) + "px"; m2Vis.style.height = ((data.extH - (data.fW * 2) - effM1T - effM1B) * ratio) + "px"; 
        m2Vis.style.borderWidth = (data.m2 * ratio) + 'px'; m2Vis.style.borderColor = mat2Color; viewObj.style.setProperty('--m2-color', mat2Color); fVis.appendChild(m2Vis);
    }
    
    const artVis = document.createElement('div'); artVis.className = 'art-visual'; artVis.id = 'dash-art-visual';
    artVis.style.border = isCanvas ? "4px solid #111" : "1px solid #aaa";
    
    let artTopOffset = (data.fType === 'color') ? ((effM1T + effM2) * ratio) : ((data.fW + effM1T + effM2) * ratio);
    let artLeftOffset = (data.fType === 'color') ? ((effM1L + effM2) * ratio) : ((data.fW + effM1L + effM2) * ratio);
    
    artVis.style.top = artTopOffset + "px"; artVis.style.left = artLeftOffset + "px";
    artVis.style.width = (Math.max(0, finalW) * ratio) + "px"; artVis.style.height = (Math.max(0, finalH) * ratio) + "px";
    artVis.innerText = `${dashFmt(Math.max(0, finalW))}${dashUnit === 'in' ? '"' : ' cm'} × ${dashFmt(Math.max(0, finalH))}${dashUnit === 'in' ? '"' : ' cm'}`;
    fVis.appendChild(artVis);
}

// CSV IMPORT PARSER
function importDashCSV(e) {
    const f = e.target.files[0]; if(!f) return;
    const r = new FileReader();
    r.onload = ev => {
        const text = ev.target.result;
        const lines = text.split('\n');
        let dataStartIndex = -1;
        for(let i=0; i<lines.length; i++) {
            if (lines[i].includes('LEVEL') && lines[i].includes('Item Code')) { dataStartIndex = i + 1; break; }
        }
        if(dataStartIndex === -1) return alert("Invalid CSV format. Cannot find column headers.");
        
        const parseCSVRow = (row) => {
            const matches = row.match(/(\s*"[^"]+"\s*|\s*[^,]+|,)(?=,|$)/g);
            if(!matches) return [];
            return matches.map(m => m.replace(/^,/, '').replace(/^"|"$/g, '').trim());
        };
        
        const newData = [];
        for(let i=dataStartIndex; i<lines.length; i++) {
            if(!lines[i].trim()) continue;
            const cols = parseCSVRow(lines[i]);
            if(cols.length < 5) continue;

            const d = JSON.parse(JSON.stringify(dashDefaultData));
            d.level = cols[0] || ''; d.qty = parseInt(cols[1]) || 0; d.id = cols[2] || generateNextItemCode();
            d.product = cols[3] || 'Framed Art'; d.location = cols[4] || ''; d.imageCode = cols[5] || '';
            d.extW = parseFloat(cols[6]) || 24; d.extH = parseFloat(cols[7]) || 24;
            d.canvasDepth = cols[12] && cols[12] !== "N/A" ? parseFloat(cols[12]) : "";
            d.canvasWrap = cols[13] && cols[13] !== "N/A" ? parseFloat(cols[13]) : "";
            
            d.m1ColorName = cols[14] !== "None" ? cols[14] : "B 97 White"; d.m1ColorHex = cols[15] !== "None" ? cols[15] : "#ffffff";
            d.m1A = (cols[16] !== "None" && !isNaN(parseFloat(cols[16])));
            d.m1T = d.m1A ? parseFloat(cols[16]) : 0;
            d.m1B = d.m1A ? parseFloat(cols[17]) : 0;
            d.m1L = d.m1A ? parseFloat(cols[18]) : 0;
            d.m1R = d.m1A ? parseFloat(cols[19]) : 0;
            
            d.m2ColorName = cols[20] !== "None" ? cols[20] : "B 97 White"; d.m2ColorHex = cols[21] !== "None" ? cols[21] : "#ffffff";
            d.m2A = (cols[22] !== "None" && !isNaN(parseFloat(cols[22])));
            d.m2 = d.m2A ? parseFloat(cols[22]) : 0;
            
            d.glass = cols[23] || ''; d.fCode = cols[24] || ''; d.fW = parseFloat(cols[25]) || 1.25;
            d.hardware = cols[26] || ''; d.backing = cols[27] || ''; d.mount = cols[28] || '';
            d.notes = cols[29] || ''; d.prodNotes = cols[30] || '';
            
            newData.push(d);
        }
        if(newData.length > 0) {
            dashProjectData = newData; dashSelectedRowIndex = 0; loadDashDataIntoControls(dashProjectData[dashSelectedRowIndex]);
            renderDashTable(); elevations.forEach(elev => elev.frames = []); recalculateDashboardQuantities(); alert(`Imported ${newData.length} items successfully.`);
        }
    };
    r.readAsText(f); e.target.value = '';
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
    
    if (id === 'm1') {
        const m1On = b.classList.contains('active');
        document.querySelectorAll('.m1-input').forEach(e => e.disabled = !m1On);
        // Mat 2 is nested inside Mat 1. If Mat 1 turns off, force Mat 2 off too,
        // and disable the M2 toggle so the user can't enable an orphaned mat.
        const m2Btn = document.getElementById('m2Toggle');
        if (!m1On && m2Btn.classList.contains('active')) {
            m2Btn.classList.remove('active');
            m2Btn.innerText = 'OFF';
            document.getElementById('m2').disabled = true;
        }
        m2Btn.disabled = !m1On;
        m2Btn.style.opacity = m1On ? '1' : '0.4';
        m2Btn.style.cursor = m1On ? 'pointer' : 'not-allowed';
    } else {
        document.getElementById('m2').disabled = !b.classList.contains('active');
    }
    syncDashAndCalculate();
}

function toggleDashLock() {
    const b = document.getElementById('m1Lock');
    b.classList.toggle('active'); b.innerText = b.classList.contains('active') ? 'LOCKED' : 'UNLOCKED';
    handleDashMatSync('m1T');
}

// Frame Style toggle (Library / Color). Writes through to the hidden #fType input
// so the existing syncDashAndCalculate plumbing keeps working unchanged.
function setFrameStyle(val) {
    document.getElementById('fType').value = val;
    document.getElementById('fTypeBtnLibrary').classList.toggle('active', val === 'image');
    document.getElementById('fTypeBtnSolid').classList.toggle('active', val === 'color');
    applyFrameStyleDimming(val);
    syncDashAndCalculate();
}

// Dim the accessory controls on the side that's NOT in use:
//   Library mode → dim the color swatch
//   Color mode   → dim the folder icon AND the library row (vendor/collection/swatch/upload)
function applyFrameStyleDimming(val) {
    const libSide = (val !== 'image'); // dim library accessories when NOT in library mode
    const colSide = (val !== 'color'); // dim color swatch when NOT in color mode
    document.getElementById('libFolderBtn').classList.toggle('fstyle-disabled', libSide);
    document.getElementById('fColor').classList.toggle('fstyle-disabled', colSide);
    // The vendor/collection/frame/upload row is also library-side
    const imageControls = document.getElementById('imageControls');
    if (imageControls) imageControls.classList.toggle('fstyle-disabled', libSide);
}

function toggleMatLink() {
    const b = document.getElementById('matLinkBtn'); 
    b.classList.toggle('active'); 
    b.style.color = b.classList.contains('active') ? 'var(--accent)' : 'var(--text-muted)';
    syncDashAndCalculate();
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

function syncDashLibraryFolder(e) {
    const f = e.target.files;
    if(!f || f.length === 0) return;
    dashLocalLibrary = {}; let c = 0;
    
    for(let file of f) {
        const ext = file.name.split('.').pop().toLowerCase();
        const isImage = file.type.startsWith('image/') || ['jpg', 'jpeg', 'png', 'webp', 'svg'].includes(ext);
        if(!isImage) continue;
        
        const parts = file.webkitRelativePath.split('/');
        const filename = parts[parts.length - 1]; 
        const vendor = parts.length > 2 ? parts[parts.length - 3] : (parts.length > 1 ? parts[parts.length - 2] : parts[0]);
        const collection = parts.length > 2 ? parts[parts.length - 2] : "General";
        
        const baseName = filename.substring(0, filename.lastIndexOf('.'));
        const nP = baseName.split('_');
        const code = nP.slice(0, -1).join(' ') || baseName;
        const w = nP.length > 1 && !isNaN(parseFloat(nP[nP.length - 1])) ? parseFloat(nP[nP.length - 1]) : 1.25;
        
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

// Track object URLs created for dropdown thumbnails so we can revoke them when the
// dropdown is rebuilt or hidden — avoids leaking memory across many list refreshes.
let dashSwatchThumbUrls = [];
function _clearDashSwatchThumbUrls() {
    dashSwatchThumbUrls.forEach(u => { try { URL.revokeObjectURL(u); } catch(e) {} });
    dashSwatchThumbUrls = [];
}

function updateDashCustomSwatchDropdown() {
    const v = document.getElementById('libVendor').value; const c = document.getElementById('libCollection').value;
    const s = document.getElementById('swatchDropdownList'); s.innerHTML = '';
    _clearDashSwatchThumbUrls();
    if(!v || !c || !dashLocalLibrary[v][c]) return;
    dashLocalLibrary[v][c].forEach((i, idx) => {
        const li = document.createElement('li');
        // Build a row: [22x22 swatch preview] [code (width")]
        li.style.cssText = 'display:flex; align-items:center; gap:6px;';
        const thumbUrl = URL.createObjectURL(i.file);
        dashSwatchThumbUrls.push(thumbUrl);
        const thumb = document.createElement('span');
        thumb.style.cssText = `flex-shrink:0; width:22px; height:22px; border-radius:3px; border:1px solid var(--border-color); background-image:url(${thumbUrl}); background-size:cover; background-position:center;`;
        const txt = document.createElement('span');
        txt.style.cssText = 'flex:1; min-width:0; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;';
        txt.textContent = `${i.code} (${i.width}")`;
        li.appendChild(thumb); li.appendChild(txt);
        // Searchable text for the dropdown's existing filter logic, if any
        li.dataset.label = `${i.code} (${i.width}")`;

        li.onmouseenter = () => {
            if(dashTempHoverUrl) URL.revokeObjectURL(dashTempHoverUrl);
            dashTempHoverUrl = URL.createObjectURL(i.file);
            document.getElementById('swatchThumbPreview').style.backgroundImage = `url(${dashTempHoverUrl})`;
        };
        li.onclick = () => { document.getElementById('swatchSelectedDisplay').textContent = li.dataset.label; s.style.display = 'none'; loadDashFromCustomLibrary(idx); };
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
        dashProjectData[dashSelectedRowIndex].fType = 'image'; dashProjectData[dashSelectedRowIndex].fW = w; dashProjectData[dashSelectedRowIndex].fCode = item.code;
        dashProjectData[dashSelectedRowIndex].swatchDataUrl = u; dashProjectData[dashSelectedRowIndex].swatchName = item.code;
        document.getElementById('view-dashboard').style.setProperty('--frame-bg', `url(${u})`);
        // Sync the Library/Solid toggle and trigger the redraw via syncDashAndCalculate
        document.getElementById('fType').value = 'image';
        document.getElementById('fTypeBtnLibrary').classList.add('active');
        document.getElementById('fTypeBtnSolid').classList.remove('active');
        dashActiveImageObj.src = u; dashActiveImageObj.onload = () => syncDashAndCalculate();
    };
    r.readAsDataURL(item.file);
}

function loadDashCustomSwatch(e) {
    const f = e.target.files[0]; if(!f) return;
    const n = f.name.split('.')[0]; const r = new FileReader();
    r.onload = e => {
        document.getElementById('fType').value = 'image';
        document.getElementById('fTypeBtnLibrary').classList.add('active');
        document.getElementById('fTypeBtnSolid').classList.remove('active');
        document.getElementById('view-dashboard').style.setProperty('--frame-bg', `url(${e.target.result})`);
        dashProjectData[dashSelectedRowIndex].fType = 'image'; dashProjectData[dashSelectedRowIndex].swatchDataUrl = e.target.result; dashProjectData[dashSelectedRowIndex].swatchName = n;
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

// =========================================================================
// SHARED FRAME CANVAS RENDERER
// Used by both Frame Dashboard PNG export and Elevation PNG export so the
// frames look identical in both. Returns a canvas with the frame drawn,
// fully padded with shadow space.
// =========================================================================
function renderFrameToCanvas(d, swatchImg, opts) {
    opts = opts || {};
    const dpi = opts.dpi || 72;
    const pad = opts.pad !== undefined ? opts.pad : 40;
    const w = d.extW * dpi;
    const h = d.extH * dpi;
    const fw = d.fW * dpi;

    const c = document.createElement('canvas');
    c.width = w + (pad * 2);
    c.height = h + (pad * 2);
    const x = c.getContext('2d');
    x.translate(pad, pad);

    // Outer drop shadow under the whole frame
    x.shadowColor = 'rgba(0,0,0,0.4)'; x.shadowBlur = 25; x.shadowOffsetY = 15;
    x.fillStyle = '#000'; x.fillRect(0, 0, w, h);
    x.shadowColor = 'transparent';

    // Rail fill: either patterned swatch (Library) or solid color
    function fR(img, rw, rh, sx, sy) {
        if (!img || img.src === emptyImgUrl || d.fType === 'color' || !img.complete || !img.naturalWidth) {
            x.fillStyle = d.fColor || '#1a1a1a'; x.fillRect(sx, sy, rw, rh); return;
        }
        const s = rw / img.width;
        const pt = x.createPattern(img, 'repeat');
        const m = new DOMMatrix().translate(sx, sy).scale(s, s);
        pt.setTransform(m); x.fillStyle = pt; x.fillRect(sx, sy, rw, rh);
    }

    // Mitered frame rails (4 trapezoidal clips)
    x.save(); x.beginPath(); x.moveTo(0,0); x.lineTo(fw,fw); x.lineTo(fw,h-fw); x.lineTo(0,h); x.closePath(); x.clip(); fR(swatchImg,fw,h,0,0); x.restore();
    x.save(); x.beginPath(); x.moveTo(w,0); x.lineTo(w-fw,fw); x.lineTo(w-fw,h-fw); x.lineTo(w,h); x.closePath(); x.clip(); x.translate(w,0); x.scale(-1,1); fR(swatchImg,fw,h,0,0); x.restore();
    x.save(); x.beginPath(); x.moveTo(0,0); x.lineTo(w,0); x.lineTo(w-fw,fw); x.lineTo(fw,fw); x.closePath(); x.clip(); x.translate(w/2,fw/2); x.rotate(90*Math.PI/180); x.translate(-fw/2,-w/2); fR(swatchImg,fw,w,0,0); x.restore();
    x.save(); x.beginPath(); x.moveTo(0,h); x.lineTo(w,h); x.lineTo(w-fw,h-fw); x.lineTo(fw,h-fw); x.closePath(); x.clip(); x.translate(w/2,h-fw/2); x.rotate(-90*Math.PI/180); x.translate(-fw/2,-w/2); fR(swatchImg,fw,w,0,0); x.restore();

    // Drop-inset shadow helper for mat & art bevels
    function dIS(bx, by, bw, bh, bl, os, op) {
        x.save(); x.beginPath(); x.rect(bx,by,bw,bh); x.clip();
        x.shadowColor = `rgba(0,0,0,${op})`; x.shadowBlur = bl; x.shadowOffsetY = os;
        x.lineWidth = 20; x.strokeStyle = '#000'; x.strokeRect(bx-10,by-10,bw+20,bh+20);
        x.restore();
    }

    const iX = fw, iY = fw, iW = w - (fw*2), iH = h - (fw*2);
    const isC = (d.product === "Framed Canvas (Floater)");
    // Mat 2 only renders when Mat 1 is also active (M2 sits inside M1)
    const m1On = (d.m1A !== false && !isC);
    const m2On = (m1On && d.m2A === true && !isC);
    const mT = m1On ? d.m1T : 0, mB = m1On ? d.m1B : 0, mL = m1On ? d.m1L : 0, mR = m1On ? d.m1R : 0;
    const m2 = m2On ? d.m2 : 0;
    const mat1Color = d.m1ColorHex || '#ffffff';
    const mat2Color = d.m2ColorHex || '#ffffff';

    if (m1On) {
        x.fillStyle = mat1Color; x.fillRect(iX, iY, iW, iH);
        dIS(iX, iY, iW, iH, 25, 10, 0.45);
        x.strokeStyle = "#cccccc"; x.lineWidth = 1; x.strokeRect(iX, iY, iW, iH);
    }
    const m2X = iX + (mL*dpi), m2Y = iY + (mT*dpi);
    const m2W = iW - ((mL+mR)*dpi), m2H = iH - ((mT+mB)*dpi);
    if (m2On) {
        x.fillStyle = mat2Color; x.fillRect(m2X, m2Y, m2W, m2H);
        dIS(m2X, m2Y, m2W, m2H, 15, 6, 0.35);
        x.strokeStyle = "#cccccc"; x.lineWidth = 1; x.strokeRect(m2X, m2Y, m2W, m2H);
    }

    // Art opening
    const aX = m2X + (m2*dpi), aY = m2Y + (m2*dpi);
    const aW = m2W - (m2*2*dpi), aH = m2H - (m2*2*dpi);
    x.clearRect(aX, aY, aW, aH);

    if (isC) {
        x.strokeStyle = "#111111"; x.lineWidth = 10; x.strokeRect(aX+5, aY+5, aW-10, aH-10);
        dIS(aX, aY, aW, aH, 20, 5, 0.8);
    } else {
        dIS(aX, aY, aW, aH, 8, 3, 0.25);
        x.strokeStyle = "#aaaaaa"; x.lineWidth = 1; x.strokeRect(aX, aY, aW, aH);
    }

    // Optional art-opening size label (for elevation export — dashboard already shows this elsewhere)
    if (opts.showArtLabel && aW > 0 && aH > 0) {
        const unitSuffix = (opts.unit || 'in') === 'in' ? '"' : ' cm';
        const fontSize = Math.max(10, Math.min(aW, aH) * 0.08);
        x.fillStyle = 'rgba(60,60,60,0.65)';
        x.font = `bold ${fontSize}px system-ui, -apple-system, sans-serif`;
        x.textAlign = 'center'; x.textBaseline = 'middle';
        const cx = aX + aW/2, cy = aY + aH/2;
        const lineH = fontSize * 1.15;
        x.fillText(`${(d.extW - d.fW*2 - mL - mR - m2*2).toFixed(1)}${unitSuffix}`, cx, cy - lineH);
        x.fillText(`x`, cx, cy);
        x.fillText(`${(d.extH - d.fW*2 - mT - mB - m2*2).toFixed(1)}${unitSuffix}`, cx, cy + lineH);
    }

    return { canvas: c, pad: pad, frameW: w, frameH: h };
}

function exportDashNativePNG() {
    const d = dashProjectData[dashSelectedRowIndex];
    const { canvas } = renderFrameToCanvas(d, dashActiveImageObj, { dpi: 72, pad: 40 });
    const a = document.createElement('a');
    a.download = `${d.id || 'Frame'}.png`;
    a.href = canvas.toDataURL("image/png");
    a.click();
}

function exportDashCSV() {
    const g = (id) => document.getElementById(id).value; const u = ` (${dashUnit})`;
    let csv = `,RFI,PROJECT NAME,${g('g_projName')},,,,,,,,,,,,,,,,,,,,,,,,,,,,,,\n,,DESCRIPTION,${g('g_desc')},,,,,,,,,,,,,,,,,,,,,,,,,,,,,,\n,,DATE,${g('g_date')},,,,,,,,,,,,,,,,,,,,,,,,,,,,,,\n,,ISSUED BY,${g('g_issued')},,,,,,,,,,,,,,,,,,,,,,,,,,,,,,\n,,CLIENT NAME,${g('g_client')},,,,,,,,,,,,,,,,,,,,,,,,,,,,,,\n,,Attn:,${g('g_attn')},,,,,,,,,,,,,,,,,,,,,,,,,,,,,,\n,,Delivery,${g('g_delivery')},,,,,,,,,,,,,,,,,,,,,,,,,,,,,,\n\nLEVEL,Qty,Item Code,PRODUCT,LOCATION,Image code,Overall Width${u},Overall Height${u},Art Size W${u},Art Size H${u},Print File W${u},Print File H${u},Canvas Stretcher Depth${u},Canvas Image Wrap${u},Mat 1 Name,Mat 1 Hex,Mat Top,Mat Bottom,Mat Left,Mat Right,Mat 2 Name,Mat 2 Hex,Mat 2 Reveal,Glass,Frame Code,Frame (Width)${u},Security Hardware,Backing/Substrate,Mount,Notes,Production Notes,Image_Filename\n`;
    dashProjectData.forEach(r => {
        const iC = (r.product === "Framed Canvas (Floater)");
        const mT = (r.m1A !== false && !iC) ? r.m1T : 0; const mB = (r.m1A !== false && !isC) ? r.m1B : 0; const mL = (r.m1A !== false && !iC) ? r.m1L : 0; const mR = (r.m1A !== false && !iC) ? r.m1R : 0; const m2 = (r.m2A && !iC) ? r.m2 : 0;
        const fW = r.extW - (r.fW*2) - mL - mR - (m2*2); const fH = r.extH - (r.fW*2) - mT - mB - (m2*2);
        const iW = iC ? "N/A" : dashFmt(Math.max(0, fW) + (r.bleed*2)); const iH = iC ? "N/A" : dashFmt(Math.max(0, fH) + (r.bleed*2));
        const d = [r.level, r.qty, r.id, r.product, r.location, r.imageCode, dashFmt(r.extW), dashFmt(r.extH), dashFmt(Math.max(0,fW)), dashFmt(Math.max(0,fH)), iW, iH, r.canvasDepth ? dashFmt(r.canvasDepth) : "N/A", r.canvasWrap ? dashFmt(r.canvasWrap) : "N/A", (r.m1A !== false && !iC) ? r.m1ColorName : "None", (r.m1A !== false && !iC) ? r.m1ColorHex : "None", (r.m1A !== false && !iC) ? dashFmt(r.m1T) : "None", (r.m1A !== false && !iC) ? dashFmt(r.m1B) : "None", (r.m1A !== false && !iC) ? dashFmt(r.m1L) : "None", (r.m1A !== false && !iC) ? dashFmt(r.m1R) : "None", (r.m2A && !iC) ? r.m2ColorName : "None", (r.m2A && !iC) ? r.m2ColorHex : "None", (r.m2A && !iC) ? dashFmt(r.m2) : "None", r.glass, r.fCode, dashFmt(r.fW), r.hardware, r.backing, r.mount, r.notes, r.prodNotes, `${r.id}.png`];
        csv += d.map(s => `"${String(s).replace(/"/g, '""')}"`).join(',') + '\n';
    });
    const b = new Blob([csv], { type: 'text/csv;charset=utf-8;' }); const a = document.createElement("a"); a.href = URL.createObjectURL(b); a.download = `RFI_Project_Tracker.csv`; a.click();
}

function renderDashTable() {
    const tbody = document.getElementById('rfiBody'); tbody.innerHTML = '';
    dashProjectData.forEach((row, index) => {
        const isCanvas = (row.product === "Framed Canvas (Floater)");
        const m1T = (row.m1A !== false && !isCanvas) ? row.m1T : 0; const m1B = (row.m1A !== false && !isCanvas) ? row.m1B : 0; const m1L = (row.m1A !== false && !isCanvas) ? row.m1L : 0; const m1R = (row.m1A !== false && !isCanvas) ? row.m1R : 0;
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
            <td id="calc-openW-${index}" style="padding: 4px 8px; color:var(--accent); font-weight:bold;">${dashFmt(Math.max(0, finalW))}</td><td id="calc-openH-${index}" style="padding: 4px 8px; color:var(--accent); font-weight:bold;">${dashFmt(Math.max(0, finalH))}</td><td id="calc-printW-${index}" style="color:var(--accent); font-weight:bold; padding: 4px 8px;">${imgW}</td><td id="calc-printH-${index}" style="color:var(--accent); font-weight:bold; padding: 4px 8px;">${imgH}</td>
            <td><input class="tbl-in" type="number" step="0.125" value="${row.canvasDepth || ''}" oninput="dashHtIn(${index}, 'canvasDepth', this.value)" style="width:45px;"></td>
            <td><input class="tbl-in" type="number" step="0.125" value="${row.canvasWrap || ''}" oninput="dashHtIn(${index}, 'canvasWrap', this.value)" style="width:45px;"></td>
            <td><input class="tbl-in" type="text" value="${row.m1ColorName}" oninput="dashHtIn(${index}, 'm1ColorName', this.value)" style="width:80px;"></td>
            <td><input type="color" value="${row.m1ColorHex}" disabled style="width:24px; border:none; padding:0; background:transparent;"></td>
            <td><input class="tbl-in" type="number" step="0.125" value="${dashFmt(row.m1T)}" oninput="dashHtIn(${index}, 'm1T', this.value)" style="width:40px;"></td><td><input class="tbl-in" type="number" step="0.125" value="${dashFmt(row.m1B)}" oninput="dashHtIn(${index}, 'm1B', this.value)" style="width:40px;"></td><td><input class="tbl-in" type="number" step="0.125" value="${dashFmt(row.m1L)}" oninput="dashHtIn(${index}, 'm1L', this.value)" style="width:40px;"></td><td><input class="tbl-in" type="number" step="0.125" value="${dashFmt(row.m1R)}" oninput="dashHtIn(${index}, 'm1R', this.value)" style="width:40px;"></td>
            <td><input class="tbl-in" type="text" value="${row.m2ColorName}" oninput="dashHtIn(${index}, 'm2ColorName', this.value)" style="width:80px;"></td>
            <td><input type="color" value="${row.m2ColorHex}" disabled style="width:24px; border:none; padding:0; background:transparent;"></td>
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
function updateDimFontSize() {
    const size = document.getElementById('elevDimFontSize').value || 12;
    document.documentElement.style.setProperty('--dim-font-size', size + 'px');
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
    e.stopPropagation(); 
    pendingDuplicateIndex = idx;
    document.getElementById('duplicateModal').style.display = 'flex';
}

function closeDuplicateModal() {
    document.getElementById('duplicateModal').style.display = 'none';
    pendingDuplicateIndex = null;
}

function confirmDuplicate(type) {
    if (pendingDuplicateIndex === null) return;
    const idx = pendingDuplicateIndex;
    
    const temp = elevFrames[idx]; 
    const nF = JSON.parse(JSON.stringify(temp));
    nF.letter = getElevLetter(elevFrames.length); 
    const moveFactor = elevUnit === 'in' ? 10 : 25.4;
    nF.x = temp.x + moveFactor; nF.y = temp.y - moveFactor; 

    if(type === 'new') {
        const dashSrc = dashProjectData.find(d => d.id === temp.id);
        const newDash = JSON.parse(JSON.stringify(dashSrc || dashDefaultData));
        newDash.id = generateNextItemCode();
        newDash.qty = 0;
        dashProjectData.push(newDash);
        nF.id = newDash.id;
    }

    elevFrames.push(nF); 
    initElevControls(); drawElevAll(); recalculateDashboardQuantities(); 
    closeDuplicateModal();
}

function toggleElevLayer(id, btn) {
    const layer = document.getElementById(id);
    const isHidden = (layer.style.display === 'none' || layer.style.display === '');
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
    elevFrames.forEach(f => f.isGrouped = !anyGrouped);
    btn.classList.toggle('active', !anyGrouped);
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

    const personHeightIn = 72; 
    const personHeight = elevUnit === 'in' ? personHeightIn : parseFloat((personHeightIn * 2.54).toFixed(2)); 
    const pWrap = document.getElementById('person-wrap');
    document.getElementById('person').style.height = (personHeight * elevScale) + 'px';
    pWrap.style.left = (elevPersonPos.x * elevScale) + 'px';

    const frameLayer = document.getElementById('frame-layer'); frameLayer.innerHTML = '';
    const labelLayer = document.getElementById('label-layer'); labelLayer.innerHTML = '';
    const odLayer = document.getElementById('od-layer'); odLayer.innerHTML = '';
    const centerLayer = document.getElementById('frame-center-layer'); centerLayer.innerHTML = '';
    
    elevFrames.forEach((f, idx) => {
        if(!f.active) return;
        
        const el = document.createElement('div'); el.className = 'draggable frame-vis';
        el.style.cssText = `width:${f.w*elevScale}px; height:${f.h*elevScale}px; left:${f.x*elevScale}px; bottom:${f.y*elevScale}px;`;
        
        if (f.fType === 'color') {
            el.classList.add('frame-vis-solid');
            el.style.border = `${f.fW * elevScale}px solid ${f.fColor || '#1a1a1a'}`;
            el.style.setProperty('--frame-color', f.fColor || '#1a1a1a');
            // Dynamically scale shadows identically to the CSS engine for perfect export matching
            el.style.boxShadow = `0 0 0 1.5px ${f.fColor || '#1a1a1a'}, 0 ${10 * elevScale}px ${30 * elevScale}px rgba(0,0,0,0.3), 0 ${4 * elevScale}px ${8 * elevScale}px rgba(0,0,0,0.2)`;
        } else {
            el.classList.add('frame-vis-image');
            el.style.setProperty('--fW', (f.fW * elevScale) + 'px');
            el.style.setProperty('--frame-W', (f.w * elevScale) + 'px');
            el.style.setProperty('--frame-bg', `url(${f.swatchDataUrl})`);
            el.style.boxShadow = `0 ${10 * elevScale}px ${30 * elevScale}px rgba(0,0,0,0.3), 0 ${4 * elevScale}px ${8 * elevScale}px rgba(0,0,0,0.2)`;
            const rails = ['top', 'bottom', 'left', 'right'];
            rails.forEach(pos => {
                const rail = document.createElement('div'); rail.className = `frame-rail rail-${pos}`; rail.innerHTML = `<div class="rail-bg"></div>`; el.appendChild(rail);
            });
        }

        let offsetW = (f.fType === 'color') ? 0 : (f.fW * elevScale);

        // Mat 2 is nested inside Mat 1. If Mat 1 is off, Mat 2 is implicitly off too.
        const m1Active = f.m1A !== false;
        const m2Active = m1Active && f.m2A;

        if (m1Active) {
            const m1 = document.createElement('div'); m1.className = 'mat-visual';
            const m1Color = f.m1ColorHex || '#ffffff';
            m1.style.cssText = `top:${offsetW}px; left:${offsetW}px; width:${(f.w - f.fW*2)*elevScale}px; height:${(f.h - f.fW*2)*elevScale}px; border-top-width:${f.m1T*elevScale}px; border-bottom-width:${f.m1B*elevScale}px; border-left-width:${f.m1L*elevScale}px; border-right-width:${f.m1R*elevScale}px;`;
            // Explicit borderColor — must NOT depend on inherited theme color
            m1.style.borderColor = m1Color;
            el.style.setProperty('--m1-color', m1Color);
            m1.style.boxShadow = `0 0 0 1.5px ${m1Color}, inset 0 ${2 * elevScale}px ${6 * elevScale}px rgba(0,0,0,0.3), 0 ${2 * elevScale}px ${5 * elevScale}px rgba(0,0,0,0.15)`;
            el.appendChild(m1);
        }
        
        let m2Val = m2Active ? f.m2 : 0;
        if (m2Active) {
            const m2 = document.createElement('div'); m2.className = 'mat2-visual';
            const m2Color = f.m2ColorHex || '#ffffff';
            let m2TopOffset = (f.fType === 'color') ? (f.m1T * elevScale) : ((f.fW + f.m1T) * elevScale);
            let m2LeftOffset = (f.fType === 'color') ? (f.m1L * elevScale) : ((f.fW + f.m1L) * elevScale);
            m2.style.cssText = `top:${m2TopOffset}px; left:${m2LeftOffset}px; width:${(f.w - f.fW*2 - f.m1L - f.m1R)*elevScale}px; height:${(f.h - f.fW*2 - f.m1T - f.m1B)*elevScale}px; border-width:${m2Val*elevScale}px;`;
            // Explicit borderColor — must NOT depend on inherited theme color
            m2.style.borderColor = m2Color;
            el.style.setProperty('--m2-color', m2Color);
            m2.style.boxShadow = `0 0 0 1.5px ${m2Color}, inset 0 ${2 * elevScale}px ${6 * elevScale}px rgba(0,0,0,0.3), 0 ${2 * elevScale}px ${5 * elevScale}px rgba(0,0,0,0.15)`;
            el.appendChild(m2);
        }
        
        // Use effective values — when Mat 1 is off, mat dimensions don't push the art inward
        const effM1T = m1Active ? f.m1T : 0; const effM1B = m1Active ? f.m1B : 0;
        const effM1L = m1Active ? f.m1L : 0; const effM1R = m1Active ? f.m1R : 0;

        const artW = f.w - f.fW*2 - effM1L - effM1R - m2Val*2; const artH = f.h - f.fW*2 - effM1T - effM1B - m2Val*2;
        const art = document.createElement('div'); art.className = 'art-visual';
        
        let artTopOffset = (f.fType === 'color') ? ((effM1T + m2Val) * elevScale) : ((f.fW + effM1T + m2Val) * elevScale);
        let artLeftOffset = (f.fType === 'color') ? ((effM1L + m2Val) * elevScale) : ((f.fW + effM1L + m2Val) * elevScale);
        
        art.style.cssText = `top:${artTopOffset}px; left:${artLeftOffset}px; width:${artW*elevScale}px; height:${artH*elevScale}px;`;
        art.style.boxShadow = `inset 0 ${2 * elevScale}px ${8 * elevScale}px rgba(0,0,0,0.2)`;
        
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
        const chPad = elevUnit === 'in' ? 6 : 15.24;
        const chHalf = elevUnit === 'in' ? 3 : 7.62;
        crossH.style.width = ((f.w + chPad) * elevScale) + 'px'; crossH.style.left = ((f.x - chHalf) * elevScale) + 'px'; crossH.style.bottom = ((f.y + f.h/2) * elevScale) + 'px';
        const crossV = document.createElement('div'); crossV.className = 'crosshair-v';
        crossV.style.height = ((f.h + chPad) * elevScale) + 'px'; crossV.style.left = ((f.x + f.w/2) * elevScale) + 'px'; crossV.style.bottom = ((f.y - chHalf) * elevScale) + 'px';
        centerLayer.appendChild(crossH); centerLayer.appendChild(crossV);
        
        makeElevDraggable(el, idx); frameLayer.appendChild(el);
    });

    makeElevDraggable(pWrap, 'person');
    drawElevTargetedSpacing(); drawElevGuides(wallW, wallH);
}

function makeElevDraggable(el, idx) {
    el.onmousedown = function(e) {
        if (e.target.tagName === 'INPUT' || e.target.tagName === 'BUTTON') return;
        e.preventDefault(); let sx = e.clientX, sy = e.clientY;
        document.onmousemove = function(e) {
            let dx = (sx - e.clientX)/elevScale, dy = (sy - e.clientY)/elevScale; sx = e.clientX; sy = e.clientY;
            const snap = elevUnit === 'in' ? 1 : 2.54;
            if(idx === 'person') { 
                elevPersonPos.x -= dx; 
            } else { 
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
    
    const offsetDist = elevUnit === 'in' ? 6 : 15.24;
    createElevArchDim(0, wallH + offsetDist, wallW, wallH + offsetDist, 'h', `${elevFmt(wallW)}${elevUnit === 'in' ? '"' : ' cm'}`, archLayer, true);
    createElevArchDim(-offsetDist, 0, -offsetDist, wallH, 'v', `${elevFmt(wallH)}${elevUnit === 'in' ? '"' : ' cm'}`, archLayer, true);

    const pHeight = elevUnit === 'in' ? 72 : parseFloat((72 * 2.54).toFixed(2));
    const pLabel = elevUnit === 'in' ? `72"` : `182.9 cm`;
    const pX = elevPersonPos.x - (elevUnit === 'in' ? 4 : 10.16); 
    createElevArchDim(pX, 0, pX, pHeight, 'v', pLabel, archLayer, true);
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

    const offset = (elevUnit === 'in' ? 6 : 15.24) * elevScale; 

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
            <span class="arch-label-new">${label}</span>
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
        dim.innerHTML = `<div class="dim-line-segment-v"></div><span class="arch-label-new">${label}</span><div class="dim-line-segment-v"></div>`;
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

// Helper: load an image from a data URL and resolve when ready (or reject on error)
function _loadImg(dataUrl) {
    return new Promise((resolve) => {
        if (!dataUrl) { resolve(null); return; }
        const img = new Image();
        img.onload = () => resolve(img);
        img.onerror = () => resolve(null);
        img.src = dataUrl;
    });
}

// html2canvas sometimes fails to capture <img src="*.svg"> when the SVG is fetched
// at render time (it can drop silently due to CORS or timing). Preconvert the
// person SVG to a base64 data URL once and cache it; we swap the img's src to
// the data URL during export so html2canvas sees an inline image it can rasterize.
let _personSvgDataUrl = null;
async function _getPersonSvgDataUrl() {
    if (_personSvgDataUrl) return _personSvgDataUrl;
    const personImg = document.getElementById('person');
    if (!personImg) return null;
    const src = personImg.getAttribute('src');
    if (!src || src.startsWith('data:')) { _personSvgDataUrl = src; return _personSvgDataUrl; }
    try {
        const res = await fetch(src);
        if (!res.ok) return null;
        const ct = res.headers.get('content-type') || 'image/svg+xml';
        const text = await res.text();
        // base64-encode the SVG text safely (handles unicode)
        const b64 = btoa(unescape(encodeURIComponent(text)));
        _personSvgDataUrl = `data:${ct};base64,${b64}`;
        return _personSvgDataUrl;
    } catch (e) {
        console.warn('Could not inline person SVG:', e);
        return null;
    }
}

async function exportElevPNG() {
    const ws = document.querySelector('#view-elevation .workspace');
    const wrap = document.getElementById('export-wrap');
    const wall = document.getElementById('wall');

    // Force light theme for export (consistency with print/PDF), restore after.
    const wasDark = !document.body.classList.contains('light-theme');
    document.body.classList.add('light-theme');
    drawElevAll();
    await new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)));

    const oldOverflow = ws.style.overflow; ws.style.overflow = 'visible';
    const oldWallBg = wall.style.background; wall.style.background = 'transparent';

    // Swap the person's SVG src for an inline data URL so html2canvas can definitely
    // rasterize it (external SVG <img> tags sometimes silently drop on capture).
    const personImg = document.getElementById('person');
    let personOriginalSrc = null;
    if (personImg) {
        const dataUrl = await _getPersonSvgDataUrl();
        if (dataUrl && personImg.src !== dataUrl) {
            personOriginalSrc = personImg.getAttribute('src');
            personImg.src = dataUrl;
            // Wait for the swapped image to be decoded so it's painted before capture
            await new Promise(res => {
                if (personImg.complete && personImg.naturalWidth) return res();
                personImg.onload = () => res();
                personImg.onerror = () => res();
            });
        }
    }

    // Track everything we mutate so we can put it back exactly as it was.
    const frameLayer = document.getElementById('frame-layer');
    const frameDivs = Array.from(frameLayer.children);
    const restoreList = []; // {div, hiddenChildren: [{el, prevDisplay}], overlayCanvas}

    try {
        // Pre-load all swatch images in parallel
        const visibleFrames = elevFrames.filter(f => f.active);
        const imagePromises = visibleFrames.map(f =>
            (f.fType === 'image' && f.swatchDataUrl) ? _loadImg(f.swatchDataUrl) : Promise.resolve(null)
        );
        const loadedImages = await Promise.all(imagePromises);

        // For each visible frame: render a high-res canvas using the same routine
        // as the Frame Dashboard PNG export, then drop it on top of the frame's
        // existing children (which we hide) so html2canvas sees the crisp version.
        let frameIdx = 0;
        frameDivs.forEach(div => {
            // Skip non-frame siblings (shouldn't be any, but defensive)
            if (!div.classList.contains('frame-vis')) return;

            const f = visibleFrames[frameIdx];
            if (!f) return;
            const swatchImg = loadedImages[frameIdx];
            frameIdx++;

            // Elevation frames store dimensions as f.w / f.h (set from d.extW / d.extH at import time).
            // The renderer expects extW / extH (dashboard schema), so adapt here. Bail out if
            // dimensions are zero/missing to avoid a 0-sized canvas (drawImage would throw).
            if (!f.w || !f.h || !f.fW) return;
            const adaptedFrame = Object.assign({}, f, { extW: f.w, extH: f.h });

            // Native render at 72 dpi with NO padding — pad would push the rendered
            // frame outside the container and leak shadow into adjacent frames.
            const { canvas: nativeCanvas } = renderFrameToCanvas(adaptedFrame, swatchImg, {
                dpi: 72,
                pad: 0,
                showArtLabel: true,
                unit: elevUnit,
            });

            // Sanity check the canvas before we try to drawImage from it
            if (!nativeCanvas.width || !nativeCanvas.height) return;

            // Hide the frame's existing children (rails, mats, art-visual)
            const hiddenChildren = [];
            Array.from(div.children).forEach(child => {
                hiddenChildren.push({ el: child, prevDisplay: child.style.display });
                child.style.display = 'none';
            });

            // Add the canvas as a single overlay child filling the container.
            // The native canvas is f.extW*72 by f.extH*72 px; we let CSS scale it
            // to fit the container, which is already sized correctly via elevScale.
            const overlay = document.createElement('canvas');
            overlay.width = nativeCanvas.width;
            overlay.height = nativeCanvas.height;
            overlay.getContext('2d').drawImage(nativeCanvas, 0, 0);
            overlay.style.cssText = `position:absolute; top:0; left:0; width:100%; height:100%; pointer-events:none; display:block;`;
            div.appendChild(overlay);

            // Make sure the container itself doesn't draw a competing border/shadow.
            div.dataset._origBoxShadow = div.style.boxShadow || '';
            div.dataset._origBorder = div.style.border || '';
            div.style.boxShadow = 'none';
            div.style.border = 'none';

            restoreList.push({ div, hiddenChildren, overlayCanvas: overlay });
        });

        // Wait for the browser to commit the DOM changes
        await new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)));

        // Capture the elevation
        const canvas = await html2canvas(wrap, {
            backgroundColor: null,
            scale: 3,
            useCORS: true,
            allowTaint: true,
        });
        const a = document.createElement('a');
        a.download = `${elevations[currentElevIndex].name.replace(/[^a-z0-9]/gi, '_')}.png`;
        a.href = canvas.toDataURL("image/png");
        a.click();
    } catch (err) {
        console.error(err);
        alert("Image Export Failed: " + (err && err.message ? err.message : "Unknown error") +
              "\n\nIf you opened this file directly (file://...), browser security blocks local file access. Please serve the folder via a local web server (e.g. VS Code Live Server) and try again.");
    } finally {
        // Restore every frame we touched
        restoreList.forEach(({ div, hiddenChildren, overlayCanvas }) => {
            if (overlayCanvas && overlayCanvas.parentNode === div) div.removeChild(overlayCanvas);
            hiddenChildren.forEach(({ el, prevDisplay }) => { el.style.display = prevDisplay; });
            div.style.boxShadow = div.dataset._origBoxShadow || '';
            div.style.border = div.dataset._origBorder || '';
            delete div.dataset._origBoxShadow;
            delete div.dataset._origBorder;
        });

        if (wasDark) document.body.classList.remove('light-theme');
        ws.style.overflow = oldOverflow;
        wall.style.background = oldWallBg;
        // Restore the person's original src if we swapped it
        if (personImg && personOriginalSrc !== null) {
            personImg.src = personOriginalSrc;
        }
        // Final clean re-render in restored theme
        drawElevAll();
    }
}

// BOOT UP THE ENGINE
initMasterApp();
