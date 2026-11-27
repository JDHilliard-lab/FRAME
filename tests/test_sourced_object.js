// SOURCED OBJECT IS ITS OWN PRODUCT (20.85).
//
// Asked: "Sourced objects should be an object only which usually have width, height, and
// depth in the measurements, then would include notes on hardware because it is usually
// custom. Also, it should be a dashed line when dropping it into the elevations with the
// opportunity for designers to drop in an image of the object." Then, with the studio's FSO
// tracker: "you have to find pricing, keep track of links with the projects".
//
// Before this it was in the product list with no branch anywhere, so it drew a moulding and
// mats on the wall and printed Frame Code, Mat 1, Glass and Art Dimensions on its spec.
const { JSDOM } = require('jsdom');
const fs = require('fs');
const path = require('path');

(async () => {
  const root = path.join(__dirname, '..');
  const src = fs.readFileSync(path.join(root, 'app.js'), 'utf8');
  const htmlSrc = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
  const dom = new JSDOM(htmlSrc, { url: 'https://example.com/', pretendToBeVisual: true, runScripts: 'outside-only' });
  const { window } = dom;
  // A context that RECORDS what is drawn, so a check can tell a dashed outline from a frame.
  window.__ops = [];
  const noop = () => ctx2d;
  const ctx2d = new Proxy({}, { get: (t, k) => {
    if (k === 'measureText') return () => ({ width: 0 });
    if (k === 'getImageData') return (x, y, w, h) => ({ data: new Uint8ClampedArray(4), width: w, height: h });
    if (k === 'createLinearGradient' || k === 'createRadialGradient' || k === 'createPattern') return () => ({ addColorStop: () => {} });
    if (k === 'canvas') return undefined;
    if (k in t) return t[k];
    return (...a) => { window.__ops.push({ op: String(k), a: a, stroke: t.strokeStyle, fill: t.fillStyle }); return ctx2d; };
  }, set: (t, k, v) => { t[k] = v; return true; } });
  window.HTMLCanvasElement.prototype.getContext = () => ctx2d;
  window.HTMLCanvasElement.prototype.toDataURL = () => 'data:image/png;base64,RASTER';
  window.Element.prototype.getBoundingClientRect = function () { return { left: 10, top: 10, right: 110, bottom: 90, width: 100, height: 80, x: 10, y: 10 }; };
  window.fetch = () => Promise.reject(new Error('no network in test'));
  window.open = (u) => { window.__opened = u; };
  global.window = window; global.document = window.document;
  global.navigator = window.navigator;

  const testBlock = `
    window.__testResults = [];
    const __check = async (label, fn) => { try { await fn(); window.__testResults.push({ label, ok: true }); } catch (e) { window.__testResults.push({ label, ok: false, err: e.message }); } };
    scheduleAutosave = () => {}; _autosaveToldUser = true;
    const OBJ = () => Object.assign(_cloneData(dashDefaultData), { id: 'OBJ.1', product: 'Sourced Object', extW: 12, extH: 30, fW: 1.25,
      m1A: true, m1T: 3, m1B: 3, m1L: 3, m1R: 3, bleed: 0.25, objDepth: 8, hardware: 'Pin mount, 2 pins', qty: 1 });
    const frameOf = (extra) => Object.assign({ id: 'OBJ.1', letter: 'A', product: 'Sourced Object', w: 12, h: 30, x: 40, y: 40, fW: 1.25, fType: 'color', fColor: '#111111',
      m1A: true, m1T: 3, m1B: 3, m1L: 3, m1R: 3, m2: 0, m2A: false, active: true, dimTo: [], distToggles: { ceiling: false, floor: false, left: false, right: false } }, extra || {});
    const wall = (extra) => {
      elevUnit = 'in'; dashUnit = 'in';
      elevations = [{ id: 'w1', name: 'Lobby', wallW: 120, wallH: 100, personPos: { x: -60 }, frames: [frameOf(extra)] }];
      currentElevIndex = 0; elevFrames = elevations[0].frames; elevZoomFactor = 1;
      const ws = document.querySelector('#view-elevation .workspace');
      Object.defineProperty(ws, 'clientWidth', { get: () => 1100, configurable: true });
      Object.defineProperty(ws, 'clientHeight', { get: () => 800, configurable: true });
      drawElevAll();
      return document.querySelector('#frame-layer .frame-vis[data-frame-letter="A"]');
    };
    const capture = async () => {
      const keepS = _frameShadowCanvas, keepP = frameFx.pageShadow;
      _frameShadowCanvas = () => ({ canvas: { toDataURL: () => 'data:image/png;base64,SHADOW' }, padIn: 0.5 });
      frameFx.pageShadow = true;
      try { const res = await exportElevSVG({ returnBlob: true, forCapture: true }); return String(res && res.picSvg || ''); }
      finally { _frameShadowCanvas = keepS; frameFx.pageShadow = keepP; }
    };
    const fakeImg = { naturalWidth: 400, naturalHeight: 1000, width: 400, height: 1000 };
    const setForm = () => {
      dashProjectData = [OBJ()]; dashSelectedRowIndex = 0;
      loadDashDataIntoControls(dashProjectData[0]);
    };
    const shown = (el) => !!el && el.style.display !== 'none';

    window.__run = (async () => {
      // ── the numbers ───────────────────────────────────────────────────────
      await __check('EXACT ASK: the spec is Application, Hardware and W x H x D, nothing framed', async () => {
        dashUnit = 'in';
        const lines = buildSpecStrings(OBJ(), { dualUnit: '' }).lines;
        const labels = lines.map(l => l.label).join('|');
        if (labels !== 'Application|Hardware|Overall Dimensions') throw new Error('rows: ' + labels);
        const od = lines.find(l => l.label === 'Overall Dimensions').value;
        if (od.indexOf('12') < 0 || od.indexOf('30') < 0 || od.indexOf('8') < 0 || !/D$/.test(od)) throw new Error('Overall Dimensions reads ' + od);
        if (lines[0].value !== 'Sourced Object') throw new Error('Application ' + lines[0].value);
        const noDepth = buildSpecStrings(Object.assign(OBJ(), { objDepth: '' }), { dualUnit: '' }).lines.find(l => l.label === 'Overall Dimensions').value;
        if (/D$/.test(noDepth)) throw new Error('a depth was printed with none given: ' + noDepth);
      });
      await __check('opening and file size are the object itself: no mats, no bleed', async () => {
        const z = _rowOpeningAndPrint(OBJ());
        if (z.openW !== 12 || z.openH !== 30 || z.printW !== 12 || z.printH !== 30) throw new Error(JSON.stringify(z));
      });

      // ── the canvas renderer (PNG, spec pages, thumbnails) ──────────────────
      await __check('EXACT ASK: with no photo the object draws a dashed outline, and no frame or shadow', async () => {
        dashOuterShadowsOn = true;
        window.__ops = [];
        renderFrameToCanvas({ product: 'Sourced Object', extW: 12, extH: 30, fW: 1.25, m1A: true, m1T: 3, m1B: 3, m1L: 3, m1R: 3 }, null, { dpi: 72, pad: 0 });
        const ops = window.__ops;
        const dash = ops.find(o => o.op === 'setLineDash' && o.a[0] && o.a[0].length);
        if (!dash) throw new Error('no dashed line: ' + ops.map(o => o.op).join(','));
        if (!ops.some(o => o.op === 'strokeRect')) throw new Error('nothing outlined');
        if (ops.some(o => o.op === 'fillRect')) throw new Error('something was filled (a frame body or its shadow)');
      });
      await __check('EXACT ASK: a dropped photo shows in place of the outline', async () => {
        window.__ops = [];
        renderFrameToCanvas({ product: 'Sourced Object', extW: 12, extH: 30, fW: 1.25 }, null, { dpi: 72, pad: 0, artworkImg: fakeImg });
        const ops = window.__ops.map(o => o.op);
        if (ops.indexOf('drawImage') < 0) throw new Error('the photo was not drawn');
        if (ops.indexOf('strokeRect') >= 0) throw new Error('the outline is still drawn over the photo');
      });
      await __check('a wireframe wall shows the grey block with the dashed outline', async () => {
        window.__ops = [];
        renderFrameToCanvas({ product: 'Sourced Object', extW: 12, extH: 30, fW: 1.25 }, null, { dpi: 72, pad: 0, artworkImg: fakeImg, wireframe: true, wireframeLetter: 'A' });
        const ops = window.__ops;
        if (!ops.some(o => o.op === 'fillRect' && o.fill === ELEV_WF_FILL)) throw new Error('no wireframe block');
        if (ops.some(o => o.op === 'drawImage')) throw new Error('the photo was drawn on a wireframe');
        if (!ops.some(o => o.op === 'strokeRect')) throw new Error('no outline on the wireframe');
      });

      // ── the wall on screen ────────────────────────────────────────────────
      await __check('EXACT ASK: on the wall it is a dashed box with no moulding, mat or shadow', async () => {
        dashOuterShadowsOn = true;
        const el = wall();
        if (!el) throw new Error('no frame drawn');
        if (el.querySelector('.frame-rail, .mat-visual, .mat2-visual')) throw new Error('a rail or mat was drawn');
        if ((el.style.boxShadow || 'none') !== 'none') throw new Error('shadow ' + el.style.boxShadow);
        const art = el.querySelector('.art-visual');
        if (!art || (art.getAttribute('style') || '').indexOf('dashed') < 0) throw new Error('the opening is not dashed: ' + (art && art.getAttribute('style')));
        if (art.style.width !== (12 * elevScale) + 'px') throw new Error('the photo area is not the whole object: ' + art.style.width);
        if (!art.querySelector('.art-add-ph')) throw new Error('nothing says a photo can be dropped on it');
      });
      await __check('with a photo the wall shows the photo and drops the outline', async () => {
        const el = wall({ artworkUrl: 'data:image/png;base64,OBJPHOTO', artworkW: 400, artworkH: 1000 });
        const art = el.querySelector('.art-visual');
        if (!art.querySelector('img.art-img')) throw new Error('no photo');
        if ((art.getAttribute('style') || '').indexOf('dashed') >= 0) throw new Error('outline still on');
      });

      // ── the SVG / PDF capture ─────────────────────────────────────────────
      await __check('EXACT ASK: the PDF and SVG draw the dashed outline as a vector, no raster frame, no shadow', async () => {
        wall();
        const svg = await capture();
        if (svg.indexOf('data-object-outline') < 0) throw new Error('no outline in the capture');
        if (svg.indexOf('stroke-dasharray') < 0) throw new Error('the outline is not dashed');
        if (svg.indexOf('RASTER') >= 0) throw new Error('a raster frame was drawn for the object');
        if (svg.indexOf('SHADOW') >= 0) throw new Error('the object cast a frame shadow');
      });
      await __check('with a photo the capture carries the photo and no outline', async () => {
        wall({ artworkUrl: 'data:image/png;base64,OBJPHOTO', artworkW: 400, artworkH: 1000 });
        const keep = _imgDownsized; _imgDownsized = async (u) => u;
        let svg = ''; try { svg = await capture(); } finally { _imgDownsized = keep; }
        if (svg.indexOf('OBJPHOTO') < 0) throw new Error('photo missing');
        if (svg.indexOf('data-object-outline') >= 0) throw new Error('outline over the photo');
      });

      // ── the dashboard form ────────────────────────────────────────────────
      await __check('EXACT ASK: the form shows W x H x Depth and Hardware (custom), and hides frame, mats, glass and bleed', async () => {
        setForm();
        const cells = Array.from(document.querySelectorAll('.frame-dim-cell'));
        if (cells.some(shown)) throw new Error('Fr. W / Fr. H / Rabbet still shown');
        if (!shown(document.querySelector('.obj-dim-cell'))) throw new Error('no Depth box');
        if (document.getElementById('m_objDepth').value !== '8') throw new Error('Depth box reads ' + document.getElementById('m_objDepth').value);
        ['frameStyleSection', 'matWrapper', 'bleedSettings', 'canvasSettings', 'egdSettings'].forEach(id => { if (shown(document.getElementById(id))) throw new Error(id + ' still shown'); });
        ['m_glass', 'm_mount', 'm_backing'].forEach(id => { if (shown(document.getElementById(id).parentElement)) throw new Error(id + ' still shown'); });
        const hw = document.getElementById('m_hardware');
        if (!shown(hw.parentElement) || hw.parentElement.querySelector('label').textContent !== 'Hardware (custom)') throw new Error('Hardware is not offered as custom');
        if (!shown(document.getElementById('objectSettings'))) throw new Error('no Sourcing panel');
      });
      await __check('the dashboard preview is a dashed box asking for a photo, not the wallcovering prompt', async () => {
        setForm(); updateDashVisualsFromDOM();
        const av = document.getElementById('dash-art-visual');
        if (!av) throw new Error('no preview opening');
        if (av.textContent.indexOf('wallcovering') >= 0 || av.textContent.indexOf('graphic here') >= 0) throw new Error('shows the flat-graphic prompt: ' + av.textContent);
        if (!av.querySelector('.art-add-btn')) throw new Error('no add-photo button');
        if ((av.getAttribute('style') || '').indexOf('dashed') < 0) throw new Error('not dashed');
        if (document.querySelector('#dash-frame-visual .frame-rail, #dash-frame-visual .mat-visual')) throw new Error('a rail or mat in the preview');
      });
      await __check('a typed depth is stored, survives a switch to Framed Art and back, and converts with the unit', async () => {
        setForm();
        document.getElementById('m_objDepth').value = '6.5';
        syncDashAndCalculate();
        if (dashProjectData[0].objDepth !== 6.5) throw new Error('stored ' + dashProjectData[0].objDepth);
        document.getElementById('m_product').value = 'Framed Art'; handleDashProductChange();
        if (shown(document.getElementById('objectSettings'))) throw new Error('Sourcing still shown on framed art');
        if (!Array.from(document.querySelectorAll('.frame-dim-cell')).every(shown)) throw new Error('frame boxes not back');
        if (shown(document.getElementById('m_glass').parentElement) === false) throw new Error('Glass not back');
        if (dashProjectData[0].objDepth !== 6.5) throw new Error('the depth was lost on the way: ' + dashProjectData[0].objDepth);
        document.getElementById('m_product').value = 'Sourced Object'; handleDashProductChange();
        if (document.getElementById('m_objDepth').value !== '6.5') throw new Error('box reads ' + document.getElementById('m_objDepth').value);
        setUnit('cm');
        if (Math.abs(dashProjectData[0].objDepth - 16.51) > 1e-6) throw new Error('not converted: ' + dashProjectData[0].objDepth);
        setUnit('in');
      });
      await __check('switching to an object clears the framed-art default hardware, never a typed one', async () => {
        dashProjectData = [Object.assign(_cloneData(dashDefaultData), { id: 'ART.1', product: 'Framed Art' })]; dashSelectedRowIndex = 0;
        loadDashDataIntoControls(dashProjectData[0]);
        document.getElementById('m_product').value = 'Sourced Object'; handleDashProductChange();
        if (dashProjectData[0].hardware !== '') throw new Error('kept ' + dashProjectData[0].hardware);
        setForm();
        if (dashProjectData[0].hardware !== 'Pin mount, 2 pins') throw new Error('a typed hardware note was cleared');
      });

      // ── sourcing (the FSO tracker) ────────────────────────────────────────
      await __check('EXACT ASK: link, cost, lead time, stock and alternatives are kept on the piece', async () => {
        setForm();
        _objSourcingSet('vendorName', 'Global Views');
        _objSourcingSet('sourceUrl', ' https://shop.example/vase ');
        _objSourcingSet('unitCost', '245.5');
        _objSourcingSet('leadTime', '2-3 weeks');
        _objSourcingSet('stockQty', '40');
        _objSourcingSet('altLinks', 'https://shop.example/vase-b' + String.fromCharCode(10) + 'https://shop.example/vase-c');
        const r = dashProjectData[0];
        if (r.vendorName !== 'Global Views' || r.sourceUrl !== 'https://shop.example/vase' || r.leadTime !== '2-3 weeks' || r.stockQty !== 40) throw new Error(JSON.stringify(r));
        if (r.unitCost !== 245.5 || r.costSrc !== 'quote') throw new Error('the cost is not a quote Production can price: ' + r.unitCost + ' ' + r.costSrc);
        _objSourcingSet('leadTime', '');
        if ('leadTime' in r) throw new Error('a cleared field left an empty key');
        document.getElementById('m_objLink').value = 'https://shop.example/vase';
        _objOpenLink();
        if (window.__opened !== 'https://shop.example/vase') throw new Error('Open went to ' + window.__opened);
      });
      await __check('under 10 in stock, or fewer than the walls need, is flagged', async () => {
        if (_objStockText({ stockQty: 40, qty: 2 })) throw new Error('40 flagged');
        if (_objStockText({ qty: 2 })) throw new Error('flagged with no stock given');
        if (_objStockText({ stockQty: 6, qty: 2 }).indexOf('Low stock') !== 0) throw new Error('6 in stock not flagged');
        if (_objStockText({ stockQty: 3, qty: 5 }).indexOf('Not enough stock') !== 0) throw new Error('3 for 5 not flagged');
        setForm(); _objSourcingSet('stockQty', '4');
        const note = document.getElementById('objStockNote');
        if (!shown(note) || note.textContent.indexOf('Low stock') !== 0) throw new Error('the form shows no flag');
      });
      await __check('Preflight lists an object with no link or no cost', async () => {
        const keep = _deckSpecRows;
        _deckSpecRows = () => [Object.assign(OBJ(), { id: 'OBJ.2' }), Object.assign(OBJ(), { id: 'OBJ.3', sourceUrl: 'https://x.example/a', unitCost: 10, costSrc: 'quote' }), { id: 'ART.1', product: 'Framed Art' }];
        let out; try { out = _objectPreflight(); } finally { _deckSpecRows = keep; }
        const txt = out.map(o => o.text).join(' / ');
        if (txt.indexOf('OBJ.2') < 0 || txt.indexOf('purchasing link') < 0 || txt.indexOf('cost') < 0) throw new Error(txt);
        if (txt.indexOf('OBJ.3') >= 0 || txt.indexOf('ART.1') >= 0) throw new Error('flagged a complete object or framed art: ' + txt);
      });
      await __check('the Purchases sheet uses the tracker columns with the design half filled', async () => {
        const r = Object.assign(OBJ(), { vendorName: 'Global Views', sourceUrl: 'https://shop.example/vase', unitCost: 245.5, costSrc: 'quote', leadTime: '2-3 weeks', stockQty: 6, altLinks: 'a' + String.fromCharCode(10) + 'b' });
        const csv = _purchaseTrackerCSV([r]);
        const head = csv.split(String.fromCharCode(10))[0];
        ['Source', 'Units', 'Item Code', 'Purchasing Link', 'Lead Time', 'Custom Notes', 'Cost Each', 'Ordered', 'Tracking Link', 'Delivered', 'Hardware', 'Alternatives For Reselection'].forEach(h => {
          if (head.indexOf('"' + h) < 0) throw new Error('no column ' + h);
        });
        ['Global Views', 'https://shop.example/vase', '2-3 weeks', '245.5', 'Pin mount, 2 pins', 'a | b', 'Low stock', '"8"'].forEach(v => {
          if (csv.indexOf(v) < 0) throw new Error('missing ' + v + ' in ' + csv);
        });
      });
    })();
  `;
  try {
    window.eval(src + '\n' + testBlock);
    await window.__run;
  } catch (e) { console.error('LOAD/RUN FAILED:', e.message); process.exit(1); }
  const all = window.__testResults || [];
  const failures = [];
  all.forEach(r => { console.log((r.ok ? 'OK:  ' : 'FAIL:') + ' ' + r.label + (r.ok ? '' : ' -> ' + r.err)); if (!r.ok) failures.push(r.label); });
  console.log('\n--- Summary ---');
  if (failures.length) { console.log(failures.length + ' FAILURES'); process.exit(1); }
  console.log('ALL PASSED (' + all.length + ')');
  process.exit(0);
})();
