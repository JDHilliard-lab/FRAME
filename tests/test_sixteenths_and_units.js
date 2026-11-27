// SIXTEENTHS SURVIVE, AND A RESTORED PROJECT KEEPS ITS UNIT (20.84).
//
// Found checking a review's "Rabbet shows 0.813 while the other fields show 1 or 2 decimals":
// it was not a display quirk. dashFmt rounded to THREE places, and it both fills the form and
// writes data (picking a frame, the default frame, unit conversion). So the default 13/16"
// rabbet was stored as 0.813 from the moment it was applied, a 24-1/16" piece became 24.063
// on the next edit of anything (the form reads every field back), and Chrome marked the
// Rabbet box invalid because 0.813 is not a 1/16" step.
//
// Found on the way: the autosave payload and the undo / version snapshot never recorded the
// UNIT. Continue on a cm project put 60.96 (24") back under a fresh boot's inches, so the
// piece came back 60.96" wide.
const fs = require('fs');
const path = require('path');
const { JSDOM } = require('jsdom');

const root = path.join(__dirname, '..');
const APP = fs.readFileSync(path.join(root, 'app.js'), 'utf8');
const HTML = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const NL = String.fromCharCode(10);

const results = [];
const check = (label, fn) => {
  try { fn(); results.push({ label, ok: true }); }
  catch (e) { results.push({ label, ok: false, err: e.message + (process.env.DBG ? e.stack : '') }); }
};

const dom = new JSDOM(HTML, { url: 'https://example.com/', pretendToBeVisual: true, runScripts: 'outside-only' });
const { window } = dom;
window.HTMLCanvasElement.prototype.getContext = () => ({});
window.fetch = () => Promise.reject(new Error('no network in test'));
global.window = window; global.document = window.document; global.navigator = window.navigator;
const quiet = console.error; console.error = () => {}; const qlog = console.log; console.log = () => {};
window.eval(APP + NL + [
  '_autosaveToldUser = true;',
  'window.__fx = {',
  '  get unit() { return dashUnit; }, get elevUnitNow() { return elevUnit; },',
  '  setUnitRaw: function (u) { dashUnit = u; elevUnit = u; },',
  '  rows: function () { return dashProjectData; }, elevs: function () { return elevations; },',
  '  select: function (i) { dashSelectedRowIndex = i; },',
  '  payload: function () { return _cloneData(_autosavePayload()); },',
  '  snapshot: function () { return snapshotProjectState(); },',
  '  dashFmt: function (v) { return dashFmt(v); },',
  '};',
].join(NL));
console.error = quiet; console.log = qlog;
const fx = window.__fx;
const doc = window.document;
window.scheduleAutosave = () => {};

// ── the rounding ─────────────────────────────────────────────────────────
check('EXACT BUG: dashFmt keeps every sixteenth (13/16 is 0.8125, not 0.813)', () => {
  [[0.8125, 0.8125], [24.0625, 24.0625], [1.0625, 1.0625], [0.0625, 0.0625], [1.875, 1.875], [24, 24]].forEach(([v, want]) => {
    if (fx.dashFmt(v) !== want) throw new Error('dashFmt(' + v + ') = ' + fx.dashFmt(v));
  });
  if (fx.dashFmt(0.75 * 2.54) !== 1.905) throw new Error('float residue not trimmed: ' + fx.dashFmt(0.75 * 2.54));
});
check('EXACT BUG: the default frame stores its 13/16" rabbet exactly', () => {
  fx.setUnitRaw('in');
  const row = window._defaultFrameApply({ id: 'X', product: 'Framed Art' });
  if (row.rabbetDepth !== 0.8125) throw new Error('rabbet stored as ' + row.rabbetDepth);
});
check('EXACT BUG: editing another field does not round the frame sizes the form shows', () => {
  fx.setUnitRaw('in');
  const r = fx.rows()[0];
  Object.assign(r, { rabbetDepth: 0.8125, fW: 1.0625, extW: 24.0625 });
  fx.select(0);
  window.loadDashDataIntoControls(r);
  const shown = doc.getElementById('rabbetDepth').value;
  if (shown !== '0.8125') throw new Error('Rabbet box shows ' + shown);
  window.syncDashAndCalculate();
  if (r.rabbetDepth !== 0.8125 || r.fW !== 1.0625 || r.extW !== 24.0625) throw new Error('stored after an edit: ' + JSON.stringify({ rabbet: r.rabbetDepth, fW: r.fW, extW: r.extW }));
});

// ── putting back what the old rounding did ─────────────────────────────
check('only the exact three-place rounding of an odd sixteenth goes back', () => {
  const u = window._unroundSixteenth;
  [[0.813, 0.8125], [24.063, 24.0625], [1.063, 1.0625], [0.063, 0.0625], ['0.813', 0.8125], [2.938, 2.9375]].forEach(([v, want]) => {
    if (u(v) !== want) throw new Error(JSON.stringify(v) + ' -> ' + u(v) + ', wanted ' + want);
  });
  // Typed values and exact ones are left alone.
  [24.06, 0.81, 0.875, 0.8125, 24, 1.5, 0.812, 3.333, '', 'abc', '0.8130'].forEach((v) => {
    if (u(v) !== v) throw new Error(JSON.stringify(v) + ' was changed to ' + u(v));
  });
});
check('opening an inch project puts drifted rows and wall frames back; a cm project is untouched', () => {
  const mk = (unit) => ({ type: 'master-studio-v6', dashUnit: unit, elevUnit: unit,
    dashProjectData: [{ id: 'ART.1', rabbetDepth: 0.813, fW: 1.063, extW: 24.063, extH: 36 }],
    elevations: [{ wallH: unit === 'in' ? 120 : 305, frames: [{ id: 'ART.1', w: 24.063, rabbetDepth: 0.813, x: 10.063 }] }] });
  const a = window._migrateLoadedProject(mk('in'));
  const r = a.dashProjectData[0], f = a.elevations[0].frames[0];
  if (r.rabbetDepth !== 0.8125 || r.fW !== 1.0625 || r.extW !== 24.0625 || r.extH !== 36) throw new Error('row ' + JSON.stringify(r));
  if (f.w !== 24.0625 || f.rabbetDepth !== 0.8125) throw new Error('frame ' + JSON.stringify(f));
  if (f.x !== 10.063) throw new Error('a POSITION was snapped: ' + f.x);
  const c = window._migrateLoadedProject(mk('cm'));
  if (c.dashProjectData[0].rabbetDepth !== 0.813) throw new Error('a cm project was snapped');
});

// ── the unit travels with the numbers ───────────────────────────────────
check('EXACT BUG: Continue on a cm project comes back in cm, not inches', () => {
  fx.setUnitRaw('in');
  const r = fx.rows()[0]; r.extW = 24; r.extH = 36;
  window.setUnit('cm');
  const p = fx.payload();
  if (p.data.dashUnit !== 'cm') throw new Error('the autosave payload does not record the unit');
  fx.setUnitRaw('in');                       // what a fresh boot starts with
  window._autosaveRestore(p);
  if (fx.unit !== 'cm' || fx.elevUnitNow !== 'cm') throw new Error('restored under ' + fx.unit + '/' + fx.elevUnitNow);
  if (Math.abs(fx.rows()[0].extW - 60.96) > 1e-9) throw new Error('width ' + fx.rows()[0].extW);
  window.setUnit('in');
});
check('undo across a unit switch puts the unit back with the numbers', () => {
  fx.setUnitRaw('in');
  const snapIn = fx.snapshot();
  if (snapIn.dashUnit !== 'in') throw new Error('the history snapshot does not record the unit');
  window.setUnit('cm');
  window.restoreProjectState(snapIn);
  if (fx.unit !== 'in') throw new Error('restored inch numbers under ' + fx.unit);
});
check('an older snapshot with no unit leaves the live unit alone', () => {
  fx.setUnitRaw('cm');
  const s = fx.snapshot(); delete s.dashUnit;
  window.restoreProjectState(s);
  if (fx.unit !== 'cm') throw new Error('changed to ' + fx.unit);
  fx.setUnitRaw('in');
});

// ── the Project Code box ────────────────────────────────────────────────
check('EXACT BUG: the Project Code example fits its box (it was cut off)', () => {
  const inp = doc.getElementById('specPdfCode');
  // Measured in Chrome: the box holds 162px of 12.8px monospace, about 21 characters.
  if (inp.placeholder.length > 20) throw new Error('"' + inp.placeholder + '" is ' + inp.placeholder.length + ' characters');
  if (!inp.title) throw new Error('the full example moved to the tooltip and is missing');
});

results.forEach(r => console.log((r.ok ? 'OK:   ' : 'FAIL: ') + r.label + (r.ok ? '' : ' -> ' + r.err)));
const bad = results.filter(r => !r.ok);
console.log('');
console.log('--- Summary ---');
if (bad.length) { console.log('FAILED (' + bad.length + ' of ' + results.length + ')'); process.exitCode = 1; }
else console.log('ALL PASSED (' + results.length + ')');
