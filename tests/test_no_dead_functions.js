// EVERY TOP-LEVEL FUNCTION IN app.js IS CALLED FROM SOMEWHERE (20.83).
//
// A review counted 74 functions nothing called; the real number was 84, and removing them
// exposed 10 more that only the dead ones had called (82 in all, about 800 lines). They were
// the tails of features that moved: the + LAYOUT / + SET buttons (options moved to Deck
// Studio in 17.89), the toolbar fit and align menus, the units wrappers left by 18.47, the
// on-wall dimension drag handles. Each one still read as live code to anyone searching the
// file, and some were still pinned by tests that guarded nothing.
//
// A function counts as used when its name appears anywhere outside its own declaration in
// app.js, templates.js, index.html or sw.js, comments aside. Strings count, so an inline
// onclick in an HTML template is a call. A name only a TEST calls is not used.
const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..');
const read = (f) => fs.readFileSync(path.join(root, f), 'utf8');

// Kept on purpose, each with its reason. A new entry needs one too.
const KEEP = {
  setHelpVideoUrl: 'console API: turns on the Help video tab (CLAUDE.md)',
  _shadeToHex: 'inverse of _hexToShade; the round trip is tested',
  _aGetMeta: 'autosave IndexedDB reader, the autosave test drives it',
  _pdfIsCancelled: 'the PDF cancel test reads the flag through it',
  _specSlotsFor: 'spec-slot map lookup the spec-slot test drives',
  _pageTypeHasOv: 'title-band per-page override lookup the title-band test drives',
  _isCatalogueDeck: 'catalogue deck predicate the catalogue test cross-checks',
  recodeToPattern: 'recode-everything-to-the-pattern, tested; its button left in 20.50',
  _dsNormalizeElZ: 'layer order normaliser the v20 test drives',
  _fitIn: 'aspect-fit helper the safe-frame test drives',
  updateEstimates: 'pricing refresh the pricing test drives',
  applyUnderlayCrop: 'on-wall crop mode, unreachable since 20.67; remove with that whole subsystem',
};

// Blank out comments, keep everything else (strings included).
function stripComments(src) {
  let o = '', i = 0, lastSig = '';
  const n = src.length, stack = [];
  const kw = /(?:^|[^\w$])(return|typeof|case|in|of|delete|void|throw|new|else|do|yield|await)$/;
  const tpl = () => {
    while (i < n) {
      const d = src[i];
      if (d === '\\') { o += src.slice(i, i + 2); i += 2; continue; }
      if (d === '$' && src[i + 1] === '{') { o += '${'; i += 2; stack.push(1); lastSig = '{'; return; }
      o += d; i++;
      if (d === '`') { lastSig = 'a'; return; }
    }
  };
  while (i < n) {
    const c = src[i];
    if (c === '/' && src[i + 1] === '/') { while (i < n && src[i] !== '\n') i++; continue; }
    if (c === '/' && src[i + 1] === '*') { const e = src.indexOf('*/', i + 2); o += ' '; i = e + 2; continue; }
    if (c === '"' || c === "'") {
      let j = i + 1;
      while (j < n && src[j] !== c) { if (src[j] === '\\') j++; j++; }
      o += src.slice(i, j + 1); i = j + 1; lastSig = 'a'; continue;
    }
    if (c === '`') { o += c; i++; tpl(); continue; }
    if (c === '/') {
      const before = o.slice(-12).replace(/\s+$/, '');
      if (lastSig === '' || '(,=:[!&|?{};+-*%<>~^'.indexOf(lastSig) >= 0 || kw.test(before)) {
        let j = i + 1, cls = false;
        while (j < n) { const d = src[j]; if (d === '\\') { j += 2; continue; } if (d === '[') cls = true; else if (d === ']') cls = false; else if ((d === '/' && !cls) || d === '\n') break; j++; }
        j++; while (/[a-z]/.test(src[j] || '')) j++;
        o += src.slice(i, j); i = j; lastSig = 'a'; continue;
      }
    }
    if (stack.length) {
      if (c === '{') stack[stack.length - 1]++;
      else if (c === '}' && --stack[stack.length - 1] === 0) { stack.pop(); o += c; i++; tpl(); continue; }
    }
    if (!/\s/.test(c)) lastSig = /[\w$]/.test(c) ? 'a' : c;
    o += c; i++;
  }
  return o;
}

const results = [];
const check = (label, fn) => { try { fn(); results.push({ label, ok: true }); } catch (e) { results.push({ label, ok: false, err: e.message }); } };

const APP = stripComments(read('app.js'));
const all = APP + '\n' + stripComments(read('templates.js')) + '\n' + stripComments(read('sw.js')) + '\n'
  + read('index.html').replace(/<!--[\s\S]*?-->/g, ' ');
const counts = new Map();
(all.match(/[A-Za-z_$][\w$]*/g) || []).forEach((w) => counts.set(w, (counts.get(w) || 0) + 1));
const declared = [];
{ const re = /^(?:async\s+)?function\s*\*?\s*([A-Za-z_$][\w$]*)\s*\(/gm; let m; while ((m = re.exec(APP))) declared.push(m[1]); }
const dead = declared.filter((nm) => (counts.get(nm) || 0) <= 1);

check('the scan finds the top-level functions (it is reading app.js)', () => {
  if (declared.length < 1000) throw new Error('only ' + declared.length + ' declarations found');
  if ((counts.get('setUnit') || 0) < 3) throw new Error('setUnit, called in several places, counted ' + counts.get('setUnit'));
});
check('EXACT BUG: no top-level function in app.js is left with no caller', () => {
  const bad = dead.filter((nm) => !KEEP[nm]);
  if (bad.length) throw new Error(bad.length + ' uncalled: ' + bad.slice(0, 12).join(', ') + ' (delete it, or add it to KEEP with a reason)');
});
check('every KEEP entry is still uncalled, so the list cannot rot', () => {
  const stale = Object.keys(KEEP).filter((nm) => declared.indexOf(nm) < 0 || dead.indexOf(nm) < 0);
  if (stale.length) throw new Error('now called or gone, take it off KEEP: ' + stale.join(', '));
});
check('the removed leftovers stay removed', () => {
  ['catArrangementAction', '_catLayoutChooser', 'openUnitMenu', 'setElevUnit', 'setDashUnit', '_dsMoveLayoutPage',
   '_dsDeleteInserted', '_elevReturnToDeck', 'createDesignFolder', '_drawPlaceholderPage'].forEach((nm) => {
    if (declared.indexOf(nm) >= 0) throw new Error(nm + ' is back');
  });
});

results.forEach(r => console.log((r.ok ? 'OK:   ' : 'FAIL: ') + r.label + (r.ok ? '' : ' -> ' + r.err)));
const bad = results.filter(r => !r.ok);
console.log('');
console.log('--- Summary ---');
if (bad.length) { console.log('FAILED (' + bad.length + ' of ' + results.length + ')'); process.exitCode = 1; }
else console.log('ALL PASSED (' + results.length + ')');
