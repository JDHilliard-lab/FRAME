// NOTHING THE APP NEEDS COMES FROM THE INTERNET (20.82).
//
// The offline work in 19.26 kept the app itself on the device, but three files still came from
// outside: html2canvas from cdnjs (Elevation PNG export), the Oswald font from Google Fonts,
// and a flaticon image as the scale figure's fallback. The service worker only kept them in its
// runtime cache, which drops its oldest entries past 1500, so a long frame-library browse could
// evict html2canvas and PNG export then failed on a site with no signal. They are local files
// now, and sw.js fetches them at install.
//
// Same release: UI text loses its em dashes (house style), app.js loses a literal NUL byte
// (grep and other tools treated the whole file as binary), and the duplicate JSZip copy in
// scripts/ is gone.
const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..');
const read = (f) => fs.readFileSync(path.join(root, f), 'utf8');
const HTML = read('index.html');
const SW = read('sw.js');
const APP_BYTES = fs.readFileSync(path.join(root, 'app.js'));
const APP = APP_BYTES.toString('utf8');
const PKG = JSON.parse(read('package.json'));

const results = [];
const check = (label, fn) => { try { fn(); results.push({ label, ok: true }); } catch (e) { results.push({ label, ok: false, err: e.message }); } };
const stripHtmlComments = (s) => s.replace(/<!--[\s\S]*?-->/g, '');
const LIVE_HTML = stripHtmlComments(HTML);

// ── index.html loads nothing from outside ───────────────────────────────
check('EXACT BUG: no <script src> points at the internet (html2canvas came from cdnjs)', () => {
  const ext = (LIVE_HTML.match(/<script[^>]+src="https?:[^"]*"/g) || []);
  if (ext.length) throw new Error('external script(s): ' + ext.join(', '));
});
check('EXACT BUG: no stylesheet or font comes from Google Fonts', () => {
  if (LIVE_HTML.indexOf('fonts.googleapis.com') >= 0 || LIVE_HTML.indexOf('fonts.gstatic.com') >= 0) throw new Error('index.html still loads Google Fonts');
  const ext = (LIVE_HTML.match(/<link[^>]+href="https?:[^"]*"/g) || []);
  if (ext.length) throw new Error('external link(s): ' + ext.join(', '));
});
check('EXACT BUG: the scale figure has no internet fallback (flaticon)', () => {
  if (LIVE_HTML.indexOf('flaticon') >= 0) throw new Error('index.html still points at flaticon');
  const ext = (LIVE_HTML.match(/(src|onerror)="[^"]*https?:\/\/[^"]*"/g) || []);
  if (ext.length) throw new Error('external src/onerror: ' + ext.join(', '));
});
check('no url(http...) inside index.html styles', () => {
  if (/url\(\s*['"]?https?:/.test(LIVE_HTML)) throw new Error('a CSS url() points at the internet');
});

// ── the local copies exist and are real ─────────────────────────────────
check('lib-html2canvas.min.js is loaded and is the real library', () => {
  if (LIVE_HTML.indexOf('<script src="lib-html2canvas.min.js"></script>') < 0) throw new Error('index.html does not load lib-html2canvas.min.js');
  const f = path.join(root, 'lib-html2canvas.min.js');
  if (!fs.existsSync(f)) throw new Error('lib-html2canvas.min.js is missing');
  const s = fs.readFileSync(f, 'utf8');
  if (s.length < 100000 || s.indexOf('html2canvas') < 0) throw new Error('lib-html2canvas.min.js does not look like html2canvas (' + s.length + ' bytes)');
});
check('Oswald is declared with @font-face from local woff2 files that exist', () => {
  const faces = LIVE_HTML.match(/@font-face\s*\{[^}]*font-family:\s*'Oswald'[^}]*\}/g) || [];
  if (faces.length < 1) throw new Error('no @font-face for Oswald');
  faces.forEach((f) => {
    const m = f.match(/url\('([^']+)'\)/);
    if (!m) throw new Error('Oswald face has no url: ' + f.slice(0, 80));
    const buf = fs.readFileSync(path.join(root, m[1]));
    if (buf.slice(0, 4).toString('latin1') !== 'wOF2') throw new Error(m[1] + ' is not a woff2 file');
  });
  if (!faces.some(f => f.indexOf('U+0000-00FF') >= 0)) throw new Error('no Oswald face covers basic latin');
});

// ── the service worker keeps them from install, not just in the runtime cache ─
const STATIC = (() => {
  const m = SW.match(/const STATIC = \[([\s\S]*?)\];/);
  if (!m) return null;
  return (m[1].match(/'([^']+)'/g) || []).map(s => s.slice(1, -1));
})();
check('EXACT BUG: sw.js fetches html2canvas, Oswald and the scale figure at install', () => {
  if (!STATIC) throw new Error('could not read STATIC from sw.js');
  ['lib-html2canvas.min.js', 'fonts/Oswald-latin.woff2', 'fonts/Oswald-latin-ext.woff2', 'Character_Lady_walk.svg'].forEach((f) => {
    if (STATIC.indexOf(f) < 0) throw new Error(f + ' is not in STATIC, so the runtime cache can drop it and it fails offline');
  });
});
check('every STATIC file in sw.js exists, so install never caches a 404', () => {
  STATIC.filter(f => f !== './').forEach((f) => {
    if (!fs.existsSync(path.join(root, f))) throw new Error('sw.js lists ' + f + ' but there is no such file');
  });
});

// ── housekeeping ───────────────────────────────────────────────────────
check('EXACT BUG: app.js has no literal NUL byte (tools treated it as a binary file)', () => {
  const at = APP_BYTES.indexOf(0);
  if (at >= 0) throw new Error('NUL byte at line ' + APP_BYTES.slice(0, at).toString('utf8').split('\n').length + '; write it as an escape instead');
});
check('the image-code join key still uses the NUL character, written as an escape', () => {
  if (APP.indexOf("vals.join('\\u0000')") < 0) throw new Error("vals.join('\\u0000') not found; the dedupe key changed");
});
check('each vendored library exists exactly once in the repo', () => {
  const seen = {};
  const walk = (d) => fs.readdirSync(d, { withFileTypes: true }).forEach((e) => {
    if (e.name === 'node_modules' || e.name.startsWith('.')) return;
    const p = path.join(d, e.name);
    if (e.isDirectory()) walk(p);
    else if (/^lib-.*\.min\.js$/.test(e.name)) (seen[e.name] = seen[e.name] || []).push(path.relative(root, p));
  });
  walk(root);
  const dup = Object.keys(seen).filter(k => seen[k].length > 1);
  if (dup.length) throw new Error('duplicate copies: ' + dup.map(k => seen[k].join(' + ')).join('; '));
});
check('npm test runs the suite', () => {
  if (!PKG.scripts || PKG.scripts.test !== 'node tests/run-all.js') throw new Error('package.json scripts.test is ' + JSON.stringify(PKG.scripts && PKG.scripts.test));
});

// ── no em dashes in UI text ─────────────────────────────────────────────
check('index.html shows no em dash (comments aside)', () => {
  const nc = LIVE_HTML.replace(/\/\*[\s\S]*?\*\//g, '');
  // 20.85: the entity form too. Nine &mdash; in option text and tooltips passed the 20.82 sweep.
  const lines = nc.split('\n').filter(l => l.indexOf('—') >= 0 || l.indexOf('&mdash;') >= 0 || l.indexOf('&#8212;') >= 0);
  if (lines.length) throw new Error(lines.length + ' line(s), e.g. ' + lines[0].trim().slice(0, 100));
});

// String and template literal chunks of app.js containing an em dash, written out or as an
// escape. Comments are skipped; a regex literal is skipped by the usual previous-token rule.
function emDashStrings(src) {
  const out = []; const n = src.length; let i = 0, lastSig = '';
  const kw = /(?:^|[^\w$])(return|typeof|case|in|of|delete|void|throw|new|else|do|yield|await)$/;
  const stack = [];
  const has = s => s.indexOf('—') >= 0 || s.indexOf('\\u2014') >= 0;
  const str = (q, from) => {
    let j = from;
    while (j < n) {
      const c = src[j];
      if (c === '\\') { j += 2; continue; }
      if (q === '`' && c === '$' && src[j + 1] === '{') { if (has(src.slice(from, j))) out.push(src.slice(from, j)); i = j + 2; stack.push(1); lastSig = '{'; return; }
      if (c === q) { if (has(src.slice(from, j))) out.push(src.slice(from, j)); i = j + 1; lastSig = 'a'; return; }
      j++;
    }
    i = n;
  };
  while (i < n) {
    const c = src[i];
    if (c === ' ' || c === '\t' || c === '\n' || c === '\r') { i++; continue; }
    if (c === '/' && src[i + 1] === '/') { while (i < n && src[i] !== '\n') i++; continue; }
    if (c === '/' && src[i + 1] === '*') { i = src.indexOf('*/', i + 2) + 2; continue; }
    if (c === '"' || c === "'" || c === '`') { str(c, i + 1); continue; }
    if (c === '/') {
      const before = src.slice(Math.max(0, i - 12), i).replace(/\s+$/, '');
      if (lastSig === '' || '(,=:[!&|?{};+-*%<>~^'.indexOf(lastSig) >= 0 || kw.test(before)) {
        let j = i + 1, cls = false;
        while (j < n) { const d = src[j]; if (d === '\\') { j += 2; continue; } if (d === '[') cls = true; else if (d === ']') cls = false; else if ((d === '/' && !cls) || d === '\n') break; j++; }
        i = j + 1; while (/[a-z]/.test(src[i] || '')) i++; lastSig = 'a'; continue;
      }
    }
    if (stack.length) {
      if (c === '{') stack[stack.length - 1]++;
      else if (c === '}' && --stack[stack.length - 1] === 0) { stack.pop(); str('`', i + 1); continue; }
    }
    lastSig = /[\w$]/.test(c) ? 'a' : c; i++;
  }
  return out;
}
// What may keep one: a lone dash meaning "no value", the studio's style names in the seed
// data, and headings printed on client PDF pages (left as the studio's layouts set them).
const EM_ALLOWED = [
  (s) => /^(—|\\u2014)$/.test(s),
  (s) => /^(sty_thank_you_|sty_quote_|Thank You |Quote )/.test(s),
  (s) => /^(PLACEHOLDER |PROPOSED FLOOR PLAN |PRINT FILES )/.test(s),
  (s) => /FLOORPLAN — $/.test(s),
  (s) => s === ' \\u2014 ',   // _catLayoutTag's "A — name", printed on catalogue pages
];
check('app.js UI text has no em dashes (printed headings and placeholders aside)', () => {
  const found = emDashStrings(APP);
  if (found.length < 5) throw new Error('scanner found only ' + found.length + ' allowed strings; it is not reading app.js');
  const bad = found.filter(s => !EM_ALLOWED.some(ok => ok(s)));
  const tag = found.filter(s => s === ' \\u2014 ').length;
  if (tag > 1) bad.push('(' + tag + ' bare joiners, only _catLayoutTag may keep one)');
  if (bad.length) throw new Error(bad.length + ' string(s), e.g. ' + bad.slice(0, 3).map(s => JSON.stringify(s.slice(0, 70))).join(', '));
});

results.forEach(r => console.log((r.ok ? 'OK:   ' : 'FAIL: ') + r.label + (r.ok ? '' : ' -> ' + r.err)));
const bad = results.filter(r => !r.ok);
console.log('');
console.log('--- Summary ---');
if (bad.length) { console.log('FAILED (' + bad.length + ' of ' + results.length + ')'); process.exitCode = 1; }
else console.log('ALL PASSED (' + results.length + ')');
