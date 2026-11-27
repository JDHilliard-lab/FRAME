// TYPE A GAP, AND A TAPE MEASURE (20.87).
//
// Asked: "is there anything you would add to the elevation section, think of other
// softwares that use dimension mark ups like SketchUp", then "sure build it" for the two
// picked first:
//   1. Type a number to move things (SketchUp types a distance after a drag; Revit edits a
//      dimension's number to move the element). A gap here could only be dragged until the
//      number read right. Double-click a spacing or edge-gap number, type, Enter.
//   2. A tape measure that does not print. Measure (M) adds a line to the drawing, so every
//      quick check became mark-up to delete.
const { JSDOM } = require('jsdom');
const fs = require('fs');
const path = require('path');

(async () => {
  const root = path.join(__dirname, '..');
  const src = fs.readFileSync(path.join(root, 'app.js'), 'utf8');
  const htmlSrc = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
  const dom = new JSDOM(htmlSrc, { url: 'https://example.com/', pretendToBeVisual: true, runScripts: 'outside-only' });
  const { window } = dom;
  window.HTMLCanvasElement.prototype.getContext = () => ({});
  window.fetch = () => Promise.reject(new Error('no network in test'));
  global.window = window; global.document = window.document; global.navigator = window.navigator;

  const testBlock = `
    window.__testResults = [];
    const __check = (label, fn) => { try { fn(); window.__testResults.push({ label, ok: true }); } catch (e) { window.__testResults.push({ label, ok: false, err: e.message }); } };
    _autosaveToldUser = true; scheduleAutosave = () => {};
    editorialContent = editorialContent || {};
    const $ = (id) => document.getElementById(id);
    const near = (a, b) => Math.abs(a - b) < 1e-6;
    const seed = (frames, opts) => {
      elevUnit = (opts && opts.unit) || 'in';
      elevations = [{ name: 'Wall 1', wallW: 185, wallH: 108, personPos: { x: -60 }, frames: frames },
                    { name: 'Wall 2', wallW: 185, wallH: 108, personPos: { x: -60 }, frames: [] }];
      currentElevIndex = 0; currentView = 'elevation';
      elevFrames = elevations[0].frames;
      elevZoomFactor = 1;
      $('wallW').value = '185'; $('wallH').value = '108';
      const ws = document.querySelector('#view-elevation .workspace');
      Object.defineProperty(ws, 'clientWidth', { get: () => 1100, configurable: true });
      Object.defineProperty(ws, 'clientHeight', { get: () => 800, configurable: true });
      $('dim-layer').style.display = 'block';
      dimVisibility.edgeGap = true;
      drawElevAll();
    };
    const pair = () => [
      { letter: 'A', id: 'A1', x: 20, y: 40, w: 30, h: 24, active: true, dimTo: ['B'] },
      { letter: 'B', id: 'B1', x: 80, y: 40, w: 30, h: 24, active: true },
    ];
    const labelFor = (id) => {
      // The dim element carries no id; find the number whose double-tap targets this id by
      // reading the drawn gap text is fragile, so go through the layer in draw order.
      const all = document.querySelectorAll('#dim-layer .arch-dim .arch-label-new, #floor-ceiling-layer .arch-dim .arch-label-new');
      return all;
    };
    const dbl = (el) => {
      el.dispatchEvent(new window.MouseEvent('mousedown', { bubbles: true }));
      // The first press redraws: take the fresh node in the same place.
      return el;
    };
    const press = (k, el) => (el || document).dispatchEvent(new window.KeyboardEvent('keydown', { key: k, bubbles: true }));

    // ── reading what was typed ───────────────────────────────────────────
    __check('sizes are read the way a designer types them', () => {
      const P = (t, u) => _parseLenInput(t, u || 'in');
      [['4', 4], ['3.5', 3.5], ['3 1/2', 3.5], ['3-1/2', 3.5], ['1/2', 0.5], ["1'6\\"", 18], ["2'", 24], ["1' 6", 18],
       ["1'-6 1/2\\"", 18.5], ['4"', 4], ['4in', 4], ['0', 0]].forEach(([t, want]) => {
        const v = P(t); if (v === null || !near(v, want)) throw new Error(JSON.stringify(t) + ' read as ' + v + ', wanted ' + want);
      });
      if (!near(P('10cm'), 10 / 2.54)) throw new Error('10cm read as ' + P('10cm'));
      if (!near(P('254mm'), 10)) throw new Error('254mm read as ' + P('254mm'));
      if (!near(P('4in', 'cm'), 10.16)) throw new Error('4in on a cm wall read as ' + P('4in', 'cm'));
      if (!near(P('10', 'cm'), 10)) throw new Error('a bare number on a cm wall is cm');
      ['', 'abc', '-3', '3/0', '1/2/3'].forEach(t => { if (P(t) !== null) throw new Error(JSON.stringify(t) + ' was accepted as ' + P(t)); });
    });

    // ── typing a gap ─────────────────────────────────────────────────────
    __check('EXACT ASK: double-click a spacing number, type 6, Enter: the right frame moves, the left stays', () => {
      seed(pair());
      const lbl = document.querySelector('#dim-layer .arch-dim .arch-label-new');
      if (!lbl) throw new Error('fixture: no spacing number drawn');
      lbl.dispatchEvent(new window.MouseEvent('mousedown', { bubbles: true }));
      const lbl2 = document.querySelector('#dim-layer .arch-dim .arch-label-new');
      lbl2.dispatchEvent(new window.MouseEvent('mousedown', { bubbles: true }));
      const inp = $('dimTypeInput');
      if (!inp) throw new Error('no box to type in after a double-click');
      if (inp.value !== '30') throw new Error('box starts at ' + inp.value + ', the gap is 30');
      const before = undoStack.length;
      inp.value = '6';
      press('Enter', inp);
      const [A, B] = elevFrames;
      if (A.x !== 20) throw new Error('the left frame moved to ' + A.x);
      if (!near(B.x, 56)) throw new Error('right frame at ' + B.x + ', a 6" gap puts it at 56');
      if ($('dimTypeInput')) throw new Error('the box stayed open');
      if (undoStack.length <= before) throw new Error('not undoable');
    });
    __check('a vertical gap moves the upper frame', () => {
      seed([{ letter: 'A', id: 'A1', x: 20, y: 20, w: 30, h: 24, active: true, dimTo: ['B'] },
            { letter: 'B', id: 'B1', x: 20, y: 60, w: 30, h: 24, active: true }]);
      if (!_dimTypeCommit('spacing-v-A-B', '3 1/2')) throw new Error('commit refused');
      if (elevFrames[0].y !== 20 || !near(elevFrames[1].y, 47.5)) throw new Error(JSON.stringify(elevFrames.map(f => f.y)));
    });
    __check('edge gaps move the frame to the typed distance from that edge', () => {
      seed([{ letter: 'A', id: 'A1', x: 20, y: 40, w: 30, h: 24, active: true, distToggles: { right: true, floor: true } }]);
      _dimTypeCommit('edge-A-right', '12');
      const A = elevFrames[0];
      if (!near(185 - (A.x + A.w), 12)) throw new Error('right gap is ' + (185 - (A.x + A.w)));
      _dimTypeCommit('edge-A-floor', "4'");
      if (!near(A.y, 48)) throw new Error('floor gap is ' + A.y);
      _dimTypeCommit('edge-A-left', '10');
      if (!near(A.x, 10)) throw new Error('left gap is ' + A.x);
      _dimTypeCommit('edge-A-ceiling', '20');
      if (!near(108 - (A.y + A.h), 20)) throw new Error('ceiling gap is ' + (108 - (A.y + A.h)));
    });
    __check('a double-click on an edge-gap number opens the box too', () => {
      seed([{ letter: 'A', id: 'A1', x: 20, y: 40, w: 30, h: 24, active: true, distToggles: { left: true } }]);
      const q = () => document.querySelector('#floor-ceiling-layer .arch-dim .arch-label-new');
      if (!q()) throw new Error('fixture: no edge gap drawn');
      q().dispatchEvent(new window.MouseEvent('mousedown', { bubbles: true }));
      q().dispatchEvent(new window.MouseEvent('mousedown', { bubbles: true }));
      if (!$('dimTypeInput')) throw new Error('no box');
      press('Escape', $('dimTypeInput'));
    });
    __check('Esc leaves the gap as it was', () => {
      seed(pair());
      const q = () => document.querySelector('#dim-layer .arch-dim .arch-label-new');
      q().dispatchEvent(new window.MouseEvent('mousedown', { bubbles: true }));
      q().dispatchEvent(new window.MouseEvent('mousedown', { bubbles: true }));
      const inp = $('dimTypeInput'); inp.value = '2';
      press('Escape', inp);
      if (elevFrames[1].x !== 80) throw new Error('moved on Esc');
      if ($('dimTypeInput')) throw new Error('box still open');
    });
    __check('something that is not a size moves nothing', () => {
      seed(pair());
      if (_dimTypeCommit('spacing-h-A-B', 'wide')) throw new Error('accepted');
      if (elevFrames[1].x !== 80) throw new Error('moved');
    });
    __check('a grouped frame takes its group; the frame measured from never moves', () => {
      seed([{ letter: 'A', id: 'A1', x: 20, y: 40, w: 30, h: 24, active: true, dimTo: ['B'] },
            { letter: 'B', id: 'B1', x: 80, y: 40, w: 30, h: 24, active: true, isGrouped: true },
            { letter: 'C', id: 'C1', x: 120, y: 40, w: 30, h: 24, active: true, isGrouped: true }]);
      _dimTypeCommit('spacing-h-A-B', '10');
      if (!near(elevFrames[1].x, 60) || !near(elevFrames[2].x, 100)) throw new Error('group did not move together: ' + elevFrames.map(f => f.x));
      seed([{ letter: 'A', id: 'A1', x: 20, y: 40, w: 30, h: 24, active: true, dimTo: ['B'], isGrouped: true },
            { letter: 'B', id: 'B1', x: 80, y: 40, w: 30, h: 24, active: true, isGrouped: true }]);
      _dimTypeCommit('spacing-h-A-B', '10');
      if (elevFrames[0].x !== 20 || !near(elevFrames[1].x, 60)) throw new Error('both grouped: ' + elevFrames.map(f => f.x));
    });
    __check('selected frames move together', () => {
      seed([{ letter: 'A', id: 'A1', x: 20, y: 40, w: 30, h: 24, active: true, dimTo: ['B'] },
            { letter: 'B', id: 'B1', x: 80, y: 40, w: 30, h: 24, active: true, selected: true },
            { letter: 'C', id: 'C1', x: 120, y: 40, w: 30, h: 24, active: true, selected: true }]);
      _dimTypeCommit('spacing-h-A-B', '0');
      if (!near(elevFrames[1].x, 50) || !near(elevFrames[2].x, 90)) throw new Error(elevFrames.map(f => f.x).join(','));
    });
    __check('on a centimetre wall a bare number is centimetres', () => {
      seed([{ letter: 'A', id: 'A1', x: 50, y: 100, w: 60, h: 60, active: true, dimTo: ['B'] },
            { letter: 'B', id: 'B1', x: 200, y: 100, w: 60, h: 60, active: true }], { unit: 'cm' });
      _dimTypeCommit('spacing-h-A-B', '10');
      if (!near(elevFrames[1].x, 120)) throw new Error('B at ' + elevFrames[1].x);
    });
    __check('glazing and wall dimensions do not offer the box (nothing there to move)', () => {
      if (_dimTypeTarget('glazing-0-1') !== null) throw new Error('glazing dim offered a box');
      if (_dimTypeTarget('spacing-h-A-Z') !== null) throw new Error('a gap to a frame that is not there offered a box');
    });

    // ── the tape ─────────────────────────────────────────────────────────
    // jsdom has no layout: the wall rect is 0 x 0 at the origin, so a point at wall
    // inches (x, y) is clientX = x * scale, clientY = -y * scale. Alt = no snapping.
    const at = (x, y) => ({ clientX: x * elevScale, clientY: -y * elevScale, altKey: true, button: 0, bubbles: true, cancelable: true });
    const click = (x, y) => $('wall').dispatchEvent(new window.MouseEvent('mousedown', at(x, y)));
    const move = (x, y) => $('wall').dispatchEvent(new window.MouseEvent('mousemove', at(x, y)));
    const readout = () => { const r = document.querySelector('#tape-layer .tape-readout'); return r ? r.textContent : ''; };

    __check('the readout says across, up and direct', () => {
      elevUnit = 'in';
      const t = _tapeReadout({ x: 0, y: 0 }, { x: 30, y: 40 });
      ['30" across', '40" up', '50" direct'].forEach(w => { if (t.indexOf(w) < 0) throw new Error(t); });
      if (_tapeReadout({ x: 0, y: 10 }, { x: 24, y: 10 }) !== '24" across') throw new Error('level: ' + _tapeReadout({ x: 0, y: 10 }, { x: 24, y: 10 }));
      if (_tapeReadout({ x: 5, y: 50 }, { x: 5, y: 8 }).indexOf('42" down') !== 0) throw new Error('plumb');
    });
    __check('EXACT ASK: T picks up the tape; two clicks read a distance and nothing is drawn into the wall', () => {
      seed(pair());
      const snap = JSON.stringify(elevations);
      const before = undoStack.length;
      press('t');
      if (!tapeToolActive) throw new Error('T did not arm the tape');
      if (!$('tapeToolBtn').classList.contains('active')) throw new Error('dock button not lit');
      click(10, 20); move(40, 20);
      if (readout() !== '30" across') throw new Error('live reading: ' + readout());
      click(40, 60);
      if (readout().indexOf('40" up') < 0 || readout().indexOf('30" across') < 0) throw new Error('held reading: ' + readout());
      move(100, 100);
      if (readout().indexOf('40" up') < 0) throw new Error('a held reading followed the mouse: ' + readout());
      if (JSON.stringify(elevations) !== snap) throw new Error('the tape changed the project');
      if (undoStack.length !== before) throw new Error('the tape pushed history');
      if (elevFrames[0].x !== 20) throw new Error('a click through the tape moved a frame');
    });
    __check('the tape never exports', () => {
      const l = $('tape-layer');
      if (!l) throw new Error('no tape layer');
      if (l.getAttribute('data-export-skip') !== '1' || l.getAttribute('data-html2canvas-ignore') !== 'true') throw new Error('not skipped by the exports');
    });
    __check('a third click starts a new reading', () => {
      click(0, 0); move(12, 0);
      if (readout() !== '12" across') throw new Error(readout());
    });
    __check('Esc clears the reading, then puts the tape down', () => {
      press('Escape');
      if (!tapeToolActive) throw new Error('first Esc put the tape down');
      if (readout()) throw new Error('reading still showing');
      press('Escape');
      if (tapeToolActive) throw new Error('second Esc left it armed');
      if ($('tape-layer')) throw new Error('layer left behind');
    });
    __check('one wall tool at a time: the tape and Measure put each other down', () => {
      toggleTapeTool(true);
      toggleLineTool(true);
      if (tapeToolActive) throw new Error('Measure did not put the tape down');
      toggleTapeTool(true);
      if (lineToolActive) throw new Error('the tape did not put Measure down');
      toggleMarkTool(true);
      if (tapeToolActive) throw new Error('Lines did not put the tape down');
      toggleMarkTool(false);
      toggleTapeTool(true);
      _elevDockSelect();
      if (tapeToolActive) throw new Error('Select did not put the tape down');
    });
    __check('a reading belongs to its wall: changing walls clears it', () => {
      seed(pair());
      toggleTapeTool(true); click(0, 0); click(30, 0);
      _elevLoadWall(1);
      if (_tapeA || _tapeB) throw new Error('reading carried to the next wall');
      toggleTapeTool(false);
    });
  `;

  const quiet = console.error; console.error = () => {};
  try {
    window.eval(src + '\n' + testBlock);
  } catch (e) {
    console.error = quiet;
    console.error('LOAD/RUN FAILED:', e.message);
    process.exit(1);
  }
  console.error = quiet;
  const all = window.__testResults || [];
  const bad = all.filter(r => !r.ok);
  all.forEach(r => console.log((r.ok ? 'OK:  ' : 'FAIL:') + ' ' + r.label + (r.ok ? '' : ' -> ' + r.err)));
  console.log('\n--- Summary ---');
  console.log(bad.length ? bad.length + ' FAILURES' : 'ALL PASSED (' + all.length + ')');
  process.exit(bad.length ? 1 : 0);
})();
