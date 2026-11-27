// 19.53, "when I select all guides off my CL does not get hidden it is the centre between I
// drew using the new line tool", and "when I turn them back on all the layout guides get
// turned on including grid, letter labels, frame dim... I want to be able to click on and off
// the buttons I currently have selected. Maybe I need two on off buttons."
//   - free lines are a guide: All off hides them, and they have their own button
//   - My guides: off remembers the selection (in the project), on puts exactly it back
//   - All: every guide on or off (its off remembers too). Removed in 20.86: one switch.
const { JSDOM } = require('jsdom');
const fs = require('fs');
const path = require('path');
(async () => {
  const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
  const htmlSrc = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
  const dom = new JSDOM(htmlSrc, { url: 'https://example.com/', pretendToBeVisual: true, runScripts: 'outside-only' });
  const { window } = dom;
  window.HTMLCanvasElement.prototype.getContext = () => null;
  window.fetch = () => Promise.reject(new Error('no network in test'));
  global.window = window; global.document = window.document; global.navigator = window.navigator;
  const testBlock = `
    window.__testResults = [];
    const __check = (label, fn) => { try { fn(); window.__testResults.push({ label, ok: true }); } catch (e) { window.__testResults.push({ label, ok: false, err: e.message }); } };
    _autosaveToldUser = true; scheduleAutosave = () => {}; pushHistory = () => {};
    const L = (id) => document.getElementById(id).style.display === 'block';
    const pick = () => {   // a designer's selection: spacing, hang/CL and free lines, nothing else
      _elevGuidesAllSet(false);
      ['dim-layer', 'guide-layer', 'mark-layer'].forEach(id => { document.getElementById(id).style.display = 'block'; });
      dimVisibility.wallDims = true; saveDimVisibility();
    };
    __check('EXACT BUG: My guides off hides the free lines (a drawn CL)', () => {
      pick();
      if (!L('mark-layer')) throw new Error('fixture: free lines not shown');
      toggleMyElevGuides();
      if (L('mark-layer')) throw new Error('free lines still showing with guides off');
    });
    __check('EXACT ASK: My guides off then on puts back exactly the selection, not grid / letters / frame dims', () => {
      pick();
      toggleMyElevGuides();
      if (_elevAnyGuideOn()) throw new Error('something still on after My guides off');
      toggleMyElevGuides();
      ['dim-layer', 'guide-layer', 'mark-layer'].forEach(id => { if (!L(id)) throw new Error(id + ' not back'); });
      ['grid-layer', 'label-layer', 'od-layer'].forEach(id => { if (L(id)) throw new Error(id + ' came on though it was off'); });
      if (!dimVisibility.wallDims) throw new Error('wall dims not back');
      if (dimVisibility.edgeGap) throw new Error('edge gaps came on');
    });
    __check('with nothing remembered yet, My guides on shows every guide, and that is never remembered as mine', () => {
      _elevGuidesAllSet(false);
      editorialContent.elevGuidesMine = null;
      toggleMyElevGuides();    // nothing remembered: everything
      ['dim-layer', 'guide-layer', 'grid-layer'].forEach(id => { if (!L(id)) throw new Error(id + ' not on'); });
      toggleMyElevGuides();    // off: "everything" is not a selection
      if (editorialContent.elevGuidesMine) throw new Error('remembered the everything state as mine');
    });
    __check('a guide switched by hand after that IS the new selection', () => {
      _elevGuidesAllSet(false); editorialContent.elevGuidesMine = null;
      toggleMyElevGuides();    // everything
      toggleElevLayer('grid-layer', document.getElementById('gridToggle'));   // grid off by hand: a real choice
      toggleMyElevGuides();    // off, remembers this hand-made state
      toggleMyElevGuides();    // on
      if (L('grid-layer')) throw new Error('grid back though it was switched off by hand');
      if (!L('label-layer')) throw new Error('the hand-made selection (all but grid) was not remembered');
    });
    __check('the remembered selection is saved in the project', () => {
      pick(); toggleMyElevGuides();
      const m = editorialContent.elevGuidesMine;
      if (!m || !m.layers || m.layers['guide-layer'] !== true || m.layers['grid-layer'] !== false) throw new Error(JSON.stringify(m));
    });
    __check('20.86: one master switch, and it says what it will do next', () => {
      if (document.getElementById('allGuidesToggle')) throw new Error('the All on / off button is back');
      if (typeof toggleAllElevGuides !== 'undefined') throw new Error('toggleAllElevGuides is back');
      pick(); _elevGuidesMasterSync();
      if (document.getElementById('myGuidesLbl').textContent !== 'My guides off') throw new Error('label while on');
      toggleMyElevGuides();
      if (document.getElementById('myGuidesLbl').textContent !== 'My guides on') throw new Error('label while off');
    });
    __check('arming the Lines tool shows free lines if they were hidden', () => {
      _elevGuidesAllSet(false);
      try { toggleMarkTool(true); } catch (e) {}
      if (!L('mark-layer')) throw new Error('drawing into a hidden layer');
      try { toggleMarkTool(false); } catch (e) {}
    });
  `;
  try { window.eval(src + '\n' + testBlock); } catch (e) { console.error('LOAD/RUN FAILED:', e.message); process.exit(1); }
  const all = window.__testResults || [];
  const bad = all.filter(r => !r.ok);
  all.forEach(r => console.log((r.ok ? 'OK:  ' : 'FAIL:') + ' ' + r.label + (r.ok ? '' : ' -> ' + r.err)));
  console.log('\n--- Summary ---');
  console.log(bad.length ? bad.length + ' FAILURES' : 'ALL PASSED (' + all.length + ')');
  process.exit(bad.length ? 1 : 0);
})();
