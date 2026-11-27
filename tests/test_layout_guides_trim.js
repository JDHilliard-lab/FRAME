// LAYOUT GUIDES, TRIMMED (20.86).
//
// Asked: "can I clean up or refine the layout guides in the elevations section ... have you
// noticed any I do not use or seem useless now?" Read from the studio's saved projects (HUB
// Phase B is the one that stores its switches): free lines were never drawn in any of them
// while Lines and Free lines had a button each; grid and centres were never left on; Auto
// space sat among the show/hide tiles though it ADDS or CLEARS spacing; EQ is a style, not a
// guide; and All on lit the grid, centres and outlines on top of the guides in use.
//   - one Lines button shows or hides measurement lines AND free lines
//   - Auto space is a labelled action that says what the next click does
//   - EQ is a tick box in Elevation settings
//   - Grid and Centers live in Drafting aids, folded until opened
//   - My guides is the one master switch
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
    const $ = (id) => document.getElementById(id);
    const wall = () => elevations[currentElevIndex];
    const freeLineOnly = () => {
      wall().customLines = [];
      wall().marks = [{ id: 'mk1', kind: 'line', a: { x: 10, y: 40 }, b: { x: 60, y: 40 }, arrow: 'none' }];
    };

    // ── Lines ────────────────────────────────────────────────────────────
    __check('EXACT ASK: there is one Lines button; Free lines has none of its own', () => {
      if ($('markLayerToggle')) throw new Error('#markLayerToggle is still in the panel');
      const b = $('customLinesToggle');
      if (!b || !b.closest('details.lg-group[data-lg="dims"]')) throw new Error('Lines is not in Dimensions');
      if (b.title.indexOf('free lines') < 0) throw new Error('its tooltip does not say it covers free lines: ' + b.title);
    });
    __check('EXACT BUG: a wall with only free lines (a drawn CL) has Lines available, not faded', () => {
      freeLineOnly();
      syncLayoutGuideButtonStates();
      if ($('customLinesToggle').getAttribute('aria-disabled') === 'true') throw new Error('Lines is faded on a wall with a free line');
    });
    __check('Lines off hides measurement lines AND free lines; on brings both back', () => {
      freeLineOnly();
      $('custom-lines-layer').style.display = 'block'; $('mark-layer').style.display = 'block'; dimVisibility.customLines = true;
      toggleCustomLinesVisibility($('customLinesToggle'));
      if (L('mark-layer') || L('custom-lines-layer') || dimVisibility.customLines) throw new Error('something still showing: ' + JSON.stringify({ mark: L('mark-layer'), measure: L('custom-lines-layer'), flag: dimVisibility.customLines }));
      if ($('customLinesToggle').classList.contains('active')) throw new Error('button still lit');
      toggleCustomLinesVisibility($('customLinesToggle'));
      if (!L('mark-layer') || !L('custom-lines-layer') || !dimVisibility.customLines) throw new Error('not both back');
      if (!$('customLinesToggle').classList.contains('active')) throw new Error('button not lit');
    });
    __check('switching Lines is saved with the project, free lines included', () => {
      freeLineOnly();
      $('custom-lines-layer').style.display = 'block'; $('mark-layer').style.display = 'block'; dimVisibility.customLines = true;
      toggleCustomLinesVisibility($('customLinesToggle'));
      const g = editorialContent.elevGuides;
      if (!g || g.layers['mark-layer'] !== false) throw new Error('free lines off was not recorded: ' + JSON.stringify(g && g.layers));
      toggleCustomLinesVisibility($('customLinesToggle'));
    });
    __check('a project saved with the old Free lines switch still opens and shows its state', () => {
      _elevGuidesApply({ layers: { 'mark-layer': false, 'custom-lines-layer': true }, dims: { customLines: true } }, false);
      if (L('mark-layer')) throw new Error('free lines shown though the project had them off');
      _elevGuidesApply({ layers: { 'mark-layer': true }, dims: {} }, false);
      if (!L('mark-layer')) throw new Error('free lines not shown');
    });
    __check('arming the free-line tool lights Lines', () => {
      freeLineOnly();
      toggleCustomLinesVisibility($('customLinesToggle'));   // off
      if (L('mark-layer')) throw new Error('fixture');
      try { toggleMarkTool(true); } catch (e) {}
      if (!L('mark-layer')) throw new Error('drawing into a hidden layer');
      if (!$('customLinesToggle').classList.contains('active')) throw new Error('Lines button not lit');
      try { toggleMarkTool(false); } catch (e) {}
    });

    // ── Auto space ───────────────────────────────────────────────────────
    __check('EXACT ASK: Auto space is a labelled action, not one of the show/hide tiles', () => {
      const b = $('autoSpacingToggle');
      if (!b) throw new Error('missing');
      if (b.closest('.layout-guides-grid')) throw new Error('still inside the tile grid');
      if (!$('autoSpacingLbl') || !$('autoSpacingLbl').textContent.trim()) throw new Error('no label');
    });
    __check('Auto space says what the next click does', () => {
      elevFrames.forEach(f => { f.dimTo = []; });
      _syncMagnetBtn();
      if ($('autoSpacingLbl').textContent.indexOf('Auto space') < 0) throw new Error('with no spacing: ' + $('autoSpacingLbl').textContent);
      if (elevFrames.length) {
        const f = elevFrames.find(x => x.active) || elevFrames[0]; f.active = true; f.dimTo = ['X'];
        _syncMagnetBtn();
        if ($('autoSpacingLbl').textContent.indexOf('Clear') < 0) throw new Error('with spacing: ' + $('autoSpacingLbl').textContent);
        f.dimTo = [];
      }
    });

    // ── EQ ───────────────────────────────────────────────────────────────
    __check('EXACT ASK: EQ is a tick box in Elevation settings, not a guide tile', () => {
      const b = $('spacingEQToggle');
      if (!b || b.type !== 'checkbox') throw new Error('not a tick box');
      if (!b.closest('#precisionModal')) throw new Error('not in Elevation settings');
      if (b.closest('#elev-sec-layout')) throw new Error('still in Layout Guides');
    });
    __check('ticking EQ sets it, and the box follows the setting', () => {
      const b = $('spacingEQToggle');
      dimVisibility.spacingEQ = false; b.checked = true; toggleSpacingEQ(b);
      if (dimVisibility.spacingEQ !== true) throw new Error('tick did not set it');
      b.checked = false; toggleSpacingEQ(b);
      if (dimVisibility.spacingEQ !== false) throw new Error('untick did not clear it');
      dimVisibility.spacingEQ = true; syncLayoutGuideButtonStates();
      if (!b.checked) throw new Error('box does not show the setting');
      dimVisibility.spacingEQ = false; syncLayoutGuideButtonStates();
    });

    // ── Drafting aids and the master switch ─────────────────────────────
    __check('EXACT ASK: Grid and Centers are in Drafting aids, folded until opened', () => {
      const d = document.querySelector('details.lg-group[data-lg="aids"]');
      if (!d) throw new Error('no Drafting aids group');
      if (!d.contains($('gridToggle')) || !d.contains($('frameCenterToggle'))) throw new Error('grid / centers not in it');
      if (d.hasAttribute('open')) throw new Error('starts open');
    });
    __check('opening Drafting aids is remembered on this machine', () => {
      localStorage.setItem('frame_elev_guide_groups', JSON.stringify({ aids: true }));
      const d = document.querySelector('details.lg-group[data-lg="aids"]');
      d.open = false;
      _elevGuideGroupsInit();
      if (!d.open) throw new Error('opened group came back folded');
      localStorage.removeItem('frame_elev_guide_groups');
    });
    __check('one master switch: no All on / off', () => {
      if ($('allGuidesToggle') || $('allGuidesLbl')) throw new Error('All button still there');
      if (!$('myGuidesToggle')) throw new Error('My guides is gone');
    });
  `;
  const quiet = console.error; console.error = () => {};
  try { window.eval(src + '\n' + testBlock); } catch (e) { console.error = quiet; console.error('LOAD/RUN FAILED:', e.message); process.exit(1); }
  console.error = quiet;
  const all = window.__testResults || [];
  const bad = all.filter(r => !r.ok);
  all.forEach(r => console.log((r.ok ? 'OK:  ' : 'FAIL:') + ' ' + r.label + (r.ok ? '' : ' -> ' + r.err)));
  console.log('\n--- Summary ---');
  console.log(bad.length ? bad.length + ' FAILURES' : 'ALL PASSED (' + all.length + ')');
  process.exit(bad.length ? 1 : 0);
})();
