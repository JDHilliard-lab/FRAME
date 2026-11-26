# Farmhand (FRAME)

A browser-based presentation builder for Farmboy Fine Arts. It replaces manual InDesign work:
artwork lists, wall elevations, artwork spec pages and client PDF decks, built from one project.

- **Stable:** https://jdhilliard-lab.github.io/FRAME/
- **Dev:** https://jdhilliard-lab.github.io/FRAME-dev/

Projects are saved on your device (autosave and backups live in the browser). Once opened,
the app also works offline.

## Working on it

Plain JavaScript with no build step: `index.html`, `app.js`, `templates.js` and `style.css`,
plus vendored libraries (`lib-jspdf.min.js`, `lib-jszip.min.js`, `lib-html2canvas.min.js`).
Open `index.html` in a browser to run it.

```
npm install    # jsdom, for the tests
npm test       # runs node tests/run-all.js; must print ALL GREEN before shipping
```

Changes go to dev first. `node tools/promote.js` (dry run) then `--push` copies dev to stable.
Rules for each area of the app are in `CLAUDE.md` and `docs/claude/`.
