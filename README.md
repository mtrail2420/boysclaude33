# Trail & Bish Dynasty

A static site, no server or build step required. Every page fetches its data from
plain JSON files in `data/`, so the whole thing runs off files on disk.

## Pages
- `index.html` — Scoreboard (home)
- `roster.html` — Full Roster (all picks, search/filter/sort)
- `draft-classes.html` — Year-by-year draft classes
- `wildcards.html` — Wildcard Boys / Cooked Meter
- `analytics.html` — Power Index, Position DNA, Timeline
- `rules.html` — Scoring formula and tier thresholds

## Local preview
Static sites need to be served over HTTP for `fetch()` to work (opening the
file directly with a double-click will not load the data). From this folder:

```
python3 -m http.server 8000
```

Then open `http://localhost:8000`.

## Deploy
Push this folder to a GitHub repo and turn on GitHub Pages (Settings → Pages →
deploy from branch), or drag the folder into Vercel/Netlify. No build command
needed — it's just static files.

## Updating the data each year
When the workbook (`Trail_Bish_Dynasty_Premium.xlsx`) is updated — March after
awards, or April after the draft — regenerate the JSON:

```
python3 scripts/build_data.py /path/to/Trail_Bish_Dynasty_Premium.xlsx
```

This overwrites `data/players.json`, `data/man_status.json`, and
`data/wildcards.json`. Nothing else needs to change — commit and redeploy.

## Changing the look
- `assets/theme.css` — shared colors, fonts, base layout used on every page
- `assets/nav.js` — the bottom tab bar
- Each page has its own `<style>` block for page-specific layout
