# YES TECH LED Layout Designer

Small browser-based planning tool for YES TECH `MG9`, `MG12`, and `MG13` LED panels.

## What it does

- Tracks total stock and used stock for each panel type
- Lets you add panels manually and drag them around a snap-aware layout canvas
- Snaps panels by connector anchor points instead of only by loose bounding boxes
- Generates authored 5-panel-high text layouts from a glyph library
- Supports multiline sample-sheet layouts for visual glyph tuning
- Shows overall layout width, height, and total panel count

## Real-world assumptions used

- `MG9`: `500 x 500 mm`
- `MG12`: triangle panel in a `500 x 500 mm` creative footprint
- `MG13`: quarter-circle panel in a `500 x 500 mm` creative footprint

## Default stock loaded into the tool

- `MG9`: `320`
- `MG12`: `20`
- `MG13`: `20`

## Glyph tuning

The reference glyphs live in `glyph-library.js`.

- `window.GLYPH_LIBRARY` holds the per-character authored rows
- `window.GLYPH_TOKEN_MAP` maps each token to an MG panel type and rotation
- `Load Reference Sheet` in the UI fills the text box with the sample alphabet/numbers for fast comparison

If your exact cabinets use different mechanical connector positions, we can tune the snap points easily in `app.js`.

## Run it

Open `index.html` in a browser.
