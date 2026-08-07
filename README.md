# YES TECH LED Layout Designer

Small browser-based planning tool for YES TECH `MG9`, `MG12`, and `MG13` LED panels.

## What's New in v1.6.0

- Added a Bulk Grid Placement tool (Panel Library section): enter columns x rows (MG9 squares only, up to 60 per side / 1200 panels total), click **Create Grid** to arm placement mode, and a semi-transparent preview of the whole grid follows the cursor — snapping to the canvas grid and nearby panels exactly like a single panel. Click once to drop the entire grid at once; placement mode ends automatically afterwards, and every placed panel behaves exactly like an individually added one (selectable, draggable, counted in stock, etc.)
- Added a custom favicon/app icon so the tool looks polished in the browser tab and when saved as a shortcut
- Reworked the whole UI into a dark theme (sidebar, toolbar, canvas, inputs, buttons, and connector/selection colours all retuned for a dark background)

## What's New in v1.5.0

- Fixed canvas auto-expansion: the canvas now grows automatically whenever a panel is placed outside the current view, instead of staying frozen at its previous size. Existing panel positions are unaffected — only the visible/exportable canvas area grows
- Added MG9 Push-Out panels: mark any MG9 as Push-Out from the Selected Panel panel (and revert it back to a standard MG9 the same way). Push-Out panels get a distinct violet colour and a "PO" label on the canvas, while still counting as normal MG9 panels for stock, orientation, and panel totals
- PDF exports now show a total Push-Out count and a legend swatch explaining the Push-Out colour/label when any are present in the layout
- Fixed PNG (and PDF) export cropping/scaling incorrectly when the canvas was zoomed in — exports now always render the complete layout at its actual bounds, regardless of the on-screen zoom level or scroll position

## What's New in v1.4.0

- The version badge is now a link to the GitHub release/changelog for the running version
- Fixed group snapping: dragging or pasting multiple selected panels now snaps the whole group together as one rigid unit when it lands near other panels, instead of letting individual panels in the group snap to different neighbours and fall out of alignment

## What's New in v1.3.0

- Panels can now be rotated in 45° steps or to any custom angle: use the **Rotate 45°** / **Rotate 90°** buttons, type an exact value in the **Angle (°)** field, or press `R` (90°) / `Shift+R` (45°)
- Shaped panels rotated off-axis are counted against their nearest 90° orientation bucket for stock

## What's New in v1.2.0

- Triangle panels no longer join on their long (hypotenuse) side — only on their two legs
- Shaped panels (`MG12`, `MG13`) now track stock per orientation (↖ Left-Up, ↙ Left-Down, ↗ Right-Up, ↘ Right-Down), and availability is checked against the matching orientation bucket
- PDF export renders on a white background (no more black page) and lists total panels used plus a per-type / per-orientation breakdown
- New PNG preview export: pick one solid colour and download a transparent-background preview of the whole layout
- Copy (`Ctrl/Cmd+C`) and Paste-at-cursor (`Ctrl/Cmd+V`) keep multi-panel selections grouped with their exact spacing; Duplicate is also fixed to preserve grouping
- Drag a marquee on the empty canvas to select multiple panels at once
- Corrected reference glyphs for `2`, `3`, `B`, `K`, `N`, `Y`
- Rebuilt rendering and connection detection (spatial hashing, geometry caching, transform-based drag) so large layouts stay smooth

## What it does

- Tracks stock and used counts for each panel type, and per orientation for shaped panels
- Lets you add panels manually and drag them around a snap-aware layout canvas
- Snaps panels by connector anchor points instead of only by loose bounding boxes
- Generates authored 5-panel-high text layouts from a glyph library
- Supports multiline sample-sheet layouts for visual glyph tuning
- Shows overall layout width, height, and total panel count
- Exports to PDF (documentation) and PNG (single-colour preview)

## Real-world assumptions used

- `MG9`: `500 x 500 mm`
- `MG12`: triangle panel in a `500 x 500 mm` creative footprint
- `MG13`: quarter-circle panel in a `500 x 500 mm` creative footprint

## Default stock loaded into the tool

- `MG9`: `320`
- `MG12`: `20` (split evenly across the four orientations — `5` each)
- `MG13`: `20` (split evenly across the four orientations — `5` each)

Older project files that stored a single number for `MG12`/`MG13` are migrated automatically by
splitting that number evenly across the four orientation buckets.

## Glyph tuning

The reference glyphs live in `glyph-library.js`.

- `window.GLYPH_LIBRARY` holds the per-character authored rows
- `window.GLYPH_TOKEN_MAP` maps each token to an MG panel type and rotation
- `Load Reference Sheet` in the UI fills the text box with the sample alphabet/numbers for fast comparison

If your exact cabinets use different mechanical connector positions, we can tune the snap points easily in `app.js`.

## Run it

Open `index.html` in a browser.
