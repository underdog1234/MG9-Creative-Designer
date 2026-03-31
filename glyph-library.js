// Hand-authored 5-panel-high glyph library.
// Edit these rows directly to tune the lettering against the YES TECH sample sheet.
// Simple glyphs use string rows. Tricky glyphs can use explicit panel coordinates:
// { width, height, panels: [{ token, x, y }] } where x/y are in panel units and can use 0.5 offsets.
//
// Token legend:
// . = empty
// S = MG9 square
// A = MG12 triangle, right angle at bottom-left
// B = MG12 triangle, right angle at top-left
// C = MG12 triangle, right angle at top-right
// D = MG12 triangle, right angle at bottom-right
// a = MG13 quarter-circle, rounded top-left
// b = MG13 quarter-circle, rounded top-right
// c = MG13 quarter-circle, rounded bottom-right
// d = MG13 quarter-circle, rounded bottom-left

window.REFERENCE_SAMPLE_TEXT = `123456789
ABCDEFGHI
JKLMNOPQR
STUVWXYZ`;

window.GLYPH_TOKEN_MAP = {
  S: { type: "MG9", rotation: 0 },
  A: { type: "MG12", rotation: 0 },
  B: { type: "MG12", rotation: 90 },
  C: { type: "MG12", rotation: 180 },
  D: { type: "MG12", rotation: 270 },
  a: { type: "MG13", rotation: 0 },
  b: { type: "MG13", rotation: 90 },
  c: { type: "MG13", rotation: 180 },
  d: { type: "MG13", rotation: 270 },
};

window.GLYPH_LIBRARY = {
  "0": ["aSSb", "S..S", "S..S", "S..S", "dSSc"],
  "1": ["S", "S", "S", "S", "S"],
  "2": ["SSSb", "..Sc", ".SS.", "S...", "dSSS"],
  "3": ["SSSb", "..S.", ".SS.", "..S.", "dSSc"],
  "4": ["S..S", "S..S", "dSSc", "...S", "...S"],
  "5": ["aSSS", "S...", "dSSb", "...S", "dSSc"],
  "6": ["aSSb", "S...", "aSSb", "S..S", "dSSc"],
  "7": {
    width: 2.5,
    height: 5,
    panels: [
      { token: "S", x: 1.5, y: 0 },
      { token: "S", x: 1.0, y: 1 },
      { token: "S", x: 0.5, y: 2 },
      { token: "S", x: 0.0, y: 3 },
    ],
  },
  "8": ["aSSb", "S..S", "dSSc", "S..S", "dSSc"],
  "9": ["aSSb", "S..S", "dSSc", "...S", "dSSc"],
  A: ["aSSb", "S..S", "SSSS", "S..S", "S..S"],
  B: ["SSSb", "S..S", "SSSc", "S..S", "dSSc"],
  C: ["aSSS", "S...", "S...", "S...", "dSSS"],
  D: ["SSSb", "S..S", "S..S", "S..S", "dSSc"],
  E: ["SSSC", "S...", "SSS.", "S...", "SSSD"],
  F: ["SSSC", "S...", "SSC.", "S...", "S..."],
  G: ["aSSb", "S...", "S.Sb", "S..S", "dSSc"],
  H: ["S..S", "S..S", "SSSS", "S..S", "S..S"],
  I: ["SSS", ".S.", ".S.", ".S.", "SSS"],
  J: ["..S", "..S", "..S", "..S", "dSS"],
  K: ["S..C", "S.B.", "SS..", "S.D.", "S..A"],
  L: ["S...", "S...", "S...", "S...", "dSSS"],
  M: ["S.B.S", "SBASB", "S.A.S", "S...S", "S...S"],
  N: ["SB..S", "S.B.S", "S..BS", "S...S", "S...S"],
  O: ["aSSb", "S..S", "S..S", "S..S", "dSSc"],
  P: ["aSSb", "S..S", "dSSc", "S...", "S..."],
  Q: ["aSSb", "S..S", "S..S", "S..S", "dSSD"],
  R: ["aSSb", "S..S", "dSSc", "S..B", "S...A"],
  S: ["aSSb", "S...", "dSSb", "...S", "dSSc"],
  T: ["SSSS", ".SS.", ".SS.", ".SS.", ".SS."],
  U: ["S..S", "S..S", "S..S", "S..S", "dSSc"],
  V: ["S...S", "S...S", ".S.S.", ".B.D.", "..A.."],
  W: ["S...S", "S...S", "S.A.S", "SB.BS", "S...S"],
  X: ["B...C", ".B.D.", "..A..", ".B.D.", "A...D"],
  Y: {
    width: 2.5,
    height: 5,
    panels: [
      { token: "S", x: 0.5, y: 0 },
      { token: "S", x: 1.5, y: 0 },
      { token: "S", x: 1.0, y: 1 },
      { token: "S", x: 1.0, y: 2 },
      { token: "S", x: 1.0, y: 3 },
      { token: "S", x: 1.0, y: 4 },
    ],
  },
  Z: ["SSSC", "..S.", ".S..", "S...", "DSSS"],
  " ": ["..", "..", "..", "..", ".."],
};
