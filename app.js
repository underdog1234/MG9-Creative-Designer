const SVG_NS = "http://www.w3.org/2000/svg";
const APP_VERSION = "1.7.1";
const MM_TO_UNITS = 0.25;
const PANEL_SIZE_MM = 500;
const PANEL_SIZE_UNITS = PANEL_SIZE_MM * MM_TO_UNITS;
const HALF_PANEL = PANEL_SIZE_UNITS / 2;
const SNAP_DISTANCE_UNITS = 8;
const PLACEMENT_LOCK_DISTANCE = 52;
const ROTATION_STEP = 90;
const ROTATION_STEP_FINE = 45;
const CANVAS_PADDING = 180;
// Native LED pixel resolution per panel edge. A panel is PANEL_SIZE_MM (500mm)
// physically but PANEL_PIXEL_SIZE pixels of actual LEDs, so raster exports must
// scale by pixels-per-unit derived from this, not by an arbitrary on-screen size.
const PANEL_PIXEL_SIZE = 168;
const EXPORT_PIXELS_PER_UNIT = PANEL_PIXEL_SIZE / PANEL_SIZE_UNITS;
const AUTO_REFRESH_DELAY_MS = 120;
const MAX_GRID_DIMENSION = 60;
const MAX_GRID_CELLS = 1200;
const REFERENCE_SAMPLE_TEXT =
  window.REFERENCE_SAMPLE_TEXT ||
  `123456789
ABCDEFGHI
JKLMNOPQR
STUVWXYZ`;

const PANEL_TYPES = {
  MG9: {
    id: "MG9",
    name: "MG9 Square",
    widthMm: 500,
    heightMm: 500,
    color: "#e48a52",
    shapeKind: "rect",
  },
  MG12: {
    id: "MG12",
    name: "MG12 Triangle",
    widthMm: 500,
    heightMm: 500,
    color: "#d45b5b",
    shapeKind: "triangle",
  },
  MG13: {
    id: "MG13",
    name: "MG13 Quarter Circle",
    widthMm: 500,
    heightMm: 500,
    color: "#4f6d8c",
    shapeKind: "sector",
  },
};

// Push-Out is a per-instance flag on MG9 panels (not a separate stock type),
// so push-out panels stay in the normal MG9 pool for counts/stock/orientation
// and only differ visually and in export labelling.
const PUSH_OUT_COLOR = "#8a4fd6";
const PUSH_OUT_LABEL = "PO";

const LETTER_PATTERNS = {
  A: ["01110", "10001", "11111", "10001", "10001"],
  B: ["11110", "10001", "11110", "10001", "11110"],
  C: ["01111", "10000", "10000", "10000", "01111"],
  D: ["11110", "10001", "10001", "10001", "11110"],
  E: ["11111", "10000", "11110", "10000", "11111"],
  F: ["11111", "10000", "11110", "10000", "10000"],
  G: ["01111", "10000", "10111", "10001", "01111"],
  H: ["10001", "10001", "11111", "10001", "10001"],
  I: ["11111", "00100", "00100", "00100", "11111"],
  J: ["00111", "00010", "00010", "10010", "01100"],
  K: ["10001", "10010", "11100", "10010", "10001"],
  L: ["10000", "10000", "10000", "10000", "11111"],
  M: ["10001", "11011", "10101", "10001", "10001"],
  N: ["10001", "11001", "10101", "10011", "10001"],
  O: ["01110", "10001", "10001", "10001", "01110"],
  P: ["11110", "10001", "11110", "10000", "10000"],
  Q: ["01110", "10001", "10001", "10011", "01111"],
  R: ["11110", "10001", "11110", "10010", "10001"],
  S: ["01111", "10000", "01110", "00001", "11110"],
  T: ["11111", "00100", "00100", "00100", "00100"],
  U: ["10001", "10001", "10001", "10001", "01110"],
  V: ["10001", "10001", "10001", "01010", "00100"],
  W: ["10001", "10001", "10101", "11011", "10001"],
  X: ["10001", "01010", "00100", "01010", "10001"],
  Y: ["10001", "01010", "00100", "00100", "00100"],
  Z: ["11111", "00010", "00100", "01000", "11111"],
  "0": ["01110", "10011", "10101", "11001", "01110"],
  "1": ["00100", "01100", "00100", "00100", "01110"],
  "2": ["01110", "00001", "00110", "01000", "11111"],
  "3": ["11110", "00001", "00110", "00001", "11110"],
  "4": ["10010", "10010", "11111", "00010", "00010"],
  "5": ["11111", "10000", "11110", "00001", "11110"],
  "6": ["01111", "10000", "11110", "10001", "01110"],
  "7": ["11111", "00010", "00100", "01000", "01000"],
  "8": ["01110", "10001", "01110", "10001", "01110"],
  "9": ["01110", "10001", "01111", "00001", "11110"],
  " ": ["000", "000", "000", "000", "000"],
  "&": ["01110", "10010", "01100", "10100", "01101"],
  "@": ["01110", "10001", "10111", "10000", "01110"],
  "#": ["01010", "11111", "01010", "11111", "01010"],
  "!": ["00100", "00100", "00100", "00000", "00100"],
  "?": ["01110", "00001", "00110", "00000", "00100"],
};

const GLYPH_TOKEN_MAP = window.GLYPH_TOKEN_MAP || {
  S: { type: "MG9", rotation: 0 },
};

const GLYPH_LIBRARY = window.GLYPH_LIBRARY || {
  " ": ["..", "..", "..", "..", ".."],
};

// Shaped panels (triangle + quarter-circle) are stocked per orientation. Each
// physical piece points one of four ways, keyed by the location of its
// right-angle corner after rotation.
const SHAPED_TYPES = ["MG12", "MG13"];
const ORIENTATIONS = [
  { key: "LU", icon: "↖", label: "Left Up" },
  { key: "LD", icon: "↙", label: "Left Down" },
  { key: "RU", icon: "↗", label: "Right Up" },
  { key: "RD", icon: "↘", label: "Right Down" },
];
// Right-angle corner after rotation -> orientation bucket (SVG clockwise).
const TRIANGLE_ORIENTATION = { 0: "LD", 90: "LU", 180: "RU", 270: "RD" };
const SECTOR_ORIENTATION = { 0: "RD", 90: "LD", 180: "LU", 270: "RU" };

function normalizeRotation(rotation) {
  return (((Math.round((Number(rotation) || 0) / 90) * 90) % 360) + 360) % 360;
}

function isShapedType(type) {
  return SHAPED_TYPES.includes(type);
}

function getPanelOrientation(panel) {
  const type = PANEL_TYPES[panel.type];
  if (!type) return null;
  const rot = normalizeRotation(panel.rotation);
  if (type.shapeKind === "triangle") return TRIANGLE_ORIENTATION[rot];
  if (type.shapeKind === "sector") return SECTOR_ORIENTATION[rot];
  return null;
}

function splitEvenly(total) {
  const count = Math.max(0, Math.floor(Number(total) || 0));
  const base = Math.floor(count / 4);
  const remainder = count % 4;
  const buckets = {};
  ORIENTATIONS.forEach((orientation, index) => {
    buckets[orientation.key] = base + (index < remainder ? 1 : 0);
  });
  return buckets;
}

function normalizeInventory(raw) {
  const inventory = { MG9: Math.max(0, Number(raw?.MG9) || 0) };
  SHAPED_TYPES.forEach((type) => {
    const value = raw?.[type];
    if (value && typeof value === "object") {
      inventory[type] = {};
      ORIENTATIONS.forEach((orientation) => {
        inventory[type][orientation.key] = Math.max(0, Number(value[orientation.key]) || 0);
      });
    } else {
      inventory[type] = splitEvenly(value);
    }
  });
  return inventory;
}

function cloneInventory(inventory) {
  return normalizeInventory(inventory);
}

function getStock(type, orientation) {
  if (isShapedType(type)) return Number(state.inventory[type]?.[orientation]) || 0;
  return Number(state.inventory[type]) || 0;
}

function getStockTotal(type) {
  if (isShapedType(type)) {
    return ORIENTATIONS.reduce(
      (total, orientation) => total + (Number(state.inventory[type]?.[orientation.key]) || 0),
      0
    );
  }
  return Number(state.inventory[type]) || 0;
}

const DEFAULT_INVENTORY_RAW = {
  MG9: 320,
  MG12: 20,
  MG13: 20,
};

const defaultInventory = normalizeInventory(DEFAULT_INVENTORY_RAW);

const state = {
  inventory: cloneInventory(defaultInventory),
  panels: [],
  selectedId: null,
  selectedIds: [],
  projectName: "untitled-layout",
  zoom: 1,
  drag: null,
  marquee: null,
  clipboard: null,
  lastPointer: null,
  nextId: 1,
  connectionMap: new Map(),
  autoTextTimer: null,
  history: [],
  currentViewBox: "0 0 2400 1600",
  manualViewLocked: false,
  collapsedSections: {
    textLayout: false,
    selectedPanel: false,
    panelLibrary: true,
    inventory: true,
    help: true,
  },
  placement: {
    active: false,
    type: null,
    grid: null,
    pointer: null,
    preview: null,
  },
};

const els = {
  appVersionBadge: document.querySelector("#appVersionBadge"),
  inventoryList: document.querySelector("#inventoryList"),
  library: document.querySelector("#library"),
  replaceTypeLibrary: document.querySelector("#replaceTypeLibrary"),
  gridColsInput: document.querySelector("#gridColsInput"),
  gridRowsInput: document.querySelector("#gridRowsInput"),
  createGridBtn: document.querySelector("#createGridBtn"),
  canvas: document.querySelector("#layoutCanvas"),
  canvasPanels: document.querySelector("#canvasPanels"),
  canvasGuides: document.querySelector("#canvasGuides"),
  canvasPreview: document.querySelector("#canvasPreview"),
  canvasMarquee: document.querySelector("#canvasMarquee"),
  canvasBackground: document.querySelector("#canvasBackground"),
  canvasGrid: document.querySelector("#canvasGrid"),
  canvasSubgrid: document.querySelector("#canvasSubgrid"),
  selectedMeta: document.querySelector("#selectedMeta"),
  layoutWidth: document.querySelector("#layoutWidth"),
  layoutHeight: document.querySelector("#layoutHeight"),
  panelCount: document.querySelector("#panelCount"),
  connectionStatus: document.querySelector("#connectionStatus"),
  zoomLabel: document.querySelector("#zoomLabel"),
  projectNameInput: document.querySelector("#projectNameInput"),
  textInput: document.querySelector("#textInput"),
  heightMetersInput: document.querySelector("#heightMetersInput"),
  actualWidthOutput: document.querySelector("#actualWidthOutput"),
  spacingInput: document.querySelector("#spacingInput"),
  letterStyleInput: document.querySelector("#letterStyleInput"),
  textSummary: document.querySelector("#textSummary"),
  generateTextBtn: document.querySelector("#generateTextBtn"),
  loadReferenceBtn: document.querySelector("#loadReferenceBtn"),
  saveProjectBtn: document.querySelector("#saveProjectBtn"),
  openProjectBtn: document.querySelector("#openProjectBtn"),
  savePdfBtn: document.querySelector("#savePdfBtn"),
  exportPngBtn: document.querySelector("#exportPngBtn"),
  pngColorInput: document.querySelector("#pngColorInput"),
  rotateBtn: document.querySelector("#rotateBtn"),
  rotateFineBtn: document.querySelector("#rotateFineBtn"),
  rotationInput: document.querySelector("#rotationInput"),
  duplicateBtn: document.querySelector("#duplicateBtn"),
  pushOutBtn: document.querySelector("#pushOutBtn"),
  copyBtn: document.querySelector("#copyBtn"),
  pasteBtn: document.querySelector("#pasteBtn"),
  deleteBtn: document.querySelector("#deleteBtn"),
  undoBtn: document.querySelector("#undoBtn"),
  clearLayoutBtn: document.querySelector("#clearLayoutBtn"),
  resetInventoryBtn: document.querySelector("#resetInventoryBtn"),
  zoomInBtn: document.querySelector("#zoomInBtn"),
  zoomOutBtn: document.querySelector("#zoomOutBtn"),
  projectFileInput: document.querySelector("#projectFileInput"),
  inventoryRowTemplate: document.querySelector("#inventoryRowTemplate"),
  libraryCardTemplate: document.querySelector("#libraryCardTemplate"),
  sectionToggles: document.querySelectorAll("[data-section-toggle]"),
};

const EDGE_CONNECTORS_5 = [-0.8, -0.4, 0, 0.4, 0.8].map((value) => value * HALF_PANEL);
const EDGE_CONNECTORS_3 = [-0.5, 0, 0.5].map((value) => value * HALF_PANEL);
const TEMPLATE_RESOLUTION = 16;

function mmToUnits(mm) {
  return mm * MM_TO_UNITS;
}

function unitsToMm(units) {
  return units / MM_TO_UNITS;
}

function mmToMetersText(mm) {
  return `${(mm / 1000).toFixed(2)} m`;
}

function hexToRgb(hex) {
  const clean = hex.replace("#", "");
  const value = parseInt(clean, 16);
  return [(value >> 16) & 255, (value >> 8) & 255, value & 255];
}

function sanitizeProjectName(name) {
  const cleaned = String(name || "")
    .trim()
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, "-")
    .replace(/\s+/g, " ");
  return cleaned || "untitled-layout";
}

function isTypingTarget(target) {
  if (!(target instanceof HTMLElement)) return false;
  return Boolean(target.closest("input, textarea, select, [contenteditable='true']"));
}

function snapToIncrement(value, increment) {
  return Math.round(value / increment) * increment;
}

function degToRad(deg) {
  return (deg * Math.PI) / 180;
}

function rotatePoint(point, rotation) {
  const angle = degToRad(rotation);
  const cos = Math.cos(angle);
  const sin = Math.sin(angle);
  return {
    x: point.x * cos - point.y * sin,
    y: point.x * sin + point.y * cos,
  };
}

function getGraphemes(text) {
  if (window.Intl?.Segmenter) {
    const segmenter = new Intl.Segmenter(undefined, { granularity: "grapheme" });
    return [...segmenter.segment(text)].map((item) => item.segment);
  }
  return Array.from(text);
}

function buildConnectorGrid() {
  const anchors = [];

  EDGE_CONNECTORS_3.forEach((x) => {
    anchors.push({ x, y: -HALF_PANEL });
    anchors.push({ x, y: HALF_PANEL });
  });

  EDGE_CONNECTORS_5.forEach((y) => {
    anchors.push({ x: -HALF_PANEL, y });
    anchors.push({ x: HALF_PANEL, y });
  });

  return anchors;
}

const SHARED_CONNECTORS = buildConnectorGrid();

function triangleEdgeConnectors() {
  // Triangles may only join on their two legs (left + bottom edges). The
  // hypotenuse (long side) intentionally exposes no connectors, so two
  // triangles can never snap/connect along their long side.
  return [
    ...EDGE_CONNECTORS_5.map((y) => ({ x: -HALF_PANEL, y })),
    ...EDGE_CONNECTORS_5.map((x) => ({ x, y: HALF_PANEL })),
  ];
}

function sectorEdgeConnectors() {
  return [
    ...EDGE_CONNECTORS_5.map((x) => ({ x, y: HALF_PANEL })),
    ...EDGE_CONNECTORS_5.map((y) => ({ x: HALF_PANEL, y })),
  ];
}

function getPanelBaseGeometry(panelType) {
  if (panelType.shapeKind === "triangle") {
    return {
      width: PANEL_SIZE_UNITS,
      height: PANEL_SIZE_UNITS,
      points: [
        { x: -HALF_PANEL, y: -HALF_PANEL },
        { x: -HALF_PANEL, y: HALF_PANEL },
        { x: HALF_PANEL, y: HALF_PANEL },
      ],
      anchors: triangleEdgeConnectors(),
      labelOffset: { x: -28, y: 36 },
    };
  }

  if (panelType.shapeKind === "sector") {
    return {
      width: PANEL_SIZE_UNITS,
      height: PANEL_SIZE_UNITS,
      path: `M ${-HALF_PANEL} ${HALF_PANEL} L ${HALF_PANEL} ${HALF_PANEL} L ${HALF_PANEL} ${-HALF_PANEL} A ${PANEL_SIZE_UNITS} ${PANEL_SIZE_UNITS} 0 0 0 ${-HALF_PANEL} ${HALF_PANEL} Z`,
      anchors: sectorEdgeConnectors(),
      labelOffset: { x: 36, y: 40 },
    };
  }

  return {
    width: PANEL_SIZE_UNITS,
    height: PANEL_SIZE_UNITS,
    points: [
      { x: -HALF_PANEL, y: -HALF_PANEL },
      { x: HALF_PANEL, y: -HALF_PANEL },
      { x: HALF_PANEL, y: HALF_PANEL },
      { x: -HALF_PANEL, y: HALF_PANEL },
    ],
    anchors: SHARED_CONNECTORS,
    labelOffset: { x: 0, y: 0 },
  };
}

// Rotated geometry only depends on (type, rotation) and rotation is always a
// multiple of 90, so cache the handful of variants instead of recomputing the
// trigonometry on every access.
const geometryCache = new Map();

function getGeometry(panel) {
  // Rotate by the panel's true angle (any degree, not just multiples of 90) so
  // custom / 45-degree rotations render correctly. Orientation-stock bucketing
  // still uses normalizeRotation elsewhere.
  const rotation = (((Number(panel.rotation) || 0) % 360) + 360) % 360;
  const key = `${panel.type}:${rotation}`;
  let geometry = geometryCache.get(key);
  if (!geometry) {
    const base = getPanelBaseGeometry(PANEL_TYPES[panel.type]);
    geometry = {
      width: base.width,
      height: base.height,
      path: base.path,
      points: base.points ? base.points.map((point) => rotatePoint(point, rotation)) : undefined,
      anchors: base.anchors.map((anchor) => rotatePoint(anchor, rotation)),
      labelOffset: rotatePoint(base.labelOffset, rotation),
    };
    geometryCache.set(key, geometry);
  }
  return geometry;
}

function createPanel(type, x = 320, y = 260, rotation = 0, pushOut = false) {
  return {
    id: `panel-${state.nextId++}`,
    type,
    x,
    y,
    rotation,
    pushOut: type === "MG9" && Boolean(pushOut),
  };
}

function getPanelById(id) {
  return state.panels.find((panel) => panel.id === id);
}

function getSelectedPanels() {
  const ids = new Set(state.selectedIds.length ? state.selectedIds : state.selectedId ? [state.selectedId] : []);
  return state.panels.filter((panel) => ids.has(panel.id));
}

function snapshotState() {
  return {
    inventory: cloneInventory(state.inventory),
    panels: state.panels.map((panel) => ({ ...panel })),
    selectedId: state.selectedId,
    selectedIds: [...state.selectedIds],
    projectName: state.projectName,
    placement: {
      active: state.placement.active,
      type: state.placement.type,
      grid: state.placement.grid,
    },
  };
}

function pushHistory() {
  state.history.push(snapshotState());
  if (state.history.length > 100) state.history.shift();
}

function undoLastAction() {
  const snapshot = state.history.pop();
  if (!snapshot) return;
  state.inventory = cloneInventory(snapshot.inventory);
  state.panels = snapshot.panels.map((panel) => ({ ...panel }));
  state.selectedId = snapshot.selectedId;
  state.selectedIds = [...snapshot.selectedIds];
  state.projectName = snapshot.projectName || "untitled-layout";
  state.placement.active = snapshot.placement?.active ?? false;
  state.placement.type = snapshot.placement?.type ?? null;
  state.placement.grid = snapshot.placement?.grid ?? null;
  state.placement.preview = null;
  state.placement.pointer = null;
  els.projectNameInput.value = state.projectName;
  applyAvailability();
  computeConnections();
  render();
}

function getUsedCounts() {
  return state.panels.reduce((counts, panel) => {
    counts[panel.type] = (counts[panel.type] || 0) + 1;
    return counts;
  }, {});
}

function applyAvailability() {
  const usedByType = {};
  const usedByOrientation = { MG12: {}, MG13: {} };

  state.panels.forEach((panel) => {
    if (isShapedType(panel.type)) {
      const orientation = getPanelOrientation(panel);
      const counts = usedByOrientation[panel.type];
      counts[orientation] = (counts[orientation] || 0) + 1;
      panel.available = counts[orientation] <= getStock(panel.type, orientation);
    } else {
      usedByType[panel.type] = (usedByType[panel.type] || 0) + 1;
      panel.available = usedByType[panel.type] <= getStock(panel.type);
    }
  });
}

function getUsedOrientationCounts() {
  const counts = {
    MG9: 0,
    MG12: { LU: 0, LD: 0, RU: 0, RD: 0 },
    MG13: { LU: 0, LD: 0, RU: 0, RD: 0 },
  };
  state.panels.forEach((panel) => {
    if (isShapedType(panel.type)) {
      const orientation = getPanelOrientation(panel);
      counts[panel.type][orientation] = (counts[panel.type][orientation] || 0) + 1;
    } else if (panel.type === "MG9") {
      counts.MG9 += 1;
    }
  });
  return counts;
}

function getGlobalAnchors(panel) {
  const geometry = getGeometry(panel);
  return geometry.anchors.map((anchor, index) => ({
    panelId: panel.id,
    index,
    x: panel.x + anchor.x,
    y: panel.y + anchor.y,
  }));
}

// Connection detection uses a uniform spatial hash so it scales ~linearly with
// panel count instead of comparing every anchor against every other anchor.
const CONNECTION_CELL = SNAP_DISTANCE_UNITS;

function getConnectionMapForPanels(panels) {
  const panelConnections = new Map();

  panels.forEach((panel) => {
    panelConnections.set(panel.id, {
      connected: panels.length <= 1,
      anchorIndices: new Set(),
    });
  });

  if (panels.length <= 1) return panelConnections;

  const buckets = new Map();
  const anchors = [];
  panels.forEach((panel) => {
    getGlobalAnchors(panel).forEach((anchor) => {
      anchors.push(anchor);
      const key = `${Math.floor(anchor.x / CONNECTION_CELL)}:${Math.floor(anchor.y / CONNECTION_CELL)}`;
      let bucket = buckets.get(key);
      if (!bucket) {
        bucket = [];
        buckets.set(key, bucket);
      }
      bucket.push(anchor);
    });
  });

  anchors.forEach((anchor) => {
    const cellX = Math.floor(anchor.x / CONNECTION_CELL);
    const cellY = Math.floor(anchor.y / CONNECTION_CELL);
    for (let gx = cellX - 1; gx <= cellX + 1; gx += 1) {
      for (let gy = cellY - 1; gy <= cellY + 1; gy += 1) {
        const bucket = buckets.get(`${gx}:${gy}`);
        if (!bucket) continue;
        for (const other of bucket) {
          if (other.panelId === anchor.panelId) continue;
          if (Math.hypot(anchor.x - other.x, anchor.y - other.y) <= SNAP_DISTANCE_UNITS) {
            panelConnections.get(anchor.panelId).anchorIndices.add(anchor.index);
            break;
          }
        }
      }
    }
  });

  panels.forEach((panel) => {
    panelConnections.get(panel.id).connected =
      panelConnections.get(panel.id).anchorIndices.size >= 2;
  });

  return panelConnections;
}

function computeConnections() {
  const panelConnections = getConnectionMapForPanels(state.panels);
  state.connectionMap = panelConnections;
}

function isPanelConnected(panelId) {
  return state.connectionMap.get(panelId)?.connected ?? false;
}

function addPanel(type) {
  pushHistory();
  state.manualViewLocked = true;
  const panel = findPlacementForNewPanel(type);
  state.panels.push(panel);
  applyAvailability();
  state.selectedId = panel.id;
  state.selectedIds = [panel.id];
  computeConnections();
  render();
}

function removeSelectedPanel() {
  const selectedIds = new Set(getSelectedPanels().map((panel) => panel.id));
  if (!selectedIds.size) return;
  pushHistory();
  state.manualViewLocked = true;
  state.panels = state.panels.filter((panel) => !selectedIds.has(panel.id));
  applyAvailability();
  state.selectedId = null;
  state.selectedIds = [];
  computeConnections();
  render();
}

function duplicateSelectedPanel() {
  const selectedPanels = getSelectedPanels();
  if (!selectedPanels.length) return;
  pushHistory();
  state.manualViewLocked = true;
  // Offset the whole selection by the same delta so the group keeps its exact
  // relative spacing/rotation instead of drifting apart.
  const offset = PANEL_SIZE_UNITS;
  const duplicates = selectedPanels.map((selected) =>
    createPanel(selected.type, selected.x + offset, selected.y + offset, selected.rotation, selected.pushOut)
  );
  state.panels.push(...duplicates);
  applyAvailability();
  state.selectedId = duplicates[0]?.id || null;
  state.selectedIds = duplicates.map((panel) => panel.id);
  computeConnections();
  render();
}

function copySelection() {
  const selectedPanels = getSelectedPanels();
  if (!selectedPanels.length) return;
  const minX = Math.min(...selectedPanels.map((panel) => panel.x));
  const minY = Math.min(...selectedPanels.map((panel) => panel.y));
  state.clipboard = selectedPanels.map((panel) => ({
    type: panel.type,
    rotation: panel.rotation,
    dx: panel.x - minX,
    dy: panel.y - minY,
    pushOut: panel.pushOut,
  }));
}

function pasteClipboard() {
  if (!state.clipboard || !state.clipboard.length) return;
  pushHistory();
  state.manualViewLocked = true;

  const fallbackBase = getPanelById(state.selectedId) || state.panels[state.panels.length - 1];
  const anchor = state.lastPointer || {
    x: (fallbackBase?.x ?? 500) + PANEL_SIZE_UNITS,
    y: (fallbackBase?.y ?? 500) + PANEL_SIZE_UNITS,
  };
  const originX = snapToIncrement(anchor.x, HALF_PANEL);
  const originY = snapToIncrement(anchor.y, HALF_PANEL);

  const pasted = state.clipboard.map((item) =>
    createPanel(item.type, originX + item.dx, originY + item.dy, item.rotation, item.pushOut)
  );
  state.panels.push(...pasted);
  snapPanelGroup(pasted.map((panel) => panel.id), PLACEMENT_LOCK_DISTANCE);
  applyAvailability();
  state.selectedId = pasted[0]?.id || null;
  state.selectedIds = pasted.map((panel) => panel.id);
  computeConnections();
  render();
}

function rotateSelectedPanel(step = ROTATION_STEP) {
  const selectedPanels = getSelectedPanels();
  if (!selectedPanels.length) return;
  pushHistory();
  state.manualViewLocked = true;
  selectedPanels.forEach((selected) => {
    selected.rotation = (((selected.rotation + step) % 360) + 360) % 360;
  });
  // Rotating moves a panel's connectors, which can turn a flush join with a
  // neighbour into a small gap; pull it back to exact alignment if one is
  // still in reach so "connected" always means zero-gap.
  selectedPanels.forEach((selected) => snapPanel(selected, SNAP_DISTANCE_UNITS));
  applyAvailability();
  computeConnections();
  render();
}

function setSelectedRotation(angle) {
  const selectedPanels = getSelectedPanels();
  if (!selectedPanels.length) return;
  const normalized = (((Math.round(Number(angle) || 0) % 360) + 360) % 360);
  pushHistory();
  state.manualViewLocked = true;
  selectedPanels.forEach((selected) => {
    selected.rotation = normalized;
  });
  selectedPanels.forEach((selected) => snapPanel(selected, SNAP_DISTANCE_UNITS));
  applyAvailability();
  computeConnections();
  render();
}

function setSelected(id, options = {}) {
  const { additive = false, toggle = false } = options;
  const selected = new Set(state.selectedIds);

  if (toggle && selected.has(id)) selected.delete(id);
  else if (additive || toggle) selected.add(id);
  else {
    selected.clear();
    selected.add(id);
  }

  state.selectedIds = [...selected];
  state.selectedId = state.selectedIds[0] || null;
  renderSelectedMeta();
  renderCanvas();
}

function clearSelection() {
  state.selectedId = null;
  state.selectedIds = [];
  renderSelectedMeta();
  renderCanvas();
}

function setSectionCollapsed(sectionKey, collapsed) {
  state.collapsedSections[sectionKey] = collapsed;
  renderSections();
}

function renderSections() {
  els.sectionToggles.forEach((toggle) => {
    const sectionKey = toggle.dataset.sectionToggle;
    const collapsed = Boolean(state.collapsedSections[sectionKey]);
    const panel = toggle.closest(".section-panel");
    const body = panel?.querySelector(`[data-section-body="${sectionKey}"]`);
    toggle.setAttribute("aria-expanded", String(!collapsed));
    const icon = toggle.querySelector(".section-toggle-icon");
    if (icon) icon.textContent = collapsed ? "+" : "−";
    panel?.classList.toggle("is-collapsed", collapsed);
    if (body) body.hidden = collapsed;
  });
}

function resetPlacement() {
  state.placement.active = false;
  state.placement.type = null;
  state.placement.grid = null;
  state.placement.pointer = null;
  state.placement.preview = null;
}

function setPlacementMode(type = null, gridConfig = null) {
  state.placement.active = Boolean(type);
  state.placement.type = type;
  state.placement.grid = gridConfig;
  state.placement.pointer = null;
  state.placement.preview = null;
  renderLibrary();
  renderCanvasPreview();
  renderSelectedMeta();
}

function startGridPlacement() {
  const clamp = (value, min, max) => Math.min(max, Math.max(min, Math.round(value) || min));
  const cols = clamp(Number(els.gridColsInput?.value), 1, MAX_GRID_DIMENSION);
  const rows = clamp(Number(els.gridRowsInput?.value), 1, MAX_GRID_DIMENSION);

  if (cols * rows > MAX_GRID_CELLS) {
    window.alert(
      `That grid would be ${cols * rows} panels, which is too large for one placement (limit ${MAX_GRID_CELLS}). Reduce the columns or rows.`
    );
    return;
  }

  if (els.gridColsInput) els.gridColsInput.value = cols;
  if (els.gridRowsInput) els.gridRowsInput.value = rows;
  setPlacementMode("MG9", { cols, rows });
}

function refreshAfterInventoryChange() {
  applyAvailability();
  computeConnections();
  renderInventory();
  renderLibrary();
  renderCanvas();
}

function renderInventory() {
  const usedCounts = getUsedCounts();
  const usedOrientation = getUsedOrientationCounts();
  els.inventoryList.innerHTML = "";

  Object.values(PANEL_TYPES).forEach((type) => {
    if (isShapedType(type.id)) {
      const group = document.createElement("div");
      group.className = "inventory-group";

      const head = document.createElement("div");
      head.className = "inventory-group-head";
      const name = document.createElement("strong");
      name.textContent = type.name;
      const size = document.createElement("span");
      size.textContent = `${type.widthMm} x ${type.heightMm} mm`;
      head.append(name, size);
      group.appendChild(head);

      ORIENTATIONS.forEach((orientation) => {
        const row = document.createElement("div");
        row.className = "inventory-orient-row";

        const label = document.createElement("span");
        label.className = "orient-label";
        label.textContent = `${orientation.icon} ${orientation.label}`;

        const input = document.createElement("input");
        input.type = "number";
        input.min = "0";
        input.step = "1";
        input.value = getStock(type.id, orientation.key);
        input.addEventListener("change", (event) => {
          state.inventory[type.id][orientation.key] = Math.max(0, Number(event.target.value) || 0);
          refreshAfterInventoryChange();
        });

        const used = usedOrientation[type.id][orientation.key] || 0;
        const remaining = Math.max(getStock(type.id, orientation.key) - used, 0);
        const counts = document.createElement("span");
        counts.className = "orient-counts";
        counts.textContent = `Used ${used} · Left ${remaining}`;

        row.append(label, input, counts);
        group.appendChild(row);
      });

      els.inventoryList.appendChild(group);
      return;
    }

    const row = els.inventoryRowTemplate.content.firstElementChild.cloneNode(true);
    row.querySelector(".inventory-name").textContent = type.name;
    row.querySelector(".inventory-size").textContent = `${type.widthMm} x ${type.heightMm} mm`;

    const input = row.querySelector("input");
    input.value = getStockTotal(type.id);
    input.addEventListener("change", (event) => {
      state.inventory[type.id] = Math.max(0, Number(event.target.value) || 0);
      refreshAfterInventoryChange();
    });

    const used = usedCounts[type.id] || 0;
    const remaining = Math.max(getStockTotal(type.id) - used, 0);
    row.querySelector(".used-count").textContent = `Used: ${used}`;
    row.querySelector(".remaining-count").textContent = `Remaining: ${remaining}`;
    els.inventoryList.appendChild(row);
  });
}

function renderLibrary() {
  const usedCounts = getUsedCounts();
  els.library.innerHTML = "";

  Object.values(PANEL_TYPES).forEach((type) => {
    const card = els.libraryCardTemplate.content.firstElementChild.cloneNode(true);
    const shape = card.querySelector(".library-shape");
    const remaining = Math.max(getStockTotal(type.id) - (usedCounts[type.id] || 0), 0);

    shape.dataset.shape = type.shapeKind === "rect" ? "rect" : type.shapeKind;
    shape.style.background = `linear-gradient(135deg, ${type.color}88, ${type.color}33)`;

    card.querySelector(".library-name").textContent = type.name;
    card.querySelector(".library-dimensions").textContent = `${type.widthMm} x ${type.heightMm} mm - ${remaining} left`;
    card.classList.toggle("active", state.placement.active && state.placement.type === type.id);
    card.addEventListener("click", () => {
      if (state.placement.active && state.placement.type === type.id) setPlacementMode(null);
      else setPlacementMode(type.id);
    });
    els.library.appendChild(card);
  });
}

function replaceSelectedPanelType(type) {
  const selectedPanels = getSelectedPanels();
  if (!selectedPanels.length) return;
  pushHistory();
  selectedPanels.forEach((panel) => {
    panel.type = type;
    if (type !== "MG9") panel.pushOut = false;
  });
  // Different panel shapes have different connector layouts, so swapping
  // type can turn a flush join into a small gap; re-settle against any
  // neighbour still in reach.
  selectedPanels.forEach((panel) => snapPanel(panel, SNAP_DISTANCE_UNITS));
  applyAvailability();
  computeConnections();
  render();
}

// Push-Out is a display/labelling flag only: it never changes panel.type, so
// it automatically stays included in every existing MG9 count, stock check,
// and orientation calculation without any further wiring.
function togglePushOut() {
  const selectedMG9 = getSelectedPanels().filter((panel) => panel.type === "MG9");
  if (!selectedMG9.length) return;
  pushHistory();
  const allPushOut = selectedMG9.every((panel) => panel.pushOut);
  selectedMG9.forEach((panel) => {
    panel.pushOut = !allPushOut;
  });
  render();
}

function getPushOutCount() {
  return state.panels.filter((panel) => panel.type === "MG9" && panel.pushOut).length;
}

function renderReplaceTypeLibrary() {
  els.replaceTypeLibrary.innerHTML = "";
  const selectedPanels = getSelectedPanels();
  const singleType = selectedPanels.length === 1 ? selectedPanels[0].type : null;

  Object.values(PANEL_TYPES).forEach((type) => {
    const card = els.libraryCardTemplate.content.firstElementChild.cloneNode(true);
    const shape = card.querySelector(".library-shape");
    shape.dataset.shape = type.shapeKind === "rect" ? "rect" : type.shapeKind;
    shape.style.background = `linear-gradient(135deg, ${type.color}88, ${type.color}33)`;
    card.classList.add("small-card");
    card.classList.toggle("active", singleType === type.id);
    card.querySelector(".library-name").textContent = type.id;
    card.querySelector(".library-dimensions").textContent = "Replace selected";
    card.disabled = selectedPanels.length === 0;
    card.addEventListener("click", () => replaceSelectedPanelType(type.id));
    els.replaceTypeLibrary.appendChild(card);
  });
}

// Panels render inside a `<g transform="translate(x y)">` with all child
// geometry in local (rotation-baked) coordinates, so repositioning during a
// drag only updates the group's transform instead of rebuilding child nodes.
function buildPanelChildren(panel, options = {}) {
  const { anchorClass = "anchor-point", connectedSet = null } = options;
  const type = PANEL_TYPES[panel.type];
  const geometry = getGeometry(panel);
  const fragment = document.createDocumentFragment();
  const isPushOut = panel.type === "MG9" && Boolean(panel.pushOut);
  const fillColor = isPushOut ? PUSH_OUT_COLOR : type.color;

  if (type.shapeKind === "sector") {
    const path = document.createElementNS(SVG_NS, "path");
    path.setAttribute("d", geometry.path);
    if (panel.rotation) path.setAttribute("transform", `rotate(${panel.rotation})`);
    path.setAttribute("class", "panel-shape");
    path.setAttribute("fill", fillColor);
    fragment.appendChild(path);
  } else {
    const polygon = document.createElementNS(SVG_NS, "polygon");
    polygon.setAttribute("points", geometry.points.map((point) => `${point.x},${point.y}`).join(" "));
    polygon.setAttribute("class", "panel-shape");
    polygon.setAttribute("fill", fillColor);
    fragment.appendChild(polygon);
  }

  const text = document.createElementNS(SVG_NS, "text");
  text.setAttribute("x", geometry.labelOffset.x);
  text.setAttribute("y", geometry.labelOffset.y);
  text.setAttribute("class", "panel-label");
  text.textContent = isPushOut ? PUSH_OUT_LABEL : panel.type;
  fragment.appendChild(text);

  if (isPushOut) {
    const title = document.createElementNS(SVG_NS, "title");
    title.textContent = "Push-Out panel (MG9)";
    fragment.appendChild(title);
  }

  const connected = connectedSet || state.connectionMap.get(panel.id)?.anchorIndices || new Set();
  geometry.anchors.forEach((anchor, index) => {
    const circle = document.createElementNS(SVG_NS, "circle");
    circle.setAttribute("cx", anchor.x);
    circle.setAttribute("cy", anchor.y);
    circle.setAttribute("r", 3.2);
    circle.setAttribute("class", `${anchorClass}${connected.has(index) ? " connected" : ""}`);
    fragment.appendChild(circle);
  });

  return fragment;
}

function panelClasses(panel) {
  const isConnected = isPanelConnected(panel.id);
  return [
    "panel-group",
    panel.id === state.selectedId ? "selected" : "",
    state.selectedIds.includes(panel.id) && panel.id !== state.selectedId ? "multi-selected" : "",
    panel.available === false ? "unavailable" : "",
    !isConnected ? "disconnected" : "",
    panel.type === "MG9" && panel.pushOut ? "push-out" : "",
  ]
    .filter(Boolean)
    .join(" ");
}

function getPlacementPreview(type, pointerX, pointerY, rotation = 0) {
  const rawX = pointerX;
  const rawY = pointerY;
  const preview = {
    id: "preview",
    type,
    x: rawX,
    y: rawY,
    rotation,
  };

  const previewAnchors = getGlobalAnchors(preview);
  const candidates = new Map();

  state.panels.forEach((panel) => {
    getGlobalAnchors(panel).forEach((targetAnchor) => {
      previewAnchors.forEach((anchor) => {
        const dx = targetAnchor.x - anchor.x;
        const dy = targetAnchor.y - anchor.y;
        const candidateX = rawX + dx;
        const candidateY = rawY + dy;
        const distanceFromPointer = Math.hypot(candidateX - pointerX, candidateY - pointerY);
        if (distanceFromPointer > PLACEMENT_LOCK_DISTANCE) return;
        const key = `${Math.round(candidateX * 100) / 100}:${Math.round(candidateY * 100) / 100}`;
        if (!candidates.has(key)) {
          candidates.set(key, {
            x: candidateX,
            y: candidateY,
            distanceFromPointer,
          });
        }
      });
    });
  });

  let best = {
    x: snapToIncrement(rawX, HALF_PANEL),
    y: snapToIncrement(rawY, HALF_PANEL),
    snapped: false,
    valid: !state.panels.some(
      (panel) =>
        Math.abs(panel.x - snapToIncrement(rawX, HALF_PANEL)) < SNAP_DISTANCE_UNITS &&
        Math.abs(panel.y - snapToIncrement(rawY, HALF_PANEL)) < SNAP_DISTANCE_UNITS
    ),
    connectorMatches: 0,
    connectedIndices: new Set(),
  };

  candidates.forEach((candidate) => {
    const positioned = { ...preview, x: candidate.x, y: candidate.y };
    const previewConnections = getConnectionMapForPanels([...state.panels, positioned]);
    const connectorMatches = previewConnections.get("preview")?.anchorIndices.size || 0;
    const connectedIndices = previewConnections.get("preview")?.anchorIndices || new Set();
    const overlaps = state.panels.some(
      (panel) =>
        Math.abs(panel.x - candidate.x) < SNAP_DISTANCE_UNITS &&
        Math.abs(panel.y - candidate.y) < SNAP_DISTANCE_UNITS
    );
    const scored = {
      x: candidate.x,
      y: candidate.y,
      snapped: connectorMatches > 0,
      valid: !overlaps,
      connectorMatches,
      distanceFromPointer: candidate.distanceFromPointer,
      connectedIndices,
    };

    if (!best.snapped && scored.snapped) {
      best = scored;
      return;
    }

    if (best.snapped === scored.snapped) {
      const betterDistance = scored.distanceFromPointer < (best.distanceFromPointer ?? Infinity) - 0.1;
      const betterMatches = scored.connectorMatches > (best.connectorMatches || 0);
      if (betterDistance || (!betterDistance && betterMatches)) {
        best = scored;
      }
    }
  });

  return {
    panel: { ...preview, x: best.x, y: best.y },
    snapped: best.snapped,
    valid: best.valid,
    connectedIndices: best.connectedIndices || new Set(),
  };
}

// The grid's top-left cell reuses the normal single-panel snap logic (anchor
// matching against existing panels, falling back to the half-panel grid), so
// the whole grid inherits the same connector-snap and grid-snap behaviour for
// free. Every other cell is just an offset from that anchored corner.
function getGridPlacementPreview(cols, rows, pointerX, pointerY) {
  const origin = getPlacementPreview("MG9", pointerX, pointerY, 0);
  const cells = [];

  for (let row = 0; row < rows; row += 1) {
    for (let col = 0; col < cols; col += 1) {
      cells.push({
        x: origin.panel.x + col * PANEL_SIZE_UNITS,
        y: origin.panel.y + row * PANEL_SIZE_UNITS,
      });
    }
  }

  const overlaps = cells.some((cell) =>
    state.panels.some(
      (panel) => Math.abs(panel.x - cell.x) < SNAP_DISTANCE_UNITS && Math.abs(panel.y - cell.y) < SNAP_DISTANCE_UNITS
    )
  );

  return {
    cells,
    snapped: origin.snapped,
    valid: origin.valid && !overlaps,
  };
}

function renderGridPreview(preview) {
  const group = document.createElementNS(SVG_NS, "g");
  group.setAttribute(
    "class",
    `preview-group grid-preview-group${preview.snapped && preview.valid ? " snap-ready" : ""}${
      preview.valid ? "" : " invalid-preview"
    }`
  );
  preview.cells.forEach((cell) => {
    const rect = document.createElementNS(SVG_NS, "rect");
    rect.setAttribute("x", cell.x - HALF_PANEL);
    rect.setAttribute("y", cell.y - HALF_PANEL);
    rect.setAttribute("width", PANEL_SIZE_UNITS);
    rect.setAttribute("height", PANEL_SIZE_UNITS);
    rect.setAttribute("class", "panel-shape");
    rect.setAttribute("fill", PANEL_TYPES.MG9.color);
    group.appendChild(rect);
  });
  els.canvasPreview.appendChild(group);
}

function renderCanvasPreview() {
  els.canvasPreview.innerHTML = "";
  if (!state.placement.active || !state.placement.preview) return;

  if (state.placement.grid) {
    renderGridPreview(state.placement.preview);
    return;
  }

  const { panel, snapped, valid, connectedIndices } = state.placement.preview;
  const group = document.createElementNS(SVG_NS, "g");
  group.setAttribute("transform", `translate(${panel.x} ${panel.y})`);
  group.setAttribute(
    "class",
    `preview-group${snapped && valid ? " snap-ready" : ""}${valid ? "" : " invalid-preview"}`
  );
  group.appendChild(
    buildPanelChildren(panel, { anchorClass: "preview-anchor", connectedSet: connectedIndices })
  );
  els.canvasPreview.appendChild(group);
}

// Keep a stable id -> <g> map so full renders reuse existing nodes (and their
// listeners) instead of tearing down and rebuilding the whole panel layer.
const panelElements = new Map();

function onPanelClick(event) {
  event.stopPropagation();
  if (suppressClick) {
    suppressClick = false;
    return;
  }
  const id = event.currentTarget.dataset.id;
  setSelected(id, {
    additive: event.shiftKey || event.ctrlKey || event.metaKey,
    toggle: event.ctrlKey || event.metaKey,
  });
}

function renderPanelElement(panel) {
  let group = panelElements.get(panel.id);
  if (!group) {
    group = document.createElementNS(SVG_NS, "g");
    group.dataset.id = panel.id;
    group.addEventListener("pointerdown", onPanelPointerDown);
    group.addEventListener("click", onPanelClick);
    panelElements.set(panel.id, group);
  }
  group.setAttribute("class", panelClasses(panel));
  group.setAttribute("transform", `translate(${panel.x} ${panel.y})`);
  group.replaceChildren(buildPanelChildren(panel));
  els.canvasPanels.appendChild(group);
}

function syncPanelElements() {
  const alive = new Set(state.panels.map((panel) => panel.id));
  panelElements.forEach((element, id) => {
    if (!alive.has(id)) {
      element.remove();
      panelElements.delete(id);
    }
  });
  state.panels.forEach((panel) => renderPanelElement(panel));
}

function renderCanvas() {
  computeConnections();
  syncPanelElements();
  renderCanvasPreview();
  updateMetrics();
  renderGuides();
}

function getRotatedBoundingBox(panel) {
  const corners = [
    rotatePoint({ x: -HALF_PANEL, y: -HALF_PANEL }, panel.rotation),
    rotatePoint({ x: HALF_PANEL, y: -HALF_PANEL }, panel.rotation),
    rotatePoint({ x: HALF_PANEL, y: HALF_PANEL }, panel.rotation),
    rotatePoint({ x: -HALF_PANEL, y: HALF_PANEL }, panel.rotation),
  ];
  const xs = corners.map((point) => point.x + panel.x);
  const ys = corners.map((point) => point.y + panel.y);
  return {
    minX: Math.min(...xs),
    maxX: Math.max(...xs),
    minY: Math.min(...ys),
    maxY: Math.max(...ys),
  };
}

function getLayoutBounds() {
  if (state.panels.length === 0) {
    return {
      minX: 0,
      maxX: 2400,
      minY: 0,
      maxY: 1600,
      width: 2400,
      height: 1600,
      hasPanels: false,
    };
  }

  const merged = state.panels.map(getRotatedBoundingBox).reduce(
    (acc, bound) => ({
      minX: Math.min(acc.minX, bound.minX),
      maxX: Math.max(acc.maxX, bound.maxX),
      minY: Math.min(acc.minY, bound.minY),
      maxY: Math.max(acc.maxY, bound.maxY),
    }),
    { minX: Infinity, maxX: -Infinity, minY: Infinity, maxY: -Infinity }
  );

  return {
    ...merged,
    width: merged.maxX - merged.minX,
    height: merged.maxY - merged.minY,
    hasPanels: true,
  };
}

function parseViewBox(viewBoxStr) {
  const [x, y, width, height] = viewBoxStr.split(" ").map(Number);
  return { x, y, width, height };
}

function unionRect(a, b) {
  const minX = Math.min(a.x, b.x);
  const minY = Math.min(a.y, b.y);
  const maxX = Math.max(a.x + a.width, b.x + b.width);
  const maxY = Math.max(a.y + a.height, b.y + b.height);
  return { x: minX, y: minY, width: maxX - minX, height: maxY - minY };
}

function setCanvasExtent(x, y, width, height) {
  els.canvasBackground.setAttribute("x", x);
  els.canvasBackground.setAttribute("y", y);
  els.canvasBackground.setAttribute("width", width);
  els.canvasBackground.setAttribute("height", height);
  els.canvasGrid.setAttribute("x", x);
  els.canvasGrid.setAttribute("y", y);
  els.canvasGrid.setAttribute("width", width);
  els.canvasGrid.setAttribute("height", height);
  els.canvasSubgrid.setAttribute("x", x);
  els.canvasSubgrid.setAttribute("y", y);
  els.canvasSubgrid.setAttribute("width", width);
  els.canvasSubgrid.setAttribute("height", height);
}

function updateCanvasView(bounds) {
  if (state.drag) {
    els.canvas.setAttribute("viewBox", state.currentViewBox);
    return;
  }

  if (state.manualViewLocked && state.panels.length) {
    // Keep the user's current pan/zoom framing (don't recenter/re-fit), but
    // grow the canvas — never shrink it — so a panel placed outside the
    // current view is always pulled into the visible/exportable area.
    const current = parseViewBox(state.currentViewBox);
    const required = {
      x: bounds.minX - CANVAS_PADDING,
      y: bounds.minY - CANVAS_PADDING,
      width: bounds.width + CANVAS_PADDING * 2,
      height: bounds.height + CANVAS_PADDING * 2,
    };
    const expanded = unionRect(current, required);
    if (
      expanded.x !== current.x ||
      expanded.y !== current.y ||
      expanded.width !== current.width ||
      expanded.height !== current.height
    ) {
      state.currentViewBox = `${expanded.x} ${expanded.y} ${expanded.width} ${expanded.height}`;
      setCanvasExtent(expanded.x, expanded.y, expanded.width, expanded.height);
    }
    els.canvas.setAttribute("viewBox", state.currentViewBox);
    return;
  }

  const width = Math.max(bounds.width + CANVAS_PADDING * 2, 2400);
  const height = Math.max(bounds.height + CANVAS_PADDING * 2, 1600);
  const x = bounds.hasPanels ? bounds.minX - CANVAS_PADDING : 0;
  const y = bounds.hasPanels ? bounds.minY - CANVAS_PADDING : 0;

  state.currentViewBox = `${x} ${y} ${width} ${height}`;
  els.canvas.setAttribute("viewBox", state.currentViewBox);
  setCanvasExtent(x, y, width, height);
}

function renderGuides() {
  els.canvasGuides.innerHTML = "";
  const bounds = getLayoutBounds();
  updateCanvasView(bounds);

  if (!bounds.hasPanels) return;

  const widthMm = unitsToMm(bounds.width);
  const heightMm = unitsToMm(bounds.height);
  const topY = bounds.minY - 90;
  const leftX = bounds.minX - 90;

  const widthLine = document.createElementNS(SVG_NS, "line");
  widthLine.setAttribute("x1", bounds.minX);
  widthLine.setAttribute("y1", topY);
  widthLine.setAttribute("x2", bounds.maxX);
  widthLine.setAttribute("y2", topY);
  widthLine.setAttribute("class", "guide-arrow");
  els.canvasGuides.appendChild(widthLine);

  const widthText = document.createElementNS(SVG_NS, "text");
  widthText.setAttribute("x", bounds.minX + bounds.width / 2);
  widthText.setAttribute("y", topY - 18);
  widthText.setAttribute("class", "guide-label");
  widthText.setAttribute("text-anchor", "middle");
  widthText.textContent = mmToMetersText(widthMm);
  els.canvasGuides.appendChild(widthText);

  const heightLine = document.createElementNS(SVG_NS, "line");
  heightLine.setAttribute("x1", leftX);
  heightLine.setAttribute("y1", bounds.minY);
  heightLine.setAttribute("x2", leftX);
  heightLine.setAttribute("y2", bounds.maxY);
  heightLine.setAttribute("class", "guide-arrow");
  els.canvasGuides.appendChild(heightLine);

  const heightText = document.createElementNS(SVG_NS, "text");
  heightText.setAttribute("x", leftX - 24);
  heightText.setAttribute("y", bounds.minY + bounds.height / 2);
  heightText.setAttribute("class", "guide-label");
  heightText.setAttribute("text-anchor", "middle");
  heightText.setAttribute("transform", `rotate(-90 ${leftX - 24} ${bounds.minY + bounds.height / 2})`);
  heightText.textContent = mmToMetersText(heightMm);
  els.canvasGuides.appendChild(heightText);
}

function updateMetrics() {
  const bounds = getLayoutBounds();
  els.panelCount.textContent = String(state.panels.length);

  if (!bounds.hasPanels) {
    els.layoutWidth.textContent = "0.00 m";
    els.layoutHeight.textContent = "0.00 m";
    els.connectionStatus.textContent = "Ready";
    return;
  }

  els.layoutWidth.textContent = mmToMetersText(unitsToMm(bounds.width));
  els.layoutHeight.textContent = mmToMetersText(unitsToMm(bounds.height));

  const disconnectedCount = state.panels.filter((panel) => !isPanelConnected(panel.id)).length;
  els.connectionStatus.textContent =
    disconnectedCount > 0 ? `${disconnectedCount} disconnected` : "All connected";
}

function syncRotationInput(selectedPanels) {
  if (!els.rotationInput) return;
  if (!selectedPanels.length) {
    els.rotationInput.disabled = true;
    els.rotationInput.value = 0;
    return;
  }
  els.rotationInput.disabled = false;
  els.rotationInput.value = Math.round(selectedPanels[0].rotation) % 360;
}

function syncPushOutButton(selectedPanels) {
  if (!els.pushOutBtn) return;
  const selectedMG9 = selectedPanels.filter((panel) => panel.type === "MG9");
  if (!selectedMG9.length) {
    els.pushOutBtn.disabled = true;
    els.pushOutBtn.textContent = "Mark as Push-Out";
    return;
  }
  els.pushOutBtn.disabled = false;
  const allPushOut = selectedMG9.every((panel) => panel.pushOut);
  els.pushOutBtn.textContent = allPushOut ? "Revert to Standard MG9" : "Mark as Push-Out";
}

function renderSelectedMeta() {
  const selectedPanels = getSelectedPanels();
  const panel = getPanelById(state.selectedId);
  syncRotationInput(selectedPanels);
  syncPushOutButton(selectedPanels);
  if (!panel || !selectedPanels.length) {
    els.selectedMeta.classList.add("empty");
    const gridConfig = state.placement.grid;
    els.selectedMeta.textContent = state.placement.active
      ? gridConfig
        ? `Grid placement mode: ${gridConfig.cols} x ${gridConfig.rows} MG9 grid. Click the canvas to place it (placement mode ends automatically) or press Esc to cancel.`
        : `Placement mode: ${state.placement.type}. Click the canvas to place panels or press Esc to cancel.`
      : "Select a panel on the canvas.";
    return;
  }

  if (selectedPanels.length > 1) {
    const pushOutSelected = selectedPanels.filter((item) => item.type === "MG9" && item.pushOut).length;
    els.selectedMeta.classList.remove("empty");
    els.selectedMeta.innerHTML = `
      <strong>${selectedPanels.length} panels selected</strong><br />
      Move, rotate, duplicate, or delete them as one group.${
        pushOutSelected ? `<br />Push-Out panels in selection: ${pushOutSelected}` : ""
      }
    `;
    return;
  }

  const type = PANEL_TYPES[panel.type];
  const connectionState = isPanelConnected(panel.id) ? "Connected" : "Needs connection";
  const stockState = panel.available === false ? "Unavailable" : "Available";
  const isPushOut = panel.type === "MG9" && Boolean(panel.pushOut);
  els.selectedMeta.classList.remove("empty");
  els.selectedMeta.innerHTML = `
    <strong>${type.name}${isPushOut ? " — Push-Out" : ""}</strong><br />
    Size: ${type.widthMm} x ${type.heightMm} mm<br />
    Rotation: ${panel.rotation} deg<br />
    Position: ${Math.round(unitsToMm(panel.x))} mm, ${Math.round(unitsToMm(panel.y))} mm<br />
    Status: ${connectionState}<br />
    Stock: ${stockState}${panel.type === "MG9" ? `<br />Push-Out: ${isPushOut ? "Yes" : "No"}` : ""}
  `;
}

const CHANGELOG_URL = "https://github.com/underdog1234/MG9-Creative-Designer/releases";

function render() {
  els.appVersionBadge.textContent = `v${APP_VERSION}`;
  if (els.appVersionBadge instanceof HTMLAnchorElement) {
    els.appVersionBadge.href = `${CHANGELOG_URL}/tag/v${APP_VERSION}`;
  }
  els.projectNameInput.value = state.projectName;
  renderSections();
  renderInventory();
  renderLibrary();
  renderReplaceTypeLibrary();
  renderSelectedMeta();
  renderCanvas();
  updateZoom();
}

function updateZoom() {
  els.canvas.style.transformOrigin = "top left";
  els.canvas.style.transform = `scale(${state.zoom})`;
  els.zoomLabel.textContent = `${Math.round(state.zoom * 100)}%`;
}

function screenToSvg(event) {
  const point = els.canvas.createSVGPoint();
  point.x = event.clientX;
  point.y = event.clientY;
  const transformed = point.matrixTransform(els.canvas.getScreenCTM().inverse());
  return { x: transformed.x, y: transformed.y };
}

function onPanelPointerDown(event) {
  event.stopPropagation();
  suppressClick = false;
  const group = event.currentTarget;
  const panel = getPanelById(group.dataset.id);
  if (!panel) return;

  if (!state.selectedIds.includes(panel.id)) {
    setSelected(panel.id, {
      additive: event.shiftKey || event.ctrlKey || event.metaKey,
      toggle: event.ctrlKey || event.metaKey,
    });
  }
  pushHistory();
  state.manualViewLocked = true;
  const pointer = screenToSvg(event);
  const selectedPanels = getSelectedPanels();
  state.drag = {
    panelIds: selectedPanels.map((item) => item.id),
    originPointerX: pointer.x,
    originPointerY: pointer.y,
    originalPanels: selectedPanels.map((item) => ({
      id: item.id,
      x: item.x,
      y: item.y,
      rotation: item.rotation,
    })),
  };
}

function isBackgroundTarget(target) {
  return !(target instanceof Element) || !target.closest(".panel-group");
}

let dragFramePending = false;
// A drag that actually moved the pointer would otherwise be followed by a
// synthetic click that collapses a multi-selection to a single panel; this flag
// swallows exactly that one click.
let suppressClick = false;

function applyDrag() {
  dragFramePending = false;
  if (!state.drag || !state.drag.lastPointer) return;
  const dx = state.drag.lastPointer.x - state.drag.originPointerX;
  const dy = state.drag.lastPointer.y - state.drag.originPointerY;
  if (Math.hypot(dx, dy) > 2) state.drag.moved = true;
  state.drag.originalPanels.forEach((original) => {
    const panel = getPanelById(original.id);
    if (!panel) return;
    panel.x = original.x + dx;
    panel.y = original.y + dy;
    const element = panelElements.get(panel.id);
    if (element) element.setAttribute("transform", `translate(${panel.x} ${panel.y})`);
  });
  renderGuides();
  renderSelectedMeta();
}

// During a drag we only move existing DOM nodes (transform) and defer the
// expensive connection recompute until the pointer is released; movement is
// coalesced with requestAnimationFrame.
function handleDragMove(event) {
  state.drag.lastPointer = screenToSvg(event);
  if (dragFramePending) return;
  dragFramePending = true;
  requestAnimationFrame(applyDrag);
}

function finishDrag() {
  snapPanelGroup(state.drag.panelIds, PLACEMENT_LOCK_DISTANCE);
  computeConnections();
  if (state.drag.moved) suppressClick = true;
  state.drag = null;
  render();
}

function renderMarquee() {
  els.canvasMarquee.innerHTML = "";
  const marquee = state.marquee;
  if (!marquee || !marquee.moved) return;
  const rect = document.createElementNS(SVG_NS, "rect");
  rect.setAttribute("x", Math.min(marquee.startX, marquee.x));
  rect.setAttribute("y", Math.min(marquee.startY, marquee.y));
  rect.setAttribute("width", Math.abs(marquee.x - marquee.startX));
  rect.setAttribute("height", Math.abs(marquee.y - marquee.startY));
  rect.setAttribute("class", "marquee-rect");
  els.canvasMarquee.appendChild(rect);
}

function handleMarqueeMove(event) {
  const pointer = screenToSvg(event);
  const marquee = state.marquee;
  marquee.x = pointer.x;
  marquee.y = pointer.y;
  if (!marquee.moved && Math.hypot(marquee.x - marquee.startX, marquee.y - marquee.startY) > 4) {
    marquee.moved = true;
  }
  renderMarquee();
}

function finalizeMarquee() {
  const marquee = state.marquee;
  state.marquee = null;
  els.canvasMarquee.innerHTML = "";
  if (!marquee.moved) {
    if (!marquee.additive) clearSelection();
    return;
  }
  const minX = Math.min(marquee.startX, marquee.x);
  const maxX = Math.max(marquee.startX, marquee.x);
  const minY = Math.min(marquee.startY, marquee.y);
  const maxY = Math.max(marquee.startY, marquee.y);
  const selected = new Set(marquee.additive ? state.selectedIds : []);
  state.panels.forEach((panel) => {
    const box = getRotatedBoundingBox(panel);
    if (box.minX <= maxX && box.maxX >= minX && box.minY <= maxY && box.maxY >= minY) {
      selected.add(panel.id);
    }
  });
  state.selectedIds = [...selected];
  state.selectedId = state.selectedIds[0] || null;
  renderSelectedMeta();
  renderCanvas();
}

function onCanvasPointerDown(event) {
  suppressClick = false;
  if (event.button !== 0) return;
  if (state.placement.active) return;
  if (!isBackgroundTarget(event.target)) return;
  const pointer = screenToSvg(event);
  state.marquee = {
    startX: pointer.x,
    startY: pointer.y,
    x: pointer.x,
    y: pointer.y,
    additive: event.shiftKey || event.ctrlKey || event.metaKey,
    moved: false,
  };
}

function onWindowPointerMove(event) {
  if (state.marquee) {
    handleMarqueeMove(event);
    return;
  }
  if (state.drag) {
    handleDragMove(event);
  }
}

function onWindowPointerUp() {
  if (state.marquee) {
    finalizeMarquee();
    return;
  }
  if (state.drag) {
    finishDrag();
  }
}

function onCanvasHover(event) {
  state.lastPointer = screenToSvg(event);
  if (!state.placement.active || state.drag) return;
  state.placement.pointer = state.lastPointer;
  state.placement.preview = state.placement.grid
    ? getGridPlacementPreview(
        state.placement.grid.cols,
        state.placement.grid.rows,
        state.lastPointer.x,
        state.lastPointer.y
      )
    : getPlacementPreview(state.placement.type, state.lastPointer.x, state.lastPointer.y);
  renderCanvasPreview();
}

function onCanvasLeave() {
  if (!state.placement.active || state.drag) return;
  state.placement.pointer = null;
  state.placement.preview = null;
  renderCanvasPreview();
}

function onCanvasClick(event) {
  if (!state.placement.active) return;
  if (!isBackgroundTarget(event.target)) return;
  const pointer = screenToSvg(event);

  if (state.placement.grid) {
    const { cols, rows } = state.placement.grid;
    const preview = getGridPlacementPreview(cols, rows, pointer.x, pointer.y);
    if (!preview.valid) return;
    pushHistory();
    state.manualViewLocked = true;
    const newPanels = preview.cells.map((cell) => createPanel("MG9", cell.x, cell.y, 0));
    state.panels.push(...newPanels);
    state.selectedId = newPanels[newPanels.length - 1]?.id || null;
    state.selectedIds = newPanels.map((item) => item.id);
    applyAvailability();
    computeConnections();
    // Bulk placement is a one-shot action: place the whole grid, then drop
    // out of placement mode automatically (unlike single-panel placement,
    // which stays armed for placing several in a row).
    resetPlacement();
    render();
    return;
  }

  const preview = getPlacementPreview(state.placement.type, pointer.x, pointer.y);
  if (!preview.valid) return;
  pushHistory();
  state.manualViewLocked = true;
  const panel = createPanel(preview.panel.type, preview.panel.x, preview.panel.y, preview.panel.rotation);
  state.panels.push(panel);
  state.selectedId = panel.id;
  state.selectedIds = [panel.id];
  applyAvailability();
  computeConnections();
  state.placement.pointer = pointer;
  state.placement.preview = getPlacementPreview(state.placement.type, pointer.x, pointer.y);
  render();
}

function snapPanel(panel, threshold = SNAP_DISTANCE_UNITS) {
  const geometry = getGeometry(panel);
  let bestMatch = null;

  state.panels.forEach((other) => {
    if (other.id === panel.id) return;
    const otherGeometry = getGeometry(other);

    geometry.anchors.forEach((anchor) => {
      otherGeometry.anchors.forEach((otherAnchor) => {
        const currentAnchor = { x: panel.x + anchor.x, y: panel.y + anchor.y };
        const targetAnchor = { x: other.x + otherAnchor.x, y: other.y + otherAnchor.y };
        const dx = targetAnchor.x - currentAnchor.x;
        const dy = targetAnchor.y - currentAnchor.y;
        const distance = Math.hypot(dx, dy);

        if (distance <= threshold && (!bestMatch || distance < bestMatch.distance)) {
          bestMatch = { dx, dy, distance };
        }
      });
    });
  });

  if (bestMatch) {
    panel.x += bestMatch.dx;
    panel.y += bestMatch.dy;
  } else {
    panel.x = snapToIncrement(panel.x, HALF_PANEL);
    panel.y = snapToIncrement(panel.y, HALF_PANEL);
  }
}

// Snaps a whole selection as one rigid body: finds the single best anchor
// match between any panel in the group and any panel outside it, then moves
// every panel in the group by that same offset. Snapping each panel
// independently (the old behaviour) let different panels in a dragged group
// lock onto different neighbours, which silently pulled the group out of
// alignment with itself.
function snapPanelGroup(panelIds, threshold = SNAP_DISTANCE_UNITS) {
  const idSet = new Set(panelIds);
  const groupPanels = panelIds.map((id) => getPanelById(id)).filter(Boolean);
  if (!groupPanels.length) return;

  if (groupPanels.length === 1) {
    snapPanel(groupPanels[0], threshold);
    return;
  }

  const externalPanels = state.panels.filter((panel) => !idSet.has(panel.id));
  let bestMatch = null;

  groupPanels.forEach((panel) => {
    const geometry = getGeometry(panel);
    externalPanels.forEach((other) => {
      const otherGeometry = getGeometry(other);
      geometry.anchors.forEach((anchor) => {
        otherGeometry.anchors.forEach((otherAnchor) => {
          const currentAnchor = { x: panel.x + anchor.x, y: panel.y + anchor.y };
          const targetAnchor = { x: other.x + otherAnchor.x, y: other.y + otherAnchor.y };
          const dx = targetAnchor.x - currentAnchor.x;
          const dy = targetAnchor.y - currentAnchor.y;
          const distance = Math.hypot(dx, dy);

          if (distance <= threshold && (!bestMatch || distance < bestMatch.distance)) {
            bestMatch = { dx, dy, distance };
          }
        });
      });
    });
  });

  if (bestMatch) {
    groupPanels.forEach((panel) => {
      panel.x += bestMatch.dx;
      panel.y += bestMatch.dy;
    });
    return;
  }

  // No neighbour close enough: settle the group onto the half-panel grid as a
  // unit, using one panel as the alignment reference so relative spacing
  // inside the group is preserved exactly.
  const reference = groupPanels[0];
  const dx = snapToIncrement(reference.x, HALF_PANEL) - reference.x;
  const dy = snapToIncrement(reference.y, HALF_PANEL) - reference.y;
  if (dx || dy) {
    groupPanels.forEach((panel) => {
      panel.x += dx;
      panel.y += dy;
    });
  }
}

// "Connected" (the green anchor / All-connected status) is decided by anchors
// being within SNAP_DISTANCE_UNITS of each other, but proximity alone doesn't
// move anything -- a panel can show as connected while still sitting a few
// units away from its neighbour, i.e. joined-but-with-a-gap. This pass closes
// that gap for the whole layout: any panel whose anchor is within tolerance
// of another panel's anchor (but not already exactly on it) gets nudged by
// the exact vector needed to make them flush, so "connected" and "flush"
// become the same thing everywhere, not just right after a drag.
//
// The spatial hash is built incrementally as panels are processed (not
// snapshotted up front), so each panel only settles against neighbours'
// already-updated positions. That matters: snapshotting positions for the
// whole pass and moving everyone from that stale snapshot lets two mutually
// touching panels each try to close the *same* gap independently, doubling
// the correction and overshooting into an overlap instead of a flush join.
// Settling one side at a time against live positions closes each gap exactly
// once. A few passes let chains of near-touching panels fully settle.
function settleAllPanelsFlush(threshold = SNAP_DISTANCE_UNITS) {
  const MAX_PASSES = 4;
  let anyMoved = false;

  for (let pass = 0; pass < MAX_PASSES; pass += 1) {
    const buckets = new Map();
    const insertAnchor = (anchor) => {
      const key = `${Math.floor(anchor.x / CONNECTION_CELL)}:${Math.floor(anchor.y / CONNECTION_CELL)}`;
      let bucket = buckets.get(key);
      if (!bucket) {
        bucket = [];
        buckets.set(key, bucket);
      }
      bucket.push(anchor);
    };

    let movedThisPass = false;

    state.panels.forEach((panel) => {
      let bestMatch = null;

      getGlobalAnchors(panel).forEach((anchor) => {
        const cellX = Math.floor(anchor.x / CONNECTION_CELL);
        const cellY = Math.floor(anchor.y / CONNECTION_CELL);
        for (let gx = cellX - 1; gx <= cellX + 1; gx += 1) {
          for (let gy = cellY - 1; gy <= cellY + 1; gy += 1) {
            const bucket = buckets.get(`${gx}:${gy}`);
            if (!bucket) continue;
            for (const other of bucket) {
              if (other.panelId === panel.id) continue;
              const dx = other.x - anchor.x;
              const dy = other.y - anchor.y;
              const distance = Math.hypot(dx, dy);
              if (distance > 0 && distance <= threshold && (!bestMatch || distance < bestMatch.distance)) {
                bestMatch = { dx, dy, distance };
              }
            }
          }
        }
      });

      if (bestMatch) {
        panel.x += bestMatch.dx;
        panel.y += bestMatch.dy;
        movedThisPass = true;
        anyMoved = true;
      }

      // Insert this panel's final-for-now anchors so panels processed later
      // in this same pass settle against where it actually ended up.
      getGlobalAnchors(panel).forEach(insertAnchor);
    });

    if (!movedThisPass) break;
  }

  return anyMoved;
}

function isLayoutValid() {
  if (state.panels.length <= 1) return true;
  return state.panels.every((panel) => isPanelConnected(panel.id));
}

function findPlacementForNewPanel(type, rotation = 0) {
  if (state.panels.length === 0) {
    return createPanel(type, 500, 500, rotation);
  }

  const basePanel = getPanelById(state.selectedId) || state.panels[0];
  const candidateOffsets = [
    { x: PANEL_SIZE_UNITS, y: 0 },
    { x: 0, y: PANEL_SIZE_UNITS },
    { x: -PANEL_SIZE_UNITS, y: 0 },
    { x: 0, y: -PANEL_SIZE_UNITS },
  ];

  for (const offset of candidateOffsets) {
    const candidate = createPanel(type, basePanel.x + offset.x, basePanel.y + offset.y, rotation);
    snapPanel(candidate);
    const overlaps = state.panels.some(
      (panel) => Math.hypot(panel.x - candidate.x, panel.y - candidate.y) < SNAP_DISTANCE_UNITS
    );
    if (!overlaps) {
      return candidate;
    }
  }

  return createPanel(type, basePanel.x + PANEL_SIZE_UNITS, basePanel.y, rotation);
}

function clearLayout() {
  pushHistory();
  state.manualViewLocked = false;
  resetPlacement();
  state.panels = [];
  state.selectedId = null;
  state.selectedIds = [];
  applyAvailability();
  computeConnections();
  render();
}

function resetInventoryUsed() {
  pushHistory();
  state.manualViewLocked = false;
  resetPlacement();
  state.panels = [];
  state.selectedId = null;
  state.selectedIds = [];
  state.inventory = cloneInventory(defaultInventory);
  applyAvailability();
  computeConnections();
  render();
}

function normalizePatternRows(pattern) {
  const width = Math.max(...pattern.map((row) => row.length));
  return pattern.map((row) => row.padEnd(width, "0"));
}

function trimPattern(pattern) {
  let rows = [...pattern];

  while (rows.length > 1 && /^0+$/.test(rows[0])) rows.shift();
  while (rows.length > 1 && /^0+$/.test(rows[rows.length - 1])) rows.pop();

  const width = rows[0].length;
  let left = 0;
  let right = width - 1;

  while (left < width - 1 && rows.every((row) => row[left] === "0")) left += 1;
  while (right > 0 && rows.every((row) => row[right] === "0")) right -= 1;

  rows = rows.map((row) => row.slice(left, right + 1));
  return normalizePatternRows(rows);
}

function rasterizeGlyph(glyph) {
  const canvas = document.createElement("canvas");
  const size = 112;
  const cols = 7;
  const rows = 7;
  canvas.width = size;
  canvas.height = size;
  const context = canvas.getContext("2d");

  context.clearRect(0, 0, size, size);
  context.fillStyle = "#000";
  context.textAlign = "center";
  context.textBaseline = "middle";
  context.font = `700 78px "Segoe UI Emoji", "Apple Color Emoji", "Noto Color Emoji", "Segoe UI Symbol", sans-serif`;
  context.fillText(glyph, size / 2, size / 2 + 2);

  const { data } = context.getImageData(0, 0, size, size);
  const rowStrings = [];

  for (let row = 0; row < rows; row += 1) {
    let rowString = "";
    for (let col = 0; col < cols; col += 1) {
      let active = 0;
      const startX = Math.floor((col / cols) * size);
      const endX = Math.floor(((col + 1) / cols) * size);
      const startY = Math.floor((row / rows) * size);
      const endY = Math.floor(((row + 1) / rows) * size);

      for (let y = startY; y < endY; y += 1) {
        for (let x = startX; x < endX; x += 1) {
          const index = (y * size + x) * 4;
          if (data[index + 3] > 32) active += 1;
        }
      }

      rowString += active > ((endX - startX) * (endY - startY)) / 6 ? "1" : "0";
    }
    rowStrings.push(rowString);
  }

  if (rowStrings.every((row) => /^0+$/.test(row))) {
    return normalizePatternRows(["11111", "10001", "10001", "10001", "10001", "10001", "11111"]);
  }

  return trimPattern(rowStrings);
}

function getPatternForGlyph(glyph) {
  const upper = glyph.toUpperCase();
  if (LETTER_PATTERNS[upper]) return normalizePatternRows(LETTER_PATTERNS[upper]);
  return rasterizeGlyph(glyph);
}

function getTextBlueprint(graphemes) {
  return graphemes.map((glyph) => ({
    glyph,
    pattern: getPatternForGlyph(glyph),
  }));
}

function buildSourceBitmap(blueprint, spacingPanels) {
  const rows = Math.max(...blueprint.map((item) => item.pattern.length), 1);
  const cols = blueprint.reduce((total, item, index) => {
    const gap = index === blueprint.length - 1 ? 0 : spacingPanels;
    return total + item.pattern[0].length + gap;
  }, 0);

  const bitmap = Array.from({ length: rows }, () => Array(cols).fill(0));
  let cursor = 0;

  blueprint.forEach((item, index) => {
    item.pattern.forEach((row, rowIndex) => {
      [...row].forEach((cell, colIndex) => {
        if (cell === "1") {
          bitmap[rowIndex][cursor + colIndex] = 1;
        }
      });
    });

    cursor += item.pattern[0].length;
    if (index < blueprint.length - 1) cursor += spacingPanels;
  });

  return bitmap;
}

function resampleBitmap(sourceBitmap, targetCols, targetRows) {
  const sourceRows = sourceBitmap.length;
  const sourceCols = sourceBitmap[0]?.length || 1;
  const scaled = Array.from({ length: targetRows }, () => Array(targetCols).fill(0));

  for (let targetRow = 0; targetRow < targetRows; targetRow += 1) {
    const sourceY0 = (targetRow * sourceRows) / targetRows;
    const sourceY1 = ((targetRow + 1) * sourceRows) / targetRows;

    for (let targetCol = 0; targetCol < targetCols; targetCol += 1) {
      const sourceX0 = (targetCol * sourceCols) / targetCols;
      const sourceX1 = ((targetCol + 1) * sourceCols) / targetCols;
      let covered = 0;
      let totalArea = 0;

      for (let sourceRow = Math.floor(sourceY0); sourceRow < Math.ceil(sourceY1); sourceRow += 1) {
        for (let sourceCol = Math.floor(sourceX0); sourceCol < Math.ceil(sourceX1); sourceCol += 1) {
          const overlapX = Math.max(
            0,
            Math.min(sourceX1, sourceCol + 1) - Math.max(sourceX0, sourceCol)
          );
          const overlapY = Math.max(
            0,
            Math.min(sourceY1, sourceRow + 1) - Math.max(sourceY0, sourceRow)
          );
          const area = overlapX * overlapY;

          if (area <= 0) continue;
          totalArea += area;
          if (sourceBitmap[sourceRow]?.[sourceCol]) covered += area;
        }
      }

      scaled[targetRow][targetCol] = covered >= totalArea * 0.35 ? 1 : 0;
    }
  }

  return scaled;
}

function rotateTemplate(mask, rotation) {
  const size = mask.length;
  let working = mask.map((row) => [...row]);
  const turns = ((rotation % 360) + 360) % 360 / 90;

  for (let turn = 0; turn < turns; turn += 1) {
    const rotated = Array.from({ length: size }, () => Array(size).fill(0));
    for (let y = 0; y < size; y += 1) {
      for (let x = 0; x < size; x += 1) {
        rotated[x][size - 1 - y] = working[y][x];
      }
    }
    working = rotated;
  }

  return working;
}

function createBaseTemplate(type) {
  const size = TEMPLATE_RESOLUTION;
  const mask = Array.from({ length: size }, () => Array(size).fill(0));

  for (let y = 0; y < size; y += 1) {
    for (let x = 0; x < size; x += 1) {
      const nx = x / (size - 1);
      const ny = y / (size - 1);

      if (type === "MG9") {
        mask[y][x] = 1;
      } else if (type === "MG12") {
        mask[y][x] = ny >= nx ? 1 : 0;
      } else if (type === "MG13") {
        const dx = 1 - nx;
        const dy = 1 - ny;
        mask[y][x] = dx * dx + dy * dy <= 1 ? 1 : 0;
      }
    }
  }

  return mask;
}

const PIECE_TEMPLATES = [
  { type: "MG9", rotation: 0, mask: createBaseTemplate("MG9") },
  ...[0, 90, 180, 270].map((rotation) => ({
    type: "MG12",
    rotation,
    mask: rotateTemplate(createBaseTemplate("MG12"), rotation),
  })),
  ...[0, 90, 180, 270].map((rotation) => ({
    type: "MG13",
    rotation,
    mask: rotateTemplate(createBaseTemplate("MG13"), rotation),
  })),
];

function sampleCellMask(sourceBitmap, targetCols, targetRows, cellCol, cellRow) {
  const sourceRows = sourceBitmap.length;
  const sourceCols = sourceBitmap[0]?.length || 1;
  const mask = Array.from({ length: TEMPLATE_RESOLUTION }, () => Array(TEMPLATE_RESOLUTION).fill(0));

  for (let y = 0; y < TEMPLATE_RESOLUTION; y += 1) {
    for (let x = 0; x < TEMPLATE_RESOLUTION; x += 1) {
      const globalX = cellCol + (x + 0.5) / TEMPLATE_RESOLUTION;
      const globalY = cellRow + (y + 0.5) / TEMPLATE_RESOLUTION;
      const sourceX = Math.min(sourceCols - 1, Math.max(0, Math.floor((globalX / targetCols) * sourceCols)));
      const sourceY = Math.min(sourceRows - 1, Math.max(0, Math.floor((globalY / targetRows) * sourceRows)));
      mask[y][x] = sourceBitmap[sourceY]?.[sourceX] ? 1 : 0;
    }
  }

  return mask;
}

function scoreTemplate(mask, template) {
  let score = 0;
  let filled = 0;

  for (let y = 0; y < TEMPLATE_RESOLUTION; y += 1) {
    for (let x = 0; x < TEMPLATE_RESOLUTION; x += 1) {
      const expected = mask[y][x];
      const actual = template.mask[y][x];
      if (expected) filled += 1;
      if (expected && actual) score += 1.2;
      else if (!expected && !actual) score += 0.15;
      else if (!expected && actual) score -= 0.9;
      else score -= 0.55;
    }
  }

  return { score, fillRatio: filled / (TEMPLATE_RESOLUTION * TEMPLATE_RESOLUTION) };
}

function chooseCreativeType(cell, cellSet, style) {
  if (style === "rect") {
    return { type: "MG9", rotation: 0 };
  }

  const up = cellSet.has(`${cell.col},${cell.row - 1}`);
  const down = cellSet.has(`${cell.col},${cell.row + 1}`);
  const left = cellSet.has(`${cell.col - 1},${cell.row}`);
  const right = cellSet.has(`${cell.col + 1},${cell.row}`);
  const upLeft = cellSet.has(`${cell.col - 1},${cell.row - 1}`);
  const upRight = cellSet.has(`${cell.col + 1},${cell.row - 1}`);
  const downLeft = cellSet.has(`${cell.col - 1},${cell.row + 1}`);
  const downRight = cellSet.has(`${cell.col + 1},${cell.row + 1}`);

  if (!up && !left && right && down && downRight) return { type: "MG13", rotation: 0 };
  if (!up && !right && left && down && downLeft) return { type: "MG13", rotation: 90 };
  if (!down && !right && left && up && upLeft) return { type: "MG13", rotation: 180 };
  if (!down && !left && right && up && upRight) return { type: "MG13", rotation: 270 };

  if ((right && downRight && !up && !left) || (down && downRight && !up && !left)) {
    return { type: "MG12", rotation: 0 };
  }
  if ((left && downLeft && !up && !right) || (down && downLeft && !up && !right)) {
    return { type: "MG12", rotation: 90 };
  }
  if ((left && upLeft && !down && !right) || (up && upLeft && !down && !right)) {
    return { type: "MG12", rotation: 180 };
  }
  if ((right && upRight && !down && !left) || (up && upRight && !down && !left)) {
    return { type: "MG12", rotation: 270 };
  }

  if ([up, down, left, right].filter(Boolean).length <= 1) {
    return { type: "MG9", rotation: 0 };
  }

  return { type: "MG9", rotation: 0 };
}

function normalizeGlyphEntry(entry) {
  if (Array.isArray(entry)) {
    const width = Math.max(...entry.map((row) => row.length));
    const panels = [];

    entry.forEach((row, rowIndex) => {
      [...row].forEach((token, colIndex) => {
        if (token !== ".") panels.push({ token, x: colIndex, y: rowIndex });
      });
    });

    return { width, height: entry.length, panels };
  }

  return entry;
}

function buildLibraryGlyphPanels(lines, targetHeightPanels, spacingPanels, style) {
  const scale = Math.max(1, Math.round(targetHeightPanels / 5));
  const lineGapPanels = Math.max(1, Math.round(scale * 1.5));
  let cursorRow = 0;
  let maxWidth = 0;
  const panels = [];

  lines.forEach((graphemes, lineIndex) => {
    let cursorCol = 0;

    graphemes.forEach((glyph, glyphIndex) => {
      const glyphEntry = normalizeGlyphEntry(GLYPH_LIBRARY[glyph.toUpperCase()]);

      glyphEntry.panels.forEach((cell) => {
        const mapped = style === "rect" ? GLYPH_TOKEN_MAP.S : GLYPH_TOKEN_MAP[cell.token] || GLYPH_TOKEN_MAP.S;

        for (let scaleY = 0; scaleY < scale; scaleY += 1) {
          for (let scaleX = 0; scaleX < scale; scaleX += 1) {
            panels.push({
              type: mapped.type,
              rotation: mapped.rotation,
              x: 500 + (cursorCol + cell.x * scale + scaleX) * PANEL_SIZE_UNITS,
              y: 500 + (cursorRow + cell.y * scale + scaleY) * PANEL_SIZE_UNITS,
            });
          }
        }
      });

      cursorCol += glyphEntry.width * scale;
      if (glyphIndex < graphemes.length - 1) cursorCol += spacingPanels;
    });

    maxWidth = Math.max(maxWidth, cursorCol || 1);
    cursorRow += 5 * scale;
    if (lineIndex < lines.length - 1) cursorRow += lineGapPanels;
  });

  return {
    generatedPanels: panels,
    targetWidthPanels: maxWidth || 1,
    targetHeightPanels: cursorRow || 5 * scale,
    fromLibrary: true,
  };
}

function buildTextPanels() {
  const lines = (els.textInput.value || "")
    .replace(/\r/g, "")
    .split("\n")
    .map((line) => {
      const graphemes = getGraphemes(line);
      return graphemes.length ? graphemes : [" "];
    });
  const visibleLines = lines.length ? lines : [[" "]];
  const visibleGraphemes = visibleLines.flat();
  const spacingPanels = Math.max(0, Math.round((Number(els.spacingInput.value) || 0) * 1000 / PANEL_SIZE_MM));
  const targetHeightPanels = Math.max(1, Math.round(((Number(els.heightMetersInput.value) || 1) * 1000) / PANEL_SIZE_MM));
  const style = els.letterStyleInput.value;
  const allSupported = visibleGraphemes.every((glyph) => GLYPH_LIBRARY[glyph.toUpperCase()]);

  if (allSupported) {
    return buildLibraryGlyphPanels(visibleLines, targetHeightPanels, spacingPanels, style);
  }

  const blueprint = getTextBlueprint(visibleGraphemes);
  const sourceBitmap = buildSourceBitmap(blueprint, spacingPanels);
  const sourceRows = sourceBitmap.length || 1;
  const sourceCols = sourceBitmap[0]?.length || 1;
  const targetWidthPanels = Math.max(1, Math.round((sourceCols / sourceRows) * targetHeightPanels));
  const scaledBitmap = resampleBitmap(sourceBitmap, targetWidthPanels, targetHeightPanels);
  const generatedPanels = [];
  const activeCells = [];

  scaledBitmap.forEach((row, rowIndex) => {
    row.forEach((value, colIndex) => {
      if (value) activeCells.push({ col: colIndex, row: rowIndex });
    });
  });

  const cellSet = new Set(activeCells.map((cell) => `${cell.col},${cell.row}`));

  activeCells.forEach((cell) => {
    const piece = chooseCreativeType(cell, cellSet, style);
    generatedPanels.push({
      type: piece.type,
      rotation: piece.rotation,
      x: 500 + cell.col * PANEL_SIZE_UNITS,
      y: 500 + cell.row * PANEL_SIZE_UNITS,
    });
  });

  return {
    generatedPanels,
    targetWidthPanels,
    targetHeightPanels,
    fromLibrary: false,
  };
}

function checkStockForGeneratedPanels(generatedPanels) {
  const requestedCounts = {};
  const requestedOrientation = { MG12: {}, MG13: {} };

  generatedPanels.forEach((panel) => {
    requestedCounts[panel.type] = (requestedCounts[panel.type] || 0) + 1;
    if (isShapedType(panel.type)) {
      const orientation = getPanelOrientation(panel);
      const counts = requestedOrientation[panel.type];
      counts[orientation] = (counts[orientation] || 0) + 1;
    }
  });

  const shortages = [];
  if ((requestedCounts.MG9 || 0) > getStockTotal("MG9")) {
    shortages.push({ label: "MG9", need: requestedCounts.MG9, have: getStockTotal("MG9") });
  }
  SHAPED_TYPES.forEach((type) => {
    ORIENTATIONS.forEach((orientation) => {
      const need = requestedOrientation[type][orientation.key] || 0;
      const have = getStock(type, orientation.key);
      if (need > have) {
        shortages.push({ label: `${type} ${orientation.icon}`, need, have });
      }
    });
  });

  return { requestedCounts, requestedOrientation, shortages };
}

function optimizeGeneratedPanels(generatedPanels) {
  const workingPanels = generatedPanels.map((panel, index) => ({ ...panel, id: `gen-${index}` }));
  let changed = true;
  while (changed) {
    changed = false;
    const connectionMap = getConnectionMapForPanels(workingPanels);

    workingPanels.forEach((panel, index) => {
      const connectorCount = connectionMap.get(panel.id)?.anchorIndices.size || 0;
      if (panel.type !== "MG9" && connectorCount < 2) {
        const options = [
          { ...panel, rotation: 0 },
          { ...panel, rotation: 90 },
          { ...panel, rotation: 180 },
          { ...panel, rotation: 270 },
          { ...panel, type: "MG9", rotation: 0 },
        ];
        let bestOption = { ...panel, type: "MG9", rotation: 0 };
        let bestCount = 0;

        options.forEach((option) => {
          const candidatePanels = workingPanels.map((candidate, candidateIndex) =>
            candidateIndex === index ? { ...option } : { ...candidate }
          );
          const candidateMap = getConnectionMapForPanels(candidatePanels);
          const count = candidateMap.get(candidatePanels[index].id)?.anchorIndices.size || 0;
          if (count > bestCount) {
            bestCount = count;
            bestOption = { ...option };
          }
        });

        workingPanels[index] = bestOption;
        changed = true;
      }
    });
  }

  return workingPanels;
}

function updateTextSummary(generatedPanels, targetWidthPanels, targetHeightPanels, stockCheck) {
  if (generatedPanels.length === 0) {
    els.textSummary.textContent = "No visible panels for this text.";
    els.actualWidthOutput.value = "0.0";
    return;
  }

  const xs = generatedPanels.map((panel) => panel.x);
  const ys = generatedPanels.map((panel) => panel.y);
  const widthMm = (Math.max(...xs) - Math.min(...xs)) / MM_TO_UNITS + PANEL_SIZE_MM;
  const heightMm = (Math.max(...ys) - Math.min(...ys)) / MM_TO_UNITS + PANEL_SIZE_MM;
  const unavailableCount = stockCheck.shortages.reduce(
    (total, shortage) => total + Math.max(0, shortage.need - shortage.have),
    0
  );
  const stockText = unavailableCount > 0 ? ` ${unavailableCount} panels shown as unavailable.` : "";
  els.actualWidthOutput.value = (widthMm / 1000).toFixed(2);
  els.textSummary.textContent = `Height ${mmToMetersText(targetHeightPanels * PANEL_SIZE_MM)}. Actual ${mmToMetersText(widthMm)} x ${mmToMetersText(heightMm)}.${stockText}`;
}

function generateTextLayout(showAlerts = false, recordHistory = true) {
  if (recordHistory) pushHistory();
  state.manualViewLocked = false;
  resetPlacement();
  const { generatedPanels, targetWidthPanels, targetHeightPanels, fromLibrary } = buildTextPanels();
  const optimizedPanels = fromLibrary ? generatedPanels : optimizeGeneratedPanels(generatedPanels);
  const stockCheck = checkStockForGeneratedPanels(optimizedPanels);
  updateTextSummary(optimizedPanels, targetWidthPanels, targetHeightPanels, stockCheck);

  if (showAlerts && stockCheck.shortages.length) {
    const shortageText = stockCheck.shortages
      .map((shortage) => `${shortage.label}: need ${shortage.need}, have ${shortage.have}`)
      .join("\n");
    window.alert(`Layout generated with unavailable panels shown in grey.\n${shortageText}`);
  }

  state.panels = optimizedPanels.map((panel) => createPanel(panel.type, panel.x, panel.y, panel.rotation));
  applyAvailability();
  state.selectedId = state.panels[0]?.id || null;
  state.selectedIds = state.selectedId ? [state.selectedId] : [];
  computeConnections();
  render();
}

function scheduleAutoGenerate() {
  clearTimeout(state.autoTextTimer);
  state.autoTextTimer = setTimeout(() => generateTextLayout(false, false), AUTO_REFRESH_DELAY_MS);
}

function clearAndResetInventory() {
  state.inventory = cloneInventory(defaultInventory);
  clearLayout();
}

function serializeProject() {
  return {
    version: APP_VERSION,
    projectName: sanitizeProjectName(state.projectName),
    inventory: cloneInventory(state.inventory),
    panels: state.panels.map((panel) => ({
      type: panel.type,
      x: panel.x,
      y: panel.y,
      rotation: panel.rotation,
      pushOut: Boolean(panel.pushOut),
    })),
    textLayout: {
      text: els.textInput.value,
      heightMeters: Number(els.heightMetersInput.value) || 2.5,
      spacing: Number(els.spacingInput.value) || 0.5,
      style: els.letterStyleInput.value,
    },
    ui: {
      collapsedSections: { ...state.collapsedSections },
    },
  };
}

function downloadFile(filename, content, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function saveProject() {
  state.projectName = sanitizeProjectName(els.projectNameInput.value);
  els.projectNameInput.value = state.projectName;
  const project = serializeProject();
  downloadFile(`${project.projectName}.json`, JSON.stringify(project, null, 2), "application/json");
}

function restoreProject(project) {
  if (!project || typeof project !== "object") throw new Error("Invalid project file.");
  if (!Array.isArray(project.panels)) throw new Error("Project file is missing panels.");

  state.projectName = sanitizeProjectName(project.projectName);
  state.inventory = normalizeInventory({ ...DEFAULT_INVENTORY_RAW, ...(project.inventory || {}) });
  state.collapsedSections = {
    textLayout: false,
    selectedPanel: false,
    panelLibrary: true,
    inventory: true,
    help: true,
    ...(project.ui?.collapsedSections || {}),
  };
  state.panels = [];
  state.nextId = 1;

  (project.panels || []).forEach((panel) => {
    state.panels.push(
      createPanel(
        panel.type,
        Number(panel.x) || 500,
        Number(panel.y) || 500,
        Number(panel.rotation) || 0,
        Boolean(panel.pushOut)
      )
    );
  });

  els.projectNameInput.value = state.projectName;
  els.textInput.value = project.textLayout?.text ?? els.textInput.value;
  els.heightMetersInput.value = String(project.textLayout?.heightMeters ?? 2.5);
  els.spacingInput.value = String(project.textLayout?.spacing ?? 0.5);
  els.letterStyleInput.value = project.textLayout?.style ?? "mixed";

  state.selectedId = null;
  state.selectedIds = [];
  state.manualViewLocked = false;
  resetPlacement();
  // Loaded positions come straight from the file with no snapping applied,
  // so near-but-not-exact joins (rounding, hand-edited coordinates, etc.)
  // would otherwise show as "connected" while still having a visible gap.
  settleAllPanelsFlush();
  applyAvailability();
  computeConnections();
  render();
}

function openProjectFile(event) {
  const [file] = event.target.files || [];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const parsed = JSON.parse(String(reader.result || "{}"));
      restoreProject(parsed);
    } catch (error) {
      window.alert(error instanceof Error ? error.message : "Could not open project file.");
    }
    event.target.value = "";
  };
  reader.readAsText(file);
}

// Shared SVG -> raster canvas pipeline for both PDF and PNG export. The grid
// and sub-grid rects only get their fill from external CSS, so a plain clone
// serialized without stylesheets defaults them to black and paints the whole
// page. We explicitly neutralize those fills (root cause of the "black
// background" export) and inline the styling the raster needs.
async function rasterizeLayout({ transparent = false, panelColor = null } = {}) {
  const clone = els.canvas.cloneNode(true);
  // The live canvas carries the on-screen zoom as an inline CSS transform
  // (see updateZoom). cloneNode copies that verbatim, which would bake the
  // current zoom level into the exported raster and crop it to whatever was
  // on screen. Exports must always cover the full layout bounds regardless
  // of zoom/pan, so strip it before rasterizing.
  clone.style.transform = "";
  clone.style.transformOrigin = "";
  clone.querySelector("#canvasPreview")?.replaceChildren();
  clone.querySelector("#canvasGuides")?.replaceChildren();
  clone.querySelector("#canvasMarquee")?.replaceChildren();
  clone.querySelector("#canvasGrid")?.setAttribute("fill", "none");
  clone.querySelector("#canvasSubgrid")?.setAttribute("fill", "none");

  const background = clone.querySelector("#canvasBackground");
  if (background) background.setAttribute("fill", transparent ? "none" : "#ffffff");

  clone.querySelectorAll(".panel-shape").forEach((shape) => {
    if (panelColor) {
      shape.setAttribute("fill", panelColor);
      shape.setAttribute("stroke", "none");
    } else {
      shape.setAttribute("stroke", "#1f2826");
      shape.setAttribute("stroke-width", "2.2");
    }
  });
  clone.querySelectorAll(".anchor-point").forEach((anchor) => anchor.remove());
  if (panelColor) clone.querySelectorAll(".panel-label").forEach((label) => label.remove());

  clone.setAttribute("xmlns", SVG_NS);
  // Crop tightly to the actual panel layout instead of the full on-screen
  // canvas viewBox, which includes empty grid space around the layout, and
  // size the raster from the LED layout's real pixel resolution (each panel
  // contributes exactly PANEL_PIXEL_SIZE x PANEL_PIXEL_SIZE pixels) rather
  // than an arbitrary on-screen scale.
  const bounds = getLayoutBounds();
  const svgWidth = bounds.width;
  const svgHeight = bounds.height;
  clone.setAttribute("viewBox", `${bounds.minX} ${bounds.minY} ${svgWidth} ${svgHeight}`);
  clone.setAttribute("width", String(svgWidth));
  clone.setAttribute("height", String(svgHeight));

  const svgMarkup = new XMLSerializer().serializeToString(clone);
  const url = URL.createObjectURL(new Blob([svgMarkup], { type: "image/svg+xml;charset=utf-8" }));

  try {
    const image = await new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = () => resolve(img);
      img.onerror = reject;
      img.src = url;
    });

    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.round(svgWidth * EXPORT_PIXELS_PER_UNIT));
    canvas.height = Math.max(1, Math.round(svgHeight * EXPORT_PIXELS_PER_UNIT));
    const context = canvas.getContext("2d");
    if (!transparent) {
      context.fillStyle = "#ffffff";
      context.fillRect(0, 0, canvas.width, canvas.height);
    }
    context.drawImage(image, 0, 0, canvas.width, canvas.height);
    return canvas;
  } finally {
    URL.revokeObjectURL(url);
  }
}

async function saveToPdf() {
  const jspdfApi = window.jspdf?.jsPDF;
  if (!jspdfApi) {
    window.alert("PDF export library did not load.");
    return;
  }

  const bounds = getLayoutBounds();
  const canvas = await rasterizeLayout({ transparent: false });

  const pdf = new jspdfApi({
    orientation: canvas.width >= canvas.height ? "landscape" : "portrait",
    unit: "pt",
    format: "a4",
  });

  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();
  const margin = 36;
  const headerY = 36;
  const projectName = sanitizeProjectName(state.projectName);
  const widthText = bounds.hasPanels ? mmToMetersText(unitsToMm(bounds.width)) : "0.00 m";
  const heightText = bounds.hasPanels ? mmToMetersText(unitsToMm(bounds.height)) : "0.00 m";
  const used = getUsedOrientationCounts();
  const pushOutCount = getPushOutCount();
  const orientationLine = (counts) =>
    `LU ${counts.LU}   LD ${counts.LD}   RU ${counts.RU}   RD ${counts.RD}`;

  let cursorY = headerY;
  pdf.setFont("helvetica", "bold");
  pdf.setFontSize(18);
  pdf.text(projectName, margin, cursorY);
  cursorY += 16;
  pdf.setFont("helvetica", "normal");
  pdf.setFontSize(10);
  pdf.text(`Version v${APP_VERSION}`, margin, cursorY);
  cursorY += 18;
  pdf.text(`Layout width: ${widthText}`, margin, cursorY);
  cursorY += 14;
  pdf.text(`Layout height: ${heightText}`, margin, cursorY);
  cursorY += 18;
  pdf.setFont("helvetica", "bold");
  pdf.text(`Total panels used: ${state.panels.length}`, margin, cursorY);
  cursorY += 16;
  pdf.setFont("helvetica", "normal");
  pdf.text(
    `MG9 squares: ${used.MG9}${pushOutCount ? ` (including ${pushOutCount} Push-Out)` : ""}`,
    margin,
    cursorY
  );
  cursorY += 14;
  pdf.text(`MG12 triangles  -  ${orientationLine(used.MG12)}`, margin, cursorY);
  cursorY += 14;
  pdf.text(`MG13 quarter-circles  -  ${orientationLine(used.MG13)}`, margin, cursorY);
  cursorY += 18;

  if (pushOutCount) {
    const [r, g, b] = hexToRgb(PUSH_OUT_COLOR);
    pdf.setFillColor(r, g, b);
    pdf.setDrawColor(74, 42, 134);
    pdf.rect(margin, cursorY - 9, 12, 12, "FD");
    pdf.setFont("helvetica", "bold");
    pdf.text("Legend:", margin + 18, cursorY);
    pdf.setFont("helvetica", "normal");
    pdf.text(
      `Violet panel labelled "PO" = MG9 Push-Out panel (${pushOutCount} in this layout)`,
      margin + 64,
      cursorY
    );
    cursorY += 20;
  }

  const imageTop = cursorY + 10;
  const availableWidth = pageWidth - margin * 2;
  const availableHeight = pageHeight - imageTop - margin;
  const imageRatio = canvas.width / canvas.height;
  let imageWidth = availableWidth;
  let imageHeight = imageWidth / imageRatio;
  if (imageHeight > availableHeight) {
    imageHeight = availableHeight;
    imageWidth = imageHeight * imageRatio;
  }

  pdf.addImage(
    canvas.toDataURL("image/png"),
    "PNG",
    margin,
    imageTop,
    imageWidth,
    imageHeight,
    undefined,
    "FAST"
  );
  pdf.save(`${projectName}-v${APP_VERSION}.pdf`);
}

async function exportPngPreview() {
  if (!state.panels.length) {
    window.alert("Add panels to the layout before exporting a PNG preview.");
    return;
  }
  const color = els.pngColorInput?.value || "#e48a52";
  const canvas = await rasterizeLayout({ transparent: true, panelColor: color });
  const projectName = sanitizeProjectName(state.projectName);
  canvas.toBlob((blob) => {
    if (!blob) {
      window.alert("Could not generate the PNG preview.");
      return;
    }
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `${projectName}-preview.png`;
    link.click();
    URL.revokeObjectURL(url);
  }, "image/png");
}

function wireLiveTextEvents() {
  els.projectNameInput.addEventListener("input", (event) => {
    state.projectName = sanitizeProjectName(event.target.value);
  });
  els.textInput.addEventListener("input", scheduleAutoGenerate);

  els.heightMetersInput.addEventListener("input", () => {
    scheduleAutoGenerate();
  });

  els.spacingInput.addEventListener("input", scheduleAutoGenerate);
  els.letterStyleInput.addEventListener("change", scheduleAutoGenerate);
}

function wireEvents() {
  els.canvas.addEventListener("click", onCanvasClick);
  els.canvas.addEventListener("pointerdown", onCanvasPointerDown);
  els.canvas.addEventListener("pointermove", onCanvasHover);
  els.canvas.addEventListener("pointerleave", onCanvasLeave);
  window.addEventListener("pointermove", onWindowPointerMove);
  window.addEventListener("pointerup", onWindowPointerUp);

  els.generateTextBtn.addEventListener("click", () => generateTextLayout(true));
  els.loadReferenceBtn.addEventListener("click", () => {
    els.textInput.value = REFERENCE_SAMPLE_TEXT;
    generateTextLayout(false);
  });
  els.saveProjectBtn.addEventListener("click", saveProject);
  els.openProjectBtn.addEventListener("click", () => els.projectFileInput.click());
  els.savePdfBtn.addEventListener("click", () => {
    void saveToPdf();
  });
  els.exportPngBtn?.addEventListener("click", () => {
    void exportPngPreview();
  });
  els.undoBtn.addEventListener("click", undoLastAction);
  els.rotateBtn.addEventListener("click", () => rotateSelectedPanel(ROTATION_STEP));
  els.rotateFineBtn?.addEventListener("click", () => rotateSelectedPanel(ROTATION_STEP_FINE));
  els.rotationInput?.addEventListener("change", (event) => setSelectedRotation(event.target.value));
  els.duplicateBtn.addEventListener("click", duplicateSelectedPanel);
  els.pushOutBtn?.addEventListener("click", togglePushOut);
  els.copyBtn?.addEventListener("click", copySelection);
  els.pasteBtn?.addEventListener("click", pasteClipboard);
  els.deleteBtn.addEventListener("click", removeSelectedPanel);
  els.createGridBtn?.addEventListener("click", startGridPlacement);
  els.clearLayoutBtn.addEventListener("click", clearLayout);
  els.resetInventoryBtn.addEventListener("click", clearAndResetInventory);
  els.projectFileInput.addEventListener("change", openProjectFile);
  els.sectionToggles.forEach((toggle) => {
    toggle.addEventListener("click", () => {
      const sectionKey = toggle.dataset.sectionToggle;
      setSectionCollapsed(sectionKey, !state.collapsedSections[sectionKey]);
    });
  });
  els.zoomInBtn.addEventListener("click", () => {
    state.zoom = Math.min(2, state.zoom + 0.1);
    updateZoom();
  });
  els.zoomOutBtn.addEventListener("click", () => {
    state.zoom = Math.max(0.5, state.zoom - 0.1);
    updateZoom();
  });

  window.addEventListener("keydown", (event) => {
    const typing = isTypingTarget(event.target);
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "z") {
      event.preventDefault();
      undoLastAction();
    }
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "s") {
      event.preventDefault();
      saveProject();
      return;
    }
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "o") {
      event.preventDefault();
      els.projectFileInput.click();
      return;
    }
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "p") {
      event.preventDefault();
      void saveToPdf();
      return;
    }
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "c") {
      if (typing) return;
      if (getSelectedPanels().length) {
        event.preventDefault();
        copySelection();
      }
      return;
    }
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "v") {
      if (typing) return;
      if (state.clipboard?.length) {
        event.preventDefault();
        pasteClipboard();
      }
      return;
    }
    if (event.key === "Escape") {
      setPlacementMode(null);
      return;
    }
    if (typing) return;
    if (event.key === "Delete" || event.key === "Backspace") removeSelectedPanel();
    if (event.key.toLowerCase() === "r") {
      rotateSelectedPanel(event.shiftKey ? ROTATION_STEP_FINE : ROTATION_STEP);
    }
    if (event.key.toLowerCase() === "d") duplicateSelectedPanel();
    if (event.key.toLowerCase() === "g") generateTextLayout(true);
    if (event.key === "1") setPlacementMode("MG9");
    if (event.key === "2") setPlacementMode("MG12");
    if (event.key === "3") setPlacementMode("MG13");
  });

  wireLiveTextEvents();
}

wireEvents();
applyAvailability();
computeConnections();
generateTextLayout(false);
