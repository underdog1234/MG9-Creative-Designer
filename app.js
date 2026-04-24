const SVG_NS = "http://www.w3.org/2000/svg";
const APP_VERSION = "1.1.0";
const MM_TO_UNITS = 0.25;
const PANEL_SIZE_MM = 500;
const PANEL_SIZE_UNITS = PANEL_SIZE_MM * MM_TO_UNITS;
const HALF_PANEL = PANEL_SIZE_UNITS / 2;
const SNAP_DISTANCE_UNITS = 8;
const PLACEMENT_LOCK_DISTANCE = 52;
const ROTATION_STEP = 90;
const CANVAS_PADDING = 180;
const AUTO_REFRESH_DELAY_MS = 120;
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

const defaultInventory = {
  MG9: 320,
  MG12: 20,
  MG13: 20,
};

const state = {
  inventory: { ...defaultInventory },
  panels: [],
  selectedId: null,
  selectedIds: [],
  projectName: "untitled-layout",
  zoom: 1,
  drag: null,
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
    pointer: null,
    preview: null,
  },
};

const els = {
  appVersionBadge: document.querySelector("#appVersionBadge"),
  inventoryList: document.querySelector("#inventoryList"),
  library: document.querySelector("#library"),
  replaceTypeLibrary: document.querySelector("#replaceTypeLibrary"),
  canvas: document.querySelector("#layoutCanvas"),
  canvasPanels: document.querySelector("#canvasPanels"),
  canvasGuides: document.querySelector("#canvasGuides"),
  canvasPreview: document.querySelector("#canvasPreview"),
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
  rotateBtn: document.querySelector("#rotateBtn"),
  duplicateBtn: document.querySelector("#duplicateBtn"),
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
  const diagonal = EDGE_CONNECTORS_5.map((value) => ({
    x: value,
    y: value,
  }));
  return [
    ...EDGE_CONNECTORS_5.map((y) => ({ x: -HALF_PANEL, y })),
    ...EDGE_CONNECTORS_5.map((x) => ({ x, y: HALF_PANEL })),
    ...diagonal,
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

function getGeometry(panel) {
  const base = getPanelBaseGeometry(PANEL_TYPES[panel.type]);
  return {
    ...base,
    points: base.points?.map((point) => rotatePoint(point, panel.rotation)),
    anchors: base.anchors.map((anchor) => {
      const rotated = rotatePoint(anchor, panel.rotation);
      return { x: rotated.x, y: rotated.y };
    }),
    labelOffset: rotatePoint(base.labelOffset, panel.rotation),
  };
}

function createPanel(type, x = 320, y = 260, rotation = 0) {
  return {
    id: `panel-${state.nextId++}`,
    type,
    x,
    y,
    rotation,
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
    inventory: { ...state.inventory },
    panels: state.panels.map((panel) => ({ ...panel })),
    selectedId: state.selectedId,
    selectedIds: [...state.selectedIds],
    projectName: state.projectName,
    placement: {
      active: state.placement.active,
      type: state.placement.type,
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
  state.inventory = { ...snapshot.inventory };
  state.panels = snapshot.panels.map((panel) => ({ ...panel }));
  state.selectedId = snapshot.selectedId;
  state.selectedIds = [...snapshot.selectedIds];
  state.projectName = snapshot.projectName || "untitled-layout";
  state.placement.active = snapshot.placement?.active ?? false;
  state.placement.type = snapshot.placement?.type ?? null;
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

  state.panels.forEach((panel) => {
    usedByType[panel.type] = (usedByType[panel.type] || 0) + 1;
    panel.available = usedByType[panel.type] <= (Number(state.inventory[panel.type]) || 0);
  });
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

function getConnectionMapForPanels(panels) {
  const panelConnections = new Map();

  panels.forEach((panel) => {
    panelConnections.set(panel.id, {
      connected: panels.length <= 1,
      anchorIndices: new Set(),
    });
  });

  for (let i = 0; i < panels.length; i += 1) {
    const panelA = panels[i];
    const anchorsA = getGlobalAnchors(panelA);

    for (let j = i + 1; j < panels.length; j += 1) {
      const panelB = panels[j];
      const anchorsB = getGlobalAnchors(panelB);

      anchorsA.forEach((anchorA) => {
        anchorsB.forEach((anchorB) => {
          if (Math.hypot(anchorA.x - anchorB.x, anchorA.y - anchorB.y) <= SNAP_DISTANCE_UNITS) {
            panelConnections.get(panelA.id).anchorIndices.add(anchorA.index);
            panelConnections.get(panelB.id).anchorIndices.add(anchorB.index);
          }
        });
      });
    }
  }

  if (panels.length > 1) {
    panels.forEach((panel) => {
      panelConnections.get(panel.id).connected =
        panelConnections.get(panel.id).anchorIndices.size >= 2;
    });
  }

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
  const duplicates = selectedPanels.map((selected, index) =>
    createPanel(selected.type, selected.x + PANEL_SIZE_UNITS, selected.y + index * HALF_PANEL, selected.rotation)
  );
  state.panels.push(...duplicates);
  applyAvailability();
  state.selectedId = duplicates[0]?.id || null;
  state.selectedIds = duplicates.map((panel) => panel.id);
  computeConnections();
  render();
}

function rotateSelectedPanel() {
  const selectedPanels = getSelectedPanels();
  if (!selectedPanels.length) return;
  pushHistory();
  state.manualViewLocked = true;
  selectedPanels.forEach((selected) => {
    selected.rotation = (selected.rotation + ROTATION_STEP) % 360;
  });
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

function setPlacementMode(type = null) {
  state.placement.active = Boolean(type);
  state.placement.type = type;
  state.placement.pointer = null;
  state.placement.preview = null;
  renderLibrary();
  renderCanvasPreview();
  renderSelectedMeta();
}

function renderInventory() {
  const usedCounts = getUsedCounts();
  els.inventoryList.innerHTML = "";

  Object.values(PANEL_TYPES).forEach((type) => {
    const row = els.inventoryRowTemplate.content.firstElementChild.cloneNode(true);
    row.querySelector(".inventory-name").textContent = type.name;
    row.querySelector(".inventory-size").textContent = `${type.widthMm} x ${type.heightMm} mm`;

    const input = row.querySelector("input");
    input.value = state.inventory[type.id];
    input.addEventListener("change", (event) => {
      state.inventory[type.id] = Math.max(0, Number(event.target.value) || 0);
      applyAvailability();
      computeConnections();
      renderInventory();
      renderLibrary();
      renderCanvas();
    });

    const used = usedCounts[type.id] || 0;
    const remaining = Math.max((Number(state.inventory[type.id]) || 0) - used, 0);
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
    const remaining = Math.max((Number(state.inventory[type.id]) || 0) - (usedCounts[type.id] || 0), 0);

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
  });
  applyAvailability();
  computeConnections();
  render();
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

function polygonPointsString(points, x, y) {
  return points.map((point) => `${point.x + x},${point.y + y}`).join(" ");
}

function createShapeElement(panel, geometry) {
  const type = PANEL_TYPES[panel.type];

  if (type.shapeKind === "sector") {
    const path = document.createElementNS(SVG_NS, "path");
    path.setAttribute("d", geometry.path);
    path.setAttribute("transform", `translate(${panel.x} ${panel.y}) rotate(${panel.rotation})`);
    path.setAttribute("class", "panel-shape");
    path.setAttribute("fill", type.color);
    return path;
  }

  const polygon = document.createElementNS(SVG_NS, "polygon");
  polygon.setAttribute("points", polygonPointsString(geometry.points, panel.x, panel.y));
  polygon.setAttribute("class", "panel-shape");
  polygon.setAttribute("fill", type.color);
  return polygon;
}

function createTextElement(panel, geometry) {
  const text = document.createElementNS(SVG_NS, "text");
  text.setAttribute("x", panel.x + geometry.labelOffset.x);
  text.setAttribute("y", panel.y + geometry.labelOffset.y);
  text.setAttribute("class", "panel-label");
  text.textContent = panel.type;
  return text;
}

function createAnchors(panel, geometry) {
  const fragment = document.createDocumentFragment();
  const connectedAnchors = state.connectionMap.get(panel.id)?.anchorIndices ?? new Set();

  geometry.anchors.forEach((anchor, index) => {
    const circle = document.createElementNS(SVG_NS, "circle");
    circle.setAttribute("cx", panel.x + anchor.x);
    circle.setAttribute("cy", panel.y + anchor.y);
    circle.setAttribute("r", 3.2);
    circle.setAttribute("class", `anchor-point${connectedAnchors.has(index) ? " connected" : ""}`);
    fragment.appendChild(circle);
  });
  return fragment;
}

function createPreviewAnchors(panel, geometry, connectedIndices = new Set()) {
  const fragment = document.createDocumentFragment();
  geometry.anchors.forEach((anchor, index) => {
    const circle = document.createElementNS(SVG_NS, "circle");
    circle.setAttribute("cx", panel.x + anchor.x);
    circle.setAttribute("cy", panel.y + anchor.y);
    circle.setAttribute("r", 3.2);
    circle.setAttribute("class", `preview-anchor${connectedIndices.has(index) ? " connected" : ""}`);
    fragment.appendChild(circle);
  });
  return fragment;
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

function renderCanvasPreview() {
  els.canvasPreview.innerHTML = "";
  if (!state.placement.active || !state.placement.preview) return;

  const { panel, snapped, valid, connectedIndices } = state.placement.preview;
  const geometry = getGeometry(panel);
  const group = document.createElementNS(SVG_NS, "g");
  group.setAttribute(
    "class",
    `preview-group${snapped && valid ? " snap-ready" : ""}${valid ? "" : " invalid-preview"}`
  );
  group.appendChild(createShapeElement(panel, geometry));
  group.appendChild(createTextElement(panel, geometry));
  group.appendChild(createPreviewAnchors(panel, geometry, connectedIndices));
  els.canvasPreview.appendChild(group);
}

function renderCanvas() {
  computeConnections();
  els.canvasPanels.innerHTML = "";

  state.panels.forEach((panel) => {
    const geometry = getGeometry(panel);
    const isConnected = isPanelConnected(panel.id);
    const classes = [
      "panel-group",
      panel.id === state.selectedId ? "selected" : "",
      state.selectedIds.includes(panel.id) && panel.id !== state.selectedId ? "multi-selected" : "",
      panel.available === false ? "unavailable" : "",
      !isConnected ? "disconnected" : "",
    ]
      .filter(Boolean)
      .join(" ");

    const group = document.createElementNS(SVG_NS, "g");
    group.dataset.id = panel.id;
    group.setAttribute("class", classes);

    group.appendChild(createShapeElement(panel, geometry));
    group.appendChild(createTextElement(panel, geometry));
    group.appendChild(createAnchors(panel, geometry));

    group.addEventListener("pointerdown", onPanelPointerDown);
    group.addEventListener("click", (event) => {
      event.stopPropagation();
      setSelected(panel.id, {
        additive: event.shiftKey || event.ctrlKey || event.metaKey,
        toggle: event.ctrlKey || event.metaKey,
      });
    });

    els.canvasPanels.appendChild(group);
  });

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

function updateCanvasView(bounds) {
  if (state.drag || (state.manualViewLocked && state.panels.length)) {
    els.canvas.setAttribute("viewBox", state.currentViewBox);
    return;
  }
  const width = Math.max(bounds.width + CANVAS_PADDING * 2, 2400);
  const height = Math.max(bounds.height + CANVAS_PADDING * 2, 1600);
  const x = bounds.hasPanels ? bounds.minX - CANVAS_PADDING : 0;
  const y = bounds.hasPanels ? bounds.minY - CANVAS_PADDING : 0;

  state.currentViewBox = `${x} ${y} ${width} ${height}`;
  els.canvas.setAttribute("viewBox", state.currentViewBox);
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

function renderSelectedMeta() {
  const selectedPanels = getSelectedPanels();
  const panel = getPanelById(state.selectedId);
  if (!panel || !selectedPanels.length) {
    els.selectedMeta.classList.add("empty");
    els.selectedMeta.textContent = state.placement.active
      ? `Placement mode: ${state.placement.type}. Click the canvas to place panels or press Esc to cancel.`
      : "Select a panel on the canvas.";
    return;
  }

  if (selectedPanels.length > 1) {
    els.selectedMeta.classList.remove("empty");
    els.selectedMeta.innerHTML = `
      <strong>${selectedPanels.length} panels selected</strong><br />
      Move, rotate, duplicate, or delete them as one group.
    `;
    return;
  }

  const type = PANEL_TYPES[panel.type];
  const connectionState = isPanelConnected(panel.id) ? "Connected" : "Needs connection";
  const stockState = panel.available === false ? "Unavailable" : "Available";
  els.selectedMeta.classList.remove("empty");
  els.selectedMeta.innerHTML = `
    <strong>${type.name}</strong><br />
    Size: ${type.widthMm} x ${type.heightMm} mm<br />
    Rotation: ${panel.rotation} deg<br />
    Position: ${Math.round(unitsToMm(panel.x))} mm, ${Math.round(unitsToMm(panel.y))} mm<br />
    Status: ${connectionState}<br />
    Stock: ${stockState}
  `;
}

function render() {
  els.appVersionBadge.textContent = `v${APP_VERSION}`;
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

function onCanvasPointerMove(event) {
  if (!state.drag) return;
  const pointer = screenToSvg(event);
  const dx = pointer.x - state.drag.originPointerX;
  const dy = pointer.y - state.drag.originPointerY;
  state.drag.originalPanels.forEach((original) => {
    const panel = getPanelById(original.id);
    if (!panel) return;
    panel.x = original.x + dx;
    panel.y = original.y + dy;
  });
  computeConnections();
  renderCanvas();
  renderSelectedMeta();
}

function onCanvasHover(event) {
  if (!state.placement.active || state.drag) return;
  const pointer = screenToSvg(event);
  state.placement.pointer = pointer;
  state.placement.preview = getPlacementPreview(state.placement.type, pointer.x, pointer.y);
  renderCanvasPreview();
}

function onCanvasLeave() {
  if (!state.placement.active || state.drag) return;
  state.placement.pointer = null;
  state.placement.preview = null;
  renderCanvasPreview();
}

function onCanvasPointerUp() {
  if (!state.drag) return;
  state.drag.panelIds.forEach((panelId) => {
    const panel = getPanelById(panelId);
    if (!panel) return;
    snapPanel(panel, PLACEMENT_LOCK_DISTANCE);
  });
  computeConnections();

  state.drag = null;
  render();
}

function onCanvasClick(event) {
  const isBackground =
    event.target === els.canvas ||
    event.target.classList.contains("canvas-bg") ||
    event.target.classList.contains("canvas-grid") ||
    event.target.classList.contains("canvas-subgrid");

  if (isBackground && state.placement.active) {
    const pointer = screenToSvg(event);
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
    return;
  }

  if (isBackground) {
    clearSelection();
  }
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
  state.placement.active = false;
  state.placement.type = null;
  state.placement.pointer = null;
  state.placement.preview = null;
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
  state.placement.active = false;
  state.placement.type = null;
  state.placement.pointer = null;
  state.placement.preview = null;
  state.panels = [];
  state.selectedId = null;
  state.selectedIds = [];
  state.inventory = { ...defaultInventory };
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
  const requestedCounts = generatedPanels.reduce((counts, panel) => {
    counts[panel.type] = (counts[panel.type] || 0) + 1;
    return counts;
  }, {});

  const shortages = Object.keys(requestedCounts).filter(
    (type) => requestedCounts[type] > (Number(state.inventory[type]) || 0)
  );

  return { requestedCounts, shortages };
}

function optimizeGeneratedPanels(generatedPanels) {
  const workingPanels = generatedPanels.map((panel) => ({ ...panel }));
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
    (total, type) => total + (stockCheck.requestedCounts[type] - (Number(state.inventory[type]) || 0)),
    0
  );
  const stockText = unavailableCount > 0 ? ` ${unavailableCount} panels shown as unavailable.` : "";
  els.actualWidthOutput.value = (widthMm / 1000).toFixed(2);
  els.textSummary.textContent = `Height ${mmToMetersText(targetHeightPanels * PANEL_SIZE_MM)}. Actual ${mmToMetersText(widthMm)} x ${mmToMetersText(heightMm)}.${stockText}`;
}

function generateTextLayout(showAlerts = false, recordHistory = true) {
  if (recordHistory) pushHistory();
  state.manualViewLocked = false;
  state.placement.active = false;
  state.placement.type = null;
  state.placement.pointer = null;
  state.placement.preview = null;
  const { generatedPanels, targetWidthPanels, targetHeightPanels, fromLibrary } = buildTextPanels();
  const optimizedPanels = fromLibrary ? generatedPanels : optimizeGeneratedPanels(generatedPanels);
  const stockCheck = checkStockForGeneratedPanels(optimizedPanels);
  updateTextSummary(optimizedPanels, targetWidthPanels, targetHeightPanels, stockCheck);

  if (showAlerts && stockCheck.shortages.length) {
    const shortageText = stockCheck.shortages
      .map((type) => `${type}: need ${stockCheck.requestedCounts[type]}, have ${state.inventory[type] || 0}`)
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
  state.inventory = { ...defaultInventory };
  clearLayout();
}

function serializeProject() {
  return {
    version: APP_VERSION,
    projectName: sanitizeProjectName(state.projectName),
    inventory: { ...state.inventory },
    panels: state.panels.map((panel) => ({
      type: panel.type,
      x: panel.x,
      y: panel.y,
      rotation: panel.rotation,
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
  state.inventory = { ...defaultInventory, ...(project.inventory || {}) };
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
      createPanel(panel.type, Number(panel.x) || 500, Number(panel.y) || 500, Number(panel.rotation) || 0)
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
  state.placement.active = false;
  state.placement.type = null;
  state.placement.pointer = null;
  state.placement.preview = null;
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

async function saveToPdf() {
  const jspdfApi = window.jspdf?.jsPDF;
  if (!jspdfApi) {
    window.alert("PDF export library did not load.");
    return;
  }

  const bounds = getLayoutBounds();
  const serializer = new XMLSerializer();
  const clone = els.canvas.cloneNode(true);
  const previewLayer = clone.querySelector("#canvasPreview");
  if (previewLayer) previewLayer.innerHTML = "";
  const background = clone.querySelector("#canvasBackground");
  if (background) background.setAttribute("style", "fill:#ffffff");
  clone.setAttribute("xmlns", SVG_NS);

  const viewBox = els.canvas.getAttribute("viewBox") || "0 0 2400 1600";
  const [, , widthStr, heightStr] = viewBox.split(" ");
  const svgWidth = Number(widthStr) || 2400;
  const svgHeight = Number(heightStr) || 1600;
  clone.setAttribute("width", String(svgWidth));
  clone.setAttribute("height", String(svgHeight));

  const svgMarkup = serializer.serializeToString(clone);
  const svgBlob = new Blob([svgMarkup], { type: "image/svg+xml;charset=utf-8" });
  const url = URL.createObjectURL(svgBlob);

  try {
    const image = await new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = () => resolve(img);
      img.onerror = reject;
      img.src = url;
    });

    const canvas = document.createElement("canvas");
    const scale = Math.min(1.6, 2200 / Math.max(svgWidth, svgHeight));
    canvas.width = Math.max(1, Math.round(svgWidth * scale));
    canvas.height = Math.max(1, Math.round(svgHeight * scale));
    const context = canvas.getContext("2d");
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, canvas.width, canvas.height);
    context.drawImage(image, 0, 0, canvas.width, canvas.height);

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

    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(18);
    pdf.text(projectName, margin, headerY);
    pdf.setFont("helvetica", "normal");
    pdf.setFontSize(10);
    pdf.text(`Version v${APP_VERSION}`, margin, headerY + 16);
    pdf.text(`Layout width: ${widthText}`, margin, headerY + 34);
    pdf.text(`Layout height: ${heightText}`, margin, headerY + 48);
    pdf.text(`Panels used: ${state.panels.length}`, margin, headerY + 62);

    const imageTop = headerY + 80;
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
  } finally {
    URL.revokeObjectURL(url);
  }
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
  els.canvas.addEventListener("pointermove", onCanvasHover);
  els.canvas.addEventListener("pointerleave", onCanvasLeave);
  window.addEventListener("pointermove", onCanvasPointerMove);
  window.addEventListener("pointerup", onCanvasPointerUp);

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
  els.undoBtn.addEventListener("click", undoLastAction);
  els.rotateBtn.addEventListener("click", rotateSelectedPanel);
  els.duplicateBtn.addEventListener("click", duplicateSelectedPanel);
  els.deleteBtn.addEventListener("click", removeSelectedPanel);
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
    if (event.key === "Escape") {
      setPlacementMode(null);
      return;
    }
    if (typing) return;
    if (event.key === "Delete" || event.key === "Backspace") removeSelectedPanel();
    if (event.key.toLowerCase() === "r") rotateSelectedPanel();
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
