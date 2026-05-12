const DATA_URLS = {
  fr: "./texts_fr.json",
  en: "./texts_en.json",
};

const VIEWBOX = { width: 1672, height: 941 };
const MARKDOWN_BASE_STAGE = { width: 560, height: 315 };
const STORAGE_KEY_PREFIX = "joan-comic-texts";

let data;
let currentLanguage = "fr";
let editMode = false;
let presentationMode = "grid";
let activeIndex = 0;
let selectedFrameId = null;
let selectedBalloonId = null;
let pendingInsertIndex = null;
let resizeTimer = null;
let presentationRedrawTimers = [];
let markdownFitTimers = [];
const drawings = new Map();
const renderedBalloons = new Map();
const markdownLayers = new Map();

const gallery = document.querySelector("#gallery");
const statusEl = document.querySelector("#status");
const languageSelect = document.querySelector("#languageSelect");
const presentationModeSelect = document.querySelector("#presentationMode");
const prevFrameBtn = document.querySelector("#prevFrame");
const nextFrameBtn = document.querySelector("#nextFrame");
const editToggle = document.querySelector("#editToggle");
const addBalloonBtn = document.querySelector("#addBalloon");
const exportBtn = document.querySelector("#exportJson");
const resetBtn = document.querySelector("#resetLocal");
const editorEmpty = document.querySelector("#editorEmpty");
const editorFields = document.querySelector("#editorFields");
const markdownFields = document.querySelector("#markdownFields");
const newFrameFields = document.querySelector("#newFrameFields");
const newFrameType = document.querySelector("#newFrameType");
const newFramePath = document.querySelector("#newFramePath");
const newFramePathLabel = document.querySelector("#newFramePathLabel");
const newFrameMarkdown = document.querySelector("#newFrameMarkdown");
const newFrameMarkdownLabel = document.querySelector("#newFrameMarkdownLabel");
const newFrameFontSize = document.querySelector("#newFrameFontSize");
const newFrameFontLabel = document.querySelector("#newFrameFontLabel");
const confirmNewFrame = document.querySelector("#confirmNewFrame");
const cancelNewFrame = document.querySelector("#cancelNewFrame");
const balloonText = document.querySelector("#balloonText");
const balloonType = document.querySelector("#balloonType");
const fontSize = document.querySelector("#fontSize");
const balloonWidth = document.querySelector("#balloonWidth");
const balloonHeight = document.querySelector("#balloonHeight");
const centerX = document.querySelector("#centerX");
const centerY = document.querySelector("#centerY");
const sourceX = document.querySelector("#sourceX");
const sourceY = document.querySelector("#sourceY");
const deleteBalloon = document.querySelector("#deleteBalloon");
const markdownText = document.querySelector("#markdownText");
const markdownFontSize = document.querySelector("#markdownFontSize");
const markdownFontFamily = document.querySelector("#markdownFontFamily");

init().catch((error) => {
  console.error(error);
  statusEl.textContent = `Erreur: ${error.message}`;
});

async function init() {
  if (!window.SVG) throw new Error("svg.js n'est pas chargé");

  currentLanguage = languageSelect?.value || "fr";
  const loaded = await loadJsonSource(currentLanguage);
  const local = localStorage.getItem(storageKey());
  data = local ? JSON.parse(local) : loaded;
  document.documentElement.lang = currentLanguage;

  renderGallery();
  wireUi();
  selectFrame(data.frames[0]?.id);
  updatePresentation();
  statusEl.textContent = "Chargé";
}

function wireUi() {
  languageSelect.addEventListener("change", async () => {
    await switchLanguage(languageSelect.value);
  });

  presentationModeSelect.addEventListener("change", () => {
    presentationMode = presentationModeSelect.value;
    document.body.classList.toggle("presentation-gallery", presentationMode === "gallery");
    updatePresentation();
    schedulePresentationRedraw();
    scheduleMarkdownFrameFit();
  });

  prevFrameBtn.addEventListener("click", () => moveActive(-1));
  nextFrameBtn.addEventListener("click", () => moveActive(1));
  window.addEventListener("keydown", (event) => {
    if (presentationMode !== "gallery") return;
    if (event.key === "ArrowLeft") moveActive(-1);
    if (event.key === "ArrowRight") moveActive(1);
  });
  window.addEventListener("resize", () => {
    clearTimeout(resizeTimer);
    resizeTimer = setTimeout(() => {
      redrawAllBalloons();
      scheduleMarkdownFrameFit();
      updatePresentation();
    }, 80);
  });

  editToggle.addEventListener("click", () => {
    editMode = !editMode;
    document.body.classList.toggle("editing", editMode);
    addBalloonBtn.disabled = !editMode || !selectedFrameId;
    editToggle.textContent = editMode ? "Quitter édition" : "Mode édition";
    updateEditorVisibility();
    redrawAllBalloons();
  });

  addBalloonBtn.addEventListener("click", () => {
    const frame = selectedFrame();
    if (!frame) return;
    const balloon = {
      id: crypto.randomUUID(),
      type: "dialogue",
      text: "Nouveau dialogue",
      center: { x: 0.5, y: 0.18 },
      source: { x: 0.5, y: 0.42 },
      width: 360,
      height: 150,
      fontSize: 30,
    };
    frame.balloons.push(balloon);
    drawBalloonForFrame(frame, balloon);
    selectBalloon(frame.id, balloon.id);
    persist();
  });

  exportBtn.addEventListener("click", () => {
    const blob = new Blob([`${JSON.stringify(data, null, 2)}\n`], { type: "application/json" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `texts_${currentLanguage}.json`;
    link.click();
    URL.revokeObjectURL(link.href);
  });

  resetBtn.addEventListener("click", async () => {
    statusEl.textContent = `Rechargement de texts_${currentLanguage}.json...`;
    localStorage.removeItem(storageKey());
    data = await loadJsonSource(currentLanguage);
    selectedFrameId = data.frames[0]?.id || null;
    selectedBalloonId = null;
    pendingInsertIndex = null;
    activeIndex = 0;
    renderGallery();
    if (selectedFrameId) selectFrame(selectedFrameId);
    updateEditorVisibility();
    updatePresentation();
    statusEl.textContent = `texts_${currentLanguage}.json rechargé`;
  });

  for (const input of [balloonText, balloonType, fontSize, balloonWidth, balloonHeight, centerX, centerY, sourceX, sourceY]) {
    input.addEventListener("input", updateSelectedFromPanel);
  }

  for (const input of [markdownText, markdownFontSize, markdownFontFamily]) {
    input.addEventListener("input", updateSelectedMarkdownFrame);
  }

  newFrameType.addEventListener("change", updateNewFrameTypeVisibility);
  confirmNewFrame.addEventListener("click", confirmInsertFrame);
  cancelNewFrame.addEventListener("click", () => {
    pendingInsertIndex = null;
    updateEditorVisibility();
  });

  deleteBalloon.addEventListener("click", () => {
    const frame = selectedFrame();
    if (!frame || !selectedBalloonId) return;
    frame.balloons = frame.balloons.filter((balloon) => balloon.id !== selectedBalloonId);
    removeRenderedBalloon(frame.id, selectedBalloonId);
    selectedBalloonId = null;
    updateEditorVisibility();
    persist();
  });
}

async function switchLanguage(language) {
  if (!DATA_URLS[language] || language === currentLanguage) return;
  statusEl.textContent = `Chargement de texts_${language}.json...`;
  currentLanguage = language;
  document.documentElement.lang = language;
  const loaded = await loadJsonSource(language);
  const local = localStorage.getItem(storageKey());
  data = local ? JSON.parse(local) : loaded;
  selectedFrameId = data.frames[activeIndex]?.id || data.frames[0]?.id || null;
  selectedBalloonId = null;
  pendingInsertIndex = null;
  renderGallery();
  if (selectedFrameId) selectFrame(selectedFrameId);
  updateEditorVisibility();
  updatePresentation();
  schedulePresentationRedraw();
  statusEl.textContent = `texts_${language}.json chargé`;
}

async function loadJsonSource(language = currentLanguage) {
  const url = DATA_URLS[language] || DATA_URLS.fr;
  const separator = url.includes("?") ? "&" : "?";
  const response = await fetch(`${url}${separator}v=${Date.now()}`, { cache: "no-store" });
  if (!response.ok) throw new Error(`Impossible de charger ${url}`);
  return response.json();
}

function renderGallery() {
  gallery.innerHTML = "";
  drawings.clear();
  renderedBalloons.clear();
  markdownLayers.clear();

  data.frames.forEach((frame, index) => {
    const card = document.createElement("article");
    card.className = "frame-card";
    card.dataset.frameId = frame.id;
    card.dataset.index = String(index);

    const stage = document.createElement("div");
    stage.className = "stage";
    const controls = createFrameControls(index);
    card.append(stage, controls);
    gallery.append(card);

    card.addEventListener("click", () => {
      activeIndex = index;
      selectFrame(frame.id);
      updatePresentation();
    });

    const draw = SVG().addTo(stage).viewbox(0, 0, VIEWBOX.width, VIEWBOX.height);
    drawings.set(frame.id, draw);
    if (frame.markdown != null) {
      renderMarkdownFrame(stage, frame);
    } else {
      draw.image(imageHref(frame.image)).size(VIEWBOX.width, VIEWBOX.height).move(0, 0);
    }
    frame.balloons.forEach((balloon) => drawBalloonForFrame(frame, balloon));
  });
}

function imageHref(image) {
  const path = image.startsWith("scenes/") ? image.slice("scenes/".length) : image;
  return `./scenes/${encodeURIComponent(path).replaceAll("%2F", "/")}`;
}

function drawBalloonForFrame(frame, balloon) {
  const draw = drawings.get(frame.id);
  if (!draw) return;
  removeRenderedBalloon(frame.id, balloon.id);
  const rendered = createBalloon(draw, frame, balloon);
  renderedBalloons.set(balloonKey(frame.id, balloon.id), rendered);
}

function createBalloon(draw, frame, balloon) {
  const group = draw.group().attr({ "data-balloon-id": balloon.id });
  const rendered = { group, shapes: [], htmlNode: null, handles: {}, balloon, frame };
  buildBalloon(rendered);

  group.on("click", (event) => {
    if (!editMode) return;
    event.stopPropagation();
    selectBalloon(frame.id, balloon.id);
  });

  return rendered;
}

function buildBalloon(rendered) {
  const { group, balloon, frame } = rendered;
  group.clear();
  rendered.shapes = [];
  rendered.handles = {};
  rendered.htmlNode?.remove();
  rendered.htmlNode = null;

  const center = toPx(balloon.center);
  const source = toPx(balloon.source);
  const measured = renderBalloonHtml(rendered);
  const textPaddingX = balloon.type === "letter" ? 0 : 42;
  const textPaddingY = balloon.type === "letter" ? 0 : 30;
  const size = balloonSize(balloon, measured, textPaddingX, textPaddingY);
  const { width, height } = size;
  balloon.height = Math.round(height);
  const left = center.x - width / 2;
  const top = center.y - height / 2;
  const fillOpacity = 0.8;
  const strokeWidth = balloon.type === "scream" ? 5 : 3;

  if (balloon.type !== "letter") {
    const edge = tailEdge(center, source, width, height);
    const tail = group.path(`M ${source.x} ${source.y} L ${edge.x - 24} ${edge.y} L ${edge.x + 24} ${edge.y} Z`)
      .fill({ color: "#fffdf7", opacity: fillOpacity })
      .stroke({ color: "#16110d", width: strokeWidth, linejoin: "round" });
    rendered.tail = tail;
    rendered.shapes.push(tail);

    const bubble = group.rect(width, height)
      .move(left, top)
      .radius(12)
      .fill({ color: balloon.type === "caption" ? "#fff2bd" : "#fffdf7", opacity: fillOpacity })
      .stroke({ color: selectedBalloonId === balloon.id ? "#d09a42" : "#16110d", width: selectedBalloonId === balloon.id ? 6 : strokeWidth });
    rendered.bubble = bubble;
    rendered.shapes.push(bubble);

    if (balloon.type === "thought" || balloon.type === "caption") {
      tail.hide();
    }
    if (balloon.type === "thought") {
      rendered.shapes.push(
        group.circle(30).center((source.x + center.x) / 2, (source.y + center.y) / 2).fill({ color: "#fffdf7", opacity: fillOpacity }).stroke({ color: "#16110d", width: 3 }),
        group.circle(18).center(source.x, source.y).fill({ color: "#fffdf7", opacity: fillOpacity }).stroke({ color: "#16110d", width: 3 }),
      );
    }
  }

  positionBalloonHtml(rendered, center, width, height, textPaddingX, textPaddingY);
  if (editMode) addEditHandles(rendered);
}

function balloonSize(balloon, measured, paddingX, paddingY) {
  if (balloon.type === "letter") return measured;

  const contentWidth = measured.width + paddingX * 2;
  const contentHeight = measured.height + paddingY * 2;
  if (balloon.type !== "dialogue") {
    return { width: contentWidth, height: contentHeight };
  }

  const ratio = 16 / 9;
  let width = Math.max(contentWidth, contentHeight * ratio);
  let height = width / ratio;
  if (height < contentHeight) {
    height = contentHeight;
    width = height * ratio;
  }
  return { width, height };
}

function renderBalloonHtml(rendered) {
  const { balloon, frame } = rendered;
  const draw = drawings.get(frame.id);
  const stage = draw?.node.parentElement;
  if (!stage) return { width: balloon.width || 360, height: balloon.height || 150 };

  const stageWidth = stage.clientWidth || stage.getBoundingClientRect().width;
  const stageHeight = stage.clientHeight || stage.getBoundingClientRect().height;
  const scaleX = stageWidth / VIEWBOX.width || 1;
  const scaleY = stageHeight / VIEWBOX.height || scaleX;
  const scale = Math.min(scaleX, scaleY);
  const center = normalizedPoint(balloon.center);
  const node = document.createElement("div");
  node.className = `balloon-html ${balloon.type || "dialogue"}`;
  node.innerHTML = markdownToHtml(balloon.text || "");
  node.style.fontSize = `${(balloon.fontSize || 28) * scale}px`;
  node.style.maxWidth = `${(balloon.width || 360) * scale}px`;
  node.style.left = `${center.x * stageWidth}px`;
  node.style.top = `${center.y * stageHeight}px`;
  node.style.transform = "translate(-50%, -50%)";
  stage.append(node);

  const rect = node.getBoundingClientRect();
  rendered.htmlNode = node;
  return {
    width: Math.max(20, rect.width / scale),
    height: Math.max(20, rect.height / scale),
  };
}

function positionBalloonHtml(rendered, center, width, height, paddingX, paddingY) {
  const { balloon, htmlNode } = rendered;
  if (!htmlNode) return;
  const draw = drawings.get(rendered.frame.id);
  const stage = draw?.node.parentElement;
  if (!stage) return;

  const stageWidth = stage.clientWidth || stage.getBoundingClientRect().width;
  const stageHeight = stage.clientHeight || stage.getBoundingClientRect().height;
  const scaleX = stageWidth / VIEWBOX.width || 1;
  const scaleY = stageHeight / VIEWBOX.height || scaleX;
  const scale = Math.min(scaleX, scaleY);

  if (balloon.type === "letter") {
    htmlNode.style.left = `${center.x * scale}px`;
    htmlNode.style.top = `${center.y * scale}px`;
    htmlNode.style.width = "";
    htmlNode.style.height = "";
    htmlNode.style.transform = "translate(-50%, -50%)";
    return;
  }

  htmlNode.style.left = `${(center.x - width / 2 + paddingX) * scale}px`;
  htmlNode.style.top = `${(center.y - height / 2 + paddingY) * scale}px`;
  htmlNode.style.width = `${Math.max(1, width - paddingX * 2) * scale}px`;
  htmlNode.style.height = `${Math.max(1, height - paddingY * 2) * scale}px`;
  htmlNode.style.transform = "none";
}

function addEditHandles(rendered) {
  const { group, balloon, frame } = rendered;
  const center = toPx(balloon.center);
  const source = toPx(balloon.source);

  const centerHandle = group.circle(20)
    .addClass("edit-handle")
    .center(center.x, center.y)
    .fill("#d09a42")
    .stroke({ color: "#16110d", width: 3 });
  rendered.handles.center = centerHandle;
  centerHandle.on("pointerdown", (event) => startCenterDrag(event, rendered));

  const sourceHandle = group.circle(18)
    .addClass("edit-handle")
    .center(source.x, source.y)
    .fill(selectedBalloonId === balloon.id ? "#d86c5f" : "#f2c14e")
    .stroke({ color: "#16110d", width: 3 });
  rendered.handles.source = sourceHandle;
  sourceHandle.on("pointerdown", (event) => startSourceDrag(event, frame, balloon));
}

function startCenterDrag(event, rendered) {
  if (!editMode) return;
  event.preventDefault();
  event.stopPropagation();
  const { balloon, frame } = rendered;
  selectBalloon(frame.id, balloon.id);
  const start = pointerSvg(event, drawings.get(frame.id));
  const initial = { ...balloon.center };

  const move = (moveEvent) => {
    const current = pointerSvg(moveEvent, drawings.get(frame.id));
    const dx = current.x - start.x;
    const dy = current.y - start.y;
    balloon.center.x = clamp(initial.x + dx / VIEWBOX.width, 0.03, 0.97);
    balloon.center.y = clamp(initial.y + dy / VIEWBOX.height, 0.03, 0.97);
    buildBalloon(rendered);
    loadPanel();
  };

  const up = () => {
    window.removeEventListener("pointermove", move);
    window.removeEventListener("pointerup", up);
    persist();
  };

  window.addEventListener("pointermove", move);
  window.addEventListener("pointerup", up);
}

function startSourceDrag(event, frame, balloon) {
  if (!editMode) return;
  event.preventDefault();
  event.stopPropagation();
  selectBalloon(frame.id, balloon.id);
  const move = (moveEvent) => {
    const point = pointerSvg(moveEvent, drawings.get(frame.id));
    balloon.source.x = clamp(point.x / VIEWBOX.width, 0.02, 0.98);
    balloon.source.y = clamp(point.y / VIEWBOX.height, 0.02, 0.98);
    const rendered = renderedBalloons.get(balloonKey(frame.id, balloon.id));
    if (rendered) buildBalloon(rendered);
    loadPanel();
  };
  const up = () => {
    window.removeEventListener("pointermove", move);
    window.removeEventListener("pointerup", up);
    persist();
  };
  window.addEventListener("pointermove", move);
  window.addEventListener("pointerup", up);
}

function pointerSvg(event, draw) {
  const svg = draw.root().node;
  const point = svg.createSVGPoint();
  point.x = event.clientX;
  point.y = event.clientY;
  return point.matrixTransform(svg.getScreenCTM().inverse());
}

function selectFrame(frameId) {
  if (!frameId) return;
  selectedFrameId = frameId;
  selectedBalloonId = null;
  pendingInsertIndex = null;
  activeIndex = Math.max(0, data.frames.findIndex((frame) => frame.id === frameId));
  addBalloonBtn.disabled = !editMode;
  document.querySelectorAll(".frame-card").forEach((card) => {
    card.classList.toggle("selected", card.dataset.frameId === frameId);
  });
  updateEditorVisibility();
  updatePresentation();
}

function selectBalloon(frameId, balloonId) {
  selectFrame(frameId);
  selectedBalloonId = balloonId;
  loadPanel();
  updateEditorVisibility();
  redrawAllBalloons();
}

function loadPanel() {
  const balloon = selectedBalloon();
  if (!balloon) return;
  balloonText.value = balloon.text || "";
  balloonType.value = balloon.type || "dialogue";
  fontSize.value = balloon.fontSize || 28;
  balloonWidth.value = balloon.width || 360;
  balloonHeight.value = balloon.height || 150;
  centerX.value = formatCoord(balloon.center.x);
  centerY.value = formatCoord(balloon.center.y);
  sourceX.value = formatCoord(balloon.source.x);
  sourceY.value = formatCoord(balloon.source.y);
}

function loadMarkdownPanel() {
  const frame = selectedFrame();
  if (!frame || frame.markdown == null) return;
  markdownText.value = frame.markdown || "";
  markdownFontSize.value = frame.fontSize || 38;
  markdownFontFamily.value = frame.fontFamily || "Bastarda";
}

function updateSelectedFromPanel() {
  const balloon = selectedBalloon();
  if (!balloon) return;
  balloon.text = balloonText.value;
  balloon.type = balloonType.value;
  balloon.fontSize = Number(fontSize.value) || 28;
  balloon.width = Number(balloonWidth.value) || 360;
  balloon.height = Number(balloonHeight.value) || 150;
  balloon.center.x = clamp(Number(centerX.value) || 0, 0, 1);
  balloon.center.y = clamp(Number(centerY.value) || 0, 0, 1);
  balloon.source.x = clamp(Number(sourceX.value) || 0, 0, 1);
  balloon.source.y = clamp(Number(sourceY.value) || 0, 0, 1);
  persist();
  const rendered = renderedBalloons.get(balloonKey(selectedFrameId, balloon.id));
  if (rendered) buildBalloon(rendered);
}

function updateSelectedMarkdownFrame() {
  const frame = selectedFrame();
  if (!frame || frame.markdown == null) return;
  frame.markdown = markdownText.value;
  frame.fontSize = Number(markdownFontSize.value) || 38;
  frame.fontFamily = markdownFontFamily.value || "Bastarda";
  redrawFrame(frame);
  persist();
}

function selectedFrame() {
  return data.frames.find((frame) => frame.id === selectedFrameId);
}

function selectedBalloon() {
  return selectedFrame()?.balloons.find((balloon) => balloon.id === selectedBalloonId);
}

function redrawAllBalloons() {
  for (const frame of data.frames) {
    for (const balloon of frame.balloons) {
      drawBalloonForFrame(frame, balloon);
    }
  }
}

function schedulePresentationRedraw() {
  presentationRedrawTimers.forEach((timer) => clearTimeout(timer));
  presentationRedrawTimers = [];
  requestAnimationFrame(() => {
    redrawAllBalloons();
  });
  for (const delay of [80, 240, 460]) {
    presentationRedrawTimers.push(setTimeout(() => {
      redrawAllBalloons();
    }, delay));
  }
}

function scheduleMarkdownFrameFit() {
  markdownFitTimers.forEach((timer) => clearTimeout(timer));
  markdownFitTimers = [];
  requestAnimationFrame(() => {
    fitAllMarkdownFrames();
  });
  for (const delay of [80, 240, 460]) {
    markdownFitTimers.push(setTimeout(() => {
      fitAllMarkdownFrames();
    }, delay));
  }
}

function fitAllMarkdownFrames() {
  markdownLayers.forEach((layer, frameId) => {
    const frame = data.frames.find((item) => item.id === frameId);
    if (frame) fitMarkdownFrame(layer, frame);
  });
}

function removeRenderedBalloon(frameId, balloonId) {
  const key = balloonKey(frameId, balloonId);
  const rendered = renderedBalloons.get(key);
  rendered?.group.remove();
  rendered?.htmlNode?.remove();
  renderedBalloons.delete(key);
}

function createFrameControls(index) {
  const controls = document.createElement("div");
  controls.className = "frame-controls";

  if (index === 0) {
    controls.append(createInsertButton(index, "before"));
  }

  controls.append(createInsertButton(index + 1, "after"));

  const deleteButton = document.createElement("button");
  deleteButton.type = "button";
  deleteButton.className = "delete-frame";
  deleteButton.textContent = "×";
  deleteButton.title = "Supprimer cette scène";
  deleteButton.addEventListener("click", (event) => {
    event.stopPropagation();
    deleteFrameAt(index);
  });
  controls.append(deleteButton);

  return controls;
}

function createInsertButton(index, position) {
  const button = document.createElement("button");
  button.type = "button";
  button.className = `insert-frame ${position}`;
  button.textContent = "+";
  button.title = "Insérer une nouvelle scène ici";
  button.addEventListener("click", (event) => {
    event.stopPropagation();
    insertSceneAt(index);
  });
  return button;
}

function insertSceneAt(index) {
  if (!editMode) return;
  pendingInsertIndex = index;
  newFrameType.value = "image";
  newFramePath.value = "scenes/nouvelle-scene.png";
  newFrameMarkdown.value = "";
  newFrameFontSize.value = 38;
  updateNewFrameTypeVisibility();
  updateEditorVisibility();
}

function confirmInsertFrame() {
  if (!editMode || pendingInsertIndex == null) return;
  const isMarkdown = newFrameType.value === "markdown";
  const markdown = newFrameMarkdown.value.trim();
  const normalizedImage = newFramePath.value.trim().replace(/^\.?\//, "").replace(/^scenes\//, "");
  if (isMarkdown && !markdown) return;
  if (!isMarkdown && !normalizedImage) return;
  const id = uniqueFrameId(isMarkdown ? markdown.slice(0, 32) || "markdown" : normalizedImage);
  const frame = {
    id,
    ...(isMarkdown ? { markdown, fontSize: Number(newFrameFontSize.value) || 38, fontFamily: "Bastarda" } : { image: normalizedImage }),
    balloons: [],
  };
  const index = pendingInsertIndex;
  data.frames.splice(index, 0, frame);
  activeIndex = index;
  selectedFrameId = id;
  selectedBalloonId = null;
  pendingInsertIndex = null;
  persist();
  renderGallery();
  selectFrame(id);
}

function deleteFrameAt(index) {
  if (!editMode) return;
  const frame = data.frames[index];
  if (!frame) return;
  data.frames.splice(index, 1);
  activeIndex = clamp(Math.min(index, data.frames.length - 1), 0, Math.max(0, data.frames.length - 1));
  selectedFrameId = data.frames[activeIndex]?.id || null;
  selectedBalloonId = null;
  persist();
  renderGallery();
  if (selectedFrameId) selectFrame(selectedFrameId);
  updateEditorVisibility();
}

function uniqueFrameId(image) {
  const base = image
    .replace(/\.[^.]+$/, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-|-$/g, "") || "scene";
  let id = base;
  let counter = 2;
  const used = new Set(data.frames.map((frame) => frame.id));
  while (used.has(id)) {
    id = `${base}-${counter}`;
    counter += 1;
  }
  return id;
}

function updateEditorVisibility() {
  const insertingFrame = Boolean(editMode && pendingInsertIndex != null);
  const hasSelection = Boolean(editMode && selectedBalloon());
  const hasMarkdownFrame = Boolean(editMode && !insertingFrame && !selectedBalloonId && selectedFrame()?.markdown != null);
  newFrameFields.hidden = !insertingFrame;
  editorFields.hidden = insertingFrame || !hasSelection;
  markdownFields.hidden = !hasMarkdownFrame;
  editorEmpty.hidden = insertingFrame || hasSelection || hasMarkdownFrame;
  if (hasMarkdownFrame) loadMarkdownPanel();
}

function updateNewFrameTypeVisibility() {
  const isMarkdown = newFrameType.value === "markdown";
  newFramePathLabel.hidden = isMarkdown;
  newFrameMarkdownLabel.hidden = !isMarkdown;
  newFrameFontLabel.hidden = !isMarkdown;
}

function moveActive(delta) {
  activeIndex = clamp(activeIndex + delta, 0, data.frames.length - 1);
  selectFrame(data.frames[activeIndex].id);
  schedulePresentationRedraw();
}

function updatePresentation() {
  document.body.classList.toggle("presentation-gallery", presentationMode === "gallery");
  document.querySelectorAll(".frame-card").forEach((card) => {
    const index = Number(card.dataset.index);
    card.classList.toggle("active", index === activeIndex);
  });
  if (presentationMode !== "gallery") {
    gallery.style.removeProperty("--gallery-offset");
    return;
  }
  requestAnimationFrame(() => {
    const active = document.querySelector(`.frame-card[data-index="${activeIndex}"]`);
    if (!active) return;
    const shellRect = gallery.parentElement.getBoundingClientRect();
    const activeRect = active.getBoundingClientRect();
    const current = Number.parseFloat(gallery.style.getPropertyValue("--gallery-offset")) || 0;
    const desired = current + shellRect.left + shellRect.width / 2 - (activeRect.left + activeRect.width / 2);
    gallery.style.setProperty("--gallery-offset", `${desired}px`);
  });
}

function renderMarkdownFrame(stage, frame) {
  removeMarkdownLayer(frame.id);
  const layer = document.createElement("div");
  layer.className = "markdown-frame";
  layer.style.setProperty("--markdown-font-size", `${frame.fontSize || 38}px`);
  layer.style.setProperty("--markdown-font-family", fontFamilyValue(frame.fontFamily || "Bastarda"));
  layer.innerHTML = markdownToHtml(frame.markdown);
  stage.prepend(layer);
  markdownLayers.set(frame.id, layer);
  requestAnimationFrame(() => fitMarkdownFrame(layer, frame));
  document.fonts?.ready.then(() => fitMarkdownFrame(layer, frame));
}

function fontFamilyValue(fontFamily) {
  if (fontFamily === "system") return "system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif";
  return `"${fontFamily}", serif`;
}

function fitMarkdownFrame(layer, frame) {
  const stage = layer.parentElement;
  if (!stage) return;

  const stageWidth = stage.clientWidth || stage.getBoundingClientRect().width;
  const stageHeight = stage.clientHeight || stage.getBoundingClientRect().height;
  const scaleX = stageWidth / MARKDOWN_BASE_STAGE.width || 1;
  const scaleY = stageHeight / MARKDOWN_BASE_STAGE.height || scaleX;
  const baseFontSize = Number(frame.fontSize) || 38;
  const targetFontSize = Math.max(1, baseFontSize * Math.min(1, scaleX, scaleY));

  layer.style.setProperty("--markdown-font-size", `${targetFontSize}px`);

  if (!markdownLayerOverflows(layer)) return;

  let low = 1;
  let high = targetFontSize;
  for (let index = 0; index < 10; index += 1) {
    const mid = (low + high) / 2;
    layer.style.setProperty("--markdown-font-size", `${mid}px`);
    if (markdownLayerOverflows(layer)) {
      high = mid;
    } else {
      low = mid;
    }
  }
  layer.style.setProperty("--markdown-font-size", `${low}px`);
}

function markdownLayerOverflows(layer) {
  return layer.scrollHeight > layer.clientHeight + 1 || layer.scrollWidth > layer.clientWidth + 1;
}

function redrawFrame(frame) {
  const draw = drawings.get(frame.id);
  if (!draw) return;
  draw.clear();
  removeMarkdownLayer(frame.id);
  renderedBalloons.forEach((rendered, key) => {
    if (key.startsWith(`${frame.id}:`)) renderedBalloons.delete(key);
  });
  if (frame.markdown != null) {
    const stage = draw.node.parentElement;
    if (stage) renderMarkdownFrame(stage, frame);
  } else {
    draw.image(imageHref(frame.image)).size(VIEWBOX.width, VIEWBOX.height).move(0, 0);
  }
  frame.balloons.forEach((balloon) => drawBalloonForFrame(frame, balloon));
}

function removeMarkdownLayer(frameId) {
  const layer = markdownLayers.get(frameId);
  if (!layer) return;
  layer.remove();
  markdownLayers.delete(frameId);
}

function markdownToHtml(markdown) {
  const blocks = String(markdown)
    .replace(/\r\n/g, "\n")
    .split(/\n{2,}/)
    .map((block) => block.trim())
    .filter(Boolean);

  if (!blocks.length) return "";

  return blocks.map((block) => {
    const lines = block.split("\n").map((line) => line.trim()).filter(Boolean);
    if (!lines.length) return "";
    if (lines.length === 1 && lines[0].startsWith("# ")) {
      return `<h1>${formatInlineMarkdown(lines[0].slice(2))}</h1>`;
    }
    if (lines.length === 1 && lines[0].startsWith("## ")) {
      return `<h2>${formatInlineMarkdown(lines[0].slice(3))}</h2>`;
    }
    if (lines.every((line) => /^[-*]\s+/.test(line))) {
      const items = lines.map((line) => `<li>${formatInlineMarkdown(line.replace(/^[-*]\s+/, ""))}</li>`).join("");
      return `<ul>${items}</ul>`;
    }
    return `<p>${lines.map((line) => formatInlineMarkdown(line)).join("<br>")}</p>`;
  }).join("");
}

function formatInlineMarkdown(value) {
  return escapeHtml(value)
    .replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>")
    .replace(/\*([^*]+)\*/g, "<em>$1</em>");
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function persist() {
  localStorage.setItem(storageKey(), JSON.stringify(data));
}

function storageKey(language = currentLanguage) {
  return `${STORAGE_KEY_PREFIX}-${language}-v2`;
}

function balloonKey(frameId, balloonId) {
  return `${frameId}:${balloonId}`;
}

function toPx(point) {
  const normalized = normalizedPoint(point);
  return {
    x: normalized.x * VIEWBOX.width,
    y: normalized.y * VIEWBOX.height,
  };
}

function normalizedPoint(point) {
  return {
    x: clamp(Number(point?.x) || 0, 0, 1),
    y: clamp(Number(point?.y) || 0, 0, 1),
  };
}

function tailEdge(center, source, width, height) {
  const dx = source.x - center.x;
  const dy = source.y - center.y;
  const scale = Math.max(Math.abs(dx) / (width / 2), Math.abs(dy) / (height / 2), 0.001);
  return {
    x: center.x + dx / scale,
    y: center.y + dy / scale,
  };
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function formatCoord(value) {
  return Number(value).toFixed(3);
}
