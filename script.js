const state = {
  images: [],
  customFonts: [],
  systemFonts: [],
  bodyEdits: new Map(),
  titleEdits: new Map(),
  imageEdits: new Map(),
  editingText: false,
  savedSelection: null,
  activeEditable: null,
  draggedImage: null,
  columns: 2,
  seed: Math.floor(Math.random() * 100000),
  zoom: 1,
};

const builtInFonts = [
  {
    id: 'builtin:futurist-mono',
    label: 'Futurist mono',
    family: '"Courier New", "SFMono-Regular", monospace',
    role: 'mono',
  },
  {
    id: 'builtin:editorial-serif',
    label: 'Editorial serif',
    family: 'Georgia, "Times New Roman", serif',
    role: 'serif',
  },
  {
    id: 'builtin:condensed-grotesk',
    label: 'Condensed grotesk',
    family: '"Arial Narrow", "Avenir Next Condensed", Impact, sans-serif',
    role: 'condensed',
  },
  {
    id: 'builtin:swiss-grotesk',
    label: 'Swiss grotesk',
    family: 'Helvetica, Arial, sans-serif',
    role: 'sans',
  },
  {
    id: 'builtin:typewriter',
    label: 'Typewriter',
    family: '"American Typewriter", "Courier New", monospace',
    role: 'type',
  },
  {
    id: 'builtin:system-ui',
    label: 'System UI',
    family: 'Inter, ui-sans-serif, system-ui, sans-serif',
    role: 'sans',
  },
];

const layoutNames = [
  'Controlled chaos',
  'Circuit poetry',
  'Xerox grid',
  'Poster notes',
  'Off-register manual',
  'Small type panic',
];

const palettes = ['#ff2a1c', '#0077ff', '#1f9d55', '#d01b8c', '#f08b00', '#202020'];

const els = {
  text: document.querySelector('#sourceText'),
  imageInput: document.querySelector('#imageInput'),
  fontInput: document.querySelector('#fontInput'),
  displayFont: document.querySelector('#displayFont'),
  bodyFont: document.querySelector('#bodyFont'),
  pageCount: document.querySelector('#pageCount'),
  accentColor: document.querySelector('#accentColor'),
  typeScale: document.querySelector('#typeScale'),
  imageEnergy: document.querySelector('#imageEnergy'),
  monoImages: document.querySelector('#monoImages'),
  imageCount: document.querySelector('#imageCount'),
  fontCount: document.querySelector('#fontCount'),
  systemFontBtn: document.querySelector('#systemFontBtn'),
  systemFontStatus: document.querySelector('#systemFontStatus'),
  generate: document.querySelector('#generateBtn'),
  randomize: document.querySelector('#randomizeBtn'),
  print: document.querySelector('#printBtn'),
  printRoot: document.querySelector('#printRoot'),
  pages: document.querySelector('#pages'),
  imageStrip: document.querySelector('#imageStrip'),
  layoutName: document.querySelector('#layoutName'),
  editText: document.querySelector('#editTextBtn'),
  formatTools: document.querySelector('#formatTools'),
  bold: document.querySelector('#boldBtn'),
  italic: document.querySelector('#italicBtn'),
  highlight: document.querySelector('#highlightBtn'),
  highlightColor: document.querySelector('#highlightColor'),
  fit: document.querySelector('#fitBtn'),
  zoomIn: document.querySelector('#zoomInBtn'),
  zoomOut: document.querySelector('#zoomOutBtn'),
  segments: [...document.querySelectorAll('.segment')],
};

function seededRandom(seed) {
  let value = seed % 2147483647;
  if (value <= 0) value += 2147483646;
  return () => {
    value = (value * 16807) % 2147483647;
    return (value - 1) / 2147483646;
  };
}

function pick(items, rand) {
  return items[Math.floor(rand() * items.length)];
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function splitText(text, pageCount) {
  const clean = text.trim().replace(/\n{3,}/g, '\n\n');
  if (!clean) return [];

  const words = clean.split(/\s+/);
  const perPage = Math.max(1, Math.ceil(words.length / pageCount));
  const chunks = [];

  for (let i = 0; i < pageCount; i += 1) {
    const start = i * perPage;
    const end = i === pageCount - 1 ? words.length : (i + 1) * perPage;
    const chunk = words.slice(start, end).join(' ');
    if (chunk) chunks.push(chunk);
  }

  return chunks;
}

function titleFromText(text) {
  const firstLine = text.trim().split(/\n/).find(Boolean);
  if (!firstLine) return 'Untitled Blook';
  return firstLine.slice(0, 74);
}

function bodyFromText(text) {
  const lines = text.trim().split(/\n/);
  const firstIndex = lines.findIndex((line) => line.trim());
  if (firstIndex === -1 || lines.filter((line) => line.trim()).length < 2) return text;
  lines.splice(firstIndex, 1);
  return lines.join('\n').trim();
}

function phraseFromChunk(chunk) {
  const words = chunk
    .replace(/[^\w\s-]/g, ' ')
    .split(/\s+/)
    .filter((word) => word.length > 3);

  if (words.length < 2) return 'Fanzine';
  const start = Math.max(0, Math.floor(words.length * 0.22));
  return words.slice(start, start + 3).join(' ');
}

function allFonts() {
  return [...builtInFonts, ...state.customFonts, ...state.systemFonts];
}

function quoteFontFamily(name) {
  return `"${name.replace(/\\/g, '\\\\').replace(/"/g, '\\"')}", sans-serif`;
}

function cssFont(fontId) {
  return allFonts().find((font) => font.id === fontId)?.family ?? builtInFonts[0].family;
}

function addFontGroup(select, label, fonts) {
  if (!fonts.length) return;
  const group = document.createElement('optgroup');
  group.label = label;
  fonts.forEach((font) => {
    const option = document.createElement('option');
    option.value = font.id;
    option.textContent = font.label;
    group.append(option);
  });
  select.append(group);
}

function refreshFontSelects() {
  for (const select of [els.displayFont, els.bodyFont]) {
    const previous = select.value;
    select.innerHTML = '';
    addFontGroup(select, 'Built in', builtInFonts);
    addFontGroup(select, 'Uploaded', state.customFonts);
    addFontGroup(select, 'System', state.systemFonts);

    const fallback =
      select === els.displayFont ? 'builtin:condensed-grotesk' : 'builtin:futurist-mono';
    select.value = allFonts().some((font) => font.id === previous) ? previous : fallback;
  }
}

function makeImageStrip() {
  els.imageStrip.innerHTML = '';
  state.images.forEach((image) => {
    const img = document.createElement('img');
    img.className = 'thumb';
    img.src = image.url;
    img.alt = image.name;
    els.imageStrip.append(img);
  });
}

function readFiles(files, mapper) {
  return Promise.all(
    [...files].map(
      (file) =>
        new Promise((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = () => resolve(mapper(file, reader.result));
          reader.onerror = reject;
          reader.readAsDataURL(file);
        }),
    ),
  );
}

function addCustomFont(file, url) {
  const family = `UserFont_${state.customFonts.length}_${file.name.replace(/\W+/g, '_')}`;
  const style = document.createElement('style');
  style.textContent = `@font-face{font-family:${family};src:url("${url}");font-display:swap;}`;
  document.head.append(style);
  return {
    id: `upload:${family}`,
    label: file.name.replace(/\.(ttf|otf|woff2?|TTF|OTF|WOFF2?)$/, ''),
    family,
    role: 'custom',
  };
}

function timeoutAfter(ms) {
  return new Promise((_, reject) => {
    window.setTimeout(() => reject(new Error('Font access timed out')), ms);
  });
}

async function loadSystemFonts() {
  if (!('queryLocalFonts' in window)) {
    els.systemFontStatus.textContent = 'Not supported in this browser';
    return;
  }

  if (!window.isSecureContext) {
    els.systemFontStatus.textContent = 'Use localhost or HTTPS';
    return;
  }

  els.systemFontBtn.disabled = true;
  els.systemFontStatus.textContent = 'Requesting access...';

  try {
    const fonts = await Promise.race([window.queryLocalFonts(), timeoutAfter(15000)]);
    const families = new Map();

    fonts.forEach((font) => {
      if (!font.family || families.has(font.family)) return;
      families.set(font.family, {
        id: `system:${font.family}`,
        label: font.family,
        family: quoteFontFamily(font.family),
        role: 'system',
      });
    });

    state.systemFonts = [...families.values()].sort((a, b) => a.label.localeCompare(b.label));
    els.systemFontStatus.textContent = `${state.systemFonts.length} fonts loaded`;
    refreshFontSelects();
    render();
  } catch (error) {
    if (error.message === 'Font access timed out') {
      els.systemFontStatus.textContent = 'Permission prompt timed out';
    } else {
      els.systemFontStatus.textContent =
        error.name === 'NotAllowedError' ? 'Permission denied' : 'Could not load fonts';
    }
  } finally {
    els.systemFontBtn.disabled = false;
  }
}

function layoutRecipe(index, rand, settings) {
  const types = ['schematic', 'xerox', 'poster', 'manual'];
  const layout = pick(types, rand);
  const energy = settings.imageEnergy / 100;
  const fit = (size, min, max) => min + rand() * Math.max(0, max - min - size);
  const width = 38 + rand() * 48 * energy;
  const height = 14 + rand() * 32 * energy;
  const textWidth = settings.columns === 2 ? 76 + rand() * 12 : 58 + rand() * 28;
  const textHeight = 30 + rand() * 42;
  const imageTop = fit(height, 7, 91);
  const imageLeft = fit(width, 6, 94);
  const textLeft = fit(textWidth, 7, 93);
  const textTop = rand() > 0.5 ? fit(textHeight, 10, 54) : fit(textHeight, 34, 92);
  const displayLeft = fit(54, 7, 92);
  const displayTop = rand() > 0.5 ? 8 + rand() * 14 : 70 + rand() * 13;

  return {
    layout,
    text: {
      left: clamp(textLeft, 5, 32),
      top: clamp(textTop, 8, 70),
      width: clamp(textWidth, 54, 88),
      height: clamp(textHeight, 28, 72),
    },
    image: {
      left: clamp(imageLeft, 4, 54),
      top: clamp(imageTop, 4, 76),
      width: clamp(width, 28, 86),
      height: clamp(height, 12, 50),
    },
    display: {
      left: clamp(displayLeft, 5, 38),
      top: clamp(displayTop, 6, 82),
    },
    spaced: rand() > 0.38,
    outline: rand() > 0.72,
    blackTitle: rand() > 0.82,
    rotate: rand() > 0.5 ? 'rotate-left' : 'rotate-right',
  };
}

function placePercent(el, box) {
  el.style.left = `${box.left}%`;
  el.style.top = `${box.top}%`;
  el.style.width = `${box.width}%`;
  el.style.height = `${box.height}%`;
}

function addRules(page, rand, accent) {
  const count = 2 + Math.floor(rand() * 5);
  for (let i = 0; i < count; i += 1) {
    const rule = document.createElement('span');
    rule.className = `rule ${rand() > 0.78 ? 'red' : ''}`;
    const horizontal = rand() > 0.34;
    rule.style.left = `${4 + rand() * 82}%`;
    rule.style.top = `${5 + rand() * 84}%`;
    rule.style.width = horizontal ? `${8 + rand() * 38}%` : '1px';
    rule.style.height = horizontal ? '1px' : `${6 + rand() * 28}%`;
    rule.style.backgroundColor = rule.classList.contains('red') ? accent : '';
    page.append(rule);
  }
}

function clearPageTextEdits() {
  state.bodyEdits.clear();
  state.titleEdits.clear();
}

function clearImageEdits() {
  state.imageEdits.clear();
}

function updateEditTextMode() {
  els.editText.classList.toggle('is-active', state.editingText);
  els.editText.setAttribute('aria-pressed', String(state.editingText));
  els.pages.classList.toggle('is-editing-text', state.editingText);

  for (const control of [els.bold, els.italic, els.highlight, els.highlightColor]) {
    control.disabled = !state.editingText;
  }

  els.pages.querySelectorAll('.editable-text').forEach((el) => {
    if (state.editingText) {
      el.setAttribute('contenteditable', 'true');
      el.setAttribute('spellcheck', 'false');
      el.setAttribute('tabindex', '0');
    } else {
      el.removeAttribute('contenteditable');
      el.removeAttribute('tabindex');
    }
  });

  if (!state.editingText) {
    state.savedSelection = null;
    state.activeEditable = null;
  }
}

function makeEditableText(el, index, type) {
  el.classList.add('editable-text');
  el.dataset.pageIndex = String(index);
  el.dataset.editType = type;
  updateEditTextMode();
}

function insertPlainText(text) {
  const selection = window.getSelection();
  if (!selection || !selection.rangeCount) return;

  selection.deleteFromDocument();
  selection.getRangeAt(0).insertNode(document.createTextNode(text));
  selection.collapseToEnd();
}

function imageEditKey(index, slot) {
  return `${index}:${slot}`;
}

function applyImageEdit(figure, index, slot) {
  figure.classList.add('movable-image');
  figure.dataset.pageIndex = String(index);
  figure.dataset.imageSlot = slot;

  const edit = state.imageEdits.get(imageEditKey(index, slot));
  if (!edit) return;

  figure.style.left = `${edit.left}%`;
  figure.style.top = `${edit.top}%`;
}

function startImageDrag(event) {
  const figure = event.target.closest('.movable-image');
  if (!figure || !state.editingText) return;

  const page = figure.closest('.book-page');
  if (!page) return;

  event.preventDefault();
  event.stopPropagation();

  const pageRect = page.getBoundingClientRect();
  const figureRect = figure.getBoundingClientRect();
  const width = parseFloat(figure.style.width) || (figureRect.width / pageRect.width) * 100;
  const height = parseFloat(figure.style.height) || (figureRect.height / pageRect.height) * 100;

  state.draggedImage = {
    figure,
    pageRect,
    offsetX: event.clientX - figureRect.left,
    offsetY: event.clientY - figureRect.top,
    width,
    height,
    key: imageEditKey(figure.dataset.pageIndex, figure.dataset.imageSlot),
  };

  figure.classList.add('is-dragging');
  figure.setPointerCapture?.(event.pointerId);
}

function moveDraggedImage(event) {
  if (!state.draggedImage) return;

  const drag = state.draggedImage;
  const left = clamp(
    ((event.clientX - drag.pageRect.left - drag.offsetX) / drag.pageRect.width) * 100,
    0,
    100 - drag.width,
  );
  const top = clamp(
    ((event.clientY - drag.pageRect.top - drag.offsetY) / drag.pageRect.height) * 100,
    0,
    100 - drag.height,
  );

  drag.figure.style.left = `${left}%`;
  drag.figure.style.top = `${top}%`;
  state.imageEdits.set(drag.key, { left, top });
}

function stopImageDrag() {
  if (!state.draggedImage) return;

  state.draggedImage.figure.classList.remove('is-dragging');
  state.draggedImage = null;
}

function editableFromNode(node) {
  const element = node?.nodeType === Node.TEXT_NODE ? node.parentElement : node;
  return element?.closest?.('.editable-text') ?? null;
}

function editableFromSelection() {
  const selection = window.getSelection();
  if (!selection || !selection.rangeCount) return null;
  return editableFromNode(selection.anchorNode);
}

function saveTextSelection() {
  if (!state.editingText) return;

  const selection = window.getSelection();
  if (!selection || !selection.rangeCount) return;

  const editable = editableFromSelection();
  if (!editable) return;

  state.savedSelection = selection.getRangeAt(0).cloneRange();
  state.activeEditable = editable;
}

function restoreTextSelection() {
  if (!state.savedSelection || !state.activeEditable?.isConnected) return false;

  state.activeEditable.focus({ preventScroll: true });
  const selection = window.getSelection();
  selection.removeAllRanges();
  selection.addRange(state.savedSelection);
  return true;
}

function storeEditableOverride(editable) {
  if (!editable) return;

  const pageIndex = Number(editable.dataset.pageIndex);
  const targetMap = editable.dataset.editType === 'title' ? state.titleEdits : state.bodyEdits;
  targetMap.set(pageIndex, editable.innerHTML);
}

function updateFormatState() {
  if (!state.editingText) {
    els.bold.classList.remove('is-active');
    els.italic.classList.remove('is-active');
    return;
  }

  els.bold.classList.toggle('is-active', document.queryCommandState('bold'));
  els.italic.classList.toggle('is-active', document.queryCommandState('italic'));
}

function runTextCommand(command, value = null) {
  if (!state.editingText) return;

  restoreTextSelection();
  const editable = editableFromSelection() ?? state.activeEditable;
  if (!editable) return;

  document.execCommand('styleWithCSS', false, true);
  document.execCommand(command, false, value);
  storeEditableOverride(editable);
  saveTextSelection();
  updateFormatState();
}

function applyTextBackground() {
  if (!state.editingText) return;

  restoreTextSelection();
  const editable = editableFromSelection() ?? state.activeEditable;
  if (!editable) return;

  document.execCommand('styleWithCSS', false, true);
  document.execCommand('backColor', false, els.highlightColor.value);
  storeEditableOverride(editable);
  saveTextSelection();
}

function clonePageForBooklet(page) {
  const clone = page.cloneNode(true);
  clone.querySelectorAll('[contenteditable], [tabindex]').forEach((el) => {
    el.removeAttribute('contenteditable');
    el.removeAttribute('tabindex');
  });
  return clone;
}

function makeBookletSlot(page, side) {
  const slot = document.createElement('div');
  slot.className = `booklet-slot booklet-slot-${side}`;

  if (!page) {
    slot.classList.add('is-blank');
    return slot;
  }

  const frame = document.createElement('div');
  frame.className = 'booklet-page-frame';
  frame.append(clonePageForBooklet(page));
  slot.append(frame);
  return slot;
}

function makeBookletSheet(leftPage, rightPage, index) {
  const sheet = document.createElement('section');
  sheet.className = 'booklet-sheet';
  sheet.dataset.sheet = String(index + 1);
  sheet.append(makeBookletSlot(leftPage, 'left'));
  sheet.append(makeBookletSlot(rightPage, 'right'));
  return sheet;
}

function buildBookletPrint() {
  const pages = [...els.pages.querySelectorAll('.book-page:not(.empty-page)')];
  els.printRoot.innerHTML = '';

  if (!pages.length) return false;

  const documentEl = document.createElement('div');
  documentEl.className = 'booklet-document';

  let low = 0;
  let high = pages.length - 1;
  let sheetIndex = 0;

  while (low <= high) {
    let leftPage, rightPage;
    if (sheetIndex % 2 === 0) {
      leftPage = low === high ? null : pages[high];
      rightPage = pages[low];
    } else {
      leftPage = pages[low];
      rightPage = low === high ? null : pages[high];
    }
    documentEl.append(makeBookletSheet(leftPage, rightPage, sheetIndex));
    low += 1;
    high -= 1;
    sheetIndex += 1;
  }

  els.printRoot.append(documentEl);
  return true;
}

function cleanupBookletPrint() {
  document.body.classList.remove('is-booklet-printing');
  els.printRoot.innerHTML = '';
}

function createBookletPdf() {
  if (!buildBookletPrint()) return;

  document.body.classList.add('is-booklet-printing');
  window.requestAnimationFrame(() => {
    window.requestAnimationFrame(() => window.print());
  });
}

function renderCoverPage(settings) {
  const page = document.createElement('article');
  page.className = 'book-page cover-page';
  page.style.setProperty('--display-font', settings.displayFont);
  page.style.setProperty('--body-font', settings.bodyFont);
  page.style.setProperty('--type-scale', settings.typeScale);
  page.style.setProperty('--page-accent', settings.accent);

  const display = document.createElement('h2');
  display.className = 'display-line cover-title';
  if (state.titleEdits.has(-1)) {
    display.innerHTML = state.titleEdits.get(-1);
  } else {
    display.textContent = settings.title;
  }
  display.style.fontFamily = settings.displayFont;
  makeEditableText(display, -1, 'title');
  page.append(display);

  const text = document.createElement('p');
  text.className = 'cover-text';
  if (state.bodyEdits.has(-1)) {
    text.innerHTML = state.bodyEdits.get(-1);
  } else {
    text.textContent = 'subtitle / author / year';
  }
  makeEditableText(text, -1, 'body');
  page.append(text);

  return page;
}

function renderBackCoverPage(settings) {
  const page = document.createElement('article');
  page.className = 'book-page back-cover';
  page.style.setProperty('--page-accent', settings.accent);
  return page;
}

function renderPage(chunk, index, total, settings, rand) {
  const page = document.createElement('article');
  const recipe = layoutRecipe(index, rand, settings);
  page.className = 'book-page';
  page.dataset.layout = recipe.layout;
  page.style.setProperty('--display-font', settings.displayFont);
  page.style.setProperty('--type-scale', settings.typeScale);
  page.style.setProperty('--body-size', `${settings.bodySize}px`);
  page.style.setProperty('--body-leading', settings.leading);
  page.style.setProperty('--page-accent', settings.accent);

  const folio = document.createElement('span');
  folio.className = 'folio-mark';
  folio.textContent = `${settings.title} / ${String(index + 1).padStart(2, '0')}`;
  page.append(folio);

  const pageNumber = document.createElement('span');
  pageNumber.className = 'page-number';
  pageNumber.textContent = `${index + 1}/${total}`;
  page.append(pageNumber);

  if (state.images.length) {
    const image =
      state.images[(index + Math.floor(rand() * state.images.length)) % state.images.length];
    const figure = document.createElement('figure');
    figure.className = `image-block ${settings.monoImages ? 'mono' : ''} ${rand() > 0.42 ? 'line' : ''} ${recipe.rotate}`;
    placePercent(figure, recipe.image);
    applyImageEdit(figure, index, 'primary');

    const img = document.createElement('img');
    img.src = image.url;
    img.alt = image.name;
    figure.append(img);
    page.append(figure);
  }

  if (state.images.length > 1 && rand() > 0.45) {
    const second = state.images[(index + 1) % state.images.length];
    const figure = document.createElement('figure');
    figure.className = `image-block ${settings.monoImages ? 'mono' : ''}`;
    placePercent(figure, {
      left: clamp(8 + rand() * 72, 4, 82),
      top: clamp(8 + rand() * 62, 4, 78),
      width: 12 + rand() * 18,
      height: 12 + rand() * 22,
    });
    applyImageEdit(figure, index, 'secondary');
    figure.style.zIndex = rand() > 0.5 ? '1' : '4';
    const img = document.createElement('img');
    img.src = second.url;
    img.alt = second.name;
    figure.append(img);
    page.append(figure);
  }

  const display = document.createElement('h2');
  display.className = `display-line ${recipe.outline ? 'outline' : ''} ${recipe.blackTitle ? 'black' : ''}`;
  if (state.titleEdits.has(index)) {
    display.innerHTML = state.titleEdits.get(index);
  } else {
    display.textContent = index === 0 ? settings.title : phraseFromChunk(chunk);
  }
  display.style.left = `${recipe.display.left}%`;
  display.style.top = `${recipe.display.top}%`;
  display.style.fontFamily = settings.displayFont;
  makeEditableText(display, index, 'title');
  page.append(display);

  const textZone = document.createElement('div');
  textZone.className = `text-zone columns-${settings.columns} ${recipe.spaced ? 'spaced' : ''}`;
  textZone.style.fontFamily = settings.bodyFont;
  placePercent(textZone, recipe.text);

  const paragraph = document.createElement('p');
  if (state.bodyEdits.has(index)) {
    paragraph.innerHTML = state.bodyEdits.get(index);
  } else {
    paragraph.textContent = chunk;
  }
  makeEditableText(paragraph, index, 'body');
  textZone.append(paragraph);
  page.append(textZone);

  const caption = document.createElement('span');
  caption.className = 'caption';
  caption.textContent = pick(['index', 'source', 'fragment', 'plate', 'scan', 'note'], rand);
  caption.style.left = `${clamp(recipe.image.left + 1, 4, 88)}%`;
  caption.style.top = `${clamp(recipe.image.top + recipe.image.height + 1, 5, 92)}%`;
  page.append(caption);

  addRules(page, rand, settings.accent);

  return page;
}

function currentSettings() {
  const pageCount = clamp(Number(els.pageCount.value) || 4, 1, 24);
  const typeScale = Number(els.typeScale.value) || 1;
  const displayFont = cssFont(els.displayFont.value);
  const bodyFont = cssFont(els.bodyFont.value);
  return {
    pageCount,
    columns: state.columns,
    accent: els.accentColor.value,
    typeScale,
    imageEnergy: Number(els.imageEnergy.value) || 0,
    monoImages: els.monoImages.checked,
    displayFont,
    bodyFont,
    bodySize: state.columns === 2 ? 9.4 : 10.6,
    leading: state.columns === 2 ? 1.38 : 1.45,
    title: titleFromText(els.text.value),
  };
}

function render() {
  document.documentElement.style.setProperty('--accent', els.accentColor.value);
  const settings = currentSettings();
  const chunks = splitText(bodyFromText(els.text.value), settings.pageCount);
  els.pages.innerHTML = '';

  if (!chunks.length) {
    const empty = document.querySelector('#emptyPageTemplate').content.cloneNode(true);
    els.pages.append(empty);
    updateEditTextMode();
    return;
  }

  const rand = seededRandom(state.seed);
  const total = settings.pageCount;
  els.pages.append(renderCoverPage(settings));
  for (let index = 0; index < total; index += 1) {
    const chunk = chunks[index] || '';
    els.pages.append(renderPage(chunk, index, total, settings, rand));
  }
  els.pages.append(renderBackCoverPage(settings));
  updateEditTextMode();
}

function randomize() {
  const rand = seededRandom(Date.now() % 1000000);
  const fonts = allFonts();
  clearPageTextEdits();
  clearImageEdits();
  state.seed = Math.floor(rand() * 1000000);
  state.columns = rand() > 0.48 ? 2 : 1;
  els.displayFont.value = pick(fonts, rand).id;
  els.bodyFont.value = pick(fonts, rand).id;
  els.pageCount.value = String(pick([2, 4, 6, 8, 10, 12], rand));
  els.typeScale.value = String((0.82 + rand() * 0.46).toFixed(2));
  els.imageEnergy.value = String(Math.floor(35 + rand() * 64));
  els.accentColor.value = pick(palettes, rand);
  els.monoImages.checked = rand() > 0.28;
  els.layoutName.textContent = pick(layoutNames, rand);
  updateSegments();
  render();
}

function updateSegments() {
  els.segments.forEach((segment) => {
    segment.classList.toggle('is-active', Number(segment.dataset.columns) === state.columns);
  });
}

function fitPreview() {
  const stageWidth = document.querySelector('#previewStage').clientWidth - 56;
  const page = document.querySelector('.book-page');
  if (!page) return;
  const naturalWidth = page.getBoundingClientRect().width / state.zoom;
  state.zoom = clamp(stageWidth / naturalWidth, 0.52, 1);
  document.documentElement.style.setProperty('--page-scale', state.zoom);
}

els.generate.addEventListener('click', render);
els.randomize.addEventListener('click', randomize);
els.print.addEventListener('click', createBookletPdf);
els.systemFontBtn.addEventListener('click', loadSystemFonts);
els.editText.addEventListener('click', () => {
  state.editingText = !state.editingText;
  updateEditTextMode();
});
for (const control of [els.bold, els.italic, els.highlight, els.highlightColor]) {
  control.addEventListener('mousedown', saveTextSelection);
}
els.bold.addEventListener('click', () => runTextCommand('bold'));
els.italic.addEventListener('click', () => runTextCommand('italic'));
els.highlight.addEventListener('click', applyTextBackground);
els.highlightColor.addEventListener('input', applyTextBackground);
els.highlightColor.addEventListener('change', applyTextBackground);
els.fit.addEventListener('click', fitPreview);
els.zoomIn.addEventListener('click', () => {
  state.zoom = clamp(state.zoom + 0.1, 0.45, 1.4);
  document.documentElement.style.setProperty('--page-scale', state.zoom);
});
els.zoomOut.addEventListener('click', () => {
  state.zoom = clamp(state.zoom - 0.1, 0.45, 1.4);
  document.documentElement.style.setProperty('--page-scale', state.zoom);
});

els.segments.forEach((segment) => {
  segment.addEventListener('click', () => {
    state.columns = Number(segment.dataset.columns);
    updateSegments();
    render();
  });
});

els.text.addEventListener('input', () => {
  clearPageTextEdits();
  render();
});
els.text.addEventListener('change', () => {
  clearPageTextEdits();
  render();
});

for (const el of [
  els.pageCount,
  els.accentColor,
  els.typeScale,
  els.imageEnergy,
  els.monoImages,
  els.displayFont,
  els.bodyFont,
]) {
  el.addEventListener('input', render);
  el.addEventListener('change', render);
}

els.pages.addEventListener('input', (event) => {
  const editable = event.target.closest('.editable-text');
  if (!editable || !state.editingText) return;

  storeEditableOverride(editable);
  saveTextSelection();
});

els.pages.addEventListener('paste', (event) => {
  const editable = event.target.closest('.editable-text');
  if (!editable || !state.editingText) return;

  event.preventDefault();
  insertPlainText(event.clipboardData.getData('text/plain'));
  editable.dispatchEvent(new Event('input', { bubbles: true }));
});

els.pages.addEventListener('keyup', () => {
  saveTextSelection();
  updateFormatState();
});
els.pages.addEventListener('mouseup', () => {
  saveTextSelection();
  updateFormatState();
});
els.pages.addEventListener('pointerdown', startImageDrag);
window.addEventListener('pointermove', moveDraggedImage);
window.addEventListener('pointerup', stopImageDrag);
window.addEventListener('pointercancel', stopImageDrag);

els.imageInput.addEventListener('change', async (event) => {
  const images = await readFiles(event.target.files, (file, url) => ({
    name: file.name,
    url,
  }));
  state.images = images;
  clearImageEdits();
  els.imageCount.textContent = `${images.length} file${images.length === 1 ? '' : 's'}`;
  makeImageStrip();
  render();
});

els.fontInput.addEventListener('change', async (event) => {
  const fonts = await readFiles(event.target.files, addCustomFont);
  state.customFonts.push(...fonts);
  els.fontCount.textContent = `${state.customFonts.length} file${state.customFonts.length === 1 ? '' : 's'}`;
  refreshFontSelects();
  render();
});

window.addEventListener('resize', () => {
  if (state.zoom < 1) fitPreview();
});
window.addEventListener('afterprint', cleanupBookletPrint);
document.addEventListener('selectionchange', () => {
  saveTextSelection();
  updateFormatState();
});

refreshFontSelects();
updateSegments();
render();
