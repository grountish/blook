const state = {
  images: [],
  customFonts: [],
  systemFonts: [],
  columns: 2,
  seed: Math.floor(Math.random() * 100000),
  zoom: 1,
};

const builtInFonts = [
  { id: "builtin:futurist-mono", label: "Futurist mono", family: '"Courier New", "SFMono-Regular", monospace', role: "mono" },
  { id: "builtin:editorial-serif", label: "Editorial serif", family: 'Georgia, "Times New Roman", serif', role: "serif" },
  { id: "builtin:condensed-grotesk", label: "Condensed grotesk", family: '"Arial Narrow", "Avenir Next Condensed", Impact, sans-serif', role: "condensed" },
  { id: "builtin:swiss-grotesk", label: "Swiss grotesk", family: 'Helvetica, Arial, sans-serif', role: "sans" },
  { id: "builtin:typewriter", label: "Typewriter", family: '"American Typewriter", "Courier New", monospace', role: "type" },
  { id: "builtin:system-ui", label: "System UI", family: 'Inter, ui-sans-serif, system-ui, sans-serif', role: "sans" },
];

const layoutNames = [
  "Controlled chaos",
  "Circuit poetry",
  "Xerox grid",
  "Poster notes",
  "Off-register manual",
  "Small type panic",
];

const palettes = [
  "#ff2a1c",
  "#0077ff",
  "#1f9d55",
  "#d01b8c",
  "#f08b00",
  "#202020",
];

const els = {
  text: document.querySelector("#sourceText"),
  imageInput: document.querySelector("#imageInput"),
  fontInput: document.querySelector("#fontInput"),
  displayFont: document.querySelector("#displayFont"),
  bodyFont: document.querySelector("#bodyFont"),
  pageCount: document.querySelector("#pageCount"),
  accentColor: document.querySelector("#accentColor"),
  typeScale: document.querySelector("#typeScale"),
  imageEnergy: document.querySelector("#imageEnergy"),
  monoImages: document.querySelector("#monoImages"),
  imageCount: document.querySelector("#imageCount"),
  fontCount: document.querySelector("#fontCount"),
  systemFontBtn: document.querySelector("#systemFontBtn"),
  systemFontStatus: document.querySelector("#systemFontStatus"),
  generate: document.querySelector("#generateBtn"),
  randomize: document.querySelector("#randomizeBtn"),
  print: document.querySelector("#printBtn"),
  pages: document.querySelector("#pages"),
  imageStrip: document.querySelector("#imageStrip"),
  layoutName: document.querySelector("#layoutName"),
  fit: document.querySelector("#fitBtn"),
  zoomIn: document.querySelector("#zoomInBtn"),
  zoomOut: document.querySelector("#zoomOutBtn"),
  segments: [...document.querySelectorAll(".segment")],
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
  const clean = text.trim().replace(/\n{3,}/g, "\n\n");
  if (!clean) return [];

  const words = clean.split(/\s+/);
  const perPage = Math.max(1, Math.ceil(words.length / pageCount));
  const chunks = [];

  for (let i = 0; i < pageCount; i += 1) {
    const start = i * perPage;
    const end = i === pageCount - 1 ? words.length : (i + 1) * perPage;
    const chunk = words.slice(start, end).join(" ");
    if (chunk) chunks.push(chunk);
  }

  return chunks;
}

function titleFromText(text) {
  const firstLine = text.trim().split(/\n/).find(Boolean);
  if (!firstLine) return "Untitled Blook";
  return firstLine.slice(0, 74);
}

function bodyFromText(text) {
  const lines = text.trim().split(/\n/);
  const firstIndex = lines.findIndex((line) => line.trim());
  if (firstIndex === -1 || lines.filter((line) => line.trim()).length < 2) return text;
  lines.splice(firstIndex, 1);
  return lines.join("\n").trim();
}

function phraseFromChunk(chunk) {
  const words = chunk
    .replace(/[^\w\s-]/g, " ")
    .split(/\s+/)
    .filter((word) => word.length > 3);

  if (words.length < 2) return "Fanzine";
  const start = Math.max(0, Math.floor(words.length * 0.22));
  return words.slice(start, start + 3).join(" ");
}

function allFonts() {
  return [...builtInFonts, ...state.customFonts, ...state.systemFonts];
}

function quoteFontFamily(name) {
  return `"${name.replace(/\\/g, "\\\\").replace(/"/g, '\\"')}", sans-serif`;
}

function cssFont(fontId) {
  return allFonts().find((font) => font.id === fontId)?.family ?? builtInFonts[0].family;
}

function addFontGroup(select, label, fonts) {
  if (!fonts.length) return;
  const group = document.createElement("optgroup");
  group.label = label;
  fonts.forEach((font) => {
    const option = document.createElement("option");
    option.value = font.id;
    option.textContent = font.label;
    group.append(option);
  });
  select.append(group);
}

function refreshFontSelects() {
  for (const select of [els.displayFont, els.bodyFont]) {
    const previous = select.value;
    select.innerHTML = "";
    addFontGroup(select, "Built in", builtInFonts);
    addFontGroup(select, "Uploaded", state.customFonts);
    addFontGroup(select, "System", state.systemFonts);

    const fallback = select === els.displayFont ? "builtin:condensed-grotesk" : "builtin:futurist-mono";
    select.value = allFonts().some((font) => font.id === previous) ? previous : fallback;
  }
}

function makeImageStrip() {
  els.imageStrip.innerHTML = "";
  state.images.forEach((image) => {
    const img = document.createElement("img");
    img.className = "thumb";
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
  const family = `UserFont_${state.customFonts.length}_${file.name.replace(/\W+/g, "_")}`;
  const style = document.createElement("style");
  style.textContent = `@font-face{font-family:${family};src:url("${url}");font-display:swap;}`;
  document.head.append(style);
  return {
    id: `upload:${family}`,
    label: file.name.replace(/\.(ttf|otf|woff2?|TTF|OTF|WOFF2?)$/, ""),
    family,
    role: "custom",
  };
}

function timeoutAfter(ms) {
  return new Promise((_, reject) => {
    window.setTimeout(() => reject(new Error("Font access timed out")), ms);
  });
}

async function loadSystemFonts() {
  if (!("queryLocalFonts" in window)) {
    els.systemFontStatus.textContent = "Not supported in this browser";
    return;
  }

  if (!window.isSecureContext) {
    els.systemFontStatus.textContent = "Use localhost or HTTPS";
    return;
  }

  els.systemFontBtn.disabled = true;
  els.systemFontStatus.textContent = "Requesting access...";

  try {
    const fonts = await Promise.race([window.queryLocalFonts(), timeoutAfter(15000)]);
    const families = new Map();

    fonts.forEach((font) => {
      if (!font.family || families.has(font.family)) return;
      families.set(font.family, {
        id: `system:${font.family}`,
        label: font.family,
        family: quoteFontFamily(font.family),
        role: "system",
      });
    });

    state.systemFonts = [...families.values()].sort((a, b) => a.label.localeCompare(b.label));
    els.systemFontStatus.textContent = `${state.systemFonts.length} fonts loaded`;
    refreshFontSelects();
    render();
  } catch (error) {
    if (error.message === "Font access timed out") {
      els.systemFontStatus.textContent = "Permission prompt timed out";
    } else {
      els.systemFontStatus.textContent = error.name === "NotAllowedError" ? "Permission denied" : "Could not load fonts";
    }
  } finally {
    els.systemFontBtn.disabled = false;
  }
}

function layoutRecipe(index, rand, settings) {
  const types = ["schematic", "xerox", "poster", "manual"];
  const layout = pick(types, rand);
  const energy = settings.imageEnergy / 100;
  const width = 30 + rand() * 42 * energy;
  const height = 18 + rand() * 30 * energy;
  const imageTop = rand() > 0.5 ? 6 + rand() * 22 : 48 + rand() * 28;
  const imageLeft = rand() > 0.5 ? 5 + rand() * 18 : 42 + rand() * 26;
  const textWidth = settings.columns === 2 ? 43 + rand() * 16 : 28 + rand() * 24;
  const textHeight = 44 + rand() * 22;
  const textLeft = index % 2 === 0 ? 7 + rand() * 18 : 45 + rand() * 12;
  const textTop = rand() > 0.5 ? 10 + rand() * 16 : 35 + rand() * 20;
  const displayLeft = rand() > 0.5 ? 7 + rand() * 16 : 47 + rand() * 16;
  const displayTop = rand() > 0.5 ? 7 + rand() * 10 : 72 + rand() * 10;

  return {
    layout,
    text: {
      left: clamp(textLeft, 5, 67),
      top: clamp(textTop, 8, 70),
      width: clamp(textWidth, 24, 62),
      height: clamp(textHeight, 30, 72),
    },
    image: {
      left: clamp(imageLeft, 4, 74),
      top: clamp(imageTop, 4, 76),
      width: clamp(width, 22, 72),
      height: clamp(height, 16, 58),
    },
    display: {
      left: clamp(displayLeft, 5, 72),
      top: clamp(displayTop, 6, 82),
    },
    spaced: rand() > 0.38,
    outline: rand() > 0.72,
    blackTitle: rand() > 0.82,
    rotate: rand() > 0.5 ? "rotate-left" : "rotate-right",
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
    const rule = document.createElement("span");
    rule.className = `rule ${rand() > 0.78 ? "red" : ""}`;
    const horizontal = rand() > 0.34;
    rule.style.left = `${4 + rand() * 82}%`;
    rule.style.top = `${5 + rand() * 84}%`;
    rule.style.width = horizontal ? `${8 + rand() * 38}%` : "1px";
    rule.style.height = horizontal ? "1px" : `${6 + rand() * 28}%`;
    rule.style.backgroundColor = rule.classList.contains("red") ? accent : "";
    page.append(rule);
  }
}

function renderPage(chunk, index, total, settings, rand) {
  const page = document.createElement("article");
  const recipe = layoutRecipe(index, rand, settings);
  page.className = "book-page";
  page.dataset.layout = recipe.layout;
  page.style.setProperty("--display-font", settings.displayFont);
  page.style.setProperty("--type-scale", settings.typeScale);
  page.style.setProperty("--body-size", `${settings.bodySize}px`);
  page.style.setProperty("--body-leading", settings.leading);
  page.style.setProperty("--page-accent", settings.accent);

  const folio = document.createElement("span");
  folio.className = "folio-mark";
  folio.textContent = `${settings.title} / ${String(index + 1).padStart(2, "0")}`;
  page.append(folio);

  const pageNumber = document.createElement("span");
  pageNumber.className = "page-number";
  pageNumber.textContent = `${index + 1}/${total}`;
  page.append(pageNumber);

  if (state.images.length) {
    const image = state.images[(index + Math.floor(rand() * state.images.length)) % state.images.length];
    const figure = document.createElement("figure");
    figure.className = `image-block ${settings.monoImages ? "mono" : ""} ${rand() > 0.42 ? "line" : ""} ${recipe.rotate}`;
    placePercent(figure, recipe.image);

    const img = document.createElement("img");
    img.src = image.url;
    img.alt = image.name;
    figure.append(img);
    page.append(figure);
  }

  if (state.images.length > 1 && rand() > 0.45) {
    const second = state.images[(index + 1) % state.images.length];
    const figure = document.createElement("figure");
    figure.className = `image-block ${settings.monoImages ? "mono" : ""}`;
    placePercent(figure, {
      left: clamp(8 + rand() * 72, 4, 82),
      top: clamp(8 + rand() * 62, 4, 78),
      width: 12 + rand() * 18,
      height: 12 + rand() * 22,
    });
    figure.style.zIndex = rand() > 0.5 ? "1" : "4";
    const img = document.createElement("img");
    img.src = second.url;
    img.alt = second.name;
    figure.append(img);
    page.append(figure);
  }

  const display = document.createElement("h2");
  display.className = `display-line ${recipe.outline ? "outline" : ""} ${recipe.blackTitle ? "black" : ""}`;
  display.textContent = index === 0 ? settings.title : phraseFromChunk(chunk);
  display.style.left = `${recipe.display.left}%`;
  display.style.top = `${recipe.display.top}%`;
  display.style.fontFamily = settings.displayFont;
  page.append(display);

  const textZone = document.createElement("div");
  textZone.className = `text-zone columns-${settings.columns} ${recipe.spaced ? "spaced" : ""}`;
  textZone.style.fontFamily = settings.bodyFont;
  placePercent(textZone, recipe.text);

  const paragraph = document.createElement("p");
  paragraph.textContent = chunk;
  textZone.append(paragraph);
  page.append(textZone);

  const caption = document.createElement("span");
  caption.className = "caption";
  caption.textContent = pick(["index", "source", "fragment", "plate", "scan", "note"], rand);
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
    bodySize: state.columns === 2 ? 10.6 : 12.2,
    leading: state.columns === 2 ? 1.42 : 1.48,
    title: titleFromText(els.text.value),
  };
}

function render() {
  document.documentElement.style.setProperty("--accent", els.accentColor.value);
  const settings = currentSettings();
  const chunks = splitText(bodyFromText(els.text.value), settings.pageCount);
  els.pages.innerHTML = "";

  if (!chunks.length) {
    const empty = document.querySelector("#emptyPageTemplate").content.cloneNode(true);
    els.pages.append(empty);
    return;
  }

  const rand = seededRandom(state.seed);
  const total = settings.pageCount;
  for (let index = 0; index < total; index += 1) {
    const chunk = chunks[index] || "";
    els.pages.append(renderPage(chunk, index, total, settings, rand));
  }
}

function randomize() {
  const rand = seededRandom(Date.now() % 1000000);
  const fonts = allFonts();
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
    segment.classList.toggle("is-active", Number(segment.dataset.columns) === state.columns);
  });
}

function fitPreview() {
  const stageWidth = document.querySelector("#previewStage").clientWidth - 56;
  const page = document.querySelector(".book-page");
  if (!page) return;
  const naturalWidth = page.getBoundingClientRect().width / state.zoom;
  state.zoom = clamp(stageWidth / naturalWidth, 0.52, 1);
  document.documentElement.style.setProperty("--page-scale", state.zoom);
}

els.generate.addEventListener("click", render);
els.randomize.addEventListener("click", randomize);
els.print.addEventListener("click", () => window.print());
els.systemFontBtn.addEventListener("click", loadSystemFonts);
els.fit.addEventListener("click", fitPreview);
els.zoomIn.addEventListener("click", () => {
  state.zoom = clamp(state.zoom + 0.1, 0.45, 1.4);
  document.documentElement.style.setProperty("--page-scale", state.zoom);
});
els.zoomOut.addEventListener("click", () => {
  state.zoom = clamp(state.zoom - 0.1, 0.45, 1.4);
  document.documentElement.style.setProperty("--page-scale", state.zoom);
});

els.segments.forEach((segment) => {
  segment.addEventListener("click", () => {
    state.columns = Number(segment.dataset.columns);
    updateSegments();
    render();
  });
});

for (const el of [els.text, els.pageCount, els.accentColor, els.typeScale, els.imageEnergy, els.monoImages, els.displayFont, els.bodyFont]) {
  el.addEventListener("input", render);
  el.addEventListener("change", render);
}

els.imageInput.addEventListener("change", async (event) => {
  const images = await readFiles(event.target.files, (file, url) => ({
    name: file.name,
    url,
  }));
  state.images = images;
  els.imageCount.textContent = `${images.length} file${images.length === 1 ? "" : "s"}`;
  makeImageStrip();
  render();
});

els.fontInput.addEventListener("change", async (event) => {
  const fonts = await readFiles(event.target.files, addCustomFont);
  state.customFonts.push(...fonts);
  els.fontCount.textContent = `${state.customFonts.length} file${state.customFonts.length === 1 ? "" : "s"}`;
  refreshFontSelects();
  render();
});

window.addEventListener("resize", () => {
  if (state.zoom < 1) fitPreview();
});

refreshFontSelects();
updateSegments();
render();
