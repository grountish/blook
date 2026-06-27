const FAVORITE_FONTS_KEY = 'blook-favorite-fonts';
const TEXTURE_BACKGROUND_KEY = 'blook-apply-textures';
const TEXTURE_POOL_KEY = 'blook-texture-pool';
const AUTO_PAGES_KEY = 'blook-auto-pages';
const ASCII_CHARS = ' .,:;i1tfLCG08@';
// Upper bound on auto-fitted pages, also the manual ceiling.
const MAX_PAGES = 60;

const textureImages = [
  'textures/texture1.jpg',
  'textures/texture2.jpg',
  'textures/texture3.jpg',
  'textures/texture4.jpg',
  'textures/texture5.jpg',
];

const state = {
  images: [],
  asciiPages: [],

  systemFonts: [],
  favoriteFontPresets: readFavoriteFontPresets(),
  bodyEdits: new Map(),
  titleEdits: new Map(),
  imageEdits: new Map(),
  editingText: false,
  savedSelection: null,
  activeEditable: null,
  draggedImage: null,
  draggedZone: null,
  textZoneEdits: new Map(),
  columns: 2,
  imageTreatment: 'mono',
  seed: Math.floor(Math.random() * 100000),
  zoom: 1,
  activeFavoritePresetId: '',
  applyTextures: readTexturePreference(),
  texturePool: readTexturePool(),
  autoPages: readAutoPagesPreference(),
  wordDistribution: null,
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

// Curated Google Fonts catalog + pairings.
// Pairings are sourced from public typographic pairing guides (Fontpair.co and
// Google Fonts' own pairing recommendations) rather than random local fonts, so
// every display/body combination is a known-good aesthetic match.
// `weights` only lists axes that exist on each family — one invalid token would
// make Google's combined CSS request fail for the whole catalog.
const googleFontCatalog = [
  { name: 'Playfair Display', role: 'serif', stack: 'Georgia, serif', weights: '400;700' },
  { name: 'Source Sans 3', role: 'sans', stack: 'Helvetica, Arial, sans-serif', weights: '400;700' },
  { name: 'Archivo Black', role: 'sans', stack: '"Arial Black", sans-serif', weights: '400' },
  { name: 'Roboto', role: 'sans', stack: 'Arial, sans-serif', weights: '400;700' },
  { name: 'Oswald', role: 'sans', stack: '"Arial Narrow", sans-serif', weights: '400;700' },
  { name: 'Open Sans', role: 'sans', stack: 'Arial, sans-serif', weights: '400;700' },
  { name: 'Bebas Neue', role: 'sans', stack: 'Impact, sans-serif', weights: '400' },
  { name: 'Roboto Mono', role: 'mono', stack: '"Courier New", monospace', weights: '400;700' },
  { name: 'Abril Fatface', role: 'serif', stack: 'Georgia, serif', weights: '400' },
  { name: 'Lato', role: 'sans', stack: 'Helvetica, Arial, sans-serif', weights: '400;700' },
  { name: 'Montserrat', role: 'sans', stack: 'Helvetica, Arial, sans-serif', weights: '400;700' },
  { name: 'Merriweather', role: 'serif', stack: 'Georgia, serif', weights: '400;700' },
  { name: 'Anton', role: 'sans', stack: 'Impact, sans-serif', weights: '400' },
  { name: 'Space Grotesk', role: 'sans', stack: '"Helvetica Neue", sans-serif', weights: '400;700' },
  { name: 'Inter', role: 'sans', stack: 'system-ui, sans-serif', weights: '400;700' },
  { name: 'DM Serif Display', role: 'serif', stack: 'Georgia, serif', weights: '400' },
  { name: 'DM Sans', role: 'sans', stack: 'Helvetica, Arial, sans-serif', weights: '400;700' },
  { name: 'Syne', role: 'sans', stack: '"Helvetica Neue", sans-serif', weights: '400;700' },
  { name: 'Libre Baskerville', role: 'serif', stack: 'Georgia, serif', weights: '400;700' },
  { name: 'Libre Franklin', role: 'sans', stack: 'Helvetica, Arial, sans-serif', weights: '400;700' },
  { name: 'Cormorant Garamond', role: 'serif', stack: 'Garamond, Georgia, serif', weights: '400;700' },
  { name: 'Proza Libre', role: 'sans', stack: 'Helvetica, Arial, sans-serif', weights: '400;700' },
  { name: 'Fraunces', role: 'serif', stack: 'Georgia, serif', weights: '400;700' },
  { name: 'Manrope', role: 'sans', stack: 'Helvetica, Arial, sans-serif', weights: '400;700' },
  { name: 'IBM Plex Sans', role: 'sans', stack: 'Helvetica, Arial, sans-serif', weights: '400;700' },
  { name: 'IBM Plex Mono', role: 'mono', stack: '"Courier New", monospace', weights: '400;700' },
  { name: 'Spectral', role: 'serif', stack: 'Georgia, serif', weights: '400;700' },
  { name: 'Karla', role: 'sans', stack: 'Helvetica, Arial, sans-serif', weights: '400;700' },
  // Experimental / characterful faces — terminal, glitch, brutalist, blackletter.
  { name: 'Unbounded', role: 'display', stack: '"Arial Black", sans-serif', weights: '400;700' },
  { name: 'Bricolage Grotesque', role: 'sans', stack: 'Helvetica, Arial, sans-serif', weights: '400;700' },
  { name: 'Big Shoulders Display', role: 'display', stack: '"Arial Narrow", sans-serif', weights: '400;700' },
  { name: 'Space Mono', role: 'mono', stack: '"Courier New", monospace', weights: '400;700' },
  { name: 'VT323', role: 'mono', stack: '"Courier New", monospace', weights: '400' },
  { name: 'Major Mono Display', role: 'mono', stack: '"Courier New", monospace', weights: '400' },
  { name: 'Monoton', role: 'display', stack: '"Arial Black", sans-serif', weights: '400' },
  { name: 'Bungee', role: 'display', stack: 'Impact, sans-serif', weights: '400' },
  { name: 'Rubik Mono One', role: 'display', stack: '"Arial Black", sans-serif', weights: '400' },
  { name: 'Rubik Glitch', role: 'display', stack: 'Impact, sans-serif', weights: '400' },
  { name: 'Silkscreen', role: 'display', stack: 'monospace', weights: '400;700' },
  { name: 'Pirata One', role: 'display', stack: 'Georgia, serif', weights: '400' },
  { name: 'Instrument Serif', role: 'serif', stack: 'Georgia, serif', weights: '400' },
  { name: 'Gloock', role: 'serif', stack: 'Georgia, serif', weights: '400' },
  { name: 'Redacted Script', role: 'display', stack: 'cursive', weights: '400;700' },
];

const googleFonts = googleFontCatalog.map((font) => ({
  id: `google:${font.name}`,
  label: font.name,
  family: `"${font.name}", ${font.stack}`,
  role: font.role,
}));

// Display / body pairs. `display`/`body` reference catalog `name` values.
// `classic` = safe, conventional matches; `experimental` = deliberately bold,
// clashing, or characterful — fitting a fanzine's punk/photocopy aesthetic.
const classicPairings = [
  { vibe: 'Editorial elegance', display: 'Playfair Display', body: 'Source Sans 3' },
  { vibe: 'Bold & modern', display: 'Archivo Black', body: 'Roboto' },
  { vibe: 'Clean condensed', display: 'Oswald', body: 'Open Sans' },
  { vibe: 'Poster + machine', display: 'Bebas Neue', body: 'Roboto Mono' },
  { vibe: 'Fashion serif', display: 'Abril Fatface', body: 'Lato' },
  { vibe: 'Classic contrast', display: 'Montserrat', body: 'Merriweather' },
  { vibe: 'Impact headline', display: 'Anton', body: 'Roboto' },
  { vibe: 'Contemporary geometric', display: 'Space Grotesk', body: 'Inter' },
  { vibe: 'High-contrast serif', display: 'DM Serif Display', body: 'DM Sans' },
  { vibe: 'Art-house brutal', display: 'Syne', body: 'Inter' },
  { vibe: 'Literary classic', display: 'Libre Baskerville', body: 'Libre Franklin' },
  { vibe: 'Refined editorial', display: 'Cormorant Garamond', body: 'Proza Libre' },
  { vibe: 'Warm modern serif', display: 'Fraunces', body: 'Manrope' },
  { vibe: 'Technical systems', display: 'IBM Plex Sans', body: 'IBM Plex Mono' },
  { vibe: 'Journalistic', display: 'Spectral', body: 'Karla' },
];

const experimentalPairings = [
  { vibe: 'Neon glitch', display: 'Monoton', body: 'Space Mono' },
  { vibe: 'Sign painter', display: 'Bungee', body: 'Inter' },
  { vibe: 'Brutalist block', display: 'Rubik Mono One', body: 'Roboto Mono' },
  { vibe: 'Maximalist shout', display: 'Unbounded', body: 'IBM Plex Mono' },
  { vibe: 'Blackletter punk', display: 'Pirata One', body: 'Karla' },
  { vibe: 'Lo-fi terminal', display: 'Major Mono Display', body: 'Space Mono' },
  { vibe: 'CRT boot screen', display: 'VT323', body: 'Space Mono' },
  { vibe: 'Industrial slab', display: 'Big Shoulders Display', body: 'DM Sans' },
  { vibe: 'Editorial brutalism', display: 'Bricolage Grotesque', body: 'Spectral' },
  { vibe: 'Off-kilter editorial', display: 'Instrument Serif', body: 'Space Grotesk' },
  { vibe: 'Heavy serif drama', display: 'Gloock', body: 'Manrope' },
  { vibe: 'Corrupted signal', display: 'Rubik Glitch', body: 'Roboto Mono' },
  { vibe: 'Pixel zine', display: 'Silkscreen', body: 'Roboto Mono' },
  { vibe: 'Censored draft', display: 'Redacted Script', body: 'Karla' },
  { vibe: 'Headline + machine', display: 'Anton', body: 'Space Mono' },
];

const decoratePairing = (tier) => (pairing) => ({
  ...pairing,
  tier,
  id: `pair:${pairing.display}|${pairing.body}`,
  displayId: `google:${pairing.display}`,
  bodyId: `google:${pairing.body}`,
  label: `${pairing.display} / ${pairing.body}`,
});

const fontPairings = [
  ...classicPairings.map(decoratePairing('classic')),
  ...experimentalPairings.map(decoratePairing('experimental')),
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
  keywords: document.querySelector('#sourceKeywords'),
  imageInput: document.querySelector('#imageInput'),
  imageSearchInput: document.querySelector('#imageSearchInput'),
  imageSearchBtn: document.querySelector('#imageSearchBtn'),
  imageSearchStatus: document.querySelector('#imageSearchStatus'),
  asciiInput: document.querySelector('#asciiInput'),
  asciiChars: document.querySelector('#asciiChars'),
  asciiColumns: document.querySelector('#asciiColumns'),
  asciiInsertAfter: document.querySelector('#asciiInsertAfter'),
  displayFont: document.querySelector('#displayFont'),
  bodyFont: document.querySelector('#bodyFont'),
  displayFontTrigger: document.querySelector('#displayFontTrigger'),
  displayFontMenu: document.querySelector('#displayFontMenu'),
  bodyFontTrigger: document.querySelector('#bodyFontTrigger'),
  bodyFontMenu: document.querySelector('#bodyFontMenu'),
  pageCount: document.querySelector('#pageCount'),
  accentColor: document.querySelector('#accentColor'),
  typeScale: document.querySelector('#typeScale'),
  typeSpacing: document.querySelector('#typeSpacing'),
  imageEnergy: document.querySelector('#imageEnergy'),
  imageBlend: document.querySelector('#imageBlend'),
  grit: document.querySelector('#grit'),
  showDots: document.querySelector('#showDots'),
  applyTextures: document.querySelector('#applyTextures'),
  autoPages: document.querySelector('#autoPages'),
  imageCount: document.querySelector('#imageCount'),
  asciiStatus: document.querySelector('#asciiStatus'),
  asciiList: document.querySelector('#asciiList'),
  textureStrip: document.querySelector('#textureStrip'),
  systemFontBtn: document.querySelector('#systemFontBtn'),
  systemFontStatus: document.querySelector('#systemFontStatus'),
  favoriteFontBtn: document.querySelector('#favoriteFontBtn'),
  favoriteFontStatus: document.querySelector('#favoriteFontStatus'),
  favoriteFontsSelect: document.querySelector('#favoriteFontsSelect'),
  removeFavoriteFontBtn: document.querySelector('#removeFavoriteFontBtn'),
  fontPairingSelect: document.querySelector('#fontPairingSelect'),
  shufflePairingBtn: document.querySelector('#shufflePairingBtn'),
  generate: document.querySelector('#generateBtn'),
  save: document.querySelector('#saveBtn'),
  library: document.querySelector('#libraryBtn'),
  closeLibrary: document.querySelector('#closeLibraryBtn'),
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
  segments: [...document.querySelectorAll('#columnSegment .segment')],
  treatmentSegments: [...document.querySelectorAll('#imageTreatment .segment')],
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

function readTexturePreference() {
  try {
    return window.localStorage.getItem(TEXTURE_BACKGROUND_KEY) === 'true';
  } catch {
    return false;
  }
}

function writeTexturePreference(enabled) {
  try {
    window.localStorage.setItem(TEXTURE_BACKGROUND_KEY, String(Boolean(enabled)));
  } catch {}
}

function readAutoPagesPreference() {
  try {
    return window.localStorage.getItem(AUTO_PAGES_KEY) === 'true';
  } catch {
    return false;
  }
}

function writeAutoPagesPreference(enabled) {
  try {
    window.localStorage.setItem(AUTO_PAGES_KEY, String(Boolean(enabled)));
  } catch {}
}

function debounce(fn, wait) {
  let timer;
  return (...args) => {
    window.clearTimeout(timer);
    timer = window.setTimeout(() => fn(...args), wait);
  };
}

function readTexturePool() {
  try {
    const raw = window.localStorage.getItem(TEXTURE_POOL_KEY);
    if (!raw) return textureImages.map((texture) => ({ src: texture, active: true }));

    const parsed = JSON.parse(raw);
    const activeSet = new Set(Array.isArray(parsed) ? parsed : []);
    return textureImages.map((texture) => ({
      src: texture,
      active: activeSet.has(texture),
    }));
  } catch {
    return textureImages.map((texture) => ({ src: texture, active: true }));
  }
}

function writeTexturePool() {
  try {
    const activeTextures = state.texturePool.filter((item) => item.active).map((item) => item.src);
    window.localStorage.setItem(TEXTURE_POOL_KEY, JSON.stringify(activeTextures));
  } catch {}
}

function textWords(text) {
  const clean = text.trim().replace(/\n{3,}/g, '\n\n');
  return clean ? clean.split(/\s+/) : [];
}

// Break body text into paragraphs. Prefer blank-line separation; fall back to
// single newlines when the text has no blank lines. Each paragraph's internal
// whitespace is collapsed so it reflows cleanly inside a page.
function splitParagraphs(text) {
  const clean = text.replace(/\r\n/g, '\n').trim();
  if (!clean) return [];
  let paras = clean.split(/\n{2,}/);
  if (paras.length <= 1) paras = clean.split(/\n+/);
  return paras.map((p) => p.replace(/\s+/g, ' ').trim()).filter(Boolean);
}

function wordCount(str) {
  const trimmed = str.trim();
  return trimmed ? trimmed.split(/\s+/).length : 0;
}

// Pack whole paragraphs into pages so every page break lands on a paragraph
// boundary (a page never starts mid-paragraph). `targets` are relative per-page
// weights (capacity). Each page aims for its share of the *remaining* words, so
// chunky paragraphs early on don't starve later pages into blanks; we also keep
// at least one paragraph in reserve per remaining page. Within a page we snap to
// the nearest paragraph boundary around the target. The last page takes the rest.
function packParagraphs(paragraphs, targets) {
  if (!paragraphs.length) return [];
  const pages = [];
  let idx = 0;
  let remainingWords = paragraphs.reduce((sum, para) => sum + wordCount(para), 0);
  let remainingWeight = targets.reduce((sum, weight) => sum + (weight || 0), 0) || targets.length;
  for (let p = 0; p < targets.length; p += 1) {
    const pagesLeft = targets.length - p;
    const weight = targets[p] || remainingWeight / pagesLeft;
    if (idx >= paragraphs.length) {
      pages.push('');
      remainingWeight -= weight;
      continue;
    }
    if (p === targets.length - 1) {
      pages.push(paragraphs.slice(idx).join('\n\n'));
      idx = paragraphs.length;
      continue;
    }
    const target = remainingWords * (weight / (remainingWeight || 1));
    // Reserve one paragraph for each later page so none ends up empty.
    const maxTake = Math.max(1, paragraphs.length - idx - (pagesLeft - 1));
    const start = idx;
    let words = 0;
    do {
      words += wordCount(paragraphs[idx]);
      idx += 1;
    } while (
      idx - start < maxTake &&
      idx < paragraphs.length &&
      target - words > wordCount(paragraphs[idx]) / 2
    );
    pages.push(paragraphs.slice(start, idx).join('\n\n'));
    remainingWords -= words;
    remainingWeight -= weight;
  }
  // Any paragraphs left over (ran past the last target) join the final page.
  if (idx < paragraphs.length) {
    const tail = paragraphs.slice(idx).join('\n\n');
    pages[pages.length - 1] = [pages[pages.length - 1], tail].filter(Boolean).join('\n\n');
  }
  return pages;
}

function splitText(text, pageCount) {
  const paragraphs = splitParagraphs(text);
  if (!paragraphs.length) return [];
  const total = paragraphs.reduce((sum, para) => sum + wordCount(para), 0);
  const targets = Array.from({ length: pageCount }, () => total / pageCount);
  return packParagraphs(paragraphs, targets);
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

function googleFontsHref() {
  const families = googleFontCatalog
    .map((font) => `family=${encodeURIComponent(font.name).replace(/%20/g, '+')}:wght@${font.weights}`)
    .join('&');
  return `https://fonts.googleapis.com/css2?${families}&display=swap`;
}

// Inject the combined stylesheet for the whole curated catalog once. Declaring
// every @font-face is cheap; the browser only downloads a family's files when a
// glyph is actually rendered in it (e.g. in the picker preview or on a page).
function loadGoogleFontCatalog() {
  if (document.querySelector('#google-fonts-link')) return;

  const preconnect = document.createElement('link');
  preconnect.rel = 'preconnect';
  preconnect.href = 'https://fonts.googleapis.com';

  const preconnectStatic = document.createElement('link');
  preconnectStatic.rel = 'preconnect';
  preconnectStatic.href = 'https://fonts.gstatic.com';
  preconnectStatic.crossOrigin = 'anonymous';

  const stylesheet = document.createElement('link');
  stylesheet.id = 'google-fonts-link';
  stylesheet.rel = 'stylesheet';
  stylesheet.href = googleFontsHref();

  document.head.append(preconnect, preconnectStatic, stylesheet);
}

function allFonts() {
  const fonts = [...favoriteFontPool(), ...builtInFonts, ...googleFonts, ...state.systemFonts];
  const unique = new Map();
  fonts.forEach((font) => {
    if (!font?.id || unique.has(font.id)) return;
    unique.set(font.id, font);
  });
  return [...unique.values()];
}

function normalizeFontRecord(font) {
  if (!font?.id || !font?.label || !font?.family) return null;
  return {
    id: font.id,
    label: font.label,
    family: font.family,
    role: font.role || 'system',
  };
}

function normalizeFavoritePreset(preset) {
  if (!preset || typeof preset !== 'object') return null;

  const display = normalizeFontRecord(preset.display);
  const body = normalizeFontRecord(preset.body);
  if (!display && !body) return null;

  return {
    id: preset.id || `${display?.id || 'none'}|${body?.id || 'none'}`,
    display,
    body,
    savedAt: preset.savedAt || new Date().toISOString(),
  };
}

function readFavoriteFontPresets() {
  try {
    const raw = window.localStorage.getItem(FAVORITE_FONTS_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) {
      return parsed.map(normalizeFavoritePreset).filter(Boolean);
    }

    if (Array.isArray(parsed?.items)) {
      return parsed.items.map(normalizeFavoritePreset).filter(Boolean);
    }

    const legacy = normalizeFavoritePreset(parsed);
    return legacy ? [legacy] : [];
  } catch {
    return [];
  }
}

function writeFavoriteFontPresets(favoriteFontPresets) {
  try {
    window.localStorage.setItem(FAVORITE_FONTS_KEY, JSON.stringify(favoriteFontPresets));
  } catch {}
}

function favoriteFontPool() {
  const fonts = [];
  state.favoriteFontPresets.forEach((preset) => {
    if (preset.display) fonts.push(preset.display);
    if (preset.body) fonts.push(preset.body);
  });

  const unique = new Map();
  fonts.forEach((font) => {
    if (!font?.id || unique.has(font.id)) return;
    unique.set(font.id, font);
  });
  return [...unique.values()];
}

function filterFontsById(fonts, excludedIds = new Set()) {
  return fonts.filter((font) => font?.id && !excludedIds.has(font.id));
}

function favoriteFontSummary() {
  const count = state.favoriteFontPresets.length;
  if (!count) return 'No favorites saved';
  return `${count} favorite font set${count === 1 ? '' : 's'} saved`;
}

function updateFavoriteFontStatus(message = null) {
  els.favoriteFontStatus.textContent = message || favoriteFontSummary();
}

function updateTextureMode() {
  state.applyTextures = Boolean(els.applyTextures.checked);
  document.body.classList.toggle('has-textures', state.applyTextures);
  writeTexturePreference(state.applyTextures);
}

function activeTextures() {
  return state.texturePool.filter((item) => item.active).map((item) => item.src);
}

function favoritePresetLabel(preset) {
  const display = preset.display?.label || 'Display font';
  const body = preset.body?.label || 'Text font';
  return `${display} / ${body}`;
}

function textureForKey(key) {
  const textures = activeTextures();
  if (!textures.length) return null;

  let hash = 2166136261;
  const input = `${state.seed}:${key}`;
  for (let i = 0; i < input.length; i += 1) {
    hash ^= input.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }
  return textures[Math.abs(hash) % textures.length];
}

function applyTextureToPage(page, key) {
  const texture = textureForKey(key);
  page.style.setProperty('--page-texture-image', texture ? `url("${texture}")` : 'none');
}

function renderTextureStrip() {
  const strip = els.textureStrip;
  strip.innerHTML = '';

  state.texturePool.forEach((texture, index) => {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = `thumb texture-thumb ${texture.active ? 'is-active' : ''}`;
    button.setAttribute('aria-pressed', String(texture.active));
    button.title = texture.src.split('/').pop();
    button.style.backgroundImage = `url("${texture.src}")`;
    button.addEventListener('click', () => {
      state.texturePool[index].active = !state.texturePool[index].active;
      writeTexturePool();
      renderTextureStrip();
      render();
    });
    strip.append(button);
  });
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
    option.style.fontFamily = font.family;
    group.append(option);
  });
  select.append(group);
}

function fontGroups() {
  const favoriteFonts = favoriteFontPool();
  const usedIds = new Set(favoriteFonts.map((font) => font.id));
  const builtIn = filterFontsById(builtInFonts, usedIds);
  builtIn.forEach((font) => usedIds.add(font.id));
  const google = filterFontsById(googleFonts, usedIds);
  google.forEach((font) => usedIds.add(font.id));
  const system = filterFontsById(state.systemFonts, usedIds);

  return [
    ['Favorites', favoriteFonts],
    ['Google pairings', google],
    ['Built in', builtIn],
    ['System', system],
  ];
}

function syncFontSelectPreview(select) {
  const font = allFonts().find((item) => item.id === select.value);
  select.style.fontFamily = font?.family || '';
  const trigger = select === els.displayFont ? els.displayFontTrigger : els.bodyFontTrigger;
  if (trigger) {
    trigger.textContent = font?.label || 'Choose font';
    trigger.style.fontFamily = font?.family || '';
  }
}

function closeFontPicker(kind) {
  const menu = kind === 'display' ? els.displayFontMenu : els.bodyFontMenu;
  const trigger = kind === 'display' ? els.displayFontTrigger : els.bodyFontTrigger;
  menu.hidden = true;
  trigger.setAttribute('aria-expanded', 'false');
}

function closeAllFontPickers() {
  closeFontPicker('display');
  closeFontPicker('body');
}

function openFontPicker(kind) {
  const menu = kind === 'display' ? els.displayFontMenu : els.bodyFontMenu;
  const trigger = kind === 'display' ? els.displayFontTrigger : els.bodyFontTrigger;
  const isOpen = !menu.hidden;
  closeAllFontPickers();
  if (isOpen) return;
  menu.hidden = false;
  trigger.setAttribute('aria-expanded', 'true');
}

function buildFontPickerMenu(menu, select) {
  menu.innerHTML = '';
  const currentValue = select.value;

  fontGroups().forEach(([label, fonts]) => {
    if (!fonts.length) return;

    const group = document.createElement('div');
    group.className = 'font-picker-group';

    const heading = document.createElement('div');
    heading.className = 'font-picker-group-label';
    heading.textContent = label;
    group.append(heading);

    fonts.forEach((font) => {
      const option = document.createElement('button');
      option.type = 'button';
      option.className = 'font-picker-option';
      option.dataset.fontId = font.id;
      option.textContent = font.label;
      option.style.fontFamily = font.family;
      option.setAttribute('role', 'option');
      option.setAttribute('aria-selected', String(font.id === currentValue));
      option.classList.toggle('is-selected', font.id === currentValue);
      option.addEventListener('click', () => {
        select.value = font.id;
        syncFontSelectPreview(select);
        closeAllFontPickers();
        render();
      });
      group.append(option);
    });

    menu.append(group);
  });

  if (!menu.children.length) {
    const empty = document.createElement('div');
    empty.className = 'font-picker-empty';
    empty.textContent = 'No fonts available';
    menu.append(empty);
  }
}

function setFontPickerValue(select, value) {
  select.value = value;
  syncFontSelectPreview(select);
}

function refreshFontSelects() {
  for (const select of [els.displayFont, els.bodyFont]) {
    const previous = select.value;
    select.innerHTML = '';
    fontGroups().forEach(([label, fonts]) => addFontGroup(select, label, fonts));

    const fallback =
      select === els.displayFont ? 'builtin:condensed-grotesk' : 'builtin:futurist-mono';
    select.value = allFonts().some((font) => font.id === previous) ? previous : fallback;
    syncFontSelectPreview(select);
  }

  buildFontPickerMenu(els.displayFontMenu, els.displayFont);
  buildFontPickerMenu(els.bodyFontMenu, els.bodyFont);
}

function renderFavoriteFontSelect() {
  const select = els.favoriteFontsSelect;
  const presets = state.favoriteFontPresets;
  select.innerHTML = '';

  if (!presets.length) {
    const option = document.createElement('option');
    option.value = '';
    option.textContent = 'No favorites saved';
    select.append(option);
    select.disabled = true;
    els.removeFavoriteFontBtn.disabled = true;
    state.activeFavoritePresetId = '';
    return;
  }

  select.disabled = false;
  els.removeFavoriteFontBtn.disabled = false;

  presets.forEach((preset) => {
    const option = document.createElement('option');
    option.value = preset.id;
    option.textContent = favoritePresetLabel(preset);
    select.append(option);
  });

  const activeId =
    presets.some((preset) => preset.id === state.activeFavoritePresetId)
      ? state.activeFavoritePresetId
      : presets[0].id;
  state.activeFavoritePresetId = activeId;
  select.value = activeId;
}

function saveFavoriteFonts() {
  const display = normalizeFontRecord(allFonts().find((font) => font.id === els.displayFont.value));
  const body = normalizeFontRecord(allFonts().find((font) => font.id === els.bodyFont.value));
  if (!display && !body) return;

  const preset = normalizeFavoritePreset({ display, body });
  if (!preset) return;

  state.favoriteFontPresets = [
    preset,
    ...state.favoriteFontPresets.filter((item) => item.id !== preset.id),
  ];
  state.activeFavoritePresetId = preset.id;
  writeFavoriteFontPresets(state.favoriteFontPresets);
  refreshFontSelects();
  renderFavoriteFontSelect();
  updateFavoriteFontStatus('Favorite saved');
  window.setTimeout(() => updateFavoriteFontStatus(), 1400);
}

function applyFavoriteFontPresetById(id) {
  const preset = state.favoriteFontPresets.find((item) => item.id === id);
  if (!preset) return;

  state.activeFavoritePresetId = id;

  if (preset.display && allFonts().some((font) => font.id === preset.display.id)) {
    setFontPickerValue(els.displayFont, preset.display.id);
  }

  if (preset.body && allFonts().some((font) => font.id === preset.body.id)) {
    setFontPickerValue(els.bodyFont, preset.body.id);
  }

  render();
}

function removeFavoriteFontPreset(id = state.activeFavoritePresetId) {
  if (!id) return;

  state.favoriteFontPresets = state.favoriteFontPresets.filter((preset) => preset.id !== id);
  state.activeFavoritePresetId = state.favoriteFontPresets[0]?.id || '';
  writeFavoriteFontPresets(state.favoriteFontPresets);
  refreshFontSelects();
  renderFavoriteFontSelect();
  updateFavoriteFontStatus();
}

function renderPairingSelect() {
  const select = els.fontPairingSelect;
  if (!select) return;
  select.innerHTML = '';

  const placeholder = document.createElement('option');
  placeholder.value = '';
  placeholder.textContent = 'Curated pairing…';
  select.append(placeholder);

  const tiers = [
    ['Classic', fontPairings.filter((pairing) => pairing.tier === 'classic')],
    ['Experimental', fontPairings.filter((pairing) => pairing.tier === 'experimental')],
  ];

  tiers.forEach(([label, pairings]) => {
    if (!pairings.length) return;
    const group = document.createElement('optgroup');
    group.label = label;
    pairings.forEach((pairing) => {
      const option = document.createElement('option');
      option.value = pairing.id;
      option.textContent = `${pairing.vibe} · ${pairing.label}`;
      group.append(option);
    });
    select.append(group);
  });
}

function applyPairingById(id) {
  const pairing = fontPairings.find((item) => item.id === id);
  if (!pairing) return;

  setFontPickerValue(els.displayFont, pairing.displayId);
  setFontPickerValue(els.bodyFont, pairing.bodyId);
  if (els.fontPairingSelect) els.fontPairingSelect.value = pairing.id;
  render();
}

function selectedImages() {
  return state.images.filter((image) => image.selected !== false);
}

function updateImageCount() {
  const total = state.images.length;
  if (!total) {
    els.imageCount.textContent = '0 files';
    return;
  }
  const used = selectedImages().length;
  els.imageCount.textContent = `${used}/${total} selected`;
}

function makeImageStrip() {
  els.imageStrip.innerHTML = '';
  state.images.forEach((image, index) => {
    const img = document.createElement('img');
    img.className = classes('thumb', image.selected === false && 'is-deselected');
    img.src = image.url;
    img.alt = image.name;
    img.dataset.index = String(index);
    img.title = `${image.name} — click to ${image.selected === false ? 'use' : 'skip'}`;
    els.imageStrip.append(img);
  });
}

// One delegated listener on the strip — survives every re-render and can't get
// detached from individual thumbnails mid-interaction.
els.imageStrip.addEventListener('click', (event) => {
  const thumb = event.target.closest('.thumb');
  if (!thumb || !els.imageStrip.contains(thumb)) return;
  const image = state.images[Number(thumb.dataset.index)];
  if (!image) return;
  image.selected = image.selected === false;
  thumb.classList.toggle('is-deselected', image.selected === false);
  thumb.title = `${image.name} — click to ${image.selected === false ? 'use' : 'skip'}`;
  updateImageCount();
  render();
});

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

function ensureAsciiCharSet(value) {
  const clean = String(value || '')
    .replace(/[\r\n\t]/g, '')
    .replace(/\s{2,}/g, ' ')
    .slice(0, 80);
  return clean.length >= 2 ? clean : ASCII_CHARS;
}

function currentTextPageTotal() {
  const pageCount = clamp(Number(els.pageCount.value) || 4, 1, MAX_PAGES);
  const hasBodyText = bodyFromText(els.text.value).trim().length > 0;
  return hasBodyText ? pageCount : 0;
}

function normalizeAsciiInsertAfter(insertAfter, textPageTotal = currentTextPageTotal()) {
  return clamp(Number(insertAfter) || 0, 0, textPageTotal);
}

function ensureAsciiRecord(record, textPageTotal = currentTextPageTotal()) {
  if (!record || typeof record.text !== 'string') return null;
  return {
    id: record.id || `${Date.now()}-${Math.floor(Math.random() * 1e6)}`,
    name: record.name || 'ASCII page',
    text: record.text,
    sourceUrl: record.sourceUrl || '',
    columns: clamp(Number(record.columns) || 100, 40, 140),
    charSet: ensureAsciiCharSet(record.charSet),
    insertAfter: normalizeAsciiInsertAfter(record.insertAfter, textPageTotal),
  };
}

function updateAsciiInsertOptions() {
  if (!els.asciiInsertAfter) return;
  const currentValue = Number(els.asciiInsertAfter.value || 0);
  const textPageTotal = currentTextPageTotal();
  els.asciiInsertAfter.innerHTML = '';

  for (let after = 0; after <= textPageTotal; after += 1) {
    const option = document.createElement('option');
    option.value = String(after);
    if (!textPageTotal) {
      option.textContent = 'Only page slot';
    } else if (after === 0) {
      option.textContent = 'Before page 1';
    } else if (after === textPageTotal) {
      option.textContent = `After page ${textPageTotal}`;
    } else {
      option.textContent = `After page ${after}`;
    }
    els.asciiInsertAfter.append(option);
  }

  els.asciiInsertAfter.value = String(normalizeAsciiInsertAfter(currentValue, textPageTotal));
}

function updateAsciiStatus() {
  if (!els.asciiStatus) return;
  if (!state.asciiPages.length) {
    els.asciiStatus.textContent = 'No image';
    return;
  }

  els.asciiStatus.textContent = `${state.asciiPages.length} page${state.asciiPages.length === 1 ? '' : 's'} ready`;
}

function renderAsciiList() {
  if (!els.asciiList) return;
  const textPageTotal = currentTextPageTotal();
  els.asciiList.innerHTML = '';

  if (!state.asciiPages.length) return;

  state.asciiPages.forEach((asciiPage) => {
    const row = document.createElement('div');
    row.className = 'ascii-item';

    const label = document.createElement('small');
    label.className = 'ascii-item-label';
    label.textContent = asciiPage.name || 'ASCII page';
    row.append(label);

    const sizeInput = document.createElement('input');
    sizeInput.type = 'number';
    sizeInput.min = '40';
    sizeInput.max = '140';
    sizeInput.step = '2';
    sizeInput.className = 'ascii-item-size';
    sizeInput.title = 'ASCII width (columns)';
    sizeInput.setAttribute('aria-label', 'ASCII width');
    sizeInput.placeholder = 'Width';
    sizeInput.value = String(asciiPage.columns || 100);
    row.append(sizeInput);

    const charsInput = document.createElement('input');
    charsInput.type = 'text';
    charsInput.className = 'ascii-item-chars';
    charsInput.title = 'Character set (dark to light)';
    charsInput.setAttribute('aria-label', 'ASCII characters');
    charsInput.placeholder = 'Character set';
    charsInput.value = ensureAsciiCharSet(asciiPage.charSet);
    charsInput.spellcheck = false;
    row.append(charsInput);

    const select = document.createElement('select');
    select.className = 'ascii-item-select';
    select.title = 'Insert position';
    select.setAttribute('aria-label', 'Insert position');
    for (let after = 0; after <= textPageTotal; after += 1) {
      const option = document.createElement('option');
      option.value = String(after);
      if (!textPageTotal) {
        option.textContent = 'Only page slot';
      } else if (after === 0) {
        option.textContent = 'Before page 1';
      } else if (after === textPageTotal) {
        option.textContent = `After page ${textPageTotal}`;
      } else {
        option.textContent = `After page ${after}`;
      }
      select.append(option);
    }
    select.value = String(normalizeAsciiInsertAfter(asciiPage.insertAfter, textPageTotal));
    select.addEventListener('change', () => {
      asciiPage.insertAfter = normalizeAsciiInsertAfter(select.value, textPageTotal);
      render();
    });
    row.append(select);

    const updateButton = document.createElement('button');
    updateButton.type = 'button';
    updateButton.className = 'secondary-button ascii-item-update';
    updateButton.textContent = 'Update';
    updateButton.disabled = !asciiPage.sourceUrl;
    updateButton.addEventListener('click', async () => {
      if (!asciiPage.sourceUrl) return;
      const columns = clamp(Number(sizeInput.value) || 100, 40, 140);
      const charSet = ensureAsciiCharSet(charsInput.value);
      sizeInput.value = String(columns);
      charsInput.value = charSet;
      updateButton.disabled = true;
      const originalLabel = updateButton.textContent;
      updateButton.textContent = 'Updating...';
      try {
        asciiPage.columns = columns;
        asciiPage.charSet = charSet;
        asciiPage.text = await imageToAscii(asciiPage.sourceUrl, columns, charSet);
        render();
      } catch {
        if (els.asciiStatus) els.asciiStatus.textContent = 'Conversion failed';
      } finally {
        updateButton.textContent = originalLabel;
        updateButton.disabled = false;
      }
    });
    row.append(updateButton);

    const removeButton = document.createElement('button');
    removeButton.type = 'button';
    removeButton.className = 'secondary-button ascii-item-remove';
    removeButton.textContent = 'Remove';
    removeButton.addEventListener('click', () => {
      state.asciiPages = state.asciiPages.filter((entry) => entry.id !== asciiPage.id);
      updateAsciiStatus();
      render();
    });
    row.append(removeButton);

    els.asciiList.append(row);
  });
}

function loadImage(url) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error('Could not decode image'));
    img.src = url;
  });
}

async function imageToAscii(url, targetColumns = 100, chars = ASCII_CHARS) {
  const image = await loadImage(url);
  const columns = clamp(targetColumns, 40, 140);
  const characterSet = ensureAsciiCharSet(chars);
  const rowRatio = 0.5;
  const rows = clamp(Math.round((image.height / image.width) * columns * rowRatio), 26, 180);
  const canvas = document.createElement('canvas');
  canvas.width = columns;
  canvas.height = rows;
  const ctx = canvas.getContext('2d', { willReadFrequently: true });
  ctx.drawImage(image, 0, 0, columns, rows);
  const { data } = ctx.getImageData(0, 0, columns, rows);
  const lines = [];

  for (let y = 0; y < rows; y += 1) {
    let line = '';
    for (let x = 0; x < columns; x += 1) {
      const offset = (y * columns + x) * 4;
      const r = data[offset];
      const g = data[offset + 1];
      const b = data[offset + 2];
      const alpha = data[offset + 3] / 255;
      const luminance = ((0.2126 * r + 0.7152 * g + 0.0722 * b) * alpha + 255 * (1 - alpha)) / 255;
      const index = Math.round((1 - luminance) * (characterSet.length - 1));
      line += characterSet[index];
    }
    lines.push(line);
  }

  return lines.join('\n');
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
    renderFavoriteFontSelect();
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
function addPoints(page, rand, accent) {
  const count = 2 + Math.floor(rand() * 2);
  for (let i = 0; i < count; i += 1) {
    const point = document.createElement('span');
    point.className = `point ${rand() > 0.78 ? 'red' : ''}`;
    const horizontal = rand() > 0.34;
    point.style.left = `${2 + rand() * 32}%`;
    point.style.top = `${6 + rand() * 54}%`;
    const fixRandom = rand() * 1 + 0.77;
    point.style.width = `${fixRandom * 2}px`;
    point.style.height = `${fixRandom * 2}px`;
    point.style.backgroundColor = point.classList.contains('red') ? accent : '';
    page.append(point);
  }
}

// Join truthy class names, dropping false/'' so we never emit stray spaces.
function classes(...names) {
  return names.filter(Boolean).join(' ');
}


function addTape(page, rand, box) {
  const tape = document.createElement('span');
  tape.className = 'tape';
  const width = 42 + rand() * 36;
  const corner = Math.floor(rand() * 4);
  const anchorLeft = corner % 2 === 0 ? box.left : box.left + box.width;
  const anchorTop = corner < 2 ? box.top : box.top + box.height;
  tape.style.width = `${width}px`;
  tape.style.left = `calc(${clamp(anchorLeft, 2, 96)}% - ${width / 2}px)`;
  tape.style.top = `calc(${clamp(anchorTop, 2, 96)}% - 11px)`;
  tape.style.transform = `rotate(${(rand() * 40 - 20).toFixed(1)}deg)`;
  page.append(tape);
}

function addStaples(page, rand) {
  const baseTop = 24 + rand() * 10;
  for (let i = 0; i < 2; i += 1) {
    const staple = document.createElement('span');
    staple.className = 'staple';
    staple.style.left = `${3.4 + rand() * 1.6}%`;
    staple.style.top = `${baseTop + i * 40}%`;
    staple.style.transform = `rotate(${(rand() * 10 - 5).toFixed(1)}deg)`;
    page.append(staple);
  }
}

function addGrain(page, grit) {
  const grain = document.createElement('span');
  grain.className = 'grain';
  grain.style.opacity = (0.1 + grit * 0.32).toFixed(2);
  page.append(grain);
}

// Scatter the collage layer; every motif's odds scale with the Grit slider.
function addCollage(page, rand, settings) {
  const grit = settings.grit;
  if (grit <= 0) return;
  if (rand() < 0.5 * grit) addStaples(page, rand);
  if (grit > 0.05 && settings.showDots) addGrain(page, grit);
}

function clearPageTextEdits() {
  state.bodyEdits.clear();
  state.titleEdits.clear();
}

function clearImageEdits() {
  state.imageEdits.clear();
  state.textZoneEdits.clear();
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

  const resizer = document.createElement('div');
  resizer.className = 'zone-resize-handle';
  resizer.setAttribute('aria-hidden', 'true');
  figure.append(resizer);

  const edit = state.imageEdits.get(imageEditKey(index, slot));
  if (!edit) return;

  figure.style.left = `${edit.left}%`;
  figure.style.top = `${edit.top}%`;
  if (edit.width) figure.style.width = `${edit.width}%`;
  if (edit.height) figure.style.height = `${edit.height}%`;
}

function startImageDrag(event) {
  if (!state.editingText) return;

  const resizer = event.target.closest('.zone-resize-handle');
  const figure = event.target.closest('.movable-image');
  if (!figure) return;

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
    mode: resizer ? 'resize' : 'move',
    offsetX: event.clientX - figureRect.left,
    offsetY: event.clientY - figureRect.top,
    startX: event.clientX,
    startY: event.clientY,
    startWidth: width,
    startHeight: height,
    startLeft: ((figureRect.left - pageRect.left) / pageRect.width) * 100,
    startTop: ((figureRect.top - pageRect.top) / pageRect.height) * 100,
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

  if (drag.mode === 'move') {
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
    state.imageEdits.set(drag.key, { left, top, width: drag.width, height: drag.height });
  } else {
    const dx = ((event.clientX - drag.startX) / drag.pageRect.width) * 100;
    const dy = ((event.clientY - drag.startY) / drag.pageRect.height) * 100;

    const width = clamp(drag.startWidth + dx, 5, 96);
    const height = clamp(drag.startHeight + dy, 5, 92);

    drag.figure.style.width = `${width}%`;
    drag.figure.style.height = `${height}%`;
    state.imageEdits.set(drag.key, {
      left: drag.startLeft,
      top: drag.startTop,
      width,
      height,
    });
  }
}

function stopImageDrag() {
  if (!state.draggedImage) return;

  state.draggedImage.figure.classList.remove('is-dragging');
  state.draggedImage = null;
}

function startTextZoneDrag(event) {
  if (!state.editingText) return;

  const handle = event.target.closest('.zone-drag-handle');
  const resizer = event.target.closest('.zone-resize-handle');
  if (!handle && !resizer) return;

  const textZone = (handle || resizer).closest('.text-zone');
  const page = textZone?.closest('.book-page');
  if (!textZone || !page) return;

  event.preventDefault();
  event.stopPropagation();

  const pageRect = page.getBoundingClientRect();
  const zoneRect = textZone.getBoundingClientRect();
  const left = ((zoneRect.left - pageRect.left) / pageRect.width) * 100;
  const top = ((zoneRect.top - pageRect.top) / pageRect.height) * 100;
  const width = (zoneRect.width / pageRect.width) * 100;
  const height = (zoneRect.height / pageRect.height) * 100;

  state.draggedZone = {
    textZone,
    pageRect,
    mode: handle ? 'move' : 'resize',
    index: Number(textZone.dataset.pageIndex),
    startX: event.clientX,
    startY: event.clientY,
    startLeft: left,
    startTop: top,
    startWidth: width,
    startHeight: height,
  };

  textZone.classList.add('is-dragging-zone');
  event.target.setPointerCapture?.(event.pointerId);
}

function moveTextZoneDrag(event) {
  if (!state.draggedZone) return;

  const drag = state.draggedZone;
  const dx = ((event.clientX - drag.startX) / drag.pageRect.width) * 100;
  const dy = ((event.clientY - drag.startY) / drag.pageRect.height) * 100;

  let left = drag.startLeft;
  let top = drag.startTop;
  let width = drag.startWidth;
  let height = drag.startHeight;

  if (drag.mode === 'move') {
    left = clamp(drag.startLeft + dx, 0, 85);
    top = clamp(drag.startTop + dy, 0, 85);
  } else {
    width = clamp(drag.startWidth + dx, 15, 96);
    height = clamp(drag.startHeight + dy, 12, 92);
  }

  drag.textZone.style.left = `${left}%`;
  drag.textZone.style.top = `${top}%`;
  drag.textZone.style.width = `${width}%`;
  drag.textZone.style.height = `${height}%`;
  state.textZoneEdits.set(drag.index, { left, top, width, height });
}

function stopTextZoneDrag() {
  if (!state.draggedZone) return;

  state.draggedZone.textZone.classList.remove('is-dragging-zone');
  state.draggedZone = null;
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

  const pageKey = editable.dataset.pageIndex;
  if (!pageKey) return;
  const pageIndex = /^-?\d+$/.test(pageKey) ? Number(pageKey) : pageKey;
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

  // Saddle-stitch binding folds the stack in groups of four, so the leaf count
  // must be a multiple of four for the sheets to close cleanly. Pad with blank
  // leaves (rendered as null slots) up to the next multiple of four, inserting
  // them just before the back cover so they fall in the centre of the fold and
  // the back cover stays the outermost last leaf.
  const remainder = pages.length % 4;
  if (remainder !== 0) {
    const lastIsBackCover = pages[pages.length - 1].classList.contains('back-cover');
    const insertAt = lastIsBackCover ? pages.length - 1 : pages.length;
    const blanks = Array.from({ length: 4 - remainder }, () => null);
    pages.splice(insertAt, 0, ...blanks);
  }

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
  page.style.setProperty('--type-spacing', settings.typeSpacing);
  page.style.setProperty('--page-accent', settings.accent);
  applyTextureToPage(page, 'cover');

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
  applyTextureToPage(page, 'back-cover');
  return page;
}

function renderAsciiPage(asciiPage, index, total, settings) {
  const page = document.createElement('article');
  page.className = 'book-page ascii-page';
  const asciiKey = `ascii:${asciiPage.id}`;
  page.style.setProperty('--page-accent', settings.accent);
  page.style.setProperty('--display-font', settings.displayFont);
  page.style.setProperty('--type-scale', settings.typeScale);
  page.style.setProperty('--type-spacing', settings.typeSpacing);
  applyTextureToPage(page, `ascii-${asciiPage.id}`);

  const folio = document.createElement('span');
  folio.className = 'folio-mark';
  folio.textContent = `${settings.title} / ${String(index + 1).padStart(2, '0')}`;
  page.append(folio);

  const pageNumber = document.createElement('span');
  pageNumber.className = 'page-number';
  pageNumber.textContent = `${index + 1}/${total}`;
  page.append(pageNumber);

  const display = document.createElement('h2');
  display.className = 'display-line ascii-display';
  const asciiName = asciiPage.name
    ? asciiPage.name.replace(/\.[^/.]+$/, '').replace(/[_-]+/g, ' ').trim()
    : '';
  if (state.titleEdits.has(asciiKey)) {
    display.innerHTML = state.titleEdits.get(asciiKey);
  } else {
    display.textContent = asciiName || settings.title;
  }
  display.style.fontFamily = settings.displayFont;
  makeEditableText(display, asciiKey, 'title');
  page.append(display);

  const pre = document.createElement('pre');
  pre.className = 'ascii-art';
  pre.textContent = asciiPage.text;
  page.append(pre);

  const keywords = els.keywords.value
    .trim()
    .split(/\s*,\s*/)
    .filter(Boolean);
  const caption = document.createElement('span');
  caption.className = 'caption ascii-caption';
  if (state.bodyEdits.has(asciiKey)) {
    caption.innerHTML = state.bodyEdits.get(asciiKey);
  } else {
    caption.textContent = keywords.length ? keywords[index % keywords.length] : 'ASCII';
  }
  makeEditableText(caption, asciiKey, 'body');
  page.append(caption);

  return page;
}

function renderPage(chunk, index, total, settings, rand, outputIndex) {
  const page = document.createElement('article');
  // "random" → each page rolls its own 1 or 2 columns. Only consume rand() in
  // random mode so fixed 1/2 layouts stay identical to before.
  const pageColumns = settings.columns === 'random' ? (rand() > 0.5 ? 2 : 1) : settings.columns;
  const recipe = layoutRecipe(index, rand, { ...settings, columns: pageColumns });
  page.className = 'book-page';
  page.dataset.layout = recipe.layout;
  page.style.setProperty('--display-font', settings.displayFont);
  page.style.setProperty('--type-scale', settings.typeScale);
  page.style.setProperty('--type-spacing', settings.typeSpacing);
  page.style.setProperty('--body-size', `${pageColumns === 2 ? 9.4 : 10.6}px`);
  page.style.setProperty('--body-leading', pageColumns === 2 ? 1.38 : 1.45);
  page.style.setProperty('--page-accent', settings.accent);
  // Empty = "Auto", leave the var unset so each treatment keeps its own blend.
  if (settings.imageBlend) page.style.setProperty('--image-blend', settings.imageBlend);
  applyTextureToPage(page, `page-${index}`);

  const folio = document.createElement('span');
  folio.className = 'folio-mark';
  folio.textContent = `${settings.title} / ${String(outputIndex + 1).padStart(2, '0')}`;
  page.append(folio);

  const pageNumber = document.createElement('span');
  pageNumber.className = 'page-number';
  pageNumber.textContent = `${outputIndex + 1}/${total}`;
  page.append(pageNumber);

  const imgs = selectedImages();

  if (imgs.length) {
    const image = imgs[(index + Math.floor(rand() * imgs.length)) % imgs.length];
    const figure = document.createElement('figure');
    const torn = settings.grit > 0 && rand() < 0.45 * settings.grit ? pick(['torn-a', 'torn-b'], rand) : '';
    figure.className = classes('image-block', settings.imageTreatment, settings.imageBlend && 'has-blend', rand() > 0.42 && 'line', recipe.rotate, torn);
    placePercent(figure, recipe.image);
    applyImageEdit(figure, index, 'primary');

    const img = document.createElement('img');
    img.src = image.url;
    img.alt = image.name;
    figure.append(img);
    page.append(figure);

    if (settings.grit > 0 && rand() < 0.65 * settings.grit) {
      addTape(page, rand, recipe.image);
    }
  }

  if (imgs.length > 1 && rand() > 0.45) {
    const second = imgs[(index + 1) % imgs.length];
    const figure = document.createElement('figure');
    const torn = settings.grit > 0 && rand() < 0.4 * settings.grit ? pick(['torn-a', 'torn-b'], rand) : '';
    figure.className = classes('image-block', settings.imageTreatment, settings.imageBlend && 'has-blend', torn);
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
  display.className = classes('display-line', recipe.outline && 'outline', recipe.blackTitle && 'black');
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
  textZone.className = `text-zone columns-${pageColumns} ${recipe.spaced ? 'spaced' : ''}`;
  textZone.style.fontFamily = settings.bodyFont;
  textZone.dataset.pageIndex = String(index);
  textZone.dataset.words = String(chunk.trim() ? chunk.trim().split(/\s+/).length : 0);
  placePercent(textZone, recipe.text);

  const zoneEdit = state.textZoneEdits.get(index);
  if (zoneEdit) {
    textZone.style.left = `${zoneEdit.left}%`;
    textZone.style.top = `${zoneEdit.top}%`;
    textZone.style.width = `${zoneEdit.width}%`;
    textZone.style.height = `${zoneEdit.height}%`;
  }

  const dragHandle = document.createElement('div');
  dragHandle.className = 'zone-drag-handle';
  dragHandle.setAttribute('aria-hidden', 'true');
  textZone.append(dragHandle);

  const paragraph = document.createElement('p');
  if (state.bodyEdits.has(index)) {
    paragraph.innerHTML = state.bodyEdits.get(index);
  } else {
    paragraph.textContent = chunk;
  }
  makeEditableText(paragraph, index, 'body');
  textZone.append(paragraph);

  const resizeHandle = document.createElement('div');
  resizeHandle.className = 'zone-resize-handle';
  resizeHandle.setAttribute('aria-hidden', 'true');
  textZone.append(resizeHandle);

  page.append(textZone);

  const caption = document.createElement('span');
  caption.className = 'caption';
  const keywords = els.keywords.value
    .trim()
    .split(/\s*,\s*/)
    .filter(Boolean);
  // Fall back to the stock label pool when no keywords are supplied, so every
  // page still gets its accent tag.
  const tagPool = keywords.length ? keywords : ['index', 'source', 'fragment', 'plate', 'scan', 'note'];
  caption.textContent = pick(tagPool, rand);
  caption.style.left = `${clamp(recipe.image.left + 1, 4, 88)}%`;
  caption.style.top = `${clamp(recipe.image.top + recipe.image.height + 1, 5, 92)}%`;
  page.append(caption);

  addRules(page, rand, settings.accent);
  if (settings.showDots) addPoints(page, rand, settings.accent);
  addCollage(page, rand, settings);
  return page;
}

function currentSettings() {
  const pageCount = clamp(Number(els.pageCount.value) || 4, 1, MAX_PAGES);
  const typeScale = Number(els.typeScale.value) || 1;
  const typeSpacing = Number(els.typeSpacing.value) || 1;
  const displayFont = cssFont(els.displayFont.value);
  const bodyFont = cssFont(els.bodyFont.value);
  return {
    pageCount,
    columns: state.columns,
    accent: els.accentColor.value,
    typeScale,
    typeSpacing,
    imageEnergy: Number(els.imageEnergy.value) || 0,
    imageTreatment: state.imageTreatment,
    imageBlend: els.imageBlend.value,
    grit: clamp((Number(els.grit.value) || 0) / 100, 0, 1),
    showDots: els.showDots.checked,
    displayFont,
    bodyFont,
    bodySize: state.columns === 2 ? 9.4 : 10.6,
    leading: state.columns === 2 ? 1.38 : 1.45,
    title: titleFromText(els.text.value),
  };
}

function renderPages() {
  document.documentElement.style.setProperty('--accent', els.accentColor.value);
  updateAsciiInsertOptions();
  const textPageTotal = currentTextPageTotal();
  state.asciiPages = state.asciiPages
    .map((record) => ensureAsciiRecord(record, textPageTotal))
    .filter(Boolean);
  renderAsciiList();
  updateAsciiStatus();
  const settings = currentSettings();
  const bodyText = bodyFromText(els.text.value);
  const distribution = state.wordDistribution;
  const chunks =
    distribution && distribution.length === settings.pageCount
      ? packParagraphs(splitParagraphs(bodyText), distribution)
      : splitText(bodyText, settings.pageCount);
  const total = textPageTotal + state.asciiPages.length;
  els.pages.innerHTML = '';

  if (!total) {
    const empty = document.querySelector('#emptyPageTemplate').content.cloneNode(true);
    els.pages.append(empty);
    updateEditTextMode();
    return;
  }
  const rand = seededRandom(state.seed);
  const asciiBuckets = new Map();
  state.asciiPages.forEach((page) => {
    const slot = normalizeAsciiInsertAfter(page.insertAfter, textPageTotal);
    if (!asciiBuckets.has(slot)) asciiBuckets.set(slot, []);
    asciiBuckets.get(slot).push(page);
  });
  let renderedPages = 0;
  els.pages.append(renderCoverPage(settings));

  const renderAsciiAtSlot = (slot) => {
    const pagesAtSlot = asciiBuckets.get(slot) || [];
    pagesAtSlot.forEach((asciiPage) => {
      els.pages.append(renderAsciiPage(asciiPage, renderedPages, total, settings));
      renderedPages += 1;
    });
  };

  renderAsciiAtSlot(0);
  for (let index = 0; index < textPageTotal; index += 1) {
    const chunk = chunks[index] || '';
    els.pages.append(renderPage(chunk, index, total, settings, rand, renderedPages));
    renderedPages += 1;
    renderAsciiAtSlot(index + 1);
  }
  els.pages.append(renderBackCoverPage(settings));
  updateEditTextMode();
}

// Per-text-page fill measurement. `ratio` is how full the zone is: <1 under-filled,
// ~1 just full, >1 clipped (zones are `overflow: hidden`). The paragraph's own
// height — not the zone's clamped scrollHeight — reveals under-fill. In multi-column
// layouts overflow appears horizontally (content spilling into extra columns), so
// when that happens we trust the horizontal measure; otherwise the vertical one.
function measureTextFills() {
  const zones = [
    ...els.pages.querySelectorAll(
      '.book-page:not(.cover-page):not(.ascii-page):not(.back-cover) .text-zone',
    ),
  ];
  return zones.map((zone) => {
    const paragraph = zone.querySelector('p');
    const contentHeight = paragraph ? paragraph.offsetHeight : zone.scrollHeight;
    const verticalFill = contentHeight / Math.max(1, zone.clientHeight);
    const horizontalFill = zone.scrollWidth / Math.max(1, zone.clientWidth);
    const ratio = horizontalFill > 1.01 ? horizontalFill : verticalFill;
    return { words: Number(zone.dataset.words || 0), ratio };
  });
}

// Capacity (in words) of a page = words it currently holds / how full it is.
// Pages that are nearly empty give an unreliable estimate, so they fall back to
// the average of the pages we can measure.
function estimateCapacities(fills) {
  const measured = fills
    .filter((fill) => fill.words > 0 && fill.ratio > 0.02)
    .map((fill) => fill.words / fill.ratio);
  const average = measured.length
    ? measured.reduce((sum, cap) => sum + cap, 0) / measured.length
    : 0;
  return { caps: fills.map((fill) => (fill.words > 0 && fill.ratio > 0.02 ? fill.words / fill.ratio : average)), average };
}

function renderAtPageCount(count, distribution = null) {
  els.pageCount.value = String(count);
  state.wordDistribution = distribution;
  renderPages();
}

// Allocate `words` across `count` pages in proportion to each page's measured
// capacity, so every page fills to roughly the same level instead of the
// tightest page capping the rest. Iterates a couple of times to converge.
function distributeProportionally(wordTotal, count, iterations = 2) {
  let counts = Array.from({ length: count }, () => wordTotal / count);
  let fills = [];
  for (let i = 0; i < iterations; i += 1) {
    renderAtPageCount(count, counts);
    fills = measureTextFills();
    const { caps, average } = estimateCapacities(fills);
    const totalCap = caps.reduce((sum, cap) => sum + cap, 0) || average * count || wordTotal;
    counts = caps.map((cap) => (wordTotal * cap) / totalCap);
  }
  const maxRatio = fills.reduce((max, fill) => Math.max(max, fill.ratio), 0);
  return { counts, fits: maxRatio <= 1.01 };
}

// Auto-fit: pack every page to capacity, then use the fewest pages that still
// show all the text. We probe once with an even split to estimate average page
// capacity, jump straight to the estimated page count, then nudge up/down so the
// result is both gap-free and minimal.
function autoFitPages() {
  const words = textWords(bodyFromText(els.text.value));
  if (!words.length) {
    state.wordDistribution = null;
    renderPages();
    return;
  }

  // Probe with an even split to learn the average per-page capacity.
  const probe = clamp(Number(els.pageCount.value) || 4, 1, MAX_PAGES);
  renderAtPageCount(probe, null);
  const { average } = estimateCapacities(measureTextFills());
  const avgCap = average > 0 ? average : words.length;

  let n = clamp(Math.ceil(words.length / avgCap), 1, MAX_PAGES);
  let result = distributeProportionally(words.length, n);

  // Grow if the estimate was optimistic and text still clips.
  while (!result.fits && n < MAX_PAGES) {
    n += 1;
    result = distributeProportionally(words.length, n);
  }

  // Shrink if a page could be removed without clipping (estimate was generous).
  while (n > 1) {
    const candidate = distributeProportionally(words.length, n - 1);
    if (!candidate.fits) break;
    n -= 1;
    result = candidate;
  }

  // Ensure the final DOM reflects the chosen distribution.
  renderAtPageCount(n, result.counts);
}

function render() {
  if (state.autoPages) {
    autoFitPages();
  } else {
    state.wordDistribution = null;
    renderPages();
  }
}

const debouncedRender = debounce(render, 140);

function updateAutoPagesMode() {
  state.autoPages = Boolean(els.autoPages.checked);
  els.pageCount.disabled = state.autoPages;
  els.pageCount.title = state.autoPages
    ? 'Page count is set automatically to fit all text'
    : '';
  writeAutoPagesPreference(state.autoPages);
}

function randomize() {
  const rand = seededRandom(Date.now() % 1000000);
  const fonts = allFonts();
  clearPageTextEdits();
  clearImageEdits();
  state.seed = Math.floor(rand() * 1000000);
  state.columns = rand() > 0.48 ? 2 : 1;
  // Prefer a curated pairing so randomized fonts stay aesthetically matched;
  // occasionally fall back to fully random fonts for variety.
  if (fontPairings.length && rand() > 0.25) {
    const pairing = pick(fontPairings, rand);
    setFontPickerValue(els.displayFont, pairing.displayId);
    setFontPickerValue(els.bodyFont, pairing.bodyId);
    if (els.fontPairingSelect) els.fontPairingSelect.value = pairing.id;
  } else {
    setFontPickerValue(els.displayFont, pick(fonts, rand).id);
    setFontPickerValue(els.bodyFont, pick(fonts, rand).id);
    if (els.fontPairingSelect) els.fontPairingSelect.value = '';
  }
  els.typeScale.value = String((0.52 + rand() * 2).toFixed(2));
  els.typeSpacing.value = Number((0.82 + rand() * 2).toFixed(2)) - 0.3;
  els.imageEnergy.value = String(Math.floor(35 + rand() * 64));
  els.accentColor.value = pick(palettes, rand);
  state.imageTreatment = pick(['mono', 'mono', 'duotone', 'halftone'], rand);
  els.grit.value = String(Math.floor(18 + rand() * 62));
  if (els.layoutName) {
    els.layoutName.textContent = pick(layoutNames, rand);
  }
  updateSegments();
  updateTreatmentSegments();
  render();
}

function updateSegments() {
  els.segments.forEach((segment) => {
    segment.classList.toggle('is-active', segment.dataset.columns === String(state.columns));
  });
}

function updateTreatmentSegments() {
  els.treatmentSegments.forEach((segment) => {
    segment.classList.toggle('is-active', segment.dataset.treatment === state.imageTreatment);
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

// — Book library (IndexedDB) —

const DB_NAME = 'blook-studio';
const DB_VERSION = 1;
const BOOK_STORE = 'books';

function openDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = (e) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains(BOOK_STORE)) {
        db.createObjectStore(BOOK_STORE, { keyPath: 'id' });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function dbPut(record) {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const req = db.transaction(BOOK_STORE, 'readwrite').objectStore(BOOK_STORE).put(record);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function dbList() {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const req = db.transaction(BOOK_STORE, 'readonly').objectStore(BOOK_STORE).getAll();
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function dbGet(id) {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const req = db.transaction(BOOK_STORE, 'readonly').objectStore(BOOK_STORE).get(id);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function dbRemove(id) {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const req = db.transaction(BOOK_STORE, 'readwrite').objectStore(BOOK_STORE).delete(id);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

function serializeBook() {
  return {
    id: Date.now(),
    name: titleFromText(els.text.value) || 'Untitled',
    savedAt: new Date().toISOString(),
    text: els.text.value,
    seed: state.seed,
    columns: state.columns,
    pageCountVal: els.pageCount.value,
    typeScaleVal: els.typeScale.value,
    typeSpacingVal: els.typeSpacing.value,
    imageEnergyVal: els.imageEnergy.value,
    imageTreatment: state.imageTreatment,
    imageBlendVal: els.imageBlend.value,
    gritVal: els.grit.value,
    showDots: els.showDots.checked,
    accentColor: els.accentColor.value,
    displayFontId: els.displayFont.value,
    bodyFontId: els.bodyFont.value,
    bodyEdits: [...state.bodyEdits.entries()],
    titleEdits: [...state.titleEdits.entries()],
    imageEdits: [...state.imageEdits.entries()],
    textZoneEdits: [...state.textZoneEdits.entries()],
    images: state.images,
    asciiPages: state.asciiPages,
  };
}

async function saveCurrentBook() {
  els.save.disabled = true;
  await dbPut(serializeBook());
  const orig = els.save.textContent;
  els.save.textContent = 'Saved ✓';
  setTimeout(() => {
    els.save.textContent = orig;
    els.save.disabled = false;
  }, 1400);
}

async function loadBook(id) {
  const record = await dbGet(id);
  if (!record) return;

  state.images = record.images || [];
  state.asciiPages = Array.isArray(record.asciiPages)
    ? record.asciiPages.filter((page) => page && typeof page.text === 'string')
    : [];
  makeImageStrip();
  updateImageCount();
  updateAsciiStatus();

  els.text.value = record.text || '';
  els.pageCount.value = record.pageCountVal;
  els.typeScale.value = record.typeScaleVal;
  els.typeSpacing.value = record.typeSpacingVal;
  els.imageEnergy.value = record.imageEnergyVal;
  state.imageTreatment = record.imageTreatment || 'mono';
  els.imageBlend.value = record.imageBlendVal || '';
  els.grit.value = record.gritVal ?? els.grit.value;
  els.showDots.checked = record.showDots ?? false;
  els.accentColor.value = record.accentColor;
  state.columns = record.columns;
  state.seed = record.seed;

  state.bodyEdits = new Map(record.bodyEdits || []);
  state.titleEdits = new Map(record.titleEdits || []);
  state.imageEdits = new Map(record.imageEdits || []);
  state.textZoneEdits = new Map(record.textZoneEdits || []);

  refreshFontSelects();
  if (record.displayFontId) setFontPickerValue(els.displayFont, record.displayFontId);
  if (record.bodyFontId) setFontPickerValue(els.bodyFont, record.bodyFontId);

  updateSegments();
  updateTreatmentSegments();
  closeLibrary();
  render();
}

async function deleteBook(id, name) {
  if (!confirm(`Delete "${name}"?`)) return;
  await dbRemove(id);
  renderLibraryGrid();
}

function formatSavedDate(iso) {
  return new Date(iso).toLocaleDateString(undefined, {
    month: 'short',
    day: 'numeric',
    year: 'numeric',
  });
}

async function renderLibraryGrid() {
  const grid = document.querySelector('#libraryGrid');
  grid.innerHTML = '';
  const books = await dbList();

  if (!books.length) {
    const empty = document.createElement('p');
    empty.className = 'library-empty';
    empty.textContent = 'No saved books yet. Generate a layout and hit Save.';
    grid.append(empty);
    return;
  }

  books
    .sort((a, b) => b.id - a.id)
    .forEach((book) => {
      const card = document.createElement('div');
      card.className = 'library-card';

      const meta = document.createElement('div');
      meta.className = 'library-card-meta';

      const accent = document.createElement('div');
      accent.className = 'library-card-accent';
      accent.style.backgroundColor = book.accentColor || '#ff2a1c';

      const title = document.createElement('strong');
      title.className = 'library-card-title';
      title.textContent = book.name;

      const date = document.createElement('span');
      date.className = 'library-card-date';
      date.textContent = formatSavedDate(book.savedAt);

      const stats = document.createElement('span');
      stats.className = 'library-card-stats';
      stats.textContent = `${book.pageCountVal || '?'} pages · ${(book.images || []).length} images`;

      meta.append(accent, title, date, stats);

      const actions = document.createElement('div');
      actions.className = 'library-card-actions';

      const loadBtn = document.createElement('button');
      loadBtn.className = 'primary-button';
      loadBtn.textContent = 'Load';
      loadBtn.addEventListener('click', () => loadBook(book.id));

      const delBtn = document.createElement('button');
      delBtn.className = 'secondary-button delete-btn';
      delBtn.textContent = 'Delete';
      delBtn.addEventListener('click', () => deleteBook(book.id, book.name));

      actions.append(loadBtn, delBtn);
      card.append(meta, actions);
      grid.append(card);
    });
}

async function openLibrary() {
  const overlay = document.querySelector('#libraryOverlay');
  overlay.hidden = false;
  overlay.removeAttribute('aria-hidden');
  await renderLibraryGrid();
}

function closeLibrary() {
  const overlay = document.querySelector('#libraryOverlay');
  overlay.hidden = true;
  overlay.setAttribute('aria-hidden', 'true');
}

els.generate.addEventListener('click', render);
els.save.addEventListener('click', saveCurrentBook);
els.library.addEventListener('click', openLibrary);
els.closeLibrary.addEventListener('click', closeLibrary);
els.displayFontTrigger.addEventListener('click', () => openFontPicker('display'));
els.bodyFontTrigger.addEventListener('click', () => openFontPicker('body'));
els.displayFontMenu.addEventListener('click', (event) => {
  if (!event.target.closest('.font-picker-option')) return;
  closeFontPicker('display');
});
els.bodyFontMenu.addEventListener('click', (event) => {
  if (!event.target.closest('.font-picker-option')) return;
  closeFontPicker('body');
});
els.favoriteFontsSelect.addEventListener('change', () => {
  state.activeFavoritePresetId = els.favoriteFontsSelect.value;
  applyFavoriteFontPresetById(state.activeFavoritePresetId);
  updateFavoriteFontStatus();
});
els.removeFavoriteFontBtn.addEventListener('click', () => {
  removeFavoriteFontPreset(els.favoriteFontsSelect.value);
});
els.fontPairingSelect.addEventListener('change', () => {
  if (els.fontPairingSelect.value) applyPairingById(els.fontPairingSelect.value);
});
els.shufflePairingBtn.addEventListener('click', () => {
  const pairing = fontPairings[Math.floor(Math.random() * fontPairings.length)];
  applyPairingById(pairing.id);
});
els.applyTextures.addEventListener('input', updateTextureMode);
els.applyTextures.addEventListener('change', updateTextureMode);
els.autoPages.addEventListener('change', () => {
  updateAutoPagesMode();
  render();
});
document.querySelector('#libraryOverlay').addEventListener('click', (e) => {
  if (e.target === e.currentTarget) closeLibrary();
});
els.randomize.addEventListener('click', randomize);
els.print.addEventListener('click', createBookletPdf);
els.systemFontBtn.addEventListener('click', loadSystemFonts);
els.favoriteFontBtn.addEventListener('click', saveFavoriteFonts);
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
document.addEventListener('click', (event) => {
  if (!event.target.closest('.font-picker')) closeAllFontPickers();
});
document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape') closeAllFontPickers();
});
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
    state.columns = segment.dataset.columns === 'random' ? 'random' : Number(segment.dataset.columns);
    updateSegments();
    render();
  });
});

els.treatmentSegments.forEach((segment) => {
  segment.addEventListener('click', () => {
    state.imageTreatment = segment.dataset.treatment;
    updateTreatmentSegments();
    render();
  });
});

els.text.addEventListener('input', () => {
  clearPageTextEdits();
  debouncedRender();
});
els.text.addEventListener('change', () => {
  clearPageTextEdits();
  render();
});

for (const el of [
  els.pageCount,
  els.accentColor,
  els.typeScale,
  els.typeSpacing,
  els.imageEnergy,
  els.imageBlend,
  els.grit,
  els.showDots,
  els.displayFont,
  els.bodyFont,
]) {
  el.addEventListener('input', () => {
    if (el === els.displayFont || el === els.bodyFont) syncFontSelectPreview(el);
    // Auto-fit re-runs a multi-render search, so debounce high-frequency input
    // (slider drags) to keep the preview smooth.
    if (state.autoPages) debouncedRender();
    else render();
  });
  el.addEventListener('change', () => {
    if (el === els.displayFont || el === els.bodyFont) syncFontSelectPreview(el);
    render();
  });
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
els.pages.addEventListener('pointerdown', startTextZoneDrag);
window.addEventListener('pointermove', moveDraggedImage);
window.addEventListener('pointermove', moveTextZoneDrag);
window.addEventListener('pointerup', stopImageDrag);
window.addEventListener('pointerup', stopTextZoneDrag);
window.addEventListener('pointercancel', stopImageDrag);
window.addEventListener('pointercancel', stopTextZoneDrag);

els.imageInput.addEventListener('change', async (event) => {
  const images = await readFiles(event.target.files, (file, url) => ({
    name: file.name,
    url,
    selected: true,
  }));
  state.images = images;
  clearImageEdits();
  updateImageCount();
  makeImageStrip();
  render();
});

// Openverse: free Creative Commons image search, no API key required.
// https://api.openverse.org/v1/images/  — returns proxied thumbnails that are
// safe to use as plain <img src> (display + native print, no canvas taint).
// Anonymous requests cap page_size at 20, so page through to reach `count`.
async function searchOpenverse(query, count = 40) {
  const perPage = 20;
  const pages = Math.ceil(count / perPage);
  const collected = [];
  for (let page = 1; page <= pages; page += 1) {
    const params = new URLSearchParams({
      q: query,
      page_size: String(perPage),
      page: String(page),
    });
    const res = await fetch(`https://api.openverse.org/v1/images/?${params}`);
    if (!res.ok) {
      if (collected.length) break; // keep whatever earlier pages returned
      throw new Error(`Openverse responded ${res.status}`);
    }
    const data = await res.json();
    const batch = (data.results || [])
      .map((r) => ({ name: r.title || query, url: r.thumbnail || r.url, selected: false }))
      .filter((r) => r.url);
    collected.push(...batch);
    if (batch.length < perPage) break; // no more results
  }
  return collected.slice(0, count);
}

async function runImageSearch() {
  const query = els.imageSearchInput.value.trim();
  if (!query) return;
  els.imageSearchStatus.textContent = 'Searching…';
  els.imageSearchBtn.disabled = true;
  try {
    const found = await searchOpenverse(query, 40);
    if (!found.length) {
      els.imageSearchStatus.textContent = `No results for "${query}"`;
      return;
    }
    state.images = state.images.concat(found);
    updateImageCount();
    els.imageSearchStatus.textContent = `${found.length} found for "${query}" — click thumbnails to use them`;
    makeImageStrip();
    render();
  } catch (err) {
    console.error(err);
    els.imageSearchStatus.textContent = 'Search failed — try again';
  } finally {
    els.imageSearchBtn.disabled = false;
  }
}

els.imageSearchBtn.addEventListener('click', runImageSearch);
els.imageSearchInput.addEventListener('keydown', (event) => {
  if (event.key === 'Enter') {
    event.preventDefault();
    runImageSearch();
  }
});

els.asciiInput.addEventListener('change', async (event) => {
  const file = event.target.files?.[0];
  if (!file) return;

  if (els.asciiStatus) els.asciiStatus.textContent = 'Converting...';
  try {
    const chars = ensureAsciiCharSet(els.asciiChars.value);
    const columns = clamp(Number(els.asciiColumns.value) || 100, 40, 140);
    els.asciiChars.value = chars;
    els.asciiColumns.value = String(columns);

    const [{ name, url }] = await readFiles([file], (item, dataUrl) => ({
      name: item.name,
      url: dataUrl,
    }));
    const ascii = await imageToAscii(url, columns, chars);
    state.asciiPages.push(
      ensureAsciiRecord({
        id: `${Date.now()}-${Math.floor(Math.random() * 1e6)}`,
        name,
        text: ascii,
        sourceUrl: url,
        columns,
        charSet: chars,
        insertAfter: Number(els.asciiInsertAfter.value || 0),
      }),
    );
    updateAsciiStatus();
    render();
  } catch {
    if (els.asciiStatus) els.asciiStatus.textContent = 'Conversion failed';
  } finally {
    event.target.value = '';
  }
});

window.addEventListener('resize', () => {
  if (state.zoom < 1) fitPreview();
});
window.addEventListener('afterprint', cleanupBookletPrint);
document.addEventListener('selectionchange', () => {
  saveTextSelection();
  updateFormatState();
});

loadGoogleFontCatalog();
refreshFontSelects();
renderFavoriteFontSelect();
renderPairingSelect();
updateFavoriteFontStatus();
renderTextureStrip();
updateAsciiStatus();
els.applyTextures.checked = state.applyTextures;
updateTextureMode();
els.autoPages.checked = state.autoPages;
updateAutoPagesMode();
closeAllFontPickers();
updateSegments();
updateTreatmentSegments();
render();

// Webfonts load asynchronously and have different metrics than the fallback
// stacks, which can change how much text fits — re-fit once they settle.
if (document.fonts?.ready) {
  document.fonts.ready.then(() => {
    if (state.autoPages) render();
  });
}
