/**
 * tokens.js
 *
 * Loads forge-style-dict/tokens.json and resolves {Token.Path} references to
 * their final concrete values.
 *
 * Public API
 * ──────────
 *   resolveToken(dotPath, [setName])
 *       Resolve a dot-path from a specific token set (or any set if omitted).
 *       e.g. resolveToken("Color.Blue.300", "Primitives/Color")
 *            resolveToken("Color.Primary",  "Color Modes / Dark")
 *
 *   C       – dark-mode semantic color map  (bare 6-digit hex, no #)
 *   FONT    – font family aliases
 *   PALETTE – 7-color data-series palette from dark-mode tokens
 *   tokens  – full flat map of every { setName, path } → resolved value
 *
 * Resolution rules
 * ────────────────
 * 1. Primitive color tokens  (Primitives/Color)  are the ground truth for
 *    refs like {Color.Blue.300}.
 * 2. Semantic tokens in a given Color Mode reference primitives (or other
 *    semantic tokens in the same or earlier sets).
 * 3. When resolving a ref that is NOT found in the current set, the resolver
 *    falls back through the set priority order below.
 */

"use strict";

const path = require("path");
const rawTokens = require(path.join(__dirname, "forge-style-dict/tokens.json"));

// ─── Set priority: earlier = lower priority (later sets may shadow) ────────
// We define an explicit lookup order for cross-set reference resolution.
// Primitives always come first; specific mode sets are queried last.
const SET_PRIORITY = [
  "Primitives/Default",
  "Primitives/Color",
  "Primitives/Design",
  "Scale/Default",
  "Font/Default",
  "Font/Adobe",
  "Font/Open",
  "Color Modes / Default",
  "Color Modes / Dark",   // ← intentionally before BW; BW is not used
  // "Color Modes / BW" excluded – we never want BW to shadow Dark
  // "Color Modes / Inverted" excluded for same reason
  "Media/Desktop",
  "Media/Mobile",
  "Media/Document",
  "Media/Slide",
];

// ─── 1. Index every token set into its own flat path→tokenObj map ─────────
/** @type {Map<string, Map<string, {value: any, type: string}>>} */
const setMaps = new Map();

function indexSet(obj, pathParts, map) {
  if (!obj || typeof obj !== "object") return;
  if ("value" in obj && "type" in obj) {
    map.set(pathParts.join("."), obj);
    return;
  }
  for (const [k, v] of Object.entries(obj)) {
    if (k.startsWith("$")) continue;
    indexSet(v, [...pathParts, k], map);
  }
}

for (const [setName, setObj] of Object.entries(rawTokens)) {
  if (setName.startsWith("$")) continue;
  const map = new Map();
  indexSet(setObj, [], map);
  setMaps.set(setName, map);
}

// ─── 2. Resolver ──────────────────────────────────────────────────────────
// Cache key = "setName\0dotPath" (or "\0dotPath" for any-set lookups)
const RESOLVE_CACHE = new Map();
const RESOLVING = new Set();

/**
 * Look up a raw token object by dotPath, searching through the set priority
 * list.  If `preferredSet` is given it is searched first.
 */
function findRaw(dotPath, preferredSet) {
  // 1. Try the preferred set first
  if (preferredSet) {
    const m = setMaps.get(preferredSet);
    if (m && m.has(dotPath)) return { tok: m.get(dotPath), set: preferredSet };
  }
  // 2. Walk the priority list
  for (const setName of SET_PRIORITY) {
    if (setName === preferredSet) continue; // already tried
    const m = setMaps.get(setName);
    if (m && m.has(dotPath)) return { tok: m.get(dotPath), set: setName };
  }
  return null;
}

/**
 * Resolve a raw value string (or object) to a concrete value.
 * Colors come back as bare 6-digit hex UPPERCASE.
 * Single {Ref} tokens are resolved recursively.
 * Composite objects (typography) have each field resolved.
 */
function resolveValue(raw, type, originSet) {
  if (raw === null || raw === undefined) return raw;

  if (typeof raw === "string") {
    // Plain hex  (#RRGGBB or RRGGBB)
    const hexMatch = raw.match(/^#?([0-9A-Fa-f]{6})$/);
    if (hexMatch) return hexMatch[1].toUpperCase();

    // Single token reference  {Foo.Bar.Baz}
    const singleRef = raw.match(/^\{([^}]+)\}$/);
    if (singleRef) {
      return resolveToken(singleRef[1], originSet);
    }

    // Anything else (font names, format codes, expressions, …)
    return raw;
  }

  if (typeof raw === "number") return raw;

  // Composite / object value (typography token etc.)
  if (typeof raw === "object") {
    const out = {};
    for (const [k, v] of Object.entries(raw)) {
      out[k] = resolveValue(v, null, originSet);
    }
    return out;
  }

  return raw;
}

/**
 * Resolve a dot-path token to its concrete value.
 *
 * @param {string} dotPath     e.g. "Color.Blue.300" or "fontFamilies.Heading"
 * @param {string} [setHint]   optional preferred token set to search first
 * @returns {string|number|object|undefined}
 */
function resolveToken(dotPath, setHint) {
  const cacheKey = `${setHint || ""}\0${dotPath}`;
  if (RESOLVE_CACHE.has(cacheKey)) return RESOLVE_CACHE.get(cacheKey);
  if (RESOLVING.has(cacheKey)) {
    console.warn(`[tokens] Circular reference: ${dotPath}`);
    return undefined;
  }

  const found = findRaw(dotPath, setHint);
  if (!found) return undefined;

  RESOLVING.add(cacheKey);
  // Resolve references in the context of the *found* set so that, e.g.,
  // "Color.Primary" in "Color Modes / Dark" resolves its own ref correctly.
  const resolved = resolveValue(found.tok.value, found.tok.type, found.set);
  RESOLVING.delete(cacheKey);

  RESOLVE_CACHE.set(cacheKey, resolved);
  return resolved;
}

// ─── 3. Convenience colour helper ─────────────────────────────────────────
// Shorthand: resolve from dark mode, falling back to primitives automatically.
const DARK = "Color Modes / Dark";

function dc(dotPath) {
  return resolveToken(dotPath, DARK);
}

// ─── 4. Exported colour / font maps ───────────────────────────────────────

/**
 * Dark-mode semantic color tokens (bare 6-digit hex, no leading #).
 */
const C = {
  // Backgrounds / surfaces
  surface:        dc("Color.Surface"),           // deep dark navy
  surfaceVariant: dc("Color.Surface Variant"),   // slightly lighter
  containerHigh:  dc("Color.Container High"),    // header / footer bar
  containerMed:   dc("Color.Container Medium"),

  // Text on dark backgrounds
  onSurface:      dc("Color.On Surface"),        // primary text
  onSurfaceLow:   dc("Color.On Surface Low"),    // secondary / muted text
  white:          dc("Color.Heading Text"),      // pure white heading text

  // Brand / data colours
  primary:        dc("Color.Primary"),           // blue
  primaryLight:   dc("Color.Base.400"),          // lighter blue
  secondary:      dc("Color.Secondary"),         // mid blue
  accent:         dc("Color.Accent"),            // amber / yellow
  accentText:     dc("Color.Accent Text"),       // red
  success:        resolveToken("Color.Green.500", "Primitives/Color"),
  outline:        dc("Color.Outline"),           // grid-line / border colour

  // Extended ramp (Base palette from dark mode)
  baseDeep:  dc("Color.Base.100"),
  baseMid:   dc("Color.Base.200"),
  base300:   dc("Color.Base.300"),
  base500:   dc("Color.Base.500"),
};

/**
 * Font family tokens resolved from Font/Default.
 */
const FONT = {
  heading: resolveToken("fontFamilies.Heading", "Font/Default"),  // Empirica NYCMayor
  body:    resolveToken("fontFamilies.Body",    "Font/Default"),  // Empirica Text
  ui:      resolveToken("fontFamilies.UI",      "Font/Default"),  // Instrument Sans
  accent:  resolveToken("fontFamilies.Accent",  "Font/Default"),  // Instrument Sans
};

/**
 * Seven-color data-series palette built from dark-mode tokens.
 */
const PALETTE = [
  C.accent,      // amber  – warm standout
  C.primary,     // blue
  C.accentText,  // red
  C.base300,     // soft blue-purple
  C.success,     // green
  C.base500,     // light periwinkle
  C.baseMid,     // mid blue
];

// ─── 5. Full resolved flat map ────────────────────────────────────────────
// Pre-resolve every token from every set for inspection / other consumers.
const tokens = {};
for (const [setName, map] of setMaps) {
  for (const [p] of map) {
    tokens[`${setName}/${p}`] = resolveToken(p, setName);
  }
}

module.exports = {
  /** Resolve any dot-path to its concrete value (optionally set-scoped). */
  resolveToken,
  /** Fully resolved flat map keyed as "SetName/dot.path". */
  tokens,
  /** Dark-mode semantic color aliases (bare 6-digit hex). */
  C,
  /** Font family aliases. */
  FONT,
  /** Seven-color data-series palette. */
  PALETTE,
};
