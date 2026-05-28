/** RGB channels 0–255 and alpha 0–1. */
export interface RgbColor {
  r: number;
  g: number;
  b: number;
  a: number;
}

const SCHEME_ALIASES: Record<string, string> = {
  bg1: "lt1",
  bg2: "lt2",
  tx1: "dk1",
  tx2: "dk2",
};

function clamp01(n: number): number {
  return Math.min(1, Math.max(0, n));
}

function parseHex(hex: string): RgbColor | null {
  const m = hex.replace(/^#/, "").match(/^([0-9a-fA-F]{6})$/);
  if (!m) return null;
  const n = parseInt(m[1]!, 16);
  return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255, a: 1 };
}

function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
  const rn = r / 255;
  const gn = g / 255;
  const bn = b / 255;
  const max = Math.max(rn, gn, bn);
  const min = Math.min(rn, gn, bn);
  const l = (max + min) / 2;
  if (max === min) return [0, l, 0];

  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h = 0;
  switch (max) {
    case rn:
      h = ((gn - bn) / d + (gn < bn ? 6 : 0)) / 6;
      break;
    case gn:
      h = ((bn - rn) / d + 2) / 6;
      break;
    default:
      h = ((rn - gn) / d + 4) / 6;
      break;
  }
  return [h, l, s];
}

function hslToRgb(h: number, l: number, s: number): [number, number, number] {
  if (s === 0) {
    const v = Math.round(l * 255);
    return [v, v, v];
  }

  const hue2rgb = (p: number, q: number, t: number) => {
    let tt = t;
    if (tt < 0) tt += 1;
    if (tt > 1) tt -= 1;
    if (tt < 1 / 6) return p + (q - p) * 6 * tt;
    if (tt < 1 / 2) return q;
    if (tt < 2 / 3) return p + (q - p) * (2 / 3 - tt) * 6;
    return p;
  };

  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return [
    Math.round(hue2rgb(p, q, h + 1 / 3) * 255),
    Math.round(hue2rgb(p, q, h) * 255),
    Math.round(hue2rgb(p, q, h - 1 / 3) * 255),
  ];
}

function applyTint(color: RgbColor, amount: number): RgbColor {
  const [h, l, s] = rgbToHsl(color.r, color.g, color.b);
  const nextL = l * (1 - amount) + amount;
  const [r, g, b] = hslToRgb(h, nextL, s);
  return { r, g, b, a: color.a };
}

function applyShade(color: RgbColor, amount: number): RgbColor {
  const [h, l, s] = rgbToHsl(color.r, color.g, color.b);
  const [r, g, b] = hslToRgb(h, l * (1 - amount), s);
  return { r, g, b, a: color.a };
}

function applyLumMod(color: RgbColor, amount: number): RgbColor {
  const [h, l, s] = rgbToHsl(color.r, color.g, color.b);
  const [r, g, b] = hslToRgb(h, Math.min(1, l * amount), s);
  return { r, g, b, a: color.a };
}

function applyLumOff(color: RgbColor, amount: number): RgbColor {
  const [h, l, s] = rgbToHsl(color.r, color.g, color.b);
  const [r, g, b] = hslToRgb(h, clamp01(l + amount), s);
  return { r, g, b, a: color.a };
}

function pct(el: Element | null | undefined): number | null {
  if (!el) return null;
  const raw = el.getAttribute("val") ?? el.getAttribute("amt");
  if (raw == null || raw === "") return null;
  const n = Number(raw);
  return Number.isFinite(n) ? n / 100000 : null;
}

function applyModifier(color: RgbColor, mod: Element): RgbColor {
  const tag = mod.localName;
  const amount = pct(mod);
  if (amount == null) return color;

  switch (tag) {
    case "tint":
      return applyTint(color, amount);
    case "shade":
      return applyShade(color, amount);
    case "lumMod":
      return applyLumMod(color, amount);
    case "lumOff":
      return applyLumOff(color, amount);
    case "alpha":
      return { ...color, a: clamp01(amount) };
    case "alphaMod":
      return { ...color, a: clamp01(color.a * amount) };
    case "alphaOff":
      return { ...color, a: clamp01(color.a + amount) };
    default:
      return color;
  }
}

function resolveSchemeKey(val: string, themeColors?: Record<string, string>): string | undefined {
  if (!themeColors) return undefined;
  const resolved = SCHEME_ALIASES[val] ?? val;
  return themeColors[resolved] ?? themeColors[val];
}

/** Resolve base RGB from an a:srgbClr / a:schemeClr / a:sysClr element. */
export function colorFromClrElement(
  clrEl: Element,
  themeColors?: Record<string, string>
): RgbColor | null {
  const tag = clrEl.localName;

  if (tag === "srgbClr") {
    const parsed = parseHex(clrEl.getAttribute("val") ?? "");
    return parsed;
  }

  if (tag === "schemeClr") {
    const val = clrEl.getAttribute("val");
    if (!val) return null;
    const hex = resolveSchemeKey(val, themeColors);
    return hex ? parseHex(hex) : null;
  }

  if (tag === "sysClr") {
    const parsed = parseHex(clrEl.getAttribute("lastClr") ?? "");
    return parsed;
  }

  return null;
}

/** Apply child tint/shade/lum/alpha modifiers in document order. */
export function applyColorModifiers(base: RgbColor, clrEl: Element): RgbColor {
  let color = base;
  for (const child of Array.from(clrEl.children)) {
    color = applyModifier(color, child);
  }
  return color;
}

export function formatCssColor(color: RgbColor): string {
  const r = Math.round(clamp01(color.r / 255) * 255);
  const g = Math.round(clamp01(color.g / 255) * 255);
  const b = Math.round(clamp01(color.b / 255) * 255);
  const a = clamp01(color.a);

  if (a >= 0.999) {
    return `#${r.toString(16).padStart(2, "0")}${g.toString(16).padStart(2, "0")}${b.toString(16).padStart(2, "0")}`.toUpperCase();
  }
  return `rgba(${r}, ${g}, ${b}, ${Math.round(a * 1000) / 1000})`;
}

/** Find the first color element under el and resolve it with modifiers. */
export function resolveColorFromElement(
  el: Element | null,
  themeColors?: Record<string, string>
): string | undefined {
  if (!el) return undefined;

  for (const tag of ["srgbClr", "schemeClr", "sysClr"] as const) {
    const clrEl = el.localName === tag ? el : el.getElementsByTagNameNS("*", tag)[0];
    if (!clrEl) continue;

    const base = colorFromClrElement(clrEl, themeColors);
    if (!base) continue;

    return formatCssColor(applyColorModifiers(base, clrEl));
  }

  return undefined;
}
