import { applyColorModifiers, colorFromClrElement, formatCssColor } from "./colorModifiers";

export interface DropShadow {
  /** Blur radius in px (a:outerShdw @blurRad). */
  blurRadPx: number;
  /** Offset distance in px (a:outerShdw @dist). */
  distPx: number;
  /** Direction in degrees (a:outerShdw @dir, 60000ths of a degree). */
  dirDeg: number;
  color: string;
  /** Opacity 0–1 (from a:alpha on shadow color). */
  opacity: number;
}

const EMU_PER_PX = 9525;

function emuToPx(emu: number): number {
  return emu / EMU_PER_PX;
}

function parseShadowColor(
  colorEl: Element | null,
  themeColors: Record<string, string>,
): { color: string; opacity: number } {
  if (!colorEl) return { color: "#000000", opacity: 0.5 };

  const base = colorFromClrElement(colorEl, themeColors);
  if (!base) return { color: "#000000", opacity: 0.5 };

  const resolved = applyColorModifiers(base, colorEl);
  return {
    color: formatCssColor({ ...resolved, a: 1 }),
    opacity: resolved.a,
  };
}

/** Parse a:outerShdw from spPr/a:effectLst or any element containing one. */
export function parseOuterShdw(
  container: Element | null | undefined,
  themeColors: Record<string, string>,
): DropShadow | undefined {
  const outerShdw = container?.getElementsByTagNameNS("*", "outerShdw")[0]
    ?? (container?.localName === "outerShdw" ? container : null);
  if (!outerShdw) return undefined;

  const blurRad = Number(outerShdw.getAttribute("blurRad") ?? 0);
  const dist = Number(outerShdw.getAttribute("dist") ?? 0);
  const dir = Number(outerShdw.getAttribute("dir") ?? 0);

  const colorEl =
    outerShdw.getElementsByTagNameNS("*", "srgbClr")[0]
    ?? outerShdw.getElementsByTagNameNS("*", "schemeClr")[0]
    ?? null;
  const { color, opacity } = parseShadowColor(colorEl, themeColors);

  return {
    blurRadPx: emuToPx(blurRad),
    distPx: emuToPx(dist),
    dirDeg: dir / 60000,
    color,
    opacity,
  };
}

/** SVG feDropShadow needs opaque source pixels; use a white silhouette when fill is transparent. */
export function shadowCasterFill(fillColor: string | undefined): string {
  if (fillColor && fillColor !== "transparent" && fillColor !== "none") return fillColor;
  return "white";
}

/** Shadow offset from OOXML direction (0° = east, clockwise). */
export function shadowOffsetPx(shadow: DropShadow): { dx: number; dy: number } {
  const rad = (shadow.dirDeg * Math.PI) / 180;
  return {
    dx: shadow.distPx * Math.cos(rad),
    dy: shadow.distPx * Math.sin(rad),
  };
}

/** CSS box-shadow approximation of a:outerShdw. */
export function shadowToCss(shadow: DropShadow): string {
  const { dx, dy } = shadowOffsetPx(shadow);
  const blur = Math.max(shadow.blurRadPx, 0);
  const spread = 0;
  const rgba = hexToRgba(shadow.color, shadow.opacity);
  return `${dx}px ${dy}px ${blur}px ${spread}px ${rgba}`;
}

function hexToRgba(hex: string, alpha: number): string {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex.trim());
  if (!m) return `rgba(0,0,0,${alpha})`;
  const n = parseInt(m[1]!, 16);
  const r = (n >> 16) & 255;
  const g = (n >> 8) & 255;
  const b = n & 255;
  return `rgba(${r},${g},${b},${alpha})`;
}

/** SVG filter defs for a:outerShdw (for use on shapes/images). */
export function shadowToSvgFilterDef(shadow: DropShadow, filterId: string): string {
  const { dx, dy } = shadowOffsetPx(shadow);
  const stdDev = Math.max(shadow.blurRadPx / 2, 0);
  return `<filter id="${filterId}" x="-50%" y="-50%" width="200%" height="200%">
    <feDropShadow dx="${dx}" dy="${dy}" stdDeviation="${stdDev}" flood-color="${shadow.color}" flood-opacity="${shadow.opacity}" />
  </filter>`;
}
