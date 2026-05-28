import { ShapeElement } from "../models/SlideElement";
import { XmlHelper } from "./XmlHelper";
import { dashStyleFromPrst } from "./lineStyle";
import { DropShadow, parseOuterShdw } from "./shadowEffect";

export interface ExtractedShapeStyle {
  shapeType: string;
  fillColor: string;
  borderColor?: string;
  strokeWidth?: number;
  dashStyle?: ShapeElement["dashStyle"];
  headEnd?: { type?: string; w?: string; len?: string };
  tailEnd?: { type?: string; w?: string; len?: string };
  shadow?: DropShadow;
}

function spPrHasDirectNoFill(spPr: Element | null | undefined): boolean {
  if (!spPr) return false;
  return Array.from(spPr.children).some((child) => child.localName === "noFill");
}

function lineHasNoFill(ln: Element | null): boolean {
  return !!ln?.getElementsByTagNameNS("*", "noFill")[0];
}

function readLnRefIndex(owner: Element): number {
  const lnRef = owner.getElementsByTagNameNS("*", "style")[0]
    ?.getElementsByTagNameNS("*", "lnRef")[0] ?? null;
  return parseInt(lnRef?.getAttribute("idx") || "0", 10);
}

function readLineWidthPx(ln: Element | null): number | undefined {
  if (!ln) return undefined;
  const wAttr = ln.getAttribute("w");
  if (wAttr == null || wAttr === "") return undefined;
  const w = Number(wAttr);
  return Number.isFinite(w) ? w / 9525 : undefined;
}

function readLineColorFromLn(
  ln: Element | null,
  themeColors: Record<string, string>,
): string | undefined {
  if (!ln || lineHasNoFill(ln)) return undefined;

  const solidFill = ln.getElementsByTagNameNS("*", "solidFill")[0] ?? null;
  const solid = XmlHelper.getColorFromElement(solidFill, themeColors);
  if (solid && solid !== "transparent") return solid;

  const gradFill = ln.getElementsByTagNameNS("*", "gradFill")[0] ?? null;
  if (gradFill) {
    for (const stop of Array.from(gradFill.getElementsByTagNameNS("*", "gs"))) {
      const c = XmlHelper.getColorFromElement(stop, themeColors);
      if (c && c !== "transparent") return c;
    }
  }

  return undefined;
}

function readLineDecorations(ln: Element | null): Pick<ExtractedShapeStyle, "headEnd" | "tailEnd" | "dashStyle"> {
  if (!ln || lineHasNoFill(ln)) return {};

  const headEndEl = ln.getElementsByTagNameNS("*", "headEnd")[0] ?? null;
  const tailEndEl = ln.getElementsByTagNameNS("*", "tailEnd")[0] ?? null;
  const prstDash = ln.getElementsByTagNameNS("*", "prstDash")[0] ?? null;

  return {
    headEnd: headEndEl
      ? {
          type: headEndEl.getAttribute("type") || undefined,
          w: headEndEl.getAttribute("w") || undefined,
          len: headEndEl.getAttribute("len") || undefined,
        }
      : undefined,
    tailEnd: tailEndEl
      ? {
          type: tailEndEl.getAttribute("type") || undefined,
          w: tailEndEl.getAttribute("w") || undefined,
          len: tailEndEl.getAttribute("len") || undefined,
        }
      : undefined,
    dashStyle: dashStyleFromPrst(prstDash?.getAttribute("val")),
  };
}

function readPresetShapeType(spPr: Element | null | undefined, fallback = "rect"): string {
  const prstGeom = spPr?.getElementsByTagNameNS("*", "prstGeom")[0];
  return prstGeom?.getAttribute("prst") ?? fallback;
}

/** PowerPoint preset default: adj val 16667 → ~16.67% of the shorter side. */
export const DEFAULT_ROUND_RECT_ADJ = 16667 / 100000;

/**
 * Arc corner radius in px for a roundRect preset.
 * DrawingML uses adj as a fraction of min(width, height); SVG rx is the same arc radius.
 */
export function roundRectCornerRadiusPx(
  widthPx: number,
  heightPx: number,
  adjFraction?: number,
): number {
  const shortSide = Math.min(widthPx, heightPx);
  if (shortSide <= 0) return 0;
  const fraction = adjFraction && adjFraction > 0 ? adjFraction : DEFAULT_ROUND_RECT_ADJ;
  return Math.min(shortSide * fraction, shortSide / 2);
}

/** roundRect corner radius as a fraction of the shorter side (adj / 100000). */
export function readRoundRectAdj(spPr: Element | null | undefined): number | undefined {
  const avLst = spPr?.getElementsByTagNameNS("*", "prstGeom")[0]
    ?.getElementsByTagNameNS("*", "avLst")[0];
  const gd = avLst?.getElementsByTagNameNS("*", "gd")[0];
  const fmla = gd?.getAttribute("fmla") ?? "";
  const m = /val\s+(\d+)/.exec(fmla);
  if (!m) return undefined;
  const n = Number(m[1]);
  return Number.isFinite(n) ? n / 100000 : undefined;
}

/**
 * Extract fill, outline, and preset geometry from a shape/picture spPr.
 * When pictureFill is true, theme fillRef is not used as a fallback (blip is the fill).
 */
export function extractShapeStyle(
  owner: Element,
  spPr: Element | null | undefined,
  themeColors: Record<string, string>,
  themeDoc: Document | null = null,
  opts: { pictureFill?: boolean; shapeType?: string } = {},
): ExtractedShapeStyle {
  const shapeType = opts.shapeType ?? readPresetShapeType(spPr);
  let fillColor = "transparent";
  let borderColor: string | undefined;
  let strokeWidth: number | undefined;
  let headEnd: ExtractedShapeStyle["headEnd"];
  let tailEnd: ExtractedShapeStyle["tailEnd"];
  let dashStyle: ExtractedShapeStyle["dashStyle"];

  const ln = spPr?.getElementsByTagNameNS("*", "ln")[0] ?? null;
  const fillExplicitlyNone = spPrHasDirectNoFill(spPr);
  const lineExplicitlyNone = lineHasNoFill(ln);
  const lnRefIdx = readLnRefIndex(owner);
  const themeLn = lnRefIdx > 0 && themeDoc ? XmlHelper.getThemeLineElement(themeDoc, lnRefIdx) : null;

  if (spPr) {
    if (fillExplicitlyNone) {
      fillColor = "transparent";
    } else {
      const solidFill = spPr.getElementsByTagNameNS("*", "solidFill")[0] ?? null;
      const gradFill = spPr.getElementsByTagNameNS("*", "gradFill")[0] ?? null;
      fillColor =
        XmlHelper.getColorFromElement(solidFill, themeColors)
        ?? XmlHelper.getColorFromElement(gradFill?.getElementsByTagNameNS("*", "gs")[0] ?? null, themeColors)
        ?? "transparent";
    }

    if (lineExplicitlyNone) {
      borderColor = undefined;
      strokeWidth = undefined;
    } else {
      borderColor =
        readLineColorFromLn(ln, themeColors)
        ?? (lnRefIdx > 0 ? XmlHelper.resolveStyleRefColor(owner, "lnRef", themeColors, themeDoc) : undefined)
        ?? readLineColorFromLn(themeLn, themeColors);

      strokeWidth = readLineWidthPx(ln) ?? readLineWidthPx(themeLn);

      const decorations = ln && !lineHasNoFill(ln)
        ? readLineDecorations(ln)
        : readLineDecorations(themeLn);
      headEnd = decorations.headEnd;
      tailEnd = decorations.tailEnd;
      dashStyle = decorations.dashStyle;
    }
  }

  if (!opts.pictureFill && !fillExplicitlyNone && fillColor === "transparent") {
    fillColor =
      XmlHelper.resolveStyleRefColor(owner, "fillRef", themeColors, themeDoc)
      ?? fillColor;
  }

  if (!opts.pictureFill && !fillExplicitlyNone && fillColor === "transparent") {
    const style = owner.getElementsByTagNameNS("*", "style")[0];
    const fillRef = style?.getElementsByTagNameNS("*", "fillRef")[0];
    fillColor = XmlHelper.getColorFromElement(fillRef ?? null, themeColors) ?? fillColor;
  }

  const shadowFromSpPr = parseOuterShdw(spPr, themeColors);
  let shadow = shadowFromSpPr;
  if (!shadow && themeDoc) {
    const style = owner.getElementsByTagNameNS("*", "style")[0];
    const effectRef = style?.getElementsByTagNameNS("*", "effectRef")[0];
    const idx = parseInt(effectRef?.getAttribute("idx") || "0", 10);
    if (idx > 0) {
      const effectStyle = XmlHelper.getThemeEffectStyleElement(themeDoc, idx);
      shadow = parseOuterShdw(effectStyle, themeColors);
    }
  }

  return {
    shapeType,
    fillColor,
    borderColor,
    strokeWidth,
    dashStyle,
    headEnd,
    tailEnd,
    shadow,
  };
}

/** True when spPr carries a visible preset frame (fill, outline, shadow, or non-rect clip). */
export function hasVisibleShapeFrame(style: ExtractedShapeStyle): boolean {
  return (
    style.shapeType !== "rect"
    || style.fillColor !== "transparent"
    || style.borderColor !== undefined
    || (style.strokeWidth !== undefined && style.strokeWidth > 0)
    || style.shadow !== undefined
  );
}
