import { XmlHelper } from "./XmlHelper";

const EMU_PER_PX = 9525;

export interface ListLevelDefaults {
  kind: "p" | "ul" | "ol";
  listStyle?: string;
  align?: string;
  marginLeft?: number;
  indent?: number;
  spaceBefore?: number;
  lineHeight?: number;
  lineHeightPt?: number;
  bulletColor?: string;
  bulletChar?: string;
}

function mapAutoNumToCss(typ: string): string {
  const t = typ.toLowerCase();
  if (t.includes("alphauc")) return "upper-alpha";
  if (t.includes("alphalc")) return "lower-alpha";
  if (t.includes("romanu")) return "upper-roman";
  if (t.includes("romanl")) return "lower-roman";
  return "decimal";
}

/** Parse a:lstStyle lvlNpPr or master txStyle lvlNpPr into list + spacing defaults. */
export function parseLvlPPrElement(
  lvlPPr: Element,
  themeColors: Record<string, string>,
): ListLevelDefaults {
  let kind: "p" | "ul" | "ol" = "p";
  let listStyle: string | undefined;
  let bulletChar: string | undefined;

  if (lvlPPr.querySelector("*|buNone")) {
    kind = "p";
  } else if (lvlPPr.querySelector("*|buAutoNum")) {
    kind = "ol";
    listStyle = mapAutoNumToCss(
      lvlPPr.querySelector("*|buAutoNum")?.getAttribute("type") || "arabicPeriod",
    );
  } else if (lvlPPr.querySelector("*|buChar")) {
    kind = "ul";
    bulletChar = lvlPPr.querySelector("*|buChar")?.getAttribute("char") || undefined;
    listStyle = "disc";
  }

  const align = lvlPPr.getAttribute("algn") || undefined;
  const marLAttr = lvlPPr.getAttribute("marL");
  const marginLeft = marLAttr ? Number(marLAttr) / EMU_PER_PX : undefined;
  const indentAttr = lvlPPr.getAttribute("indent");
  const indent = indentAttr ? Number(indentAttr) / EMU_PER_PX : undefined;

  const spaceBefore = parseParagraphSpacing(lvlPPr, "spcBef");
  const { lineHeight, lineHeightPt } = parseLineSpacing(lvlPPr);

  let bulletColor: string | undefined;
  const buClrEl = lvlPPr.getElementsByTagNameNS("*", "buClr")[0] ?? null;
  if (buClrEl) {
    bulletColor = XmlHelper.getColorFromElement(buClrEl, themeColors) ?? undefined;
  }
  if (!bulletColor && kind !== "p") {
    const defRPr = lvlPPr.querySelector("*|defRPr");
    const solidFill = defRPr?.querySelector("*|solidFill");
    bulletColor = XmlHelper.getColorFromElement(solidFill || null, themeColors) ?? undefined;
  }

  return {
    kind,
    listStyle,
    align,
    marginLeft,
    indent,
    spaceBefore,
    lineHeight,
    lineHeightPt,
    bulletColor,
    bulletChar,
  };
}

/** Read lvl1pPr–lvl9pPr children under a:lstStyle or p:bodyStyle. */
export function parseListLevelDefaults(
  parent: Element | null,
  themeColors: Record<string, string>,
): Record<number, ListLevelDefaults> {
  const out: Record<number, ListLevelDefaults> = {};
  if (!parent) return out;
  for (const child of Array.from(parent.children)) {
    const m = child.localName.match(/^lvl(\d+)pPr$/);
    if (!m) continue;
    const idx = parseInt(m[1], 10) - 1;
    out[idx] = parseLvlPPrElement(child, themeColors);
  }
  return out;
}

/** True when a:lstStyle defines per-level paragraph properties (not just defPPr). */
export function lstStyleDefinesLevels(lstStyle: Element | null): boolean {
  if (!lstStyle) return false;
  return Array.from(lstStyle.children).some((c) => /^lvl\d+pPr$/.test(c.localName));
}

function paragraphHasExplicitBulletOnPPr(pPr: Element | null): boolean {
  if (!pPr) return false;
  return !!(
    pPr.querySelector("*|buNone")
    || pPr.querySelector("*|buChar")
    || pPr.querySelector("*|buAutoNum")
    || pPr.querySelector("*|buFont")
    || pPr.querySelector("*|buClr")
    || pPr.querySelector("*|buSzPct")
  );
}

export interface ParagraphListResolveContext {
  pPr: Element | null;
  level: number;
  shapeLstLevels: Record<number, ListLevelDefaults>;
  shapeLstStyleDefinesLevels: boolean;
  layoutLstLevels?: Record<number, ListLevelDefaults>;
  masterLstLevels?: Record<number, ListLevelDefaults>;
  masterListKey: "title" | "body" | "other";
  ph: Element | null;
  paragraphCount: number;
  maxContentFontSizePt?: number;
}

/**
 * Resolve whether a paragraph is plain text or a list item (DrawingML bullet cascade).
 */
export function resolveParagraphListKind(
  ctx: ParagraphListResolveContext,
  themeColors: Record<string, string>,
): ListLevelDefaults {
  const { pPr, level } = ctx;

  if (pPr?.querySelector("*|buNone")) {
    return { kind: "p" };
  }
  if (pPr?.querySelector("*|buAutoNum")) {
    const buAutoNum = pPr.querySelector("*|buAutoNum");
    return {
      kind: "ol",
      listStyle: mapAutoNumToCss(buAutoNum?.getAttribute("type") || "arabicPeriod"),
      ...readSpacingFromPPr(pPr, themeColors),
    };
  }
  if (pPr?.querySelector("*|buChar")) {
    const buChar = pPr.querySelector("*|buChar");
    return {
      kind: "ul",
      bulletChar: buChar?.getAttribute("char") || undefined,
      listStyle: "disc",
      ...readSpacingFromPPr(pPr, themeColors),
    };
  }
  if (paragraphHasExplicitBulletOnPPr(pPr)) {
    return { kind: "ul", listStyle: "disc", ...readSpacingFromPPr(pPr, themeColors) };
  }

  if (ctx.shapeLstStyleDefinesLevels && ctx.shapeLstLevels[level]) {
    return ctx.shapeLstLevels[level];
  }

  if (ctx.ph && ctx.layoutLstLevels?.[level]) {
    return ctx.layoutLstLevels[level];
  }

  // Single large heading lines are not list bullets unless explicitly marked on pPr.
  if (ctx.paragraphCount === 1 && (ctx.maxContentFontSizePt ?? 0) >= 36) {
    return { kind: "p" };
  }

  if (shouldInheritMasterListLevels(ctx) && ctx.masterLstLevels?.[level]) {
    return ctx.masterLstLevels[level];
  }

  return { kind: "p" };
}

function shouldInheritMasterListLevels(ctx: ParagraphListResolveContext): boolean {
  if (!ctx.ph || !ctx.masterLstLevels) return false;
  if (ctx.masterListKey === "title") return false;
  if (ctx.masterListKey !== "body") return false;

  const type = ctx.ph.getAttribute("type");
  if (type === "body" || type === "obj" || type === "subTitle") return true;

  if (!type && ctx.ph.getAttribute("idx") === "1") {
    if (ctx.layoutLstLevels?.[0]?.kind === "p") return false;
    return true;
  }

  return false;
}

function readSpacingFromPPr(
  pPr: Element,
  themeColors: Record<string, string>,
): Pick<ListLevelDefaults, "align" | "marginLeft" | "indent" | "spaceBefore" | "lineHeight" | "lineHeightPt" | "bulletColor"> {
  const marLAttr = pPr.getAttribute("marL");
  const marginLeft = marLAttr ? Number(marLAttr) / EMU_PER_PX : undefined;
  const indentAttr = pPr.getAttribute("indent");
  const indent = indentAttr ? Number(indentAttr) / EMU_PER_PX : undefined;
  const align = pPr.getAttribute("algn") || undefined;
  const spaceBefore = parseParagraphSpacing(pPr, "spcBef");
  const { lineHeight, lineHeightPt } = parseLineSpacing(pPr);

  let bulletColor: string | undefined;
  const buClrEl = pPr.getElementsByTagNameNS("*", "buClr")[0] ?? null;
  if (buClrEl) {
    bulletColor = XmlHelper.getColorFromElement(buClrEl, themeColors) ?? undefined;
  }

  return { align, marginLeft, indent, spaceBefore, lineHeight, lineHeightPt, bulletColor };
}

function parseParagraphSpacing(pPr: Element | null, tag: "spcBef" | "spcAft"): number | undefined {
  if (!pPr) return undefined;
  const spcEl = pPr.getElementsByTagNameNS("*", tag)[0] ?? null;
  if (!spcEl) return undefined;
  const spcPts = spcEl.getElementsByTagNameNS("*", "spcPts")[0] ?? null;
  if (spcPts) {
    const v = Number(spcPts.getAttribute("val") || 0);
    if (Number.isFinite(v) && v > 0) return v / 100;
  }
  const spcPct = spcEl.getElementsByTagNameNS("*", "spcPct")[0] ?? null;
  if (spcPct) {
    const v = Number(spcPct.getAttribute("val") || 0);
    if (Number.isFinite(v) && v > 0) return (v / 100000) * 12;
  }
  return undefined;
}

const POWERPOINT_SINGLE_LINE_FACTOR = 1.2;

function parseLineSpacing(pr: Element | null): { lineHeight?: number; lineHeightPt?: number } {
  if (!pr) return {};
  const lnSpc = pr.getElementsByTagNameNS("*", "lnSpc")[0] ?? null;
  if (!lnSpc) return {};
  const spcPts = lnSpc.getElementsByTagNameNS("*", "spcPts")[0] ?? null;
  if (spcPts) {
    const v = Number(spcPts.getAttribute("val") || 0);
    if (Number.isFinite(v) && v > 0) return { lineHeightPt: v / 100 };
  }
  const spcPct = lnSpc.getElementsByTagNameNS("*", "spcPct")[0] ?? null;
  if (spcPct) {
    const v = Number(spcPct.getAttribute("val") || 0);
    if (Number.isFinite(v) && v > 0) {
      return { lineHeight: (v / 100000) * POWERPOINT_SINGLE_LINE_FACTOR };
    }
  }
  return {};
}

/**
 * Map OOXML paragraph marL + indent to list padding (hanging bullet layout).
 * Text starts at marL; bullet sits at marL + indent (often indent = -marL).
 */
/** Minimum ul padding so outside markers are not clipped by overflow:hidden parents. */
export const MIN_BULLET_GUTTER_PX = 24;

export function listBulletLayoutCss(
  marginLeft?: number,
  indent?: number,
): { ulPaddingLeft: number; liPaddingLeft: number } {
  const ml = marginLeft ?? 0;
  const ind = indent ?? 0;
  let ulPaddingLeft = Math.max(0, ml + ind);
  let liPaddingLeft = Math.max(0, ml - ulPaddingLeft);
  if (ulPaddingLeft === 0 && liPaddingLeft === 0) {
    ulPaddingLeft = MIN_BULLET_GUTTER_PX;
  }
  return { ulPaddingLeft, liPaddingLeft };
}
