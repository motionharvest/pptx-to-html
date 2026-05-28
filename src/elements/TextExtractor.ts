import { TextElement } from "../models/SlideElement";
import { XmlHelper } from "../core/XmlHelper";
import { textRunInlineStyles } from "../core/scriptTextStyle";
import { parseTextDirection } from "../core/textDirection";
import {
  listBulletLayoutCss,
  lstStyleDefinesLevels,
  parseListLevelDefaults,
  resolveParagraphListKind,
  type ListLevelDefaults,
} from "../core/listLevelDefaults";
import { splitTextByUrls } from "../core/textUrls";

export interface PlaceholderDefaults {
  x?: number;
  y?: number;
  cx?: number;
  cy?: number;
  anchor?: string;
  vert?: string;
  lIns?: string;
  tIns?: string;
  rIns?: string;
  bIns?: string;
  fontSize?: number;
  fontFamily?: string;
  bold?: boolean;
  italic?: boolean;
  color?: string;
  align?: string;
  /** Placeholder type from layout/master (when slide ph omits type). */
  placeholderType?: string;
  /** Per-level list defaults from layout placeholder lstStyle. */
  listLevels?: Record<number, ListLevelDefaults>;
}

interface TextRunDefaults {
  bold?: boolean;
  italic?: boolean;
  color?: string;
  fontSize?: number;
  fontFamily?: string;
}

interface MasterStyleDefaults extends TextRunDefaults {
  align?: string;
  listLevels?: Record<number, ListLevelDefaults>;
}

export type MasterTextStyleKey = "title" | "body" | "other";
export type MasterTextStyles = Partial<Record<MasterTextStyleKey, MasterStyleDefaults>>;

const PLACEHOLDER_TYPE_ALIASES: Record<string, string[]> = {
  ctrTitle: ["title"],
  title: ["ctrTitle"],
};

interface ParsedRun {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  superscript?: boolean;
  subscript?: boolean;
  color?: string;
  fontSize?: number;
  fontFamily?: string;
  isBreak?: boolean;
}

interface ParsedParagraph {
  runs: ParsedRun[];
  align?: string;
  level: number;
  listKind: "p" | "ul" | "ol";
  listStyle?: string;
  spaceBefore?: number;
  spaceAfter?: number;
  lineHeight?: number;
  lineHeightPt?: number;
  bulletColor?: string;
  marginLeft?: number;
  indent?: number;
}

export class TextExtractor {
  static extract(spTree: Element | null, themeColors: Record<string, string>, opts: { context?: "slide" | "layout" | "master"; placeholderDefaults?: Map<string, PlaceholderDefaults>; masterTextStyles?: MasterTextStyles; preserveTextStructure?: boolean } = {}): TextElement[] {
    if (!spTree) return [];
    const elements: TextElement[] = [];
    const shapes = spTree.getElementsByTagNameNS("*", "sp");
    for (const shape of Array.from(shapes)) {
      const el = this.extractFromSp(shape, themeColors, opts);
      if (el) elements.push(el);
    }
    return elements;
  }

  static extractFromSp(
    shape: Element,
    themeColors: Record<string, string>,
    opts: { context?: "slide" | "layout" | "master"; placeholderDefaults?: Map<string, PlaceholderDefaults>; masterTextStyles?: MasterTextStyles; preserveTextStructure?: boolean } = {},
  ): TextElement | null {
      const nvPr = shape.getElementsByTagNameNS("*", "nvPr")[0] ?? null;
      const ph = nvPr?.getElementsByTagNameNS("*", "ph")[0] ?? null;
      if (opts.context && opts.context !== "slide" && ph) return null;
      const txBody = shape.getElementsByTagNameNS("*", "txBody")[0];
      if (!txBody) return null;

      // Resolve placeholder defaults for this shape
      const phDefaults = resolvePlaceholder(ph, opts.placeholderDefaults, opts.masterTextStyles);

      const paragraphs = txBody.getElementsByTagNameNS("*", "p");
      const bodyPr = txBody.getElementsByTagNameNS("*", "bodyPr")[0] ?? null;
      const textDirection = parseTextDirection(bodyPr, phDefaults?.vert);
      const anchor = bodyPr?.getAttribute("anchor") || phDefaults?.anchor || undefined;
      const verticalAlign = anchor === "ctr" ? "middle" : anchor === "b" ? "bottom" : "top";
      const lIns = bodyPr?.getAttribute("lIns") ?? phDefaults?.lIns;
      const tIns = bodyPr?.getAttribute("tIns") ?? phDefaults?.tIns;
      const rIns = bodyPr?.getAttribute("rIns") ?? phDefaults?.rIns;
      const bIns = bodyPr?.getAttribute("bIns") ?? phDefaults?.bIns;
      const padding = {
        left: lIns ? Number(lIns) / 9525 : 0,
        top: tIns ? Number(tIns) / 9525 : 0,
        right: rIns ? Number(rIns) / 9525 : 0,
        bottom: bIns ? Number(bIns) / 9525 : 0,
      };

      let bulletChar: string | undefined;
      const lvlDefaults: Record<number, ListLevelDefaults> = {};
      const lstStyle = txBody.querySelector("*|lstStyle");
      const shapeLstStyleDefinesLevels = lstStyleDefinesLevels(lstStyle);
      if (shapeLstStyleDefinesLevels && lstStyle) {
        const shapeLevels = parseListLevelDefaults(lstStyle, themeColors);
        for (const [idx, lvl] of Object.entries(shapeLevels)) {
          lvlDefaults[Number(idx)] = lvl;
          if (lvl.bulletChar) bulletChar = lvl.bulletChar;
        }
      }

      const effectivePhType = ph?.getAttribute("type") || phDefaults?.placeholderType || undefined;
      const masterListKey = masterStyleKeyForPlaceholder(
        effectivePhType,
        ph?.getAttribute("idx"),
      );
      const masterListLevels = opts.masterTextStyles?.[masterListKey]?.listLevels;
      const layoutListLevels = phDefaults?.listLevels;

      // Text run defaults cascade: placeholder → shape-level fallbacks
      const textDefaults: TextRunDefaults = {
        bold: phDefaults?.bold,
        italic: phDefaults?.italic,
        color: phDefaults?.color,
        fontSize: phDefaults?.fontSize,
        fontFamily: phDefaults?.fontFamily,
      };

      const contentParagraphCount = Array.from(paragraphs).filter((p) => {
        const runs = extractRunsFromParagraph(p, themeColors, textDefaults);
        return runs.some((r) => !r.isBreak && r.text.trim());
      }).length;

      const parsedParagraphs: ParsedParagraph[] = [];
      let horizontalAlign: "left" | "center" | "right" | "justify" | undefined;
      let defaultFontName = phDefaults?.fontFamily || "Arial";
      let defaultFontSize = phDefaults?.fontSize || 18;
      let fontDefaultsCaptured = !!phDefaults?.fontFamily || !!phDefaults?.fontSize;

      for (const p of Array.from(paragraphs)) {
        const pPr = p.getElementsByTagNameNS("*", "pPr")[0] ?? null;
        const lvlAttr = pPr?.getAttribute("lvl");
        const lvl = lvlAttr ? parseInt(lvlAttr, 10) : 0;
        const algn =
          pPr?.getAttribute("algn")
          || lvlDefaults[lvl]?.align
          || masterListLevels?.[lvl]?.align
          || undefined;

        const runs = extractRunsFromParagraph(p, themeColors, textDefaults);

        if (!fontDefaultsCaptured) {
          for (const run of runs) {
            if (run.isBreak || !run.text.trim()) continue;
            if (run.fontFamily) defaultFontName = run.fontFamily;
            if (run.fontSize) defaultFontSize = run.fontSize;
            fontDefaultsCaptured = true;
            break;
          }
        }

        let maxContentFontSizePt = 0;
        for (const run of runs) {
          if (run.fontSize && run.fontSize > maxContentFontSizePt) maxContentFontSizePt = run.fontSize;
        }
        if (!maxContentFontSizePt && phDefaults?.fontSize) maxContentFontSizePt = phDefaults.fontSize;

        const listResolved = resolveParagraphListKind(
          {
            pPr,
            level: lvl,
            shapeLstLevels: lvlDefaults,
            shapeLstStyleDefinesLevels,
            layoutLstLevels: layoutListLevels,
            masterLstLevels: masterListLevels,
            masterListKey,
            ph,
            paragraphCount: contentParagraphCount,
            maxContentFontSizePt,
          },
          themeColors,
        );
        const kind = listResolved.kind;
        const listStyle = listResolved.listStyle;
        if (listResolved.bulletChar) bulletChar = listResolved.bulletChar;

        let spaceBefore = parseParagraphSpacing(pPr, "spcBef");
        const spaceAfter = parseParagraphSpacing(pPr, "spcAft");
        let { lineHeight, lineHeightPt } = parseLineSpacing(pPr);

        const marLAttr = pPr?.getAttribute("marL");
        let marginLeft = marLAttr ? Number(marLAttr) / 9525 : undefined;
        const indentAttr = pPr?.getAttribute("indent");
        let indent = indentAttr ? Number(indentAttr) / 9525 : undefined;

        const inheritedLvl = listResolved;
        if (marginLeft == null && inheritedLvl.marginLeft != null) marginLeft = inheritedLvl.marginLeft;
        if (indent == null && inheritedLvl.indent != null) indent = inheritedLvl.indent;
        if (!spaceBefore && inheritedLvl.spaceBefore) spaceBefore = inheritedLvl.spaceBefore;
        if (!lineHeight && inheritedLvl.lineHeight) lineHeight = inheritedLvl.lineHeight;
        if (!lineHeightPt && inheritedLvl.lineHeightPt) lineHeightPt = inheritedLvl.lineHeightPt;

        let bulletColor: string | undefined;
        const buClrEl = pPr?.getElementsByTagNameNS("*", "buClr")[0] ?? null;
        if (buClrEl) {
          bulletColor = XmlHelper.getColorFromElement(buClrEl, themeColors) ?? undefined;
        }
        if (!bulletColor && kind !== "p") {
          bulletColor = inheritedLvl.bulletColor;
        }
        if (!bulletColor && kind !== "p") {
          const firstContentRun = runs.find((r) => !r.isBreak && r.text.trim());
          bulletColor = firstContentRun?.color;
        }

        parsedParagraphs.push({ runs, align: algn, level: isNaN(lvl) ? 0 : lvl, listKind: kind, listStyle, spaceBefore, spaceAfter, lineHeight, lineHeightPt, bulletColor, marginLeft, indent });
      }

      // Shape-level align: only when all paragraphs agree; mixed → left + per-paragraph html
      horizontalAlign = resolveShapeHorizontalAlign(parsedParagraphs, phDefaults, ph);

      // Resolve default color: slide lstStyle → defRPr → shape fill → placeholder → black
      let defaultColor: string | undefined;
      const lstDefRPr = txBody.querySelector("*|lstStyle *|defRPr");
      const lstFill = lstDefRPr?.querySelector("*|solidFill");
      defaultColor = XmlHelper.getColorFromElement(lstFill || null, themeColors);
      if (!defaultColor) {
        const defRPr = txBody.querySelector("*|defRPr");
        const defFill = defRPr?.querySelector("*|solidFill");
        defaultColor = XmlHelper.getColorFromElement(defFill || null, themeColors);
      }
      if (!defaultColor) {
        const spPr = shape.querySelector("p\\:spPr, spPr");
        const shapeFill = spPr?.querySelector("*|solidFill");
        defaultColor = XmlHelper.getColorFromElement(shapeFill || null, themeColors);
      }
      if (!defaultColor && phDefaults?.color) {
        defaultColor = phDefaults.color;
      }

      // Build plain text content (preserve Shift+Enter breaks and blank lines)
      const textParts = parsedParagraphs.map(paragraphToPlainText);
      const content = opts.preserveTextStructure
        ? textParts.join("\n")
        : textParts.map((t) => t.replace(/\n+/g, " ").trim()).filter((t) => t.length > 0).join(" ").trim();
      if (!content.trim() && !textBoxHasVisibleContent(parsedParagraphs)) return null;

      // Build segments with per-run formatting
      const segments: NonNullable<TextElement["segments"]> = [];
      for (let pi = 0; pi < parsedParagraphs.length; pi++) {
        const para = parsedParagraphs[pi];
        let isFirstContentRun = true;
        for (const run of para.runs) {
          if (run.isBreak) {
            segments.push({ text: "\n", breakBefore: true });
            continue;
          }
          if (!run.text) continue;
          segments.push({
            text: run.text,
            bold: run.bold || undefined,
            italic: run.italic || undefined,
            underline: run.underline || undefined,
            superscript: run.superscript || undefined,
            subscript: run.subscript || undefined,
            color: run.color,
            fontSize: run.fontSize,
            fontFamily: run.fontFamily,
            paragraphBreakBefore: pi > 0 && isFirstContentRun,
          });
          isFirstContentRun = false;
        }
      }

      const richHtml = buildRichHtml(parsedParagraphs, bulletChar, defaultFontName, defaultFontSize, defaultColor, phDefaults, ph);

      // Geometry: slide xfrm → placeholder defaults → fallback
      const xfrm = shape.getElementsByTagNameNS("*", "xfrm")[0];
      const off = xfrm?.getElementsByTagNameNS("*", "off")[0] ?? null;
      const ext = xfrm?.getElementsByTagNameNS("*", "ext")[0] ?? null;

      let x: number, y: number, cx: number, cy: number;
      if (off && ext) {
        x = XmlHelper.getAttrAsNumber(off, "x");
        y = XmlHelper.getAttrAsNumber(off, "y");
        cx = XmlHelper.getAttrAsNumber(ext, "cx");
        cy = XmlHelper.getAttrAsNumber(ext, "cy");
      } else if (phDefaults?.cx != null) {
        x = phDefaults.x ?? 0;
        y = phDefaults.y ?? 0;
        cx = phDefaults.cx;
        cy = phDefaults.cy ?? 500000;
      } else {
        x = 0; y = 0; cx = 1000000; cy = 500000;
      }
      const bodyLineSpacing = parseLineSpacing(bodyPr);
      const paraLineSpacing = parsedParagraphs.find((p) => p.lineHeight || p.lineHeightPt);
      const lineHeight = bodyLineSpacing.lineHeight ?? paraLineSpacing?.lineHeight;
      const lineHeightPt = bodyLineSpacing.lineHeightPt ?? paraLineSpacing?.lineHeightPt;
      return {
        type: "text",
        content,
        position: { x, y },
        size: { width: cx, height: cy },
        font: { name: defaultFontName, size: defaultFontSize, color: defaultColor || "#000000" },
        align: { horizontal: horizontalAlign ?? "left", vertical: verticalAlign },
        textDirection,
        padding,
        html: richHtml,
        segments,
        lineHeight,
        lineHeightPt,
      };
  }

  /** Extract default paragraph/run properties from slide master txStyles (title/body/other). */
  static extractMasterTextStyles(masterDoc: Document | null, themeColors: Record<string, string>): MasterTextStyles {
    if (!masterDoc) return {};
    const styles: MasterTextStyles = {};
    const mapping: Array<[MasterTextStyleKey, string]> = [
      ["title", "titleStyle"],
      ["body", "bodyStyle"],
      ["other", "otherStyle"],
    ];
    for (const [key, tag] of mapping) {
      const styleEl = masterDoc.getElementsByTagNameNS("*", tag)[0] ?? null;
      const lvl1pPr = styleEl?.querySelector("*|lvl1pPr");
      const defRPr = lvl1pPr?.querySelector("*|defRPr") ?? styleEl?.querySelector("*|defPPr *|defRPr");
      const defaults: MasterStyleDefaults = defRPr ? parseDefRPrDefaults(defRPr, themeColors) : {};
      const algn = lvl1pPr?.getAttribute("algn");
      if (algn) defaults.align = algn;
      const listLevels = parseListLevelDefaults(styleEl, themeColors);
      if (Object.keys(listLevels).length > 0) defaults.listLevels = listLevels;
      if (Object.keys(defaults).length > 0) styles[key] = defaults;
    }
    return styles;
  }
}

const IMPLICIT_PLACEHOLDER_ALIGN: Record<string, string> = {
  title: "ctr",
  ctrTitle: "ctr",
  subTitle: "ctr",
};

/**
 * Match a slide placeholder to its layout/master definition.
 * Tries idx first (more specific), then type (well-known placeholders like title/body).
 */
function masterStyleKeyForPlaceholder(type?: string | null, idx?: string | null): MasterTextStyleKey {
  if (!type) {
    // idx 1 is the standard content/body placeholder when type is omitted
    if (idx === "1") return "body";
    return "other";
  }
  if (type === "title" || type === "ctrTitle") return "title";
  if (type === "body" || type === "obj" || type === "subTitle" || type === "dt") return "body";
  return "other";
}

function lookupPlaceholderByType(
  type: string | null | undefined,
  defaults?: Map<string, PlaceholderDefaults>,
): PlaceholderDefaults | undefined {
  if (!type || !defaults) return undefined;
  let matched = defaults.get(`type:${type}`);
  if (!matched) {
    for (const alias of PLACEHOLDER_TYPE_ALIASES[type] || []) {
      matched = defaults.get(`type:${alias}`);
      if (matched) break;
    }
  }
  return matched;
}

/** Layout idx can match while geometry lives on the master's type entry — inherit missing fields. */
function mergePlaceholderDefaults(
  base?: PlaceholderDefaults,
  override?: PlaceholderDefaults,
): PlaceholderDefaults | undefined {
  if (!base && !override) return undefined;
  const merged: PlaceholderDefaults = { ...base, ...override };
  const inheritKeys: Array<keyof PlaceholderDefaults> = [
    "x", "y", "cx", "cy", "anchor", "lIns", "tIns", "rIns", "bIns", "align",
    "fontSize", "fontFamily", "bold", "italic", "color", "placeholderType", "listLevels",
  ];
  for (const key of inheritKeys) {
    if (merged[key] == null && base?.[key] != null) {
      (merged as any)[key] = base[key];
    }
  }
  return merged;
}

function resolvePlaceholder(
  ph: Element | null,
  defaults?: Map<string, PlaceholderDefaults>,
  masterTextStyles?: MasterTextStyles,
): PlaceholderDefaults | undefined {
  if (!ph) return undefined;

  const idx = ph.getAttribute("idx");
  const type = ph.getAttribute("type");
  const byType = lookupPlaceholderByType(type, defaults);
  const byIdx = idx && defaults ? defaults.get(`idx:${idx}`) : undefined;
  const matched = mergePlaceholderDefaults(byType, byIdx);
  const effectiveType = type || matched?.placeholderType;

  const styleDefaults = masterTextStyles?.[masterStyleKeyForPlaceholder(effectiveType, idx)];
  const runDefaults: PlaceholderDefaults | undefined = styleDefaults
    ? {
        bold: styleDefaults.bold,
        italic: styleDefaults.italic,
        color: styleDefaults.color,
        fontSize: styleDefaults.fontSize,
        fontFamily: styleDefaults.fontFamily,
      }
    : undefined;
  if (runDefaults || matched) {
    return {
      ...runDefaults,
      ...matched,
      placeholderType: effectiveType || matched?.placeholderType,
      listLevels: matched?.listLevels,
    };
  }
  return undefined;
}

function mapOoxmlAlign(algn: string): "left" | "center" | "right" | "justify" {
  if (algn === "ctr") return "center";
  if (algn === "r") return "right";
  if (algn.startsWith("just")) return "justify";
  return "left";
}

function parseDefRPrDefaults(defRPr: Element, themeColors: Record<string, string>): TextRunDefaults {
  const defaults: TextRunDefaults = {};
  const sz = defRPr.getAttribute("sz");
  if (sz) {
    const n = parseInt(sz, 10);
    if (Number.isFinite(n)) defaults.fontSize = n / 100;
  }
  if (defRPr.getAttribute("b") === "1") defaults.bold = true;
  if (defRPr.getAttribute("i") === "1") defaults.italic = true;
  const solidFill = defRPr.querySelector("*|solidFill");
  const color = XmlHelper.getColorFromElement(solidFill || null, themeColors);
  if (color) defaults.color = color;
  const latin = defRPr.getElementsByTagNameNS("*", "latin")[0];
  const fontFamily = latin?.getAttribute("typeface");
  if (fontFamily) defaults.fontFamily = fontFamily;
  return defaults;
}

function extractRunsFromParagraph(p: Element, themeColors: Record<string, string>, textDefaults?: TextRunDefaults): ParsedRun[] {
  const runs: ParsedRun[] = [];
  for (const child of Array.from(p.childNodes) as any[]) {
    if (!(child instanceof Element)) continue;
    const ln = child.localName;
    if (ln === "r") {
      runs.push(parseTextRun(child, themeColors, textDefaults));
    } else if (ln === "br") {
      runs.push({ text: "\n", isBreak: true });
    } else if (ln === "fld") {
      const t = child.getElementsByTagNameNS("*", "t")[0]?.textContent ?? "";
      if (t) {
        const rPr = child.getElementsByTagNameNS("*", "rPr")[0] ?? null;
        runs.push(buildRunFromProps(t, rPr, themeColors, textDefaults));
      }
    } else if (ln === "tab") {
      runs.push({ text: "\t" });
    }
  }
  return runs;
}

function parseTextRun(r: Element, themeColors: Record<string, string>, textDefaults?: TextRunDefaults): ParsedRun {
  const rPr = r.getElementsByTagNameNS("*", "rPr")[0] ?? null;
  const text = r.getElementsByTagNameNS("*", "t")[0]?.textContent ?? "";
  return buildRunFromProps(text, rPr, themeColors, textDefaults);
}

/**
 * Extract formatting from a run's <a:rPr>, falling back to inherited defaults
 * for any property the run doesn't explicitly set.
 */
function buildRunFromProps(text: string, rPr: Element | null, themeColors: Record<string, string>, textDefaults?: TextRunDefaults): ParsedRun {
  if (!rPr) {
    return {
      text,
      bold: textDefaults?.bold || undefined,
      italic: textDefaults?.italic || undefined,
      color: textDefaults?.color,
      fontSize: textDefaults?.fontSize,
      fontFamily: textDefaults?.fontFamily,
    };
  }

  // Distinguish "not set" (null) from "explicitly off" ("0") vs "on" ("1")
  const bAttr = rPr.getAttribute("b");
  const bold = bAttr === "1" ? true : bAttr === "0" ? false : textDefaults?.bold;

  const iAttr = rPr.getAttribute("i");
  const italic = iAttr === "1" ? true : iAttr === "0" ? false : textDefaults?.italic;

  const uVal = rPr.getAttribute("u");
  let underline = uVal && uVal !== "none" ? true : undefined;

  const hasHyperlink = !!(
    rPr.querySelector("*|hlinkClick") || rPr.querySelector("*|hlinkMouseOver")
  );
  if (hasHyperlink) underline = true;

  const solidFill = rPr.querySelector("*|solidFill");
  let color = XmlHelper.getColorFromElement(solidFill || null, themeColors) ?? textDefaults?.color;
  if (hasHyperlink && !solidFill && themeColors.hlink) {
    color = themeColors.hlink;
  }

  let fontSize: number | undefined;
  const sz = rPr.getAttribute("sz");
  if (sz) {
    const n = parseInt(sz, 10);
    if (Number.isFinite(n)) fontSize = n / 100;
  }
  fontSize = fontSize ?? textDefaults?.fontSize;

  const latin = rPr.getElementsByTagNameNS("*", "latin")[0];
  const fontFamily = latin?.getAttribute("typeface") || textDefaults?.fontFamily || undefined;

  const { superscript, subscript } = parseBaselineOffset(rPr.getAttribute("baseline"));

  return { text, bold: bold || undefined, italic: italic || undefined, underline, superscript, subscript, color, fontSize, fontFamily };
}

/** a:rPr @baseline — positive raises (superscript), negative lowers (subscript). */
function parseBaselineOffset(baseline: string | null): { superscript?: boolean; subscript?: boolean } {
  if (!baseline) return {};
  const n = parseInt(baseline, 10);
  if (!Number.isFinite(n) || n === 0) return {};
  return n > 0 ? { superscript: true } : { subscript: true };
}

function paragraphHasText(para: ParsedParagraph): boolean {
  return para.runs.some((r) => !r.isBreak && r.text.trim().length > 0);
}

function paragraphHasLineBreaks(para: ParsedParagraph): boolean {
  return para.runs.some((r) => r.isBreak);
}

/** Empty <a:p> used as vertical spacing (Enter), not the default lone empty paragraph. */
function paragraphIsEmptySpacer(para: ParsedParagraph, all: ParsedParagraph[]): boolean {
  if (para.runs.length > 0) return false;
  return all.some((p) => p !== para && (paragraphHasText(p) || paragraphHasLineBreaks(p)));
}

function paragraphHasSpacing(para: ParsedParagraph, all: ParsedParagraph[]): boolean {
  return paragraphHasText(para) || paragraphHasLineBreaks(para) || paragraphIsEmptySpacer(para, all);
}

function paragraphToPlainText(para: ParsedParagraph): string {
  let t = "";
  for (const r of para.runs) {
    t += r.isBreak ? "\n" : r.text;
  }
  return t;
}

function textBoxHasVisibleContent(paragraphs: ParsedParagraph[]): boolean {
  return paragraphs.some((p) => paragraphHasSpacing(p, paragraphs));
}

/** PowerPoint single line spacing is ~1.2× font size; spcPct is a multiple of that. */
const POWERPOINT_SINGLE_LINE_FACTOR = 1.2;

function canUseFlowLayout(paragraphs: ParsedParagraph[]): boolean {
  return paragraphs.every(
    (p) =>
      p.listKind === "p" &&
      !p.spaceBefore &&
      !p.spaceAfter &&
      (p.marginLeft == null || p.marginLeft === 0) &&
      p.level === 0,
  );
}

function inheritedPlaceholderAlign(
  phDefaults?: PlaceholderDefaults,
  ph?: Element | null,
): string | undefined {
  return phDefaults?.align
    || (ph ? IMPLICIT_PLACEHOLDER_ALIGN[ph.getAttribute("type") || ""] : undefined);
}

function resolveParagraphAlignCss(
  para: ParsedParagraph,
  phDefaults?: PlaceholderDefaults,
  ph?: Element | null,
): "left" | "center" | "right" | "justify" {
  if (para.align) return mapOoxmlAlign(para.align);
  const inherited = inheritedPlaceholderAlign(phDefaults, ph);
  if (inherited) return mapOoxmlAlign(inherited);
  return "left";
}

function resolveShapeHorizontalAlign(
  paragraphs: ParsedParagraph[],
  phDefaults?: PlaceholderDefaults,
  ph?: Element | null,
): "left" | "center" | "right" | "justify" {
  const contentParas = paragraphs.filter((p) => paragraphHasText(p));
  const effectiveAligns = contentParas.map((p) => resolveParagraphAlignCss(p, phDefaults, ph));
  const unique = new Set(effectiveAligns);
  if (unique.size === 1) return [...unique][0];
  if (unique.size > 1) return "left";

  const inherited = inheritedPlaceholderAlign(phDefaults, ph);
  if (inherited) return mapOoxmlAlign(inherited);
  return "left";
}

function buildFlowLayoutHtml(
  paragraphs: ParsedParagraph[],
  defaultFont: string,
  defaultSize: number,
  defaultColor: string | undefined,
  phDefaults?: PlaceholderDefaults,
  ph?: Element | null,
): string {
  const lines: string[] = [];
  for (const para of paragraphs) {
    if (!paragraphHasSpacing(para, paragraphs)) continue;
    const styles = ["margin:0", "padding:0", "display:block"];
    styles.push(`text-align:${resolveParagraphAlignCss(para, phDefaults, ph)}`);
    const inner = paragraphHasText(para) || paragraphHasLineBreaks(para)
      ? renderRunsToHtml(para.runs, defaultFont, defaultSize, defaultColor)
      : "<br>";
    lines.push(`<div style="${styles.join(";")}">${inner}</div>`);
  }
  if (lines.length === 0) return "";

  const wrapperStyles = ["margin:0", "padding:0", "display:block"];
  const lineHeight = paragraphs.find((p) => p.lineHeight)?.lineHeight;
  const lineHeightPt = paragraphs.find((p) => p.lineHeightPt)?.lineHeightPt;
  if (lineHeight) wrapperStyles.push(`line-height:${lineHeight}`);
  else if (lineHeightPt) wrapperStyles.push(`line-height:${lineHeightPt}pt`);

  return `<div style="${wrapperStyles.join(";")}">${lines.join("")}</div>`;
}

function paragraphLineGapPt(para: ParsedParagraph, defaultSize: number): number | undefined {
  if (para.lineHeightPt != null) {
    const gap = para.lineHeightPt - defaultSize;
    return gap > 0 ? gap : undefined;
  }
  if (para.lineHeight != null && para.lineHeight > 1) {
    return (para.lineHeight - 1) * defaultSize;
  }
  return undefined;
}

function buildRichHtml(
  paragraphs: ParsedParagraph[],
  bulletChar: string | undefined,
  defaultFont: string,
  defaultSize: number,
  defaultColor: string | undefined,
  phDefaults?: PlaceholderDefaults,
  ph?: Element | null,
): string {
  if (canUseFlowLayout(paragraphs)) {
    return buildFlowLayoutHtml(paragraphs, defaultFont, defaultSize, defaultColor, phDefaults, ph);
  }

  const parts: string[] = [];
  let openList: { kind: "ul" | "ol"; listStyle?: string } | null = null;
  const bullet = bulletChar ? escapeHtml(bulletChar) : "\u2022";

  for (let pi = 0; pi < paragraphs.length; pi++) {
    const para = paragraphs[pi];
    const runsHtml = renderRunsToHtml(para.runs, defaultFont, defaultSize, defaultColor);
    const hasContent = paragraphHasText(para);
    const hasSpacing = paragraphHasSpacing(para, paragraphs);

    const spacingStyles: string[] = [];
    if (para.spaceBefore && pi > 0) spacingStyles.push(`margin-top:${para.spaceBefore}pt`);
    if (para.spaceAfter && pi < paragraphs.length - 1) spacingStyles.push(`margin-bottom:${para.spaceAfter}pt`);
    const lineGap = paragraphLineGapPt(para, defaultSize);
    if (lineGap && pi < paragraphs.length - 1 && hasContent) spacingStyles.push(`margin-bottom:${lineGap}pt`);

    if (para.listKind === "p") {
      if (openList) {
        parts.push(openList.kind === "ul" ? "</ul>" : "</ol>");
        openList = null;
      }
      if (!hasSpacing) continue;
      const ml = para.marginLeft ?? (para.level > 0 ? para.level * 24 : 0);
      if (ml > 0) spacingStyles.push(`margin-left:${ml}px`);
      spacingStyles.push(`text-align:${resolveParagraphAlignCss(para, phDefaults, ph)}`);
      const styleAttr = spacingStyles.length ? ` style="${spacingStyles.join(";")}"` : "";
      const inner = hasContent || paragraphHasLineBreaks(para)
        ? runsHtml
        : "<br>";
      parts.push(`<div${styleAttr}>${inner}</div>`);
    } else {
      if (!openList || openList.kind !== para.listKind) {
        if (openList) parts.push(openList.kind === "ul" ? "</ul>" : "</ol>");
        const marL = para.marginLeft ?? para.level * 24;
        const { ulPaddingLeft } = listBulletLayoutCss(marL, para.indent);
        const commonListCss = `list-style-position:outside;padding-left:${ulPaddingLeft}px;margin:0;list-style-type:${para.listKind === "ol" ? (para.listStyle || "decimal") : "disc"}`;
        const style = ` style="${commonListCss}"`;
        parts.push(para.listKind === "ul" ? `<ul${style}>` : `<ol${style}>`);
        if (para.listKind === "ul") parts.push(`<style>.pptx-bullet::marker{content:"${bullet} ";}</style>`);
        openList = { kind: para.listKind, listStyle: para.listStyle };
      }
      const marL = para.marginLeft ?? para.level * 24;
      const { liPaddingLeft } = listBulletLayoutCss(marL, para.indent);
      if (liPaddingLeft > 0) spacingStyles.push(`padding-left:${liPaddingLeft}px`);
      if (para.bulletColor) spacingStyles.push(`color:${para.bulletColor}`);
      spacingStyles.push(`text-align:${resolveParagraphAlignCss(para, phDefaults, ph)}`);
      parts.push(`<li class="pptx-bullet" style="${spacingStyles.join(";")}">${runsHtml}</li>`);
    }
  }
  if (openList) parts.push(openList.kind === "ul" ? "</ul>" : "</ol>");
  return parts.join("");
}

function renderRunsToHtml(
  runs: ParsedRun[],
  defaultFont: string,
  defaultSize: number,
  defaultColor: string | undefined,
): string {
  return runs
    .flatMap((run) => {
      if (run.isBreak) return ["<br>"];
      return splitTextByUrls(run.text).map((part) => {
        const partRun: ParsedRun = {
          ...run,
          text: part.text,
          underline: run.underline || part.isUrl || undefined,
        };
        const styles = textRunInlineStyles(partRun, defaultFont, defaultSize, defaultColor);
        const escaped = escapeHtml(part.text);
        if (styles.length === 0) return escaped;
        return `<span style="${styles.join(";")}">${escaped}</span>`;
      });
    })
    .join("");
}

/**
 * Extract spcBef or spcAft from a:pPr. Returns value in pt.
 * Supports both spcPts (hundredths of a point) and spcPct (percentage of font size).
 * Uses getElementsByTagNameNS for reliable namespace handling with @xmldom/xmldom.
 */
function parseParagraphSpacing(pPr: Element | null, tag: "spcBef" | "spcAft"): number | undefined {
  if (!pPr) return undefined;
  const spcEl = pPr.getElementsByTagNameNS("*", tag)[0] ?? null;
  if (!spcEl) return undefined;
  const spcPts = spcEl.getElementsByTagNameNS("*", "spcPts")[0] ?? null;
  if (spcPts) {
    const v = Number(spcPts.getAttribute("val") || 0);
    return Number.isFinite(v) ? v / 100 : undefined;
  }
  const spcPct = spcEl.getElementsByTagNameNS("*", "spcPct")[0] ?? null;
  if (spcPct) {
    const v = Number(spcPct.getAttribute("val") || 0);
    return Number.isFinite(v) ? v / 1000 : undefined;
  }
  return undefined;
}

function mapAutoNumToCss(typ: string): string { const t = typ.toLowerCase(); if (t.includes("alphauc")) return "upper-alpha"; if (t.includes("alphalc")) return "lower-alpha"; if (t.includes("romanu")) return "upper-roman"; if (t.includes("romanl")) return "lower-roman"; return "decimal"; }
function escapeHtml(str: string): string { return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;"); }
/**
 * Extract a:lnSpc from a:bodyPr or a:pPr.
 * spcPct is a percentage multiplier (100000 = 100%); spcPts is hundredths of a point.
 */
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
