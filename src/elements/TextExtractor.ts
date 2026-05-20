import { TextElement } from "../models/SlideElement";
import { XmlHelper } from "../core/XmlHelper";

export interface PlaceholderDefaults {
  x?: number;
  y?: number;
  cx?: number;
  cy?: number;
  anchor?: string;
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
}

interface TextRunDefaults {
  bold?: boolean;
  italic?: boolean;
  color?: string;
  fontSize?: number;
  fontFamily?: string;
}

interface ParsedRun {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
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
}

export class TextExtractor {
  static extract(spTree: Element | null, themeColors: Record<string, string>, opts: { context?: "slide" | "layout" | "master"; placeholderDefaults?: Map<string, PlaceholderDefaults>; preserveTextStructure?: boolean } = {}): TextElement[] {
    if (!spTree) return [];
    const elements: TextElement[] = [];
    const shapes = spTree.getElementsByTagNameNS("*", "sp");
    for (const shape of Array.from(shapes)) {
      const nvPr = shape.getElementsByTagNameNS("*", "nvPr")[0] ?? null;
      const ph = nvPr?.getElementsByTagNameNS("*", "ph")[0] ?? null;
      if (opts.context && opts.context !== "slide" && ph) continue;
      const txBody = shape.getElementsByTagNameNS("*", "txBody")[0];
      if (!txBody) continue;

      // Resolve placeholder defaults for this shape
      const phDefaults = resolvePlaceholder(ph, opts.placeholderDefaults);

      const paragraphs = txBody.getElementsByTagNameNS("*", "p");
      const bodyPr = txBody.getElementsByTagNameNS("*", "bodyPr")[0] ?? null;
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
      const lvlDefaults: Record<number, { kind: "p" | "ul" | "ol"; listStyle?: string }> = {};
      const lstStyle = txBody.querySelector("*|lstStyle");
      if (lstStyle) {
        for (const node of Array.from(lstStyle.children)) {
          const m = node.localName.match(/^lvl(\d+)pPr$/);
          if (!m) continue;
          const idx = parseInt(m[1], 10) - 1;
          let kind: "p" | "ul" | "ol" = "p";
          let listStyle: string | undefined;
          if (node.querySelector("*|buNone")) kind = "p";
          else if (node.querySelector("*|buAutoNum")) { kind = "ol"; listStyle = mapAutoNumToCss(node.querySelector("*|buAutoNum")?.getAttribute("type") || "arabicPeriod"); }
          else if (node.querySelector("*|buChar")) { kind = "ul"; bulletChar = node.querySelector("*|buChar")?.getAttribute("char") || bulletChar; listStyle = "disc"; }
          lvlDefaults[idx] = { kind, listStyle };
        }
      }

      // Text run defaults cascade: placeholder → shape-level fallbacks
      const textDefaults: TextRunDefaults = {
        bold: phDefaults?.bold,
        italic: phDefaults?.italic,
        color: phDefaults?.color,
        fontSize: phDefaults?.fontSize,
        fontFamily: phDefaults?.fontFamily,
      };

      const parsedParagraphs: ParsedParagraph[] = [];
      let horizontalAlign: "left" | "center" | "right" | "justify" | undefined;
      let defaultFontName = phDefaults?.fontFamily || "Arial";
      let defaultFontSize = phDefaults?.fontSize || 18;
      let fontDefaultsCaptured = !!phDefaults?.fontFamily || !!phDefaults?.fontSize;

      for (const p of Array.from(paragraphs)) {
        const pPr = p.getElementsByTagNameNS("*", "pPr")[0] ?? null;
        const algn = pPr?.getAttribute("algn") || undefined;
        if (algn && !horizontalAlign) horizontalAlign = algn === "ctr" ? "center" : algn === "r" ? "right" : algn.startsWith("just") ? "justify" : "left";

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

        const lvlAttr = pPr?.getAttribute("lvl");
        const lvl = lvlAttr ? parseInt(lvlAttr, 10) : 0;
        let kind: "p" | "ul" | "ol" = "p";
        let listStyle: string | undefined;
        if (pPr?.querySelector("*|buNone")) kind = "p";
        else if (pPr?.querySelector("*|buAutoNum")) { kind = "ol"; listStyle = mapAutoNumToCss(pPr.querySelector("*|buAutoNum")?.getAttribute("type") || "arabicPeriod"); }
        else if (pPr?.querySelector("*|buChar")) { kind = "ul"; bulletChar = pPr.querySelector("*|buChar")?.getAttribute("char") || bulletChar; listStyle = "disc"; }
        else if (lvlDefaults[lvl]) { kind = lvlDefaults[lvl].kind; listStyle = lvlDefaults[lvl].listStyle; }

        parsedParagraphs.push({ runs, align: algn, level: isNaN(lvl) ? 0 : lvl, listKind: kind, listStyle });
      }

      // Apply placeholder alignment, then implicit type defaults
      if (!horizontalAlign) {
        const a = phDefaults?.align
          || (ph ? IMPLICIT_PLACEHOLDER_ALIGN[ph.getAttribute("type") || ""] : undefined);
        if (a) horizontalAlign = a === "ctr" ? "center" : a === "r" ? "right" : a.startsWith("just") ? "justify" : "left";
      }

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

      // Build plain text content
      const textParts: string[] = [];
      for (const para of parsedParagraphs) {
        let t = "";
        for (const r of para.runs) {
          t += r.isBreak ? "\n" : r.text;
        }
        t = t.trim();
        if (t) t.split(/\n+/).forEach((part) => { if (part.trim()) textParts.push(part); });
      }
      const content = opts.preserveTextStructure ? textParts.join("\n").trim() : textParts.join(" ").trim();
      if (!content) continue;

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
            color: run.color,
            fontSize: run.fontSize,
            fontFamily: run.fontFamily,
            paragraphBreakBefore: pi > 0 && isFirstContentRun,
          });
          isFirstContentRun = false;
        }
      }

      const richHtml = buildRichHtml(parsedParagraphs, bulletChar, defaultFontName, defaultFontSize, defaultColor);

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
      const lineHeight = getLineSpacing(bodyPr);
      elements.push({
        type: "text",
        content,
        position: { x, y },
        size: { width: cx, height: cy },
        font: { name: defaultFontName, size: defaultFontSize, color: defaultColor || "#000000" },
        align: { horizontal: horizontalAlign ?? "left", vertical: verticalAlign },
        padding,
        html: richHtml,
        segments,
        lineHeight,
      });
    }
    return elements;
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
function resolvePlaceholder(ph: Element | null, defaults?: Map<string, PlaceholderDefaults>): PlaceholderDefaults | undefined {
  if (!ph || !defaults) return undefined;
  const idx = ph.getAttribute("idx");
  const type = ph.getAttribute("type");
  if (idx) {
    const byIdx = defaults.get(`idx:${idx}`);
    if (byIdx) return byIdx;
  }
  if (type) {
    const byType = defaults.get(`type:${type}`);
    if (byType) return byType;
  }
  return undefined;
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
  const underline = uVal && uVal !== "none" ? true : undefined;

  const solidFill = rPr.querySelector("*|solidFill");
  const color = XmlHelper.getColorFromElement(solidFill || null, themeColors) ?? textDefaults?.color;

  let fontSize: number | undefined;
  const sz = rPr.getAttribute("sz");
  if (sz) {
    const n = parseInt(sz, 10);
    if (Number.isFinite(n)) fontSize = n / 100;
  }
  fontSize = fontSize ?? textDefaults?.fontSize;

  const latin = rPr.getElementsByTagNameNS("*", "latin")[0];
  const fontFamily = latin?.getAttribute("typeface") || textDefaults?.fontFamily || undefined;

  return { text, bold: bold || undefined, italic: italic || undefined, underline, color, fontSize, fontFamily };
}

function buildRichHtml(
  paragraphs: ParsedParagraph[],
  bulletChar: string | undefined,
  defaultFont: string,
  defaultSize: number,
  defaultColor: string | undefined,
): string {
  const parts: string[] = [];
  let openList: { kind: "ul" | "ol"; listStyle?: string } | null = null;
  const bullet = bulletChar ? escapeHtml(bulletChar) : "\u2022";

  for (const para of paragraphs) {
    const runsHtml = renderRunsToHtml(para.runs, defaultFont, defaultSize, defaultColor);
    const hasContent = para.runs.some((r) => !r.isBreak && r.text.trim());

    if (para.listKind === "p") {
      if (openList) {
        parts.push(openList.kind === "ul" ? "</ul>" : "</ol>");
        openList = null;
      }
      const ml = para.level > 0 ? ` style="margin-left:${para.level * 24}px"` : "";
      parts.push(hasContent ? `<div${ml}>${runsHtml}</div>` : `<div${ml}>&nbsp;</div>`);
    } else {
      if (!openList || openList.kind !== para.listKind) {
        if (openList) parts.push(openList.kind === "ul" ? "</ul>" : "</ol>");
        const commonListCss = `list-style-position: inside; padding-left: 0; margin: 0;`;
        const style = para.listKind === "ol"
          ? ` style="${commonListCss} list-style-type: ${para.listStyle || "decimal"};"`
          : ` style="${commonListCss}"`;
        parts.push(para.listKind === "ul" ? `<ul${style}>` : `<ol${style}>`);
        if (para.listKind === "ul") parts.push(`<style>.pptx-bullet::marker{content:"${bullet} ";}</style>`);
        openList = { kind: para.listKind, listStyle: para.listStyle };
      }
      parts.push(`<li class="pptx-bullet" style="margin-left:${para.level * 24}px">${runsHtml}</li>`);
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
  return runs.map((run) => {
    if (run.isBreak) return "<br>";
    const styles: string[] = [];
    if (run.bold) styles.push("font-weight:bold");
    if (run.italic) styles.push("font-style:italic");
    if (run.underline) styles.push("text-decoration:underline");
    if (run.color && run.color !== defaultColor) styles.push(`color:${run.color}`);
    if (run.fontSize && run.fontSize !== defaultSize) styles.push(`font-size:${run.fontSize}pt`);
    if (run.fontFamily && run.fontFamily !== defaultFont) styles.push(`font-family:${run.fontFamily}`);
    const escaped = escapeHtml(run.text);
    if (styles.length === 0) return escaped;
    return `<span style="${styles.join(";")}">${escaped}</span>`;
  }).join("");
}

function mapAutoNumToCss(typ: string): string { const t = typ.toLowerCase(); if (t.includes("alphauc")) return "upper-alpha"; if (t.includes("alphalc")) return "lower-alpha"; if (t.includes("romanu")) return "upper-roman"; if (t.includes("romanl")) return "lower-roman"; return "decimal"; }
function escapeHtml(str: string): string { return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;"); }
function getLineSpacing(bodyPr: Element | null): number | undefined { const lnSpc = bodyPr?.querySelector("*|lnSpc"); const spcPct = lnSpc?.querySelector("*|spcPct"); const spcPts = lnSpc?.querySelector("*|spcPts"); if (spcPts) { const v = Number(spcPts.getAttribute("val") || 0); return Number.isFinite(v) && v > 0 ? v / 100 : undefined; } if (spcPct) { const v = Number(spcPct.getAttribute("val") || 0); return Number.isFinite(v) && v > 0 ? v / 100000 : undefined; } return undefined; }
