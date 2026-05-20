import { TextElement } from "../models/SlideElement";
import { XmlHelper } from "../core/XmlHelper";

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
  static extract(spTree: Element | null, themeColors: Record<string, string>, opts: { context?: "slide" | "layout" | "master"; placeholderGeom?: Record<string, { x: number; y: number; cx: number; cy: number }>; preserveTextStructure?: boolean } = {}): TextElement[] {
    if (!spTree) return [];
    const elements: TextElement[] = [];
    const shapes = spTree.getElementsByTagNameNS("*", "sp");
    for (const shape of Array.from(shapes)) {
      const nvPr = shape.getElementsByTagNameNS("*", "nvPr")[0] ?? null;
      const ph = nvPr?.getElementsByTagNameNS("*", "ph")[0] ?? null;
      if (opts.context && opts.context !== "slide" && ph) continue;
      const txBody = shape.getElementsByTagNameNS("*", "txBody")[0];
      if (!txBody) continue;
      const paragraphs = txBody.getElementsByTagNameNS("*", "p");
      const bodyPr = txBody.getElementsByTagNameNS("*", "bodyPr")[0] ?? null;
      const anchor = bodyPr?.getAttribute("anchor") || undefined;
      const verticalAlign = anchor === "ctr" ? "middle" : anchor === "b" ? "bottom" : "top";
      const lIns = bodyPr?.getAttribute("lIns");
      const tIns = bodyPr?.getAttribute("tIns");
      const rIns = bodyPr?.getAttribute("rIns");
      const bIns = bodyPr?.getAttribute("bIns");
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

      const parsedParagraphs: ParsedParagraph[] = [];
      let horizontalAlign: "left" | "center" | "right" | "justify" | undefined;
      let defaultFontName = "Arial";
      let defaultFontSize = 18;
      let fontDefaultsCaptured = false;

      for (const p of Array.from(paragraphs)) {
        const pPr = p.getElementsByTagNameNS("*", "pPr")[0] ?? null;
        const algn = pPr?.getAttribute("algn") || undefined;
        if (algn && !horizontalAlign) horizontalAlign = algn === "ctr" ? "center" : algn === "r" ? "right" : algn.startsWith("just") ? "justify" : "left";

        const runs = extractRunsFromParagraph(p, themeColors);

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

      // Resolve default color from the fallback chain (NOT from individual runs)
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

      // Build plain text content (backward-compatible with old getParagraphText behavior)
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

      const xfrm = shape.getElementsByTagNameNS("*", "xfrm")[0];
      const off = xfrm?.getElementsByTagNameNS("*", "off")[0] ?? null;
      const ext = xfrm?.getElementsByTagNameNS("*", "ext")[0] ?? null;

      let x: number, y: number, cx: number, cy: number;
      if (off && ext) {
        x = XmlHelper.getAttrAsNumber(off, "x");
        y = XmlHelper.getAttrAsNumber(off, "y");
        cx = XmlHelper.getAttrAsNumber(ext, "cx");
        cy = XmlHelper.getAttrAsNumber(ext, "cy");
      } else if (opts.placeholderGeom) {
        const phIdx = ph?.getAttribute("idx") || undefined;
        const g = phIdx ? opts.placeholderGeom[phIdx] : undefined;
        x = g?.x ?? 0;
        y = g?.y ?? 0;
        cx = g?.cx ?? 1000000;
        cy = g?.cy ?? 500000;
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

function extractRunsFromParagraph(p: Element, themeColors: Record<string, string>): ParsedRun[] {
  const runs: ParsedRun[] = [];
  for (const child of Array.from(p.childNodes) as any[]) {
    if (!(child instanceof Element)) continue;
    const ln = child.localName;
    if (ln === "r") {
      runs.push(parseTextRun(child, themeColors));
    } else if (ln === "br") {
      runs.push({ text: "\n", isBreak: true });
    } else if (ln === "fld") {
      const t = child.getElementsByTagNameNS("*", "t")[0]?.textContent ?? "";
      if (t) {
        const rPr = child.getElementsByTagNameNS("*", "rPr")[0] ?? null;
        runs.push(buildRunFromProps(t, rPr, themeColors));
      }
    } else if (ln === "tab") {
      runs.push({ text: "\t" });
    }
  }
  return runs;
}

function parseTextRun(r: Element, themeColors: Record<string, string>): ParsedRun {
  const rPr = r.getElementsByTagNameNS("*", "rPr")[0] ?? null;
  const text = r.getElementsByTagNameNS("*", "t")[0]?.textContent ?? "";
  return buildRunFromProps(text, rPr, themeColors);
}

function buildRunFromProps(text: string, rPr: Element | null, themeColors: Record<string, string>): ParsedRun {
  if (!rPr) return { text };
  const bold = rPr.getAttribute("b") === "1" || undefined;
  const italic = rPr.getAttribute("i") === "1" || undefined;
  const uVal = rPr.getAttribute("u");
  const underline = uVal && uVal !== "none" ? true : undefined;
  const solidFill = rPr.querySelector("*|solidFill");
  const color = XmlHelper.getColorFromElement(solidFill || null, themeColors);
  let fontSize: number | undefined;
  const sz = rPr.getAttribute("sz");
  if (sz) {
    const n = parseInt(sz, 10);
    if (Number.isFinite(n)) fontSize = n / 100;
  }
  const latin = rPr.getElementsByTagNameNS("*", "latin")[0];
  const fontFamily = latin?.getAttribute("typeface") || undefined;
  return { text, bold, italic, underline, color, fontSize, fontFamily };
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
