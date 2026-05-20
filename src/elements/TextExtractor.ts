import { TextElement } from "../models/SlideElement";
import { XmlHelper } from "../core/XmlHelper";

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
      const textRuns: string[] = [];
      const segments: NonNullable<TextElement["segments"]> = [];
      let fontName = "Arial";
      let fontSize = 18;
      let color: string | undefined;
      let horizontalAlign: "left" | "center" | "right" | "justify" | undefined;
      let bulletChar: string | undefined;
      const paraItems: { kind: "p" | "ul" | "ol"; text: string; lvl: number; listStyle?: string }[] = [];
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
      for (const p of Array.from(paragraphs)) {
        const pPr = p.getElementsByTagNameNS("*", "pPr")[0] ?? null;
        const algn = pPr?.getAttribute("algn") || undefined;
        if (algn && !horizontalAlign) horizontalAlign = algn === "ctr" ? "center" : algn === "r" ? "right" : algn.startsWith("just") ? "justify" : "left";
        const runs = p.getElementsByTagNameNS("*", "r");
        const paraText = getParagraphText(p);
        if (paraText) paraText.split(/\n+/).forEach((t) => { if (t) textRuns.push(t); });
        if (opts.preserveTextStructure && paraText) {
          segments.push({ text: paraText, paragraphBreakBefore: segments.length > 0 });
        }
        for (const r of Array.from(runs)) {
          const rPr = r.getElementsByTagNameNS("*", "rPr")[0];
          if (rPr) {
            const latin = rPr.getElementsByTagNameNS("*", "latin")[0];
            fontName = latin?.getAttribute("typeface") ?? fontName;
            const sz = rPr.getAttribute("sz");
            if (sz) {
              const n = parseInt(sz, 10);
              if (Number.isFinite(n)) fontSize = n / 100;
            }
            const solidFill = rPr.querySelector("*|solidFill");
            const candidate = XmlHelper.getColorFromElement(solidFill || null, themeColors);
            if (candidate) color = candidate;
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
        paraItems.push({ kind, text: paraText, lvl: isNaN(lvl) ? 0 : lvl, listStyle });
      }
      if (!color) {
        const lstDefRPr = txBody.querySelector("*|lstStyle *|defRPr");
        const lstFill = lstDefRPr?.querySelector("*|solidFill");
        const c0 = XmlHelper.getColorFromElement(lstFill || null, themeColors);
        if (c0) color = c0;

        const defRPr = txBody.querySelector("*|defRPr");
        const defFill = defRPr?.querySelector("*|solidFill");
        const fallback1 = XmlHelper.getColorFromElement(defFill || null, themeColors);
        if (fallback1) color = fallback1;

        if (!color) {
          const spPr = shape.querySelector("p\\:spPr, spPr");
          const shapeFill = spPr?.querySelector("*|solidFill");
          const fallback2 = XmlHelper.getColorFromElement(shapeFill || null, themeColors);
          if (fallback2) color = fallback2;
        }
      }
      const content = opts.preserveTextStructure ? textRuns.join("\n").trim() : textRuns.join(" ").trim();
      if (!content) continue;
      const xfrm = shape.getElementsByTagNameNS("*", "xfrm")[0];
      let off = xfrm?.getElementsByTagNameNS("*", "off")[0] ?? null;
      let ext = xfrm?.getElementsByTagNameNS("*", "ext")[0] ?? null;

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
      let richHtml: string | undefined;
      if (paraItems.some((it) => it.kind !== "p")) {
        const parts: string[] = [];
        let open: { kind: "ul" | "ol"; listStyle?: string } | null = null;
        const bullet = bulletChar ? escapeHtml(bulletChar) : "•";
        for (const it of paraItems) {
          if (it.kind === "p") {
            if (open) { parts.push(open.kind === "ul" ? "</ul>" : "</ol>"); open = null; }
            if (it.text.trim()) parts.push(`<div style="margin-left:${it.lvl * 24}px">${escapeHtml(it.text).replace(/\n/g, "<br>")}</div>`);
            continue;
          }
          if (!open || open.kind !== it.kind) {
            if (open) parts.push(open.kind === "ul" ? "</ul>" : "</ol>");
            const commonListCss = `list-style-position: inside; padding-left: 0; margin: 0;`;
            const style = it.kind === "ol" ? ` style="${commonListCss} list-style-type: ${it.listStyle || "decimal"};"` : ` style="${commonListCss}"`;
            parts.push(it.kind === "ul" ? `<ul${style}>` : `<ol${style}>`);
            if (it.kind === "ul") parts.push(`<style>.pptx-bullet::marker{content:"${bullet} ";}</style>`);
            open = { kind: it.kind, listStyle: it.listStyle };
          }
          parts.push(`<li class="pptx-bullet" style="margin-left:${it.lvl * 24}px">${escapeHtml(it.text).replace(/\n/g, "<br>")}</li>`);
        }
        if (open) parts.push(open.kind === "ul" ? "</ul>" : "</ol>");
        richHtml = parts.join("");
      }
      elements.push({ type: "text", content, position: { x, y }, size: { width: cx, height: cy }, font: { name: fontName, size: fontSize, color }, align: { horizontal: horizontalAlign ?? "left", vertical: verticalAlign }, padding, html: richHtml, segments: opts.preserveTextStructure ? segments : undefined, lineHeight });
    }
    return elements;
  }
}

function getParagraphText(p: Element): string {
  let out = "";
  for (const child of Array.from(p.childNodes) as any[]) {
    if (!(child instanceof Element)) continue;
    const ln = child.localName;
    if (ln === "r") out += child.getElementsByTagNameNS("*", "t")[0]?.textContent ?? "";
    else if (ln === "br") out += "\n";
    else if (ln === "fld") for (const r of Array.from(child.getElementsByTagNameNS("*", "r"))) out += r.getElementsByTagNameNS("*", "t")[0]?.textContent ?? "";
    else if (ln === "tab") out += "\t";
  }
  return out;
}

function mapAutoNumToCss(typ: string): string { const t = typ.toLowerCase(); if (t.includes("alphauc")) return "upper-alpha"; if (t.includes("alphalc")) return "lower-alpha"; if (t.includes("romanu")) return "upper-roman"; if (t.includes("romanl")) return "lower-roman"; return "decimal"; }
function escapeHtml(str: string): string { return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;"); }
function getLineSpacing(bodyPr: Element | null): number | undefined { const lnSpc = bodyPr?.querySelector("*|lnSpc"); const spcPct = lnSpc?.querySelector("*|spcPct"); const spcPts = lnSpc?.querySelector("*|spcPts"); if (spcPts) { const v = Number(spcPts.getAttribute("val") || 0); return Number.isFinite(v) && v > 0 ? v / 100 : undefined; } if (spcPct) { const v = Number(spcPct.getAttribute("val") || 0); return Number.isFinite(v) && v > 0 ? v / 100000 : undefined; } return undefined; }
