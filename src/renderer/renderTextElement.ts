import { TextElement } from "../models/SlideElement";

/**
 * Renders a text element as an absolutely positioned HTML <div>.
 * @param el Text element to render.
 * @returns HTML string representing the text element.
 */
export function renderTextElement(el: TextElement): string {
  const nf = (n: number, fb = 0) => (Number.isFinite(n) ? n : fb);
  const x = nf(el.position?.x, 0) / 9525;
  const y = nf(el.position?.y, 0) / 9525;
  const w = nf(el.size?.width, 0) / 9525;
  const h = nf(el.size?.height, 0) / 9525;
  const pad = el.padding || { left: 0, top: 0, right: 0, bottom: 0 };
  const textAlign = el.align?.horizontal || "left";
  const justify = el.align?.vertical === "middle" ? "center" : el.align?.vertical === "bottom" ? "flex-end" : "flex-start";
  const inner = el.html
    ? el.html
    : el.segments?.length
      ? el.segments.map((seg) => `${seg.paragraphBreakBefore ? "<br>" : ""}${escape(seg.text).replace(/\n/g, "<br>")}`).join("")
      : escape(el.content).replace(/\n/g, "<br>");
  const lineHeight = el.lineHeight ? `line-height: ${el.lineHeight};` : "line-height: 1.0;";
  return `<div style="
    position: absolute;
    left: ${x}px;
    top: ${y}px;
    width: ${w}px;
    height: ${h}px;
    display: flex;
    flex-direction: column;
    justify-content: ${justify};
    text-align: ${textAlign};
    line-height: 1.0;
    padding: ${pad.top}px ${pad.right}px ${pad.bottom}px ${pad.left}px;
    font-family: ${el.font?.name || "Arial"};
    font-size: ${nf(Number(el.font?.size), 12)}pt;
    color: ${el.font?.color || "#000"};
    ${lineHeight}
    overflow: hidden;
    white-space: pre-wrap;
  ">${inner}</div>`;
}

function escape(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
