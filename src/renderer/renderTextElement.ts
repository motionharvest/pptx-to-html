import { TextElement } from "../models/SlideElement";
import { textRunInlineStyles } from "../core/scriptTextStyle";
import { TEXT_DIRECTION_CENTER_WRAPPER, textDirectionInnerStyles } from "../core/textDirection";

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
  const flowStyles = ["width:100%", "margin:0", "padding:0", "min-height:0", "align-self:stretch"];
  if (!el.html) {
    if (el.lineHeight) flowStyles.push(`line-height:${el.lineHeight}`);
    else if (el.lineHeightPt) flowStyles.push(`line-height:${el.lineHeightPt}pt`);
  }

  const rawInner = el.html
    ? el.html
    : el.segments?.length
      ? el.segments.map((seg) => {
          const prefix = seg.paragraphBreakBefore ? "<br>" : "";
          if (seg.breakBefore) return `${prefix}<br>`;
          const styles = textRunInlineStyles(
            seg,
            el.font?.name || "Arial",
            nf(Number(el.font?.size), 12),
            el.font?.color || "#000",
          );
          const text = escape(seg.text).replace(/\n/g, "<br>");
          return styles.length > 0 ? `${prefix}<span style="${styles.join(";")}">${text}</span>` : `${prefix}${text}`;
        }).join("")
      : escape(el.content).replace(/\n/g, "<br>");
  const inner = el.textDirection
    ? `<div style="${TEXT_DIRECTION_CENTER_WRAPPER}"><div style="${textDirectionInnerStyles(el.textDirection).join(";")}">${rawInner}</div></div>`
    : `<div style="${flowStyles.join(";")}">${rawInner}</div>`;

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
    line-height: 1;
    padding: ${pad.top}px ${pad.right}px ${pad.bottom}px ${pad.left}px;
    font-family: ${el.font?.name || "Arial"};
    font-size: ${nf(Number(el.font?.size), 12)}pt;
    color: ${el.font?.color || "#000"};
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
