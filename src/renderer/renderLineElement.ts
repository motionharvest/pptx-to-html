import { LineElement } from "../models/SlideElement";

export function renderLineElement(el: LineElement, options: { scaleStrokes?: boolean } = {}): string {
  const nf = (n: number, fb = 0) => (Number.isFinite(n) ? n : fb);
  const x = nf(el.position?.x, 0) / 9525;
  const y = nf(el.position?.y, 0) / 9525;
  const width = Math.max(nf(el.size?.width, 0) / 9525, 1);
  const height = Math.max(nf(el.size?.height, 0) / 9525, 1);
  const sw = el.strokeWidth && el.strokeWidth > 0 ? el.strokeWidth : 1;
  const dash = el.dashStyle === "dashed" ? "stroke-dasharray: 8 6;" : el.dashStyle === "dotted" ? "stroke-dasharray: 2 6;" : "";

  const defs = buildMarkerDefs(el.headEnd, el.tailEnd, el.color || "#000");
  const markerStartAttr = defs.startId ? `marker-start="url(#${defs.startId})"` : "";
  const markerEndAttr = defs.endId ? `marker-end="url(#${defs.endId})"` : "";

  return `<svg viewBox="0 0 ${width} ${height}" style="position:absolute; left:${x}px; top:${y}px; width:${width}px; height:${height}px; overflow:visible;" overflow="visible">
    ${defs.defs}
    <line x1="0" y1="${height / 2}" x2="${width}" y2="${height / 2}" stroke="${el.color || "#000"}" stroke-width="${sw}" ${options.scaleStrokes ? "" : "vector-effect=\"non-scaling-stroke\""} style="${dash}" ${markerStartAttr} ${markerEndAttr} />
  </svg>`;
}

function buildMarkerDefs(headEnd: any, tailEnd: any, color: string): { defs: string; startId?: string; endId?: string } {
  const parts: string[] = [];
  let startId: string | undefined;
  let endId: string | undefined;
  if (headEnd && headEnd.type && headEnd.type !== "none") { startId = `mstart-${Math.random().toString(36).slice(2, 8)}`; parts.push(markerDef(startId, headEnd, color)); }
  if (tailEnd && tailEnd.type && tailEnd.type !== "none") { endId = `mend-${Math.random().toString(36).slice(2, 8)}`; parts.push(markerDef(endId, tailEnd, color)); }
  return { defs: parts.length ? `<defs>${parts.join("\n")}</defs>` : "", startId, endId };
}

function markerDef(id: string, spec: any, color: string): string {
  return `<marker id="${id}" markerUnits="strokeWidth" markerWidth="8" markerHeight="8" refX="8" refY="4" orient="auto-start-reverse"><polygon points="0,0 8,4 0,8" fill="${color}" /></marker>`;
}
