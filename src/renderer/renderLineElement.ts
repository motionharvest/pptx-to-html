import { LineElement } from "../models/SlideElement";
import { buildMarkerDefs, dashStyleToSvgAttr, insetSegmentForArrowEnds } from "../core/lineStyle";

export function renderLineElement(el: LineElement, options: { scaleStrokes?: boolean } = {}): string {
  const nf = (n: number, fb = 0) => (Number.isFinite(n) ? n : fb);
  const x = nf(el.position?.x, 0) / 9525;
  const y = nf(el.position?.y, 0) / 9525;
  const rawW = nf(el.size?.width, 0) / 9525;
  const rawH = nf(el.size?.height, 0) / 9525;
  const width = Math.max(rawW, 1);
  const height = Math.max(rawH, 1);
  const sw = el.strokeWidth && el.strokeWidth > 0 ? el.strokeWidth : 1;
  const dashAttr = dashStyleToSvgAttr(el.dashStyle);
  const strokeColor = el.color || "#000";

  const defs = buildMarkerDefs(el.headEnd, el.tailEnd, strokeColor);
  const markerStartAttr = defs.startId ? `marker-start="url(#${defs.startId})"` : "";
  const markerEndAttr = defs.endId ? `marker-end="url(#${defs.endId})"` : "";
  const hasArrow = !!(defs.startId || defs.endId);
  const lineCapAttr = hasArrow ? 'stroke-linecap="butt"' : "";

  let x1 = 0, y1 = 0, x2 = rawW, y2 = rawH;
  if (el.flipH) { x1 = rawW; x2 = 0; }
  if (el.flipV) { y1 = rawH; y2 = 0; }

  if (rawW < 1) { x1 = width / 2; x2 = width / 2; }
  if (rawH < 1) { y1 = height / 2; y2 = height / 2; }

  if (hasArrow) {
    ({ x1, y1, x2, y2 } = insetSegmentForArrowEnds(x1, y1, x2, y2, el.headEnd, el.tailEnd, sw));
  }

  return `<svg viewBox="0 0 ${width} ${height}" style="position:absolute; left:${x}px; top:${y}px; width:${width}px; height:${height}px; overflow:visible;" overflow="visible">
    ${defs.defs}
    <line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" stroke="${strokeColor}" stroke-width="${sw}" ${lineCapAttr} ${options.scaleStrokes ? "" : "vector-effect=\"non-scaling-stroke\""} ${dashAttr} ${markerStartAttr} ${markerEndAttr} />
  </svg>`;
}
