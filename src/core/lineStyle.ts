export type DashStyle = "solid" | "dashed" | "dotted" | "sysDash" | "sysDot" | "lgDash" | "dashDot";

export interface ArrowEndSpec {
  type?: string;
  w?: string;
  len?: string;
}

/** Map DrawingML a:prstDash @val to SVG stroke-dasharray. */
export function dashStyleFromPrst(val: string | null | undefined): DashStyle | undefined {
  if (!val) return undefined;
  const v = val.toLowerCase();
  if (v === "solid") return "solid";
  if (v.includes("sysdot")) return "sysDot";
  if (v.includes("sysdash")) return "sysDash";
  if (v.includes("lgdashdotdot")) return "dashDot";
  if (v.includes("lgdashdot")) return "dashDot";
  if (v.includes("lgdash")) return "lgDash";
  if (v.includes("dashdot")) return "dashDot";
  if (v.includes("dot")) return "dotted";
  if (v.includes("dash")) return "dashed";
  return "solid";
}

export function dashStyleToSvgAttr(style: DashStyle | undefined): string {
  switch (style) {
    case "dotted":
    case "sysDot":
      return 'stroke-dasharray="2 4"';
    case "sysDash":
      return 'stroke-dasharray="6 4"';
    case "dashed":
      return 'stroke-dasharray="8 4"';
    case "lgDash":
      return 'stroke-dasharray="12 4"';
    case "dashDot":
      return 'stroke-dasharray="8 4 2 4"';
    case "solid":
    default:
      return "";
  }
}

export function buildMarkerDefs(
  headEnd: ArrowEndSpec | undefined,
  tailEnd: ArrowEndSpec | undefined,
  color: string,
  idPrefix = "m"
): { defs: string; startId?: string; endId?: string } {
  const parts: string[] = [];
  let startId: string | undefined;
  let endId: string | undefined;

  if (headEnd?.type && headEnd.type !== "none") {
    startId = `${idPrefix}start-${Math.random().toString(36).slice(2, 8)}`;
    parts.push(markerDef(startId, headEnd, color, "auto-start-reverse"));
  }
  if (tailEnd?.type && tailEnd.type !== "none") {
    endId = `${idPrefix}end-${Math.random().toString(36).slice(2, 8)}`;
    parts.push(markerDef(endId, tailEnd, color, "auto"));
  }

  return { defs: parts.length ? `<defs>${parts.join("\n")}</defs>` : "", startId, endId };
}

/** Arrow length in px (DrawingML len × stroke width). */
export function arrowLengthPx(spec: ArrowEndSpec | undefined, strokeWidthPx: number): number {
  if (!spec?.type || spec.type === "none") return 0;
  return mapArrowSize(spec.len) * strokeWidthPx;
}

/** How far to inset the line when marker refX=1 (tip extends (len−1)×sw beyond endpoint). */
export function arrowLineInsetPx(spec: ArrowEndSpec | undefined, strokeWidthPx: number): number {
  if (!spec?.type || spec.type === "none") return 0;
  const lenUnits = mapArrowSize(spec.len);
  return Math.max(0, lenUnits - 1) * strokeWidthPx;
}

/**
 * Shorten a line segment so stroke ends at the arrow base while markers (refX=1)
 * place the tip at the original endpoints.
 */
export function insetSegmentForArrowEnds(
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  headEnd: ArrowEndSpec | undefined,
  tailEnd: ArrowEndSpec | undefined,
  strokeWidthPx: number,
): { x1: number; y1: number; x2: number; y2: number } {
  const headInset = arrowLineInsetPx(headEnd, strokeWidthPx);
  const tailInset = arrowLineInsetPx(tailEnd, strokeWidthPx);
  if (headInset <= 0 && tailInset <= 0) return { x1, y1, x2, y2 };

  const dx = x2 - x1;
  const dy = y2 - y1;
  const len = Math.hypot(dx, dy);
  if (len <= 0) return { x1, y1, x2, y2 };

  const ux = dx / len;
  const uy = dy / len;
  let h = headInset;
  let t = tailInset;
  if (h + t >= len) {
    const scale = Math.max(0, (len - 1) / (h + t));
    h *= scale;
    t *= scale;
  }

  return {
    x1: x1 + ux * h,
    y1: y1 + uy * h,
    x2: x2 - ux * t,
    y2: y2 - uy * t,
  };
}

/** Arrow length/width as multiples of stroke width (Office-compatible defaults). */
function mapArrowSize(size?: string): number {
  switch ((size || "med").toLowerCase()) {
    case "sm":
    case "small":
      return 2;
    case "lg":
    case "large":
      return 5;
    case "med":
    case "medium":
    default:
      return 3;
  }
}

function markerDef(id: string, spec: ArrowEndSpec, color: string, orient: string): string {
  const arrowLen = mapArrowSize(spec.len);
  const arrowW = mapArrowSize(spec.w);
  const refX = 1;
  const refY = arrowW / 2;

  switch ((spec.type || "triangle").toLowerCase()) {
    case "diamond":
      return `<marker id="${id}" markerUnits="strokeWidth" markerWidth="${arrowLen}" markerHeight="${arrowW}"
                      refX="${refX}" refY="${refY}" orient="${orient}">
                <polygon points="${arrowLen / 2},0 ${arrowLen},${arrowW / 2} ${arrowLen / 2},${arrowW} 0,${arrowW / 2}" fill="${color}" />
              </marker>`;
    case "oval":
      return `<marker id="${id}" markerUnits="strokeWidth" markerWidth="${arrowLen}" markerHeight="${arrowW}"
                      refX="${refX}" refY="${refY}" orient="${orient}">
                <ellipse cx="${arrowLen / 2}" cy="${arrowW / 2}" rx="${arrowLen / 2}" ry="${arrowW / 2}" fill="${color}" />
              </marker>`;
    case "stealth":
      return `<marker id="${id}" markerUnits="strokeWidth" markerWidth="${arrowLen}" markerHeight="${arrowW}"
                      refX="${refX}" refY="${refY}" orient="${orient}">
                <polygon points="${arrowLen},${arrowW / 2} 0,0 0,${arrowW}" fill="${color}" />
              </marker>`;
    case "arrow":
    case "triangle":
    default:
      return `<marker id="${id}" markerUnits="strokeWidth" markerWidth="${arrowLen}" markerHeight="${arrowW}"
                      refX="${refX}" refY="${refY}" orient="${orient}">
                <polygon points="0,0 ${arrowLen},${arrowW / 2} 0,${arrowW}" fill="${color}" />
              </marker>`;
  }
}
