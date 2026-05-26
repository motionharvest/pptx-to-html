/**
 * Converts DrawingML custom geometry (a:custGeom / a:pathLst) to SVG path data.
 */

export interface CustGeomPath {
  d: string;
  /** Path fill mode from @fill — "none" skips fill */
  fillMode?: string;
}

export interface CustGeomData {
  viewBoxW: number;
  viewBoxH: number;
  paths: CustGeomPath[];
}

function ptCoords(pt: Element | null): { x: number; y: number } {
  if (!pt) return { x: 0, y: 0 };
  return {
    x: Number(pt.getAttribute("x") || 0),
    y: Number(pt.getAttribute("y") || 0),
  };
}

function appendArc(
  cmds: string[],
  prev: { x: number; y: number },
  arcTo: Element
): { x: number; y: number } {
  const wR = Number(arcTo.getAttribute("wR") || 0);
  const hR = Number(arcTo.getAttribute("hR") || 0);
  const stAng = Number(arcTo.getAttribute("stAng") || 0) / 60000;
  const swAng = Number(arcTo.getAttribute("swAng") || 0) / 60000;

  if (!wR || !hR || !swAng) return prev;

  const endX = prev.x + wR * Math.cos((stAng + swAng) * Math.PI / 180);
  const endY = prev.y + hR * Math.sin((stAng + swAng) * Math.PI / 180);
  const largeArc = Math.abs(swAng) > 180 ? 1 : 0;
  const sweep = swAng > 0 ? 1 : 0;

  cmds.push(
    `A ${wR} ${hR} ${stAng} ${largeArc} ${sweep} ${endX} ${endY}`
  );
  return { x: endX, y: endY };
}

function pathElementToSvgD(pathEl: Element): string {
  const cmds: string[] = [];
  let cur = { x: 0, y: 0 };

  for (const child of Array.from(pathEl.children)) {
    const tag = child.localName;
    if (tag === "moveTo") {
      const pt = child.getElementsByTagNameNS("*", "pt")[0];
      cur = ptCoords(pt);
      cmds.push(`M ${cur.x} ${cur.y}`);
    } else if (tag === "lnTo") {
      const pt = child.getElementsByTagNameNS("*", "pt")[0];
      cur = ptCoords(pt);
      cmds.push(`L ${cur.x} ${cur.y}`);
    } else if (tag === "cubicBezTo") {
      const pts = child.getElementsByTagNameNS("*", "pt");
      if (pts.length >= 3) {
        const p1 = ptCoords(pts[0]);
        const p2 = ptCoords(pts[1]);
        cur = ptCoords(pts[2]);
        cmds.push(`C ${p1.x} ${p1.y} ${p2.x} ${p2.y} ${cur.x} ${cur.y}`);
      }
    } else if (tag === "quadBezTo") {
      const pts = child.getElementsByTagNameNS("*", "pt");
      if (pts.length >= 2) {
        const p1 = ptCoords(pts[0]);
        cur = ptCoords(pts[1]);
        cmds.push(`Q ${p1.x} ${p1.y} ${cur.x} ${cur.y}`);
      }
    } else if (tag === "arcTo") {
      cur = appendArc(cmds, cur, child);
    } else if (tag === "close") {
      cmds.push("Z");
    }
  }

  return cmds.join(" ");
}

/**
 * Parses a:custGeom under spPr into SVG-ready path data.
 */
export function parseCustGeom(spPr: Element | null): CustGeomData | null {
  if (!spPr) return null;

  const custGeom = spPr.getElementsByTagNameNS("*", "custGeom")[0];
  if (!custGeom) return null;

  const pathLst = custGeom.getElementsByTagNameNS("*", "pathLst")[0];
  if (!pathLst) return null;

  const pathEls = Array.from(pathLst.getElementsByTagNameNS("*", "path"));
  if (!pathEls.length) return null;

  let viewBoxW = 0;
  let viewBoxH = 0;
  const paths: CustGeomPath[] = [];

  for (const pathEl of pathEls) {
    const w = Number(pathEl.getAttribute("w") || 0);
    const h = Number(pathEl.getAttribute("h") || 0);
    if (w > viewBoxW) viewBoxW = w;
    if (h > viewBoxH) viewBoxH = h;

    const d = pathElementToSvgD(pathEl);
    if (!d) continue;

    paths.push({
      d,
      fillMode: pathEl.getAttribute("fill") || "norm",
    });
  }

  if (!paths.length) return null;
  if (!viewBoxW) viewBoxW = 100000;
  if (!viewBoxH) viewBoxH = 100000;

  return { viewBoxW, viewBoxH, paths };
}
