import { CustomGeometry } from "../models/SlideElement";

const EMU_PER_PX = 9525;

function escapeAttr(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
}

function buildPathElements(
  geom: CustomGeometry,
  fill: string,
  stroke: string | undefined,
  strokeWidthPx?: number
): string {
  const sw = strokeWidthPx && strokeWidthPx > 0 ? strokeWidthPx : 0;
  const strokeAttr = stroke && stroke !== "transparent" ? stroke : "none";

  return geom.paths
    .map((p) => {
      if (p.fillMode === "none") {
        return `<path d="${p.d}" fill="none" stroke="${strokeAttr}" stroke-width="${sw}"/>`;
      }
      return `<path d="${p.d}" fill="${fill}" stroke="${strokeAttr}" stroke-width="${sw}"/>`;
    })
    .join("\n");
}

/**
 * Renders custom geometry as an absolutely positioned SVG.
 */
export function renderCustGeomSvg(options: {
  x: number;
  y: number;
  width: number;
  height: number;
  geom: CustomGeometry;
  fill: string;
  stroke?: string;
  strokeWidthPx?: number;
  rotationDeg?: number;
  imageSrc?: string;
}): string {
  const {
    x,
    y,
    width,
    height,
    geom,
    fill,
    stroke,
    strokeWidthPx,
    rotationDeg,
    imageSrc,
  } = options;

  const rotationStyle = rotationDeg
    ? `transform: rotate(${rotationDeg}deg); transform-origin: center;`
    : "";

  const pathMarkup = buildPathElements(geom, fill, stroke, strokeWidthPx);

  if (imageSrc) {
    const clipId = `cg-${Math.random().toString(36).slice(2, 10)}`;
    const clipPaths = geom.paths
      .filter((p) => p.fillMode !== "none")
      .map((p) => `<path d="${p.d}"/>`)
      .join("");

    return `<svg xmlns="http://www.w3.org/2000/svg" style="
    position: absolute;
    left: ${x}px;
    top: ${y}px;
    width: ${width}px;
    height: ${height}px;
    overflow: visible;
    ${rotationStyle}
  " viewBox="0 0 ${geom.viewBoxW} ${geom.viewBoxH}" preserveAspectRatio="none">
    <defs>
      <clipPath id="${clipId}" clipPathUnits="userSpaceOnUse">${clipPaths}</clipPath>
    </defs>
    <image href="${escapeAttr(imageSrc)}" x="0" y="0" width="100%" height="100%" preserveAspectRatio="xMidYMid slice" clip-path="url(#${clipId})"/>
  </svg>`;
  }

  return `<svg xmlns="http://www.w3.org/2000/svg" style="
    position: absolute;
    left: ${x}px;
    top: ${y}px;
    width: ${width}px;
    height: ${height}px;
    overflow: visible;
    ${rotationStyle}
  " viewBox="0 0 ${geom.viewBoxW} ${geom.viewBoxH}" preserveAspectRatio="none">
    ${pathMarkup}
  </svg>`;
}

export function emuToPx(n: number, fallback = 0): number {
  return (Number.isFinite(n) ? n : fallback) / EMU_PER_PX;
}
