import { CustomGeometry, ImageCrop } from "../models/SlideElement";
import { computeCroppedImageLayout } from "../core/imageCropLayout";
import { buildMarkerDefs, dashStyleToSvgAttr, DashStyle } from "../core/lineStyle";

const EMU_PER_PX = 9525;

function escapeAttr(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
}

/** Rectangular clip that matches the custGeom viewBox — clipPath is redundant. */
function isFullViewBoxRectClip(geom: CustomGeometry): boolean {
  const clips = geom.paths.filter((p) => p.fillMode !== "none");
  if (clips.length !== 1) return false;
  const norm = clips[0].d.replace(/\s+/g, " ").trim();
  const { viewBoxW: w, viewBoxH: h } = geom;
  const closed = `M 0 0 L ${w} 0 L ${w} ${h} L 0 ${h} L 0 0 Z`;
  const open = `M 0 0 L ${w} 0 L ${w} ${h} L 0 ${h} Z`;
  return norm === closed || norm === open;
}

function buildCustGeomImageMarkup(
  imageSrc: string,
  frameW: number,
  frameH: number,
  geom: CustomGeometry,
  crop?: ImageCrop,
  fillRect?: ImageCrop,
  naturalSize?: { width: number; height: number },
  preserveAspectRatio?: boolean,
): string {
  const sx = frameW > 0 ? geom.viewBoxW / frameW : 1;
  const sy = frameH > 0 ? geom.viewBoxH / frameH : 1;
  const href = escapeAttr(imageSrc);

  if (crop) {
    const layout = computeCroppedImageLayout({
      frameW,
      frameH,
      crop,
      fill: fillRect,
      naturalSize,
      preserveAspectRatio,
    });
    if (layout) {
      const { imgW, imgH, imgLeft, imgTop } = layout;
      return `<image href="${href}" x="${imgLeft * sx}" y="${imgTop * sy}" width="${imgW * sx}" height="${imgH * sy}" preserveAspectRatio="none"/>`;
    }
  }

  return `<image href="${href}" x="0" y="0" width="${geom.viewBoxW}" height="${geom.viewBoxH}" preserveAspectRatio="xMidYMid slice"/>`;
}

function buildPathElements(
  geom: CustomGeometry,
  fill: string,
  stroke: string | undefined,
  strokeWidthPx?: number,
  dashStyle?: DashStyle,
  scaleStrokes?: boolean,
): string {
  const sw = strokeWidthPx && strokeWidthPx > 0 ? strokeWidthPx : 0;
  const strokeColor = stroke && stroke !== "transparent" ? stroke : "none";
  const dashAttr = dashStyleToSvgAttr(dashStyle);
  // custGeom paths use EMU-scale viewBox coords; without this, strokes shrink to ~0px when scaled to fit.
  const vectorEffect =
    sw > 0 && strokeColor !== "none" && !scaleStrokes
      ? 'vector-effect="non-scaling-stroke"'
      : "";

  return geom.paths
    .map((p) => {
      if (p.fillMode === "none") {
        return `<path d="${p.d}" fill="none" stroke="${strokeColor}" stroke-width="${sw}" ${vectorEffect} ${dashAttr}/>`;
      }
      return `<path d="${p.d}" fill="${fill}" stroke="${strokeColor}" stroke-width="${sw}" ${vectorEffect} ${dashAttr}/>`;
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
  crop?: ImageCrop;
  fillRect?: ImageCrop;
  naturalSize?: { width: number; height: number };
  preserveAspectRatio?: boolean;
  dashStyle?: DashStyle;
  headEnd?: { type?: string; w?: string; len?: string };
  tailEnd?: { type?: string; w?: string; len?: string };
  scaleStrokes?: boolean;
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
    crop,
    fillRect,
    naturalSize,
    preserveAspectRatio,
    dashStyle,
    headEnd,
    tailEnd,
    scaleStrokes,
  } = options;

  const rotationStyle = rotationDeg
    ? `transform: rotate(${rotationDeg}deg); transform-origin: center;`
    : "";

  const strokeColor = stroke && stroke !== "transparent" ? stroke : undefined;
  const markerStroke = strokeColor || fill;
  const defs = buildMarkerDefs(headEnd, tailEnd, markerStroke || "#000", "cg");
  const markerStartAttr = defs.startId ? `marker-start="url(#${defs.startId})"` : "";
  const markerEndAttr = defs.endId ? `marker-end="url(#${defs.endId})"` : "";

  const sw = strokeWidthPx && strokeWidthPx > 0 ? strokeWidthPx : 0;
  const minStrokeDim = scaleStrokes ? 0 : sw;
  const effW = Math.max(width, minStrokeDim);
  const effH = Math.max(height, minStrokeDim);

  const pathMarkup = buildPathElements(geom, fill, stroke, strokeWidthPx, dashStyle, scaleStrokes);
  const markerAttrs = markerStartAttr || markerEndAttr
    ? pathMarkup.replace(/<path /, `<path ${markerStartAttr} ${markerEndAttr} `)
    : pathMarkup;

  if (imageSrc) {
    const imageMarkup = buildCustGeomImageMarkup(
      imageSrc,
      width,
      height,
      geom,
      crop,
      fillRect,
      naturalSize,
      preserveAspectRatio,
    );
    const skipClip = isFullViewBoxRectClip(geom);
    const clipPaths = geom.paths
      .filter((p) => p.fillMode !== "none")
      .map((p) => `<path d="${p.d}"/>`)
      .join("");
    const clipId = `cg-${Math.random().toString(36).slice(2, 10)}`;
    const clipDef = skipClip
      ? ""
      : `<defs><clipPath id="${clipId}" clipPathUnits="userSpaceOnUse">${clipPaths}</clipPath></defs>`;
    const clippedImage = skipClip
      ? imageMarkup
      : `<g clip-path="url(#${clipId})">${imageMarkup}</g>`;

    return `<svg xmlns="http://www.w3.org/2000/svg" style="
    position: absolute;
    left: ${x}px;
    top: ${y}px;
    width: ${effW}px;
    height: ${effH}px;
    display: block;
    overflow: hidden;
    ${rotationStyle}
  " viewBox="0 0 ${geom.viewBoxW} ${geom.viewBoxH}" preserveAspectRatio="none">
    ${clipDef}
    ${clippedImage}
  </svg>`;
  }

  return `<svg xmlns="http://www.w3.org/2000/svg" style="
    position: absolute;
    left: ${x}px;
    top: ${y}px;
    width: ${effW}px;
    height: ${effH}px;
    overflow: visible;
    ${rotationStyle}
  " viewBox="0 0 ${geom.viewBoxW} ${geom.viewBoxH}" preserveAspectRatio="none">
    ${defs.defs}
    ${markerAttrs}
  </svg>`;
}

export function emuToPx(n: number, fallback = 0): number {
  return (Number.isFinite(n) ? n : fallback) / EMU_PER_PX;
}
