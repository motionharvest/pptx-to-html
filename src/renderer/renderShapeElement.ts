import { ShapeElement } from "../models/SlideElement";
import { getSvgPathForShape } from "./shapePathMap";
import { emuToPx, renderCustGeomSvg } from "./renderCustGeom";
import { buildMarkerDefs, dashStyleToSvgAttr } from "../core/lineStyle";
import { shadowCasterFill, shadowToSvgFilterDef } from "../core/shadowEffect";
import { roundRectCornerRadiusPx } from "../core/shapeStyle";

/**
 * Renders a shape element as an absolutely positioned HTML or SVG element.
 * Supports all recognized PPTX shape types using SVG when necessary.
 * @param el Shape element to render.
 * @returns HTML string representing the shape.
 */
export function renderShapeElement(el: ShapeElement, options: { scaleStrokes?: boolean } = {}): string {
    const x = emuToPx(el.position?.x ?? 0);
    const y = emuToPx(el.position?.y ?? 0);
    const width = emuToPx(el.size?.width ?? 0);
    const height = emuToPx(el.size?.height ?? 0);

    if (el.customGeometry) {
      return renderCustGeomSvg({
        x,
        y,
        width,
        height,
        geom: el.customGeometry,
        fill: el.fillColor,
        stroke: el.borderColor,
        strokeWidthPx: el.strokeWidth,
        rotationDeg: el.rotationDeg && !isNaN(el.rotationDeg) ? el.rotationDeg : undefined,
        dashStyle: el.dashStyle,
        headEnd: el.headEnd,
        tailEnd: el.tailEnd,
        scaleStrokes: options.scaleStrokes,
      });
    }

    const rotation = el.rotationDeg && !isNaN(el.rotationDeg) ? el.rotationDeg : 0;

    // Basic preset shapes — SVG stroke is centered on the path (matches PowerPoint outline).
    if (el.shapeType === "rect" || el.shapeType === "ellipse" || el.shapeType === "roundRect") {
        return renderPresetShapeSvg(
          el.shapeType,
          x,
          y,
          width,
          height,
          el.fillColor,
          el.borderColor,
          el.strokeWidth,
          rotation,
          el.dashStyle,
          el.shadow,
          el.roundRectAdj,
        );
    }

    // SVG-based shapes using prefixed definition
    const raw = getSvgPathForShape(el.shapeType);
    return shapeSvg(
      x,
      y,
      width,
      height,
      el.fillColor,
      el.borderColor,
      raw,
      el.strokeWidth && Number.isFinite(el.strokeWidth) ? el.strokeWidth : undefined,
      rotation,
      el.headEnd,
      el.tailEnd,
      options.scaleStrokes === true,
      el.dashStyle
    );
}

/** Render rect/ellipse/roundRect with stroke centered on the geometry edge (DrawingML default). */
function renderPresetShapeSvg(
  shapeType: "rect" | "ellipse" | "roundRect",
  x: number,
  y: number,
  width: number,
  height: number,
  fill: string,
  stroke: string | undefined,
  strokeWidthPx: number | undefined,
  rotationDeg: number,
  dashStyle?: ShapeElement["dashStyle"],
  shadow?: ShapeElement["shadow"],
  roundRectAdj?: number,
): string {
  const sw = strokeWidthPx && strokeWidthPx > 0 ? strokeWidthPx : 0;
  const strokeColor = stroke && stroke !== "transparent" ? stroke : "none";
  const fillColor = fill && fill !== "transparent" ? fill : "none";
  const renderFill = shadow ? shadowCasterFill(fillColor) : fillColor;
  const dashAttr = dashStyleToSvgAttr(dashStyle);
  const rotationStyle = rotationDeg
    ? `transform: rotate(${rotationDeg}deg); transform-origin: center;`
    : "";
  const filterId = shadow ? `sh-${Math.random().toString(36).slice(2, 10)}` : "";
  const filterDef = shadow ? shadowToSvgFilterDef(shadow, filterId) : "";
  const filterAttr = shadow ? `filter="url(#${filterId})"` : "";

  let markup = "";
  if (shapeType === "ellipse") {
    const cx = width / 2;
    const cy = height / 2;
    markup = `<ellipse cx="${cx}" cy="${cy}" rx="${cx}" ry="${cy}" fill="${renderFill}" stroke="${strokeColor}" stroke-width="${sw}" ${dashAttr} ${filterAttr} />`;
  } else if (shapeType === "roundRect") {
    const rx = roundRectCornerRadiusPx(width, height, roundRectAdj);
    markup = `<rect x="0" y="0" width="${width}" height="${height}" rx="${rx}" ry="${rx}" fill="${renderFill}" stroke="${strokeColor}" stroke-width="${sw}" ${dashAttr} ${filterAttr} />`;
  } else {
    markup = `<rect x="0" y="0" width="${width}" height="${height}" fill="${renderFill}" stroke="${strokeColor}" stroke-width="${sw}" ${dashAttr} ${filterAttr} />`;
  }

  return `<svg xmlns="http://www.w3.org/2000/svg" style="
    position: absolute;
    left: ${x}px;
    top: ${y}px;
    width: ${width}px;
    height: ${height}px;
    overflow: visible;
    ${rotationStyle}
  " viewBox="0 0 ${width} ${height}" preserveAspectRatio="none">
    ${filterDef ? `<defs>${filterDef}</defs>` : ""}
    ${markup}
  </svg>`;
}

function shapeSvg(
  x: number,
  y: number,
  width: number,
  height: number,
  fill: string,
  stroke: string | undefined,
  raw: string,
  strokeWidthPx?: number,
  rotationDeg?: number,
  headEnd?: { type?: string; w?: string; len?: string },
  tailEnd?: { type?: string; w?: string; len?: string },
  scaleStrokes?: boolean,
  dashStyle?: ShapeElement["dashStyle"]
): string {
  const strokeColor = resolveStrokeColor(stroke, fill) || "#000";
  const [typeRaw, ...rest] = raw.trim().split(/\s+/);
  const type = typeRaw.toUpperCase().replace("_ARROW", "");
  const isArrow = typeRaw.endsWith("_ARROW");
  const data = rest.join(" ");

  const svgHeight = height;
  const svgWidth = width;
  const sw = strokeWidthPx && strokeWidthPx > 0 ? strokeWidthPx : 2;
  const dashAttr = dashStyleToSvgAttr(dashStyle);
  const vectorEffect = scaleStrokes ? "" : "vector-effect=\"non-scaling-stroke\"";

  const rotationStyle = rotationDeg ? `transform: rotate(${rotationDeg}deg); transform-origin: center;` : "";
  const commonStyle = `
    position: absolute;
    left: ${x}px;
    top: ${y}px;
    width: ${svgWidth}px;
    height: ${svgHeight}px;
    ${rotationStyle}
  `;

  switch (type) {
    case "PATH": {
      const defs = buildMarkerDefs(headEnd, tailEnd, strokeColor);
      const markerStartAttr = defs.startId ? `marker-start="url(#${defs.startId})"` : "";
      const markerEndAttr = defs.endId ? `marker-end="url(#${defs.endId})"` : "";
      return `<svg viewBox="0 0 100 100" style="${commonStyle}" overflow="visible">
        ${defs.defs}
        <path d="${data}" fill="none" stroke="${strokeColor}" stroke-width="${sw}" ${vectorEffect} ${dashAttr} ${markerStartAttr} ${markerEndAttr} />
      </svg>`;
    }

    case "POLYLINE":
    case "LINE": {
      const coords = data
        .split(/[\s,]+/)
        .map((v) => parseFloat(v))
        .filter((v) => !isNaN(v));

      if (coords.length < 4 || coords.length % 2 !== 0) {
        if (typeof console !== "undefined" && console.warn) {
          console.warn(`[pptx-to-html] Invalid POLYLINE/LINE shape data: "${data}"`);
        }
        return "";
      }

      const effectiveWidth = width > 0 ? width : Math.max(sw * 2, 2);
      const effectiveHeight = height > 0 ? height : Math.max(sw * 2, 2);

      const defs = buildMarkerDefs(
        headEnd,
        tailEnd ?? (isArrow ? { type: "triangle", w: "med", len: "med" } : undefined),
        strokeColor
      );
      const markerStartAttr = defs.startId ? `marker-start="url(#${defs.startId})"` : "";
      const markerEndAttr = defs.endId ? `marker-end="url(#${defs.endId})"` : "";

      const scaledPairs: string[] = [];
      for (let i = 0; i < coords.length; i += 2) {
        const px = (coords[i] / 100) * effectiveWidth;
        const py = (coords[i + 1] / 100) * effectiveHeight;
        scaledPairs.push(`${px},${py}`);
      }
      const scaledPoints = scaledPairs.join(" ");

      return `<svg viewBox="0 0 ${effectiveWidth} ${effectiveHeight}"
            style="
              position: absolute;
              left: ${x}px;
              top: ${y}px;
              width: ${effectiveWidth}px;
              height: ${effectiveHeight}px;
              ${rotationStyle}
            "
            overflow="visible">
          ${defs.defs}
          <polyline points="${scaledPoints}"
                    fill="none"
                    stroke="${strokeColor}"
                    stroke-width="${sw}"
                    ${vectorEffect}
                    ${dashAttr}
                    ${markerStartAttr} ${markerEndAttr} />
        </svg>`;
    }

    case "POLYGON":
    default:
      return `<svg viewBox="0 0 100 100" style="${commonStyle}">
        <polygon points="${data}" fill="${fill}" stroke="${strokeColor}" stroke-width="${sw}" ${vectorEffect} />
      </svg>`;
  }
}

function resolveStrokeColor(stroke?: string, fill?: string): string | undefined {
  if (stroke && stroke !== "transparent") return stroke;
  if (fill && fill !== "transparent") return fill;
  return undefined;
}
