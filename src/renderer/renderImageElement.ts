import { ImageElement, ShapeFrameStyle } from "../models/SlideElement";
import { computeCroppedImageLayout } from "../core/imageCropLayout";
import { shadowCasterFill, shadowToSvgFilterDef } from "../core/shadowEffect";
import { emuToPx, renderCustGeomSvg } from "./renderCustGeom";
import { roundRectCornerRadiusPx } from "../core/shapeStyle";

function randomId(prefix: string): string {
  return `${prefix}-${Math.random().toString(36).slice(2, 10)}`;
}

function buildTransformStyle(el: ImageElement): string {
  const parts: string[] = [];
  const rotation = el.rotationDeg && !isNaN(el.rotationDeg) ? el.rotationDeg : 0;
  if (rotation) parts.push(`rotate(${rotation}deg)`);
  const scaleX = el.flipH ? -1 : 1;
  const scaleY = el.flipV ? -1 : 1;
  if (scaleX !== 1 || scaleY !== 1) parts.push(`scale(${scaleX}, ${scaleY})`);
  if (!parts.length) return "";
  return `transform: ${parts.join(" ")}; transform-origin: center;`;
}

function opacityStyle(el: ImageElement): string {
  return el.imageOpacity !== undefined && el.imageOpacity < 1
    ? `opacity: ${el.imageOpacity};`
    : "";
}

/**
 * Renders an image element as an absolutely positioned <img> or clipped SVG.
 * @param el Image element to render.
 * @returns HTML string representing the image element.
 */
export function renderImageElement(el: ImageElement): string {
  const x = emuToPx(el.position?.x ?? 0);
  const y = emuToPx(el.position?.y ?? 0);
  const width = emuToPx(el.size?.width ?? 0);
  const height = emuToPx(el.size?.height ?? 0);
  const rotationStyle = buildTransformStyle(el);
  const opacity = opacityStyle(el);

  if (el.customGeometry) {
    return renderCustGeomSvg({
      x,
      y,
      width,
      height,
      geom: el.customGeometry,
      fill: "transparent",
      imageSrc: el.src,
      crop: el.crop,
      fillRect: el.fill,
      naturalSize: el.naturalSize,
      preserveAspectRatio: el.preserveAspectRatio,
      rotationDeg: el.rotationDeg && !isNaN(el.rotationDeg) ? el.rotationDeg : undefined,
    });
  }

  if (
    el.frame
    && (el.frame.shapeType === "ellipse" || el.frame.shapeType === "roundRect")
  ) {
    return renderFramedImage(el, rotationStyle, opacity);
  }

  if (el.crop) {
    return renderCroppedImage(el, x, y, width, height, rotationStyle, opacity);
  }

  return `<img src="${el.src}" style="
    position: absolute;
    left: ${x}px;
    top: ${y}px;
    width: ${width}px;
    height: ${height}px;
    object-fit: cover;
    ${rotationStyle}
    ${opacity}
  " />`;
}

/** Render image with an ellipse/roundRect frame (clip and/or backing plate + shadow). */
function renderFramedImage(el: ImageElement, rotationStyle: string, opacity: string): string {
  const frame = el.frame!;
  const frameX = emuToPx(frame.position.x);
  const frameY = emuToPx(frame.position.y);
  const frameW = emuToPx(frame.size.width);
  const frameH = emuToPx(frame.size.height);
  const imgX = emuToPx(el.position.x);
  const imgY = emuToPx(el.position.y);
  const imgW = emuToPx(el.size.width);
  const imgH = emuToPx(el.size.height);

  if (frame.clipImage) {
    return renderClippedFrameImage(el, frame, frameX, frameY, frameW, frameH, rotationStyle, opacity);
  }

  return renderPlateFrameImage(
    el,
    frame,
    imgX,
    imgY,
    imgW,
    imgH,
    frameX - imgX,
    frameY - imgY,
    frameW,
    frameH,
    rotationStyle,
    opacity,
  );
}

/** Photo clipped to circle/roundRect (team headshots). */
function renderClippedFrameImage(
  el: ImageElement,
  frame: ShapeFrameStyle,
  frameX: number,
  frameY: number,
  frameW: number,
  frameH: number,
  rotationStyle: string,
  opacity: string,
): string {
  const clipId = randomId("fc");
  const filterId = frame.shadow ? randomId("fs") : "";
  const filterDef = frame.shadow ? shadowToSvgFilterDef(frame.shadow, filterId) : "";
  const filterAttr = frame.shadow ? `filter="url(#${filterId})"` : "";

  const fillColor = frame.fillColor && frame.fillColor !== "transparent" ? frame.fillColor : "none";
  const strokeColor = frame.borderColor && frame.borderColor !== "transparent" ? frame.borderColor : "none";
  const sw = frame.strokeWidth && frame.strokeWidth > 0 ? frame.strokeWidth : 0;
  const shadowFill = frame.shadow ? shadowCasterFill(fillColor) : fillColor;

  const shapeMarkup = frame.shapeType === "roundRect"
    ? buildRoundRectMarkup(frameW, frameH, shadowFill, strokeColor, sw, 0, 0, frame.roundRectAdj)
    : buildEllipseMarkup(frameW, frameH, shadowFill, strokeColor, sw);

  const clipMarkup = frame.shapeType === "roundRect"
    ? buildRoundRectMarkup(frameW, frameH, "white", "none", 0, 0, 0, frame.roundRectAdj)
    : buildEllipseMarkup(frameW, frameH, "white", "none", 0);

  const innerHtml = buildImageContentHtml(el, 0, 0, frameW, frameH);

  return `<svg xmlns="http://www.w3.org/2000/svg" style="
    position: absolute;
    left: ${frameX}px;
    top: ${frameY}px;
    width: ${frameW}px;
    height: ${frameH}px;
    overflow: visible;
    ${rotationStyle}
    ${opacity}
  " viewBox="0 0 ${frameW} ${frameH}" preserveAspectRatio="none">
    <defs>
      ${filterDef}
      <clipPath id="${clipId}" clipPathUnits="userSpaceOnUse">${clipMarkup}</clipPath>
    </defs>
    ${fillColor !== "none" || strokeColor !== "none" || frame.shadow
      ? `<g ${filterAttr}>${shapeMarkup}</g>`
      : ""}
    <foreignObject x="0" y="0" width="${frameW}" height="${frameH}" clip-path="url(#${clipId})">
      <div xmlns="http://www.w3.org/1999/xhtml" style="width:${frameW}px;height:${frameH}px;overflow:hidden;margin:0;padding:0;">
        ${innerHtml}
      </div>
    </foreignObject>
  </svg>`;
}

/** Icon on a colored circle plate (slide 7 step icons). */
function renderPlateFrameImage(
  el: ImageElement,
  frame: ShapeFrameStyle,
  svgX: number,
  svgY: number,
  svgW: number,
  svgH: number,
  ovalX: number,
  ovalY: number,
  ovalW: number,
  ovalH: number,
  rotationStyle: string,
  opacity: string,
): string {
  const filterId = frame.shadow ? randomId("fs") : "";
  const filterDef = frame.shadow ? shadowToSvgFilterDef(frame.shadow, filterId) : "";
  const filterAttr = frame.shadow ? `filter="url(#${filterId})"` : "";

  const fillColor = frame.fillColor && frame.fillColor !== "transparent" ? frame.fillColor : "none";
  const strokeColor = frame.borderColor && frame.borderColor !== "transparent" ? frame.borderColor : "none";
  const sw = frame.strokeWidth && frame.strokeWidth > 0 ? frame.strokeWidth : 0;
  const shadowFill = frame.shadow ? shadowCasterFill(fillColor) : fillColor;

  const ovalMarkup = frame.shapeType === "roundRect"
    ? buildRoundRectMarkup(ovalW, ovalH, shadowFill, strokeColor, sw, ovalX, ovalY, frame.roundRectAdj)
    : buildEllipseMarkup(ovalW, ovalH, shadowFill, strokeColor, sw, ovalX, ovalY);

  const innerHtml = buildImageContentHtml(el, 0, 0, svgW, svgH);

  return `<svg xmlns="http://www.w3.org/2000/svg" style="
    position: absolute;
    left: ${svgX}px;
    top: ${svgY}px;
    width: ${svgW}px;
    height: ${svgH}px;
    overflow: visible;
    ${rotationStyle}
    ${opacity}
  " viewBox="0 0 ${Math.max(svgW, ovalX + ovalW)} ${Math.max(svgH, ovalY + ovalH)}" preserveAspectRatio="none">
    <defs>${filterDef}</defs>
    <g ${filterAttr}>${ovalMarkup}</g>
    <foreignObject x="0" y="0" width="${svgW}" height="${svgH}">
      <div xmlns="http://www.w3.org/1999/xhtml" style="width:${svgW}px;height:${svgH}px;overflow:visible;margin:0;padding:0;">
        ${innerHtml}
      </div>
    </foreignObject>
  </svg>`;
}

function buildImageContentHtml(
  el: ImageElement,
  originX: number,
  originY: number,
  frameW: number,
  frameH: number,
): string {
  if (el.crop) {
    const layout = computeCroppedImageLayout({
      frameW,
      frameH,
      crop: el.crop,
      fill: el.fill,
      naturalSize: el.naturalSize,
      preserveAspectRatio: el.preserveAspectRatio,
    });
    if (layout) {
      const { imgW, imgH, imgLeft, imgTop } = layout;
      return `<img src="${el.src}" style="position:absolute;left:${originX + imgLeft}px;top:${originY + imgTop}px;width:${imgW}px;height:${imgH}px;max-width:none;max-height:none;" />`;
    }
  }

  return `<img src="${el.src}" style="position:absolute;left:${originX}px;top:${originY}px;width:${frameW}px;height:${frameH}px;object-fit:cover;" />`;
}

function buildEllipseMarkup(
  w: number,
  h: number,
  fill: string,
  stroke: string,
  sw: number,
  ox = 0,
  oy = 0,
): string {
  const cx = ox + w / 2;
  const cy = oy + h / 2;
  const rx = w / 2;
  const ry = h / 2;
  return `<ellipse cx="${cx}" cy="${cy}" rx="${rx}" ry="${ry}" fill="${fill}" stroke="${stroke}" stroke-width="${sw}" />`;
}

function buildRoundRectMarkup(
  w: number,
  h: number,
  fill: string,
  stroke: string,
  sw: number,
  ox = 0,
  oy = 0,
  adj?: number,
): string {
  const rx = roundRectCornerRadiusPx(w, h, adj);
  return `<rect x="${ox}" y="${oy}" width="${w}" height="${h}" rx="${rx}" ry="${rx}" fill="${fill}" stroke="${stroke}" stroke-width="${sw}" />`;
}

/** Render a:srcRect crop scaled into a:fillRect (a:stretch) within the shape frame. */
function renderCroppedImage(
  el: ImageElement,
  x: number,
  y: number,
  width: number,
  height: number,
  rotationStyle: string,
  opacity: string,
): string {
  const layout = computeCroppedImageLayout({
    frameW: width,
    frameH: height,
    crop: el.crop!,
    fill: el.fill,
    naturalSize: el.naturalSize,
    preserveAspectRatio: el.preserveAspectRatio,
  });

  if (!layout) {
    return `<img src="${el.src}" style="position:absolute;left:${x}px;top:${y}px;width:${width}px;height:${height}px;${rotationStyle}" />`;
  }

  const { imgW, imgH, imgLeft, imgTop } = layout;

  return `<div style="
    position: absolute;
    left: ${x}px;
    top: ${y}px;
    width: ${width}px;
    height: ${height}px;
    overflow: hidden;
    ${rotationStyle}
    ${opacity}
  "><img src="${el.src}" style="
    position: absolute;
    left: ${imgLeft}px;
    top: ${imgTop}px;
    width: ${imgW}px;
    height: ${imgH}px;
    max-width: none;
    max-height: none;
  " /></div>`;
}
