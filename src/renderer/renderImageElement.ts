import { ImageElement } from "../models/SlideElement";
import { emuToPx, renderCustGeomSvg } from "./renderCustGeom";

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

  if (el.customGeometry) {
    return renderCustGeomSvg({
      x,
      y,
      width,
      height,
      geom: el.customGeometry,
      fill: "transparent",
      imageSrc: el.src,
    });
  }

  return `<img src="${el.src}" style="
    position: absolute;
    left: ${x}px;
    top: ${y}px;
    width: ${width}px;
    height: ${height}px;
    object-fit: cover;
  " />`;
}
