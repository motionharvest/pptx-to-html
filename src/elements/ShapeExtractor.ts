import { ShapeElement, LineElement } from "../models/SlideElement";
import { XmlHelper } from "../core/XmlHelper";
import { parseCustGeom } from "../core/custGeom";
import { extractShapeStyle, readRoundRectAdj } from "../core/shapeStyle";

const STRAIGHT_LINE_PRESETS = new Set(["line", "straightConnector1"]);

/**
 * Responsible for extracting shape elements (including connectors) from a slide XML node.
 */
export class ShapeExtractor {
  /**
   * Extracts shape and connector elements from the <spTree> element (legacy batch order).
   */
  static extract(
    spTree: Element | null,
    themeColors: Record<string, string>,
    themeDoc: Document | null = null
  ): (ShapeElement | LineElement)[] {
    if (!spTree) return [];

    const elements: (ShapeElement | LineElement)[] = [];
    const allShapes = [
      ...Array.from(spTree.getElementsByTagNameNS("*", "sp")),
      ...Array.from(spTree.getElementsByTagNameNS("*", "cxnSp")),
    ];

    for (const shape of allShapes) {
      const el =
        shape.localName === "cxnSp"
          ? this.extractFromCxnSp(shape, themeColors, themeDoc)
          : this.extractFromSp(shape, themeColors, themeDoc);
      if (el) elements.push(el);
    }

    return elements;
  }

  /** Extract a p:sp as a shape or straight line (null when handled by ImageExtractor). */
  static extractFromSp(
    shape: Element,
    themeColors: Record<string, string>,
    themeDoc: Document | null = null,
  ): ShapeElement | LineElement | null {
    const spPr = shape.getElementsByTagNameNS("*", "spPr")[0];
    const customGeometry = parseCustGeom(spPr ?? null) ?? undefined;
    const hasPictureFill = XmlHelper.shapeHasPictureFill(shape, themeDoc);
    if (hasPictureFill && !customGeometry) {
      return null;
    }

    return this.buildShapeOrLine(shape, spPr ?? null, customGeometry, themeColors, themeDoc);
  }

  /**
   * Shadow-only p:pic (roundRect/ellipse frame with a:outerShdw, no blipFill).
   * NY Proposal headshots use a separate pic under the photo for the drop shadow.
   */
  static extractFromPicFrame(
    pic: Element,
    themeColors: Record<string, string>,
    themeDoc: Document | null = null,
  ): ShapeElement | null {
    const spPr = pic.getElementsByTagNameNS("*", "spPr")[0];
    if (!spPr) return null;

    const blipFill = spPr.getElementsByTagNameNS("*", "blipFill")[0];
    if (blipFill?.getElementsByTagNameNS("*", "blip")[0]) return null;

    const style = extractShapeStyle(pic, spPr, themeColors, themeDoc);
    if (style.shapeType !== "ellipse" && style.shapeType !== "roundRect") return null;
    if (
      style.fillColor === "transparent"
      && style.borderColor === undefined
      && style.shadow === undefined
    ) {
      return null;
    }

    const xfrm = spPr.getElementsByTagNameNS("*", "xfrm")[0]
      ?? pic.getElementsByTagNameNS("*", "xfrm")[0];
    const off = xfrm?.getElementsByTagNameNS("*", "off")[0];
    const ext = xfrm?.getElementsByTagNameNS("*", "ext")[0];
    const x = off ? XmlHelper.getAttrAsNumber(off, "x") : 0;
    const y = off ? XmlHelper.getAttrAsNumber(off, "y") : 0;
    const cx = ext ? XmlHelper.getAttrAsNumber(ext, "cx") : 1000000;
    const cy = ext ? XmlHelper.getAttrAsNumber(ext, "cy") : 500000;
    const rotAttr = xfrm?.getAttribute("rot");
    const rotationDeg = rotAttr ? Number(rotAttr) / 60000 : undefined;

    return {
      type: "shape",
      shapeType: style.shapeType,
      position: { x, y },
      size: { width: cx, height: cy },
      fillColor: style.fillColor,
      borderColor: style.borderColor,
      strokeWidth: style.strokeWidth,
      rotationDeg: rotationDeg && !isNaN(rotationDeg) ? rotationDeg : undefined,
      roundRectAdj:
        style.shapeType === "roundRect" ? readRoundRectAdj(spPr) : undefined,
      shadow: style.shadow,
    };
  }

  /** Extract a p:cxnSp connector as a line element. */
  static extractFromCxnSp(
    cxnSp: Element,
    themeColors: Record<string, string>,
    themeDoc: Document | null = null,
  ): LineElement | null {
    const spPr = cxnSp.getElementsByTagNameNS("*", "spPr")[0];
    const customGeometry = parseCustGeom(spPr ?? null) ?? undefined;
    const built = this.buildShapeOrLine(cxnSp, spPr ?? null, customGeometry, themeColors, themeDoc);
    return built?.type === "line" ? built : null;
  }

  private static buildShapeOrLine(
    shape: Element,
    spPr: Element | null,
    customGeometry: ReturnType<typeof parseCustGeom> | undefined,
    themeColors: Record<string, string>,
    themeDoc: Document | null,
  ): ShapeElement | LineElement | null {
    const xfrm = shape.getElementsByTagNameNS("*", "xfrm")[0];
    const off = xfrm?.getElementsByTagNameNS("*", "off")[0];
    const ext = xfrm?.getElementsByTagNameNS("*", "ext")[0];

    const x = off ? XmlHelper.getAttrAsNumber(off, "x") : 0;
    const y = off ? XmlHelper.getAttrAsNumber(off, "y") : 0;
    const cx = ext ? XmlHelper.getAttrAsNumber(ext, "cx") : 1000000;
    const cy = ext ? XmlHelper.getAttrAsNumber(ext, "cy") : 500000;
    const rotAttr = xfrm?.getAttribute("rot");
    const rotationDeg = rotAttr ? Number(rotAttr) / 60000 : undefined;
    const flipH = xfrm?.getAttribute("flipH") === "1";
    const flipV = xfrm?.getAttribute("flipV") === "1";

    const shapeType = customGeometry ? "custom" : undefined;
    const style = extractShapeStyle(shape, spPr, themeColors, themeDoc, { shapeType });
    const resolvedShapeType = customGeometry ? "custom" : style.shapeType;

    if (STRAIGHT_LINE_PRESETS.has(resolvedShapeType)) {
      const lineColor = style.borderColor ?? style.fillColor;
      return {
        type: "line",
        position: { x, y },
        size: { width: cx, height: cy },
        color: lineColor !== "transparent" ? lineColor : "#000000",
        strokeWidth: style.strokeWidth,
        dashStyle: style.dashStyle,
        rotationDeg,
        flipH: flipH || undefined,
        flipV: flipV || undefined,
        headEnd: style.headEnd,
        tailEnd: style.tailEnd,
      };
    }

    return {
      type: "shape",
      shapeType: resolvedShapeType,
      position: { x, y },
      size: { width: cx, height: cy },
      fillColor: style.fillColor,
      borderColor: style.borderColor,
      strokeWidth: style.strokeWidth,
      rotationDeg,
      headEnd: style.headEnd,
      tailEnd: style.tailEnd,
      dashStyle: style.dashStyle,
      customGeometry: customGeometry ?? undefined,
      roundRectAdj:
        resolvedShapeType === "roundRect" ? readRoundRectAdj(spPr) : undefined,
      shadow: style.shadow,
    };
  }
}
