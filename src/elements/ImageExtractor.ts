import { ImageElement, ImageCrop, ShapeFrameStyle } from "../models/SlideElement";
import { XmlHelper } from "../core/XmlHelper";
import { parseCustGeom } from "../core/custGeom";
import { readImageDimensions } from "../core/imageDimensions";
import { extractShapeStyle, hasVisibleShapeFrame, readRoundRectAdj } from "../core/shapeStyle";
import JSZip from "jszip";

export interface ResolvedBlip {
  src: string;
  naturalSize?: { width: number; height: number };
}

export interface ThemeImageContext {
  themeDoc: Document | null;
  themeRels: Document | null;
  themeBasePath?: string;
  themeColors?: Record<string, string>;
}

/**
 * Responsible for extracting image elements from a slide XML node.
 */
export class ImageExtractor {
  /**
   * Extracts image elements from the <spTree> element using rels from slide relationships.
   * @param spTree The <spTree> element of the slide.
   * @param rels XML Document for slide relationships (ppt/slides/_rels/slideX.xml.rels).
   * @param zip The JSZip archive of the entire .pptx file.
   * @returns List of ImageElement extracted.
   */
  static async extract(
    spTree: Element | null,
    rels: Document,
    zip: JSZip,
    basePath: string = "ppt/slides",
    options: { imageSource?: "data-uri" | "zip-path" } = {},
    themeContext?: ThemeImageContext
  ): Promise<ImageElement[]> {
    if (!spTree) return [];

    const elements: ImageElement[] = [];

    const pics = spTree.getElementsByTagNameNS("*", "pic");
    for (const pic of Array.from(pics)) {
      const el = await this.extractFromPic(pic, rels, zip, basePath, options, themeContext);
      if (el) elements.push(el);
    }

    const shapes = spTree.getElementsByTagNameNS("*", "sp");
    for (const shape of Array.from(shapes)) {
      const el = await this.extractFromShape(shape, rels, zip, basePath, options, themeContext);
      if (el) elements.push(el);
    }

    return elements;
  }

  static async extractFromPic(
    pic: Element,
    rels: Document,
    zip: JSZip,
    basePath: string,
    options: { imageSource?: "data-uri" | "zip-path" },
    themeContext?: ThemeImageContext,
  ): Promise<ImageElement | null> {
    const blipFill = pic.getElementsByTagNameNS("*", "blipFill")[0];
    const blip = blipFill?.getElementsByTagNameNS("*", "blip")[0]
      ?? pic.getElementsByTagNameNS("*", "blip")[0];
    if (!blip) return null;

    const resolved = await this.resolveBlip(blip, rels, zip, basePath, options);
    if (!resolved) return null;

    const xfrm = pic.getElementsByTagNameNS("*", "xfrm")[0];
    const { x, y, cx, cy, rotationDeg, flipH, flipV } = this.readTransform(xfrm);
    const embedId = blip.getAttribute("r:embed") ?? blip.getAttribute("r:link") ?? "";
    const spPr = pic.getElementsByTagNameNS("*", "spPr")[0];
    const customGeometry = parseCustGeom(spPr ?? null) ?? undefined;
    const { crop, fill, preserveAspectRatio, imageOpacity } = this.parseBlipFill(blipFill, pic);
    const frame = this.extractFrameFromSpPr(
      pic,
      spPr,
      { x, y, width: cx, height: cy },
      themeContext,
      false,
    );

    return {
      type: "image",
      relId: embedId,
      src: resolved.src,
      position: { x, y },
      size: { width: cx, height: cy },
      rotationDeg,
      crop,
      fill,
      naturalSize: resolved.naturalSize,
      preserveAspectRatio,
      frame,
      customGeometry,
      imageOpacity,
      flipH,
      flipV,
    };
  }

  static async extractFromShape(
    shape: Element,
    rels: Document,
    zip: JSZip,
    basePath: string,
    options: { imageSource?: "data-uri" | "zip-path" },
    themeContext?: ThemeImageContext
  ): Promise<ImageElement | null> {
    const spPr = shape.getElementsByTagNameNS("*", "spPr")[0];
    let blipFill = spPr?.getElementsByTagNameNS("*", "blipFill")[0] ?? null;
    let blipRels = rels;
    let blipBase = basePath;

    if (!blipFill && themeContext?.themeDoc && themeContext.themeRels) {
      const style = shape.getElementsByTagNameNS("*", "style")[0];
      const fillRef = style?.getElementsByTagNameNS("*", "fillRef")[0];
      const idx = parseInt(fillRef?.getAttribute("idx") || "0", 10);
      if (idx > 0 && idx !== 1000) {
        const themeFill = XmlHelper.getThemeFillElement(themeContext.themeDoc, idx);
        if (themeFill?.localName === "blipFill") {
          blipFill = themeFill;
          blipRels = themeContext.themeRels;
          blipBase = themeContext.themeBasePath ?? "ppt/theme";
        }
      }
    }

    if (!blipFill) return null;

    const blip = blipFill.getElementsByTagNameNS("*", "blip")[0];
    if (!blip) return null;

    const resolved = await this.resolveBlip(blip, blipRels, zip, blipBase, options);
    if (!resolved) return null;

    const xfrm = shape.getElementsByTagNameNS("*", "xfrm")[0];
    const { x, y, cx, cy, rotationDeg, flipH, flipV } = this.readTransform(xfrm);
    const embedId = blip.getAttribute("r:embed") ?? blip.getAttribute("r:link") ?? "";
    const customGeometry = parseCustGeom(spPr ?? null) ?? undefined;
    const { crop, fill, preserveAspectRatio, imageOpacity } = this.parseBlipFill(blipFill);
    const frame = this.extractFrameFromSpPr(
      shape,
      spPr,
      { x, y, width: cx, height: cy },
      themeContext,
      true,
    );

    return {
      type: "image",
      relId: embedId,
      src: resolved.src,
      position: { x, y },
      size: { width: cx, height: cy },
      rotationDeg,
      crop,
      fill,
      naturalSize: resolved.naturalSize,
      preserveAspectRatio,
      frame,
      customGeometry,
      imageOpacity,
      flipH,
      flipV,
    };
  }

  static async resolveBlip(
    blip: Element,
    rels: Document | null,
    zip: JSZip,
    basePath: string,
    options: { imageSource?: "data-uri" | "zip-path" } = {}
  ): Promise<ResolvedBlip | null> {
    if (!rels) return null;

    const embedId = blip.getAttribute("r:embed") ?? blip.getAttribute("r:link") ?? "";
    if (!embedId) return null;

    const rel = XmlHelper.findRelationshipById(rels, embedId);
    const target = rel?.getAttribute("Target");
    if (!target) return null;

    const candidates = [
      this.normalizePath(target, basePath),
      target.replace(/^\.\.\//, ""),
      target.startsWith("/") ? target.slice(1) : null,
    ].filter((p): p is string => Boolean(p));

    const baseName = target.split("/").pop();
    if (baseName) {
      candidates.push(`ppt/media/${baseName}`);
    }

    let imageFile: JSZip.JSZipObject | null = null;
    let normalizedPath = candidates[0];
    for (const candidate of candidates) {
      const file = zip.file(candidate);
      if (file) {
        imageFile = file;
        normalizedPath = candidate;
        break;
      }
    }

    if (!imageFile && baseName) {
      const match = Object.keys(zip.files).find(
        (p) => p.endsWith(`/${baseName}`) && !zip.files[p].dir
      );
      if (match) {
        imageFile = zip.file(match)!;
        normalizedPath = match;
      }
    }

    if (!imageFile) return null;

    const bytes = await imageFile.async("uint8array");
    const naturalSize = readImageDimensions(bytes);
    const src = options.imageSource === "zip-path"
      ? normalizedPath
      : `data:image/${normalizedPath.split(".").pop()?.toLowerCase() || "png"};base64,${bytesToBase64(bytes)}`;

    return { src, naturalSize };
  }

  private static extractFrameFromSpPr(
    owner: Element,
    spPr: Element | null | undefined,
    bounds: { x: number; y: number; width: number; height: number },
    themeContext: ThemeImageContext | undefined,
    pictureFill: boolean,
  ): ShapeFrameStyle | undefined {
    const themeColors = themeContext?.themeColors ?? {};
    const style = extractShapeStyle(owner, spPr, themeColors, themeContext?.themeDoc ?? null, {
      pictureFill,
    });
    if (!hasVisibleShapeFrame(style)) return undefined;

    return {
      shapeType: style.shapeType,
      fillColor: style.fillColor,
      borderColor: style.borderColor,
      strokeWidth: style.strokeWidth,
      position: { x: bounds.x, y: bounds.y },
      size: { width: bounds.width, height: bounds.height },
      shadow: style.shadow,
      clipImage: style.shapeType === "ellipse" || style.shapeType === "roundRect",
      roundRectAdj: style.shapeType === "roundRect" ? readRoundRectAdj(spPr) : undefined,
    };
  }

  private static parseBlipFill(
    blipFill: Element | null | undefined,
    pic?: Element,
  ): { crop?: ImageCrop; fill?: ImageCrop; preserveAspectRatio?: boolean; imageOpacity?: number } {
    const blip = blipFill?.getElementsByTagNameNS("*", "blip")[0];
    const alphaModFix = blip?.getElementsByTagNameNS("*", "alphaModFix")[0];
    const alphaAmt = alphaModFix ? Number(alphaModFix.getAttribute("amt") ?? 100000) : undefined;
    const imageOpacity = alphaAmt !== undefined && Number.isFinite(alphaAmt) ? alphaAmt / 100000 : undefined;

    const crop = this.parseRelativeRect(blipFill?.getElementsByTagNameNS("*", "srcRect")[0]);
    const stretchEl = blipFill?.getElementsByTagNameNS("*", "stretch")[0];
    const fill = stretchEl
      ? this.parseRelativeRect(stretchEl.getElementsByTagNameNS("*", "fillRect")[0]) ?? { left: 0, top: 0, right: 0, bottom: 0 }
      : undefined;

    const picLocks = pic?.getElementsByTagNameNS("*", "picLocks")[0]
      ?? pic?.querySelector("*|cNvPicPr *|picLocks");
    const preserveAspectRatio = picLocks?.getAttribute("noChangeAspect") === "1" ? true : undefined;

    return { crop, fill, preserveAspectRatio, imageOpacity };
  }

  /** Parse a:srcRect / a:fillRect insets (percent * 1000 → fraction). */
  private static parseRelativeRect(rect: Element | null | undefined): ImageCrop | undefined {
    if (!rect) return undefined;

    const left = Number(rect.getAttribute("l") || 0) / 100000;
    const top = Number(rect.getAttribute("t") || 0) / 100000;
    const right = Number(rect.getAttribute("r") || 0) / 100000;
    const bottom = Number(rect.getAttribute("b") || 0) / 100000;
    if (left === 0 && top === 0 && right === 0 && bottom === 0) return undefined;
    return { left, top, right, bottom };
  }

  private static readTransform(
    xfrm: Element | null | undefined
  ): { x: number; y: number; cx: number; cy: number; rotationDeg?: number; flipH?: boolean; flipV?: boolean } {
    const off = xfrm?.getElementsByTagNameNS("*", "off")[0];
    const ext = xfrm?.getElementsByTagNameNS("*", "ext")[0];
    const rotAttr = xfrm?.getAttribute("rot");
    const rotationDeg = rotAttr ? Number(rotAttr) / 60000 : undefined;
    return {
      x: off ? XmlHelper.getAttrAsNumber(off, "x") : 0,
      y: off ? XmlHelper.getAttrAsNumber(off, "y") : 0,
      cx: ext ? XmlHelper.getAttrAsNumber(ext, "cx") : 1000000,
      cy: ext ? XmlHelper.getAttrAsNumber(ext, "cy") : 500000,
      rotationDeg: rotationDeg && !isNaN(rotationDeg) ? rotationDeg : undefined,
      flipH: xfrm?.getAttribute("flipH") === "1" || undefined,
      flipV: xfrm?.getAttribute("flipV") === "1" || undefined,
    };
  }

  /**
   * Normalizes a relative path from a slide rels file.
   * @param target Path from the relationship XML (e.g. "../media/image1.png")
   * @param basePath Base folder (e.g. "ppt/slides")
   * @returns Normalized path inside the zip (e.g. "ppt/media/image1.png")
   */
  private static normalizePath(target: string, basePath: string): string {
    const parts = (basePath + "/" + target).split("/");
    const resolved: string[] = [];
    for (const part of parts) {
      if (part === "..") resolved.pop();
      else if (part !== ".") resolved.push(part);
    }
    return resolved.join("/");
  }
}

function bytesToBase64(bytes: Uint8Array): string {
  if (typeof Buffer !== "undefined") {
    return Buffer.from(bytes).toString("base64");
  }
  let binary = "";
  for (let i = 0; i < bytes.length; i++) binary += String.fromCharCode(bytes[i]!);
  return btoa(binary);
}
