import { ImageElement } from "../models/SlideElement";
import { XmlHelper } from "../core/XmlHelper";
import { parseCustGeom } from "../core/custGeom";
import JSZip from "jszip";

export interface ThemeImageContext {
  themeDoc: Document | null;
  themeRels: Document | null;
  themeBasePath?: string;
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
      const el = await this.extractFromPic(pic, rels, zip, basePath, options);
      if (el) elements.push(el);
    }

    const shapes = spTree.getElementsByTagNameNS("*", "sp");
    for (const shape of Array.from(shapes)) {
      const el = await this.extractFromShape(shape, rels, zip, basePath, options, themeContext);
      if (el) elements.push(el);
    }

    return elements;
  }

  private static async extractFromPic(
    pic: Element,
    rels: Document,
    zip: JSZip,
    basePath: string,
    options: { imageSource?: "data-uri" | "zip-path" }
  ): Promise<ImageElement | null> {
    const blipFill = pic.getElementsByTagNameNS("*", "blipFill")[0];
    const blip = blipFill?.getElementsByTagNameNS("*", "blip")[0]
      ?? pic.getElementsByTagNameNS("*", "blip")[0];
    if (!blip) return null;

    const src = await this.resolveBlip(blip, rels, zip, basePath, options);
    if (!src) return null;

    const xfrm = pic.getElementsByTagNameNS("*", "xfrm")[0];
    const { x, y, cx, cy } = this.readTransform(xfrm);
    const embedId = blip.getAttribute("r:embed") ?? blip.getAttribute("r:link") ?? "";
    const spPr = pic.getElementsByTagNameNS("*", "spPr")[0];
    const customGeometry = parseCustGeom(spPr ?? null) ?? undefined;

    return {
      type: "image",
      relId: embedId,
      src,
      position: { x, y },
      size: { width: cx, height: cy },
      customGeometry,
    };
  }

  private static async extractFromShape(
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

    const src = await this.resolveBlip(blip, blipRels, zip, blipBase, options);
    if (!src) return null;

    const xfrm = shape.getElementsByTagNameNS("*", "xfrm")[0];
    const { x, y, cx, cy } = this.readTransform(xfrm);
    const embedId = blip.getAttribute("r:embed") ?? blip.getAttribute("r:link") ?? "";

    const customGeometry = parseCustGeom(spPr ?? null) ?? undefined;

    return {
      type: "image",
      relId: embedId,
      src,
      position: { x, y },
      size: { width: cx, height: cy },
      customGeometry,
    };
  }

  static async resolveBlip(
    blip: Element,
    rels: Document | null,
    zip: JSZip,
    basePath: string,
    options: { imageSource?: "data-uri" | "zip-path" } = {}
  ): Promise<string | null> {
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

    return options.imageSource === "zip-path"
      ? normalizedPath
      : await this.toDataUri(imageFile, normalizedPath);
  }

  private static readTransform(xfrm: Element | null | undefined): { x: number; y: number; cx: number; cy: number } {
    const off = xfrm?.getElementsByTagNameNS("*", "off")[0];
    const ext = xfrm?.getElementsByTagNameNS("*", "ext")[0];
    return {
      x: off ? XmlHelper.getAttrAsNumber(off, "x") : 0,
      y: off ? XmlHelper.getAttrAsNumber(off, "y") : 0,
      cx: ext ? XmlHelper.getAttrAsNumber(ext, "cx") : 1000000,
      cy: ext ? XmlHelper.getAttrAsNumber(ext, "cy") : 500000,
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

  private static async toDataUri(imageFile: JSZip.JSZipObject, normalizedPath: string): Promise<string> {
    const binary = await imageFile.async("base64");
    const extImg = normalizedPath.split(".").pop()?.toLowerCase() || "png";
    return `data:image/${extImg};base64,${binary}`;
  }
}
