import JSZip from "jszip";
import { XmlHelper } from "./XmlHelper";
import { TextExtractor, PlaceholderDefaults } from "../elements/TextExtractor";
import { ImageExtractor, ThemeImageContext } from "../elements/ImageExtractor";
import { ShapeExtractor } from "../elements/ShapeExtractor";
import { TableExtractor } from "../elements/TableExtractor";
import { ChartExtractor } from "../elements/ChartExtractor";
import { SlideElement } from "../models/SlideElement";

/**
 * Responsible for extracting all slides from the .pptx file as lists of SlideElement.
 */
export class SlideExtractor {
  constructor(
    private zip: JSZip,
    private options: { imageSource?: "data-uri" | "zip-path" } = {}
  ) {}

  /**
   * Extracts all slides in order and parses their visual elements.
   * @returns An array of SlideElement lists (one per slide).
   */
  async extractSlides(): Promise<SlideElement[][]> {
    // Load theme colors
    const themeXmlStr = await this.zip.file("ppt/theme/theme1.xml")?.async("string");
    const themeXml = themeXmlStr ? XmlHelper.parseXml(themeXmlStr) : null;
    const themeColors = XmlHelper.extractThemeColors(themeXml);
    const themeTableStyles = XmlHelper.extractThemeTableStyles(themeXml);
    const themeRelsPath = "ppt/theme/_rels/theme1.xml.rels";
    const themeRelsStr = await this.zip.file(themeRelsPath)?.async("string");
    const themeRelsXml = themeRelsStr ? XmlHelper.parseXml(themeRelsStr) : null;
    const themeContext: ThemeImageContext = {
      themeDoc: themeXml,
      themeRels: themeRelsXml,
      themeBasePath: "ppt/theme",
    };

    const slidePaths = Object.keys(this.zip.files)
      .filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f))
      .sort((a, b) => {
        const numA = parseInt(a.match(/slide(\d+)\.xml$/)?.[1] || "0", 10);
        const numB = parseInt(b.match(/slide(\d+)\.xml$/)?.[1] || "0", 10);
        return numA - numB;
      });

    const slides: SlideElement[][] = [];

    for (const slidePath of slidePaths) {
      const slideXmlStr = await this.zip.file(slidePath)!.async("string");
      const slideXml = XmlHelper.parseXml(slideXmlStr);

      // Intentar encontrar el spTree usando namespace-tolerancia
      const spTree =
        slideXml.querySelector("p\\:spTree") ||
        slideXml.getElementsByTagNameNS("*", "spTree")[0];

      if (!spTree) {
        console.warn(`Warning: no <spTree> found in ${slidePath}`);
        slides.push([]);
        continue;
      }

      const relsPath = slidePath.replace("slides/", "slides/_rels/") + ".rels";
      const relsXml = this.zip.file(relsPath)
        ? XmlHelper.parseXml(await this.zip.file(relsPath)!.async("string"))
        : XmlHelper.parseXml(`<Relationships/>`);

      // Resolve layout from slide rels
      const layoutRel = XmlHelper.findRelationshipByTypeSuffix(relsXml, "/slideLayout");
      const layoutTarget = layoutRel?.getAttribute("Target") || undefined;
      let layoutSpTree: Element | null = null;
      let layoutRelsXml: Document | null = null;
      if (layoutTarget) {
        const layoutPath = this.resolvePath(layoutTarget, "ppt/slides");
        const layoutXmlStr = await this.zip.file(layoutPath)?.async("string");
        if (layoutXmlStr) {
          const layoutXml = XmlHelper.parseXml(layoutXmlStr);
          layoutSpTree = layoutXml.querySelector("p\\:spTree") || layoutXml.getElementsByTagNameNS("*", "spTree")[0] || null;
          const layoutRelsPath = layoutPath.replace("slideLayouts/", "slideLayouts/_rels/") + ".rels";
          layoutRelsXml = this.zip.file(layoutRelsPath)
            ? XmlHelper.parseXml(await this.zip.file(layoutRelsPath)!.async("string"))
            : XmlHelper.parseXml(`<Relationships/>`);
        }
      }

      // Resolve master from layout rels
      let masterSpTree: Element | null = null;
      let masterRelsXml: Document | null = null;
      if (layoutRelsXml) {
        const masterRel = XmlHelper.findRelationshipByTypeSuffix(layoutRelsXml, "/slideMaster");
        const masterTarget = masterRel?.getAttribute("Target") || undefined;
        if (masterTarget) {
          const masterPath = this.resolvePath(masterTarget, "ppt/slideLayouts");
          const masterXmlStr = await this.zip.file(masterPath)?.async("string");
          if (masterXmlStr) {
            const masterXml = XmlHelper.parseXml(masterXmlStr);
            masterSpTree = masterXml.querySelector("p\\:spTree") || masterXml.getElementsByTagNameNS("*", "spTree")[0] || null;
            const masterRelsPath = masterPath.replace("slideMasters/", "slideMasters/_rels/") + ".rels";
            masterRelsXml = this.zip.file(masterRelsPath)
              ? XmlHelper.parseXml(await this.zip.file(masterRelsPath)!.async("string"))
              : XmlHelper.parseXml(`<Relationships/>`);
          }
        }
      }

      // Extract background from slide/layout/master (slide overrides layout overrides master)
      const slideBg = await this.extractBackground(slideXml, relsXml, "ppt/slides", this.zip, themeColors, themeXml, themeRelsXml);
      const layoutBg = layoutRelsXml ? await this.extractBackground(layoutSpTree?.ownerDocument || null, layoutRelsXml, "ppt/slideLayouts", this.zip, themeColors, themeXml, themeRelsXml) : null;
      const masterBg = masterRelsXml ? await this.extractBackground(masterSpTree?.ownerDocument || null, masterRelsXml, "ppt/slideMasters", this.zip, themeColors, themeXml, themeRelsXml) : null;
      const bgElement = slideBg || layoutBg || masterBg;

      // Extract elements from master → layout → slide (respecting z-order: back to front)
      const masterText = masterSpTree ? TextExtractor.extract(masterSpTree, themeColors, { context: "master" }) : [];
      const masterImages = masterSpTree && masterRelsXml
        ? await ImageExtractor.extract(masterSpTree, masterRelsXml, this.zip, "ppt/slideMasters", this.options, themeContext)
        : [];
      const masterShapes = masterSpTree ? ShapeExtractor.extract(masterSpTree, themeColors, themeXml) : [];

      const layoutText = layoutSpTree ? TextExtractor.extract(layoutSpTree, themeColors, { context: "layout" }) : [];
      const layoutImages = layoutSpTree && layoutRelsXml
        ? await ImageExtractor.extract(layoutSpTree, layoutRelsXml, this.zip, "ppt/slideLayouts", this.options, themeContext)
        : [];
      const layoutShapes = layoutSpTree ? ShapeExtractor.extract(layoutSpTree, themeColors, themeXml) : [];

      const masterDefaults = this.extractPlaceholderDefaults(masterSpTree, themeColors);
      const layoutDefaults = this.extractPlaceholderDefaults(layoutSpTree, themeColors);
      const mergedDefaults = new Map<string, PlaceholderDefaults>(masterDefaults);
      for (const [key, layoutVal] of layoutDefaults) {
        const existing = mergedDefaults.get(key);
        if (existing) {
          const merged = { ...existing };
          for (const [k, v] of Object.entries(layoutVal)) {
            if (v !== undefined) (merged as any)[k] = v;
          }
          mergedDefaults.set(key, merged);
        } else {
          mergedDefaults.set(key, layoutVal);
        }
      }
      const slideText = TextExtractor.extract(spTree, themeColors, { context: "slide", placeholderDefaults: mergedDefaults });
      const slideImages = await ImageExtractor.extract(spTree, relsXml, this.zip, "ppt/slides", this.options, themeContext);
      const slideTables = TableExtractor.extract(spTree, themeColors, themeTableStyles);
      const slideCharts = await ChartExtractor.extract(spTree, relsXml, this.zip, themeColors);
      const slideShapes = ShapeExtractor.extract(spTree, themeColors, themeXml);

      slides.push([
        ...(bgElement ? [bgElement] : []),
        ...masterShapes,
        ...masterImages,
        ...masterText,
        ...layoutShapes,
        ...layoutImages,
        ...layoutText,
        ...slideShapes,
        ...slideTables,
        ...slideCharts,
        ...slideImages,
        ...slideText,
      ]);
    }

    return slides;
  }

  /** Normalize a relative path against a base directory inside ppt folder */
  private resolvePath(target: string, baseDir: string): string {
    const parts = (baseDir + "/" + target).split("/");
    const resolved: string[] = [];
    for (const part of parts) {
      if (part === "..") {
        if (resolved.length) resolved.pop();
      } else if (part !== "." && part !== "") {
        resolved.push(part);
      }
    }
    return resolved.join("/");
  }

  private extractPlaceholderDefaults(spTree: Element | null, themeColors: Record<string, string>): Map<string, PlaceholderDefaults> {
    const map = new Map<string, PlaceholderDefaults>();
    if (!spTree) return map;
    const shapes = spTree.getElementsByTagNameNS("*", "sp");
    for (const shape of Array.from(shapes)) {
      const nvPr = shape.getElementsByTagNameNS("*", "nvPr")[0] ?? null;
      const ph = nvPr?.getElementsByTagNameNS("*", "ph")[0] ?? null;
      if (!ph) continue;
      const type = ph.getAttribute("type") || undefined;
      const idx = ph.getAttribute("idx") || undefined;
      if (!type && !idx) continue;

      const defaults: PlaceholderDefaults = {};

      const xfrm = shape.getElementsByTagNameNS("*", "xfrm")[0] ?? null;
      const off = xfrm?.getElementsByTagNameNS("*", "off")[0] ?? null;
      const ext = xfrm?.getElementsByTagNameNS("*", "ext")[0] ?? null;
      if (off && ext) {
        defaults.x = XmlHelper.getAttrAsNumber(off, "x");
        defaults.y = XmlHelper.getAttrAsNumber(off, "y");
        defaults.cx = XmlHelper.getAttrAsNumber(ext, "cx");
        defaults.cy = XmlHelper.getAttrAsNumber(ext, "cy");
      }

      const txBody = shape.getElementsByTagNameNS("*", "txBody")[0] ?? null;
      const bodyPr = txBody?.getElementsByTagNameNS("*", "bodyPr")[0] ?? null;
      if (bodyPr) {
        const anchor = bodyPr.getAttribute("anchor");
        if (anchor) defaults.anchor = anchor;
        const lI = bodyPr.getAttribute("lIns"); if (lI) defaults.lIns = lI;
        const tI = bodyPr.getAttribute("tIns"); if (tI) defaults.tIns = tI;
        const rI = bodyPr.getAttribute("rIns"); if (rI) defaults.rIns = rI;
        const bI = bodyPr.getAttribute("bIns"); if (bI) defaults.bIns = bI;
      }

      const lstStyle = txBody?.querySelector("*|lstStyle");
      const lvl1pPr = lstStyle?.querySelector("*|lvl1pPr");
      if (lvl1pPr) {
        const algn = lvl1pPr.getAttribute("algn");
        if (algn) defaults.align = algn;

        const defRPr = lvl1pPr.querySelector("*|defRPr");
        if (defRPr) {
          const sz = defRPr.getAttribute("sz");
          if (sz) { const n = parseInt(sz, 10); if (Number.isFinite(n)) defaults.fontSize = n / 100; }
          if (defRPr.getAttribute("b") === "1") defaults.bold = true;
          if (defRPr.getAttribute("i") === "1") defaults.italic = true;
          const solidFill = defRPr.querySelector("*|solidFill");
          const color = XmlHelper.getColorFromElement(solidFill || null, themeColors);
          if (color) defaults.color = color;
          const latin = defRPr.getElementsByTagNameNS("*", "latin")[0];
          const fontFamily = latin?.getAttribute("typeface");
          if (fontFamily) defaults.fontFamily = fontFamily;
        }
      }

      if (type) map.set(`type:${type}`, defaults);
      if (idx) map.set(`idx:${idx}`, defaults);
    }
    return map;
  }

  private async extractBackground(
    doc: Document | null,
    rels: Document | null,
    baseDir: string,
    zip: JSZip,
    themeColors: Record<string, string>,
    themeDoc: Document | null = null,
    themeRels: Document | null = null
  ): Promise<SlideElement | null> {
    if (!doc) return null;
    const bg = doc.getElementsByTagNameNS("*", "bg")[0];
    if (!bg) return null;

    const bgPr = bg.getElementsByTagNameNS("*", "bgPr")[0] || null;
    const bgRef = bg.getElementsByTagNameNS("*", "bgRef")[0] || null;

    // Direct solid fill on bgPr
    const solidFill = bgPr?.getElementsByTagNameNS("*", "solidFill")[0] || null;
    const color = XmlHelper.getColorFromElement(solidFill, themeColors);
    if (color) {
      return { type: "background", fillColor: color } as SlideElement;
    }

    // Direct image fill on bgPr
    const bgPrImage = await this.backgroundFromBlipFill(
      bgPr?.getElementsByTagNameNS("*", "blipFill")[0] || null,
      rels,
      baseDir,
      zip
    );
    if (bgPrImage) return bgPrImage;

    // bgRef → theme style matrix (common on slide masters)
    if (bgRef && themeDoc) {
      const idx = parseInt(bgRef.getAttribute("idx") || "0", 10);
      if (idx > 0 && idx !== 1000) {
        const themeFill = XmlHelper.getThemeFillElement(themeDoc, idx);
        if (themeFill) {
          if (themeFill.localName === "solidFill") {
            const themeColor = XmlHelper.getColorFromElement(themeFill, themeColors);
            if (themeColor) {
              return { type: "background", fillColor: themeColor } as SlideElement;
            }
          }
          if (themeFill.localName === "blipFill" && themeRels) {
            const themeImage = await this.backgroundFromBlipFill(
              themeFill,
              themeRels,
              "ppt/theme",
              zip
            );
            if (themeImage) return themeImage;
          }
        }
      }

      // Fallback: scheme tint only (no picture resolved)
      const schemeClr = bgRef.getElementsByTagNameNS("*", "schemeClr")[0] || null;
      const schemeVal = schemeClr?.getAttribute("val") || undefined;
      if (schemeVal && themeColors[schemeVal]) {
        return { type: "background", fillColor: themeColors[schemeVal] } as SlideElement;
      }
    }

    return null;
  }

  private async backgroundFromBlipFill(
    blipFill: Element | null,
    rels: Document | null,
    baseDir: string,
    zip: JSZip
  ): Promise<SlideElement | null> {
    const blip = blipFill?.getElementsByTagNameNS("*", "blip")[0] || null;
    if (!blip || !rels) return null;

    const imageSrc = await ImageExtractor.resolveBlip(blip, rels, zip, baseDir, this.options);
    if (!imageSrc) return null;

    return { type: "background", imageSrc } as SlideElement;
  }
}
