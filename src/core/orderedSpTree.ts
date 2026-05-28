import JSZip from "jszip";
import { SlideElement, ShapeElement, ImageElement } from "../models/SlideElement";
import { XmlHelper } from "./XmlHelper";
import { TextExtractor, PlaceholderDefaults, MasterTextStyles } from "../elements/TextExtractor";
import { ImageExtractor, ThemeImageContext } from "../elements/ImageExtractor";
import { ShapeExtractor } from "../elements/ShapeExtractor";
import { TableExtractor } from "../elements/TableExtractor";
import { ChartExtractor } from "../elements/ChartExtractor";

const SP_TREE_SKIP = new Set(["nvGrpSpPr", "grpSpPr"]);

export interface OrderedSpTreeContext {
  themeColors: Record<string, string>;
  themeDoc: Document | null;
  rels: Document;
  zip: JSZip;
  basePath: string;
  options: { imageSource?: "data-uri" | "zip-path" };
  themeContext: ThemeImageContext;
  themeTableStyles?: Record<string, { fills: Record<string, string>; fontColors: Record<string, string> }>;
  textOpts?: {
    context?: "slide" | "layout" | "master";
    placeholderDefaults?: Map<string, PlaceholderDefaults>;
    masterTextStyles?: MasterTextStyles;
  };
}

/**
 * Walk p:spTree children in document order (PowerPoint z-order: back → front).
 */
export async function extractOrderedSpTree(
  spTree: Element | null,
  ctx: OrderedSpTreeContext,
): Promise<SlideElement[]> {
  if (!spTree) return [];
  const elements: SlideElement[] = [];
  await walkSpTreeNode(spTree, ctx, elements);
  return elements;
}

async function walkSpTreeNode(
  node: Element,
  ctx: OrderedSpTreeContext,
  out: SlideElement[],
): Promise<void> {
  for (const child of Array.from(node.children)) {
    const tag = child.localName;
    if (SP_TREE_SKIP.has(tag)) continue;

    switch (tag) {
      case "grpSp":
        await walkSpTreeNode(child, ctx, out);
        break;

      case "pic": {
        const image = await ImageExtractor.extractFromPic(
          child,
          ctx.rels,
          ctx.zip,
          ctx.basePath,
          ctx.options,
          ctx.themeContext,
        );
        if (image) {
          out.push(image);
          break;
        }

        const frameShape = ShapeExtractor.extractFromPicFrame(
          child,
          ctx.themeColors,
          ctx.themeDoc,
        );
        if (frameShape) out.push(frameShape);
        break;
      }

      case "sp": {
        const fromSp = await extractFromSpNode(child, ctx);
        out.push(...fromSp);
        break;
      }

      case "cxnSp": {
        const line = ShapeExtractor.extractFromCxnSp(child, ctx.themeColors, ctx.themeDoc);
        if (line) out.push(line);
        break;
      }

      case "graphicFrame": {
        const table = TableExtractor.extractFromGraphicFrame(
          child,
          ctx.themeColors,
          ctx.themeTableStyles,
        );
        if (table) out.push(table);

        const chart = await ChartExtractor.extractFromGraphicFrame(
          child,
          ctx.rels,
          ctx.zip,
          ctx.themeColors,
          ctx.basePath,
        );
        if (chart) out.push(chart);
        break;
      }

      default:
        break;
    }
  }
}

/** A p:sp may yield an image, shape/line, and/or text — in that paint order. */
async function extractFromSpNode(
  shape: Element,
  ctx: OrderedSpTreeContext,
): Promise<SlideElement[]> {
  const elements: SlideElement[] = [];

  const image = await ImageExtractor.extractFromShape(
    shape,
    ctx.rels,
    ctx.zip,
    ctx.basePath,
    ctx.options,
    ctx.themeContext,
  );
  if (image) {
    elements.push(image);
    return elements;
  }

  const visual = ShapeExtractor.extractFromSp(shape, ctx.themeColors, ctx.themeDoc);
  if (visual) elements.push(visual);

  if (ctx.textOpts) {
    const text = TextExtractor.extractFromSp(shape, ctx.themeColors, ctx.textOpts);
    if (text) elements.push(text);
  }

  return elements;
}

/** Collect shape/image subsets after ordered extraction (e.g. for frame attachment). */
export function partitionSlideElements(elements: SlideElement[]): {
  shapes: ShapeElement[];
  images: ImageElement[];
} {
  return {
    shapes: elements.filter((el): el is ShapeElement => el.type === "shape"),
    images: elements.filter((el): el is ImageElement => el.type === "image"),
  };
}
