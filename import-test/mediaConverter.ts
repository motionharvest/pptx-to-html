import { extname } from "node:path";
import { createCanvas, Canvas } from "canvas";
import { convertEmfToDataUrl, convertWmfToDataUrl } from "emf-converter";
import { JSDOM } from "jsdom";
import UTIF from "utif";
import UPNG from "upng-js";

const WEB_FRIENDLY_EXTENSIONS = new Set([".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg"]);

let canvasPolyfillReady = false;

export interface ConvertedMedia {
  filename: string;
  data: Buffer;
  converted: boolean;
}

export interface MediaExtractionResult {
  files: string[];
  pathMap: Map<string, string>;
}

function ensureCanvasPolyfill(): void {
  if (canvasPolyfillReady) {
    return;
  }

  const dom = new JSDOM("<!DOCTYPE html><html><body></body></html>");
  const { document } = dom.window;
  const originalCreateElement = document.createElement.bind(document);

  document.createElement = ((tagName: string, options?: ElementCreationOptions) => {
    if (String(tagName).toLowerCase() === "canvas") {
      return createCanvas(300, 150) as unknown as HTMLCanvasElement;
    }
    return originalCreateElement(tagName, options);
  }) as typeof document.createElement;

  (globalThis as typeof globalThis & { document: Document; HTMLCanvasElement: typeof Canvas }).document =
    document;
  (globalThis as typeof globalThis & { HTMLCanvasElement: typeof Canvas }).HTMLCanvasElement = Canvas;

  canvasPolyfillReady = true;
}

function toArrayBuffer(data: Buffer): ArrayBuffer {
  return data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) as ArrayBuffer;
}

function dataUrlToBuffer(dataUrl: string): Buffer {
  const base64 = dataUrl.split(",")[1];
  if (!base64) {
    throw new Error("Invalid PNG data URL");
  }
  return Buffer.from(base64, "base64");
}

function replaceExtension(filename: string, newExt: string): string {
  const extension = extname(filename);
  return `${filename.slice(0, -extension.length)}${newExt}`;
}

function convertTiffToPng(data: Buffer): Buffer {
  const arrayBuffer = toArrayBuffer(data);
  const ifds = UTIF.decode(arrayBuffer);
  if (!ifds.length) {
    throw new Error("No TIFF image found");
  }

  UTIF.decodeImage(arrayBuffer, ifds[0]);
  const rgba = UTIF.toRGBA8(ifds[0]);
  const png = UPNG.encode([rgba.buffer], ifds[0].width, ifds[0].height, 0);
  return Buffer.from(png);
}

async function convertEmfToPng(data: Buffer): Promise<Buffer> {
  ensureCanvasPolyfill();
  const dataUrl = await convertEmfToDataUrl(toArrayBuffer(data));
  if (!dataUrl) {
    throw new Error("EMF conversion failed");
  }
  return dataUrlToBuffer(dataUrl);
}

async function convertWmfToPng(data: Buffer): Promise<Buffer> {
  ensureCanvasPolyfill();
  const dataUrl = await convertWmfToDataUrl(toArrayBuffer(data));
  if (!dataUrl) {
    throw new Error("WMF conversion failed");
  }
  return dataUrlToBuffer(dataUrl);
}

export async function convertMediaBuffer(filename: string, data: Buffer): Promise<ConvertedMedia> {
  const extension = extname(filename).toLowerCase();

  if (WEB_FRIENDLY_EXTENSIONS.has(extension)) {
    return { filename, data, converted: false };
  }

  try {
    if (extension === ".tif" || extension === ".tiff") {
      return {
        filename: replaceExtension(filename, ".png"),
        data: convertTiffToPng(data),
        converted: true,
      };
    }

    if (extension === ".emf") {
      return {
        filename: replaceExtension(filename, ".png"),
        data: await convertEmfToPng(data),
        converted: true,
      };
    }

    if (extension === ".wmf") {
      return {
        filename: replaceExtension(filename, ".png"),
        data: await convertWmfToPng(data),
        converted: true,
      };
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.warn(`[import-test] Could not convert ${filename}: ${message}`);
    return { filename, data, converted: false };
  }

  console.warn(`[import-test] Unsupported media format, keeping original: ${filename}`);
  return { filename, data, converted: false };
}

export function rewriteMediaPaths(html: string, pathMap: Map<string, string>): string {
  let result = html;

  if (pathMap.size > 0) {
    const entries = [...pathMap.entries()].sort((a, b) => b[0].length - a[0].length);
    for (const [from, to] of entries) {
      result = result.split(from).join(to);
    }
  }

  return result.replace(/ppt\/media\//g, "media/");
}
