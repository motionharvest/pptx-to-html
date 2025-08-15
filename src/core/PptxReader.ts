import JSZip from "jszip";
import { SlideExtractor } from "./SlideExtractor";
import { SlideElement } from "../models/SlideElement";
import { XmlHelper } from "./XmlHelper";

/**
 * Handles loading a .pptx file and extracting slide elements from it.
 */
export class PptxReader {
  private zip!: JSZip;
  private baseWidthPx?: number;
  private baseHeightPx?: number;

  /**
   * Loads and parses a .pptx binary buffer.
   * @param buffer The binary content of a .pptx file.
   * @returns A list of slides, each represented as an array of SlideElement.
   */
  async load(buffer: ArrayBuffer): Promise<SlideElement[][]> {
    this.zip = await JSZip.loadAsync(buffer);
    await this.computeSlideBaseSize();
    const extractor = new SlideExtractor(this.zip);
    return extractor.extractSlides();
  }

  /**
   * Returns slide base size in pixels derived from ppt/presentation.xml sldSz (if available).
   * Defaults to 960x540 when not found.
   */
  async getBaseSizePx(): Promise<{ width: number; height: number }> {
    if (this.baseWidthPx && this.baseHeightPx) {
      return { width: this.baseWidthPx, height: this.baseHeightPx };
    }
    await this.computeSlideBaseSize();
    return {
      width: Number.isFinite(this.baseWidthPx as any) && (this as any).baseWidthPx > 0 ? (this.baseWidthPx as number) : 960,
      height: Number.isFinite(this.baseHeightPx as any) && (this as any).baseHeightPx > 0 ? (this.baseHeightPx as number) : 540,
    };
  }

  private async computeSlideBaseSize(): Promise<void> {
    try {
      if (!this.zip) return; // allow calling before load(), will use defaults
      const presFile = this.zip.file("ppt/presentation.xml");
      if (!presFile) return;
      const xmlStr = await presFile.async("string");
      const doc = XmlHelper.parseXml(xmlStr);
      const sldSz = doc.getElementsByTagNameNS("*", "sldSz")[0];
      const cx = sldSz ? Number(sldSz.getAttribute("cx") || 0) : 0;
      const cy = sldSz ? Number(sldSz.getAttribute("cy") || 0) : 0;
      if (Number.isFinite(cx) && Number.isFinite(cy) && cx > 0 && cy > 0) {
        this.baseWidthPx = cx / 9525;
        this.baseHeightPx = cy / 9525;
      }
    } catch {
      // ignore; fall back to defaults
    }
  }
}
