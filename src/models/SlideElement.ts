import type { TextDirection } from "../core/textDirection";

export interface Position {
  /** Horizontal position in EMUs (English Metric Units) */
  x: number;

  /** Vertical position in EMUs */
  y: number;
}

export interface Size {
  /** Width in EMUs */
  width: number;

  /** Height in EMUs */
  height: number;
}

export interface TextElement {
  type: "text";
  content: string;
  position: Position;
  size: Size;
  font: { name: string; size: number; color: string };
  align?: { horizontal?: "left" | "center" | "right" | "justify"; vertical?: "top" | "middle" | "bottom" };
  /** a:bodyPr @vert — rotate / stack text (e.g. chart axis labels). */
  textDirection?: TextDirection;
  padding?: { left: number; top: number; right: number; bottom: number };
  html?: string;
  segments?: Array<{ text: string; color?: string; bold?: boolean; italic?: boolean; underline?: boolean; superscript?: boolean; subscript?: boolean; fontSize?: number; fontFamily?: string; breakBefore?: boolean; paragraphBreakBefore?: boolean }>;

  /** Optional line height multiplier (CSS unitless line-height) */
  lineHeight?: number;

  /** Optional absolute line height in points */
  lineHeightPt?: number;
}

export interface CustomGeometry {
  viewBoxW: number;
  viewBoxH: number;
  paths: Array<{ d: string; fillMode?: string }>;
}

export interface ImageCrop {
  /** Crop inset as a fraction of the source image (OOXML srcRect / 100000). */
  left: number;
  top: number;
  right: number;
  bottom: number;
}

export interface DropShadow {
  blurRadPx: number;
  distPx: number;
  dirDeg: number;
  color: string;
  opacity: number;
}

/** Fill and outline from a backing preset shape (e.g. circle under an icon image). */
export interface ShapeFrameStyle {
  shapeType: string;
  fillColor: string;
  borderColor?: string;
  strokeWidth?: number;
  position: Position;
  size: Size;
  shadow?: DropShadow;
  /** When true, clip the image to the frame shape (photo headshot). When false, draw frame behind icon. */
  clipImage?: boolean;
  /** roundRect corner radius as fraction of the shorter side (a:avLst adj / 100000). */
  roundRectAdj?: number;
}

export interface ImageElement {
  type: "image";
  relId: string;
  src: string;
  position: Position;
  size: Size;
  /** Rotation in degrees (from a:xfrm @rot, stored as 60000ths of a degree in PPTX). */
  rotationDeg?: number;
  /** Crop from a:srcRect (source image inset). */
  crop?: ImageCrop;
  /** Destination area from a:stretch/a:fillRect (shape inset). Defaults to full frame. */
  fill?: ImageCrop;
  /** Pixel dimensions of the embedded media file. */
  naturalSize?: { width: number; height: number };
  /** a:picLocks @noChangeAspect — uniform scale (cover) vs stretch-to-fill. */
  preserveAspectRatio?: boolean;
  /** Preset shape fill/outline when the image sits on a framed shape (from spPr or backing oval). */
  frame?: ShapeFrameStyle;
  /** When set, image is clipped to custom geometry paths (a:custGeom). */
  customGeometry?: CustomGeometry;
  /** a:blip/a:alphaModFix opacity multiplier (0–1). */
  imageOpacity?: number;
  flipH?: boolean;
  flipV?: boolean;
}

export interface ShapeElement {
  type: "shape";
  shapeType: string;
  position: Position;
  size: Size;
  fillColor: string;
  borderColor?: string;
  strokeWidth?: number;
  rotationDeg?: number;
  headEnd?: { type?: string; w?: string; len?: string };
  tailEnd?: { type?: string; w?: string; len?: string };
  dashStyle?: "solid" | "dashed" | "dotted" | "sysDash" | "sysDot" | "lgDash" | "dashDot";
  /** Vector paths from a:custGeom (rendered as SVG, not a preset rect). */
  customGeometry?: CustomGeometry;
  /** roundRect corner radius as fraction of the shorter side (a:avLst adj / 100000). */
  roundRectAdj?: number;
  shadow?: DropShadow;
}

export interface LineElement {
  type: "line";
  position: Position;
  size: Size;
  color: string;
  strokeWidth?: number;
  dashStyle?: "solid" | "dashed" | "dotted" | "sysDash" | "sysDot" | "lgDash" | "dashDot";
  rotationDeg?: number;
  flipH?: boolean;
  flipV?: boolean;
  headEnd?: { type?: string; w?: string; len?: string };
  tailEnd?: { type?: string; w?: string; len?: string };
}

export interface BackgroundElement { type: "background"; fillColor?: string; imageSrc?: string; }
export interface TableCell { text: string; font?: { name?: string; size?: number; color?: string }; align?: { horizontal?: "left" | "center" | "right" | "justify"; vertical?: "top" | "middle" | "bottom" }; textDirection?: TextDirection; padding?: { left: number; top: number; right: number; bottom: number }; fillColor?: string; colSpan?: number; rowSpan?: number; borders?: { top?: { color?: string; width?: number; style?: "solid" | "dashed" | "dotted" }; right?: { color?: string; width?: number; style?: "solid" | "dashed" | "dotted" }; bottom?: { color?: string; width?: number; style?: "solid" | "dashed" | "dotted" }; left?: { color?: string; width?: number; style?: "solid" | "dashed" | "dotted" } } }
export interface TableRow {
  cells: TableCell[];
  /** Row height in EMUs (a:tr @h). */
  height?: number;
}
export interface TableElement { type: "table"; position: Position; size: Size; columns: number[]; rows: TableRow[]; tableStyle?: any; tableStyleId?: string; tableFillColor?: string; style?: any }
export type SlideElement = TextElement | ImageElement | ShapeElement | LineElement | BackgroundElement | TableElement | ChartElement;
export type ChartType = "bar" | "column" | "line" | "pie" | "area" | "scatter";
export interface ChartSeries { name?: string; values?: number[]; points?: { x: number; y: number }[]; color?: string; valueFormat?: string }
export interface ChartElement { type: "chart"; chartType: ChartType; position: Position; size: Size; categories: (string | number)[]; series: ChartSeries[]; palette?: string[]; title?: string; showLegend?: boolean; showDataLabels?: boolean; stackedMode?: "none" | "stacked" | "percent"; valueFormat?: string }
