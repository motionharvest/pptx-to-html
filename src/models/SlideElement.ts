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
  padding?: { left: number; top: number; right: number; bottom: number };
  html?: string;
  segments?: Array<{ text: string; color?: string; bold?: boolean; italic?: boolean; underline?: boolean; fontSize?: number; fontFamily?: string; breakBefore?: boolean; paragraphBreakBefore?: boolean }>;

  /** Optional line height multiplier */
  lineHeight?: number;
}

export interface CustomGeometry {
  viewBoxW: number;
  viewBoxH: number;
  paths: Array<{ d: string; fillMode?: string }>;
}

export interface ImageElement {
  type: "image";
  relId: string;
  src: string;
  position: Position;
  size: Size;
  /** When set, image is clipped to custom geometry paths (a:custGeom). */
  customGeometry?: CustomGeometry;
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
  /** Vector paths from a:custGeom (rendered as SVG, not a preset rect). */
  customGeometry?: CustomGeometry;
}

export interface LineElement {
  type: "line";
  position: Position;
  size: Size;
  color: string;
  strokeWidth?: number;
  dashStyle?: "solid" | "dashed" | "dotted";
  rotationDeg?: number;
  flipH?: boolean;
  flipV?: boolean;
  headEnd?: { type?: string; w?: string; len?: string };
  tailEnd?: { type?: string; w?: string; len?: string };
}

export interface BackgroundElement { type: "background"; fillColor?: string; imageSrc?: string; }
export interface TableCell { text: string; font?: { name?: string; size?: number; color?: string }; align?: { horizontal?: "left" | "center" | "right" | "justify"; vertical?: "top" | "middle" | "bottom" }; padding?: { left: number; top: number; right: number; bottom: number }; fillColor?: string; colSpan?: number; rowSpan?: number; borders?: { top?: { color?: string; width?: number; style?: "solid" | "dashed" | "dotted" }; right?: { color?: string; width?: number; style?: "solid" | "dashed" | "dotted" }; bottom?: { color?: string; width?: number; style?: "solid" | "dashed" | "dotted" }; left?: { color?: string; width?: number; style?: "solid" | "dashed" | "dotted" } } }
export interface TableRow { cells: TableCell[] }
export interface TableElement { type: "table"; position: Position; size: Size; columns: number[]; rows: TableRow[]; tableStyle?: any; tableStyleId?: string; tableFillColor?: string; style?: any }
export type SlideElement = TextElement | ImageElement | ShapeElement | LineElement | BackgroundElement | TableElement | ChartElement;
export type ChartType = "bar" | "column" | "line" | "pie" | "area" | "scatter";
export interface ChartSeries { name?: string; values?: number[]; points?: { x: number; y: number }[]; color?: string; valueFormat?: string }
export interface ChartElement { type: "chart"; chartType: ChartType; position: Position; size: Size; categories: (string | number)[]; series: ChartSeries[]; palette?: string[]; title?: string; showLegend?: boolean; showDataLabels?: boolean; stackedMode?: "none" | "stacked" | "percent"; valueFormat?: string }
