/** OOXML ST_TextVerticalType (a:bodyPr @vert). */
export type TextDirection =
  | "vert"
  | "vert270"
  | "eaVert"
  | "mongolianVert"
  | "wordArtVert"
  | "wordArtVertRtl";

export function parseTextDirection(
  bodyPr: Element | null,
  fallback?: string | null,
): TextDirection | undefined {
  const vert = bodyPr?.getAttribute("vert") ?? fallback ?? undefined;
  if (!vert || vert === "horz") return undefined;
  if (
    vert === "vert" ||
    vert === "vert270" ||
    vert === "eaVert" ||
    vert === "mongolianVert" ||
    vert === "wordArtVert" ||
    vert === "wordArtVertRtl"
  ) {
    return vert;
  }
  return undefined;
}

/** CSS for the rotated / vertical text inner wrapper. */
export function textDirectionInnerStyles(direction: TextDirection): string[] {
  switch (direction) {
    case "vert":
      return ["transform:rotate(90deg)", "transform-origin:center center", "white-space:nowrap"];
    case "vert270":
      return ["transform:rotate(-90deg)", "transform-origin:center center", "white-space:nowrap"];
    case "eaVert":
      return ["writing-mode:vertical-rl", "text-orientation:upright"];
    case "mongolianVert":
      return ["writing-mode:vertical-lr"];
    case "wordArtVert":
      return ["writing-mode:vertical-rl", "text-orientation:upright"];
    case "wordArtVertRtl":
      return ["writing-mode:vertical-rl", "text-orientation:upright", "direction:rtl"];
    default:
      return [];
  }
}

export const TEXT_DIRECTION_CENTER_WRAPPER =
  "display:flex;align-items:center;justify-content:center;width:100%;height:100%;min-height:0;align-self:stretch";
