/** PPTX keeps superscript/subscript at the same sz; Office renders glyphs smaller (~65%). */
export const SCRIPT_FONT_SCALE = 0.65;

export interface TextRunStyleInput {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  superscript?: boolean;
  subscript?: boolean;
  color?: string;
  fontSize?: number;
  fontFamily?: string;
}

export function textRunInlineStyles(
  run: TextRunStyleInput,
  defaultFont: string,
  defaultSize: number,
  defaultColor: string | undefined,
): string[] {
  const styles: string[] = [];
  if (run.bold) styles.push("font-weight:bold");
  if (run.italic) styles.push("font-style:italic");
  if (run.underline) styles.push("text-decoration:underline");
  if (run.superscript) styles.push("vertical-align:super");
  else if (run.subscript) styles.push("vertical-align:sub");
  if (run.color && run.color !== defaultColor) styles.push(`color:${run.color}`);

  const baseSize = run.fontSize ?? defaultSize;
  if (run.superscript || run.subscript) {
    styles.push(`font-size:${formatPt(baseSize * SCRIPT_FONT_SCALE)}pt`);
  } else if (run.fontSize && run.fontSize !== defaultSize) {
    styles.push(`font-size:${formatPt(run.fontSize)}pt`);
  }

  if (run.fontFamily && run.fontFamily !== defaultFont) styles.push(`font-family:${run.fontFamily}`);
  return styles;
}

function formatPt(size: number): string {
  const rounded = Math.round(size * 1000) / 1000;
  return Number.isInteger(rounded) ? String(rounded) : String(rounded);
}
