/** Detect http(s) and www. URLs in plain text (for underline when not marked in OOXML). */
export const URL_REGEX = /\bhttps?:\/\/[^\s<]+|\bwww\.[^\s<]+/gi;

export function findUrlRanges(text: string): Array<{ start: number; end: number }> {
  const ranges: Array<{ start: number; end: number }> = [];
  if (!text) return ranges;
  const re = new RegExp(URL_REGEX.source, URL_REGEX.flags);
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) {
    ranges.push({ start: m.index, end: m.index + m[0].length });
  }
  return ranges;
}

export function splitTextByUrls(text: string): Array<{ text: string; isUrl: boolean }> {
  const ranges = findUrlRanges(text);
  if (!ranges.length) return [{ text, isUrl: false }];

  const parts: Array<{ text: string; isUrl: boolean }> = [];
  let pos = 0;
  for (const { start, end } of ranges) {
    if (start > pos) parts.push({ text: text.slice(pos, start), isUrl: false });
    parts.push({ text: text.slice(start, end), isUrl: true });
    pos = end;
  }
  if (pos < text.length) parts.push({ text: text.slice(pos), isUrl: false });
  return parts;
}
