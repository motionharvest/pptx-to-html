# pptx-to-html

Convert PowerPoint (.pptx) files into HTML slides using TypeScript. This library parses the OOXML structures inside a PPTX and renders each slide as a self-contained HTML snippet with absolutely positioned elements (text, images, shapes, tables, and basic charts).

- Zero-runtime CSS required — styles are inlined.
- Theme colors, basic table and chart styling supported.
- Works in modern browsers. Node is supported via an injectable DOM parser.

> Based on the ISO/IEC 29500:2012 "Office Open XML File Formats — Fundamentals And Markup Language Reference".

## Installation

```bash
npm install pptx-to-html
```

## Quick Start (Browser)

```ts
import { pptxToHtml } from "pptx-to-html";

async function handleFile(file: File) {
  const arrayBuffer = await file.arrayBuffer();
  const slidesHtml = await pptxToHtml(arrayBuffer, {
    width: 960,
    height: 540,
    scaleToFit: true,      // scale content to the container size
    letterbox: true        // preserve aspect ratio with black bars
  });

  // Render slides into the page
  const container = document.getElementById("slides")!;
  container.innerHTML = slidesHtml.join("\n");
}

document.getElementById("file")!.addEventListener("change", (e) => {
  const input = e.target as HTMLInputElement;
  if (input.files && input.files[0]) handleFile(input.files[0]);
});
```

## Quick Start (Node)

In Node, there is no built-in `DOMParser`. Provide one via `domParserFactory` or install a DOM parser in your app (e.g. `@xmldom/xmldom`).

```bash
npm install @xmldom/xmldom
```

```ts
import { readFile } from "node:fs/promises";
import { pptxToHtml } from "pptx-to-html";
import { DOMParser } from "@xmldom/xmldom";

async function main() {
  const buf = await readFile("./example.pptx");
  const slidesHtml = await pptxToHtml(buf.buffer, {
    width: 960,
    height: 540,
    scaleToFit: true,
    domParserFactory: () => new DOMParser(),
  });
  console.log(slidesHtml.join("\n\n"));
}

main().catch(console.error);
```

## API

### `pptxToHtml(buffer, config?) => Promise<string[]>`
- `buffer`: `ArrayBuffer` of the `.pptx` file contents.
- Returns: an array of HTML strings, one per slide.

### Options

| Option              | Type        | Default           | Description |
|---------------------|-------------|-------------------|-------------|
| `width`             | `number`    | PPT base width or `960` | Target container width (px). If `scaleToFit` is `true`, this is the outer container width used to scale content. |
| `height`            | `number`    | PPT base height or `540` | Target container height (px). If `scaleToFit` is `true`, this is the outer container height used to scale content. |
| `scaleToFit`        | `boolean`   | `false`           | When `true`, scales the original slide coordinate system (EMU to px) to fit the target width/height. |
| `letterbox`         | `boolean`   | `true` if `scaleToFit` else `false` | When scaling, preserve aspect ratio with black bars around content. If `false`, content stretches to fill. |
| `domParserFactory`  | `() => { parseFromString(xml: string, mime: string): Document }` | `undefined` | Node-only. Provide a DOM parser factory if no global `DOMParser` exists. If omitted, the library tries to `require('@xmldom/xmldom')` at runtime if present in the host app. |

Notes:
- The library reads the slide base size from the PPT (`ppt/presentation.xml` sldSz). When unavailable, it defaults to `960x540`.
- All element coordinates and sizes are normalized from EMUs to pixels (EMU/9525).

## What It Renders

- Text boxes (paragraphs, basic list bullets and numbering).
- Images.
- Shapes and connectors (common presets; arrows use SVG markers).
- Tables (header/first column emphasis, banded rows/columns, table-level and cell-level fills, borders).
- Charts (column, bar, line, area, pie, scatter) with basic axes, labels and theme palette fallback.

## Limitations

- Fidelity: Not a pixel-perfect engine. Complex layouts or advanced effects may differ from PowerPoint.
- Unsupported features: animations, transitions, SmartArt, embedded audio/video, 3D, complex text effects, advanced chart options.
- Fonts: No font embedding; rendering uses the client’s available fonts (default fallbacks provided).
- Theme resolution: Color resolution follows common OOXML patterns; some vendor-specific or advanced theme constructs may not be fully recognized.
- Large files: Very large PPTX files may require substantial memory for ZIP extraction and base64 images.

## Environment Support

- Browsers: modern evergreen browsers (Chromium, Firefox, Safari) with ES modules support.
- Node: supported. Provide a DOM parser via `domParserFactory`, or install one in your app (e.g. `@xmldom/xmldom`). The library will also attempt a best-effort `require('@xmldom/xmldom')` if no parser is provided and it’s installed.

## Rationale and References

This library parses OOXML files according to the Office Open XML standards and common PowerPoint authoring patterns.

- Based on ISO/IEC 29500:2012 — "Office Open XML File Formats — Fundamentals And Markup Language Reference".

## License

MIT

