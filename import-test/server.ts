import { createServer, type IncomingMessage, type ServerResponse } from "node:http";
import { mkdir, readdir, rm, stat, writeFile } from "node:fs/promises";
import { createReadStream, existsSync } from "node:fs";
import { basename, extname, join, normalize, relative, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import Busboy from "busboy";
import JSZip from "jszip";
import { JSDOM } from "jsdom";
import { pptxToHtml } from "../dist/index.js";
import { convertMediaBuffer, rewriteMediaPaths } from "./mediaConverter.ts";

const domParserWindow = new JSDOM("").window;

// The library checks `instanceof Element` while walking XML nodes.
(globalThis as typeof globalThis & { Element: typeof Element }).Element = domParserWindow.Element;

function createDomParser() {
  return new domParserWindow.DOMParser();
}

const __dirname = fileURLToPath(new URL(".", import.meta.url));
const PUBLIC_DIR = join(__dirname, "public");
const OUTPUT_DIR = join(__dirname, "output");
const PORT = Number(process.env.PORT || 3456);

interface ImportResult {
  deckName: string;
  slideCount: number;
  mediaCount: number;
  slides: string[];
  media: string[];
}

function sanitizeDeckName(name: string): string {
  const cleaned = name.replace(/[^a-zA-Z0-9._-]+/g, "-").replace(/^-+|-+$/g, "");
  return cleaned || "imported-deck";
}

function wrapSlideHtml(slideHtml: string, title: string): string {
  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>${title}</title>
  <style>
    body {
      margin: 0;
      min-height: 100vh;
      display: grid;
      place-items: center;
      background: #111827;
    }
  </style>
</head>
<body>
${slideHtml}
</body>
</html>
`;
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function renderResultsPage(result: ImportResult): string {
  const slideItems = result.slides
    .map((slideFile, index) => {
      const label = `Slide ${index + 1}`;
      return `<li><button type="button" class="file-item" data-preview="${escapeHtml(slideFile)}" data-kind="html"><span class="file-label">${label}</span><span class="file-name">${escapeHtml(slideFile)}</span></button></li>`;
    })
    .join("\n");

  const mediaItems = result.media
    .map((mediaFile) => {
      return `<li><button type="button" class="file-item" data-preview="media/${escapeHtml(mediaFile)}" data-kind="image"><span class="file-label">${escapeHtml(mediaFile)}</span><span class="file-name">media/${escapeHtml(mediaFile)}</span></button></li>`;
    })
    .join("\n");

  const mediaSection = result.media.length
    ? `<section class="file-group"><h2>Media</h2><ul class="file-list">${mediaItems}</ul></section>`
    : "";

  const firstPreview = result.slides[0] || (result.media[0] ? `media/${result.media[0]}` : "");
  const firstKind = result.slides[0] ? "html" : result.media[0] ? "image" : "";

  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>${escapeHtml(result.deckName)} import results</title>
  <style>
    :root {
      color-scheme: dark;
      font-family: system-ui, sans-serif;
      line-height: 1.5;
    }

    * { box-sizing: border-box; }

    html,
    body {
      height: 100%;
      margin: 0;
      overflow: hidden;
      background: #0f172a;
      color: #e2e8f0;
    }

    .layout {
      display: grid;
      grid-template-columns: minmax(240px, 320px) minmax(0, 1fr);
      height: 100vh;
      overflow: hidden;
    }

    .sidebar {
      display: flex;
      flex-direction: column;
      gap: 1rem;
      height: 100vh;
      padding: 1.25rem;
      border-right: 1px solid #334155;
      background: #111827;
      overflow-x: hidden;
      overflow-y: auto;
      overscroll-behavior: contain;
    }

    .sidebar-header h1 {
      margin: 0 0 0.35rem;
      font-size: 1.25rem;
      word-break: break-word;
    }

    .meta {
      margin: 0;
      color: #94a3b8;
      font-size: 0.92rem;
    }

    .sidebar-actions {
      margin-top: 0.75rem;
    }

    .sidebar-actions a {
      color: #38bdf8;
      text-decoration: none;
      font-size: 0.92rem;
    }

    .file-group h2 {
      margin: 0 0 0.5rem;
      font-size: 0.8rem;
      letter-spacing: 0.06em;
      text-transform: uppercase;
      color: #64748b;
    }

    .file-list {
      list-style: none;
      margin: 0;
      padding: 0;
      display: grid;
      gap: 0.35rem;
    }

    .file-item {
      width: 100%;
      border: 1px solid transparent;
      border-radius: 0.65rem;
      padding: 0.65rem 0.75rem;
      background: #1e293b;
      color: inherit;
      text-align: left;
      cursor: pointer;
      font: inherit;
    }

    .file-item:hover {
      border-color: #475569;
      background: #243244;
    }

    .file-item.active {
      border-color: #38bdf8;
      background: rgba(56, 189, 248, 0.12);
    }

    .file-label {
      display: block;
      font-weight: 600;
      font-size: 0.95rem;
    }

    .file-name {
      display: block;
      margin-top: 0.15rem;
      color: #94a3b8;
      font-size: 0.78rem;
      word-break: break-all;
    }

    .preview-panel {
      display: flex;
      flex-direction: column;
      height: 100vh;
      min-height: 0;
      overflow: hidden;
      background: #0b1220;
    }

    .preview-toolbar {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 1rem;
      flex: 0 0 auto;
      padding: 0.85rem 1rem;
      border-bottom: 1px solid #334155;
      background: #111827;
    }

    .preview-title {
      margin: 0;
      font-size: 0.95rem;
      color: #cbd5e1;
      word-break: break-all;
    }

    .preview-open {
      color: #38bdf8;
      text-decoration: none;
      font-size: 0.88rem;
      white-space: nowrap;
    }

    .preview-stage {
      position: relative;
      flex: 1 1 auto;
      min-height: 0;
      overflow: hidden;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 1rem;
    }

    .preview-shell {
      display: flex;
      align-items: center;
      justify-content: center;
      width: 100%;
      height: 100%;
      overflow: hidden;
    }

    .preview-frame {
      border: 0;
      border-radius: 0.5rem;
      background: #fff;
      box-shadow: 0 20px 50px rgba(0, 0, 0, 0.35);
      transform-origin: center center;
    }

    .preview-image {
      max-width: 100%;
      max-height: 100%;
      width: auto;
      height: auto;
      object-fit: contain;
      border-radius: 0.5rem;
      box-shadow: 0 20px 50px rgba(0, 0, 0, 0.35);
      background: #fff;
    }

    .preview-empty {
      color: #64748b;
      font-size: 0.95rem;
    }

    @media (max-width: 900px) {
      .layout {
        grid-template-columns: 1fr;
        grid-template-rows: minmax(0, 40vh) minmax(0, 1fr);
      }

      .sidebar {
        height: auto;
        border-right: 0;
        border-bottom: 1px solid #334155;
      }

      .preview-panel {
        height: auto;
      }
    }
  </style>
</head>
<body>
  <div class="layout">
    <aside class="sidebar">
      <div class="sidebar-header">
        <h1>${escapeHtml(result.deckName)}</h1>
        <p class="meta">${result.slideCount} slide(s), ${result.mediaCount} media file(s)</p>
        <div class="sidebar-actions"><a href="/">Import another deck</a></div>
      </div>
      <section class="file-group">
        <h2>Slides</h2>
        <ul class="file-list">${slideItems}</ul>
      </section>
      ${mediaSection}
    </aside>
    <main class="preview-panel">
      <div class="preview-toolbar">
        <p class="preview-title" id="preview-title">${firstPreview ? escapeHtml(firstPreview) : "Select a file to preview"}</p>
        <a class="preview-open" id="preview-open" href="${firstPreview ? `./${escapeHtml(firstPreview)}` : "#"}" target="_blank" rel="noopener"${firstPreview ? "" : " hidden"}>Open in new tab</a>
      </div>
      <div class="preview-stage" id="preview-stage">
        ${firstPreview
          ? firstKind === "image"
            ? `<img class="preview-image" id="preview-image" src="./${escapeHtml(firstPreview)}" alt="${escapeHtml(firstPreview)}" />`
            : `<div class="preview-shell"><iframe class="preview-frame" id="preview-frame" title="Slide preview" scrolling="no" src="./${escapeHtml(firstPreview)}"></iframe></div>`
          : `<p class="preview-empty">Select a slide or media file from the left.</p>`}
      </div>
    </main>
  </div>
  <script>
    const previewTitle = document.getElementById("preview-title");
    const previewOpen = document.getElementById("preview-open");
    const previewStage = document.getElementById("preview-stage");
    const fileItems = Array.from(document.querySelectorAll(".file-item"));
    let activeResizeHandler = null;

    function setActiveItem(button) {
      fileItems.forEach((item) => item.classList.toggle("active", item === button));
    }

    function clearPreviewResizeHandler() {
      if (activeResizeHandler) {
        window.removeEventListener("resize", activeResizeHandler);
        activeResizeHandler = null;
      }
    }

    function fitHtmlPreview(frame) {
      const runFit = () => {
        const padding = 32;
        const availW = Math.max(previewStage.clientWidth - padding, 1);
        const availH = Math.max(previewStage.clientHeight - padding, 1);
        let slideW = 1280;
        let slideH = 720;

        try {
          const doc = frame.contentDocument;
          const slide = doc?.querySelector(".slide, .slide-container");
          if (slide) {
            slideW = slide.offsetWidth || slideW;
            slideH = slide.offsetHeight || slideH;
          } else if (doc?.body) {
            slideW = doc.body.scrollWidth || slideW;
            slideH = doc.body.scrollHeight || slideH;
          }

          if (doc?.documentElement) {
            doc.documentElement.style.overflow = "hidden";
          }
          if (doc?.body) {
            doc.body.style.overflow = "hidden";
            doc.body.style.margin = "0";
          }
        } catch {
          // keep default slide dimensions
        }

        const scale = Math.min(availW / slideW, availH / slideH, 1);
        frame.style.width = slideW + "px";
        frame.style.height = slideH + "px";
        frame.style.transform = "scale(" + scale + ")";
      };

      clearPreviewResizeHandler();
      activeResizeHandler = runFit;
      window.addEventListener("resize", runFit);
      frame.addEventListener("load", runFit);
      runFit();
    }

    function showPreview(url, kind, label) {
      previewTitle.textContent = label || url;
      previewOpen.href = "./" + url;
      previewOpen.hidden = false;
      clearPreviewResizeHandler();
      previewStage.innerHTML = "";

      if (kind === "image") {
        const img = document.createElement("img");
        img.className = "preview-image";
        img.id = "preview-image";
        img.src = "./" + url;
        img.alt = label || url;
        previewStage.appendChild(img);
        return;
      }

      const shell = document.createElement("div");
      shell.className = "preview-shell";
      const frame = document.createElement("iframe");
      frame.className = "preview-frame";
      frame.id = "preview-frame";
      frame.title = "Slide preview";
      frame.scrolling = "no";
      frame.src = "./" + url;
      shell.appendChild(frame);
      previewStage.appendChild(shell);
      fitHtmlPreview(frame);
    }

    fileItems.forEach((button) => {
      button.addEventListener("click", () => {
        const url = button.dataset.preview;
        const kind = button.dataset.kind || "html";
        const label = button.querySelector(".file-label")?.textContent || url;
        setActiveItem(button);
        showPreview(url, kind, label);
        history.replaceState(null, "", "?file=" + encodeURIComponent(url));
      });
    });

    const initial = fileItems[0];
    if (initial) {
      initial.classList.add("active");
    }

    const initialFrame = document.getElementById("preview-frame");
    if (initialFrame) {
      fitHtmlPreview(initialFrame);
    }

    const params = new URLSearchParams(window.location.search);
    const requested = params.get("file");
    if (requested) {
      const match = fileItems.find((button) => button.dataset.preview === requested);
      if (match) {
        match.click();
      }
    }
  </script>
</body>
</html>
`;
}

async function collectMediaPathsFromZipAsync(zip: JSZip): Promise<string[]> {
  const paths = new Set<string>();

  for (const zipPath of Object.keys(zip.files)) {
    if (zipPath.startsWith("ppt/media/") && !zip.files[zipPath].dir) {
      paths.add(zipPath);
    }
  }

  for (const zipPath of Object.keys(zip.files)) {
    if (!zipPath.endsWith(".rels") || zip.files[zipPath].dir) continue;

    const relsXml = await zip.file(zipPath)!.async("string");
    const relsDir = zipPath.replace("/_rels/", "/").replace(".rels", "");
    const baseDir = relsDir.includes("/") ? relsDir.slice(0, relsDir.lastIndexOf("/")) : relsDir;

    const targets = relsXml.match(/Target="([^"]+)"/g) ?? [];
    for (const raw of targets) {
      const target = raw.slice(8, -1);
      const lower = target.toLowerCase();
      if (!/\.(png|jpe?g|gif|bmp|tif{1,2}|emf|wmf|svg|webp)$/i.test(lower)) continue;

      const parts = (baseDir + "/" + target).split("/");
      const resolved: string[] = [];
      for (const part of parts) {
        if (part === "..") resolved.pop();
        else if (part !== "." && part !== "") resolved.push(part);
      }
      const full = resolved.join("/");
      if (zip.files[full] && !zip.files[full].dir) {
        paths.add(full);
      }
    }
  }

  return [...paths];
}

async function extractMedia(
  buffer: ArrayBuffer,
  mediaDir: string
): Promise<{ files: string[]; pathMap: Map<string, string> }> {
  const zip = await JSZip.loadAsync(buffer);
  const mediaPaths = await collectMediaPathsFromZipAsync(zip);

  await mkdir(mediaDir, { recursive: true });

  const files: string[] = [];
  const pathMap = new Map<string, string>();

  for (const zipPath of mediaPaths) {
    const file = zip.file(zipPath);
    if (!file) continue;

    const originalFilename = basename(zipPath);
    const rawData = await file.async("nodebuffer");
    const converted = await convertMediaBuffer(originalFilename, rawData);
    await writeFile(join(mediaDir, converted.filename), converted.data);

    if (!files.includes(converted.filename)) {
      files.push(converted.filename);
    }

    const outputPath = `media/${converted.filename}`;
    pathMap.set(zipPath, outputPath);
    pathMap.set(`ppt/media/${originalFilename}`, outputPath);
    if (converted.converted) {
      pathMap.set(`media/${originalFilename}`, outputPath);
    }
  }

  return { files: files.sort(), pathMap };
}

async function importPptx(buffer: ArrayBuffer, originalName: string): Promise<ImportResult> {
  const deckName = sanitizeDeckName(basename(originalName, extname(originalName)));
  const deckDir = join(OUTPUT_DIR, deckName);
  const mediaDir = join(deckDir, "media");

  await rm(deckDir, { recursive: true, force: true });
  await mkdir(deckDir, { recursive: true });

  const { files: media, pathMap } = await extractMedia(buffer, mediaDir);
  const slidesHtml = await pptxToHtml(buffer, {
    imageSource: "zip-path",
    domParserFactory: createDomParser,
  });

  const slides: string[] = [];
  for (let index = 0; index < slidesHtml.length; index += 1) {
    const slideFile = `slide-${String(index + 1).padStart(3, "0")}.html`;
    const html = wrapSlideHtml(
      rewriteMediaPaths(slidesHtml[index], pathMap),
      `${deckName} - Slide ${index + 1}`
    );
    await writeFile(join(deckDir, slideFile), html, "utf8");
    slides.push(slideFile);
  }

  const result: ImportResult = {
    deckName,
    slideCount: slides.length,
    mediaCount: media.length,
    slides,
    media,
  };

  await writeFile(join(deckDir, "index.html"), renderResultsPage(result), "utf8");
  return result;
}

async function listImportedDecks(): Promise<string[]> {
  try {
    const entries = await readdir(OUTPUT_DIR, { withFileTypes: true });
    const decks: string[] = [];

    for (const entry of entries) {
      if (!entry.isDirectory()) continue;
      const indexPath = join(OUTPUT_DIR, entry.name, "index.html");
      try {
        await stat(indexPath);
        decks.push(entry.name);
      } catch {
        // skip folders without a results page
      }
    }

    return decks.sort((a, b) => a.localeCompare(b));
  } catch {
    return [];
  }
}

async function readUploadedPptx(req: IncomingMessage): Promise<{ buffer: Buffer; filename: string }> {
  return new Promise((resolvePromise, reject) => {
    const busboy = Busboy({ headers: req.headers });
    let uploadBuffer: Buffer | null = null;
    let filename = "upload.pptx";
    let foundFile = false;

    busboy.on("file", (_field, fileStream, info) => {
      foundFile = true;
      filename = info.filename || filename;
      const chunks: Buffer[] = [];

      fileStream.on("data", (chunk) => {
        chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
      });

      fileStream.on("end", () => {
        uploadBuffer = Buffer.concat(chunks);
      });
    });

    busboy.on("finish", () => {
      if (!foundFile || !uploadBuffer) {
        reject(new Error("No .pptx file uploaded"));
        return;
      }
      resolvePromise({ buffer: uploadBuffer, filename });
    });

    busboy.on("error", reject);
    req.pipe(busboy);
  });
}

function sendJson(res: ServerResponse, statusCode: number, payload: unknown): void {
  res.writeHead(statusCode, { "Content-Type": "application/json; charset=utf-8" });
  res.end(JSON.stringify(payload));
}

function resolveOutputPath(urlPath: string): string | null {
  const decoded = decodeURIComponent(urlPath);
  const relativePath = decoded.replace(/^\/output\/?/, "");
  if (!relativePath || relativePath.includes("..")) {
    return null;
  }

  const absolutePath = resolve(OUTPUT_DIR, relativePath);
  const normalizedOutput = normalize(OUTPUT_DIR);
  if (!absolutePath.startsWith(normalizedOutput)) {
    return null;
  }

  return absolutePath;
}

function contentTypeFor(path: string): string {
  const ext = extname(path).toLowerCase();
  switch (ext) {
    case ".html":
      return "text/html; charset=utf-8";
    case ".css":
      return "text/css; charset=utf-8";
    case ".js":
      return "text/javascript; charset=utf-8";
    case ".png":
      return "image/png";
    case ".jpg":
    case ".jpeg":
      return "image/jpeg";
    case ".gif":
      return "image/gif";
    case ".webp":
      return "image/webp";
    case ".svg":
      return "image/svg+xml";
    default:
      return "application/octet-stream";
  }
}

async function serveStaticFile(res: ServerResponse, filePath: string): Promise<void> {
  if (!existsSync(filePath)) {
    res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
    res.end("Not found");
    return;
  }

  res.writeHead(200, { "Content-Type": contentTypeFor(filePath) });
  createReadStream(filePath).pipe(res);
}

const server = createServer(async (req, res) => {
  const method = req.method || "GET";
  const url = new URL(req.url || "/", `http://${req.headers.host || "localhost"}`);

  try {
    if (method === "GET" && url.pathname === "/api/imports") {
      const decks = await listImportedDecks();
      sendJson(res, 200, { decks });
      return;
    }

    if (method === "GET" && url.pathname === "/") {
      await serveStaticFile(res, join(PUBLIC_DIR, "index.html"));
      return;
    }

    if (method === "POST" && url.pathname === "/import") {
      const { buffer, filename } = await readUploadedPptx(req);
      if (!filename.toLowerCase().endsWith(".pptx")) {
        sendJson(res, 400, { error: "Uploaded file must be a .pptx presentation" });
        return;
      }

      const result = await importPptx(buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength), filename);
      sendJson(res, 200, {
        deckName: result.deckName,
        slideCount: result.slideCount,
        mediaCount: result.mediaCount,
        resultsUrl: `/output/${encodeURIComponent(result.deckName)}/`,
      });
      return;
    }

    if (method === "GET" && url.pathname.startsWith("/output/")) {
      const outputPath = resolveOutputPath(url.pathname);
      if (!outputPath) {
        res.writeHead(400, { "Content-Type": "text/plain; charset=utf-8" });
        res.end("Bad request");
        return;
      }

      if (existsSync(outputPath) && !outputPath.endsWith("/")) {
        const stats = await import("node:fs/promises").then((fs) => fs.stat(outputPath));
        if (stats.isDirectory()) {
          await serveStaticFile(res, join(outputPath, "index.html"));
          return;
        }
      }

      await serveStaticFile(res, outputPath);
      return;
    }

    res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
    res.end("Not found");
  } catch (error) {
    const message = error instanceof Error ? error.message : "Unexpected server error";
    if (method === "POST" && url.pathname === "/import") {
      sendJson(res, 500, { error: message });
      return;
    }

    res.writeHead(500, { "Content-Type": "text/plain; charset=utf-8" });
    res.end(message);
  }
});

await mkdir(OUTPUT_DIR, { recursive: true });

server.listen(PORT, () => {
  console.log(`PPTX import test server running at http://localhost:${PORT}`);
  console.log(`Output directory: ${relative(process.cwd(), OUTPUT_DIR) || OUTPUT_DIR}`);
});
