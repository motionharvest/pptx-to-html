/**
 * Regression check against import-test/reference/ American Express deck.
 * Run: npm run build --prefix .. && node test-reference-deck.mjs
 */
import { readFileSync, writeFileSync, mkdirSync } from "fs";
import { join, dirname } from "path";
import { fileURLToPath } from "url";
import { execSync } from "child_process";
import { JSDOM } from "jsdom";
import { pptxToHtml } from "../dist/index.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const REF_DIR = join(__dirname, "reference", "American Express Early Pay_06-04 Deck - Copy");
const PPTX = join(__dirname, "tmp-reference-test.pptx");

execSync(`cd "${REF_DIR}" && zip -qr "${PPTX}" .`, { stdio: "inherit" });

const domParserWindow = new JSDOM("").window;
globalThis.Element = domParserWindow.Element;

const buf = readFileSync(PPTX);
const html = (await pptxToHtml(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength), {
  domParserFactory: () => new domParserWindow.DOMParser(),
}))[0];

const checks = {
  hasBlueSvg: html.includes('fill="#006FCF"'),
  noBlueRect: !html.includes("background-color: #006FCF"),
  hasGraySvg: html.includes('fill="#D9D9D6"'),
  hasSvg: html.includes("<svg"),
  fullWidthWave: /left: 0px;[\s\S]*?width: 1280px/.test(html),
};

console.log("Reference deck checks:", checks);

mkdirSync(join(__dirname, "tmp"), { recursive: true });
writeFileSync(join(__dirname, "tmp", "reference-slide1.html"), html);

const failed = Object.entries(checks).filter(([, v]) => !v);
if (failed.length) {
  console.error("Failed:", failed.map(([k]) => k).join(", "));
  process.exit(1);
}

console.log("OK — wrote tmp/reference-slide1.html for inspection");
