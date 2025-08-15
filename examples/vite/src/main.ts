import { pptxToHtml } from "../../../dist/index.js";

const input = document.getElementById("pptxInput") as HTMLInputElement;
const container = document.getElementById("slides")!;

input.addEventListener("change", async () => {
  const file = input.files?.[0];
  if (!file) return;

  const buffer = await file.arrayBuffer();
  const slides = await pptxToHtml(buffer, { width: 400, height: 240, scaleToFit: true });

  container.innerHTML = slides
    .map(
      (html) => `<div class="slide-container">${html}</div>`
    )
    .join("\n");
});