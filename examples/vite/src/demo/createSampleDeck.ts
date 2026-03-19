import JSZip from "jszip";

const PRESENTATION_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types";

function textShape(id: number, name: string, x: number, y: number, cx: number, cy: number, paragraphs: string[], fontSize = 1800, color = "123052") {
  const runs = paragraphs
    .map((paragraph) => {
      const isBullet = paragraph.startsWith("• ");
      const text = escapeXml(isBullet ? paragraph.slice(2) : paragraph);

      return `
        <a:p>
          ${isBullet ? '<a:pPr lvl="1"><a:buChar char="•"/></a:pPr>' : ""}
          <a:r>
            <a:rPr lang="en-US" sz="${fontSize}" dirty="0">
              <a:solidFill><a:srgbClr val="${color}"/></a:solidFill>
              <a:latin typeface="Aptos"/>
            </a:rPr>
            <a:t>${text}</a:t>
          </a:r>
          <a:endParaRPr lang="en-US" sz="${fontSize}" dirty="0"/>
        </a:p>`;
    })
    .join("");

  return `
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="${id}" name="${escapeXml(name)}"/>
        <p:cNvSpPr txBox="1"/>
        <p:nvPr/>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <a:off x="${x}" y="${y}"/>
          <a:ext cx="${cx}" cy="${cy}"/>
        </a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
        <a:noFill/>
        <a:ln><a:noFill/></a:ln>
      </p:spPr>
      <p:txBody>
        <a:bodyPr wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720"/>
        <a:lstStyle/>
        ${runs}
      </p:txBody>
    </p:sp>`;
}

function shape(id: number, name: string, x: number, y: number, cx: number, cy: number, fill: string, stroke: string, preset = "roundRect") {
  return `
    <p:sp>
      <p:nvSpPr>
        <p:cNvPr id="${id}" name="${escapeXml(name)}"/>
        <p:cNvSpPr/>
        <p:nvPr/>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm>
          <a:off x="${x}" y="${y}"/>
          <a:ext cx="${cx}" cy="${cy}"/>
        </a:xfrm>
        <a:prstGeom prst="${preset}"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>
        <a:ln w="19050"><a:solidFill><a:srgbClr val="${stroke}"/></a:solidFill></a:ln>
      </p:spPr>
    </p:sp>`;
}

function slideXml(content: string) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="${DRAWING_NS}" xmlns:r="${REL_NS}" xmlns:p="${PRESENTATION_NS}">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
      ${content}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`;
}

function escapeXml(value: string) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

export async function createSampleDeck(): Promise<ArrayBuffer> {
  const zip = new JSZip();

  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="${CONTENT_TYPES_NS}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`);

  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PACKAGE_REL_NS}">
  <Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`);

  zip.file("ppt/presentation.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="${DRAWING_NS}" xmlns:r="${REL_NS}" xmlns:p="${PRESENTATION_NS}">
  <p:sldMasterIdLst/>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
    <p:sldId id="257" r:id="rId2"/>
  </p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`);

  zip.file("ppt/_rels/presentation.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PACKAGE_REL_NS}">
  <Relationship Id="rId1" Type="${REL_NS}/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId2" Type="${REL_NS}/slide" Target="slides/slide2.xml"/>
</Relationships>`);

  zip.file("ppt/slides/_rels/slide1.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PACKAGE_REL_NS}"/>`);
  zip.file("ppt/slides/_rels/slide2.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PACKAGE_REL_NS}"/>`);

  zip.file("ppt/slides/slide1.xml", slideXml(`
    ${shape(2, "Summary card", 457200, 457200, 10972800, 1968500, "E9F2FF", "2F6FEB")}
    ${shape(3, "Accent bar", 457200, 2641600, 4572000, 2743200, "FFF4D6", "E0B140")}
    ${shape(4, "Status chip", 5486400, 2641600, 5486400, 731520, "F4F9F7", "3A8A61")}
    ${textShape(5, "Slide title", 731520, 731520, 10058400, 731520, ["PowerPoint slides rendered as HTML"], 2400, "102040")}
    ${textShape(6, "Slide summary", 731520, 1371600, 10058400, 640080, ["Browser-first parsing for PPTX files with inline HTML output."], 1600, "355070")}
    ${textShape(7, "Feature list", 777240, 2870200, 3937000, 2194560, ["Sample capabilities", "• Text extraction", "• Shape rendering", "• Scaled slide output"], 1500, "4E4020")}
    ${textShape(8, "Metric heading", 5753100, 2857500, 5029200, 548640, ["Generated sample deck"], 1500, "1F5136")}
    ${textShape(9, "Metric copy", 5753100, 3467100, 5029200, 1645920, ["Use the sample deck button to test the library without uploading a file."], 1500, "25543C")}
  `));

  zip.file("ppt/slides/slide2.xml", slideXml(`
    ${shape(10, "Roadmap panel", 640080, 640080, 10911840, 5486400, "FAFBFD", "D4DCE8", "rect")}
    ${shape(11, "Milestone one", 1097280, 1645920, 2377440, 1097280, "E8F7EE", "4D9A6B")}
    ${shape(12, "Milestone two", 4206240, 1645920, 2377440, 1097280, "FFF4D6", "C49B2C")}
    ${shape(13, "Milestone three", 7315200, 1645920, 2377440, 1097280, "FBE9EF", "C4587A")}
    ${textShape(14, "Roadmap title", 1097280, 960120, 10058400, 548640, ["Implementation map"], 2200, "12263A")}
    ${textShape(15, "Milestone one text", 1277640, 1913880, 2011680, 548640, ["1. Parse PPTX"], 1600, "24513A")}
    ${textShape(16, "Milestone two text", 4386600, 1913880, 2011680, 548640, ["2. Render slides"], 1600, "6B4B00")}
    ${textShape(17, "Milestone three text", 7495560, 1913880, 2011680, 548640, ["3. Publish docs"], 1600, "6C2741")}
    ${textShape(18, "Roadmap bullets", 1277640, 3291840, 9144000, 1645920, ["The same API used in production is exercised in Storybook.", "• Generated sample deck for zero-setup testing", "• File upload path for real presentations", "• Static docs ready for GitHub Pages"], 1500, "33485C")}
  `));

  return zip.generateAsync({ type: "arraybuffer" });
}
