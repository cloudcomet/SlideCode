const pptxgen = require("pptxgenjs");
const { convert } = require('pptx-to-pdf');
const fs = require('fs');

async function run() {
const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";

const NAVY = "1E2761";
const DARK = "141A3D";
const ICE = "CADCFC";
const LIGHT = "EAF0FB";
const WHITE = "FFFFFF";
const GOLD = "C9962E";
const MUTED = "5B6B8C";

function footer(slide, number) {
  slide.addText(
    "TECHNICAL SEMINAR | RESEARCH PRESENTATION",
    { x: 0.6, y: 7.1, w: 7, h: 0.2, fontSize: 8, color: MUTED, charSpacing: 1 }
  );
  slide.addText(String(number), { x: 12.2, y: 7.1, w: 0.5, h: 0.2, fontSize: 8, align: "right", color: MUTED });
}

function sectionTitle(slide, eyebrow, heading) {
  slide.addText(eyebrow.toUpperCase(), { x: 0.6, y: 0.4, w: 10, h: 0.3, fontSize: 11, bold: true, color: GOLD, charSpacing: 2 });
  slide.addText(heading, { x: 0.6, y: 0.8, w: 12, h: 0.7, fontSize: 29, bold: true, color: NAVY, fontFace: "Cambria" });
}

{
  const slide = pptx.addSlide();
  slide.background = { color: NAVY };
  slide.addText("ARTIFICIAL INTELLIGENCE & DATA SCIENCE", { x: 0.7, y: 1.6, w: 8, h: 0.4, fontSize: 12, bold: true, color: GOLD, charSpacing: 2 });
}

const buffer = await pptx.write({ outputType: 'nodebuffer' });
try {
   const pdf = await convert(buffer);
   fs.writeFileSync('test18.pdf', pdf);
   console.log("PDF generated");
} catch (e) {
   console.error("Convert error:", e);
}
}
run();
