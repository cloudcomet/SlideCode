const { convert } = require('pptx-to-pdf');
const PptxGenJS = require('pptxgenjs');

async function run() {
const pptx = new PptxGenJS();
pptx.layout = "LAYOUT_WIDE";

const slide = pptx.addSlide();

slide.background = { color: "FFFFFF" };

slide.addText("JavaScript → PowerPoint", {
  x: 1,
  y: 1,
  w: 11,
  h: 1,
  fontSize: 32,
  bold: true,
  color: "1E2761",
  align: "center"
});

slide.addText(
  "Basic conversion test with text, styling and positioning.",
  {
    x: 2,
    y: 2.5,
    w: 9,
    h: 1,
    fontSize: 18,
    color: "555555",
    align: "center"
  }
);
  const buffer = await pptx.write({ outputType: 'nodebuffer' });
  try {
     const pdf = await convert(buffer);
     console.log("PDF generated, size:", pdf.byteLength);
  } catch (e) {
     console.error("Convert error:", e);
  }
}
run();
