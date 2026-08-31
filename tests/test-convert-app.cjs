const { convert } = require('pptx-to-pdf');
const PptxGenJS = require('pptxgenjs');
const fs = require('fs');

async function run() {
  const pptx = new PptxGenJS();
  const slide = pptx.addSlide();
  slide.addText("Welcome to Code-to-PPTX", {
    x: 1.5,
    y: 1.5,
    w: 7,
    h: 1,
    fontSize: 32,
    bold: true,
    color: "363636",
    align: "center",
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: 1.5,
    y: 2.75,
    w: 7,
    h: 0.5,
    fill: { color: "0088CC" },
  });

  try {
    const buffer = await pptx.write({ outputType: 'nodebuffer' });
    const pdf = await convert(buffer);
    console.log("PDF generated, size:", pdf.byteLength);
  } catch (e) {
    console.error("Error from convert:", e);
  }
}
run();
