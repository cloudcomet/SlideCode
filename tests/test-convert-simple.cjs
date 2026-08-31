const { convert } = require('pptx-to-pdf');
const PptxGenJS = require('pptxgenjs');

async function run() {
  const pptx = new PptxGenJS();
  const slide = pptx.addSlide();
  slide.addText("Hello");
  const buffer = await pptx.write({ outputType: 'nodebuffer' });
  try {
     const pdf = await convert(buffer);
     console.log("PDF generated, size:", pdf.byteLength);
  } catch (e) {
     console.error("Convert error:", e);
  }
}
run();
