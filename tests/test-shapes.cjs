const { convert } = require('pptx-to-pdf');
const PptxGenJS = require('pptxgenjs');
const fs = require('fs');

async function run() {
const pptx = new PptxGenJS();
const slide = pptx.addSlide();
slide.addShape(pptx.ShapeType.rect, { x: 0.5, y: 0.5, w: 1, h: 1 });
const buffer = await pptx.write({ outputType: 'nodebuffer' });
try {
   const pdf = await convert(buffer);
   fs.writeFileSync('test-shapes.pdf', pdf);
   console.log("PDF generated");
} catch (e) {
   console.error("Convert error:", e);
}
}
run();
