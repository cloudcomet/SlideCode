const PptxGenJS = require('pptxgenjs');
const fs = require('fs');

async function run() {
const pptx = new PptxGenJS();
pptx.layout = "LAYOUT_WIDE";

const slide = pptx.addSlide();
slide.addText("Hello");
const buffer = await pptx.write({ outputType: 'nodebuffer' });
fs.writeFileSync('test.pptx', buffer);
}
run();
