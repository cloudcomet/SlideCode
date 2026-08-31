const pptxgen = require("pptxgenjs");
const { convert } = require('pptx-to-pdf');
const fs = require('fs');

async function run() {
const pptx = new pptxgen();
const slide = pptx.addSlide();
slide.addText("नमस्ते दुनिया", { x: 1, y: 1 });
const buffer = await pptx.write({ outputType: 'nodebuffer' });
const pdf = await convert(buffer);
fs.writeFileSync('test09.pdf', pdf);
}
run();
