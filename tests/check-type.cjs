const { convert } = require('pptx-to-pdf');
const PptxGenJS = require('pptxgenjs');

async function run() {
const pptx = new PptxGenJS();
pptx.addSlide().addText("Hello");
const buffer = await pptx.write({ outputType: 'nodebuffer' });
const pdf = await convert(buffer);
console.log(pdf.constructor.name);
console.log(Buffer.isBuffer(pdf));
}
run();
