const { convert } = require('pptx-to-pdf');
const PptxGenJS = require('pptxgenjs');
const fs = require('fs');

async function run() {
  const pptx = new PptxGenJS();
  try {
    const buffer = await pptx.write({ outputType: 'nodebuffer' });
    const pdf = await convert(buffer);
    console.log("PDF generated, size:", pdf.byteLength);
  } catch (e) {
    console.error("Error from convert:", e);
  }
}
run();
