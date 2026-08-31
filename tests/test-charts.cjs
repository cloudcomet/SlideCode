const { convert } = require('pptx-to-pdf');
const PptxGenJS = require('pptxgenjs');
const fs = require('fs');

async function run() {
const pptx = new PptxGenJS();
const slide = pptx.addSlide();
slide.addChart(pptx.ChartType.line, [{ name: "Revenue", labels: ["Jan"], values: [120] }], { x: 0.7, y: 1.3, w: 1, h: 1 });
const buffer = await pptx.write({ outputType: 'nodebuffer' });
try {
   const pdf = await convert(buffer);
   fs.writeFileSync('test-charts.pdf', pdf);
   console.log("PDF generated");
} catch (e) {
   console.error("Convert error:", e);
}
}
run();
