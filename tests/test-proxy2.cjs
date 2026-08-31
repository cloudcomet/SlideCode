const PptxGenJS = require('pptxgenjs');

async function test(code) {
  const pptx = new PptxGenJS();
  const executeFn = new Function('pptx', code);
  try {
     executeFn(pptx);
     console.log("Success");
  } catch (e) {
     console.error("Exec error:", e.message);
  }
}

test(`
  const pptx = new PptxGenJS();
`);
