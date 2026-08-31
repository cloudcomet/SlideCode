const PptxGenJS = require('pptxgenjs');
const { convert } = require('pptx-to-pdf');

async function test(code) {
  const pptx = new PptxGenJS();
  
  const mockRequire = (moduleName) => {
    if (moduleName === 'pptxgenjs') return PptxGenJS;
    return require(moduleName);
  };
  
  let targetPptx = pptx;
  
  const ProxyPptxGenJS = new Proxy(PptxGenJS, {
    construct(target, args) {
      const instance = new target(...args);
      targetPptx = instance;
      return instance;
    }
  });

  const executeFn = new Function('pptx', 'require', 'PptxGenJS', code);
  executeFn(pptx, mockRequire, ProxyPptxGenJS);
  
  const buffer = await targetPptx.write({ outputType: 'nodebuffer' });
  try {
     const pdf = await convert(buffer);
     console.log("PDF generated!");
  } catch (e) {
     console.error("Parse error:", e.message);
  }
}

async function runAll() {
  console.log("Test 1: Global pptx");
  await test(`
    const slide = pptx.addSlide();
    slide.addText("Hello");
  `);

  console.log("Test 2: New pptx");
  await test(`
    const p = new PptxGenJS();
    const slide = p.addSlide();
    slide.addText("Hello");
  `);

  console.log("Test 3: Empty pptx");
  await test(`
    // no slide
  `);
}

runAll();

