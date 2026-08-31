const AsyncFunction = Object.getPrototypeOf(async function(){}).constructor;

const code1 = `
  pptx.layout = "LAYOUT_WIDE";
  console.log("no local decl:", pptx.layout);
`;

const code2 = `
  const pptxgen = require("pptxgenjs");
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_NARROW";
  console.log("with local decl:", pptx.layout);
`;

try {
  global.pptx = { layout: "GLOBAL" };
  
  const executeFn1 = new AsyncFunction('require', 'PptxGenJS', code1);
  executeFn1(() => class { layout = "NEW" }, {}).then(() => {
    const executeFn2 = new AsyncFunction('require', 'PptxGenJS', code2);
    return executeFn2(() => class { layout = "NEW" }, {});
  }).then(() => console.log("done"));
  
} catch(e) {
  console.error("error:", e);
}
