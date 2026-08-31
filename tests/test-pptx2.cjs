const pptxgen = require("pptxgenjs");

let pres = new pptxgen();
let s = pres.addSlide();
s.addTable([ ["test"] ]);
pres.writeFile({ fileName: "test3.pptx" }).then(() => {
  const { convert } = require("pptx-to-pdf");
  const fs = require("fs");
  const buffer = fs.readFileSync("test3.pptx");
  convert(buffer).then(pdfBuf => {
     fs.writeFileSync("test3.pdf", pdfBuf);
     const pdftotext = require("child_process").execSync;
     console.log("PDF text:", pdftotext("pdftotext test3.pdf -").toString());
  });
});
