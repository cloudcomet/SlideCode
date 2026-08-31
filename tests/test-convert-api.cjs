const { convert } = require("pptx-to-pdf");
const fs = require("fs");
const pptx = fs.readFileSync("full-test.pptx");
convert(pptx).then(pdf => {
  console.log("Converted!", pdf.length);
}).catch(err => {
  console.error("Conversion error:", err);
});
