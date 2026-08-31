const fs = require('fs');
async function test() {
  const pptxBuffer = fs.readFileSync('test.pptx');
  const pptxBlob = new Blob([pptxBuffer], { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
  
  const formData = new FormData();
  formData.append("pptx", pptxBlob, "presentation.pptx");
  
  const response = await fetch("http://localhost:3000/api/convert-pdf", {
    method: "POST",
    body: formData,
  });
  
  const pdfBlob = await response.blob();
  const pdfBuffer = Buffer.from(await pdfBlob.arrayBuffer());
  fs.writeFileSync('test-frontend.pdf', pdfBuffer);
  console.log("PDF size:", pdfBuffer.length);
}
test();
