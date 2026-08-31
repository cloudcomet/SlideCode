import fs from 'fs';

async function run() {
  const pptxBuffer = fs.readFileSync('test.pptx');
  const blob = new Blob([pptxBuffer], { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
  
  const formData = new FormData();
  formData.append("pptx", blob, "presentation.pptx");
  
  const response = await fetch("http://localhost:3000/api/convert-pdf", {
    method: "POST",
    body: formData
  });
  
  console.log(response.status);
  const pdfBuffer = await response.arrayBuffer();
  console.log("PDF buffer size:", pdfBuffer.byteLength);
  
  fs.writeFileSync('test-output.pdf', Buffer.from(pdfBuffer));
}
run();
