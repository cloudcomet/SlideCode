import fs from 'fs';

async function run() {
  const pptxBuffer = Buffer.from('hello'); // not a pptx
  const blob = new Blob([pptxBuffer], { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
  
  const formData = new FormData();
  formData.append("pptx", blob, "presentation.pptx");
  
  const response = await fetch("http://localhost:3000/api/convert-pdf", {
    method: "POST",
    body: formData
  });
  
  console.log("Status:", response.status);
  const text = await response.text();
  console.log("Response:", text);
}
run();
