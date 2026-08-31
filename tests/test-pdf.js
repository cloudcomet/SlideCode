import PptxGenJS from "pptxgenjs";
import { convert } from "pptx-to-pdf";
import fs from "fs/promises";

async function run() {
  const pptx = new PptxGenJS();
  pptx.addSlide().addText("Slide 1");
  pptx.addSlide().addText("Slide 2");
  pptx.addSlide().addText("Slide 3");
  
  const buffer = await pptx.write({ outputType: "arraybuffer" });
  const pdfBuffer = await convert(buffer);
  
  await fs.writeFile("test.pdf", pdfBuffer);
  console.log("Written test.pdf");
}
run();
