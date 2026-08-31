import PptxGenJS from "pptxgenjs";
import fs from "fs/promises";
import { convert } from "pptx-to-pdf";

async function run() {
  const pptx = new PptxGenJS();
  const buffer = await pptx.write({ outputType: "arraybuffer" });
  const pdfBuffer = await convert(buffer);
  await fs.writeFile("test-empty.pdf", Buffer.from(pdfBuffer));
  console.log("Written test-empty.pdf, size:", pdfBuffer.length);
}
run();
