import { convert } from "pptx-to-pdf";
import fs from "fs/promises";

async function run() {
  try {
    const buffer = await fs.readFile("test2.pptx");
    const pdfBuffer = await convert(buffer, {
      onWarning: (msg) => console.log("WARN:", msg)
    });
    await fs.writeFile("test2-local.pdf", pdfBuffer);
    console.log("Converted successfully");
  } catch(e) {
    console.error("ERROR:", e);
  }
}
run();
