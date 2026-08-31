import fs from "fs/promises";
import { convert } from "pptx-to-pdf";
async function run() {
  const buf = await fs.readFile("full-test.pptx");
  const pdf = await convert(buf);
  await fs.writeFile("output.pdf", pdf);
}
run();
