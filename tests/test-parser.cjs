const { parsePresentation } = require('pptx-to-pdf/dist/parser/presentation-parser.js');
const { PPTXArchive } = require('pptx-to-pdf/dist/parser/pptx-archive.js');
const fs = require('fs/promises');

async function run() {
  const buf = await fs.readFile('test2.pptx');
  const archive = new PPTXArchive();
  await archive.load(buf);
  const pres = await parsePresentation(archive);
  const slide2 = pres.slides[1]; // Slide 5
  const table = slide2.elements.find(e => e.type === 'table');
  console.log("Table rows:", table.rows.length);
  console.log("Row 0, Cell 0 text paragraphs:", table.rows[0].cells[0].text?.paragraphs?.length);
  console.log("Row 1, Cell 1 text paragraphs:", table.rows[1].cells[1].text?.paragraphs?.length);
  
  const slide4 = pres.slides[3]; // Slide 11
  const table4 = slide4.elements.find(e => e.type === 'table');
  console.log("Slide 11 Table rows:", table4.rows.length);
  console.log("Slide 11 Row 1, Cell 1 text paragraphs:", table4.rows[1].cells[1].text?.paragraphs?.length);
  console.log("Slide 11 Row 1, Cell 1 runs:", table4.rows[1].cells[1].text?.paragraphs[0]?.runs);
}
run().catch(console.error);
