import * as pdfjsLib from 'pdfjs-dist/legacy/build/pdf.mjs';

const file = process.argv[2] || 'test-shapes.pdf';
const loadingTask = pdfjsLib.getDocument(file);
loadingTask.promise.then(function(pdf) {
  console.log('PDF loaded, pages:', pdf.numPages);
}, function (reason) {
  console.error('Error:', reason);
});
