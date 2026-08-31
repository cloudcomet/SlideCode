const fs = require('fs');
const pdf = require('pdf-parse');
pdf(fs.readFileSync('test2.pdf')).then(data => {
  console.log(data.text);
}).catch(console.error);
