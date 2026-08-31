fetch("http://localhost:3000/api/convert-pdf", { method: 'POST' }).then(res => console.log(res.status)).catch(err => console.error(err));
