const functions = require('firebase-functions');
const Busboy = require('busboy');
const XLSX = require('xlsx');

// // Create and Deploy Your First Cloud Functions
// // https://firebase.google.com/docs/functions/write-firebase-functions
//
exports.helloWorld = functions.https.onRequest((request, response) => {
 response.send("Hello from Firebase!");
});

exports.main = functions.https.onRequest((req, res) => {
  var bb = new Busboy({
    headers: {
      'content-type': req.headers['content-type']
    }
  });
  let fields = {};
  let files = {};
  bb.on('field', (fieldname, val) => {
    fields[fieldname] = val;
  });
  bb.on('file', (fieldname, file, filename) => {
    var buffers = [];
    file.on('data', (data) => {
      buffers.push(data);
    });
    file.on('end', () => {
      files[fieldname] = [Buffer.concat(buffers), filename];
    });
  });
  bb.on('finish', () => {
    let f = files[Object.keys(files)[0]];
    const wb = XLSX.read(f[0], { type: "buffer" });
    // Convert to CSV
    res.send(XLSX.utils.sheet_to_csv(wb.Sheets[wb.SheetNames[0]]));
  });
  bb.end(req.body)
});
