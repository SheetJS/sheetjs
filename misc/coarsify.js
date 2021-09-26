/* based on the `coarse` project README */
const fs = require('fs');
const coarse = require('coarse');
 
const svg = fs.readFileSync(process.argv[2]);
const roughened = coarse(svg);
 
fs.writeFileSync(process.argv[3], roughened);

