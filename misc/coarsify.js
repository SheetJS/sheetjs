/* based on the `coarse` project README */
const fs = require('fs');
const coarse = require('coarse');
 
const svg = fs.readFileSync(process.argv[2], "utf8");
let roughened = coarse(svg);
const viewbox = roughened.match(/viewBox="(.*?)"/)[1].split(/\s+/);
const v = viewbox.map(x => parseFloat(x));
v[0] -= 40; v[1] += 40; v[2] += 80; v[3] += 80;
roughened = roughened.replace(/<title>G<\/title>/, `$&<polygon fill="white" stroke="" points="${v[0]},${v[1]} ${v[0]},${v[1]-v[3]} ${v[0]+v[2]},${v[1]-v[3]} ${v[0]+v[2]},${v[1]} ${v[0]},${v[1]}"/>`);
 
fs.writeFileSync(process.argv[3], roughened);

