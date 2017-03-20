#!/usr/bin/env node
fs.writeFileSync('tests/fixtures.js', Object.keys(paths).concat(['multiformat.lst', './misc/ssf.json', './test_files/biff5/number_format_greek.xls']).map(function(x) { return paths[x] || x; }).map(function(x) { return [x, fs.readFileSync(x).toString('base64')]}).map(function(w) { return "fs['" + w[0] + "'] = '" + w[1] + "';\n"; }).join(""));
