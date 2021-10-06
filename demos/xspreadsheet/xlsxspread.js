/**
 * Converts data from SheetJS to x-spreadsheet
 * 
 * @param  {Object} wb SheetJS workbook object
 * @param  {Boolean} [keepMerges=false] Does the conversion keep merged cells? (default: false)
 * @param  {Boolean} [keepFormulas=false] Does the conversion keep formulas (=true) or result value (=false)? (default: false)
 * 
 * @returns {Object[]} An x-spreadsheet data
 */
function stox(wb, keepMerges, keepFormulas) {
    keepMerges = keepMerges === undefined ? false : keepMerges;
    keepFormulas = keepFormulas === undefined ? false : keepFormulas;

    var out = [];
    wb.SheetNames.forEach(function (name) {
        var o = { name: name, rows: {} };
        var ws = wb.Sheets[name];
        var range = XLSX.utils.decode_range(ws['!ref']);
        // sheet_to_json will lost empty row and col at begin as default
        range.s = { r: 0, c: 0 }
        var aoa = XLSX.utils.sheet_to_json(ws, {
            raw: false,
            header: 1,
            range: range,
        });

        aoa.forEach(function (r, i) {
            var cells = {};
            r.forEach(function (c, j) {
                cells[j] = { text: c };

                if (keepFormulas) {
                    var cellRef = XLSX.utils.encode_cell({ r: i, c: j })

                    if (
                        ws[cellRef] != undefined &&
                        ws[cellRef].f != undefined
                    ) {
                        cells[j].text = "=" + ws[cellRef].f;
                    }
                }
            });
            o.rows[i] = { cells: cells };
        });

        if (keepMerges) {
            o.merges = [];
            ws["!merges"].forEach(function (merge, i) {
                //Needed to support merged cells with empty content
                if (o.rows[merge.s.r] == undefined) {
                    o.rows[merge.s.r] = { cells: {} };
                }
                if (o.rows[merge.s.r].cells[merge.s.c] == undefined) {
                    o.rows[merge.s.r].cells[merge.s.c] = {};
                }

                o.rows[merge.s.r].cells[merge.s.c].merge = [
                    merge.e.r - merge.s.r,
                    merge.e.c - merge.s.c,
                ];

                o.merges[i] =
                    XLSX.utils.encode_cell(merge.s) +
                    ":" +
                    XLSX.utils.encode_cell(merge.e);
            });
        }

        out.push(o);
    });

    return out;
}

/**
 * Converts data from x-spreadsheet to SheetJS
 *
 * @param  {Object[]} sdata An x-spreadsheet data object
 * @param  {Boolean} [keepMerges=false] Does the conversion keep merged cells? (default: false)
 * @param  {Boolean} [keepFormulas=false] Does the conversion keep formulas (=true) or result value (=false)? (default: false)
 *
 * @returns {Object} A SheetJS workbook object
 */
function xtos(sdata, keepMerges, keepFormulas) {
    keepMerges = keepMerges === undefined ? false : keepMerges;
    keepFormulas = keepFormulas === undefined ? false : keepFormulas;

    let out = XLSX.utils.book_new();
    sdata.forEach(function (xws) {
        var ws = {};
        var rowobj = xws.rows;
        for (var ri = 0; ri < rowobj.len; ++ri) {
            var row = rowobj[ri];
            if (!row) continue;

            var minCoord, maxCoord;
            Object.keys(row.cells).forEach(function (k) {
                var idx = +k;
                if (isNaN(idx)) return;

                var lastRef = XLSX.utils.encode_cell({ r: ri, c: idx });
                if (minCoord == undefined) {
                    minCoord = {
                        r: ri,
                        c: idx,
                    };
                } else {
                    if (ri < minCoord.r) minCoord.r = ri;
                    if (idx < minCoord.c) minCoord.c = idx;
                }
                if (maxCoord == undefined) {
                    maxCoord = {
                        r: ri,
                        c: idx,
                    };
                } else {
                    if (ri > maxCoord.r) maxCoord.r = ri;
                    if (idx > maxCoord.c) maxCoord.c = idx;
                }

                var cellText = row.cells[k].text,
                    type = "s";
                if (!cellText) {
                    cellText = "";
                    type = "z";
                } else if (!isNaN(parseFloat(cellText))) {
                    cellText = parseFloat(cellText);
                    type = "n";
                } else if (cellText === "true" || cellText === "false") {
                    cellText = Boolean(cellText);
                    type = "b";
                }

                ws[lastRef] = {
                    v: cellText,
                    t: type,
                };

                if (keepFormulas && type == "s" && cellText[0] == "=") {
                    ws[lastRef].f = cellText.slice(1);
                }

                if (keepMerges && row.cells[k].merge != undefined) {
                    if (ws["!merges"] == undefined) ws["!merges"] = [];

                    ws["!merges"].push({
                        s: {
                            r: ri,
                            c: idx,
                        },
                        e: {
                            r: ri + row.cells[k].merge[0],
                            c: idx + row.cells[k].merge[1],
                        },
                    });
                }
            });

            ws["!ref"] =
                XLSX.utils.encode_cell({ r: minCoord.r, c: minCoord.c }) +
                ":" +
                XLSX.utils.encode_cell({ r: maxCoord.r, c: maxCoord.c });
        }

        XLSX.utils.book_append_sheet(out, ws, xws.name);
    });

    return out;
}
