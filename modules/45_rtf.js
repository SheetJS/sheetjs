function rtf_to_sheet(d, opts) {
  switch (opts.type) {
    case "base64":
      return rtf_to_sheet_str(Base64_decode(d), opts);
    case "binary":
      return rtf_to_sheet_str(d, opts);
    case "buffer":
      return rtf_to_sheet_str(has_buf && Buffer.isBuffer(d) ? d.toString("binary") : a2s(d), opts);
    case "array":
      return rtf_to_sheet_str(cc2str(d), opts);
  }
  throw new Error("Unrecognized type " + opts.type);
}
function rtf_to_sheet_str(str, opts) {
  var o = opts || {};
  var ws = o.dense ? [] : {};
  var rows = str.match(/\\trowd[\s\S]*?\\row\b/g);
  if (!rows)
    throw new Error("RTF missing table");
  var range = { s: { c: 0, r: 0 }, e: { c: 0, r: rows.length - 1 } };
  rows.forEach(function(rowtf, R) {
    if (Array.isArray(ws))
      ws[R] = [];
    var rtfre = /\\[\w\-]+\b/g;
    var last_index = 0;
    var res;
    var C = -1;
    var payload = [];
    while ((res = rtfre.exec(rowtf)) != null) {
      var data = rowtf.slice(last_index, rtfre.lastIndex - res[0].length);
      if (data.charCodeAt(0) == 32)
        data = data.slice(1);
      if (data.length)
        payload.push(data);
      switch (res[0]) {
        case "\\cell":
          ++C;
          if (payload.length) {
            var cell = { v: payload.join(""), t: "s" };
            if (cell.v == "TRUE" || cell.v == "FALSE") {
              cell.v = cell.v == "TRUE";
              cell.t = "b";
            } else if (!isNaN(fuzzynum(cell.v))) {
              cell.t = "n";
              if (o.cellText !== false)
                cell.w = cell.v;
              cell.v = fuzzynum(cell.v);
            }
            if (Array.isArray(ws))
              ws[R][C] = cell;
            else
              ws[encode_cell({ r: R, c: C })] = cell;
          }
          payload = [];
          break;
        case "\\par":
          payload.push("\n");
          break;
      }
      last_index = rtfre.lastIndex;
    }
    if (C > range.e.c)
      range.e.c = C;
  });
  ws["!ref"] = encode_range(range);
  return ws;
}
function rtf_to_workbook(d, opts) {
  var wb = sheet_to_workbook(rtf_to_sheet(d, opts), opts);
  wb.bookType = "rtf";
  return wb;
}
function sheet_to_rtf(ws, opts) {
  var o = ["{\\rtf1\\ansi"];
  if (!ws["!ref"])
    return o[0] + "}";
  var r = safe_decode_range(ws["!ref"]), cell;
  var dense = Array.isArray(ws);
  for (var R = r.s.r; R <= r.e.r; ++R) {
    o.push("\\trowd\\trautofit1");
    for (var C = r.s.c; C <= r.e.c; ++C)
      o.push("\\cellx" + (C + 1));
    o.push("\\pard\\intbl");
    for (C = r.s.c; C <= r.e.c; ++C) {
      var coord = encode_cell({ r: R, c: C });
      cell = dense ? (ws[R] || [])[C] : ws[coord];
      if (!cell || cell.v == null && (!cell.f || cell.F)) {
        o.push(" \\cell");
        continue;
      }
      o.push(" " + (cell.w || (format_cell(cell), cell.w) || "").replace(/[\r\n]/g, "\\par "));
      o.push("\\cell");
    }
    o.push("\\pard\\intbl\\row");
  }
  return o.join("") + "}";
}
