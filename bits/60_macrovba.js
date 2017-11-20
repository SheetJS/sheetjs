RELS.DS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet";
RELS.MS = "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet";

/* macro and dialog sheet stubs */
function parse_ds_bin(data/*:any*/, opts, idx/*:number*/, rels, wb, themes, styles)/*:Worksheet*/ { return {'!type':'dialog'}; }
function parse_ds_xml(data/*:any*/, opts, idx/*:number*/, rels, wb, themes, styles)/*:Worksheet*/ { return {'!type':'dialog'}; }
function parse_ms_bin(data/*:any*/, opts, idx/*:number*/, rels, wb, themes, styles)/*:Worksheet*/ { return {'!type':'macro'}; }
function parse_ms_xml(data/*:any*/, opts, idx/*:number*/, rels, wb, themes, styles)/*:Worksheet*/ { return {'!type':'macro'}; }
