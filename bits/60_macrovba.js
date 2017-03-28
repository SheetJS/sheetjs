RELS.DS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet";
RELS.MS = "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet";

/* macro and dialog sheet stubs */
function parse_ds_bin() { return {'!type':'dialog'}; }
function parse_ds_xml() { return {'!type':'dialog'}; }
function parse_ms_bin() { return {'!type':'macro'}; }
function parse_ms_xml() { return {'!type':'macro'}; }
