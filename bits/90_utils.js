function decode_row(rowstr) { return parseInt(unfix_row(rowstr),10) - 1; }
function encode_row(row) { return "" + (row + 1); }
function fix_row(cstr) { return cstr.replace(/([A-Z]|^)(\d+)$/,"$1$$$2"); }
function unfix_row(cstr) { return cstr.replace(/\$(\d+)$/,"$1"); }

function decode_col(colstr) { var c = unfix_col(colstr), d = 0, i = 0; for(; i !== c.length; ++i) d = 26*d + c.charCodeAt(i) - 64; return d - 1; }
function encode_col(col) { var s=""; for(++col; col; col=Math.floor((col-1)/26)) s = String.fromCharCode(((col-1)%26) + 65) + s; return s; }
function fix_col(cstr) { return cstr.replace(/^([A-Z])/,"$$$1"); }
function unfix_col(cstr) { return cstr.replace(/^\$([A-Z])/,"$1"); }

function split_cell(cstr) { return cstr.replace(/(\$?[A-Z]*)(\$?\d*)/,"$1,$2").split(","); }
function decode_cell(cstr) { var splt = split_cell(cstr); return { c:decode_col(splt[0]), r:decode_row(splt[1]) }; }
function encode_cell(cell) { return encode_col(cell.c) + encode_row(cell.r); }
function fix_cell(cstr) { return fix_col(fix_row(cstr)); }
function unfix_cell(cstr) { return unfix_col(unfix_row(cstr)); }
function decode_range(range) { var x =range.split(":").map(decode_cell); return {s:x[0],e:x[x.length-1]}; }
function encode_range(cs,ce) {
	if(ce === undefined || typeof ce === 'number') return encode_range(cs.s, cs.e);
	if(typeof cs !== 'string') cs = encode_cell(cs); if(typeof ce !== 'string') ce = encode_cell(ce);
	return cs == ce ? cs : cs + ":" + ce;
}

function safe_decode_range(range) {
	var o = {s:{c:0,r:0},e:{c:0,r:0}};
	var idx = 0, i = 0, cc = 0;
	var len = range.length;
	for(idx = 0; i < len; ++i) {
		if((cc=range.charCodeAt(i)-64) < 1 || cc > 26) break;
		idx = 26*idx + cc;
	}
	o.s.c = --idx;

	for(idx = 0; i < len; ++i) {
		if((cc=range.charCodeAt(i)-48) < 0 || cc > 9) break;
		idx = 10*idx + cc;
	}
	o.s.r = --idx;

	if(i === len || range.charCodeAt(++i) === 58) { o.e.c=o.s.c; o.e.r=o.s.r; return o; }

	for(idx = 0; i != len; ++i) {
		if((cc=range.charCodeAt(i)-64) < 1 || cc > 26) break;
		idx = 26*idx + cc;
	}
	o.e.c = --idx;

	for(idx = 0; i != len; ++i) {
		if((cc=range.charCodeAt(i)-48) < 0 || cc > 9) break;
		idx = 10*idx + cc;
	}
	o.e.r = --idx;
	return o;
}

function safe_format_cell(cell, v) {
	if(cell.z !== undefined) try { return (cell.w = SSF.format(cell.z, v)); } catch(e) { }
	if(!cell.XF) return v;
	try { return (cell.w = SSF.format(cell.XF.ifmt||0, v)); } catch(e) { return ''+v; }
}

function format_cell(cell, v) {
	if(cell == null || cell.t == null) return "";
	if(cell.w !== undefined) return cell.w;
	if(v === undefined) return safe_format_cell(cell, cell.v);
	return safe_format_cell(cell, v);
}

function sheet_to_json(sheet, opts){
	var val, row, range, header = 0, offset = 1, r, hdr = [], isempty, R, C, v;
	var o = opts != null ? opts : {};
	var raw = o.raw;
	if(sheet == null || sheet["!ref"] == null) return [];
	range = o.range !== undefined ? o.range : sheet["!ref"];
	if(o.header === 1) header = 1;
	else if(o.header === "A") header = 2;
	else if(Array.isArray(o.header)) header = 3;
	switch(typeof range) {
		case 'string': r = safe_decode_range(range); break;
		case 'number': r = safe_decode_range(sheet["!ref"]); r.s.r = range; break;
		default: r = range;
	}
	if(header > 0) offset = 0;
	var rr = encode_row(r.s.r);
	var cols = new Array(r.e.c-r.s.c+1);
	var out = new Array(r.e.r-r.s.r-offset+1);
	var outi = 0;
	for(C = r.s.c; C <= r.e.c; ++C) {
		cols[C] = encode_col(C);
		val = sheet[cols[C] + rr];
		switch(header) {
			case 1: hdr[C] = C; break;
			case 2: hdr[C] = cols[C]; break;
			case 3: hdr[C] = o.header[C - r.s.c]; break;
			default:
				if(val === undefined) continue;
				hdr[C] = format_cell(val);
		}
	}

	for (R = r.s.r + offset; R <= r.e.r; ++R) {
		rr = encode_row(R);
		isempty = true;
		row = header === 1 ? [] : Object.create({ __rowNum__ : R });
		for (C = r.s.c; C <= r.e.c; ++C) {
			val = sheet[cols[C] + rr];
			if(val === undefined || val.t === undefined) continue;
			v = val.v;
			switch(val.t){
				case 'e': continue;
				case 's': break;
				case 'b': case 'n': break;
				default: throw 'unrecognized type ' + val.t;
			}
			if(v !== undefined) {
				row[hdr[C]] = raw ? v : format_cell(val,v);
				isempty = false;
			}
		}
		if(isempty === false) out[outi++] = row;
	}
	out.length = outi;
	return out;
}

function sheet_to_row_object_array(sheet, opts) { return sheet_to_json(sheet, opts != null ? opts : {}); }

function sheet_to_csv(sheet, opts) {
	var out = "", txt = "", qreg = /"/g;
	var o = opts == null ? {} : opts;
	if(sheet == null || sheet["!ref"] == null) return "";
	var r = safe_decode_range(sheet["!ref"]);
	var FS = o.FS !== undefined ? o.FS : ",", fs = FS.charCodeAt(0);
	var RS = o.RS !== undefined ? o.RS : "\n", rs = RS.charCodeAt(0);
	var row = "", rr = "", cols = [];
	var i = 0, cc = 0, val;
	var R = 0, C = 0;
	for(C = r.s.c; C <= r.e.c; ++C) cols[C] = encode_col(C);
	for(R = r.s.r; R <= r.e.r; ++R) {
		row = "";
		rr = encode_row(R);
		for(C = r.s.c; C <= r.e.c; ++C) {
			val = sheet[cols[C] + rr];
			txt = val !== undefined ? ''+format_cell(val) : "";
			for(i = 0, cc = 0; i !== txt.length; ++i) if((cc = txt.charCodeAt(i)) === fs || cc === rs || cc === 34) {
				txt = "\"" + txt.replace(qreg, '""') + "\""; break; }
			row += (C === r.s.c ? "" : FS) + txt;
		}
		out += row + RS;
	}
	return out;
}
var make_csv = sheet_to_csv;

function sheet_to_formulae(sheet) {
	var cmds, y = "", x, val="";
	if(sheet == null || sheet["!ref"] == null) return "";
	var r = safe_decode_range(sheet['!ref']), rr = "", cols = [], C;
	cmds = new Array((r.e.r-r.s.r+1)*(r.e.c-r.s.c+1));
	var i = 0;
	for(C = r.s.c; C <= r.e.c; ++C) cols[C] = encode_col(C);
	for(var R = r.s.r; R <= r.e.r; ++R) {
		rr = encode_row(R);
		for(C = r.s.c; C <= r.e.c; ++C) {
			y = cols[C] + rr;
			x = sheet[y];
			val = "";
			if(x === undefined) continue;
			if(x.f != null) val = x.f;
			else if(x.w !== undefined) val = "'" + x.w;
			else if(x.v === undefined) continue;
			else val = ""+x.v;
			cmds[i++] = y + "=" + val;
		}
	}
	cmds.length = i;
	return cmds;
}

var utils = {
	encode_col: encode_col,
	encode_row: encode_row,
	encode_cell: encode_cell,
	encode_range: encode_range,
	decode_col: decode_col,
	decode_row: decode_row,
	split_cell: split_cell,
	decode_cell: decode_cell,
	decode_range: decode_range,
	format_cell: format_cell,
	get_formulae: sheet_to_formulae,
	make_csv: sheet_to_csv,
	make_json: sheet_to_json,
	make_formulae: sheet_to_formulae,
	sheet_to_csv: sheet_to_csv,
	sheet_to_json: sheet_to_json,
	sheet_to_formulae: sheet_to_formulae,
	sheet_to_row_object_array: sheet_to_row_object_array
};

if ((typeof 'module' != 'undefined'  && typeof require != 'undefined') || (typeof $ != 'undefined')) {
  var StyleBuilder = function (options) {

    if(typeof module !== "undefined" && typeof require !== 'undefined' ) {
      var cheerio = require('cheerio');
      createElement = function(str) { return cheerio(cheerio(str, null, null, {xmlMode: true})); };
    }
    else if (typeof jQuery !== 'undefined' || typeof $ !== 'undefined') {
      createElement = function(str) { return $(str); }
    }
    else {
      createElement = function() { }
    }


    var customNumFmtId = 164;


    var table_fmt = {
      0:  'General',
      1:  '0',
      2:  '0.00',
      3:  '#,##0',
      4:  '#,##0.00',
      9:  '0%',
      10: '0.00%',
      11: '0.00E+00',
      12: '# ?/?',
      13: '# ??/??',
      14: 'm/d/yy',
      15: 'd-mmm-yy',
      16: 'd-mmm',
      17: 'mmm-yy',
      18: 'h:mm AM/PM',
      19: 'h:mm:ss AM/PM',
      20: 'h:mm',
      21: 'h:mm:ss',
      22: 'm/d/yy h:mm',
      37: '#,##0 ;(#,##0)',
      38: '#,##0 ;[Red](#,##0)',
      39: '#,##0.00;(#,##0.00)',
      40: '#,##0.00;[Red](#,##0.00)',
      45: 'mm:ss',
      46: '[h]:mm:ss',
      47: 'mmss.0',
      48: '##0.0E+0',
      49: '@',
      56: '"上午/下午 "hh"時"mm"分"ss"秒 "',
      65535: 'General'
    };
    var fmt_table = {};

    for (var idx in table_fmt) {
      fmt_table[table_fmt[idx]] = idx;

    }


    var baseXmlprefix = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    var baseXml =
        '<styleSheet xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"\
        xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" mc:Ignorable="x14ac">\
            <numFmts count="1">\
              <numFmt numFmtId="164" formatCode="0.00%"/>\
            </numFmts>\
            <fonts count="0" x14ac:knownFonts="1"></fonts>\
            <fills count="0"></fills>\
            <borders count="0"></borders>\
            <cellStyleXfs count="1">\
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>\
            </cellStyleXfs>\
            <cellXfs count="0"></cellXfs>\
            <cellStyles count="1">\
              <cellStyle name="Normal" xfId="0" builtinId="0"/>\
            </cellStyles>\
            <dxfs count="0"/>\
            <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleMedium4"/>\
        </styleSheet>';

    _hashIndex = {};
    _listIndex = [];

    return {

      initialize: function (options) {
        if (typeof cheerio !== 'undefined') {
          this.$styles = cheerio.load(baseXml, {xmlMode: true});
          this.$styles.find = function(q) { return this(q)}
        }
        else {
          this.$styles = $(baseXml);
        }


        // need to specify styles at index 0 and 1.
        // the second style MUST be gray125 for some reason

        var defaultStyle = options.defaultCellStyle;
        if (!defaultStyle) defaultStyle = {
          font: {name: 'Calibri', sz: '11'},
          fill: { fgColor: { patternType: "none"}},
          border: {},
          numFmt: null
        };
        if (!defaultStyle.border) { defaultStyle.border = {}}

        var gray125Style = JSON.parse(JSON.stringify(defaultStyle));
        gray125Style.fill = { fgColor: { patternType: "gray125"}}

        this.addStyles([defaultStyle, gray125Style]);
        return this;
      },

      // create a style entry and returns an integer index that can be used in the cell .s property
      // these format of this object follows the emerging Common Spreadsheet Format
      addStyle: function (attributes) {
        var attributes = this._duckTypeStyle(attributes);
        var hashKey = JSON.stringify(attributes);
        var index = _hashIndex[hashKey];
        if (index == undefined) {
          index = this._addXf(attributes || {});
          _hashIndex[hashKey] = index;

        }
        else {
          index = _hashIndex[hashKey];
        }
        return index;
      },

      // create style entries and returns array of integer indexes that can be used in cell .s property
      addStyles: function (styles) {
        var self = this;
        return styles.map(function (style) {
          return self.addStyle(style);
        })
      },

      _duckTypeStyle: function(attributes) {

        if (typeof attributes == 'object' && (attributes.patternFill || attributes.fgColor)) {
          return {fill: attributes }; // this must be read via XLSX.parseFile(...)
        }
        else if (attributes.font || attributes.numFmt || attributes.border || attributes.fill) {
          return attributes;
        }
        else {
          return this._getStyleCSS(attributes)
        }
      },

      _getStyleCSS: function(css) {
        return css; //TODO
      },

      // Create an <xf> record for the style as well as corresponding <font>, <fill>, <border>, <numfmts>
      // Right now this is simple and creates a <font>, <fill>, <border>, <numfmts> for every <xf>
      // We could perhaps get fancier and avoid duplicating  auxiliary entries as Excel presumably intended, but bother.
      _addXf: function (attributes) {


        var fontId = this._addFont(attributes.font);
        var fillId = this._addFill(attributes.fill);
        var borderId = this._addBorder(attributes.border);
        var numFmtId = this._addNumFmt(attributes.numFmt);

        var $xf = createElement('<xf></xf>')
            .attr("numFmtId", numFmtId)
            .attr("fontId", fontId)
            .attr("fillId", fillId)
            .attr("borderId", 0)
            .attr("xfId", "0");

        if (fontId > 0) {
          $xf.attr('applyFont', "1");
        }
        if (fillId > 0) {
          $xf.attr('applyFill', "1");
        }
        if (borderId > 0) {
          $xf.attr('applyBorder', "1");
        }
        if (numFmtId > 0) {
          $xf.attr('applyNumberFormat', "1");
        }

        if (attributes.alignment) {
          var $alignment = createElement('<alignment></alignment>');
          if (attributes.alignment.horizontal) { $alignment.attr('horizontal', attributes.alignment.horizontal);}
          if (attributes.alignment.vertical)  { $alignment.attr('vertical', attributes.alignment.vertical);}
          if (attributes.alignment.indent)  { $alignment.attr('indent', attributes.alignment.indent);}
          if (attributes.alignment.wrapText)  { $alignment.attr('wrapText', attributes.alignment.wrapText);}
          $xf.append($alignment).attr('applyAlignment',1)

        }

        var $cellXfs = this.$styles.find('cellXfs');

        $cellXfs.append($xf);
        var count = +$cellXfs.attr('count') + 1;

        $cellXfs.attr('count', count);
        return count - 1;
      },

      _addFont: function (attributes) {
        if (!attributes) {
          return 0;
        }

        var $font = createElement('<font/>', null, null, {xmlMode: true});

        $font.append(createElement('<sz/>').attr('val', attributes.sz))
            .append(createElement('<color/>').attr('theme', '1'))
            .append(createElement('<name/>').attr('val', attributes.name))
//              .append(createElement('<family/>').attr('val', '2'))
//              .append(createElement('<scheme/>').attr('val', 'minor'));

        if (attributes.bold) $font.append('<b/>');
        if (attributes.underline) $font.append('<u/>');
        if (attributes.italic) $font.append('<i/>');

        if (attributes.color) {
          if (attributes.color.theme) {
            $font.append(createElement('<color/>').attr('theme', attributes.color.theme));
          } else if (attributes.color.rgb) {
            $font.append(createElement('<color/>').attr('rgb', attributes.color.rgb));
          }
        }


        var $fonts = this.$styles.find('fonts');
        $fonts.append($font);

        var count = $fonts.children().length;
        $fonts.attr('count', count);
        return count - 1;
      },

      _addNumFmt: function (numFmt) {
        if (!numFmt) {
          return 0;
        }

        if (typeof numFmt == 'string') {
          var numFmtIdx = fmt_table[numFmt];
          if (numFmtIdx >= 0) {
            return numFmtIdx; // we found a match against built in formats
          }
        }

        if (numFmt == +numFmt) {
          return numFmt; // we're matching an integer against some known code
        }

        var $numFmt = createElement('<numFmt/>', null, null, {xmlMode: true})
            .attr("numFmtId", ++customNumFmtId )
            .attr("formatCode", numFmt);

        var $numFmts = this.$styles.find('numFmts');
        $numFmts.append($numFmt);

        var count = $numFmts.children().length;
        $numFmts.attr('count', count);
        return customNumFmtId;
      },

      _addFill: function (attributes) {

        if (!attributes) {
          return 0;
        }
        var $patternFill = createElement('<patternFill></patternFill>', null, null, {xmlMode: true})
            .attr('patternType', attributes.patternType || 'solid');

        if (attributes.fgColor) {
          //Excel doesn't like it when we set both rgb and theme+tint, but xlsx.parseFile() sets both
          //var $fgColor = createElement('<fgColor/>', null, null, {xmlMode: true}).attr(attributes.fgColor)
          if (attributes.fgColor.rgb) {

            if (attributes.fgColor.rgb.length == 6) {
              attributes.fgColor.rgb = "FF" + attributes.fgColor.rgb /// add alpha to an RGB as Excel expects aRGB
            }
            var $fgColor = createElement('<fgColor/>', null, null, {xmlMode: true}).
                attr('rgb', attributes.fgColor.rgb);
            $patternFill.append($fgColor);
          }
          else if (attributes.fgColor.theme) {
            var $fgColor = createElement('<fgColor/>', null, null, {xmlMode: true});
            $fgColor.attr('theme', attributes.fgColor.theme);
            if (attributes.fgColor.tint) {
              $fgColor.attr('tint', attributes.fgColor.tint);
            }
            $patternFill.append($fgColor);
          }

          if (!attributes.bgColor) {
            attributes.bgColor = { "indexed": "64"}
          }
        }

        if (attributes.bgColor) {
          var $bgColor = createElement('<bgColor/>', null, null, {xmlMode: true}).attr(attributes.bgColor);
          $patternFill.append($bgColor);
        }

        var $fill = createElement('<fill></fill>')
            .append($patternFill);

        this.$styles.find('fills').append($fill);
        var $fills = this.$styles.find('fills')
        $fills.append($fill);

        var count = $fills.children().length;
        $fills.attr('count', count);
        return count - 1;
      },

      _addBorder: function (attributes) {
        if (!attributes) {
          return 0;
        }
        var $border = createElement('<border></border>')
            .append('<left></left>')
            .append('<right></right>')
            .append('<top></top>')
            .append('<bottom></bottom>')
            .append('<diagonal></diagonal>');

        var $borders = this.$styles.find('borders');
        $borders.append($border);

        var count = $borders.children().length;
        $borders.attr('count', count);
        return count;
      },

      toXml: function () {
        if (this.$styles.find('numFmts').children().length == 0) {
          this.$styles.find('numFmts').remove();
        }
        if (this.$styles.xml) { return this.$styles.xml(); }
        else { return baseXmlprefix + this.$styles.html(); }
      }
    }.initialize(options||{});
  }
}
