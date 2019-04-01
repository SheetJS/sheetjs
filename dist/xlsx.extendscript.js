/*! xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/*! shim.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* ES3/5 Compatibility shims and other utilities for older browsers. */

// From https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object/keys
if(!Object.keys) Object.keys = (function() {
  var hasOwnProperty = Object.prototype.hasOwnProperty,
      hasDontEnumBug = !({toString: null}).propertyIsEnumerable('toString'),
      dontEnums = [
        'toString',
        'toLocaleString',
        'valueOf',
        'hasOwnProperty',
        'isPrototypeOf',
        'propertyIsEnumerable',
        'constructor'
      ],
      dontEnumsLength = dontEnums.length;

  return function(obj) {
    if(typeof obj !== 'object' && typeof obj !== 'function' || obj === null) throw new TypeError('Object.keys called on non-object');

    var result = [];

    for(var prop in obj) if(hasOwnProperty.call(obj, prop)) result.push(prop);

    if(hasDontEnumBug)
      for(var i=0; i < dontEnumsLength; ++i)
        if(hasOwnProperty.call(obj, dontEnums[i])) result.push(dontEnums[i]);
    return result;
  };
})();

if(!String.prototype.trim) String.prototype.trim = function() {
  var s = this.replace(/^\s+/, '');
  for(var i = s.length - 1; i >=0 ; --i) if(!s.charAt(i).match(/^\s/)) return s.slice(0,i+1);
  return "";
};

if(!Array.prototype.forEach) Array.prototype.forEach = function(cb) {
  var len = (this.length>>>0), self = (arguments[1]||void 0);
  for(var i=0; i<len; ++i) if(i in this) self ? cb.call(self, this[i], i, this) : cb(this[i], i, this);
};

if(!Array.prototype.map) Array.prototype.map = function(cb) {
  var len = (this.length>>>0), self = (arguments[1]||void 0), A = new Array(len);
  for(var i=0; i<len; ++i) if(i in this) A[i] = self ? cb.call(self, this[i], i, this) : cb(this[i], i, this);
  return A;
};

if(!Array.prototype.indexOf) Array.prototype.indexOf = function(needle) {
  var len = (this.length>>>0), i = ((arguments[1]|0)||0);
  for(i<0 && (i+=len)<0 && (i=0); i<len; ++i) if(this[i] === needle) return i;
  return -1;
};

if(!Array.prototype.lastIndexOf) Array.prototype.lastIndexOf = function(needle) {
  var len = (this.length>>>0), i = len - 1;
  for(; i>=0; --i) if(this[i] === needle) return i;
  return -1;
};

if(!Array.isArray) Array.isArray = function(obj) { return Object.prototype.toString.call(obj) === "[object Array]"; };

if(!Date.prototype.toISOString) Date.prototype.toISOString = (function() {
  function p(n,i) { return ('0000000' + n).slice(-(i||2)); }

  return function _toISOString() {
    var y = this.getUTCFullYear(), yr = "";
    if(y>9999)   yr = '+' + p( y, 6);
    else if(y<0) yr = '-' + p(-y, 6);
    else         yr =       p( y, 4);

    return [
      yr, p(this.getUTCMonth()+1), p(this.getUTCDate())
    ].join('-') + 'T' + [
      p(this.getUTCHours()), p(this.getUTCMinutes()), p(this.getUTCSeconds())
    ].join(':') + '.' + p(this.getUTCMilliseconds(),3) + 'Z';
  };
}());

if(typeof ArrayBuffer !== 'undefined' && !ArrayBuffer.prototype.slice) ArrayBuffer.prototype.slice = function(start, end) {
  if(start == null) start = 0;
  if(start < 0) { start += this.byteLength; if(start < 0) start = 0; }
  if(start >= this.byteLength) return new Uint8Array(0);
  if(end == null) end = this.byteLength;
  if(end < 0) { end += this.byteLength; if(end < 0) end = 0; }
  if(end > this.byteLength) end = this.byteLength;
  if(start > end) return new Uint8Array(0);
  var out = new ArrayBuffer(end - start);
  var view = new Uint8Array(out);
  var data = new Uint8Array(this, start, end - start)
  /* IE10 should have Uint8Array#set */
  if(view.set) view.set(data); else while(start <= --end) view[end - start] = data[end];
  return out;
};
if(typeof Uint8Array !== 'undefined' && !Uint8Array.prototype.slice) Uint8Array.prototype.slice = function(start, end) {
  if(start == null) start = 0;
  if(start < 0) { start += this.length; if(start < 0) start = 0; }
  if(start >= this.length) return new Uint8Array(0);
  if(end == null) end = this.length;
  if(end < 0) { end += this.length; if(end < 0) end = 0; }
  if(end > this.length) end = this.length;
  if(start > end) return new Uint8Array(0);
  var out = new Uint8Array(end - start);
  while(start <= --end) out[end - start] = this[end];
  return out;
};

// VBScript + ActiveX fallback for IE5+
var IE_SaveFile = (function() { try {
  if(typeof IE_SaveFile_Impl == "undefined") document.write([
'<script type="text/vbscript" language="vbscript">',
'IE_GetProfileAndPath_Key = "HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\User Shell Folders\\"',
'Function IE_GetProfileAndPath(key): Set wshell = CreateObject("WScript.Shell"): IE_GetProfileAndPath = wshell.RegRead(IE_GetProfileAndPath_Key & key): IE_GetProfileAndPath = wshell.ExpandEnvironmentStrings("%USERPROFILE%") & "!" & IE_GetProfileAndPath: End Function',
'Function IE_SaveFile_Impl(FileName, payload): Dim data, plen, i, bit: data = CStr(payload): plen = Len(data): Set fso = CreateObject("Scripting.FileSystemObject"): fso.CreateTextFile FileName, True: Set f = fso.GetFile(FileName): Set stream = f.OpenAsTextStream(2, 0): For i = 1 To plen Step 3: bit = Mid(data, i, 2): stream.write Chr(CLng("&h" & bit)): Next: stream.Close: IE_SaveFile_Impl = True: End Function',
'|/script>'.replace("|","<")
  ].join("\r\n"));
  if(typeof IE_SaveFile_Impl == "undefined") return void 0;
  var IE_GetPath = (function() {
    var DDP1 = "";
    try { DDP1 = IE_GetProfileAndPath("{374DE290-123F-4565-9164-39C4925E467B}"); } catch(e) { try { DDP1 = IE_GetProfileAndPath("Personal"); } catch(e) { try { DDP1 = IE_GetProfileAndPath("Desktop"); } catch(e) { throw e; }}}
    var o = DDP1.split("!");
    DDP = o[1].replace("%USERPROFILE%", o[0]);
    return function(path) { return DDP + "\\" + path; };
  })();
  function fix_data(data) {
    var out = [];
    var T = typeof data == "string";
    for(var i = 0; i < data.length; ++i) out.push(("00"+(T ? data.charCodeAt(i) : data[i]).toString(16)).slice(-2));
    var o = out.join("|");
    return o;
  }
  return function(data, filename) { return IE_SaveFile_Impl(IE_GetPath(filename), fix_data(data)); };
} catch(e) { return void 0; }})();
var IE_LoadFile = (function() { try {
  if(typeof IE_LoadFile_Impl == "undefined") document.write([
'<script type="text/vbscript" language="vbscript">',
'Function IE_LoadFile_Impl(FileName): Dim out(), plen, i, cc: Set fso = CreateObject("Scripting.FileSystemObject"): Set f = fso.GetFile(FileName): Set stream = f.OpenAsTextStream(1, 0): plen = f.Size: ReDim out(plen): For i = 1 To plen Step 1: cc = Hex(Asc(stream.read(1))): If Len(cc) < 2 Then: cc = "0" & cc: End If: out(i) = cc: Next: IE_LoadFile_Impl = Join(out,""): End Function',
'|/script>'.replace("|","<")
  ].join("\r\n"));
  if(typeof IE_LoadFile_Impl == "undefined") return void 0;
  function fix_data(data) {
    var out = [];
    for(var i = 0; i < data.length; i+=2) out.push(String.fromCharCode(parseInt(data.slice(i, i+2), 16)));
    var o = out.join("");
    return o;
  }
  return function(filename) { return fix_data(IE_LoadFile_Impl(filename)); };
} catch(e) { return void 0; }})();

// getComputedStyle polyfill from https://gist.github.com/8HNHoFtE/5891086
if(typeof window !== 'undefined' && typeof window.getComputedStyle !== 'function') {
  window.getComputedStyle = function(e,t){return this.el=e,this.getPropertyValue=function(t){var n=/(\-([a-z]){1})/g;return t=="float"&&(t="styleFloat"),n.test(t)&&(t=t.replace(n,function(){return arguments[2].toUpperCase()})),e.currentStyle[t]?e.currentStyle[t]:null},this}
}
var DO_NOT_EXPORT_CODEPAGE = true;
var DO_NOT_EXPORT_JSZIP = true;
/*

JSZip - A Javascript class for generating and reading zip files
<http://stuartk.com/jszip>

(c) 2009-2014 Stuart Knightley <stuart [at] stuartk.com>
Dual licenced under the MIT license or GPLv3. See https://raw.github.com/Stuk/jszip/master/LICENSE.markdown.

JSZip uses the library pako released under the MIT license :
https://github.com/nodeca/pako/blob/master/LICENSE

Note: since JSZip 3 removed critical functionality, this version assigns to the
`JSZipSync` variable.  Another JSZip version can be loaded in parallel.
*/
(function(e){
	if("object"==typeof exports&&"undefined"!=typeof module&&"undefined"==typeof DO_NOT_EXPORT_JSZIP)module.exports=e();
	else if("function"==typeof define&&define.amd&&"undefined"==typeof DO_NOT_EXPORT_JSZIP){JSZipSync=e();define([],e);}
	else{
		var f;
		"undefined"!=typeof window?f=window:
		"undefined"!=typeof global?f=global:
		"undefined"!=typeof $ && $.global?f=$.global:
		"undefined"!=typeof self&&(f=self),f.JSZipSync=e()
	}
}(function(){var define,module,exports;return (function e(t,n,r){function s(o,u){if(!n[o]){if(!t[o]){var a=typeof require=="function"&&require;if(!u&&a)return a(o,!0);if(i)return i(o,!0);throw new Error("Cannot find module '"+o+"'")}var f=n[o]={exports:{}};t[o][0].call(f.exports,function(e){var n=t[o][1][e];return s(n?n:e)},f,f.exports,e,t,n,r)}return n[o].exports}var i=typeof require=="function"&&require;for(var o=0;o<r.length;o++)s(r[o]);return s})({1:[function(_dereq_,module,exports){
'use strict';
// private property
var _keyStr = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=";


// public method for encoding
exports.encode = function(input, utf8) {
    var output = "";
    var chr1, chr2, chr3, enc1, enc2, enc3, enc4;
    var i = 0;

    while (i < input.length) {

        chr1 = input.charCodeAt(i++);
        chr2 = input.charCodeAt(i++);
        chr3 = input.charCodeAt(i++);

        enc1 = chr1 >> 2;
        enc2 = ((chr1 & 3) << 4) | (chr2 >> 4);
        enc3 = ((chr2 & 15) << 2) | (chr3 >> 6);
        enc4 = chr3 & 63;

        if (isNaN(chr2)) {
            enc3 = enc4 = 64;
        }
        else if (isNaN(chr3)) {
            enc4 = 64;
        }

        output = output + _keyStr.charAt(enc1) + _keyStr.charAt(enc2) + _keyStr.charAt(enc3) + _keyStr.charAt(enc4);

    }

    return output;
};

// public method for decoding
exports.decode = function(input, utf8) {
    var output = "";
    var chr1, chr2, chr3;
    var enc1, enc2, enc3, enc4;
    var i = 0;

    input = input.replace(/[^A-Za-z0-9\+\/\=]/g, "");

    while (i < input.length) {

        enc1 = _keyStr.indexOf(input.charAt(i++));
        enc2 = _keyStr.indexOf(input.charAt(i++));
        enc3 = _keyStr.indexOf(input.charAt(i++));
        enc4 = _keyStr.indexOf(input.charAt(i++));

        chr1 = (enc1 << 2) | (enc2 >> 4);
        chr2 = ((enc2 & 15) << 4) | (enc3 >> 2);
        chr3 = ((enc3 & 3) << 6) | enc4;

        output = output + String.fromCharCode(chr1);

        if (enc3 != 64) {
            output = output + String.fromCharCode(chr2);
        }
        if (enc4 != 64) {
            output = output + String.fromCharCode(chr3);
        }

    }

    return output;

};

},{}],2:[function(_dereq_,module,exports){
'use strict';
function CompressedObject() {
    this.compressedSize = 0;
    this.uncompressedSize = 0;
    this.crc32 = 0;
    this.compressionMethod = null;
    this.compressedContent = null;
}

CompressedObject.prototype = {
    /**
     * Return the decompressed content in an unspecified format.
     * The format will depend on the decompressor.
     * @return {Object} the decompressed content.
     */
    getContent: function() {
        return null; // see implementation
    },
    /**
     * Return the compressed content in an unspecified format.
     * The format will depend on the compressed conten source.
     * @return {Object} the compressed content.
     */
    getCompressedContent: function() {
        return null; // see implementation
    }
};
module.exports = CompressedObject;

},{}],3:[function(_dereq_,module,exports){
'use strict';
exports.STORE = {
    magic: "\x00\x00",
    compress: function(content) {
        return content; // no compression
    },
    uncompress: function(content) {
        return content; // no compression
    },
    compressInputType: null,
    uncompressInputType: null
};
exports.DEFLATE = _dereq_('./flate');

},{"./flate":8}],4:[function(_dereq_,module,exports){
'use strict';

var utils = _dereq_('./utils');

var table = [
    0x00000000, 0x77073096, 0xEE0E612C, 0x990951BA,
    0x076DC419, 0x706AF48F, 0xE963A535, 0x9E6495A3,
    0x0EDB8832, 0x79DCB8A4, 0xE0D5E91E, 0x97D2D988,
    0x09B64C2B, 0x7EB17CBD, 0xE7B82D07, 0x90BF1D91,
    0x1DB71064, 0x6AB020F2, 0xF3B97148, 0x84BE41DE,
    0x1ADAD47D, 0x6DDDE4EB, 0xF4D4B551, 0x83D385C7,
    0x136C9856, 0x646BA8C0, 0xFD62F97A, 0x8A65C9EC,
    0x14015C4F, 0x63066CD9, 0xFA0F3D63, 0x8D080DF5,
    0x3B6E20C8, 0x4C69105E, 0xD56041E4, 0xA2677172,
    0x3C03E4D1, 0x4B04D447, 0xD20D85FD, 0xA50AB56B,
    0x35B5A8FA, 0x42B2986C, 0xDBBBC9D6, 0xACBCF940,
    0x32D86CE3, 0x45DF5C75, 0xDCD60DCF, 0xABD13D59,
    0x26D930AC, 0x51DE003A, 0xC8D75180, 0xBFD06116,
    0x21B4F4B5, 0x56B3C423, 0xCFBA9599, 0xB8BDA50F,
    0x2802B89E, 0x5F058808, 0xC60CD9B2, 0xB10BE924,
    0x2F6F7C87, 0x58684C11, 0xC1611DAB, 0xB6662D3D,
    0x76DC4190, 0x01DB7106, 0x98D220BC, 0xEFD5102A,
    0x71B18589, 0x06B6B51F, 0x9FBFE4A5, 0xE8B8D433,
    0x7807C9A2, 0x0F00F934, 0x9609A88E, 0xE10E9818,
    0x7F6A0DBB, 0x086D3D2D, 0x91646C97, 0xE6635C01,
    0x6B6B51F4, 0x1C6C6162, 0x856530D8, 0xF262004E,
    0x6C0695ED, 0x1B01A57B, 0x8208F4C1, 0xF50FC457,
    0x65B0D9C6, 0x12B7E950, 0x8BBEB8EA, 0xFCB9887C,
    0x62DD1DDF, 0x15DA2D49, 0x8CD37CF3, 0xFBD44C65,
    0x4DB26158, 0x3AB551CE, 0xA3BC0074, 0xD4BB30E2,
    0x4ADFA541, 0x3DD895D7, 0xA4D1C46D, 0xD3D6F4FB,
    0x4369E96A, 0x346ED9FC, 0xAD678846, 0xDA60B8D0,
    0x44042D73, 0x33031DE5, 0xAA0A4C5F, 0xDD0D7CC9,
    0x5005713C, 0x270241AA, 0xBE0B1010, 0xC90C2086,
    0x5768B525, 0x206F85B3, 0xB966D409, 0xCE61E49F,
    0x5EDEF90E, 0x29D9C998, 0xB0D09822, 0xC7D7A8B4,
    0x59B33D17, 0x2EB40D81, 0xB7BD5C3B, 0xC0BA6CAD,
    0xEDB88320, 0x9ABFB3B6, 0x03B6E20C, 0x74B1D29A,
    0xEAD54739, 0x9DD277AF, 0x04DB2615, 0x73DC1683,
    0xE3630B12, 0x94643B84, 0x0D6D6A3E, 0x7A6A5AA8,
    0xE40ECF0B, 0x9309FF9D, 0x0A00AE27, 0x7D079EB1,
    0xF00F9344, 0x8708A3D2, 0x1E01F268, 0x6906C2FE,
    0xF762575D, 0x806567CB, 0x196C3671, 0x6E6B06E7,
    0xFED41B76, 0x89D32BE0, 0x10DA7A5A, 0x67DD4ACC,
    0xF9B9DF6F, 0x8EBEEFF9, 0x17B7BE43, 0x60B08ED5,
    0xD6D6A3E8, 0xA1D1937E, 0x38D8C2C4, 0x4FDFF252,
    0xD1BB67F1, 0xA6BC5767, 0x3FB506DD, 0x48B2364B,
    0xD80D2BDA, 0xAF0A1B4C, 0x36034AF6, 0x41047A60,
    0xDF60EFC3, 0xA867DF55, 0x316E8EEF, 0x4669BE79,
    0xCB61B38C, 0xBC66831A, 0x256FD2A0, 0x5268E236,
    0xCC0C7795, 0xBB0B4703, 0x220216B9, 0x5505262F,
    0xC5BA3BBE, 0xB2BD0B28, 0x2BB45A92, 0x5CB36A04,
    0xC2D7FFA7, 0xB5D0CF31, 0x2CD99E8B, 0x5BDEAE1D,
    0x9B64C2B0, 0xEC63F226, 0x756AA39C, 0x026D930A,
    0x9C0906A9, 0xEB0E363F, 0x72076785, 0x05005713,
    0x95BF4A82, 0xE2B87A14, 0x7BB12BAE, 0x0CB61B38,
    0x92D28E9B, 0xE5D5BE0D, 0x7CDCEFB7, 0x0BDBDF21,
    0x86D3D2D4, 0xF1D4E242, 0x68DDB3F8, 0x1FDA836E,
    0x81BE16CD, 0xF6B9265B, 0x6FB077E1, 0x18B74777,
    0x88085AE6, 0xFF0F6A70, 0x66063BCA, 0x11010B5C,
    0x8F659EFF, 0xF862AE69, 0x616BFFD3, 0x166CCF45,
    0xA00AE278, 0xD70DD2EE, 0x4E048354, 0x3903B3C2,
    0xA7672661, 0xD06016F7, 0x4969474D, 0x3E6E77DB,
    0xAED16A4A, 0xD9D65ADC, 0x40DF0B66, 0x37D83BF0,
    0xA9BCAE53, 0xDEBB9EC5, 0x47B2CF7F, 0x30B5FFE9,
    0xBDBDF21C, 0xCABAC28A, 0x53B39330, 0x24B4A3A6,
    0xBAD03605, 0xCDD70693, 0x54DE5729, 0x23D967BF,
    0xB3667A2E, 0xC4614AB8, 0x5D681B02, 0x2A6F2B94,
    0xB40BBE37, 0xC30C8EA1, 0x5A05DF1B, 0x2D02EF8D
];

/**
 *
 *  Javascript crc32
 *  http://www.webtoolkit.info/
 *
 */
module.exports = function crc32(input, crc) {
    if (typeof input === "undefined" || !input.length) {
        return 0;
    }

    var isArray = utils.getTypeOf(input) !== "string";

    if (typeof(crc) == "undefined") {
        crc = 0;
    }
    var x = 0;
    var y = 0;
    var b = 0;

    crc = crc ^ (-1);
    for (var i = 0, iTop = input.length; i < iTop; i++) {
        b = isArray ? input[i] : input.charCodeAt(i);
        y = (crc ^ b) & 0xFF;
        x = table[y];
        crc = (crc >>> 8) ^ x;
    }

    return crc ^ (-1);
};
// vim: set shiftwidth=4 softtabstop=4:

},{"./utils":21}],5:[function(_dereq_,module,exports){
'use strict';
var utils = _dereq_('./utils');

function DataReader(data) {
    this.data = null; // type : see implementation
    this.length = 0;
    this.index = 0;
}
DataReader.prototype = {
    /**
     * Check that the offset will not go too far.
     * @param {string} offset the additional offset to check.
     * @throws {Error} an Error if the offset is out of bounds.
     */
    checkOffset: function(offset) {
        this.checkIndex(this.index + offset);
    },
    /**
     * Check that the specifed index will not be too far.
     * @param {string} newIndex the index to check.
     * @throws {Error} an Error if the index is out of bounds.
     */
    checkIndex: function(newIndex) {
        if (this.length < newIndex || newIndex < 0) {
            throw new Error("End of data reached (data length = " + this.length + ", asked index = " + (newIndex) + "). Corrupted zip ?");
        }
    },
    /**
     * Change the index.
     * @param {number} newIndex The new index.
     * @throws {Error} if the new index is out of the data.
     */
    setIndex: function(newIndex) {
        this.checkIndex(newIndex);
        this.index = newIndex;
    },
    /**
     * Skip the next n bytes.
     * @param {number} n the number of bytes to skip.
     * @throws {Error} if the new index is out of the data.
     */
    skip: function(n) {
        this.setIndex(this.index + n);
    },
    /**
     * Get the byte at the specified index.
     * @param {number} i the index to use.
     * @return {number} a byte.
     */
    byteAt: function(i) {
        // see implementations
    },
    /**
     * Get the next number with a given byte size.
     * @param {number} size the number of bytes to read.
     * @return {number} the corresponding number.
     */
    readInt: function(size) {
        var result = 0,
            i;
        this.checkOffset(size);
        for (i = this.index + size - 1; i >= this.index; i--) {
            result = (result << 8) + this.byteAt(i);
        }
        this.index += size;
        return result;
    },
    /**
     * Get the next string with a given byte size.
     * @param {number} size the number of bytes to read.
     * @return {string} the corresponding string.
     */
    readString: function(size) {
        return utils.transformTo("string", this.readData(size));
    },
    /**
     * Get raw data without conversion, <size> bytes.
     * @param {number} size the number of bytes to read.
     * @return {Object} the raw data, implementation specific.
     */
    readData: function(size) {
        // see implementations
    },
    /**
     * Find the last occurence of a zip signature (4 bytes).
     * @param {string} sig the signature to find.
     * @return {number} the index of the last occurence, -1 if not found.
     */
    lastIndexOfSignature: function(sig) {
        // see implementations
    },
    /**
     * Get the next date.
     * @return {Date} the date.
     */
    readDate: function() {
        var dostime = this.readInt(4);
        return new Date(
        ((dostime >> 25) & 0x7f) + 1980, // year
        ((dostime >> 21) & 0x0f) - 1, // month
        (dostime >> 16) & 0x1f, // day
        (dostime >> 11) & 0x1f, // hour
        (dostime >> 5) & 0x3f, // minute
        (dostime & 0x1f) << 1); // second
    }
};
module.exports = DataReader;

},{"./utils":21}],6:[function(_dereq_,module,exports){
'use strict';
exports.base64 = false;
exports.binary = false;
exports.dir = false;
exports.createFolders = false;
exports.date = null;
exports.compression = null;
exports.comment = null;

},{}],7:[function(_dereq_,module,exports){
'use strict';
var utils = _dereq_('./utils');

/**
 * @deprecated
 * This function will be removed in a future version without replacement.
 */
exports.string2binary = function(str) {
    return utils.string2binary(str);
};

/**
 * @deprecated
 * This function will be removed in a future version without replacement.
 */
exports.string2Uint8Array = function(str) {
    return utils.transformTo("uint8array", str);
};

/**
 * @deprecated
 * This function will be removed in a future version without replacement.
 */
exports.uint8Array2String = function(array) {
    return utils.transformTo("string", array);
};

/**
 * @deprecated
 * This function will be removed in a future version without replacement.
 */
exports.string2Blob = function(str) {
    var buffer = utils.transformTo("arraybuffer", str);
    return utils.arrayBuffer2Blob(buffer);
};

/**
 * @deprecated
 * This function will be removed in a future version without replacement.
 */
exports.arrayBuffer2Blob = function(buffer) {
    return utils.arrayBuffer2Blob(buffer);
};

/**
 * @deprecated
 * This function will be removed in a future version without replacement.
 */
exports.transformTo = function(outputType, input) {
    return utils.transformTo(outputType, input);
};

/**
 * @deprecated
 * This function will be removed in a future version without replacement.
 */
exports.getTypeOf = function(input) {
    return utils.getTypeOf(input);
};

/**
 * @deprecated
 * This function will be removed in a future version without replacement.
 */
exports.checkSupport = function(type) {
    return utils.checkSupport(type);
};

/**
 * @deprecated
 * This value will be removed in a future version without replacement.
 */
exports.MAX_VALUE_16BITS = utils.MAX_VALUE_16BITS;

/**
 * @deprecated
 * This value will be removed in a future version without replacement.
 */
exports.MAX_VALUE_32BITS = utils.MAX_VALUE_32BITS;


/**
 * @deprecated
 * This function will be removed in a future version without replacement.
 */
exports.pretty = function(str) {
    return utils.pretty(str);
};

/**
 * @deprecated
 * This function will be removed in a future version without replacement.
 */
exports.findCompression = function(compressionMethod) {
    return utils.findCompression(compressionMethod);
};

/**
 * @deprecated
 * This function will be removed in a future version without replacement.
 */
exports.isRegExp = function (object) {
    return utils.isRegExp(object);
};


},{"./utils":21}],8:[function(_dereq_,module,exports){
'use strict';
var USE_TYPEDARRAY = (typeof Uint8Array !== 'undefined') && (typeof Uint16Array !== 'undefined') && (typeof Uint32Array !== 'undefined');

var pako = _dereq_("pako");
exports.uncompressInputType = USE_TYPEDARRAY ? "uint8array" : "array";
exports.compressInputType = USE_TYPEDARRAY ? "uint8array" : "array";

exports.magic = "\x08\x00";
exports.compress = function(input) {
    return pako.deflateRaw(input);
};
exports.uncompress =  function(input) {
    return pako.inflateRaw(input);
};

},{"pako":24}],9:[function(_dereq_,module,exports){
'use strict';

var base64 = _dereq_('./base64');

/**
Usage:
   zip = new JSZip();
   zip.file("hello.txt", "Hello, World!").file("tempfile", "nothing");
   zip.folder("images").file("smile.gif", base64Data, {base64: true});
   zip.file("Xmas.txt", "Ho ho ho !", {date : new Date("December 25, 2007 00:00:01")});
   zip.remove("tempfile");

   base64zip = zip.generate();

**/

/**
 * Representation a of zip file in js
 * @constructor
 * @param {String=|ArrayBuffer=|Uint8Array=} data the data to load, if any (optional).
 * @param {Object=} options the options for creating this objects (optional).
 */
function JSZipSync(data, options) {
    // if this constructor is used without `new`, it adds `new` before itself:
    if(!(this instanceof JSZipSync)) return new JSZipSync(data, options);

    // object containing the files :
    // {
    //   "folder/" : {...},
    //   "folder/data.txt" : {...}
    // }
    this.files = {};

    this.comment = null;

    // Where we are in the hierarchy
    this.root = "";
    if (data) {
        this.load(data, options);
    }
    this.clone = function() {
        var newObj = new JSZipSync();
        for (var i in this) {
            if (typeof this[i] !== "function") {
                newObj[i] = this[i];
            }
        }
        return newObj;
    };
}
JSZipSync.prototype = _dereq_('./object');
JSZipSync.prototype.load = _dereq_('./load');
JSZipSync.support = _dereq_('./support');
JSZipSync.defaults = _dereq_('./defaults');

/**
 * @deprecated
 * This namespace will be removed in a future version without replacement.
 */
JSZipSync.utils = _dereq_('./deprecatedPublicUtils');

JSZipSync.base64 = {
    /**
     * @deprecated
     * This method will be removed in a future version without replacement.
     */
    encode : function(input) {
        return base64.encode(input);
    },
    /**
     * @deprecated
     * This method will be removed in a future version without replacement.
     */
    decode : function(input) {
        return base64.decode(input);
    }
};
JSZipSync.compressions = _dereq_('./compressions');
module.exports = JSZipSync;

},{"./base64":1,"./compressions":3,"./defaults":6,"./deprecatedPublicUtils":7,"./load":10,"./object":13,"./support":17}],10:[function(_dereq_,module,exports){
'use strict';
var base64 = _dereq_('./base64');
var ZipEntries = _dereq_('./zipEntries');
module.exports = function(data, options) {
    var files, zipEntries, i, input;
    options = options || {};
    if (options.base64) {
        data = base64.decode(data);
    }

    zipEntries = new ZipEntries(data, options);
    files = zipEntries.files;
    for (i = 0; i < files.length; i++) {
        input = files[i];
        this.file(input.fileName, input.decompressed, {
            binary: true,
            optimizedBinaryString: true,
            date: input.date,
            dir: input.dir,
            comment : input.fileComment.length ? input.fileComment : null,
            createFolders: options.createFolders
        });
    }
    if (zipEntries.zipComment.length) {
        this.comment = zipEntries.zipComment;
    }

    return this;
};

},{"./base64":1,"./zipEntries":22}],11:[function(_dereq_,module,exports){
(function (Buffer){
'use strict';
var Buffer_from = /*::(*/function(){}/*:: :any)*/;
if(typeof Buffer !== 'undefined') {
	var nbfs = !Buffer.from;
	if(!nbfs) try { Buffer.from("foo", "utf8"); } catch(e) { nbfs = true; }
	Buffer_from = nbfs ? function(buf, enc) { return (enc) ? new Buffer(buf, enc) : new Buffer(buf); } : Buffer.from.bind(Buffer);
	// $FlowIgnore
	if(!Buffer.alloc) Buffer.alloc = function(n) { return new Buffer(n); };
}
module.exports = function(data, encoding){
    return typeof data == 'number' ? Buffer.alloc(data) : Buffer_from(data, encoding);
};
module.exports.test = function(b){
    return Buffer.isBuffer(b);
};
}).call(this,(typeof Buffer !== "undefined" ? Buffer : undefined))
},{}],12:[function(_dereq_,module,exports){
'use strict';
var Uint8ArrayReader = _dereq_('./uint8ArrayReader');

function NodeBufferReader(data) {
    this.data = data;
    this.length = this.data.length;
    this.index = 0;
}
NodeBufferReader.prototype = new Uint8ArrayReader();

/**
 * @see DataReader.readData
 */
NodeBufferReader.prototype.readData = function(size) {
    this.checkOffset(size);
    var result = this.data.slice(this.index, this.index + size);
    this.index += size;
    return result;
};
module.exports = NodeBufferReader;

},{"./uint8ArrayReader":18}],13:[function(_dereq_,module,exports){
'use strict';
var support = _dereq_('./support');
var utils = _dereq_('./utils');
var crc32 = _dereq_('./crc32');
var signature = _dereq_('./signature');
var defaults = _dereq_('./defaults');
var base64 = _dereq_('./base64');
var compressions = _dereq_('./compressions');
var CompressedObject = _dereq_('./compressedObject');
var nodeBuffer = _dereq_('./nodeBuffer');
var utf8 = _dereq_('./utf8');
var StringWriter = _dereq_('./stringWriter');
var Uint8ArrayWriter = _dereq_('./uint8ArrayWriter');

/**
 * Returns the raw data of a ZipObject, decompress the content if necessary.
 * @param {ZipObject} file the file to use.
 * @return {String|ArrayBuffer|Uint8Array|Buffer} the data.
 */
var getRawData = function(file) {
    if (file._data instanceof CompressedObject) {
        file._data = file._data.getContent();
        file.options.binary = true;
        file.options.base64 = false;

        if (utils.getTypeOf(file._data) === "uint8array") {
            var copy = file._data;
            // when reading an arraybuffer, the CompressedObject mechanism will keep it and subarray() a Uint8Array.
            // if we request a file in the same format, we might get the same Uint8Array or its ArrayBuffer (the original zip file).
            file._data = new Uint8Array(copy.length);
            // with an empty Uint8Array, Opera fails with a "Offset larger than array size"
            if (copy.length !== 0) {
                file._data.set(copy, 0);
            }
        }
    }
    return file._data;
};

/**
 * Returns the data of a ZipObject in a binary form. If the content is an unicode string, encode it.
 * @param {ZipObject} file the file to use.
 * @return {String|ArrayBuffer|Uint8Array|Buffer} the data.
 */
var getBinaryData = function(file) {
    var result = getRawData(file),
        type = utils.getTypeOf(result);
    if (type === "string") {
        if (!file.options.binary) {
            // unicode text !
            // unicode string => binary string is a painful process, check if we can avoid it.
            if (support.nodebuffer) {
                return nodeBuffer(result, "utf-8");
            }
        }
        return file.asBinary();
    }
    return result;
};

/**
 * Transform this._data into a string.
 * @param {function} filter a function String -> String, applied if not null on the result.
 * @return {String} the string representing this._data.
 */
var dataToString = function(asUTF8) {
    var result = getRawData(this);
    if (result === null || typeof result === "undefined") {
        return "";
    }
    // if the data is a base64 string, we decode it before checking the encoding !
    if (this.options.base64) {
        result = base64.decode(result);
    }
    if (asUTF8 && this.options.binary) {
        // JSZip.prototype.utf8decode supports arrays as input
        // skip to array => string step, utf8decode will do it.
        result = out.utf8decode(result);
    }
    else {
        // no utf8 transformation, do the array => string step.
        result = utils.transformTo("string", result);
    }

    if (!asUTF8 && !this.options.binary) {
        result = utils.transformTo("string", out.utf8encode(result));
    }
    return result;
};
/**
 * A simple object representing a file in the zip file.
 * @constructor
 * @param {string} name the name of the file
 * @param {String|ArrayBuffer|Uint8Array|Buffer} data the data
 * @param {Object} options the options of the file
 */
var ZipObject = function(name, data, options) {
    this.name = name;
    this.dir = options.dir;
    this.date = options.date;
    this.comment = options.comment;

    this._data = data;
    this.options = options;

    /*
     * This object contains initial values for dir and date.
     * With them, we can check if the user changed the deprecated metadata in
     * `ZipObject#options` or not.
     */
    this._initialMetadata = {
      dir : options.dir,
      date : options.date
    };
};

ZipObject.prototype = {
    /**
     * Return the content as UTF8 string.
     * @return {string} the UTF8 string.
     */
    asText: function() {
        return dataToString.call(this, true);
    },
    /**
     * Returns the binary content.
     * @return {string} the content as binary.
     */
    asBinary: function() {
        return dataToString.call(this, false);
    },
    /**
     * Returns the content as a nodejs Buffer.
     * @return {Buffer} the content as a Buffer.
     */
    asNodeBuffer: function() {
        var result = getBinaryData(this);
        return utils.transformTo("nodebuffer", result);
    },
    /**
     * Returns the content as an Uint8Array.
     * @return {Uint8Array} the content as an Uint8Array.
     */
    asUint8Array: function() {
        var result = getBinaryData(this);
        return utils.transformTo("uint8array", result);
    },
    /**
     * Returns the content as an ArrayBuffer.
     * @return {ArrayBuffer} the content as an ArrayBufer.
     */
    asArrayBuffer: function() {
        return this.asUint8Array().buffer;
    }
};

/**
 * Transform an integer into a string in hexadecimal.
 * @private
 * @param {number} dec the number to convert.
 * @param {number} bytes the number of bytes to generate.
 * @returns {string} the result.
 */
var decToHex = function(dec, bytes) {
    var hex = "",
        i;
    for (i = 0; i < bytes; i++) {
        hex += String.fromCharCode(dec & 0xff);
        dec = dec >>> 8;
    }
    return hex;
};

/**
 * Merge the objects passed as parameters into a new one.
 * @private
 * @param {...Object} var_args All objects to merge.
 * @return {Object} a new object with the data of the others.
 */
var extend = function() {
    var result = {}, i, attr;
    for (i = 0; i < arguments.length; i++) { // arguments is not enumerable in some browsers
        for (attr in arguments[i]) {
            if (arguments[i].hasOwnProperty(attr) && typeof result[attr] === "undefined") {
                result[attr] = arguments[i][attr];
            }
        }
    }
    return result;
};

/**
 * Transforms the (incomplete) options from the user into the complete
 * set of options to create a file.
 * @private
 * @param {Object} o the options from the user.
 * @return {Object} the complete set of options.
 */
var prepareFileAttrs = function(o) {
    o = o || {};
    if (o.base64 === true && (o.binary === null || o.binary === undefined)) {
        o.binary = true;
    }
    o = extend(o, defaults);
    o.date = o.date || new Date();
    if (o.compression !== null) o.compression = o.compression.toUpperCase();

    return o;
};

/**
 * Add a file in the current folder.
 * @private
 * @param {string} name the name of the file
 * @param {String|ArrayBuffer|Uint8Array|Buffer} data the data of the file
 * @param {Object} o the options of the file
 * @return {Object} the new file.
 */
var fileAdd = function(name, data, o) {
    // be sure sub folders exist
    var dataType = utils.getTypeOf(data),
        parent;

    o = prepareFileAttrs(o);

    if (o.createFolders && (parent = parentFolder(name))) {
        folderAdd.call(this, parent, true);
    }

    if (o.dir || data === null || typeof data === "undefined") {
        o.base64 = false;
        o.binary = false;
        data = null;
    }
    else if (dataType === "string") {
        if (o.binary && !o.base64) {
            // optimizedBinaryString == true means that the file has already been filtered with a 0xFF mask
            if (o.optimizedBinaryString !== true) {
                // this is a string, not in a base64 format.
                // Be sure that this is a correct "binary string"
                data = utils.string2binary(data);
            }
        }
    }
    else { // arraybuffer, uint8array, ...
        o.base64 = false;
        o.binary = true;

        if (!dataType && !(data instanceof CompressedObject)) {
            throw new Error("The data of '" + name + "' is in an unsupported format !");
        }

        // special case : it's way easier to work with Uint8Array than with ArrayBuffer
        if (dataType === "arraybuffer") {
            data = utils.transformTo("uint8array", data);
        }
    }

    var object = new ZipObject(name, data, o);
    this.files[name] = object;
    return object;
};

/**
 * Find the parent folder of the path.
 * @private
 * @param {string} path the path to use
 * @return {string} the parent folder, or ""
 */
var parentFolder = function (path) {
    if (path.slice(-1) == '/') {
        path = path.substring(0, path.length - 1);
    }
    var lastSlash = path.lastIndexOf('/');
    return (lastSlash > 0) ? path.substring(0, lastSlash) : "";
};

/**
 * Add a (sub) folder in the current folder.
 * @private
 * @param {string} name the folder's name
 * @param {boolean=} [createFolders] If true, automatically create sub
 *  folders. Defaults to false.
 * @return {Object} the new folder.
 */
var folderAdd = function(name, createFolders) {
    // Check the name ends with a /
    if (name.slice(-1) != "/") {
        name += "/"; // IE doesn't like substr(-1)
    }

    createFolders = (typeof createFolders !== 'undefined') ? createFolders : false;

    // Does this folder already exist?
    if (!this.files[name]) {
        fileAdd.call(this, name, null, {
            dir: true,
            createFolders: createFolders
        });
    }
    return this.files[name];
};

/**
 * Generate a JSZip.CompressedObject for a given zipOject.
 * @param {ZipObject} file the object to read.
 * @param {JSZip.compression} compression the compression to use.
 * @return {JSZip.CompressedObject} the compressed result.
 */
var generateCompressedObjectFrom = function(file, compression) {
    var result = new CompressedObject(),
        content;

    // the data has not been decompressed, we might reuse things !
    if (file._data instanceof CompressedObject) {
        result.uncompressedSize = file._data.uncompressedSize;
        result.crc32 = file._data.crc32;

        if (result.uncompressedSize === 0 || file.dir) {
            compression = compressions['STORE'];
            result.compressedContent = "";
            result.crc32 = 0;
        }
        else if (file._data.compressionMethod === compression.magic) {
            result.compressedContent = file._data.getCompressedContent();
        }
        else {
            content = file._data.getContent();
            // need to decompress / recompress
            result.compressedContent = compression.compress(utils.transformTo(compression.compressInputType, content));
        }
    }
    else {
        // have uncompressed data
        content = getBinaryData(file);
        if (!content || content.length === 0 || file.dir) {
            compression = compressions['STORE'];
            content = "";
        }
        result.uncompressedSize = content.length;
        result.crc32 = crc32(content);
        result.compressedContent = compression.compress(utils.transformTo(compression.compressInputType, content));
    }

    result.compressedSize = result.compressedContent.length;
    result.compressionMethod = compression.magic;

    return result;
};

/**
 * Generate the various parts used in the construction of the final zip file.
 * @param {string} name the file name.
 * @param {ZipObject} file the file content.
 * @param {JSZip.CompressedObject} compressedObject the compressed object.
 * @param {number} offset the current offset from the start of the zip file.
 * @return {object} the zip parts.
 */
var generateZipParts = function(name, file, compressedObject, offset) {
    var data = compressedObject.compressedContent,
        utfEncodedFileName = utils.transformTo("string", utf8.utf8encode(file.name)),
        comment = file.comment || "",
        utfEncodedComment = utils.transformTo("string", utf8.utf8encode(comment)),
        useUTF8ForFileName = utfEncodedFileName.length !== file.name.length,
        useUTF8ForComment = utfEncodedComment.length !== comment.length,
        o = file.options,
        dosTime,
        dosDate,
        extraFields = "",
        unicodePathExtraField = "",
        unicodeCommentExtraField = "",
        dir, date;


    // handle the deprecated options.dir
    if (file._initialMetadata.dir !== file.dir) {
        dir = file.dir;
    } else {
        dir = o.dir;
    }

    // handle the deprecated options.date
    if(file._initialMetadata.date !== file.date) {
        date = file.date;
    } else {
        date = o.date;
    }


    dosTime = date.getHours();
    dosTime = dosTime << 6;
    dosTime = dosTime | date.getMinutes();
    dosTime = dosTime << 5;
    dosTime = dosTime | date.getSeconds() / 2;

    dosDate = date.getFullYear() - 1980;
    dosDate = dosDate << 4;
    dosDate = dosDate | (date.getMonth() + 1);
    dosDate = dosDate << 5;
    dosDate = dosDate | date.getDate();

    if (useUTF8ForFileName) {
        // set the unicode path extra field. unzip needs at least one extra
        // field to correctly handle unicode path, so using the path is as good
        // as any other information. This could improve the situation with
        // other archive managers too.
        // This field is usually used without the utf8 flag, with a non
        // unicode path in the header (winrar, winzip). This helps (a bit)
        // with the messy Windows' default compressed folders feature but
        // breaks on p7zip which doesn't seek the unicode path extra field.
        // So for now, UTF-8 everywhere !
        unicodePathExtraField =
            // Version
            decToHex(1, 1) +
            // NameCRC32
            decToHex(crc32(utfEncodedFileName), 4) +
            // UnicodeName
            utfEncodedFileName;

        extraFields +=
            // Info-ZIP Unicode Path Extra Field
            "\x75\x70" +
            // size
            decToHex(unicodePathExtraField.length, 2) +
            // content
            unicodePathExtraField;
    }

    if(useUTF8ForComment) {

        unicodeCommentExtraField =
            // Version
            decToHex(1, 1) +
            // CommentCRC32
            decToHex(this.crc32(utfEncodedComment), 4) +
            // UnicodeName
            utfEncodedComment;

        extraFields +=
            // Info-ZIP Unicode Path Extra Field
            "\x75\x63" +
            // size
            decToHex(unicodeCommentExtraField.length, 2) +
            // content
            unicodeCommentExtraField;
    }

    var header = "";

    // version needed to extract
    header += "\x0A\x00";
    // general purpose bit flag
    // set bit 11 if utf8
    header += (useUTF8ForFileName || useUTF8ForComment) ? "\x00\x08" : "\x00\x00";
    // compression method
    header += compressedObject.compressionMethod;
    // last mod file time
    header += decToHex(dosTime, 2);
    // last mod file date
    header += decToHex(dosDate, 2);
    // crc-32
    header += decToHex(compressedObject.crc32, 4);
    // compressed size
    header += decToHex(compressedObject.compressedSize, 4);
    // uncompressed size
    header += decToHex(compressedObject.uncompressedSize, 4);
    // file name length
    header += decToHex(utfEncodedFileName.length, 2);
    // extra field length
    header += decToHex(extraFields.length, 2);


    var fileRecord = signature.LOCAL_FILE_HEADER + header + utfEncodedFileName + extraFields;

    var dirRecord = signature.CENTRAL_FILE_HEADER +
    // version made by (00: DOS)
    "\x14\x00" +
    // file header (common to file and central directory)
    header +
    // file comment length
    decToHex(utfEncodedComment.length, 2) +
    // disk number start
    "\x00\x00" +
    // internal file attributes TODO
    "\x00\x00" +
    // external file attributes
    (dir === true ? "\x10\x00\x00\x00" : "\x00\x00\x00\x00") +
    // relative offset of local header
    decToHex(offset, 4) +
    // file name
    utfEncodedFileName +
    // extra field
    extraFields +
    // file comment
    utfEncodedComment;

    return {
        fileRecord: fileRecord,
        dirRecord: dirRecord,
        compressedObject: compressedObject
    };
};


// return the actual prototype of JSZip
var out = {
    /**
     * Read an existing zip and merge the data in the current JSZip object.
     * The implementation is in jszip-load.js, don't forget to include it.
     * @param {String|ArrayBuffer|Uint8Array|Buffer} stream  The stream to load
     * @param {Object} options Options for loading the stream.
     *  options.base64 : is the stream in base64 ? default : false
     * @return {JSZip} the current JSZip object
     */
    load: function(stream, options) {
        throw new Error("Load method is not defined. Is the file jszip-load.js included ?");
    },

    /**
     * Filter nested files/folders with the specified function.
     * @param {Function} search the predicate to use :
     * function (relativePath, file) {...}
     * It takes 2 arguments : the relative path and the file.
     * @return {Array} An array of matching elements.
     */
    filter: function(search) {
        var result = [],
            filename, relativePath, file, fileClone;
        for (filename in this.files) {
            if (!this.files.hasOwnProperty(filename)) {
                continue;
            }
            file = this.files[filename];
            // return a new object, don't let the user mess with our internal objects :)
            fileClone = new ZipObject(file.name, file._data, extend(file.options));
            relativePath = filename.slice(this.root.length, filename.length);
            if (filename.slice(0, this.root.length) === this.root && // the file is in the current root
            search(relativePath, fileClone)) { // and the file matches the function
                result.push(fileClone);
            }
        }
        return result;
    },

    /**
     * Add a file to the zip file, or search a file.
     * @param   {string|RegExp} name The name of the file to add (if data is defined),
     * the name of the file to find (if no data) or a regex to match files.
     * @param   {String|ArrayBuffer|Uint8Array|Buffer} data  The file data, either raw or base64 encoded
     * @param   {Object} o     File options
     * @return  {JSZip|Object|Array} this JSZip object (when adding a file),
     * a file (when searching by string) or an array of files (when searching by regex).
     */
    file: function(name, data, o) {
        if (arguments.length === 1) {
            if (utils.isRegExp(name)) {
                var regexp = name;
                return this.filter(function(relativePath, file) {
                    return !file.dir && regexp.test(relativePath);
                });
            }
            else { // text
                return this.filter(function(relativePath, file) {
                    return !file.dir && relativePath === name;
                })[0] || null;
            }
        }
        else { // more than one argument : we have data !
            name = this.root + name;
            fileAdd.call(this, name, data, o);
        }
        return this;
    },

    /**
     * Add a directory to the zip file, or search.
     * @param   {String|RegExp} arg The name of the directory to add, or a regex to search folders.
     * @return  {JSZip} an object with the new directory as the root, or an array containing matching folders.
     */
    folder: function(arg) {
        if (!arg) {
            return this;
        }

        if (utils.isRegExp(arg)) {
            return this.filter(function(relativePath, file) {
                return file.dir && arg.test(relativePath);
            });
        }

        // else, name is a new folder
        var name = this.root + arg;
        var newFolder = folderAdd.call(this, name);

        // Allow chaining by returning a new object with this folder as the root
        var ret = this.clone();
        ret.root = newFolder.name;
        return ret;
    },

    /**
     * Delete a file, or a directory and all sub-files, from the zip
     * @param {string} name the name of the file to delete
     * @return {JSZip} this JSZip object
     */
    remove: function(name) {
        name = this.root + name;
        var file = this.files[name];
        if (!file) {
            // Look for any folders
            if (name.slice(-1) != "/") {
                name += "/";
            }
            file = this.files[name];
        }

        if (file && !file.dir) {
            // file
            delete this.files[name];
        } else {
            // maybe a folder, delete recursively
            var kids = this.filter(function(relativePath, file) {
                return file.name.slice(0, name.length) === name;
            });
            for (var i = 0; i < kids.length; i++) {
                delete this.files[kids[i].name];
            }
        }

        return this;
    },

    /**
     * Generate the complete zip file
     * @param {Object} options the options to generate the zip file :
     * - base64, (deprecated, use type instead) true to generate base64.
     * - compression, "STORE" by default.
     * - type, "base64" by default. Values are : string, base64, uint8array, arraybuffer, blob.
     * @return {String|Uint8Array|ArrayBuffer|Buffer|Blob} the zip file
     */
    generate: function(options) {
        options = extend(options || {}, {
            base64: true,
            compression: "STORE",
            type: "base64",
            comment: null
        });

        utils.checkSupport(options.type);

        var zipData = [],
            localDirLength = 0,
            centralDirLength = 0,
            writer, i,
            utfEncodedComment = utils.transformTo("string", this.utf8encode(options.comment || this.comment || ""));

        // first, generate all the zip parts.
        for (var name in this.files) {
            if (!this.files.hasOwnProperty(name)) {
                continue;
            }
            var file = this.files[name];

            var compressionName = file.options.compression || options.compression.toUpperCase();
            var compression = compressions[compressionName];
            if (!compression) {
                throw new Error(compressionName + " is not a valid compression method !");
            }

            var compressedObject = generateCompressedObjectFrom.call(this, file, compression);

            var zipPart = generateZipParts.call(this, name, file, compressedObject, localDirLength);
            localDirLength += zipPart.fileRecord.length + compressedObject.compressedSize;
            centralDirLength += zipPart.dirRecord.length;
            zipData.push(zipPart);
        }

        var dirEnd = "";

        // end of central dir signature
        dirEnd = signature.CENTRAL_DIRECTORY_END +
        // number of this disk
        "\x00\x00" +
        // number of the disk with the start of the central directory
        "\x00\x00" +
        // total number of entries in the central directory on this disk
        decToHex(zipData.length, 2) +
        // total number of entries in the central directory
        decToHex(zipData.length, 2) +
        // size of the central directory   4 bytes
        decToHex(centralDirLength, 4) +
        // offset of start of central directory with respect to the starting disk number
        decToHex(localDirLength, 4) +
        // .ZIP file comment length
        decToHex(utfEncodedComment.length, 2) +
        // .ZIP file comment
        utfEncodedComment;


        // we have all the parts (and the total length)
        // time to create a writer !
        var typeName = options.type.toLowerCase();
        if(typeName==="uint8array"||typeName==="arraybuffer"||typeName==="blob"||typeName==="nodebuffer") {
            writer = new Uint8ArrayWriter(localDirLength + centralDirLength + dirEnd.length);
        }else{
            writer = new StringWriter(localDirLength + centralDirLength + dirEnd.length);
        }

        for (i = 0; i < zipData.length; i++) {
            writer.append(zipData[i].fileRecord);
            writer.append(zipData[i].compressedObject.compressedContent);
        }
        for (i = 0; i < zipData.length; i++) {
            writer.append(zipData[i].dirRecord);
        }

        writer.append(dirEnd);

        var zip = writer.finalize();



        switch(options.type.toLowerCase()) {
            // case "zip is an Uint8Array"
            case "uint8array" :
            case "arraybuffer" :
            case "nodebuffer" :
               return utils.transformTo(options.type.toLowerCase(), zip);
            case "blob" :
               return utils.arrayBuffer2Blob(utils.transformTo("arraybuffer", zip));
            // case "zip is a string"
            case "base64" :
               return (options.base64) ? base64.encode(zip) : zip;
            default : // case "string" :
               return zip;
         }

    },

    /**
     * @deprecated
     * This method will be removed in a future version without replacement.
     */
    crc32: function (input, crc) {
        return crc32(input, crc);
    },

    /**
     * @deprecated
     * This method will be removed in a future version without replacement.
     */
    utf8encode: function (string) {
        return utils.transformTo("string", utf8.utf8encode(string));
    },

    /**
     * @deprecated
     * This method will be removed in a future version without replacement.
     */
    utf8decode: function (input) {
        return utf8.utf8decode(input);
    }
};
module.exports = out;

},{"./base64":1,"./compressedObject":2,"./compressions":3,"./crc32":4,"./defaults":6,"./nodeBuffer":11,"./signature":14,"./stringWriter":16,"./support":17,"./uint8ArrayWriter":19,"./utf8":20,"./utils":21}],14:[function(_dereq_,module,exports){
'use strict';
exports.LOCAL_FILE_HEADER = "PK\x03\x04";
exports.CENTRAL_FILE_HEADER = "PK\x01\x02";
exports.CENTRAL_DIRECTORY_END = "PK\x05\x06";
exports.ZIP64_CENTRAL_DIRECTORY_LOCATOR = "PK\x06\x07";
exports.ZIP64_CENTRAL_DIRECTORY_END = "PK\x06\x06";
exports.DATA_DESCRIPTOR = "PK\x07\x08";

},{}],15:[function(_dereq_,module,exports){
'use strict';
var DataReader = _dereq_('./dataReader');
var utils = _dereq_('./utils');

function StringReader(data, optimizedBinaryString) {
    this.data = data;
    if (!optimizedBinaryString) {
        this.data = utils.string2binary(this.data);
    }
    this.length = this.data.length;
    this.index = 0;
}
StringReader.prototype = new DataReader();
/**
 * @see DataReader.byteAt
 */
StringReader.prototype.byteAt = function(i) {
    return this.data.charCodeAt(i);
};
/**
 * @see DataReader.lastIndexOfSignature
 */
StringReader.prototype.lastIndexOfSignature = function(sig) {
    return this.data.lastIndexOf(sig);
};
/**
 * @see DataReader.readData
 */
StringReader.prototype.readData = function(size) {
    this.checkOffset(size);
    // this will work because the constructor applied the "& 0xff" mask.
    var result = this.data.slice(this.index, this.index + size);
    this.index += size;
    return result;
};
module.exports = StringReader;

},{"./dataReader":5,"./utils":21}],16:[function(_dereq_,module,exports){
'use strict';

var utils = _dereq_('./utils');

/**
 * An object to write any content to a string.
 * @constructor
 */
var StringWriter = function() {
    this.data = [];
};
StringWriter.prototype = {
    /**
     * Append any content to the current string.
     * @param {Object} input the content to add.
     */
    append: function(input) {
        input = utils.transformTo("string", input);
        this.data.push(input);
    },
    /**
     * Finalize the construction an return the result.
     * @return {string} the generated string.
     */
    finalize: function() {
        return this.data.join("");
    }
};

module.exports = StringWriter;

},{"./utils":21}],17:[function(_dereq_,module,exports){
(function (Buffer){
'use strict';
exports.base64 = true;
exports.array = true;
exports.string = true;
exports.arraybuffer = typeof ArrayBuffer !== "undefined" && typeof Uint8Array !== "undefined";
// contains true if JSZip can read/generate nodejs Buffer, false otherwise.
// Browserify will provide a Buffer implementation for browsers, which is
// an augmented Uint8Array (i.e., can be used as either Buffer or U8).
exports.nodebuffer = typeof Buffer !== "undefined";
// contains true if JSZip can read/generate Uint8Array, false otherwise.
exports.uint8array = typeof Uint8Array !== "undefined";

if (typeof ArrayBuffer === "undefined") {
    exports.blob = false;
}
else {
    var buffer = new ArrayBuffer(0);
    try {
        exports.blob = new Blob([buffer], {
            type: "application/zip"
        }).size === 0;
    }
    catch (e) {
        try {
            var Builder = window.BlobBuilder || window.WebKitBlobBuilder || window.MozBlobBuilder || window.MSBlobBuilder;
            var builder = new Builder();
            builder.append(buffer);
            exports.blob = builder.getBlob('application/zip').size === 0;
        }
        catch (e) {
            exports.blob = false;
        }
    }
}

}).call(this,(typeof Buffer !== "undefined" ? Buffer : undefined))
},{}],18:[function(_dereq_,module,exports){
'use strict';
var DataReader = _dereq_('./dataReader');

function Uint8ArrayReader(data) {
    if (data) {
        this.data = data;
        this.length = this.data.length;
        this.index = 0;
    }
}
Uint8ArrayReader.prototype = new DataReader();
/**
 * @see DataReader.byteAt
 */
Uint8ArrayReader.prototype.byteAt = function(i) {
    return this.data[i];
};
/**
 * @see DataReader.lastIndexOfSignature
 */
Uint8ArrayReader.prototype.lastIndexOfSignature = function(sig) {
    var sig0 = sig.charCodeAt(0),
        sig1 = sig.charCodeAt(1),
        sig2 = sig.charCodeAt(2),
        sig3 = sig.charCodeAt(3);
    for (var i = this.length - 4; i >= 0; --i) {
        if (this.data[i] === sig0 && this.data[i + 1] === sig1 && this.data[i + 2] === sig2 && this.data[i + 3] === sig3) {
            return i;
        }
    }

    return -1;
};
/**
 * @see DataReader.readData
 */
Uint8ArrayReader.prototype.readData = function(size) {
    this.checkOffset(size);
    if(size === 0) {
        // in IE10, when using subarray(idx, idx), we get the array [0x00] instead of [].
        return new Uint8Array(0);
    }
    var result = this.data.subarray(this.index, this.index + size);
    this.index += size;
    return result;
};
module.exports = Uint8ArrayReader;

},{"./dataReader":5}],19:[function(_dereq_,module,exports){
'use strict';

var utils = _dereq_('./utils');

/**
 * An object to write any content to an Uint8Array.
 * @constructor
 * @param {number} length The length of the array.
 */
var Uint8ArrayWriter = function(length) {
    this.data = new Uint8Array(length);
    this.index = 0;
};
Uint8ArrayWriter.prototype = {
    /**
     * Append any content to the current array.
     * @param {Object} input the content to add.
     */
    append: function(input) {
        if (input.length !== 0) {
            // with an empty Uint8Array, Opera fails with a "Offset larger than array size"
            input = utils.transformTo("uint8array", input);
            this.data.set(input, this.index);
            this.index += input.length;
        }
    },
    /**
     * Finalize the construction an return the result.
     * @return {Uint8Array} the generated array.
     */
    finalize: function() {
        return this.data;
    }
};

module.exports = Uint8ArrayWriter;

},{"./utils":21}],20:[function(_dereq_,module,exports){
'use strict';

var utils = _dereq_('./utils');
var support = _dereq_('./support');
var nodeBuffer = _dereq_('./nodeBuffer');

/**
 * The following functions come from pako, from pako/lib/utils/strings
 * released under the MIT license, see pako https://github.com/nodeca/pako/
 */

// Table with utf8 lengths (calculated by first byte of sequence)
// Note, that 5 & 6-byte values and some 4-byte values can not be represented in JS,
// because max possible codepoint is 0x10ffff
var _utf8len = new Array(256);
for (var i=0; i<256; i++) {
  _utf8len[i] = (i >= 252 ? 6 : i >= 248 ? 5 : i >= 240 ? 4 : i >= 224 ? 3 : i >= 192 ? 2 : 1);
}
_utf8len[254]=_utf8len[254]=1; // Invalid sequence start

// convert string to array (typed, when possible)
var string2buf = function (str) {
    var buf, c, c2, m_pos, i, str_len = str.length, buf_len = 0;

    // count binary size
    for (m_pos = 0; m_pos < str_len; m_pos++) {
        c = str.charCodeAt(m_pos);
        if (((c & 0xfc00) === 0xd800) && (m_pos+1 < str_len)) {
            c2 = str.charCodeAt(m_pos+1);
            if ((c2 & 0xfc00) === 0xdc00) {
                c = 0x10000 + ((c - 0xd800) << 10) + (c2 - 0xdc00);
                m_pos++;
            }
        }
        buf_len += (c < 0x80) ? 1 : ((c < 0x800) ? 2 : ((c < 0x10000) ? 3 : 4));
    }

    // allocate buffer
    if (support.uint8array) {
        buf = new Uint8Array(buf_len);
    } else {
        buf = new Array(buf_len);
    }

    // convert
    for (i=0, m_pos = 0; i < buf_len; m_pos++) {
        c = str.charCodeAt(m_pos);
        if ((c & 0xfc00) === 0xd800 && (m_pos+1 < str_len)) {
            c2 = str.charCodeAt(m_pos+1);
            if ((c2 & 0xfc00) === 0xdc00) {
                c = 0x10000 + ((c - 0xd800) << 10) + (c2 - 0xdc00);
                m_pos++;
            }
        }
        if (c < 0x80) {
            /* one byte */
            buf[i++] = c;
        } else if (c < 0x800) {
            /* two bytes */
            buf[i++] = 0xC0 | (c >>> 6);
            buf[i++] = 0x80 | (c & 0x3f);
        } else if (c < 0x10000) {
            /* three bytes */
            buf[i++] = 0xE0 | (c >>> 12);
            buf[i++] = 0x80 | ((c >>> 6) & 0x3f);
            buf[i++] = 0x80 | (c & 0x3f);
        } else {
            /* four bytes */
            buf[i++] = 0xf0 | (c >>> 18);
            buf[i++] = 0x80 | ((c >>> 12) & 0x3f);
            buf[i++] = 0x80 | ((c >>> 6) & 0x3f);
            buf[i++] = 0x80 | (c & 0x3f);
        }
    }

    return buf;
};

// Calculate max possible position in utf8 buffer,
// that will not break sequence. If that's not possible
// - (very small limits) return max size as is.
//
// buf[] - utf8 bytes array
// max   - length limit (mandatory);
var utf8border = function(buf, max) {
    var pos;

    max = max || buf.length;
    if (max > buf.length) { max = buf.length; }

    // go back from last position, until start of sequence found
    pos = max-1;
    while (pos >= 0 && (buf[pos] & 0xC0) === 0x80) { pos--; }

    // Fuckup - very small and broken sequence,
    // return max, because we should return something anyway.
    if (pos < 0) { return max; }

    // If we came to start of buffer - that means vuffer is too small,
    // return max too.
    if (pos === 0) { return max; }

    return (pos + _utf8len[buf[pos]] > max) ? pos : max;
};

// convert array to string
var buf2string = function (buf) {
    var str, i, out, c, c_len;
    var len = buf.length;

    // Reserve max possible length (2 words per char)
    // NB: by unknown reasons, Array is significantly faster for
    //     String.fromCharCode.apply than Uint16Array.
    var utf16buf = new Array(len*2);

    for (out=0, i=0; i<len;) {
        c = buf[i++];
        // quick process ascii
        if (c < 0x80) { utf16buf[out++] = c; continue; }

        c_len = _utf8len[c];
        // skip 5 & 6 byte codes
        if (c_len > 4) { utf16buf[out++] = 0xfffd; i += c_len-1; continue; }

        // apply mask on first byte
        c &= c_len === 2 ? 0x1f : c_len === 3 ? 0x0f : 0x07;
        // join the rest
        while (c_len > 1 && i < len) {
            c = (c << 6) | (buf[i++] & 0x3f);
            c_len--;
        }

        // terminated by end of string?
        if (c_len > 1) { utf16buf[out++] = 0xfffd; continue; }

        if (c < 0x10000) {
            utf16buf[out++] = c;
        } else {
            c -= 0x10000;
            utf16buf[out++] = 0xd800 | ((c >> 10) & 0x3ff);
            utf16buf[out++] = 0xdc00 | (c & 0x3ff);
        }
    }

    // shrinkBuf(utf16buf, out)
    if (utf16buf.length !== out) {
        if(utf16buf.subarray) {
            utf16buf = utf16buf.subarray(0, out);
        } else {
            utf16buf.length = out;
        }
    }

    // return String.fromCharCode.apply(null, utf16buf);
    return utils.applyFromCharCode(utf16buf);
};


// That's all for the pako functions.


/**
 * Transform a javascript string into an array (typed if possible) of bytes,
 * UTF-8 encoded.
 * @param {String} str the string to encode
 * @return {Array|Uint8Array|Buffer} the UTF-8 encoded string.
 */
exports.utf8encode = function utf8encode(str) {
    if (support.nodebuffer) {
        return nodeBuffer(str, "utf-8");
    }

    return string2buf(str);
};


/**
 * Transform a bytes array (or a representation) representing an UTF-8 encoded
 * string into a javascript string.
 * @param {Array|Uint8Array|Buffer} buf the data de decode
 * @return {String} the decoded string.
 */
exports.utf8decode = function utf8decode(buf) {
    if (support.nodebuffer) {
        return utils.transformTo("nodebuffer", buf).toString("utf-8");
    }

    buf = utils.transformTo(support.uint8array ? "uint8array" : "array", buf);

    // return buf2string(buf);
    // Chrome prefers to work with "small" chunks of data
    // for the method buf2string.
    // Firefox and Chrome has their own shortcut, IE doesn't seem to really care.
    var result = [], k = 0, len = buf.length, chunk = 65536;
    while (k < len) {
        var nextBoundary = utf8border(buf, Math.min(k + chunk, len));
        if (support.uint8array) {
            result.push(buf2string(buf.subarray(k, nextBoundary)));
        } else {
            result.push(buf2string(buf.slice(k, nextBoundary)));
        }
        k = nextBoundary;
    }
    return result.join("");

};
// vim: set shiftwidth=4 softtabstop=4:

},{"./nodeBuffer":11,"./support":17,"./utils":21}],21:[function(_dereq_,module,exports){
'use strict';
var support = _dereq_('./support');
var compressions = _dereq_('./compressions');
var nodeBuffer = _dereq_('./nodeBuffer');
/**
 * Convert a string to a "binary string" : a string containing only char codes between 0 and 255.
 * @param {string} str the string to transform.
 * @return {String} the binary string.
 */
exports.string2binary = function(str) {
    var result = "";
    for (var i = 0; i < str.length; i++) {
        result += String.fromCharCode(str.charCodeAt(i) & 0xff);
    }
    return result;
};
exports.arrayBuffer2Blob = function(buffer) {
    exports.checkSupport("blob");

    try {
        // Blob constructor
        return new Blob([buffer], {
            type: "application/zip"
        });
    }
    catch (e) {

        try {
            // deprecated, browser only, old way
            var Builder = window.BlobBuilder || window.WebKitBlobBuilder || window.MozBlobBuilder || window.MSBlobBuilder;
            var builder = new Builder();
            builder.append(buffer);
            return builder.getBlob('application/zip');
        }
        catch (e) {

            // well, fuck ?!
            throw new Error("Bug : can't construct the Blob.");
        }
    }


};
/**
 * The identity function.
 * @param {Object} input the input.
 * @return {Object} the same input.
 */
function identity(input) {
    return input;
}

/**
 * Fill in an array with a string.
 * @param {String} str the string to use.
 * @param {Array|ArrayBuffer|Uint8Array|Buffer} array the array to fill in (will be mutated).
 * @return {Array|ArrayBuffer|Uint8Array|Buffer} the updated array.
 */
function stringToArrayLike(str, array) {
    for (var i = 0; i < str.length; ++i) {
        array[i] = str.charCodeAt(i) & 0xFF;
    }
    return array;
}

/**
 * Transform an array-like object to a string.
 * @param {Array|ArrayBuffer|Uint8Array|Buffer} array the array to transform.
 * @return {String} the result.
 */
function arrayLikeToString(array) {
    // Performances notes :
    // --------------------
    // String.fromCharCode.apply(null, array) is the fastest, see
    // see http://jsperf.com/converting-a-uint8array-to-a-string/2
    // but the stack is limited (and we can get huge arrays !).
    //
    // result += String.fromCharCode(array[i]); generate too many strings !
    //
    // This code is inspired by http://jsperf.com/arraybuffer-to-string-apply-performance/2
    var chunk = 65536;
    var result = [],
        len = array.length,
        type = exports.getTypeOf(array),
        k = 0,
        canUseApply = true;
      try {
         switch(type) {
            case "uint8array":
               String.fromCharCode.apply(null, new Uint8Array(0));
               break;
            case "nodebuffer":
               String.fromCharCode.apply(null, nodeBuffer(0));
               break;
         }
      } catch(e) {
         canUseApply = false;
      }

      // no apply : slow and painful algorithm
      // default browser on android 4.*
      if (!canUseApply) {
         var resultStr = "";
         for(var i = 0; i < array.length;i++) {
            resultStr += String.fromCharCode(array[i]);
         }
    return resultStr;
    }
    while (k < len && chunk > 1) {
        try {
            if (type === "array" || type === "nodebuffer") {
                result.push(String.fromCharCode.apply(null, array.slice(k, Math.min(k + chunk, len))));
            }
            else {
                result.push(String.fromCharCode.apply(null, array.subarray(k, Math.min(k + chunk, len))));
            }
            k += chunk;
        }
        catch (e) {
            chunk = Math.floor(chunk / 2);
        }
    }
    return result.join("");
}

exports.applyFromCharCode = arrayLikeToString;


/**
 * Copy the data from an array-like to an other array-like.
 * @param {Array|ArrayBuffer|Uint8Array|Buffer} arrayFrom the origin array.
 * @param {Array|ArrayBuffer|Uint8Array|Buffer} arrayTo the destination array which will be mutated.
 * @return {Array|ArrayBuffer|Uint8Array|Buffer} the updated destination array.
 */
function arrayLikeToArrayLike(arrayFrom, arrayTo) {
    for (var i = 0; i < arrayFrom.length; i++) {
        arrayTo[i] = arrayFrom[i];
    }
    return arrayTo;
}

// a matrix containing functions to transform everything into everything.
var transform = {};

// string to ?
transform["string"] = {
    "string": identity,
    "array": function(input) {
        return stringToArrayLike(input, new Array(input.length));
    },
    "arraybuffer": function(input) {
        return transform["string"]["uint8array"](input).buffer;
    },
    "uint8array": function(input) {
        return stringToArrayLike(input, new Uint8Array(input.length));
    },
    "nodebuffer": function(input) {
        return stringToArrayLike(input, nodeBuffer(input.length));
    }
};

// array to ?
transform["array"] = {
    "string": arrayLikeToString,
    "array": identity,
    "arraybuffer": function(input) {
        return (new Uint8Array(input)).buffer;
    },
    "uint8array": function(input) {
        return new Uint8Array(input);
    },
    "nodebuffer": function(input) {
        return nodeBuffer(input);
    }
};

// arraybuffer to ?
transform["arraybuffer"] = {
    "string": function(input) {
        return arrayLikeToString(new Uint8Array(input));
    },
    "array": function(input) {
        return arrayLikeToArrayLike(new Uint8Array(input), new Array(input.byteLength));
    },
    "arraybuffer": identity,
    "uint8array": function(input) {
        return new Uint8Array(input);
    },
    "nodebuffer": function(input) {
        return nodeBuffer(new Uint8Array(input));
    }
};

// uint8array to ?
transform["uint8array"] = {
    "string": arrayLikeToString,
    "array": function(input) {
        return arrayLikeToArrayLike(input, new Array(input.length));
    },
    "arraybuffer": function(input) {
        return input.buffer;
    },
    "uint8array": identity,
    "nodebuffer": function(input) {
        return nodeBuffer(input);
    }
};

// nodebuffer to ?
transform["nodebuffer"] = {
    "string": arrayLikeToString,
    "array": function(input) {
        return arrayLikeToArrayLike(input, new Array(input.length));
    },
    "arraybuffer": function(input) {
        return transform["nodebuffer"]["uint8array"](input).buffer;
    },
    "uint8array": function(input) {
        return arrayLikeToArrayLike(input, new Uint8Array(input.length));
    },
    "nodebuffer": identity
};

/**
 * Transform an input into any type.
 * The supported output type are : string, array, uint8array, arraybuffer, nodebuffer.
 * If no output type is specified, the unmodified input will be returned.
 * @param {String} outputType the output type.
 * @param {String|Array|ArrayBuffer|Uint8Array|Buffer} input the input to convert.
 * @throws {Error} an Error if the browser doesn't support the requested output type.
 */
exports.transformTo = function(outputType, input) {
    if (!input) {
        // undefined, null, etc
        // an empty string won't harm.
        input = "";
    }
    if (!outputType) {
        return input;
    }
    exports.checkSupport(outputType);
    var inputType = exports.getTypeOf(input);
    var result = transform[inputType][outputType](input);
    return result;
};

/**
 * Return the type of the input.
 * The type will be in a format valid for JSZip.utils.transformTo : string, array, uint8array, arraybuffer.
 * @param {Object} input the input to identify.
 * @return {String} the (lowercase) type of the input.
 */
exports.getTypeOf = function(input) {
    if (typeof input === "string") {
        return "string";
    }
    if (Object.prototype.toString.call(input) === "[object Array]") {
        return "array";
    }
    if (support.nodebuffer && nodeBuffer.test(input)) {
        return "nodebuffer";
    }
    if (support.uint8array && input instanceof Uint8Array) {
        return "uint8array";
    }
    if (support.arraybuffer && input instanceof ArrayBuffer) {
        return "arraybuffer";
    }
};

/**
 * Throw an exception if the type is not supported.
 * @param {String} type the type to check.
 * @throws {Error} an Error if the browser doesn't support the requested type.
 */
exports.checkSupport = function(type) {
    var supported = support[type.toLowerCase()];
    if (!supported) {
        throw new Error(type + " is not supported by this browser");
    }
};
exports.MAX_VALUE_16BITS = 65535;
exports.MAX_VALUE_32BITS = -1; // well, "\xFF\xFF\xFF\xFF\xFF\xFF\xFF\xFF" is parsed as -1

/**
 * Prettify a string read as binary.
 * @param {string} str the string to prettify.
 * @return {string} a pretty string.
 */
exports.pretty = function(str) {
    var res = '',
        code, i;
    for (i = 0; i < (str || "").length; i++) {
        code = str.charCodeAt(i);
        res += '\\x' + (code < 16 ? "0" : "") + code.toString(16).toUpperCase();
    }
    return res;
};

/**
 * Find a compression registered in JSZip.
 * @param {string} compressionMethod the method magic to find.
 * @return {Object|null} the JSZip compression object, null if none found.
 */
exports.findCompression = function(compressionMethod) {
    for (var method in compressions) {
        if (!compressions.hasOwnProperty(method)) {
            continue;
        }
        if (compressions[method].magic === compressionMethod) {
            return compressions[method];
        }
    }
    return null;
};
/**
* Cross-window, cross-Node-context regular expression detection
* @param  {Object}  object Anything
* @return {Boolean}        true if the object is a regular expression,
* false otherwise
*/
exports.isRegExp = function (object) {
    return Object.prototype.toString.call(object) === "[object RegExp]";
};


},{"./compressions":3,"./nodeBuffer":11,"./support":17}],22:[function(_dereq_,module,exports){
'use strict';
var StringReader = _dereq_('./stringReader');
var NodeBufferReader = _dereq_('./nodeBufferReader');
var Uint8ArrayReader = _dereq_('./uint8ArrayReader');
var utils = _dereq_('./utils');
var sig = _dereq_('./signature');
var ZipEntry = _dereq_('./zipEntry');
var support = _dereq_('./support');
var jszipProto = _dereq_('./object');
//  class ZipEntries {{{
/**
 * All the entries in the zip file.
 * @constructor
 * @param {String|ArrayBuffer|Uint8Array} data the binary stream to load.
 * @param {Object} loadOptions Options for loading the stream.
 */
function ZipEntries(data, loadOptions) {
    this.files = [];
    this.loadOptions = loadOptions;
    if (data) {
        this.load(data);
    }
}
ZipEntries.prototype = {
    /**
     * Check that the reader is on the speficied signature.
     * @param {string} expectedSignature the expected signature.
     * @throws {Error} if it is an other signature.
     */
    checkSignature: function(expectedSignature) {
        var signature = this.reader.readString(4);
        if (signature !== expectedSignature) {
            throw new Error("Corrupted zip or bug : unexpected signature " + "(" + utils.pretty(signature) + ", expected " + utils.pretty(expectedSignature) + ")");
        }
    },
    /**
     * Read the end of the central directory.
     */
    readBlockEndOfCentral: function() {
        this.diskNumber = this.reader.readInt(2);
        this.diskWithCentralDirStart = this.reader.readInt(2);
        this.centralDirRecordsOnThisDisk = this.reader.readInt(2);
        this.centralDirRecords = this.reader.readInt(2);
        this.centralDirSize = this.reader.readInt(4);
        this.centralDirOffset = this.reader.readInt(4);

        this.zipCommentLength = this.reader.readInt(2);
        // warning : the encoding depends of the system locale
        // On a linux machine with LANG=en_US.utf8, this field is utf8 encoded.
        // On a windows machine, this field is encoded with the localized windows code page.
        this.zipComment = this.reader.readString(this.zipCommentLength);
        // To get consistent behavior with the generation part, we will assume that
        // this is utf8 encoded.
        this.zipComment = jszipProto.utf8decode(this.zipComment);
    },
    /**
     * Read the end of the Zip 64 central directory.
     * Not merged with the method readEndOfCentral :
     * The end of central can coexist with its Zip64 brother,
     * I don't want to read the wrong number of bytes !
     */
    readBlockZip64EndOfCentral: function() {
        this.zip64EndOfCentralSize = this.reader.readInt(8);
        this.versionMadeBy = this.reader.readString(2);
        this.versionNeeded = this.reader.readInt(2);
        this.diskNumber = this.reader.readInt(4);
        this.diskWithCentralDirStart = this.reader.readInt(4);
        this.centralDirRecordsOnThisDisk = this.reader.readInt(8);
        this.centralDirRecords = this.reader.readInt(8);
        this.centralDirSize = this.reader.readInt(8);
        this.centralDirOffset = this.reader.readInt(8);

        this.zip64ExtensibleData = {};
        var extraDataSize = this.zip64EndOfCentralSize - 44,
            index = 0,
            extraFieldId,
            extraFieldLength,
            extraFieldValue;
        while (index < extraDataSize) {
            extraFieldId = this.reader.readInt(2);
            extraFieldLength = this.reader.readInt(4);
            extraFieldValue = this.reader.readString(extraFieldLength);
            this.zip64ExtensibleData[extraFieldId] = {
                id: extraFieldId,
                length: extraFieldLength,
                value: extraFieldValue
            };
        }
    },
    /**
     * Read the end of the Zip 64 central directory locator.
     */
    readBlockZip64EndOfCentralLocator: function() {
        this.diskWithZip64CentralDirStart = this.reader.readInt(4);
        this.relativeOffsetEndOfZip64CentralDir = this.reader.readInt(8);
        this.disksCount = this.reader.readInt(4);
        if (this.disksCount > 1) {
            throw new Error("Multi-volumes zip are not supported");
        }
    },
    /**
     * Read the local files, based on the offset read in the central part.
     */
    readLocalFiles: function() {
        var i, file;
        for (i = 0; i < this.files.length; i++) {
            file = this.files[i];
            this.reader.setIndex(file.localHeaderOffset);
            this.checkSignature(sig.LOCAL_FILE_HEADER);
            file.readLocalPart(this.reader);
            file.handleUTF8();
        }
    },
    /**
     * Read the central directory.
     */
    readCentralDir: function() {
        var file;

        this.reader.setIndex(this.centralDirOffset);
        while (this.reader.readString(4) === sig.CENTRAL_FILE_HEADER) {
            file = new ZipEntry({
                zip64: this.zip64
            }, this.loadOptions);
            file.readCentralPart(this.reader);
            this.files.push(file);
        }
    },
    /**
     * Read the end of central directory.
     */
    readEndOfCentral: function() {
        var offset = this.reader.lastIndexOfSignature(sig.CENTRAL_DIRECTORY_END);
        if (offset === -1) {
            throw new Error("Corrupted zip : can't find end of central directory");
        }
        this.reader.setIndex(offset);
        this.checkSignature(sig.CENTRAL_DIRECTORY_END);
        this.readBlockEndOfCentral();


        /* extract from the zip spec :
            4)  If one of the fields in the end of central directory
                record is too small to hold required data, the field
                should be set to -1 (0xFFFF or 0xFFFFFFFF) and the
                ZIP64 format record should be created.
            5)  The end of central directory record and the
                Zip64 end of central directory locator record must
                reside on the same disk when splitting or spanning
                an archive.
         */
        if (this.diskNumber === utils.MAX_VALUE_16BITS || this.diskWithCentralDirStart === utils.MAX_VALUE_16BITS || this.centralDirRecordsOnThisDisk === utils.MAX_VALUE_16BITS || this.centralDirRecords === utils.MAX_VALUE_16BITS || this.centralDirSize === utils.MAX_VALUE_32BITS || this.centralDirOffset === utils.MAX_VALUE_32BITS) {
            this.zip64 = true;

            /*
            Warning : the zip64 extension is supported, but ONLY if the 64bits integer read from
            the zip file can fit into a 32bits integer. This cannot be solved : Javascript represents
            all numbers as 64-bit double precision IEEE 754 floating point numbers.
            So, we have 53bits for integers and bitwise operations treat everything as 32bits.
            see https://developer.mozilla.org/en-US/docs/JavaScript/Reference/Operators/Bitwise_Operators
            and http://www.ecma-international.org/publications/files/ECMA-ST/ECMA-262.pdf section 8.5
            */

            // should look for a zip64 EOCD locator
            offset = this.reader.lastIndexOfSignature(sig.ZIP64_CENTRAL_DIRECTORY_LOCATOR);
            if (offset === -1) {
                throw new Error("Corrupted zip : can't find the ZIP64 end of central directory locator");
            }
            this.reader.setIndex(offset);
            this.checkSignature(sig.ZIP64_CENTRAL_DIRECTORY_LOCATOR);
            this.readBlockZip64EndOfCentralLocator();

            // now the zip64 EOCD record
            this.reader.setIndex(this.relativeOffsetEndOfZip64CentralDir);
            this.checkSignature(sig.ZIP64_CENTRAL_DIRECTORY_END);
            this.readBlockZip64EndOfCentral();
        }
    },
    prepareReader: function(data) {
        var type = utils.getTypeOf(data);
        if (type === "string" && !support.uint8array) {
            this.reader = new StringReader(data, this.loadOptions.optimizedBinaryString);
        }
        else if (type === "nodebuffer") {
            this.reader = new NodeBufferReader(data);
        }
        else {
            this.reader = new Uint8ArrayReader(utils.transformTo("uint8array", data));
        }
    },
    /**
     * Read a zip file and create ZipEntries.
     * @param {String|ArrayBuffer|Uint8Array|Buffer} data the binary string representing a zip file.
     */
    load: function(data) {
        this.prepareReader(data);
        this.readEndOfCentral();
        this.readCentralDir();
        this.readLocalFiles();
    }
};
// }}} end of ZipEntries
module.exports = ZipEntries;

},{"./nodeBufferReader":12,"./object":13,"./signature":14,"./stringReader":15,"./support":17,"./uint8ArrayReader":18,"./utils":21,"./zipEntry":23}],23:[function(_dereq_,module,exports){
'use strict';
var StringReader = _dereq_('./stringReader');
var utils = _dereq_('./utils');
var CompressedObject = _dereq_('./compressedObject');
var jszipProto = _dereq_('./object');
// class ZipEntry {{{
/**
 * An entry in the zip file.
 * @constructor
 * @param {Object} options Options of the current file.
 * @param {Object} loadOptions Options for loading the stream.
 */
function ZipEntry(options, loadOptions) {
    this.options = options;
    this.loadOptions = loadOptions;
}
ZipEntry.prototype = {
    /**
     * say if the file is encrypted.
     * @return {boolean} true if the file is encrypted, false otherwise.
     */
    isEncrypted: function() {
        // bit 1 is set
        return (this.bitFlag & 0x0001) === 0x0001;
    },
    /**
     * say if the file has utf-8 filename/comment.
     * @return {boolean} true if the filename/comment is in utf-8, false otherwise.
     */
    useUTF8: function() {
        // bit 11 is set
        return (this.bitFlag & 0x0800) === 0x0800;
    },
    /**
     * Prepare the function used to generate the compressed content from this ZipFile.
     * @param {DataReader} reader the reader to use.
     * @param {number} from the offset from where we should read the data.
     * @param {number} length the length of the data to read.
     * @return {Function} the callback to get the compressed content (the type depends of the DataReader class).
     */
    prepareCompressedContent: function(reader, from, length) {
        return function() {
            var previousIndex = reader.index;
            reader.setIndex(from);
            var compressedFileData = reader.readData(length);
            reader.setIndex(previousIndex);

            return compressedFileData;
        };
    },
    /**
     * Prepare the function used to generate the uncompressed content from this ZipFile.
     * @param {DataReader} reader the reader to use.
     * @param {number} from the offset from where we should read the data.
     * @param {number} length the length of the data to read.
     * @param {JSZip.compression} compression the compression used on this file.
     * @param {number} uncompressedSize the uncompressed size to expect.
     * @return {Function} the callback to get the uncompressed content (the type depends of the DataReader class).
     */
    prepareContent: function(reader, from, length, compression, uncompressedSize) {
        return function() {

            var compressedFileData = utils.transformTo(compression.uncompressInputType, this.getCompressedContent());
            var uncompressedFileData = compression.uncompress(compressedFileData);

            if (uncompressedFileData.length !== uncompressedSize) {
                throw new Error("Bug : uncompressed data size mismatch");
            }

            return uncompressedFileData;
        };
    },
    /**
     * Read the local part of a zip file and add the info in this object.
     * @param {DataReader} reader the reader to use.
     */
    readLocalPart: function(reader) {
        var compression, localExtraFieldsLength;

        // we already know everything from the central dir !
        // If the central dir data are false, we are doomed.
        // On the bright side, the local part is scary  : zip64, data descriptors, both, etc.
        // The less data we get here, the more reliable this should be.
        // Let's skip the whole header and dash to the data !
        reader.skip(22);
        // in some zip created on windows, the filename stored in the central dir contains \ instead of /.
        // Strangely, the filename here is OK.
        // I would love to treat these zip files as corrupted (see http://www.info-zip.org/FAQ.html#backslashes
        // or APPNOTE#4.4.17.1, "All slashes MUST be forward slashes '/'") but there are a lot of bad zip generators...
        // Search "unzip mismatching "local" filename continuing with "central" filename version" on
        // the internet.
        //
        // I think I see the logic here : the central directory is used to display
        // content and the local directory is used to extract the files. Mixing / and \
        // may be used to display \ to windows users and use / when extracting the files.
        // Unfortunately, this lead also to some issues : http://seclists.org/fulldisclosure/2009/Sep/394
        this.fileNameLength = reader.readInt(2);
        localExtraFieldsLength = reader.readInt(2); // can't be sure this will be the same as the central dir
        this.fileName = reader.readString(this.fileNameLength);
        reader.skip(localExtraFieldsLength);

        if (this.compressedSize == -1 || this.uncompressedSize == -1) {
            throw new Error("Bug or corrupted zip : didn't get enough informations from the central directory " + "(compressedSize == -1 || uncompressedSize == -1)");
        }

        compression = utils.findCompression(this.compressionMethod);
        if (compression === null) { // no compression found
            throw new Error("Corrupted zip : compression " + utils.pretty(this.compressionMethod) + " unknown (inner file : " + this.fileName + ")");
        }
        this.decompressed = new CompressedObject();
        this.decompressed.compressedSize = this.compressedSize;
        this.decompressed.uncompressedSize = this.uncompressedSize;
        this.decompressed.crc32 = this.crc32;
        this.decompressed.compressionMethod = this.compressionMethod;
        this.decompressed.getCompressedContent = this.prepareCompressedContent(reader, reader.index, this.compressedSize, compression);
        this.decompressed.getContent = this.prepareContent(reader, reader.index, this.compressedSize, compression, this.uncompressedSize);

        // we need to compute the crc32...
        if (this.loadOptions.checkCRC32) {
            this.decompressed = utils.transformTo("string", this.decompressed.getContent());
            if (jszipProto.crc32(this.decompressed) !== this.crc32) {
                throw new Error("Corrupted zip : CRC32 mismatch");
            }
        }
    },

    /**
     * Read the central part of a zip file and add the info in this object.
     * @param {DataReader} reader the reader to use.
     */
    readCentralPart: function(reader) {
        this.versionMadeBy = reader.readString(2);
        this.versionNeeded = reader.readInt(2);
        this.bitFlag = reader.readInt(2);
        this.compressionMethod = reader.readString(2);
        this.date = reader.readDate();
        this.crc32 = reader.readInt(4);
        this.compressedSize = reader.readInt(4);
        this.uncompressedSize = reader.readInt(4);
        this.fileNameLength = reader.readInt(2);
        this.extraFieldsLength = reader.readInt(2);
        this.fileCommentLength = reader.readInt(2);
        this.diskNumberStart = reader.readInt(2);
        this.internalFileAttributes = reader.readInt(2);
        this.externalFileAttributes = reader.readInt(4);
        this.localHeaderOffset = reader.readInt(4);

        if (this.isEncrypted()) {
            throw new Error("Encrypted zip are not supported");
        }

        this.fileName = reader.readString(this.fileNameLength);
        this.readExtraFields(reader);
        this.parseZIP64ExtraField(reader);
        this.fileComment = reader.readString(this.fileCommentLength);

        // warning, this is true only for zip with madeBy == DOS (plateform dependent feature)
        this.dir = this.externalFileAttributes & 0x00000010 ? true : false;
    },
    /**
     * Parse the ZIP64 extra field and merge the info in the current ZipEntry.
     * @param {DataReader} reader the reader to use.
     */
    parseZIP64ExtraField: function(reader) {

        if (!this.extraFields[0x0001]) {
            return;
        }

        // should be something, preparing the extra reader
        var extraReader = new StringReader(this.extraFields[0x0001].value);

        // I really hope that these 64bits integer can fit in 32 bits integer, because js
        // won't let us have more.
        if (this.uncompressedSize === utils.MAX_VALUE_32BITS) {
            this.uncompressedSize = extraReader.readInt(8);
        }
        if (this.compressedSize === utils.MAX_VALUE_32BITS) {
            this.compressedSize = extraReader.readInt(8);
        }
        if (this.localHeaderOffset === utils.MAX_VALUE_32BITS) {
            this.localHeaderOffset = extraReader.readInt(8);
        }
        if (this.diskNumberStart === utils.MAX_VALUE_32BITS) {
            this.diskNumberStart = extraReader.readInt(4);
        }
    },
    /**
     * Read the central part of a zip file and add the info in this object.
     * @param {DataReader} reader the reader to use.
     */
    readExtraFields: function(reader) {
        var start = reader.index,
            extraFieldId,
            extraFieldLength,
            extraFieldValue;

        this.extraFields = this.extraFields || {};

        while (reader.index < start + this.extraFieldsLength) {
            extraFieldId = reader.readInt(2);
            extraFieldLength = reader.readInt(2);
            extraFieldValue = reader.readString(extraFieldLength);

            this.extraFields[extraFieldId] = {
                id: extraFieldId,
                length: extraFieldLength,
                value: extraFieldValue
            };
        }
    },
    /**
     * Apply an UTF8 transformation if needed.
     */
    handleUTF8: function() {
        if (this.useUTF8()) {
            this.fileName = jszipProto.utf8decode(this.fileName);
            this.fileComment = jszipProto.utf8decode(this.fileComment);
        } else {
            var upath = this.findExtraFieldUnicodePath();
            if (upath !== null) {
                this.fileName = upath;
            }
            var ucomment = this.findExtraFieldUnicodeComment();
            if (ucomment !== null) {
                this.fileComment = ucomment;
            }
        }
    },

    /**
     * Find the unicode path declared in the extra field, if any.
     * @return {String} the unicode path, null otherwise.
     */
    findExtraFieldUnicodePath: function() {
        var upathField = this.extraFields[0x7075];
        if (upathField) {
            var extraReader = new StringReader(upathField.value);

            // wrong version
            if (extraReader.readInt(1) !== 1) {
                return null;
            }

            // the crc of the filename changed, this field is out of date.
            if (jszipProto.crc32(this.fileName) !== extraReader.readInt(4)) {
                return null;
            }

            return jszipProto.utf8decode(extraReader.readString(upathField.length - 5));
        }
        return null;
    },

    /**
     * Find the unicode comment declared in the extra field, if any.
     * @return {String} the unicode comment, null otherwise.
     */
    findExtraFieldUnicodeComment: function() {
        var ucommentField = this.extraFields[0x6375];
        if (ucommentField) {
            var extraReader = new StringReader(ucommentField.value);

            // wrong version
            if (extraReader.readInt(1) !== 1) {
                return null;
            }

            // the crc of the comment changed, this field is out of date.
            if (jszipProto.crc32(this.fileComment) !== extraReader.readInt(4)) {
                return null;
            }

            return jszipProto.utf8decode(extraReader.readString(ucommentField.length - 5));
        }
        return null;
    }
};
module.exports = ZipEntry;

},{"./compressedObject":2,"./object":13,"./stringReader":15,"./utils":21}],24:[function(_dereq_,module,exports){
// Top level file is just a mixin of submodules & constants
'use strict';

var assign    = _dereq_('./lib/utils/common').assign;

var deflate   = _dereq_('./lib/deflate');
var inflate   = _dereq_('./lib/inflate');
var constants = _dereq_('./lib/zlib/constants');

var pako = {};

assign(pako, deflate, inflate, constants);

module.exports = pako;
},{"./lib/deflate":25,"./lib/inflate":26,"./lib/utils/common":27,"./lib/zlib/constants":30}],25:[function(_dereq_,module,exports){
'use strict';


var zlib_deflate = _dereq_('./zlib/deflate.js');
var utils = _dereq_('./utils/common');
var strings = _dereq_('./utils/strings');
var msg = _dereq_('./zlib/messages');
var zstream = _dereq_('./zlib/zstream');


/* Public constants ==========================================================*/
/* ===========================================================================*/

var Z_NO_FLUSH      = 0;
var Z_FINISH        = 4;

var Z_OK            = 0;
var Z_STREAM_END    = 1;

var Z_DEFAULT_COMPRESSION = -1;

var Z_DEFAULT_STRATEGY    = 0;

var Z_DEFLATED  = 8;

/* ===========================================================================*/


/**
 * class Deflate
 *
 * Generic JS-style wrapper for zlib calls. If you don't need
 * streaming behaviour - use more simple functions: [[deflate]],
 * [[deflateRaw]] and [[gzip]].
 **/

/* internal
 * Deflate.chunks -> Array
 *
 * Chunks of output data, if [[Deflate#onData]] not overriden.
 **/

/**
 * Deflate.result -> Uint8Array|Array
 *
 * Compressed result, generated by default [[Deflate#onData]]
 * and [[Deflate#onEnd]] handlers. Filled after you push last chunk
 * (call [[Deflate#push]] with `Z_FINISH` / `true` param).
 **/

/**
 * Deflate.err -> Number
 *
 * Error code after deflate finished. 0 (Z_OK) on success.
 * You will not need it in real life, because deflate errors
 * are possible only on wrong options or bad `onData` / `onEnd`
 * custom handlers.
 **/

/**
 * Deflate.msg -> String
 *
 * Error message, if [[Deflate.err]] != 0
 **/


/**
 * new Deflate(options)
 * - options (Object): zlib deflate options.
 *
 * Creates new deflator instance with specified params. Throws exception
 * on bad params. Supported options:
 *
 * - `level`
 * - `windowBits`
 * - `memLevel`
 * - `strategy`
 *
 * [http://zlib.net/manual.html#Advanced](http://zlib.net/manual.html#Advanced)
 * for more information on these.
 *
 * Additional options, for internal needs:
 *
 * - `chunkSize` - size of generated data chunks (16K by default)
 * - `raw` (Boolean) - do raw deflate
 * - `gzip` (Boolean) - create gzip wrapper
 * - `to` (String) - if equal to 'string', then result will be "binary string"
 *    (each char code [0..255])
 * - `header` (Object) - custom header for gzip
 *   - `text` (Boolean) - true if compressed data believed to be text
 *   - `time` (Number) - modification time, unix timestamp
 *   - `os` (Number) - operation system code
 *   - `extra` (Array) - array of bytes with extra data (max 65536)
 *   - `name` (String) - file name (binary string)
 *   - `comment` (String) - comment (binary string)
 *   - `hcrc` (Boolean) - true if header crc should be added
 *
 * ##### Example:
 *
 * ```javascript
 * var pako = require('pako')
 *   , chunk1 = Uint8Array([1,2,3,4,5,6,7,8,9])
 *   , chunk2 = Uint8Array([10,11,12,13,14,15,16,17,18,19]);
 *
 * var deflate = new pako.Deflate({ level: 3});
 *
 * deflate.push(chunk1, false);
 * deflate.push(chunk2, true);  // true -> last chunk
 *
 * if (deflate.err) { throw new Error(deflate.err); }
 *
 * console.log(deflate.result);
 * ```
 **/
var Deflate = function(options) {

  this.options = utils.assign({
    level: Z_DEFAULT_COMPRESSION,
    method: Z_DEFLATED,
    chunkSize: 16384,
    windowBits: 15,
    memLevel: 8,
    strategy: Z_DEFAULT_STRATEGY,
    to: ''
  }, options || {});

  var opt = this.options;

  if (opt.raw && (opt.windowBits > 0)) {
    opt.windowBits = -opt.windowBits;
  }

  else if (opt.gzip && (opt.windowBits > 0) && (opt.windowBits < 16)) {
    opt.windowBits += 16;
  }

  this.err    = 0;      // error code, if happens (0 = Z_OK)
  this.msg    = '';     // error message
  this.ended  = false;  // used to avoid multiple onEnd() calls
  this.chunks = [];     // chunks of compressed data

  this.strm = new zstream();
  this.strm.avail_out = 0;

  var status = zlib_deflate.deflateInit2(
    this.strm,
    opt.level,
    opt.method,
    opt.windowBits,
    opt.memLevel,
    opt.strategy
  );

  if (status !== Z_OK) {
    throw new Error(msg[status]);
  }

  if (opt.header) {
    zlib_deflate.deflateSetHeader(this.strm, opt.header);
  }
};

/**
 * Deflate#push(data[, mode]) -> Boolean
 * - data (Uint8Array|Array|String): input data. Strings will be converted to
 *   utf8 byte sequence.
 * - mode (Number|Boolean): 0..6 for corresponding Z_NO_FLUSH..Z_TREE modes.
 *   See constants. Skipped or `false` means Z_NO_FLUSH, `true` meansh Z_FINISH.
 *
 * Sends input data to deflate pipe, generating [[Deflate#onData]] calls with
 * new compressed chunks. Returns `true` on success. The last data block must have
 * mode Z_FINISH (or `true`). That flush internal pending buffers and call
 * [[Deflate#onEnd]].
 *
 * On fail call [[Deflate#onEnd]] with error code and return false.
 *
 * We strongly recommend to use `Uint8Array` on input for best speed (output
 * array format is detected automatically). Also, don't skip last param and always
 * use the same type in your code (boolean or number). That will improve JS speed.
 *
 * For regular `Array`-s make sure all elements are [0..255].
 *
 * ##### Example
 *
 * ```javascript
 * push(chunk, false); // push one of data chunks
 * ...
 * push(chunk, true);  // push last chunk
 * ```
 **/
Deflate.prototype.push = function(data, mode) {
  var strm = this.strm;
  var chunkSize = this.options.chunkSize;
  var status, _mode;

  if (this.ended) { return false; }

  _mode = (mode === ~~mode) ? mode : ((mode === true) ? Z_FINISH : Z_NO_FLUSH);

  // Convert data if needed
  if (typeof data === 'string') {
    // If we need to compress text, change encoding to utf8.
    strm.input = strings.string2buf(data);
  } else {
    strm.input = data;
  }

  strm.next_in = 0;
  strm.avail_in = strm.input.length;

  do {
    if (strm.avail_out === 0) {
      strm.output = new utils.Buf8(chunkSize);
      strm.next_out = 0;
      strm.avail_out = chunkSize;
    }
    status = zlib_deflate.deflate(strm, _mode);    /* no bad return value */

    if (status !== Z_STREAM_END && status !== Z_OK) {
      this.onEnd(status);
      this.ended = true;
      return false;
    }
    if (strm.avail_out === 0 || (strm.avail_in === 0 && _mode === Z_FINISH)) {
      if (this.options.to === 'string') {
        this.onData(strings.buf2binstring(utils.shrinkBuf(strm.output, strm.next_out)));
      } else {
        this.onData(utils.shrinkBuf(strm.output, strm.next_out));
      }
    }
  } while ((strm.avail_in > 0 || strm.avail_out === 0) && status !== Z_STREAM_END);

  // Finalize on the last chunk.
  if (_mode === Z_FINISH) {
    status = zlib_deflate.deflateEnd(this.strm);
    this.onEnd(status);
    this.ended = true;
    return status === Z_OK;
  }

  return true;
};


/**
 * Deflate#onData(chunk) -> Void
 * - chunk (Uint8Array|Array|String): ouput data. Type of array depends
 *   on js engine support. When string output requested, each chunk
 *   will be string.
 *
 * By default, stores data blocks in `chunks[]` property and glue
 * those in `onEnd`. Override this handler, if you need another behaviour.
 **/
Deflate.prototype.onData = function(chunk) {
  this.chunks.push(chunk);
};


/**
 * Deflate#onEnd(status) -> Void
 * - status (Number): deflate status. 0 (Z_OK) on success,
 *   other if not.
 *
 * Called once after you tell deflate that input stream complete
 * or error happenned. By default - join collected chunks,
 * free memory and fill `results` / `err` properties.
 **/
Deflate.prototype.onEnd = function(status) {
  // On success - join
  if (status === Z_OK) {
    if (this.options.to === 'string') {
      this.result = this.chunks.join('');
    } else {
      this.result = utils.flattenChunks(this.chunks);
    }
  }
  this.chunks = [];
  this.err = status;
  this.msg = this.strm.msg;
};


/**
 * deflate(data[, options]) -> Uint8Array|Array|String
 * - data (Uint8Array|Array|String): input data to compress.
 * - options (Object): zlib deflate options.
 *
 * Compress `data` with deflate alrorythm and `options`.
 *
 * Supported options are:
 *
 * - level
 * - windowBits
 * - memLevel
 * - strategy
 *
 * [http://zlib.net/manual.html#Advanced](http://zlib.net/manual.html#Advanced)
 * for more information on these.
 *
 * Sugar (options):
 *
 * - `raw` (Boolean) - say that we work with raw stream, if you don't wish to specify
 *   negative windowBits implicitly.
 * - `to` (String) - if equal to 'string', then result will be "binary string"
 *    (each char code [0..255])
 *
 * ##### Example:
 *
 * ```javascript
 * var pako = require('pako')
 *   , data = Uint8Array([1,2,3,4,5,6,7,8,9]);
 *
 * console.log(pako.deflate(data));
 * ```
 **/
function deflate(input, options) {
  var deflator = new Deflate(options);

  deflator.push(input, true);

  // That will never happens, if you don't cheat with options :)
  if (deflator.err) { throw deflator.msg; }

  return deflator.result;
}


/**
 * deflateRaw(data[, options]) -> Uint8Array|Array|String
 * - data (Uint8Array|Array|String): input data to compress.
 * - options (Object): zlib deflate options.
 *
 * The same as [[deflate]], but creates raw data, without wrapper
 * (header and adler32 crc).
 **/
function deflateRaw(input, options) {
  options = options || {};
  options.raw = true;
  return deflate(input, options);
}


/**
 * gzip(data[, options]) -> Uint8Array|Array|String
 * - data (Uint8Array|Array|String): input data to compress.
 * - options (Object): zlib deflate options.
 *
 * The same as [[deflate]], but create gzip wrapper instead of
 * deflate one.
 **/
function gzip(input, options) {
  options = options || {};
  options.gzip = true;
  return deflate(input, options);
}


exports.Deflate = Deflate;
exports.deflate = deflate;
exports.deflateRaw = deflateRaw;
exports.gzip = gzip;
},{"./utils/common":27,"./utils/strings":28,"./zlib/deflate.js":32,"./zlib/messages":37,"./zlib/zstream":39}],26:[function(_dereq_,module,exports){
'use strict';


var zlib_inflate = _dereq_('./zlib/inflate.js');
var utils = _dereq_('./utils/common');
var strings = _dereq_('./utils/strings');
var c = _dereq_('./zlib/constants');
var msg = _dereq_('./zlib/messages');
var zstream = _dereq_('./zlib/zstream');
var gzheader = _dereq_('./zlib/gzheader');


/**
 * class Inflate
 *
 * Generic JS-style wrapper for zlib calls. If you don't need
 * streaming behaviour - use more simple functions: [[inflate]]
 * and [[inflateRaw]].
 **/

/* internal
 * inflate.chunks -> Array
 *
 * Chunks of output data, if [[Inflate#onData]] not overriden.
 **/

/**
 * Inflate.result -> Uint8Array|Array|String
 *
 * Uncompressed result, generated by default [[Inflate#onData]]
 * and [[Inflate#onEnd]] handlers. Filled after you push last chunk
 * (call [[Inflate#push]] with `Z_FINISH` / `true` param).
 **/

/**
 * Inflate.err -> Number
 *
 * Error code after inflate finished. 0 (Z_OK) on success.
 * Should be checked if broken data possible.
 **/

/**
 * Inflate.msg -> String
 *
 * Error message, if [[Inflate.err]] != 0
 **/


/**
 * new Inflate(options)
 * - options (Object): zlib inflate options.
 *
 * Creates new inflator instance with specified params. Throws exception
 * on bad params. Supported options:
 *
 * - `windowBits`
 *
 * [http://zlib.net/manual.html#Advanced](http://zlib.net/manual.html#Advanced)
 * for more information on these.
 *
 * Additional options, for internal needs:
 *
 * - `chunkSize` - size of generated data chunks (16K by default)
 * - `raw` (Boolean) - do raw inflate
 * - `to` (String) - if equal to 'string', then result will be converted
 *   from utf8 to utf16 (javascript) string. When string output requested,
 *   chunk length can differ from `chunkSize`, depending on content.
 *
 * By default, when no options set, autodetect deflate/gzip data format via
 * wrapper header.
 *
 * ##### Example:
 *
 * ```javascript
 * var pako = require('pako')
 *   , chunk1 = Uint8Array([1,2,3,4,5,6,7,8,9])
 *   , chunk2 = Uint8Array([10,11,12,13,14,15,16,17,18,19]);
 *
 * var inflate = new pako.Inflate({ level: 3});
 *
 * inflate.push(chunk1, false);
 * inflate.push(chunk2, true);  // true -> last chunk
 *
 * if (inflate.err) { throw new Error(inflate.err); }
 *
 * console.log(inflate.result);
 * ```
 **/
var Inflate = function(options) {

  this.options = utils.assign({
    chunkSize: 16384,
    windowBits: 0,
    to: ''
  }, options || {});

  var opt = this.options;

  // Force window size for `raw` data, if not set directly,
  // because we have no header for autodetect.
  if (opt.raw && (opt.windowBits >= 0) && (opt.windowBits < 16)) {
    opt.windowBits = -opt.windowBits;
    if (opt.windowBits === 0) { opt.windowBits = -15; }
  }

  // If `windowBits` not defined (and mode not raw) - set autodetect flag for gzip/deflate
  if ((opt.windowBits >= 0) && (opt.windowBits < 16) &&
      !(options && options.windowBits)) {
    opt.windowBits += 32;
  }

  // Gzip header has no info about windows size, we can do autodetect only
  // for deflate. So, if window size not set, force it to max when gzip possible
  if ((opt.windowBits > 15) && (opt.windowBits < 48)) {
    // bit 3 (16) -> gzipped data
    // bit 4 (32) -> autodetect gzip/deflate
    if ((opt.windowBits & 15) === 0) {
      opt.windowBits |= 15;
    }
  }

  this.err    = 0;      // error code, if happens (0 = Z_OK)
  this.msg    = '';     // error message
  this.ended  = false;  // used to avoid multiple onEnd() calls
  this.chunks = [];     // chunks of compressed data

  this.strm   = new zstream();
  this.strm.avail_out = 0;

  var status  = zlib_inflate.inflateInit2(
    this.strm,
    opt.windowBits
  );

  if (status !== c.Z_OK) {
    throw new Error(msg[status]);
  }

  this.header = new gzheader();

  zlib_inflate.inflateGetHeader(this.strm, this.header);
};

/**
 * Inflate#push(data[, mode]) -> Boolean
 * - data (Uint8Array|Array|String): input data
 * - mode (Number|Boolean): 0..6 for corresponding Z_NO_FLUSH..Z_TREE modes.
 *   See constants. Skipped or `false` means Z_NO_FLUSH, `true` meansh Z_FINISH.
 *
 * Sends input data to inflate pipe, generating [[Inflate#onData]] calls with
 * new output chunks. Returns `true` on success. The last data block must have
 * mode Z_FINISH (or `true`). That flush internal pending buffers and call
 * [[Inflate#onEnd]].
 *
 * On fail call [[Inflate#onEnd]] with error code and return false.
 *
 * We strongly recommend to use `Uint8Array` on input for best speed (output
 * format is detected automatically). Also, don't skip last param and always
 * use the same type in your code (boolean or number). That will improve JS speed.
 *
 * For regular `Array`-s make sure all elements are [0..255].
 *
 * ##### Example
 *
 * ```javascript
 * push(chunk, false); // push one of data chunks
 * ...
 * push(chunk, true);  // push last chunk
 * ```
 **/
Inflate.prototype.push = function(data, mode) {
  var strm = this.strm;
  var chunkSize = this.options.chunkSize;
  var status, _mode;
  var next_out_utf8, tail, utf8str;

  if (this.ended) { return false; }
  _mode = (mode === ~~mode) ? mode : ((mode === true) ? c.Z_FINISH : c.Z_NO_FLUSH);

  // Convert data if needed
  if (typeof data === 'string') {
    // Only binary strings can be decompressed on practice
    strm.input = strings.binstring2buf(data);
  } else {
    strm.input = data;
  }

  strm.next_in = 0;
  strm.avail_in = strm.input.length;

  do {
    if (strm.avail_out === 0) {
      strm.output = new utils.Buf8(chunkSize);
      strm.next_out = 0;
      strm.avail_out = chunkSize;
    }

    status = zlib_inflate.inflate(strm, c.Z_NO_FLUSH);    /* no bad return value */

    if (status !== c.Z_STREAM_END && status !== c.Z_OK) {
      this.onEnd(status);
      this.ended = true;
      return false;
    }

    if (strm.next_out) {
      if (strm.avail_out === 0 || status === c.Z_STREAM_END || (strm.avail_in === 0 && _mode === c.Z_FINISH)) {

        if (this.options.to === 'string') {

          next_out_utf8 = strings.utf8border(strm.output, strm.next_out);

          tail = strm.next_out - next_out_utf8;
          utf8str = strings.buf2string(strm.output, next_out_utf8);

          // move tail
          strm.next_out = tail;
          strm.avail_out = chunkSize - tail;
          if (tail) { utils.arraySet(strm.output, strm.output, next_out_utf8, tail, 0); }

          this.onData(utf8str);

        } else {
          this.onData(utils.shrinkBuf(strm.output, strm.next_out));
        }
      }
    }
  } while ((strm.avail_in > 0) && status !== c.Z_STREAM_END);

  if (status === c.Z_STREAM_END) {
    _mode = c.Z_FINISH;
  }
  // Finalize on the last chunk.
  if (_mode === c.Z_FINISH) {
    status = zlib_inflate.inflateEnd(this.strm);
    this.onEnd(status);
    this.ended = true;
    return status === c.Z_OK;
  }

  return true;
};


/**
 * Inflate#onData(chunk) -> Void
 * - chunk (Uint8Array|Array|String): ouput data. Type of array depends
 *   on js engine support. When string output requested, each chunk
 *   will be string.
 *
 * By default, stores data blocks in `chunks[]` property and glue
 * those in `onEnd`. Override this handler, if you need another behaviour.
 **/
Inflate.prototype.onData = function(chunk) {
  this.chunks.push(chunk);
};


/**
 * Inflate#onEnd(status) -> Void
 * - status (Number): inflate status. 0 (Z_OK) on success,
 *   other if not.
 *
 * Called once after you tell inflate that input stream complete
 * or error happenned. By default - join collected chunks,
 * free memory and fill `results` / `err` properties.
 **/
Inflate.prototype.onEnd = function(status) {
  // On success - join
  if (status === c.Z_OK) {
    if (this.options.to === 'string') {
      // Glue & convert here, until we teach pako to send
      // utf8 alligned strings to onData
      this.result = this.chunks.join('');
    } else {
      this.result = utils.flattenChunks(this.chunks);
    }
  }
  this.chunks = [];
  this.err = status;
  this.msg = this.strm.msg;
};


/**
 * inflate(data[, options]) -> Uint8Array|Array|String
 * - data (Uint8Array|Array|String): input data to decompress.
 * - options (Object): zlib inflate options.
 *
 * Decompress `data` with inflate/ungzip and `options`. Autodetect
 * format via wrapper header by default. That's why we don't provide
 * separate `ungzip` method.
 *
 * Supported options are:
 *
 * - windowBits
 *
 * [http://zlib.net/manual.html#Advanced](http://zlib.net/manual.html#Advanced)
 * for more information.
 *
 * Sugar (options):
 *
 * - `raw` (Boolean) - say that we work with raw stream, if you don't wish to specify
 *   negative windowBits implicitly.
 * - `to` (String) - if equal to 'string', then result will be converted
 *   from utf8 to utf16 (javascript) string. When string output requested,
 *   chunk length can differ from `chunkSize`, depending on content.
 *
 *
 * ##### Example:
 *
 * ```javascript
 * var pako = require('pako')
 *   , input = pako.deflate([1,2,3,4,5,6,7,8,9])
 *   , output;
 *
 * try {
 *   output = pako.inflate(input);
 * } catch (err)
 *   console.log(err);
 * }
 * ```
 **/
function inflate(input, options) {
  var inflator = new Inflate(options);

  inflator.push(input, true);

  // That will never happens, if you don't cheat with options :)
  if (inflator.err) { throw inflator.msg; }

  return inflator.result;
}


/**
 * inflateRaw(data[, options]) -> Uint8Array|Array|String
 * - data (Uint8Array|Array|String): input data to decompress.
 * - options (Object): zlib inflate options.
 *
 * The same as [[inflate]], but creates raw data, without wrapper
 * (header and adler32 crc).
 **/
function inflateRaw(input, options) {
  options = options || {};
  options.raw = true;
  return inflate(input, options);
}


/**
 * ungzip(data[, options]) -> Uint8Array|Array|String
 * - data (Uint8Array|Array|String): input data to decompress.
 * - options (Object): zlib inflate options.
 *
 * Just shortcut to [[inflate]], because it autodetects format
 * by header.content. Done for convenience.
 **/


exports.Inflate = Inflate;
exports.inflate = inflate;
exports.inflateRaw = inflateRaw;
exports.ungzip  = inflate;

},{"./utils/common":27,"./utils/strings":28,"./zlib/constants":30,"./zlib/gzheader":33,"./zlib/inflate.js":35,"./zlib/messages":37,"./zlib/zstream":39}],27:[function(_dereq_,module,exports){
'use strict';


var TYPED_OK =  (typeof Uint8Array !== 'undefined') &&
                (typeof Uint16Array !== 'undefined') &&
                (typeof Int32Array !== 'undefined');


exports.assign = function (obj /*from1, from2, from3, ...*/) {
  var sources = Array.prototype.slice.call(arguments, 1);
  while (sources.length) {
    var source = sources.shift();
    if (!source) { continue; }

    if (typeof(source) !== 'object') {
      throw new TypeError(source + 'must be non-object');
    }

    for (var p in source) {
      if (source.hasOwnProperty(p)) {
        obj[p] = source[p];
      }
    }
  }

  return obj;
};


// reduce buffer size, avoiding mem copy
exports.shrinkBuf = function (buf, size) {
  if (buf.length === size) { return buf; }
  if (buf.subarray) { return buf.subarray(0, size); }
  buf.length = size;
  return buf;
};


var fnTyped = {
  arraySet: function (dest, src, src_offs, len, dest_offs) {
    if (src.subarray && dest.subarray) {
      dest.set(src.subarray(src_offs, src_offs+len), dest_offs);
      return;
    }
    // Fallback to ordinary array
    for(var i=0; i<len; i++) {
      dest[dest_offs + i] = src[src_offs + i];
    }
  },
  // Join array of chunks to single array.
  flattenChunks: function(chunks) {
    var i, l, len, pos, chunk, result;

    // calculate data length
    len = 0;
    for (i=0, l=chunks.length; i<l; i++) {
      len += chunks[i].length;
    }

    // join chunks
    result = new Uint8Array(len);
    pos = 0;
    for (i=0, l=chunks.length; i<l; i++) {
      chunk = chunks[i];
      result.set(chunk, pos);
      pos += chunk.length;
    }

    return result;
  }
};

var fnUntyped = {
  arraySet: function (dest, src, src_offs, len, dest_offs) {
    for(var i=0; i<len; i++) {
      dest[dest_offs + i] = src[src_offs + i];
    }
  },
  // Join array of chunks to single array.
  flattenChunks: function(chunks) {
    return [].concat.apply([], chunks);
  }
};


// Enable/Disable typed arrays use, for testing
//
exports.setTyped = function (on) {
  if (on) {
    exports.Buf8  = Uint8Array;
    exports.Buf16 = Uint16Array;
    exports.Buf32 = Int32Array;
    exports.assign(exports, fnTyped);
  } else {
    exports.Buf8  = Array;
    exports.Buf16 = Array;
    exports.Buf32 = Array;
    exports.assign(exports, fnUntyped);
  }
};

exports.setTyped(TYPED_OK);
},{}],28:[function(_dereq_,module,exports){
// String encode/decode helpers
'use strict';


var utils = _dereq_('./common');


// Quick check if we can use fast array to bin string conversion
//
// - apply(Array) can fail on Android 2.2
// - apply(Uint8Array) can fail on iOS 5.1 Safary
//
var STR_APPLY_OK = true;
var STR_APPLY_UIA_OK = true;

try { String.fromCharCode.apply(null, [0]); } catch(__) { STR_APPLY_OK = false; }
try { String.fromCharCode.apply(null, new Uint8Array(1)); } catch(__) { STR_APPLY_UIA_OK = false; }


// Table with utf8 lengths (calculated by first byte of sequence)
// Note, that 5 & 6-byte values and some 4-byte values can not be represented in JS,
// because max possible codepoint is 0x10ffff
var _utf8len = new utils.Buf8(256);
for (var i=0; i<256; i++) {
  _utf8len[i] = (i >= 252 ? 6 : i >= 248 ? 5 : i >= 240 ? 4 : i >= 224 ? 3 : i >= 192 ? 2 : 1);
}
_utf8len[254]=_utf8len[254]=1; // Invalid sequence start


// convert string to array (typed, when possible)
exports.string2buf = function (str) {
  var buf, c, c2, m_pos, i, str_len = str.length, buf_len = 0;

  // count binary size
  for (m_pos = 0; m_pos < str_len; m_pos++) {
    c = str.charCodeAt(m_pos);
    if ((c & 0xfc00) === 0xd800 && (m_pos+1 < str_len)) {
      c2 = str.charCodeAt(m_pos+1);
      if ((c2 & 0xfc00) === 0xdc00) {
        c = 0x10000 + ((c - 0xd800) << 10) + (c2 - 0xdc00);
        m_pos++;
      }
    }
    buf_len += c < 0x80 ? 1 : c < 0x800 ? 2 : c < 0x10000 ? 3 : 4;
  }

  // allocate buffer
  buf = new utils.Buf8(buf_len);

  // convert
  for (i=0, m_pos = 0; i < buf_len; m_pos++) {
    c = str.charCodeAt(m_pos);
    if ((c & 0xfc00) === 0xd800 && (m_pos+1 < str_len)) {
      c2 = str.charCodeAt(m_pos+1);
      if ((c2 & 0xfc00) === 0xdc00) {
        c = 0x10000 + ((c - 0xd800) << 10) + (c2 - 0xdc00);
        m_pos++;
      }
    }
    if (c < 0x80) {
      /* one byte */
      buf[i++] = c;
    } else if (c < 0x800) {
      /* two bytes */
      buf[i++] = 0xC0 | (c >>> 6);
      buf[i++] = 0x80 | (c & 0x3f);
    } else if (c < 0x10000) {
      /* three bytes */
      buf[i++] = 0xE0 | (c >>> 12);
      buf[i++] = 0x80 | (c >>> 6 & 0x3f);
      buf[i++] = 0x80 | (c & 0x3f);
    } else {
      /* four bytes */
      buf[i++] = 0xf0 | (c >>> 18);
      buf[i++] = 0x80 | (c >>> 12 & 0x3f);
      buf[i++] = 0x80 | (c >>> 6 & 0x3f);
      buf[i++] = 0x80 | (c & 0x3f);
    }
  }

  return buf;
};

// Helper (used in 2 places)
function buf2binstring(buf, len) {
  // use fallback for big arrays to avoid stack overflow
  if (len < 65537) {
    if ((buf.subarray && STR_APPLY_UIA_OK) || (!buf.subarray && STR_APPLY_OK)) {
      return String.fromCharCode.apply(null, utils.shrinkBuf(buf, len));
    }
  }

  var result = '';
  for(var i=0; i < len; i++) {
    result += String.fromCharCode(buf[i]);
  }
  return result;
}


// Convert byte array to binary string
exports.buf2binstring = function(buf) {
  return buf2binstring(buf, buf.length);
};


// Convert binary string (typed, when possible)
exports.binstring2buf = function(str) {
  var buf = new utils.Buf8(str.length);
  for(var i=0, len=buf.length; i < len; i++) {
    buf[i] = str.charCodeAt(i);
  }
  return buf;
};


// convert array to string
exports.buf2string = function (buf, max) {
  var i, out, c, c_len;
  var len = max || buf.length;

  // Reserve max possible length (2 words per char)
  // NB: by unknown reasons, Array is significantly faster for
  //     String.fromCharCode.apply than Uint16Array.
  var utf16buf = new Array(len*2);

  for (out=0, i=0; i<len;) {
    c = buf[i++];
    // quick process ascii
    if (c < 0x80) { utf16buf[out++] = c; continue; }

    c_len = _utf8len[c];
    // skip 5 & 6 byte codes
    if (c_len > 4) { utf16buf[out++] = 0xfffd; i += c_len-1; continue; }

    // apply mask on first byte
    c &= c_len === 2 ? 0x1f : c_len === 3 ? 0x0f : 0x07;
    // join the rest
    while (c_len > 1 && i < len) {
      c = (c << 6) | (buf[i++] & 0x3f);
      c_len--;
    }

    // terminated by end of string?
    if (c_len > 1) { utf16buf[out++] = 0xfffd; continue; }

    if (c < 0x10000) {
      utf16buf[out++] = c;
    } else {
      c -= 0x10000;
      utf16buf[out++] = 0xd800 | ((c >> 10) & 0x3ff);
      utf16buf[out++] = 0xdc00 | (c & 0x3ff);
    }
  }

  return buf2binstring(utf16buf, out);
};


// Calculate max possible position in utf8 buffer,
// that will not break sequence. If that's not possible
// - (very small limits) return max size as is.
//
// buf[] - utf8 bytes array
// max   - length limit (mandatory);
exports.utf8border = function(buf, max) {
  var pos;

  max = max || buf.length;
  if (max > buf.length) { max = buf.length; }

  // go back from last position, until start of sequence found
  pos = max-1;
  while (pos >= 0 && (buf[pos] & 0xC0) === 0x80) { pos--; }

  // Fuckup - very small and broken sequence,
  // return max, because we should return something anyway.
  if (pos < 0) { return max; }

  // If we came to start of buffer - that means vuffer is too small,
  // return max too.
  if (pos === 0) { return max; }

  return (pos + _utf8len[buf[pos]] > max) ? pos : max;
};

},{"./common":27}],29:[function(_dereq_,module,exports){
'use strict';

// Note: adler32 takes 12% for level 0 and 2% for level 6.
// It doesn't worth to make additional optimizationa as in original.
// Small size is preferable.

function adler32(adler, buf, len, pos) {
  var s1 = (adler & 0xffff) |0
    , s2 = ((adler >>> 16) & 0xffff) |0
    , n = 0;

  while (len !== 0) {
    // Set limit ~ twice less than 5552, to keep
    // s2 in 31-bits, because we force signed ints.
    // in other case %= will fail.
    n = len > 2000 ? 2000 : len;
    len -= n;

    do {
      s1 = (s1 + buf[pos++]) |0;
      s2 = (s2 + s1) |0;
    } while (--n);

    s1 %= 65521;
    s2 %= 65521;
  }

  return (s1 | (s2 << 16)) |0;
}


module.exports = adler32;
},{}],30:[function(_dereq_,module,exports){
module.exports = {

  /* Allowed flush values; see deflate() and inflate() below for details */
  Z_NO_FLUSH:         0,
  Z_PARTIAL_FLUSH:    1,
  Z_SYNC_FLUSH:       2,
  Z_FULL_FLUSH:       3,
  Z_FINISH:           4,
  Z_BLOCK:            5,
  Z_TREES:            6,

  /* Return codes for the compression/decompression functions. Negative values
  * are errors, positive values are used for special but normal events.
  */
  Z_OK:               0,
  Z_STREAM_END:       1,
  Z_NEED_DICT:        2,
  Z_ERRNO:           -1,
  Z_STREAM_ERROR:    -2,
  Z_DATA_ERROR:      -3,
  //Z_MEM_ERROR:     -4,
  Z_BUF_ERROR:       -5,
  //Z_VERSION_ERROR: -6,

  /* compression levels */
  Z_NO_COMPRESSION:         0,
  Z_BEST_SPEED:             1,
  Z_BEST_COMPRESSION:       9,
  Z_DEFAULT_COMPRESSION:   -1,


  Z_FILTERED:               1,
  Z_HUFFMAN_ONLY:           2,
  Z_RLE:                    3,
  Z_FIXED:                  4,
  Z_DEFAULT_STRATEGY:       0,

  /* Possible values of the data_type field (though see inflate()) */
  Z_BINARY:                 0,
  Z_TEXT:                   1,
  //Z_ASCII:                1, // = Z_TEXT (deprecated)
  Z_UNKNOWN:                2,

  /* The deflate compression method */
  Z_DEFLATED:               8
  //Z_NULL:                 null // Use -1 or null inline, depending on var type
};
},{}],31:[function(_dereq_,module,exports){
'use strict';

// Note: we can't get significant speed boost here.
// So write code to minimize size - no pregenerated tables
// and array tools dependencies.


// Use ordinary array, since untyped makes no boost here
function makeTable() {
  var c, table = [];

  for(var n =0; n < 256; n++){
    c = n;
    for(var k =0; k < 8; k++){
      c = ((c&1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1));
    }
    table[n] = c;
  }

  return table;
}

// Create table on load. Just 255 signed longs. Not a problem.
var crcTable = makeTable();


function crc32(crc, buf, len, pos) {
  var t = crcTable
    , end = pos + len;

  crc = crc ^ (-1);

  for (var i = pos; i < end; i++ ) {
    crc = (crc >>> 8) ^ t[(crc ^ buf[i]) & 0xFF];
  }

  return (crc ^ (-1)); // >>> 0;
}


module.exports = crc32;
},{}],32:[function(_dereq_,module,exports){
'use strict';

var utils   = _dereq_('../utils/common');
var trees   = _dereq_('./trees');
var adler32 = _dereq_('./adler32');
var crc32   = _dereq_('./crc32');
var msg   = _dereq_('./messages');

/* Public constants ==========================================================*/
/* ===========================================================================*/


/* Allowed flush values; see deflate() and inflate() below for details */
var Z_NO_FLUSH      = 0;
var Z_PARTIAL_FLUSH = 1;
//var Z_SYNC_FLUSH    = 2;
var Z_FULL_FLUSH    = 3;
var Z_FINISH        = 4;
var Z_BLOCK         = 5;
//var Z_TREES         = 6;


/* Return codes for the compression/decompression functions. Negative values
 * are errors, positive values are used for special but normal events.
 */
var Z_OK            = 0;
var Z_STREAM_END    = 1;
//var Z_NEED_DICT     = 2;
//var Z_ERRNO         = -1;
var Z_STREAM_ERROR  = -2;
var Z_DATA_ERROR    = -3;
//var Z_MEM_ERROR     = -4;
var Z_BUF_ERROR     = -5;
//var Z_VERSION_ERROR = -6;


/* compression levels */
//var Z_NO_COMPRESSION      = 0;
//var Z_BEST_SPEED          = 1;
//var Z_BEST_COMPRESSION    = 9;
var Z_DEFAULT_COMPRESSION = -1;


var Z_FILTERED            = 1;
var Z_HUFFMAN_ONLY        = 2;
var Z_RLE                 = 3;
var Z_FIXED               = 4;
var Z_DEFAULT_STRATEGY    = 0;

/* Possible values of the data_type field (though see inflate()) */
//var Z_BINARY              = 0;
//var Z_TEXT                = 1;
//var Z_ASCII               = 1; // = Z_TEXT
var Z_UNKNOWN             = 2;


/* The deflate compression method */
var Z_DEFLATED  = 8;

/*============================================================================*/


var MAX_MEM_LEVEL = 9;
/* Maximum value for memLevel in deflateInit2 */
var MAX_WBITS = 15;
/* 32K LZ77 window */
var DEF_MEM_LEVEL = 8;


var LENGTH_CODES  = 29;
/* number of length codes, not counting the special END_BLOCK code */
var LITERALS      = 256;
/* number of literal bytes 0..255 */
var L_CODES       = LITERALS + 1 + LENGTH_CODES;
/* number of Literal or Length codes, including the END_BLOCK code */
var D_CODES       = 30;
/* number of distance codes */
var BL_CODES      = 19;
/* number of codes used to transfer the bit lengths */
var HEAP_SIZE     = 2*L_CODES + 1;
/* maximum heap size */
var MAX_BITS  = 15;
/* All codes must not exceed MAX_BITS bits */

var MIN_MATCH = 3;
var MAX_MATCH = 258;
var MIN_LOOKAHEAD = (MAX_MATCH + MIN_MATCH + 1);

var PRESET_DICT = 0x20;

var INIT_STATE = 42;
var EXTRA_STATE = 69;
var NAME_STATE = 73;
var COMMENT_STATE = 91;
var HCRC_STATE = 103;
var BUSY_STATE = 113;
var FINISH_STATE = 666;

var BS_NEED_MORE      = 1; /* block not completed, need more input or more output */
var BS_BLOCK_DONE     = 2; /* block flush performed */
var BS_FINISH_STARTED = 3; /* finish started, need only more output at next deflate */
var BS_FINISH_DONE    = 4; /* finish done, accept no more input or output */

var OS_CODE = 0x03; // Unix :) . Don't detect, use this default.

function err(strm, errorCode) {
  strm.msg = msg[errorCode];
  return errorCode;
}

function rank(f) {
  return ((f) << 1) - ((f) > 4 ? 9 : 0);
}

function zero(buf) { var len = buf.length; while (--len >= 0) { buf[len] = 0; } }


/* =========================================================================
 * Flush as much pending output as possible. All deflate() output goes
 * through this function so some applications may wish to modify it
 * to avoid allocating a large strm->output buffer and copying into it.
 * (See also read_buf()).
 */
function flush_pending(strm) {
  var s = strm.state;

  //_tr_flush_bits(s);
  var len = s.pending;
  if (len > strm.avail_out) {
    len = strm.avail_out;
  }
  if (len === 0) { return; }

  utils.arraySet(strm.output, s.pending_buf, s.pending_out, len, strm.next_out);
  strm.next_out += len;
  s.pending_out += len;
  strm.total_out += len;
  strm.avail_out -= len;
  s.pending -= len;
  if (s.pending === 0) {
    s.pending_out = 0;
  }
}


function flush_block_only (s, last) {
  trees._tr_flush_block(s, (s.block_start >= 0 ? s.block_start : -1), s.strstart - s.block_start, last);
  s.block_start = s.strstart;
  flush_pending(s.strm);
}


function put_byte(s, b) {
  s.pending_buf[s.pending++] = b;
}


/* =========================================================================
 * Put a short in the pending buffer. The 16-bit value is put in MSB order.
 * IN assertion: the stream state is correct and there is enough room in
 * pending_buf.
 */
function putShortMSB(s, b) {
//  put_byte(s, (Byte)(b >> 8));
//  put_byte(s, (Byte)(b & 0xff));
  s.pending_buf[s.pending++] = (b >>> 8) & 0xff;
  s.pending_buf[s.pending++] = b & 0xff;
}


/* ===========================================================================
 * Read a new buffer from the current input stream, update the adler32
 * and total number of bytes read.  All deflate() input goes through
 * this function so some applications may wish to modify it to avoid
 * allocating a large strm->input buffer and copying from it.
 * (See also flush_pending()).
 */
function read_buf(strm, buf, start, size) {
  var len = strm.avail_in;

  if (len > size) { len = size; }
  if (len === 0) { return 0; }

  strm.avail_in -= len;

  utils.arraySet(buf, strm.input, strm.next_in, len, start);
  if (strm.state.wrap === 1) {
    strm.adler = adler32(strm.adler, buf, len, start);
  }

  else if (strm.state.wrap === 2) {
    strm.adler = crc32(strm.adler, buf, len, start);
  }

  strm.next_in += len;
  strm.total_in += len;

  return len;
}


/* ===========================================================================
 * Set match_start to the longest match starting at the given string and
 * return its length. Matches shorter or equal to prev_length are discarded,
 * in which case the result is equal to prev_length and match_start is
 * garbage.
 * IN assertions: cur_match is the head of the hash chain for the current
 *   string (strstart) and its distance is <= MAX_DIST, and prev_length >= 1
 * OUT assertion: the match length is not greater than s->lookahead.
 */
function longest_match(s, cur_match) {
  var chain_length = s.max_chain_length;      /* max hash chain length */
  var scan = s.strstart; /* current string */
  var match;                       /* matched string */
  var len;                           /* length of current match */
  var best_len = s.prev_length;              /* best match length so far */
  var nice_match = s.nice_match;             /* stop if match long enough */
  var limit = (s.strstart > (s.w_size - MIN_LOOKAHEAD)) ?
      s.strstart - (s.w_size - MIN_LOOKAHEAD) : 0/*NIL*/;

  var _win = s.window; // shortcut

  var wmask = s.w_mask;
  var prev  = s.prev;

  /* Stop when cur_match becomes <= limit. To simplify the code,
   * we prevent matches with the string of window index 0.
   */

  var strend = s.strstart + MAX_MATCH;
  var scan_end1  = _win[scan + best_len - 1];
  var scan_end   = _win[scan + best_len];

  /* The code is optimized for HASH_BITS >= 8 and MAX_MATCH-2 multiple of 16.
   * It is easy to get rid of this optimization if necessary.
   */
  // Assert(s->hash_bits >= 8 && MAX_MATCH == 258, "Code too clever");

  /* Do not waste too much time if we already have a good match: */
  if (s.prev_length >= s.good_match) {
    chain_length >>= 2;
  }
  /* Do not look for matches beyond the end of the input. This is necessary
   * to make deflate deterministic.
   */
  if (nice_match > s.lookahead) { nice_match = s.lookahead; }

  // Assert((ulg)s->strstart <= s->window_size-MIN_LOOKAHEAD, "need lookahead");

  do {
    // Assert(cur_match < s->strstart, "no future");
    match = cur_match;

    /* Skip to next match if the match length cannot increase
     * or if the match length is less than 2.  Note that the checks below
     * for insufficient lookahead only occur occasionally for performance
     * reasons.  Therefore uninitialized memory will be accessed, and
     * conditional jumps will be made that depend on those values.
     * However the length of the match is limited to the lookahead, so
     * the output of deflate is not affected by the uninitialized values.
     */

    if (_win[match + best_len]     !== scan_end  ||
        _win[match + best_len - 1] !== scan_end1 ||
        _win[match]                !== _win[scan] ||
        _win[++match]              !== _win[scan + 1]) {
      continue;
    }

    /* The check at best_len-1 can be removed because it will be made
     * again later. (This heuristic is not always a win.)
     * It is not necessary to compare scan[2] and match[2] since they
     * are always equal when the other bytes match, given that
     * the hash keys are equal and that HASH_BITS >= 8.
     */
    scan += 2;
    match++;
    // Assert(*scan == *match, "match[2]?");

    /* We check for insufficient lookahead only every 8th comparison;
     * the 256th check will be made at strstart+258.
     */
    do {
      /*jshint noempty:false*/
    } while (_win[++scan] === _win[++match] && _win[++scan] === _win[++match] &&
             _win[++scan] === _win[++match] && _win[++scan] === _win[++match] &&
             _win[++scan] === _win[++match] && _win[++scan] === _win[++match] &&
             _win[++scan] === _win[++match] && _win[++scan] === _win[++match] &&
             scan < strend);

    // Assert(scan <= s->window+(unsigned)(s->window_size-1), "wild scan");

    len = MAX_MATCH - (strend - scan);
    scan = strend - MAX_MATCH;

    if (len > best_len) {
      s.match_start = cur_match;
      best_len = len;
      if (len >= nice_match) {
        break;
      }
      scan_end1  = _win[scan + best_len - 1];
      scan_end   = _win[scan + best_len];
    }
  } while ((cur_match = prev[cur_match & wmask]) > limit && --chain_length !== 0);

  if (best_len <= s.lookahead) {
    return best_len;
  }
  return s.lookahead;
}


/* ===========================================================================
 * Fill the window when the lookahead becomes insufficient.
 * Updates strstart and lookahead.
 *
 * IN assertion: lookahead < MIN_LOOKAHEAD
 * OUT assertions: strstart <= window_size-MIN_LOOKAHEAD
 *    At least one byte has been read, or avail_in == 0; reads are
 *    performed for at least two bytes (required for the zip translate_eol
 *    option -- not supported here).
 */
function fill_window(s) {
  var _w_size = s.w_size;
  var p, n, m, more, str;

  //Assert(s->lookahead < MIN_LOOKAHEAD, "already enough lookahead");

  do {
    more = s.window_size - s.lookahead - s.strstart;

    // JS ints have 32 bit, block below not needed
    /* Deal with !@#$% 64K limit: */
    //if (sizeof(int) <= 2) {
    //    if (more == 0 && s->strstart == 0 && s->lookahead == 0) {
    //        more = wsize;
    //
    //  } else if (more == (unsigned)(-1)) {
    //        /* Very unlikely, but possible on 16 bit machine if
    //         * strstart == 0 && lookahead == 1 (input done a byte at time)
    //         */
    //        more--;
    //    }
    //}


    /* If the window is almost full and there is insufficient lookahead,
     * move the upper half to the lower one to make room in the upper half.
     */
    if (s.strstart >= _w_size + (_w_size - MIN_LOOKAHEAD)) {

      utils.arraySet(s.window, s.window, _w_size, _w_size, 0);
      s.match_start -= _w_size;
      s.strstart -= _w_size;
      /* we now have strstart >= MAX_DIST */
      s.block_start -= _w_size;

      /* Slide the hash table (could be avoided with 32 bit values
       at the expense of memory usage). We slide even when level == 0
       to keep the hash table consistent if we switch back to level > 0
       later. (Using level 0 permanently is not an optimal usage of
       zlib, so we don't care about this pathological case.)
       */

      n = s.hash_size;
      p = n;
      do {
        m = s.head[--p];
        s.head[p] = (m >= _w_size ? m - _w_size : 0);
      } while (--n);

      n = _w_size;
      p = n;
      do {
        m = s.prev[--p];
        s.prev[p] = (m >= _w_size ? m - _w_size : 0);
        /* If n is not on any hash chain, prev[n] is garbage but
         * its value will never be used.
         */
      } while (--n);

      more += _w_size;
    }
    if (s.strm.avail_in === 0) {
      break;
    }

    /* If there was no sliding:
     *    strstart <= WSIZE+MAX_DIST-1 && lookahead <= MIN_LOOKAHEAD - 1 &&
     *    more == window_size - lookahead - strstart
     * => more >= window_size - (MIN_LOOKAHEAD-1 + WSIZE + MAX_DIST-1)
     * => more >= window_size - 2*WSIZE + 2
     * In the BIG_MEM or MMAP case (not yet supported),
     *   window_size == input_size + MIN_LOOKAHEAD  &&
     *   strstart + s->lookahead <= input_size => more >= MIN_LOOKAHEAD.
     * Otherwise, window_size == 2*WSIZE so more >= 2.
     * If there was sliding, more >= WSIZE. So in all cases, more >= 2.
     */
    //Assert(more >= 2, "more < 2");
    n = read_buf(s.strm, s.window, s.strstart + s.lookahead, more);
    s.lookahead += n;

    /* Initialize the hash value now that we have some input: */
    if (s.lookahead + s.insert >= MIN_MATCH) {
      str = s.strstart - s.insert;
      s.ins_h = s.window[str];

      /* UPDATE_HASH(s, s->ins_h, s->window[str + 1]); */
      s.ins_h = ((s.ins_h << s.hash_shift) ^ s.window[str + 1]) & s.hash_mask;
//#if MIN_MATCH != 3
//        Call update_hash() MIN_MATCH-3 more times
//#endif
      while (s.insert) {
        /* UPDATE_HASH(s, s->ins_h, s->window[str + MIN_MATCH-1]); */
        s.ins_h = ((s.ins_h << s.hash_shift) ^ s.window[str + MIN_MATCH-1]) & s.hash_mask;

        s.prev[str & s.w_mask] = s.head[s.ins_h];
        s.head[s.ins_h] = str;
        str++;
        s.insert--;
        if (s.lookahead + s.insert < MIN_MATCH) {
          break;
        }
      }
    }
    /* If the whole input has less than MIN_MATCH bytes, ins_h is garbage,
     * but this is not important since only literal bytes will be emitted.
     */

  } while (s.lookahead < MIN_LOOKAHEAD && s.strm.avail_in !== 0);

  /* If the WIN_INIT bytes after the end of the current data have never been
   * written, then zero those bytes in order to avoid memory check reports of
   * the use of uninitialized (or uninitialised as Julian writes) bytes by
   * the longest match routines.  Update the high water mark for the next
   * time through here.  WIN_INIT is set to MAX_MATCH since the longest match
   * routines allow scanning to strstart + MAX_MATCH, ignoring lookahead.
   */
//  if (s.high_water < s.window_size) {
//    var curr = s.strstart + s.lookahead;
//    var init = 0;
//
//    if (s.high_water < curr) {
//      /* Previous high water mark below current data -- zero WIN_INIT
//       * bytes or up to end of window, whichever is less.
//       */
//      init = s.window_size - curr;
//      if (init > WIN_INIT)
//        init = WIN_INIT;
//      zmemzero(s->window + curr, (unsigned)init);
//      s->high_water = curr + init;
//    }
//    else if (s->high_water < (ulg)curr + WIN_INIT) {
//      /* High water mark at or above current data, but below current data
//       * plus WIN_INIT -- zero out to current data plus WIN_INIT, or up
//       * to end of window, whichever is less.
//       */
//      init = (ulg)curr + WIN_INIT - s->high_water;
//      if (init > s->window_size - s->high_water)
//        init = s->window_size - s->high_water;
//      zmemzero(s->window + s->high_water, (unsigned)init);
//      s->high_water += init;
//    }
//  }
//
//  Assert((ulg)s->strstart <= s->window_size - MIN_LOOKAHEAD,
//    "not enough room for search");
}

/* ===========================================================================
 * Copy without compression as much as possible from the input stream, return
 * the current block state.
 * This function does not insert new strings in the dictionary since
 * uncompressible data is probably not useful. This function is used
 * only for the level=0 compression option.
 * NOTE: this function should be optimized to avoid extra copying from
 * window to pending_buf.
 */
function deflate_stored(s, flush) {
  /* Stored blocks are limited to 0xffff bytes, pending_buf is limited
   * to pending_buf_size, and each stored block has a 5 byte header:
   */
  var max_block_size = 0xffff;

  if (max_block_size > s.pending_buf_size - 5) {
    max_block_size = s.pending_buf_size - 5;
  }

  /* Copy as much as possible from input to output: */
  for (;;) {
    /* Fill the window as much as possible: */
    if (s.lookahead <= 1) {

      //Assert(s->strstart < s->w_size+MAX_DIST(s) ||
      //  s->block_start >= (long)s->w_size, "slide too late");
//      if (!(s.strstart < s.w_size + (s.w_size - MIN_LOOKAHEAD) ||
//        s.block_start >= s.w_size)) {
//        throw  new Error("slide too late");
//      }

      fill_window(s);
      if (s.lookahead === 0 && flush === Z_NO_FLUSH) {
        return BS_NEED_MORE;
      }

      if (s.lookahead === 0) {
        break;
      }
      /* flush the current block */
    }
    //Assert(s->block_start >= 0L, "block gone");
//    if (s.block_start < 0) throw new Error("block gone");

    s.strstart += s.lookahead;
    s.lookahead = 0;

    /* Emit a stored block if pending_buf will be full: */
    var max_start = s.block_start + max_block_size;

    if (s.strstart === 0 || s.strstart >= max_start) {
      /* strstart == 0 is possible when wraparound on 16-bit machine */
      s.lookahead = s.strstart - max_start;
      s.strstart = max_start;
      /*** FLUSH_BLOCK(s, 0); ***/
      flush_block_only(s, false);
      if (s.strm.avail_out === 0) {
        return BS_NEED_MORE;
      }
      /***/


    }
    /* Flush if we may have to slide, otherwise block_start may become
     * negative and the data will be gone:
     */
    if (s.strstart - s.block_start >= (s.w_size - MIN_LOOKAHEAD)) {
      /*** FLUSH_BLOCK(s, 0); ***/
      flush_block_only(s, false);
      if (s.strm.avail_out === 0) {
        return BS_NEED_MORE;
      }
      /***/
    }
  }

  s.insert = 0;

  if (flush === Z_FINISH) {
    /*** FLUSH_BLOCK(s, 1); ***/
    flush_block_only(s, true);
    if (s.strm.avail_out === 0) {
      return BS_FINISH_STARTED;
    }
    /***/
    return BS_FINISH_DONE;
  }

  if (s.strstart > s.block_start) {
    /*** FLUSH_BLOCK(s, 0); ***/
    flush_block_only(s, false);
    if (s.strm.avail_out === 0) {
      return BS_NEED_MORE;
    }
    /***/
  }

  return BS_NEED_MORE;
}

/* ===========================================================================
 * Compress as much as possible from the input stream, return the current
 * block state.
 * This function does not perform lazy evaluation of matches and inserts
 * new strings in the dictionary only for unmatched strings or for short
 * matches. It is used only for the fast compression options.
 */
function deflate_fast(s, flush) {
  var hash_head;        /* head of the hash chain */
  var bflush;           /* set if current block must be flushed */

  for (;;) {
    /* Make sure that we always have enough lookahead, except
     * at the end of the input file. We need MAX_MATCH bytes
     * for the next match, plus MIN_MATCH bytes to insert the
     * string following the next match.
     */
    if (s.lookahead < MIN_LOOKAHEAD) {
      fill_window(s);
      if (s.lookahead < MIN_LOOKAHEAD && flush === Z_NO_FLUSH) {
        return BS_NEED_MORE;
      }
      if (s.lookahead === 0) {
        break; /* flush the current block */
      }
    }

    /* Insert the string window[strstart .. strstart+2] in the
     * dictionary, and set hash_head to the head of the hash chain:
     */
    hash_head = 0/*NIL*/;
    if (s.lookahead >= MIN_MATCH) {
      /*** INSERT_STRING(s, s.strstart, hash_head); ***/
      s.ins_h = ((s.ins_h << s.hash_shift) ^ s.window[s.strstart + MIN_MATCH - 1]) & s.hash_mask;
      hash_head = s.prev[s.strstart & s.w_mask] = s.head[s.ins_h];
      s.head[s.ins_h] = s.strstart;
      /***/
    }

    /* Find the longest match, discarding those <= prev_length.
     * At this point we have always match_length < MIN_MATCH
     */
    if (hash_head !== 0/*NIL*/ && ((s.strstart - hash_head) <= (s.w_size - MIN_LOOKAHEAD))) {
      /* To simplify the code, we prevent matches with the string
       * of window index 0 (in particular we have to avoid a match
       * of the string with itself at the start of the input file).
       */
      s.match_length = longest_match(s, hash_head);
      /* longest_match() sets match_start */
    }
    if (s.match_length >= MIN_MATCH) {
      // check_match(s, s.strstart, s.match_start, s.match_length); // for debug only

      /*** _tr_tally_dist(s, s.strstart - s.match_start,
                     s.match_length - MIN_MATCH, bflush); ***/
      bflush = trees._tr_tally(s, s.strstart - s.match_start, s.match_length - MIN_MATCH);

      s.lookahead -= s.match_length;

      /* Insert new strings in the hash table only if the match length
       * is not too large. This saves time but degrades compression.
       */
      if (s.match_length <= s.max_lazy_match/*max_insert_length*/ && s.lookahead >= MIN_MATCH) {
        s.match_length--; /* string at strstart already in table */
        do {
          s.strstart++;
          /*** INSERT_STRING(s, s.strstart, hash_head); ***/
          s.ins_h = ((s.ins_h << s.hash_shift) ^ s.window[s.strstart + MIN_MATCH - 1]) & s.hash_mask;
          hash_head = s.prev[s.strstart & s.w_mask] = s.head[s.ins_h];
          s.head[s.ins_h] = s.strstart;
          /***/
          /* strstart never exceeds WSIZE-MAX_MATCH, so there are
           * always MIN_MATCH bytes ahead.
           */
        } while (--s.match_length !== 0);
        s.strstart++;
      } else
      {
        s.strstart += s.match_length;
        s.match_length = 0;
        s.ins_h = s.window[s.strstart];
        /* UPDATE_HASH(s, s.ins_h, s.window[s.strstart+1]); */
        s.ins_h = ((s.ins_h << s.hash_shift) ^ s.window[s.strstart + 1]) & s.hash_mask;

//#if MIN_MATCH != 3
//                Call UPDATE_HASH() MIN_MATCH-3 more times
//#endif
        /* If lookahead < MIN_MATCH, ins_h is garbage, but it does not
         * matter since it will be recomputed at next deflate call.
         */
      }
    } else {
      /* No match, output a literal byte */
      //Tracevv((stderr,"%c", s.window[s.strstart]));
      /*** _tr_tally_lit(s, s.window[s.strstart], bflush); ***/
      bflush = trees._tr_tally(s, 0, s.window[s.strstart]);

      s.lookahead--;
      s.strstart++;
    }
    if (bflush) {
      /*** FLUSH_BLOCK(s, 0); ***/
      flush_block_only(s, false);
      if (s.strm.avail_out === 0) {
        return BS_NEED_MORE;
      }
      /***/
    }
  }
  s.insert = ((s.strstart < (MIN_MATCH-1)) ? s.strstart : MIN_MATCH-1);
  if (flush === Z_FINISH) {
    /*** FLUSH_BLOCK(s, 1); ***/
    flush_block_only(s, true);
    if (s.strm.avail_out === 0) {
      return BS_FINISH_STARTED;
    }
    /***/
    return BS_FINISH_DONE;
  }
  if (s.last_lit) {
    /*** FLUSH_BLOCK(s, 0); ***/
    flush_block_only(s, false);
    if (s.strm.avail_out === 0) {
      return BS_NEED_MORE;
    }
    /***/
  }
  return BS_BLOCK_DONE;
}

/* ===========================================================================
 * Same as above, but achieves better compression. We use a lazy
 * evaluation for matches: a match is finally adopted only if there is
 * no better match at the next window position.
 */
function deflate_slow(s, flush) {
  var hash_head;          /* head of hash chain */
  var bflush;              /* set if current block must be flushed */

  var max_insert;

  /* Process the input block. */
  for (;;) {
    /* Make sure that we always have enough lookahead, except
     * at the end of the input file. We need MAX_MATCH bytes
     * for the next match, plus MIN_MATCH bytes to insert the
     * string following the next match.
     */
    if (s.lookahead < MIN_LOOKAHEAD) {
      fill_window(s);
      if (s.lookahead < MIN_LOOKAHEAD && flush === Z_NO_FLUSH) {
        return BS_NEED_MORE;
      }
      if (s.lookahead === 0) { break; } /* flush the current block */
    }

    /* Insert the string window[strstart .. strstart+2] in the
     * dictionary, and set hash_head to the head of the hash chain:
     */
    hash_head = 0/*NIL*/;
    if (s.lookahead >= MIN_MATCH) {
      /*** INSERT_STRING(s, s.strstart, hash_head); ***/
      s.ins_h = ((s.ins_h << s.hash_shift) ^ s.window[s.strstart + MIN_MATCH - 1]) & s.hash_mask;
      hash_head = s.prev[s.strstart & s.w_mask] = s.head[s.ins_h];
      s.head[s.ins_h] = s.strstart;
      /***/
    }

    /* Find the longest match, discarding those <= prev_length.
     */
    s.prev_length = s.match_length;
    s.prev_match = s.match_start;
    s.match_length = MIN_MATCH-1;

    if (hash_head !== 0/*NIL*/ && s.prev_length < s.max_lazy_match &&
        s.strstart - hash_head <= (s.w_size-MIN_LOOKAHEAD)/*MAX_DIST(s)*/) {
      /* To simplify the code, we prevent matches with the string
       * of window index 0 (in particular we have to avoid a match
       * of the string with itself at the start of the input file).
       */
      s.match_length = longest_match(s, hash_head);
      /* longest_match() sets match_start */

      if (s.match_length <= 5 &&
         (s.strategy === Z_FILTERED || (s.match_length === MIN_MATCH && s.strstart - s.match_start > 4096/*TOO_FAR*/))) {

        /* If prev_match is also MIN_MATCH, match_start is garbage
         * but we will ignore the current match anyway.
         */
        s.match_length = MIN_MATCH-1;
      }
    }
    /* If there was a match at the previous step and the current
     * match is not better, output the previous match:
     */
    if (s.prev_length >= MIN_MATCH && s.match_length <= s.prev_length) {
      max_insert = s.strstart + s.lookahead - MIN_MATCH;
      /* Do not insert strings in hash table beyond this. */

      //check_match(s, s.strstart-1, s.prev_match, s.prev_length);

      /***_tr_tally_dist(s, s.strstart - 1 - s.prev_match,
                     s.prev_length - MIN_MATCH, bflush);***/
      bflush = trees._tr_tally(s, s.strstart - 1- s.prev_match, s.prev_length - MIN_MATCH);
      /* Insert in hash table all strings up to the end of the match.
       * strstart-1 and strstart are already inserted. If there is not
       * enough lookahead, the last two strings are not inserted in
       * the hash table.
       */
      s.lookahead -= s.prev_length-1;
      s.prev_length -= 2;
      do {
        if (++s.strstart <= max_insert) {
          /*** INSERT_STRING(s, s.strstart, hash_head); ***/
          s.ins_h = ((s.ins_h << s.hash_shift) ^ s.window[s.strstart + MIN_MATCH - 1]) & s.hash_mask;
          hash_head = s.prev[s.strstart & s.w_mask] = s.head[s.ins_h];
          s.head[s.ins_h] = s.strstart;
          /***/
        }
      } while (--s.prev_length !== 0);
      s.match_available = 0;
      s.match_length = MIN_MATCH-1;
      s.strstart++;

      if (bflush) {
        /*** FLUSH_BLOCK(s, 0); ***/
        flush_block_only(s, false);
        if (s.strm.avail_out === 0) {
          return BS_NEED_MORE;
        }
        /***/
      }

    } else if (s.match_available) {
      /* If there was no match at the previous position, output a
       * single literal. If there was a match but the current match
       * is longer, truncate the previous match to a single literal.
       */
      //Tracevv((stderr,"%c", s->window[s->strstart-1]));
      /*** _tr_tally_lit(s, s.window[s.strstart-1], bflush); ***/
      bflush = trees._tr_tally(s, 0, s.window[s.strstart-1]);

      if (bflush) {
        /*** FLUSH_BLOCK_ONLY(s, 0) ***/
        flush_block_only(s, false);
        /***/
      }
      s.strstart++;
      s.lookahead--;
      if (s.strm.avail_out === 0) {
        return BS_NEED_MORE;
      }
    } else {
      /* There is no previous match to compare with, wait for
       * the next step to decide.
       */
      s.match_available = 1;
      s.strstart++;
      s.lookahead--;
    }
  }
  //Assert (flush != Z_NO_FLUSH, "no flush?");
  if (s.match_available) {
    //Tracevv((stderr,"%c", s->window[s->strstart-1]));
    /*** _tr_tally_lit(s, s.window[s.strstart-1], bflush); ***/
    bflush = trees._tr_tally(s, 0, s.window[s.strstart-1]);

    s.match_available = 0;
  }
  s.insert = s.strstart < MIN_MATCH-1 ? s.strstart : MIN_MATCH-1;
  if (flush === Z_FINISH) {
    /*** FLUSH_BLOCK(s, 1); ***/
    flush_block_only(s, true);
    if (s.strm.avail_out === 0) {
      return BS_FINISH_STARTED;
    }
    /***/
    return BS_FINISH_DONE;
  }
  if (s.last_lit) {
    /*** FLUSH_BLOCK(s, 0); ***/
    flush_block_only(s, false);
    if (s.strm.avail_out === 0) {
      return BS_NEED_MORE;
    }
    /***/
  }

  return BS_BLOCK_DONE;
}


/* ===========================================================================
 * For Z_RLE, simply look for runs of bytes, generate matches only of distance
 * one.  Do not maintain a hash table.  (It will be regenerated if this run of
 * deflate switches away from Z_RLE.)
 */
function deflate_rle(s, flush) {
  var bflush;            /* set if current block must be flushed */
  var prev;              /* byte at distance one to match */
  var scan, strend;      /* scan goes up to strend for length of run */

  var _win = s.window;

  for (;;) {
    /* Make sure that we always have enough lookahead, except
     * at the end of the input file. We need MAX_MATCH bytes
     * for the longest run, plus one for the unrolled loop.
     */
    if (s.lookahead <= MAX_MATCH) {
      fill_window(s);
      if (s.lookahead <= MAX_MATCH && flush === Z_NO_FLUSH) {
        return BS_NEED_MORE;
      }
      if (s.lookahead === 0) { break; } /* flush the current block */
    }

    /* See how many times the previous byte repeats */
    s.match_length = 0;
    if (s.lookahead >= MIN_MATCH && s.strstart > 0) {
      scan = s.strstart - 1;
      prev = _win[scan];
      if (prev === _win[++scan] && prev === _win[++scan] && prev === _win[++scan]) {
        strend = s.strstart + MAX_MATCH;
        do {
          /*jshint noempty:false*/
        } while (prev === _win[++scan] && prev === _win[++scan] &&
                 prev === _win[++scan] && prev === _win[++scan] &&
                 prev === _win[++scan] && prev === _win[++scan] &&
                 prev === _win[++scan] && prev === _win[++scan] &&
                 scan < strend);
        s.match_length = MAX_MATCH - (strend - scan);
        if (s.match_length > s.lookahead) {
          s.match_length = s.lookahead;
        }
      }
      //Assert(scan <= s->window+(uInt)(s->window_size-1), "wild scan");
    }

    /* Emit match if have run of MIN_MATCH or longer, else emit literal */
    if (s.match_length >= MIN_MATCH) {
      //check_match(s, s.strstart, s.strstart - 1, s.match_length);

      /*** _tr_tally_dist(s, 1, s.match_length - MIN_MATCH, bflush); ***/
      bflush = trees._tr_tally(s, 1, s.match_length - MIN_MATCH);

      s.lookahead -= s.match_length;
      s.strstart += s.match_length;
      s.match_length = 0;
    } else {
      /* No match, output a literal byte */
      //Tracevv((stderr,"%c", s->window[s->strstart]));
      /*** _tr_tally_lit(s, s.window[s.strstart], bflush); ***/
      bflush = trees._tr_tally(s, 0, s.window[s.strstart]);

      s.lookahead--;
      s.strstart++;
    }
    if (bflush) {
      /*** FLUSH_BLOCK(s, 0); ***/
      flush_block_only(s, false);
      if (s.strm.avail_out === 0) {
        return BS_NEED_MORE;
      }
      /***/
    }
  }
  s.insert = 0;
  if (flush === Z_FINISH) {
    /*** FLUSH_BLOCK(s, 1); ***/
    flush_block_only(s, true);
    if (s.strm.avail_out === 0) {
      return BS_FINISH_STARTED;
    }
    /***/
    return BS_FINISH_DONE;
  }
  if (s.last_lit) {
    /*** FLUSH_BLOCK(s, 0); ***/
    flush_block_only(s, false);
    if (s.strm.avail_out === 0) {
      return BS_NEED_MORE;
    }
    /***/
  }
  return BS_BLOCK_DONE;
}

/* ===========================================================================
 * For Z_HUFFMAN_ONLY, do not look for matches.  Do not maintain a hash table.
 * (It will be regenerated if this run of deflate switches away from Huffman.)
 */
function deflate_huff(s, flush) {
  var bflush;             /* set if current block must be flushed */

  for (;;) {
    /* Make sure that we have a literal to write. */
    if (s.lookahead === 0) {
      fill_window(s);
      if (s.lookahead === 0) {
        if (flush === Z_NO_FLUSH) {
          return BS_NEED_MORE;
        }
        break;      /* flush the current block */
      }
    }

    /* Output a literal byte */
    s.match_length = 0;
    //Tracevv((stderr,"%c", s->window[s->strstart]));
    /*** _tr_tally_lit(s, s.window[s.strstart], bflush); ***/
    bflush = trees._tr_tally(s, 0, s.window[s.strstart]);
    s.lookahead--;
    s.strstart++;
    if (bflush) {
      /*** FLUSH_BLOCK(s, 0); ***/
      flush_block_only(s, false);
      if (s.strm.avail_out === 0) {
        return BS_NEED_MORE;
      }
      /***/
    }
  }
  s.insert = 0;
  if (flush === Z_FINISH) {
    /*** FLUSH_BLOCK(s, 1); ***/
    flush_block_only(s, true);
    if (s.strm.avail_out === 0) {
      return BS_FINISH_STARTED;
    }
    /***/
    return BS_FINISH_DONE;
  }
  if (s.last_lit) {
    /*** FLUSH_BLOCK(s, 0); ***/
    flush_block_only(s, false);
    if (s.strm.avail_out === 0) {
      return BS_NEED_MORE;
    }
    /***/
  }
  return BS_BLOCK_DONE;
}

/* Values for max_lazy_match, good_match and max_chain_length, depending on
 * the desired pack level (0..9). The values given below have been tuned to
 * exclude worst case performance for pathological files. Better values may be
 * found for specific files.
 */
var Config = function (good_length, max_lazy, nice_length, max_chain, func) {
  this.good_length = good_length;
  this.max_lazy = max_lazy;
  this.nice_length = nice_length;
  this.max_chain = max_chain;
  this.func = func;
};

var configuration_table;

configuration_table = [
  /*      good lazy nice chain */
  new Config(0, 0, 0, 0, deflate_stored),          /* 0 store only */
  new Config(4, 4, 8, 4, deflate_fast),            /* 1 max speed, no lazy matches */
  new Config(4, 5, 16, 8, deflate_fast),           /* 2 */
  new Config(4, 6, 32, 32, deflate_fast),          /* 3 */

  new Config(4, 4, 16, 16, deflate_slow),          /* 4 lazy matches */
  new Config(8, 16, 32, 32, deflate_slow),         /* 5 */
  new Config(8, 16, 128, 128, deflate_slow),       /* 6 */
  new Config(8, 32, 128, 256, deflate_slow),       /* 7 */
  new Config(32, 128, 258, 1024, deflate_slow),    /* 8 */
  new Config(32, 258, 258, 4096, deflate_slow)     /* 9 max compression */
];


/* ===========================================================================
 * Initialize the "longest match" routines for a new zlib stream
 */
function lm_init(s) {
  s.window_size = 2 * s.w_size;

  /*** CLEAR_HASH(s); ***/
  zero(s.head); // Fill with NIL (= 0);

  /* Set the default configuration parameters:
   */
  s.max_lazy_match = configuration_table[s.level].max_lazy;
  s.good_match = configuration_table[s.level].good_length;
  s.nice_match = configuration_table[s.level].nice_length;
  s.max_chain_length = configuration_table[s.level].max_chain;

  s.strstart = 0;
  s.block_start = 0;
  s.lookahead = 0;
  s.insert = 0;
  s.match_length = s.prev_length = MIN_MATCH - 1;
  s.match_available = 0;
  s.ins_h = 0;
}


function DeflateState() {
  this.strm = null;            /* pointer back to this zlib stream */
  this.status = 0;            /* as the name implies */
  this.pending_buf = null;      /* output still pending */
  this.pending_buf_size = 0;  /* size of pending_buf */
  this.pending_out = 0;       /* next pending byte to output to the stream */
  this.pending = 0;           /* nb of bytes in the pending buffer */
  this.wrap = 0;              /* bit 0 true for zlib, bit 1 true for gzip */
  this.gzhead = null;         /* gzip header information to write */
  this.gzindex = 0;           /* where in extra, name, or comment */
  this.method = Z_DEFLATED; /* can only be DEFLATED */
  this.last_flush = -1;   /* value of flush param for previous deflate call */

  this.w_size = 0;  /* LZ77 window size (32K by default) */
  this.w_bits = 0;  /* log2(w_size)  (8..16) */
  this.w_mask = 0;  /* w_size - 1 */

  this.window = null;
  /* Sliding window. Input bytes are read into the second half of the window,
   * and move to the first half later to keep a dictionary of at least wSize
   * bytes. With this organization, matches are limited to a distance of
   * wSize-MAX_MATCH bytes, but this ensures that IO is always
   * performed with a length multiple of the block size.
   */

  this.window_size = 0;
  /* Actual size of window: 2*wSize, except when the user input buffer
   * is directly used as sliding window.
   */

  this.prev = null;
  /* Link to older string with same hash index. To limit the size of this
   * array to 64K, this link is maintained only for the last 32K strings.
   * An index in this array is thus a window index modulo 32K.
   */

  this.head = null;   /* Heads of the hash chains or NIL. */

  this.ins_h = 0;       /* hash index of string to be inserted */
  this.hash_size = 0;   /* number of elements in hash table */
  this.hash_bits = 0;   /* log2(hash_size) */
  this.hash_mask = 0;   /* hash_size-1 */

  this.hash_shift = 0;
  /* Number of bits by which ins_h must be shifted at each input
   * step. It must be such that after MIN_MATCH steps, the oldest
   * byte no longer takes part in the hash key, that is:
   *   hash_shift * MIN_MATCH >= hash_bits
   */

  this.block_start = 0;
  /* Window position at the beginning of the current output block. Gets
   * negative when the window is moved backwards.
   */

  this.match_length = 0;      /* length of best match */
  this.prev_match = 0;        /* previous match */
  this.match_available = 0;   /* set if previous match exists */
  this.strstart = 0;          /* start of string to insert */
  this.match_start = 0;       /* start of matching string */
  this.lookahead = 0;         /* number of valid bytes ahead in window */

  this.prev_length = 0;
  /* Length of the best match at previous step. Matches not greater than this
   * are discarded. This is used in the lazy match evaluation.
   */

  this.max_chain_length = 0;
  /* To speed up deflation, hash chains are never searched beyond this
   * length.  A higher limit improves compression ratio but degrades the
   * speed.
   */

  this.max_lazy_match = 0;
  /* Attempt to find a better match only when the current match is strictly
   * smaller than this value. This mechanism is used only for compression
   * levels >= 4.
   */
  // That's alias to max_lazy_match, don't use directly
  //this.max_insert_length = 0;
  /* Insert new strings in the hash table only if the match length is not
   * greater than this length. This saves time but degrades compression.
   * max_insert_length is used only for compression levels <= 3.
   */

  this.level = 0;     /* compression level (1..9) */
  this.strategy = 0;  /* favor or force Huffman coding*/

  this.good_match = 0;
  /* Use a faster search when the previous match is longer than this */

  this.nice_match = 0; /* Stop searching when current match exceeds this */

              /* used by trees.c: */

  /* Didn't use ct_data typedef below to suppress compiler warning */

  // struct ct_data_s dyn_ltree[HEAP_SIZE];   /* literal and length tree */
  // struct ct_data_s dyn_dtree[2*D_CODES+1]; /* distance tree */
  // struct ct_data_s bl_tree[2*BL_CODES+1];  /* Huffman tree for bit lengths */

  // Use flat array of DOUBLE size, with interleaved fata,
  // because JS does not support effective
  this.dyn_ltree  = new utils.Buf16(HEAP_SIZE * 2);
  this.dyn_dtree  = new utils.Buf16((2*D_CODES+1) * 2);
  this.bl_tree    = new utils.Buf16((2*BL_CODES+1) * 2);
  zero(this.dyn_ltree);
  zero(this.dyn_dtree);
  zero(this.bl_tree);

  this.l_desc   = null;         /* desc. for literal tree */
  this.d_desc   = null;         /* desc. for distance tree */
  this.bl_desc  = null;         /* desc. for bit length tree */

  //ush bl_count[MAX_BITS+1];
  this.bl_count = new utils.Buf16(MAX_BITS+1);
  /* number of codes at each bit length for an optimal tree */

  //int heap[2*L_CODES+1];      /* heap used to build the Huffman trees */
  this.heap = new utils.Buf16(2*L_CODES+1);  /* heap used to build the Huffman trees */
  zero(this.heap);

  this.heap_len = 0;               /* number of elements in the heap */
  this.heap_max = 0;               /* element of largest frequency */
  /* The sons of heap[n] are heap[2*n] and heap[2*n+1]. heap[0] is not used.
   * The same heap array is used to build all trees.
   */

  this.depth = new utils.Buf16(2*L_CODES+1); //uch depth[2*L_CODES+1];
  zero(this.depth);
  /* Depth of each subtree used as tie breaker for trees of equal frequency
   */

  this.l_buf = 0;          /* buffer index for literals or lengths */

  this.lit_bufsize = 0;
  /* Size of match buffer for literals/lengths.  There are 4 reasons for
   * limiting lit_bufsize to 64K:
   *   - frequencies can be kept in 16 bit counters
   *   - if compression is not successful for the first block, all input
   *     data is still in the window so we can still emit a stored block even
   *     when input comes from standard input.  (This can also be done for
   *     all blocks if lit_bufsize is not greater than 32K.)
   *   - if compression is not successful for a file smaller than 64K, we can
   *     even emit a stored file instead of a stored block (saving 5 bytes).
   *     This is applicable only for zip (not gzip or zlib).
   *   - creating new Huffman trees less frequently may not provide fast
   *     adaptation to changes in the input data statistics. (Take for
   *     example a binary file with poorly compressible code followed by
   *     a highly compressible string table.) Smaller buffer sizes give
   *     fast adaptation but have of course the overhead of transmitting
   *     trees more frequently.
   *   - I can't count above 4
   */

  this.last_lit = 0;      /* running index in l_buf */

  this.d_buf = 0;
  /* Buffer index for distances. To simplify the code, d_buf and l_buf have
   * the same number of elements. To use different lengths, an extra flag
   * array would be necessary.
   */

  this.opt_len = 0;       /* bit length of current block with optimal trees */
  this.static_len = 0;    /* bit length of current block with static trees */
  this.matches = 0;       /* number of string matches in current block */
  this.insert = 0;        /* bytes at end of window left to insert */


  this.bi_buf = 0;
  /* Output buffer. bits are inserted starting at the bottom (least
   * significant bits).
   */
  this.bi_valid = 0;
  /* Number of valid bits in bi_buf.  All bits above the last valid bit
   * are always zero.
   */

  // Used for window memory init. We safely ignore it for JS. That makes
  // sense only for pointers and memory check tools.
  //this.high_water = 0;
  /* High water mark offset in window for initialized bytes -- bytes above
   * this are set to zero in order to avoid memory check warnings when
   * longest match routines access bytes past the input.  This is then
   * updated to the new high water mark.
   */
}


function deflateResetKeep(strm) {
  var s;

  if (!strm || !strm.state) {
    return err(strm, Z_STREAM_ERROR);
  }

  strm.total_in = strm.total_out = 0;
  strm.data_type = Z_UNKNOWN;

  s = strm.state;
  s.pending = 0;
  s.pending_out = 0;

  if (s.wrap < 0) {
    s.wrap = -s.wrap;
    /* was made negative by deflate(..., Z_FINISH); */
  }
  s.status = (s.wrap ? INIT_STATE : BUSY_STATE);
  strm.adler = (s.wrap === 2) ?
    0  // crc32(0, Z_NULL, 0)
  :
    1; // adler32(0, Z_NULL, 0)
  s.last_flush = Z_NO_FLUSH;
  trees._tr_init(s);
  return Z_OK;
}


function deflateReset(strm) {
  var ret = deflateResetKeep(strm);
  if (ret === Z_OK) {
    lm_init(strm.state);
  }
  return ret;
}


function deflateSetHeader(strm, head) {
  if (!strm || !strm.state) { return Z_STREAM_ERROR; }
  if (strm.state.wrap !== 2) { return Z_STREAM_ERROR; }
  strm.state.gzhead = head;
  return Z_OK;
}


function deflateInit2(strm, level, method, windowBits, memLevel, strategy) {
  if (!strm) { // === Z_NULL
    return Z_STREAM_ERROR;
  }
  var wrap = 1;

  if (level === Z_DEFAULT_COMPRESSION) {
    level = 6;
  }

  if (windowBits < 0) { /* suppress zlib wrapper */
    wrap = 0;
    windowBits = -windowBits;
  }

  else if (windowBits > 15) {
    wrap = 2;           /* write gzip wrapper instead */
    windowBits -= 16;
  }


  if (memLevel < 1 || memLevel > MAX_MEM_LEVEL || method !== Z_DEFLATED ||
    windowBits < 8 || windowBits > 15 || level < 0 || level > 9 ||
    strategy < 0 || strategy > Z_FIXED) {
    return err(strm, Z_STREAM_ERROR);
  }


  if (windowBits === 8) {
    windowBits = 9;
  }
  /* until 256-byte window bug fixed */

  var s = new DeflateState();

  strm.state = s;
  s.strm = strm;

  s.wrap = wrap;
  s.gzhead = null;
  s.w_bits = windowBits;
  s.w_size = 1 << s.w_bits;
  s.w_mask = s.w_size - 1;

  s.hash_bits = memLevel + 7;
  s.hash_size = 1 << s.hash_bits;
  s.hash_mask = s.hash_size - 1;
  s.hash_shift = ~~((s.hash_bits + MIN_MATCH - 1) / MIN_MATCH);

  s.window = new utils.Buf8(s.w_size * 2);
  s.head = new utils.Buf16(s.hash_size);
  s.prev = new utils.Buf16(s.w_size);

  // Don't need mem init magic for JS.
  //s.high_water = 0;  /* nothing written to s->window yet */

  s.lit_bufsize = 1 << (memLevel + 6); /* 16K elements by default */

  s.pending_buf_size = s.lit_bufsize * 4;
  s.pending_buf = new utils.Buf8(s.pending_buf_size);

  s.d_buf = s.lit_bufsize >> 1;
  s.l_buf = (1 + 2) * s.lit_bufsize;

  s.level = level;
  s.strategy = strategy;
  s.method = method;

  return deflateReset(strm);
}

function deflateInit(strm, level) {
  return deflateInit2(strm, level, Z_DEFLATED, MAX_WBITS, DEF_MEM_LEVEL, Z_DEFAULT_STRATEGY);
}


function deflate(strm, flush) {
  var old_flush, s;
  var beg, val; // for gzip header write only

  if (!strm || !strm.state ||
    flush > Z_BLOCK || flush < 0) {
    return strm ? err(strm, Z_STREAM_ERROR) : Z_STREAM_ERROR;
  }

  s = strm.state;

  if (!strm.output ||
      (!strm.input && strm.avail_in !== 0) ||
      (s.status === FINISH_STATE && flush !== Z_FINISH)) {
    return err(strm, (strm.avail_out === 0) ? Z_BUF_ERROR : Z_STREAM_ERROR);
  }

  s.strm = strm; /* just in case */
  old_flush = s.last_flush;
  s.last_flush = flush;

  /* Write the header */
  if (s.status === INIT_STATE) {

    if (s.wrap === 2) { // GZIP header
      strm.adler = 0;  //crc32(0L, Z_NULL, 0);
      put_byte(s, 31);
      put_byte(s, 139);
      put_byte(s, 8);
      if (!s.gzhead) { // s->gzhead == Z_NULL
        put_byte(s, 0);
        put_byte(s, 0);
        put_byte(s, 0);
        put_byte(s, 0);
        put_byte(s, 0);
        put_byte(s, s.level === 9 ? 2 :
                    (s.strategy >= Z_HUFFMAN_ONLY || s.level < 2 ?
                     4 : 0));
        put_byte(s, OS_CODE);
        s.status = BUSY_STATE;
      }
      else {
        put_byte(s, (s.gzhead.text ? 1 : 0) +
                    (s.gzhead.hcrc ? 2 : 0) +
                    (!s.gzhead.extra ? 0 : 4) +
                    (!s.gzhead.name ? 0 : 8) +
                    (!s.gzhead.comment ? 0 : 16)
                );
        put_byte(s, s.gzhead.time & 0xff);
        put_byte(s, (s.gzhead.time >> 8) & 0xff);
        put_byte(s, (s.gzhead.time >> 16) & 0xff);
        put_byte(s, (s.gzhead.time >> 24) & 0xff);
        put_byte(s, s.level === 9 ? 2 :
                    (s.strategy >= Z_HUFFMAN_ONLY || s.level < 2 ?
                     4 : 0));
        put_byte(s, s.gzhead.os & 0xff);
        if (s.gzhead.extra && s.gzhead.extra.length) {
          put_byte(s, s.gzhead.extra.length & 0xff);
          put_byte(s, (s.gzhead.extra.length >> 8) & 0xff);
        }
        if (s.gzhead.hcrc) {
          strm.adler = crc32(strm.adler, s.pending_buf, s.pending, 0);
        }
        s.gzindex = 0;
        s.status = EXTRA_STATE;
      }
    }
    else // DEFLATE header
    {
      var header = (Z_DEFLATED + ((s.w_bits - 8) << 4)) << 8;
      var level_flags = -1;

      if (s.strategy >= Z_HUFFMAN_ONLY || s.level < 2) {
        level_flags = 0;
      } else if (s.level < 6) {
        level_flags = 1;
      } else if (s.level === 6) {
        level_flags = 2;
      } else {
        level_flags = 3;
      }
      header |= (level_flags << 6);
      if (s.strstart !== 0) { header |= PRESET_DICT; }
      header += 31 - (header % 31);

      s.status = BUSY_STATE;
      putShortMSB(s, header);

      /* Save the adler32 of the preset dictionary: */
      if (s.strstart !== 0) {
        putShortMSB(s, strm.adler >>> 16);
        putShortMSB(s, strm.adler & 0xffff);
      }
      strm.adler = 1; // adler32(0L, Z_NULL, 0);
    }
  }

//#ifdef GZIP
  if (s.status === EXTRA_STATE) {
    if (s.gzhead.extra/* != Z_NULL*/) {
      beg = s.pending;  /* start of bytes to update crc */

      while (s.gzindex < (s.gzhead.extra.length & 0xffff)) {
        if (s.pending === s.pending_buf_size) {
          if (s.gzhead.hcrc && s.pending > beg) {
            strm.adler = crc32(strm.adler, s.pending_buf, s.pending - beg, beg);
          }
          flush_pending(strm);
          beg = s.pending;
          if (s.pending === s.pending_buf_size) {
            break;
          }
        }
        put_byte(s, s.gzhead.extra[s.gzindex] & 0xff);
        s.gzindex++;
      }
      if (s.gzhead.hcrc && s.pending > beg) {
        strm.adler = crc32(strm.adler, s.pending_buf, s.pending - beg, beg);
      }
      if (s.gzindex === s.gzhead.extra.length) {
        s.gzindex = 0;
        s.status = NAME_STATE;
      }
    }
    else {
      s.status = NAME_STATE;
    }
  }
  if (s.status === NAME_STATE) {
    if (s.gzhead.name/* != Z_NULL*/) {
      beg = s.pending;  /* start of bytes to update crc */
      //int val;

      do {
        if (s.pending === s.pending_buf_size) {
          if (s.gzhead.hcrc && s.pending > beg) {
            strm.adler = crc32(strm.adler, s.pending_buf, s.pending - beg, beg);
          }
          flush_pending(strm);
          beg = s.pending;
          if (s.pending === s.pending_buf_size) {
            val = 1;
            break;
          }
        }
        // JS specific: little magic to add zero terminator to end of string
        if (s.gzindex < s.gzhead.name.length) {
          val = s.gzhead.name.charCodeAt(s.gzindex++) & 0xff;
        } else {
          val = 0;
        }
        put_byte(s, val);
      } while (val !== 0);

      if (s.gzhead.hcrc && s.pending > beg){
        strm.adler = crc32(strm.adler, s.pending_buf, s.pending - beg, beg);
      }
      if (val === 0) {
        s.gzindex = 0;
        s.status = COMMENT_STATE;
      }
    }
    else {
      s.status = COMMENT_STATE;
    }
  }
  if (s.status === COMMENT_STATE) {
    if (s.gzhead.comment/* != Z_NULL*/) {
      beg = s.pending;  /* start of bytes to update crc */
      //int val;

      do {
        if (s.pending === s.pending_buf_size) {
          if (s.gzhead.hcrc && s.pending > beg) {
            strm.adler = crc32(strm.adler, s.pending_buf, s.pending - beg, beg);
          }
          flush_pending(strm);
          beg = s.pending;
          if (s.pending === s.pending_buf_size) {
            val = 1;
            break;
          }
        }
        // JS specific: little magic to add zero terminator to end of string
        if (s.gzindex < s.gzhead.comment.length) {
          val = s.gzhead.comment.charCodeAt(s.gzindex++) & 0xff;
        } else {
          val = 0;
        }
        put_byte(s, val);
      } while (val !== 0);

      if (s.gzhead.hcrc && s.pending > beg) {
        strm.adler = crc32(strm.adler, s.pending_buf, s.pending - beg, beg);
      }
      if (val === 0) {
        s.status = HCRC_STATE;
      }
    }
    else {
      s.status = HCRC_STATE;
    }
  }
  if (s.status === HCRC_STATE) {
    if (s.gzhead.hcrc) {
      if (s.pending + 2 > s.pending_buf_size) {
        flush_pending(strm);
      }
      if (s.pending + 2 <= s.pending_buf_size) {
        put_byte(s, strm.adler & 0xff);
        put_byte(s, (strm.adler >> 8) & 0xff);
        strm.adler = 0; //crc32(0L, Z_NULL, 0);
        s.status = BUSY_STATE;
      }
    }
    else {
      s.status = BUSY_STATE;
    }
  }
//#endif

  /* Flush as much pending output as possible */
  if (s.pending !== 0) {
    flush_pending(strm);
    if (strm.avail_out === 0) {
      /* Since avail_out is 0, deflate will be called again with
       * more output space, but possibly with both pending and
       * avail_in equal to zero. There won't be anything to do,
       * but this is not an error situation so make sure we
       * return OK instead of BUF_ERROR at next call of deflate:
       */
      s.last_flush = -1;
      return Z_OK;
    }

    /* Make sure there is something to do and avoid duplicate consecutive
     * flushes. For repeated and useless calls with Z_FINISH, we keep
     * returning Z_STREAM_END instead of Z_BUF_ERROR.
     */
  } else if (strm.avail_in === 0 && rank(flush) <= rank(old_flush) &&
    flush !== Z_FINISH) {
    return err(strm, Z_BUF_ERROR);
  }

  /* User must not provide more input after the first FINISH: */
  if (s.status === FINISH_STATE && strm.avail_in !== 0) {
    return err(strm, Z_BUF_ERROR);
  }

  /* Start a new block or continue the current one.
   */
  if (strm.avail_in !== 0 || s.lookahead !== 0 ||
    (flush !== Z_NO_FLUSH && s.status !== FINISH_STATE)) {
    var bstate = (s.strategy === Z_HUFFMAN_ONLY) ? deflate_huff(s, flush) :
      (s.strategy === Z_RLE ? deflate_rle(s, flush) :
        configuration_table[s.level].func(s, flush));

    if (bstate === BS_FINISH_STARTED || bstate === BS_FINISH_DONE) {
      s.status = FINISH_STATE;
    }
    if (bstate === BS_NEED_MORE || bstate === BS_FINISH_STARTED) {
      if (strm.avail_out === 0) {
        s.last_flush = -1;
        /* avoid BUF_ERROR next call, see above */
      }
      return Z_OK;
      /* If flush != Z_NO_FLUSH && avail_out == 0, the next call
       * of deflate should use the same flush parameter to make sure
       * that the flush is complete. So we don't have to output an
       * empty block here, this will be done at next call. This also
       * ensures that for a very small output buffer, we emit at most
       * one empty block.
       */
    }
    if (bstate === BS_BLOCK_DONE) {
      if (flush === Z_PARTIAL_FLUSH) {
        trees._tr_align(s);
      }
      else if (flush !== Z_BLOCK) { /* FULL_FLUSH or SYNC_FLUSH */

        trees._tr_stored_block(s, 0, 0, false);
        /* For a full flush, this empty block will be recognized
         * as a special marker by inflate_sync().
         */
        if (flush === Z_FULL_FLUSH) {
          /*** CLEAR_HASH(s); ***/             /* forget history */
          zero(s.head); // Fill with NIL (= 0);

          if (s.lookahead === 0) {
            s.strstart = 0;
            s.block_start = 0;
            s.insert = 0;
          }
        }
      }
      flush_pending(strm);
      if (strm.avail_out === 0) {
        s.last_flush = -1; /* avoid BUF_ERROR at next call, see above */
        return Z_OK;
      }
    }
  }
  //Assert(strm->avail_out > 0, "bug2");
  //if (strm.avail_out <= 0) { throw new Error("bug2");}

  if (flush !== Z_FINISH) { return Z_OK; }
  if (s.wrap <= 0) { return Z_STREAM_END; }

  /* Write the trailer */
  if (s.wrap === 2) {
    put_byte(s, strm.adler & 0xff);
    put_byte(s, (strm.adler >> 8) & 0xff);
    put_byte(s, (strm.adler >> 16) & 0xff);
    put_byte(s, (strm.adler >> 24) & 0xff);
    put_byte(s, strm.total_in & 0xff);
    put_byte(s, (strm.total_in >> 8) & 0xff);
    put_byte(s, (strm.total_in >> 16) & 0xff);
    put_byte(s, (strm.total_in >> 24) & 0xff);
  }
  else
  {
    putShortMSB(s, strm.adler >>> 16);
    putShortMSB(s, strm.adler & 0xffff);
  }

  flush_pending(strm);
  /* If avail_out is zero, the application will call deflate again
   * to flush the rest.
   */
  if (s.wrap > 0) { s.wrap = -s.wrap; }
  /* write the trailer only once! */
  return s.pending !== 0 ? Z_OK : Z_STREAM_END;
}

function deflateEnd(strm) {
  var status;

  if (!strm/*== Z_NULL*/ || !strm.state/*== Z_NULL*/) {
    return Z_STREAM_ERROR;
  }

  status = strm.state.status;
  if (status !== INIT_STATE &&
    status !== EXTRA_STATE &&
    status !== NAME_STATE &&
    status !== COMMENT_STATE &&
    status !== HCRC_STATE &&
    status !== BUSY_STATE &&
    status !== FINISH_STATE
  ) {
    return err(strm, Z_STREAM_ERROR);
  }

  strm.state = null;

  return status === BUSY_STATE ? err(strm, Z_DATA_ERROR) : Z_OK;
}

/* =========================================================================
 * Copy the source state to the destination state
 */
//function deflateCopy(dest, source) {
//
//}

exports.deflateInit = deflateInit;
exports.deflateInit2 = deflateInit2;
exports.deflateReset = deflateReset;
exports.deflateResetKeep = deflateResetKeep;
exports.deflateSetHeader = deflateSetHeader;
exports.deflate = deflate;
exports.deflateEnd = deflateEnd;
exports.deflateInfo = 'pako deflate (from Nodeca project)';

/* Not implemented
exports.deflateBound = deflateBound;
exports.deflateCopy = deflateCopy;
exports.deflateSetDictionary = deflateSetDictionary;
exports.deflateParams = deflateParams;
exports.deflatePending = deflatePending;
exports.deflatePrime = deflatePrime;
exports.deflateTune = deflateTune;
*/
},{"../utils/common":27,"./adler32":29,"./crc32":31,"./messages":37,"./trees":38}],33:[function(_dereq_,module,exports){
'use strict';


function GZheader() {
  /* true if compressed data believed to be text */
  this.text       = 0;
  /* modification time */
  this.time       = 0;
  /* extra flags (not used when writing a gzip file) */
  this.xflags     = 0;
  /* operating system */
  this.os         = 0;
  /* pointer to extra field or Z_NULL if none */
  this.extra      = null;
  /* extra field length (valid if extra != Z_NULL) */
  this.extra_len  = 0; // Actually, we don't need it in JS,
                       // but leave for few code modifications

  //
  // Setup limits is not necessary because in js we should not preallocate memory
  // for inflate use constant limit in 65536 bytes
  //

  /* space at extra (only when reading header) */
  // this.extra_max  = 0;
  /* pointer to zero-terminated file name or Z_NULL */
  this.name       = '';
  /* space at name (only when reading header) */
  // this.name_max   = 0;
  /* pointer to zero-terminated comment or Z_NULL */
  this.comment    = '';
  /* space at comment (only when reading header) */
  // this.comm_max   = 0;
  /* true if there was or will be a header crc */
  this.hcrc       = 0;
  /* true when done reading gzip header (not used when writing a gzip file) */
  this.done       = false;
}

module.exports = GZheader;
},{}],34:[function(_dereq_,module,exports){
'use strict';

// See state defs from inflate.js
var BAD = 30;       /* got a data error -- remain here until reset */
var TYPE = 12;      /* i: waiting for type bits, including last-flag bit */

/*
   Decode literal, length, and distance codes and write out the resulting
   literal and match bytes until either not enough input or output is
   available, an end-of-block is encountered, or a data error is encountered.
   When large enough input and output buffers are supplied to inflate(), for
   example, a 16K input buffer and a 64K output buffer, more than 95% of the
   inflate execution time is spent in this routine.

   Entry assumptions:

        state.mode === LEN
        strm.avail_in >= 6
        strm.avail_out >= 258
        start >= strm.avail_out
        state.bits < 8

   On return, state.mode is one of:

        LEN -- ran out of enough output space or enough available input
        TYPE -- reached end of block code, inflate() to interpret next block
        BAD -- error in block data

   Notes:

    - The maximum input bits used by a length/distance pair is 15 bits for the
      length code, 5 bits for the length extra, 15 bits for the distance code,
      and 13 bits for the distance extra.  This totals 48 bits, or six bytes.
      Therefore if strm.avail_in >= 6, then there is enough input to avoid
      checking for available input while decoding.

    - The maximum bytes that a single length/distance pair can output is 258
      bytes, which is the maximum length that can be coded.  inflate_fast()
      requires strm.avail_out >= 258 for each loop to avoid checking for
      output space.
 */
module.exports = function inflate_fast(strm, start) {
  var state;
  var _in;                    /* local strm.input */
  var last;                   /* have enough input while in < last */
  var _out;                   /* local strm.output */
  var beg;                    /* inflate()'s initial strm.output */
  var end;                    /* while out < end, enough space available */
//#ifdef INFLATE_STRICT
  var dmax;                   /* maximum distance from zlib header */
//#endif
  var wsize;                  /* window size or zero if not using window */
  var whave;                  /* valid bytes in the window */
  var wnext;                  /* window write index */
  var window;                 /* allocated sliding window, if wsize != 0 */
  var hold;                   /* local strm.hold */
  var bits;                   /* local strm.bits */
  var lcode;                  /* local strm.lencode */
  var dcode;                  /* local strm.distcode */
  var lmask;                  /* mask for first level of length codes */
  var dmask;                  /* mask for first level of distance codes */
  var here;                   /* retrieved table entry */
  var op;                     /* code bits, operation, extra bits, or */
                              /*  window position, window bytes to copy */
  var len;                    /* match length, unused bytes */
  var dist;                   /* match distance */
  var from;                   /* where to copy match from */
  var from_source;


  var input, output; // JS specific, because we have no pointers

  /* copy state to local variables */
  state = strm.state;
  //here = state.here;
  _in = strm.next_in;
  input = strm.input;
  last = _in + (strm.avail_in - 5);
  _out = strm.next_out;
  output = strm.output;
  beg = _out - (start - strm.avail_out);
  end = _out + (strm.avail_out - 257);
//#ifdef INFLATE_STRICT
  dmax = state.dmax;
//#endif
  wsize = state.wsize;
  whave = state.whave;
  wnext = state.wnext;
  window = state.window;
  hold = state.hold;
  bits = state.bits;
  lcode = state.lencode;
  dcode = state.distcode;
  lmask = (1 << state.lenbits) - 1;
  dmask = (1 << state.distbits) - 1;


  /* decode literals and length/distances until end-of-block or not enough
     input data or output space */

  top:
  do {
    if (bits < 15) {
      hold += input[_in++] << bits;
      bits += 8;
      hold += input[_in++] << bits;
      bits += 8;
    }

    here = lcode[hold & lmask];

    dolen:
    for (;;) { // Goto emulation
      op = here >>> 24/*here.bits*/;
      hold >>>= op;
      bits -= op;
      op = (here >>> 16) & 0xff/*here.op*/;
      if (op === 0) {                          /* literal */
        //Tracevv((stderr, here.val >= 0x20 && here.val < 0x7f ?
        //        "inflate:         literal '%c'\n" :
        //        "inflate:         literal 0x%02x\n", here.val));
        output[_out++] = here & 0xffff/*here.val*/;
      }
      else if (op & 16) {                     /* length base */
        len = here & 0xffff/*here.val*/;
        op &= 15;                           /* number of extra bits */
        if (op) {
          if (bits < op) {
            hold += input[_in++] << bits;
            bits += 8;
          }
          len += hold & ((1 << op) - 1);
          hold >>>= op;
          bits -= op;
        }
        //Tracevv((stderr, "inflate:         length %u\n", len));
        if (bits < 15) {
          hold += input[_in++] << bits;
          bits += 8;
          hold += input[_in++] << bits;
          bits += 8;
        }
        here = dcode[hold & dmask];

        dodist:
        for (;;) { // goto emulation
          op = here >>> 24/*here.bits*/;
          hold >>>= op;
          bits -= op;
          op = (here >>> 16) & 0xff/*here.op*/;

          if (op & 16) {                      /* distance base */
            dist = here & 0xffff/*here.val*/;
            op &= 15;                       /* number of extra bits */
            if (bits < op) {
              hold += input[_in++] << bits;
              bits += 8;
              if (bits < op) {
                hold += input[_in++] << bits;
                bits += 8;
              }
            }
            dist += hold & ((1 << op) - 1);
//#ifdef INFLATE_STRICT
            if (dist > dmax) {
              strm.msg = 'invalid distance too far back';
              state.mode = BAD;
              break top;
            }
//#endif
            hold >>>= op;
            bits -= op;
            //Tracevv((stderr, "inflate:         distance %u\n", dist));
            op = _out - beg;                /* max distance in output */
            if (dist > op) {                /* see if copy from window */
              op = dist - op;               /* distance back in window */
              if (op > whave) {
                if (state.sane) {
                  strm.msg = 'invalid distance too far back';
                  state.mode = BAD;
                  break top;
                }

// (!) This block is disabled in zlib defailts,
// don't enable it for binary compatibility
//#ifdef INFLATE_ALLOW_INVALID_DISTANCE_TOOFAR_ARRR
//                if (len <= op - whave) {
//                  do {
//                    output[_out++] = 0;
//                  } while (--len);
//                  continue top;
//                }
//                len -= op - whave;
//                do {
//                  output[_out++] = 0;
//                } while (--op > whave);
//                if (op === 0) {
//                  from = _out - dist;
//                  do {
//                    output[_out++] = output[from++];
//                  } while (--len);
//                  continue top;
//                }
//#endif
              }
              from = 0; // window index
              from_source = window;
              if (wnext === 0) {           /* very common case */
                from += wsize - op;
                if (op < len) {         /* some from window */
                  len -= op;
                  do {
                    output[_out++] = window[from++];
                  } while (--op);
                  from = _out - dist;  /* rest from output */
                  from_source = output;
                }
              }
              else if (wnext < op) {      /* wrap around window */
                from += wsize + wnext - op;
                op -= wnext;
                if (op < len) {         /* some from end of window */
                  len -= op;
                  do {
                    output[_out++] = window[from++];
                  } while (--op);
                  from = 0;
                  if (wnext < len) {  /* some from start of window */
                    op = wnext;
                    len -= op;
                    do {
                      output[_out++] = window[from++];
                    } while (--op);
                    from = _out - dist;      /* rest from output */
                    from_source = output;
                  }
                }
              }
              else {                      /* contiguous in window */
                from += wnext - op;
                if (op < len) {         /* some from window */
                  len -= op;
                  do {
                    output[_out++] = window[from++];
                  } while (--op);
                  from = _out - dist;  /* rest from output */
                  from_source = output;
                }
              }
              while (len > 2) {
                output[_out++] = from_source[from++];
                output[_out++] = from_source[from++];
                output[_out++] = from_source[from++];
                len -= 3;
              }
              if (len) {
                output[_out++] = from_source[from++];
                if (len > 1) {
                  output[_out++] = from_source[from++];
                }
              }
            }
            else {
              from = _out - dist;          /* copy direct from output */
              do {                        /* minimum length is three */
                output[_out++] = output[from++];
                output[_out++] = output[from++];
                output[_out++] = output[from++];
                len -= 3;
              } while (len > 2);
              if (len) {
                output[_out++] = output[from++];
                if (len > 1) {
                  output[_out++] = output[from++];
                }
              }
            }
          }
          else if ((op & 64) === 0) {          /* 2nd level distance code */
            here = dcode[(here & 0xffff)/*here.val*/ + (hold & ((1 << op) - 1))];
            continue dodist;
          }
          else {
            strm.msg = 'invalid distance code';
            state.mode = BAD;
            break top;
          }

          break; // need to emulate goto via "continue"
        }
      }
      else if ((op & 64) === 0) {              /* 2nd level length code */
        here = lcode[(here & 0xffff)/*here.val*/ + (hold & ((1 << op) - 1))];
        continue dolen;
      }
      else if (op & 32) {                     /* end-of-block */
        //Tracevv((stderr, "inflate:         end of block\n"));
        state.mode = TYPE;
        break top;
      }
      else {
        strm.msg = 'invalid literal/length code';
        state.mode = BAD;
        break top;
      }

      break; // need to emulate goto via "continue"
    }
  } while (_in < last && _out < end);

  /* return unused bytes (on entry, bits < 8, so in won't go too far back) */
  len = bits >> 3;
  _in -= len;
  bits -= len << 3;
  hold &= (1 << bits) - 1;

  /* update state and return */
  strm.next_in = _in;
  strm.next_out = _out;
  strm.avail_in = (_in < last ? 5 + (last - _in) : 5 - (_in - last));
  strm.avail_out = (_out < end ? 257 + (end - _out) : 257 - (_out - end));
  state.hold = hold;
  state.bits = bits;
  return;
};

},{}],35:[function(_dereq_,module,exports){
'use strict';


var utils = _dereq_('../utils/common');
var adler32 = _dereq_('./adler32');
var crc32   = _dereq_('./crc32');
var inflate_fast = _dereq_('./inffast');
var inflate_table = _dereq_('./inftrees');

var CODES = 0;
var LENS = 1;
var DISTS = 2;

/* Public constants ==========================================================*/
/* ===========================================================================*/


/* Allowed flush values; see deflate() and inflate() below for details */
//var Z_NO_FLUSH      = 0;
//var Z_PARTIAL_FLUSH = 1;
//var Z_SYNC_FLUSH    = 2;
//var Z_FULL_FLUSH    = 3;
var Z_FINISH        = 4;
var Z_BLOCK         = 5;
var Z_TREES         = 6;


/* Return codes for the compression/decompression functions. Negative values
 * are errors, positive values are used for special but normal events.
 */
var Z_OK            = 0;
var Z_STREAM_END    = 1;
var Z_NEED_DICT     = 2;
//var Z_ERRNO         = -1;
var Z_STREAM_ERROR  = -2;
var Z_DATA_ERROR    = -3;
var Z_MEM_ERROR     = -4;
var Z_BUF_ERROR     = -5;
//var Z_VERSION_ERROR = -6;

/* The deflate compression method */
var Z_DEFLATED  = 8;


/* STATES ====================================================================*/
/* ===========================================================================*/


var    HEAD = 1;       /* i: waiting for magic header */
var    FLAGS = 2;      /* i: waiting for method and flags (gzip) */
var    TIME = 3;       /* i: waiting for modification time (gzip) */
var    OS = 4;         /* i: waiting for extra flags and operating system (gzip) */
var    EXLEN = 5;      /* i: waiting for extra length (gzip) */
var    EXTRA = 6;      /* i: waiting for extra bytes (gzip) */
var    NAME = 7;       /* i: waiting for end of file name (gzip) */
var    COMMENT = 8;    /* i: waiting for end of comment (gzip) */
var    HCRC = 9;       /* i: waiting for header crc (gzip) */
var    DICTID = 10;    /* i: waiting for dictionary check value */
var    DICT = 11;      /* waiting for inflateSetDictionary() call */
var        TYPE = 12;      /* i: waiting for type bits, including last-flag bit */
var        TYPEDO = 13;    /* i: same, but skip check to exit inflate on new block */
var        STORED = 14;    /* i: waiting for stored size (length and complement) */
var        COPY_ = 15;     /* i/o: same as COPY below, but only first time in */
var        COPY = 16;      /* i/o: waiting for input or output to copy stored block */
var        TABLE = 17;     /* i: waiting for dynamic block table lengths */
var        LENLENS = 18;   /* i: waiting for code length code lengths */
var        CODELENS = 19;  /* i: waiting for length/lit and distance code lengths */
var            LEN_ = 20;      /* i: same as LEN below, but only first time in */
var            LEN = 21;       /* i: waiting for length/lit/eob code */
var            LENEXT = 22;    /* i: waiting for length extra bits */
var            DIST = 23;      /* i: waiting for distance code */
var            DISTEXT = 24;   /* i: waiting for distance extra bits */
var            MATCH = 25;     /* o: waiting for output space to copy string */
var            LIT = 26;       /* o: waiting for output space to write literal */
var    CHECK = 27;     /* i: waiting for 32-bit check value */
var    LENGTH = 28;    /* i: waiting for 32-bit length (gzip) */
var    DONE = 29;      /* finished check, done -- remain here until reset */
var    BAD = 30;       /* got a data error -- remain here until reset */
var    MEM = 31;       /* got an inflate() memory error -- remain here until reset */
var    SYNC = 32;      /* looking for synchronization bytes to restart inflate() */

/* ===========================================================================*/



var ENOUGH_LENS = 852;
var ENOUGH_DISTS = 592;
//var ENOUGH =  (ENOUGH_LENS+ENOUGH_DISTS);

var MAX_WBITS = 15;
/* 32K LZ77 window */
var DEF_WBITS = MAX_WBITS;


function ZSWAP32(q) {
  return  (((q >>> 24) & 0xff) +
          ((q >>> 8) & 0xff00) +
          ((q & 0xff00) << 8) +
          ((q & 0xff) << 24));
}


function InflateState() {
  this.mode = 0;             /* current inflate mode */
  this.last = false;          /* true if processing last block */
  this.wrap = 0;              /* bit 0 true for zlib, bit 1 true for gzip */
  this.havedict = false;      /* true if dictionary provided */
  this.flags = 0;             /* gzip header method and flags (0 if zlib) */
  this.dmax = 0;              /* zlib header max distance (INFLATE_STRICT) */
  this.check = 0;             /* protected copy of check value */
  this.total = 0;             /* protected copy of output count */
  // TODO: may be {}
  this.head = null;           /* where to save gzip header information */

  /* sliding window */
  this.wbits = 0;             /* log base 2 of requested window size */
  this.wsize = 0;             /* window size or zero if not using window */
  this.whave = 0;             /* valid bytes in the window */
  this.wnext = 0;             /* window write index */
  this.window = null;         /* allocated sliding window, if needed */

  /* bit accumulator */
  this.hold = 0;              /* input bit accumulator */
  this.bits = 0;              /* number of bits in "in" */

  /* for string and stored block copying */
  this.length = 0;            /* literal or length of data to copy */
  this.offset = 0;            /* distance back to copy string from */

  /* for table and code decoding */
  this.extra = 0;             /* extra bits needed */

  /* fixed and dynamic code tables */
  this.lencode = null;          /* starting table for length/literal codes */
  this.distcode = null;         /* starting table for distance codes */
  this.lenbits = 0;           /* index bits for lencode */
  this.distbits = 0;          /* index bits for distcode */

  /* dynamic table building */
  this.ncode = 0;             /* number of code length code lengths */
  this.nlen = 0;              /* number of length code lengths */
  this.ndist = 0;             /* number of distance code lengths */
  this.have = 0;              /* number of code lengths in lens[] */
  this.next = null;              /* next available space in codes[] */

  this.lens = new utils.Buf16(320); /* temporary storage for code lengths */
  this.work = new utils.Buf16(288); /* work area for code table building */

  /*
   because we don't have pointers in js, we use lencode and distcode directly
   as buffers so we don't need codes
  */
  //this.codes = new utils.Buf32(ENOUGH);       /* space for code tables */
  this.lendyn = null;              /* dynamic table for length/literal codes (JS specific) */
  this.distdyn = null;             /* dynamic table for distance codes (JS specific) */
  this.sane = 0;                   /* if false, allow invalid distance too far */
  this.back = 0;                   /* bits back of last unprocessed length/lit */
  this.was = 0;                    /* initial length of match */
}

function inflateResetKeep(strm) {
  var state;

  if (!strm || !strm.state) { return Z_STREAM_ERROR; }
  state = strm.state;
  strm.total_in = strm.total_out = state.total = 0;
  strm.msg = ''; /*Z_NULL*/
  if (state.wrap) {       /* to support ill-conceived Java test suite */
    strm.adler = state.wrap & 1;
  }
  state.mode = HEAD;
  state.last = 0;
  state.havedict = 0;
  state.dmax = 32768;
  state.head = null/*Z_NULL*/;
  state.hold = 0;
  state.bits = 0;
  //state.lencode = state.distcode = state.next = state.codes;
  state.lencode = state.lendyn = new utils.Buf32(ENOUGH_LENS);
  state.distcode = state.distdyn = new utils.Buf32(ENOUGH_DISTS);

  state.sane = 1;
  state.back = -1;
  //Tracev((stderr, "inflate: reset\n"));
  return Z_OK;
}

function inflateReset(strm) {
  var state;

  if (!strm || !strm.state) { return Z_STREAM_ERROR; }
  state = strm.state;
  state.wsize = 0;
  state.whave = 0;
  state.wnext = 0;
  return inflateResetKeep(strm);

}

function inflateReset2(strm, windowBits) {
  var wrap;
  var state;

  /* get the state */
  if (!strm || !strm.state) { return Z_STREAM_ERROR; }
  state = strm.state;

  /* extract wrap request from windowBits parameter */
  if (windowBits < 0) {
    wrap = 0;
    windowBits = -windowBits;
  }
  else {
    wrap = (windowBits >> 4) + 1;
    if (windowBits < 48) {
      windowBits &= 15;
    }
  }

  /* set number of window bits, free window if different */
  if (windowBits && (windowBits < 8 || windowBits > 15)) {
    return Z_STREAM_ERROR;
  }
  if (state.window !== null && state.wbits !== windowBits) {
    state.window = null;
  }

  /* update state and reset the rest of it */
  state.wrap = wrap;
  state.wbits = windowBits;
  return inflateReset(strm);
}

function inflateInit2(strm, windowBits) {
  var ret;
  var state;

  if (!strm) { return Z_STREAM_ERROR; }
  //strm.msg = Z_NULL;                 /* in case we return an error */

  state = new InflateState();

  //if (state === Z_NULL) return Z_MEM_ERROR;
  //Tracev((stderr, "inflate: allocated\n"));
  strm.state = state;
  state.window = null/*Z_NULL*/;
  ret = inflateReset2(strm, windowBits);
  if (ret !== Z_OK) {
    strm.state = null/*Z_NULL*/;
  }
  return ret;
}

function inflateInit(strm) {
  return inflateInit2(strm, DEF_WBITS);
}


/*
 Return state with length and distance decoding tables and index sizes set to
 fixed code decoding.  Normally this returns fixed tables from inffixed.h.
 If BUILDFIXED is defined, then instead this routine builds the tables the
 first time it's called, and returns those tables the first time and
 thereafter.  This reduces the size of the code by about 2K bytes, in
 exchange for a little execution time.  However, BUILDFIXED should not be
 used for threaded applications, since the rewriting of the tables and virgin
 may not be thread-safe.
 */
var virgin = true;

var lenfix, distfix; // We have no pointers in JS, so keep tables separate

function fixedtables(state) {
  /* build fixed huffman tables if first call (may not be thread safe) */
  if (virgin) {
    var sym;

    lenfix = new utils.Buf32(512);
    distfix = new utils.Buf32(32);

    /* literal/length table */
    sym = 0;
    while (sym < 144) { state.lens[sym++] = 8; }
    while (sym < 256) { state.lens[sym++] = 9; }
    while (sym < 280) { state.lens[sym++] = 7; }
    while (sym < 288) { state.lens[sym++] = 8; }

    inflate_table(LENS,  state.lens, 0, 288, lenfix,   0, state.work, {bits: 9});

    /* distance table */
    sym = 0;
    while (sym < 32) { state.lens[sym++] = 5; }

    inflate_table(DISTS, state.lens, 0, 32,   distfix, 0, state.work, {bits: 5});

    /* do this just once */
    virgin = false;
  }

  state.lencode = lenfix;
  state.lenbits = 9;
  state.distcode = distfix;
  state.distbits = 5;
}


/*
 Update the window with the last wsize (normally 32K) bytes written before
 returning.  If window does not exist yet, create it.  This is only called
 when a window is already in use, or when output has been written during this
 inflate call, but the end of the deflate stream has not been reached yet.
 It is also called to create a window for dictionary data when a dictionary
 is loaded.

 Providing output buffers larger than 32K to inflate() should provide a speed
 advantage, since only the last 32K of output is copied to the sliding window
 upon return from inflate(), and since all distances after the first 32K of
 output will fall in the output data, making match copies simpler and faster.
 The advantage may be dependent on the size of the processor's data caches.
 */
function updatewindow(strm, src, end, copy) {
  var dist;
  var state = strm.state;

  /* if it hasn't been done already, allocate space for the window */
  if (state.window === null) {
    state.wsize = 1 << state.wbits;
    state.wnext = 0;
    state.whave = 0;

    state.window = new utils.Buf8(state.wsize);
  }

  /* copy state->wsize or less output bytes into the circular window */
  if (copy >= state.wsize) {
    utils.arraySet(state.window,src, end - state.wsize, state.wsize, 0);
    state.wnext = 0;
    state.whave = state.wsize;
  }
  else {
    dist = state.wsize - state.wnext;
    if (dist > copy) {
      dist = copy;
    }
    //zmemcpy(state->window + state->wnext, end - copy, dist);
    utils.arraySet(state.window,src, end - copy, dist, state.wnext);
    copy -= dist;
    if (copy) {
      //zmemcpy(state->window, end - copy, copy);
      utils.arraySet(state.window,src, end - copy, copy, 0);
      state.wnext = copy;
      state.whave = state.wsize;
    }
    else {
      state.wnext += dist;
      if (state.wnext === state.wsize) { state.wnext = 0; }
      if (state.whave < state.wsize) { state.whave += dist; }
    }
  }
  return 0;
}

function inflate(strm, flush) {
  var state;
  var input, output;          // input/output buffers
  var next;                   /* next input INDEX */
  var put;                    /* next output INDEX */
  var have, left;             /* available input and output */
  var hold;                   /* bit buffer */
  var bits;                   /* bits in bit buffer */
  var _in, _out;              /* save starting available input and output */
  var copy;                   /* number of stored or match bytes to copy */
  var from;                   /* where to copy match bytes from */
  var from_source;
  var here = 0;               /* current decoding table entry */
  var here_bits, here_op, here_val; // paked "here" denormalized (JS specific)
  //var last;                   /* parent table entry */
  var last_bits, last_op, last_val; // paked "last" denormalized (JS specific)
  var len;                    /* length to copy for repeats, bits to drop */
  var ret;                    /* return code */
  var hbuf = new utils.Buf8(4);    /* buffer for gzip header crc calculation */
  var opts;

  var n; // temporary var for NEED_BITS

  var order = /* permutation of code lengths */
    [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];


  if (!strm || !strm.state || !strm.output ||
      (!strm.input && strm.avail_in !== 0)) {
    return Z_STREAM_ERROR;
  }

  state = strm.state;
  if (state.mode === TYPE) { state.mode = TYPEDO; }    /* skip check */


  //--- LOAD() ---
  put = strm.next_out;
  output = strm.output;
  left = strm.avail_out;
  next = strm.next_in;
  input = strm.input;
  have = strm.avail_in;
  hold = state.hold;
  bits = state.bits;
  //---

  _in = have;
  _out = left;
  ret = Z_OK;

  inf_leave: // goto emulation
  for (;;) {
    switch (state.mode) {
    case HEAD:
      if (state.wrap === 0) {
        state.mode = TYPEDO;
        break;
      }
      //=== NEEDBITS(16);
      while (bits < 16) {
        if (have === 0) { break inf_leave; }
        have--;
        hold += input[next++] << bits;
        bits += 8;
      }
      //===//
      if ((state.wrap & 2) && hold === 0x8b1f) {  /* gzip header */
        state.check = 0/*crc32(0L, Z_NULL, 0)*/;
        //=== CRC2(state.check, hold);
        hbuf[0] = hold & 0xff;
        hbuf[1] = (hold >>> 8) & 0xff;
        state.check = crc32(state.check, hbuf, 2, 0);
        //===//

        //=== INITBITS();
        hold = 0;
        bits = 0;
        //===//
        state.mode = FLAGS;
        break;
      }
      state.flags = 0;           /* expect zlib header */
      if (state.head) {
        state.head.done = false;
      }
      if (!(state.wrap & 1) ||   /* check if zlib header allowed */
        (((hold & 0xff)/*BITS(8)*/ << 8) + (hold >> 8)) % 31) {
        strm.msg = 'incorrect header check';
        state.mode = BAD;
        break;
      }
      if ((hold & 0x0f)/*BITS(4)*/ !== Z_DEFLATED) {
        strm.msg = 'unknown compression method';
        state.mode = BAD;
        break;
      }
      //--- DROPBITS(4) ---//
      hold >>>= 4;
      bits -= 4;
      //---//
      len = (hold & 0x0f)/*BITS(4)*/ + 8;
      if (state.wbits === 0) {
        state.wbits = len;
      }
      else if (len > state.wbits) {
        strm.msg = 'invalid window size';
        state.mode = BAD;
        break;
      }
      state.dmax = 1 << len;
      //Tracev((stderr, "inflate:   zlib header ok\n"));
      strm.adler = state.check = 1/*adler32(0L, Z_NULL, 0)*/;
      state.mode = hold & 0x200 ? DICTID : TYPE;
      //=== INITBITS();
      hold = 0;
      bits = 0;
      //===//
      break;
    case FLAGS:
      //=== NEEDBITS(16); */
      while (bits < 16) {
        if (have === 0) { break inf_leave; }
        have--;
        hold += input[next++] << bits;
        bits += 8;
      }
      //===//
      state.flags = hold;
      if ((state.flags & 0xff) !== Z_DEFLATED) {
        strm.msg = 'unknown compression method';
        state.mode = BAD;
        break;
      }
      if (state.flags & 0xe000) {
        strm.msg = 'unknown header flags set';
        state.mode = BAD;
        break;
      }
      if (state.head) {
        state.head.text = ((hold >> 8) & 1);
      }
      if (state.flags & 0x0200) {
        //=== CRC2(state.check, hold);
        hbuf[0] = hold & 0xff;
        hbuf[1] = (hold >>> 8) & 0xff;
        state.check = crc32(state.check, hbuf, 2, 0);
        //===//
      }
      //=== INITBITS();
      hold = 0;
      bits = 0;
      //===//
      state.mode = TIME;
      /* falls through */
    case TIME:
      //=== NEEDBITS(32); */
      while (bits < 32) {
        if (have === 0) { break inf_leave; }
        have--;
        hold += input[next++] << bits;
        bits += 8;
      }
      //===//
      if (state.head) {
        state.head.time = hold;
      }
      if (state.flags & 0x0200) {
        //=== CRC4(state.check, hold)
        hbuf[0] = hold & 0xff;
        hbuf[1] = (hold >>> 8) & 0xff;
        hbuf[2] = (hold >>> 16) & 0xff;
        hbuf[3] = (hold >>> 24) & 0xff;
        state.check = crc32(state.check, hbuf, 4, 0);
        //===
      }
      //=== INITBITS();
      hold = 0;
      bits = 0;
      //===//
      state.mode = OS;
      /* falls through */
    case OS:
      //=== NEEDBITS(16); */
      while (bits < 16) {
        if (have === 0) { break inf_leave; }
        have--;
        hold += input[next++] << bits;
        bits += 8;
      }
      //===//
      if (state.head) {
        state.head.xflags = (hold & 0xff);
        state.head.os = (hold >> 8);
      }
      if (state.flags & 0x0200) {
        //=== CRC2(state.check, hold);
        hbuf[0] = hold & 0xff;
        hbuf[1] = (hold >>> 8) & 0xff;
        state.check = crc32(state.check, hbuf, 2, 0);
        //===//
      }
      //=== INITBITS();
      hold = 0;
      bits = 0;
      //===//
      state.mode = EXLEN;
      /* falls through */
    case EXLEN:
      if (state.flags & 0x0400) {
        //=== NEEDBITS(16); */
        while (bits < 16) {
          if (have === 0) { break inf_leave; }
          have--;
          hold += input[next++] << bits;
          bits += 8;
        }
        //===//
        state.length = hold;
        if (state.head) {
          state.head.extra_len = hold;
        }
        if (state.flags & 0x0200) {
          //=== CRC2(state.check, hold);
          hbuf[0] = hold & 0xff;
          hbuf[1] = (hold >>> 8) & 0xff;
          state.check = crc32(state.check, hbuf, 2, 0);
          //===//
        }
        //=== INITBITS();
        hold = 0;
        bits = 0;
        //===//
      }
      else if (state.head) {
        state.head.extra = null/*Z_NULL*/;
      }
      state.mode = EXTRA;
      /* falls through */
    case EXTRA:
      if (state.flags & 0x0400) {
        copy = state.length;
        if (copy > have) { copy = have; }
        if (copy) {
          if (state.head) {
            len = state.head.extra_len - state.length;
            if (!state.head.extra) {
              // Use untyped array for more conveniend processing later
              state.head.extra = new Array(state.head.extra_len);
            }
            utils.arraySet(
              state.head.extra,
              input,
              next,
              // extra field is limited to 65536 bytes
              // - no need for additional size check
              copy,
              /*len + copy > state.head.extra_max - len ? state.head.extra_max : copy,*/
              len
            );
            //zmemcpy(state.head.extra + len, next,
            //        len + copy > state.head.extra_max ?
            //        state.head.extra_max - len : copy);
          }
          if (state.flags & 0x0200) {
            state.check = crc32(state.check, input, copy, next);
          }
          have -= copy;
          next += copy;
          state.length -= copy;
        }
        if (state.length) { break inf_leave; }
      }
      state.length = 0;
      state.mode = NAME;
      /* falls through */
    case NAME:
      if (state.flags & 0x0800) {
        if (have === 0) { break inf_leave; }
        copy = 0;
        do {
          // TODO: 2 or 1 bytes?
          len = input[next + copy++];
          /* use constant limit because in js we should not preallocate memory */
          if (state.head && len &&
              (state.length < 65536 /*state.head.name_max*/)) {
            state.head.name += String.fromCharCode(len);
          }
        } while (len && copy < have);

        if (state.flags & 0x0200) {
          state.check = crc32(state.check, input, copy, next);
        }
        have -= copy;
        next += copy;
        if (len) { break inf_leave; }
      }
      else if (state.head) {
        state.head.name = null;
      }
      state.length = 0;
      state.mode = COMMENT;
      /* falls through */
    case COMMENT:
      if (state.flags & 0x1000) {
        if (have === 0) { break inf_leave; }
        copy = 0;
        do {
          len = input[next + copy++];
          /* use constant limit because in js we should not preallocate memory */
          if (state.head && len &&
              (state.length < 65536 /*state.head.comm_max*/)) {
            state.head.comment += String.fromCharCode(len);
          }
        } while (len && copy < have);
        if (state.flags & 0x0200) {
          state.check = crc32(state.check, input, copy, next);
        }
        have -= copy;
        next += copy;
        if (len) { break inf_leave; }
      }
      else if (state.head) {
        state.head.comment = null;
      }
      state.mode = HCRC;
      /* falls through */
    case HCRC:
      if (state.flags & 0x0200) {
        //=== NEEDBITS(16); */
        while (bits < 16) {
          if (have === 0) { break inf_leave; }
          have--;
          hold += input[next++] << bits;
          bits += 8;
        }
        //===//
        if (hold !== (state.check & 0xffff)) {
          strm.msg = 'header crc mismatch';
          state.mode = BAD;
          break;
        }
        //=== INITBITS();
        hold = 0;
        bits = 0;
        //===//
      }
      if (state.head) {
        state.head.hcrc = ((state.flags >> 9) & 1);
        state.head.done = true;
      }
      strm.adler = state.check = 0 /*crc32(0L, Z_NULL, 0)*/;
      state.mode = TYPE;
      break;
    case DICTID:
      //=== NEEDBITS(32); */
      while (bits < 32) {
        if (have === 0) { break inf_leave; }
        have--;
        hold += input[next++] << bits;
        bits += 8;
      }
      //===//
      strm.adler = state.check = ZSWAP32(hold);
      //=== INITBITS();
      hold = 0;
      bits = 0;
      //===//
      state.mode = DICT;
      /* falls through */
    case DICT:
      if (state.havedict === 0) {
        //--- RESTORE() ---
        strm.next_out = put;
        strm.avail_out = left;
        strm.next_in = next;
        strm.avail_in = have;
        state.hold = hold;
        state.bits = bits;
        //---
        return Z_NEED_DICT;
      }
      strm.adler = state.check = 1/*adler32(0L, Z_NULL, 0)*/;
      state.mode = TYPE;
      /* falls through */
    case TYPE:
      if (flush === Z_BLOCK || flush === Z_TREES) { break inf_leave; }
      /* falls through */
    case TYPEDO:
      if (state.last) {
        //--- BYTEBITS() ---//
        hold >>>= bits & 7;
        bits -= bits & 7;
        //---//
        state.mode = CHECK;
        break;
      }
      //=== NEEDBITS(3); */
      while (bits < 3) {
        if (have === 0) { break inf_leave; }
        have--;
        hold += input[next++] << bits;
        bits += 8;
      }
      //===//
      state.last = (hold & 0x01)/*BITS(1)*/;
      //--- DROPBITS(1) ---//
      hold >>>= 1;
      bits -= 1;
      //---//

      switch ((hold & 0x03)/*BITS(2)*/) {
      case 0:                             /* stored block */
        //Tracev((stderr, "inflate:     stored block%s\n",
        //        state.last ? " (last)" : ""));
        state.mode = STORED;
        break;
      case 1:                             /* fixed block */
        fixedtables(state);
        //Tracev((stderr, "inflate:     fixed codes block%s\n",
        //        state.last ? " (last)" : ""));
        state.mode = LEN_;             /* decode codes */
        if (flush === Z_TREES) {
          //--- DROPBITS(2) ---//
          hold >>>= 2;
          bits -= 2;
          //---//
          break inf_leave;
        }
        break;
      case 2:                             /* dynamic block */
        //Tracev((stderr, "inflate:     dynamic codes block%s\n",
        //        state.last ? " (last)" : ""));
        state.mode = TABLE;
        break;
      case 3:
        strm.msg = 'invalid block type';
        state.mode = BAD;
      }
      //--- DROPBITS(2) ---//
      hold >>>= 2;
      bits -= 2;
      //---//
      break;
    case STORED:
      //--- BYTEBITS() ---// /* go to byte boundary */
      hold >>>= bits & 7;
      bits -= bits & 7;
      //---//
      //=== NEEDBITS(32); */
      while (bits < 32) {
        if (have === 0) { break inf_leave; }
        have--;
        hold += input[next++] << bits;
        bits += 8;
      }
      //===//
      if ((hold & 0xffff) !== ((hold >>> 16) ^ 0xffff)) {
        strm.msg = 'invalid stored block lengths';
        state.mode = BAD;
        break;
      }
      state.length = hold & 0xffff;
      //Tracev((stderr, "inflate:       stored length %u\n",
      //        state.length));
      //=== INITBITS();
      hold = 0;
      bits = 0;
      //===//
      state.mode = COPY_;
      if (flush === Z_TREES) { break inf_leave; }
      /* falls through */
    case COPY_:
      state.mode = COPY;
      /* falls through */
    case COPY:
      copy = state.length;
      if (copy) {
        if (copy > have) { copy = have; }
        if (copy > left) { copy = left; }
        if (copy === 0) { break inf_leave; }
        //--- zmemcpy(put, next, copy); ---
        utils.arraySet(output, input, next, copy, put);
        //---//
        have -= copy;
        next += copy;
        left -= copy;
        put += copy;
        state.length -= copy;
        break;
      }
      //Tracev((stderr, "inflate:       stored end\n"));
      state.mode = TYPE;
      break;
    case TABLE:
      //=== NEEDBITS(14); */
      while (bits < 14) {
        if (have === 0) { break inf_leave; }
        have--;
        hold += input[next++] << bits;
        bits += 8;
      }
      //===//
      state.nlen = (hold & 0x1f)/*BITS(5)*/ + 257;
      //--- DROPBITS(5) ---//
      hold >>>= 5;
      bits -= 5;
      //---//
      state.ndist = (hold & 0x1f)/*BITS(5)*/ + 1;
      //--- DROPBITS(5) ---//
      hold >>>= 5;
      bits -= 5;
      //---//
      state.ncode = (hold & 0x0f)/*BITS(4)*/ + 4;
      //--- DROPBITS(4) ---//
      hold >>>= 4;
      bits -= 4;
      //---//
//#ifndef PKZIP_BUG_WORKAROUND
      if (state.nlen > 286 || state.ndist > 30) {
        strm.msg = 'too many length or distance symbols';
        state.mode = BAD;
        break;
      }
//#endif
      //Tracev((stderr, "inflate:       table sizes ok\n"));
      state.have = 0;
      state.mode = LENLENS;
      /* falls through */
    case LENLENS:
      while (state.have < state.ncode) {
        //=== NEEDBITS(3);
        while (bits < 3) {
          if (have === 0) { break inf_leave; }
          have--;
          hold += input[next++] << bits;
          bits += 8;
        }
        //===//
        state.lens[order[state.have++]] = (hold & 0x07);//BITS(3);
        //--- DROPBITS(3) ---//
        hold >>>= 3;
        bits -= 3;
        //---//
      }
      while (state.have < 19) {
        state.lens[order[state.have++]] = 0;
      }
      // We have separate tables & no pointers. 2 commented lines below not needed.
      //state.next = state.codes;
      //state.lencode = state.next;
      // Switch to use dynamic table
      state.lencode = state.lendyn;
      state.lenbits = 7;

      opts = {bits: state.lenbits};
      ret = inflate_table(CODES, state.lens, 0, 19, state.lencode, 0, state.work, opts);
      state.lenbits = opts.bits;

      if (ret) {
        strm.msg = 'invalid code lengths set';
        state.mode = BAD;
        break;
      }
      //Tracev((stderr, "inflate:       code lengths ok\n"));
      state.have = 0;
      state.mode = CODELENS;
      /* falls through */
    case CODELENS:
      while (state.have < state.nlen + state.ndist) {
        for (;;) {
          here = state.lencode[hold & ((1 << state.lenbits) - 1)];/*BITS(state.lenbits)*/
          here_bits = here >>> 24;
          here_op = (here >>> 16) & 0xff;
          here_val = here & 0xffff;

          if ((here_bits) <= bits) { break; }
          //--- PULLBYTE() ---//
          if (have === 0) { break inf_leave; }
          have--;
          hold += input[next++] << bits;
          bits += 8;
          //---//
        }
        if (here_val < 16) {
          //--- DROPBITS(here.bits) ---//
          hold >>>= here_bits;
          bits -= here_bits;
          //---//
          state.lens[state.have++] = here_val;
        }
        else {
          if (here_val === 16) {
            //=== NEEDBITS(here.bits + 2);
            n = here_bits + 2;
            while (bits < n) {
              if (have === 0) { break inf_leave; }
              have--;
              hold += input[next++] << bits;
              bits += 8;
            }
            //===//
            //--- DROPBITS(here.bits) ---//
            hold >>>= here_bits;
            bits -= here_bits;
            //---//
            if (state.have === 0) {
              strm.msg = 'invalid bit length repeat';
              state.mode = BAD;
              break;
            }
            len = state.lens[state.have - 1];
            copy = 3 + (hold & 0x03);//BITS(2);
            //--- DROPBITS(2) ---//
            hold >>>= 2;
            bits -= 2;
            //---//
          }
          else if (here_val === 17) {
            //=== NEEDBITS(here.bits + 3);
            n = here_bits + 3;
            while (bits < n) {
              if (have === 0) { break inf_leave; }
              have--;
              hold += input[next++] << bits;
              bits += 8;
            }
            //===//
            //--- DROPBITS(here.bits) ---//
            hold >>>= here_bits;
            bits -= here_bits;
            //---//
            len = 0;
            copy = 3 + (hold & 0x07);//BITS(3);
            //--- DROPBITS(3) ---//
            hold >>>= 3;
            bits -= 3;
            //---//
          }
          else {
            //=== NEEDBITS(here.bits + 7);
            n = here_bits + 7;
            while (bits < n) {
              if (have === 0) { break inf_leave; }
              have--;
              hold += input[next++] << bits;
              bits += 8;
            }
            //===//
            //--- DROPBITS(here.bits) ---//
            hold >>>= here_bits;
            bits -= here_bits;
            //---//
            len = 0;
            copy = 11 + (hold & 0x7f);//BITS(7);
            //--- DROPBITS(7) ---//
            hold >>>= 7;
            bits -= 7;
            //---//
          }
          if (state.have + copy > state.nlen + state.ndist) {
            strm.msg = 'invalid bit length repeat';
            state.mode = BAD;
            break;
          }
          while (copy--) {
            state.lens[state.have++] = len;
          }
        }
      }

      /* handle error breaks in while */
      if (state.mode === BAD) { break; }

      /* check for end-of-block code (better have one) */
      if (state.lens[256] === 0) {
        strm.msg = 'invalid code -- missing end-of-block';
        state.mode = BAD;
        break;
      }

      /* build code tables -- note: do not change the lenbits or distbits
         values here (9 and 6) without reading the comments in inftrees.h
         concerning the ENOUGH constants, which depend on those values */
      state.lenbits = 9;

      opts = {bits: state.lenbits};
      ret = inflate_table(LENS, state.lens, 0, state.nlen, state.lencode, 0, state.work, opts);
      // We have separate tables & no pointers. 2 commented lines below not needed.
      // state.next_index = opts.table_index;
      state.lenbits = opts.bits;
      // state.lencode = state.next;

      if (ret) {
        strm.msg = 'invalid literal/lengths set';
        state.mode = BAD;
        break;
      }

      state.distbits = 6;
      //state.distcode.copy(state.codes);
      // Switch to use dynamic table
      state.distcode = state.distdyn;
      opts = {bits: state.distbits};
      ret = inflate_table(DISTS, state.lens, state.nlen, state.ndist, state.distcode, 0, state.work, opts);
      // We have separate tables & no pointers. 2 commented lines below not needed.
      // state.next_index = opts.table_index;
      state.distbits = opts.bits;
      // state.distcode = state.next;

      if (ret) {
        strm.msg = 'invalid distances set';
        state.mode = BAD;
        break;
      }
      //Tracev((stderr, 'inflate:       codes ok\n'));
      state.mode = LEN_;
      if (flush === Z_TREES) { break inf_leave; }
      /* falls through */
    case LEN_:
      state.mode = LEN;
      /* falls through */
    case LEN:
      if (have >= 6 && left >= 258) {
        //--- RESTORE() ---
        strm.next_out = put;
        strm.avail_out = left;
        strm.next_in = next;
        strm.avail_in = have;
        state.hold = hold;
        state.bits = bits;
        //---
        inflate_fast(strm, _out);
        //--- LOAD() ---
        put = strm.next_out;
        output = strm.output;
        left = strm.avail_out;
        next = strm.next_in;
        input = strm.input;
        have = strm.avail_in;
        hold = state.hold;
        bits = state.bits;
        //---

        if (state.mode === TYPE) {
          state.back = -1;
        }
        break;
      }
      state.back = 0;
      for (;;) {
        here = state.lencode[hold & ((1 << state.lenbits) -1)];  /*BITS(state.lenbits)*/
        here_bits = here >>> 24;
        here_op = (here >>> 16) & 0xff;
        here_val = here & 0xffff;

        if (here_bits <= bits) { break; }
        //--- PULLBYTE() ---//
        if (have === 0) { break inf_leave; }
        have--;
        hold += input[next++] << bits;
        bits += 8;
        //---//
      }
      if (here_op && (here_op & 0xf0) === 0) {
        last_bits = here_bits;
        last_op = here_op;
        last_val = here_val;
        for (;;) {
          here = state.lencode[last_val +
                  ((hold & ((1 << (last_bits + last_op)) -1))/*BITS(last.bits + last.op)*/ >> last_bits)];
          here_bits = here >>> 24;
          here_op = (here >>> 16) & 0xff;
          here_val = here & 0xffff;

          if ((last_bits + here_bits) <= bits) { break; }
          //--- PULLBYTE() ---//
          if (have === 0) { break inf_leave; }
          have--;
          hold += input[next++] << bits;
          bits += 8;
          //---//
        }
        //--- DROPBITS(last.bits) ---//
        hold >>>= last_bits;
        bits -= last_bits;
        //---//
        state.back += last_bits;
      }
      //--- DROPBITS(here.bits) ---//
      hold >>>= here_bits;
      bits -= here_bits;
      //---//
      state.back += here_bits;
      state.length = here_val;
      if (here_op === 0) {
        //Tracevv((stderr, here.val >= 0x20 && here.val < 0x7f ?
        //        "inflate:         literal '%c'\n" :
        //        "inflate:         literal 0x%02x\n", here.val));
        state.mode = LIT;
        break;
      }
      if (here_op & 32) {
        //Tracevv((stderr, "inflate:         end of block\n"));
        state.back = -1;
        state.mode = TYPE;
        break;
      }
      if (here_op & 64) {
        strm.msg = 'invalid literal/length code';
        state.mode = BAD;
        break;
      }
      state.extra = here_op & 15;
      state.mode = LENEXT;
      /* falls through */
    case LENEXT:
      if (state.extra) {
        //=== NEEDBITS(state.extra);
        n = state.extra;
        while (bits < n) {
          if (have === 0) { break inf_leave; }
          have--;
          hold += input[next++] << bits;
          bits += 8;
        }
        //===//
        state.length += hold & ((1 << state.extra) -1)/*BITS(state.extra)*/;
        //--- DROPBITS(state.extra) ---//
        hold >>>= state.extra;
        bits -= state.extra;
        //---//
        state.back += state.extra;
      }
      //Tracevv((stderr, "inflate:         length %u\n", state.length));
      state.was = state.length;
      state.mode = DIST;
      /* falls through */
    case DIST:
      for (;;) {
        here = state.distcode[hold & ((1 << state.distbits) -1)];/*BITS(state.distbits)*/
        here_bits = here >>> 24;
        here_op = (here >>> 16) & 0xff;
        here_val = here & 0xffff;

        if ((here_bits) <= bits) { break; }
        //--- PULLBYTE() ---//
        if (have === 0) { break inf_leave; }
        have--;
        hold += input[next++] << bits;
        bits += 8;
        //---//
      }
      if ((here_op & 0xf0) === 0) {
        last_bits = here_bits;
        last_op = here_op;
        last_val = here_val;
        for (;;) {
          here = state.distcode[last_val +
                  ((hold & ((1 << (last_bits + last_op)) -1))/*BITS(last.bits + last.op)*/ >> last_bits)];
          here_bits = here >>> 24;
          here_op = (here >>> 16) & 0xff;
          here_val = here & 0xffff;

          if ((last_bits + here_bits) <= bits) { break; }
          //--- PULLBYTE() ---//
          if (have === 0) { break inf_leave; }
          have--;
          hold += input[next++] << bits;
          bits += 8;
          //---//
        }
        //--- DROPBITS(last.bits) ---//
        hold >>>= last_bits;
        bits -= last_bits;
        //---//
        state.back += last_bits;
      }
      //--- DROPBITS(here.bits) ---//
      hold >>>= here_bits;
      bits -= here_bits;
      //---//
      state.back += here_bits;
      if (here_op & 64) {
        strm.msg = 'invalid distance code';
        state.mode = BAD;
        break;
      }
      state.offset = here_val;
      state.extra = (here_op) & 15;
      state.mode = DISTEXT;
      /* falls through */
    case DISTEXT:
      if (state.extra) {
        //=== NEEDBITS(state.extra);
        n = state.extra;
        while (bits < n) {
          if (have === 0) { break inf_leave; }
          have--;
          hold += input[next++] << bits;
          bits += 8;
        }
        //===//
        state.offset += hold & ((1 << state.extra) -1)/*BITS(state.extra)*/;
        //--- DROPBITS(state.extra) ---//
        hold >>>= state.extra;
        bits -= state.extra;
        //---//
        state.back += state.extra;
      }
//#ifdef INFLATE_STRICT
      if (state.offset > state.dmax) {
        strm.msg = 'invalid distance too far back';
        state.mode = BAD;
        break;
      }
//#endif
      //Tracevv((stderr, "inflate:         distance %u\n", state.offset));
      state.mode = MATCH;
      /* falls through */
    case MATCH:
      if (left === 0) { break inf_leave; }
      copy = _out - left;
      if (state.offset > copy) {         /* copy from window */
        copy = state.offset - copy;
        if (copy > state.whave) {
          if (state.sane) {
            strm.msg = 'invalid distance too far back';
            state.mode = BAD;
            break;
          }
// (!) This block is disabled in zlib defailts,
// don't enable it for binary compatibility
//#ifdef INFLATE_ALLOW_INVALID_DISTANCE_TOOFAR_ARRR
//          Trace((stderr, "inflate.c too far\n"));
//          copy -= state.whave;
//          if (copy > state.length) { copy = state.length; }
//          if (copy > left) { copy = left; }
//          left -= copy;
//          state.length -= copy;
//          do {
//            output[put++] = 0;
//          } while (--copy);
//          if (state.length === 0) { state.mode = LEN; }
//          break;
//#endif
        }
        if (copy > state.wnext) {
          copy -= state.wnext;
          from = state.wsize - copy;
        }
        else {
          from = state.wnext - copy;
        }
        if (copy > state.length) { copy = state.length; }
        from_source = state.window;
      }
      else {                              /* copy from output */
        from_source = output;
        from = put - state.offset;
        copy = state.length;
      }
      if (copy > left) { copy = left; }
      left -= copy;
      state.length -= copy;
      do {
        output[put++] = from_source[from++];
      } while (--copy);
      if (state.length === 0) { state.mode = LEN; }
      break;
    case LIT:
      if (left === 0) { break inf_leave; }
      output[put++] = state.length;
      left--;
      state.mode = LEN;
      break;
    case CHECK:
      if (state.wrap) {
        //=== NEEDBITS(32);
        while (bits < 32) {
          if (have === 0) { break inf_leave; }
          have--;
          // Use '|' insdead of '+' to make sure that result is signed
          hold |= input[next++] << bits;
          bits += 8;
        }
        //===//
        _out -= left;
        strm.total_out += _out;
        state.total += _out;
        if (_out) {
          strm.adler = state.check =
              /*UPDATE(state.check, put - _out, _out);*/
              (state.flags ? crc32(state.check, output, _out, put - _out) : adler32(state.check, output, _out, put - _out));

        }
        _out = left;
        // NB: crc32 stored as signed 32-bit int, ZSWAP32 returns signed too
        if ((state.flags ? hold : ZSWAP32(hold)) !== state.check) {
          strm.msg = 'incorrect data check';
          state.mode = BAD;
          break;
        }
        //=== INITBITS();
        hold = 0;
        bits = 0;
        //===//
        //Tracev((stderr, "inflate:   check matches trailer\n"));
      }
      state.mode = LENGTH;
      /* falls through */
    case LENGTH:
      if (state.wrap && state.flags) {
        //=== NEEDBITS(32);
        while (bits < 32) {
          if (have === 0) { break inf_leave; }
          have--;
          hold += input[next++] << bits;
          bits += 8;
        }
        //===//
        if (hold !== (state.total & 0xffffffff)) {
          strm.msg = 'incorrect length check';
          state.mode = BAD;
          break;
        }
        //=== INITBITS();
        hold = 0;
        bits = 0;
        //===//
        //Tracev((stderr, "inflate:   length matches trailer\n"));
      }
      state.mode = DONE;
      /* falls through */
    case DONE:
      ret = Z_STREAM_END;
      break inf_leave;
    case BAD:
      ret = Z_DATA_ERROR;
      break inf_leave;
    case MEM:
      return Z_MEM_ERROR;
    case SYNC:
      /* falls through */
    default:
      return Z_STREAM_ERROR;
    }
  }

  // inf_leave <- here is real place for "goto inf_leave", emulated via "break inf_leave"

  /*
     Return from inflate(), updating the total counts and the check value.
     If there was no progress during the inflate() call, return a buffer
     error.  Call updatewindow() to create and/or update the window state.
     Note: a memory error from inflate() is non-recoverable.
   */

  //--- RESTORE() ---
  strm.next_out = put;
  strm.avail_out = left;
  strm.next_in = next;
  strm.avail_in = have;
  state.hold = hold;
  state.bits = bits;
  //---

  if (state.wsize || (_out !== strm.avail_out && state.mode < BAD &&
                      (state.mode < CHECK || flush !== Z_FINISH))) {
    if (updatewindow(strm, strm.output, strm.next_out, _out - strm.avail_out)) {
      state.mode = MEM;
      return Z_MEM_ERROR;
    }
  }
  _in -= strm.avail_in;
  _out -= strm.avail_out;
  strm.total_in += _in;
  strm.total_out += _out;
  state.total += _out;
  if (state.wrap && _out) {
    strm.adler = state.check = /*UPDATE(state.check, strm.next_out - _out, _out);*/
      (state.flags ? crc32(state.check, output, _out, strm.next_out - _out) : adler32(state.check, output, _out, strm.next_out - _out));
  }
  strm.data_type = state.bits + (state.last ? 64 : 0) +
                    (state.mode === TYPE ? 128 : 0) +
                    (state.mode === LEN_ || state.mode === COPY_ ? 256 : 0);
  if (((_in === 0 && _out === 0) || flush === Z_FINISH) && ret === Z_OK) {
    ret = Z_BUF_ERROR;
  }
  return ret;
}

function inflateEnd(strm) {

  if (!strm || !strm.state /*|| strm->zfree == (free_func)0*/) {
    return Z_STREAM_ERROR;
  }

  var state = strm.state;
  if (state.window) {
    state.window = null;
  }
  strm.state = null;
  return Z_OK;
}

function inflateGetHeader(strm, head) {
  var state;

  /* check state */
  if (!strm || !strm.state) { return Z_STREAM_ERROR; }
  state = strm.state;
  if ((state.wrap & 2) === 0) { return Z_STREAM_ERROR; }

  /* save header structure */
  state.head = head;
  head.done = false;
  return Z_OK;
}


exports.inflateReset = inflateReset;
exports.inflateReset2 = inflateReset2;
exports.inflateResetKeep = inflateResetKeep;
exports.inflateInit = inflateInit;
exports.inflateInit2 = inflateInit2;
exports.inflate = inflate;
exports.inflateEnd = inflateEnd;
exports.inflateGetHeader = inflateGetHeader;
exports.inflateInfo = 'pako inflate (from Nodeca project)';

/* Not implemented
exports.inflateCopy = inflateCopy;
exports.inflateGetDictionary = inflateGetDictionary;
exports.inflateMark = inflateMark;
exports.inflatePrime = inflatePrime;
exports.inflateSetDictionary = inflateSetDictionary;
exports.inflateSync = inflateSync;
exports.inflateSyncPoint = inflateSyncPoint;
exports.inflateUndermine = inflateUndermine;
*/
},{"../utils/common":27,"./adler32":29,"./crc32":31,"./inffast":34,"./inftrees":36}],36:[function(_dereq_,module,exports){
'use strict';


var utils = _dereq_('../utils/common');

var MAXBITS = 15;
var ENOUGH_LENS = 852;
var ENOUGH_DISTS = 592;
//var ENOUGH = (ENOUGH_LENS+ENOUGH_DISTS);

var CODES = 0;
var LENS = 1;
var DISTS = 2;

var lbase = [ /* Length codes 257..285 base */
  3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31,
  35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258, 0, 0
];

var lext = [ /* Length codes 257..285 extra */
  16, 16, 16, 16, 16, 16, 16, 16, 17, 17, 17, 17, 18, 18, 18, 18,
  19, 19, 19, 19, 20, 20, 20, 20, 21, 21, 21, 21, 16, 72, 78
];

var dbase = [ /* Distance codes 0..29 base */
  1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193,
  257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145,
  8193, 12289, 16385, 24577, 0, 0
];

var dext = [ /* Distance codes 0..29 extra */
  16, 16, 16, 16, 17, 17, 18, 18, 19, 19, 20, 20, 21, 21, 22, 22,
  23, 23, 24, 24, 25, 25, 26, 26, 27, 27,
  28, 28, 29, 29, 64, 64
];

module.exports = function inflate_table(type, lens, lens_index, codes, table, table_index, work, opts)
{
  var bits = opts.bits;
      //here = opts.here; /* table entry for duplication */

  var len = 0;               /* a code's length in bits */
  var sym = 0;               /* index of code symbols */
  var min = 0, max = 0;          /* minimum and maximum code lengths */
  var root = 0;              /* number of index bits for root table */
  var curr = 0;              /* number of index bits for current table */
  var drop = 0;              /* code bits to drop for sub-table */
  var left = 0;                   /* number of prefix codes available */
  var used = 0;              /* code entries in table used */
  var huff = 0;              /* Huffman code */
  var incr;              /* for incrementing code, index */
  var fill;              /* index for replicating entries */
  var low;               /* low bits for current root entry */
  var mask;              /* mask for low root bits */
  var next;             /* next available space in table */
  var base = null;     /* base value table to use */
  var base_index = 0;
//  var shoextra;    /* extra bits table to use */
  var end;                    /* use base and extra for symbol > end */
  var count = new utils.Buf16(MAXBITS+1); //[MAXBITS+1];    /* number of codes of each length */
  var offs = new utils.Buf16(MAXBITS+1); //[MAXBITS+1];     /* offsets in table for each length */
  var extra = null;
  var extra_index = 0;

  var here_bits, here_op, here_val;

  /*
   Process a set of code lengths to create a canonical Huffman code.  The
   code lengths are lens[0..codes-1].  Each length corresponds to the
   symbols 0..codes-1.  The Huffman code is generated by first sorting the
   symbols by length from short to long, and retaining the symbol order
   for codes with equal lengths.  Then the code starts with all zero bits
   for the first code of the shortest length, and the codes are integer
   increments for the same length, and zeros are appended as the length
   increases.  For the deflate format, these bits are stored backwards
   from their more natural integer increment ordering, and so when the
   decoding tables are built in the large loop below, the integer codes
   are incremented backwards.

   This routine assumes, but does not check, that all of the entries in
   lens[] are in the range 0..MAXBITS.  The caller must assure this.
   1..MAXBITS is interpreted as that code length.  zero means that that
   symbol does not occur in this code.

   The codes are sorted by computing a count of codes for each length,
   creating from that a table of starting indices for each length in the
   sorted table, and then entering the symbols in order in the sorted
   table.  The sorted table is work[], with that space being provided by
   the caller.

   The length counts are used for other purposes as well, i.e. finding
   the minimum and maximum length codes, determining if there are any
   codes at all, checking for a valid set of lengths, and looking ahead
   at length counts to determine sub-table sizes when building the
   decoding tables.
   */

  /* accumulate lengths for codes (assumes lens[] all in 0..MAXBITS) */
  for (len = 0; len <= MAXBITS; len++) {
    count[len] = 0;
  }
  for (sym = 0; sym < codes; sym++) {
    count[lens[lens_index + sym]]++;
  }

  /* bound code lengths, force root to be within code lengths */
  root = bits;
  for (max = MAXBITS; max >= 1; max--) {
    if (count[max] !== 0) { break; }
  }
  if (root > max) {
    root = max;
  }
  if (max === 0) {                     /* no symbols to code at all */
    //table.op[opts.table_index] = 64;  //here.op = (var char)64;    /* invalid code marker */
    //table.bits[opts.table_index] = 1;   //here.bits = (var char)1;
    //table.val[opts.table_index++] = 0;   //here.val = (var short)0;
    table[table_index++] = (1 << 24) | (64 << 16) | 0;


    //table.op[opts.table_index] = 64;
    //table.bits[opts.table_index] = 1;
    //table.val[opts.table_index++] = 0;
    table[table_index++] = (1 << 24) | (64 << 16) | 0;

    opts.bits = 1;
    return 0;     /* no symbols, but wait for decoding to report error */
  }
  for (min = 1; min < max; min++) {
    if (count[min] !== 0) { break; }
  }
  if (root < min) {
    root = min;
  }

  /* check for an over-subscribed or incomplete set of lengths */
  left = 1;
  for (len = 1; len <= MAXBITS; len++) {
    left <<= 1;
    left -= count[len];
    if (left < 0) {
      return -1;
    }        /* over-subscribed */
  }
  if (left > 0 && (type === CODES || max !== 1)) {
    return -1;                      /* incomplete set */
  }

  /* generate offsets into symbol table for each length for sorting */
  offs[1] = 0;
  for (len = 1; len < MAXBITS; len++) {
    offs[len + 1] = offs[len] + count[len];
  }

  /* sort symbols by length, by symbol order within each length */
  for (sym = 0; sym < codes; sym++) {
    if (lens[lens_index + sym] !== 0) {
      work[offs[lens[lens_index + sym]]++] = sym;
    }
  }

  /*
   Create and fill in decoding tables.  In this loop, the table being
   filled is at next and has curr index bits.  The code being used is huff
   with length len.  That code is converted to an index by dropping drop
   bits off of the bottom.  For codes where len is less than drop + curr,
   those top drop + curr - len bits are incremented through all values to
   fill the table with replicated entries.

   root is the number of index bits for the root table.  When len exceeds
   root, sub-tables are created pointed to by the root entry with an index
   of the low root bits of huff.  This is saved in low to check for when a
   new sub-table should be started.  drop is zero when the root table is
   being filled, and drop is root when sub-tables are being filled.

   When a new sub-table is needed, it is necessary to look ahead in the
   code lengths to determine what size sub-table is needed.  The length
   counts are used for this, and so count[] is decremented as codes are
   entered in the tables.

   used keeps track of how many table entries have been allocated from the
   provided *table space.  It is checked for LENS and DIST tables against
   the constants ENOUGH_LENS and ENOUGH_DISTS to guard against changes in
   the initial root table size constants.  See the comments in inftrees.h
   for more information.

   sym increments through all symbols, and the loop terminates when
   all codes of length max, i.e. all codes, have been processed.  This
   routine permits incomplete codes, so another loop after this one fills
   in the rest of the decoding tables with invalid code markers.
   */

  /* set up for code type */
  // poor man optimization - use if-else instead of switch,
  // to avoid deopts in old v8
  if (type === CODES) {
      base = extra = work;    /* dummy value--not used */
      end = 19;
  } else if (type === LENS) {
      base = lbase;
      base_index -= 257;
      extra = lext;
      extra_index -= 257;
      end = 256;
  } else {                    /* DISTS */
      base = dbase;
      extra = dext;
      end = -1;
  }

  /* initialize opts for loop */
  huff = 0;                   /* starting code */
  sym = 0;                    /* starting code symbol */
  len = min;                  /* starting code length */
  next = table_index;              /* current table to fill in */
  curr = root;                /* current table index bits */
  drop = 0;                   /* current bits to drop from code for index */
  low = -1;                   /* trigger new sub-table when len > root */
  used = 1 << root;          /* use root table entries */
  mask = used - 1;            /* mask for comparing low */

  /* check available table space */
  if ((type === LENS && used > ENOUGH_LENS) ||
    (type === DISTS && used > ENOUGH_DISTS)) {
    return 1;
  }

  var i=0;
  /* process all codes and make table entries */
  for (;;) {
    i++;
    /* create table entry */
    here_bits = len - drop;
    if (work[sym] < end) {
      here_op = 0;
      here_val = work[sym];
    }
    else if (work[sym] > end) {
      here_op = extra[extra_index + work[sym]];
      here_val = base[base_index + work[sym]];
    }
    else {
      here_op = 32 + 64;         /* end of block */
      here_val = 0;
    }

    /* replicate for those indices with low len bits equal to huff */
    incr = 1 << (len - drop);
    fill = 1 << curr;
    min = fill;                 /* save offset to next table */
    do {
      fill -= incr;
      table[next + (huff >> drop) + fill] = (here_bits << 24) | (here_op << 16) | here_val |0;
    } while (fill !== 0);

    /* backwards increment the len-bit code huff */
    incr = 1 << (len - 1);
    while (huff & incr) {
      incr >>= 1;
    }
    if (incr !== 0) {
      huff &= incr - 1;
      huff += incr;
    } else {
      huff = 0;
    }

    /* go to next symbol, update count, len */
    sym++;
    if (--count[len] === 0) {
      if (len === max) { break; }
      len = lens[lens_index + work[sym]];
    }

    /* create new sub-table if needed */
    if (len > root && (huff & mask) !== low) {
      /* if first time, transition to sub-tables */
      if (drop === 0) {
        drop = root;
      }

      /* increment past last table */
      next += min;            /* here min is 1 << curr */

      /* determine length of next table */
      curr = len - drop;
      left = 1 << curr;
      while (curr + drop < max) {
        left -= count[curr + drop];
        if (left <= 0) { break; }
        curr++;
        left <<= 1;
      }

      /* check for enough space */
      used += 1 << curr;
      if ((type === LENS && used > ENOUGH_LENS) ||
        (type === DISTS && used > ENOUGH_DISTS)) {
        return 1;
      }

      /* point entry in root table to sub-table */
      low = huff & mask;
      /*table.op[low] = curr;
      table.bits[low] = root;
      table.val[low] = next - opts.table_index;*/
      table[low] = (root << 24) | (curr << 16) | (next - table_index) |0;
    }
  }

  /* fill in remaining table entry if code is incomplete (guaranteed to have
   at most one remaining entry, since if the code is incomplete, the
   maximum code length that was allowed to get this far is one bit) */
  if (huff !== 0) {
    //table.op[next + huff] = 64;            /* invalid code marker */
    //table.bits[next + huff] = len - drop;
    //table.val[next + huff] = 0;
    table[next + huff] = ((len - drop) << 24) | (64 << 16) |0;
  }

  /* set return parameters */
  //opts.table_index += used;
  opts.bits = root;
  return 0;
};

},{"../utils/common":27}],37:[function(_dereq_,module,exports){
'use strict';

module.exports = {
  '2':    'need dictionary',     /* Z_NEED_DICT       2  */
  '1':    'stream end',          /* Z_STREAM_END      1  */
  '0':    '',                    /* Z_OK              0  */
  '-1':   'file error',          /* Z_ERRNO         (-1) */
  '-2':   'stream error',        /* Z_STREAM_ERROR  (-2) */
  '-3':   'data error',          /* Z_DATA_ERROR    (-3) */
  '-4':   'insufficient memory', /* Z_MEM_ERROR     (-4) */
  '-5':   'buffer error',        /* Z_BUF_ERROR     (-5) */
  '-6':   'incompatible version' /* Z_VERSION_ERROR (-6) */
};
},{}],38:[function(_dereq_,module,exports){
'use strict';


var utils = _dereq_('../utils/common');

/* Public constants ==========================================================*/
/* ===========================================================================*/


//var Z_FILTERED          = 1;
//var Z_HUFFMAN_ONLY      = 2;
//var Z_RLE               = 3;
var Z_FIXED               = 4;
//var Z_DEFAULT_STRATEGY  = 0;

/* Possible values of the data_type field (though see inflate()) */
var Z_BINARY              = 0;
var Z_TEXT                = 1;
//var Z_ASCII             = 1; // = Z_TEXT
var Z_UNKNOWN             = 2;

/*============================================================================*/


function zero(buf) { var len = buf.length; while (--len >= 0) { buf[len] = 0; } }

// From zutil.h

var STORED_BLOCK = 0;
var STATIC_TREES = 1;
var DYN_TREES    = 2;
/* The three kinds of block type */

var MIN_MATCH    = 3;
var MAX_MATCH    = 258;
/* The minimum and maximum match lengths */

// From deflate.h
/* ===========================================================================
 * Internal compression state.
 */

var LENGTH_CODES  = 29;
/* number of length codes, not counting the special END_BLOCK code */

var LITERALS      = 256;
/* number of literal bytes 0..255 */

var L_CODES       = LITERALS + 1 + LENGTH_CODES;
/* number of Literal or Length codes, including the END_BLOCK code */

var D_CODES       = 30;
/* number of distance codes */

var BL_CODES      = 19;
/* number of codes used to transfer the bit lengths */

var HEAP_SIZE     = 2*L_CODES + 1;
/* maximum heap size */

var MAX_BITS      = 15;
/* All codes must not exceed MAX_BITS bits */

var Buf_size      = 16;
/* size of bit buffer in bi_buf */


/* ===========================================================================
 * Constants
 */

var MAX_BL_BITS = 7;
/* Bit length codes must not exceed MAX_BL_BITS bits */

var END_BLOCK   = 256;
/* end of block literal code */

var REP_3_6     = 16;
/* repeat previous bit length 3-6 times (2 bits of repeat count) */

var REPZ_3_10   = 17;
/* repeat a zero length 3-10 times  (3 bits of repeat count) */

var REPZ_11_138 = 18;
/* repeat a zero length 11-138 times  (7 bits of repeat count) */

var extra_lbits =   /* extra bits for each length code */
  [0,0,0,0,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3,3,4,4,4,4,5,5,5,5,0];

var extra_dbits =   /* extra bits for each distance code */
  [0,0,0,0,1,1,2,2,3,3,4,4,5,5,6,6,7,7,8,8,9,9,10,10,11,11,12,12,13,13];

var extra_blbits =  /* extra bits for each bit length code */
  [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,2,3,7];

var bl_order =
  [16,17,18,0,8,7,9,6,10,5,11,4,12,3,13,2,14,1,15];
/* The lengths of the bit length codes are sent in order of decreasing
 * probability, to avoid transmitting the lengths for unused bit length codes.
 */

/* ===========================================================================
 * Local data. These are initialized only once.
 */

// We pre-fill arrays with 0 to avoid uninitialized gaps

var DIST_CODE_LEN = 512; /* see definition of array dist_code below */

// !!!! Use flat array insdead of structure, Freq = i*2, Len = i*2+1
var static_ltree  = new Array((L_CODES+2) * 2);
zero(static_ltree);
/* The static literal tree. Since the bit lengths are imposed, there is no
 * need for the L_CODES extra codes used during heap construction. However
 * The codes 286 and 287 are needed to build a canonical tree (see _tr_init
 * below).
 */

var static_dtree  = new Array(D_CODES * 2);
zero(static_dtree);
/* The static distance tree. (Actually a trivial tree since all codes use
 * 5 bits.)
 */

var _dist_code    = new Array(DIST_CODE_LEN);
zero(_dist_code);
/* Distance codes. The first 256 values correspond to the distances
 * 3 .. 258, the last 256 values correspond to the top 8 bits of
 * the 15 bit distances.
 */

var _length_code  = new Array(MAX_MATCH-MIN_MATCH+1);
zero(_length_code);
/* length code for each normalized match length (0 == MIN_MATCH) */

var base_length   = new Array(LENGTH_CODES);
zero(base_length);
/* First normalized length for each code (0 = MIN_MATCH) */

var base_dist     = new Array(D_CODES);
zero(base_dist);
/* First normalized distance for each code (0 = distance of 1) */


var StaticTreeDesc = function (static_tree, extra_bits, extra_base, elems, max_length) {

  this.static_tree  = static_tree;  /* static tree or NULL */
  this.extra_bits   = extra_bits;   /* extra bits for each code or NULL */
  this.extra_base   = extra_base;   /* base index for extra_bits */
  this.elems        = elems;        /* max number of elements in the tree */
  this.max_length   = max_length;   /* max bit length for the codes */

  // show if `static_tree` has data or dummy - needed for monomorphic objects
  this.has_stree    = static_tree && static_tree.length;
};


var static_l_desc;
var static_d_desc;
var static_bl_desc;


var TreeDesc = function(dyn_tree, stat_desc) {
  this.dyn_tree = dyn_tree;     /* the dynamic tree */
  this.max_code = 0;            /* largest code with non zero frequency */
  this.stat_desc = stat_desc;   /* the corresponding static tree */
};



function d_code(dist) {
  return dist < 256 ? _dist_code[dist] : _dist_code[256 + (dist >>> 7)];
}


/* ===========================================================================
 * Output a short LSB first on the stream.
 * IN assertion: there is enough room in pendingBuf.
 */
function put_short (s, w) {
//    put_byte(s, (uch)((w) & 0xff));
//    put_byte(s, (uch)((ush)(w) >> 8));
  s.pending_buf[s.pending++] = (w) & 0xff;
  s.pending_buf[s.pending++] = (w >>> 8) & 0xff;
}


/* ===========================================================================
 * Send a value on a given number of bits.
 * IN assertion: length <= 16 and value fits in length bits.
 */
function send_bits(s, value, length) {
  if (s.bi_valid > (Buf_size - length)) {
    s.bi_buf |= (value << s.bi_valid) & 0xffff;
    put_short(s, s.bi_buf);
    s.bi_buf = value >> (Buf_size - s.bi_valid);
    s.bi_valid += length - Buf_size;
  } else {
    s.bi_buf |= (value << s.bi_valid) & 0xffff;
    s.bi_valid += length;
  }
}


function send_code(s, c, tree) {
  send_bits(s, tree[c*2]/*.Code*/, tree[c*2 + 1]/*.Len*/);
}


/* ===========================================================================
 * Reverse the first len bits of a code, using straightforward code (a faster
 * method would use a table)
 * IN assertion: 1 <= len <= 15
 */
function bi_reverse(code, len) {
  var res = 0;
  do {
    res |= code & 1;
    code >>>= 1;
    res <<= 1;
  } while (--len > 0);
  return res >>> 1;
}


/* ===========================================================================
 * Flush the bit buffer, keeping at most 7 bits in it.
 */
function bi_flush(s) {
  if (s.bi_valid === 16) {
    put_short(s, s.bi_buf);
    s.bi_buf = 0;
    s.bi_valid = 0;

  } else if (s.bi_valid >= 8) {
    s.pending_buf[s.pending++] = s.bi_buf & 0xff;
    s.bi_buf >>= 8;
    s.bi_valid -= 8;
  }
}


/* ===========================================================================
 * Compute the optimal bit lengths for a tree and update the total bit length
 * for the current block.
 * IN assertion: the fields freq and dad are set, heap[heap_max] and
 *    above are the tree nodes sorted by increasing frequency.
 * OUT assertions: the field len is set to the optimal bit length, the
 *     array bl_count contains the frequencies for each bit length.
 *     The length opt_len is updated; static_len is also updated if stree is
 *     not null.
 */
function gen_bitlen(s, desc)
//    deflate_state *s;
//    tree_desc *desc;    /* the tree descriptor */
{
  var tree            = desc.dyn_tree;
  var max_code        = desc.max_code;
  var stree           = desc.stat_desc.static_tree;
  var has_stree       = desc.stat_desc.has_stree;
  var extra           = desc.stat_desc.extra_bits;
  var base            = desc.stat_desc.extra_base;
  var max_length      = desc.stat_desc.max_length;
  var h;              /* heap index */
  var n, m;           /* iterate over the tree elements */
  var bits;           /* bit length */
  var xbits;          /* extra bits */
  var f;              /* frequency */
  var overflow = 0;   /* number of elements with bit length too large */

  for (bits = 0; bits <= MAX_BITS; bits++) {
    s.bl_count[bits] = 0;
  }

  /* In a first pass, compute the optimal bit lengths (which may
   * overflow in the case of the bit length tree).
   */
  tree[s.heap[s.heap_max]*2 + 1]/*.Len*/ = 0; /* root of the heap */

  for (h = s.heap_max+1; h < HEAP_SIZE; h++) {
    n = s.heap[h];
    bits = tree[tree[n*2 +1]/*.Dad*/ * 2 + 1]/*.Len*/ + 1;
    if (bits > max_length) {
      bits = max_length;
      overflow++;
    }
    tree[n*2 + 1]/*.Len*/ = bits;
    /* We overwrite tree[n].Dad which is no longer needed */

    if (n > max_code) { continue; } /* not a leaf node */

    s.bl_count[bits]++;
    xbits = 0;
    if (n >= base) {
      xbits = extra[n-base];
    }
    f = tree[n * 2]/*.Freq*/;
    s.opt_len += f * (bits + xbits);
    if (has_stree) {
      s.static_len += f * (stree[n*2 + 1]/*.Len*/ + xbits);
    }
  }
  if (overflow === 0) { return; }

  // Trace((stderr,"\nbit length overflow\n"));
  /* This happens for example on obj2 and pic of the Calgary corpus */

  /* Find the first bit length which could increase: */
  do {
    bits = max_length-1;
    while (s.bl_count[bits] === 0) { bits--; }
    s.bl_count[bits]--;      /* move one leaf down the tree */
    s.bl_count[bits+1] += 2; /* move one overflow item as its brother */
    s.bl_count[max_length]--;
    /* The brother of the overflow item also moves one step up,
     * but this does not affect bl_count[max_length]
     */
    overflow -= 2;
  } while (overflow > 0);

  /* Now recompute all bit lengths, scanning in increasing frequency.
   * h is still equal to HEAP_SIZE. (It is simpler to reconstruct all
   * lengths instead of fixing only the wrong ones. This idea is taken
   * from 'ar' written by Haruhiko Okumura.)
   */
  for (bits = max_length; bits !== 0; bits--) {
    n = s.bl_count[bits];
    while (n !== 0) {
      m = s.heap[--h];
      if (m > max_code) { continue; }
      if (tree[m*2 + 1]/*.Len*/ !== bits) {
        // Trace((stderr,"code %d bits %d->%d\n", m, tree[m].Len, bits));
        s.opt_len += (bits - tree[m*2 + 1]/*.Len*/)*tree[m*2]/*.Freq*/;
        tree[m*2 + 1]/*.Len*/ = bits;
      }
      n--;
    }
  }
}


/* ===========================================================================
 * Generate the codes for a given tree and bit counts (which need not be
 * optimal).
 * IN assertion: the array bl_count contains the bit length statistics for
 * the given tree and the field len is set for all tree elements.
 * OUT assertion: the field code is set for all tree elements of non
 *     zero code length.
 */
function gen_codes(tree, max_code, bl_count)
//    ct_data *tree;             /* the tree to decorate */
//    int max_code;              /* largest code with non zero frequency */
//    ushf *bl_count;            /* number of codes at each bit length */
{
  var next_code = new Array(MAX_BITS+1); /* next code value for each bit length */
  var code = 0;              /* running code value */
  var bits;                  /* bit index */
  var n;                     /* code index */

  /* The distribution counts are first used to generate the code values
   * without bit reversal.
   */
  for (bits = 1; bits <= MAX_BITS; bits++) {
    next_code[bits] = code = (code + bl_count[bits-1]) << 1;
  }
  /* Check that the bit counts in bl_count are consistent. The last code
   * must be all ones.
   */
  //Assert (code + bl_count[MAX_BITS]-1 == (1<<MAX_BITS)-1,
  //        "inconsistent bit counts");
  //Tracev((stderr,"\ngen_codes: max_code %d ", max_code));

  for (n = 0;  n <= max_code; n++) {
    var len = tree[n*2 + 1]/*.Len*/;
    if (len === 0) { continue; }
    /* Now reverse the bits */
    tree[n*2]/*.Code*/ = bi_reverse(next_code[len]++, len);

    //Tracecv(tree != static_ltree, (stderr,"\nn %3d %c l %2d c %4x (%x) ",
    //     n, (isgraph(n) ? n : ' '), len, tree[n].Code, next_code[len]-1));
  }
}


/* ===========================================================================
 * Initialize the various 'constant' tables.
 */
function tr_static_init() {
  var n;        /* iterates over tree elements */
  var bits;     /* bit counter */
  var length;   /* length value */
  var code;     /* code value */
  var dist;     /* distance index */
  var bl_count = new Array(MAX_BITS+1);
  /* number of codes at each bit length for an optimal tree */

  // do check in _tr_init()
  //if (static_init_done) return;

  /* For some embedded targets, global variables are not initialized: */
/*#ifdef NO_INIT_GLOBAL_POINTERS
  static_l_desc.static_tree = static_ltree;
  static_l_desc.extra_bits = extra_lbits;
  static_d_desc.static_tree = static_dtree;
  static_d_desc.extra_bits = extra_dbits;
  static_bl_desc.extra_bits = extra_blbits;
#endif*/

  /* Initialize the mapping length (0..255) -> length code (0..28) */
  length = 0;
  for (code = 0; code < LENGTH_CODES-1; code++) {
    base_length[code] = length;
    for (n = 0; n < (1<<extra_lbits[code]); n++) {
      _length_code[length++] = code;
    }
  }
  //Assert (length == 256, "tr_static_init: length != 256");
  /* Note that the length 255 (match length 258) can be represented
   * in two different ways: code 284 + 5 bits or code 285, so we
   * overwrite length_code[255] to use the best encoding:
   */
  _length_code[length-1] = code;

  /* Initialize the mapping dist (0..32K) -> dist code (0..29) */
  dist = 0;
  for (code = 0 ; code < 16; code++) {
    base_dist[code] = dist;
    for (n = 0; n < (1<<extra_dbits[code]); n++) {
      _dist_code[dist++] = code;
    }
  }
  //Assert (dist == 256, "tr_static_init: dist != 256");
  dist >>= 7; /* from now on, all distances are divided by 128 */
  for ( ; code < D_CODES; code++) {
    base_dist[code] = dist << 7;
    for (n = 0; n < (1<<(extra_dbits[code]-7)); n++) {
      _dist_code[256 + dist++] = code;
    }
  }
  //Assert (dist == 256, "tr_static_init: 256+dist != 512");

  /* Construct the codes of the static literal tree */
  for (bits = 0; bits <= MAX_BITS; bits++) {
    bl_count[bits] = 0;
  }

  n = 0;
  while (n <= 143) {
    static_ltree[n*2 + 1]/*.Len*/ = 8;
    n++;
    bl_count[8]++;
  }
  while (n <= 255) {
    static_ltree[n*2 + 1]/*.Len*/ = 9;
    n++;
    bl_count[9]++;
  }
  while (n <= 279) {
    static_ltree[n*2 + 1]/*.Len*/ = 7;
    n++;
    bl_count[7]++;
  }
  while (n <= 287) {
    static_ltree[n*2 + 1]/*.Len*/ = 8;
    n++;
    bl_count[8]++;
  }
  /* Codes 286 and 287 do not exist, but we must include them in the
   * tree construction to get a canonical Huffman tree (longest code
   * all ones)
   */
  gen_codes(static_ltree, L_CODES+1, bl_count);

  /* The static distance tree is trivial: */
  for (n = 0; n < D_CODES; n++) {
    static_dtree[n*2 + 1]/*.Len*/ = 5;
    static_dtree[n*2]/*.Code*/ = bi_reverse(n, 5);
  }

  // Now data ready and we can init static trees
  static_l_desc = new StaticTreeDesc(static_ltree, extra_lbits, LITERALS+1, L_CODES, MAX_BITS);
  static_d_desc = new StaticTreeDesc(static_dtree, extra_dbits, 0,          D_CODES, MAX_BITS);
  static_bl_desc =new StaticTreeDesc(new Array(0), extra_blbits, 0,         BL_CODES, MAX_BL_BITS);

  //static_init_done = true;
}


/* ===========================================================================
 * Initialize a new block.
 */
function init_block(s) {
  var n; /* iterates over tree elements */

  /* Initialize the trees. */
  for (n = 0; n < L_CODES;  n++) { s.dyn_ltree[n*2]/*.Freq*/ = 0; }
  for (n = 0; n < D_CODES;  n++) { s.dyn_dtree[n*2]/*.Freq*/ = 0; }
  for (n = 0; n < BL_CODES; n++) { s.bl_tree[n*2]/*.Freq*/ = 0; }

  s.dyn_ltree[END_BLOCK*2]/*.Freq*/ = 1;
  s.opt_len = s.static_len = 0;
  s.last_lit = s.matches = 0;
}


/* ===========================================================================
 * Flush the bit buffer and align the output on a byte boundary
 */
function bi_windup(s)
{
  if (s.bi_valid > 8) {
    put_short(s, s.bi_buf);
  } else if (s.bi_valid > 0) {
    //put_byte(s, (Byte)s->bi_buf);
    s.pending_buf[s.pending++] = s.bi_buf;
  }
  s.bi_buf = 0;
  s.bi_valid = 0;
}

/* ===========================================================================
 * Copy a stored block, storing first the length and its
 * one's complement if requested.
 */
function copy_block(s, buf, len, header)
//DeflateState *s;
//charf    *buf;    /* the input data */
//unsigned len;     /* its length */
//int      header;  /* true if block header must be written */
{
  bi_windup(s);        /* align on byte boundary */

  if (header) {
    put_short(s, len);
    put_short(s, ~len);
  }
//  while (len--) {
//    put_byte(s, *buf++);
//  }
  utils.arraySet(s.pending_buf, s.window, buf, len, s.pending);
  s.pending += len;
}

/* ===========================================================================
 * Compares to subtrees, using the tree depth as tie breaker when
 * the subtrees have equal frequency. This minimizes the worst case length.
 */
function smaller(tree, n, m, depth) {
  var _n2 = n*2;
  var _m2 = m*2;
  return (tree[_n2]/*.Freq*/ < tree[_m2]/*.Freq*/ ||
         (tree[_n2]/*.Freq*/ === tree[_m2]/*.Freq*/ && depth[n] <= depth[m]));
}

/* ===========================================================================
 * Restore the heap property by moving down the tree starting at node k,
 * exchanging a node with the smallest of its two sons if necessary, stopping
 * when the heap property is re-established (each father smaller than its
 * two sons).
 */
function pqdownheap(s, tree, k)
//    deflate_state *s;
//    ct_data *tree;  /* the tree to restore */
//    int k;               /* node to move down */
{
  var v = s.heap[k];
  var j = k << 1;  /* left son of k */
  while (j <= s.heap_len) {
    /* Set j to the smallest of the two sons: */
    if (j < s.heap_len &&
      smaller(tree, s.heap[j+1], s.heap[j], s.depth)) {
      j++;
    }
    /* Exit if v is smaller than both sons */
    if (smaller(tree, v, s.heap[j], s.depth)) { break; }

    /* Exchange v with the smallest son */
    s.heap[k] = s.heap[j];
    k = j;

    /* And continue down the tree, setting j to the left son of k */
    j <<= 1;
  }
  s.heap[k] = v;
}


// inlined manually
// var SMALLEST = 1;

/* ===========================================================================
 * Send the block data compressed using the given Huffman trees
 */
function compress_block(s, ltree, dtree)
//    deflate_state *s;
//    const ct_data *ltree; /* literal tree */
//    const ct_data *dtree; /* distance tree */
{
  var dist;           /* distance of matched string */
  var lc;             /* match length or unmatched char (if dist == 0) */
  var lx = 0;         /* running index in l_buf */
  var code;           /* the code to send */
  var extra;          /* number of extra bits to send */

  if (s.last_lit !== 0) {
    do {
      dist = (s.pending_buf[s.d_buf + lx*2] << 8) | (s.pending_buf[s.d_buf + lx*2 + 1]);
      lc = s.pending_buf[s.l_buf + lx];
      lx++;

      if (dist === 0) {
        send_code(s, lc, ltree); /* send a literal byte */
        //Tracecv(isgraph(lc), (stderr," '%c' ", lc));
      } else {
        /* Here, lc is the match length - MIN_MATCH */
        code = _length_code[lc];
        send_code(s, code+LITERALS+1, ltree); /* send the length code */
        extra = extra_lbits[code];
        if (extra !== 0) {
          lc -= base_length[code];
          send_bits(s, lc, extra);       /* send the extra length bits */
        }
        dist--; /* dist is now the match distance - 1 */
        code = d_code(dist);
        //Assert (code < D_CODES, "bad d_code");

        send_code(s, code, dtree);       /* send the distance code */
        extra = extra_dbits[code];
        if (extra !== 0) {
          dist -= base_dist[code];
          send_bits(s, dist, extra);   /* send the extra distance bits */
        }
      } /* literal or match pair ? */

      /* Check that the overlay between pending_buf and d_buf+l_buf is ok: */
      //Assert((uInt)(s->pending) < s->lit_bufsize + 2*lx,
      //       "pendingBuf overflow");

    } while (lx < s.last_lit);
  }

  send_code(s, END_BLOCK, ltree);
}


/* ===========================================================================
 * Construct one Huffman tree and assigns the code bit strings and lengths.
 * Update the total bit length for the current block.
 * IN assertion: the field freq is set for all tree elements.
 * OUT assertions: the fields len and code are set to the optimal bit length
 *     and corresponding code. The length opt_len is updated; static_len is
 *     also updated if stree is not null. The field max_code is set.
 */
function build_tree(s, desc)
//    deflate_state *s;
//    tree_desc *desc; /* the tree descriptor */
{
  var tree     = desc.dyn_tree;
  var stree    = desc.stat_desc.static_tree;
  var has_stree = desc.stat_desc.has_stree;
  var elems    = desc.stat_desc.elems;
  var n, m;          /* iterate over heap elements */
  var max_code = -1; /* largest code with non zero frequency */
  var node;          /* new node being created */

  /* Construct the initial heap, with least frequent element in
   * heap[SMALLEST]. The sons of heap[n] are heap[2*n] and heap[2*n+1].
   * heap[0] is not used.
   */
  s.heap_len = 0;
  s.heap_max = HEAP_SIZE;

  for (n = 0; n < elems; n++) {
    if (tree[n * 2]/*.Freq*/ !== 0) {
      s.heap[++s.heap_len] = max_code = n;
      s.depth[n] = 0;

    } else {
      tree[n*2 + 1]/*.Len*/ = 0;
    }
  }

  /* The pkzip format requires that at least one distance code exists,
   * and that at least one bit should be sent even if there is only one
   * possible code. So to avoid special checks later on we force at least
   * two codes of non zero frequency.
   */
  while (s.heap_len < 2) {
    node = s.heap[++s.heap_len] = (max_code < 2 ? ++max_code : 0);
    tree[node * 2]/*.Freq*/ = 1;
    s.depth[node] = 0;
    s.opt_len--;

    if (has_stree) {
      s.static_len -= stree[node*2 + 1]/*.Len*/;
    }
    /* node is 0 or 1 so it does not have extra bits */
  }
  desc.max_code = max_code;

  /* The elements heap[heap_len/2+1 .. heap_len] are leaves of the tree,
   * establish sub-heaps of increasing lengths:
   */
  for (n = (s.heap_len >> 1/*int /2*/); n >= 1; n--) { pqdownheap(s, tree, n); }

  /* Construct the Huffman tree by repeatedly combining the least two
   * frequent nodes.
   */
  node = elems;              /* next internal node of the tree */
  do {
    //pqremove(s, tree, n);  /* n = node of least frequency */
    /*** pqremove ***/
    n = s.heap[1/*SMALLEST*/];
    s.heap[1/*SMALLEST*/] = s.heap[s.heap_len--];
    pqdownheap(s, tree, 1/*SMALLEST*/);
    /***/

    m = s.heap[1/*SMALLEST*/]; /* m = node of next least frequency */

    s.heap[--s.heap_max] = n; /* keep the nodes sorted by frequency */
    s.heap[--s.heap_max] = m;

    /* Create a new node father of n and m */
    tree[node * 2]/*.Freq*/ = tree[n * 2]/*.Freq*/ + tree[m * 2]/*.Freq*/;
    s.depth[node] = (s.depth[n] >= s.depth[m] ? s.depth[n] : s.depth[m]) + 1;
    tree[n*2 + 1]/*.Dad*/ = tree[m*2 + 1]/*.Dad*/ = node;

    /* and insert the new node in the heap */
    s.heap[1/*SMALLEST*/] = node++;
    pqdownheap(s, tree, 1/*SMALLEST*/);

  } while (s.heap_len >= 2);

  s.heap[--s.heap_max] = s.heap[1/*SMALLEST*/];

  /* At this point, the fields freq and dad are set. We can now
   * generate the bit lengths.
   */
  gen_bitlen(s, desc);

  /* The field len is now set, we can generate the bit codes */
  gen_codes(tree, max_code, s.bl_count);
}


/* ===========================================================================
 * Scan a literal or distance tree to determine the frequencies of the codes
 * in the bit length tree.
 */
function scan_tree(s, tree, max_code)
//    deflate_state *s;
//    ct_data *tree;   /* the tree to be scanned */
//    int max_code;    /* and its largest code of non zero frequency */
{
  var n;                     /* iterates over all tree elements */
  var prevlen = -1;          /* last emitted length */
  var curlen;                /* length of current code */

  var nextlen = tree[0*2 + 1]/*.Len*/; /* length of next code */

  var count = 0;             /* repeat count of the current code */
  var max_count = 7;         /* max repeat count */
  var min_count = 4;         /* min repeat count */

  if (nextlen === 0) {
    max_count = 138;
    min_count = 3;
  }
  tree[(max_code+1)*2 + 1]/*.Len*/ = 0xffff; /* guard */

  for (n = 0; n <= max_code; n++) {
    curlen = nextlen;
    nextlen = tree[(n+1)*2 + 1]/*.Len*/;

    if (++count < max_count && curlen === nextlen) {
      continue;

    } else if (count < min_count) {
      s.bl_tree[curlen * 2]/*.Freq*/ += count;

    } else if (curlen !== 0) {

      if (curlen !== prevlen) { s.bl_tree[curlen * 2]/*.Freq*/++; }
      s.bl_tree[REP_3_6*2]/*.Freq*/++;

    } else if (count <= 10) {
      s.bl_tree[REPZ_3_10*2]/*.Freq*/++;

    } else {
      s.bl_tree[REPZ_11_138*2]/*.Freq*/++;
    }

    count = 0;
    prevlen = curlen;

    if (nextlen === 0) {
      max_count = 138;
      min_count = 3;

    } else if (curlen === nextlen) {
      max_count = 6;
      min_count = 3;

    } else {
      max_count = 7;
      min_count = 4;
    }
  }
}


/* ===========================================================================
 * Send a literal or distance tree in compressed form, using the codes in
 * bl_tree.
 */
function send_tree(s, tree, max_code)
//    deflate_state *s;
//    ct_data *tree; /* the tree to be scanned */
//    int max_code;       /* and its largest code of non zero frequency */
{
  var n;                     /* iterates over all tree elements */
  var prevlen = -1;          /* last emitted length */
  var curlen;                /* length of current code */

  var nextlen = tree[0*2 + 1]/*.Len*/; /* length of next code */

  var count = 0;             /* repeat count of the current code */
  var max_count = 7;         /* max repeat count */
  var min_count = 4;         /* min repeat count */

  /* tree[max_code+1].Len = -1; */  /* guard already set */
  if (nextlen === 0) {
    max_count = 138;
    min_count = 3;
  }

  for (n = 0; n <= max_code; n++) {
    curlen = nextlen;
    nextlen = tree[(n+1)*2 + 1]/*.Len*/;

    if (++count < max_count && curlen === nextlen) {
      continue;

    } else if (count < min_count) {
      do { send_code(s, curlen, s.bl_tree); } while (--count !== 0);

    } else if (curlen !== 0) {
      if (curlen !== prevlen) {
        send_code(s, curlen, s.bl_tree);
        count--;
      }
      //Assert(count >= 3 && count <= 6, " 3_6?");
      send_code(s, REP_3_6, s.bl_tree);
      send_bits(s, count-3, 2);

    } else if (count <= 10) {
      send_code(s, REPZ_3_10, s.bl_tree);
      send_bits(s, count-3, 3);

    } else {
      send_code(s, REPZ_11_138, s.bl_tree);
      send_bits(s, count-11, 7);
    }

    count = 0;
    prevlen = curlen;
    if (nextlen === 0) {
      max_count = 138;
      min_count = 3;

    } else if (curlen === nextlen) {
      max_count = 6;
      min_count = 3;

    } else {
      max_count = 7;
      min_count = 4;
    }
  }
}


/* ===========================================================================
 * Construct the Huffman tree for the bit lengths and return the index in
 * bl_order of the last bit length code to send.
 */
function build_bl_tree(s) {
  var max_blindex;  /* index of last bit length code of non zero freq */

  /* Determine the bit length frequencies for literal and distance trees */
  scan_tree(s, s.dyn_ltree, s.l_desc.max_code);
  scan_tree(s, s.dyn_dtree, s.d_desc.max_code);

  /* Build the bit length tree: */
  build_tree(s, s.bl_desc);
  /* opt_len now includes the length of the tree representations, except
   * the lengths of the bit lengths codes and the 5+5+4 bits for the counts.
   */

  /* Determine the number of bit length codes to send. The pkzip format
   * requires that at least 4 bit length codes be sent. (appnote.txt says
   * 3 but the actual value used is 4.)
   */
  for (max_blindex = BL_CODES-1; max_blindex >= 3; max_blindex--) {
    if (s.bl_tree[bl_order[max_blindex]*2 + 1]/*.Len*/ !== 0) {
      break;
    }
  }
  /* Update opt_len to include the bit length tree and counts */
  s.opt_len += 3*(max_blindex+1) + 5+5+4;
  //Tracev((stderr, "\ndyn trees: dyn %ld, stat %ld",
  //        s->opt_len, s->static_len));

  return max_blindex;
}


/* ===========================================================================
 * Send the header for a block using dynamic Huffman trees: the counts, the
 * lengths of the bit length codes, the literal tree and the distance tree.
 * IN assertion: lcodes >= 257, dcodes >= 1, blcodes >= 4.
 */
function send_all_trees(s, lcodes, dcodes, blcodes)
//    deflate_state *s;
//    int lcodes, dcodes, blcodes; /* number of codes for each tree */
{
  var rank;                    /* index in bl_order */

  //Assert (lcodes >= 257 && dcodes >= 1 && blcodes >= 4, "not enough codes");
  //Assert (lcodes <= L_CODES && dcodes <= D_CODES && blcodes <= BL_CODES,
  //        "too many codes");
  //Tracev((stderr, "\nbl counts: "));
  send_bits(s, lcodes-257, 5); /* not +255 as stated in appnote.txt */
  send_bits(s, dcodes-1,   5);
  send_bits(s, blcodes-4,  4); /* not -3 as stated in appnote.txt */
  for (rank = 0; rank < blcodes; rank++) {
    //Tracev((stderr, "\nbl code %2d ", bl_order[rank]));
    send_bits(s, s.bl_tree[bl_order[rank]*2 + 1]/*.Len*/, 3);
  }
  //Tracev((stderr, "\nbl tree: sent %ld", s->bits_sent));

  send_tree(s, s.dyn_ltree, lcodes-1); /* literal tree */
  //Tracev((stderr, "\nlit tree: sent %ld", s->bits_sent));

  send_tree(s, s.dyn_dtree, dcodes-1); /* distance tree */
  //Tracev((stderr, "\ndist tree: sent %ld", s->bits_sent));
}


/* ===========================================================================
 * Check if the data type is TEXT or BINARY, using the following algorithm:
 * - TEXT if the two conditions below are satisfied:
 *    a) There are no non-portable control characters belonging to the
 *       "black list" (0..6, 14..25, 28..31).
 *    b) There is at least one printable character belonging to the
 *       "white list" (9 {TAB}, 10 {LF}, 13 {CR}, 32..255).
 * - BINARY otherwise.
 * - The following partially-portable control characters form a
 *   "gray list" that is ignored in this detection algorithm:
 *   (7 {BEL}, 8 {BS}, 11 {VT}, 12 {FF}, 26 {SUB}, 27 {ESC}).
 * IN assertion: the fields Freq of dyn_ltree are set.
 */
function detect_data_type(s) {
  /* black_mask is the bit mask of black-listed bytes
   * set bits 0..6, 14..25, and 28..31
   * 0xf3ffc07f = binary 11110011111111111100000001111111
   */
  var black_mask = 0xf3ffc07f;
  var n;

  /* Check for non-textual ("black-listed") bytes. */
  for (n = 0; n <= 31; n++, black_mask >>>= 1) {
    if ((black_mask & 1) && (s.dyn_ltree[n*2]/*.Freq*/ !== 0)) {
      return Z_BINARY;
    }
  }

  /* Check for textual ("white-listed") bytes. */
  if (s.dyn_ltree[9 * 2]/*.Freq*/ !== 0 || s.dyn_ltree[10 * 2]/*.Freq*/ !== 0 ||
      s.dyn_ltree[13 * 2]/*.Freq*/ !== 0) {
    return Z_TEXT;
  }
  for (n = 32; n < LITERALS; n++) {
    if (s.dyn_ltree[n * 2]/*.Freq*/ !== 0) {
      return Z_TEXT;
    }
  }

  /* There are no "black-listed" or "white-listed" bytes:
   * this stream either is empty or has tolerated ("gray-listed") bytes only.
   */
  return Z_BINARY;
}


var static_init_done = false;

/* ===========================================================================
 * Initialize the tree data structures for a new zlib stream.
 */
function _tr_init(s)
{

  if (!static_init_done) {
    tr_static_init();
    static_init_done = true;
  }

  s.l_desc  = new TreeDesc(s.dyn_ltree, static_l_desc);
  s.d_desc  = new TreeDesc(s.dyn_dtree, static_d_desc);
  s.bl_desc = new TreeDesc(s.bl_tree, static_bl_desc);

  s.bi_buf = 0;
  s.bi_valid = 0;

  /* Initialize the first block of the first file: */
  init_block(s);
}


/* ===========================================================================
 * Send a stored block
 */
function _tr_stored_block(s, buf, stored_len, last)
//DeflateState *s;
//charf *buf;       /* input block */
//ulg stored_len;   /* length of input block */
//int last;         /* one if this is the last block for a file */
{
  send_bits(s, (STORED_BLOCK<<1)+(last ? 1 : 0), 3);    /* send block type */
  copy_block(s, buf, stored_len, true); /* with header */
}


/* ===========================================================================
 * Send one empty static block to give enough lookahead for inflate.
 * This takes 10 bits, of which 7 may remain in the bit buffer.
 */
function _tr_align(s) {
  send_bits(s, STATIC_TREES<<1, 3);
  send_code(s, END_BLOCK, static_ltree);
  bi_flush(s);
}


/* ===========================================================================
 * Determine the best encoding for the current block: dynamic trees, static
 * trees or store, and output the encoded block to the zip file.
 */
function _tr_flush_block(s, buf, stored_len, last)
//DeflateState *s;
//charf *buf;       /* input block, or NULL if too old */
//ulg stored_len;   /* length of input block */
//int last;         /* one if this is the last block for a file */
{
  var opt_lenb, static_lenb;  /* opt_len and static_len in bytes */
  var max_blindex = 0;        /* index of last bit length code of non zero freq */

  /* Build the Huffman trees unless a stored block is forced */
  if (s.level > 0) {

    /* Check if the file is binary or text */
    if (s.strm.data_type === Z_UNKNOWN) {
      s.strm.data_type = detect_data_type(s);
    }

    /* Construct the literal and distance trees */
    build_tree(s, s.l_desc);
    // Tracev((stderr, "\nlit data: dyn %ld, stat %ld", s->opt_len,
    //        s->static_len));

    build_tree(s, s.d_desc);
    // Tracev((stderr, "\ndist data: dyn %ld, stat %ld", s->opt_len,
    //        s->static_len));
    /* At this point, opt_len and static_len are the total bit lengths of
     * the compressed block data, excluding the tree representations.
     */

    /* Build the bit length tree for the above two trees, and get the index
     * in bl_order of the last bit length code to send.
     */
    max_blindex = build_bl_tree(s);

    /* Determine the best encoding. Compute the block lengths in bytes. */
    opt_lenb = (s.opt_len+3+7) >>> 3;
    static_lenb = (s.static_len+3+7) >>> 3;

    // Tracev((stderr, "\nopt %lu(%lu) stat %lu(%lu) stored %lu lit %u ",
    //        opt_lenb, s->opt_len, static_lenb, s->static_len, stored_len,
    //        s->last_lit));

    if (static_lenb <= opt_lenb) { opt_lenb = static_lenb; }

  } else {
    // Assert(buf != (char*)0, "lost buf");
    opt_lenb = static_lenb = stored_len + 5; /* force a stored block */
  }

  if ((stored_len+4 <= opt_lenb) && (buf !== -1)) {
    /* 4: two words for the lengths */

    /* The test buf != NULL is only necessary if LIT_BUFSIZE > WSIZE.
     * Otherwise we can't have processed more than WSIZE input bytes since
     * the last block flush, because compression would have been
     * successful. If LIT_BUFSIZE <= WSIZE, it is never too late to
     * transform a block into a stored block.
     */
    _tr_stored_block(s, buf, stored_len, last);

  } else if (s.strategy === Z_FIXED || static_lenb === opt_lenb) {

    send_bits(s, (STATIC_TREES<<1) + (last ? 1 : 0), 3);
    compress_block(s, static_ltree, static_dtree);

  } else {
    send_bits(s, (DYN_TREES<<1) + (last ? 1 : 0), 3);
    send_all_trees(s, s.l_desc.max_code+1, s.d_desc.max_code+1, max_blindex+1);
    compress_block(s, s.dyn_ltree, s.dyn_dtree);
  }
  // Assert (s->compressed_len == s->bits_sent, "bad compressed size");
  /* The above check is made mod 2^32, for files larger than 512 MB
   * and uLong implemented on 32 bits.
   */
  init_block(s);

  if (last) {
    bi_windup(s);
  }
  // Tracev((stderr,"\ncomprlen %lu(%lu) ", s->compressed_len>>3,
  //       s->compressed_len-7*last));
}

/* ===========================================================================
 * Save the match info and tally the frequency counts. Return true if
 * the current block must be flushed.
 */
function _tr_tally(s, dist, lc)
//    deflate_state *s;
//    unsigned dist;  /* distance of matched string */
//    unsigned lc;    /* match length-MIN_MATCH or unmatched char (if dist==0) */
{
  //var out_length, in_length, dcode;

  s.pending_buf[s.d_buf + s.last_lit * 2]     = (dist >>> 8) & 0xff;
  s.pending_buf[s.d_buf + s.last_lit * 2 + 1] = dist & 0xff;

  s.pending_buf[s.l_buf + s.last_lit] = lc & 0xff;
  s.last_lit++;

  if (dist === 0) {
    /* lc is the unmatched char */
    s.dyn_ltree[lc*2]/*.Freq*/++;
  } else {
    s.matches++;
    /* Here, lc is the match length - MIN_MATCH */
    dist--;             /* dist = match distance - 1 */
    //Assert((ush)dist < (ush)MAX_DIST(s) &&
    //       (ush)lc <= (ush)(MAX_MATCH-MIN_MATCH) &&
    //       (ush)d_code(dist) < (ush)D_CODES,  "_tr_tally: bad match");

    s.dyn_ltree[(_length_code[lc]+LITERALS+1) * 2]/*.Freq*/++;
    s.dyn_dtree[d_code(dist) * 2]/*.Freq*/++;
  }

// (!) This block is disabled in zlib defailts,
// don't enable it for binary compatibility

//#ifdef TRUNCATE_BLOCK
//  /* Try to guess if it is profitable to stop the current block here */
//  if ((s.last_lit & 0x1fff) === 0 && s.level > 2) {
//    /* Compute an upper bound for the compressed length */
//    out_length = s.last_lit*8;
//    in_length = s.strstart - s.block_start;
//
//    for (dcode = 0; dcode < D_CODES; dcode++) {
//      out_length += s.dyn_dtree[dcode*2]/*.Freq*/ * (5 + extra_dbits[dcode]);
//    }
//    out_length >>>= 3;
//    //Tracev((stderr,"\nlast_lit %u, in %ld, out ~%ld(%ld%%) ",
//    //       s->last_lit, in_length, out_length,
//    //       100L - out_length*100L/in_length));
//    if (s.matches < (s.last_lit>>1)/*int /2*/ && out_length < (in_length>>1)/*int /2*/) {
//      return true;
//    }
//  }
//#endif

  return (s.last_lit === s.lit_bufsize-1);
  /* We avoid equality with lit_bufsize because of wraparound at 64K
   * on 16 bit machines and because stored blocks are restricted to
   * 64K-1 bytes.
   */
}

exports._tr_init  = _tr_init;
exports._tr_stored_block = _tr_stored_block;
exports._tr_flush_block  = _tr_flush_block;
exports._tr_tally = _tr_tally;
exports._tr_align = _tr_align;
},{"../utils/common":27}],39:[function(_dereq_,module,exports){
'use strict';


function ZStream() {
  /* next input byte */
  this.input = null; // JS specific, because we have no pointers
  this.next_in = 0;
  /* number of bytes available at input */
  this.avail_in = 0;
  /* total number of input bytes read so far */
  this.total_in = 0;
  /* next output byte should be put there */
  this.output = null; // JS specific, because we have no pointers
  this.next_out = 0;
  /* remaining free space at output */
  this.avail_out = 0;
  /* total number of bytes output so far */
  this.total_out = 0;
  /* last error message, NULL if no error */
  this.msg = ''/*Z_NULL*/;
  /* not visible by applications */
  this.state = null;
  /* best guess about the data type: binary or text */
  this.data_type = 2/*Z_UNKNOWN*/;
  /* adler32 value of the uncompressed data */
  this.adler = 0;
}

module.exports = ZStream;
},{}]},{},[9])
(9)
}));
/*! xlsx.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
/*exported XLSX */
/*global global, exports, module, require:false, process:false, Buffer:false, ArrayBuffer:false */
var XLSX = {};
function make_xlsx_lib(XLSX){
XLSX.version = '0.14.2';
var current_codepage = 1200, current_ansi = 1252;
/*global cptable:true, window */
if(typeof module !== "undefined" && typeof require !== 'undefined') {
	if(typeof cptable === 'undefined') {
		if(typeof global !== 'undefined') global.cptable = undefined;
		else if(typeof window !== 'undefined') window.cptable = undefined;
	}
}

var VALID_ANSI = [ 874, 932, 936, 949, 950 ];
for(var i = 0; i <= 8; ++i) VALID_ANSI.push(1250 + i);
/* ECMA-376 Part I 18.4.1 charset to codepage mapping */
var CS2CP = ({
0:    1252, /* ANSI */
1:   65001, /* DEFAULT */
2:   65001, /* SYMBOL */
77:  10000, /* MAC */
128:   932, /* SHIFTJIS */
129:   949, /* HANGUL */
130:  1361, /* JOHAB */
134:   936, /* GB2312 */
136:   950, /* CHINESEBIG5 */
161:  1253, /* GREEK */
162:  1254, /* TURKISH */
163:  1258, /* VIETNAMESE */
177:  1255, /* HEBREW */
178:  1256, /* ARABIC */
186:  1257, /* BALTIC */
204:  1251, /* RUSSIAN */
222:   874, /* THAI */
238:  1250, /* EASTEUROPE */
255:  1252, /* OEM */
69:   6969  /* MISC */
});

var set_ansi = function(cp) { if(VALID_ANSI.indexOf(cp) == -1) return; current_ansi = CS2CP[0] = cp; };
function reset_ansi() { set_ansi(1252); }

var set_cp = function(cp) { current_codepage = cp; set_ansi(cp); };
function reset_cp() { set_cp(1200); reset_ansi(); }

function char_codes(data) { var o = []; for(var i = 0, len = data.length; i < len; ++i) o[i] = data.charCodeAt(i); return o; }

function utf16leread(data) {
	var o = [];
	for(var i = 0; i < (data.length>>1); ++i) o[i] = String.fromCharCode(data.charCodeAt(2*i) + (data.charCodeAt(2*i+1)<<8));
	return o.join("");
}
function utf16beread(data) {
	var o = [];
	for(var i = 0; i < (data.length>>1); ++i) o[i] = String.fromCharCode(data.charCodeAt(2*i+1) + (data.charCodeAt(2*i)<<8));
	return o.join("");
}

var debom = function(data) {
	var c1 = data.charCodeAt(0), c2 = data.charCodeAt(1);
	if(c1 == 0xFF && c2 == 0xFE) return utf16leread(data.slice(2));
	if(c1 == 0xFE && c2 == 0xFF) return utf16beread(data.slice(2));
	if(c1 == 0xFEFF) return data.slice(1);
	return data;
};

var _getchar = function _gc1(x) { return String.fromCharCode(x); };
if(typeof cptable !== 'undefined') {
	set_cp = function(cp) { current_codepage = cp; };
	debom = function(data) {
		if(data.charCodeAt(0) === 0xFF && data.charCodeAt(1) === 0xFE) { return cptable.utils.decode(1200, char_codes(data.slice(2))); }
		return data;
	};
	_getchar = function _gc2(x) {
		if(current_codepage === 1200) return String.fromCharCode(x);
		return cptable.utils.decode(current_codepage, [x&255,x>>8])[0];
	};
}
var DENSE = null;
var DIF_XL = true;
var Base64 = (function make_b64(){
	var map = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=";
	return {
		encode: function(input) {
			var o = "";
			var c1=0, c2=0, c3=0, e1=0, e2=0, e3=0, e4=0;
			for(var i = 0; i < input.length; ) {
				c1 = input.charCodeAt(i++);
				e1 = (c1 >> 2);

				c2 = input.charCodeAt(i++);
				e2 = ((c1 & 3) << 4) | (c2 >> 4);

				c3 = input.charCodeAt(i++);
				e3 = ((c2 & 15) << 2) | (c3 >> 6);
				e4 = (c3 & 63);
				if (isNaN(c2)) { e3 = e4 = 64; }
				else if (isNaN(c3)) { e4 = 64; }
				o += map.charAt(e1) + map.charAt(e2) + map.charAt(e3) + map.charAt(e4);
			}
			return o;
		},
		decode: function b64_decode(input) {
			var o = "";
			var c1=0, c2=0, c3=0, e1=0, e2=0, e3=0, e4=0;
			input = input.replace(/[^\w\+\/\=]/g, "");
			for(var i = 0; i < input.length;) {
				e1 = map.indexOf(input.charAt(i++));
				e2 = map.indexOf(input.charAt(i++));
				c1 = (e1 << 2) | (e2 >> 4);
				o += String.fromCharCode(c1);

				e3 = map.indexOf(input.charAt(i++));
				c2 = ((e2 & 15) << 4) | (e3 >> 2);
				if (e3 !== 64) { o += String.fromCharCode(c2); }

				e4 = map.indexOf(input.charAt(i++));
				c3 = ((e3 & 3) << 6) | e4;
				if (e4 !== 64) { o += String.fromCharCode(c3); }
			}
			return o;
		}
	};
})();
var has_buf = (typeof Buffer !== 'undefined' && typeof process !== 'undefined' && typeof process.versions !== 'undefined' && !!process.versions.node);

var Buffer_from = function(){};

if(typeof Buffer !== 'undefined') {
	var nbfs = !Buffer.from;
	if(!nbfs) try { Buffer.from("foo", "utf8"); } catch(e) { nbfs = true; }
	Buffer_from = nbfs ? function(buf, enc) { return (enc) ? new Buffer(buf, enc) : new Buffer(buf); } : Buffer.from.bind(Buffer);
	// $FlowIgnore
	if(!Buffer.alloc) Buffer.alloc = function(n) { return new Buffer(n); };
	// $FlowIgnore
	if(!Buffer.allocUnsafe) Buffer.allocUnsafe = function(n) { return new Buffer(n); };
}

function new_raw_buf(len) {
	/* jshint -W056 */
	return has_buf ? Buffer.alloc(len) : new Array(len);
	/* jshint +W056 */
}

function new_unsafe_buf(len) {
	/* jshint -W056 */
	return has_buf ? Buffer.allocUnsafe(len) : new Array(len);
	/* jshint +W056 */
}

var s2a = function s2a(s) {
	// $FlowIgnore
	if(has_buf) return Buffer_from(s, "binary");
	return s.split("").map(function(x){ return x.charCodeAt(0) & 0xff; });
};

function s2ab(s) {
	if(typeof ArrayBuffer === 'undefined') return s2a(s);
	var buf = new ArrayBuffer(s.length), view = new Uint8Array(buf);
	for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
	return buf;
}

function a2s(data) {
	if(Array.isArray(data)) return data.map(_chr).join("");
	var o = []; for(var i = 0; i < data.length; ++i) o[i] = _chr(data[i]); return o.join("");
}

function a2u(data) {
	if(typeof Uint8Array === 'undefined') throw new Error("Unsupported");
	return new Uint8Array(data);
}

function ab2a(data) {
	if(typeof ArrayBuffer == 'undefined') throw new Error("Unsupported");
	if(data instanceof ArrayBuffer) return ab2a(new Uint8Array(data));
var o = new Array(data.length);
	for(var i = 0; i < data.length; ++i) o[i] = data[i];
	return o;
}

var bconcat = function(bufs) { return [].concat.apply([], bufs); };

var chr0 = /\u0000/g, chr1 = /[\u0001-\u0006]/g;
/* ssf.js (C) 2013-present SheetJS -- http://sheetjs.com */
/*jshint -W041 */
var SSF = ({});
var make_ssf = function make_ssf(SSF){
SSF.version = '0.10.2';
function _strrev(x) { var o = "", i = x.length-1; while(i>=0) o += x.charAt(i--); return o; }
function fill(c,l) { var o = ""; while(o.length < l) o+=c; return o; }
function pad0(v,d){var t=""+v; return t.length>=d?t:fill('0',d-t.length)+t;}
function pad_(v,d){var t=""+v;return t.length>=d?t:fill(' ',d-t.length)+t;}
function rpad_(v,d){var t=""+v; return t.length>=d?t:t+fill(' ',d-t.length);}
function pad0r1(v,d){var t=""+Math.round(v); return t.length>=d?t:fill('0',d-t.length)+t;}
function pad0r2(v,d){var t=""+v; return t.length>=d?t:fill('0',d-t.length)+t;}
var p2_32 = Math.pow(2,32);
function pad0r(v,d){if(v>p2_32||v<-p2_32) return pad0r1(v,d); var i = Math.round(v); return pad0r2(i,d); }
function isgeneral(s, i) { i = i || 0; return s.length >= 7 + i && (s.charCodeAt(i)|32) === 103 && (s.charCodeAt(i+1)|32) === 101 && (s.charCodeAt(i+2)|32) === 110 && (s.charCodeAt(i+3)|32) === 101 && (s.charCodeAt(i+4)|32) === 114 && (s.charCodeAt(i+5)|32) === 97 && (s.charCodeAt(i+6)|32) === 108; }
var days = [
	['Sun', 'Sunday'],
	['Mon', 'Monday'],
	['Tue', 'Tuesday'],
	['Wed', 'Wednesday'],
	['Thu', 'Thursday'],
	['Fri', 'Friday'],
	['Sat', 'Saturday']
];
var months = [
	['J', 'Jan', 'January'],
	['F', 'Feb', 'February'],
	['M', 'Mar', 'March'],
	['A', 'Apr', 'April'],
	['M', 'May', 'May'],
	['J', 'Jun', 'June'],
	['J', 'Jul', 'July'],
	['A', 'Aug', 'August'],
	['S', 'Sep', 'September'],
	['O', 'Oct', 'October'],
	['N', 'Nov', 'November'],
	['D', 'Dec', 'December']
];
function init_table(t) {
	t[0]=  'General';
	t[1]=  '0';
	t[2]=  '0.00';
	t[3]=  '#,##0';
	t[4]=  '#,##0.00';
	t[9]=  '0%';
	t[10]= '0.00%';
	t[11]= '0.00E+00';
	t[12]= '# ?/?';
	t[13]= '# ??/??';
	t[14]= 'm/d/yy';
	t[15]= 'd-mmm-yy';
	t[16]= 'd-mmm';
	t[17]= 'mmm-yy';
	t[18]= 'h:mm AM/PM';
	t[19]= 'h:mm:ss AM/PM';
	t[20]= 'h:mm';
	t[21]= 'h:mm:ss';
	t[22]= 'm/d/yy h:mm';
	t[37]= '#,##0 ;(#,##0)';
	t[38]= '#,##0 ;[Red](#,##0)';
	t[39]= '#,##0.00;(#,##0.00)';
	t[40]= '#,##0.00;[Red](#,##0.00)';
	t[45]= 'mm:ss';
	t[46]= '[h]:mm:ss';
	t[47]= 'mmss.0';
	t[48]= '##0.0E+0';
	t[49]= '@';
	t[56]= '"上午/下午 "hh"時"mm"分"ss"秒 "';
	t[65535]= 'General';
}

var table_fmt = {};
init_table(table_fmt);
function frac(x, D, mixed) {
	var sgn = x < 0 ? -1 : 1;
	var B = x * sgn;
	var P_2 = 0, P_1 = 1, P = 0;
	var Q_2 = 1, Q_1 = 0, Q = 0;
	var A = Math.floor(B);
	while(Q_1 < D) {
		A = Math.floor(B);
		P = A * P_1 + P_2;
		Q = A * Q_1 + Q_2;
		if((B - A) < 0.00000005) break;
		B = 1 / (B - A);
		P_2 = P_1; P_1 = P;
		Q_2 = Q_1; Q_1 = Q;
	}
	if(Q > D) { if(Q_1 > D) { Q = Q_2; P = P_2; } else { Q = Q_1; P = P_1; } }
	if(!mixed) return [0, sgn * P, Q];
	var q = Math.floor(sgn * P/Q);
	return [q, sgn*P - q*Q, Q];
}
function parse_date_code(v,opts,b2) {
	if(v > 2958465 || v < 0) return null;
	var date = (v|0), time = Math.floor(86400 * (v - date)), dow=0;
	var dout=[];
	var out={D:date, T:time, u:86400*(v-date)-time,y:0,m:0,d:0,H:0,M:0,S:0,q:0};
	if(Math.abs(out.u) < 1e-6) out.u = 0;
	if(opts && opts.date1904) date += 1462;
	if(out.u > 0.9999) {
		out.u = 0;
		if(++time == 86400) { out.T = time = 0; ++date; ++out.D; }
	}
	if(date === 60) {dout = b2 ? [1317,10,29] : [1900,2,29]; dow=3;}
	else if(date === 0) {dout = b2 ? [1317,8,29] : [1900,1,0]; dow=6;}
	else {
		if(date > 60) --date;
		/* 1 = Jan 1 1900 in Gregorian */
		var d = new Date(1900, 0, 1);
		d.setDate(d.getDate() + date - 1);
		dout = [d.getFullYear(), d.getMonth()+1,d.getDate()];
		dow = d.getDay();
		if(date < 60) dow = (dow + 6) % 7;
		if(b2) dow = fix_hijri(d, dout);
	}
	out.y = dout[0]; out.m = dout[1]; out.d = dout[2];
	out.S = time % 60; time = Math.floor(time / 60);
	out.M = time % 60; time = Math.floor(time / 60);
	out.H = time;
	out.q = dow;
	return out;
}
SSF.parse_date_code = parse_date_code;
var basedate = new Date(1899, 11, 31, 0, 0, 0);
var dnthresh = basedate.getTime();
var base1904 = new Date(1900, 2, 1, 0, 0, 0);
function datenum_local(v, date1904) {
	var epoch = v.getTime();
	if(date1904) epoch -= 1461*24*60*60*1000;
	else if(v >= base1904) epoch += 24*60*60*1000;
	return (epoch - (dnthresh + (v.getTimezoneOffset() - basedate.getTimezoneOffset()) * 60000)) / (24 * 60 * 60 * 1000);
}
function general_fmt_int(v) { return v.toString(10); }
SSF._general_int = general_fmt_int;
var general_fmt_num = (function make_general_fmt_num() {
var gnr1 = /\.(\d*[1-9])0+$/, gnr2 = /\.0*$/, gnr4 = /\.(\d*[1-9])0+/, gnr5 = /\.0*[Ee]/, gnr6 = /(E[+-])(\d)$/;
function gfn2(v) {
	var w = (v<0?12:11);
	var o = gfn5(v.toFixed(12)); if(o.length <= w) return o;
	o = v.toPrecision(10); if(o.length <= w) return o;
	return v.toExponential(5);
}
function gfn3(v) {
	var o = v.toFixed(11).replace(gnr1,".$1");
	if(o.length > (v<0?12:11)) o = v.toPrecision(6);
	return o;
}
function gfn4(o) {
	for(var i = 0; i != o.length; ++i) if((o.charCodeAt(i) | 0x20) === 101) return o.replace(gnr4,".$1").replace(gnr5,"E").replace("e","E").replace(gnr6,"$10$2");
	return o;
}
function gfn5(o) {
	return o.indexOf(".") > -1 ? o.replace(gnr2,"").replace(gnr1,".$1") : o;
}
return function general_fmt_num(v) {
	var V = Math.floor(Math.log(Math.abs(v))*Math.LOG10E), o;
	if(V >= -4 && V <= -1) o = v.toPrecision(10+V);
	else if(Math.abs(V) <= 9) o = gfn2(v);
	else if(V === 10) o = v.toFixed(10).substr(0,12);
	else o = gfn3(v);
	return gfn5(gfn4(o));
};})();
SSF._general_num = general_fmt_num;
function general_fmt(v, opts) {
	switch(typeof v) {
		case 'string': return v;
		case 'boolean': return v ? "TRUE" : "FALSE";
		case 'number': return (v|0) === v ? general_fmt_int(v) : general_fmt_num(v);
		case 'undefined': return "";
		case 'object':
			if(v == null) return "";
			if(v instanceof Date) return format(14, datenum_local(v, opts && opts.date1904), opts);
	}
	throw new Error("unsupported value in General format: " + v);
}
SSF._general = general_fmt;
function fix_hijri() { return 0; }
/*jshint -W086 */
function write_date(type, fmt, val, ss0) {
	var o="", ss=0, tt=0, y = val.y, out, outl = 0;
	switch(type) {
		case 98: /* 'b' buddhist year */
			y = val.y + 543;
			/* falls through */
		case 121: /* 'y' year */
		switch(fmt.length) {
			case 1: case 2: out = y % 100; outl = 2; break;
			default: out = y % 10000; outl = 4; break;
		} break;
		case 109: /* 'm' month */
		switch(fmt.length) {
			case 1: case 2: out = val.m; outl = fmt.length; break;
			case 3: return months[val.m-1][1];
			case 5: return months[val.m-1][0];
			default: return months[val.m-1][2];
		} break;
		case 100: /* 'd' day */
		switch(fmt.length) {
			case 1: case 2: out = val.d; outl = fmt.length; break;
			case 3: return days[val.q][0];
			default: return days[val.q][1];
		} break;
		case 104: /* 'h' 12-hour */
		switch(fmt.length) {
			case 1: case 2: out = 1+(val.H+11)%12; outl = fmt.length; break;
			default: throw 'bad hour format: ' + fmt;
		} break;
		case 72: /* 'H' 24-hour */
		switch(fmt.length) {
			case 1: case 2: out = val.H; outl = fmt.length; break;
			default: throw 'bad hour format: ' + fmt;
		} break;
		case 77: /* 'M' minutes */
		switch(fmt.length) {
			case 1: case 2: out = val.M; outl = fmt.length; break;
			default: throw 'bad minute format: ' + fmt;
		} break;
		case 115: /* 's' seconds */
			if(fmt != 's' && fmt != 'ss' && fmt != '.0' && fmt != '.00' && fmt != '.000') throw 'bad second format: ' + fmt;
			if(val.u === 0 && (fmt == "s" || fmt == "ss")) return pad0(val.S, fmt.length);
if(ss0 >= 2) tt = ss0 === 3 ? 1000 : 100;
			else tt = ss0 === 1 ? 10 : 1;
			ss = Math.round((tt)*(val.S + val.u));
			if(ss >= 60*tt) ss = 0;
			if(fmt === 's') return ss === 0 ? "0" : ""+ss/tt;
			o = pad0(ss,2 + ss0);
			if(fmt === 'ss') return o.substr(0,2);
			return "." + o.substr(2,fmt.length-1);
		case 90: /* 'Z' absolute time */
		switch(fmt) {
			case '[h]': case '[hh]': out = val.D*24+val.H; break;
			case '[m]': case '[mm]': out = (val.D*24+val.H)*60+val.M; break;
			case '[s]': case '[ss]': out = ((val.D*24+val.H)*60+val.M)*60+Math.round(val.S+val.u); break;
			default: throw 'bad abstime format: ' + fmt;
		} outl = fmt.length === 3 ? 1 : 2; break;
		case 101: /* 'e' era */
			out = y; outl = 1;
	}
	if(outl > 0) return pad0(out, outl); else return "";
}
/*jshint +W086 */
function commaify(s) {
	var w = 3;
	if(s.length <= w) return s;
	var j = (s.length % w), o = s.substr(0,j);
	for(; j!=s.length; j+=w) o+=(o.length > 0 ? "," : "") + s.substr(j,w);
	return o;
}
var write_num = (function make_write_num(){
var pct1 = /%/g;
function write_num_pct(type, fmt, val){
	var sfmt = fmt.replace(pct1,""), mul = fmt.length - sfmt.length;
	return write_num(type, sfmt, val * Math.pow(10,2*mul)) + fill("%",mul);
}
function write_num_cm(type, fmt, val){
	var idx = fmt.length - 1;
	while(fmt.charCodeAt(idx-1) === 44) --idx;
	return write_num(type, fmt.substr(0,idx), val / Math.pow(10,3*(fmt.length-idx)));
}
function write_num_exp(fmt, val){
	var o;
	var idx = fmt.indexOf("E") - fmt.indexOf(".") - 1;
	if(fmt.match(/^#+0.0E\+0$/)) {
		if(val == 0) return "0.0E+0";
		else if(val < 0) return "-" + write_num_exp(fmt, -val);
		var period = fmt.indexOf("."); if(period === -1) period=fmt.indexOf('E');
		var ee = Math.floor(Math.log(val)*Math.LOG10E)%period;
		if(ee < 0) ee += period;
		o = (val/Math.pow(10,ee)).toPrecision(idx+1+(period+ee)%period);
		if(o.indexOf("e") === -1) {
			var fakee = Math.floor(Math.log(val)*Math.LOG10E);
			if(o.indexOf(".") === -1) o = o.charAt(0) + "." + o.substr(1) + "E+" + (fakee - o.length+ee);
			else o += "E+" + (fakee - ee);
			while(o.substr(0,2) === "0.") {
				o = o.charAt(0) + o.substr(2,period) + "." + o.substr(2+period);
				o = o.replace(/^0+([1-9])/,"$1").replace(/^0+\./,"0.");
			}
			o = o.replace(/\+-/,"-");
		}
		o = o.replace(/^([+-]?)(\d*)\.(\d*)[Ee]/,function($$,$1,$2,$3) { return $1 + $2 + $3.substr(0,(period+ee)%period) + "." + $3.substr(ee) + "E"; });
	} else o = val.toExponential(idx);
	if(fmt.match(/E\+00$/) && o.match(/e[+-]\d$/)) o = o.substr(0,o.length-1) + "0" + o.charAt(o.length-1);
	if(fmt.match(/E\-/) && o.match(/e\+/)) o = o.replace(/e\+/,"e");
	return o.replace("e","E");
}
var frac1 = /# (\?+)( ?)\/( ?)(\d+)/;
function write_num_f1(r, aval, sign) {
	var den = parseInt(r[4],10), rr = Math.round(aval * den), base = Math.floor(rr/den);
	var myn = (rr - base*den), myd = den;
	return sign + (base === 0 ? "" : ""+base) + " " + (myn === 0 ? fill(" ", r[1].length + 1 + r[4].length) : pad_(myn,r[1].length) + r[2] + "/" + r[3] + pad0(myd,r[4].length));
}
function write_num_f2(r, aval, sign) {
	return sign + (aval === 0 ? "" : ""+aval) + fill(" ", r[1].length + 2 + r[4].length);
}
var dec1 = /^#*0*\.([0#]+)/;
var closeparen = /\).*[0#]/;
var phone = /\(###\) ###\\?-####/;
function hashq(str) {
	var o = "", cc;
	for(var i = 0; i != str.length; ++i) switch((cc=str.charCodeAt(i))) {
		case 35: break;
		case 63: o+= " "; break;
		case 48: o+= "0"; break;
		default: o+= String.fromCharCode(cc);
	}
	return o;
}
function rnd(val, d) { var dd = Math.pow(10,d); return ""+(Math.round(val * dd)/dd); }
function dec(val, d) {
	if (d < ('' + Math.round((val-Math.floor(val))*Math.pow(10,d))).length) {
		return 0;
	}
	return Math.round((val-Math.floor(val))*Math.pow(10,d));
}
function carry(val, d) {
	if (d < ('' + Math.round((val-Math.floor(val))*Math.pow(10,d))).length) {
		return 1;
	}
	return 0;
}
function flr(val) { if(val < 2147483647 && val > -2147483648) return ""+(val >= 0 ? (val|0) : (val-1|0)); return ""+Math.floor(val); }
function write_num_flt(type, fmt, val) {
	if(type.charCodeAt(0) === 40 && !fmt.match(closeparen)) {
		var ffmt = fmt.replace(/\( */,"").replace(/ \)/,"").replace(/\)/,"");
		if(val >= 0) return write_num_flt('n', ffmt, val);
		return '(' + write_num_flt('n', ffmt, -val) + ')';
	}
	if(fmt.charCodeAt(fmt.length - 1) === 44) return write_num_cm(type, fmt, val);
	if(fmt.indexOf('%') !== -1) return write_num_pct(type, fmt, val);
	if(fmt.indexOf('E') !== -1) return write_num_exp(fmt, val);
	if(fmt.charCodeAt(0) === 36) return "$"+write_num_flt(type,fmt.substr(fmt.charAt(1)==' '?2:1),val);
	var o;
	var r, ri, ff, aval = Math.abs(val), sign = val < 0 ? "-" : "";
	if(fmt.match(/^00+$/)) return sign + pad0r(aval,fmt.length);
	if(fmt.match(/^[#?]+$/)) {
		o = pad0r(val,0); if(o === "0") o = "";
		return o.length > fmt.length ? o : hashq(fmt.substr(0,fmt.length-o.length)) + o;
	}
	if((r = fmt.match(frac1))) return write_num_f1(r, aval, sign);
	if(fmt.match(/^#+0+$/)) return sign + pad0r(aval,fmt.length - fmt.indexOf("0"));
	if((r = fmt.match(dec1))) {
		o = rnd(val, r[1].length).replace(/^([^\.]+)$/,"$1."+hashq(r[1])).replace(/\.$/,"."+hashq(r[1])).replace(/\.(\d*)$/,function($$, $1) { return "." + $1 + fill("0", hashq(r[1]).length-$1.length); });
		return fmt.indexOf("0.") !== -1 ? o : o.replace(/^0\./,".");
	}
	fmt = fmt.replace(/^#+([0.])/, "$1");
	if((r = fmt.match(/^(0*)\.(#*)$/))) {
		return sign + rnd(aval, r[2].length).replace(/\.(\d*[1-9])0*$/,".$1").replace(/^(-?\d*)$/,"$1.").replace(/^0\./,r[1].length?"0.":".");
	}
	if((r = fmt.match(/^#{1,3},##0(\.?)$/))) return sign + commaify(pad0r(aval,0));
	if((r = fmt.match(/^#,##0\.([#0]*0)$/))) {
		return val < 0 ? "-" + write_num_flt(type, fmt, -val) : commaify(""+(Math.floor(val) + carry(val, r[1].length))) + "." + pad0(dec(val, r[1].length),r[1].length);
	}
	if((r = fmt.match(/^#,#*,#0/))) return write_num_flt(type,fmt.replace(/^#,#*,/,""),val);
	if((r = fmt.match(/^([0#]+)(\\?-([0#]+))+$/))) {
		o = _strrev(write_num_flt(type, fmt.replace(/[\\-]/g,""), val));
		ri = 0;
		return _strrev(_strrev(fmt.replace(/\\/g,"")).replace(/[0#]/g,function(x){return ri<o.length?o.charAt(ri++):x==='0'?'0':"";}));
	}
	if(fmt.match(phone)) {
		o = write_num_flt(type, "##########", val);
		return "(" + o.substr(0,3) + ") " + o.substr(3, 3) + "-" + o.substr(6);
	}
	var oa = "";
	if((r = fmt.match(/^([#0?]+)( ?)\/( ?)([#0?]+)/))) {
		ri = Math.min(r[4].length,7);
		ff = frac(aval, Math.pow(10,ri)-1, false);
		o = "" + sign;
		oa = write_num("n", r[1], ff[1]);
		if(oa.charAt(oa.length-1) == " ") oa = oa.substr(0,oa.length-1) + "0";
		o += oa + r[2] + "/" + r[3];
		oa = rpad_(ff[2],ri);
		if(oa.length < r[4].length) oa = hashq(r[4].substr(r[4].length-oa.length)) + oa;
		o += oa;
		return o;
	}
	if((r = fmt.match(/^# ([#0?]+)( ?)\/( ?)([#0?]+)/))) {
		ri = Math.min(Math.max(r[1].length, r[4].length),7);
		ff = frac(aval, Math.pow(10,ri)-1, true);
		return sign + (ff[0]||(ff[1] ? "" : "0")) + " " + (ff[1] ? pad_(ff[1],ri) + r[2] + "/" + r[3] + rpad_(ff[2],ri): fill(" ", 2*ri+1 + r[2].length + r[3].length));
	}
	if((r = fmt.match(/^[#0?]+$/))) {
		o = pad0r(val, 0);
		if(fmt.length <= o.length) return o;
		return hashq(fmt.substr(0,fmt.length-o.length)) + o;
	}
	if((r = fmt.match(/^([#0?]+)\.([#0]+)$/))) {
		o = "" + val.toFixed(Math.min(r[2].length,10)).replace(/([^0])0+$/,"$1");
		ri = o.indexOf(".");
		var lres = fmt.indexOf(".") - ri, rres = fmt.length - o.length - lres;
		return hashq(fmt.substr(0,lres) + o + fmt.substr(fmt.length-rres));
	}
	if((r = fmt.match(/^00,000\.([#0]*0)$/))) {
		ri = dec(val, r[1].length);
		return val < 0 ? "-" + write_num_flt(type, fmt, -val) : commaify(flr(val)).replace(/^\d,\d{3}$/,"0$&").replace(/^\d*$/,function($$) { return "00," + ($$.length < 3 ? pad0(0,3-$$.length) : "") + $$; }) + "." + pad0(ri,r[1].length);
	}
	switch(fmt) {
		case "###,##0.00": return write_num_flt(type, "#,##0.00", val);
		case "###,###":
		case "##,###":
		case "#,###": var x = commaify(pad0r(aval,0)); return x !== "0" ? sign + x : "";
		case "###,###.00": return write_num_flt(type, "###,##0.00",val).replace(/^0\./,".");
		case "#,###.00": return write_num_flt(type, "#,##0.00",val).replace(/^0\./,".");
		default:
	}
	throw new Error("unsupported format |" + fmt + "|");
}
function write_num_cm2(type, fmt, val){
	var idx = fmt.length - 1;
	while(fmt.charCodeAt(idx-1) === 44) --idx;
	return write_num(type, fmt.substr(0,idx), val / Math.pow(10,3*(fmt.length-idx)));
}
function write_num_pct2(type, fmt, val){
	var sfmt = fmt.replace(pct1,""), mul = fmt.length - sfmt.length;
	return write_num(type, sfmt, val * Math.pow(10,2*mul)) + fill("%",mul);
}
function write_num_exp2(fmt, val){
	var o;
	var idx = fmt.indexOf("E") - fmt.indexOf(".") - 1;
	if(fmt.match(/^#+0.0E\+0$/)) {
		if(val == 0) return "0.0E+0";
		else if(val < 0) return "-" + write_num_exp2(fmt, -val);
		var period = fmt.indexOf("."); if(period === -1) period=fmt.indexOf('E');
		var ee = Math.floor(Math.log(val)*Math.LOG10E)%period;
		if(ee < 0) ee += period;
		o = (val/Math.pow(10,ee)).toPrecision(idx+1+(period+ee)%period);
		if(!o.match(/[Ee]/)) {
			var fakee = Math.floor(Math.log(val)*Math.LOG10E);
			if(o.indexOf(".") === -1) o = o.charAt(0) + "." + o.substr(1) + "E+" + (fakee - o.length+ee);
			else o += "E+" + (fakee - ee);
			o = o.replace(/\+-/,"-");
		}
		o = o.replace(/^([+-]?)(\d*)\.(\d*)[Ee]/,function($$,$1,$2,$3) { return $1 + $2 + $3.substr(0,(period+ee)%period) + "." + $3.substr(ee) + "E"; });
	} else o = val.toExponential(idx);
	if(fmt.match(/E\+00$/) && o.match(/e[+-]\d$/)) o = o.substr(0,o.length-1) + "0" + o.charAt(o.length-1);
	if(fmt.match(/E\-/) && o.match(/e\+/)) o = o.replace(/e\+/,"e");
	return o.replace("e","E");
}
function write_num_int(type, fmt, val) {
	if(type.charCodeAt(0) === 40 && !fmt.match(closeparen)) {
		var ffmt = fmt.replace(/\( */,"").replace(/ \)/,"").replace(/\)/,"");
		if(val >= 0) return write_num_int('n', ffmt, val);
		return '(' + write_num_int('n', ffmt, -val) + ')';
	}
	if(fmt.charCodeAt(fmt.length - 1) === 44) return write_num_cm2(type, fmt, val);
	if(fmt.indexOf('%') !== -1) return write_num_pct2(type, fmt, val);
	if(fmt.indexOf('E') !== -1) return write_num_exp2(fmt, val);
	if(fmt.charCodeAt(0) === 36) return "$"+write_num_int(type,fmt.substr(fmt.charAt(1)==' '?2:1),val);
	var o;
	var r, ri, ff, aval = Math.abs(val), sign = val < 0 ? "-" : "";
	if(fmt.match(/^00+$/)) return sign + pad0(aval,fmt.length);
	if(fmt.match(/^[#?]+$/)) {
		o = (""+val); if(val === 0) o = "";
		return o.length > fmt.length ? o : hashq(fmt.substr(0,fmt.length-o.length)) + o;
	}
	if((r = fmt.match(frac1))) return write_num_f2(r, aval, sign);
	if(fmt.match(/^#+0+$/)) return sign + pad0(aval,fmt.length - fmt.indexOf("0"));
	if((r = fmt.match(dec1))) {
o = (""+val).replace(/^([^\.]+)$/,"$1."+hashq(r[1])).replace(/\.$/,"."+hashq(r[1]));
		o = o.replace(/\.(\d*)$/,function($$, $1) {
return "." + $1 + fill("0", hashq(r[1]).length-$1.length); });
		return fmt.indexOf("0.") !== -1 ? o : o.replace(/^0\./,".");
	}
	fmt = fmt.replace(/^#+([0.])/, "$1");
	if((r = fmt.match(/^(0*)\.(#*)$/))) {
		return sign + (""+aval).replace(/\.(\d*[1-9])0*$/,".$1").replace(/^(-?\d*)$/,"$1.").replace(/^0\./,r[1].length?"0.":".");
	}
	if((r = fmt.match(/^#{1,3},##0(\.?)$/))) return sign + commaify((""+aval));
	if((r = fmt.match(/^#,##0\.([#0]*0)$/))) {
		return val < 0 ? "-" + write_num_int(type, fmt, -val) : commaify((""+val)) + "." + fill('0',r[1].length);
	}
	if((r = fmt.match(/^#,#*,#0/))) return write_num_int(type,fmt.replace(/^#,#*,/,""),val);
	if((r = fmt.match(/^([0#]+)(\\?-([0#]+))+$/))) {
		o = _strrev(write_num_int(type, fmt.replace(/[\\-]/g,""), val));
		ri = 0;
		return _strrev(_strrev(fmt.replace(/\\/g,"")).replace(/[0#]/g,function(x){return ri<o.length?o.charAt(ri++):x==='0'?'0':"";}));
	}
	if(fmt.match(phone)) {
		o = write_num_int(type, "##########", val);
		return "(" + o.substr(0,3) + ") " + o.substr(3, 3) + "-" + o.substr(6);
	}
	var oa = "";
	if((r = fmt.match(/^([#0?]+)( ?)\/( ?)([#0?]+)/))) {
		ri = Math.min(r[4].length,7);
		ff = frac(aval, Math.pow(10,ri)-1, false);
		o = "" + sign;
		oa = write_num("n", r[1], ff[1]);
		if(oa.charAt(oa.length-1) == " ") oa = oa.substr(0,oa.length-1) + "0";
		o += oa + r[2] + "/" + r[3];
		oa = rpad_(ff[2],ri);
		if(oa.length < r[4].length) oa = hashq(r[4].substr(r[4].length-oa.length)) + oa;
		o += oa;
		return o;
	}
	if((r = fmt.match(/^# ([#0?]+)( ?)\/( ?)([#0?]+)/))) {
		ri = Math.min(Math.max(r[1].length, r[4].length),7);
		ff = frac(aval, Math.pow(10,ri)-1, true);
		return sign + (ff[0]||(ff[1] ? "" : "0")) + " " + (ff[1] ? pad_(ff[1],ri) + r[2] + "/" + r[3] + rpad_(ff[2],ri): fill(" ", 2*ri+1 + r[2].length + r[3].length));
	}
	if((r = fmt.match(/^[#0?]+$/))) {
		o = "" + val;
		if(fmt.length <= o.length) return o;
		return hashq(fmt.substr(0,fmt.length-o.length)) + o;
	}
	if((r = fmt.match(/^([#0]+)\.([#0]+)$/))) {
		o = "" + val.toFixed(Math.min(r[2].length,10)).replace(/([^0])0+$/,"$1");
		ri = o.indexOf(".");
		var lres = fmt.indexOf(".") - ri, rres = fmt.length - o.length - lres;
		return hashq(fmt.substr(0,lres) + o + fmt.substr(fmt.length-rres));
	}
	if((r = fmt.match(/^00,000\.([#0]*0)$/))) {
		return val < 0 ? "-" + write_num_int(type, fmt, -val) : commaify(""+val).replace(/^\d,\d{3}$/,"0$&").replace(/^\d*$/,function($$) { return "00," + ($$.length < 3 ? pad0(0,3-$$.length) : "") + $$; }) + "." + pad0(0,r[1].length);
	}
	switch(fmt) {
		case "###,###":
		case "##,###":
		case "#,###": var x = commaify(""+aval); return x !== "0" ? sign + x : "";
		default:
			if(fmt.match(/\.[0#?]*$/)) return write_num_int(type, fmt.slice(0,fmt.lastIndexOf(".")), val) + hashq(fmt.slice(fmt.lastIndexOf(".")));
	}
	throw new Error("unsupported format |" + fmt + "|");
}
return function write_num(type, fmt, val) {
	return (val|0) === val ? write_num_int(type, fmt, val) : write_num_flt(type, fmt, val);
};})();
function split_fmt(fmt) {
	var out = [];
	var in_str = false/*, cc*/;
	for(var i = 0, j = 0; i < fmt.length; ++i) switch((/*cc=*/fmt.charCodeAt(i))) {
		case 34: /* '"' */
			in_str = !in_str; break;
		case 95: case 42: case 92: /* '_' '*' '\\' */
			++i; break;
		case 59: /* ';' */
			out[out.length] = fmt.substr(j,i-j);
			j = i+1;
	}
	out[out.length] = fmt.substr(j);
	if(in_str === true) throw new Error("Format |" + fmt + "| unterminated string ");
	return out;
}
SSF._split = split_fmt;
var abstime = /\[[HhMmSs]*\]/;
function fmt_is_date(fmt) {
	var i = 0, /*cc = 0,*/ c = "", o = "";
	while(i < fmt.length) {
		switch((c = fmt.charAt(i))) {
			case 'G': if(isgeneral(fmt, i)) i+= 6; i++; break;
			case '"': for(;(/*cc=*/fmt.charCodeAt(++i)) !== 34 && i < fmt.length;) ++i; ++i; break;
			case '\\': i+=2; break;
			case '_': i+=2; break;
			case '@': ++i; break;
			case 'B': case 'b':
				if(fmt.charAt(i+1) === "1" || fmt.charAt(i+1) === "2") return true;
				/* falls through */
			case 'M': case 'D': case 'Y': case 'H': case 'S': case 'E':
				/* falls through */
			case 'm': case 'd': case 'y': case 'h': case 's': case 'e': case 'g': return true;
			case 'A': case 'a':
				if(fmt.substr(i, 3).toUpperCase() === "A/P") return true;
				if(fmt.substr(i, 5).toUpperCase() === "AM/PM") return true;
				++i; break;
			case '[':
				o = c;
				while(fmt.charAt(i++) !== ']' && i < fmt.length) o += fmt.charAt(i);
				if(o.match(abstime)) return true;
				break;
			case '.':
				/* falls through */
			case '0': case '#':
				while(i < fmt.length && ("0#?.,E+-%".indexOf(c=fmt.charAt(++i)) > -1 || (c=='\\' && fmt.charAt(i+1) == "-" && "0#".indexOf(fmt.charAt(i+2))>-1))){/* empty */}
				break;
			case '?': while(fmt.charAt(++i) === c){/* empty */} break;
			case '*': ++i; if(fmt.charAt(i) == ' ' || fmt.charAt(i) == '*') ++i; break;
			case '(': case ')': ++i; break;
			case '1': case '2': case '3': case '4': case '5': case '6': case '7': case '8': case '9':
				while(i < fmt.length && "0123456789".indexOf(fmt.charAt(++i)) > -1){/* empty */} break;
			case ' ': ++i; break;
			default: ++i; break;
		}
	}
	return false;
}
SSF.is_date = fmt_is_date;
function eval_fmt(fmt, v, opts, flen) {
	var out = [], o = "", i = 0, c = "", lst='t', dt, j, cc;
	var hr='H';
	/* Tokenize */
	while(i < fmt.length) {
		switch((c = fmt.charAt(i))) {
			case 'G': /* General */
				if(!isgeneral(fmt, i)) throw new Error('unrecognized character ' + c + ' in ' +fmt);
				out[out.length] = {t:'G', v:'General'}; i+=7; break;
			case '"': /* Literal text */
				for(o="";(cc=fmt.charCodeAt(++i)) !== 34 && i < fmt.length;) o += String.fromCharCode(cc);
				out[out.length] = {t:'t', v:o}; ++i; break;
			case '\\': var w = fmt.charAt(++i), t = (w === "(" || w === ")") ? w : 't';
				out[out.length] = {t:t, v:w}; ++i; break;
			case '_': out[out.length] = {t:'t', v:" "}; i+=2; break;
			case '@': /* Text Placeholder */
				out[out.length] = {t:'T', v:v}; ++i; break;
			case 'B': case 'b':
				if(fmt.charAt(i+1) === "1" || fmt.charAt(i+1) === "2") {
					if(dt==null) { dt=parse_date_code(v, opts, fmt.charAt(i+1) === "2"); if(dt==null) return ""; }
					out[out.length] = {t:'X', v:fmt.substr(i,2)}; lst = c; i+=2; break;
				}
				/* falls through */
			case 'M': case 'D': case 'Y': case 'H': case 'S': case 'E':
				c = c.toLowerCase();
				/* falls through */
			case 'm': case 'd': case 'y': case 'h': case 's': case 'e': case 'g':
				if(v < 0) return "";
				if(dt==null) { dt=parse_date_code(v, opts); if(dt==null) return ""; }
				o = c; while(++i < fmt.length && fmt.charAt(i).toLowerCase() === c) o+=c;
				if(c === 'm' && lst.toLowerCase() === 'h') c = 'M';
				if(c === 'h') c = hr;
				out[out.length] = {t:c, v:o}; lst = c; break;
			case 'A': case 'a':
				var q={t:c, v:c};
				if(dt==null) dt=parse_date_code(v, opts);
				if(fmt.substr(i, 3).toUpperCase() === "A/P") { if(dt!=null) q.v = dt.H >= 12 ? "P" : "A"; q.t = 'T'; hr='h';i+=3;}
				else if(fmt.substr(i,5).toUpperCase() === "AM/PM") { if(dt!=null) q.v = dt.H >= 12 ? "PM" : "AM"; q.t = 'T'; i+=5; hr='h'; }
				else { q.t = "t"; ++i; }
				if(dt==null && q.t === 'T') return "";
				out[out.length] = q; lst = c; break;
			case '[':
				o = c;
				while(fmt.charAt(i++) !== ']' && i < fmt.length) o += fmt.charAt(i);
				if(o.slice(-1) !== ']') throw 'unterminated "[" block: |' + o + '|';
				if(o.match(abstime)) {
					if(dt==null) { dt=parse_date_code(v, opts); if(dt==null) return ""; }
					out[out.length] = {t:'Z', v:o.toLowerCase()};
					lst = o.charAt(1);
				} else if(o.indexOf("$") > -1) {
					o = (o.match(/\$([^-\[\]]*)/)||[])[1]||"$";
					if(!fmt_is_date(fmt)) out[out.length] = {t:'t',v:o};
				}
				break;
			/* Numbers */
			case '.':
				if(dt != null) {
					o = c; while(++i < fmt.length && (c=fmt.charAt(i)) === "0") o += c;
					out[out.length] = {t:'s', v:o}; break;
				}
				/* falls through */
			case '0': case '#':
				o = c; while((++i < fmt.length && "0#?.,E+-%".indexOf(c=fmt.charAt(i)) > -1) || (c=='\\' && fmt.charAt(i+1) == "-" && i < fmt.length - 2 && "0#".indexOf(fmt.charAt(i+2))>-1)) o += c;
				out[out.length] = {t:'n', v:o}; break;
			case '?':
				o = c; while(fmt.charAt(++i) === c) o+=c;
				out[out.length] = {t:c, v:o}; lst = c; break;
			case '*': ++i; if(fmt.charAt(i) == ' ' || fmt.charAt(i) == '*') ++i; break; // **
			case '(': case ')': out[out.length] = {t:(flen===1?'t':c), v:c}; ++i; break;
			case '1': case '2': case '3': case '4': case '5': case '6': case '7': case '8': case '9':
				o = c; while(i < fmt.length && "0123456789".indexOf(fmt.charAt(++i)) > -1) o+=fmt.charAt(i);
				out[out.length] = {t:'D', v:o}; break;
			case ' ': out[out.length] = {t:c, v:c}; ++i; break;
			default:
				if(",$-+/():!^&'~{}<>=€acfijklopqrtuvwxzP".indexOf(c) === -1) throw new Error('unrecognized character ' + c + ' in ' + fmt);
				out[out.length] = {t:'t', v:c}; ++i; break;
		}
	}
	var bt = 0, ss0 = 0, ssm;
	for(i=out.length-1, lst='t'; i >= 0; --i) {
		switch(out[i].t) {
			case 'h': case 'H': out[i].t = hr; lst='h'; if(bt < 1) bt = 1; break;
			case 's':
				if((ssm=out[i].v.match(/\.0+$/))) ss0=Math.max(ss0,ssm[0].length-1);
				if(bt < 3) bt = 3;
			/* falls through */
			case 'd': case 'y': case 'M': case 'e': lst=out[i].t; break;
			case 'm': if(lst === 's') { out[i].t = 'M'; if(bt < 2) bt = 2; } break;
			case 'X': /*if(out[i].v === "B2");*/
				break;
			case 'Z':
				if(bt < 1 && out[i].v.match(/[Hh]/)) bt = 1;
				if(bt < 2 && out[i].v.match(/[Mm]/)) bt = 2;
				if(bt < 3 && out[i].v.match(/[Ss]/)) bt = 3;
		}
	}
	switch(bt) {
		case 0: break;
		case 1:
if(dt.u >= 0.5) { dt.u = 0; ++dt.S; }
			if(dt.S >=  60) { dt.S = 0; ++dt.M; }
			if(dt.M >=  60) { dt.M = 0; ++dt.H; }
			break;
		case 2:
if(dt.u >= 0.5) { dt.u = 0; ++dt.S; }
			if(dt.S >=  60) { dt.S = 0; ++dt.M; }
			break;
	}
	/* replace fields */
	var nstr = "", jj;
	for(i=0; i < out.length; ++i) {
		switch(out[i].t) {
			case 't': case 'T': case ' ': case 'D': break;
			case 'X': out[i].v = ""; out[i].t = ";"; break;
			case 'd': case 'm': case 'y': case 'h': case 'H': case 'M': case 's': case 'e': case 'b': case 'Z':
out[i].v = write_date(out[i].t.charCodeAt(0), out[i].v, dt, ss0);
				out[i].t = 't'; break;
			case 'n': case '(': case '?':
				jj = i+1;
				while(out[jj] != null && (
					(c=out[jj].t) === "?" || c === "D" ||
					((c === " " || c === "t") && out[jj+1] != null && (out[jj+1].t === '?' || out[jj+1].t === "t" && out[jj+1].v === '/')) ||
					(out[i].t === '(' && (c === ' ' || c === 'n' || c === ')')) ||
					(c === 't' && (out[jj].v === '/' || out[jj].v === ' ' && out[jj+1] != null && out[jj+1].t == '?'))
				)) {
					out[i].v += out[jj].v;
					out[jj] = {v:"", t:";"}; ++jj;
				}
				nstr += out[i].v;
				i = jj-1; break;
			case 'G': out[i].t = 't'; out[i].v = general_fmt(v,opts); break;
		}
	}
	var vv = "", myv, ostr;
	if(nstr.length > 0) {
		if(nstr.charCodeAt(0) == 40) /* '(' */ {
			myv = (v<0&&nstr.charCodeAt(0) === 45 ? -v : v);
			ostr = write_num('(', nstr, myv);
		} else {
			myv = (v<0 && flen > 1 ? -v : v);
			ostr = write_num('n', nstr, myv);
			if(myv < 0 && out[0] && out[0].t == 't') {
				ostr = ostr.substr(1);
				out[0].v = "-" + out[0].v;
			}
		}
		jj=ostr.length-1;
		var decpt = out.length;
		for(i=0; i < out.length; ++i) if(out[i] != null && out[i].t != 't' && out[i].v.indexOf(".") > -1) { decpt = i; break; }
		var lasti=out.length;
		if(decpt === out.length && ostr.indexOf("E") === -1) {
			for(i=out.length-1; i>= 0;--i) {
				if(out[i] == null || 'n?('.indexOf(out[i].t) === -1) continue;
				if(jj>=out[i].v.length-1) { jj -= out[i].v.length; out[i].v = ostr.substr(jj+1, out[i].v.length); }
				else if(jj < 0) out[i].v = "";
				else { out[i].v = ostr.substr(0, jj+1); jj = -1; }
				out[i].t = 't';
				lasti = i;
			}
			if(jj>=0 && lasti<out.length) out[lasti].v = ostr.substr(0,jj+1) + out[lasti].v;
		}
		else if(decpt !== out.length && ostr.indexOf("E") === -1) {
			jj = ostr.indexOf(".")-1;
			for(i=decpt; i>= 0; --i) {
				if(out[i] == null || 'n?('.indexOf(out[i].t) === -1) continue;
				j=out[i].v.indexOf(".")>-1&&i===decpt?out[i].v.indexOf(".")-1:out[i].v.length-1;
				vv = out[i].v.substr(j+1);
				for(; j>=0; --j) {
					if(jj>=0 && (out[i].v.charAt(j) === "0" || out[i].v.charAt(j) === "#")) vv = ostr.charAt(jj--) + vv;
				}
				out[i].v = vv;
				out[i].t = 't';
				lasti = i;
			}
			if(jj>=0 && lasti<out.length) out[lasti].v = ostr.substr(0,jj+1) + out[lasti].v;
			jj = ostr.indexOf(".")+1;
			for(i=decpt; i<out.length; ++i) {
				if(out[i] == null || ('n?('.indexOf(out[i].t) === -1 && i !== decpt)) continue;
				j=out[i].v.indexOf(".")>-1&&i===decpt?out[i].v.indexOf(".")+1:0;
				vv = out[i].v.substr(0,j);
				for(; j<out[i].v.length; ++j) {
					if(jj<ostr.length) vv += ostr.charAt(jj++);
				}
				out[i].v = vv;
				out[i].t = 't';
				lasti = i;
			}
		}
	}
	for(i=0; i<out.length; ++i) if(out[i] != null && 'n(?'.indexOf(out[i].t)>-1) {
		myv = (flen >1 && v < 0 && i>0 && out[i-1].v === "-" ? -v:v);
		out[i].v = write_num(out[i].t, out[i].v, myv);
		out[i].t = 't';
	}
	var retval = "";
	for(i=0; i !== out.length; ++i) if(out[i] != null) retval += out[i].v;
	return retval;
}
SSF._eval = eval_fmt;
var cfregex = /\[[=<>]/;
var cfregex2 = /\[(=|>[=]?|<[>=]?)(-?\d+(?:\.\d*)?)\]/;
function chkcond(v, rr) {
	if(rr == null) return false;
	var thresh = parseFloat(rr[2]);
	switch(rr[1]) {
		case "=":  if(v == thresh) return true; break;
		case ">":  if(v >  thresh) return true; break;
		case "<":  if(v <  thresh) return true; break;
		case "<>": if(v != thresh) return true; break;
		case ">=": if(v >= thresh) return true; break;
		case "<=": if(v <= thresh) return true; break;
	}
	return false;
}
function choose_fmt(f, v) {
	var fmt = split_fmt(f);
	var l = fmt.length, lat = fmt[l-1].indexOf("@");
	if(l<4 && lat>-1) --l;
	if(fmt.length > 4) throw new Error("cannot find right format for |" + fmt.join("|") + "|");
	if(typeof v !== "number") return [4, fmt.length === 4 || lat>-1?fmt[fmt.length-1]:"@"];
	switch(fmt.length) {
		case 1: fmt = lat>-1 ? ["General", "General", "General", fmt[0]] : [fmt[0], fmt[0], fmt[0], "@"]; break;
		case 2: fmt = lat>-1 ? [fmt[0], fmt[0], fmt[0], fmt[1]] : [fmt[0], fmt[1], fmt[0], "@"]; break;
		case 3: fmt = lat>-1 ? [fmt[0], fmt[1], fmt[0], fmt[2]] : [fmt[0], fmt[1], fmt[2], "@"]; break;
		case 4: break;
	}
	var ff = v > 0 ? fmt[0] : v < 0 ? fmt[1] : fmt[2];
	if(fmt[0].indexOf("[") === -1 && fmt[1].indexOf("[") === -1) return [l, ff];
	if(fmt[0].match(cfregex) != null || fmt[1].match(cfregex) != null) {
		var m1 = fmt[0].match(cfregex2);
		var m2 = fmt[1].match(cfregex2);
		return chkcond(v, m1) ? [l, fmt[0]] : chkcond(v, m2) ? [l, fmt[1]] : [l, fmt[m1 != null && m2 != null ? 2 : 1]];
	}
	return [l, ff];
}
function format(fmt,v,o) {
	if(o == null) o = {};
	var sfmt = "";
	switch(typeof fmt) {
		case "string":
			if(fmt == "m/d/yy" && o.dateNF) sfmt = o.dateNF;
			else sfmt = fmt;
			break;
		case "number":
			if(fmt == 14 && o.dateNF) sfmt = o.dateNF;
			else sfmt = (o.table != null ? (o.table) : table_fmt)[fmt];
			break;
	}
	if(isgeneral(sfmt,0)) return general_fmt(v, o);
	if(v instanceof Date) v = datenum_local(v, o.date1904);
	var f = choose_fmt(sfmt, v);
	if(isgeneral(f[1])) return general_fmt(v, o);
	if(v === true) v = "TRUE"; else if(v === false) v = "FALSE";
	else if(v === "" || v == null) return "";
	return eval_fmt(f[1], v, o, f[0]);
}
function load_entry(fmt, idx) {
	if(typeof idx != 'number') {
		idx = +idx || -1;
for(var i = 0; i < 0x0188; ++i) {
if(table_fmt[i] == undefined) { if(idx < 0) idx = i; continue; }
			if(table_fmt[i] == fmt) { idx = i; break; }
		}
if(idx < 0) idx = 0x187;
	}
table_fmt[idx] = fmt;
	return idx;
}
SSF.load = load_entry;
SSF._table = table_fmt;
SSF.get_table = function get_table() { return table_fmt; };
SSF.load_table = function load_table(tbl) {
	for(var i=0; i!=0x0188; ++i)
		if(tbl[i] !== undefined) load_entry(tbl[i], i);
};
SSF.init_table = init_table;
SSF.format = format;
};
make_ssf(SSF);
/* map from xlml named formats to SSF TODO: localize */
var XLMLFormatMap/*{[string]:string}*/ = ({
	"General Number": "General",
	"General Date": SSF._table[22],
	"Long Date": "dddd, mmmm dd, yyyy",
	"Medium Date": SSF._table[15],
	"Short Date": SSF._table[14],
	"Long Time": SSF._table[19],
	"Medium Time": SSF._table[18],
	"Short Time": SSF._table[20],
	"Currency": '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
	"Fixed": SSF._table[2],
	"Standard": SSF._table[4],
	"Percent": SSF._table[10],
	"Scientific": SSF._table[11],
	"Yes/No": '"Yes";"Yes";"No";@',
	"True/False": '"True";"True";"False";@',
	"On/Off": '"Yes";"Yes";"No";@'
});

var SSFImplicit/*{[number]:string}*/ = ({
	"5": '"$"#,##0_);\\("$"#,##0\\)',
	"6": '"$"#,##0_);[Red]\\("$"#,##0\\)',
	"7": '"$"#,##0.00_);\\("$"#,##0.00\\)',
	"8": '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
	"23": 'General', "24": 'General', "25": 'General', "26": 'General',
	"27": 'm/d/yy', "28": 'm/d/yy', "29": 'm/d/yy', "30": 'm/d/yy', "31": 'm/d/yy',
	"32": 'h:mm:ss', "33": 'h:mm:ss', "34": 'h:mm:ss', "35": 'h:mm:ss',
	"36": 'm/d/yy',
	"41": '_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)',
	"42": '_("$"* #,##0_);_("$"* \(#,##0\);_("$"* "-"_);_(@_)',
	"43": '_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)',
	"44": '_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)',
	"50": 'm/d/yy', "51": 'm/d/yy', "52": 'm/d/yy', "53": 'm/d/yy', "54": 'm/d/yy',
	"55": 'm/d/yy', "56": 'm/d/yy', "57": 'm/d/yy', "58": 'm/d/yy',
	"59": '0',
	"60": '0.00',
	"61": '#,##0',
	"62": '#,##0.00',
	"63": '"$"#,##0_);\\("$"#,##0\\)',
	"64": '"$"#,##0_);[Red]\\("$"#,##0\\)',
	"65": '"$"#,##0.00_);\\("$"#,##0.00\\)',
	"66": '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
	"67": '0%',
	"68": '0.00%',
	"69": '# ?/?',
	"70": '# ??/??',
	"71": 'm/d/yy',
	"72": 'm/d/yy',
	"73": 'd-mmm-yy',
	"74": 'd-mmm',
	"75": 'mmm-yy',
	"76": 'h:mm',
	"77": 'h:mm:ss',
	"78": 'm/d/yy h:mm',
	"79": 'mm:ss',
	"80": '[h]:mm:ss',
	"81": 'mmss.0'
});

/* dateNF parse TODO: move to SSF */
var dateNFregex = /[dD]+|[mM]+|[yYeE]+|[Hh]+|[Ss]+/g;
function dateNF_regex(dateNF) {
	var fmt = typeof dateNF == "number" ? SSF._table[dateNF] : dateNF;
	fmt = fmt.replace(dateNFregex, "(\\d+)");
	return new RegExp("^" + fmt + "$");
}
function dateNF_fix(str, dateNF, match) {
	var Y = -1, m = -1, d = -1, H = -1, M = -1, S = -1;
	(dateNF.match(dateNFregex)||[]).forEach(function(n, i) {
		var v = parseInt(match[i+1], 10);
		switch(n.toLowerCase().charAt(0)) {
			case 'y': Y = v; break; case 'd': d = v; break;
			case 'h': H = v; break; case 's': S = v; break;
			case 'm': if(H >= 0) M = v; else m = v; break;
		}
	});
	if(S >= 0 && M == -1 && m >= 0) { M = m; m = -1; }
	var datestr = (("" + (Y>=0?Y: new Date().getFullYear())).slice(-4) + "-" + ("00" + (m>=1?m:1)).slice(-2) + "-" + ("00" + (d>=1?d:1)).slice(-2));
	if(datestr.length == 7) datestr = "0" + datestr;
	if(datestr.length == 8) datestr = "20" + datestr;
	var timestr = (("00" + (H>=0?H:0)).slice(-2) + ":" + ("00" + (M>=0?M:0)).slice(-2) + ":" + ("00" + (S>=0?S:0)).slice(-2));
	if(H == -1 && M == -1 && S == -1) return datestr;
	if(Y == -1 && m == -1 && d == -1) return timestr;
	return datestr + "T" + timestr;
}

var DO_NOT_EXPORT_CFB = true;
/* cfb.js (C) 2013-present SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
/*jshint eqnull:true */
/*exported CFB */
/*global module, require:false, process:false, Buffer:false, Uint8Array:false, Uint16Array:false */

/* crc32.js (C) 2014-present SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
/*exported CRC32 */
var CRC32;
(function (factory) {
	/*jshint ignore:start */
	/*eslint-disable */
	factory(CRC32 = {});
	/*eslint-enable */
	/*jshint ignore:end */
}(function(CRC32) {
CRC32.version = '1.2.0';
/* see perf/crc32table.js */
/*global Int32Array */
function signed_crc_table() {
	var c = 0, table = new Array(256);

	for(var n =0; n != 256; ++n){
		c = n;
		c = ((c&1) ? (-306674912 ^ (c >>> 1)) : (c >>> 1));
		c = ((c&1) ? (-306674912 ^ (c >>> 1)) : (c >>> 1));
		c = ((c&1) ? (-306674912 ^ (c >>> 1)) : (c >>> 1));
		c = ((c&1) ? (-306674912 ^ (c >>> 1)) : (c >>> 1));
		c = ((c&1) ? (-306674912 ^ (c >>> 1)) : (c >>> 1));
		c = ((c&1) ? (-306674912 ^ (c >>> 1)) : (c >>> 1));
		c = ((c&1) ? (-306674912 ^ (c >>> 1)) : (c >>> 1));
		c = ((c&1) ? (-306674912 ^ (c >>> 1)) : (c >>> 1));
		table[n] = c;
	}

	return typeof Int32Array !== 'undefined' ? new Int32Array(table) : table;
}

var T = signed_crc_table();
function crc32_bstr(bstr, seed) {
	var C = seed ^ -1, L = bstr.length - 1;
	for(var i = 0; i < L;) {
		C = (C>>>8) ^ T[(C^bstr.charCodeAt(i++))&0xFF];
		C = (C>>>8) ^ T[(C^bstr.charCodeAt(i++))&0xFF];
	}
	if(i === L) C = (C>>>8) ^ T[(C ^ bstr.charCodeAt(i))&0xFF];
	return C ^ -1;
}

function crc32_buf(buf, seed) {
	if(buf.length > 10000) return crc32_buf_8(buf, seed);
	var C = seed ^ -1, L = buf.length - 3;
	for(var i = 0; i < L;) {
		C = (C>>>8) ^ T[(C^buf[i++])&0xFF];
		C = (C>>>8) ^ T[(C^buf[i++])&0xFF];
		C = (C>>>8) ^ T[(C^buf[i++])&0xFF];
		C = (C>>>8) ^ T[(C^buf[i++])&0xFF];
	}
	while(i < L+3) C = (C>>>8) ^ T[(C^buf[i++])&0xFF];
	return C ^ -1;
}

function crc32_buf_8(buf, seed) {
	var C = seed ^ -1, L = buf.length - 7;
	for(var i = 0; i < L;) {
		C = (C>>>8) ^ T[(C^buf[i++])&0xFF];
		C = (C>>>8) ^ T[(C^buf[i++])&0xFF];
		C = (C>>>8) ^ T[(C^buf[i++])&0xFF];
		C = (C>>>8) ^ T[(C^buf[i++])&0xFF];
		C = (C>>>8) ^ T[(C^buf[i++])&0xFF];
		C = (C>>>8) ^ T[(C^buf[i++])&0xFF];
		C = (C>>>8) ^ T[(C^buf[i++])&0xFF];
		C = (C>>>8) ^ T[(C^buf[i++])&0xFF];
	}
	while(i < L+7) C = (C>>>8) ^ T[(C^buf[i++])&0xFF];
	return C ^ -1;
}

function crc32_str(str, seed) {
	var C = seed ^ -1;
	for(var i = 0, L=str.length, c, d; i < L;) {
		c = str.charCodeAt(i++);
		if(c < 0x80) {
			C = (C>>>8) ^ T[(C ^ c)&0xFF];
		} else if(c < 0x800) {
			C = (C>>>8) ^ T[(C ^ (192|((c>>6)&31)))&0xFF];
			C = (C>>>8) ^ T[(C ^ (128|(c&63)))&0xFF];
		} else if(c >= 0xD800 && c < 0xE000) {
			c = (c&1023)+64; d = str.charCodeAt(i++)&1023;
			C = (C>>>8) ^ T[(C ^ (240|((c>>8)&7)))&0xFF];
			C = (C>>>8) ^ T[(C ^ (128|((c>>2)&63)))&0xFF];
			C = (C>>>8) ^ T[(C ^ (128|((d>>6)&15)|((c&3)<<4)))&0xFF];
			C = (C>>>8) ^ T[(C ^ (128|(d&63)))&0xFF];
		} else {
			C = (C>>>8) ^ T[(C ^ (224|((c>>12)&15)))&0xFF];
			C = (C>>>8) ^ T[(C ^ (128|((c>>6)&63)))&0xFF];
			C = (C>>>8) ^ T[(C ^ (128|(c&63)))&0xFF];
		}
	}
	return C ^ -1;
}
CRC32.table = T;
CRC32.bstr = crc32_bstr;
CRC32.buf = crc32_buf;
CRC32.str = crc32_str;
}));
/* [MS-CFB] v20171201 */
var CFB = (function _CFB(){
var exports = {};
exports.version = '1.1.0';
/* [MS-CFB] 2.6.4 */
function namecmp(l, r) {
	var L = l.split("/"), R = r.split("/");
	for(var i = 0, c = 0, Z = Math.min(L.length, R.length); i < Z; ++i) {
		if((c = L[i].length - R[i].length)) return c;
		if(L[i] != R[i]) return L[i] < R[i] ? -1 : 1;
	}
	return L.length - R.length;
}
function dirname(p) {
	if(p.charAt(p.length - 1) == "/") return (p.slice(0,-1).indexOf("/") === -1) ? p : dirname(p.slice(0, -1));
	var c = p.lastIndexOf("/");
	return (c === -1) ? p : p.slice(0, c+1);
}

function filename(p) {
	if(p.charAt(p.length - 1) == "/") return filename(p.slice(0, -1));
	var c = p.lastIndexOf("/");
	return (c === -1) ? p : p.slice(c+1);
}
/* -------------------------------------------------------------------------- */
/* DOS Date format:
   high|YYYYYYYm.mmmddddd.HHHHHMMM.MMMSSSSS|low
   add 1980 to stored year
   stored second should be doubled
*/

/* write JS date to buf as a DOS date */
function write_dos_date(buf, date) {
	if(typeof date === "string") date = new Date(date);
	var hms = date.getHours();
	hms = hms << 6 | date.getMinutes();
	hms = hms << 5 | (date.getSeconds()>>>1);
	buf.write_shift(2, hms);
	var ymd = (date.getFullYear() - 1980);
	ymd = ymd << 4 | (date.getMonth()+1);
	ymd = ymd << 5 | date.getDate();
	buf.write_shift(2, ymd);
}

/* read four bytes from buf and interpret as a DOS date */
function parse_dos_date(buf) {
	var hms = buf.read_shift(2) & 0xFFFF;
	var ymd = buf.read_shift(2) & 0xFFFF;
	var val = new Date();
	var d = ymd & 0x1F; ymd >>>= 5;
	var m = ymd & 0x0F; ymd >>>= 4;
	val.setMilliseconds(0);
	val.setFullYear(ymd + 1980);
	val.setMonth(m-1);
	val.setDate(d);
	var S = hms & 0x1F; hms >>>= 5;
	var M = hms & 0x3F; hms >>>= 6;
	val.setHours(hms);
	val.setMinutes(M);
	val.setSeconds(S<<1);
	return val;
}
function parse_extra_field(blob) {
	prep_blob(blob, 0);
	var o = {};
	var flags = 0;
	while(blob.l <= blob.length - 4) {
		var type = blob.read_shift(2);
		var sz = blob.read_shift(2), tgt = blob.l + sz;
		var p = {};
		switch(type) {
			/* UNIX-style Timestamps */
			case 0x5455: {
				flags = blob.read_shift(1);
				if(flags & 1) p.mtime = blob.read_shift(4);
				/* for some reason, CD flag corresponds to LFH */
				if(sz > 5) {
					if(flags & 2) p.atime = blob.read_shift(4);
					if(flags & 4) p.ctime = blob.read_shift(4);
				}
				if(p.mtime) p.mt = new Date(p.mtime*1000);
			}
			break;
		}
		blob.l = tgt;
		o[type] = p;
	}
	return o;
}
var fs;
function get_fs() { return fs || (fs = require('fs')); }
function parse(file, options) {
if(file[0] == 0x50 && file[1] == 0x4b) return parse_zip(file, options);
if(file.length < 512) throw new Error("CFB file size " + file.length + " < 512");
var mver = 3;
var ssz = 512;
var nmfs = 0; // number of mini FAT sectors
var difat_sec_cnt = 0;
var dir_start = 0;
var minifat_start = 0;
var difat_start = 0;

var fat_addrs = []; // locations of FAT sectors

/* [MS-CFB] 2.2 Compound File Header */
var blob = file.slice(0,512);
prep_blob(blob, 0);

/* major version */
var mv = check_get_mver(blob);
mver = mv[0];
switch(mver) {
	case 3: ssz = 512; break; case 4: ssz = 4096; break;
	case 0: if(mv[1] == 0) return parse_zip(file, options);
	/* falls through */
	default: throw new Error("Major Version: Expected 3 or 4 saw " + mver);
}

/* reprocess header */
if(ssz !== 512) { blob = file.slice(0,ssz); prep_blob(blob, 28 /* blob.l */); }
/* Save header for final object */
var header = file.slice(0,ssz);

check_shifts(blob, mver);

// Number of Directory Sectors
var dir_cnt = blob.read_shift(4, 'i');
if(mver === 3 && dir_cnt !== 0) throw new Error('# Directory Sectors: Expected 0 saw ' + dir_cnt);

// Number of FAT Sectors
blob.l += 4;

// First Directory Sector Location
dir_start = blob.read_shift(4, 'i');

// Transaction Signature
blob.l += 4;

// Mini Stream Cutoff Size
blob.chk('00100000', 'Mini Stream Cutoff Size: ');

// First Mini FAT Sector Location
minifat_start = blob.read_shift(4, 'i');

// Number of Mini FAT Sectors
nmfs = blob.read_shift(4, 'i');

// First DIFAT sector location
difat_start = blob.read_shift(4, 'i');

// Number of DIFAT Sectors
difat_sec_cnt = blob.read_shift(4, 'i');

// Grab FAT Sector Locations
for(var q = -1, j = 0; j < 109; ++j) { /* 109 = (512 - blob.l)>>>2; */
	q = blob.read_shift(4, 'i');
	if(q<0) break;
	fat_addrs[j] = q;
}

/** Break the file up into sectors */
var sectors = sectorify(file, ssz);

sleuth_fat(difat_start, difat_sec_cnt, sectors, ssz, fat_addrs);

/** Chains */
var sector_list = make_sector_list(sectors, dir_start, fat_addrs, ssz);

sector_list[dir_start].name = "!Directory";
if(nmfs > 0 && minifat_start !== ENDOFCHAIN) sector_list[minifat_start].name = "!MiniFAT";
sector_list[fat_addrs[0]].name = "!FAT";
sector_list.fat_addrs = fat_addrs;
sector_list.ssz = ssz;

/* [MS-CFB] 2.6.1 Compound File Directory Entry */
var files = {}, Paths = [], FileIndex = [], FullPaths = [];
read_directory(dir_start, sector_list, sectors, Paths, nmfs, files, FileIndex, minifat_start);

build_full_paths(FileIndex, FullPaths, Paths);
Paths.shift();

var o = {
	FileIndex: FileIndex,
	FullPaths: FullPaths
};

// $FlowIgnore
if(options && options.raw) o.raw = {header: header, sectors: sectors};
return o;
} // parse

/* [MS-CFB] 2.2 Compound File Header -- read up to major version */
function check_get_mver(blob) {
	if(blob[blob.l] == 0x50 && blob[blob.l + 1] == 0x4b) return [0, 0];
	// header signature 8
	blob.chk(HEADER_SIGNATURE, 'Header Signature: ');

	// clsid 16
	blob.chk(HEADER_CLSID, 'CLSID: ');

	// minor version 2
	var mver = blob.read_shift(2, 'u');

	return [blob.read_shift(2,'u'), mver];
}
function check_shifts(blob, mver) {
	var shift = 0x09;

	// Byte Order
	//blob.chk('feff', 'Byte Order: '); // note: some writers put 0xffff
	blob.l += 2;

	// Sector Shift
	switch((shift = blob.read_shift(2))) {
		case 0x09: if(mver != 3) throw new Error('Sector Shift: Expected 9 saw ' + shift); break;
		case 0x0c: if(mver != 4) throw new Error('Sector Shift: Expected 12 saw ' + shift); break;
		default: throw new Error('Sector Shift: Expected 9 or 12 saw ' + shift);
	}

	// Mini Sector Shift
	blob.chk('0600', 'Mini Sector Shift: ');

	// Reserved
	blob.chk('000000000000', 'Reserved: ');
}

/** Break the file up into sectors */
function sectorify(file, ssz) {
	var nsectors = Math.ceil(file.length/ssz)-1;
	var sectors = [];
	for(var i=1; i < nsectors; ++i) sectors[i-1] = file.slice(i*ssz,(i+1)*ssz);
	sectors[nsectors-1] = file.slice(nsectors*ssz);
	return sectors;
}

/* [MS-CFB] 2.6.4 Red-Black Tree */
function build_full_paths(FI, FP, Paths) {
	var i = 0, L = 0, R = 0, C = 0, j = 0, pl = Paths.length;
	var dad = [], q = [];

	for(; i < pl; ++i) { dad[i]=q[i]=i; FP[i]=Paths[i]; }

	for(; j < q.length; ++j) {
		i = q[j];
		L = FI[i].L; R = FI[i].R; C = FI[i].C;
		if(dad[i] === i) {
			if(L !== -1 /*NOSTREAM*/ && dad[L] !== L) dad[i] = dad[L];
			if(R !== -1 && dad[R] !== R) dad[i] = dad[R];
		}
		if(C !== -1 /*NOSTREAM*/) dad[C] = i;
		if(L !== -1) { dad[L] = dad[i]; if(q.lastIndexOf(L) < j) q.push(L); }
		if(R !== -1) { dad[R] = dad[i]; if(q.lastIndexOf(R) < j) q.push(R); }
	}
	for(i=1; i < pl; ++i) if(dad[i] === i) {
		if(R !== -1 /*NOSTREAM*/ && dad[R] !== R) dad[i] = dad[R];
		else if(L !== -1 && dad[L] !== L) dad[i] = dad[L];
	}

	for(i=1; i < pl; ++i) {
		if(FI[i].type === 0 /* unknown */) continue;
		j = dad[i];
		if(j === 0) FP[i] = FP[0] + "/" + FP[i];
		else while(j !== 0 && j !== dad[j]) {
			FP[i] = FP[j] + "/" + FP[i];
			j = dad[j];
		}
		dad[i] = 0;
	}

	FP[0] += "/";
	for(i=1; i < pl; ++i) {
		if(FI[i].type !== 2 /* stream */) FP[i] += "/";
	}
}

function get_mfat_entry(entry, payload, mini) {
	var start = entry.start, size = entry.size;
	//return (payload.slice(start*MSSZ, start*MSSZ + size));
	var o = [];
	var idx = start;
	while(mini && size > 0 && idx >= 0) {
		o.push(payload.slice(idx * MSSZ, idx * MSSZ + MSSZ));
		size -= MSSZ;
		idx = __readInt32LE(mini, idx * 4);
	}
	if(o.length === 0) return (new_buf(0));
	return (bconcat(o).slice(0, entry.size));
}

/** Chase down the rest of the DIFAT chain to build a comprehensive list
    DIFAT chains by storing the next sector number as the last 32 bits */
function sleuth_fat(idx, cnt, sectors, ssz, fat_addrs) {
	var q = ENDOFCHAIN;
	if(idx === ENDOFCHAIN) {
		if(cnt !== 0) throw new Error("DIFAT chain shorter than expected");
	} else if(idx !== -1 /*FREESECT*/) {
		var sector = sectors[idx], m = (ssz>>>2)-1;
		if(!sector) return;
		for(var i = 0; i < m; ++i) {
			if((q = __readInt32LE(sector,i*4)) === ENDOFCHAIN) break;
			fat_addrs.push(q);
		}
		sleuth_fat(__readInt32LE(sector,ssz-4),cnt - 1, sectors, ssz, fat_addrs);
	}
}

/** Follow the linked list of sectors for a given starting point */
function get_sector_list(sectors, start, fat_addrs, ssz, chkd) {
	var buf = [], buf_chain = [];
	if(!chkd) chkd = [];
	var modulus = ssz - 1, j = 0, jj = 0;
	for(j=start; j>=0;) {
		chkd[j] = true;
		buf[buf.length] = j;
		buf_chain.push(sectors[j]);
		var addr = fat_addrs[Math.floor(j*4/ssz)];
		jj = ((j*4) & modulus);
		if(ssz < 4 + jj) throw new Error("FAT boundary crossed: " + j + " 4 "+ssz);
		if(!sectors[addr]) break;
		j = __readInt32LE(sectors[addr], jj);
	}
	return {nodes: buf, data:__toBuffer([buf_chain])};
}

/** Chase down the sector linked lists */
function make_sector_list(sectors, dir_start, fat_addrs, ssz) {
	var sl = sectors.length, sector_list = ([]);
	var chkd = [], buf = [], buf_chain = [];
	var modulus = ssz - 1, i=0, j=0, k=0, jj=0;
	for(i=0; i < sl; ++i) {
		buf = ([]);
		k = (i + dir_start); if(k >= sl) k-=sl;
		if(chkd[k]) continue;
		buf_chain = [];
		for(j=k; j>=0;) {
			chkd[j] = true;
			buf[buf.length] = j;
			buf_chain.push(sectors[j]);
			var addr = fat_addrs[Math.floor(j*4/ssz)];
			jj = ((j*4) & modulus);
			if(ssz < 4 + jj) throw new Error("FAT boundary crossed: " + j + " 4 "+ssz);
			if(!sectors[addr]) break;
			j = __readInt32LE(sectors[addr], jj);
		}
		sector_list[k] = ({nodes: buf, data:__toBuffer([buf_chain])});
	}
	return sector_list;
}

/* [MS-CFB] 2.6.1 Compound File Directory Entry */
function read_directory(dir_start, sector_list, sectors, Paths, nmfs, files, FileIndex, mini) {
	var minifat_store = 0, pl = (Paths.length?2:0);
	var sector = sector_list[dir_start].data;
	var i = 0, namelen = 0, name;
	for(; i < sector.length; i+= 128) {
		var blob = sector.slice(i, i+128);
		prep_blob(blob, 64);
		namelen = blob.read_shift(2);
		name = __utf16le(blob,0,namelen-pl);
		Paths.push(name);
		var o = ({
			name:  name,
			type:  blob.read_shift(1),
			color: blob.read_shift(1),
			L:     blob.read_shift(4, 'i'),
			R:     blob.read_shift(4, 'i'),
			C:     blob.read_shift(4, 'i'),
			clsid: blob.read_shift(16),
			state: blob.read_shift(4, 'i'),
			start: 0,
			size: 0
		});
		var ctime = blob.read_shift(2) + blob.read_shift(2) + blob.read_shift(2) + blob.read_shift(2);
		if(ctime !== 0) o.ct = read_date(blob, blob.l-8);
		var mtime = blob.read_shift(2) + blob.read_shift(2) + blob.read_shift(2) + blob.read_shift(2);
		if(mtime !== 0) o.mt = read_date(blob, blob.l-8);
		o.start = blob.read_shift(4, 'i');
		o.size = blob.read_shift(4, 'i');
		if(o.size < 0 && o.start < 0) { o.size = o.type = 0; o.start = ENDOFCHAIN; o.name = ""; }
		if(o.type === 5) { /* root */
			minifat_store = o.start;
			if(nmfs > 0 && minifat_store !== ENDOFCHAIN) sector_list[minifat_store].name = "!StreamData";
			/*minifat_size = o.size;*/
		} else if(o.size >= 4096 /* MSCSZ */) {
			o.storage = 'fat';
			if(sector_list[o.start] === undefined) sector_list[o.start] = get_sector_list(sectors, o.start, sector_list.fat_addrs, sector_list.ssz);
			sector_list[o.start].name = o.name;
			o.content = (sector_list[o.start].data.slice(0,o.size));
		} else {
			o.storage = 'minifat';
			if(o.size < 0) o.size = 0;
			else if(minifat_store !== ENDOFCHAIN && o.start !== ENDOFCHAIN && sector_list[minifat_store]) {
				o.content = get_mfat_entry(o, sector_list[minifat_store].data, (sector_list[mini]||{}).data);
			}
		}
		if(o.content) prep_blob(o.content, 0);
		files[name] = o;
		FileIndex.push(o);
	}
}

function read_date(blob, offset) {
	return new Date(( ( (__readUInt32LE(blob,offset+4)/1e7)*Math.pow(2,32)+__readUInt32LE(blob,offset)/1e7 ) - 11644473600)*1000);
}

function read_file(filename, options) {
	get_fs();
	return parse(fs.readFileSync(filename), options);
}

function read(blob, options) {
	switch(options && options.type || "base64") {
		case "file": return read_file(blob, options);
		case "base64": return parse(s2a(Base64.decode(blob)), options);
		case "binary": return parse(s2a(blob), options);
	}
	return parse(blob, options);
}

function init_cfb(cfb, opts) {
	var o = opts || {}, root = o.root || "Root Entry";
	if(!cfb.FullPaths) cfb.FullPaths = [];
	if(!cfb.FileIndex) cfb.FileIndex = [];
	if(cfb.FullPaths.length !== cfb.FileIndex.length) throw new Error("inconsistent CFB structure");
	if(cfb.FullPaths.length === 0) {
		cfb.FullPaths[0] = root + "/";
		cfb.FileIndex[0] = ({ name: root, type: 5 });
	}
	if(o.CLSID) cfb.FileIndex[0].clsid = o.CLSID;
	seed_cfb(cfb);
}
function seed_cfb(cfb) {
	var nm = "\u0001Sh33tJ5";
	if(CFB.find(cfb, "/" + nm)) return;
	var p = new_buf(4); p[0] = 55; p[1] = p[3] = 50; p[2] = 54;
	cfb.FileIndex.push(({ name: nm, type: 2, content:p, size:4, L:69, R:69, C:69 }));
	cfb.FullPaths.push(cfb.FullPaths[0] + nm);
	rebuild_cfb(cfb);
}
function rebuild_cfb(cfb, f) {
	init_cfb(cfb);
	var gc = false, s = false;
	for(var i = cfb.FullPaths.length - 1; i >= 0; --i) {
		var _file = cfb.FileIndex[i];
		switch(_file.type) {
			case 0:
				if(s) gc = true;
				else { cfb.FileIndex.pop(); cfb.FullPaths.pop(); }
				break;
			case 1: case 2: case 5:
				s = true;
				if(isNaN(_file.R * _file.L * _file.C)) gc = true;
				if(_file.R > -1 && _file.L > -1 && _file.R == _file.L) gc = true;
				break;
			default: gc = true; break;
		}
	}
	if(!gc && !f) return;

	var now = new Date(1987, 1, 19), j = 0;
	var data = [];
	for(i = 0; i < cfb.FullPaths.length; ++i) {
		if(cfb.FileIndex[i].type === 0) continue;
		data.push([cfb.FullPaths[i], cfb.FileIndex[i]]);
	}
	for(i = 0; i < data.length; ++i) {
		var dad = dirname(data[i][0]);
		s = false;
		for(j = 0; j < data.length; ++j) if(data[j][0] === dad) s = true;
		if(!s) data.push([dad, ({
			name: filename(dad).replace("/",""),
			type: 1,
			clsid: HEADER_CLSID,
			ct: now, mt: now,
			content: null
		})]);
	}

	data.sort(function(x,y) { return namecmp(x[0], y[0]); });
	cfb.FullPaths = []; cfb.FileIndex = [];
	for(i = 0; i < data.length; ++i) { cfb.FullPaths[i] = data[i][0]; cfb.FileIndex[i] = data[i][1]; }
	for(i = 0; i < data.length; ++i) {
		var elt = cfb.FileIndex[i];
		var nm = cfb.FullPaths[i];

		elt.name =  filename(nm).replace("/","");
		elt.L = elt.R = elt.C = -(elt.color = 1);
		elt.size = elt.content ? elt.content.length : 0;
		elt.start = 0;
		elt.clsid = (elt.clsid || HEADER_CLSID);
		if(i === 0) {
			elt.C = data.length > 1 ? 1 : -1;
			elt.size = 0;
			elt.type = 5;
		} else if(nm.slice(-1) == "/") {
			for(j=i+1;j < data.length; ++j) if(dirname(cfb.FullPaths[j])==nm) break;
			elt.C = j >= data.length ? -1 : j;
			for(j=i+1;j < data.length; ++j) if(dirname(cfb.FullPaths[j])==dirname(nm)) break;
			elt.R = j >= data.length ? -1 : j;
			elt.type = 1;
		} else {
			if(dirname(cfb.FullPaths[i+1]||"") == dirname(nm)) elt.R = i + 1;
			elt.type = 2;
		}
	}

}

function _write(cfb, options) {
	var _opts = options || {};
	rebuild_cfb(cfb);
	if(_opts.fileType == 'zip') return write_zip(cfb, _opts);
	var L = (function(cfb){
		var mini_size = 0, fat_size = 0;
		for(var i = 0; i < cfb.FileIndex.length; ++i) {
			var file = cfb.FileIndex[i];
			if(!file.content) continue;
var flen = file.content.length;
			if(flen > 0){
				if(flen < 0x1000) mini_size += (flen + 0x3F) >> 6;
				else fat_size += (flen + 0x01FF) >> 9;
			}
		}
		var dir_cnt = (cfb.FullPaths.length +3) >> 2;
		var mini_cnt = (mini_size + 7) >> 3;
		var mfat_cnt = (mini_size + 0x7F) >> 7;
		var fat_base = mini_cnt + fat_size + dir_cnt + mfat_cnt;
		var fat_cnt = (fat_base + 0x7F) >> 7;
		var difat_cnt = fat_cnt <= 109 ? 0 : Math.ceil((fat_cnt-109)/0x7F);
		while(((fat_base + fat_cnt + difat_cnt + 0x7F) >> 7) > fat_cnt) difat_cnt = ++fat_cnt <= 109 ? 0 : Math.ceil((fat_cnt-109)/0x7F);
		var L =  [1, difat_cnt, fat_cnt, mfat_cnt, dir_cnt, fat_size, mini_size, 0];
		cfb.FileIndex[0].size = mini_size << 6;
		L[7] = (cfb.FileIndex[0].start=L[0]+L[1]+L[2]+L[3]+L[4]+L[5])+((L[6]+7) >> 3);
		return L;
	})(cfb);
	var o = new_buf(L[7] << 9);
	var i = 0, T = 0;
	{
		for(i = 0; i < 8; ++i) o.write_shift(1, HEADER_SIG[i]);
		for(i = 0; i < 8; ++i) o.write_shift(2, 0);
		o.write_shift(2, 0x003E);
		o.write_shift(2, 0x0003);
		o.write_shift(2, 0xFFFE);
		o.write_shift(2, 0x0009);
		o.write_shift(2, 0x0006);
		for(i = 0; i < 3; ++i) o.write_shift(2, 0);
		o.write_shift(4, 0);
		o.write_shift(4, L[2]);
		o.write_shift(4, L[0] + L[1] + L[2] + L[3] - 1);
		o.write_shift(4, 0);
		o.write_shift(4, 1<<12);
		o.write_shift(4, L[3] ? L[0] + L[1] + L[2] - 1: ENDOFCHAIN);
		o.write_shift(4, L[3]);
		o.write_shift(-4, L[1] ? L[0] - 1: ENDOFCHAIN);
		o.write_shift(4, L[1]);
		for(i = 0; i < 109; ++i) o.write_shift(-4, i < L[2] ? L[1] + i : -1);
	}
	if(L[1]) {
		for(T = 0; T < L[1]; ++T) {
			for(; i < 236 + T * 127; ++i) o.write_shift(-4, i < L[2] ? L[1] + i : -1);
			o.write_shift(-4, T === L[1] - 1 ? ENDOFCHAIN : T + 1);
		}
	}
	var chainit = function(w) {
		for(T += w; i<T-1; ++i) o.write_shift(-4, i+1);
		if(w) { ++i; o.write_shift(-4, ENDOFCHAIN); }
	};
	T = i = 0;
	for(T+=L[1]; i<T; ++i) o.write_shift(-4, consts.DIFSECT);
	for(T+=L[2]; i<T; ++i) o.write_shift(-4, consts.FATSECT);
	chainit(L[3]);
	chainit(L[4]);
	var j = 0, flen = 0;
	var file = cfb.FileIndex[0];
	for(; j < cfb.FileIndex.length; ++j) {
		file = cfb.FileIndex[j];
		if(!file.content) continue;
flen = file.content.length;
		if(flen < 0x1000) continue;
		file.start = T;
		chainit((flen + 0x01FF) >> 9);
	}
	chainit((L[6] + 7) >> 3);
	while(o.l & 0x1FF) o.write_shift(-4, consts.ENDOFCHAIN);
	T = i = 0;
	for(j = 0; j < cfb.FileIndex.length; ++j) {
		file = cfb.FileIndex[j];
		if(!file.content) continue;
flen = file.content.length;
		if(!flen || flen >= 0x1000) continue;
		file.start = T;
		chainit((flen + 0x3F) >> 6);
	}
	while(o.l & 0x1FF) o.write_shift(-4, consts.ENDOFCHAIN);
	for(i = 0; i < L[4]<<2; ++i) {
		var nm = cfb.FullPaths[i];
		if(!nm || nm.length === 0) {
			for(j = 0; j < 17; ++j) o.write_shift(4, 0);
			for(j = 0; j < 3; ++j) o.write_shift(4, -1);
			for(j = 0; j < 12; ++j) o.write_shift(4, 0);
			continue;
		}
		file = cfb.FileIndex[i];
		if(i === 0) file.start = file.size ? file.start - 1 : ENDOFCHAIN;
		var _nm = (i === 0 && _opts.root) || file.name;
		flen = 2*(_nm.length+1);
		o.write_shift(64, _nm, "utf16le");
		o.write_shift(2, flen);
		o.write_shift(1, file.type);
		o.write_shift(1, file.color);
		o.write_shift(-4, file.L);
		o.write_shift(-4, file.R);
		o.write_shift(-4, file.C);
		if(!file.clsid) for(j = 0; j < 4; ++j) o.write_shift(4, 0);
		else o.write_shift(16, file.clsid, "hex");
		o.write_shift(4, file.state || 0);
		o.write_shift(4, 0); o.write_shift(4, 0);
		o.write_shift(4, 0); o.write_shift(4, 0);
		o.write_shift(4, file.start);
		o.write_shift(4, file.size); o.write_shift(4, 0);
	}
	for(i = 1; i < cfb.FileIndex.length; ++i) {
		file = cfb.FileIndex[i];
if(file.size >= 0x1000) {
			o.l = (file.start+1) << 9;
			for(j = 0; j < file.size; ++j) o.write_shift(1, file.content[j]);
			for(; j & 0x1FF; ++j) o.write_shift(1, 0);
		}
	}
	for(i = 1; i < cfb.FileIndex.length; ++i) {
		file = cfb.FileIndex[i];
if(file.size > 0 && file.size < 0x1000) {
			for(j = 0; j < file.size; ++j) o.write_shift(1, file.content[j]);
			for(; j & 0x3F; ++j) o.write_shift(1, 0);
		}
	}
	while(o.l < o.length) o.write_shift(1, 0);
	return o;
}
/* [MS-CFB] 2.6.4 (Unicode 3.0.1 case conversion) */
function find(cfb, path) {
	var UCFullPaths = cfb.FullPaths.map(function(x) { return x.toUpperCase(); });
	var UCPaths = UCFullPaths.map(function(x) { var y = x.split("/"); return y[y.length - (x.slice(-1) == "/" ? 2 : 1)]; });
	var k = false;
	if(path.charCodeAt(0) === 47 /* "/" */) { k = true; path = UCFullPaths[0].slice(0, -1) + path; }
	else k = path.indexOf("/") !== -1;
	var UCPath = path.toUpperCase();
	var w = k === true ? UCFullPaths.indexOf(UCPath) : UCPaths.indexOf(UCPath);
	if(w !== -1) return cfb.FileIndex[w];

	var m = !UCPath.match(chr1);
	UCPath = UCPath.replace(chr0,'');
	if(m) UCPath = UCPath.replace(chr1,'!');
	for(w = 0; w < UCFullPaths.length; ++w) {
		if((m ? UCFullPaths[w].replace(chr1,'!') : UCFullPaths[w]).replace(chr0,'') == UCPath) return cfb.FileIndex[w];
		if((m ? UCPaths[w].replace(chr1,'!') : UCPaths[w]).replace(chr0,'') == UCPath) return cfb.FileIndex[w];
	}
	return null;
}
/** CFB Constants */
var MSSZ = 64; /* Mini Sector Size = 1<<6 */
//var MSCSZ = 4096; /* Mini Stream Cutoff Size */
/* 2.1 Compound File Sector Numbers and Types */
var ENDOFCHAIN = -2;
/* 2.2 Compound File Header */
var HEADER_SIGNATURE = 'd0cf11e0a1b11ae1';
var HEADER_SIG = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
var HEADER_CLSID = '00000000000000000000000000000000';
var consts = {
	/* 2.1 Compund File Sector Numbers and Types */
	MAXREGSECT: -6,
	DIFSECT: -4,
	FATSECT: -3,
	ENDOFCHAIN: ENDOFCHAIN,
	FREESECT: -1,
	/* 2.2 Compound File Header */
	HEADER_SIGNATURE: HEADER_SIGNATURE,
	HEADER_MINOR_VERSION: '3e00',
	MAXREGSID: -6,
	NOSTREAM: -1,
	HEADER_CLSID: HEADER_CLSID,
	/* 2.6.1 Compound File Directory Entry */
	EntryTypes: ['unknown','storage','stream','lockbytes','property','root']
};

function write_file(cfb, filename, options) {
	get_fs();
	var o = _write(cfb, options);
fs.writeFileSync(filename, o);
}

function a2s(o) {
	var out = new Array(o.length);
	for(var i = 0; i < o.length; ++i) out[i] = String.fromCharCode(o[i]);
	return out.join("");
}

function write(cfb, options) {
	var o = _write(cfb, options);
	switch(options && options.type) {
		case "file": get_fs(); fs.writeFileSync(options.filename, (o)); return o;
		case "binary": return a2s(o);
		case "base64": return Base64.encode(a2s(o));
	}
	return o;
}
/* node < 8.1 zlib does not expose bytesRead, so default to pure JS */
var _zlib;
function use_zlib(zlib) { try {
	var InflateRaw = zlib.InflateRaw;
	var InflRaw = new InflateRaw();
	InflRaw._processChunk(new Uint8Array([3, 0]), InflRaw._finishFlushFlag);
	if(InflRaw.bytesRead) _zlib = zlib;
	else throw new Error("zlib does not expose bytesRead");
} catch(e) {console.error("cannot use native zlib: " + (e.message || e)); } }

function _inflateRawSync(payload, usz) {
	if(!_zlib) return _inflate(payload, usz);
	var InflateRaw = _zlib.InflateRaw;
	var InflRaw = new InflateRaw();
	var out = InflRaw._processChunk(payload.slice(payload.l), InflRaw._finishFlushFlag);
	payload.l += InflRaw.bytesRead;
	return out;
}

function _deflateRawSync(payload) {
	return _zlib ? _zlib.deflateRawSync(payload) : _deflate(payload);
}
var CLEN_ORDER = [ 16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15 ];

/*  LEN_ID = [ 257, 258, 259, 260, 261, 262, 263, 264, 265, 266, 267, 268, 269, 270, 271, 272, 273, 274, 275, 276, 277, 278, 279, 280, 281, 282, 283, 284, 285 ]; */
var LEN_LN = [   3,   4,   5,   6,   7,   8,   9,  10,  11,  13 , 15,  17,  19,  23,  27,  31,  35,  43,  51,  59,  67,  83,  99, 115, 131, 163, 195, 227, 258 ];

/*  DST_ID = [  0,  1,  2,  3,  4,  5,  6,  7,  8,  9, 10, 11, 12, 13,  14,  15,  16,  17,  18,  19,   20,   21,   22,   23,   24,   25,   26,    27,    28,    29 ]; */
var DST_LN = [  1,  2,  3,  4,  5,  7,  9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577 ];

function bit_swap_8(n) { var t = (((((n<<1)|(n<<11)) & 0x22110) | (((n<<5)|(n<<15)) & 0x88440))); return ((t>>16) | (t>>8) |t)&0xFF; }

var use_typed_arrays = typeof Uint8Array !== 'undefined';

var bitswap8 = use_typed_arrays ? new Uint8Array(1<<8) : [];
for(var q = 0; q < (1<<8); ++q) bitswap8[q] = bit_swap_8(q);

function bit_swap_n(n, b) {
	var rev = bitswap8[n & 0xFF];
	if(b <= 8) return rev >>> (8-b);
	rev = (rev << 8) | bitswap8[(n>>8)&0xFF];
	if(b <= 16) return rev >>> (16-b);
	rev = (rev << 8) | bitswap8[(n>>16)&0xFF];
	return rev >>> (24-b);
}

/* helpers for unaligned bit reads */
function read_bits_2(buf, bl) { var w = (bl&7), h = (bl>>>3); return ((buf[h]|(w <= 6 ? 0 : buf[h+1]<<8))>>>w)& 0x03; }
function read_bits_3(buf, bl) { var w = (bl&7), h = (bl>>>3); return ((buf[h]|(w <= 5 ? 0 : buf[h+1]<<8))>>>w)& 0x07; }
function read_bits_4(buf, bl) { var w = (bl&7), h = (bl>>>3); return ((buf[h]|(w <= 4 ? 0 : buf[h+1]<<8))>>>w)& 0x0F; }
function read_bits_5(buf, bl) { var w = (bl&7), h = (bl>>>3); return ((buf[h]|(w <= 3 ? 0 : buf[h+1]<<8))>>>w)& 0x1F; }
function read_bits_7(buf, bl) { var w = (bl&7), h = (bl>>>3); return ((buf[h]|(w <= 1 ? 0 : buf[h+1]<<8))>>>w)& 0x7F; }

/* works up to n = 3 * 8 + 1 = 25 */
function read_bits_n(buf, bl, n) {
	var w = (bl&7), h = (bl>>>3), f = ((1<<n)-1);
	var v = buf[h] >>> w;
	if(n < 8 - w) return v & f;
	v |= buf[h+1]<<(8-w);
	if(n < 16 - w) return v & f;
	v |= buf[h+2]<<(16-w);
	if(n < 24 - w) return v & f;
	v |= buf[h+3]<<(24-w);
	return v & f;
}

/* until ArrayBuffer#realloc is a thing, fake a realloc */
function realloc(b, sz) {
	var L = b.length, M = 2*L > sz ? 2*L : sz + 5, i = 0;
	if(L >= sz) return b;
	if(has_buf) {
		var o = new_unsafe_buf(M);
		// $FlowIgnore
		if(b.copy) b.copy(o);
		else for(; i < b.length; ++i) o[i] = b[i];
		return o;
	} else if(use_typed_arrays) {
		var a = new Uint8Array(M);
		if(a.set) a.set(b);
		else for(; i < b.length; ++i) a[i] = b[i];
		return a;
	}
	b.length = M;
	return b;
}

/* zero-filled arrays for older browsers */
function zero_fill_array(n) {
	var o = new Array(n);
	for(var i = 0; i < n; ++i) o[i] = 0;
	return o;
}var _deflate = (function() {
var _deflateRaw = (function() {
	return function deflateRaw(data, out) {
		var boff = 0;
		while(boff < data.length) {
			var L = Math.min(0xFFFF, data.length - boff);
			var h = boff + L == data.length;
			/* TODO: this is only type 0 stored */
			out.write_shift(1, +h);
			out.write_shift(2, L);
			out.write_shift(2, (~L) & 0xFFFF);
			while(L-- > 0) out[out.l++] = data[boff++];
		}
		return out.l;
	};
})();

return function(data) {
	var buf = new_buf(50+Math.floor(data.length*1.1));
	var off = _deflateRaw(data, buf);
	return buf.slice(0, off);
};
})();
/* modified inflate function also moves original read head */

/* build tree (used for literals and lengths) */
function build_tree(clens, cmap, MAX) {
	var maxlen = 1, w = 0, i = 0, j = 0, ccode = 0, L = clens.length;

	var bl_count  = use_typed_arrays ? new Uint16Array(32) : zero_fill_array(32);
	for(i = 0; i < 32; ++i) bl_count[i] = 0;

	for(i = L; i < MAX; ++i) clens[i] = 0;
	L = clens.length;

	var ctree = use_typed_arrays ? new Uint16Array(L) : zero_fill_array(L); // []

	/* build code tree */
	for(i = 0; i < L; ++i) {
		bl_count[(w = clens[i])]++;
		if(maxlen < w) maxlen = w;
		ctree[i] = 0;
	}
	bl_count[0] = 0;
	for(i = 1; i <= maxlen; ++i) bl_count[i+16] = (ccode = (ccode + bl_count[i-1])<<1);
	for(i = 0; i < L; ++i) {
		ccode = clens[i];
		if(ccode != 0) ctree[i] = bl_count[ccode+16]++;
	}

	/* cmap[maxlen + 4 bits] = (off&15) + (lit<<4) reverse mapping */
	var cleni = 0;
	for(i = 0; i < L; ++i) {
		cleni = clens[i];
		if(cleni != 0) {
			ccode = bit_swap_n(ctree[i], maxlen)>>(maxlen-cleni);
			for(j = (1<<(maxlen + 4 - cleni)) - 1; j>=0; --j)
				cmap[ccode|(j<<cleni)] = (cleni&15) | (i<<4);
		}
	}
	return maxlen;
}

var fix_lmap = use_typed_arrays ? new Uint16Array(512) : zero_fill_array(512);
var fix_dmap = use_typed_arrays ? new Uint16Array(32)  : zero_fill_array(32);
if(!use_typed_arrays) {
	for(var i = 0; i < 512; ++i) fix_lmap[i] = 0;
	for(i = 0; i < 32; ++i) fix_dmap[i] = 0;
}
(function() {
	var dlens = [];
	var i = 0;
	for(;i<32; i++) dlens.push(5);
	build_tree(dlens, fix_dmap, 32);

	var clens = [];
	i = 0;
	for(; i<=143; i++) clens.push(8);
	for(; i<=255; i++) clens.push(9);
	for(; i<=279; i++) clens.push(7);
	for(; i<=287; i++) clens.push(8);
	build_tree(clens, fix_lmap, 288);
})();

var dyn_lmap = use_typed_arrays ? new Uint16Array(32768) : zero_fill_array(32768);
var dyn_dmap = use_typed_arrays ? new Uint16Array(32768) : zero_fill_array(32768);
var dyn_cmap = use_typed_arrays ? new Uint16Array(128)   : zero_fill_array(128);
var dyn_len_1 = 1, dyn_len_2 = 1;

/* 5.5.3 Expanding Huffman Codes */
function dyn(data, boff) {
	/* nomenclature from RFC1951 refers to bit values; these are offset by the implicit constant */
	var _HLIT = read_bits_5(data, boff) + 257; boff += 5;
	var _HDIST = read_bits_5(data, boff) + 1; boff += 5;
	var _HCLEN = read_bits_4(data, boff) + 4; boff += 4;
	var w = 0;

	/* grab and store code lengths */
	var clens = use_typed_arrays ? new Uint8Array(19) : zero_fill_array(19);
	var ctree = [ 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 ];
	var maxlen = 1;
	var bl_count =  use_typed_arrays ? new Uint8Array(8) : zero_fill_array(8);
	var next_code = use_typed_arrays ? new Uint8Array(8) : zero_fill_array(8);
	var L = clens.length; /* 19 */
	for(var i = 0; i < _HCLEN; ++i) {
		clens[CLEN_ORDER[i]] = w = read_bits_3(data, boff);
		if(maxlen < w) maxlen = w;
		bl_count[w]++;
		boff += 3;
	}

	/* build code tree */
	var ccode = 0;
	bl_count[0] = 0;
	for(i = 1; i <= maxlen; ++i) next_code[i] = ccode = (ccode + bl_count[i-1])<<1;
	for(i = 0; i < L; ++i) if((ccode = clens[i]) != 0) ctree[i] = next_code[ccode]++;
	/* cmap[7 bits from stream] = (off&7) + (lit<<3) */
	var cleni = 0;
	for(i = 0; i < L; ++i) {
		cleni = clens[i];
		if(cleni != 0) {
			ccode = bitswap8[ctree[i]]>>(8-cleni);
			for(var j = (1<<(7-cleni))-1; j>=0; --j) dyn_cmap[ccode|(j<<cleni)] = (cleni&7) | (i<<3);
		}
	}

	/* read literal and dist codes at once */
	var hcodes = [];
	maxlen = 1;
	for(; hcodes.length < _HLIT + _HDIST;) {
		ccode = dyn_cmap[read_bits_7(data, boff)];
		boff += ccode & 7;
		switch((ccode >>>= 3)) {
			case 16:
				w = 3 + read_bits_2(data, boff); boff += 2;
				ccode = hcodes[hcodes.length - 1];
				while(w-- > 0) hcodes.push(ccode);
				break;
			case 17:
				w = 3 + read_bits_3(data, boff); boff += 3;
				while(w-- > 0) hcodes.push(0);
				break;
			case 18:
				w = 11 + read_bits_7(data, boff); boff += 7;
				while(w -- > 0) hcodes.push(0);
				break;
			default:
				hcodes.push(ccode);
				if(maxlen < ccode) maxlen = ccode;
				break;
		}
	}

	/* build literal / length trees */
	var h1 = hcodes.slice(0, _HLIT), h2 = hcodes.slice(_HLIT);
	for(i = _HLIT; i < 286; ++i) h1[i] = 0;
	for(i = _HDIST; i < 30; ++i) h2[i] = 0;
	dyn_len_1 = build_tree(h1, dyn_lmap, 286);
	dyn_len_2 = build_tree(h2, dyn_dmap, 30);
	return boff;
}

/* return [ data, bytesRead ] */
function inflate(data, usz) {
	/* shortcircuit for empty buffer [0x03, 0x00] */
	if(data[0] == 3 && !(data[1] & 0x3)) { return [new_raw_buf(usz), 2]; }

	/* bit offset */
	var boff = 0;

	/* header includes final bit and type bits */
	var header = 0;

	var outbuf = new_unsafe_buf(usz ? usz : (1<<18));
	var woff = 0;
	var OL = outbuf.length>>>0;
	var max_len_1 = 0, max_len_2 = 0;

	while((header&1) == 0) {
		header = read_bits_3(data, boff); boff += 3;
		if((header >>> 1) == 0) {
			/* Stored block */
			if(boff & 7) boff += 8 - (boff&7);
			/* 2 bytes sz, 2 bytes bit inverse */
			var sz = data[boff>>>3] | data[(boff>>>3)+1]<<8;
			boff += 32;
			/* push sz bytes */
			if(!usz && OL < woff + sz) { outbuf = realloc(outbuf, woff + sz); OL = outbuf.length; }
			if(typeof data.copy === 'function') {
				// $FlowIgnore
				data.copy(outbuf, woff, boff>>>3, (boff>>>3)+sz);
				woff += sz; boff += 8*sz;
			} else while(sz-- > 0) { outbuf[woff++] = data[boff>>>3]; boff += 8; }
			continue;
		} else if((header >>> 1) == 1) {
			/* Fixed Huffman */
			max_len_1 = 9; max_len_2 = 5;
		} else {
			/* Dynamic Huffman */
			boff = dyn(data, boff);
			max_len_1 = dyn_len_1; max_len_2 = dyn_len_2;
		}
		if(!usz && (OL < woff + 32767)) { outbuf = realloc(outbuf, woff + 32767); OL = outbuf.length; }
		for(;;) { // while(true) is apparently out of vogue in modern JS circles
			/* ingest code and move read head */
			var bits = read_bits_n(data, boff, max_len_1);
			var code = (header>>>1) == 1 ? fix_lmap[bits] : dyn_lmap[bits];
			boff += code & 15; code >>>= 4;
			/* 0-255 are literals, 256 is end of block token, 257+ are copy tokens */
			if(((code>>>8)&0xFF) === 0) outbuf[woff++] = code;
			else if(code == 256) break;
			else {
				code -= 257;
				var len_eb = (code < 8) ? 0 : ((code-4)>>2); if(len_eb > 5) len_eb = 0;
				var tgt = woff + LEN_LN[code];
				/* length extra bits */
				if(len_eb > 0) {
					tgt += read_bits_n(data, boff, len_eb);
					boff += len_eb;
				}

				/* dist code */
				bits = read_bits_n(data, boff, max_len_2);
				code = (header>>>1) == 1 ? fix_dmap[bits] : dyn_dmap[bits];
				boff += code & 15; code >>>= 4;
				var dst_eb = (code < 4 ? 0 : (code-2)>>1);
				var dst = DST_LN[code];
				/* dist extra bits */
				if(dst_eb > 0) {
					dst += read_bits_n(data, boff, dst_eb);
					boff += dst_eb;
				}

				/* in the common case, manual byte copy is faster than TA set / Buffer copy */
				if(!usz && OL < tgt) { outbuf = realloc(outbuf, tgt); OL = outbuf.length; }
				while(woff < tgt) { outbuf[woff] = outbuf[woff - dst]; ++woff; }
			}
		}
	}
	return [usz ? outbuf : outbuf.slice(0, woff), (boff+7)>>>3];
}

function _inflate(payload, usz) {
	var data = payload.slice(payload.l||0);
	var out = inflate(data, usz);
	payload.l += out[1];
	return out[0];
}

function warn_or_throw(wrn, msg) {
	if(wrn) { if(typeof console !== 'undefined') console.error(msg); }
	else throw new Error(msg);
}

function parse_zip(file, options) {
	var blob = file;
	prep_blob(blob, 0);

	var FileIndex = [], FullPaths = [];
	var o = {
		FileIndex: FileIndex,
		FullPaths: FullPaths
	};
	init_cfb(o, { root: options.root });

	/* find end of central directory, start just after signature */
	var i = blob.length - 4;
	while((blob[i] != 0x50 || blob[i+1] != 0x4b || blob[i+2] != 0x05 || blob[i+3] != 0x06) && i >= 0) --i;
	blob.l = i + 4;

	/* parse end of central directory */
	blob.l += 4;
	var fcnt = blob.read_shift(2);
	blob.l += 6;
	var start_cd = blob.read_shift(4);

	/* parse central directory */
	blob.l = start_cd;

	for(i = 0; i < fcnt; ++i) {
		/* trust local file header instead of CD entry */
		blob.l += 20;
		var csz = blob.read_shift(4);
		var usz = blob.read_shift(4);
		var namelen = blob.read_shift(2);
		var efsz = blob.read_shift(2);
		var fcsz = blob.read_shift(2);
		blob.l += 8;
		var offset = blob.read_shift(4);
		var EF = parse_extra_field(blob.slice(blob.l+namelen, blob.l+namelen+efsz));
		blob.l += namelen + efsz + fcsz;

		var L = blob.l;
		blob.l = offset + 4;
		parse_local_file(blob, csz, usz, o, EF);
		blob.l = L;
	}

	return o;
}


/* head starts just after local file header signature */
function parse_local_file(blob, csz, usz, o, EF) {
	/* [local file header] */
	blob.l += 2;
	var flags = blob.read_shift(2);
	var meth = blob.read_shift(2);
	var date = parse_dos_date(blob);

	if(flags & 0x2041) throw new Error("Unsupported ZIP encryption");
	var crc32 = blob.read_shift(4);
	var _csz = blob.read_shift(4);
	var _usz = blob.read_shift(4);

	var namelen = blob.read_shift(2);
	var efsz = blob.read_shift(2);

	// TODO: flags & (1<<11) // UTF8
	var name = ""; for(var i = 0; i < namelen; ++i) name += String.fromCharCode(blob[blob.l++]);
	if(efsz) {
		var ef = parse_extra_field(blob.slice(blob.l, blob.l + efsz));
		if((ef[0x5455]||{}).mt) date = ef[0x5455].mt;
		if(((EF||{})[0x5455]||{}).mt) date = EF[0x5455].mt;
	}
	blob.l += efsz;

	/* [encryption header] */

	/* [file data] */
	var data = blob.slice(blob.l, blob.l + _csz);
	switch(meth) {
		case 8: data = _inflateRawSync(blob, _usz); break;
		case 0: break;
		default: throw new Error("Unsupported ZIP Compression method " + meth);
	}

	/* [data descriptor] */
	var wrn = false;
	if(flags & 8) {
		crc32 = blob.read_shift(4);
		if(crc32 == 0x08074b50) { crc32 = blob.read_shift(4); wrn = true; }
		_csz = blob.read_shift(4);
		_usz = blob.read_shift(4);
	}

	if(_csz != csz) warn_or_throw(wrn, "Bad compressed size: " + csz + " != " + _csz);
	if(_usz != usz) warn_or_throw(wrn, "Bad uncompressed size: " + usz + " != " + _usz);
	var _crc32 = CRC32.buf(data, 0);
	if(crc32 != _crc32) warn_or_throw(wrn, "Bad CRC32 checksum: " + crc32 + " != " + _crc32);
	cfb_add(o, name, data, {unsafe: true, mt: date});
}
function write_zip(cfb, options) {
	var _opts = options || {};
	var out = [], cdirs = [];
	var o = new_buf(1);
	var method = (_opts.compression ? 8 : 0), flags = 0;
	var desc = false;
	if(desc) flags |= 8;
	var i = 0, j = 0;

	var start_cd = 0, fcnt = 0;
	var root = cfb.FullPaths[0], fp = root, fi = cfb.FileIndex[0];
	var crcs = [];
	var sz_cd = 0;

	for(i = 1; i < cfb.FullPaths.length; ++i) {
		fp = cfb.FullPaths[i].slice(root.length); fi = cfb.FileIndex[i];
		if(!fi.size || !fi.content || fp == "\u0001Sh33tJ5") continue;
		var start = start_cd;

		/* TODO: CP437 filename */
		var namebuf = new_buf(fp.length);
		for(j = 0; j < fp.length; ++j) namebuf.write_shift(1, fp.charCodeAt(j) & 0x7F);
		namebuf = namebuf.slice(0, namebuf.l);
		crcs[fcnt] = CRC32.buf(fi.content, 0);

		var outbuf = fi.content;
		if(method == 8) outbuf = _deflateRawSync(outbuf);

		/* local file header */
		o = new_buf(30);
		o.write_shift(4, 0x04034b50);
		o.write_shift(2, 20);
		o.write_shift(2, flags);
		o.write_shift(2, method);
		/* TODO: last mod file time/date */
		if(fi.mt) write_dos_date(o, fi.mt);
		else o.write_shift(4, 0);
		o.write_shift(-4, (flags & 8) ? 0 : crcs[fcnt]);
		o.write_shift(4,  (flags & 8) ? 0 : outbuf.length);
		o.write_shift(4,  (flags & 8) ? 0 : fi.content.length);
		o.write_shift(2, namebuf.length);
		o.write_shift(2, 0);

		start_cd += o.length;
		out.push(o);
		start_cd += namebuf.length;
		out.push(namebuf);

		/* TODO: encryption header ? */
		start_cd += outbuf.length;
		out.push(outbuf);

		/* data descriptor */
		if(flags & 8) {
			o = new_buf(12);
			o.write_shift(-4, crcs[fcnt]);
			o.write_shift(4, outbuf.length);
			o.write_shift(4, fi.content.length);
			start_cd += o.l;
			out.push(o);
		}

		/* central directory */
		o = new_buf(46);
		o.write_shift(4, 0x02014b50);
		o.write_shift(2, 0);
		o.write_shift(2, 20);
		o.write_shift(2, flags);
		o.write_shift(2, method);
		o.write_shift(4, 0); /* TODO: last mod file time/date */
		o.write_shift(-4, crcs[fcnt]);

		o.write_shift(4, outbuf.length);
		o.write_shift(4, fi.content.length);
		o.write_shift(2, namebuf.length);
		o.write_shift(2, 0);
		o.write_shift(2, 0);
		o.write_shift(2, 0);
		o.write_shift(2, 0);
		o.write_shift(4, 0);
		o.write_shift(4, start);

		sz_cd += o.l;
		cdirs.push(o);
		sz_cd += namebuf.length;
		cdirs.push(namebuf);
		++fcnt;
	}

	/* end of central directory */
	o = new_buf(22);
	o.write_shift(4, 0x06054b50);
	o.write_shift(2, 0);
	o.write_shift(2, 0);
	o.write_shift(2, fcnt);
	o.write_shift(2, fcnt);
	o.write_shift(4, sz_cd);
	o.write_shift(4, start_cd);
	o.write_shift(2, 0);

	return bconcat(([bconcat((out)), bconcat(cdirs), o]));
}
function cfb_new(opts) {
	var o = ({});
	init_cfb(o, opts);
	return o;
}

function cfb_add(cfb, name, content, opts) {
	var unsafe = opts && opts.unsafe;
	if(!unsafe) init_cfb(cfb);
	var file = !unsafe && CFB.find(cfb, name);
	if(!file) {
		var fpath = cfb.FullPaths[0];
		if(name.slice(0, fpath.length) == fpath) fpath = name;
		else {
			if(fpath.slice(-1) != "/") fpath += "/";
			fpath = (fpath + name).replace("//","/");
		}
		file = ({name: filename(name), type: 2});
		cfb.FileIndex.push(file);
		cfb.FullPaths.push(fpath);
		if(!unsafe) CFB.utils.cfb_gc(cfb);
	}
file.content = (content);
	file.size = content ? content.length : 0;
	if(opts) {
		if(opts.CLSID) file.clsid = opts.CLSID;
		if(opts.mt) file.mt = opts.mt;
		if(opts.ct) file.ct = opts.ct;
	}
	return file;
}

function cfb_del(cfb, name) {
	init_cfb(cfb);
	var file = CFB.find(cfb, name);
	if(file) for(var j = 0; j < cfb.FileIndex.length; ++j) if(cfb.FileIndex[j] == file) {
		cfb.FileIndex.splice(j, 1);
		cfb.FullPaths.splice(j, 1);
		return true;
	}
	return false;
}

function cfb_mov(cfb, old_name, new_name) {
	init_cfb(cfb);
	var file = CFB.find(cfb, old_name);
	if(file) for(var j = 0; j < cfb.FileIndex.length; ++j) if(cfb.FileIndex[j] == file) {
		cfb.FileIndex[j].name = filename(new_name);
		cfb.FullPaths[j] = new_name;
		return true;
	}
	return false;
}

function cfb_gc(cfb) { rebuild_cfb(cfb, true); }

exports.find = find;
exports.read = read;
exports.parse = parse;
exports.write = write;
exports.writeFile = write_file;
exports.utils = {
	cfb_new: cfb_new,
	cfb_add: cfb_add,
	cfb_del: cfb_del,
	cfb_mov: cfb_mov,
	cfb_gc: cfb_gc,
	ReadShift: ReadShift,
	CheckField: CheckField,
	prep_blob: prep_blob,
	bconcat: bconcat,
	use_zlib: use_zlib,
	_deflateRaw: _deflate,
	_inflateRaw: _inflate,
	consts: consts
};

return exports;
})();

if(typeof require !== 'undefined' && typeof module !== 'undefined' && typeof DO_NOT_EXPORT_CFB === 'undefined') { module.exports = CFB; }
var _fs;
if(typeof require !== 'undefined') try { _fs = require('fs'); } catch(e) {}

/* normalize data for blob ctor */
function blobify(data) {
	if(typeof data === "string") return s2ab(data);
	if(Array.isArray(data)) return a2u(data);
	return data;
}
/* write or download file */
function write_dl(fname, payload, enc) {
	/*global IE_SaveFile, Blob, navigator, saveAs, URL, document, File, chrome */
	if(typeof _fs !== 'undefined' && _fs.writeFileSync) return enc ? _fs.writeFileSync(fname, payload, enc) : _fs.writeFileSync(fname, payload);
	var data = (enc == "utf8") ? utf8write(payload) : payload;
if(typeof IE_SaveFile !== 'undefined') return IE_SaveFile(data, fname);
	if(typeof Blob !== 'undefined') {
		var blob = new Blob([blobify(data)], {type:"application/octet-stream"});
if(typeof navigator !== 'undefined' && navigator.msSaveBlob) return navigator.msSaveBlob(blob, fname);
if(typeof saveAs !== 'undefined') return saveAs(blob, fname);
		if(typeof URL !== 'undefined' && typeof document !== 'undefined' && document.createElement && URL.createObjectURL) {
			var url = URL.createObjectURL(blob);
if(typeof chrome === 'object' && typeof (chrome.downloads||{}).download == "function") {
				if(URL.revokeObjectURL && typeof setTimeout !== 'undefined') setTimeout(function() { URL.revokeObjectURL(url); }, 60000);
				return chrome.downloads.download({ url: url, filename: fname, saveAs: true});
			}
			var a = document.createElement("a");
			if(a.download != null) {
a.download = fname; a.href = url; document.body.appendChild(a); a.click();
document.body.removeChild(a);
				if(URL.revokeObjectURL && typeof setTimeout !== 'undefined') setTimeout(function() { URL.revokeObjectURL(url); }, 60000);
				return url;
			}
		}
	}
	// $FlowIgnore
	if(typeof $ !== 'undefined' && typeof File !== 'undefined' && typeof Folder !== 'undefined') try { // extendscript
		// $FlowIgnore
		var out = File(fname); out.open("w"); out.encoding = "binary";
		if(Array.isArray(payload)) payload = a2s(payload);
		out.write(payload); out.close(); return payload;
	} catch(e) { if(!e.message || !e.message.match(/onstruct/)) throw e; }
	throw new Error("cannot save file " + fname);
}

/* read binary data from file */
function read_binary(path) {
	if(typeof _fs !== 'undefined') return _fs.readFileSync(path);
	// $FlowIgnore
	if(typeof $ !== 'undefined' && typeof File !== 'undefined' && typeof Folder !== 'undefined') try { // extendscript
		// $FlowIgnore
		var infile = File(path); infile.open("r"); infile.encoding = "binary";
		var data = infile.read(); infile.close();
		return data;
	} catch(e) { if(!e.message || !e.message.match(/onstruct/)) throw e; }
	throw new Error("Cannot access file " + path);
}
function keys(o) {
	var ks = Object.keys(o), o2 = [];
	for(var i = 0; i < ks.length; ++i) if(o.hasOwnProperty(ks[i])) o2.push(ks[i]);
	return o2;
}

function evert_key(obj, key) {
	var o = ([]), K = keys(obj);
	for(var i = 0; i !== K.length; ++i) if(o[obj[K[i]][key]] == null) o[obj[K[i]][key]] = K[i];
	return o;
}

function evert(obj) {
	var o = ([]), K = keys(obj);
	for(var i = 0; i !== K.length; ++i) o[obj[K[i]]] = K[i];
	return o;
}

function evert_num(obj) {
	var o = ([]), K = keys(obj);
	for(var i = 0; i !== K.length; ++i) o[obj[K[i]]] = parseInt(K[i],10);
	return o;
}

function evert_arr(obj) {
	var o = ([]), K = keys(obj);
	for(var i = 0; i !== K.length; ++i) {
		if(o[obj[K[i]]] == null) o[obj[K[i]]] = [];
		o[obj[K[i]]].push(K[i]);
	}
	return o;
}

var basedate = new Date(1899, 11, 30, 0, 0, 0); // 2209161600000
var dnthresh = basedate.getTime() + (new Date().getTimezoneOffset() - basedate.getTimezoneOffset()) * 60000;
function datenum(v, date1904) {
	var epoch = v.getTime();
	if(date1904) epoch -= 1462*24*60*60*1000;
	return (epoch - dnthresh) / (24 * 60 * 60 * 1000);
}
function numdate(v) {
	var out = new Date();
	out.setTime(v * 24 * 60 * 60 * 1000 + dnthresh);
	return out;
}

/* ISO 8601 Duration */
function parse_isodur(s) {
	var sec = 0, mt = 0, time = false;
	var m = s.match(/P([0-9\.]+Y)?([0-9\.]+M)?([0-9\.]+D)?T([0-9\.]+H)?([0-9\.]+M)?([0-9\.]+S)?/);
	if(!m) throw new Error("|" + s + "| is not an ISO8601 Duration");
	for(var i = 1; i != m.length; ++i) {
		if(!m[i]) continue;
		mt = 1;
		if(i > 3) time = true;
		switch(m[i].slice(m[i].length-1)) {
			case 'Y':
				throw new Error("Unsupported ISO Duration Field: " + m[i].slice(m[i].length-1));
			case 'D': mt *= 24;
				/* falls through */
			case 'H': mt *= 60;
				/* falls through */
			case 'M':
				if(!time) throw new Error("Unsupported ISO Duration Field: M");
				else mt *= 60;
				/* falls through */
			case 'S': break;
		}
		sec += mt * parseInt(m[i], 10);
	}
	return sec;
}

var good_pd_date = new Date('2017-02-19T19:06:09.000Z');
if(isNaN(good_pd_date.getFullYear())) good_pd_date = new Date('2/19/17');
var good_pd = good_pd_date.getFullYear() == 2017;
/* parses a date as a local date */
function parseDate(str, fixdate) {
	var d = new Date(str);
	if(good_pd) {
if(fixdate > 0) d.setTime(d.getTime() + d.getTimezoneOffset() * 60 * 1000);
		else if(fixdate < 0) d.setTime(d.getTime() - d.getTimezoneOffset() * 60 * 1000);
		return d;
	}
	if(str instanceof Date) return str;
	if(good_pd_date.getFullYear() == 1917 && !isNaN(d.getFullYear())) {
		var s = d.getFullYear();
		if(str.indexOf("" + s) > -1) return d;
		d.setFullYear(d.getFullYear() + 100); return d;
	}
	var n = str.match(/\d+/g)||["2017","2","19","0","0","0"];
	var out = new Date(+n[0], +n[1] - 1, +n[2], (+n[3]||0), (+n[4]||0), (+n[5]||0));
	if(str.indexOf("Z") > -1) out = new Date(out.getTime() - out.getTimezoneOffset() * 60 * 1000);
	return out;
}

function cc2str(arr) {
	var o = "";
	for(var i = 0; i != arr.length; ++i) o += String.fromCharCode(arr[i]);
	return o;
}

function dup(o) {
	if(typeof JSON != 'undefined' && !Array.isArray(o)) return JSON.parse(JSON.stringify(o));
	if(typeof o != 'object' || o == null) return o;
	if(o instanceof Date) return new Date(o.getTime());
	var out = {};
	for(var k in o) if(o.hasOwnProperty(k)) out[k] = dup(o[k]);
	return out;
}

function fill(c,l) { var o = ""; while(o.length < l) o+=c; return o; }

/* TODO: stress test */
function fuzzynum(s) {
	var v = Number(s);
	if(!isNaN(v)) return v;
	var wt = 1;
	var ss = s.replace(/([\d]),([\d])/g,"$1$2").replace(/[$]/g,"").replace(/[%]/g, function() { wt *= 100; return "";});
	if(!isNaN(v = Number(ss))) return v / wt;
	ss = ss.replace(/[(](.*)[)]/,function($$, $1) { wt = -wt; return $1;});
	if(!isNaN(v = Number(ss))) return v / wt;
	return v;
}
function fuzzydate(s) {
	var o = new Date(s), n = new Date(NaN);
	var y = o.getYear(), m = o.getMonth(), d = o.getDate();
	if(isNaN(d)) return n;
	if(y < 0 || y > 8099) return n;
	if((m > 0 || d > 1) && y != 101) return o;
	if(s.toLowerCase().match(/jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec/)) return o;
	if(s.match(/[^-0-9:,\/\\]/)) return n;
	return o;
}

var safe_split_regex = "abacaba".split(/(:?b)/i).length == 5;
function split_regex(str, re, def) {
	if(safe_split_regex || typeof re == "string") return str.split(re);
	var p = str.split(re), o = [p[0]];
	for(var i = 1; i < p.length; ++i) { o.push(def); o.push(p[i]); }
	return o;
}
function getdatastr(data) {
	if(!data) return null;
	if(data.data) return debom(data.data);
	if(data.asNodeBuffer && has_buf) return debom(data.asNodeBuffer().toString('binary'));
	if(data.asBinary) return debom(data.asBinary());
	if(data._data && data._data.getContent) return debom(cc2str(Array.prototype.slice.call(data._data.getContent(),0)));
	return null;
}

function getdatabin(data) {
	if(!data) return null;
	if(data.data) return char_codes(data.data);
	if(data.asNodeBuffer && has_buf) return data.asNodeBuffer();
	if(data._data && data._data.getContent) {
		var o = data._data.getContent();
		if(typeof o == "string") return char_codes(o);
		return Array.prototype.slice.call(o);
	}
	return null;
}

function getdata(data) { return (data && data.name.slice(-4) === ".bin") ? getdatabin(data) : getdatastr(data); }

/* Part 2 Section 10.1.2 "Mapping Content Types" Names are case-insensitive */
/* OASIS does not comment on filename case sensitivity */
function safegetzipfile(zip, file) {
	var k = keys(zip.files);
	var f = file.toLowerCase(), g = f.replace(/\//g,'\\');
	for(var i=0; i<k.length; ++i) {
		var n = k[i].toLowerCase();
		if(f == n || g == n) return zip.files[k[i]];
	}
	return null;
}

function getzipfile(zip, file) {
	var o = safegetzipfile(zip, file);
	if(o == null) throw new Error("Cannot find file " + file + " in zip");
	return o;
}

function getzipdata(zip, file, safe) {
	if(!safe) return getdata(getzipfile(zip, file));
	if(!file) return null;
	try { return getzipdata(zip, file); } catch(e) { return null; }
}

function getzipstr(zip, file, safe) {
	if(!safe) return getdatastr(getzipfile(zip, file));
	if(!file) return null;
	try { return getzipstr(zip, file); } catch(e) { return null; }
}

function zipentries(zip) {
	var k = keys(zip.files), o = [];
	for(var i = 0; i < k.length; ++i) if(k[i].slice(-1) != '/') o.push(k[i]);
	return o.sort();
}

var jszip;
/*global JSZipSync:true */
if(typeof JSZipSync !== 'undefined') jszip = JSZipSync;
if(typeof exports !== 'undefined') {
	if(typeof module !== 'undefined' && module.exports) {
		if(typeof jszip === 'undefined') jszip = undefined;
	}
}

function resolve_path(path, base) {
	var result = base.split('/');
	if(base.slice(-1) != "/") result.pop(); // folder path
	var target = path.split('/');
	while (target.length !== 0) {
		var step = target.shift();
		if (step === '..') result.pop();
		else if (step !== '.') result.push(step);
	}
	return result.join('/');
}
var XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
var attregexg=/([^"\s?>\/]+)\s*=\s*((?:")([^"]*)(?:")|(?:')([^']*)(?:')|([^'">\s]+))/g;
var tagregex=/<[\/\?]?[a-zA-Z0-9:]+(?:\s+[^"\s?>\/]+\s*=\s*(?:"[^"]*"|'[^']*'|[^'">\s=]+))*\s?[\/\?]?>/g;
if(!(XML_HEADER.match(tagregex))) tagregex = /<[^>]*>/g;
var nsregex=/<\w*:/, nsregex2 = /<(\/?)\w+:/;
function parsexmltag(tag, skip_root) {
	var z = ({});
	var eq = 0, c = 0;
	for(; eq !== tag.length; ++eq) if((c = tag.charCodeAt(eq)) === 32 || c === 10 || c === 13) break;
	if(!skip_root) z[0] = tag.slice(0, eq);
	if(eq === tag.length) return z;
	var m = tag.match(attregexg), j=0, v="", i=0, q="", cc="", quot = 1;
	if(m) for(i = 0; i != m.length; ++i) {
		cc = m[i];
		for(c=0; c != cc.length; ++c) if(cc.charCodeAt(c) === 61) break;
		q = cc.slice(0,c).trim();
		while(cc.charCodeAt(c+1) == 32) ++c;
		quot = ((eq=cc.charCodeAt(c+1)) == 34 || eq == 39) ? 1 : 0;
		v = cc.slice(c+1+quot, cc.length-quot);
		for(j=0;j!=q.length;++j) if(q.charCodeAt(j) === 58) break;
		if(j===q.length) {
			if(q.indexOf("_") > 0) q = q.slice(0, q.indexOf("_")); // from ods
			z[q] = v;
			z[q.toLowerCase()] = v;
		}
		else {
			var k = (j===5 && q.slice(0,5)==="xmlns"?"xmlns":"")+q.slice(j+1);
			if(z[k] && q.slice(j-3,j) == "ext") continue; // from ods
			z[k] = v;
			z[k.toLowerCase()] = v;
		}
	}
	return z;
}
function strip_ns(x) { return x.replace(nsregex2, "<$1"); }

var encodings = {
	'&quot;': '"',
	'&apos;': "'",
	'&gt;': '>',
	'&lt;': '<',
	'&amp;': '&'
};
var rencoding = evert(encodings);
//var rencstr = "&<>'\"".split("");

// TODO: CP remap (need to read file version to determine OS)
var unescapexml = (function() {
	/* 22.4.2.4 bstr (Basic String) */
	var encregex = /&(?:quot|apos|gt|lt|amp|#x?([\da-fA-F]+));/g, coderegex = /_x([\da-fA-F]{4})_/g;
	return function unescapexml(text) {
		var s = text + '', i = s.indexOf("<![CDATA[");
		if(i == -1) return s.replace(encregex, function($$, $1) { return encodings[$$]||String.fromCharCode(parseInt($1,$$.indexOf("x")>-1?16:10))||$$; }).replace(coderegex,function(m,c) {return String.fromCharCode(parseInt(c,16));});
		var j = s.indexOf("]]>");
		return unescapexml(s.slice(0, i)) + s.slice(i+9,j) + unescapexml(s.slice(j+3));
	};
})();

var decregex=/[&<>'"]/g, charegex = /[\u0000-\u0008\u000b-\u001f]/g;
function escapexml(text){
	var s = text + '';
	return s.replace(decregex, function(y) { return rencoding[y]; }).replace(charegex,function(s) { return "_x" + ("000"+s.charCodeAt(0).toString(16)).slice(-4) + "_";});
}
function escapexmltag(text){ return escapexml(text).replace(/ /g,"_x0020_"); }

var htmlcharegex = /[\u0000-\u001f]/g;
function escapehtml(text){
	var s = text + '';
	return s.replace(decregex, function(y) { return rencoding[y]; }).replace(/\n/g, "<br/>").replace(htmlcharegex,function(s) { return "&#x" + ("000"+s.charCodeAt(0).toString(16)).slice(-4) + ";"; });
}

function escapexlml(text){
	var s = text + '';
	return s.replace(decregex, function(y) { return rencoding[y]; }).replace(htmlcharegex,function(s) { return "&#x" + (s.charCodeAt(0).toString(16)).toUpperCase() + ";"; });
}

/* TODO: handle codepages */
var xlml_fixstr = (function() {
	var entregex = /&#(\d+);/g;
	function entrepl($$,$1) { return String.fromCharCode(parseInt($1,10)); }
	return function xlml_fixstr(str) { return str.replace(entregex,entrepl); };
})();
var xlml_unfixstr = (function() {
	return function xlml_unfixstr(str) { return str.replace(/(\r\n|[\r\n])/g,"\&#10;"); };
})();

function parsexmlbool(value) {
	switch(value) {
		case 1: case true: case '1': case 'true': case 'TRUE': return true;
		/* case '0': case 'false': case 'FALSE':*/
		default: return false;
	}
}

var utf8read = function utf8reada(orig) {
	var out = "", i = 0, c = 0, d = 0, e = 0, f = 0, w = 0;
	while (i < orig.length) {
		c = orig.charCodeAt(i++);
		if (c < 128) { out += String.fromCharCode(c); continue; }
		d = orig.charCodeAt(i++);
		if (c>191 && c<224) { f = ((c & 31) << 6); f |= (d & 63); out += String.fromCharCode(f); continue; }
		e = orig.charCodeAt(i++);
		if (c < 240) { out += String.fromCharCode(((c & 15) << 12) | ((d & 63) << 6) | (e & 63)); continue; }
		f = orig.charCodeAt(i++);
		w = (((c & 7) << 18) | ((d & 63) << 12) | ((e & 63) << 6) | (f & 63))-65536;
		out += String.fromCharCode(0xD800 + ((w>>>10)&1023));
		out += String.fromCharCode(0xDC00 + (w&1023));
	}
	return out;
};

var utf8write = function(orig) {
	var out = [], i = 0, c = 0, d = 0;
	while(i < orig.length) {
		c = orig.charCodeAt(i++);
		switch(true) {
			case c < 128: out.push(String.fromCharCode(c)); break;
			case c < 2048:
				out.push(String.fromCharCode(192 + (c >> 6)));
				out.push(String.fromCharCode(128 + (c & 63)));
				break;
			case c >= 55296 && c < 57344:
				c -= 55296; d = orig.charCodeAt(i++) - 56320 + (c<<10);
				out.push(String.fromCharCode(240 + ((d >>18) & 7)));
				out.push(String.fromCharCode(144 + ((d >>12) & 63)));
				out.push(String.fromCharCode(128 + ((d >> 6) & 63)));
				out.push(String.fromCharCode(128 + (d & 63)));
				break;
			default:
				out.push(String.fromCharCode(224 + (c >> 12)));
				out.push(String.fromCharCode(128 + ((c >> 6) & 63)));
				out.push(String.fromCharCode(128 + (c & 63)));
		}
	}
	return out.join("");
};

if(has_buf) {
	var utf8readb = function utf8readb(data) {
		var out = Buffer.alloc(2*data.length), w, i, j = 1, k = 0, ww=0, c;
		for(i = 0; i < data.length; i+=j) {
			j = 1;
			if((c=data.charCodeAt(i)) < 128) w = c;
			else if(c < 224) { w = (c&31)*64+(data.charCodeAt(i+1)&63); j=2; }
			else if(c < 240) { w=(c&15)*4096+(data.charCodeAt(i+1)&63)*64+(data.charCodeAt(i+2)&63); j=3; }
			else { j = 4;
				w = (c & 7)*262144+(data.charCodeAt(i+1)&63)*4096+(data.charCodeAt(i+2)&63)*64+(data.charCodeAt(i+3)&63);
				w -= 65536; ww = 0xD800 + ((w>>>10)&1023); w = 0xDC00 + (w&1023);
			}
			if(ww !== 0) { out[k++] = ww&255; out[k++] = ww>>>8; ww = 0; }
			out[k++] = w%256; out[k++] = w>>>8;
		}
		return out.slice(0,k).toString('ucs2');
	};
	var corpus = "foo bar baz\u00e2\u0098\u0083\u00f0\u009f\u008d\u00a3";
	if(utf8read(corpus) == utf8readb(corpus)) utf8read = utf8readb;
	// $FlowIgnore
	var utf8readc = function utf8readc(data) { return Buffer_from(data, 'binary').toString('utf8'); };
	if(utf8read(corpus) == utf8readc(corpus)) utf8read = utf8readc;

	// $FlowIgnore
	utf8write = function(data) { return Buffer_from(data, 'utf8').toString("binary"); };
}

// matches <foo>...</foo> extracts content
var matchtag = (function() {
	var mtcache = ({});
	return function matchtag(f,g) {
		var t = f+"|"+(g||"");
		if(mtcache[t]) return mtcache[t];
		return (mtcache[t] = new RegExp('<(?:\\w+:)?'+f+'(?: xml:space="preserve")?(?:[^>]*)>([\\s\\S]*?)</(?:\\w+:)?'+f+'>',((g||""))));
	};
})();

var htmldecode = (function() {
	var entities = [
		['nbsp', ' '], ['middot', '·'],
		['quot', '"'], ['apos', "'"], ['gt',   '>'], ['lt',   '<'], ['amp',  '&']
	].map(function(x) { return [new RegExp('&' + x[0] + ';', "g"), x[1]]; });
	return function htmldecode(str) {
		var o = str.replace(/^[\t\n\r ]+/, "").replace(/[\t\n\r ]+$/,"").replace(/[\t\n\r ]+/g, " ").replace(/<\s*[bB][rR]\s*\/?>/g,"\n").replace(/<[^>]*>/g,"");
		for(var i = 0; i < entities.length; ++i) o = o.replace(entities[i][0], entities[i][1]);
		return o;
	};
})();

var vtregex = (function(){ var vt_cache = {};
	return function vt_regex(bt) {
		if(vt_cache[bt] !== undefined) return vt_cache[bt];
		return (vt_cache[bt] = new RegExp("<(?:vt:)?" + bt + ">([\\s\\S]*?)</(?:vt:)?" + bt + ">", 'g') );
};})();
var vtvregex = /<\/?(?:vt:)?variant>/g, vtmregex = /<(?:vt:)([^>]*)>([\s\S]*)</;
function parseVector(data, opts) {
	var h = parsexmltag(data);

	var matches = data.match(vtregex(h.baseType))||[];
	var res = [];
	if(matches.length != h.size) {
		if(opts.WTF) throw new Error("unexpected vector length " + matches.length + " != " + h.size);
		return res;
	}
	matches.forEach(function(x) {
		var v = x.replace(vtvregex,"").match(vtmregex);
		if(v) res.push({v:utf8read(v[2]), t:v[1]});
	});
	return res;
}

var wtregex = /(^\s|\s$|\n)/;
function writetag(f,g) { return '<' + f + (g.match(wtregex)?' xml:space="preserve"' : "") + '>' + g + '</' + f + '>'; }

function wxt_helper(h) { return keys(h).map(function(k) { return " " + k + '="' + h[k] + '"';}).join(""); }
function writextag(f,g,h) { return '<' + f + ((h != null) ? wxt_helper(h) : "") + ((g != null) ? (g.match(wtregex)?' xml:space="preserve"' : "") + '>' + g + '</' + f : "/") + '>';}

function write_w3cdtf(d, t) { try { return d.toISOString().replace(/\.\d*/,""); } catch(e) { if(t) throw e; } return ""; }

function write_vt(s) {
	switch(typeof s) {
		case 'string': return writextag('vt:lpwstr', s);
		case 'number': return writextag((s|0)==s?'vt:i4':'vt:r8', String(s));
		case 'boolean': return writextag('vt:bool',s?'true':'false');
	}
	if(s instanceof Date) return writextag('vt:filetime', write_w3cdtf(s));
	throw new Error("Unable to serialize " + s);
}

var XMLNS = ({
	'dc': 'http://purl.org/dc/elements/1.1/',
	'dcterms': 'http://purl.org/dc/terms/',
	'dcmitype': 'http://purl.org/dc/dcmitype/',
	'mx': 'http://schemas.microsoft.com/office/mac/excel/2008/main',
	'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
	'sjs': 'http://schemas.openxmlformats.org/package/2006/sheetjs/core-properties',
	'vt': 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
	'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
	'xsd': 'http://www.w3.org/2001/XMLSchema'
});

XMLNS.main = [
	'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
	'http://purl.oclc.org/ooxml/spreadsheetml/main',
	'http://schemas.microsoft.com/office/excel/2006/main',
	'http://schemas.microsoft.com/office/excel/2006/2'
];

var XLMLNS = ({
	'o':    'urn:schemas-microsoft-com:office:office',
	'x':    'urn:schemas-microsoft-com:office:excel',
	'ss':   'urn:schemas-microsoft-com:office:spreadsheet',
	'dt':   'uuid:C2F41010-65B3-11d1-A29F-00AA00C14882',
	'mv':   'http://macVmlSchemaUri',
	'v':    'urn:schemas-microsoft-com:vml',
	'html': 'http://www.w3.org/TR/REC-html40'
});
function read_double_le(b, idx) {
	var s = 1 - 2 * (b[idx + 7] >>> 7);
	var e = ((b[idx + 7] & 0x7f) << 4) + ((b[idx + 6] >>> 4) & 0x0f);
	var m = (b[idx+6]&0x0f);
	for(var i = 5; i >= 0; --i) m = m * 256 + b[idx + i];
	if(e == 0x7ff) return m == 0 ? (s * Infinity) : NaN;
	if(e == 0) e = -1022;
	else { e -= 1023; m += Math.pow(2,52); }
	return s * Math.pow(2, e - 52) * m;
}

function write_double_le(b, v, idx) {
	var bs = ((((v < 0) || (1/v == -Infinity)) ? 1 : 0) << 7), e = 0, m = 0;
	var av = bs ? (-v) : v;
	if(!isFinite(av)) { e = 0x7ff; m = isNaN(v) ? 0x6969 : 0; }
	else if(av == 0) e = m = 0;
	else {
		e = Math.floor(Math.log(av) / Math.LN2);
		m = av * Math.pow(2, 52 - e);
		if((e <= -1023) && (!isFinite(m) || (m < Math.pow(2,52)))) { e = -1022; }
		else { m -= Math.pow(2,52); e+=1023; }
	}
	for(var i = 0; i <= 5; ++i, m/=256) b[idx + i] = m & 0xff;
	b[idx + 6] = ((e & 0x0f) << 4) | (m & 0xf);
	b[idx + 7] = (e >> 4) | bs;
}

var __toBuffer = function(bufs) { var x=[],w=10240; for(var i=0;i<bufs[0].length;++i) if(bufs[0][i]) for(var j=0,L=bufs[0][i].length;j<L;j+=w) x.push.apply(x, bufs[0][i].slice(j,j+w)); return x; };
var ___toBuffer = __toBuffer;
var __utf16le = function(b,s,e) { var ss=[]; for(var i=s; i<e; i+=2) ss.push(String.fromCharCode(__readUInt16LE(b,i))); return ss.join("").replace(chr0,''); };
var ___utf16le = __utf16le;
var __hexlify = function(b,s,l) { var ss=[]; for(var i=s; i<s+l; ++i) ss.push(("0" + b[i].toString(16)).slice(-2)); return ss.join(""); };
var ___hexlify = __hexlify;
var __utf8 = function(b,s,e) { var ss=[]; for(var i=s; i<e; i++) ss.push(String.fromCharCode(__readUInt8(b,i))); return ss.join(""); };
var ___utf8 = __utf8;
var __lpstr = function(b,i) { var len = __readUInt32LE(b,i); return len > 0 ? __utf8(b, i+4,i+4+len-1) : "";};
var ___lpstr = __lpstr;
var __cpstr = function(b,i) { var len = __readUInt32LE(b,i); return len > 0 ? __utf8(b, i+4,i+4+len-1) : "";};
var ___cpstr = __cpstr;
var __lpwstr = function(b,i) { var len = 2*__readUInt32LE(b,i); return len > 0 ? __utf8(b, i+4,i+4+len-1) : "";};
var ___lpwstr = __lpwstr;
var __lpp4, ___lpp4;
__lpp4 = ___lpp4 = function lpp4_(b,i) { var len = __readUInt32LE(b,i); return len > 0 ? __utf16le(b, i+4,i+4+len) : "";};
var __8lpp4 = function(b,i) { var len = __readUInt32LE(b,i); return len > 0 ? __utf8(b, i+4,i+4+len) : "";};
var ___8lpp4 = __8lpp4;
var __double, ___double;
__double = ___double = function(b, idx) { return read_double_le(b, idx);};
var is_buf = function is_buf_a(a) { return Array.isArray(a); };

if(has_buf) {
	__utf16le = function(b,s,e) { if(!Buffer.isBuffer(b)) return ___utf16le(b,s,e); return b.toString('utf16le',s,e).replace(chr0,'')/*.replace(chr1,'!')*/; };
	__hexlify = function(b,s,l) { return Buffer.isBuffer(b) ? b.toString('hex',s,s+l) : ___hexlify(b,s,l); };
	__lpstr = function lpstr_b(b, i) { if(!Buffer.isBuffer(b)) return ___lpstr(b, i); var len = b.readUInt32LE(i); return len > 0 ? b.toString('utf8',i+4,i+4+len-1) : "";};
	__cpstr = function cpstr_b(b, i) { if(!Buffer.isBuffer(b)) return ___cpstr(b, i); var len = b.readUInt32LE(i); return len > 0 ? b.toString('utf8',i+4,i+4+len-1) : "";};
	__lpwstr = function lpwstr_b(b, i) { if(!Buffer.isBuffer(b)) return ___lpwstr(b, i); var len = 2*b.readUInt32LE(i); return b.toString('utf16le',i+4,i+4+len-1);};
	__lpp4 = function lpp4_b(b, i) { if(!Buffer.isBuffer(b)) return ___lpp4(b, i); var len = b.readUInt32LE(i); return b.toString('utf16le',i+4,i+4+len);};
	__8lpp4 = function lpp4_8b(b, i) { if(!Buffer.isBuffer(b)) return ___8lpp4(b, i); var len = b.readUInt32LE(i); return b.toString('utf8',i+4,i+4+len);};
	__utf8 = function utf8_b(b, s, e) { return (Buffer.isBuffer(b)) ? b.toString('utf8',s,e) : ___utf8(b,s,e); };
	__toBuffer = function(bufs) { return (bufs[0].length > 0 && Buffer.isBuffer(bufs[0][0])) ? Buffer.concat(bufs[0]) : ___toBuffer(bufs);};
	bconcat = function(bufs) { return Buffer.isBuffer(bufs[0]) ? Buffer.concat(bufs) : [].concat.apply([], bufs); };
	__double = function double_(b, i) { if(Buffer.isBuffer(b)) return b.readDoubleLE(i); return ___double(b,i); };
	is_buf = function is_buf_b(a) { return Buffer.isBuffer(a) || Array.isArray(a); };
}

/* from js-xls */
if(typeof cptable !== 'undefined') {
	__utf16le = function(b,s,e) { return cptable.utils.decode(1200, b.slice(s,e)).replace(chr0, ''); };
	__utf8 = function(b,s,e) { return cptable.utils.decode(65001, b.slice(s,e)); };
	__lpstr = function(b,i) { var len = __readUInt32LE(b,i); return len > 0 ? cptable.utils.decode(current_ansi, b.slice(i+4, i+4+len-1)) : "";};
	__cpstr = function(b,i) { var len = __readUInt32LE(b,i); return len > 0 ? cptable.utils.decode(current_codepage, b.slice(i+4, i+4+len-1)) : "";};
	__lpwstr = function(b,i) { var len = 2*__readUInt32LE(b,i); return len > 0 ? cptable.utils.decode(1200, b.slice(i+4,i+4+len-1)) : "";};
	__lpp4 = function(b,i) { var len = __readUInt32LE(b,i); return len > 0 ? cptable.utils.decode(1200, b.slice(i+4,i+4+len)) : "";};
	__8lpp4 = function(b,i) { var len = __readUInt32LE(b,i); return len > 0 ? cptable.utils.decode(65001, b.slice(i+4,i+4+len)) : "";};
}

var __readUInt8 = function(b, idx) { return b[idx]; };
var __readUInt16LE = function(b, idx) { return (b[idx+1]*(1<<8))+b[idx]; };
var __readInt16LE = function(b, idx) { var u = (b[idx+1]*(1<<8))+b[idx]; return (u < 0x8000) ? u : ((0xffff - u + 1) * -1); };
var __readUInt32LE = function(b, idx) { return b[idx+3]*(1<<24)+(b[idx+2]<<16)+(b[idx+1]<<8)+b[idx]; };
var __readInt32LE = function(b, idx) { return (b[idx+3]<<24)|(b[idx+2]<<16)|(b[idx+1]<<8)|b[idx]; };
var __readInt32BE = function(b, idx) { return (b[idx]<<24)|(b[idx+1]<<16)|(b[idx+2]<<8)|b[idx+3]; };

function ReadShift(size, t) {
	var o="", oI, oR, oo=[], w, vv, i, loc;
	switch(t) {
		case 'dbcs':
			loc = this.l;
			if(has_buf && Buffer.isBuffer(this)) o = this.slice(this.l, this.l+2*size).toString("utf16le");
			else for(i = 0; i < size; ++i) { o+=String.fromCharCode(__readUInt16LE(this, loc)); loc+=2; }
			size *= 2;
			break;

		case 'utf8': o = __utf8(this, this.l, this.l + size); break;
		case 'utf16le': size *= 2; o = __utf16le(this, this.l, this.l + size); break;

		case 'wstr':
			if(typeof cptable !== 'undefined') o = cptable.utils.decode(current_codepage, this.slice(this.l, this.l+2*size));
			else return ReadShift.call(this, size, 'dbcs');
			size = 2 * size; break;

		/* [MS-OLEDS] 2.1.4 LengthPrefixedAnsiString */
		case 'lpstr-ansi': o = __lpstr(this, this.l); size = 4 + __readUInt32LE(this, this.l); break;
		case 'lpstr-cp': o = __cpstr(this, this.l); size = 4 + __readUInt32LE(this, this.l); break;
		/* [MS-OLEDS] 2.1.5 LengthPrefixedUnicodeString */
		case 'lpwstr': o = __lpwstr(this, this.l); size = 4 + 2 * __readUInt32LE(this, this.l); break;
		/* [MS-OFFCRYPTO] 2.1.2 Length-Prefixed Padded Unicode String (UNICODE-LP-P4) */
		case 'lpp4': size = 4 +  __readUInt32LE(this, this.l); o = __lpp4(this, this.l); if(size & 0x02) size += 2; break;
		/* [MS-OFFCRYPTO] 2.1.3 Length-Prefixed UTF-8 String (UTF-8-LP-P4) */
		case '8lpp4': size = 4 +  __readUInt32LE(this, this.l); o = __8lpp4(this, this.l); if(size & 0x03) size += 4 - (size & 0x03); break;

		case 'cstr': size = 0; o = "";
			while((w=__readUInt8(this, this.l + size++))!==0) oo.push(_getchar(w));
			o = oo.join(""); break;
		case '_wstr': size = 0; o = "";
			while((w=__readUInt16LE(this,this.l +size))!==0){oo.push(_getchar(w));size+=2;}
			size+=2; o = oo.join(""); break;

		/* sbcs and dbcs support continue records in the SST way TODO codepages */
		case 'dbcs-cont': o = ""; loc = this.l;
			for(i = 0; i < size; ++i) {
				if(this.lens && this.lens.indexOf(loc) !== -1) {
					w = __readUInt8(this, loc);
					this.l = loc + 1;
					vv = ReadShift.call(this, size-i, w ? 'dbcs-cont' : 'sbcs-cont');
					return oo.join("") + vv;
				}
				oo.push(_getchar(__readUInt16LE(this, loc)));
				loc+=2;
			} o = oo.join(""); size *= 2; break;

		case 'cpstr':
			if(typeof cptable !== 'undefined') {
				o = cptable.utils.decode(current_codepage, this.slice(this.l, this.l + size));
				break;
			}
		/* falls through */
		case 'sbcs-cont': o = ""; loc = this.l;
			for(i = 0; i != size; ++i) {
				if(this.lens && this.lens.indexOf(loc) !== -1) {
					w = __readUInt8(this, loc);
					this.l = loc + 1;
					vv = ReadShift.call(this, size-i, w ? 'dbcs-cont' : 'sbcs-cont');
					return oo.join("") + vv;
				}
				oo.push(_getchar(__readUInt8(this, loc)));
				loc+=1;
			} o = oo.join(""); break;

		default:
	switch(size) {
		case 1: oI = __readUInt8(this, this.l); this.l++; return oI;
		case 2: oI = (t === 'i' ? __readInt16LE : __readUInt16LE)(this, this.l); this.l += 2; return oI;
		case 4: case -4:
			if(t === 'i' || ((this[this.l+3] & 0x80)===0)) { oI = ((size > 0) ? __readInt32LE : __readInt32BE)(this, this.l); this.l += 4; return oI; }
			else { oR = __readUInt32LE(this, this.l); this.l += 4; } return oR;
		case 8: case -8:
			if(t === 'f') {
				if(size == 8) oR = __double(this, this.l);
				else oR = __double([this[this.l+7],this[this.l+6],this[this.l+5],this[this.l+4],this[this.l+3],this[this.l+2],this[this.l+1],this[this.l+0]], 0);
				this.l += 8; return oR;
			} else size = 8;
		/* falls through */
		case 16: o = __hexlify(this, this.l, size); break;
	}}
	this.l+=size; return o;
}

var __writeUInt32LE = function(b, val, idx) { b[idx] = (val & 0xFF); b[idx+1] = ((val >>> 8) & 0xFF); b[idx+2] = ((val >>> 16) & 0xFF); b[idx+3] = ((val >>> 24) & 0xFF); };
var __writeInt32LE  = function(b, val, idx) { b[idx] = (val & 0xFF); b[idx+1] = ((val >> 8) & 0xFF); b[idx+2] = ((val >> 16) & 0xFF); b[idx+3] = ((val >> 24) & 0xFF); };
var __writeUInt16LE = function(b, val, idx) { b[idx] = (val & 0xFF); b[idx+1] = ((val >>> 8) & 0xFF); };

function WriteShift(t, val, f) {
	var size = 0, i = 0;
	if(f === 'dbcs') {
for(i = 0; i != val.length; ++i) __writeUInt16LE(this, val.charCodeAt(i), this.l + 2 * i);
		size = 2 * val.length;
	} else if(f === 'sbcs') {
		/* TODO: codepage */
val = val.replace(/[^\x00-\x7F]/g, "_");
for(i = 0; i != val.length; ++i) this[this.l + i] = (val.charCodeAt(i) & 0xFF);
		size = val.length;
	} else if(f === 'hex') {
		for(; i < t; ++i) {
this[this.l++] = (parseInt(val.slice(2*i, 2*i+2), 16)||0);
		} return this;
	} else if(f === 'utf16le') {
var end = Math.min(this.l + t, this.length);
			for(i = 0; i < Math.min(val.length, t); ++i) {
				var cc = val.charCodeAt(i);
				this[this.l++] = (cc & 0xff);
				this[this.l++] = (cc >> 8);
			}
			while(this.l < end) this[this.l++] = 0;
			return this;
	} else  switch(t) {
		case  1: size = 1; this[this.l] = val&0xFF; break;
		case  2: size = 2; this[this.l] = val&0xFF; val >>>= 8; this[this.l+1] = val&0xFF; break;
		case  3: size = 3; this[this.l] = val&0xFF; val >>>= 8; this[this.l+1] = val&0xFF; val >>>= 8; this[this.l+2] = val&0xFF; break;
		case  4: size = 4; __writeUInt32LE(this, val, this.l); break;
		case  8: size = 8; if(f === 'f') { write_double_le(this, val, this.l); break; }
		/* falls through */
		case 16: break;
		case -4: size = 4; __writeInt32LE(this, val, this.l); break;
	}
	this.l += size; return this;
}

function CheckField(hexstr, fld) {
	var m = __hexlify(this,this.l,hexstr.length>>1);
	if(m !== hexstr) throw new Error(fld + 'Expected ' + hexstr + ' saw ' + m);
	this.l += hexstr.length>>1;
}

function prep_blob(blob, pos) {
	blob.l = pos;
	blob.read_shift = ReadShift;
	blob.chk = CheckField;
	blob.write_shift = WriteShift;
}

function parsenoop(blob, length) { blob.l += length; }

function new_buf(sz) {
	var o = new_raw_buf(sz);
	prep_blob(o, 0);
	return o;
}

/* [MS-XLSB] 2.1.4 Record */
function recordhopper(data, cb, opts) {
	if(!data) return;
	var tmpbyte, cntbyte, length;
	prep_blob(data, data.l || 0);
	var L = data.length, RT = 0, tgt = 0;
	while(data.l < L) {
		RT = data.read_shift(1);
		if(RT & 0x80) RT = (RT & 0x7F) + ((data.read_shift(1) & 0x7F)<<7);
		var R = XLSBRecordEnum[RT] || XLSBRecordEnum[0xFFFF];
		tmpbyte = data.read_shift(1);
		length = tmpbyte & 0x7F;
		for(cntbyte = 1; cntbyte <4 && (tmpbyte & 0x80); ++cntbyte) length += ((tmpbyte = data.read_shift(1)) & 0x7F)<<(7*cntbyte);
		tgt = data.l + length;
		var d = (R.f||parsenoop)(data, length, opts);
		data.l = tgt;
		if(cb(d, R.n, RT)) return;
	}
}

/* control buffer usage for fixed-length buffers */
function buf_array() {
	var bufs = [], blksz = has_buf ? 256 : 2048;
	var newblk = function ba_newblk(sz) {
		var o = (new_buf(sz));
		prep_blob(o, 0);
		return o;
	};

	var curbuf = newblk(blksz);

	var endbuf = function ba_endbuf() {
		if(!curbuf) return;
		if(curbuf.length > curbuf.l) { curbuf = curbuf.slice(0, curbuf.l); curbuf.l = curbuf.length; }
		if(curbuf.length > 0) bufs.push(curbuf);
		curbuf = null;
	};

	var next = function ba_next(sz) {
		if(curbuf && (sz < (curbuf.length - curbuf.l))) return curbuf;
		endbuf();
		return (curbuf = newblk(Math.max(sz+1, blksz)));
	};

	var end = function ba_end() {
		endbuf();
		return __toBuffer([bufs]);
	};

	var push = function ba_push(buf) { endbuf(); curbuf = buf; if(curbuf.l == null) curbuf.l = curbuf.length; next(blksz); };

	return ({ next:next, push:push, end:end, _bufs:bufs });
}

function write_record(ba, type, payload, length) {
	var t = +XLSBRE[type], l;
	if(isNaN(t)) return; // TODO: throw something here?
	if(!length) length = XLSBRecordEnum[t].p || (payload||[]).length || 0;
	l = 1 + (t >= 0x80 ? 1 : 0) + 1/* + length*/;
	if(length >= 0x80) ++l; if(length >= 0x4000) ++l; if(length >= 0x200000) ++l;
	var o = ba.next(l);
	if(t <= 0x7F) o.write_shift(1, t);
	else {
		o.write_shift(1, (t & 0x7F) + 0x80);
		o.write_shift(1, (t >> 7));
	}
	for(var i = 0; i != 4; ++i) {
		if(length >= 0x80) { o.write_shift(1, (length & 0x7F)+0x80); length >>= 7; }
		else { o.write_shift(1, length); break; }
	}
	if(length > 0 && is_buf(payload)) ba.push(payload);
}
/* XLS ranges enforced */
function shift_cell_xls(cell, tgt, opts) {
	var out = dup(cell);
	if(tgt.s) {
		if(out.cRel) out.c += tgt.s.c;
		if(out.rRel) out.r += tgt.s.r;
	} else {
		if(out.cRel) out.c += tgt.c;
		if(out.rRel) out.r += tgt.r;
	}
	if(!opts || opts.biff < 12) {
		while(out.c >= 0x100) out.c -= 0x100;
		while(out.r >= 0x10000) out.r -= 0x10000;
	}
	return out;
}

function shift_range_xls(cell, range, opts) {
	var out = dup(cell);
	out.s = shift_cell_xls(out.s, range.s, opts);
	out.e = shift_cell_xls(out.e, range.s, opts);
	return out;
}

function encode_cell_xls(c, biff) {
	if(c.cRel && c.c < 0) { c = dup(c); c.c += (biff > 8) ? 0x4000 : 0x100; }
	if(c.rRel && c.r < 0) { c = dup(c); c.r += (biff > 8) ? 0x100000 : ((biff > 5) ? 0x10000 : 0x4000); }
	var s = encode_cell(c);
	if(c.cRel === 0) s = fix_col(s);
	if(c.rRel === 0) s = fix_row(s);
	return s;
}

function encode_range_xls(r, opts) {
	if(r.s.r == 0 && !r.s.rRel) {
		if(r.e.r == (opts.biff >= 12 ? 0xFFFFF : (opts.biff >= 8 ? 0x10000 : 0x4000)) && !r.e.rRel) {
			return (r.s.cRel ? "" : "$") + encode_col(r.s.c) + ":" + (r.e.cRel ? "" : "$") + encode_col(r.e.c);
		}
	}
	if(r.s.c == 0 && !r.s.cRel) {
		if(r.e.c == (opts.biff >= 12 ? 0xFFFF : 0xFF) && !r.e.cRel) {
			return (r.s.rRel ? "" : "$") + encode_row(r.s.r) + ":" + (r.e.rRel ? "" : "$") + encode_row(r.e.r);
		}
	}
	return encode_cell_xls(r.s, opts.biff) + ":" + encode_cell_xls(r.e, opts.biff);
}
var OFFCRYPTO = {};

var make_offcrypto = function(O, _crypto) {
	var crypto;
	if(typeof _crypto !== 'undefined') crypto = _crypto;
	else if(typeof require !== 'undefined') {
		try { crypto = undefined; }
		catch(e) { crypto = null; }
	}

	O.rc4 = function(key, data) {
		var S = new Array(256);
		var c = 0, i = 0, j = 0, t = 0;
		for(i = 0; i != 256; ++i) S[i] = i;
		for(i = 0; i != 256; ++i) {
			j = (j + S[i] + (key[i%key.length]).charCodeAt(0))&255;
			t = S[i]; S[i] = S[j]; S[j] = t;
		}
		// $FlowIgnore
		i = j = 0; var out = Buffer(data.length);
		for(c = 0; c != data.length; ++c) {
			i = (i + 1)&255;
			j = (j + S[i])%256;
			t = S[i]; S[i] = S[j]; S[j] = t;
			out[c] = (data[c] ^ S[(S[i]+S[j])&255]);
		}
		return out;
	};

	O.md5 = function(hex) {
		if(!crypto) throw new Error("Unsupported crypto");
		return crypto.createHash('md5').update(hex).digest('hex');
	};
};
/*global crypto:true */
make_offcrypto(OFFCRYPTO, typeof crypto !== "undefined" ? crypto : undefined);

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
function decode_range(range) { var x =range.split(":").map(decode_cell); return {s:x[0],e:x[x.length-1]}; }
function encode_range(cs,ce) {
	if(typeof ce === 'undefined' || typeof ce === 'number') {
return encode_range(cs.s, cs.e);
	}
if(typeof cs !== 'string') cs = encode_cell((cs));
	if(typeof ce !== 'string') ce = encode_cell((ce));
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
	var q = (cell.t == 'd' && v instanceof Date);
	if(cell.z != null) try { return (cell.w = SSF.format(cell.z, q ? datenum(v) : v)); } catch(e) { }
	try { return (cell.w = SSF.format((cell.XF||{}).numFmtId||(q ? 14 : 0),  q ? datenum(v) : v)); } catch(e) { return ''+v; }
}

function format_cell(cell, v, o) {
	if(cell == null || cell.t == null || cell.t == 'z') return "";
	if(cell.w !== undefined) return cell.w;
	if(cell.t == 'd' && !cell.z && o && o.dateNF) cell.z = o.dateNF;
	if(v == undefined) return safe_format_cell(cell, cell.v);
	return safe_format_cell(cell, v);
}

function sheet_to_workbook(sheet, opts) {
	var n = opts && opts.sheet ? opts.sheet : "Sheet1";
	var sheets = {}; sheets[n] = sheet;
	return { SheetNames: [n], Sheets: sheets };
}

function sheet_add_aoa(_ws, data, opts) {
	var o = opts || {};
	var dense = _ws ? Array.isArray(_ws) : o.dense;
	if(DENSE != null && dense == null) dense = DENSE;
	var ws = _ws || (dense ? ([]) : ({}));
	var _R = 0, _C = 0;
	if(ws && o.origin != null) {
		if(typeof o.origin == 'number') _R = o.origin;
		else {
			var _origin = typeof o.origin == "string" ? decode_cell(o.origin) : o.origin;
			_R = _origin.r; _C = _origin.c;
		}
	}
	var range = ({s: {c:10000000, r:10000000}, e: {c:0, r:0}});
	if(ws['!ref']) {
		var _range = safe_decode_range(ws['!ref']);
		range.s.c = _range.s.c;
		range.s.r = _range.s.r;
		range.e.c = Math.max(range.e.c, _range.e.c);
		range.e.r = Math.max(range.e.r, _range.e.r);
		if(_R == -1) range.e.r = _R = _range.e.r + 1;
	}
	for(var R = 0; R != data.length; ++R) {
		if(!data[R]) continue;
		if(!Array.isArray(data[R])) throw new Error("aoa_to_sheet expects an array of arrays");
		for(var C = 0; C != data[R].length; ++C) {
			if(typeof data[R][C] === 'undefined') continue;
			var cell = ({v: data[R][C] });
			var __R = _R + R, __C = _C + C;
			if(range.s.r > __R) range.s.r = __R;
			if(range.s.c > __C) range.s.c = __C;
			if(range.e.r < __R) range.e.r = __R;
			if(range.e.c < __C) range.e.c = __C;
			if(data[R][C] && typeof data[R][C] === 'object' && !Array.isArray(data[R][C]) && !(data[R][C] instanceof Date)) cell = data[R][C];
			else {
				if(Array.isArray(cell.v)) { cell.f = data[R][C][1]; cell.v = cell.v[0]; }
				if(cell.v === null) { if(cell.f) cell.t = 'n'; else if(!o.sheetStubs) continue; else cell.t = 'z'; }
				else if(typeof cell.v === 'number') cell.t = 'n';
				else if(typeof cell.v === 'boolean') cell.t = 'b';
				else if(cell.v instanceof Date) {
					cell.z = o.dateNF || SSF._table[14];
					if(o.cellDates) { cell.t = 'd'; cell.w = SSF.format(cell.z, datenum(cell.v)); }
					else { cell.t = 'n'; cell.v = datenum(cell.v); cell.w = SSF.format(cell.z, cell.v); }
				}
				else cell.t = 's';
			}
			if(dense) {
				if(!ws[__R]) ws[__R] = [];
				ws[__R][__C] = cell;
			} else {
				var cell_ref = encode_cell(({c:__C,r:__R}));
				ws[cell_ref] = cell;
			}
		}
	}
	if(range.s.c < 10000000) ws['!ref'] = encode_range(range);
	return ws;
}
function aoa_to_sheet(data, opts) { return sheet_add_aoa(null, data, opts); }

function write_UInt32LE(x, o) {
	if(!o) o = new_buf(4);
	o.write_shift(4, x);
	return o;
}

/* [MS-XLSB] 2.5.168 */
function parse_XLWideString(data) {
	var cchCharacters = data.read_shift(4);
	return cchCharacters === 0 ? "" : data.read_shift(cchCharacters, 'dbcs');
}
function write_XLWideString(data, o) {
	var _null = false; if(o == null) { _null = true; o = new_buf(4+2*data.length); }
	o.write_shift(4, data.length);
	if(data.length > 0) o.write_shift(0, data, 'dbcs');
	return _null ? o.slice(0, o.l) : o;
}

/* [MS-XLSB] 2.5.143 */
function parse_StrRun(data) {
	return { ich: data.read_shift(2), ifnt: data.read_shift(2) };
}
function write_StrRun(run, o) {
	if(!o) o = new_buf(4);
	o.write_shift(2, run.ich || 0);
	o.write_shift(2, run.ifnt || 0);
	return o;
}

/* [MS-XLSB] 2.5.121 */
function parse_RichStr(data, length) {
	var start = data.l;
	var flags = data.read_shift(1);
	var str = parse_XLWideString(data);
	var rgsStrRun = [];
	var z = ({ t: str, h: str });
	if((flags & 1) !== 0) { /* fRichStr */
		/* TODO: formatted string */
		var dwSizeStrRun = data.read_shift(4);
		for(var i = 0; i != dwSizeStrRun; ++i) rgsStrRun.push(parse_StrRun(data));
		z.r = rgsStrRun;
	}
	else z.r = [{ich:0, ifnt:0}];
	//if((flags & 2) !== 0) { /* fExtStr */
	//	/* TODO: phonetic string */
	//}
	data.l = start + length;
	return z;
}
function write_RichStr(str, o) {
	/* TODO: formatted string */
	var _null = false; if(o == null) { _null = true; o = new_buf(15+4*str.t.length); }
	o.write_shift(1,0);
	write_XLWideString(str.t, o);
	return _null ? o.slice(0, o.l) : o;
}
/* [MS-XLSB] 2.4.328 BrtCommentText (RichStr w/1 run) */
var parse_BrtCommentText = parse_RichStr;
function write_BrtCommentText(str, o) {
	/* TODO: formatted string */
	var _null = false; if(o == null) { _null = true; o = new_buf(23+4*str.t.length); }
	o.write_shift(1,1);
	write_XLWideString(str.t, o);
	o.write_shift(4,1);
	write_StrRun({ich:0,ifnt:0}, o);
	return _null ? o.slice(0, o.l) : o;
}

/* [MS-XLSB] 2.5.9 */
function parse_XLSBCell(data) {
	var col = data.read_shift(4);
	var iStyleRef = data.read_shift(2);
	iStyleRef += data.read_shift(1) <<16;
	data.l++; //var fPhShow = data.read_shift(1);
	return { c:col, iStyleRef: iStyleRef };
}
function write_XLSBCell(cell, o) {
	if(o == null) o = new_buf(8);
	o.write_shift(-4, cell.c);
	o.write_shift(3, cell.iStyleRef || cell.s);
	o.write_shift(1, 0); /* fPhShow */
	return o;
}


/* [MS-XLSB] 2.5.21 */
var parse_XLSBCodeName = parse_XLWideString;
var write_XLSBCodeName = write_XLWideString;

/* [MS-XLSB] 2.5.166 */
function parse_XLNullableWideString(data) {
	var cchCharacters = data.read_shift(4);
	return cchCharacters === 0 || cchCharacters === 0xFFFFFFFF ? "" : data.read_shift(cchCharacters, 'dbcs');
}
function write_XLNullableWideString(data, o) {
	var _null = false; if(o == null) { _null = true; o = new_buf(127); }
	o.write_shift(4, data.length > 0 ? data.length : 0xFFFFFFFF);
	if(data.length > 0) o.write_shift(0, data, 'dbcs');
	return _null ? o.slice(0, o.l) : o;
}

/* [MS-XLSB] 2.5.165 */
var parse_XLNameWideString = parse_XLWideString;
//var write_XLNameWideString = write_XLWideString;

/* [MS-XLSB] 2.5.114 */
var parse_RelID = parse_XLNullableWideString;
var write_RelID = write_XLNullableWideString;


/* [MS-XLS] 2.5.217 ; [MS-XLSB] 2.5.122 */
function parse_RkNumber(data) {
	var b = data.slice(data.l, data.l+4);
	var fX100 = (b[0] & 1), fInt = (b[0] & 2);
	data.l+=4;
	b[0] &= 0xFC; // b[0] &= ~3;
	var RK = fInt === 0 ? __double([0,0,0,0,b[0],b[1],b[2],b[3]],0) : __readInt32LE(b,0)>>2;
	return fX100 ? (RK/100) : RK;
}
function write_RkNumber(data, o) {
	if(o == null) o = new_buf(4);
	var fX100 = 0, fInt = 0, d100 = data * 100;
	if((data == (data | 0)) && (data >= -(1<<29)) && (data < (1 << 29))) { fInt = 1; }
	else if((d100 == (d100 | 0)) && (d100 >= -(1<<29)) && (d100 < (1 << 29))) { fInt = 1; fX100 = 1; }
	if(fInt) o.write_shift(-4, ((fX100 ? d100 : data) << 2) + (fX100 + 2));
	else throw new Error("unsupported RkNumber " + data); // TODO
}


/* [MS-XLSB] 2.5.117 RfX */
function parse_RfX(data ) {
	var cell = ({s: {}, e: {}});
	cell.s.r = data.read_shift(4);
	cell.e.r = data.read_shift(4);
	cell.s.c = data.read_shift(4);
	cell.e.c = data.read_shift(4);
	return cell;
}
function write_RfX(r, o) {
	if(!o) o = new_buf(16);
	o.write_shift(4, r.s.r);
	o.write_shift(4, r.e.r);
	o.write_shift(4, r.s.c);
	o.write_shift(4, r.e.c);
	return o;
}

/* [MS-XLSB] 2.5.153 UncheckedRfX */
var parse_UncheckedRfX = parse_RfX;
var write_UncheckedRfX = write_RfX;

/* [MS-XLS] 2.5.342 ; [MS-XLSB] 2.5.171 */
/* TODO: error checking, NaN and Infinity values are not valid Xnum */
function parse_Xnum(data) { return data.read_shift(8, 'f'); }
function write_Xnum(data, o) { return (o || new_buf(8)).write_shift(8, data, 'f'); }

/* [MS-XLSB] 2.5.97.2 */
var BErr = {
0x00: "#NULL!",
0x07: "#DIV/0!",
0x0F: "#VALUE!",
0x17: "#REF!",
0x1D: "#NAME?",
0x24: "#NUM!",
0x2A: "#N/A",
0x2B: "#GETTING_DATA",
0xFF: "#WTF?"
};
var RBErr = evert_num(BErr);

/* [MS-XLSB] 2.4.324 BrtColor */
function parse_BrtColor(data) {
	var out = {};
	var d = data.read_shift(1);

	//var fValidRGB = d & 1;
	var xColorType = d >>> 1;

	var index = data.read_shift(1);
	var nTS = data.read_shift(2, 'i');
	var bR = data.read_shift(1);
	var bG = data.read_shift(1);
	var bB = data.read_shift(1);
	data.l++; //var bAlpha = data.read_shift(1);

	switch(xColorType) {
		case 0: out.auto = 1; break;
		case 1:
			out.index = index;
			var icv = XLSIcv[index];
			/* automatic pseudo index 81 */
			if(icv) out.rgb = rgb2Hex(icv);
			break;
		case 2:
			/* if(!fValidRGB) throw new Error("invalid"); */
			out.rgb = rgb2Hex([bR, bG, bB]);
			break;
		case 3: out.theme = index; break;
	}
	if(nTS != 0) out.tint = nTS > 0 ? nTS / 32767 : nTS / 32768;

	return out;
}
function write_BrtColor(color, o) {
	if(!o) o = new_buf(8);
	if(!color||color.auto) { o.write_shift(4, 0); o.write_shift(4, 0); return o; }
	if(color.index) {
		o.write_shift(1, 0x02);
		o.write_shift(1, color.index);
	} else if(color.theme) {
		o.write_shift(1, 0x06);
		o.write_shift(1, color.theme);
	} else {
		o.write_shift(1, 0x05);
		o.write_shift(1, 0);
	}
	var nTS = color.tint || 0;
	if(nTS > 0) nTS *= 32767;
	else if(nTS < 0) nTS *= 32768;
	o.write_shift(2, nTS);
	if(!color.rgb) {
		o.write_shift(2, 0);
		o.write_shift(1, 0);
		o.write_shift(1, 0);
	} else {
		var rgb = (color.rgb || 'FFFFFF');
		o.write_shift(1, parseInt(rgb.slice(0,2),16));
		o.write_shift(1, parseInt(rgb.slice(2,4),16));
		o.write_shift(1, parseInt(rgb.slice(4,6),16));
		o.write_shift(1, 0xFF);
	}
	return o;
}

/* [MS-XLSB] 2.5.52 */
function parse_FontFlags(data) {
	var d = data.read_shift(1);
	data.l++;
	var out = {
		/* fBold: d & 0x01 */
		fItalic: d & 0x02,
		/* fUnderline: d & 0x04 */
		fStrikeout: d & 0x08,
		fOutline: d & 0x10,
		fShadow: d & 0x20,
		fCondense: d & 0x40,
		fExtend: d & 0x80
	};
	return out;
}
function write_FontFlags(font, o) {
	if(!o) o = new_buf(2);
	var grbit =
		(font.italic   ? 0x02 : 0) |
		(font.strike   ? 0x08 : 0) |
		(font.outline  ? 0x10 : 0) |
		(font.shadow   ? 0x20 : 0) |
		(font.condense ? 0x40 : 0) |
		(font.extend   ? 0x80 : 0);
	o.write_shift(1, grbit);
	o.write_shift(1, 0);
	return o;
}

/* [MS-OLEDS] 2.3.1 and 2.3.2 */
function parse_ClipboardFormatOrString(o, w) {
	// $FlowIgnore
	var ClipFmt = {2:"BITMAP",3:"METAFILEPICT",8:"DIB",14:"ENHMETAFILE"};
	var m = o.read_shift(4);
	switch(m) {
		case 0x00000000: return "";
		case 0xffffffff: case 0xfffffffe: return ClipFmt[o.read_shift(4)]||"";
	}
	if(m > 0x190) throw new Error("Unsupported Clipboard: " + m.toString(16));
	o.l -= 4;
	return o.read_shift(0, w == 1 ? "lpstr" : "lpwstr");
}
function parse_ClipboardFormatOrAnsiString(o) { return parse_ClipboardFormatOrString(o, 1); }
function parse_ClipboardFormatOrUnicodeString(o) { return parse_ClipboardFormatOrString(o, 2); }

/* [MS-OLEPS] 2.2 PropertyType */
//var VT_EMPTY    = 0x0000;
//var VT_NULL     = 0x0001;
var VT_I2       = 0x0002;
var VT_I4       = 0x0003;
//var VT_R4       = 0x0004;
//var VT_R8       = 0x0005;
//var VT_CY       = 0x0006;
//var VT_DATE     = 0x0007;
//var VT_BSTR     = 0x0008;
//var VT_ERROR    = 0x000A;
var VT_BOOL     = 0x000B;
var VT_VARIANT  = 0x000C;
//var VT_DECIMAL  = 0x000E;
//var VT_I1       = 0x0010;
//var VT_UI1      = 0x0011;
//var VT_UI2      = 0x0012;
var VT_UI4      = 0x0013;
//var VT_I8       = 0x0014;
//var VT_UI8      = 0x0015;
//var VT_INT      = 0x0016;
//var VT_UINT     = 0x0017;
var VT_LPSTR    = 0x001E;
//var VT_LPWSTR   = 0x001F;
var VT_FILETIME = 0x0040;
var VT_BLOB     = 0x0041;
//var VT_STREAM   = 0x0042;
//var VT_STORAGE  = 0x0043;
//var VT_STREAMED_Object  = 0x0044;
//var VT_STORED_Object    = 0x0045;
//var VT_BLOB_Object      = 0x0046;
var VT_CF       = 0x0047;
//var VT_CLSID    = 0x0048;
//var VT_VERSIONED_STREAM = 0x0049;
var VT_VECTOR   = 0x1000;
//var VT_ARRAY    = 0x2000;

var VT_STRING   = 0x0050; // 2.3.3.1.11 VtString
var VT_USTR     = 0x0051; // 2.3.3.1.12 VtUnalignedString
var VT_CUSTOM   = [VT_STRING, VT_USTR];

/* [MS-OSHARED] 2.3.3.2.2.1 Document Summary Information PIDDSI */
var DocSummaryPIDDSI = {
0x01: { n: 'CodePage', t: VT_I2 },
0x02: { n: 'Category', t: VT_STRING },
0x03: { n: 'PresentationFormat', t: VT_STRING },
0x04: { n: 'ByteCount', t: VT_I4 },
0x05: { n: 'LineCount', t: VT_I4 },
0x06: { n: 'ParagraphCount', t: VT_I4 },
0x07: { n: 'SlideCount', t: VT_I4 },
0x08: { n: 'NoteCount', t: VT_I4 },
0x09: { n: 'HiddenCount', t: VT_I4 },
0x0a: { n: 'MultimediaClipCount', t: VT_I4 },
0x0b: { n: 'ScaleCrop', t: VT_BOOL },
0x0c: { n: 'HeadingPairs', t: VT_VECTOR | VT_VARIANT },
0x0d: { n: 'TitlesOfParts', t: VT_VECTOR | VT_LPSTR },
0x0e: { n: 'Manager', t: VT_STRING },
0x0f: { n: 'Company', t: VT_STRING },
0x10: { n: 'LinksUpToDate', t: VT_BOOL },
0x11: { n: 'CharacterCount', t: VT_I4 },
0x13: { n: 'SharedDoc', t: VT_BOOL },
0x16: { n: 'HyperlinksChanged', t: VT_BOOL },
0x17: { n: 'AppVersion', t: VT_I4, p: 'version' },
0x18: { n: 'DigSig', t: VT_BLOB },
0x1A: { n: 'ContentType', t: VT_STRING },
0x1B: { n: 'ContentStatus', t: VT_STRING },
0x1C: { n: 'Language', t: VT_STRING },
0x1D: { n: 'Version', t: VT_STRING },
0xFF: {}
};

/* [MS-OSHARED] 2.3.3.2.1.1 Summary Information Property Set PIDSI */
var SummaryPIDSI = {
0x01: { n: 'CodePage', t: VT_I2 },
0x02: { n: 'Title', t: VT_STRING },
0x03: { n: 'Subject', t: VT_STRING },
0x04: { n: 'Author', t: VT_STRING },
0x05: { n: 'Keywords', t: VT_STRING },
0x06: { n: 'Comments', t: VT_STRING },
0x07: { n: 'Template', t: VT_STRING },
0x08: { n: 'LastAuthor', t: VT_STRING },
0x09: { n: 'RevNumber', t: VT_STRING },
0x0A: { n: 'EditTime', t: VT_FILETIME },
0x0B: { n: 'LastPrinted', t: VT_FILETIME },
0x0C: { n: 'CreatedDate', t: VT_FILETIME },
0x0D: { n: 'ModifiedDate', t: VT_FILETIME },
0x0E: { n: 'PageCount', t: VT_I4 },
0x0F: { n: 'WordCount', t: VT_I4 },
0x10: { n: 'CharCount', t: VT_I4 },
0x11: { n: 'Thumbnail', t: VT_CF },
0x12: { n: 'Application', t: VT_STRING },
0x13: { n: 'DocSecurity', t: VT_I4 },
0xFF: {}
};

/* [MS-OLEPS] 2.18 */
var SpecialProperties = {
0x80000000: { n: 'Locale', t: VT_UI4 },
0x80000003: { n: 'Behavior', t: VT_UI4 },
0x72627262: {}
};

(function() {
	for(var y in SpecialProperties) if(SpecialProperties.hasOwnProperty(y))
	DocSummaryPIDDSI[y] = SummaryPIDSI[y] = SpecialProperties[y];
})();

var DocSummaryRE = evert_key(DocSummaryPIDDSI, "n");
var SummaryRE = evert_key(SummaryPIDSI, "n");

/* [MS-XLS] 2.4.63 Country/Region codes */
var CountryEnum = {
0x0001: "US", // United States
0x0002: "CA", // Canada
0x0003: "", // Latin America (except Brazil)
0x0007: "RU", // Russia
0x0014: "EG", // Egypt
0x001E: "GR", // Greece
0x001F: "NL", // Netherlands
0x0020: "BE", // Belgium
0x0021: "FR", // France
0x0022: "ES", // Spain
0x0024: "HU", // Hungary
0x0027: "IT", // Italy
0x0029: "CH", // Switzerland
0x002B: "AT", // Austria
0x002C: "GB", // United Kingdom
0x002D: "DK", // Denmark
0x002E: "SE", // Sweden
0x002F: "NO", // Norway
0x0030: "PL", // Poland
0x0031: "DE", // Germany
0x0034: "MX", // Mexico
0x0037: "BR", // Brazil
0x003d: "AU", // Australia
0x0040: "NZ", // New Zealand
0x0042: "TH", // Thailand
0x0051: "JP", // Japan
0x0052: "KR", // Korea
0x0054: "VN", // Viet Nam
0x0056: "CN", // China
0x005A: "TR", // Turkey
0x0069: "JS", // Ramastan
0x00D5: "DZ", // Algeria
0x00D8: "MA", // Morocco
0x00DA: "LY", // Libya
0x015F: "PT", // Portugal
0x0162: "IS", // Iceland
0x0166: "FI", // Finland
0x01A4: "CZ", // Czech Republic
0x0376: "TW", // Taiwan
0x03C1: "LB", // Lebanon
0x03C2: "JO", // Jordan
0x03C3: "SY", // Syria
0x03C4: "IQ", // Iraq
0x03C5: "KW", // Kuwait
0x03C6: "SA", // Saudi Arabia
0x03CB: "AE", // United Arab Emirates
0x03CC: "IL", // Israel
0x03CE: "QA", // Qatar
0x03D5: "IR", // Iran
0xFFFF: "US"  // United States
};

/* [MS-XLS] 2.5.127 */
var XLSFillPattern = [
	null,
	'solid',
	'mediumGray',
	'darkGray',
	'lightGray',
	'darkHorizontal',
	'darkVertical',
	'darkDown',
	'darkUp',
	'darkGrid',
	'darkTrellis',
	'lightHorizontal',
	'lightVertical',
	'lightDown',
	'lightUp',
	'lightGrid',
	'lightTrellis',
	'gray125',
	'gray0625'
];

function rgbify(arr) { return arr.map(function(x) { return [(x>>16)&255,(x>>8)&255,x&255]; }); }

/* [MS-XLS] 2.5.161 */
/* [MS-XLSB] 2.5.75 Icv */
var XLSIcv = rgbify([
	/* Color Constants */
	0x000000,
	0xFFFFFF,
	0xFF0000,
	0x00FF00,
	0x0000FF,
	0xFFFF00,
	0xFF00FF,
	0x00FFFF,

	/* Overridable Defaults */
	0x000000,
	0xFFFFFF,
	0xFF0000,
	0x00FF00,
	0x0000FF,
	0xFFFF00,
	0xFF00FF,
	0x00FFFF,

	0x800000,
	0x008000,
	0x000080,
	0x808000,
	0x800080,
	0x008080,
	0xC0C0C0,
	0x808080,
	0x9999FF,
	0x993366,
	0xFFFFCC,
	0xCCFFFF,
	0x660066,
	0xFF8080,
	0x0066CC,
	0xCCCCFF,

	0x000080,
	0xFF00FF,
	0xFFFF00,
	0x00FFFF,
	0x800080,
	0x800000,
	0x008080,
	0x0000FF,
	0x00CCFF,
	0xCCFFFF,
	0xCCFFCC,
	0xFFFF99,
	0x99CCFF,
	0xFF99CC,
	0xCC99FF,
	0xFFCC99,

	0x3366FF,
	0x33CCCC,
	0x99CC00,
	0xFFCC00,
	0xFF9900,
	0xFF6600,
	0x666699,
	0x969696,
	0x003366,
	0x339966,
	0x003300,
	0x333300,
	0x993300,
	0x993366,
	0x333399,
	0x333333,

	/* Other entries to appease BIFF8/12 */
	0xFFFFFF, /* 0x40 icvForeground ?? */
	0x000000, /* 0x41 icvBackground ?? */
	0x000000, /* 0x42 icvFrame ?? */
	0x000000, /* 0x43 icv3D ?? */
	0x000000, /* 0x44 icv3DText ?? */
	0x000000, /* 0x45 icv3DHilite ?? */
	0x000000, /* 0x46 icv3DShadow ?? */
	0x000000, /* 0x47 icvHilite ?? */
	0x000000, /* 0x48 icvCtlText ?? */
	0x000000, /* 0x49 icvCtlScrl ?? */
	0x000000, /* 0x4A icvCtlInv ?? */
	0x000000, /* 0x4B icvCtlBody ?? */
	0x000000, /* 0x4C icvCtlFrame ?? */
	0x000000, /* 0x4D icvCtlFore ?? */
	0x000000, /* 0x4E icvCtlBack ?? */
	0x000000, /* 0x4F icvCtlNeutral */
	0x000000, /* 0x50 icvInfoBk ?? */
	0x000000 /* 0x51 icvInfoText ?? */
]);

/* Parts enumerated in OPC spec, MS-XLSB and MS-XLSX */
/* 12.3 Part Summary <SpreadsheetML> */
/* 14.2 Part Summary <DrawingML> */
/* [MS-XLSX] 2.1 Part Enumerations ; [MS-XLSB] 2.1.7 Part Enumeration */
var ct2type/*{[string]:string}*/ = ({
	/* Workbook */
	"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml": "workbooks",

	/* Worksheet */
	"application/vnd.ms-excel.binIndexWs": "TODO", /* Binary Index */

	/* Macrosheet */
	"application/vnd.ms-excel.intlmacrosheet": "TODO",
	"application/vnd.ms-excel.binIndexMs": "TODO", /* Binary Index */

	/* File Properties */
	"application/vnd.openxmlformats-package.core-properties+xml": "coreprops",
	"application/vnd.openxmlformats-officedocument.custom-properties+xml": "custprops",
	"application/vnd.openxmlformats-officedocument.extended-properties+xml": "extprops",

	/* Custom Data Properties */
	"application/vnd.openxmlformats-officedocument.customXmlProperties+xml": "TODO",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.customProperty": "TODO",

	/* PivotTable */
	"application/vnd.ms-excel.pivotTable": "TODO",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml": "TODO",

	/* Chart Colors */
	"application/vnd.ms-office.chartcolorstyle+xml": "TODO",

	/* Chart Style */
	"application/vnd.ms-office.chartstyle+xml": "TODO",

	/* Calculation Chain */
	"application/vnd.ms-excel.calcChain": "calcchains",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml": "calcchains",

	/* Printer Settings */
	"application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings": "TODO",

	/* ActiveX */
	"application/vnd.ms-office.activeX": "TODO",
	"application/vnd.ms-office.activeX+xml": "TODO",

	/* Custom Toolbars */
	"application/vnd.ms-excel.attachedToolbars": "TODO",

	/* External Data Connections */
	"application/vnd.ms-excel.connections": "TODO",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml": "TODO",

	/* External Links */
	"application/vnd.ms-excel.externalLink": "links",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml": "links",

	/* Metadata */
	"application/vnd.ms-excel.sheetMetadata": "TODO",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml": "TODO",

	/* PivotCache */
	"application/vnd.ms-excel.pivotCacheDefinition": "TODO",
	"application/vnd.ms-excel.pivotCacheRecords": "TODO",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml": "TODO",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml": "TODO",

	/* Query Table */
	"application/vnd.ms-excel.queryTable": "TODO",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml": "TODO",

	/* Shared Workbook */
	"application/vnd.ms-excel.userNames": "TODO",
	"application/vnd.ms-excel.revisionHeaders": "TODO",
	"application/vnd.ms-excel.revisionLog": "TODO",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.revisionHeaders+xml": "TODO",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.revisionLog+xml": "TODO",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.userNames+xml": "TODO",

	/* Single Cell Table */
	"application/vnd.ms-excel.tableSingleCells": "TODO",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.tableSingleCells+xml": "TODO",

	/* Slicer */
	"application/vnd.ms-excel.slicer": "TODO",
	"application/vnd.ms-excel.slicerCache": "TODO",
	"application/vnd.ms-excel.slicer+xml": "TODO",
	"application/vnd.ms-excel.slicerCache+xml": "TODO",

	/* Sort Map */
	"application/vnd.ms-excel.wsSortMap": "TODO",

	/* Table */
	"application/vnd.ms-excel.table": "TODO",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml": "TODO",

	/* Themes */
	"application/vnd.openxmlformats-officedocument.theme+xml": "themes",

	/* Theme Override */
	"application/vnd.openxmlformats-officedocument.themeOverride+xml": "TODO",

	/* Timeline */
	"application/vnd.ms-excel.Timeline+xml": "TODO", /* verify */
	"application/vnd.ms-excel.TimelineCache+xml": "TODO", /* verify */

	/* VBA */
	"application/vnd.ms-office.vbaProject": "vba",
	"application/vnd.ms-office.vbaProjectSignature": "vba",

	/* Volatile Dependencies */
	"application/vnd.ms-office.volatileDependencies": "TODO",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.volatileDependencies+xml": "TODO",

	/* Control Properties */
	"application/vnd.ms-excel.controlproperties+xml": "TODO",

	/* Data Model */
	"application/vnd.openxmlformats-officedocument.model+data": "TODO",

	/* Survey */
	"application/vnd.ms-excel.Survey+xml": "TODO",

	/* Drawing */
	"application/vnd.openxmlformats-officedocument.drawing+xml": "drawings",
	"application/vnd.openxmlformats-officedocument.drawingml.chart+xml": "TODO",
	"application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml": "TODO",
	"application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml": "TODO",
	"application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml": "TODO",
	"application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml": "TODO",
	"application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml": "TODO",

	/* VML */
	"application/vnd.openxmlformats-officedocument.vmlDrawing": "TODO",

	"application/vnd.openxmlformats-package.relationships+xml": "rels",
	"application/vnd.openxmlformats-officedocument.oleObject": "TODO",

	/* Image */
	"image/png": "TODO",

	"sheet": "js"
});

var CT_LIST = (function(){
	var o = {
		workbooks: {
			xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
			xlsm: "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
			xlsb: "application/vnd.ms-excel.sheet.binary.macroEnabled.main",
			xlam: "application/vnd.ms-excel.addin.macroEnabled.main+xml",
			xltx: "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"
		},
		strs: { /* Shared Strings */
			xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
			xlsb: "application/vnd.ms-excel.sharedStrings"
		},
		comments: { /* Comments */
			xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
			xlsb: "application/vnd.ms-excel.comments"
		},
		sheets: { /* Worksheet */
			xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
			xlsb: "application/vnd.ms-excel.worksheet"
		},
		charts: { /* Chartsheet */
			xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml",
			xlsb: "application/vnd.ms-excel.chartsheet"
		},
		dialogs: { /* Dialogsheet */
			xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml",
			xlsb: "application/vnd.ms-excel.dialogsheet"
		},
		macros: { /* Macrosheet (Excel 4.0 Macros) */
			xlsx: "application/vnd.ms-excel.macrosheet+xml",
			xlsb: "application/vnd.ms-excel.macrosheet"
		},
		styles: { /* Styles */
			xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
			xlsb: "application/vnd.ms-excel.styles"
		}
	};
	keys(o).forEach(function(k) { ["xlsm", "xlam"].forEach(function(v) { if(!o[k][v]) o[k][v] = o[k].xlsx; }); });
	keys(o).forEach(function(k){ keys(o[k]).forEach(function(v) { ct2type[o[k][v]] = k; }); });
	return o;
})();

var type2ct/*{[string]:Array<string>}*/ = evert_arr(ct2type);

XMLNS.CT = 'http://schemas.openxmlformats.org/package/2006/content-types';

function new_ct() {
	return ({
		workbooks:[], sheets:[], charts:[], dialogs:[], macros:[],
		rels:[], strs:[], comments:[], links:[],
		coreprops:[], extprops:[], custprops:[], themes:[], styles:[],
		calcchains:[], vba: [], drawings: [],
		TODO:[], xmlns: "" });
}

function parse_ct(data) {
	var ct = new_ct();
	if(!data || !data.match) return ct;
	var ctext = {};
	(data.match(tagregex)||[]).forEach(function(x) {
		var y = parsexmltag(x);
		switch(y[0].replace(nsregex,"<")) {
			case '<?xml': break;
			case '<Types': ct.xmlns = y['xmlns' + (y[0].match(/<(\w+):/)||["",""])[1] ]; break;
			case '<Default': ctext[y.Extension] = y.ContentType; break;
			case '<Override':
				if(ct[ct2type[y.ContentType]] !== undefined) ct[ct2type[y.ContentType]].push(y.PartName);
				break;
		}
	});
	if(ct.xmlns !== XMLNS.CT) throw new Error("Unknown Namespace: " + ct.xmlns);
	ct.calcchain = ct.calcchains.length > 0 ? ct.calcchains[0] : "";
	ct.sst = ct.strs.length > 0 ? ct.strs[0] : "";
	ct.style = ct.styles.length > 0 ? ct.styles[0] : "";
	ct.defaults = ctext;
	delete ct.calcchains;
	return ct;
}

var CTYPE_XML_ROOT = writextag('Types', null, {
	'xmlns': XMLNS.CT,
	'xmlns:xsd': XMLNS.xsd,
	'xmlns:xsi': XMLNS.xsi
});

var CTYPE_DEFAULTS = [
	['xml', 'application/xml'],
	['bin', 'application/vnd.ms-excel.sheet.binary.macroEnabled.main'],
	['vml', 'application/vnd.openxmlformats-officedocument.vmlDrawing'],
	/* from test files */
	['bmp', 'image/bmp'],
	['png', 'image/png'],
	['gif', 'image/gif'],
	['emf', 'image/x-emf'],
	['wmf', 'image/x-wmf'],
	['jpg', 'image/jpeg'], ['jpeg', 'image/jpeg'],
	['tif', 'image/tiff'], ['tiff', 'image/tiff'],
	['pdf', 'application/pdf'],
	['rels', type2ct.rels[0]]
].map(function(x) {
	return writextag('Default', null, {'Extension':x[0], 'ContentType': x[1]});
});

function write_ct(ct, opts) {
	var o = [], v;
	o[o.length] = (XML_HEADER);
	o[o.length] = (CTYPE_XML_ROOT);
	o = o.concat(CTYPE_DEFAULTS);
	var f1 = function(w) {
		if(ct[w] && ct[w].length > 0) {
			v = ct[w][0];
			o[o.length] = (writextag('Override', null, {
				'PartName': (v[0] == '/' ? "":"/") + v,
				'ContentType': CT_LIST[w][opts.bookType || 'xlsx']
			}));
		}
	};
	var f2 = function(w) {
		(ct[w]||[]).forEach(function(v) {
			o[o.length] = (writextag('Override', null, {
				'PartName': (v[0] == '/' ? "":"/") + v,
				'ContentType': CT_LIST[w][opts.bookType || 'xlsx']
			}));
		});
	};
	var f3 = function(t) {
		(ct[t]||[]).forEach(function(v) {
			o[o.length] = (writextag('Override', null, {
				'PartName': (v[0] == '/' ? "":"/") + v,
				'ContentType': type2ct[t][0]
			}));
		});
	};
	f1('workbooks');
	f2('sheets');
	f2('charts');
	f3('themes');
	['strs', 'styles'].forEach(f1);
	['coreprops', 'extprops', 'custprops'].forEach(f3);
	f3('vba');
	f3('comments');
	f3('drawings');
	if(o.length>2){ o[o.length] = ('</Types>'); o[1]=o[1].replace("/>",">"); }
	return o.join("");
}
/* 9.3 Relationships */
var RELS = ({
	WB: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
	SHEET: "http://sheetjs.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
	HLINK: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
	VML: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
	VBA: "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
});

/* 9.3.3 Representing Relationships */
function get_rels_path(file) {
	var n = file.lastIndexOf("/");
	return file.slice(0,n+1) + '_rels/' + file.slice(n+1) + ".rels";
}

function parse_rels(data, currentFilePath) {
	if (!data) return data;
	if (currentFilePath.charAt(0) !== '/') {
		currentFilePath = '/'+currentFilePath;
	}
	var rels = {};
	var hash = {};

	(data.match(tagregex)||[]).forEach(function(x) {
		var y = parsexmltag(x);
		/* 9.3.2.2 OPC_Relationships */
		if (y[0] === '<Relationship') {
			var rel = {}; rel.Type = y.Type; rel.Target = y.Target; rel.Id = y.Id; rel.TargetMode = y.TargetMode;
			var canonictarget = y.TargetMode === 'External' ? y.Target : resolve_path(y.Target, currentFilePath);
			rels[canonictarget] = rel;
			hash[y.Id] = rel;
		}
	});
	rels["!id"] = hash;
	return rels;
}

XMLNS.RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';

var RELS_ROOT = writextag('Relationships', null, {
	//'xmlns:ns0': XMLNS.RELS,
	'xmlns': XMLNS.RELS
});

/* TODO */
function write_rels(rels) {
	var o = [XML_HEADER, RELS_ROOT];
	keys(rels['!id']).forEach(function(rid) {
		o[o.length] = (writextag('Relationship', null, rels['!id'][rid]));
	});
	if(o.length>2){ o[o.length] = ('</Relationships>'); o[1]=o[1].replace("/>",">"); }
	return o.join("");
}

function add_rels(rels, rId, f, type, relobj) {
	if(!relobj) relobj = {};
	if(!rels['!id']) rels['!id'] = {};
	if(rId < 0) for(rId = 1; rels['!id']['rId' + rId]; ++rId){/* empty */}
	relobj.Id = 'rId' + rId;
	relobj.Type = type;
	relobj.Target = f;
	if(relobj.Type == RELS.HLINK) relobj.TargetMode = "External";
	if(rels['!id'][relobj.Id]) throw new Error("Cannot rewrite rId " + rId);
	rels['!id'][relobj.Id] = relobj;
	rels[('/' + relobj.Target).replace("//","/")] = relobj;
	return rId;
}
/* Open Document Format for Office Applications (OpenDocument) Version 1.2 */
/* Part 3 Section 4 Manifest File */
var CT_ODS = "application/vnd.oasis.opendocument.spreadsheet";
function parse_manifest(d, opts) {
	var str = xlml_normalize(d);
	var Rn;
	var FEtag;
	while((Rn = xlmlregex.exec(str))) switch(Rn[3]) {
		case 'manifest': break; // 4.2 <manifest:manifest>
		case 'file-entry': // 4.3 <manifest:file-entry>
			FEtag = parsexmltag(Rn[0], false);
			if(FEtag.path == '/' && FEtag.type !== CT_ODS) throw new Error("This OpenDocument is not a spreadsheet");
			break;
		case 'encryption-data': // 4.4 <manifest:encryption-data>
		case 'algorithm': // 4.5 <manifest:algorithm>
		case 'start-key-generation': // 4.6 <manifest:start-key-generation>
		case 'key-derivation': // 4.7 <manifest:key-derivation>
			throw new Error("Unsupported ODS Encryption");
		default: if(opts && opts.WTF) throw Rn;
	}
}

function write_manifest(manifest) {
	var o = [XML_HEADER];
	o.push('<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">\n');
	o.push('  <manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>\n');
	for(var i = 0; i < manifest.length; ++i) o.push('  <manifest:file-entry manifest:full-path="' + manifest[i][0] + '" manifest:media-type="' + manifest[i][1] + '"/>\n');
	o.push('</manifest:manifest>');
	return o.join("");
}

/* Part 3 Section 6 Metadata Manifest File */
function write_rdf_type(file, res, tag) {
	return [
		'  <rdf:Description rdf:about="' + file + '">\n',
		'    <rdf:type rdf:resource="http://docs.oasis-open.org/ns/office/1.2/meta/' + (tag || "odf") + '#' + res + '"/>\n',
		'  </rdf:Description>\n'
	].join("");
}
function write_rdf_has(base, file) {
	return [
		'  <rdf:Description rdf:about="' + base + '">\n',
		'    <ns0:hasPart xmlns:ns0="http://docs.oasis-open.org/ns/office/1.2/meta/pkg#" rdf:resource="' + file + '"/>\n',
		'  </rdf:Description>\n'
	].join("");
}
function write_rdf(rdf) {
	var o = [XML_HEADER];
	o.push('<rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">\n');
	for(var i = 0; i != rdf.length; ++i) {
		o.push(write_rdf_type(rdf[i][0], rdf[i][1]));
		o.push(write_rdf_has("",rdf[i][0]));
	}
	o.push(write_rdf_type("","Document", "pkg"));
	o.push('</rdf:RDF>');
	return o.join("");
}
/* TODO: pull properties */
var write_meta_ods = (function() {
	var payload = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.2"><office:meta><meta:generator>Sheet' + 'JS ' + XLSX.version + '</meta:generator></office:meta></office:document-meta>';
	return function wmo() {
		return payload;
	};
})();

/* ECMA-376 Part II 11.1 Core Properties Part */
/* [MS-OSHARED] 2.3.3.2.[1-2].1 (PIDSI/PIDDSI) */
var CORE_PROPS = [
	["cp:category", "Category"],
	["cp:contentStatus", "ContentStatus"],
	["cp:keywords", "Keywords"],
	["cp:lastModifiedBy", "LastAuthor"],
	["cp:lastPrinted", "LastPrinted"],
	["cp:revision", "RevNumber"],
	["cp:version", "Version"],
	["dc:creator", "Author"],
	["dc:description", "Comments"],
	["dc:identifier", "Identifier"],
	["dc:language", "Language"],
	["dc:subject", "Subject"],
	["dc:title", "Title"],
	["dcterms:created", "CreatedDate", 'date'],
	["dcterms:modified", "ModifiedDate", 'date']
];

XMLNS.CORE_PROPS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
RELS.CORE_PROPS  = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';

var CORE_PROPS_REGEX = (function() {
	var r = new Array(CORE_PROPS.length);
	for(var i = 0; i < CORE_PROPS.length; ++i) {
		var f = CORE_PROPS[i];
		var g = "(?:"+ f[0].slice(0,f[0].indexOf(":")) +":)"+ f[0].slice(f[0].indexOf(":")+1);
		r[i] = new RegExp("<" + g + "[^>]*>([\\s\\S]*?)<\/" + g + ">");
	}
	return r;
})();

function parse_core_props(data) {
	var p = {};
	data = utf8read(data);

	for(var i = 0; i < CORE_PROPS.length; ++i) {
		var f = CORE_PROPS[i], cur = data.match(CORE_PROPS_REGEX[i]);
		if(cur != null && cur.length > 0) p[f[1]] = cur[1];
		if(f[2] === 'date' && p[f[1]]) p[f[1]] = parseDate(p[f[1]]);
	}

	return p;
}

var CORE_PROPS_XML_ROOT = writextag('cp:coreProperties', null, {
	//'xmlns': XMLNS.CORE_PROPS,
	'xmlns:cp': XMLNS.CORE_PROPS,
	'xmlns:dc': XMLNS.dc,
	'xmlns:dcterms': XMLNS.dcterms,
	'xmlns:dcmitype': XMLNS.dcmitype,
	'xmlns:xsi': XMLNS.xsi
});

function cp_doit(f, g, h, o, p) {
	if(p[f] != null || g == null || g === "") return;
	p[f] = g;
	o[o.length] = (h ? writextag(f,g,h) : writetag(f,g));
}

function write_core_props(cp, _opts) {
	var opts = _opts || {};
	var o = [XML_HEADER, CORE_PROPS_XML_ROOT], p = {};
	if(!cp && !opts.Props) return o.join("");

	if(cp) {
		if(cp.CreatedDate != null) cp_doit("dcterms:created", typeof cp.CreatedDate === "string" ? cp.CreatedDate : write_w3cdtf(cp.CreatedDate, opts.WTF), {"xsi:type":"dcterms:W3CDTF"}, o, p);
		if(cp.ModifiedDate != null) cp_doit("dcterms:modified", typeof cp.ModifiedDate === "string" ? cp.ModifiedDate : write_w3cdtf(cp.ModifiedDate, opts.WTF), {"xsi:type":"dcterms:W3CDTF"}, o, p);
	}

	for(var i = 0; i != CORE_PROPS.length; ++i) {
		var f = CORE_PROPS[i];
		var v = opts.Props && opts.Props[f[1]] != null ? opts.Props[f[1]] : cp ? cp[f[1]] : null;
		if(v === true) v = "1";
		else if(v === false) v = "0";
		else if(typeof v == "number") v = String(v);
		if(v != null) cp_doit(f[0], v, null, o, p);
	}
	if(o.length>2){ o[o.length] = ('</cp:coreProperties>'); o[1]=o[1].replace("/>",">"); }
	return o.join("");
}
/* 15.2.12.3 Extended File Properties Part */
/* [MS-OSHARED] 2.3.3.2.[1-2].1 (PIDSI/PIDDSI) */
var EXT_PROPS = [
	["Application", "Application", "string"],
	["AppVersion", "AppVersion", "string"],
	["Company", "Company", "string"],
	["DocSecurity", "DocSecurity", "string"],
	["Manager", "Manager", "string"],
	["HyperlinksChanged", "HyperlinksChanged", "bool"],
	["SharedDoc", "SharedDoc", "bool"],
	["LinksUpToDate", "LinksUpToDate", "bool"],
	["ScaleCrop", "ScaleCrop", "bool"],
	["HeadingPairs", "HeadingPairs", "raw"],
	["TitlesOfParts", "TitlesOfParts", "raw"]
];

XMLNS.EXT_PROPS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
RELS.EXT_PROPS  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties';

var PseudoPropsPairs = [
	"Worksheets",  "SheetNames",
	"NamedRanges", "DefinedNames",
	"Chartsheets", "ChartNames"
];
function load_props_pairs(HP, TOP, props, opts) {
	var v = [];
	if(typeof HP == "string") v = parseVector(HP, opts);
	else for(var j = 0; j < HP.length; ++j) v = v.concat(HP[j].map(function(hp) { return {v:hp}; }));
	var parts = (typeof TOP == "string") ? parseVector(TOP, opts).map(function (x) { return x.v; }) : TOP;
	var idx = 0, len = 0;
	if(parts.length > 0) for(var i = 0; i !== v.length; i += 2) {
		len = +(v[i+1].v);
		switch(v[i].v) {
			case "Worksheets":
			case "工作表":
			case "Листы":
			case "أوراق العمل":
			case "ワークシート":
			case "גליונות עבודה":
			case "Arbeitsblätter":
			case "Çalışma Sayfaları":
			case "Feuilles de calcul":
			case "Fogli di lavoro":
			case "Folhas de cálculo":
			case "Planilhas":
			case "Regneark":
			case "Werkbladen":
				props.Worksheets = len;
				props.SheetNames = parts.slice(idx, idx + len);
				break;

			case "Named Ranges":
			case "名前付き一覧":
			case "Benannte Bereiche":
			case "Navngivne områder":
				props.NamedRanges = len;
				props.DefinedNames = parts.slice(idx, idx + len);
				break;

			case "Charts":
			case "Diagramme":
				props.Chartsheets = len;
				props.ChartNames = parts.slice(idx, idx + len);
				break;
		}
		idx += len;
	}
}

function parse_ext_props(data, p, opts) {
	var q = {}; if(!p) p = {};
	data = utf8read(data);

	EXT_PROPS.forEach(function(f) {
		switch(f[2]) {
			case "string": p[f[1]] = (data.match(matchtag(f[0]))||[])[1]; break;
			case "bool": p[f[1]] = (data.match(matchtag(f[0]))||[])[1] === "true"; break;
			case "raw":
				var cur = data.match(new RegExp("<" + f[0] + "[^>]*>([\\s\\S]*?)<\/" + f[0] + ">"));
				if(cur && cur.length > 0) q[f[1]] = cur[1];
				break;
		}
	});

	if(q.HeadingPairs && q.TitlesOfParts) load_props_pairs(q.HeadingPairs, q.TitlesOfParts, p, opts);

	return p;
}

var EXT_PROPS_XML_ROOT = writextag('Properties', null, {
	'xmlns': XMLNS.EXT_PROPS,
	'xmlns:vt': XMLNS.vt
});

function write_ext_props(cp) {
	var o = [], W = writextag;
	if(!cp) cp = {};
	cp.Application = "SheetJS";
	o[o.length] = (XML_HEADER);
	o[o.length] = (EXT_PROPS_XML_ROOT);

	EXT_PROPS.forEach(function(f) {
		if(cp[f[1]] === undefined) return;
		var v;
		switch(f[2]) {
			case 'string': v = String(cp[f[1]]); break;
			case 'bool': v = cp[f[1]] ? 'true' : 'false'; break;
		}
		if(v !== undefined) o[o.length] = (W(f[0], v));
	});

	/* TODO: HeadingPairs, TitlesOfParts */
	o[o.length] = (W('HeadingPairs', W('vt:vector', W('vt:variant', '<vt:lpstr>Worksheets</vt:lpstr>')+W('vt:variant', W('vt:i4', String(cp.Worksheets))), {size:2, baseType:"variant"})));
	o[o.length] = (W('TitlesOfParts', W('vt:vector', cp.SheetNames.map(function(s) { return "<vt:lpstr>" + escapexml(s) + "</vt:lpstr>"; }).join(""), {size: cp.Worksheets, baseType:"lpstr"})));
	if(o.length>2){ o[o.length] = ('</Properties>'); o[1]=o[1].replace("/>",">"); }
	return o.join("");
}
/* 15.2.12.2 Custom File Properties Part */
XMLNS.CUST_PROPS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
RELS.CUST_PROPS  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties';

var custregex = /<[^>]+>[^<]*/g;
function parse_cust_props(data, opts) {
	var p = {}, name = "";
	var m = data.match(custregex);
	if(m) for(var i = 0; i != m.length; ++i) {
		var x = m[i], y = parsexmltag(x);
		switch(y[0]) {
			case '<?xml': break;
			case '<Properties': break;
			case '<property': name = y.name; break;
			case '</property>': name = null; break;
			default: if (x.indexOf('<vt:') === 0) {
				var toks = x.split('>');
				var type = toks[0].slice(4), text = toks[1];
				/* 22.4.2.32 (CT_Variant). Omit the binary types from 22.4 (Variant Types) */
				switch(type) {
					case 'lpstr': case 'bstr': case 'lpwstr':
						p[name] = unescapexml(text);
						break;
					case 'bool':
						p[name] = parsexmlbool(text);
						break;
					case 'i1': case 'i2': case 'i4': case 'i8': case 'int': case 'uint':
						p[name] = parseInt(text, 10);
						break;
					case 'r4': case 'r8': case 'decimal':
						p[name] = parseFloat(text);
						break;
					case 'filetime': case 'date':
						p[name] = parseDate(text);
						break;
					case 'cy': case 'error':
						p[name] = unescapexml(text);
						break;
					default:
						if(type.slice(-1) == '/') break;
						if(opts.WTF && typeof console !== 'undefined') console.warn('Unexpected', x, type, toks);
				}
			} else if(x.slice(0,2) === "</") {/* empty */
			} else if(opts.WTF) throw new Error(x);
		}
	}
	return p;
}

var CUST_PROPS_XML_ROOT = writextag('Properties', null, {
	'xmlns': XMLNS.CUST_PROPS,
	'xmlns:vt': XMLNS.vt
});

function write_cust_props(cp) {
	var o = [XML_HEADER, CUST_PROPS_XML_ROOT];
	if(!cp) return o.join("");
	var pid = 1;
	keys(cp).forEach(function custprop(k) { ++pid;
		o[o.length] = (writextag('property', write_vt(cp[k]), {
			'fmtid': '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}',
			'pid': pid,
			'name': k
		}));
	});
	if(o.length>2){ o[o.length] = '</Properties>'; o[1]=o[1].replace("/>",">"); }
	return o.join("");
}
/* Common Name -> XLML Name */
var XLMLDocPropsMap = {
	Title: 'Title',
	Subject: 'Subject',
	Author: 'Author',
	Keywords: 'Keywords',
	Comments: 'Description',
	LastAuthor: 'LastAuthor',
	RevNumber: 'Revision',
	Application: 'AppName',
	/* TotalTime: 'TotalTime', */
	LastPrinted: 'LastPrinted',
	CreatedDate: 'Created',
	ModifiedDate: 'LastSaved',
	/* Pages */
	/* Words */
	/* Characters */
	Category: 'Category',
	/* PresentationFormat */
	Manager: 'Manager',
	Company: 'Company',
	/* Guid */
	/* HyperlinkBase */
	/* Bytes */
	/* Lines */
	/* Paragraphs */
	/* CharactersWithSpaces */
	AppVersion: 'Version',

	ContentStatus: 'ContentStatus', /* NOTE: missing from schema */
	Identifier: 'Identifier', /* NOTE: missing from schema */
	Language: 'Language' /* NOTE: missing from schema */
};
var evert_XLMLDPM = evert(XLMLDocPropsMap);

function xlml_set_prop(Props, tag, val) {
	tag = evert_XLMLDPM[tag] || tag;
	Props[tag] = val;
}

function xlml_write_docprops(Props, opts) {
	var o = [];
	keys(XLMLDocPropsMap).map(function(m) {
		for(var i = 0; i < CORE_PROPS.length; ++i) if(CORE_PROPS[i][1] == m) return CORE_PROPS[i];
		for(i = 0; i < EXT_PROPS.length; ++i) if(EXT_PROPS[i][1] == m) return EXT_PROPS[i];
		throw m;
	}).forEach(function(p) {
		if(Props[p[1]] == null) return;
		var m = opts && opts.Props && opts.Props[p[1]] != null ? opts.Props[p[1]] : Props[p[1]];
		switch(p[2]) {
			case 'date': m = new Date(m).toISOString().replace(/\.\d*Z/,"Z"); break;
		}
		if(typeof m == 'number') m = String(m);
		else if(m === true || m === false) { m = m ? "1" : "0"; }
		else if(m instanceof Date) m = new Date(m).toISOString().replace(/\.\d*Z/,"");
		o.push(writetag(XLMLDocPropsMap[p[1]] || p[1], m));
	});
	return writextag('DocumentProperties', o.join(""), {xmlns:XLMLNS.o });
}
function xlml_write_custprops(Props, Custprops) {
	var BLACKLIST = ["Worksheets","SheetNames"];
	var T = 'CustomDocumentProperties';
	var o = [];
	if(Props) keys(Props).forEach(function(k) {
if(!Props.hasOwnProperty(k)) return;
		for(var i = 0; i < CORE_PROPS.length; ++i) if(k == CORE_PROPS[i][1]) return;
		for(i = 0; i < EXT_PROPS.length; ++i) if(k == EXT_PROPS[i][1]) return;
		for(i = 0; i < BLACKLIST.length; ++i) if(k == BLACKLIST[i]) return;

		var m = Props[k];
		var t = "string";
		if(typeof m == 'number') { t = "float"; m = String(m); }
		else if(m === true || m === false) { t = "boolean"; m = m ? "1" : "0"; }
		else m = String(m);
		o.push(writextag(escapexmltag(k), m, {"dt:dt":t}));
	});
	if(Custprops) keys(Custprops).forEach(function(k) {
if(!Custprops.hasOwnProperty(k)) return;
		if(Props && Props.hasOwnProperty(k)) return;
		var m = Custprops[k];
		var t = "string";
		if(typeof m == 'number') { t = "float"; m = String(m); }
		else if(m === true || m === false) { t = "boolean"; m = m ? "1" : "0"; }
		else if(m instanceof Date) { t = "dateTime.tz"; m = m.toISOString(); }
		else m = String(m);
		o.push(writextag(escapexmltag(k), m, {"dt:dt":t}));
	});
	return '<' + T + ' xmlns="' + XLMLNS.o + '">' + o.join("") + '</' + T + '>';
}
/* [MS-DTYP] 2.3.3 FILETIME */
/* [MS-OLEDS] 2.1.3 FILETIME (Packet Version) */
/* [MS-OLEPS] 2.8 FILETIME (Packet Version) */
function parse_FILETIME(blob) {
	var dwLowDateTime = blob.read_shift(4), dwHighDateTime = blob.read_shift(4);
	return new Date(((dwHighDateTime/1e7*Math.pow(2,32) + dwLowDateTime/1e7) - 11644473600)*1000).toISOString().replace(/\.000/,"");
}
function write_FILETIME(time) {
	var date = (typeof time == "string") ? new Date(Date.parse(time)) : time;
	var t = date.getTime() / 1000 + 11644473600;
	var l = t % Math.pow(2,32), h = (t - l) / Math.pow(2,32);
	l *= 1e7; h *= 1e7;
	var w = (l / Math.pow(2,32)) | 0;
	if(w > 0) { l = l % Math.pow(2,32); h += w; }
	var o = new_buf(8); o.write_shift(4, l); o.write_shift(4, h); return o;
}

/* [MS-OSHARED] 2.3.3.1.4 Lpstr */
function parse_lpstr(blob, type, pad) {
	var start = blob.l;
	var str = blob.read_shift(0, 'lpstr-cp');
	if(pad) while((blob.l - start) & 3) ++blob.l;
	return str;
}

/* [MS-OSHARED] 2.3.3.1.6 Lpwstr */
function parse_lpwstr(blob, type, pad) {
	var str = blob.read_shift(0, 'lpwstr');
	if(pad) blob.l += (4 - ((str.length+1) & 3)) & 3;
	return str;
}


/* [MS-OSHARED] 2.3.3.1.11 VtString */
/* [MS-OSHARED] 2.3.3.1.12 VtUnalignedString */
function parse_VtStringBase(blob, stringType, pad) {
	if(stringType === 0x1F /*VT_LPWSTR*/) return parse_lpwstr(blob);
	return parse_lpstr(blob, stringType, pad);
}

function parse_VtString(blob, t, pad) { return parse_VtStringBase(blob, t, pad === false ? 0: 4); }
function parse_VtUnalignedString(blob, t) { if(!t) throw new Error("VtUnalignedString must have positive length"); return parse_VtStringBase(blob, t, 0); }

/* [MS-OSHARED] 2.3.3.1.9 VtVecUnalignedLpstrValue */
function parse_VtVecUnalignedLpstrValue(blob) {
	var length = blob.read_shift(4);
	var ret = [];
	for(var i = 0; i != length; ++i) ret[i] = blob.read_shift(0, 'lpstr-cp').replace(chr0,'');
	return ret;
}

/* [MS-OSHARED] 2.3.3.1.10 VtVecUnalignedLpstr */
function parse_VtVecUnalignedLpstr(blob) {
	return parse_VtVecUnalignedLpstrValue(blob);
}

/* [MS-OSHARED] 2.3.3.1.13 VtHeadingPair */
function parse_VtHeadingPair(blob) {
	var headingString = parse_TypedPropertyValue(blob, VT_USTR);
	var headerParts = parse_TypedPropertyValue(blob, VT_I4);
	return [headingString, headerParts];
}

/* [MS-OSHARED] 2.3.3.1.14 VtVecHeadingPairValue */
function parse_VtVecHeadingPairValue(blob) {
	var cElements = blob.read_shift(4);
	var out = [];
	for(var i = 0; i != cElements / 2; ++i) out.push(parse_VtHeadingPair(blob));
	return out;
}

/* [MS-OSHARED] 2.3.3.1.15 VtVecHeadingPair */
function parse_VtVecHeadingPair(blob) {
	// NOTE: When invoked, wType & padding were already consumed
	return parse_VtVecHeadingPairValue(blob);
}

/* [MS-OLEPS] 2.18.1 Dictionary (uses 2.17, 2.16) */
function parse_dictionary(blob,CodePage) {
	var cnt = blob.read_shift(4);
	var dict = ({});
	for(var j = 0; j != cnt; ++j) {
		var pid = blob.read_shift(4);
		var len = blob.read_shift(4);
		dict[pid] = blob.read_shift(len, (CodePage === 0x4B0 ?'utf16le':'utf8')).replace(chr0,'').replace(chr1,'!');
		if(CodePage === 0x4B0 && (len % 2)) blob.l += 2;
	}
	if(blob.l & 3) blob.l = (blob.l>>2+1)<<2;
	return dict;
}

/* [MS-OLEPS] 2.9 BLOB */
function parse_BLOB(blob) {
	var size = blob.read_shift(4);
	var bytes = blob.slice(blob.l,blob.l+size);
	blob.l += size;
	if((size & 3) > 0) blob.l += (4 - (size & 3)) & 3;
	return bytes;
}

/* [MS-OLEPS] 2.11 ClipboardData */
function parse_ClipboardData(blob) {
	// TODO
	var o = {};
	o.Size = blob.read_shift(4);
	//o.Format = blob.read_shift(4);
	blob.l += o.Size + 3 - (o.Size - 1) % 4;
	return o;
}

/* [MS-OLEPS] 2.15 TypedPropertyValue */
function parse_TypedPropertyValue(blob, type, _opts) {
	var t = blob.read_shift(2), ret, opts = _opts||{};
	blob.l += 2;
	if(type !== VT_VARIANT)
	if(t !== type && VT_CUSTOM.indexOf(type)===-1) throw new Error('Expected type ' + type + ' saw ' + t);
	switch(type === VT_VARIANT ? t : type) {
		case 0x02 /*VT_I2*/: ret = blob.read_shift(2, 'i'); if(!opts.raw) blob.l += 2; return ret;
		case 0x03 /*VT_I4*/: ret = blob.read_shift(4, 'i'); return ret;
		case 0x0B /*VT_BOOL*/: return blob.read_shift(4) !== 0x0;
		case 0x13 /*VT_UI4*/: ret = blob.read_shift(4); return ret;
		case 0x1E /*VT_LPSTR*/: return parse_lpstr(blob, t, 4).replace(chr0,'');
		case 0x1F /*VT_LPWSTR*/: return parse_lpwstr(blob);
		case 0x40 /*VT_FILETIME*/: return parse_FILETIME(blob);
		case 0x41 /*VT_BLOB*/: return parse_BLOB(blob);
		case 0x47 /*VT_CF*/: return parse_ClipboardData(blob);
		case 0x50 /*VT_STRING*/: return parse_VtString(blob, t, !opts.raw).replace(chr0,'');
		case 0x51 /*VT_USTR*/: return parse_VtUnalignedString(blob, t/*, 4*/).replace(chr0,'');
		case 0x100C /*VT_VECTOR|VT_VARIANT*/: return parse_VtVecHeadingPair(blob);
		case 0x101E /*VT_LPSTR*/: return parse_VtVecUnalignedLpstr(blob);
		default: throw new Error("TypedPropertyValue unrecognized type " + type + " " + t);
	}
}
function write_TypedPropertyValue(type, value) {
	var o = new_buf(4), p = new_buf(4);
	o.write_shift(4, type == 0x50 ? 0x1F : type);
	switch(type) {
		case 0x03 /*VT_I4*/: p.write_shift(-4, value); break;
		case 0x05 /*VT_I4*/: p = new_buf(8); p.write_shift(8, value, 'f'); break;
		case 0x0B /*VT_BOOL*/: p.write_shift(4, value ? 0x01 : 0x00); break;
		case 0x40 /*VT_FILETIME*/:  p = write_FILETIME(value); break;
		case 0x1F /*VT_LPWSTR*/:
		case 0x50 /*VT_STRING*/:
p = new_buf(4 + 2 * (value.length + 1) + (value.length % 2 ? 0 : 2));
			p.write_shift(4, value.length + 1);
			p.write_shift(0, value, "dbcs");
			while(p.l != p.length) p.write_shift(1, 0);
			break;
		default: throw new Error("TypedPropertyValue unrecognized type " + type + " " + value);
	}
	return bconcat([o, p]);
}

/* [MS-OLEPS] 2.20 PropertySet */
function parse_PropertySet(blob, PIDSI) {
	var start_addr = blob.l;
	var size = blob.read_shift(4);
	var NumProps = blob.read_shift(4);
	var Props = [], i = 0;
	var CodePage = 0;
	var Dictionary = -1, DictObj = ({});
	for(i = 0; i != NumProps; ++i) {
		var PropID = blob.read_shift(4);
		var Offset = blob.read_shift(4);
		Props[i] = [PropID, Offset + start_addr];
	}
	Props.sort(function(x,y) { return x[1] - y[1]; });
	var PropH = {};
	for(i = 0; i != NumProps; ++i) {
		if(blob.l !== Props[i][1]) {
			var fail = true;
			if(i>0 && PIDSI) switch(PIDSI[Props[i-1][0]].t) {
				case 0x02 /*VT_I2*/: if(blob.l+2 === Props[i][1]) { blob.l+=2; fail = false; } break;
				case 0x50 /*VT_STRING*/: if(blob.l <= Props[i][1]) { blob.l=Props[i][1]; fail = false; } break;
				case 0x100C /*VT_VECTOR|VT_VARIANT*/: if(blob.l <= Props[i][1]) { blob.l=Props[i][1]; fail = false; } break;
			}
			if((!PIDSI||i==0) && blob.l <= Props[i][1]) { fail=false; blob.l = Props[i][1]; }
			if(fail) throw new Error("Read Error: Expected address " + Props[i][1] + ' at ' + blob.l + ' :' + i);
		}
		if(PIDSI) {
			var piddsi = PIDSI[Props[i][0]];
			PropH[piddsi.n] = parse_TypedPropertyValue(blob, piddsi.t, {raw:true});
			if(piddsi.p === 'version') PropH[piddsi.n] = String(PropH[piddsi.n] >> 16) + "." + ("0000" + String(PropH[piddsi.n] & 0xFFFF)).slice(-4);
			if(piddsi.n == "CodePage") switch(PropH[piddsi.n]) {
				case 0: PropH[piddsi.n] = 1252;
					/* falls through */
				case 874:
				case 932:
				case 936:
				case 949:
				case 950:
				case 1250:
				case 1251:
				case 1253:
				case 1254:
				case 1255:
				case 1256:
				case 1257:
				case 1258:
				case 10000:
				case 1200:
				case 1201:
				case 1252:
				case 65000: case -536:
				case 65001: case -535:
					set_cp(CodePage = (PropH[piddsi.n]>>>0) & 0xFFFF); break;
				default: throw new Error("Unsupported CodePage: " + PropH[piddsi.n]);
			}
		} else {
			if(Props[i][0] === 0x1) {
				CodePage = PropH.CodePage = (parse_TypedPropertyValue(blob, VT_I2));
				set_cp(CodePage);
				if(Dictionary !== -1) {
					var oldpos = blob.l;
					blob.l = Props[Dictionary][1];
					DictObj = parse_dictionary(blob,CodePage);
					blob.l = oldpos;
				}
			} else if(Props[i][0] === 0) {
				if(CodePage === 0) { Dictionary = i; blob.l = Props[i+1][1]; continue; }
				DictObj = parse_dictionary(blob,CodePage);
			} else {
				var name = DictObj[Props[i][0]];
				var val;
				/* [MS-OSHARED] 2.3.3.2.3.1.2 + PROPVARIANT */
				switch(blob[blob.l]) {
					case 0x41 /*VT_BLOB*/: blob.l += 4; val = parse_BLOB(blob); break;
					case 0x1E /*VT_LPSTR*/: blob.l += 4; val = parse_VtString(blob, blob[blob.l-4]).replace(/\u0000+$/,""); break;
					case 0x1F /*VT_LPWSTR*/: blob.l += 4; val = parse_VtString(blob, blob[blob.l-4]).replace(/\u0000+$/,""); break;
					case 0x03 /*VT_I4*/: blob.l += 4; val = blob.read_shift(4, 'i'); break;
					case 0x13 /*VT_UI4*/: blob.l += 4; val = blob.read_shift(4); break;
					case 0x05 /*VT_R8*/: blob.l += 4; val = blob.read_shift(8, 'f'); break;
					case 0x0B /*VT_BOOL*/: blob.l += 4; val = parsebool(blob, 4); break;
					case 0x40 /*VT_FILETIME*/: blob.l += 4; val = parseDate(parse_FILETIME(blob)); break;
					default: throw new Error("unparsed value: " + blob[blob.l]);
				}
				PropH[name] = val;
			}
		}
	}
	blob.l = start_addr + size; /* step ahead to skip padding */
	return PropH;
}
var XLSPSSkip = [ "CodePage", "Thumbnail", "_PID_LINKBASE", "_PID_HLINKS", "SystemIdentifier", "FMTID" ].concat(PseudoPropsPairs);
function guess_property_type(val) {
	switch(typeof val) {
		case "boolean": return 0x0B;
		case "number": return ((val|0)==val) ? 0x03 : 0x05;
		case "string": return 0x1F;
		case "object": if(val instanceof Date) return 0x40; break;
	}
	return -1;
}
function write_PropertySet(entries, RE, PIDSI) {
	var hdr = new_buf(8), piao = [], prop = [];
	var sz = 8, i = 0;

	var pr = new_buf(8), pio = new_buf(8);
	pr.write_shift(4, 0x0002);
	pr.write_shift(4, 0x04B0);
	pio.write_shift(4, 0x0001);
	prop.push(pr); piao.push(pio);
	sz += 8 + pr.length;

	if(!RE) {
		pio = new_buf(8);
		pio.write_shift(4, 0);
		piao.unshift(pio);

		var bufs = [new_buf(4)];
		bufs[0].write_shift(4, entries.length);
		for(i = 0; i < entries.length; ++i) {
			var value = entries[i][0];
			pr = new_buf(4 + 4 + 2 * (value.length + 1) + (value.length % 2 ? 0 : 2));
			pr.write_shift(4, i+2);
			pr.write_shift(4, value.length + 1);
			pr.write_shift(0, value, "dbcs");
			while(pr.l != pr.length) pr.write_shift(1, 0);
			bufs.push(pr);
		}
		pr = bconcat(bufs);
		prop.unshift(pr);
		sz += 8 + pr.length;
	}

	for(i = 0; i < entries.length; ++i) {
		if(RE && !RE[entries[i][0]]) continue;
		if(XLSPSSkip.indexOf(entries[i][0]) > -1) continue;
		if(entries[i][1] == null) continue;

		var val = entries[i][1], idx = 0;
		if(RE) {
			idx = +RE[entries[i][0]];
			var pinfo = (PIDSI)[idx];
			if(pinfo.p == "version" && typeof val == "string") {
var arr = val.split(".");
				val = ((+arr[0])<<16) + ((+arr[1])||0);
			}
			pr = write_TypedPropertyValue(pinfo.t, val);
		} else {
			var T = guess_property_type(val);
			if(T == -1) { T = 0x1F; val = String(val); }
			pr = write_TypedPropertyValue(T, val);
		}
		prop.push(pr);

		pio = new_buf(8);
		pio.write_shift(4, !RE ? 2+i : idx);
		piao.push(pio);

		sz += 8 + pr.length;
	}

	var w = 8 * (prop.length + 1);
	for(i = 0; i < prop.length; ++i) { piao[i].write_shift(4, w); w += prop[i].length; }
	hdr.write_shift(4, sz);
	hdr.write_shift(4, prop.length);
	return bconcat([hdr].concat(piao).concat(prop));
}

/* [MS-OLEPS] 2.21 PropertySetStream */
function parse_PropertySetStream(file, PIDSI, clsid) {
	var blob = file.content;
	if(!blob) return ({});
	prep_blob(blob, 0);

	var NumSets, FMTID0, FMTID1, Offset0, Offset1 = 0;
	blob.chk('feff', 'Byte Order: ');

	/*var vers = */blob.read_shift(2); // TODO: check version
	var SystemIdentifier = blob.read_shift(4);
	var CLSID = blob.read_shift(16);
	if(CLSID !== CFB.utils.consts.HEADER_CLSID && CLSID !== clsid) throw new Error("Bad PropertySet CLSID " + CLSID);
	NumSets = blob.read_shift(4);
	if(NumSets !== 1 && NumSets !== 2) throw new Error("Unrecognized #Sets: " + NumSets);
	FMTID0 = blob.read_shift(16); Offset0 = blob.read_shift(4);

	if(NumSets === 1 && Offset0 !== blob.l) throw new Error("Length mismatch: " + Offset0 + " !== " + blob.l);
	else if(NumSets === 2) { FMTID1 = blob.read_shift(16); Offset1 = blob.read_shift(4); }
	var PSet0 = parse_PropertySet(blob, PIDSI);

	var rval = ({ SystemIdentifier: SystemIdentifier });
	for(var y in PSet0) rval[y] = PSet0[y];
	//rval.blob = blob;
	rval.FMTID = FMTID0;
	//rval.PSet0 = PSet0;
	if(NumSets === 1) return rval;
	if(Offset1 - blob.l == 2) blob.l += 2;
	if(blob.l !== Offset1) throw new Error("Length mismatch 2: " + blob.l + " !== " + Offset1);
	var PSet1;
	try { PSet1 = parse_PropertySet(blob, null); } catch(e) {/* empty */}
	for(y in PSet1) rval[y] = PSet1[y];
	rval.FMTID = [FMTID0, FMTID1]; // TODO: verify FMTID0/1
	return rval;
}
function write_PropertySetStream(entries, clsid, RE, PIDSI, entries2, clsid2) {
	var hdr = new_buf(entries2 ? 68 : 48);
	var bufs = [hdr];
	hdr.write_shift(2, 0xFFFE);
	hdr.write_shift(2, 0x0000); /* TODO: type 1 props */
	hdr.write_shift(4, 0x32363237);
	hdr.write_shift(16, CFB.utils.consts.HEADER_CLSID, "hex");
	hdr.write_shift(4, (entries2 ? 2 : 1));
	hdr.write_shift(16, clsid, "hex");
	hdr.write_shift(4, (entries2 ? 68 : 48));
	var ps0 = write_PropertySet(entries, RE, PIDSI);
	bufs.push(ps0);

	if(entries2) {
		var ps1 = write_PropertySet(entries2, null, null);
		hdr.write_shift(16, clsid2, "hex");
		hdr.write_shift(4, 68 + ps0.length);
		bufs.push(ps1);
	}
	return bconcat(bufs);
}

function parsenoop2(blob, length) { blob.read_shift(length); return null; }
function writezeroes(n, o) { if(!o) o=new_buf(n); for(var j=0; j<n; ++j) o.write_shift(1, 0); return o; }

function parslurp(blob, length, cb) {
	var arr = [], target = blob.l + length;
	while(blob.l < target) arr.push(cb(blob, target - blob.l));
	if(target !== blob.l) throw new Error("Slurp error");
	return arr;
}

function parsebool(blob, length) { return blob.read_shift(length) === 0x1; }
function writebool(v, o) { if(!o) o=new_buf(2); o.write_shift(2, +!!v); return o; }

function parseuint16(blob) { return blob.read_shift(2, 'u'); }
function writeuint16(v, o) { if(!o) o=new_buf(2); o.write_shift(2, v); return o; }
function parseuint16a(blob, length) { return parslurp(blob,length,parseuint16);}

/* --- 2.5 Structures --- */

/* [MS-XLS] 2.5.10 Bes (boolean or error) */
function parse_Bes(blob) {
	var v = blob.read_shift(1), t = blob.read_shift(1);
	return t === 0x01 ? v : v === 0x01;
}
function write_Bes(v, t, o) {
	if(!o) o = new_buf(2);
	o.write_shift(1, +v);
	o.write_shift(1, ((t == 'e') ? 1 : 0));
	return o;
}

/* [MS-XLS] 2.5.240 ShortXLUnicodeString */
function parse_ShortXLUnicodeString(blob, length, opts) {
	var cch = blob.read_shift(opts && opts.biff >= 12 ? 2 : 1);
	var encoding = 'sbcs-cont';
	var cp = current_codepage;
	if(opts && opts.biff >= 8) current_codepage = 1200;
	if(!opts || opts.biff == 8 ) {
		var fHighByte = blob.read_shift(1);
		if(fHighByte) { encoding = 'dbcs-cont'; }
	} else if(opts.biff == 12) {
		encoding = 'wstr';
	}
	if(opts.biff >= 2 && opts.biff <= 5) encoding = 'cpstr';
	var o = cch ? blob.read_shift(cch, encoding) : "";
	current_codepage = cp;
	return o;
}

/* 2.5.293 XLUnicodeRichExtendedString */
function parse_XLUnicodeRichExtendedString(blob) {
	var cp = current_codepage;
	current_codepage = 1200;
	var cch = blob.read_shift(2), flags = blob.read_shift(1);
	var /*fHighByte = flags & 0x1,*/ fExtSt = flags & 0x4, fRichSt = flags & 0x8;
	var width = 1 + (flags & 0x1); // 0x0 -> utf8, 0x1 -> dbcs
	var cRun = 0, cbExtRst;
	var z = {};
	if(fRichSt) cRun = blob.read_shift(2);
	if(fExtSt) cbExtRst = blob.read_shift(4);
	var encoding = width == 2 ? 'dbcs-cont' : 'sbcs-cont';
	var msg = cch === 0 ? "" : blob.read_shift(cch, encoding);
	if(fRichSt) blob.l += 4 * cRun; //TODO: parse this
	if(fExtSt) blob.l += cbExtRst; //TODO: parse this
	z.t = msg;
	if(!fRichSt) { z.raw = "<t>" + z.t + "</t>"; z.r = z.t; }
	current_codepage = cp;
	return z;
}

/* 2.5.296 XLUnicodeStringNoCch */
function parse_XLUnicodeStringNoCch(blob, cch, opts) {
	var retval;
	if(opts) {
		if(opts.biff >= 2 && opts.biff <= 5) return blob.read_shift(cch, 'cpstr');
		if(opts.biff >= 12) return blob.read_shift(cch, 'dbcs-cont');
	}
	var fHighByte = blob.read_shift(1);
	if(fHighByte===0) { retval = blob.read_shift(cch, 'sbcs-cont'); }
	else { retval = blob.read_shift(cch, 'dbcs-cont'); }
	return retval;
}

/* 2.5.294 XLUnicodeString */
function parse_XLUnicodeString(blob, length, opts) {
	var cch = blob.read_shift(opts && opts.biff == 2 ? 1 : 2);
	if(cch === 0) { blob.l++; return ""; }
	return parse_XLUnicodeStringNoCch(blob, cch, opts);
}
/* BIFF5 override */
function parse_XLUnicodeString2(blob, length, opts) {
	if(opts.biff > 5) return parse_XLUnicodeString(blob, length, opts);
	var cch = blob.read_shift(1);
	if(cch === 0) { blob.l++; return ""; }
	return blob.read_shift(cch, (opts.biff <= 4 || !blob.lens ) ? 'cpstr' : 'sbcs-cont');
}
/* TODO: BIFF5 and lower, codepage awareness */
function write_XLUnicodeString(str, opts, o) {
	if(!o) o = new_buf(3 + 2 * str.length);
	o.write_shift(2, str.length);
	o.write_shift(1, 1);
	o.write_shift(31, str, 'utf16le');
	return o;
}

/* [MS-XLS] 2.5.61 ControlInfo */
function parse_ControlInfo(blob) {
	var flags = blob.read_shift(1);
	blob.l++;
	var accel = blob.read_shift(2);
	blob.l += 2;
	return [flags, accel];
}

/* [MS-OSHARED] 2.3.7.6 URLMoniker TODO: flags */
function parse_URLMoniker(blob) {
	var len = blob.read_shift(4), start = blob.l;
	var extra = false;
	if(len > 24) {
		/* look ahead */
		blob.l += len - 24;
		if(blob.read_shift(16) === "795881f43b1d7f48af2c825dc4852763") extra = true;
		blob.l = start;
	}
	var url = blob.read_shift((extra?len-24:len)>>1, 'utf16le').replace(chr0,"");
	if(extra) blob.l += 24;
	return url;
}

/* [MS-OSHARED] 2.3.7.8 FileMoniker TODO: all fields */
function parse_FileMoniker(blob) {
	blob.l += 2; //var cAnti = blob.read_shift(2);
	var ansiPath = blob.read_shift(0, 'lpstr-ansi');
	blob.l += 2; //var endServer = blob.read_shift(2);
	if(blob.read_shift(2) != 0xDEAD) throw new Error("Bad FileMoniker");
	var sz = blob.read_shift(4);
	if(sz === 0) return ansiPath.replace(/\\/g,"/");
	var bytes = blob.read_shift(4);
	if(blob.read_shift(2) != 3) throw new Error("Bad FileMoniker");
	var unicodePath = blob.read_shift(bytes>>1, 'utf16le').replace(chr0,"");
	return unicodePath;
}

/* [MS-OSHARED] 2.3.7.2 HyperlinkMoniker TODO: all the monikers */
function parse_HyperlinkMoniker(blob, length) {
	var clsid = blob.read_shift(16); length -= 16;
	switch(clsid) {
		case "e0c9ea79f9bace118c8200aa004ba90b": return parse_URLMoniker(blob, length);
		case "0303000000000000c000000000000046": return parse_FileMoniker(blob, length);
		default: throw new Error("Unsupported Moniker " + clsid);
	}
}

/* [MS-OSHARED] 2.3.7.9 HyperlinkString */
function parse_HyperlinkString(blob) {
	var len = blob.read_shift(4);
	var o = len > 0 ? blob.read_shift(len, 'utf16le').replace(chr0, "") : "";
	return o;
}

/* [MS-OSHARED] 2.3.7.1 Hyperlink Object */
function parse_Hyperlink(blob, length) {
	var end = blob.l + length;
	var sVer = blob.read_shift(4);
	if(sVer !== 2) throw new Error("Unrecognized streamVersion: " + sVer);
	var flags = blob.read_shift(2);
	blob.l += 2;
	var displayName, targetFrameName, moniker, oleMoniker, Loc="", guid, fileTime;
	if(flags & 0x0010) displayName = parse_HyperlinkString(blob, end - blob.l);
	if(flags & 0x0080) targetFrameName = parse_HyperlinkString(blob, end - blob.l);
	if((flags & 0x0101) === 0x0101) moniker = parse_HyperlinkString(blob, end - blob.l);
	if((flags & 0x0101) === 0x0001) oleMoniker = parse_HyperlinkMoniker(blob, end - blob.l);
	if(flags & 0x0008) Loc = parse_HyperlinkString(blob, end - blob.l);
	if(flags & 0x0020) guid = blob.read_shift(16);
	if(flags & 0x0040) fileTime = parse_FILETIME(blob/*, 8*/);
	blob.l = end;
	var target = targetFrameName||moniker||oleMoniker||"";
	if(target && Loc) target+="#"+Loc;
	if(!target) target = "#" + Loc;
	var out = ({Target:target});
	if(guid) out.guid = guid;
	if(fileTime) out.time = fileTime;
	if(displayName) out.Tooltip = displayName;
	return out;
}
function write_Hyperlink(hl) {
	var out = new_buf(512), i = 0;
	var Target = hl.Target;
	var F = Target.indexOf("#") > -1 ? 0x1f : 0x17;
	switch(Target.charAt(0)) { case "#": F=0x1c; break; case ".": F&=~2; break; }
	out.write_shift(4,2); out.write_shift(4, F);
	var data = [8,6815827,6619237,4849780,83]; for(i = 0; i < data.length; ++i) out.write_shift(4, data[i]);
	if(F == 0x1C) {
		Target = Target.slice(1);
		out.write_shift(4, Target.length + 1);
		for(i = 0; i < Target.length; ++i) out.write_shift(2, Target.charCodeAt(i));
		out.write_shift(2, 0);
	} else if(F & 0x02) {
		data = "e0 c9 ea 79 f9 ba ce 11 8c 82 00 aa 00 4b a9 0b".split(" ");
		for(i = 0; i < data.length; ++i) out.write_shift(1, parseInt(data[i], 16));
		out.write_shift(4, 2*(Target.length + 1));
		for(i = 0; i < Target.length; ++i) out.write_shift(2, Target.charCodeAt(i));
		out.write_shift(2, 0);
	} else {
		data = "03 03 00 00 00 00 00 00 c0 00 00 00 00 00 00 46".split(" ");
		for(i = 0; i < data.length; ++i) out.write_shift(1, parseInt(data[i], 16));
		var P = 0;
		while(Target.slice(P*3,P*3+3)=="../"||Target.slice(P*3,P*3+3)=="..\\") ++P;
		out.write_shift(2, P);
		out.write_shift(4, Target.length + 1);
		for(i = 0; i < Target.length; ++i) out.write_shift(1, Target.charCodeAt(i) & 0xFF);
		out.write_shift(1, 0);
		out.write_shift(2, 0xFFFF);
		out.write_shift(2, 0xDEAD);
		for(i = 0; i < 6; ++i) out.write_shift(4, 0);
	}
	return out.slice(0, out.l);
}

/* 2.5.178 LongRGBA */
function parse_LongRGBA(blob) { var r = blob.read_shift(1), g = blob.read_shift(1), b = blob.read_shift(1), a = blob.read_shift(1); return [r,g,b,a]; }

/* 2.5.177 LongRGB */
function parse_LongRGB(blob, length) { var x = parse_LongRGBA(blob, length); x[3] = 0; return x; }


/* [MS-XLS] 2.5.19 */
function parse_XLSCell(blob) {
	var rw = blob.read_shift(2); // 0-indexed
	var col = blob.read_shift(2);
	var ixfe = blob.read_shift(2);
	return ({r:rw, c:col, ixfe:ixfe});
}
function write_XLSCell(R, C, ixfe, o) {
	if(!o) o = new_buf(6);
	o.write_shift(2, R);
	o.write_shift(2, C);
	o.write_shift(2, ixfe||0);
	return o;
}

/* [MS-XLS] 2.5.134 */
function parse_frtHeader(blob) {
	var rt = blob.read_shift(2);
	var flags = blob.read_shift(2); // TODO: parse these flags
	blob.l += 8;
	return {type: rt, flags: flags};
}



function parse_OptXLUnicodeString(blob, length, opts) { return length === 0 ? "" : parse_XLUnicodeString2(blob, length, opts); }

/* [MS-XLS] 2.5.344 */
function parse_XTI(blob, length, opts) {
	var w = opts.biff > 8 ? 4 : 2;
	var iSupBook = blob.read_shift(w), itabFirst = blob.read_shift(w,'i'), itabLast = blob.read_shift(w,'i');
	return [iSupBook, itabFirst, itabLast];
}

/* [MS-XLS] 2.5.218 */
function parse_RkRec(blob) {
	var ixfe = blob.read_shift(2);
	var RK = parse_RkNumber(blob);
	return [ixfe, RK];
}

/* [MS-XLS] 2.5.1 */
function parse_AddinUdf(blob, length, opts) {
	blob.l += 4; length -= 4;
	var l = blob.l + length;
	var udfName = parse_ShortXLUnicodeString(blob, length, opts);
	var cb = blob.read_shift(2);
	l -= blob.l;
	if(cb !== l) throw new Error("Malformed AddinUdf: padding = " + l + " != " + cb);
	blob.l += cb;
	return udfName;
}

/* [MS-XLS] 2.5.209 TODO: Check sizes */
function parse_Ref8U(blob) {
	var rwFirst = blob.read_shift(2);
	var rwLast = blob.read_shift(2);
	var colFirst = blob.read_shift(2);
	var colLast = blob.read_shift(2);
	return {s:{c:colFirst, r:rwFirst}, e:{c:colLast,r:rwLast}};
}
function write_Ref8U(r, o) {
	if(!o) o = new_buf(8);
	o.write_shift(2, r.s.r);
	o.write_shift(2, r.e.r);
	o.write_shift(2, r.s.c);
	o.write_shift(2, r.e.c);
	return o;
}

/* [MS-XLS] 2.5.211 */
function parse_RefU(blob) {
	var rwFirst = blob.read_shift(2);
	var rwLast = blob.read_shift(2);
	var colFirst = blob.read_shift(1);
	var colLast = blob.read_shift(1);
	return {s:{c:colFirst, r:rwFirst}, e:{c:colLast,r:rwLast}};
}

/* [MS-XLS] 2.5.207 */
var parse_Ref = parse_RefU;

/* [MS-XLS] 2.5.143 */
function parse_FtCmo(blob) {
	blob.l += 4;
	var ot = blob.read_shift(2);
	var id = blob.read_shift(2);
	var flags = blob.read_shift(2);
	blob.l+=12;
	return [id, ot, flags];
}

/* [MS-XLS] 2.5.149 */
function parse_FtNts(blob) {
	var out = {};
	blob.l += 4;
	blob.l += 16; // GUID TODO
	out.fSharedNote = blob.read_shift(2);
	blob.l += 4;
	return out;
}

/* [MS-XLS] 2.5.142 */
function parse_FtCf(blob) {
	var out = {};
	blob.l += 4;
	blob.cf = blob.read_shift(2);
	return out;
}

/* [MS-XLS] 2.5.140 - 2.5.154 and friends */
function parse_FtSkip(blob) { blob.l += 2; blob.l += blob.read_shift(2); }
var FtTab = {
0x00: parse_FtSkip,      /* FtEnd */
0x04: parse_FtSkip,      /* FtMacro */
0x05: parse_FtSkip,      /* FtButton */
0x06: parse_FtSkip,      /* FtGmo */
0x07: parse_FtCf,        /* FtCf */
0x08: parse_FtSkip,      /* FtPioGrbit */
0x09: parse_FtSkip,      /* FtPictFmla */
0x0A: parse_FtSkip,      /* FtCbls */
0x0B: parse_FtSkip,      /* FtRbo */
0x0C: parse_FtSkip,      /* FtSbs */
0x0D: parse_FtNts,       /* FtNts */
0x0E: parse_FtSkip,      /* FtSbsFmla */
0x0F: parse_FtSkip,      /* FtGboData */
0x10: parse_FtSkip,      /* FtEdoData */
0x11: parse_FtSkip,      /* FtRboData */
0x12: parse_FtSkip,      /* FtCblsData */
0x13: parse_FtSkip,      /* FtLbsData */
0x14: parse_FtSkip,      /* FtCblsFmla */
0x15: parse_FtCmo
};
function parse_FtArray(blob, length) {
	var tgt = blob.l + length;
	var fts = [];
	while(blob.l < tgt) {
		var ft = blob.read_shift(2);
		blob.l-=2;
		try {
			fts.push(FtTab[ft](blob, tgt - blob.l));
		} catch(e) { blob.l = tgt; return fts; }
	}
	if(blob.l != tgt) blob.l = tgt; //throw new Error("bad Object Ft-sequence");
	return fts;
}

/* --- 2.4 Records --- */

/* [MS-XLS] 2.4.21 */
function parse_BOF(blob, length) {
	var o = {BIFFVer:0, dt:0};
	o.BIFFVer = blob.read_shift(2); length -= 2;
	if(length >= 2) { o.dt = blob.read_shift(2); blob.l -= 2; }
	switch(o.BIFFVer) {
		case 0x0600: /* BIFF8 */
		case 0x0500: /* BIFF5 */
		case 0x0400: /* BIFF4 */
		case 0x0300: /* BIFF3 */
		case 0x0200: /* BIFF2 */
		case 0x0002: case 0x0007: /* BIFF2 */
			break;
		default: if(length > 6) throw new Error("Unexpected BIFF Ver " + o.BIFFVer);
	}

	blob.read_shift(length);
	return o;
}
function write_BOF(wb, t, o) {
	var h = 0x0600, w = 16;
	switch(o.bookType) {
		case 'biff8': break;
		case 'biff5': h = 0x0500; w = 8; break;
		case 'biff4': h = 0x0004; w = 6; break;
		case 'biff3': h = 0x0003; w = 6; break;
		case 'biff2': h = 0x0002; w = 4; break;
		case 'xla': break;
		default: throw new Error("unsupported BIFF version");
	}
	var out = new_buf(w);
	out.write_shift(2, h);
	out.write_shift(2, t);
	if(w > 4) out.write_shift(2, 0x7262);
	if(w > 6) out.write_shift(2, 0x07CD);
	if(w > 8) {
		out.write_shift(2, 0xC009);
		out.write_shift(2, 0x0001);
		out.write_shift(2, 0x0706);
		out.write_shift(2, 0x0000);
	}
	return out;
}


/* [MS-XLS] 2.4.146 */
function parse_InterfaceHdr(blob, length) {
	if(length === 0) return 0x04b0;
	if((blob.read_shift(2))!==0x04b0){/* empty */}
	return 0x04b0;
}


/* [MS-XLS] 2.4.349 */
function parse_WriteAccess(blob, length, opts) {
	if(opts.enc) { blob.l += length; return ""; }
	var l = blob.l;
	// TODO: make sure XLUnicodeString doesnt overrun
	var UserName = parse_XLUnicodeString2(blob, 0, opts);
	blob.read_shift(length + l - blob.l);
	return UserName;
}
function write_WriteAccess(s, opts) {
	var b8 = !opts || opts.biff == 8;
	var o = new_buf(b8 ? 112 : 54);
	o.write_shift(opts.biff == 8 ? 2 : 1, 7);
	if(b8) o.write_shift(1, 0);
	o.write_shift(4, 0x33336853);
	o.write_shift(4, (0x00534A74 | (b8 ? 0 : 0x20000000)));
	while(o.l < o.length) o.write_shift(1, (b8 ? 0 : 32));
	return o;
}

/* [MS-XLS] 2.4.351 */
function parse_WsBool(blob, length, opts) {
	var flags = opts && opts.biff == 8 || length == 2 ? blob.read_shift(2) : (blob.l += length, 0);
	return { fDialog: flags & 0x10 };
}

/* [MS-XLS] 2.4.28 */
function parse_BoundSheet8(blob, length, opts) {
	var pos = blob.read_shift(4);
	var hidden = blob.read_shift(1) & 0x03;
	var dt = blob.read_shift(1);
	switch(dt) {
		case 0: dt = 'Worksheet'; break;
		case 1: dt = 'Macrosheet'; break;
		case 2: dt = 'Chartsheet'; break;
		case 6: dt = 'VBAModule'; break;
	}
	var name = parse_ShortXLUnicodeString(blob, 0, opts);
	if(name.length === 0) name = "Sheet1";
	return { pos:pos, hs:hidden, dt:dt, name:name };
}
function write_BoundSheet8(data, opts) {
	var w = (!opts || opts.biff >= 8 ? 2 : 1);
	var o = new_buf(8 + w * data.name.length);
	o.write_shift(4, data.pos);
	o.write_shift(1, data.hs || 0);
	o.write_shift(1, data.dt);
	o.write_shift(1, data.name.length);
	if(opts.biff >= 8) o.write_shift(1, 1);
	o.write_shift(w * data.name.length, data.name, opts.biff < 8 ? 'sbcs' : 'utf16le');
	var out = o.slice(0, o.l);
	out.l = o.l; return out;
}

/* [MS-XLS] 2.4.265 TODO */
function parse_SST(blob, length) {
	var end = blob.l + length;
	var cnt = blob.read_shift(4);
	var ucnt = blob.read_shift(4);
	var strs = ([]);
	for(var i = 0; i != ucnt && blob.l < end; ++i) {
		strs.push(parse_XLUnicodeRichExtendedString(blob));
	}
	strs.Count = cnt; strs.Unique = ucnt;
	return strs;
}

/* [MS-XLS] 2.4.107 */
function parse_ExtSST(blob, length) {
	var extsst = {};
	extsst.dsst = blob.read_shift(2);
	blob.l += length-2;
	return extsst;
}


/* [MS-XLS] 2.4.221 TODO: check BIFF2-4 */
function parse_Row(blob) {
	var z = ({});
	z.r = blob.read_shift(2);
	z.c = blob.read_shift(2);
	z.cnt = blob.read_shift(2) - z.c;
	var miyRw = blob.read_shift(2);
	blob.l += 4; // reserved(2), unused(2)
	var flags = blob.read_shift(1); // various flags
	blob.l += 3; // reserved(8), ixfe(12), flags(4)
	if(flags & 0x07) z.level = flags & 0x07;
	// collapsed: flags & 0x10
	if(flags & 0x20) z.hidden = true;
	if(flags & 0x40) z.hpt = miyRw / 20;
	return z;
}


/* [MS-XLS] 2.4.125 */
function parse_ForceFullCalculation(blob) {
	var header = parse_frtHeader(blob);
	if(header.type != 0x08A3) throw new Error("Invalid Future Record " + header.type);
	var fullcalc = blob.read_shift(4);
	return fullcalc !== 0x0;
}





/* [MS-XLS] 2.4.215 rt */
function parse_RecalcId(blob) {
	blob.read_shift(2);
	return blob.read_shift(4);
}

/* [MS-XLS] 2.4.87 */
function parse_DefaultRowHeight(blob, length, opts) {
	var f = 0;
	if(!(opts && opts.biff == 2)) {
		f = blob.read_shift(2);
	}
	var miyRw = blob.read_shift(2);
	if((opts && opts.biff == 2)) {
		f = 1 - (miyRw >> 15); miyRw &= 0x7fff;
	}
	var fl = {Unsynced:f&1,DyZero:(f&2)>>1,ExAsc:(f&4)>>2,ExDsc:(f&8)>>3};
	return [fl, miyRw];
}

/* [MS-XLS] 2.4.345 TODO */
function parse_Window1(blob) {
	var xWn = blob.read_shift(2), yWn = blob.read_shift(2), dxWn = blob.read_shift(2), dyWn = blob.read_shift(2);
	var flags = blob.read_shift(2), iTabCur = blob.read_shift(2), iTabFirst = blob.read_shift(2);
	var ctabSel = blob.read_shift(2), wTabRatio = blob.read_shift(2);
	return { Pos: [xWn, yWn], Dim: [dxWn, dyWn], Flags: flags, CurTab: iTabCur,
		FirstTab: iTabFirst, Selected: ctabSel, TabRatio: wTabRatio };
}
function write_Window1() {
	var o = new_buf(18);
	o.write_shift(2, 0);
	o.write_shift(2, 0);
	o.write_shift(2, 0x7260);
	o.write_shift(2, 0x44c0);
	o.write_shift(2, 0x38);
	o.write_shift(2, 0);
	o.write_shift(2, 0);
	o.write_shift(2, 1);
	o.write_shift(2, 0x01f4);
	return o;
}
/* [MS-XLS] 2.4.346 TODO */
function parse_Window2(blob, length, opts) {
	if(opts && opts.biff >= 2 && opts.biff < 8) return {};
	var f = blob.read_shift(2);
	return { RTL: f & 0x40 };
}
function write_Window2(view) {
	var o = new_buf(18), f = 0x6b6;
	if(view && view.RTL) f |= 0x40;
	o.write_shift(2, f);
	o.write_shift(4, 0);
	o.write_shift(4, 64);
	o.write_shift(4, 0);
	o.write_shift(4, 0);
	return o;
}

/* [MS-XLS] 2.4.122 TODO */
function parse_Font(blob, length, opts) {
	var o = {
		dyHeight: blob.read_shift(2),
		fl: blob.read_shift(2)
	};
	switch((opts && opts.biff) || 8) {
		case 2: break;
		case 3: case 4: blob.l += 2; break;
		default: blob.l += 10; break;
	}
	o.name = parse_ShortXLUnicodeString(blob, 0, opts);
	return o;
}
function write_Font(data, opts) {
	var name = data.name || "Arial";
	var b5 = (opts && (opts.biff == 5)), w = (b5 ? (15 + name.length) : (16 + 2 * name.length));
	var o = new_buf(w);
	o.write_shift(2, (data.sz || 12) * 20);
	o.write_shift(4, 0);
	o.write_shift(2, 400);
	o.write_shift(4, 0);
	o.write_shift(2, 0);
	o.write_shift(1, name.length);
	if(!b5) o.write_shift(1, 1);
	o.write_shift((b5 ? 1 : 2) * name.length, name, (b5 ? "sbcs" : "utf16le"));
	return o;
}

/* [MS-XLS] 2.4.149 */
function parse_LabelSst(blob) {
	var cell = parse_XLSCell(blob);
	cell.isst = blob.read_shift(4);
	return cell;
}

/* [MS-XLS] 2.4.148 */
function parse_Label(blob, length, opts) {
	var target = blob.l + length;
	var cell = parse_XLSCell(blob, 6);
	if(opts.biff == 2) blob.l++;
	var str = parse_XLUnicodeString(blob, target - blob.l, opts);
	cell.val = str;
	return cell;
}
function write_Label(R, C, v, os, opts) {
	var b8 = !opts || opts.biff == 8;
	var o = new_buf(6 + 2 + (+b8) + (1 + b8) * v.length);
	write_XLSCell(R, C, os, o);
	o.write_shift(2, v.length);
	if(b8) o.write_shift(1, 1);
	o.write_shift((1 + b8) * v.length, v, b8 ? 'utf16le' : 'sbcs');
	return o;
}


/* [MS-XLS] 2.4.126 Number Formats */
function parse_Format(blob, length, opts) {
	var numFmtId = blob.read_shift(2);
	var fmtstr = parse_XLUnicodeString2(blob, 0, opts);
	return [numFmtId, fmtstr];
}
function write_Format(i, f, opts, o) {
	var b5 = (opts && (opts.biff == 5));
	if(!o) o = new_buf(b5 ? (3 + f.length) : (5 + 2 * f.length));
	o.write_shift(2, i);
	o.write_shift((b5 ? 1 : 2), f.length);
	if(!b5) o.write_shift(1, 1);
	o.write_shift((b5 ? 1 : 2) * f.length, f, (b5 ? 'sbcs' : 'utf16le'));
	var out = (o.length > o.l) ? o.slice(0, o.l) : o;
	if(out.l == null) out.l = out.length;
	return out;
}
var parse_BIFF2Format = parse_XLUnicodeString2;

/* [MS-XLS] 2.4.90 */
function parse_Dimensions(blob, length, opts) {
	var end = blob.l + length;
	var w = opts.biff == 8 || !opts.biff ? 4 : 2;
	var r = blob.read_shift(w), R = blob.read_shift(w);
	var c = blob.read_shift(2), C = blob.read_shift(2);
	blob.l = end;
	return {s: {r:r, c:c}, e: {r:R, c:C}};
}
function write_Dimensions(range, opts) {
	var w = opts.biff == 8 || !opts.biff ? 4 : 2;
	var o = new_buf(2*w + 6);
	o.write_shift(w, range.s.r);
	o.write_shift(w, range.e.r + 1);
	o.write_shift(2, range.s.c);
	o.write_shift(2, range.e.c + 1);
	o.write_shift(2, 0);
	return o;
}

/* [MS-XLS] 2.4.220 */
function parse_RK(blob) {
	var rw = blob.read_shift(2), col = blob.read_shift(2);
	var rkrec = parse_RkRec(blob);
	return {r:rw, c:col, ixfe:rkrec[0], rknum:rkrec[1]};
}

/* [MS-XLS] 2.4.175 */
function parse_MulRk(blob, length) {
	var target = blob.l + length - 2;
	var rw = blob.read_shift(2), col = blob.read_shift(2);
	var rkrecs = [];
	while(blob.l < target) rkrecs.push(parse_RkRec(blob));
	if(blob.l !== target) throw new Error("MulRK read error");
	var lastcol = blob.read_shift(2);
	if(rkrecs.length != lastcol - col + 1) throw new Error("MulRK length mismatch");
	return {r:rw, c:col, C:lastcol, rkrec:rkrecs};
}
/* [MS-XLS] 2.4.174 */
function parse_MulBlank(blob, length) {
	var target = blob.l + length - 2;
	var rw = blob.read_shift(2), col = blob.read_shift(2);
	var ixfes = [];
	while(blob.l < target) ixfes.push(blob.read_shift(2));
	if(blob.l !== target) throw new Error("MulBlank read error");
	var lastcol = blob.read_shift(2);
	if(ixfes.length != lastcol - col + 1) throw new Error("MulBlank length mismatch");
	return {r:rw, c:col, C:lastcol, ixfe:ixfes};
}

/* [MS-XLS] 2.5.20 2.5.249 TODO: interpret values here */
function parse_CellStyleXF(blob, length, style, opts) {
	var o = {};
	var a = blob.read_shift(4), b = blob.read_shift(4);
	var c = blob.read_shift(4), d = blob.read_shift(2);
	o.patternType = XLSFillPattern[c >> 26];

	if(!opts.cellStyles) return o;
	o.alc = a & 0x07;
	o.fWrap = (a >> 3) & 0x01;
	o.alcV = (a >> 4) & 0x07;
	o.fJustLast = (a >> 7) & 0x01;
	o.trot = (a >> 8) & 0xFF;
	o.cIndent = (a >> 16) & 0x0F;
	o.fShrinkToFit = (a >> 20) & 0x01;
	o.iReadOrder = (a >> 22) & 0x02;
	o.fAtrNum = (a >> 26) & 0x01;
	o.fAtrFnt = (a >> 27) & 0x01;
	o.fAtrAlc = (a >> 28) & 0x01;
	o.fAtrBdr = (a >> 29) & 0x01;
	o.fAtrPat = (a >> 30) & 0x01;
	o.fAtrProt = (a >> 31) & 0x01;

	o.dgLeft = b & 0x0F;
	o.dgRight = (b >> 4) & 0x0F;
	o.dgTop = (b >> 8) & 0x0F;
	o.dgBottom = (b >> 12) & 0x0F;
	o.icvLeft = (b >> 16) & 0x7F;
	o.icvRight = (b >> 23) & 0x7F;
	o.grbitDiag = (b >> 30) & 0x03;

	o.icvTop = c & 0x7F;
	o.icvBottom = (c >> 7) & 0x7F;
	o.icvDiag = (c >> 14) & 0x7F;
	o.dgDiag = (c >> 21) & 0x0F;

	o.icvFore = d & 0x7F;
	o.icvBack = (d >> 7) & 0x7F;
	o.fsxButton = (d >> 14) & 0x01;
	return o;
}
//function parse_CellXF(blob, length, opts) {return parse_CellStyleXF(blob,length,0, opts);}
//function parse_StyleXF(blob, length, opts) {return parse_CellStyleXF(blob,length,1, opts);}

/* [MS-XLS] 2.4.353 TODO: actually do this right */
function parse_XF(blob, length, opts) {
	var o = {};
	o.ifnt = blob.read_shift(2); o.numFmtId = blob.read_shift(2); o.flags = blob.read_shift(2);
	o.fStyle = (o.flags >> 2) & 0x01;
	length -= 6;
	o.data = parse_CellStyleXF(blob, length, o.fStyle, opts);
	return o;
}
function write_XF(data, ixfeP, opts, o) {
	var b5 = (opts && (opts.biff == 5));
	if(!o) o = new_buf(b5 ? 16 : 20);
	o.write_shift(2, 0);
	if(data.style) {
		o.write_shift(2, (data.numFmtId||0));
		o.write_shift(2, 0xFFF4);
	} else {
		o.write_shift(2, (data.numFmtId||0));
		o.write_shift(2, (ixfeP<<4));
	}
	o.write_shift(4, 0);
	o.write_shift(4, 0);
	if(!b5) o.write_shift(4, 0);
	o.write_shift(2, 0);
	return o;
}

/* [MS-XLS] 2.4.134 */
function parse_Guts(blob) {
	blob.l += 4;
	var out = [blob.read_shift(2), blob.read_shift(2)];
	if(out[0] !== 0) out[0]--;
	if(out[1] !== 0) out[1]--;
	if(out[0] > 7 || out[1] > 7) throw new Error("Bad Gutters: " + out.join("|"));
	return out;
}
function write_Guts(guts) {
	var o = new_buf(8);
	o.write_shift(4, 0);
	o.write_shift(2, guts[0] ? guts[0] + 1 : 0);
	o.write_shift(2, guts[1] ? guts[1] + 1 : 0);
	return o;
}

/* [MS-XLS] 2.4.24 */
function parse_BoolErr(blob, length, opts) {
	var cell = parse_XLSCell(blob, 6);
	if(opts.biff == 2) ++blob.l;
	var val = parse_Bes(blob, 2);
	cell.val = val;
	cell.t = (val === true || val === false) ? 'b' : 'e';
	return cell;
}
function write_BoolErr(R, C, v, os, opts, t) {
	var o = new_buf(8);
	write_XLSCell(R, C, os, o);
	write_Bes(v, t, o);
	return o;
}

/* [MS-XLS] 2.4.180 Number */
function parse_Number(blob) {
	var cell = parse_XLSCell(blob, 6);
	var xnum = parse_Xnum(blob, 8);
	cell.val = xnum;
	return cell;
}
function write_Number(R, C, v, os) {
	var o = new_buf(14);
	write_XLSCell(R, C, os, o);
	write_Xnum(v, o);
	return o;
}

var parse_XLHeaderFooter = parse_OptXLUnicodeString; // TODO: parse 2.4.136

/* [MS-XLS] 2.4.271 */
function parse_SupBook(blob, length, opts) {
	var end = blob.l + length;
	var ctab = blob.read_shift(2);
	var cch = blob.read_shift(2);
	opts.sbcch = cch;
	if(cch == 0x0401 || cch == 0x3A01) return [cch, ctab];
	if(cch < 0x01 || cch >0xff) throw new Error("Unexpected SupBook type: "+cch);
	var virtPath = parse_XLUnicodeStringNoCch(blob, cch);
	/* TODO: 2.5.277 Virtual Path */
	var rgst = [];
	while(end > blob.l) rgst.push(parse_XLUnicodeString(blob));
	return [cch, ctab, virtPath, rgst];
}

/* [MS-XLS] 2.4.105 TODO */
function parse_ExternName(blob, length, opts) {
	var flags = blob.read_shift(2);
	var body;
	var o = ({
		fBuiltIn: flags & 0x01,
		fWantAdvise: (flags >>> 1) & 0x01,
		fWantPict: (flags >>> 2) & 0x01,
		fOle: (flags >>> 3) & 0x01,
		fOleLink: (flags >>> 4) & 0x01,
		cf: (flags >>> 5) & 0x3FF,
		fIcon: flags >>> 15 & 0x01
	});
	if(opts.sbcch === 0x3A01) body = parse_AddinUdf(blob, length-2, opts);
	//else throw new Error("unsupported SupBook cch: " + opts.sbcch);
	o.body = body || blob.read_shift(length-2);
	if(typeof body === "string") o.Name = body;
	return o;
}

/* [MS-XLS] 2.4.150 TODO */
var XLSLblBuiltIn = [
	"_xlnm.Consolidate_Area",
	"_xlnm.Auto_Open",
	"_xlnm.Auto_Close",
	"_xlnm.Extract",
	"_xlnm.Database",
	"_xlnm.Criteria",
	"_xlnm.Print_Area",
	"_xlnm.Print_Titles",
	"_xlnm.Recorder",
	"_xlnm.Data_Form",
	"_xlnm.Auto_Activate",
	"_xlnm.Auto_Deactivate",
	"_xlnm.Sheet_Title",
	"_xlnm._FilterDatabase"
];
function parse_Lbl(blob, length, opts) {
	var target = blob.l + length;
	var flags = blob.read_shift(2);
	var chKey = blob.read_shift(1);
	var cch = blob.read_shift(1);
	var cce = blob.read_shift(opts && opts.biff == 2 ? 1 : 2);
	var itab = 0;
	if(!opts || opts.biff >= 5) {
		if(opts.biff != 5) blob.l += 2;
		itab = blob.read_shift(2);
		if(opts.biff == 5) blob.l += 2;
		blob.l += 4;
	}
	var name = parse_XLUnicodeStringNoCch(blob, cch, opts);
	if(flags & 0x20) name = XLSLblBuiltIn[name.charCodeAt(0)];
	var npflen = target - blob.l; if(opts && opts.biff == 2) --npflen;
	var rgce = target == blob.l || cce === 0 ? [] : parse_NameParsedFormula(blob, npflen, opts, cce);
	return {
		chKey: chKey,
		Name: name,
		itab: itab,
		rgce: rgce
	};
}

/* [MS-XLS] 2.4.106 TODO: verify filename encoding */
function parse_ExternSheet(blob, length, opts) {
	if(opts.biff < 8) return parse_BIFF5ExternSheet(blob, length, opts);
	var o = [], target = blob.l + length, len = blob.read_shift(opts.biff > 8 ? 4 : 2);
	while(len-- !== 0) o.push(parse_XTI(blob, opts.biff > 8 ? 12 : 6, opts));
		// [iSupBook, itabFirst, itabLast];
	if(blob.l != target) throw new Error("Bad ExternSheet: " + blob.l + " != " + target);
	return o;
}
function parse_BIFF5ExternSheet(blob, length, opts) {
	if(blob[blob.l + 1] == 0x03) blob[blob.l]++;
	var o = parse_ShortXLUnicodeString(blob, length, opts);
	return o.charCodeAt(0) == 0x03 ? o.slice(1) : o;
}

/* [MS-XLS] 2.4.176 TODO: check older biff */
function parse_NameCmt(blob, length, opts) {
	if(opts.biff < 8) { blob.l += length; return; }
	var cchName = blob.read_shift(2);
	var cchComment = blob.read_shift(2);
	var name = parse_XLUnicodeStringNoCch(blob, cchName, opts);
	var comment = parse_XLUnicodeStringNoCch(blob, cchComment, opts);
	return [name, comment];
}

/* [MS-XLS] 2.4.260 */
function parse_ShrFmla(blob, length, opts) {
	var ref = parse_RefU(blob, 6);
	blob.l++;
	var cUse = blob.read_shift(1);
	length -= 8;
	return [parse_SharedParsedFormula(blob, length, opts), cUse, ref];
}

/* [MS-XLS] 2.4.4 TODO */
function parse_Array(blob, length, opts) {
	var ref = parse_Ref(blob, 6);
	/* TODO: fAlwaysCalc */
	switch(opts.biff) {
		case 2: blob.l ++; length -= 7; break;
		case 3: case 4: blob.l += 2; length -= 8; break;
		default: blob.l += 6; length -= 12;
	}
	return [ref, parse_ArrayParsedFormula(blob, length, opts, ref)];
}

/* [MS-XLS] 2.4.173 */
function parse_MTRSettings(blob) {
	var fMTREnabled = blob.read_shift(4) !== 0x00;
	var fUserSetThreadCount = blob.read_shift(4) !== 0x00;
	var cUserThreadCount = blob.read_shift(4);
	return [fMTREnabled, fUserSetThreadCount, cUserThreadCount];
}

/* [MS-XLS] 2.5.186 TODO: BIFF5 */
function parse_NoteSh(blob, length, opts) {
	if(opts.biff < 8) return;
	var row = blob.read_shift(2), col = blob.read_shift(2);
	var flags = blob.read_shift(2), idObj = blob.read_shift(2);
	var stAuthor = parse_XLUnicodeString2(blob, 0, opts);
	if(opts.biff < 8) blob.read_shift(1);
	return [{r:row,c:col}, stAuthor, idObj, flags];
}

/* [MS-XLS] 2.4.179 */
function parse_Note(blob, length, opts) {
	/* TODO: Support revisions */
	return parse_NoteSh(blob, length, opts);
}

/* [MS-XLS] 2.4.168 */
function parse_MergeCells(blob, length) {
	var merges = [];
	var cmcs = blob.read_shift(2);
	while (cmcs--) merges.push(parse_Ref8U(blob,length));
	return merges;
}
function write_MergeCells(merges) {
	var o = new_buf(2 + merges.length * 8);
	o.write_shift(2, merges.length);
	for(var i = 0; i < merges.length; ++i) write_Ref8U(merges[i], o);
	return o;
}

/* [MS-XLS] 2.4.181 TODO: parse all the things! */
function parse_Obj(blob, length, opts) {
	if(opts && opts.biff < 8) return parse_BIFF5Obj(blob, length, opts);
	var cmo = parse_FtCmo(blob, 22); // id, ot, flags
	var fts = parse_FtArray(blob, length-22, cmo[1]);
	return { cmo: cmo, ft:fts };
}
/* from older spec */
var parse_BIFF5OT = [];
parse_BIFF5OT[0x08] = function(blob, length) {
	var tgt = blob.l + length;
	blob.l += 10; // todo
	var cf = blob.read_shift(2);
	blob.l += 4;
	blob.l += 2; //var cbPictFmla = blob.read_shift(2);
	blob.l += 2;
	blob.l += 2; //var grbit = blob.read_shift(2);
	blob.l += 4;
	var cchName = blob.read_shift(1);
	blob.l += cchName; // TODO: stName
	blob.l = tgt; // TODO: fmla
	return { fmt:cf };
};

function parse_BIFF5Obj(blob, length, opts) {
	blob.l += 4; //var cnt = blob.read_shift(4);
	var ot = blob.read_shift(2);
	var id = blob.read_shift(2);
	var grbit = blob.read_shift(2);
	blob.l += 2; //var colL = blob.read_shift(2);
	blob.l += 2; //var dxL = blob.read_shift(2);
	blob.l += 2; //var rwT = blob.read_shift(2);
	blob.l += 2; //var dyT = blob.read_shift(2);
	blob.l += 2; //var colR = blob.read_shift(2);
	blob.l += 2; //var dxR = blob.read_shift(2);
	blob.l += 2; //var rwB = blob.read_shift(2);
	blob.l += 2; //var dyB = blob.read_shift(2);
	blob.l += 2; //var cbMacro = blob.read_shift(2);
	blob.l += 6;
	length -= 36;
	var fts = [];
	fts.push((parse_BIFF5OT[ot]||parsenoop)(blob, length, opts));
	return { cmo: [id, ot, grbit], ft:fts };
}

/* [MS-XLS] 2.4.329 TODO: parse properly */
function parse_TxO(blob, length, opts) {
	var s = blob.l;
	var texts = "";
try {
	blob.l += 4;
	var ot = (opts.lastobj||{cmo:[0,0]}).cmo[1];
	var controlInfo; // eslint-disable-line no-unused-vars
	if([0,5,7,11,12,14].indexOf(ot) == -1) blob.l += 6;
	else controlInfo = parse_ControlInfo(blob, 6, opts);
	var cchText = blob.read_shift(2);
	/*var cbRuns = */blob.read_shift(2);
	/*var ifntEmpty = */parseuint16(blob, 2);
	var len = blob.read_shift(2);
	blob.l += len;
	//var fmla = parse_ObjFmla(blob, s + length - blob.l);

	for(var i = 1; i < blob.lens.length-1; ++i) {
		if(blob.l-s != blob.lens[i]) throw new Error("TxO: bad continue record");
		var hdr = blob[blob.l];
		var t = parse_XLUnicodeStringNoCch(blob, blob.lens[i+1]-blob.lens[i]-1);
		texts += t;
		if(texts.length >= (hdr ? cchText : 2*cchText)) break;
	}
	if(texts.length !== cchText && texts.length !== cchText*2) {
		throw new Error("cchText: " + cchText + " != " + texts.length);
	}

	blob.l = s + length;
	/* [MS-XLS] 2.5.272 TxORuns */
//	var rgTxoRuns = [];
//	for(var j = 0; j != cbRuns/8-1; ++j) blob.l += 8;
//	var cchText2 = blob.read_shift(2);
//	if(cchText2 !== cchText) throw new Error("TxOLastRun mismatch: " + cchText2 + " " + cchText);
//	blob.l += 6;
//	if(s + length != blob.l) throw new Error("TxO " + (s + length) + ", at " + blob.l);
	return { t: texts };
} catch(e) { blob.l = s + length; return { t: texts }; }
}

/* [MS-XLS] 2.4.140 */
function parse_HLink(blob, length) {
	var ref = parse_Ref8U(blob, 8);
	blob.l += 16; /* CLSID */
	var hlink = parse_Hyperlink(blob, length-24);
	return [ref, hlink];
}
function write_HLink(hl) {
	var O = new_buf(24);
	var ref = decode_cell(hl[0]);
	O.write_shift(2, ref.r); O.write_shift(2, ref.r);
	O.write_shift(2, ref.c); O.write_shift(2, ref.c);
	var clsid = "d0 c9 ea 79 f9 ba ce 11 8c 82 00 aa 00 4b a9 0b".split(" ");
	for(var i = 0; i < 16; ++i) O.write_shift(1, parseInt(clsid[i], 16));
	return bconcat([O, write_Hyperlink(hl[1])]);
}


/* [MS-XLS] 2.4.141 */
function parse_HLinkTooltip(blob, length) {
	blob.read_shift(2);
	var ref = parse_Ref8U(blob, 8);
	var wzTooltip = blob.read_shift((length-10)/2, 'dbcs-cont');
	wzTooltip = wzTooltip.replace(chr0,"");
	return [ref, wzTooltip];
}
function write_HLinkTooltip(hl) {
	var TT = hl[1].Tooltip;
	var O = new_buf(10 + 2 * (TT.length + 1));
	O.write_shift(2, 0x0800);
	var ref = decode_cell(hl[0]);
	O.write_shift(2, ref.r); O.write_shift(2, ref.r);
	O.write_shift(2, ref.c); O.write_shift(2, ref.c);
	for(var i = 0; i < TT.length; ++i) O.write_shift(2, TT.charCodeAt(i));
	O.write_shift(2, 0);
	return O;
}

/* [MS-XLS] 2.4.63 */
function parse_Country(blob) {
	var o = [0,0], d;
	d = blob.read_shift(2); o[0] = CountryEnum[d] || d;
	d = blob.read_shift(2); o[1] = CountryEnum[d] || d;
	return o;
}
function write_Country(o) {
	if(!o) o = new_buf(4);
	o.write_shift(2, 0x01);
	o.write_shift(2, 0x01);
	return o;
}

/* [MS-XLS] 2.4.50 ClrtClient */
function parse_ClrtClient(blob) {
	var ccv = blob.read_shift(2);
	var o = [];
	while(ccv-->0) o.push(parse_LongRGB(blob, 8));
	return o;
}

/* [MS-XLS] 2.4.188 */
function parse_Palette(blob) {
	var ccv = blob.read_shift(2);
	var o = [];
	while(ccv-->0) o.push(parse_LongRGB(blob, 8));
	return o;
}

/* [MS-XLS] 2.4.354 */
function parse_XFCRC(blob) {
	blob.l += 2;
	var o = {cxfs:0, crc:0};
	o.cxfs = blob.read_shift(2);
	o.crc = blob.read_shift(4);
	return o;
}

/* [MS-XLS] 2.4.53 TODO: parse flags */
/* [MS-XLSB] 2.4.323 TODO: parse flags */
function parse_ColInfo(blob, length, opts) {
	if(!opts.cellStyles) return parsenoop(blob, length);
	var w = opts && opts.biff >= 12 ? 4 : 2;
	var colFirst = blob.read_shift(w);
	var colLast = blob.read_shift(w);
	var coldx = blob.read_shift(w);
	var ixfe = blob.read_shift(w);
	var flags = blob.read_shift(2);
	if(w == 2) blob.l += 2;
	return {s:colFirst, e:colLast, w:coldx, ixfe:ixfe, flags:flags};
}

/* [MS-XLS] 2.4.257 */
function parse_Setup(blob, length) {
	var o = {};
	if(length < 32) return o;
	blob.l += 16;
	o.header = parse_Xnum(blob, 8);
	o.footer = parse_Xnum(blob, 8);
	blob.l += 2;
	return o;
}

/* [MS-XLS] 2.4.261 */
function parse_ShtProps(blob, length, opts) {
	var def = {area:false};
	if(opts.biff != 5) { blob.l += length; return def; }
	var d = blob.read_shift(1); blob.l += 3;
	if((d & 0x10)) def.area = true;
	return def;
}

/* [MS-XLS] 2.4.241 */
function write_RRTabId(n) {
	var out = new_buf(2 * n);
	for(var i = 0; i < n; ++i) out.write_shift(2, i+1);
	return out;
}

var parse_Blank = parse_XLSCell; /* [MS-XLS] 2.4.20 Just the cell */
var parse_Scl = parseuint16a; /* [MS-XLS] 2.4.247 num, den */
var parse_String = parse_XLUnicodeString; /* [MS-XLS] 2.4.268 */

/* --- Specific to versions before BIFF8 --- */
function parse_ImData(blob) {
	var cf = blob.read_shift(2);
	var env = blob.read_shift(2);
	var lcb = blob.read_shift(4);
	var o = {fmt:cf, env:env, len:lcb, data:blob.slice(blob.l,blob.l+lcb)};
	blob.l += lcb;
	return o;
}

/* BIFF2_??? where ??? is the name from [XLS] */
function parse_BIFF2STR(blob, length, opts) {
	var cell = parse_XLSCell(blob, 6);
	++blob.l;
	var str = parse_XLUnicodeString2(blob, length-7, opts);
	cell.t = 'str';
	cell.val = str;
	return cell;
}

function parse_BIFF2NUM(blob) {
	var cell = parse_XLSCell(blob, 6);
	++blob.l;
	var num = parse_Xnum(blob, 8);
	cell.t = 'n';
	cell.val = num;
	return cell;
}
function write_BIFF2NUM(r, c, val) {
	var out = new_buf(15);
	write_BIFF2Cell(out, r, c);
	out.write_shift(8, val, 'f');
	return out;
}

function parse_BIFF2INT(blob) {
	var cell = parse_XLSCell(blob, 6);
	++blob.l;
	var num = blob.read_shift(2);
	cell.t = 'n';
	cell.val = num;
	return cell;
}
function write_BIFF2INT(r, c, val) {
	var out = new_buf(9);
	write_BIFF2Cell(out, r, c);
	out.write_shift(2, val);
	return out;
}

function parse_BIFF2STRING(blob) {
	var cch = blob.read_shift(1);
	if(cch === 0) { blob.l++; return ""; }
	return blob.read_shift(cch, 'sbcs-cont');
}

/* TODO: convert to BIFF8 font struct */
function parse_BIFF2FONTXTRA(blob, length) {
	blob.l += 6; // unknown
	blob.l += 2; // font weight "bls"
	blob.l += 1; // charset
	blob.l += 3; // unknown
	blob.l += 1; // font family
	blob.l += length - 13;
}

/* TODO: parse rich text runs */
function parse_RString(blob, length, opts) {
	var end = blob.l + length;
	var cell = parse_XLSCell(blob, 6);
	var cch = blob.read_shift(2);
	var str = parse_XLUnicodeStringNoCch(blob, cch, opts);
	blob.l = end;
	cell.t = 'str';
	cell.val = str;
	return cell;
}
/* from js-harb (C) 2014-present  SheetJS */
var DBF = (function() {
var dbf_codepage_map = {
	/* Code Pages Supported by Visual FoxPro */
0x01:   437,           0x02:   850,
0x03:  1252,           0x04: 10000,
0x64:   852,           0x65:   866,
0x66:   865,           0x67:   861,
0x68:   895,           0x69:   620,
0x6A:   737,           0x6B:   857,
0x78:   950,           0x79:   949,
0x7A:   936,           0x7B:   932,
0x7C:   874,           0x7D:  1255,
0x7E:  1256,           0x96: 10007,
0x97: 10029,           0x98: 10006,
0xC8:  1250,           0xC9:  1251,
0xCA:  1254,           0xCB:  1253,

	/* shapefile DBF extension */
0x00: 20127,           0x08:   865,
0x09:   437,           0x0A:   850,
0x0B:   437,           0x0D:   437,
0x0E:   850,           0x0F:   437,
0x10:   850,           0x11:   437,
0x12:   850,           0x13:   932,
0x14:   850,           0x15:   437,
0x16:   850,           0x17:   865,
0x18:   437,           0x19:   437,
0x1A:   850,           0x1B:   437,
0x1C:   863,           0x1D:   850,
0x1F:   852,           0x22:   852,
0x23:   852,           0x24:   860,
0x25:   850,           0x26:   866,
0x37:   850,           0x40:   852,
0x4D:   936,           0x4E:   949,
0x4F:   950,           0x50:   874,
0x57:  1252,           0x58:  1252,
0x59:  1252,

0xFF: 16969
};

/* TODO: find an actual specification */
function dbf_to_aoa(buf, opts) {
	var out = [];
	/* TODO: browser based */
	var d = (new_raw_buf(1));
	switch(opts.type) {
		case 'base64': d = s2a(Base64.decode(buf)); break;
		case 'binary': d = s2a(buf); break;
		case 'buffer':
		case 'array': d = buf; break;
	}
	prep_blob(d, 0);
	/* header */
	var ft = d.read_shift(1);
	var memo = false;
	var vfp = false, l7 = false;
	switch(ft) {
		case 0x02: case 0x03: break;
		case 0x30: vfp = true; memo = true; break;
		case 0x31: vfp = true; break;
		case 0x83: memo = true; break;
		case 0x8B: memo = true; break;
		case 0x8C: memo = true; l7 = true; break;
		case 0xF5: memo = true; break;
		default: throw new Error("DBF Unsupported Version: " + ft.toString(16));
	}
	var /*filedate = new Date(),*/ nrow = 0, fpos = 0;
	if(ft == 0x02) nrow = d.read_shift(2);
	/*filedate = new Date(d.read_shift(1) + 1900, d.read_shift(1) - 1, d.read_shift(1));*/d.l += 3;
	if(ft != 0x02) nrow = d.read_shift(4);
	if(ft != 0x02) fpos = d.read_shift(2);
	var rlen = d.read_shift(2);

	var /*flags = 0,*/ current_cp = 1252;
	if(ft != 0x02) {
	d.l+=16;
	/*flags = */d.read_shift(1);
	//if(memo && ((flags & 0x02) === 0)) throw new Error("DBF Flags " + flags.toString(16) + " ft " + ft.toString(16));

	/* codepage present in FoxPro */
	if(d[d.l] !== 0) current_cp = dbf_codepage_map[d[d.l]];
	d.l+=1;

	d.l+=2;
	}
	if(l7) d.l += 36;
var fields = [], field = ({});
	var hend = fpos - 10 - (vfp ? 264 : 0), ww = l7 ? 32 : 11;
	while(ft == 0x02 ? d.l < d.length && d[d.l] != 0x0d: d.l < hend) {
		field = ({});
		field.name = cptable.utils.decode(current_cp, d.slice(d.l, d.l+ww)).replace(/[\u0000\r\n].*$/g,"");
		d.l += ww;
		field.type = String.fromCharCode(d.read_shift(1));
		if(ft != 0x02 && !l7) field.offset = d.read_shift(4);
		field.len = d.read_shift(1);
		if(ft == 0x02) field.offset = d.read_shift(2);
		field.dec = d.read_shift(1);
		if(field.name.length) fields.push(field);
		if(ft != 0x02) d.l += l7 ? 13 : 14;
		switch(field.type) {
			case 'B': // VFP Double
				if((!vfp || field.len != 8) && opts.WTF) console.log('Skipping ' + field.name + ':' + field.type);
				break;
			case 'G': // General
			case 'P': // Picture
				if(opts.WTF) console.log('Skipping ' + field.name + ':' + field.type);
				break;
			case 'C': // character
			case 'D': // date
			case 'F': // floating point
			case 'I': // long
			case 'L': // boolean
			case 'M': // memo
			case 'N': // number
			case 'O': // double
			case 'T': // datetime
			case 'Y': // currency
			case '0': // VFP _NullFlags
			case '@': // timestamp
			case '+': // autoincrement
				break;
			default: throw new Error('Unknown Field Type: ' + field.type);
		}
	}
	if(d[d.l] !== 0x0D) d.l = fpos-1;
	else if(ft == 0x02) d.l = 0x209;
	if(ft != 0x02) {
		if(d.read_shift(1) !== 0x0D) throw new Error("DBF Terminator not found " + d.l + " " + d[d.l]);
		d.l = fpos;
	}
	/* data */
	var R = 0, C = 0;
	out[0] = [];
	for(C = 0; C != fields.length; ++C) out[0][C] = fields[C].name;
	while(nrow-- > 0) {
		if(d[d.l] === 0x2A) { d.l+=rlen; continue; }
		++d.l;
		out[++R] = []; C = 0;
		for(C = 0; C != fields.length; ++C) {
			var dd = d.slice(d.l, d.l+fields[C].len); d.l+=fields[C].len;
			prep_blob(dd, 0);
			var s = cptable.utils.decode(current_cp, dd);
			switch(fields[C].type) {
				case 'C':
					out[R][C] = cptable.utils.decode(current_cp, dd);
					out[R][C] = out[R][C].trim();
					break;
				case 'D':
					if(s.length === 8) out[R][C] = new Date(+s.slice(0,4), +s.slice(4,6)-1, +s.slice(6,8));
					else out[R][C] = s;
					break;
				case 'F': out[R][C] = parseFloat(s.trim()); break;
				case '+': case 'I': out[R][C] = l7 ? dd.read_shift(-4, 'i') ^ 0x80000000 : dd.read_shift(4, 'i'); break;
				case 'L': switch(s.toUpperCase()) {
					case 'Y': case 'T': out[R][C] = true; break;
					case 'N': case 'F': out[R][C] = false; break;
					case ' ': case '?': out[R][C] = false; break; /* NOTE: technically uninitialized */
					default: throw new Error("DBF Unrecognized L:|" + s + "|");
					} break;
				case 'M': /* TODO: handle memo files */
					if(!memo) throw new Error("DBF Unexpected MEMO for type " + ft.toString(16));
					out[R][C] = "##MEMO##" + (l7 ? parseInt(s.trim(), 10): dd.read_shift(4));
					break;
				case 'N': out[R][C] = +s.replace(/\u0000/g,"").trim(); break;
				case '@': out[R][C] = new Date(dd.read_shift(-8, 'f') - 0x388317533400); break;
				case 'T': out[R][C] = new Date((dd.read_shift(4) - 0x253D8C) * 0x5265C00 + dd.read_shift(4)); break;
				case 'Y': out[R][C] = dd.read_shift(4,'i')/1e4; break;
				case 'O': out[R][C] = -dd.read_shift(-8, 'f'); break;
				case 'B': if(vfp && fields[C].len == 8) { out[R][C] = dd.read_shift(8,'f'); break; }
					/* falls through */
				case 'G': case 'P': dd.l += fields[C].len; break;
				case '0':
					if(fields[C].name === '_NullFlags') break;
					/* falls through */
				default: throw new Error("DBF Unsupported data type " + fields[C].type);
			}
		}
	}
	if(ft != 0x02) if(d.l < d.length && d[d.l++] != 0x1A) throw new Error("DBF EOF Marker missing " + (d.l-1) + " of " + d.length + " " + d[d.l-1].toString(16));
	if(opts && opts.sheetRows) out = out.slice(0, opts.sheetRows);
	return out;
}

function dbf_to_sheet(buf, opts) {
	var o = opts || {};
	if(!o.dateNF) o.dateNF = "yyyymmdd";
	return aoa_to_sheet(dbf_to_aoa(buf, o), o);
}

function dbf_to_workbook(buf, opts) {
	try { return sheet_to_workbook(dbf_to_sheet(buf, opts), opts); }
	catch(e) { if(opts && opts.WTF) throw e; }
	return ({SheetNames:[],Sheets:{}});
}

var _RLEN = { 'B': 8, 'C': 250, 'L': 1, 'D': 8, '?': 0, '': 0 };
function sheet_to_dbf(ws, opts) {
	var o = opts || {};
	if(o.type == "string") throw new Error("Cannot write DBF to JS string");
	var ba = buf_array();
	var aoa = sheet_to_json(ws, {header:1, cellDates:true});
	var headers = aoa[0], data = aoa.slice(1);
	var i = 0, j = 0, hcnt = 0, rlen = 1;
	for(i = 0; i < headers.length; ++i) {
		if(i == null) continue;
		++hcnt;
		if(typeof headers[i] === 'number') headers[i] = headers[i].toString(10);
		if(typeof headers[i] !== 'string') throw new Error("DBF Invalid column name " + headers[i] + " |" + (typeof headers[i]) + "|");
		if(headers.indexOf(headers[i]) !== i) for(j=0; j<1024;++j)
			if(headers.indexOf(headers[i] + "_" + j) == -1) { headers[i] += "_" + j; break; }
	}
	var range = safe_decode_range(ws['!ref']);
	var coltypes = [];
	for(i = 0; i <= range.e.c - range.s.c; ++i) {
		var col = [];
		for(j=0; j < data.length; ++j) {
			if(data[j][i] != null) col.push(data[j][i]);
		}
		if(col.length == 0 || headers[i] == null) { coltypes[i] = '?'; continue; }
		var guess = '', _guess = '';
		for(j = 0; j < col.length; ++j) {
			switch(typeof col[j]) {
				/* TODO: check if L2 compat is desired */
				case 'number': _guess = 'B'; break;
				case 'string': _guess = 'C'; break;
				case 'boolean': _guess = 'L'; break;
				case 'object': _guess = col[j] instanceof Date ? 'D' : 'C'; break;
				default: _guess = 'C';
			}
			guess = guess && guess != _guess ? 'C' : _guess;
			if(guess == 'C') break;
		}
		rlen += _RLEN[guess] || 0;
		coltypes[i] = guess;
	}

	var h = ba.next(32);
	h.write_shift(4, 0x13021130);
	h.write_shift(4, data.length);
	h.write_shift(2, 296 + 32 * hcnt);
	h.write_shift(2, rlen);
	for(i=0; i < 4; ++i) h.write_shift(4, 0);
	h.write_shift(4, 0x00000300); // TODO: CP

	for(i = 0, j = 0; i < headers.length; ++i) {
		if(headers[i] == null) continue;
		var hf = ba.next(32);
		var _f = (headers[i].slice(-10) + "\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00").slice(0, 11);
		hf.write_shift(1, _f, "sbcs");
		hf.write_shift(1, coltypes[i] == '?' ? 'C' : coltypes[i], "sbcs");
		hf.write_shift(4, j);
		hf.write_shift(1, _RLEN[coltypes[i]] || 0);
		hf.write_shift(1, 0);
		hf.write_shift(1, 0x02);
		hf.write_shift(4, 0);
		hf.write_shift(1, 0);
		hf.write_shift(4, 0);
		hf.write_shift(4, 0);
		j += _RLEN[coltypes[i]] || 0;
	}

	var hb = ba.next(264);
	hb.write_shift(4, 0x0000000D);
	for(i=0; i < 65;++i) hb.write_shift(4, 0x00000000);
	for(i=0; i < data.length; ++i) {
		var rout = ba.next(rlen);
		rout.write_shift(1, 0);
		for(j=0; j<headers.length; ++j) {
			if(headers[j] == null) continue;
			switch(coltypes[j]) {
				case 'L': rout.write_shift(1, data[i][j] == null ? 0x3F : data[i][j] ? 0x54 : 0x46); break;
				case 'B': rout.write_shift(8, data[i][j]||0, 'f'); break;
				case 'D':
					if(!data[i][j]) rout.write_shift(8, "00000000", "sbcs");
					else {
						rout.write_shift(4, ("0000"+data[i][j].getFullYear()).slice(-4), "sbcs");
						rout.write_shift(2, ("00"+(data[i][j].getMonth()+1)).slice(-2), "sbcs");
						rout.write_shift(2, ("00"+data[i][j].getDate()).slice(-2), "sbcs");
					} break;
				case 'C':
					var _s = String(data[i][j]||"");
					rout.write_shift(1, _s, "sbcs");
					for(hcnt=0; hcnt < 250-_s.length; ++hcnt) rout.write_shift(1, 0x20); break;
			}
		}
		// data
	}
	ba.next(1).write_shift(1, 0x1A);
	return ba.end();
}
	return {
		to_workbook: dbf_to_workbook,
		to_sheet: dbf_to_sheet,
		from_sheet: sheet_to_dbf
	};
})();

var SYLK = (function() {
	/* TODO: find an actual specification */
	function sylk_to_aoa(d, opts) {
		switch(opts.type) {
			case 'base64': return sylk_to_aoa_str(Base64.decode(d), opts);
			case 'binary': return sylk_to_aoa_str(d, opts);
			case 'buffer': return sylk_to_aoa_str(d.toString('binary'), opts);
			case 'array': return sylk_to_aoa_str(cc2str(d), opts);
		}
		throw new Error("Unrecognized type " + opts.type);
	}
	function sylk_to_aoa_str(str, opts) {
		var records = str.split(/[\n\r]+/), R = -1, C = -1, ri = 0, rj = 0, arr = [];
		var formats = [];
		var next_cell_format = null;
		var sht = {}, rowinfo = [], colinfo = [], cw = [];
		var Mval = 0, j;
		for (; ri !== records.length; ++ri) {
			Mval = 0;
			var rstr=records[ri].trim();
			var record=rstr.replace(/;;/g, "\u0001").split(";").map(function(x) { return x.replace(/\u0001/g, ";"); });
			var RT=record[0], val;
			if(rstr.length > 0) switch(RT) {
			case 'ID': break; /* header */
			case 'E': break; /* EOF */
			case 'B': break; /* dimensions */
			case 'O': break; /* options? */
			case 'P':
				if(record[1].charAt(0) == 'P')
					formats.push(rstr.slice(3).replace(/;;/g, ";"));
				break;
			case 'C':
			var C_seen_K = false, C_seen_X = false;
			for(rj=1; rj<record.length; ++rj) switch(record[rj].charAt(0)) {
				case 'X': C = parseInt(record[rj].slice(1))-1; C_seen_X = true; break;
				case 'Y':
					R = parseInt(record[rj].slice(1))-1; if(!C_seen_X) C = 0;
					for(j = arr.length; j <= R; ++j) arr[j] = [];
					break;
				case 'K':
					val = record[rj].slice(1);
					if(val.charAt(0) === '"') val = val.slice(1,val.length - 1);
					else if(val === 'TRUE') val = true;
					else if(val === 'FALSE') val = false;
					else if(!isNaN(fuzzynum(val))) {
						val = fuzzynum(val);
						if(next_cell_format !== null && SSF.is_date(next_cell_format)) val = numdate(val);
					} else if(!isNaN(fuzzydate(val).getDate())) {
						val = parseDate(val);
					}
					if(typeof cptable !== 'undefined' && typeof val == "string" && ((opts||{}).type != "string") && (opts||{}).codepage) val = cptable.utils.decode(opts.codepage, val);
					C_seen_K = true;
					break;
				case 'E':
					var formula = rc_to_a1(record[rj].slice(1), {r:R,c:C});
					arr[R][C] = [arr[R][C], formula];
					break;
				default: if(opts && opts.WTF) throw new Error("SYLK bad record " + rstr);
			}
			if(C_seen_K) { arr[R][C] = val; next_cell_format = null; }
			break;
			case 'F':
			var F_seen = 0;
			for(rj=1; rj<record.length; ++rj) switch(record[rj].charAt(0)) {
				case 'X': C = parseInt(record[rj].slice(1))-1; ++F_seen; break;
				case 'Y':
					R = parseInt(record[rj].slice(1))-1; /*C = 0;*/
					for(j = arr.length; j <= R; ++j) arr[j] = [];
					break;
				case 'M': Mval = parseInt(record[rj].slice(1)) / 20; break;
				case 'F': break; /* ??? */
				case 'G': break; /* hide grid */
				case 'P':
					next_cell_format = formats[parseInt(record[rj].slice(1))];
					break;
				case 'S': break; /* cell style */
				case 'D': break; /* column */
				case 'N': break; /* font */
				case 'W':
					cw = record[rj].slice(1).split(" ");
					for(j = parseInt(cw[0], 10); j <= parseInt(cw[1], 10); ++j) {
						Mval = parseInt(cw[2], 10);
						colinfo[j-1] = Mval === 0 ? {hidden:true}: {wch:Mval}; process_col(colinfo[j-1]);
					} break;
				case 'C': /* default column format */
					C = parseInt(record[rj].slice(1))-1;
					if(!colinfo[C]) colinfo[C] = {};
					break;
				case 'R': /* row properties */
					R = parseInt(record[rj].slice(1))-1;
					if(!rowinfo[R]) rowinfo[R] = {};
					if(Mval > 0) { rowinfo[R].hpt = Mval; rowinfo[R].hpx = pt2px(Mval); }
					else if(Mval === 0) rowinfo[R].hidden = true;
					break;
				default: if(opts && opts.WTF) throw new Error("SYLK bad record " + rstr);
			}
			if(F_seen < 1) next_cell_format = null; break;
			default: if(opts && opts.WTF) throw new Error("SYLK bad record " + rstr);
			}
		}
		if(rowinfo.length > 0) sht['!rows'] = rowinfo;
		if(colinfo.length > 0) sht['!cols'] = colinfo;
		if(opts && opts.sheetRows) arr = arr.slice(0, opts.sheetRows);
		return [arr, sht];
	}

	function sylk_to_sheet(d, opts) {
		var aoasht = sylk_to_aoa(d, opts);
		var aoa = aoasht[0], ws = aoasht[1];
		var o = aoa_to_sheet(aoa, opts);
		keys(ws).forEach(function(k) { o[k] = ws[k]; });
		return o;
	}

	function sylk_to_workbook(d, opts) { return sheet_to_workbook(sylk_to_sheet(d, opts), opts); }

	function write_ws_cell_sylk(cell, ws, R, C) {
		var o = "C;Y" + (R+1) + ";X" + (C+1) + ";K";
		switch(cell.t) {
			case 'n':
				o += (cell.v||0);
				if(cell.f && !cell.F) o += ";E" + a1_to_rc(cell.f, {r:R, c:C}); break;
			case 'b': o += cell.v ? "TRUE" : "FALSE"; break;
			case 'e': o += cell.w || cell.v; break;
			case 'd': o += '"' + (cell.w || cell.v) + '"'; break;
			case 's': o += '"' + cell.v.replace(/"/g,"") + '"'; break;
		}
		return o;
	}

	function write_ws_cols_sylk(out, cols) {
		cols.forEach(function(col, i) {
			var rec = "F;W" + (i+1) + " " + (i+1) + " ";
			if(col.hidden) rec += "0";
			else {
				if(typeof col.width == 'number') col.wpx = width2px(col.width);
				if(typeof col.wpx == 'number') col.wch = px2char(col.wpx);
				if(typeof col.wch == 'number') rec += Math.round(col.wch);
			}
			if(rec.charAt(rec.length - 1) != " ") out.push(rec);
		});
	}

	function write_ws_rows_sylk(out, rows) {
		rows.forEach(function(row, i) {
			var rec = "F;";
			if(row.hidden) rec += "M0;";
			else if(row.hpt) rec += "M" + 20 * row.hpt + ";";
			else if(row.hpx) rec += "M" + 20 * px2pt(row.hpx) + ";";
			if(rec.length > 2) out.push(rec + "R" + (i+1));
		});
	}

	function sheet_to_sylk(ws, opts) {
		var preamble = ["ID;PWXL;N;E"], o = [];
		var r = safe_decode_range(ws['!ref']), cell;
		var dense = Array.isArray(ws);
		var RS = "\r\n";

		preamble.push("P;PGeneral");
		preamble.push("F;P0;DG0G8;M255");
		if(ws['!cols']) write_ws_cols_sylk(preamble, ws['!cols']);
		if(ws['!rows']) write_ws_rows_sylk(preamble, ws['!rows']);

		preamble.push("B;Y" + (r.e.r - r.s.r + 1) + ";X" + (r.e.c - r.s.c + 1) + ";D" + [r.s.c,r.s.r,r.e.c,r.e.r].join(" "));
		for(var R = r.s.r; R <= r.e.r; ++R) {
			for(var C = r.s.c; C <= r.e.c; ++C) {
				var coord = encode_cell({r:R,c:C});
				cell = dense ? (ws[R]||[])[C]: ws[coord];
				if(!cell || (cell.v == null && (!cell.f || cell.F))) continue;
				o.push(write_ws_cell_sylk(cell, ws, R, C, opts));
			}
		}
		return preamble.join(RS) + RS + o.join(RS) + RS + "E" + RS;
	}

	return {
		to_workbook: sylk_to_workbook,
		to_sheet: sylk_to_sheet,
		from_sheet: sheet_to_sylk
	};
})();

var DIF = (function() {
	function dif_to_aoa(d, opts) {
		switch(opts.type) {
			case 'base64': return dif_to_aoa_str(Base64.decode(d), opts);
			case 'binary': return dif_to_aoa_str(d, opts);
			case 'buffer': return dif_to_aoa_str(d.toString('binary'), opts);
			case 'array': return dif_to_aoa_str(cc2str(d), opts);
		}
		throw new Error("Unrecognized type " + opts.type);
	}
	function dif_to_aoa_str(str, opts) {
		var records = str.split('\n'), R = -1, C = -1, ri = 0, arr = [];
		for (; ri !== records.length; ++ri) {
			if (records[ri].trim() === 'BOT') { arr[++R] = []; C = 0; continue; }
			if (R < 0) continue;
			var metadata = records[ri].trim().split(",");
			var type = metadata[0], value = metadata[1];
			++ri;
			var data = records[ri].trim();
			switch (+type) {
				case -1:
					if (data === 'BOT') { arr[++R] = []; C = 0; continue; }
					else if (data !== 'EOD') throw new Error("Unrecognized DIF special command " + data);
					break;
				case 0:
					if(data === 'TRUE') arr[R][C] = true;
					else if(data === 'FALSE') arr[R][C] = false;
					else if(!isNaN(fuzzynum(value))) arr[R][C] = fuzzynum(value);
					else if(!isNaN(fuzzydate(value).getDate())) arr[R][C] = parseDate(value);
					else arr[R][C] = value;
					++C; break;
				case 1:
					data = data.slice(1,data.length-1);
					arr[R][C++] = data !== '' ? data : null;
					break;
			}
			if (data === 'EOD') break;
		}
		if(opts && opts.sheetRows) arr = arr.slice(0, opts.sheetRows);
		return arr;
	}

	function dif_to_sheet(str, opts) { return aoa_to_sheet(dif_to_aoa(str, opts), opts); }
	function dif_to_workbook(str, opts) { return sheet_to_workbook(dif_to_sheet(str, opts), opts); }

	var sheet_to_dif = (function() {
		var push_field = function pf(o, topic, v, n, s) {
			o.push(topic);
			o.push(v + "," + n);
			o.push('"' + s.replace(/"/g,'""') + '"');
		};
		var push_value = function po(o, type, v, s) {
			o.push(type + "," + v);
			o.push(type == 1 ? '"' + s.replace(/"/g,'""') + '"' : s);
		};
		return function sheet_to_dif(ws) {
			var o = [];
			var r = safe_decode_range(ws['!ref']), cell;
			var dense = Array.isArray(ws);
			push_field(o, "TABLE", 0, 1, "sheetjs");
			push_field(o, "VECTORS", 0, r.e.r - r.s.r + 1,"");
			push_field(o, "TUPLES", 0, r.e.c - r.s.c + 1,"");
			push_field(o, "DATA", 0, 0,"");
			for(var R = r.s.r; R <= r.e.r; ++R) {
				push_value(o, -1, 0, "BOT");
				for(var C = r.s.c; C <= r.e.c; ++C) {
					var coord = encode_cell({r:R,c:C});
					cell = dense ? (ws[R]||[])[C] : ws[coord];
					if(!cell) { push_value(o, 1, 0, ""); continue;}
					switch(cell.t) {
						case 'n':
							var val = DIF_XL ? cell.w : cell.v;
							if(!val && cell.v != null) val = cell.v;
							if(val == null) {
								if(DIF_XL && cell.f && !cell.F) push_value(o, 1, 0, "=" + cell.f);
								else push_value(o, 1, 0, "");
							}
							else push_value(o, 0, val, "V");
							break;
						case 'b':
							push_value(o, 0, cell.v ? 1 : 0, cell.v ? "TRUE" : "FALSE");
							break;
						case 's':
							push_value(o, 1, 0, (!DIF_XL || isNaN(cell.v)) ? cell.v : '="' + cell.v + '"');
							break;
						case 'd':
							if(!cell.w) cell.w = SSF.format(cell.z || SSF._table[14], datenum(parseDate(cell.v)));
							if(DIF_XL) push_value(o, 0, cell.w, "V");
							else push_value(o, 1, 0, cell.w);
							break;
						default: push_value(o, 1, 0, "");
					}
				}
			}
			push_value(o, -1, 0, "EOD");
			var RS = "\r\n";
			var oo = o.join(RS);
			//while((oo.length & 0x7F) != 0) oo += "\0";
			return oo;
		};
	})();
	return {
		to_workbook: dif_to_workbook,
		to_sheet: dif_to_sheet,
		from_sheet: sheet_to_dif
	};
})();

var ETH = (function() {
	function decode(s) { return s.replace(/\\b/g,"\\").replace(/\\c/g,":").replace(/\\n/g,"\n"); }
	function encode(s) { return s.replace(/\\/g, "\\b").replace(/:/g, "\\c").replace(/\n/g,"\\n"); }

	function eth_to_aoa(str, opts) {
		var records = str.split('\n'), R = -1, C = -1, ri = 0, arr = [];
		for (; ri !== records.length; ++ri) {
			var record = records[ri].trim().split(":");
			if(record[0] !== 'cell') continue;
			var addr = decode_cell(record[1]);
			if(arr.length <= addr.r) for(R = arr.length; R <= addr.r; ++R) if(!arr[R]) arr[R] = [];
			R = addr.r; C = addr.c;
			switch(record[2]) {
				case 't': arr[R][C] = decode(record[3]); break;
				case 'v': arr[R][C] = +record[3]; break;
				case 'vtf': var _f = record[record.length - 1];
					/* falls through */
				case 'vtc':
					switch(record[3]) {
						case 'nl': arr[R][C] = +record[4] ? true : false; break;
						default: arr[R][C] = +record[4]; break;
					}
					if(record[2] == 'vtf') arr[R][C] = [arr[R][C], _f];
			}
		}
		if(opts && opts.sheetRows) arr = arr.slice(0, opts.sheetRows);
		return arr;
	}

	function eth_to_sheet(d, opts) { return aoa_to_sheet(eth_to_aoa(d, opts), opts); }
	function eth_to_workbook(d, opts) { return sheet_to_workbook(eth_to_sheet(d, opts), opts); }

	var header = [
		"socialcalc:version:1.5",
		"MIME-Version: 1.0",
		"Content-Type: multipart/mixed; boundary=SocialCalcSpreadsheetControlSave"
	].join("\n");

	var sep = [
		"--SocialCalcSpreadsheetControlSave",
		"Content-type: text/plain; charset=UTF-8"
	].join("\n") + "\n";

	/* TODO: the other parts */
	var meta = [
		"# SocialCalc Spreadsheet Control Save",
		"part:sheet"
	].join("\n");

	var end = "--SocialCalcSpreadsheetControlSave--";

	function sheet_to_eth_data(ws) {
		if(!ws || !ws['!ref']) return "";
		var o = [], oo = [], cell, coord = "";
		var r = decode_range(ws['!ref']);
		var dense = Array.isArray(ws);
		for(var R = r.s.r; R <= r.e.r; ++R) {
			for(var C = r.s.c; C <= r.e.c; ++C) {
				coord = encode_cell({r:R,c:C});
				cell = dense ? (ws[R]||[])[C] : ws[coord];
				if(!cell || cell.v == null || cell.t === 'z') continue;
				oo = ["cell", coord, 't'];
				switch(cell.t) {
					case 's': case 'str': oo.push(encode(cell.v)); break;
					case 'n':
						if(!cell.f) { oo[2]='v'; oo[3]=cell.v; }
						else { oo[2]='vtf'; oo[3]='n'; oo[4]=cell.v; oo[5]=encode(cell.f); }
						break;
					case 'b':
						oo[2] = 'vt'+(cell.f?'f':'c'); oo[3]='nl'; oo[4]=cell.v?"1":"0";
						oo[5] = encode(cell.f||(cell.v?'TRUE':'FALSE'));
						break;
					case 'd':
						var t = datenum(parseDate(cell.v));
						oo[2] = 'vtc'; oo[3] = 'nd'; oo[4] = ""+t;
						oo[5] = cell.w || SSF.format(cell.z || SSF._table[14], t);
						break;
					case 'e': continue;
				}
				o.push(oo.join(":"));
			}
		}
		o.push("sheet:c:" + (r.e.c-r.s.c+1) + ":r:" + (r.e.r-r.s.r+1) + ":tvf:1");
		o.push("valueformat:1:text-wiki");
		//o.push("copiedfrom:" + ws['!ref']); // clipboard only
		return o.join("\n");
	}

	function sheet_to_eth(ws) {
		return [header, sep, meta, sep, sheet_to_eth_data(ws), end].join("\n");
		// return ["version:1.5", sheet_to_eth_data(ws)].join("\n"); // clipboard form
	}

	return {
		to_workbook: eth_to_workbook,
		to_sheet: eth_to_sheet,
		from_sheet: sheet_to_eth
	};
})();

var PRN = (function() {
	function set_text_arr(data, arr, R, C, o) {
		if(o.raw) arr[R][C] = data;
		else if(data === 'TRUE') arr[R][C] = true;
		else if(data === 'FALSE') arr[R][C] = false;
		else if(data === ""){/* empty */}
		else if(!isNaN(fuzzynum(data))) arr[R][C] = fuzzynum(data);
		else if(!isNaN(fuzzydate(data).getDate())) arr[R][C] = parseDate(data);
		else arr[R][C] = data;
	}

	function prn_to_aoa_str(f, opts) {
		var o = opts || {};
		var arr = ([]);
		if(!f || f.length === 0) return arr;
		var lines = f.split(/[\r\n]/);
		var L = lines.length - 1;
		while(L >= 0 && lines[L].length === 0) --L;
		var start = 10, idx = 0;
		var R = 0;
		for(; R <= L; ++R) {
			idx = lines[R].indexOf(" ");
			if(idx == -1) idx = lines[R].length; else idx++;
			start = Math.max(start, idx);
		}
		for(R = 0; R <= L; ++R) {
			arr[R] = [];
			/* TODO: confirm that widths are always 10 */
			var C = 0;
			set_text_arr(lines[R].slice(0, start).trim(), arr, R, C, o);
			for(C = 1; C <= (lines[R].length - start)/10 + 1; ++C)
				set_text_arr(lines[R].slice(start+(C-1)*10,start+C*10).trim(),arr,R,C,o);
		}
		if(o.sheetRows) arr = arr.slice(0, o.sheetRows);
		return arr;
	}

	// List of accepted CSV separators
	var guess_seps = {
0x2C: ',',
0x09: "\t",
0x3B: ';'
	};

	// CSV separator weights to be used in case of equal numbers
	var guess_sep_weights = {
0x2C: 3,
0x09: 2,
0x3B: 1
	};

	function guess_sep(str) {
		var cnt = {}, instr = false, end = 0, cc = 0;
		for(;end < str.length;++end) {
			if((cc=str.charCodeAt(end)) == 0x22) instr = !instr;
			else if(!instr && cc in guess_seps) cnt[cc] = (cnt[cc]||0)+1;
		}

		cc = [];
		for(end in cnt) if ( cnt.hasOwnProperty(end) ) {
			cc.push([ cnt[end], end ]);
		}

		if ( !cc.length ) {
			cnt = guess_sep_weights;
			for(end in cnt) if ( cnt.hasOwnProperty(end) ) {
				cc.push([ cnt[end], end ]);
			}
		}

		cc.sort(function(a, b) { return a[0] - b[0] || guess_sep_weights[a[1]] - guess_sep_weights[b[1]]; });

		return guess_seps[cc.pop()[1]];
	}

	function dsv_to_sheet_str(str, opts) {
		var o = opts || {};
		var sep = "";
		if(DENSE != null && o.dense == null) o.dense = DENSE;
		var ws = o.dense ? ([]) : ({});
		var range = ({s: {c:0, r:0}, e: {c:0, r:0}});

		if(str.slice(0,4) == "sep=" && str.charCodeAt(5) == 10) { sep = str.charAt(4); str = str.slice(6); }
		else sep = guess_sep(str.slice(0,1024));
		var R = 0, C = 0, v = 0;
		var start = 0, end = 0, sepcc = sep.charCodeAt(0), instr = false, cc=0;
		str = str.replace(/\r\n/mg, "\n");
		var _re = o.dateNF != null ? dateNF_regex(o.dateNF) : null;
		function finish_cell() {
			var s = str.slice(start, end);
			var cell = ({});
			if(s.charAt(0) == '"' && s.charAt(s.length - 1) == '"') s = s.slice(1,-1).replace(/""/g,'"');
			if(s.length === 0) cell.t = 'z';
			else if(o.raw) { cell.t = 's'; cell.v = s; }
			else if(s.trim().length === 0) { cell.t = 's'; cell.v = s; }
			else if(s.charCodeAt(0) == 0x3D) {
				if(s.charCodeAt(1) == 0x22 && s.charCodeAt(s.length - 1) == 0x22) { cell.t = 's'; cell.v = s.slice(2,-1).replace(/""/g,'"'); }
				else if(fuzzyfmla(s)) { cell.t = 'n'; cell.f = s.slice(1); }
				else { cell.t = 's'; cell.v = s; } }
			else if(s == "TRUE") { cell.t = 'b'; cell.v = true; }
			else if(s == "FALSE") { cell.t = 'b'; cell.v = false; }
			else if(!isNaN(v = fuzzynum(s))) { cell.t = 'n'; if(o.cellText !== false) cell.w = s; cell.v = v; }
			else if(!isNaN(fuzzydate(s).getDate()) || _re && s.match(_re)) {
				cell.z = o.dateNF || SSF._table[14];
				var k = 0;
				if(_re && s.match(_re)){ s=dateNF_fix(s, o.dateNF, (s.match(_re)||[])); k=1; }
				if(o.cellDates) { cell.t = 'd'; cell.v = parseDate(s, k); }
				else { cell.t = 'n'; cell.v = datenum(parseDate(s, k)); }
				if(o.cellText !== false) cell.w = SSF.format(cell.z, cell.v instanceof Date ? datenum(cell.v):cell.v);
				if(!o.cellNF) delete cell.z;
			} else {
				cell.t = 's';
				cell.v = s;
			}
			if(cell.t == 'z'){}
			else if(o.dense) { if(!ws[R]) ws[R] = []; ws[R][C] = cell; }
			else ws[encode_cell({c:C,r:R})] = cell;
			start = end+1;
			if(range.e.c < C) range.e.c = C;
			if(range.e.r < R) range.e.r = R;
			if(cc == sepcc) ++C; else { C = 0; ++R; if(o.sheetRows && o.sheetRows <= R) return true; }
		}
		outer: for(;end < str.length;++end) switch((cc=str.charCodeAt(end))) {
			case 0x22: instr = !instr; break;
			case sepcc: case 0x0a: case 0x0d: if(!instr && finish_cell()) break outer; break;
			default: break;
		}
		if(end - start > 0) finish_cell();

		ws['!ref'] = encode_range(range);
		return ws;
	}

	function prn_to_sheet_str(str, opts) {
		if(str.slice(0,4) == "sep=") return dsv_to_sheet_str(str, opts);
		if(str.indexOf("\t") >= 0 || str.indexOf(",") >= 0 || str.indexOf(";") >= 0) return dsv_to_sheet_str(str, opts);
		return aoa_to_sheet(prn_to_aoa_str(str, opts), opts);
	}

	function prn_to_sheet(d, opts) {
		var str = "", bytes = opts.type == 'string' ? [0,0,0,0] : firstbyte(d, opts);
		switch(opts.type) {
			case 'base64': str = Base64.decode(d); break;
			case 'binary': str = d; break;
			case 'buffer':
				if(opts.codepage == 65001) str = d.toString('utf8');
				else if(opts.codepage && typeof cptable !== 'undefined') str = cptable.utils.decode(opts.codepage, d);
				else str = d.toString('binary');
				break;
			case 'array': str = cc2str(d); break;
			case 'string': str = d; break;
			default: throw new Error("Unrecognized type " + opts.type);
		}
		if(bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) str = utf8read(str.slice(3));
		else if((opts.type == 'binary') && typeof cptable !== 'undefined' && opts.codepage)  str = cptable.utils.decode(opts.codepage, cptable.utils.encode(1252,str));
		if(str.slice(0,19) == "socialcalc:version:") return ETH.to_sheet(opts.type == 'string' ? str : utf8read(str), opts);
		return prn_to_sheet_str(str, opts);
	}

	function prn_to_workbook(d, opts) { return sheet_to_workbook(prn_to_sheet(d, opts), opts); }

	function sheet_to_prn(ws) {
		var o = [];
		var r = safe_decode_range(ws['!ref']), cell;
		var dense = Array.isArray(ws);
		for(var R = r.s.r; R <= r.e.r; ++R) {
			var oo = [];
			for(var C = r.s.c; C <= r.e.c; ++C) {
				var coord = encode_cell({r:R,c:C});
				cell = dense ? (ws[R]||[])[C] : ws[coord];
				if(!cell || cell.v == null) { oo.push("          "); continue; }
				var w = (cell.w || (format_cell(cell), cell.w) || "").slice(0,10);
				while(w.length < 10) w += " ";
				oo.push(w + (C === 0 ? " " : ""));
			}
			o.push(oo.join(""));
		}
		return o.join("\n");
	}

	return {
		to_workbook: prn_to_workbook,
		to_sheet: prn_to_sheet,
		from_sheet: sheet_to_prn
	};
})();

/* Excel defaults to SYLK but warns if data is not valid */
function read_wb_ID(d, opts) {
	var o = opts || {}, OLD_WTF = !!o.WTF; o.WTF = true;
	try {
		var out = SYLK.to_workbook(d, o);
		o.WTF = OLD_WTF;
		return out;
	} catch(e) {
		o.WTF = OLD_WTF;
		if(!e.message.match(/SYLK bad record ID/) && OLD_WTF) throw e;
		return PRN.to_workbook(d, opts);
	}
}
var WK_ = (function() {
	function lotushopper(data, cb, opts) {
		if(!data) return;
		prep_blob(data, data.l || 0);
		var Enum = opts.Enum || WK1Enum;
		while(data.l < data.length) {
			var RT = data.read_shift(2);
			var R = Enum[RT] || Enum[0xFF];
			var length = data.read_shift(2);
			var tgt = data.l + length;
			var d = (R.f||parsenoop)(data, length, opts);
			data.l = tgt;
			if(cb(d, R.n, RT)) return;
		}
	}

	function lotus_to_workbook(d, opts) {
		switch(opts.type) {
			case 'base64': return lotus_to_workbook_buf(s2a(Base64.decode(d)), opts);
			case 'binary': return lotus_to_workbook_buf(s2a(d), opts);
			case 'buffer':
			case 'array': return lotus_to_workbook_buf(d, opts);
		}
		throw "Unsupported type " + opts.type;
	}

	function lotus_to_workbook_buf(d, opts) {
		if(!d) return d;
		var o = opts || {};
		if(DENSE != null && o.dense == null) o.dense = DENSE;
		var s = ((o.dense ? [] : {})), n = "Sheet1", sidx = 0;
		var sheets = {}, snames = [n];

		var refguess = {s: {r:0, c:0}, e: {r:0, c:0} };
		var sheetRows = o.sheetRows || 0;

		if(d[2] == 0x02) o.Enum = WK1Enum;
		else if(d[2] == 0x1a) o.Enum = WK3Enum;
		else if(d[2] == 0x0e) { o.Enum = WK3Enum; o.qpro = true; d.l = 0; }
		else throw new Error("Unrecognized LOTUS BOF " + d[2]);
		lotushopper(d, function(val, Rn, RT) {
			if(d[2] == 0x02) switch(RT) {
				case 0x00:
					o.vers = val;
					if(val >= 0x1000) o.qpro = true;
					break;
				case 0x06: refguess = val; break; /* RANGE */
				case 0x0F: /* LABEL */
					if(!o.qpro) val[1].v = val[1].v.slice(1);
					/* falls through */
				case 0x0D: /* INTEGER */
				case 0x0E: /* NUMBER */
				case 0x10: /* FORMULA */
				case 0x33: /* STRING */
					/* TODO: actual translation of the format code */
					if(RT == 0x0E && (val[2] & 0x70) == 0x70 && (val[2] & 0x0F) > 1 && (val[2] & 0x0F) < 15) {
						val[1].z = o.dateNF || SSF._table[14];
						if(o.cellDates) { val[1].t = 'd'; val[1].v = numdate(val[1].v); }
					}
					if(o.dense) {
						if(!s[val[0].r]) s[val[0].r] = [];
						s[val[0].r][val[0].c] = val[1];
					} else s[encode_cell(val[0])] = val[1];
					break;
			} else switch(RT) {
				case 0x16: /* LABEL16 */
					val[1].v = val[1].v.slice(1);
					/* falls through */
				case 0x17: /* NUMBER17 */
				case 0x18: /* NUMBER18 */
				case 0x19: /* FORMULA19 */
				case 0x25: /* NUMBER25 */
				case 0x27: /* NUMBER27 */
				case 0x28: /* FORMULA28 */
					if(val[3] > sidx) {
						s["!ref"] = encode_range(refguess);
						sheets[n] = s;
						s = (o.dense ? [] : {});
						refguess = {s: {r:0, c:0}, e: {r:0, c:0} };
						sidx = val[3]; n = "Sheet" + (sidx + 1);
						snames.push(n);
					}
					if(sheetRows > 0 && val[0].r >= sheetRows) break;
					if(o.dense) {
						if(!s[val[0].r]) s[val[0].r] = [];
						s[val[0].r][val[0].c] = val[1];
					} else s[encode_cell(val[0])] = val[1];
					if(refguess.e.c < val[0].c) refguess.e.c = val[0].c;
					if(refguess.e.r < val[0].r) refguess.e.r = val[0].r;
					break;
				default: break;
			}
		}, o);

		s["!ref"] = encode_range(refguess);
		sheets[n] = s;
		return { SheetNames: snames, Sheets:sheets };
	}

	function parse_RANGE(blob) {
		var o = {s:{c:0,r:0},e:{c:0,r:0}};
		o.s.c = blob.read_shift(2);
		o.s.r = blob.read_shift(2);
		o.e.c = blob.read_shift(2);
		o.e.r = blob.read_shift(2);
		if(o.s.c == 0xFFFF) o.s.c = o.e.c = o.s.r = o.e.r = 0;
		return o;
	}

	function parse_cell(blob, length, opts) {
		var o = [{c:0,r:0}, {t:'n',v:0}, 0];
		if(opts.qpro && opts.vers != 0x5120) {
			o[0].c = blob.read_shift(1);
			blob.l++;
			o[0].r = blob.read_shift(2);
			blob.l+=2;
		} else {
			o[2] = blob.read_shift(1);
			o[0].c = blob.read_shift(2); o[0].r = blob.read_shift(2);
		}
		return o;
	}

	function parse_LABEL(blob, length, opts) {
		var tgt = blob.l + length;
		var o = parse_cell(blob, length, opts);
		o[1].t = 's';
		if(opts.vers == 0x5120) {
			blob.l++;
			var len = blob.read_shift(1);
			o[1].v = blob.read_shift(len, 'utf8');
			return o;
		}
		if(opts.qpro) blob.l++;
		o[1].v = blob.read_shift(tgt - blob.l, 'cstr');
		return o;
	}

	function parse_INTEGER(blob, length, opts) {
		var o = parse_cell(blob, length, opts);
		o[1].v = blob.read_shift(2, 'i');
		return o;
	}

	function parse_NUMBER(blob, length, opts) {
		var o = parse_cell(blob, length, opts);
		o[1].v = blob.read_shift(8, 'f');
		return o;
	}

	function parse_FORMULA(blob, length, opts) {
		var tgt = blob.l + length;
		var o = parse_cell(blob, length, opts);
		/* TODO: formula */
		o[1].v = blob.read_shift(8, 'f');
		if(opts.qpro) blob.l = tgt;
		else {
			var flen = blob.read_shift(2);
			blob.l += flen;
		}
		return o;
	}

	function parse_cell_3(blob) {
		var o = [{c:0,r:0}, {t:'n',v:0}, 0];
		o[0].r = blob.read_shift(2); o[3] = blob[blob.l++]; o[0].c = blob[blob.l++];
		return o;
	}

	function parse_LABEL_16(blob, length) {
		var o = parse_cell_3(blob, length);
		o[1].t = 's';
		o[1].v = blob.read_shift(length - 4, 'cstr');
		return o;
	}

	function parse_NUMBER_18(blob, length) {
		var o = parse_cell_3(blob, length);
		o[1].v = blob.read_shift(2);
		var v = o[1].v >> 1;
		/* TODO: figure out all of the corner cases */
		if(o[1].v & 0x1) {
			switch(v & 0x07) {
				case 1: v = (v >> 3) * 500; break;
				case 2: v = (v >> 3) / 20; break;
				case 4: v = (v >> 3) / 2000; break;
				case 6: v = (v >> 3) / 16; break;
				case 7: v = (v >> 3) / 64; break;
				default: throw "unknown NUMBER_18 encoding " + (v & 0x07);
			}
		}
		o[1].v = v;
		return o;
	}

	function parse_NUMBER_17(blob, length) {
		var o = parse_cell_3(blob, length);
		var v1 = blob.read_shift(4);
		var v2 = blob.read_shift(4);
		var e = blob.read_shift(2);
		if(e == 0xFFFF) { o[1].v = 0; return o; }
		var s = e & 0x8000; e = (e&0x7FFF) - 16446;
		o[1].v = (s*2 - 1) * ((e > 0 ? (v2 << e) : (v2 >>> -e)) + (e > -32 ? (v1 << (e + 32)) : (v1 >>> -(e + 32))));
		return o;
	}

	function parse_FORMULA_19(blob, length) {
		var o = parse_NUMBER_17(blob, 14);
		blob.l += length - 14; /* TODO: formula */
		return o;
	}

	function parse_NUMBER_25(blob, length) {
		var o = parse_cell_3(blob, length);
		var v1 = blob.read_shift(4);
		o[1].v = v1 >> 6;
		return o;
	}

	function parse_NUMBER_27(blob, length) {
		var o = parse_cell_3(blob, length);
		var v1 = blob.read_shift(8,'f');
		o[1].v = v1;
		return o;
	}

	function parse_FORMULA_28(blob, length) {
		var o = parse_NUMBER_27(blob, 14);
		blob.l += length - 10; /* TODO: formula */
		return o;
	}

	var WK1Enum = {
0x0000: { n:"BOF", f:parseuint16 },
0x0001: { n:"EOF" },
0x0002: { n:"CALCMODE" },
0x0003: { n:"CALCORDER" },
0x0004: { n:"SPLIT" },
0x0005: { n:"SYNC" },
0x0006: { n:"RANGE", f:parse_RANGE },
0x0007: { n:"WINDOW1" },
0x0008: { n:"COLW1" },
0x0009: { n:"WINTWO" },
0x000A: { n:"COLW2" },
0x000B: { n:"NAME" },
0x000C: { n:"BLANK" },
0x000D: { n:"INTEGER", f:parse_INTEGER },
0x000E: { n:"NUMBER", f:parse_NUMBER },
0x000F: { n:"LABEL", f:parse_LABEL },
0x0010: { n:"FORMULA", f:parse_FORMULA },
0x0018: { n:"TABLE" },
0x0019: { n:"ORANGE" },
0x001A: { n:"PRANGE" },
0x001B: { n:"SRANGE" },
0x001C: { n:"FRANGE" },
0x001D: { n:"KRANGE1" },
0x0020: { n:"HRANGE" },
0x0023: { n:"KRANGE2" },
0x0024: { n:"PROTEC" },
0x0025: { n:"FOOTER" },
0x0026: { n:"HEADER" },
0x0027: { n:"SETUP" },
0x0028: { n:"MARGINS" },
0x0029: { n:"LABELFMT" },
0x002A: { n:"TITLES" },
0x002B: { n:"SHEETJS" },
0x002D: { n:"GRAPH" },
0x002E: { n:"NGRAPH" },
0x002F: { n:"CALCCOUNT" },
0x0030: { n:"UNFORMATTED" },
0x0031: { n:"CURSORW12" },
0x0032: { n:"WINDOW" },
0x0033: { n:"STRING", f:parse_LABEL },
0x0037: { n:"PASSWORD" },
0x0038: { n:"LOCKED" },
0x003C: { n:"QUERY" },
0x003D: { n:"QUERYNAME" },
0x003E: { n:"PRINT" },
0x003F: { n:"PRINTNAME" },
0x0040: { n:"GRAPH2" },
0x0041: { n:"GRAPHNAME" },
0x0042: { n:"ZOOM" },
0x0043: { n:"SYMSPLIT" },
0x0044: { n:"NSROWS" },
0x0045: { n:"NSCOLS" },
0x0046: { n:"RULER" },
0x0047: { n:"NNAME" },
0x0048: { n:"ACOMM" },
0x0049: { n:"AMACRO" },
0x004A: { n:"PARSE" },
0x00FF: { n:"", f:parsenoop }
	};

	var WK3Enum = {
0x0000: { n:"BOF" },
0x0001: { n:"EOF" },
0x0003: { n:"??" },
0x0004: { n:"??" },
0x0005: { n:"??" },
0x0006: { n:"??" },
0x0007: { n:"??" },
0x0009: { n:"??" },
0x000a: { n:"??" },
0x000b: { n:"??" },
0x000c: { n:"??" },
0x000e: { n:"??" },
0x000f: { n:"??" },
0x0010: { n:"??" },
0x0011: { n:"??" },
0x0012: { n:"??" },
0x0013: { n:"??" },
0x0015: { n:"??" },
0x0016: { n:"LABEL16", f:parse_LABEL_16},
0x0017: { n:"NUMBER17", f:parse_NUMBER_17 },
0x0018: { n:"NUMBER18", f:parse_NUMBER_18 },
0x0019: { n:"FORMULA19", f:parse_FORMULA_19},
0x001a: { n:"??" },
0x001b: { n:"??" },
0x001c: { n:"??" },
0x001d: { n:"??" },
0x001e: { n:"??" },
0x001f: { n:"??" },
0x0021: { n:"??" },
0x0025: { n:"NUMBER25", f:parse_NUMBER_25 },
0x0027: { n:"NUMBER27", f:parse_NUMBER_27 },
0x0028: { n:"FORMULA28", f:parse_FORMULA_28 },
0x00FF: { n:"", f:parsenoop }
	};
	return {
		to_workbook: lotus_to_workbook
	};
})();
/* Parse a list of <r> tags */
var parse_rs = (function parse_rs_factory() {
	var tregex = matchtag("t"), rpregex = matchtag("rPr"), rregex = /<(?:\w+:)?r>/g, rend = /<\/(?:\w+:)?r>/, nlregex = /\r\n/g;
	/* 18.4.7 rPr CT_RPrElt */
	var parse_rpr = function parse_rpr(rpr, intro, outro) {
		var font = {}, cp = 65001, align = "";
		var pass = false;
		var m = rpr.match(tagregex), i = 0;
		if(m) for(;i!=m.length; ++i) {
			var y = parsexmltag(m[i]);
			switch(y[0].replace(/\w*:/g,"")) {
				/* 18.8.12 condense CT_BooleanProperty */
				/* ** not required . */
				case '<condense': break;
				/* 18.8.17 extend CT_BooleanProperty */
				/* ** not required . */
				case '<extend': break;
				/* 18.8.36 shadow CT_BooleanProperty */
				/* ** not required . */
				case '<shadow':
					if(!y.val) break;
					/* falls through */
				case '<shadow>':
				case '<shadow/>': font.shadow = 1; break;
				case '</shadow>': break;

				/* 18.4.1 charset CT_IntProperty TODO */
				case '<charset':
					if(y.val == '1') break;
					cp = CS2CP[parseInt(y.val, 10)];
					break;

				/* 18.4.2 outline CT_BooleanProperty TODO */
				case '<outline':
					if(!y.val) break;
					/* falls through */
				case '<outline>':
				case '<outline/>': font.outline = 1; break;
				case '</outline>': break;

				/* 18.4.5 rFont CT_FontName */
				case '<rFont': font.name = y.val; break;

				/* 18.4.11 sz CT_FontSize */
				case '<sz': font.sz = y.val; break;

				/* 18.4.10 strike CT_BooleanProperty */
				case '<strike':
					if(!y.val) break;
					/* falls through */
				case '<strike>':
				case '<strike/>': font.strike = 1; break;
				case '</strike>': break;

				/* 18.4.13 u CT_UnderlineProperty */
				case '<u':
					if(!y.val) break;
					switch(y.val) {
						case 'double': font.uval = "double"; break;
						case 'singleAccounting': font.uval = "single-accounting"; break;
						case 'doubleAccounting': font.uval = "double-accounting"; break;
					}
					/* falls through */
				case '<u>':
				case '<u/>': font.u = 1; break;
				case '</u>': break;

				/* 18.8.2 b */
				case '<b':
					if(y.val == '0') break;
					/* falls through */
				case '<b>':
				case '<b/>': font.b = 1; break;
				case '</b>': break;

				/* 18.8.26 i */
				case '<i':
					if(y.val == '0') break;
					/* falls through */
				case '<i>':
				case '<i/>': font.i = 1; break;
				case '</i>': break;

				/* 18.3.1.15 color CT_Color TODO: tint, theme, auto, indexed */
				case '<color':
					if(y.rgb) font.color = y.rgb.slice(2,8);
					break;

				/* 18.8.18 family ST_FontFamily */
				case '<family': font.family = y.val; break;

				/* 18.4.14 vertAlign CT_VerticalAlignFontProperty TODO */
				case '<vertAlign': align = y.val; break;

				/* 18.8.35 scheme CT_FontScheme TODO */
				case '<scheme': break;

				/* 18.2.10 extLst CT_ExtensionList ? */
				case '<extLst': case '<extLst>': case '</extLst>': break;
				case '<ext': pass = true; break;
				case '</ext>': pass = false; break;
				default:
					if(y[0].charCodeAt(1) !== 47 && !pass) throw new Error('Unrecognized rich format ' + y[0]);
			}
		}
		var style = [];

		if(font.u) style.push("text-decoration: underline;");
		if(font.uval) style.push("text-underline-style:" + font.uval + ";");
		if(font.sz) style.push("font-size:" + font.sz + "pt;");
		if(font.outline) style.push("text-effect: outline;");
		if(font.shadow) style.push("text-shadow: auto;");
		intro.push('<span style="' + style.join("") + '">');

		if(font.b) { intro.push("<b>"); outro.push("</b>"); }
		if(font.i) { intro.push("<i>"); outro.push("</i>"); }
		if(font.strike) { intro.push("<s>"); outro.push("</s>"); }

		if(align == "superscript") align = "sup";
		else if(align == "subscript") align = "sub";
		if(align != "") { intro.push("<" + align + ">"); outro.push("</" + align + ">"); }

		outro.push("</span>");
		return cp;
	};

	/* 18.4.4 r CT_RElt */
	function parse_r(r) {
		var terms = [[],"",[]];
		/* 18.4.12 t ST_Xstring */
		var t = r.match(tregex)/*, cp = 65001*/;
		if(!t) return "";
		terms[1] = t[1];

		var rpr = r.match(rpregex);
		if(rpr) /*cp = */parse_rpr(rpr[1], terms[0], terms[2]);

		return terms[0].join("") + terms[1].replace(nlregex,'<br/>') + terms[2].join("");
	}
	return function parse_rs(rs) {
		return rs.replace(rregex,"").split(rend).map(parse_r).join("");
	};
})();

/* 18.4.8 si CT_Rst */
var sitregex = /<(?:\w+:)?t[^>]*>([^<]*)<\/(?:\w+:)?t>/g, sirregex = /<(?:\w+:)?r>/;
var sirphregex = /<(?:\w+:)?rPh.*?>([\s\S]*?)<\/(?:\w+:)?rPh>/g;
function parse_si(x, opts) {
	var html = opts ? opts.cellHTML : true;
	var z = {};
	if(!x) return null;
	//var y;
	/* 18.4.12 t ST_Xstring (Plaintext String) */
	// TODO: is whitespace actually valid here?
	if(x.match(/^\s*<(?:\w+:)?t[^>]*>/)) {
		z.t = unescapexml(utf8read(x.slice(x.indexOf(">")+1).split(/<\/(?:\w+:)?t>/)[0]||""));
		z.r = utf8read(x);
		if(html) z.h = escapehtml(z.t);
	}
	/* 18.4.4 r CT_RElt (Rich Text Run) */
	else if((/*y = */x.match(sirregex))) {
		z.r = utf8read(x);
		z.t = unescapexml(utf8read((x.replace(sirphregex, '').match(sitregex)||[]).join("").replace(tagregex,"")));
		if(html) z.h = parse_rs(z.r);
	}
	/* 18.4.3 phoneticPr CT_PhoneticPr (TODO: needed for Asian support) */
	/* 18.4.6 rPh CT_PhoneticRun (TODO: needed for Asian support) */
	return z;
}

/* 18.4 Shared String Table */
var sstr0 = /<(?:\w+:)?sst([^>]*)>([\s\S]*)<\/(?:\w+:)?sst>/;
var sstr1 = /<(?:\w+:)?(?:si|sstItem)>/g;
var sstr2 = /<\/(?:\w+:)?(?:si|sstItem)>/;
function parse_sst_xml(data, opts) {
	var s = ([]), ss = "";
	if(!data) return s;
	/* 18.4.9 sst CT_Sst */
	var sst = data.match(sstr0);
	if(sst) {
		ss = sst[2].replace(sstr1,"").split(sstr2);
		for(var i = 0; i != ss.length; ++i) {
			var o = parse_si(ss[i].trim(), opts);
			if(o != null) s[s.length] = o;
		}
		sst = parsexmltag(sst[1]); s.Count = sst.count; s.Unique = sst.uniqueCount;
	}
	return s;
}

RELS.SST = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
var straywsregex = /^\s|\s$|[\t\n\r]/;
function write_sst_xml(sst, opts) {
	if(!opts.bookSST) return "";
	var o = [XML_HEADER];
	o[o.length] = (writextag('sst', null, {
		xmlns: XMLNS.main[0],
		count: sst.Count,
		uniqueCount: sst.Unique
	}));
	for(var i = 0; i != sst.length; ++i) { if(sst[i] == null) continue;
		var s = sst[i];
		var sitag = "<si>";
		if(s.r) sitag += s.r;
		else {
			sitag += "<t";
			if(!s.t) s.t = "";
			if(s.t.match(straywsregex)) sitag += ' xml:space="preserve"';
			sitag += ">" + escapexml(s.t) + "</t>";
		}
		sitag += "</si>";
		o[o.length] = (sitag);
	}
	if(o.length>2){ o[o.length] = ('</sst>'); o[1]=o[1].replace("/>",">"); }
	return o.join("");
}
/* [MS-XLSB] 2.4.221 BrtBeginSst */
function parse_BrtBeginSst(data) {
	return [data.read_shift(4), data.read_shift(4)];
}

/* [MS-XLSB] 2.1.7.45 Shared Strings */
function parse_sst_bin(data, opts) {
	var s = ([]);
	var pass = false;
	recordhopper(data, function hopper_sst(val, R_n, RT) {
		switch(RT) {
			case 0x009F: /* 'BrtBeginSst' */
				s.Count = val[0]; s.Unique = val[1]; break;
			case 0x0013: /* 'BrtSSTItem' */
				s.push(val); break;
			case 0x00A0: /* 'BrtEndSst' */
				return true;

			case 0x0023: /* 'BrtFRTBegin' */
				pass = true; break;
			case 0x0024: /* 'BrtFRTEnd' */
				pass = false; break;

			default:
				if(R_n.indexOf("Begin") > 0){/* empty */}
				else if(R_n.indexOf("End") > 0){/* empty */}
				if(!pass || opts.WTF) throw new Error("Unexpected record " + RT + " " + R_n);
		}
	});
	return s;
}

function write_BrtBeginSst(sst, o) {
	if(!o) o = new_buf(8);
	o.write_shift(4, sst.Count);
	o.write_shift(4, sst.Unique);
	return o;
}

var write_BrtSSTItem = write_RichStr;

function write_sst_bin(sst) {
	var ba = buf_array();
	write_record(ba, "BrtBeginSst", write_BrtBeginSst(sst));
	for(var i = 0; i < sst.length; ++i) write_record(ba, "BrtSSTItem", write_BrtSSTItem(sst[i]));
	/* FRTSST */
	write_record(ba, "BrtEndSst");
	return ba.end();
}
function _JS2ANSI(str) {
	if(typeof cptable !== 'undefined') return cptable.utils.encode(current_ansi, str);
	var o = [], oo = str.split("");
	for(var i = 0; i < oo.length; ++i) o[i] = oo[i].charCodeAt(0);
	return o;
}

/* [MS-OFFCRYPTO] 2.1.4 Version */
function parse_CRYPTOVersion(blob, length) {
	var o = {};
	o.Major = blob.read_shift(2);
	o.Minor = blob.read_shift(2);
if(length >= 4) blob.l += length - 4;
	return o;
}

/* [MS-OFFCRYPTO] 2.1.5 DataSpaceVersionInfo */
function parse_DataSpaceVersionInfo(blob) {
	var o = {};
	o.id = blob.read_shift(0, 'lpp4');
	o.R = parse_CRYPTOVersion(blob, 4);
	o.U = parse_CRYPTOVersion(blob, 4);
	o.W = parse_CRYPTOVersion(blob, 4);
	return o;
}

/* [MS-OFFCRYPTO] 2.1.6.1 DataSpaceMapEntry Structure */
function parse_DataSpaceMapEntry(blob) {
	var len = blob.read_shift(4);
	var end = blob.l + len - 4;
	var o = {};
	var cnt = blob.read_shift(4);
	var comps = [];
	/* [MS-OFFCRYPTO] 2.1.6.2 DataSpaceReferenceComponent Structure */
	while(cnt-- > 0) comps.push({ t: blob.read_shift(4), v: blob.read_shift(0, 'lpp4') });
	o.name = blob.read_shift(0, 'lpp4');
	o.comps = comps;
	if(blob.l != end) throw new Error("Bad DataSpaceMapEntry: " + blob.l + " != " + end);
	return o;
}

/* [MS-OFFCRYPTO] 2.1.6 DataSpaceMap */
function parse_DataSpaceMap(blob) {
	var o = [];
	blob.l += 4; // must be 0x8
	var cnt = blob.read_shift(4);
	while(cnt-- > 0) o.push(parse_DataSpaceMapEntry(blob));
	return o;
}

/* [MS-OFFCRYPTO] 2.1.7 DataSpaceDefinition */
function parse_DataSpaceDefinition(blob) {
	var o = [];
	blob.l += 4; // must be 0x8
	var cnt = blob.read_shift(4);
	while(cnt-- > 0) o.push(blob.read_shift(0, 'lpp4'));
	return o;
}

/* [MS-OFFCRYPTO] 2.1.8 DataSpaceDefinition */
function parse_TransformInfoHeader(blob) {
	var o = {};
	/*var len = */blob.read_shift(4);
	blob.l += 4; // must be 0x1
	o.id = blob.read_shift(0, 'lpp4');
	o.name = blob.read_shift(0, 'lpp4');
	o.R = parse_CRYPTOVersion(blob, 4);
	o.U = parse_CRYPTOVersion(blob, 4);
	o.W = parse_CRYPTOVersion(blob, 4);
	return o;
}

function parse_Primary(blob) {
	/* [MS-OFFCRYPTO] 2.2.6 IRMDSTransformInfo */
	var hdr = parse_TransformInfoHeader(blob);
	/* [MS-OFFCRYPTO] 2.1.9 EncryptionTransformInfo */
	hdr.ename = blob.read_shift(0, '8lpp4');
	hdr.blksz = blob.read_shift(4);
	hdr.cmode = blob.read_shift(4);
	if(blob.read_shift(4) != 0x04) throw new Error("Bad !Primary record");
	return hdr;
}

/* [MS-OFFCRYPTO] 2.3.2 Encryption Header */
function parse_EncryptionHeader(blob, length) {
	var tgt = blob.l + length;
	var o = {};
	o.Flags = (blob.read_shift(4) & 0x3F);
	blob.l += 4;
	o.AlgID = blob.read_shift(4);
	var valid = false;
	switch(o.AlgID) {
		case 0x660E: case 0x660F: case 0x6610: valid = (o.Flags == 0x24); break;
		case 0x6801: valid = (o.Flags == 0x04); break;
		case 0: valid = (o.Flags == 0x10 || o.Flags == 0x04 || o.Flags == 0x24); break;
		default: throw 'Unrecognized encryption algorithm: ' + o.AlgID;
	}
	if(!valid) throw new Error("Encryption Flags/AlgID mismatch");
	o.AlgIDHash = blob.read_shift(4);
	o.KeySize = blob.read_shift(4);
	o.ProviderType = blob.read_shift(4);
	blob.l += 8;
	o.CSPName = blob.read_shift((tgt-blob.l)>>1, 'utf16le');
	blob.l = tgt;
	return o;
}

/* [MS-OFFCRYPTO] 2.3.3 Encryption Verifier */
function parse_EncryptionVerifier(blob, length) {
	var o = {}, tgt = blob.l + length;
	blob.l += 4; // SaltSize must be 0x10
	o.Salt = blob.slice(blob.l, blob.l+16); blob.l += 16;
	o.Verifier = blob.slice(blob.l, blob.l+16); blob.l += 16;
	/*var sz = */blob.read_shift(4);
	o.VerifierHash = blob.slice(blob.l, tgt); blob.l = tgt;
	return o;
}

/* [MS-OFFCRYPTO] 2.3.4.* EncryptionInfo Stream */
function parse_EncryptionInfo(blob) {
	var vers = parse_CRYPTOVersion(blob);
	switch(vers.Minor) {
		case 0x02: return [vers.Minor, parse_EncInfoStd(blob, vers)];
		case 0x03: return [vers.Minor, parse_EncInfoExt(blob, vers)];
		case 0x04: return [vers.Minor, parse_EncInfoAgl(blob, vers)];
	}
	throw new Error("ECMA-376 Encrypted file unrecognized Version: " + vers.Minor);
}

/* [MS-OFFCRYPTO] 2.3.4.5  EncryptionInfo Stream (Standard Encryption) */
function parse_EncInfoStd(blob) {
	var flags = blob.read_shift(4);
	if((flags & 0x3F) != 0x24) throw new Error("EncryptionInfo mismatch");
	var sz = blob.read_shift(4);
	//var tgt = blob.l + sz;
	var hdr = parse_EncryptionHeader(blob, sz);
	var verifier = parse_EncryptionVerifier(blob, blob.length - blob.l);
	return { t:"Std", h:hdr, v:verifier };
}
/* [MS-OFFCRYPTO] 2.3.4.6  EncryptionInfo Stream (Extensible Encryption) */
function parse_EncInfoExt() { throw new Error("File is password-protected: ECMA-376 Extensible"); }
/* [MS-OFFCRYPTO] 2.3.4.10 EncryptionInfo Stream (Agile Encryption) */
function parse_EncInfoAgl(blob) {
	var KeyData = ["saltSize","blockSize","keyBits","hashSize","cipherAlgorithm","cipherChaining","hashAlgorithm","saltValue"];
	blob.l+=4;
	var xml = blob.read_shift(blob.length - blob.l, 'utf8');
	var o = {};
	xml.replace(tagregex, function xml_agile(x) {
		var y = parsexmltag(x);
		switch(strip_ns(y[0])) {
			case '<?xml': break;
			case '<encryption': case '</encryption>': break;
			case '<keyData': KeyData.forEach(function(k) { o[k] = y[k]; }); break;
			case '<dataIntegrity': o.encryptedHmacKey = y.encryptedHmacKey; o.encryptedHmacValue = y.encryptedHmacValue; break;
			case '<keyEncryptors>': case '<keyEncryptors': o.encs = []; break;
			case '</keyEncryptors>': break;

			case '<keyEncryptor': o.uri = y.uri; break;
			case '</keyEncryptor>': break;
			case '<encryptedKey': o.encs.push(y); break;
			default: throw y[0];
		}
	});
	return o;
}

/* [MS-OFFCRYPTO] 2.3.5.1 RC4 CryptoAPI Encryption Header */
function parse_RC4CryptoHeader(blob, length) {
	var o = {};
	var vers = o.EncryptionVersionInfo = parse_CRYPTOVersion(blob, 4); length -= 4;
	if(vers.Minor != 2) throw new Error('unrecognized minor version code: ' + vers.Minor);
	if(vers.Major > 4 || vers.Major < 2) throw new Error('unrecognized major version code: ' + vers.Major);
	o.Flags = blob.read_shift(4); length -= 4;
	var sz = blob.read_shift(4); length -= 4;
	o.EncryptionHeader = parse_EncryptionHeader(blob, sz); length -= sz;
	o.EncryptionVerifier = parse_EncryptionVerifier(blob, length);
	return o;
}
/* [MS-OFFCRYPTO] 2.3.6.1 RC4 Encryption Header */
function parse_RC4Header(blob) {
	var o = {};
	var vers = o.EncryptionVersionInfo = parse_CRYPTOVersion(blob, 4);
	if(vers.Major != 1 || vers.Minor != 1) throw 'unrecognized version code ' + vers.Major + ' : ' + vers.Minor;
	o.Salt = blob.read_shift(16);
	o.EncryptedVerifier = blob.read_shift(16);
	o.EncryptedVerifierHash = blob.read_shift(16);
	return o;
}

/* [MS-OFFCRYPTO] 2.3.7.1 Binary Document Password Verifier Derivation */
function crypto_CreatePasswordVerifier_Method1(Password) {
	var Verifier = 0x0000, PasswordArray;
	var PasswordDecoded = _JS2ANSI(Password);
	var len = PasswordDecoded.length + 1, i, PasswordByte;
	var Intermediate1, Intermediate2, Intermediate3;
	PasswordArray = new_raw_buf(len);
	PasswordArray[0] = PasswordDecoded.length;
	for(i = 1; i != len; ++i) PasswordArray[i] = PasswordDecoded[i-1];
	for(i = len-1; i >= 0; --i) {
		PasswordByte = PasswordArray[i];
		Intermediate1 = ((Verifier & 0x4000) === 0x0000) ? 0 : 1;
		Intermediate2 = (Verifier << 1) & 0x7FFF;
		Intermediate3 = Intermediate1 | Intermediate2;
		Verifier = Intermediate3 ^ PasswordByte;
	}
	return Verifier ^ 0xCE4B;
}

/* [MS-OFFCRYPTO] 2.3.7.2 Binary Document XOR Array Initialization */
var crypto_CreateXorArray_Method1 = (function() {
	var PadArray = [0xBB, 0xFF, 0xFF, 0xBA, 0xFF, 0xFF, 0xB9, 0x80, 0x00, 0xBE, 0x0F, 0x00, 0xBF, 0x0F, 0x00];
	var InitialCode = [0xE1F0, 0x1D0F, 0xCC9C, 0x84C0, 0x110C, 0x0E10, 0xF1CE, 0x313E, 0x1872, 0xE139, 0xD40F, 0x84F9, 0x280C, 0xA96A, 0x4EC3];
	var XorMatrix = [0xAEFC, 0x4DD9, 0x9BB2, 0x2745, 0x4E8A, 0x9D14, 0x2A09, 0x7B61, 0xF6C2, 0xFDA5, 0xEB6B, 0xC6F7, 0x9DCF, 0x2BBF, 0x4563, 0x8AC6, 0x05AD, 0x0B5A, 0x16B4, 0x2D68, 0x5AD0, 0x0375, 0x06EA, 0x0DD4, 0x1BA8, 0x3750, 0x6EA0, 0xDD40, 0xD849, 0xA0B3, 0x5147, 0xA28E, 0x553D, 0xAA7A, 0x44D5, 0x6F45, 0xDE8A, 0xAD35, 0x4A4B, 0x9496, 0x390D, 0x721A, 0xEB23, 0xC667, 0x9CEF, 0x29FF, 0x53FE, 0xA7FC, 0x5FD9, 0x47D3, 0x8FA6, 0x0F6D, 0x1EDA, 0x3DB4, 0x7B68, 0xF6D0, 0xB861, 0x60E3, 0xC1C6, 0x93AD, 0x377B, 0x6EF6, 0xDDEC, 0x45A0, 0x8B40, 0x06A1, 0x0D42, 0x1A84, 0x3508, 0x6A10, 0xAA51, 0x4483, 0x8906, 0x022D, 0x045A, 0x08B4, 0x1168, 0x76B4, 0xED68, 0xCAF1, 0x85C3, 0x1BA7, 0x374E, 0x6E9C, 0x3730, 0x6E60, 0xDCC0, 0xA9A1, 0x4363, 0x86C6, 0x1DAD, 0x3331, 0x6662, 0xCCC4, 0x89A9, 0x0373, 0x06E6, 0x0DCC, 0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48C4];
	var Ror = function(Byte) { return ((Byte/2) | (Byte*128)) & 0xFF; };
	var XorRor = function(byte1, byte2) { return Ror(byte1 ^ byte2); };
	var CreateXorKey_Method1 = function(Password) {
		var XorKey = InitialCode[Password.length - 1];
		var CurrentElement = 0x68;
		for(var i = Password.length-1; i >= 0; --i) {
			var Char = Password[i];
			for(var j = 0; j != 7; ++j) {
				if(Char & 0x40) XorKey ^= XorMatrix[CurrentElement];
				Char *= 2; --CurrentElement;
			}
		}
		return XorKey;
	};
	return function(password) {
		var Password = _JS2ANSI(password);
		var XorKey = CreateXorKey_Method1(Password);
		var Index = Password.length;
		var ObfuscationArray = new_raw_buf(16);
		for(var i = 0; i != 16; ++i) ObfuscationArray[i] = 0x00;
		var Temp, PasswordLastChar, PadIndex;
		if((Index & 1) === 1) {
			Temp = XorKey >> 8;
			ObfuscationArray[Index] = XorRor(PadArray[0], Temp);
			--Index;
			Temp = XorKey & 0xFF;
			PasswordLastChar = Password[Password.length - 1];
			ObfuscationArray[Index] = XorRor(PasswordLastChar, Temp);
		}
		while(Index > 0) {
			--Index;
			Temp = XorKey >> 8;
			ObfuscationArray[Index] = XorRor(Password[Index], Temp);
			--Index;
			Temp = XorKey & 0xFF;
			ObfuscationArray[Index] = XorRor(Password[Index], Temp);
		}
		Index = 15;
		PadIndex = 15 - Password.length;
		while(PadIndex > 0) {
			Temp = XorKey >> 8;
			ObfuscationArray[Index] = XorRor(PadArray[PadIndex], Temp);
			--Index;
			--PadIndex;
			Temp = XorKey & 0xFF;
			ObfuscationArray[Index] = XorRor(Password[Index], Temp);
			--Index;
			--PadIndex;
		}
		return ObfuscationArray;
	};
})();

/* [MS-OFFCRYPTO] 2.3.7.3 Binary Document XOR Data Transformation Method 1 */
var crypto_DecryptData_Method1 = function(password, Data, XorArrayIndex, XorArray, O) {
	/* If XorArray is set, use it; if O is not set, make changes in-place */
	if(!O) O = Data;
	if(!XorArray) XorArray = crypto_CreateXorArray_Method1(password);
	var Index, Value;
	for(Index = 0; Index != Data.length; ++Index) {
		Value = Data[Index];
		Value ^= XorArray[XorArrayIndex];
		Value = ((Value>>5) | (Value<<3)) & 0xFF;
		O[Index] = Value;
		++XorArrayIndex;
	}
	return [O, XorArrayIndex, XorArray];
};

var crypto_MakeXorDecryptor = function(password) {
	var XorArrayIndex = 0, XorArray = crypto_CreateXorArray_Method1(password);
	return function(Data) {
		var O = crypto_DecryptData_Method1("", Data, XorArrayIndex, XorArray);
		XorArrayIndex = O[1];
		return O[0];
	};
};

/* 2.5.343 */
function parse_XORObfuscation(blob, length, opts, out) {
	var o = ({ key: parseuint16(blob), verificationBytes: parseuint16(blob) });
	if(opts.password) o.verifier = crypto_CreatePasswordVerifier_Method1(opts.password);
	out.valid = o.verificationBytes === o.verifier;
	if(out.valid) out.insitu = crypto_MakeXorDecryptor(opts.password);
	return o;
}

/* 2.4.117 */
function parse_FilePassHeader(blob, length, oo) {
	var o = oo || {}; o.Info = blob.read_shift(2); blob.l -= 2;
	if(o.Info === 1) o.Data = parse_RC4Header(blob, length);
	else o.Data = parse_RC4CryptoHeader(blob, length);
	return o;
}
function parse_FilePass(blob, length, opts) {
	var o = ({ Type: opts.biff >= 8 ? blob.read_shift(2) : 0 }); /* wEncryptionType */
	if(o.Type) parse_FilePassHeader(blob, length-2, o);
	else parse_XORObfuscation(blob, opts.biff >= 8 ? length : length - 2, opts, o);
	return o;
}


var RTF = (function() {
	function rtf_to_sheet(d, opts) {
		switch(opts.type) {
			case 'base64': return rtf_to_sheet_str(Base64.decode(d), opts);
			case 'binary': return rtf_to_sheet_str(d, opts);
			case 'buffer': return rtf_to_sheet_str(d.toString('binary'), opts);
			case 'array':  return rtf_to_sheet_str(cc2str(d), opts);
		}
		throw new Error("Unrecognized type " + opts.type);
	}

	function rtf_to_sheet_str(str, opts) {
		var o = opts || {};
		var ws = o.dense ? ([]) : ({});
		var range = ({s: {c:0, r:0}, e: {c:0, r:0}});

		// TODO: parse
		if(!str.match(/\\trowd/)) throw new Error("RTF missing table");

		ws['!ref'] = encode_range(range);
		return ws;
	}

	function rtf_to_workbook(d, opts) { return sheet_to_workbook(rtf_to_sheet(d, opts), opts); }

	/* TODO: this is a stub */
	function sheet_to_rtf(ws) {
		var o = ["{\\rtf1\\ansi"];
		var r = safe_decode_range(ws['!ref']), cell;
		var dense = Array.isArray(ws);
		for(var R = r.s.r; R <= r.e.r; ++R) {
			o.push("\\trowd\\trautofit1");
			for(var C = r.s.c; C <= r.e.c; ++C) o.push("\\cellx" + (C+1));
			o.push("\\pard\\intbl");
			for(C = r.s.c; C <= r.e.c; ++C) {
				var coord = encode_cell({r:R,c:C});
				cell = dense ? (ws[R]||[])[C]: ws[coord];
				if(!cell || cell.v == null && (!cell.f || cell.F)) continue;
				o.push(" " + (cell.w || (format_cell(cell), cell.w)));
				o.push("\\cell");
			}
			o.push("\\pard\\intbl\\row");
		}
		return o.join("") + "}";
	}

	return {
		to_workbook: rtf_to_workbook,
		to_sheet: rtf_to_sheet,
		from_sheet: sheet_to_rtf
	};
})();
function hex2RGB(h) {
	var o = h.slice(h[0]==="#"?1:0).slice(0,6);
	return [parseInt(o.slice(0,2),16),parseInt(o.slice(2,4),16),parseInt(o.slice(4,6),16)];
}
function rgb2Hex(rgb) {
	for(var i=0,o=1; i!=3; ++i) o = o*256 + (rgb[i]>255?255:rgb[i]<0?0:rgb[i]);
	return o.toString(16).toUpperCase().slice(1);
}

function rgb2HSL(rgb) {
	var R = rgb[0]/255, G = rgb[1]/255, B=rgb[2]/255;
	var M = Math.max(R, G, B), m = Math.min(R, G, B), C = M - m;
	if(C === 0) return [0, 0, R];

	var H6 = 0, S = 0, L2 = (M + m);
	S = C / (L2 > 1 ? 2 - L2 : L2);
	switch(M){
		case R: H6 = ((G - B) / C + 6)%6; break;
		case G: H6 = ((B - R) / C + 2); break;
		case B: H6 = ((R - G) / C + 4); break;
	}
	return [H6 / 6, S, L2 / 2];
}

function hsl2RGB(hsl){
	var H = hsl[0], S = hsl[1], L = hsl[2];
	var C = S * 2 * (L < 0.5 ? L : 1 - L), m = L - C/2;
	var rgb = [m,m,m], h6 = 6*H;

	var X;
	if(S !== 0) switch(h6|0) {
		case 0: case 6: X = C * h6; rgb[0] += C; rgb[1] += X; break;
		case 1: X = C * (2 - h6);   rgb[0] += X; rgb[1] += C; break;
		case 2: X = C * (h6 - 2);   rgb[1] += C; rgb[2] += X; break;
		case 3: X = C * (4 - h6);   rgb[1] += X; rgb[2] += C; break;
		case 4: X = C * (h6 - 4);   rgb[2] += C; rgb[0] += X; break;
		case 5: X = C * (6 - h6);   rgb[2] += X; rgb[0] += C; break;
	}
	for(var i = 0; i != 3; ++i) rgb[i] = Math.round(rgb[i]*255);
	return rgb;
}

/* 18.8.3 bgColor tint algorithm */
function rgb_tint(hex, tint) {
	if(tint === 0) return hex;
	var hsl = rgb2HSL(hex2RGB(hex));
	if (tint < 0) hsl[2] = hsl[2] * (1 + tint);
	else hsl[2] = 1 - (1 - hsl[2]) * (1 - tint);
	return rgb2Hex(hsl2RGB(hsl));
}

/* 18.3.1.13 width calculations */
/* [MS-OI29500] 2.1.595 Column Width & Formatting */
var DEF_MDW = 6, MAX_MDW = 15, MIN_MDW = 1, MDW = DEF_MDW;
function width2px(width) { return Math.floor(( width + (Math.round(128/MDW))/256 )* MDW ); }
function px2char(px) { return (Math.floor((px - 5)/MDW * 100 + 0.5))/100; }
function char2width(chr) { return (Math.round((chr * MDW + 5)/MDW*256))/256; }
//function px2char_(px) { return (((px - 5)/MDW * 100 + 0.5))/100; }
//function char2width_(chr) { return (((chr * MDW + 5)/MDW*256))/256; }
function cycle_width(collw) { return char2width(px2char(width2px(collw))); }
/* XLSX/XLSB/XLS specify width in units of MDW */
function find_mdw_colw(collw) {
	var delta = Math.abs(collw - cycle_width(collw)), _MDW = MDW;
	if(delta > 0.005) for(MDW=MIN_MDW; MDW<MAX_MDW; ++MDW) if(Math.abs(collw - cycle_width(collw)) <= delta) { delta = Math.abs(collw - cycle_width(collw)); _MDW = MDW; }
	MDW = _MDW;
}
/* XLML specifies width in terms of pixels */
/*function find_mdw_wpx(wpx) {
	var delta = Infinity, guess = 0, _MDW = MIN_MDW;
	for(MDW=MIN_MDW; MDW<MAX_MDW; ++MDW) {
		guess = char2width_(px2char_(wpx))*256;
		guess = (guess) % 1;
		if(guess > 0.5) guess--;
		if(Math.abs(guess) < delta) { delta = Math.abs(guess); _MDW = MDW; }
	}
	MDW = _MDW;
}*/

function process_col(coll) {
	if(coll.width) {
		coll.wpx = width2px(coll.width);
		coll.wch = px2char(coll.wpx);
		coll.MDW = MDW;
	} else if(coll.wpx) {
		coll.wch = px2char(coll.wpx);
		coll.width = char2width(coll.wch);
		coll.MDW = MDW;
	} else if(typeof coll.wch == 'number') {
		coll.width = char2width(coll.wch);
		coll.wpx = width2px(coll.width);
		coll.MDW = MDW;
	}
	if(coll.customWidth) delete coll.customWidth;
}

var DEF_PPI = 96, PPI = DEF_PPI;
function px2pt(px) { return px * 96 / PPI; }
function pt2px(pt) { return pt * PPI / 96; }

/* [MS-EXSPXML3] 2.4.54 ST_enmPattern */
var XLMLPatternTypeMap = {
	"None": "none",
	"Solid": "solid",
	"Gray50": "mediumGray",
	"Gray75": "darkGray",
	"Gray25": "lightGray",
	"HorzStripe": "darkHorizontal",
	"VertStripe": "darkVertical",
	"ReverseDiagStripe": "darkDown",
	"DiagStripe": "darkUp",
	"DiagCross": "darkGrid",
	"ThickDiagCross": "darkTrellis",
	"ThinHorzStripe": "lightHorizontal",
	"ThinVertStripe": "lightVertical",
	"ThinReverseDiagStripe": "lightDown",
	"ThinHorzCross": "lightGrid"
};

/* 18.8.5 borders CT_Borders */
function parse_borders(t, styles, themes, opts) {
	styles.Borders = [];
	var border = {}/*, sub_border = {}*/;
	var pass = false;
	t[0].match(tagregex).forEach(function(x) {
		var y = parsexmltag(x);
		switch(strip_ns(y[0])) {
			case '<borders': case '<borders>': case '</borders>': break;

			/* 18.8.4 border CT_Border */
			case '<border': case '<border>': case '<border/>':
				border = {};
				if (y.diagonalUp) { border.diagonalUp = y.diagonalUp; }
				if (y.diagonalDown) { border.diagonalDown = y.diagonalDown; }
				styles.Borders.push(border);
				break;
			case '</border>': break;

			/* note: not in spec, appears to be CT_BorderPr */
			case '<left/>': break;
			case '<left': case '<left>': break;
			case '</left>': break;

			/* note: not in spec, appears to be CT_BorderPr */
			case '<right/>': break;
			case '<right': case '<right>': break;
			case '</right>': break;

			/* 18.8.43 top CT_BorderPr */
			case '<top/>': break;
			case '<top': case '<top>': break;
			case '</top>': break;

			/* 18.8.6 bottom CT_BorderPr */
			case '<bottom/>': break;
			case '<bottom': case '<bottom>': break;
			case '</bottom>': break;

			/* 18.8.13 diagonal CT_BorderPr */
			case '<diagonal': case '<diagonal>': case '<diagonal/>': break;
			case '</diagonal>': break;

			/* 18.8.25 horizontal CT_BorderPr */
			case '<horizontal': case '<horizontal>': case '<horizontal/>': break;
			case '</horizontal>': break;

			/* 18.8.44 vertical CT_BorderPr */
			case '<vertical': case '<vertical>': case '<vertical/>': break;
			case '</vertical>': break;

			/* 18.8.37 start CT_BorderPr */
			case '<start': case '<start>': case '<start/>': break;
			case '</start>': break;

			/* 18.8.16 end CT_BorderPr */
			case '<end': case '<end>': case '<end/>': break;
			case '</end>': break;

			/* 18.8.? color CT_Color */
			case '<color': case '<color>': break;
			case '<color/>': case '</color>': break;

			/* 18.2.10 extLst CT_ExtensionList ? */
			case '<extLst': case '<extLst>': case '</extLst>': break;
			case '<ext': pass = true; break;
			case '</ext>': pass = false; break;
			default: if(opts && opts.WTF) {
				if(!pass) throw new Error('unrecognized ' + y[0] + ' in borders');
			}
		}
	});
}

/* 18.8.21 fills CT_Fills */
function parse_fills(t, styles, themes, opts) {
	styles.Fills = [];
	var fill = {};
	var pass = false;
	t[0].match(tagregex).forEach(function(x) {
		var y = parsexmltag(x);
		switch(strip_ns(y[0])) {
			case '<fills': case '<fills>': case '</fills>': break;

			/* 18.8.20 fill CT_Fill */
			case '<fill>': case '<fill': case '<fill/>':
				fill = {}; styles.Fills.push(fill); break;
			case '</fill>': break;

			/* 18.8.24 gradientFill CT_GradientFill */
			case '<gradientFill>': break;
			case '<gradientFill':
			case '</gradientFill>': styles.Fills.push(fill); fill = {}; break;

			/* 18.8.32 patternFill CT_PatternFill */
			case '<patternFill': case '<patternFill>':
				if(y.patternType) fill.patternType = y.patternType;
				break;
			case '<patternFill/>': case '</patternFill>': break;

			/* 18.8.3 bgColor CT_Color */
			case '<bgColor':
				if(!fill.bgColor) fill.bgColor = {};
				if(y.indexed) fill.bgColor.indexed = parseInt(y.indexed, 10);
				if(y.theme) fill.bgColor.theme = parseInt(y.theme, 10);
				if(y.tint) fill.bgColor.tint = parseFloat(y.tint);
				/* Excel uses ARGB strings */
				if(y.rgb) fill.bgColor.rgb = y.rgb.slice(-6);
				break;
			case '<bgColor/>': case '</bgColor>': break;

			/* 18.8.19 fgColor CT_Color */
			case '<fgColor':
				if(!fill.fgColor) fill.fgColor = {};
				if(y.theme) fill.fgColor.theme = parseInt(y.theme, 10);
				if(y.tint) fill.fgColor.tint = parseFloat(y.tint);
				/* Excel uses ARGB strings */
				if(y.rgb) fill.fgColor.rgb = y.rgb.slice(-6);
				break;
			case '<fgColor/>': case '</fgColor>': break;

			/* 18.8.38 stop CT_GradientStop */
			case '<stop': case '<stop/>': break;
			case '</stop>': break;

			/* 18.8.? color CT_Color */
			case '<color': case '<color/>': break;
			case '</color>': break;

			/* 18.2.10 extLst CT_ExtensionList ? */
			case '<extLst': case '<extLst>': case '</extLst>': break;
			case '<ext': pass = true; break;
			case '</ext>': pass = false; break;
			default: if(opts && opts.WTF) {
				if(!pass) throw new Error('unrecognized ' + y[0] + ' in fills');
			}
		}
	});
}

/* 18.8.23 fonts CT_Fonts */
function parse_fonts(t, styles, themes, opts) {
	styles.Fonts = [];
	var font = {};
	var pass = false;
	t[0].match(tagregex).forEach(function(x) {
		var y = parsexmltag(x);
		switch(strip_ns(y[0])) {
			case '<fonts': case '<fonts>': case '</fonts>': break;

			/* 18.8.22 font CT_Font */
			case '<font': case '<font>': break;
			case '</font>': case '<font/>':
				styles.Fonts.push(font);
				font = {};
				break;

			/* 18.8.29 name CT_FontName */
			case '<name': if(y.val) font.name = y.val; break;
			case '<name/>': case '</name>': break;

			/* 18.8.2  b CT_BooleanProperty */
			case '<b': font.bold = y.val ? parsexmlbool(y.val) : 1; break;
			case '<b/>': font.bold = 1; break;

			/* 18.8.26 i CT_BooleanProperty */
			case '<i': font.italic = y.val ? parsexmlbool(y.val) : 1; break;
			case '<i/>': font.italic = 1; break;

			/* 18.4.13 u CT_UnderlineProperty */
			case '<u':
				switch(y.val) {
					case "none": font.underline = 0x00; break;
					case "single": font.underline = 0x01; break;
					case "double": font.underline = 0x02; break;
					case "singleAccounting": font.underline = 0x21; break;
					case "doubleAccounting": font.underline = 0x22; break;
				} break;
			case '<u/>': font.underline = 1; break;

			/* 18.4.10 strike CT_BooleanProperty */
			case '<strike': font.strike = y.val ? parsexmlbool(y.val) : 1; break;
			case '<strike/>': font.strike = 1; break;

			/* 18.4.2  outline CT_BooleanProperty */
			case '<outline': font.outline = y.val ? parsexmlbool(y.val) : 1; break;
			case '<outline/>': font.outline = 1; break;

			/* 18.8.36 shadow CT_BooleanProperty */
			case '<shadow': font.shadow = y.val ? parsexmlbool(y.val) : 1; break;
			case '<shadow/>': font.shadow = 1; break;

			/* 18.8.12 condense CT_BooleanProperty */
			case '<condense': font.condense = y.val ? parsexmlbool(y.val) : 1; break;
			case '<condense/>': font.condense = 1; break;

			/* 18.8.17 extend CT_BooleanProperty */
			case '<extend': font.extend = y.val ? parsexmlbool(y.val) : 1; break;
			case '<extend/>': font.extend = 1; break;

			/* 18.4.11 sz CT_FontSize */
			case '<sz': if(y.val) font.sz = +y.val; break;
			case '<sz/>': case '</sz>': break;

			/* 18.4.14 vertAlign CT_VerticalAlignFontProperty */
			case '<vertAlign': if(y.val) font.vertAlign = y.val; break;
			case '<vertAlign/>': case '</vertAlign>': break;

			/* 18.8.18 family CT_FontFamily */
			case '<family': if(y.val) font.family = parseInt(y.val,10); break;
			case '<family/>': case '</family>': break;

			/* 18.8.35 scheme CT_FontScheme */
			case '<scheme': if(y.val) font.scheme = y.val; break;
			case '<scheme/>': case '</scheme>': break;

			/* 18.4.1 charset CT_IntProperty */
			case '<charset':
				if(y.val == '1') break;
				y.codepage = CS2CP[parseInt(y.val, 10)];
				break;

			/* 18.?.? color CT_Color */
			case '<color':
				if(!font.color) font.color = {};
				if(y.auto) font.color.auto = parsexmlbool(y.auto);

				if(y.rgb) font.color.rgb = y.rgb.slice(-6);
				else if(y.indexed) {
					font.color.index = parseInt(y.indexed, 10);
					var icv = XLSIcv[font.color.index];
					if(font.color.index == 81) icv = XLSIcv[1];
					if(!icv) throw new Error(x);
					font.color.rgb = icv[0].toString(16) + icv[1].toString(16) + icv[2].toString(16);
				} else if(y.theme) {
					font.color.theme = parseInt(y.theme, 10);
					if(y.tint) font.color.tint = parseFloat(y.tint);
					if(y.theme && themes.themeElements && themes.themeElements.clrScheme) {
						font.color.rgb = rgb_tint(themes.themeElements.clrScheme[font.color.theme].rgb, font.color.tint || 0);
					}
				}

				break;
			case '<color/>': case '</color>': break;

			/* 18.2.10 extLst CT_ExtensionList ? */
			case '<extLst': case '<extLst>': case '</extLst>': break;
			case '<ext': pass = true; break;
			case '</ext>': pass = false; break;
			default: if(opts && opts.WTF) {
				if(!pass) throw new Error('unrecognized ' + y[0] + ' in fonts');
			}
		}
	});
}

/* 18.8.31 numFmts CT_NumFmts */
function parse_numFmts(t, styles, opts) {
	styles.NumberFmt = [];
	var k/*Array<number>*/ = (keys(SSF._table));
	for(var i=0; i < k.length; ++i) styles.NumberFmt[k[i]] = SSF._table[k[i]];
	var m = t[0].match(tagregex);
	if(!m) return;
	for(i=0; i < m.length; ++i) {
		var y = parsexmltag(m[i]);
		switch(strip_ns(y[0])) {
			case '<numFmts': case '</numFmts>': case '<numFmts/>': case '<numFmts>': break;
			case '<numFmt': {
				var f=unescapexml(utf8read(y.formatCode)), j=parseInt(y.numFmtId,10);
				styles.NumberFmt[j] = f;
				if(j>0) {
					if(j > 0x188) {
						for(j = 0x188; j > 0x3c; --j) if(styles.NumberFmt[j] == null) break;
						styles.NumberFmt[j] = f;
					}
					SSF.load(f,j);
				}
			} break;
			case '</numFmt>': break;
			default: if(opts.WTF) throw new Error('unrecognized ' + y[0] + ' in numFmts');
		}
	}
}

function write_numFmts(NF) {
	var o = ["<numFmts>"];
	[[5,8],[23,26],[41,44],[/*63*/50,/*66],[164,*/392]].forEach(function(r) {
		for(var i = r[0]; i <= r[1]; ++i) if(NF[i] != null) o[o.length] = (writextag('numFmt',null,{numFmtId:i,formatCode:escapexml(NF[i])}));
	});
	if(o.length === 1) return "";
	o[o.length] = ("</numFmts>");
	o[0] = writextag('numFmts', null, { count:o.length-2 }).replace("/>", ">");
	return o.join("");
}

/* 18.8.10 cellXfs CT_CellXfs */
var cellXF_uint = [ "numFmtId", "fillId", "fontId", "borderId", "xfId" ];
var cellXF_bool = [ "applyAlignment", "applyBorder", "applyFill", "applyFont", "applyNumberFormat", "applyProtection", "pivotButton", "quotePrefix" ];
function parse_cellXfs(t, styles, opts) {
	styles.CellXf = [];
	var xf;
	var pass = false;
	t[0].match(tagregex).forEach(function(x) {
		var y = parsexmltag(x), i = 0;
		switch(strip_ns(y[0])) {
			case '<cellXfs': case '<cellXfs>': case '<cellXfs/>': case '</cellXfs>': break;

			/* 18.8.45 xf CT_Xf */
			case '<xf': case '<xf/>':
				xf = y;
				delete xf[0];
				for(i = 0; i < cellXF_uint.length; ++i) if(xf[cellXF_uint[i]])
					xf[cellXF_uint[i]] = parseInt(xf[cellXF_uint[i]], 10);
				for(i = 0; i < cellXF_bool.length; ++i) if(xf[cellXF_bool[i]])
					xf[cellXF_bool[i]] = parsexmlbool(xf[cellXF_bool[i]]);
				if(xf.numFmtId > 0x188) {
					for(i = 0x188; i > 0x3c; --i) if(styles.NumberFmt[xf.numFmtId] == styles.NumberFmt[i]) { xf.numFmtId = i; break; }
				}
				styles.CellXf.push(xf); break;
			case '</xf>': break;

			/* 18.8.1 alignment CT_CellAlignment */
			case '<alignment': case '<alignment/>':
				var alignment = {};
				if(y.vertical) alignment.vertical = y.vertical;
				if(y.horizontal) alignment.horizontal = y.horizontal;
				if(y.textRotation != null) alignment.textRotation = y.textRotation;
				if(y.indent) alignment.indent = y.indent;
				if(y.wrapText) alignment.wrapText = y.wrapText;
				xf.alignment = alignment;
				break;
			case '</alignment>': break;

			/* 18.8.33 protection CT_CellProtection */
			case '<protection': case '</protection>': case '<protection/>': break;

			/* 18.2.10 extLst CT_ExtensionList ? */
			case '<extLst': case '<extLst>': case '</extLst>': break;
			case '<ext': pass = true; break;
			case '</ext>': pass = false; break;
			default: if(opts && opts.WTF) {
				if(!pass) throw new Error('unrecognized ' + y[0] + ' in cellXfs');
			}
		}
	});
}

function write_cellXfs(cellXfs) {
	var o = [];
	o[o.length] = (writextag('cellXfs',null));
	cellXfs.forEach(function(c) { o[o.length] = (writextag('xf', null, c)); });
	o[o.length] = ("</cellXfs>");
	if(o.length === 2) return "";
	o[0] = writextag('cellXfs',null, {count:o.length-2}).replace("/>",">");
	return o.join("");
}

/* 18.8 Styles CT_Stylesheet*/
var parse_sty_xml= (function make_pstyx() {
var numFmtRegex = /<(?:\w+:)?numFmts([^>]*)>[\S\s]*?<\/(?:\w+:)?numFmts>/;
var cellXfRegex = /<(?:\w+:)?cellXfs([^>]*)>[\S\s]*?<\/(?:\w+:)?cellXfs>/;
var fillsRegex = /<(?:\w+:)?fills([^>]*)>[\S\s]*?<\/(?:\w+:)?fills>/;
var fontsRegex = /<(?:\w+:)?fonts([^>]*)>[\S\s]*?<\/(?:\w+:)?fonts>/;
var bordersRegex = /<(?:\w+:)?borders([^>]*)>[\S\s]*?<\/(?:\w+:)?borders>/;

return function parse_sty_xml(data, themes, opts) {
	var styles = {};
	if(!data) return styles;
	data = data.replace(/<!--([\s\S]*?)-->/mg,"").replace(/<!DOCTYPE[^\[]*\[[^\]]*\]>/gm,"");
	/* 18.8.39 styleSheet CT_Stylesheet */
	var t;

	/* 18.8.31 numFmts CT_NumFmts ? */
	if((t=data.match(numFmtRegex))) parse_numFmts(t, styles, opts);

	/* 18.8.23 fonts CT_Fonts ? */
	if((t=data.match(fontsRegex))) parse_fonts(t, styles, themes, opts);

	/* 18.8.21 fills CT_Fills ? */
	if((t=data.match(fillsRegex))) parse_fills(t, styles, themes, opts);

	/* 18.8.5  borders CT_Borders ? */
	if((t=data.match(bordersRegex))) parse_borders(t, styles, themes, opts);

	/* 18.8.9  cellStyleXfs CT_CellStyleXfs ? */

	/* 18.8.10 cellXfs CT_CellXfs ? */
	if((t=data.match(cellXfRegex))) parse_cellXfs(t, styles, opts);

	/* 18.8.8  cellStyles CT_CellStyles ? */
	/* 18.8.15 dxfs CT_Dxfs ? */
	/* 18.8.42 tableStyles CT_TableStyles ? */
	/* 18.8.11 colors CT_Colors ? */
	/* 18.2.10 extLst CT_ExtensionList ? */

	return styles;
};
})();

var STYLES_XML_ROOT = writextag('styleSheet', null, {
	'xmlns': XMLNS.main[0],
	'xmlns:vt': XMLNS.vt
});

RELS.STY = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

function write_sty_xml(wb, opts) {
	var o = [XML_HEADER, STYLES_XML_ROOT], w;
	if(wb.SSF && (w = write_numFmts(wb.SSF)) != null) o[o.length] = w;
	o[o.length] = ('<fonts count="1"><font><sz val="12"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>');
	o[o.length] = ('<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>');
	o[o.length] = ('<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>');
	o[o.length] = ('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>');
	if((w = write_cellXfs(opts.cellXfs))) o[o.length] = (w);
	o[o.length] = ('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>');
	o[o.length] = ('<dxfs count="0"/>');
	o[o.length] = ('<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleMedium4"/>');

	if(o.length>2){ o[o.length] = ('</styleSheet>'); o[1]=o[1].replace("/>",">"); }
	return o.join("");
}
/* [MS-XLSB] 2.4.657 BrtFmt */
function parse_BrtFmt(data, length) {
	var numFmtId = data.read_shift(2);
	var stFmtCode = parse_XLWideString(data,length-2);
	return [numFmtId, stFmtCode];
}
function write_BrtFmt(i, f, o) {
	if(!o) o = new_buf(6 + 4 * f.length);
	o.write_shift(2, i);
	write_XLWideString(f, o);
	var out = (o.length > o.l) ? o.slice(0, o.l) : o;
	if(o.l == null) o.l = o.length;
	return out;
}

/* [MS-XLSB] 2.4.659 BrtFont TODO */
function parse_BrtFont(data, length, opts) {
	var out = ({});

	out.sz = data.read_shift(2) / 20;

	var grbit = parse_FontFlags(data, 2, opts);
	if(grbit.fCondense) out.condense = 1;
	if(grbit.fExtend) out.extend = 1;
	if(grbit.fShadow) out.shadow = 1;
	if(grbit.fOutline) out.outline = 1;
	if(grbit.fStrikeout) out.strike = 1;
	if(grbit.fItalic) out.italic = 1;

	var bls = data.read_shift(2);
	if(bls === 0x02BC) out.bold = 1;

	switch(data.read_shift(2)) {
		/* case 0: out.vertAlign = "baseline"; break; */
		case 1: out.vertAlign = "superscript"; break;
		case 2: out.vertAlign = "subscript"; break;
	}

	var underline = data.read_shift(1);
	if(underline != 0) out.underline = underline;

	var family = data.read_shift(1);
	if(family > 0) out.family = family;

	var bCharSet = data.read_shift(1);
	if(bCharSet > 0) out.charset = bCharSet;

	data.l++;
	out.color = parse_BrtColor(data, 8);

	switch(data.read_shift(1)) {
		/* case 0: out.scheme = "none": break; */
		case 1: out.scheme = "major"; break;
		case 2: out.scheme = "minor"; break;
	}

	out.name = parse_XLWideString(data, length - 21);

	return out;
}
function write_BrtFont(font, o) {
	if(!o) o = new_buf(25+4*32);
	o.write_shift(2, font.sz * 20);
	write_FontFlags(font, o);
	o.write_shift(2, font.bold ? 0x02BC : 0x0190);
	var sss = 0;
	if(font.vertAlign == "superscript") sss = 1;
	else if(font.vertAlign == "subscript") sss = 2;
	o.write_shift(2, sss);
	o.write_shift(1, font.underline || 0);
	o.write_shift(1, font.family || 0);
	o.write_shift(1, font.charset || 0);
	o.write_shift(1, 0);
	write_BrtColor(font.color, o);
	var scheme = 0;
	if(font.scheme == "major") scheme = 1;
	if(font.scheme == "minor") scheme = 2;
	o.write_shift(1, scheme);
	write_XLWideString(font.name, o);
	return o.length > o.l ? o.slice(0, o.l) : o;
}

/* [MS-XLSB] 2.4.650 BrtFill */
var XLSBFillPTNames = [
	"none",
	"solid",
	"mediumGray",
	"darkGray",
	"lightGray",
	"darkHorizontal",
	"darkVertical",
	"darkDown",
	"darkUp",
	"darkGrid",
	"darkTrellis",
	"lightHorizontal",
	"lightVertical",
	"lightDown",
	"lightUp",
	"lightGrid",
	"lightTrellis",
	"gray125",
	"gray0625"
];
var rev_XLSBFillPTNames = (evert(XLSBFillPTNames));
/* TODO: gradient fill representation */
var parse_BrtFill = parsenoop;
function write_BrtFill(fill, o) {
	if(!o) o = new_buf(4*3 + 8*7 + 16*1);
	var fls = rev_XLSBFillPTNames[fill.patternType];
	if(fls == null) fls = 0x28;
	o.write_shift(4, fls);
	var j = 0;
	if(fls != 0x28) {
		/* TODO: custom FG Color */
		write_BrtColor({auto:1}, o);
		/* TODO: custom BG Color */
		write_BrtColor({auto:1}, o);

		for(; j < 12; ++j) o.write_shift(4, 0);
	} else {
		for(; j < 4; ++j) o.write_shift(4, 0);

		for(; j < 12; ++j) o.write_shift(4, 0); /* TODO */
		/* iGradientType */
		/* xnumDegree */
		/* xnumFillToLeft */
		/* xnumFillToRight */
		/* xnumFillToTop */
		/* xnumFillToBottom */
		/* cNumStop */
		/* xfillGradientStop */
	}
	return o.length > o.l ? o.slice(0, o.l) : o;
}

/* [MS-XLSB] 2.4.824 BrtXF */
function parse_BrtXF(data, length) {
	var tgt = data.l + length;
	var ixfeParent = data.read_shift(2);
	var ifmt = data.read_shift(2);
	data.l = tgt;
	return {ixfe:ixfeParent, numFmtId:ifmt };
}
function write_BrtXF(data, ixfeP, o) {
	if(!o) o = new_buf(16);
	o.write_shift(2, ixfeP||0);
	o.write_shift(2, data.numFmtId||0);
	o.write_shift(2, 0); /* iFont */
	o.write_shift(2, 0); /* iFill */
	o.write_shift(2, 0); /* ixBorder */
	o.write_shift(1, 0); /* trot */
	o.write_shift(1, 0); /* indent */
	o.write_shift(1, 0); /* flags */
	o.write_shift(1, 0); /* flags */
	o.write_shift(1, 0); /* xfGrbitAtr */
	o.write_shift(1, 0);
	return o;
}

/* [MS-XLSB] 2.5.4 Blxf TODO */
function write_Blxf(data, o) {
	if(!o) o = new_buf(10);
	o.write_shift(1, 0); /* dg */
	o.write_shift(1, 0);
	o.write_shift(4, 0); /* color */
	o.write_shift(4, 0); /* color */
	return o;
}
/* [MS-XLSB] 2.4.302 BrtBorder TODO */
var parse_BrtBorder = parsenoop;
function write_BrtBorder(border, o) {
	if(!o) o = new_buf(51);
	o.write_shift(1, 0); /* diagonal */
	write_Blxf(null, o); /* top */
	write_Blxf(null, o); /* bottom */
	write_Blxf(null, o); /* left */
	write_Blxf(null, o); /* right */
	write_Blxf(null, o); /* diag */
	return o.length > o.l ? o.slice(0, o.l) : o;
}

/* [MS-XLSB] 2.4.763 BrtStyle TODO */
function write_BrtStyle(style, o) {
	if(!o) o = new_buf(12+4*10);
	o.write_shift(4, style.xfId);
	o.write_shift(2, 1);
	o.write_shift(1, +style.builtinId);
	o.write_shift(1, 0); /* iLevel */
	write_XLNullableWideString(style.name || "", o);
	return o.length > o.l ? o.slice(0, o.l) : o;
}

/* [MS-XLSB] 2.4.272 BrtBeginTableStyles */
function write_BrtBeginTableStyles(cnt, defTableStyle, defPivotStyle) {
	var o = new_buf(4+256*2*4);
	o.write_shift(4, cnt);
	write_XLNullableWideString(defTableStyle, o);
	write_XLNullableWideString(defPivotStyle, o);
	return o.length > o.l ? o.slice(0, o.l) : o;
}

/* [MS-XLSB] 2.1.7.50 Styles */
function parse_sty_bin(data, themes, opts) {
	var styles = {};
	styles.NumberFmt = ([]);
	for(var y in SSF._table) styles.NumberFmt[y] = SSF._table[y];

	styles.CellXf = [];
	styles.Fonts = [];
	var state = [];
	var pass = false;
	recordhopper(data, function hopper_sty(val, R_n, RT) {
		switch(RT) {
			case 0x002C: /* 'BrtFmt' */
				styles.NumberFmt[val[0]] = val[1]; SSF.load(val[1], val[0]);
				break;
			case 0x002B: /* 'BrtFont' */
				styles.Fonts.push(val);
				if(val.color.theme != null && themes && themes.themeElements && themes.themeElements.clrScheme) {
					val.color.rgb = rgb_tint(themes.themeElements.clrScheme[val.color.theme].rgb, val.color.tint || 0);
				}
				break;
			case 0x0401: /* 'BrtKnownFonts' */ break;
			case 0x002D: /* 'BrtFill' */ break;
			case 0x002E: /* 'BrtBorder' */ break;
			case 0x002F: /* 'BrtXF' */
				if(state[state.length - 1] == "BrtBeginCellXFs") {
					styles.CellXf.push(val);
				}
				break;
			case 0x0030: /* 'BrtStyle' */
			case 0x01FB: /* 'BrtDXF' */
			case 0x023C: /* 'BrtMRUColor' */
			case 0x01DB: /* 'BrtIndexedColor': */
				break;

			case 0x0493: /* 'BrtDXF14' */
			case 0x0836: /* 'BrtDXF15' */
			case 0x046A: /* 'BrtSlicerStyleElement' */
			case 0x0200: /* 'BrtTableStyleElement' */
			case 0x082F: /* 'BrtTimelineStyleElement' */
			case 0x0C00: /* 'BrtUid' */
				break;

			case 0x0023: /* 'BrtFRTBegin' */
				pass = true; break;
			case 0x0024: /* 'BrtFRTEnd' */
				pass = false; break;
			case 0x0025: /* 'BrtACBegin' */
				state.push(R_n); break;
			case 0x0026: /* 'BrtACEnd' */
				state.pop(); break;

			default:
				if((R_n||"").indexOf("Begin") > 0) state.push(R_n);
				else if((R_n||"").indexOf("End") > 0) state.pop();
				else if(!pass || opts.WTF) throw new Error("Unexpected record " + RT + " " + R_n);
		}
	});
	return styles;
}

function write_FMTS_bin(ba, NF) {
	if(!NF) return;
	var cnt = 0;
	[[5,8],[23,26],[41,44],[/*63*/50,/*66],[164,*/392]].forEach(function(r) {
for(var i = r[0]; i <= r[1]; ++i) if(NF[i] != null) ++cnt;
	});

	if(cnt == 0) return;
	write_record(ba, "BrtBeginFmts", write_UInt32LE(cnt));
	[[5,8],[23,26],[41,44],[/*63*/50,/*66],[164,*/392]].forEach(function(r) {
for(var i = r[0]; i <= r[1]; ++i) if(NF[i] != null) write_record(ba, "BrtFmt", write_BrtFmt(i, NF[i]));
	});
	write_record(ba, "BrtEndFmts");
}

function write_FONTS_bin(ba) {
	var cnt = 1;

	if(cnt == 0) return;
	write_record(ba, "BrtBeginFonts", write_UInt32LE(cnt));
	write_record(ba, "BrtFont", write_BrtFont({
		sz:12,
		color: {theme:1},
		name: "Calibri",
		family: 2,
		scheme: "minor"
	}));
	/* 1*65491BrtFont [ACFONTS] */
	write_record(ba, "BrtEndFonts");
}

function write_FILLS_bin(ba) {
	var cnt = 2;

	if(cnt == 0) return;
	write_record(ba, "BrtBeginFills", write_UInt32LE(cnt));
	write_record(ba, "BrtFill", write_BrtFill({patternType:"none"}));
	write_record(ba, "BrtFill", write_BrtFill({patternType:"gray125"}));
	/* 1*65431BrtFill */
	write_record(ba, "BrtEndFills");
}

function write_BORDERS_bin(ba) {
	var cnt = 1;

	if(cnt == 0) return;
	write_record(ba, "BrtBeginBorders", write_UInt32LE(cnt));
	write_record(ba, "BrtBorder", write_BrtBorder({}));
	/* 1*65430BrtBorder */
	write_record(ba, "BrtEndBorders");
}

function write_CELLSTYLEXFS_bin(ba) {
	var cnt = 1;
	write_record(ba, "BrtBeginCellStyleXFs", write_UInt32LE(cnt));
	write_record(ba, "BrtXF", write_BrtXF({
		numFmtId:0,
		fontId:0,
		fillId:0,
		borderId:0
	}, 0xFFFF));
	/* 1*65430(BrtXF *FRT) */
	write_record(ba, "BrtEndCellStyleXFs");
}

function write_CELLXFS_bin(ba, data) {
	write_record(ba, "BrtBeginCellXFs", write_UInt32LE(data.length));
	data.forEach(function(c) { write_record(ba, "BrtXF", write_BrtXF(c,0)); });
	/* 1*65430(BrtXF *FRT) */
	write_record(ba, "BrtEndCellXFs");
}

function write_STYLES_bin(ba) {
	var cnt = 1;

	write_record(ba, "BrtBeginStyles", write_UInt32LE(cnt));
	write_record(ba, "BrtStyle", write_BrtStyle({
		xfId:0,
		builtinId:0,
		name:"Normal"
	}));
	/* 1*65430(BrtStyle *FRT) */
	write_record(ba, "BrtEndStyles");
}

function write_DXFS_bin(ba) {
	var cnt = 0;

	write_record(ba, "BrtBeginDXFs", write_UInt32LE(cnt));
	/* *2147483647(BrtDXF *FRT) */
	write_record(ba, "BrtEndDXFs");
}

function write_TABLESTYLES_bin(ba) {
	var cnt = 0;

	write_record(ba, "BrtBeginTableStyles", write_BrtBeginTableStyles(cnt, "TableStyleMedium9", "PivotStyleMedium4"));
	/* *TABLESTYLE */
	write_record(ba, "BrtEndTableStyles");
}

function write_COLORPALETTE_bin() {
	return;
	/* BrtBeginColorPalette [INDEXEDCOLORS] [MRUCOLORS] BrtEndColorPalette */
}

/* [MS-XLSB] 2.1.7.50 Styles */
function write_sty_bin(wb, opts) {
	var ba = buf_array();
	write_record(ba, "BrtBeginStyleSheet");
	write_FMTS_bin(ba, wb.SSF);
	write_FONTS_bin(ba, wb);
	write_FILLS_bin(ba, wb);
	write_BORDERS_bin(ba, wb);
	write_CELLSTYLEXFS_bin(ba, wb);
	write_CELLXFS_bin(ba, opts.cellXfs);
	write_STYLES_bin(ba, wb);
	write_DXFS_bin(ba, wb);
	write_TABLESTYLES_bin(ba, wb);
	write_COLORPALETTE_bin(ba, wb);
	/* FRTSTYLESHEET*/
	write_record(ba, "BrtEndStyleSheet");
	return ba.end();
}
RELS.THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";

/* 20.1.6.2 clrScheme CT_ColorScheme */
function parse_clrScheme(t, themes, opts) {
	themes.themeElements.clrScheme = [];
	var color = {};
	(t[0].match(tagregex)||[]).forEach(function(x) {
		var y = parsexmltag(x);
		switch(y[0]) {
			/* 20.1.6.2 clrScheme (Color Scheme) CT_ColorScheme */
			case '<a:clrScheme': case '</a:clrScheme>': break;

			/* 20.1.2.3.32 srgbClr CT_SRgbColor */
			case '<a:srgbClr':
				color.rgb = y.val; break;

			/* 20.1.2.3.33 sysClr CT_SystemColor */
			case '<a:sysClr':
				color.rgb = y.lastClr; break;

			/* 20.1.4.1.1 accent1 (Accent 1) */
			/* 20.1.4.1.2 accent2 (Accent 2) */
			/* 20.1.4.1.3 accent3 (Accent 3) */
			/* 20.1.4.1.4 accent4 (Accent 4) */
			/* 20.1.4.1.5 accent5 (Accent 5) */
			/* 20.1.4.1.6 accent6 (Accent 6) */
			/* 20.1.4.1.9 dk1 (Dark 1) */
			/* 20.1.4.1.10 dk2 (Dark 2) */
			/* 20.1.4.1.15 folHlink (Followed Hyperlink) */
			/* 20.1.4.1.19 hlink (Hyperlink) */
			/* 20.1.4.1.22 lt1 (Light 1) */
			/* 20.1.4.1.23 lt2 (Light 2) */
			case '<a:dk1>': case '</a:dk1>':
			case '<a:lt1>': case '</a:lt1>':
			case '<a:dk2>': case '</a:dk2>':
			case '<a:lt2>': case '</a:lt2>':
			case '<a:accent1>': case '</a:accent1>':
			case '<a:accent2>': case '</a:accent2>':
			case '<a:accent3>': case '</a:accent3>':
			case '<a:accent4>': case '</a:accent4>':
			case '<a:accent5>': case '</a:accent5>':
			case '<a:accent6>': case '</a:accent6>':
			case '<a:hlink>': case '</a:hlink>':
			case '<a:folHlink>': case '</a:folHlink>':
				if (y[0].charAt(1) === '/') {
					themes.themeElements.clrScheme.push(color);
					color = {};
				} else {
					color.name = y[0].slice(3, y[0].length - 1);
				}
				break;

			default: if(opts && opts.WTF) throw new Error('Unrecognized ' + y[0] + ' in clrScheme');
		}
	});
}

/* 20.1.4.1.18 fontScheme CT_FontScheme */
function parse_fontScheme() { }

/* 20.1.4.1.15 fmtScheme CT_StyleMatrix */
function parse_fmtScheme() { }

var clrsregex = /<a:clrScheme([^>]*)>[\s\S]*<\/a:clrScheme>/;
var fntsregex = /<a:fontScheme([^>]*)>[\s\S]*<\/a:fontScheme>/;
var fmtsregex = /<a:fmtScheme([^>]*)>[\s\S]*<\/a:fmtScheme>/;

/* 20.1.6.10 themeElements CT_BaseStyles */
function parse_themeElements(data, themes, opts) {
	themes.themeElements = {};

	var t;

	[
		/* clrScheme CT_ColorScheme */
		['clrScheme', clrsregex, parse_clrScheme],
		/* fontScheme CT_FontScheme */
		['fontScheme', fntsregex, parse_fontScheme],
		/* fmtScheme CT_StyleMatrix */
		['fmtScheme', fmtsregex, parse_fmtScheme]
	].forEach(function(m) {
		if(!(t=data.match(m[1]))) throw new Error(m[0] + ' not found in themeElements');
		m[2](t, themes, opts);
	});
}

var themeltregex = /<a:themeElements([^>]*)>[\s\S]*<\/a:themeElements>/;

/* 14.2.7 Theme Part */
function parse_theme_xml(data, opts) {
	/* 20.1.6.9 theme CT_OfficeStyleSheet */
	if(!data || data.length === 0) return parse_theme_xml(write_theme());

	var t;
	var themes = {};

	/* themeElements CT_BaseStyles */
	if(!(t=data.match(themeltregex))) throw new Error('themeElements not found in theme');
	parse_themeElements(t[0], themes, opts);

	return themes;
}

function write_theme(Themes, opts) {
	if(opts && opts.themeXLSX) return opts.themeXLSX;
	var o = [XML_HEADER];
	o[o.length] = '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">';
	o[o.length] =  '<a:themeElements>';

	o[o.length] =   '<a:clrScheme name="Office">';
	o[o.length] =    '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>';
	o[o.length] =    '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>';
	o[o.length] =    '<a:dk2><a:srgbClr val="1F497D"/></a:dk2>';
	o[o.length] =    '<a:lt2><a:srgbClr val="EEECE1"/></a:lt2>';
	o[o.length] =    '<a:accent1><a:srgbClr val="4F81BD"/></a:accent1>';
	o[o.length] =    '<a:accent2><a:srgbClr val="C0504D"/></a:accent2>';
	o[o.length] =    '<a:accent3><a:srgbClr val="9BBB59"/></a:accent3>';
	o[o.length] =    '<a:accent4><a:srgbClr val="8064A2"/></a:accent4>';
	o[o.length] =    '<a:accent5><a:srgbClr val="4BACC6"/></a:accent5>';
	o[o.length] =    '<a:accent6><a:srgbClr val="F79646"/></a:accent6>';
	o[o.length] =    '<a:hlink><a:srgbClr val="0000FF"/></a:hlink>';
	o[o.length] =    '<a:folHlink><a:srgbClr val="800080"/></a:folHlink>';
	o[o.length] =   '</a:clrScheme>';

	o[o.length] =   '<a:fontScheme name="Office">';
	o[o.length] =    '<a:majorFont>';
	o[o.length] =     '<a:latin typeface="Cambria"/>';
	o[o.length] =     '<a:ea typeface=""/>';
	o[o.length] =     '<a:cs typeface=""/>';
	o[o.length] =     '<a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>';
	o[o.length] =     '<a:font script="Hang" typeface="맑은 고딕"/>';
	o[o.length] =     '<a:font script="Hans" typeface="宋体"/>';
	o[o.length] =     '<a:font script="Hant" typeface="新細明體"/>';
	o[o.length] =     '<a:font script="Arab" typeface="Times New Roman"/>';
	o[o.length] =     '<a:font script="Hebr" typeface="Times New Roman"/>';
	o[o.length] =     '<a:font script="Thai" typeface="Tahoma"/>';
	o[o.length] =     '<a:font script="Ethi" typeface="Nyala"/>';
	o[o.length] =     '<a:font script="Beng" typeface="Vrinda"/>';
	o[o.length] =     '<a:font script="Gujr" typeface="Shruti"/>';
	o[o.length] =     '<a:font script="Khmr" typeface="MoolBoran"/>';
	o[o.length] =     '<a:font script="Knda" typeface="Tunga"/>';
	o[o.length] =     '<a:font script="Guru" typeface="Raavi"/>';
	o[o.length] =     '<a:font script="Cans" typeface="Euphemia"/>';
	o[o.length] =     '<a:font script="Cher" typeface="Plantagenet Cherokee"/>';
	o[o.length] =     '<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>';
	o[o.length] =     '<a:font script="Tibt" typeface="Microsoft Himalaya"/>';
	o[o.length] =     '<a:font script="Thaa" typeface="MV Boli"/>';
	o[o.length] =     '<a:font script="Deva" typeface="Mangal"/>';
	o[o.length] =     '<a:font script="Telu" typeface="Gautami"/>';
	o[o.length] =     '<a:font script="Taml" typeface="Latha"/>';
	o[o.length] =     '<a:font script="Syrc" typeface="Estrangelo Edessa"/>';
	o[o.length] =     '<a:font script="Orya" typeface="Kalinga"/>';
	o[o.length] =     '<a:font script="Mlym" typeface="Kartika"/>';
	o[o.length] =     '<a:font script="Laoo" typeface="DokChampa"/>';
	o[o.length] =     '<a:font script="Sinh" typeface="Iskoola Pota"/>';
	o[o.length] =     '<a:font script="Mong" typeface="Mongolian Baiti"/>';
	o[o.length] =     '<a:font script="Viet" typeface="Times New Roman"/>';
	o[o.length] =     '<a:font script="Uigh" typeface="Microsoft Uighur"/>';
	o[o.length] =     '<a:font script="Geor" typeface="Sylfaen"/>';
	o[o.length] =    '</a:majorFont>';
	o[o.length] =    '<a:minorFont>';
	o[o.length] =     '<a:latin typeface="Calibri"/>';
	o[o.length] =     '<a:ea typeface=""/>';
	o[o.length] =     '<a:cs typeface=""/>';
	o[o.length] =     '<a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>';
	o[o.length] =     '<a:font script="Hang" typeface="맑은 고딕"/>';
	o[o.length] =     '<a:font script="Hans" typeface="宋体"/>';
	o[o.length] =     '<a:font script="Hant" typeface="新細明體"/>';
	o[o.length] =     '<a:font script="Arab" typeface="Arial"/>';
	o[o.length] =     '<a:font script="Hebr" typeface="Arial"/>';
	o[o.length] =     '<a:font script="Thai" typeface="Tahoma"/>';
	o[o.length] =     '<a:font script="Ethi" typeface="Nyala"/>';
	o[o.length] =     '<a:font script="Beng" typeface="Vrinda"/>';
	o[o.length] =     '<a:font script="Gujr" typeface="Shruti"/>';
	o[o.length] =     '<a:font script="Khmr" typeface="DaunPenh"/>';
	o[o.length] =     '<a:font script="Knda" typeface="Tunga"/>';
	o[o.length] =     '<a:font script="Guru" typeface="Raavi"/>';
	o[o.length] =     '<a:font script="Cans" typeface="Euphemia"/>';
	o[o.length] =     '<a:font script="Cher" typeface="Plantagenet Cherokee"/>';
	o[o.length] =     '<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>';
	o[o.length] =     '<a:font script="Tibt" typeface="Microsoft Himalaya"/>';
	o[o.length] =     '<a:font script="Thaa" typeface="MV Boli"/>';
	o[o.length] =     '<a:font script="Deva" typeface="Mangal"/>';
	o[o.length] =     '<a:font script="Telu" typeface="Gautami"/>';
	o[o.length] =     '<a:font script="Taml" typeface="Latha"/>';
	o[o.length] =     '<a:font script="Syrc" typeface="Estrangelo Edessa"/>';
	o[o.length] =     '<a:font script="Orya" typeface="Kalinga"/>';
	o[o.length] =     '<a:font script="Mlym" typeface="Kartika"/>';
	o[o.length] =     '<a:font script="Laoo" typeface="DokChampa"/>';
	o[o.length] =     '<a:font script="Sinh" typeface="Iskoola Pota"/>';
	o[o.length] =     '<a:font script="Mong" typeface="Mongolian Baiti"/>';
	o[o.length] =     '<a:font script="Viet" typeface="Arial"/>';
	o[o.length] =     '<a:font script="Uigh" typeface="Microsoft Uighur"/>';
	o[o.length] =     '<a:font script="Geor" typeface="Sylfaen"/>';
	o[o.length] =    '</a:minorFont>';
	o[o.length] =   '</a:fontScheme>';

	o[o.length] =   '<a:fmtScheme name="Office">';
	o[o.length] =    '<a:fillStyleLst>';
	o[o.length] =     '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
	o[o.length] =     '<a:gradFill rotWithShape="1">';
	o[o.length] =      '<a:gsLst>';
	o[o.length] =       '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs>';
	o[o.length] =       '<a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs>';
	o[o.length] =       '<a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs>';
	o[o.length] =      '</a:gsLst>';
	o[o.length] =      '<a:lin ang="16200000" scaled="1"/>';
	o[o.length] =     '</a:gradFill>';
	o[o.length] =     '<a:gradFill rotWithShape="1">';
	o[o.length] =      '<a:gsLst>';
	o[o.length] =       '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="100000"/><a:shade val="100000"/><a:satMod val="130000"/></a:schemeClr></a:gs>';
	o[o.length] =       '<a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="50000"/><a:shade val="100000"/><a:satMod val="350000"/></a:schemeClr></a:gs>';
	o[o.length] =      '</a:gsLst>';
	o[o.length] =      '<a:lin ang="16200000" scaled="0"/>';
	o[o.length] =     '</a:gradFill>';
	o[o.length] =    '</a:fillStyleLst>';
	o[o.length] =    '<a:lnStyleLst>';
	o[o.length] =     '<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln>';
	o[o.length] =     '<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>';
	o[o.length] =     '<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>';
	o[o.length] =    '</a:lnStyleLst>';
	o[o.length] =    '<a:effectStyleLst>';
	o[o.length] =     '<a:effectStyle>';
	o[o.length] =      '<a:effectLst>';
	o[o.length] =       '<a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw>';
	o[o.length] =      '</a:effectLst>';
	o[o.length] =     '</a:effectStyle>';
	o[o.length] =     '<a:effectStyle>';
	o[o.length] =      '<a:effectLst>';
	o[o.length] =       '<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw>';
	o[o.length] =      '</a:effectLst>';
	o[o.length] =     '</a:effectStyle>';
	o[o.length] =     '<a:effectStyle>';
	o[o.length] =      '<a:effectLst>';
	o[o.length] =       '<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw>';
	o[o.length] =      '</a:effectLst>';
	o[o.length] =      '<a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d>';
	o[o.length] =      '<a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d>';
	o[o.length] =     '</a:effectStyle>';
	o[o.length] =    '</a:effectStyleLst>';
	o[o.length] =    '<a:bgFillStyleLst>';
	o[o.length] =     '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
	o[o.length] =     '<a:gradFill rotWithShape="1">';
	o[o.length] =      '<a:gsLst>';
	o[o.length] =       '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs>';
	o[o.length] =       '<a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs>';
	o[o.length] =       '<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs>';
	o[o.length] =      '</a:gsLst>';
	o[o.length] =      '<a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path>';
	o[o.length] =     '</a:gradFill>';
	o[o.length] =     '<a:gradFill rotWithShape="1">';
	o[o.length] =      '<a:gsLst>';
	o[o.length] =       '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs>';
	o[o.length] =       '<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs>';
	o[o.length] =      '</a:gsLst>';
	o[o.length] =      '<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>';
	o[o.length] =     '</a:gradFill>';
	o[o.length] =    '</a:bgFillStyleLst>';
	o[o.length] =   '</a:fmtScheme>';
	o[o.length] =  '</a:themeElements>';

	o[o.length] =  '<a:objectDefaults>';
	o[o.length] =   '<a:spDef>';
	o[o.length] =    '<a:spPr/><a:bodyPr/><a:lstStyle/><a:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="2"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></a:style>';
	o[o.length] =   '</a:spDef>';
	o[o.length] =   '<a:lnDef>';
	o[o.length] =    '<a:spPr/><a:bodyPr/><a:lstStyle/><a:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="1"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></a:style>';
	o[o.length] =   '</a:lnDef>';
	o[o.length] =  '</a:objectDefaults>';
	o[o.length] =  '<a:extraClrSchemeLst/>';
	o[o.length] = '</a:theme>';
	return o.join("");
}
/* [MS-XLS] 2.4.326 TODO: payload is a zip file */
function parse_Theme(blob, length, opts) {
	var end = blob.l + length;
	var dwThemeVersion = blob.read_shift(4);
	if(dwThemeVersion === 124226) return;
	if(!opts.cellStyles || !jszip) { blob.l = end; return; }
	var data = blob.slice(blob.l);
	blob.l = end;
	var zip; try { zip = new jszip(data); } catch(e) { return; }
	var themeXML = getzipstr(zip, "theme/theme/theme1.xml", true);
	if(!themeXML) return;
	return parse_theme_xml(themeXML, opts);
}

/* 2.5.49 */
function parse_ColorTheme(blob) { return blob.read_shift(4); }

/* 2.5.155 */
function parse_FullColorExt(blob) {
	var o = {};
	o.xclrType = blob.read_shift(2);
	o.nTintShade = blob.read_shift(2);
	switch(o.xclrType) {
		case 0: blob.l += 4; break;
		case 1: o.xclrValue = parse_IcvXF(blob, 4); break;
		case 2: o.xclrValue = parse_LongRGBA(blob, 4); break;
		case 3: o.xclrValue = parse_ColorTheme(blob, 4); break;
		case 4: blob.l += 4; break;
	}
	blob.l += 8;
	return o;
}

/* 2.5.164 TODO: read 7 bits*/
function parse_IcvXF(blob, length) {
	return parsenoop(blob, length);
}

/* 2.5.280 */
function parse_XFExtGradient(blob, length) {
	return parsenoop(blob, length);
}

/* [MS-XLS] 2.5.108 */
function parse_ExtProp(blob) {
	var extType = blob.read_shift(2);
	var cb = blob.read_shift(2) - 4;
	var o = [extType];
	switch(extType) {
		case 0x04: case 0x05: case 0x07: case 0x08:
		case 0x09: case 0x0A: case 0x0B: case 0x0D:
			o[1] = parse_FullColorExt(blob, cb); break;
		case 0x06: o[1] = parse_XFExtGradient(blob, cb); break;
		case 0x0E: case 0x0F: o[1] = blob.read_shift(cb === 1 ? 1 : 2); break;
		default: throw new Error("Unrecognized ExtProp type: " + extType + " " + cb);
	}
	return o;
}

/* 2.4.355 */
function parse_XFExt(blob, length) {
	var end = blob.l + length;
	blob.l += 2;
	var ixfe = blob.read_shift(2);
	blob.l += 2;
	var cexts = blob.read_shift(2);
	var ext = [];
	while(cexts-- > 0) ext.push(parse_ExtProp(blob, end-blob.l));
	return {ixfe:ixfe, ext:ext};
}

/* xf is an XF, see parse_XFExt for xfext */
function update_xfext(xf, xfext) {
	xfext.forEach(function(xfe) {
		switch(xfe[0]) { /* 2.5.108 extPropData */
			case 0x04: break; /* foreground color */
			case 0x05: break; /* background color */
			case 0x06: break; /* gradient fill */
			case 0x07: break; /* top cell border color */
			case 0x08: break; /* bottom cell border color */
			case 0x09: break; /* left cell border color */
			case 0x0a: break; /* right cell border color */
			case 0x0b: break; /* diagonal cell border color */
			case 0x0d: break; /* text color */
			case 0x0e: break; /* font scheme */
			case 0x0f: break; /* indentation level */
		}
	});
}

/* 18.6 Calculation Chain */
function parse_cc_xml(data) {
	var d = [];
	if(!data) return d;
	var i = 1;
	(data.match(tagregex)||[]).forEach(function(x) {
		var y = parsexmltag(x);
		switch(y[0]) {
			case '<?xml': break;
			/* 18.6.2  calcChain CT_CalcChain 1 */
			case '<calcChain': case '<calcChain>': case '</calcChain>': break;
			/* 18.6.1  c CT_CalcCell 1 */
			case '<c': delete y[0]; if(y.i) i = y.i; else y.i = i; d.push(y); break;
		}
	});
	return d;
}

//function write_cc_xml(data, opts) { }

/* [MS-XLSB] 2.6.4.1 */
function parse_BrtCalcChainItem$(data) {
	var out = {};
	out.i = data.read_shift(4);
	var cell = {};
	cell.r = data.read_shift(4);
	cell.c = data.read_shift(4);
	out.r = encode_cell(cell);
	var flags = data.read_shift(1);
	if(flags & 0x2) out.l = '1';
	if(flags & 0x8) out.a = '1';
	return out;
}

/* 18.6 Calculation Chain */
function parse_cc_bin(data, name, opts) {
	var out = [];
	var pass = false;
	recordhopper(data, function hopper_cc(val, R_n, RT) {
		switch(RT) {
			case 0x003F: /* 'BrtCalcChainItem$' */
				out.push(val); break;

			default:
				if((R_n||"").indexOf("Begin") > 0){/* empty */}
				else if((R_n||"").indexOf("End") > 0){/* empty */}
				else if(!pass || opts.WTF) throw new Error("Unexpected record " + RT + " " + R_n);
		}
	});
	return out;
}

//function write_cc_bin(data, opts) { }
/* 18.14 Supplementary Workbook Data */
function parse_xlink_xml() {
	//var opts = _opts || {};
	//if(opts.WTF) throw "XLSX External Link";
}

/* [MS-XLSB] 2.1.7.25 External Link */
function parse_xlink_bin(data, name, _opts) {
	if(!data) return data;
	var opts = _opts || {};

	var pass = false, end = false;

	recordhopper(data, function xlink_parse(val, R_n, RT) {
		if(end) return;
		switch(RT) {
			case 0x0167: /* 'BrtSupTabs' */
			case 0x016B: /* 'BrtExternTableStart' */
			case 0x016C: /* 'BrtExternTableEnd' */
			case 0x016E: /* 'BrtExternRowHdr' */
			case 0x016F: /* 'BrtExternCellBlank' */
			case 0x0170: /* 'BrtExternCellReal' */
			case 0x0171: /* 'BrtExternCellBool' */
			case 0x0172: /* 'BrtExternCellError' */
			case 0x0173: /* 'BrtExternCellString' */
			case 0x01D8: /* 'BrtExternValueMeta' */
			case 0x0241: /* 'BrtSupNameStart' */
			case 0x0242: /* 'BrtSupNameValueStart' */
			case 0x0243: /* 'BrtSupNameValueEnd' */
			case 0x0244: /* 'BrtSupNameNum' */
			case 0x0245: /* 'BrtSupNameErr' */
			case 0x0246: /* 'BrtSupNameSt' */
			case 0x0247: /* 'BrtSupNameNil' */
			case 0x0248: /* 'BrtSupNameBool' */
			case 0x0249: /* 'BrtSupNameFmla' */
			case 0x024A: /* 'BrtSupNameBits' */
			case 0x024B: /* 'BrtSupNameEnd' */
				break;

			case 0x0023: /* 'BrtFRTBegin' */
				pass = true; break;
			case 0x0024: /* 'BrtFRTEnd' */
				pass = false; break;

			default:
				if((R_n||"").indexOf("Begin") > 0){/* empty */}
				else if((R_n||"").indexOf("End") > 0){/* empty */}
				else if(!pass || opts.WTF) throw new Error("Unexpected record " + RT.toString(16) + " " + R_n);
		}
	}, opts);
}
RELS.IMG = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
RELS.DRAW = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
/* 20.5 DrawingML - SpreadsheetML Drawing */
function parse_drawing(data, rels) {
	if(!data) return "??";
	/*
	  Chartsheet Drawing:
	   - 20.5.2.35 wsDr CT_Drawing
	    - 20.5.2.1  absoluteAnchor CT_AbsoluteAnchor
	     - 20.5.2.16 graphicFrame CT_GraphicalObjectFrame
	      - 20.1.2.2.16 graphic CT_GraphicalObject
	       - 20.1.2.2.17 graphicData CT_GraphicalObjectData
          - chart reference
	   the actual type is based on the URI of the graphicData
		TODO: handle embedded charts and other types of graphics
	*/
	var id = (data.match(/<c:chart [^>]*r:id="([^"]*)"/)||["",""])[1];

	return rels['!id'][id].Target;
}

/* L.5.5.2 SpreadsheetML Comments + VML Schema */
var _shapeid = 1024;
function write_comments_vml(rId, comments) {
	var csize = [21600, 21600];
	/* L.5.2.1.2 Path Attribute */
	var bbox = ["m0,0l0",csize[1],csize[0],csize[1],csize[0],"0xe"].join(",");
	var o = [
		writextag("xml", null, { 'xmlns:v': XLMLNS.v, 'xmlns:o': XLMLNS.o, 'xmlns:x': XLMLNS.x, 'xmlns:mv': XLMLNS.mv }).replace(/\/>/,">"),
		writextag("o:shapelayout", writextag("o:idmap", null, {'v:ext':"edit", 'data':rId}), {'v:ext':"edit"}),
		writextag("v:shapetype", [
			writextag("v:stroke", null, {joinstyle:"miter"}),
			writextag("v:path", null, {gradientshapeok:"t", 'o:connecttype':"rect"})
		].join(""), {id:"_x0000_t202", 'o:spt':202, coordsize:csize.join(","),path:bbox})
	];
	while(_shapeid < rId * 1000) _shapeid += 1000;

	comments.forEach(function(x) { var c = decode_cell(x[0]);
	o = o.concat([
	'<v:shape' + wxt_helper({
		id:'_x0000_s' + (++_shapeid),
		type:"#_x0000_t202",
		style:"position:absolute; margin-left:80pt;margin-top:5pt;width:104pt;height:64pt;z-index:10" + (x[1].hidden ? ";visibility:hidden" : "") ,
		fillcolor:"#ECFAD4",
		strokecolor:"#edeaa1"
	}) + '>',
		writextag('v:fill', writextag("o:fill", null, {type:"gradientUnscaled", 'v:ext':"view"}), {'color2':"#BEFF82", 'angle':"-180", 'type':"gradient"}),
		writextag("v:shadow", null, {on:"t", 'obscured':"t"}),
		writextag("v:path", null, {'o:connecttype':"none"}),
		'<v:textbox><div style="text-align:left"></div></v:textbox>',
		'<x:ClientData ObjectType="Note">',
			'<x:MoveWithCells/>',
			'<x:SizeWithCells/>',
			/* Part 4 19.4.2.3 Anchor (Anchor) */
			writetag('x:Anchor', [c.c, 0, c.r, 0, c.c+3, 100, c.r+5, 100].join(",")),
			writetag('x:AutoFill', "False"),
			writetag('x:Row', String(c.r)),
			writetag('x:Column', String(c.c)),
			x[1].hidden ? '' : '<x:Visible/>',
		'</x:ClientData>',
	'</v:shape>'
	]); });
	o.push('</xml>');
	return o.join("");
}

RELS.CMNT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";

function parse_comments(zip, dirComments, sheets, sheetRels, opts) {
	for(var i = 0; i != dirComments.length; ++i) {
		var canonicalpath=dirComments[i];
		var comments=parse_cmnt(getzipdata(zip, canonicalpath.replace(/^\//,''), true), canonicalpath, opts);
		if(!comments || !comments.length) continue;
		// find the sheets targeted by these comments
		var sheetNames = keys(sheets);
		for(var j = 0; j != sheetNames.length; ++j) {
			var sheetName = sheetNames[j];
			var rels = sheetRels[sheetName];
			if(rels) {
				var rel = rels[canonicalpath];
				if(rel) insertCommentsIntoSheet(sheetName, sheets[sheetName], comments);
			}
		}
	}
}

function insertCommentsIntoSheet(sheetName, sheet, comments) {
	var dense = Array.isArray(sheet);
	var cell;
	comments.forEach(function(comment) {
		var r = decode_cell(comment.ref);
		if(dense) {
			if(!sheet[r.r]) sheet[r.r] = [];
			cell = sheet[r.r][r.c];
		} else cell = sheet[comment.ref];
		if (!cell) {
			cell = {};
			if(dense) sheet[r.r][r.c] = cell;
			else sheet[comment.ref] = cell;
			var range = safe_decode_range(sheet["!ref"]||"BDWGO1000001:A1");
			if(range.s.r > r.r) range.s.r = r.r;
			if(range.e.r < r.r) range.e.r = r.r;
			if(range.s.c > r.c) range.s.c = r.c;
			if(range.e.c < r.c) range.e.c = r.c;
			var encoded = encode_range(range);
			if (encoded !== sheet["!ref"]) sheet["!ref"] = encoded;
		}

		if (!cell.c) cell.c = [];
		var o = ({a: comment.author, t: comment.t, r: comment.r});
		if(comment.h) o.h = comment.h;
		cell.c.push(o);
	});
}

/* 18.7 Comments */
function parse_comments_xml(data, opts) {
	/* 18.7.6 CT_Comments */
	if(data.match(/<(?:\w+:)?comments *\/>/)) return [];
	var authors = [];
	var commentList = [];
	var authtag = data.match(/<(?:\w+:)?authors>([\s\S]*)<\/(?:\w+:)?authors>/);
	if(authtag && authtag[1]) authtag[1].split(/<\/\w*:?author>/).forEach(function(x) {
		if(x === "" || x.trim() === "") return;
		var a = x.match(/<(?:\w+:)?author[^>]*>(.*)/);
		if(a) authors.push(a[1]);
	});
	var cmnttag = data.match(/<(?:\w+:)?commentList>([\s\S]*)<\/(?:\w+:)?commentList>/);
	if(cmnttag && cmnttag[1]) cmnttag[1].split(/<\/\w*:?comment>/).forEach(function(x) {
		if(x === "" || x.trim() === "") return;
		var cm = x.match(/<(?:\w+:)?comment[^>]*>/);
		if(!cm) return;
		var y = parsexmltag(cm[0]);
		var comment = ({ author: y.authorId && authors[y.authorId] || "sheetjsghost", ref: y.ref, guid: y.guid });
		var cell = decode_cell(y.ref);
		if(opts.sheetRows && opts.sheetRows <= cell.r) return;
		var textMatch = x.match(/<(?:\w+:)?text>([\s\S]*)<\/(?:\w+:)?text>/);
		var rt = !!textMatch && !!textMatch[1] && parse_si(textMatch[1]) || {r:"",t:"",h:""};
		comment.r = rt.r;
		if(rt.r == "<t></t>") rt.t = rt.h = "";
		comment.t = rt.t.replace(/\r\n/g,"\n").replace(/\r/g,"\n");
		if(opts.cellHTML) comment.h = rt.h;
		commentList.push(comment);
	});
	return commentList;
}

var CMNT_XML_ROOT = writextag('comments', null, { 'xmlns': XMLNS.main[0] });
function write_comments_xml(data) {
	var o = [XML_HEADER, CMNT_XML_ROOT];

	var iauthor = [];
	o.push("<authors>");
	data.forEach(function(x) { x[1].forEach(function(w) { var a = escapexml(w.a);
		if(iauthor.indexOf(a) > -1) return;
		iauthor.push(a);
		o.push("<author>" + a + "</author>");
	}); });
	o.push("</authors>");
	o.push("<commentList>");
	data.forEach(function(d) {
		d[1].forEach(function(c) {
			/* 18.7.3 CT_Comment */
			o.push('<comment ref="' + d[0] + '" authorId="' + iauthor.indexOf(escapexml(c.a)) + '"><text>');
			o.push(writetag("t", c.t == null ? "" : escapexml(c.t)));
			o.push('</text></comment>');
		});
	});
	o.push("</commentList>");
	if(o.length>2) { o[o.length] = ('</comments>'); o[1]=o[1].replace("/>",">"); }
	return o.join("");
}
/* [MS-XLSB] 2.4.28 BrtBeginComment */
function parse_BrtBeginComment(data) {
	var out = {};
	out.iauthor = data.read_shift(4);
	var rfx = parse_UncheckedRfX(data, 16);
	out.rfx = rfx.s;
	out.ref = encode_cell(rfx.s);
	data.l += 16; /*var guid = parse_GUID(data); */
	return out;
}
function write_BrtBeginComment(data, o) {
	if(o == null) o = new_buf(36);
	o.write_shift(4, data[1].iauthor);
	write_UncheckedRfX((data[0]), o);
	o.write_shift(4, 0);
	o.write_shift(4, 0);
	o.write_shift(4, 0);
	o.write_shift(4, 0);
	return o;
}

/* [MS-XLSB] 2.4.327 BrtCommentAuthor */
var parse_BrtCommentAuthor = parse_XLWideString;
function write_BrtCommentAuthor(data) { return write_XLWideString(data.slice(0, 54)); }

/* [MS-XLSB] 2.1.7.8 Comments */
function parse_comments_bin(data, opts) {
	var out = [];
	var authors = [];
	var c = {};
	var pass = false;
	recordhopper(data, function hopper_cmnt(val, R_n, RT) {
		switch(RT) {
			case 0x0278: /* 'BrtCommentAuthor' */
				authors.push(val); break;
			case 0x027B: /* 'BrtBeginComment' */
				c = val; break;
			case 0x027D: /* 'BrtCommentText' */
				c.t = val.t; c.h = val.h; c.r = val.r; break;
			case 0x027C: /* 'BrtEndComment' */
				c.author = authors[c.iauthor];
				delete c.iauthor;
				if(opts.sheetRows && opts.sheetRows <= c.rfx.r) break;
				if(!c.t) c.t = "";
				delete c.rfx; out.push(c); break;

			case 0x0C00: /* 'BrtUid' */
				break;

			case 0x0023: /* 'BrtFRTBegin' */
				pass = true; break;
			case 0x0024: /* 'BrtFRTEnd' */
				pass = false; break;
			case 0x0025: /* 'BrtACBegin' */ break;
			case 0x0026: /* 'BrtACEnd' */ break;


			default:
				if((R_n||"").indexOf("Begin") > 0){/* empty */}
				else if((R_n||"").indexOf("End") > 0){/* empty */}
				else if(!pass || opts.WTF) throw new Error("Unexpected record " + RT + " " + R_n);
		}
	});
	return out;
}

function write_comments_bin(data) {
	var ba = buf_array();
	var iauthor = [];
	write_record(ba, "BrtBeginComments");

	write_record(ba, "BrtBeginCommentAuthors");
	data.forEach(function(comment) {
		comment[1].forEach(function(c) {
			if(iauthor.indexOf(c.a) > -1) return;
			iauthor.push(c.a.slice(0,54));
			write_record(ba, "BrtCommentAuthor", write_BrtCommentAuthor(c.a));
		});
	});
	write_record(ba, "BrtEndCommentAuthors");

	write_record(ba, "BrtBeginCommentList");
	data.forEach(function(comment) {
		comment[1].forEach(function(c) {
			c.iauthor = iauthor.indexOf(c.a);
			var range = {s:decode_cell(comment[0]),e:decode_cell(comment[0])};
			write_record(ba, "BrtBeginComment", write_BrtBeginComment([range, c]));
			if(c.t && c.t.length > 0) write_record(ba, "BrtCommentText", write_BrtCommentText(c));
			write_record(ba, "BrtEndComment");
			delete c.iauthor;
		});
	});
	write_record(ba, "BrtEndCommentList");

	write_record(ba, "BrtEndComments");
	return ba.end();
}
var CT_VBA = "application/vnd.ms-office.vbaProject";
function make_vba_xls(cfb) {
	var newcfb = CFB.utils.cfb_new({root:"R"});
	cfb.FullPaths.forEach(function(p, i) {
		if(p.slice(-1) === "/" || !p.match(/_VBA_PROJECT_CUR/)) return;
		var newpath = p.replace(/^[^\/]*/,"R").replace(/\/_VBA_PROJECT_CUR\u0000*/, "");
		CFB.utils.cfb_add(newcfb, newpath, cfb.FileIndex[i].content);
	});
	return CFB.write(newcfb);
}

function fill_vba_xls(cfb, vba) {
	vba.FullPaths.forEach(function(p, i) {
		if(i == 0) return;
		var newpath = p.replace(/[^\/]*[\/]/, "/_VBA_PROJECT_CUR/");
		if(newpath.slice(-1) !== "/") CFB.utils.cfb_add(cfb, newpath, vba.FileIndex[i].content);
	});
}

var VBAFMTS = [ "xlsb", "xlsm", "xlam", "biff8", "xla" ];

RELS.DS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet";
RELS.MS = "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet";

/* macro and dialog sheet stubs */
function parse_ds_bin() { return {'!type':'dialog'}; }
function parse_ds_xml() { return {'!type':'dialog'}; }
function parse_ms_bin() { return {'!type':'macro'}; }
function parse_ms_xml() { return {'!type':'macro'}; }
/* TODO: it will be useful to parse the function str */
var rc_to_a1 = (function(){
	var rcregex = /(^|[^A-Za-z])R(\[?)(-?\d+|)\]?C(\[?)(-?\d+|)\]?/g;
	var rcbase = ({r:0,c:0});
	function rcfunc($$,$1,$2,$3,$4,$5) {
		var R = $3.length>0?parseInt($3,10)|0:0, C = $5.length>0?parseInt($5,10)|0:0;
		if(C<0 && $4.length === 0) C=0;
		var cRel = false, rRel = false;
		if($4.length > 0 || $5.length == 0) cRel = true; if(cRel) C += rcbase.c; else --C;
		if($2.length > 0 || $3.length == 0) rRel = true; if(rRel) R += rcbase.r; else --R;
		return $1 + (cRel ? "" : "$") + encode_col(C) + (rRel ? "" : "$") + encode_row(R);
	}
	return function rc_to_a1(fstr, base) {
		rcbase = base;
		return fstr.replace(rcregex, rcfunc);
	};
})();

var crefregex = /(^|[^._A-Z0-9])([$]?)([A-Z]{1,2}|[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D])([$]?)([1-9]\d{0,5}|10[0-3]\d{4}|104[0-7]\d{3}|1048[0-4]\d{2}|10485[0-6]\d|104857[0-6])(?![_.\(A-Za-z0-9])/g;
var a1_to_rc =(function(){
	return function a1_to_rc(fstr, base) {
		return fstr.replace(crefregex, function($0, $1, $2, $3, $4, $5) {
			var c = decode_col($3) - ($2 ? 0 : base.c);
			var r = decode_row($5) - ($4 ? 0 : base.r);
			var R = (r == 0 ? "" : !$4 ? "[" + r + "]" : (r+1));
			var C = (c == 0 ? "" : !$2 ? "[" + c + "]" : (c+1));
			return $1 + "R" + R + "C" + C;
		});
	};
})();

/* no defined name can collide with a valid cell address A1:XFD1048576 ... except LOG10! */
function shift_formula_str(f, delta) {
	return f.replace(crefregex, function($0, $1, $2, $3, $4, $5) {
		return $1+($2=="$" ? $2+$3 : encode_col(decode_col($3)+delta.c))+($4=="$" ? $4+$5 : encode_row(decode_row($5) + delta.r));
	});
}

function shift_formula_xlsx(f, range, cell) {
	var r = decode_range(range), s = r.s, c = decode_cell(cell);
	var delta = {r:c.r - s.r, c:c.c - s.c};
	return shift_formula_str(f, delta);
}

/* TODO: parse formula */
function fuzzyfmla(f) {
	if(f.length == 1) return false;
	return true;
}

function _xlfn(f) {
	return f.replace(/_xlfn\./g,"");
}
function parseread1(blob) { blob.l+=1; return; }

/* [MS-XLS] 2.5.51 */
function parse_ColRelU(blob, length) {
	var c = blob.read_shift(length == 1 ? 1 : 2);
	return [c & 0x3FFF, (c >> 14) & 1, (c >> 15) & 1];
}

/* [MS-XLS] 2.5.198.105 ; [MS-XLSB] 2.5.97.89 */
function parse_RgceArea(blob, length, opts) {
	var w = 2;
	if(opts) {
		if(opts.biff >= 2 && opts.biff <= 5) return parse_RgceArea_BIFF2(blob, length, opts);
		else if(opts.biff == 12) w = 4;
	}
	var r=blob.read_shift(w), R=blob.read_shift(w);
	var c=parse_ColRelU(blob, 2);
	var C=parse_ColRelU(blob, 2);
	return { s:{r:r, c:c[0], cRel:c[1], rRel:c[2]}, e:{r:R, c:C[0], cRel:C[1], rRel:C[2]} };
}
/* BIFF 2-5 encodes flags in the row field */
function parse_RgceArea_BIFF2(blob) {
	var r=parse_ColRelU(blob, 2), R=parse_ColRelU(blob, 2);
	var c=blob.read_shift(1);
	var C=blob.read_shift(1);
	return { s:{r:r[0], c:c, cRel:r[1], rRel:r[2]}, e:{r:R[0], c:C, cRel:R[1], rRel:R[2]} };
}

/* [MS-XLS] 2.5.198.105 ; [MS-XLSB] 2.5.97.90 */
function parse_RgceAreaRel(blob, length, opts) {
	if(opts.biff < 8) return parse_RgceArea_BIFF2(blob, length, opts);
	var r=blob.read_shift(opts.biff == 12 ? 4 : 2), R=blob.read_shift(opts.biff == 12 ? 4 : 2);
	var c=parse_ColRelU(blob, 2);
	var C=parse_ColRelU(blob, 2);
	return { s:{r:r, c:c[0], cRel:c[1], rRel:c[2]}, e:{r:R, c:C[0], cRel:C[1], rRel:C[2]} };
}

/* [MS-XLS] 2.5.198.109 ; [MS-XLSB] 2.5.97.91 */
function parse_RgceLoc(blob, length, opts) {
	if(opts && opts.biff >= 2 && opts.biff <= 5) return parse_RgceLoc_BIFF2(blob, length, opts);
	var r = blob.read_shift(opts && opts.biff == 12 ? 4 : 2);
	var c = parse_ColRelU(blob, 2);
	return {r:r, c:c[0], cRel:c[1], rRel:c[2]};
}
function parse_RgceLoc_BIFF2(blob) {
	var r = parse_ColRelU(blob, 2);
	var c = blob.read_shift(1);
	return {r:r[0], c:c, cRel:r[1], rRel:r[2]};
}

/* [MS-XLS] 2.5.198.107, 2.5.47 */
function parse_RgceElfLoc(blob) {
	var r = blob.read_shift(2);
	var c = blob.read_shift(2);
	return {r:r, c:c & 0xFF, fQuoted:!!(c & 0x4000), cRel:c>>15, rRel:c>>15 };
}

/* [MS-XLS] 2.5.198.111 ; [MS-XLSB] 2.5.97.92 TODO */
function parse_RgceLocRel(blob, length, opts) {
	var biff = opts && opts.biff ? opts.biff : 8;
	if(biff >= 2 && biff <= 5) return parse_RgceLocRel_BIFF2(blob, length, opts);
	var r = blob.read_shift(biff >= 12 ? 4 : 2);
	var cl = blob.read_shift(2);
	var cRel = (cl & 0x4000) >> 14, rRel = (cl & 0x8000) >> 15;
	cl &= 0x3FFF;
	if(rRel == 1) while(r > 0x7FFFF) r -= 0x100000;
	if(cRel == 1) while(cl > 0x1FFF) cl = cl - 0x4000;
	return {r:r,c:cl,cRel:cRel,rRel:rRel};
}
function parse_RgceLocRel_BIFF2(blob) {
	var rl = blob.read_shift(2);
	var c = blob.read_shift(1);
	var rRel = (rl & 0x8000) >> 15, cRel = (rl & 0x4000) >> 14;
	rl &= 0x3FFF;
	if(rRel == 1 && rl >= 0x2000) rl = rl - 0x4000;
	if(cRel == 1 && c >= 0x80) c = c - 0x100;
	return {r:rl,c:c,cRel:cRel,rRel:rRel};
}

/* [MS-XLS] 2.5.198.27 ; [MS-XLSB] 2.5.97.18 */
function parse_PtgArea(blob, length, opts) {
	var type = (blob[blob.l++] & 0x60) >> 5;
	var area = parse_RgceArea(blob, opts.biff >= 2 && opts.biff <= 5 ? 6 : 8, opts);
	return [type, area];
}

/* [MS-XLS] 2.5.198.28 ; [MS-XLSB] 2.5.97.19 */
function parse_PtgArea3d(blob, length, opts) {
	var type = (blob[blob.l++] & 0x60) >> 5;
	var ixti = blob.read_shift(2, 'i');
	var w = 8;
	if(opts) switch(opts.biff) {
		case 5: blob.l += 12; w = 6; break;
		case 12: w = 12; break;
	}
	var area = parse_RgceArea(blob, w, opts);
	return [type, ixti, area];
}

/* [MS-XLS] 2.5.198.29 ; [MS-XLSB] 2.5.97.20 */
function parse_PtgAreaErr(blob, length, opts) {
	var type = (blob[blob.l++] & 0x60) >> 5;
	blob.l += opts && (opts.biff > 8) ? 12 : (opts.biff < 8 ? 6 : 8);
	return [type];
}
/* [MS-XLS] 2.5.198.30 ; [MS-XLSB] 2.5.97.21 */
function parse_PtgAreaErr3d(blob, length, opts) {
	var type = (blob[blob.l++] & 0x60) >> 5;
	var ixti = blob.read_shift(2);
	var w = 8;
	if(opts) switch(opts.biff) {
		case 5: blob.l += 12; w = 6; break;
		case 12: w = 12; break;
	}
	blob.l += w;
	return [type, ixti];
}

/* [MS-XLS] 2.5.198.31 ; [MS-XLSB] 2.5.97.22 */
function parse_PtgAreaN(blob, length, opts) {
	var type = (blob[blob.l++] & 0x60) >> 5;
	var area = parse_RgceAreaRel(blob, length - 1, opts);
	return [type, area];
}

/* [MS-XLS] 2.5.198.32 ; [MS-XLSB] 2.5.97.23 */
function parse_PtgArray(blob, length, opts) {
	var type = (blob[blob.l++] & 0x60) >> 5;
	blob.l += opts.biff == 2 ? 6 : opts.biff == 12 ? 14 : 7;
	return [type];
}

/* [MS-XLS] 2.5.198.33 ; [MS-XLSB] 2.5.97.24 */
function parse_PtgAttrBaxcel(blob) {
	var bitSemi = blob[blob.l+1] & 0x01; /* 1 = volatile */
	var bitBaxcel = 1;
	blob.l += 4;
	return [bitSemi, bitBaxcel];
}

/* [MS-XLS] 2.5.198.34 ; [MS-XLSB] 2.5.97.25 */
function parse_PtgAttrChoose(blob, length, opts) {
	blob.l +=2;
	var offset = blob.read_shift(opts && opts.biff == 2 ? 1 : 2);
	var o = [];
	/* offset is 1 less than the number of elements */
	for(var i = 0; i <= offset; ++i) o.push(blob.read_shift(opts && opts.biff == 2 ? 1 : 2));
	return o;
}

/* [MS-XLS] 2.5.198.35 ; [MS-XLSB] 2.5.97.26 */
function parse_PtgAttrGoto(blob, length, opts) {
	var bitGoto = (blob[blob.l+1] & 0xFF) ? 1 : 0;
	blob.l += 2;
	return [bitGoto, blob.read_shift(opts && opts.biff == 2 ? 1 : 2)];
}

/* [MS-XLS] 2.5.198.36 ; [MS-XLSB] 2.5.97.27 */
function parse_PtgAttrIf(blob, length, opts) {
	var bitIf = (blob[blob.l+1] & 0xFF) ? 1 : 0;
	blob.l += 2;
	return [bitIf, blob.read_shift(opts && opts.biff == 2 ? 1 : 2)];
}

/* [MS-XLSB] 2.5.97.28 */
function parse_PtgAttrIfError(blob) {
	var bitIf = (blob[blob.l+1] & 0xFF) ? 1 : 0;
	blob.l += 2;
	return [bitIf, blob.read_shift(2)];
}

/* [MS-XLS] 2.5.198.37 ; [MS-XLSB] 2.5.97.29 */
function parse_PtgAttrSemi(blob, length, opts) {
	var bitSemi = (blob[blob.l+1] & 0xFF) ? 1 : 0;
	blob.l += opts && opts.biff == 2 ? 3 : 4;
	return [bitSemi];
}

/* [MS-XLS] 2.5.198.40 ; [MS-XLSB] 2.5.97.32 */
function parse_PtgAttrSpaceType(blob) {
	var type = blob.read_shift(1), cch = blob.read_shift(1);
	return [type, cch];
}

/* [MS-XLS] 2.5.198.38 ; [MS-XLSB] 2.5.97.30 */
function parse_PtgAttrSpace(blob) {
	blob.read_shift(2);
	return parse_PtgAttrSpaceType(blob, 2);
}

/* [MS-XLS] 2.5.198.39 ; [MS-XLSB] 2.5.97.31 */
function parse_PtgAttrSpaceSemi(blob) {
	blob.read_shift(2);
	return parse_PtgAttrSpaceType(blob, 2);
}

/* [MS-XLS] 2.5.198.84 ; [MS-XLSB] 2.5.97.68 TODO */
function parse_PtgRef(blob, length, opts) {
	//var ptg = blob[blob.l] & 0x1F;
	var type = (blob[blob.l] & 0x60)>>5;
	blob.l += 1;
	var loc = parse_RgceLoc(blob, 0, opts);
	return [type, loc];
}

/* [MS-XLS] 2.5.198.88 ; [MS-XLSB] 2.5.97.72 TODO */
function parse_PtgRefN(blob, length, opts) {
	var type = (blob[blob.l] & 0x60)>>5;
	blob.l += 1;
	var loc = parse_RgceLocRel(blob, 0, opts);
	return [type, loc];
}

/* [MS-XLS] 2.5.198.85 ; [MS-XLSB] 2.5.97.69 TODO */
function parse_PtgRef3d(blob, length, opts) {
	var type = (blob[blob.l] & 0x60)>>5;
	blob.l += 1;
	var ixti = blob.read_shift(2); // XtiIndex
	if(opts && opts.biff == 5) blob.l += 12;
	var loc = parse_RgceLoc(blob, 0, opts); // TODO: or RgceLocRel
	return [type, ixti, loc];
}


/* [MS-XLS] 2.5.198.62 ; [MS-XLSB] 2.5.97.45 TODO */
function parse_PtgFunc(blob, length, opts) {
	//var ptg = blob[blob.l] & 0x1F;
	var type = (blob[blob.l] & 0x60)>>5;
	blob.l += 1;
	var iftab = blob.read_shift(opts && opts.biff <= 3 ? 1 : 2);
	return [FtabArgc[iftab], Ftab[iftab], type];
}
/* [MS-XLS] 2.5.198.63 ; [MS-XLSB] 2.5.97.46 TODO */
function parse_PtgFuncVar(blob, length, opts) {
	var type = blob[blob.l++];
	var cparams = blob.read_shift(1), tab = opts && opts.biff <= 3 ? [(type == 0x58 ? -1 : 0), blob.read_shift(1)]: parsetab(blob);
	return [cparams, (tab[0] === 0 ? Ftab : Cetab)[tab[1]]];
}

function parsetab(blob) {
	return [blob[blob.l+1]>>7, blob.read_shift(2) & 0x7FFF];
}

/* [MS-XLS] 2.5.198.41 ; [MS-XLSB] 2.5.97.33 */
function parse_PtgAttrSum(blob, length, opts) {
	blob.l += opts && opts.biff == 2 ? 3 : 4; return;
}

/* [MS-XLS] 2.5.198.58 ; [MS-XLSB] 2.5.97.40 */
function parse_PtgExp(blob, length, opts) {
	blob.l++;
	if(opts && opts.biff == 12) return [blob.read_shift(4, 'i'), 0];
	var row = blob.read_shift(2);
	var col = blob.read_shift(opts && opts.biff == 2 ? 1 : 2);
	return [row, col];
}

/* [MS-XLS] 2.5.198.57 ; [MS-XLSB] 2.5.97.39 */
function parse_PtgErr(blob) { blob.l++; return BErr[blob.read_shift(1)]; }

/* [MS-XLS] 2.5.198.66 ; [MS-XLSB] 2.5.97.49 */
function parse_PtgInt(blob) { blob.l++; return blob.read_shift(2); }

/* [MS-XLS] 2.5.198.42 ; [MS-XLSB] 2.5.97.34 */
function parse_PtgBool(blob) { blob.l++; return blob.read_shift(1)!==0;}

/* [MS-XLS] 2.5.198.79 ; [MS-XLSB] 2.5.97.63 */
function parse_PtgNum(blob) { blob.l++; return parse_Xnum(blob, 8); }

/* [MS-XLS] 2.5.198.89 ; [MS-XLSB] 2.5.97.74 */
function parse_PtgStr(blob, length, opts) { blob.l++; return parse_ShortXLUnicodeString(blob, length-1, opts); }

/* [MS-XLS] 2.5.192.112 + 2.5.192.11{3,4,5,6,7} */
/* [MS-XLSB] 2.5.97.93 + 2.5.97.9{4,5,6,7} */
function parse_SerAr(blob, biff) {
	var val = [blob.read_shift(1)];
	if(biff == 12) switch(val[0]) {
		case 0x02: val[0] = 0x04; break; /* SerBool */
		case 0x04: val[0] = 0x10; break; /* SerErr */
		case 0x00: val[0] = 0x01; break; /* SerNum */
		case 0x01: val[0] = 0x02; break; /* SerStr */
	}
	switch(val[0]) {
		case 0x04: /* SerBool -- boolean */
			val[1] = parsebool(blob, 1) ? 'TRUE' : 'FALSE';
			if(biff != 12) blob.l += 7; break;
		case 0x25: /* appears to be an alias */
		case 0x10: /* SerErr -- error */
			val[1] = BErr[blob[blob.l]];
			blob.l += ((biff == 12) ? 4 : 8); break;
		case 0x00: /* SerNil -- honestly, I'm not sure how to reproduce this */
			blob.l += 8; break;
		case 0x01: /* SerNum -- Xnum */
			val[1] = parse_Xnum(blob, 8); break;
		case 0x02: /* SerStr -- XLUnicodeString (<256 chars) */
			val[1] = parse_XLUnicodeString2(blob, 0, {biff:biff > 0 && biff < 8 ? 2 : biff}); break;
		default: throw new Error("Bad SerAr: " + val[0]); /* Unreachable */
	}
	return val;
}

/* [MS-XLS] 2.5.198.61 ; [MS-XLSB] 2.5.97.44 */
function parse_PtgExtraMem(blob, cce, opts) {
	var count = blob.read_shift((opts.biff == 12) ? 4 : 2);
	var out = [];
	for(var i = 0; i != count; ++i) out.push(((opts.biff == 12) ? parse_UncheckedRfX : parse_Ref8U)(blob, 8));
	return out;
}

/* [MS-XLS] 2.5.198.59 ; [MS-XLSB] 2.5.97.41 */
function parse_PtgExtraArray(blob, length, opts) {
	var rows = 0, cols = 0;
	if(opts.biff == 12) {
		rows = blob.read_shift(4); // DRw
		cols = blob.read_shift(4); // DCol
	} else {
		cols = 1 + blob.read_shift(1); //DColByteU
		rows = 1 + blob.read_shift(2); //DRw
	}
	if(opts.biff >= 2 && opts.biff < 8) { --rows; if(--cols == 0) cols = 0x100; }
	// $FlowIgnore
	for(var i = 0, o = []; i != rows && (o[i] = []); ++i)
		for(var j = 0; j != cols; ++j) o[i][j] = parse_SerAr(blob, opts.biff);
	return o;
}

/* [MS-XLS] 2.5.198.76 ; [MS-XLSB] 2.5.97.60 */
function parse_PtgName(blob, length, opts) {
	var type = (blob.read_shift(1) >>> 5) & 0x03;
	var w = (!opts || (opts.biff >= 8)) ? 4 : 2;
	var nameindex = blob.read_shift(w);
	switch(opts.biff) {
		case 2: blob.l += 5; break;
		case 3: case 4: blob.l += 8; break;
		case 5: blob.l += 12; break;
	}
	return [type, 0, nameindex];
}

/* [MS-XLS] 2.5.198.77 ; [MS-XLSB] 2.5.97.61 */
function parse_PtgNameX(blob, length, opts) {
	if(opts.biff == 5) return parse_PtgNameX_BIFF5(blob, length, opts);
	var type = (blob.read_shift(1) >>> 5) & 0x03;
	var ixti = blob.read_shift(2); // XtiIndex
	var nameindex = blob.read_shift(4);
	return [type, ixti, nameindex];
}
function parse_PtgNameX_BIFF5(blob) {
	var type = (blob.read_shift(1) >>> 5) & 0x03;
	var ixti = blob.read_shift(2, 'i'); // XtiIndex
	blob.l += 8;
	var nameindex = blob.read_shift(2);
	blob.l += 12;
	return [type, ixti, nameindex];
}

/* [MS-XLS] 2.5.198.70 ; [MS-XLSB] 2.5.97.54 */
function parse_PtgMemArea(blob, length, opts) {
	var type = (blob.read_shift(1) >>> 5) & 0x03;
	blob.l += (opts && opts.biff == 2 ? 3 : 4);
	var cce = blob.read_shift(opts && opts.biff == 2 ? 1 : 2);
	return [type, cce];
}

/* [MS-XLS] 2.5.198.72 ; [MS-XLSB] 2.5.97.56 */
function parse_PtgMemFunc(blob, length, opts) {
	var type = (blob.read_shift(1) >>> 5) & 0x03;
	var cce = blob.read_shift(opts && opts.biff == 2 ? 1 : 2);
	return [type, cce];
}


/* [MS-XLS] 2.5.198.86 ; [MS-XLSB] 2.5.97.69 */
function parse_PtgRefErr(blob, length, opts) {
	var type = (blob.read_shift(1) >>> 5) & 0x03;
	blob.l += 4;
	if(opts.biff < 8) blob.l--;
	if(opts.biff == 12) blob.l += 2;
	return [type];
}

/* [MS-XLS] 2.5.198.87 ; [MS-XLSB] 2.5.97.71 */
function parse_PtgRefErr3d(blob, length, opts) {
	var type = (blob[blob.l++] & 0x60) >> 5;
	var ixti = blob.read_shift(2);
	var w = 4;
	if(opts) switch(opts.biff) {
		case 5: w = 15; break;
		case 12: w = 6; break;
	}
	blob.l += w;
	return [type, ixti];
}

/* [MS-XLS] 2.5.198.71 ; [MS-XLSB] 2.5.97.55 */
var parse_PtgMemErr = parsenoop;
/* [MS-XLS] 2.5.198.73  ; [MS-XLSB] 2.5.97.57 */
var parse_PtgMemNoMem = parsenoop;
/* [MS-XLS] 2.5.198.92 */
var parse_PtgTbl = parsenoop;

function parse_PtgElfLoc(blob, length, opts) {
	blob.l += 2;
	return [parse_RgceElfLoc(blob, 4, opts)];
}
function parse_PtgElfNoop(blob) {
	blob.l += 6;
	return [];
}
/* [MS-XLS] 2.5.198.46 */
var parse_PtgElfCol = parse_PtgElfLoc;
/* [MS-XLS] 2.5.198.47 */
var parse_PtgElfColS = parse_PtgElfNoop;
/* [MS-XLS] 2.5.198.48 */
var parse_PtgElfColSV = parse_PtgElfNoop;
/* [MS-XLS] 2.5.198.49 */
var parse_PtgElfColV = parse_PtgElfLoc;
/* [MS-XLS] 2.5.198.50 */
function parse_PtgElfLel(blob) {
	blob.l += 2;
	return [parseuint16(blob), blob.read_shift(2) & 0x01];
}
/* [MS-XLS] 2.5.198.51 */
var parse_PtgElfRadical = parse_PtgElfLoc;
/* [MS-XLS] 2.5.198.52 */
var parse_PtgElfRadicalLel = parse_PtgElfLel;
/* [MS-XLS] 2.5.198.53 */
var parse_PtgElfRadicalS = parse_PtgElfNoop;
/* [MS-XLS] 2.5.198.54 */
var parse_PtgElfRw = parse_PtgElfLoc;
/* [MS-XLS] 2.5.198.55 */
var parse_PtgElfRwV = parse_PtgElfLoc;

/* [MS-XLSB] 2.5.97.52 TODO */
var PtgListRT = [
	"Data",
	"All",
	"Headers",
	"??",
	"?Data2",
	"??",
	"?DataHeaders",
	"??",
	"Totals",
	"??",
	"??",
	"??",
	"?DataTotals",
	"??",
	"??",
	"??",
	"?Current"
];
function parse_PtgList(blob) {
	blob.l += 2;
	var ixti = blob.read_shift(2);
	var flags = blob.read_shift(2);
	var idx = blob.read_shift(4);
	var c = blob.read_shift(2);
	var C = blob.read_shift(2);
	var rt = PtgListRT[(flags >> 2) & 0x1F];
	return {ixti: ixti, coltype:(flags&0x3), rt:rt, idx:idx, c:c, C:C};
}
/* [MS-XLS] 2.5.198.91 ; [MS-XLSB] 2.5.97.76 */
function parse_PtgSxName(blob) {
	blob.l += 2;
	return [blob.read_shift(4)];
}

/* [XLS] old spec */
function parse_PtgSheet(blob, length, opts) {
	blob.l += 5;
	blob.l += 2;
	blob.l += (opts.biff == 2 ? 1 : 4);
	return ["PTGSHEET"];
}
function parse_PtgEndSheet(blob, length, opts) {
	blob.l += (opts.biff == 2 ? 4 : 5);
	return ["PTGENDSHEET"];
}
function parse_PtgMemAreaN(blob) {
	var type = (blob.read_shift(1) >>> 5) & 0x03;
	var cce = blob.read_shift(2);
	return [type, cce];
}
function parse_PtgMemNoMemN(blob) {
	var type = (blob.read_shift(1) >>> 5) & 0x03;
	var cce = blob.read_shift(2);
	return [type, cce];
}
function parse_PtgAttrNoop(blob) {
	blob.l += 4;
	return [0, 0];
}

/* [MS-XLS] 2.5.198.25 ; [MS-XLSB] 2.5.97.16 */
var PtgTypes = {
0x01: { n:'PtgExp', f:parse_PtgExp },
0x02: { n:'PtgTbl', f:parse_PtgTbl },
0x03: { n:'PtgAdd', f:parseread1 },
0x04: { n:'PtgSub', f:parseread1 },
0x05: { n:'PtgMul', f:parseread1 },
0x06: { n:'PtgDiv', f:parseread1 },
0x07: { n:'PtgPower', f:parseread1 },
0x08: { n:'PtgConcat', f:parseread1 },
0x09: { n:'PtgLt', f:parseread1 },
0x0A: { n:'PtgLe', f:parseread1 },
0x0B: { n:'PtgEq', f:parseread1 },
0x0C: { n:'PtgGe', f:parseread1 },
0x0D: { n:'PtgGt', f:parseread1 },
0x0E: { n:'PtgNe', f:parseread1 },
0x0F: { n:'PtgIsect', f:parseread1 },
0x10: { n:'PtgUnion', f:parseread1 },
0x11: { n:'PtgRange', f:parseread1 },
0x12: { n:'PtgUplus', f:parseread1 },
0x13: { n:'PtgUminus', f:parseread1 },
0x14: { n:'PtgPercent', f:parseread1 },
0x15: { n:'PtgParen', f:parseread1 },
0x16: { n:'PtgMissArg', f:parseread1 },
0x17: { n:'PtgStr', f:parse_PtgStr },
0x1A: { n:'PtgSheet', f:parse_PtgSheet },
0x1B: { n:'PtgEndSheet', f:parse_PtgEndSheet },
0x1C: { n:'PtgErr', f:parse_PtgErr },
0x1D: { n:'PtgBool', f:parse_PtgBool },
0x1E: { n:'PtgInt', f:parse_PtgInt },
0x1F: { n:'PtgNum', f:parse_PtgNum },
0x20: { n:'PtgArray', f:parse_PtgArray },
0x21: { n:'PtgFunc', f:parse_PtgFunc },
0x22: { n:'PtgFuncVar', f:parse_PtgFuncVar },
0x23: { n:'PtgName', f:parse_PtgName },
0x24: { n:'PtgRef', f:parse_PtgRef },
0x25: { n:'PtgArea', f:parse_PtgArea },
0x26: { n:'PtgMemArea', f:parse_PtgMemArea },
0x27: { n:'PtgMemErr', f:parse_PtgMemErr },
0x28: { n:'PtgMemNoMem', f:parse_PtgMemNoMem },
0x29: { n:'PtgMemFunc', f:parse_PtgMemFunc },
0x2A: { n:'PtgRefErr', f:parse_PtgRefErr },
0x2B: { n:'PtgAreaErr', f:parse_PtgAreaErr },
0x2C: { n:'PtgRefN', f:parse_PtgRefN },
0x2D: { n:'PtgAreaN', f:parse_PtgAreaN },
0x2E: { n:'PtgMemAreaN', f:parse_PtgMemAreaN },
0x2F: { n:'PtgMemNoMemN', f:parse_PtgMemNoMemN },
0x39: { n:'PtgNameX', f:parse_PtgNameX },
0x3A: { n:'PtgRef3d', f:parse_PtgRef3d },
0x3B: { n:'PtgArea3d', f:parse_PtgArea3d },
0x3C: { n:'PtgRefErr3d', f:parse_PtgRefErr3d },
0x3D: { n:'PtgAreaErr3d', f:parse_PtgAreaErr3d },
0xFF: {}
};
/* These are duplicated in the PtgTypes table */
var PtgDupes = {
0x40: 0x20, 0x60: 0x20,
0x41: 0x21, 0x61: 0x21,
0x42: 0x22, 0x62: 0x22,
0x43: 0x23, 0x63: 0x23,
0x44: 0x24, 0x64: 0x24,
0x45: 0x25, 0x65: 0x25,
0x46: 0x26, 0x66: 0x26,
0x47: 0x27, 0x67: 0x27,
0x48: 0x28, 0x68: 0x28,
0x49: 0x29, 0x69: 0x29,
0x4A: 0x2A, 0x6A: 0x2A,
0x4B: 0x2B, 0x6B: 0x2B,
0x4C: 0x2C, 0x6C: 0x2C,
0x4D: 0x2D, 0x6D: 0x2D,
0x4E: 0x2E, 0x6E: 0x2E,
0x4F: 0x2F, 0x6F: 0x2F,
0x58: 0x22, 0x78: 0x22,
0x59: 0x39, 0x79: 0x39,
0x5A: 0x3A, 0x7A: 0x3A,
0x5B: 0x3B, 0x7B: 0x3B,
0x5C: 0x3C, 0x7C: 0x3C,
0x5D: 0x3D, 0x7D: 0x3D
};
(function(){for(var y in PtgDupes) PtgTypes[y] = PtgTypes[PtgDupes[y]];})();

var Ptg18 = {
0x01: { n:'PtgElfLel', f:parse_PtgElfLel },
0x02: { n:'PtgElfRw', f:parse_PtgElfRw },
0x03: { n:'PtgElfCol', f:parse_PtgElfCol },
0x06: { n:'PtgElfRwV', f:parse_PtgElfRwV },
0x07: { n:'PtgElfColV', f:parse_PtgElfColV },
0x0A: { n:'PtgElfRadical', f:parse_PtgElfRadical },
0x0B: { n:'PtgElfRadicalS', f:parse_PtgElfRadicalS },
0x0D: { n:'PtgElfColS', f:parse_PtgElfColS },
0x0F: { n:'PtgElfColSV', f:parse_PtgElfColSV },
0x10: { n:'PtgElfRadicalLel', f:parse_PtgElfRadicalLel },
0x19: { n:'PtgList', f:parse_PtgList },
0x1D: { n:'PtgSxName', f:parse_PtgSxName },
0xFF: {}
};
var Ptg19 = {
0x00: { n:'PtgAttrNoop', f:parse_PtgAttrNoop },
0x01: { n:'PtgAttrSemi', f:parse_PtgAttrSemi },
0x02: { n:'PtgAttrIf', f:parse_PtgAttrIf },
0x04: { n:'PtgAttrChoose', f:parse_PtgAttrChoose },
0x08: { n:'PtgAttrGoto', f:parse_PtgAttrGoto },
0x10: { n:'PtgAttrSum', f:parse_PtgAttrSum },
0x20: { n:'PtgAttrBaxcel', f:parse_PtgAttrBaxcel },
0x40: { n:'PtgAttrSpace', f:parse_PtgAttrSpace },
0x41: { n:'PtgAttrSpaceSemi', f:parse_PtgAttrSpaceSemi },
0x80: { n:'PtgAttrIfError', f:parse_PtgAttrIfError },
0xFF: {}
};
Ptg19[0x21] = Ptg19[0x20];

/* [MS-XLS] 2.5.198.103 ; [MS-XLSB] 2.5.97.87 */
function parse_RgbExtra(blob, length, rgce, opts) {
	if(opts.biff < 8) return parsenoop(blob, length);
	var target = blob.l + length;
	var o = [];
	for(var i = 0; i !== rgce.length; ++i) {
		switch(rgce[i][0]) {
			case 'PtgArray': /* PtgArray -> PtgExtraArray */
				rgce[i][1] = parse_PtgExtraArray(blob, 0, opts);
				o.push(rgce[i][1]);
				break;
			case 'PtgMemArea': /* PtgMemArea -> PtgExtraMem */
				rgce[i][2] = parse_PtgExtraMem(blob, rgce[i][1], opts);
				o.push(rgce[i][2]);
				break;
			case 'PtgExp': /* PtgExp -> PtgExtraCol */
				if(opts && opts.biff == 12) {
					rgce[i][1][1] = blob.read_shift(4);
					o.push(rgce[i][1]);
				} break;
			case 'PtgList': /* TODO: PtgList -> PtgExtraList */
			case 'PtgElfRadicalS': /* TODO: PtgElfRadicalS -> PtgExtraElf */
			case 'PtgElfColS': /* TODO: PtgElfColS -> PtgExtraElf */
			case 'PtgElfColSV': /* TODO: PtgElfColSV -> PtgExtraElf */
				throw "Unsupported " + rgce[i][0];
			default: break;
		}
	}
	length = target - blob.l;
	/* note: this is technically an error but Excel disregards */
	//if(target !== blob.l && blob.l !== target - length) throw new Error(target + " != " + blob.l);
	if(length !== 0) o.push(parsenoop(blob, length));
	return o;
}

/* [MS-XLS] 2.5.198.104 ; [MS-XLSB] 2.5.97.88 */
function parse_Rgce(blob, length, opts) {
	var target = blob.l + length;
	var R, id, ptgs = [];
	while(target != blob.l) {
		length = target - blob.l;
		id = blob[blob.l];
		R = PtgTypes[id];
		if(id === 0x18 || id === 0x19) R = (id === 0x18 ? Ptg18 : Ptg19)[blob[blob.l + 1]];
		if(!R || !R.f) { /*ptgs.push*/(parsenoop(blob, length)); }
		else { ptgs.push([R.n, R.f(blob, length, opts)]); }
	}
	return ptgs;
}

function stringify_array(f) {
	var o = [];
	for(var i = 0; i < f.length; ++i) {
		var x = f[i], r = [];
		for(var j = 0; j < x.length; ++j) {
			var y = x[j];
			if(y) switch(y[0]) {
				// TODO: handle embedded quotes
				case 0x02:
r.push('"' + y[1].replace(/"/g,'""') + '"'); break;
				default: r.push(y[1]);
			} else r.push("");
		}
		o.push(r.join(","));
	}
	return o.join(";");
}

/* [MS-XLS] 2.2.2 ; [MS-XLSB] 2.2.2 TODO */
var PtgBinOp = {
	PtgAdd: "+",
	PtgConcat: "&",
	PtgDiv: "/",
	PtgEq: "=",
	PtgGe: ">=",
	PtgGt: ">",
	PtgLe: "<=",
	PtgLt: "<",
	PtgMul: "*",
	PtgNe: "<>",
	PtgPower: "^",
	PtgSub: "-"
};
function formula_quote_sheet_name(sname, opts) {
	if(!sname && !(opts && opts.biff <= 5 && opts.biff >= 2)) throw new Error("empty sheet name");
	if(sname.indexOf(" ") > -1) return "'" + sname + "'";
	return sname;
}
function get_ixti_raw(supbooks, ixti, opts) {
	if(!supbooks) return "SH33TJSERR0";
	if(!supbooks.XTI) return "SH33TJSERR6";
	var XTI = supbooks.XTI[ixti];
	if(opts.biff > 8 && !supbooks.XTI[ixti]) return supbooks.SheetNames[ixti];
	if(opts.biff < 8) {
		if(ixti > 10000) ixti-= 65536;
		if(ixti < 0) ixti = -ixti;
		return ixti == 0 ? "" : supbooks.XTI[ixti - 1];
	}
	if(!XTI) return "SH33TJSERR1";
	var o = "";
	if(opts.biff > 8) switch(supbooks[XTI[0]][0]) {
		case 0x0165: /* 'BrtSupSelf' */
			o = XTI[1] == -1 ? "#REF" : supbooks.SheetNames[XTI[1]];
			return XTI[1] == XTI[2] ? o : o + ":" + supbooks.SheetNames[XTI[2]];
		case 0x0166: /* 'BrtSupSame' */
			if(opts.SID != null) return supbooks.SheetNames[opts.SID];
			return "SH33TJSSAME" + supbooks[XTI[0]][0];
		case 0x0163: /* 'BrtSupBookSrc' */
			/* falls through */
		default: return "SH33TJSSRC" + supbooks[XTI[0]][0];
	}
	switch(supbooks[XTI[0]][0][0]) {
		case 0x0401:
			o = XTI[1] == -1 ? "#REF" : (supbooks.SheetNames[XTI[1]] || "SH33TJSERR3");
			return XTI[1] == XTI[2] ? o : o + ":" + supbooks.SheetNames[XTI[2]];
		case 0x3A01: return "SH33TJSERR8";
		default:
			if(!supbooks[XTI[0]][0][3]) return "SH33TJSERR2";
			o = XTI[1] == -1 ? "#REF" : (supbooks[XTI[0]][0][3][XTI[1]] || "SH33TJSERR4");
			return XTI[1] == XTI[2] ? o : o + ":" + supbooks[XTI[0]][0][3][XTI[2]];
	}
}
function get_ixti(supbooks, ixti, opts) {
	return formula_quote_sheet_name(get_ixti_raw(supbooks, ixti, opts), opts);
}
function stringify_formula(formula/*Array<any>*/, range, cell, supbooks, opts) {
	var biff = (opts && opts.biff) || 8;
	var _range = /*range != null ? range :*/ {s:{c:0, r:0},e:{c:0, r:0}};
	var stack = [], e1, e2,  c, ixti=0, nameidx=0, r, sname="";
	if(!formula[0] || !formula[0][0]) return "";
	var last_sp = -1, sp = "";
	for(var ff = 0, fflen = formula[0].length; ff < fflen; ++ff) {
		var f = formula[0][ff];
		switch(f[0]) {
			case 'PtgUminus': /* [MS-XLS] 2.5.198.93 */
				stack.push("-" + stack.pop()); break;
			case 'PtgUplus': /* [MS-XLS] 2.5.198.95 */
				stack.push("+" + stack.pop()); break;
			case 'PtgPercent': /* [MS-XLS] 2.5.198.81 */
				stack.push(stack.pop() + "%"); break;

			case 'PtgAdd':    /* [MS-XLS] 2.5.198.26 */
			case 'PtgConcat': /* [MS-XLS] 2.5.198.43 */
			case 'PtgDiv':    /* [MS-XLS] 2.5.198.45 */
			case 'PtgEq':     /* [MS-XLS] 2.5.198.56 */
			case 'PtgGe':     /* [MS-XLS] 2.5.198.64 */
			case 'PtgGt':     /* [MS-XLS] 2.5.198.65 */
			case 'PtgLe':     /* [MS-XLS] 2.5.198.68 */
			case 'PtgLt':     /* [MS-XLS] 2.5.198.69 */
			case 'PtgMul':    /* [MS-XLS] 2.5.198.75 */
			case 'PtgNe':     /* [MS-XLS] 2.5.198.78 */
			case 'PtgPower':  /* [MS-XLS] 2.5.198.82 */
			case 'PtgSub':    /* [MS-XLS] 2.5.198.90 */
				e1 = stack.pop(); e2 = stack.pop();
				if(last_sp >= 0) {
					switch(formula[0][last_sp][1][0]) {
						case 0:
							// $FlowIgnore
							sp = fill(" ", formula[0][last_sp][1][1]); break;
						case 1:
							// $FlowIgnore
							sp = fill("\r", formula[0][last_sp][1][1]); break;
						default:
							sp = "";
							// $FlowIgnore
							if(opts.WTF) throw new Error("Unexpected PtgAttrSpaceType " + formula[0][last_sp][1][0]);
					}
					e2 = e2 + sp;
					last_sp = -1;
				}
				stack.push(e2+PtgBinOp[f[0]]+e1);
				break;

			case 'PtgIsect': /* [MS-XLS] 2.5.198.67 */
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+" "+e1);
				break;
			case 'PtgUnion': /* [MS-XLS] 2.5.198.94 */
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+","+e1);
				break;
			case 'PtgRange': /* [MS-XLS] 2.5.198.83 */
				e1 = stack.pop(); e2 = stack.pop();
				stack.push(e2+":"+e1);
				break;

			case 'PtgAttrChoose': /* [MS-XLS] 2.5.198.34 */
				break;
			case 'PtgAttrGoto': /* [MS-XLS] 2.5.198.35 */
				break;
			case 'PtgAttrIf': /* [MS-XLS] 2.5.198.36 */
				break;
			case 'PtgAttrIfError': /* [MS-XLSB] 2.5.97.28 */
				break;


			case 'PtgRef': /* [MS-XLS] 2.5.198.84 */
c = shift_cell_xls((f[1][1]), _range, opts);
				stack.push(encode_cell_xls(c, biff));
				break;
			case 'PtgRefN': /* [MS-XLS] 2.5.198.88 */
c = cell ? shift_cell_xls((f[1][1]), cell, opts) : (f[1][1]);
				stack.push(encode_cell_xls(c, biff));
				break;
			case 'PtgRef3d': /* [MS-XLS] 2.5.198.85 */
ixti = f[1][1]; c = shift_cell_xls((f[1][2]), _range, opts);
				sname = get_ixti(supbooks, ixti, opts);
				var w = sname; /* IE9 fails on defined names */ // eslint-disable-line no-unused-vars
				stack.push(sname + "!" + encode_cell_xls(c, biff));
				break;

			case 'PtgFunc': /* [MS-XLS] 2.5.198.62 */
			case 'PtgFuncVar': /* [MS-XLS] 2.5.198.63 */
				/* f[1] = [argc, func, type] */
				var argc = (f[1][0]), func = (f[1][1]);
				if(!argc) argc = 0;
				argc &= 0x7F;
				var args = argc == 0 ? [] : stack.slice(-argc);
				stack.length -= argc;
				if(func === 'User') func = args.shift();
				stack.push(func + "(" + args.join(",") + ")");
				break;

			case 'PtgBool': /* [MS-XLS] 2.5.198.42 */
				stack.push(f[1] ? "TRUE" : "FALSE"); break;
			case 'PtgInt': /* [MS-XLS] 2.5.198.66 */
				stack.push(f[1]); break;
			case 'PtgNum': /* [MS-XLS] 2.5.198.79 TODO: precision? */
				stack.push(String(f[1])); break;
			case 'PtgStr': /* [MS-XLS] 2.5.198.89 */
				// $FlowIgnore
				stack.push('"' + f[1] + '"'); break;
			case 'PtgErr': /* [MS-XLS] 2.5.198.57 */
				stack.push(f[1]); break;
			case 'PtgAreaN': /* [MS-XLS] 2.5.198.31 TODO */
r = shift_range_xls(f[1][1], cell ? {s:cell} : _range, opts);
				stack.push(encode_range_xls((r), opts));
				break;
			case 'PtgArea': /* [MS-XLS] 2.5.198.27 TODO: fixed points */
r = shift_range_xls(f[1][1], _range, opts);
				stack.push(encode_range_xls((r), opts));
				break;
			case 'PtgArea3d': /* [MS-XLS] 2.5.198.28 TODO */
ixti = f[1][1]; r = f[1][2];
				sname = get_ixti(supbooks, ixti, opts);
				stack.push(sname + "!" + encode_range_xls((r), opts));
				break;
			case 'PtgAttrSum': /* [MS-XLS] 2.5.198.41 */
				stack.push("SUM(" + stack.pop() + ")");
				break;

			case 'PtgAttrBaxcel': /* [MS-XLS] 2.5.198.33 */
			case 'PtgAttrSemi': /* [MS-XLS] 2.5.198.37 */
				break;

			case 'PtgName': /* [MS-XLS] 2.5.198.76 ; [MS-XLSB] 2.5.97.60 TODO: revisions */
				/* f[1] = type, 0, nameindex */
				nameidx = (f[1][2]);
				var lbl = (supbooks.names||[])[nameidx-1] || (supbooks[0]||[])[nameidx];
				var name = lbl ? lbl.Name : "SH33TJSNAME" + String(nameidx);
				if(name in XLSXFutureFunctions) name = XLSXFutureFunctions[name];
				stack.push(name);
				break;

			case 'PtgNameX': /* [MS-XLS] 2.5.198.77 ; [MS-XLSB] 2.5.97.61 TODO: revisions */
				/* f[1] = type, ixti, nameindex */
				var bookidx = (f[1][1]); nameidx = (f[1][2]); var externbook;
				/* TODO: Properly handle missing values */
				if(opts.biff <= 5) {
					if(bookidx < 0) bookidx = -bookidx;
					if(supbooks[bookidx]) externbook = supbooks[bookidx][nameidx];
				} else {
					var o = "";
					if(((supbooks[bookidx]||[])[0]||[])[0] == 0x3A01){/* empty */}
					else if(((supbooks[bookidx]||[])[0]||[])[0] == 0x0401){
						if(supbooks[bookidx][nameidx] && supbooks[bookidx][nameidx].itab > 0) {
							o = supbooks.SheetNames[supbooks[bookidx][nameidx].itab-1] + "!";
						}
					}
					else o = supbooks.SheetNames[nameidx-1]+ "!";
					if(supbooks[bookidx] && supbooks[bookidx][nameidx]) o += supbooks[bookidx][nameidx].Name;
					else if(supbooks[0] && supbooks[0][nameidx]) o += supbooks[0][nameidx].Name;
					else o += "SH33TJSERRX";
					stack.push(o);
					break;
				}
				if(!externbook) externbook = {Name: "SH33TJSERRY"};
				stack.push(externbook.Name);
				break;

			case 'PtgParen': /* [MS-XLS] 2.5.198.80 */
				var lp = '(', rp = ')';
				if(last_sp >= 0) {
					sp = "";
					switch(formula[0][last_sp][1][0]) {
						// $FlowIgnore
						case 2: lp = fill(" ", formula[0][last_sp][1][1]) + lp; break;
						// $FlowIgnore
						case 3: lp = fill("\r", formula[0][last_sp][1][1]) + lp; break;
						// $FlowIgnore
						case 4: rp = fill(" ", formula[0][last_sp][1][1]) + rp; break;
						// $FlowIgnore
						case 5: rp = fill("\r", formula[0][last_sp][1][1]) + rp; break;
						default:
							// $FlowIgnore
							if(opts.WTF) throw new Error("Unexpected PtgAttrSpaceType " + formula[0][last_sp][1][0]);
					}
					last_sp = -1;
				}
				stack.push(lp + stack.pop() + rp); break;

			case 'PtgRefErr': /* [MS-XLS] 2.5.198.86 */
				stack.push('#REF!'); break;

			case 'PtgRefErr3d': /* [MS-XLS] 2.5.198.87 */
				stack.push('#REF!'); break;

			case 'PtgExp': /* [MS-XLS] 2.5.198.58 TODO */
				c = {c:(f[1][1]),r:(f[1][0])};
				var q = ({c: cell.c, r:cell.r});
				if(supbooks.sharedf[encode_cell(c)]) {
					var parsedf = (supbooks.sharedf[encode_cell(c)]);
					stack.push(stringify_formula(parsedf, _range, q, supbooks, opts));
				}
				else {
					var fnd = false;
					for(e1=0;e1!=supbooks.arrayf.length; ++e1) {
						/* TODO: should be something like range_has */
						e2 = supbooks.arrayf[e1];
						if(c.c < e2[0].s.c || c.c > e2[0].e.c) continue;
						if(c.r < e2[0].s.r || c.r > e2[0].e.r) continue;
						stack.push(stringify_formula(e2[1], _range, q, supbooks, opts));
						fnd = true;
						break;
					}
					if(!fnd) stack.push(f[1]);
				}
				break;

			case 'PtgArray': /* [MS-XLS] 2.5.198.32 TODO */
				stack.push("{" + stringify_array(f[1]) + "}");
				break;

			case 'PtgMemArea': /* [MS-XLS] 2.5.198.70 TODO: confirm this is a non-display */
				//stack.push("(" + f[2].map(encode_range).join(",") + ")");
				break;

			case 'PtgAttrSpace': /* [MS-XLS] 2.5.198.38 */
			case 'PtgAttrSpaceSemi': /* [MS-XLS] 2.5.198.39 */
				last_sp = ff;
				break;

			case 'PtgTbl': /* [MS-XLS] 2.5.198.92 TODO */
				break;

			case 'PtgMemErr': /* [MS-XLS] 2.5.198.71 */
				break;

			case 'PtgMissArg': /* [MS-XLS] 2.5.198.74 */
				stack.push("");
				break;

			case 'PtgAreaErr': /* [MS-XLS] 2.5.198.29 */
				stack.push("#REF!"); break;

			case 'PtgAreaErr3d': /* [MS-XLS] 2.5.198.30 */
				stack.push("#REF!"); break;

			case 'PtgList': /* [MS-XLSB] 2.5.97.52 */
				// $FlowIgnore
				stack.push("Table" + f[1].idx + "[#" + f[1].rt + "]");
				break;

			case 'PtgMemAreaN':
			case 'PtgMemNoMemN':
			case 'PtgAttrNoop':
			case 'PtgSheet':
			case 'PtgEndSheet':
				break;

			case 'PtgMemFunc': /* [MS-XLS] 2.5.198.72 TODO */
				break;
			case 'PtgMemNoMem': /* [MS-XLS] 2.5.198.73 TODO */
				break;

			case 'PtgElfCol': /* [MS-XLS] 2.5.198.46 */
			case 'PtgElfColS': /* [MS-XLS] 2.5.198.47 */
			case 'PtgElfColSV': /* [MS-XLS] 2.5.198.48 */
			case 'PtgElfColV': /* [MS-XLS] 2.5.198.49 */
			case 'PtgElfLel': /* [MS-XLS] 2.5.198.50 */
			case 'PtgElfRadical': /* [MS-XLS] 2.5.198.51 */
			case 'PtgElfRadicalLel': /* [MS-XLS] 2.5.198.52 */
			case 'PtgElfRadicalS': /* [MS-XLS] 2.5.198.53 */
			case 'PtgElfRw': /* [MS-XLS] 2.5.198.54 */
			case 'PtgElfRwV': /* [MS-XLS] 2.5.198.55 */
				throw new Error("Unsupported ELFs");

			case 'PtgSxName': /* [MS-XLS] 2.5.198.91 TODO -- find a test case */
				throw new Error('Unrecognized Formula Token: ' + String(f));
			default: throw new Error('Unrecognized Formula Token: ' + String(f));
		}
		var PtgNonDisp = ['PtgAttrSpace', 'PtgAttrSpaceSemi', 'PtgAttrGoto'];
		if(opts.biff != 3) if(last_sp >= 0 && PtgNonDisp.indexOf(formula[0][ff][0]) == -1) {
			f = formula[0][last_sp];
			var _left = true;
			switch(f[1][0]) {
				/* note: some bad XLSB files omit the PtgParen */
				case 4: _left = false;
				/* falls through */
				case 0:
					// $FlowIgnore
					sp = fill(" ", f[1][1]); break;
				case 5: _left = false;
				/* falls through */
				case 1:
					// $FlowIgnore
					sp = fill("\r", f[1][1]); break;
				default:
					sp = "";
					// $FlowIgnore
					if(opts.WTF) throw new Error("Unexpected PtgAttrSpaceType " + f[1][0]);
			}
			stack.push((_left ? sp : "") + stack.pop() + (_left ? "" : sp));
			last_sp = -1;
		}
	}
	if(stack.length > 1 && opts.WTF) throw new Error("bad formula stack");
	return stack[0];
}

/* [MS-XLS] 2.5.198.1 TODO */
function parse_ArrayParsedFormula(blob, length, opts) {
	var target = blob.l + length, len = opts.biff == 2 ? 1 : 2;
	var rgcb, cce = blob.read_shift(len); // length of rgce
	if(cce == 0xFFFF) return [[],parsenoop(blob, length-2)];
	var rgce = parse_Rgce(blob, cce, opts);
	if(length !== cce + len) rgcb = parse_RgbExtra(blob, length - cce - len, rgce, opts);
	blob.l = target;
	return [rgce, rgcb];
}

/* [MS-XLS] 2.5.198.3 TODO */
function parse_XLSCellParsedFormula(blob, length, opts) {
	var target = blob.l + length, len = opts.biff == 2 ? 1 : 2;
	var rgcb, cce = blob.read_shift(len); // length of rgce
	if(cce == 0xFFFF) return [[],parsenoop(blob, length-2)];
	var rgce = parse_Rgce(blob, cce, opts);
	if(length !== cce + len) rgcb = parse_RgbExtra(blob, length - cce - len, rgce, opts);
	blob.l = target;
	return [rgce, rgcb];
}

/* [MS-XLS] 2.5.198.21 */
function parse_NameParsedFormula(blob, length, opts, cce) {
	var target = blob.l + length;
	var rgce = parse_Rgce(blob, cce, opts);
	var rgcb;
	if(target !== blob.l) rgcb = parse_RgbExtra(blob, target - blob.l, rgce, opts);
	return [rgce, rgcb];
}

/* [MS-XLS] 2.5.198.118 TODO */
function parse_SharedParsedFormula(blob, length, opts) {
	var target = blob.l + length;
	var rgcb, cce = blob.read_shift(2); // length of rgce
	var rgce = parse_Rgce(blob, cce, opts);
	if(cce == 0xFFFF) return [[],parsenoop(blob, length-2)];
	if(length !== cce + 2) rgcb = parse_RgbExtra(blob, target - cce - 2, rgce, opts);
	return [rgce, rgcb];
}

/* [MS-XLS] 2.5.133 TODO: how to emit empty strings? */
function parse_FormulaValue(blob) {
	var b;
	if(__readUInt16LE(blob,blob.l + 6) !== 0xFFFF) return [parse_Xnum(blob),'n'];
	switch(blob[blob.l]) {
		case 0x00: blob.l += 8; return ["String", 's'];
		case 0x01: b = blob[blob.l+2] === 0x1; blob.l += 8; return [b,'b'];
		case 0x02: b = blob[blob.l+2]; blob.l += 8; return [b,'e'];
		case 0x03: blob.l += 8; return ["",'s'];
	}
	return [];
}

/* [MS-XLS] 2.4.127 TODO */
function parse_Formula(blob, length, opts) {
	var end = blob.l + length;
	var cell = parse_XLSCell(blob, 6);
	if(opts.biff == 2) ++blob.l;
	var val = parse_FormulaValue(blob,8);
	var flags = blob.read_shift(1);
	if(opts.biff != 2) {
		blob.read_shift(1);
		if(opts.biff >= 5) {
			/*var chn = */blob.read_shift(4);
		}
	}
	var cbf = parse_XLSCellParsedFormula(blob, end - blob.l, opts);
	return {cell:cell, val:val[0], formula:cbf, shared: (flags >> 3) & 1, tt:val[1]};
}

/* XLSB Parsed Formula records have the same shape */
function parse_XLSBParsedFormula(data, length, opts) {
	var cce = data.read_shift(4);
	var rgce = parse_Rgce(data, cce, opts);
	var cb = data.read_shift(4);
	var rgcb = cb > 0 ? parse_RgbExtra(data, cb, rgce, opts) : null;
	return [rgce, rgcb];
}

/* [MS-XLSB] 2.5.97.1 ArrayParsedFormula */
var parse_XLSBArrayParsedFormula = parse_XLSBParsedFormula;
/* [MS-XLSB] 2.5.97.4 CellParsedFormula */
var parse_XLSBCellParsedFormula = parse_XLSBParsedFormula;
/* [MS-XLSB] 2.5.97.12 NameParsedFormula */
var parse_XLSBNameParsedFormula = parse_XLSBParsedFormula;
/* [MS-XLSB] 2.5.97.98 SharedParsedFormula */
var parse_XLSBSharedParsedFormula = parse_XLSBParsedFormula;
/* [MS-XLS] 2.5.198.4 */
var Cetab = {
0x0000: 'BEEP',
0x0001: 'OPEN',
0x0002: 'OPEN.LINKS',
0x0003: 'CLOSE.ALL',
0x0004: 'SAVE',
0x0005: 'SAVE.AS',
0x0006: 'FILE.DELETE',
0x0007: 'PAGE.SETUP',
0x0008: 'PRINT',
0x0009: 'PRINTER.SETUP',
0x000A: 'QUIT',
0x000B: 'NEW.WINDOW',
0x000C: 'ARRANGE.ALL',
0x000D: 'WINDOW.SIZE',
0x000E: 'WINDOW.MOVE',
0x000F: 'FULL',
0x0010: 'CLOSE',
0x0011: 'RUN',
0x0016: 'SET.PRINT.AREA',
0x0017: 'SET.PRINT.TITLES',
0x0018: 'SET.PAGE.BREAK',
0x0019: 'REMOVE.PAGE.BREAK',
0x001A: 'FONT',
0x001B: 'DISPLAY',
0x001C: 'PROTECT.DOCUMENT',
0x001D: 'PRECISION',
0x001E: 'A1.R1C1',
0x001F: 'CALCULATE.NOW',
0x0020: 'CALCULATION',
0x0022: 'DATA.FIND',
0x0023: 'EXTRACT',
0x0024: 'DATA.DELETE',
0x0025: 'SET.DATABASE',
0x0026: 'SET.CRITERIA',
0x0027: 'SORT',
0x0028: 'DATA.SERIES',
0x0029: 'TABLE',
0x002A: 'FORMAT.NUMBER',
0x002B: 'ALIGNMENT',
0x002C: 'STYLE',
0x002D: 'BORDER',
0x002E: 'CELL.PROTECTION',
0x002F: 'COLUMN.WIDTH',
0x0030: 'UNDO',
0x0031: 'CUT',
0x0032: 'COPY',
0x0033: 'PASTE',
0x0034: 'CLEAR',
0x0035: 'PASTE.SPECIAL',
0x0036: 'EDIT.DELETE',
0x0037: 'INSERT',
0x0038: 'FILL.RIGHT',
0x0039: 'FILL.DOWN',
0x003D: 'DEFINE.NAME',
0x003E: 'CREATE.NAMES',
0x003F: 'FORMULA.GOTO',
0x0040: 'FORMULA.FIND',
0x0041: 'SELECT.LAST.CELL',
0x0042: 'SHOW.ACTIVE.CELL',
0x0043: 'GALLERY.AREA',
0x0044: 'GALLERY.BAR',
0x0045: 'GALLERY.COLUMN',
0x0046: 'GALLERY.LINE',
0x0047: 'GALLERY.PIE',
0x0048: 'GALLERY.SCATTER',
0x0049: 'COMBINATION',
0x004A: 'PREFERRED',
0x004B: 'ADD.OVERLAY',
0x004C: 'GRIDLINES',
0x004D: 'SET.PREFERRED',
0x004E: 'AXES',
0x004F: 'LEGEND',
0x0050: 'ATTACH.TEXT',
0x0051: 'ADD.ARROW',
0x0052: 'SELECT.CHART',
0x0053: 'SELECT.PLOT.AREA',
0x0054: 'PATTERNS',
0x0055: 'MAIN.CHART',
0x0056: 'OVERLAY',
0x0057: 'SCALE',
0x0058: 'FORMAT.LEGEND',
0x0059: 'FORMAT.TEXT',
0x005A: 'EDIT.REPEAT',
0x005B: 'PARSE',
0x005C: 'JUSTIFY',
0x005D: 'HIDE',
0x005E: 'UNHIDE',
0x005F: 'WORKSPACE',
0x0060: 'FORMULA',
0x0061: 'FORMULA.FILL',
0x0062: 'FORMULA.ARRAY',
0x0063: 'DATA.FIND.NEXT',
0x0064: 'DATA.FIND.PREV',
0x0065: 'FORMULA.FIND.NEXT',
0x0066: 'FORMULA.FIND.PREV',
0x0067: 'ACTIVATE',
0x0068: 'ACTIVATE.NEXT',
0x0069: 'ACTIVATE.PREV',
0x006A: 'UNLOCKED.NEXT',
0x006B: 'UNLOCKED.PREV',
0x006C: 'COPY.PICTURE',
0x006D: 'SELECT',
0x006E: 'DELETE.NAME',
0x006F: 'DELETE.FORMAT',
0x0070: 'VLINE',
0x0071: 'HLINE',
0x0072: 'VPAGE',
0x0073: 'HPAGE',
0x0074: 'VSCROLL',
0x0075: 'HSCROLL',
0x0076: 'ALERT',
0x0077: 'NEW',
0x0078: 'CANCEL.COPY',
0x0079: 'SHOW.CLIPBOARD',
0x007A: 'MESSAGE',
0x007C: 'PASTE.LINK',
0x007D: 'APP.ACTIVATE',
0x007E: 'DELETE.ARROW',
0x007F: 'ROW.HEIGHT',
0x0080: 'FORMAT.MOVE',
0x0081: 'FORMAT.SIZE',
0x0082: 'FORMULA.REPLACE',
0x0083: 'SEND.KEYS',
0x0084: 'SELECT.SPECIAL',
0x0085: 'APPLY.NAMES',
0x0086: 'REPLACE.FONT',
0x0087: 'FREEZE.PANES',
0x0088: 'SHOW.INFO',
0x0089: 'SPLIT',
0x008A: 'ON.WINDOW',
0x008B: 'ON.DATA',
0x008C: 'DISABLE.INPUT',
0x008E: 'OUTLINE',
0x008F: 'LIST.NAMES',
0x0090: 'FILE.CLOSE',
0x0091: 'SAVE.WORKBOOK',
0x0092: 'DATA.FORM',
0x0093: 'COPY.CHART',
0x0094: 'ON.TIME',
0x0095: 'WAIT',
0x0096: 'FORMAT.FONT',
0x0097: 'FILL.UP',
0x0098: 'FILL.LEFT',
0x0099: 'DELETE.OVERLAY',
0x009B: 'SHORT.MENUS',
0x009F: 'SET.UPDATE.STATUS',
0x00A1: 'COLOR.PALETTE',
0x00A2: 'DELETE.STYLE',
0x00A3: 'WINDOW.RESTORE',
0x00A4: 'WINDOW.MAXIMIZE',
0x00A6: 'CHANGE.LINK',
0x00A7: 'CALCULATE.DOCUMENT',
0x00A8: 'ON.KEY',
0x00A9: 'APP.RESTORE',
0x00AA: 'APP.MOVE',
0x00AB: 'APP.SIZE',
0x00AC: 'APP.MINIMIZE',
0x00AD: 'APP.MAXIMIZE',
0x00AE: 'BRING.TO.FRONT',
0x00AF: 'SEND.TO.BACK',
0x00B9: 'MAIN.CHART.TYPE',
0x00BA: 'OVERLAY.CHART.TYPE',
0x00BB: 'SELECT.END',
0x00BC: 'OPEN.MAIL',
0x00BD: 'SEND.MAIL',
0x00BE: 'STANDARD.FONT',
0x00BF: 'CONSOLIDATE',
0x00C0: 'SORT.SPECIAL',
0x00C1: 'GALLERY.3D.AREA',
0x00C2: 'GALLERY.3D.COLUMN',
0x00C3: 'GALLERY.3D.LINE',
0x00C4: 'GALLERY.3D.PIE',
0x00C5: 'VIEW.3D',
0x00C6: 'GOAL.SEEK',
0x00C7: 'WORKGROUP',
0x00C8: 'FILL.GROUP',
0x00C9: 'UPDATE.LINK',
0x00CA: 'PROMOTE',
0x00CB: 'DEMOTE',
0x00CC: 'SHOW.DETAIL',
0x00CE: 'UNGROUP',
0x00CF: 'OBJECT.PROPERTIES',
0x00D0: 'SAVE.NEW.OBJECT',
0x00D1: 'SHARE',
0x00D2: 'SHARE.NAME',
0x00D3: 'DUPLICATE',
0x00D4: 'APPLY.STYLE',
0x00D5: 'ASSIGN.TO.OBJECT',
0x00D6: 'OBJECT.PROTECTION',
0x00D7: 'HIDE.OBJECT',
0x00D8: 'SET.EXTRACT',
0x00D9: 'CREATE.PUBLISHER',
0x00DA: 'SUBSCRIBE.TO',
0x00DB: 'ATTRIBUTES',
0x00DC: 'SHOW.TOOLBAR',
0x00DE: 'PRINT.PREVIEW',
0x00DF: 'EDIT.COLOR',
0x00E0: 'SHOW.LEVELS',
0x00E1: 'FORMAT.MAIN',
0x00E2: 'FORMAT.OVERLAY',
0x00E3: 'ON.RECALC',
0x00E4: 'EDIT.SERIES',
0x00E5: 'DEFINE.STYLE',
0x00F0: 'LINE.PRINT',
0x00F3: 'ENTER.DATA',
0x00F9: 'GALLERY.RADAR',
0x00FA: 'MERGE.STYLES',
0x00FB: 'EDITION.OPTIONS',
0x00FC: 'PASTE.PICTURE',
0x00FD: 'PASTE.PICTURE.LINK',
0x00FE: 'SPELLING',
0x0100: 'ZOOM',
0x0103: 'INSERT.OBJECT',
0x0104: 'WINDOW.MINIMIZE',
0x0109: 'SOUND.NOTE',
0x010A: 'SOUND.PLAY',
0x010B: 'FORMAT.SHAPE',
0x010C: 'EXTEND.POLYGON',
0x010D: 'FORMAT.AUTO',
0x0110: 'GALLERY.3D.BAR',
0x0111: 'GALLERY.3D.SURFACE',
0x0112: 'FILL.AUTO',
0x0114: 'CUSTOMIZE.TOOLBAR',
0x0115: 'ADD.TOOL',
0x0116: 'EDIT.OBJECT',
0x0117: 'ON.DOUBLECLICK',
0x0118: 'ON.ENTRY',
0x0119: 'WORKBOOK.ADD',
0x011A: 'WORKBOOK.MOVE',
0x011B: 'WORKBOOK.COPY',
0x011C: 'WORKBOOK.OPTIONS',
0x011D: 'SAVE.WORKSPACE',
0x0120: 'CHART.WIZARD',
0x0121: 'DELETE.TOOL',
0x0122: 'MOVE.TOOL',
0x0123: 'WORKBOOK.SELECT',
0x0124: 'WORKBOOK.ACTIVATE',
0x0125: 'ASSIGN.TO.TOOL',
0x0127: 'COPY.TOOL',
0x0128: 'RESET.TOOL',
0x0129: 'CONSTRAIN.NUMERIC',
0x012A: 'PASTE.TOOL',
0x012E: 'WORKBOOK.NEW',
0x0131: 'SCENARIO.CELLS',
0x0132: 'SCENARIO.DELETE',
0x0133: 'SCENARIO.ADD',
0x0134: 'SCENARIO.EDIT',
0x0135: 'SCENARIO.SHOW',
0x0136: 'SCENARIO.SHOW.NEXT',
0x0137: 'SCENARIO.SUMMARY',
0x0138: 'PIVOT.TABLE.WIZARD',
0x0139: 'PIVOT.FIELD.PROPERTIES',
0x013A: 'PIVOT.FIELD',
0x013B: 'PIVOT.ITEM',
0x013C: 'PIVOT.ADD.FIELDS',
0x013E: 'OPTIONS.CALCULATION',
0x013F: 'OPTIONS.EDIT',
0x0140: 'OPTIONS.VIEW',
0x0141: 'ADDIN.MANAGER',
0x0142: 'MENU.EDITOR',
0x0143: 'ATTACH.TOOLBARS',
0x0144: 'VBAActivate',
0x0145: 'OPTIONS.CHART',
0x0148: 'VBA.INSERT.FILE',
0x014A: 'VBA.PROCEDURE.DEFINITION',
0x0150: 'ROUTING.SLIP',
0x0152: 'ROUTE.DOCUMENT',
0x0153: 'MAIL.LOGON',
0x0156: 'INSERT.PICTURE',
0x0157: 'EDIT.TOOL',
0x0158: 'GALLERY.DOUGHNUT',
0x015E: 'CHART.TREND',
0x0160: 'PIVOT.ITEM.PROPERTIES',
0x0162: 'WORKBOOK.INSERT',
0x0163: 'OPTIONS.TRANSITION',
0x0164: 'OPTIONS.GENERAL',
0x0172: 'FILTER.ADVANCED',
0x0175: 'MAIL.ADD.MAILER',
0x0176: 'MAIL.DELETE.MAILER',
0x0177: 'MAIL.REPLY',
0x0178: 'MAIL.REPLY.ALL',
0x0179: 'MAIL.FORWARD',
0x017A: 'MAIL.NEXT.LETTER',
0x017B: 'DATA.LABEL',
0x017C: 'INSERT.TITLE',
0x017D: 'FONT.PROPERTIES',
0x017E: 'MACRO.OPTIONS',
0x017F: 'WORKBOOK.HIDE',
0x0180: 'WORKBOOK.UNHIDE',
0x0181: 'WORKBOOK.DELETE',
0x0182: 'WORKBOOK.NAME',
0x0184: 'GALLERY.CUSTOM',
0x0186: 'ADD.CHART.AUTOFORMAT',
0x0187: 'DELETE.CHART.AUTOFORMAT',
0x0188: 'CHART.ADD.DATA',
0x0189: 'AUTO.OUTLINE',
0x018A: 'TAB.ORDER',
0x018B: 'SHOW.DIALOG',
0x018C: 'SELECT.ALL',
0x018D: 'UNGROUP.SHEETS',
0x018E: 'SUBTOTAL.CREATE',
0x018F: 'SUBTOTAL.REMOVE',
0x0190: 'RENAME.OBJECT',
0x019C: 'WORKBOOK.SCROLL',
0x019D: 'WORKBOOK.NEXT',
0x019E: 'WORKBOOK.PREV',
0x019F: 'WORKBOOK.TAB.SPLIT',
0x01A0: 'FULL.SCREEN',
0x01A1: 'WORKBOOK.PROTECT',
0x01A4: 'SCROLLBAR.PROPERTIES',
0x01A5: 'PIVOT.SHOW.PAGES',
0x01A6: 'TEXT.TO.COLUMNS',
0x01A7: 'FORMAT.CHARTTYPE',
0x01A8: 'LINK.FORMAT',
0x01A9: 'TRACER.DISPLAY',
0x01AE: 'TRACER.NAVIGATE',
0x01AF: 'TRACER.CLEAR',
0x01B0: 'TRACER.ERROR',
0x01B1: 'PIVOT.FIELD.GROUP',
0x01B2: 'PIVOT.FIELD.UNGROUP',
0x01B3: 'CHECKBOX.PROPERTIES',
0x01B4: 'LABEL.PROPERTIES',
0x01B5: 'LISTBOX.PROPERTIES',
0x01B6: 'EDITBOX.PROPERTIES',
0x01B7: 'PIVOT.REFRESH',
0x01B8: 'LINK.COMBO',
0x01B9: 'OPEN.TEXT',
0x01BA: 'HIDE.DIALOG',
0x01BB: 'SET.DIALOG.FOCUS',
0x01BC: 'ENABLE.OBJECT',
0x01BD: 'PUSHBUTTON.PROPERTIES',
0x01BE: 'SET.DIALOG.DEFAULT',
0x01BF: 'FILTER',
0x01C0: 'FILTER.SHOW.ALL',
0x01C1: 'CLEAR.OUTLINE',
0x01C2: 'FUNCTION.WIZARD',
0x01C3: 'ADD.LIST.ITEM',
0x01C4: 'SET.LIST.ITEM',
0x01C5: 'REMOVE.LIST.ITEM',
0x01C6: 'SELECT.LIST.ITEM',
0x01C7: 'SET.CONTROL.VALUE',
0x01C8: 'SAVE.COPY.AS',
0x01CA: 'OPTIONS.LISTS.ADD',
0x01CB: 'OPTIONS.LISTS.DELETE',
0x01CC: 'SERIES.AXES',
0x01CD: 'SERIES.X',
0x01CE: 'SERIES.Y',
0x01CF: 'ERRORBAR.X',
0x01D0: 'ERRORBAR.Y',
0x01D1: 'FORMAT.CHART',
0x01D2: 'SERIES.ORDER',
0x01D3: 'MAIL.LOGOFF',
0x01D4: 'CLEAR.ROUTING.SLIP',
0x01D5: 'APP.ACTIVATE.MICROSOFT',
0x01D6: 'MAIL.EDIT.MAILER',
0x01D7: 'ON.SHEET',
0x01D8: 'STANDARD.WIDTH',
0x01D9: 'SCENARIO.MERGE',
0x01DA: 'SUMMARY.INFO',
0x01DB: 'FIND.FILE',
0x01DC: 'ACTIVE.CELL.FONT',
0x01DD: 'ENABLE.TIPWIZARD',
0x01DE: 'VBA.MAKE.ADDIN',
0x01E0: 'INSERTDATATABLE',
0x01E1: 'WORKGROUP.OPTIONS',
0x01E2: 'MAIL.SEND.MAILER',
0x01E5: 'AUTOCORRECT',
0x01E9: 'POST.DOCUMENT',
0x01EB: 'PICKLIST',
0x01ED: 'VIEW.SHOW',
0x01EE: 'VIEW.DEFINE',
0x01EF: 'VIEW.DELETE',
0x01FD: 'SHEET.BACKGROUND',
0x01FE: 'INSERT.MAP.OBJECT',
0x01FF: 'OPTIONS.MENONO',
0x0205: 'MSOCHECKS',
0x0206: 'NORMAL',
0x0207: 'LAYOUT',
0x0208: 'RM.PRINT.AREA',
0x0209: 'CLEAR.PRINT.AREA',
0x020A: 'ADD.PRINT.AREA',
0x020B: 'MOVE.BRK',
0x0221: 'HIDECURR.NOTE',
0x0222: 'HIDEALL.NOTES',
0x0223: 'DELETE.NOTE',
0x0224: 'TRAVERSE.NOTES',
0x0225: 'ACTIVATE.NOTES',
0x026C: 'PROTECT.REVISIONS',
0x026D: 'UNPROTECT.REVISIONS',
0x0287: 'OPTIONS.ME',
0x028D: 'WEB.PUBLISH',
0x029B: 'NEWWEBQUERY',
0x02A1: 'PIVOT.TABLE.CHART',
0x02F1: 'OPTIONS.SAVE',
0x02F3: 'OPTIONS.SPELL',
0x0328: 'HIDEALL.INKANNOTS'
};

/* [MS-XLS] 2.5.198.17 */
/* [MS-XLSB] 2.5.97.10 */
var Ftab = {
0x0000: 'COUNT',
0x0001: 'IF',
0x0002: 'ISNA',
0x0003: 'ISERROR',
0x0004: 'SUM',
0x0005: 'AVERAGE',
0x0006: 'MIN',
0x0007: 'MAX',
0x0008: 'ROW',
0x0009: 'COLUMN',
0x000A: 'NA',
0x000B: 'NPV',
0x000C: 'STDEV',
0x000D: 'DOLLAR',
0x000E: 'FIXED',
0x000F: 'SIN',
0x0010: 'COS',
0x0011: 'TAN',
0x0012: 'ATAN',
0x0013: 'PI',
0x0014: 'SQRT',
0x0015: 'EXP',
0x0016: 'LN',
0x0017: 'LOG10',
0x0018: 'ABS',
0x0019: 'INT',
0x001A: 'SIGN',
0x001B: 'ROUND',
0x001C: 'LOOKUP',
0x001D: 'INDEX',
0x001E: 'REPT',
0x001F: 'MID',
0x0020: 'LEN',
0x0021: 'VALUE',
0x0022: 'TRUE',
0x0023: 'FALSE',
0x0024: 'AND',
0x0025: 'OR',
0x0026: 'NOT',
0x0027: 'MOD',
0x0028: 'DCOUNT',
0x0029: 'DSUM',
0x002A: 'DAVERAGE',
0x002B: 'DMIN',
0x002C: 'DMAX',
0x002D: 'DSTDEV',
0x002E: 'VAR',
0x002F: 'DVAR',
0x0030: 'TEXT',
0x0031: 'LINEST',
0x0032: 'TREND',
0x0033: 'LOGEST',
0x0034: 'GROWTH',
0x0035: 'GOTO',
0x0036: 'HALT',
0x0037: 'RETURN',
0x0038: 'PV',
0x0039: 'FV',
0x003A: 'NPER',
0x003B: 'PMT',
0x003C: 'RATE',
0x003D: 'MIRR',
0x003E: 'IRR',
0x003F: 'RAND',
0x0040: 'MATCH',
0x0041: 'DATE',
0x0042: 'TIME',
0x0043: 'DAY',
0x0044: 'MONTH',
0x0045: 'YEAR',
0x0046: 'WEEKDAY',
0x0047: 'HOUR',
0x0048: 'MINUTE',
0x0049: 'SECOND',
0x004A: 'NOW',
0x004B: 'AREAS',
0x004C: 'ROWS',
0x004D: 'COLUMNS',
0x004E: 'OFFSET',
0x004F: 'ABSREF',
0x0050: 'RELREF',
0x0051: 'ARGUMENT',
0x0052: 'SEARCH',
0x0053: 'TRANSPOSE',
0x0054: 'ERROR',
0x0055: 'STEP',
0x0056: 'TYPE',
0x0057: 'ECHO',
0x0058: 'SET.NAME',
0x0059: 'CALLER',
0x005A: 'DEREF',
0x005B: 'WINDOWS',
0x005C: 'SERIES',
0x005D: 'DOCUMENTS',
0x005E: 'ACTIVE.CELL',
0x005F: 'SELECTION',
0x0060: 'RESULT',
0x0061: 'ATAN2',
0x0062: 'ASIN',
0x0063: 'ACOS',
0x0064: 'CHOOSE',
0x0065: 'HLOOKUP',
0x0066: 'VLOOKUP',
0x0067: 'LINKS',
0x0068: 'INPUT',
0x0069: 'ISREF',
0x006A: 'GET.FORMULA',
0x006B: 'GET.NAME',
0x006C: 'SET.VALUE',
0x006D: 'LOG',
0x006E: 'EXEC',
0x006F: 'CHAR',
0x0070: 'LOWER',
0x0071: 'UPPER',
0x0072: 'PROPER',
0x0073: 'LEFT',
0x0074: 'RIGHT',
0x0075: 'EXACT',
0x0076: 'TRIM',
0x0077: 'REPLACE',
0x0078: 'SUBSTITUTE',
0x0079: 'CODE',
0x007A: 'NAMES',
0x007B: 'DIRECTORY',
0x007C: 'FIND',
0x007D: 'CELL',
0x007E: 'ISERR',
0x007F: 'ISTEXT',
0x0080: 'ISNUMBER',
0x0081: 'ISBLANK',
0x0082: 'T',
0x0083: 'N',
0x0084: 'FOPEN',
0x0085: 'FCLOSE',
0x0086: 'FSIZE',
0x0087: 'FREADLN',
0x0088: 'FREAD',
0x0089: 'FWRITELN',
0x008A: 'FWRITE',
0x008B: 'FPOS',
0x008C: 'DATEVALUE',
0x008D: 'TIMEVALUE',
0x008E: 'SLN',
0x008F: 'SYD',
0x0090: 'DDB',
0x0091: 'GET.DEF',
0x0092: 'REFTEXT',
0x0093: 'TEXTREF',
0x0094: 'INDIRECT',
0x0095: 'REGISTER',
0x0096: 'CALL',
0x0097: 'ADD.BAR',
0x0098: 'ADD.MENU',
0x0099: 'ADD.COMMAND',
0x009A: 'ENABLE.COMMAND',
0x009B: 'CHECK.COMMAND',
0x009C: 'RENAME.COMMAND',
0x009D: 'SHOW.BAR',
0x009E: 'DELETE.MENU',
0x009F: 'DELETE.COMMAND',
0x00A0: 'GET.CHART.ITEM',
0x00A1: 'DIALOG.BOX',
0x00A2: 'CLEAN',
0x00A3: 'MDETERM',
0x00A4: 'MINVERSE',
0x00A5: 'MMULT',
0x00A6: 'FILES',
0x00A7: 'IPMT',
0x00A8: 'PPMT',
0x00A9: 'COUNTA',
0x00AA: 'CANCEL.KEY',
0x00AB: 'FOR',
0x00AC: 'WHILE',
0x00AD: 'BREAK',
0x00AE: 'NEXT',
0x00AF: 'INITIATE',
0x00B0: 'REQUEST',
0x00B1: 'POKE',
0x00B2: 'EXECUTE',
0x00B3: 'TERMINATE',
0x00B4: 'RESTART',
0x00B5: 'HELP',
0x00B6: 'GET.BAR',
0x00B7: 'PRODUCT',
0x00B8: 'FACT',
0x00B9: 'GET.CELL',
0x00BA: 'GET.WORKSPACE',
0x00BB: 'GET.WINDOW',
0x00BC: 'GET.DOCUMENT',
0x00BD: 'DPRODUCT',
0x00BE: 'ISNONTEXT',
0x00BF: 'GET.NOTE',
0x00C0: 'NOTE',
0x00C1: 'STDEVP',
0x00C2: 'VARP',
0x00C3: 'DSTDEVP',
0x00C4: 'DVARP',
0x00C5: 'TRUNC',
0x00C6: 'ISLOGICAL',
0x00C7: 'DCOUNTA',
0x00C8: 'DELETE.BAR',
0x00C9: 'UNREGISTER',
0x00CC: 'USDOLLAR',
0x00CD: 'FINDB',
0x00CE: 'SEARCHB',
0x00CF: 'REPLACEB',
0x00D0: 'LEFTB',
0x00D1: 'RIGHTB',
0x00D2: 'MIDB',
0x00D3: 'LENB',
0x00D4: 'ROUNDUP',
0x00D5: 'ROUNDDOWN',
0x00D6: 'ASC',
0x00D7: 'DBCS',
0x00D8: 'RANK',
0x00DB: 'ADDRESS',
0x00DC: 'DAYS360',
0x00DD: 'TODAY',
0x00DE: 'VDB',
0x00DF: 'ELSE',
0x00E0: 'ELSE.IF',
0x00E1: 'END.IF',
0x00E2: 'FOR.CELL',
0x00E3: 'MEDIAN',
0x00E4: 'SUMPRODUCT',
0x00E5: 'SINH',
0x00E6: 'COSH',
0x00E7: 'TANH',
0x00E8: 'ASINH',
0x00E9: 'ACOSH',
0x00EA: 'ATANH',
0x00EB: 'DGET',
0x00EC: 'CREATE.OBJECT',
0x00ED: 'VOLATILE',
0x00EE: 'LAST.ERROR',
0x00EF: 'CUSTOM.UNDO',
0x00F0: 'CUSTOM.REPEAT',
0x00F1: 'FORMULA.CONVERT',
0x00F2: 'GET.LINK.INFO',
0x00F3: 'TEXT.BOX',
0x00F4: 'INFO',
0x00F5: 'GROUP',
0x00F6: 'GET.OBJECT',
0x00F7: 'DB',
0x00F8: 'PAUSE',
0x00FB: 'RESUME',
0x00FC: 'FREQUENCY',
0x00FD: 'ADD.TOOLBAR',
0x00FE: 'DELETE.TOOLBAR',
0x00FF: 'User',
0x0100: 'RESET.TOOLBAR',
0x0101: 'EVALUATE',
0x0102: 'GET.TOOLBAR',
0x0103: 'GET.TOOL',
0x0104: 'SPELLING.CHECK',
0x0105: 'ERROR.TYPE',
0x0106: 'APP.TITLE',
0x0107: 'WINDOW.TITLE',
0x0108: 'SAVE.TOOLBAR',
0x0109: 'ENABLE.TOOL',
0x010A: 'PRESS.TOOL',
0x010B: 'REGISTER.ID',
0x010C: 'GET.WORKBOOK',
0x010D: 'AVEDEV',
0x010E: 'BETADIST',
0x010F: 'GAMMALN',
0x0110: 'BETAINV',
0x0111: 'BINOMDIST',
0x0112: 'CHIDIST',
0x0113: 'CHIINV',
0x0114: 'COMBIN',
0x0115: 'CONFIDENCE',
0x0116: 'CRITBINOM',
0x0117: 'EVEN',
0x0118: 'EXPONDIST',
0x0119: 'FDIST',
0x011A: 'FINV',
0x011B: 'FISHER',
0x011C: 'FISHERINV',
0x011D: 'FLOOR',
0x011E: 'GAMMADIST',
0x011F: 'GAMMAINV',
0x0120: 'CEILING',
0x0121: 'HYPGEOMDIST',
0x0122: 'LOGNORMDIST',
0x0123: 'LOGINV',
0x0124: 'NEGBINOMDIST',
0x0125: 'NORMDIST',
0x0126: 'NORMSDIST',
0x0127: 'NORMINV',
0x0128: 'NORMSINV',
0x0129: 'STANDARDIZE',
0x012A: 'ODD',
0x012B: 'PERMUT',
0x012C: 'POISSON',
0x012D: 'TDIST',
0x012E: 'WEIBULL',
0x012F: 'SUMXMY2',
0x0130: 'SUMX2MY2',
0x0131: 'SUMX2PY2',
0x0132: 'CHITEST',
0x0133: 'CORREL',
0x0134: 'COVAR',
0x0135: 'FORECAST',
0x0136: 'FTEST',
0x0137: 'INTERCEPT',
0x0138: 'PEARSON',
0x0139: 'RSQ',
0x013A: 'STEYX',
0x013B: 'SLOPE',
0x013C: 'TTEST',
0x013D: 'PROB',
0x013E: 'DEVSQ',
0x013F: 'GEOMEAN',
0x0140: 'HARMEAN',
0x0141: 'SUMSQ',
0x0142: 'KURT',
0x0143: 'SKEW',
0x0144: 'ZTEST',
0x0145: 'LARGE',
0x0146: 'SMALL',
0x0147: 'QUARTILE',
0x0148: 'PERCENTILE',
0x0149: 'PERCENTRANK',
0x014A: 'MODE',
0x014B: 'TRIMMEAN',
0x014C: 'TINV',
0x014E: 'MOVIE.COMMAND',
0x014F: 'GET.MOVIE',
0x0150: 'CONCATENATE',
0x0151: 'POWER',
0x0152: 'PIVOT.ADD.DATA',
0x0153: 'GET.PIVOT.TABLE',
0x0154: 'GET.PIVOT.FIELD',
0x0155: 'GET.PIVOT.ITEM',
0x0156: 'RADIANS',
0x0157: 'DEGREES',
0x0158: 'SUBTOTAL',
0x0159: 'SUMIF',
0x015A: 'COUNTIF',
0x015B: 'COUNTBLANK',
0x015C: 'SCENARIO.GET',
0x015D: 'OPTIONS.LISTS.GET',
0x015E: 'ISPMT',
0x015F: 'DATEDIF',
0x0160: 'DATESTRING',
0x0161: 'NUMBERSTRING',
0x0162: 'ROMAN',
0x0163: 'OPEN.DIALOG',
0x0164: 'SAVE.DIALOG',
0x0165: 'VIEW.GET',
0x0166: 'GETPIVOTDATA',
0x0167: 'HYPERLINK',
0x0168: 'PHONETIC',
0x0169: 'AVERAGEA',
0x016A: 'MAXA',
0x016B: 'MINA',
0x016C: 'STDEVPA',
0x016D: 'VARPA',
0x016E: 'STDEVA',
0x016F: 'VARA',
0x0170: 'BAHTTEXT',
0x0171: 'THAIDAYOFWEEK',
0x0172: 'THAIDIGIT',
0x0173: 'THAIMONTHOFYEAR',
0x0174: 'THAINUMSOUND',
0x0175: 'THAINUMSTRING',
0x0176: 'THAISTRINGLENGTH',
0x0177: 'ISTHAIDIGIT',
0x0178: 'ROUNDBAHTDOWN',
0x0179: 'ROUNDBAHTUP',
0x017A: 'THAIYEAR',
0x017B: 'RTD',

0x017C: 'CUBEVALUE',
0x017D: 'CUBEMEMBER',
0x017E: 'CUBEMEMBERPROPERTY',
0x017F: 'CUBERANKEDMEMBER',
0x0180: 'HEX2BIN',
0x0181: 'HEX2DEC',
0x0182: 'HEX2OCT',
0x0183: 'DEC2BIN',
0x0184: 'DEC2HEX',
0x0185: 'DEC2OCT',
0x0186: 'OCT2BIN',
0x0187: 'OCT2HEX',
0x0188: 'OCT2DEC',
0x0189: 'BIN2DEC',
0x018A: 'BIN2OCT',
0x018B: 'BIN2HEX',
0x018C: 'IMSUB',
0x018D: 'IMDIV',
0x018E: 'IMPOWER',
0x018F: 'IMABS',
0x0190: 'IMSQRT',
0x0191: 'IMLN',
0x0192: 'IMLOG2',
0x0193: 'IMLOG10',
0x0194: 'IMSIN',
0x0195: 'IMCOS',
0x0196: 'IMEXP',
0x0197: 'IMARGUMENT',
0x0198: 'IMCONJUGATE',
0x0199: 'IMAGINARY',
0x019A: 'IMREAL',
0x019B: 'COMPLEX',
0x019C: 'IMSUM',
0x019D: 'IMPRODUCT',
0x019E: 'SERIESSUM',
0x019F: 'FACTDOUBLE',
0x01A0: 'SQRTPI',
0x01A1: 'QUOTIENT',
0x01A2: 'DELTA',
0x01A3: 'GESTEP',
0x01A4: 'ISEVEN',
0x01A5: 'ISODD',
0x01A6: 'MROUND',
0x01A7: 'ERF',
0x01A8: 'ERFC',
0x01A9: 'BESSELJ',
0x01AA: 'BESSELK',
0x01AB: 'BESSELY',
0x01AC: 'BESSELI',
0x01AD: 'XIRR',
0x01AE: 'XNPV',
0x01AF: 'PRICEMAT',
0x01B0: 'YIELDMAT',
0x01B1: 'INTRATE',
0x01B2: 'RECEIVED',
0x01B3: 'DISC',
0x01B4: 'PRICEDISC',
0x01B5: 'YIELDDISC',
0x01B6: 'TBILLEQ',
0x01B7: 'TBILLPRICE',
0x01B8: 'TBILLYIELD',
0x01B9: 'PRICE',
0x01BA: 'YIELD',
0x01BB: 'DOLLARDE',
0x01BC: 'DOLLARFR',
0x01BD: 'NOMINAL',
0x01BE: 'EFFECT',
0x01BF: 'CUMPRINC',
0x01C0: 'CUMIPMT',
0x01C1: 'EDATE',
0x01C2: 'EOMONTH',
0x01C3: 'YEARFRAC',
0x01C4: 'COUPDAYBS',
0x01C5: 'COUPDAYS',
0x01C6: 'COUPDAYSNC',
0x01C7: 'COUPNCD',
0x01C8: 'COUPNUM',
0x01C9: 'COUPPCD',
0x01CA: 'DURATION',
0x01CB: 'MDURATION',
0x01CC: 'ODDLPRICE',
0x01CD: 'ODDLYIELD',
0x01CE: 'ODDFPRICE',
0x01CF: 'ODDFYIELD',
0x01D0: 'RANDBETWEEN',
0x01D1: 'WEEKNUM',
0x01D2: 'AMORDEGRC',
0x01D3: 'AMORLINC',
0x01D4: 'CONVERT',
0x02D4: 'SHEETJS',
0x01D5: 'ACCRINT',
0x01D6: 'ACCRINTM',
0x01D7: 'WORKDAY',
0x01D8: 'NETWORKDAYS',
0x01D9: 'GCD',
0x01DA: 'MULTINOMIAL',
0x01DB: 'LCM',
0x01DC: 'FVSCHEDULE',
0x01DD: 'CUBEKPIMEMBER',
0x01DE: 'CUBESET',
0x01DF: 'CUBESETCOUNT',
0x01E0: 'IFERROR',
0x01E1: 'COUNTIFS',
0x01E2: 'SUMIFS',
0x01E3: 'AVERAGEIF',
0x01E4: 'AVERAGEIFS'
};
var FtabArgc = {
0x0002: 1, /* ISNA */
0x0003: 1, /* ISERROR */
0x000A: 0, /* NA */
0x000F: 1, /* SIN */
0x0010: 1, /* COS */
0x0011: 1, /* TAN */
0x0012: 1, /* ATAN */
0x0013: 0, /* PI */
0x0014: 1, /* SQRT */
0x0015: 1, /* EXP */
0x0016: 1, /* LN */
0x0017: 1, /* LOG10 */
0x0018: 1, /* ABS */
0x0019: 1, /* INT */
0x001A: 1, /* SIGN */
0x001B: 2, /* ROUND */
0x001E: 2, /* REPT */
0x001F: 3, /* MID */
0x0020: 1, /* LEN */
0x0021: 1, /* VALUE */
0x0022: 0, /* TRUE */
0x0023: 0, /* FALSE */
0x0026: 1, /* NOT */
0x0027: 2, /* MOD */
0x0028: 3, /* DCOUNT */
0x0029: 3, /* DSUM */
0x002A: 3, /* DAVERAGE */
0x002B: 3, /* DMIN */
0x002C: 3, /* DMAX */
0x002D: 3, /* DSTDEV */
0x002F: 3, /* DVAR */
0x0030: 2, /* TEXT */
0x0035: 1, /* GOTO */
0x003D: 3, /* MIRR */
0x003F: 0, /* RAND */
0x0041: 3, /* DATE */
0x0042: 3, /* TIME */
0x0043: 1, /* DAY */
0x0044: 1, /* MONTH */
0x0045: 1, /* YEAR */
0x0046: 1, /* WEEKDAY */
0x0047: 1, /* HOUR */
0x0048: 1, /* MINUTE */
0x0049: 1, /* SECOND */
0x004A: 0, /* NOW */
0x004B: 1, /* AREAS */
0x004C: 1, /* ROWS */
0x004D: 1, /* COLUMNS */
0x004F: 2, /* ABSREF */
0x0050: 2, /* RELREF */
0x0053: 1, /* TRANSPOSE */
0x0055: 0, /* STEP */
0x0056: 1, /* TYPE */
0x0059: 0, /* CALLER */
0x005A: 1, /* DEREF */
0x005E: 0, /* ACTIVE.CELL */
0x005F: 0, /* SELECTION */
0x0061: 2, /* ATAN2 */
0x0062: 1, /* ASIN */
0x0063: 1, /* ACOS */
0x0065: 3, /* HLOOKUP */
0x0066: 3, /* VLOOKUP */
0x0069: 1, /* ISREF */
0x006A: 1, /* GET.FORMULA */
0x006C: 2, /* SET.VALUE */
0x006F: 1, /* CHAR */
0x0070: 1, /* LOWER */
0x0071: 1, /* UPPER */
0x0072: 1, /* PROPER */
0x0075: 2, /* EXACT */
0x0076: 1, /* TRIM */
0x0077: 4, /* REPLACE */
0x0079: 1, /* CODE */
0x007E: 1, /* ISERR */
0x007F: 1, /* ISTEXT */
0x0080: 1, /* ISNUMBER */
0x0081: 1, /* ISBLANK */
0x0082: 1, /* T */
0x0083: 1, /* N */
0x0085: 1, /* FCLOSE */
0x0086: 1, /* FSIZE */
0x0087: 1, /* FREADLN */
0x0088: 2, /* FREAD */
0x0089: 2, /* FWRITELN */
0x008A: 2, /* FWRITE */
0x008C: 1, /* DATEVALUE */
0x008D: 1, /* TIMEVALUE */
0x008E: 3, /* SLN */
0x008F: 4, /* SYD */
0x0090: 4, /* DDB */
0x00A1: 1, /* DIALOG.BOX */
0x00A2: 1, /* CLEAN */
0x00A3: 1, /* MDETERM */
0x00A4: 1, /* MINVERSE */
0x00A5: 2, /* MMULT */
0x00AC: 1, /* WHILE */
0x00AF: 2, /* INITIATE */
0x00B0: 2, /* REQUEST */
0x00B1: 3, /* POKE */
0x00B2: 2, /* EXECUTE */
0x00B3: 1, /* TERMINATE */
0x00B8: 1, /* FACT */
0x00BA: 1, /* GET.WORKSPACE */
0x00BD: 3, /* DPRODUCT */
0x00BE: 1, /* ISNONTEXT */
0x00C3: 3, /* DSTDEVP */
0x00C4: 3, /* DVARP */
0x00C5: 1, /* TRUNC */
0x00C6: 1, /* ISLOGICAL */
0x00C7: 3, /* DCOUNTA */
0x00C9: 1, /* UNREGISTER */
0x00CF: 4, /* REPLACEB */
0x00D2: 3, /* MIDB */
0x00D3: 1, /* LENB */
0x00D4: 2, /* ROUNDUP */
0x00D5: 2, /* ROUNDDOWN */
0x00D6: 1, /* ASC */
0x00D7: 1, /* DBCS */
0x00E1: 0, /* END.IF */
0x00E5: 1, /* SINH */
0x00E6: 1, /* COSH */
0x00E7: 1, /* TANH */
0x00E8: 1, /* ASINH */
0x00E9: 1, /* ACOSH */
0x00EA: 1, /* ATANH */
0x00EB: 3, /* DGET */
0x00F4: 1, /* INFO */
0x00F7: 4, /* DB */
0x00FC: 2, /* FREQUENCY */
0x0101: 1, /* EVALUATE */
0x0105: 1, /* ERROR.TYPE */
0x010F: 1, /* GAMMALN */
0x0111: 4, /* BINOMDIST */
0x0112: 2, /* CHIDIST */
0x0113: 2, /* CHIINV */
0x0114: 2, /* COMBIN */
0x0115: 3, /* CONFIDENCE */
0x0116: 3, /* CRITBINOM */
0x0117: 1, /* EVEN */
0x0118: 3, /* EXPONDIST */
0x0119: 3, /* FDIST */
0x011A: 3, /* FINV */
0x011B: 1, /* FISHER */
0x011C: 1, /* FISHERINV */
0x011D: 2, /* FLOOR */
0x011E: 4, /* GAMMADIST */
0x011F: 3, /* GAMMAINV */
0x0120: 2, /* CEILING */
0x0121: 4, /* HYPGEOMDIST */
0x0122: 3, /* LOGNORMDIST */
0x0123: 3, /* LOGINV */
0x0124: 3, /* NEGBINOMDIST */
0x0125: 4, /* NORMDIST */
0x0126: 1, /* NORMSDIST */
0x0127: 3, /* NORMINV */
0x0128: 1, /* NORMSINV */
0x0129: 3, /* STANDARDIZE */
0x012A: 1, /* ODD */
0x012B: 2, /* PERMUT */
0x012C: 3, /* POISSON */
0x012D: 3, /* TDIST */
0x012E: 4, /* WEIBULL */
0x012F: 2, /* SUMXMY2 */
0x0130: 2, /* SUMX2MY2 */
0x0131: 2, /* SUMX2PY2 */
0x0132: 2, /* CHITEST */
0x0133: 2, /* CORREL */
0x0134: 2, /* COVAR */
0x0135: 3, /* FORECAST */
0x0136: 2, /* FTEST */
0x0137: 2, /* INTERCEPT */
0x0138: 2, /* PEARSON */
0x0139: 2, /* RSQ */
0x013A: 2, /* STEYX */
0x013B: 2, /* SLOPE */
0x013C: 4, /* TTEST */
0x0145: 2, /* LARGE */
0x0146: 2, /* SMALL */
0x0147: 2, /* QUARTILE */
0x0148: 2, /* PERCENTILE */
0x014B: 2, /* TRIMMEAN */
0x014C: 2, /* TINV */
0x0151: 2, /* POWER */
0x0156: 1, /* RADIANS */
0x0157: 1, /* DEGREES */
0x015A: 2, /* COUNTIF */
0x015B: 1, /* COUNTBLANK */
0x015E: 4, /* ISPMT */
0x015F: 3, /* DATEDIF */
0x0160: 1, /* DATESTRING */
0x0161: 2, /* NUMBERSTRING */
0x0168: 1, /* PHONETIC */
0x0170: 1, /* BAHTTEXT */
0x0171: 1, /* THAIDAYOFWEEK */
0x0172: 1, /* THAIDIGIT */
0x0173: 1, /* THAIMONTHOFYEAR */
0x0174: 1, /* THAINUMSOUND */
0x0175: 1, /* THAINUMSTRING */
0x0176: 1, /* THAISTRINGLENGTH */
0x0177: 1, /* ISTHAIDIGIT */
0x0178: 1, /* ROUNDBAHTDOWN */
0x0179: 1, /* ROUNDBAHTUP */
0x017A: 1, /* THAIYEAR */
0x017E: 3, /* CUBEMEMBERPROPERTY */
0x0181: 1, /* HEX2DEC */
0x0188: 1, /* OCT2DEC */
0x0189: 1, /* BIN2DEC */
0x018C: 2, /* IMSUB */
0x018D: 2, /* IMDIV */
0x018E: 2, /* IMPOWER */
0x018F: 1, /* IMABS */
0x0190: 1, /* IMSQRT */
0x0191: 1, /* IMLN */
0x0192: 1, /* IMLOG2 */
0x0193: 1, /* IMLOG10 */
0x0194: 1, /* IMSIN */
0x0195: 1, /* IMCOS */
0x0196: 1, /* IMEXP */
0x0197: 1, /* IMARGUMENT */
0x0198: 1, /* IMCONJUGATE */
0x0199: 1, /* IMAGINARY */
0x019A: 1, /* IMREAL */
0x019E: 4, /* SERIESSUM */
0x019F: 1, /* FACTDOUBLE */
0x01A0: 1, /* SQRTPI */
0x01A1: 2, /* QUOTIENT */
0x01A4: 1, /* ISEVEN */
0x01A5: 1, /* ISODD */
0x01A6: 2, /* MROUND */
0x01A8: 1, /* ERFC */
0x01A9: 2, /* BESSELJ */
0x01AA: 2, /* BESSELK */
0x01AB: 2, /* BESSELY */
0x01AC: 2, /* BESSELI */
0x01AE: 3, /* XNPV */
0x01B6: 3, /* TBILLEQ */
0x01B7: 3, /* TBILLPRICE */
0x01B8: 3, /* TBILLYIELD */
0x01BB: 2, /* DOLLARDE */
0x01BC: 2, /* DOLLARFR */
0x01BD: 2, /* NOMINAL */
0x01BE: 2, /* EFFECT */
0x01BF: 6, /* CUMPRINC */
0x01C0: 6, /* CUMIPMT */
0x01C1: 2, /* EDATE */
0x01C2: 2, /* EOMONTH */
0x01D0: 2, /* RANDBETWEEN */
0x01D4: 3, /* CONVERT */
0x01DC: 2, /* FVSCHEDULE */
0x01DF: 1, /* CUBESETCOUNT */
0x01E0: 2, /* IFERROR */
0xFFFF: 0
};
/* [MS-XLSX] 2.2.3 Functions */
/* [MS-XLSB] 2.5.97.10 Ftab */
var XLSXFutureFunctions = {
	"_xlfn.ACOT": "ACOT",
	"_xlfn.ACOTH": "ACOTH",
	"_xlfn.AGGREGATE": "AGGREGATE",
	"_xlfn.ARABIC": "ARABIC",
	"_xlfn.AVERAGEIF": "AVERAGEIF",
	"_xlfn.AVERAGEIFS": "AVERAGEIFS",
	"_xlfn.BASE": "BASE",
	"_xlfn.BETA.DIST": "BETA.DIST",
	"_xlfn.BETA.INV": "BETA.INV",
	"_xlfn.BINOM.DIST": "BINOM.DIST",
	"_xlfn.BINOM.DIST.RANGE": "BINOM.DIST.RANGE",
	"_xlfn.BINOM.INV": "BINOM.INV",
	"_xlfn.BITAND": "BITAND",
	"_xlfn.BITLSHIFT": "BITLSHIFT",
	"_xlfn.BITOR": "BITOR",
	"_xlfn.BITRSHIFT": "BITRSHIFT",
	"_xlfn.BITXOR": "BITXOR",
	"_xlfn.CEILING.MATH": "CEILING.MATH",
	"_xlfn.CEILING.PRECISE": "CEILING.PRECISE",
	"_xlfn.CHISQ.DIST": "CHISQ.DIST",
	"_xlfn.CHISQ.DIST.RT": "CHISQ.DIST.RT",
	"_xlfn.CHISQ.INV": "CHISQ.INV",
	"_xlfn.CHISQ.INV.RT": "CHISQ.INV.RT",
	"_xlfn.CHISQ.TEST": "CHISQ.TEST",
	"_xlfn.COMBINA": "COMBINA",
	"_xlfn.CONCAT": "CONCAT",
	"_xlfn.CONFIDENCE.NORM": "CONFIDENCE.NORM",
	"_xlfn.CONFIDENCE.T": "CONFIDENCE.T",
	"_xlfn.COT": "COT",
	"_xlfn.COTH": "COTH",
	"_xlfn.COUNTIFS": "COUNTIFS",
	"_xlfn.COVARIANCE.P": "COVARIANCE.P",
	"_xlfn.COVARIANCE.S": "COVARIANCE.S",
	"_xlfn.CSC": "CSC",
	"_xlfn.CSCH": "CSCH",
	"_xlfn.DAYS": "DAYS",
	"_xlfn.DECIMAL": "DECIMAL",
	"_xlfn.ECMA.CEILING": "ECMA.CEILING",
	"_xlfn.ERF.PRECISE": "ERF.PRECISE",
	"_xlfn.ERFC.PRECISE": "ERFC.PRECISE",
	"_xlfn.EXPON.DIST": "EXPON.DIST",
	"_xlfn.F.DIST": "F.DIST",
	"_xlfn.F.DIST.RT": "F.DIST.RT",
	"_xlfn.F.INV": "F.INV",
	"_xlfn.F.INV.RT": "F.INV.RT",
	"_xlfn.F.TEST": "F.TEST",
	"_xlfn.FILTERXML": "FILTERXML",
	"_xlfn.FLOOR.MATH": "FLOOR.MATH",
	"_xlfn.FLOOR.PRECISE": "FLOOR.PRECISE",
	"_xlfn.FORECAST.ETS": "FORECAST.ETS",
	"_xlfn.FORECAST.ETS.CONFINT": "FORECAST.ETS.CONFINT",
	"_xlfn.FORECAST.ETS.SEASONALITY": "FORECAST.ETS.SEASONALITY",
	"_xlfn.FORECAST.ETS.STAT": "FORECAST.ETS.STAT",
	"_xlfn.FORECAST.LINEAR": "FORECAST.LINEAR",
	"_xlfn.FORMULATEXT": "FORMULATEXT",
	"_xlfn.GAMMA": "GAMMA",
	"_xlfn.GAMMA.DIST": "GAMMA.DIST",
	"_xlfn.GAMMA.INV": "GAMMA.INV",
	"_xlfn.GAMMALN.PRECISE": "GAMMALN.PRECISE",
	"_xlfn.GAUSS": "GAUSS",
	"_xlfn.HYPGEOM.DIST": "HYPGEOM.DIST",
	"_xlfn.IFERROR": "IFERROR",
	"_xlfn.IFNA": "IFNA",
	"_xlfn.IFS": "IFS",
	"_xlfn.IMCOSH": "IMCOSH",
	"_xlfn.IMCOT": "IMCOT",
	"_xlfn.IMCSC": "IMCSC",
	"_xlfn.IMCSCH": "IMCSCH",
	"_xlfn.IMSEC": "IMSEC",
	"_xlfn.IMSECH": "IMSECH",
	"_xlfn.IMSINH": "IMSINH",
	"_xlfn.IMTAN": "IMTAN",
	"_xlfn.ISFORMULA": "ISFORMULA",
	"_xlfn.ISO.CEILING": "ISO.CEILING",
	"_xlfn.ISOWEEKNUM": "ISOWEEKNUM",
	"_xlfn.LOGNORM.DIST": "LOGNORM.DIST",
	"_xlfn.LOGNORM.INV": "LOGNORM.INV",
	"_xlfn.MAXIFS": "MAXIFS",
	"_xlfn.MINIFS": "MINIFS",
	"_xlfn.MODE.MULT": "MODE.MULT",
	"_xlfn.MODE.SNGL": "MODE.SNGL",
	"_xlfn.MUNIT": "MUNIT",
	"_xlfn.NEGBINOM.DIST": "NEGBINOM.DIST",
	"_xlfn.NETWORKDAYS.INTL": "NETWORKDAYS.INTL",
	"_xlfn.NIGBINOM": "NIGBINOM",
	"_xlfn.NORM.DIST": "NORM.DIST",
	"_xlfn.NORM.INV": "NORM.INV",
	"_xlfn.NORM.S.DIST": "NORM.S.DIST",
	"_xlfn.NORM.S.INV": "NORM.S.INV",
	"_xlfn.NUMBERVALUE": "NUMBERVALUE",
	"_xlfn.PDURATION": "PDURATION",
	"_xlfn.PERCENTILE.EXC": "PERCENTILE.EXC",
	"_xlfn.PERCENTILE.INC": "PERCENTILE.INC",
	"_xlfn.PERCENTRANK.EXC": "PERCENTRANK.EXC",
	"_xlfn.PERCENTRANK.INC": "PERCENTRANK.INC",
	"_xlfn.PERMUTATIONA": "PERMUTATIONA",
	"_xlfn.PHI": "PHI",
	"_xlfn.POISSON.DIST": "POISSON.DIST",
	"_xlfn.QUARTILE.EXC": "QUARTILE.EXC",
	"_xlfn.QUARTILE.INC": "QUARTILE.INC",
	"_xlfn.QUERYSTRING": "QUERYSTRING",
	"_xlfn.RANK.AVG": "RANK.AVG",
	"_xlfn.RANK.EQ": "RANK.EQ",
	"_xlfn.RRI": "RRI",
	"_xlfn.SEC": "SEC",
	"_xlfn.SECH": "SECH",
	"_xlfn.SHEET": "SHEET",
	"_xlfn.SHEETS": "SHEETS",
	"_xlfn.SKEW.P": "SKEW.P",
	"_xlfn.STDEV.P": "STDEV.P",
	"_xlfn.STDEV.S": "STDEV.S",
	"_xlfn.SUMIFS": "SUMIFS",
	"_xlfn.SWITCH": "SWITCH",
	"_xlfn.T.DIST": "T.DIST",
	"_xlfn.T.DIST.2T": "T.DIST.2T",
	"_xlfn.T.DIST.RT": "T.DIST.RT",
	"_xlfn.T.INV": "T.INV",
	"_xlfn.T.INV.2T": "T.INV.2T",
	"_xlfn.T.TEST": "T.TEST",
	"_xlfn.TEXTJOIN": "TEXTJOIN",
	"_xlfn.UNICHAR": "UNICHAR",
	"_xlfn.UNICODE": "UNICODE",
	"_xlfn.VAR.P": "VAR.P",
	"_xlfn.VAR.S": "VAR.S",
	"_xlfn.WEBSERVICE": "WEBSERVICE",
	"_xlfn.WEIBULL.DIST": "WEIBULL.DIST",
	"_xlfn.WORKDAY.INTL": "WORKDAY.INTL",
	"_xlfn.XOR": "XOR",
	"_xlfn.Z.TEST": "Z.TEST"
};

/* Part 3 TODO: actually parse formulae */
function ods_to_csf_formula(f) {
	if(f.slice(0,3) == "of:") f = f.slice(3);
	/* 5.2 Basic Expressions */
	if(f.charCodeAt(0) == 61) {
		f = f.slice(1);
		if(f.charCodeAt(0) == 61) f = f.slice(1);
	}
	f = f.replace(/COM\.MICROSOFT\./g, "");
	/* Part 3 Section 5.8 References */
	f = f.replace(/\[((?:\.[A-Z]+[0-9]+)(?::\.[A-Z]+[0-9]+)?)\]/g, function($$, $1) { return $1.replace(/\./g,""); });
	/* TODO: something other than this */
	f = f.replace(/\[.(#[A-Z]*[?!])\]/g, "$1");
	return f.replace(/[;~]/g,",").replace(/\|/g,";");
}

function csf_to_ods_formula(f) {
	var o = "of:=" + f.replace(crefregex, "$1[.$2$3$4$5]").replace(/\]:\[/g,":");
	/* TODO: something other than this */
	return o.replace(/;/g, "|").replace(/,/g,";");
}

function ods_to_csf_3D(r) {
	var a = r.split(":");
	var s = a[0].split(".")[0];
	return [s, a[0].split(".")[1] + (a.length > 1 ? (":" + (a[1].split(".")[1] || a[1].split(".")[0])) : "")];
}

function csf_to_ods_3D(r) {
	return r.replace(/\./,"!");
}

var strs = {}; // shared strings
var _ssfopts = {}; // spreadsheet formatting options

RELS.WS = [
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
	"http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet"
];

/*global Map */
var browser_has_Map = typeof Map !== 'undefined';

function get_sst_id(sst, str, rev) {
	var i = 0, len = sst.length;
	if(rev) {
		if(browser_has_Map ? rev.has(str) : rev.hasOwnProperty(str)) {
			var revarr = browser_has_Map ? rev.get(str) : rev[str];
			for(; i < revarr.length; ++i) {
				if(sst[revarr[i]].t === str) { sst.Count ++; return revarr[i]; }
			}
		}
	} else for(; i < len; ++i) {
		if(sst[i].t === str) { sst.Count ++; return i; }
	}
	sst[len] = ({t:str}); sst.Count ++; sst.Unique ++;
	if(rev) {
		if(browser_has_Map) {
			if(!rev.has(str)) rev.set(str, []);
			rev.get(str).push(len);
		} else {
			if(!rev.hasOwnProperty(str)) rev[str] = [];
			rev[str].push(len);
		}
	}
	return len;
}

function col_obj_w(C, col) {
	var p = ({min:C+1,max:C+1});
	/* wch (chars), wpx (pixels) */
	var wch = -1;
	if(col.MDW) MDW = col.MDW;
	if(col.width != null) p.customWidth = 1;
	else if(col.wpx != null) wch = px2char(col.wpx);
	else if(col.wch != null) wch = col.wch;
	if(wch > -1) { p.width = char2width(wch); p.customWidth = 1; }
	else if(col.width != null) p.width = col.width;
	if(col.hidden) p.hidden = true;
	return p;
}

function default_margins(margins, mode) {
	if(!margins) return;
	var defs = [0.7, 0.7, 0.75, 0.75, 0.3, 0.3];
	if(mode == 'xlml') defs = [1, 1, 1, 1, 0.5, 0.5];
	if(margins.left   == null) margins.left   = defs[0];
	if(margins.right  == null) margins.right  = defs[1];
	if(margins.top    == null) margins.top    = defs[2];
	if(margins.bottom == null) margins.bottom = defs[3];
	if(margins.header == null) margins.header = defs[4];
	if(margins.footer == null) margins.footer = defs[5];
}

function get_cell_style(styles, cell, opts) {
	var z = opts.revssf[cell.z != null ? cell.z : "General"];
	var i = 0x3c, len = styles.length;
	if(z == null && opts.ssf) {
		for(; i < 0x188; ++i) if(opts.ssf[i] == null) {
			SSF.load(cell.z, i);
			// $FlowIgnore
			opts.ssf[i] = cell.z;
			opts.revssf[cell.z] = z = i;
			break;
		}
	}
	for(i = 0; i != len; ++i) if(styles[i].numFmtId === z) return i;
	styles[len] = {
		numFmtId:z,
		fontId:0,
		fillId:0,
		borderId:0,
		xfId:0,
		applyNumberFormat:1
	};
	return len;
}

function safe_format(p, fmtid, fillid, opts, themes, styles) {
	if(p.t === 'z') return;
	if(p.t === 'd' && typeof p.v === 'string') p.v = parseDate(p.v);
	try {
		if(opts.cellNF) p.z = SSF._table[fmtid];
	} catch(e) { if(opts.WTF) throw e; }
	if(!opts || opts.cellText !== false) try {
		if(SSF._table[fmtid] == null) SSF.load(SSFImplicit[fmtid] || "General", fmtid);
		if(p.t === 'e') p.w = p.w || BErr[p.v];
		else if(fmtid === 0) {
			if(p.t === 'n') {
				if((p.v|0) === p.v) p.w = SSF._general_int(p.v);
				else p.w = SSF._general_num(p.v);
			}
			else if(p.t === 'd') {
				var dd = datenum(p.v);
				if((dd|0) === dd) p.w = SSF._general_int(dd);
				else p.w = SSF._general_num(dd);
			}
			else if(p.v === undefined) return "";
			else p.w = SSF._general(p.v,_ssfopts);
		}
		else if(p.t === 'd') p.w = SSF.format(fmtid,datenum(p.v),_ssfopts);
		else p.w = SSF.format(fmtid,p.v,_ssfopts);
	} catch(e) { if(opts.WTF) throw e; }
	if(!opts.cellStyles) return;
	if(fillid != null) try {
		p.s = styles.Fills[fillid];
		if (p.s.fgColor && p.s.fgColor.theme && !p.s.fgColor.rgb) {
			p.s.fgColor.rgb = rgb_tint(themes.themeElements.clrScheme[p.s.fgColor.theme].rgb, p.s.fgColor.tint || 0);
			if(opts.WTF) p.s.fgColor.raw_rgb = themes.themeElements.clrScheme[p.s.fgColor.theme].rgb;
		}
		if (p.s.bgColor && p.s.bgColor.theme) {
			p.s.bgColor.rgb = rgb_tint(themes.themeElements.clrScheme[p.s.bgColor.theme].rgb, p.s.bgColor.tint || 0);
			if(opts.WTF) p.s.bgColor.raw_rgb = themes.themeElements.clrScheme[p.s.bgColor.theme].rgb;
		}
	} catch(e) { if(opts.WTF && styles.Fills) throw e; }
}

function check_ws(ws, sname, i) {
	if(ws && ws['!ref']) {
		var range = safe_decode_range(ws['!ref']);
		if(range.e.c < range.s.c || range.e.r < range.s.r) throw new Error("Bad range (" + i + "): " + ws['!ref']);
	}
}
function parse_ws_xml_dim(ws, s) {
	var d = safe_decode_range(s);
	if(d.s.r<=d.e.r && d.s.c<=d.e.c && d.s.r>=0 && d.s.c>=0) ws["!ref"] = encode_range(d);
}
var mergecregex = /<(?:\w:)?mergeCell ref="[A-Z0-9:]+"\s*[\/]?>/g;
var sheetdataregex = /<(?:\w+:)?sheetData>([\s\S]*)<\/(?:\w+:)?sheetData>/;
var hlinkregex = /<(?:\w:)?hyperlink [^>]*>/mg;
var dimregex = /"(\w*:\w*)"/;
var colregex = /<(?:\w:)?col\b[^>]*[\/]?>/g;
var afregex = /<(?:\w:)?autoFilter[^>]*([\/]|>([\s\S]*)<\/(?:\w:)?autoFilter)>/g;
var marginregex= /<(?:\w:)?pageMargins[^>]*\/>/g;
var sheetprregex = /<(?:\w:)?sheetPr(?:[^>a-z][^>]*)?\/>/;
var svsregex = /<(?:\w:)?sheetViews[^>]*(?:[\/]|>([\s\S]*)<\/(?:\w:)?sheetViews)>/;
/* 18.3 Worksheets */
function parse_ws_xml(data, opts, idx, rels, wb, themes, styles) {
	if(!data) return data;
	if(DENSE != null && opts.dense == null) opts.dense = DENSE;

	/* 18.3.1.99 worksheet CT_Worksheet */
	var s = opts.dense ? ([]) : ({});
	var refguess = ({s: {r:2000000, c:2000000}, e: {r:0, c:0} });

	var data1 = "", data2 = "";
	var mtch = data.match(sheetdataregex);
	if(mtch) {
		data1 = data.slice(0, mtch.index);
		data2 = data.slice(mtch.index + mtch[0].length);
	} else data1 = data2 = data;

	/* 18.3.1.82 sheetPr CT_SheetPr */
	var sheetPr = data1.match(sheetprregex);
	if(sheetPr) parse_ws_xml_sheetpr(sheetPr[0], s, wb, idx);

	/* 18.3.1.35 dimension CT_SheetDimension */
	// $FlowIgnore
	var ridx = (data1.match(/<(?:\w*:)?dimension/)||{index:-1}).index;
	if(ridx > 0) {
		var ref = data1.slice(ridx,ridx+50).match(dimregex);
		if(ref) parse_ws_xml_dim(s, ref[1]);
	}

	/* 18.3.1.88 sheetViews CT_SheetViews */
	var svs = data1.match(svsregex);
	if(svs && svs[1]) parse_ws_xml_sheetviews(svs[1], wb);

	/* 18.3.1.17 cols CT_Cols */
	var columns = [];
	if(opts.cellStyles) {
		/* 18.3.1.13 col CT_Col */
		var cols = data1.match(colregex);
		if(cols) parse_ws_xml_cols(columns, cols);
	}

	/* 18.3.1.80 sheetData CT_SheetData ? */
	if(mtch) parse_ws_xml_data(mtch[1], s, opts, refguess, themes, styles);

	/* 18.3.1.2  autoFilter CT_AutoFilter */
	var afilter = data2.match(afregex);
	if(afilter) s['!autofilter'] = parse_ws_xml_autofilter(afilter[0]);

	/* 18.3.1.55 mergeCells CT_MergeCells */
	var merges = [];
	var _merge = data2.match(mergecregex);
	if(_merge) for(ridx = 0; ridx != _merge.length; ++ridx)
		merges[ridx] = safe_decode_range(_merge[ridx].slice(_merge[ridx].indexOf("\"")+1));

	/* 18.3.1.48 hyperlinks CT_Hyperlinks */
	var hlink = data2.match(hlinkregex);
	if(hlink) parse_ws_xml_hlinks(s, hlink, rels);

	/* 18.3.1.62 pageMargins CT_PageMargins */
	var margins = data2.match(marginregex);
	if(margins) s['!margins'] = parse_ws_xml_margins(parsexmltag(margins[0]));

	if(!s["!ref"] && refguess.e.c >= refguess.s.c && refguess.e.r >= refguess.s.r) s["!ref"] = encode_range(refguess);
	if(opts.sheetRows > 0 && s["!ref"]) {
		var tmpref = safe_decode_range(s["!ref"]);
		if(opts.sheetRows <= +tmpref.e.r) {
			tmpref.e.r = opts.sheetRows - 1;
			if(tmpref.e.r > refguess.e.r) tmpref.e.r = refguess.e.r;
			if(tmpref.e.r < tmpref.s.r) tmpref.s.r = tmpref.e.r;
			if(tmpref.e.c > refguess.e.c) tmpref.e.c = refguess.e.c;
			if(tmpref.e.c < tmpref.s.c) tmpref.s.c = tmpref.e.c;
			s["!fullref"] = s["!ref"];
			s["!ref"] = encode_range(tmpref);
		}
	}
	if(columns.length > 0) s["!cols"] = columns;
	if(merges.length > 0) s["!merges"] = merges;
	return s;
}

function write_ws_xml_merges(merges) {
	if(merges.length === 0) return "";
	var o = '<mergeCells count="' + merges.length + '">';
	for(var i = 0; i != merges.length; ++i) o += '<mergeCell ref="' + encode_range(merges[i]) + '"/>';
	return o + '</mergeCells>';
}

/* 18.3.1.82-3 sheetPr CT_ChartsheetPr / CT_SheetPr */
function parse_ws_xml_sheetpr(sheetPr, s, wb, idx) {
	var data = parsexmltag(sheetPr);
	if(!wb.Sheets[idx]) wb.Sheets[idx] = {};
	if(data.codeName) wb.Sheets[idx].CodeName = data.codeName;
}

/* 18.3.1.85 sheetProtection CT_SheetProtection */
function write_ws_xml_protection(sp) {
	// algorithmName, hashValue, saltValue, spinCountpassword
	var o = ({sheet:1});
	var deffalse = ["objects", "scenarios", "selectLockedCells", "selectUnlockedCells"];
	var deftrue = [
		"formatColumns", "formatRows", "formatCells",
		"insertColumns", "insertRows", "insertHyperlinks",
		"deleteColumns", "deleteRows",
		"sort", "autoFilter", "pivotTables"
	];
	deffalse.forEach(function(n) { if(sp[n] != null && sp[n]) o[n] = "1"; });
	deftrue.forEach(function(n) { if(sp[n] != null && !sp[n]) o[n] = "0"; });
	/* TODO: algorithm */
	if(sp.password) o.password = crypto_CreatePasswordVerifier_Method1(sp.password).toString(16).toUpperCase();
	return writextag('sheetProtection', null, o);
}

function parse_ws_xml_hlinks(s, data, rels) {
	var dense = Array.isArray(s);
	for(var i = 0; i != data.length; ++i) {
		var val = parsexmltag(utf8read(data[i]), true);
		if(!val.ref) return;
		var rel = ((rels || {})['!id']||[])[val.id];
		if(rel) {
			val.Target = rel.Target;
			if(val.location) val.Target += "#"+val.location;
		} else {
			val.Target = "#" + val.location;
			rel = {Target: val.Target, TargetMode: 'Internal'};
		}
		val.Rel = rel;
		if(val.tooltip) { val.Tooltip = val.tooltip; delete val.tooltip; }
		var rng = safe_decode_range(val.ref);
		for(var R=rng.s.r;R<=rng.e.r;++R) for(var C=rng.s.c;C<=rng.e.c;++C) {
			var addr = encode_cell({c:C,r:R});
			if(dense) {
				if(!s[R]) s[R] = [];
				if(!s[R][C]) s[R][C] = {t:"z",v:undefined};
				s[R][C].l = val;
			} else {
				if(!s[addr]) s[addr] = {t:"z",v:undefined};
				s[addr].l = val;
			}
		}
	}
}

function parse_ws_xml_margins(margin) {
	var o = {};
	["left", "right", "top", "bottom", "header", "footer"].forEach(function(k) {
		if(margin[k]) o[k] = parseFloat(margin[k]);
	});
	return o;
}
function write_ws_xml_margins(margin) {
	default_margins(margin);
	return writextag('pageMargins', null, margin);
}

function parse_ws_xml_cols(columns, cols) {
	var seencol = false;
	for(var coli = 0; coli != cols.length; ++coli) {
		var coll = parsexmltag(cols[coli], true);
		if(coll.hidden) coll.hidden = parsexmlbool(coll.hidden);
		var colm=parseInt(coll.min, 10)-1, colM=parseInt(coll.max,10)-1;
		delete coll.min; delete coll.max; coll.width = +coll.width;
		if(!seencol && coll.width) { seencol = true; find_mdw_colw(coll.width); }
		process_col(coll);
		while(colm <= colM) columns[colm++] = dup(coll);
	}
}

function write_ws_xml_cols(ws, cols) {
	var o = ["<cols>"], col;
	for(var i = 0; i != cols.length; ++i) {
		if(!(col = cols[i])) continue;
		o[o.length] = (writextag('col', null, col_obj_w(i, col)));
	}
	o[o.length] = "</cols>";
	return o.join("");
}

function parse_ws_xml_autofilter(data) {
	var o = { ref: (data.match(/ref="([^"]*)"/)||[])[1]};
	return o;
}
function write_ws_xml_autofilter(data, ws, wb, idx) {
	var ref = typeof data.ref == "string" ? data.ref : encode_range(data.ref);
	if(!wb.Workbook) wb.Workbook = {};
	if(!wb.Workbook.Names) wb.Workbook.Names = [];
	var names = wb.Workbook.Names;
	var range = decode_range(ref);
	if(range.s.r == range.e.r) { range.e.r = decode_range(ws["!ref"]).e.r; ref = encode_range(range); }
	for(var i = 0; i < names.length; ++i) {
		var name = names[i];
		if(name.Name != '_xlnm._FilterDatabase') continue;
		if(name.Sheet != idx) continue;
		name.Ref = "'" + wb.SheetNames[idx] + "'!" + ref; break;
	}
	if(i == names.length) names.push({ Name: '_xlnm._FilterDatabase', Sheet: idx, Ref: "'" + wb.SheetNames[idx] + "'!" + ref  });
	return writextag("autoFilter", null, {ref:ref});
}

/* 18.3.1.88 sheetViews CT_SheetViews */
/* 18.3.1.87 sheetView CT_SheetView */
var sviewregex = /<(?:\w:)?sheetView(?:[^>a-z][^>]*)?\/>/;
function parse_ws_xml_sheetviews(data, wb) {
	(data.match(sviewregex)||[]).forEach(function(r) {
		var tag = parsexmltag(r);
		if(parsexmlbool(tag.rightToLeft)) {
			if(!wb.Views) wb.Views = [{}];
			if(!wb.Views[0]) wb.Views[0] = {};
			wb.Views[0].RTL = true;
		}
	});
}
function write_ws_xml_sheetviews(ws, opts, idx, wb) {
	var sview = {workbookViewId:"0"};
	// $FlowIgnore
	if( (((wb||{}).Workbook||{}).Views||[])[0] ) sview.rightToLeft = wb.Workbook.Views[0].RTL ? "1" : "0";
	return writextag("sheetViews", writextag("sheetView", null, sview), {});
}

function write_ws_xml_cell(cell, ref, ws, opts) {
	if(cell.v === undefined && cell.f === undefined || cell.t === 'z') return "";
	var vv = "";
	var oldt = cell.t, oldv = cell.v;
	switch(cell.t) {
		case 'b': vv = cell.v ? "1" : "0"; break;
		case 'n': vv = ''+cell.v; break;
		case 'e': vv = BErr[cell.v]; break;
		case 'd':
			if(opts.cellDates) vv = parseDate(cell.v, -1).toISOString();
			else {
				cell = dup(cell);
				cell.t = 'n';
				vv = ''+(cell.v = datenum(parseDate(cell.v)));
			}
			if(typeof cell.z === 'undefined') cell.z = SSF._table[14];
			break;
		default: vv = cell.v; break;
	}
	var v = writetag('v', escapexml(vv)), o = ({r:ref});
	/* TODO: cell style */
	var os = get_cell_style(opts.cellXfs, cell, opts);
	if(os !== 0) o.s = os;
	switch(cell.t) {
		case 'n': break;
		case 'd': o.t = "d"; break;
		case 'b': o.t = "b"; break;
		case 'e': o.t = "e"; break;
		default: if(cell.v == null) { delete cell.t; break; }
			if(opts.bookSST) {
				v = writetag('v', ''+get_sst_id(opts.Strings, cell.v, opts.revStrings));
				o.t = "s"; break;
			}
			o.t = "str"; break;
	}
	if(cell.t != oldt) { cell.t = oldt; cell.v = oldv; }
	if(cell.f) {
		var ff = cell.F && cell.F.slice(0, ref.length) == ref ? {t:"array", ref:cell.F} : null;
		v = writextag('f', escapexml(cell.f), ff) + (cell.v != null ? v : "");
	}
	if(cell.l) ws['!links'].push([ref, cell.l]);
	if(cell.c) ws['!comments'].push([ref, cell.c]);
	return writextag('c', v, o);
}

var parse_ws_xml_data = (function() {
	var cellregex = /<(?:\w+:)?c[ >]/, rowregex = /<\/(?:\w+:)?row>/;
	var rregex = /r=["']([^"']*)["']/, isregex = /<(?:\w+:)?is>([\S\s]*?)<\/(?:\w+:)?is>/;
	var refregex = /ref=["']([^"']*)["']/;
	var match_v = matchtag("v"), match_f = matchtag("f");

return function parse_ws_xml_data(sdata, s, opts, guess, themes, styles) {
	var ri = 0, x = "", cells = [], cref = [], idx=0, i=0, cc=0, d="", p;
	var tag, tagr = 0, tagc = 0;
	var sstr, ftag;
	var fmtid = 0, fillid = 0;
	var do_format = Array.isArray(styles.CellXf), cf;
	var arrayf = [];
	var sharedf = [];
	var dense = Array.isArray(s);
	var rows = [], rowobj = {}, rowrite = false;
	for(var marr = sdata.split(rowregex), mt = 0, marrlen = marr.length; mt != marrlen; ++mt) {
		x = marr[mt].trim();
		var xlen = x.length;
		if(xlen === 0) continue;

		/* 18.3.1.73 row CT_Row */
		for(ri = 0; ri < xlen; ++ri) if(x.charCodeAt(ri) === 62) break; ++ri;
		tag = parsexmltag(x.slice(0,ri), true);
		tagr = tag.r != null ? parseInt(tag.r, 10) : tagr+1; tagc = -1;
		if(opts.sheetRows && opts.sheetRows < tagr) continue;
		if(guess.s.r > tagr - 1) guess.s.r = tagr - 1;
		if(guess.e.r < tagr - 1) guess.e.r = tagr - 1;

		if(opts && opts.cellStyles) {
			rowobj = {}; rowrite = false;
			if(tag.ht) { rowrite = true; rowobj.hpt = parseFloat(tag.ht); rowobj.hpx = pt2px(rowobj.hpt); }
			if(tag.hidden == "1") { rowrite = true; rowobj.hidden = true; }
			if(tag.outlineLevel != null) { rowrite = true; rowobj.level = +tag.outlineLevel; }
			if(rowrite) rows[tagr-1] = rowobj;
		}

		/* 18.3.1.4 c CT_Cell */
		cells = x.slice(ri).split(cellregex);
		for(ri = 0; ri != cells.length; ++ri) {
			x = cells[ri].trim();
			if(x.length === 0) continue;
			cref = x.match(rregex); idx = ri; i=0; cc=0;
			x = "<c " + (x.slice(0,1)=="<"?">":"") + x;
			if(cref != null && cref.length === 2) {
				idx = 0; d=cref[1];
				for(i=0; i != d.length; ++i) {
					if((cc=d.charCodeAt(i)-64) < 1 || cc > 26) break;
					idx = 26*idx + cc;
				}
				--idx;
				tagc = idx;
			} else ++tagc;
			for(i = 0; i != x.length; ++i) if(x.charCodeAt(i) === 62) break; ++i;
			tag = parsexmltag(x.slice(0,i), true);
			if(!tag.r) tag.r = encode_cell({r:tagr-1, c:tagc});
			d = x.slice(i);
			p = ({t:""});

			if((cref=d.match(match_v))!= null && cref[1] !== '') p.v=unescapexml(cref[1]);
			if(opts.cellFormula) {
				if((cref=d.match(match_f))!= null && cref[1] !== '') {
					/* TODO: match against XLSXFutureFunctions */
					p.f=_xlfn(unescapexml(utf8read(cref[1])));
					if(cref[0].indexOf('t="array"') > -1) {
						p.F = (d.match(refregex)||[])[1];
						if(p.F.indexOf(":") > -1) arrayf.push([safe_decode_range(p.F), p.F]);
					} else if(cref[0].indexOf('t="shared"') > -1) {
						// TODO: parse formula
						ftag = parsexmltag(cref[0]);
						sharedf[parseInt(ftag.si, 10)] = [ftag, _xlfn(unescapexml(utf8read(cref[1])))];
					}
				} else if((cref=d.match(/<f[^>]*\/>/))) {
					ftag = parsexmltag(cref[0]);
					if(sharedf[ftag.si]) p.f = shift_formula_xlsx(sharedf[ftag.si][1], sharedf[ftag.si][0].ref, tag.r);
				}
				/* TODO: factor out contains logic */
				var _tag = decode_cell(tag.r);
				for(i = 0; i < arrayf.length; ++i)
					if(_tag.r >= arrayf[i][0].s.r && _tag.r <= arrayf[i][0].e.r)
						if(_tag.c >= arrayf[i][0].s.c && _tag.c <= arrayf[i][0].e.c)
							p.F = arrayf[i][1];
			}

			if(tag.t == null && p.v === undefined) {
				if(p.f || p.F) {
					p.v = 0; p.t = "n";
				} else if(!opts.sheetStubs) continue;
				else p.t = "z";
			}
			else p.t = tag.t || "n";
			if(guess.s.c > tagc) guess.s.c = tagc;
			if(guess.e.c < tagc) guess.e.c = tagc;
			/* 18.18.11 t ST_CellType */
			switch(p.t) {
				case 'n':
					if(p.v == "" || p.v == null) {
						if(!opts.sheetStubs) continue;
						p.t = 'z';
					} else p.v = parseFloat(p.v);
					break;
				case 's':
					if(typeof p.v == 'undefined') {
						if(!opts.sheetStubs) continue;
						p.t = 'z';
					} else {
						sstr = strs[parseInt(p.v, 10)];
						p.v = sstr.t;
						p.r = sstr.r;
						if(opts.cellHTML) p.h = sstr.h;
					}
					break;
				case 'str':
					p.t = "s";
					p.v = (p.v!=null) ? utf8read(p.v) : '';
					if(opts.cellHTML) p.h = escapehtml(p.v);
					break;
				case 'inlineStr':
					cref = d.match(isregex);
					p.t = 's';
					if(cref != null && (sstr = parse_si(cref[1]))) p.v = sstr.t; else p.v = "";
					break;
				case 'b': p.v = parsexmlbool(p.v); break;
				case 'd':
					if(opts.cellDates) p.v = parseDate(p.v, 1);
					else { p.v = datenum(parseDate(p.v, 1)); p.t = 'n'; }
					break;
				/* error string in .w, number in .v */
				case 'e':
					if(!opts || opts.cellText !== false) p.w = p.v;
					p.v = RBErr[p.v]; break;
			}
			/* formatting */
			fmtid = fillid = 0;
			if(do_format && tag.s !== undefined) {
				cf = styles.CellXf[tag.s];
				if(cf != null) {
					if(cf.numFmtId != null) fmtid = cf.numFmtId;
					if(opts.cellStyles) {
						if(cf.fillId != null) fillid = cf.fillId;
					}
				}
			}
			safe_format(p, fmtid, fillid, opts, themes, styles);
			if(opts.cellDates && do_format && p.t == 'n' && SSF.is_date(SSF._table[fmtid])) { p.t = 'd'; p.v = numdate(p.v); }
			if(dense) {
				var _r = decode_cell(tag.r);
				if(!s[_r.r]) s[_r.r] = [];
				s[_r.r][_r.c] = p;
			} else s[tag.r] = p;
		}
	}
	if(rows.length > 0) s['!rows'] = rows;
}; })();

function write_ws_xml_data(ws, opts, idx, wb) {
	var o = [], r = [], range = safe_decode_range(ws['!ref']), cell="", ref, rr = "", cols = [], R=0, C=0, rows = ws['!rows'];
	var dense = Array.isArray(ws);
	var params = ({r:rr}), row, height = -1;
	for(C = range.s.c; C <= range.e.c; ++C) cols[C] = encode_col(C);
	for(R = range.s.r; R <= range.e.r; ++R) {
		r = [];
		rr = encode_row(R);
		for(C = range.s.c; C <= range.e.c; ++C) {
			ref = cols[C] + rr;
			var _cell = dense ? (ws[R]||[])[C]: ws[ref];
			if(_cell === undefined) continue;
			if((cell = write_ws_xml_cell(_cell, ref, ws, opts, idx, wb)) != null) r.push(cell);
		}
		if(r.length > 0 || (rows && rows[R])) {
			params = ({r:rr});
			if(rows && rows[R]) {
				row = rows[R];
				if(row.hidden) params.hidden = 1;
				height = -1;
				if(row.hpx) height = px2pt(row.hpx);
				else if(row.hpt) height = row.hpt;
				if(height > -1) { params.ht = height; params.customHeight = 1; }
				if(row.level) { params.outlineLevel = row.level; }
			}
			o[o.length] = (writextag('row', r.join(""), params));
		}
	}
	if(rows) for(; R < rows.length; ++R) {
		if(rows && rows[R]) {
			params = ({r:R+1});
			row = rows[R];
			if(row.hidden) params.hidden = 1;
			height = -1;
			if (row.hpx) height = px2pt(row.hpx);
			else if (row.hpt) height = row.hpt;
			if (height > -1) { params.ht = height; params.customHeight = 1; }
			if (row.level) { params.outlineLevel = row.level; }
			o[o.length] = (writextag('row', "", params));
		}
	}
	return o.join("");
}

var WS_XML_ROOT = writextag('worksheet', null, {
	'xmlns': XMLNS.main[0],
	'xmlns:r': XMLNS.r
});

function write_ws_xml(idx, opts, wb, rels) {
	var o = [XML_HEADER, WS_XML_ROOT];
	var s = wb.SheetNames[idx], sidx = 0, rdata = "";
	var ws = wb.Sheets[s];
	if(ws == null) ws = {};
	var ref = ws['!ref'] || 'A1';
	var range = safe_decode_range(ref);
	if(range.e.c > 0x3FFF || range.e.r > 0xFFFFF) {
		if(opts.WTF) throw new Error("Range " + ref + " exceeds format limit A1:XFD1048576");
		range.e.c = Math.min(range.e.c, 0x3FFF);
		range.e.r = Math.min(range.e.c, 0xFFFFF);
		ref = encode_range(range);
	}
	if(!rels) rels = {};
	ws['!comments'] = [];
	ws['!drawing'] = [];

	if(opts.bookType !== 'xlsx' && wb.vbaraw) {
		var cname = wb.SheetNames[idx];
		try { if(wb.Workbook) cname = wb.Workbook.Sheets[idx].CodeName || cname; } catch(e) {}
		o[o.length] = (writextag('sheetPr', null, {'codeName': escapexml(cname)}));
	}

	o[o.length] = (writextag('dimension', null, {'ref': ref}));

	o[o.length] = write_ws_xml_sheetviews(ws, opts, idx, wb);

	/* TODO: store in WB, process styles */
	if(opts.sheetFormat) o[o.length] = (writextag('sheetFormatPr', null, {
		defaultRowHeight:opts.sheetFormat.defaultRowHeight||'16',
		baseColWidth:opts.sheetFormat.baseColWidth||'10',
		outlineLevelRow:opts.sheetFormat.outlineLevelRow||'7'
	}));

	if(ws['!cols'] != null && ws['!cols'].length > 0) o[o.length] = (write_ws_xml_cols(ws, ws['!cols']));

	o[sidx = o.length] = '<sheetData/>';
	ws['!links'] = [];
	if(ws['!ref'] != null) {
		rdata = write_ws_xml_data(ws, opts, idx, wb, rels);
		if(rdata.length > 0) o[o.length] = (rdata);
	}
	if(o.length>sidx+1) { o[o.length] = ('</sheetData>'); o[sidx]=o[sidx].replace("/>",">"); }

	/* sheetCalcPr */

	if(ws['!protect'] != null) o[o.length] = write_ws_xml_protection(ws['!protect']);

	/* protectedRanges */
	/* scenarios */

	if(ws['!autofilter'] != null) o[o.length] = write_ws_xml_autofilter(ws['!autofilter'], ws, wb, idx);

	/* sortState */
	/* dataConsolidate */
	/* customSheetViews */

	if(ws['!merges'] != null && ws['!merges'].length > 0) o[o.length] = (write_ws_xml_merges(ws['!merges']));

	/* phoneticPr */
	/* conditionalFormatting */
	/* dataValidations */

	var relc = -1, rel, rId = -1;
	if(ws['!links'].length > 0) {
		o[o.length] = "<hyperlinks>";
		ws['!links'].forEach(function(l) {
			if(!l[1].Target) return;
			rel = ({"ref":l[0]});
			if(l[1].Target.charAt(0) != "#") {
				rId = add_rels(rels, -1, escapexml(l[1].Target).replace(/#.*$/, ""), RELS.HLINK);
				rel["r:id"] = "rId"+rId;
			}
			if((relc = l[1].Target.indexOf("#")) > -1) rel.location = escapexml(l[1].Target.slice(relc+1));
			if(l[1].Tooltip) rel.tooltip = escapexml(l[1].Tooltip);
			o[o.length] = writextag("hyperlink",null,rel);
		});
		o[o.length] = "</hyperlinks>";
	}
	delete ws['!links'];

	/* printOptions */
	if (ws['!margins'] != null) o[o.length] =  write_ws_xml_margins(ws['!margins']);
	/* pageSetup */

	//var hfidx = o.length;
	o[o.length] = "";

	/* rowBreaks */
	/* colBreaks */
	/* customProperties */
	/* cellWatches */

	if(!opts || opts.ignoreEC || (opts.ignoreEC == (void 0))) o[o.length] = writetag("ignoredErrors", writextag("ignoredError", null, {numberStoredAsText:1, sqref:ref}));

	/* smartTags */

	if(ws['!drawing'].length > 0) {
		rId = add_rels(rels, -1, "../drawings/drawing" + (idx+1) + ".xml", RELS.DRAW);
		o[o.length] = writextag("drawing", null, {"r:id":"rId" + rId});
	}
	else delete ws['!drawing'];

	if(ws['!comments'].length > 0) {
		rId = add_rels(rels, -1, "../drawings/vmlDrawing" + (idx+1) + ".vml", RELS.VML);
		o[o.length] = writextag("legacyDrawing", null, {"r:id":"rId" + rId});
		ws['!legacy'] = rId;
	}

	/* drawingHF */
	/* picture */
	/* oleObjects */
	/* controls */
	/* webPublishItems */
	/* tableParts */
	/* extList */

	if(o.length>2) { o[o.length] = ('</worksheet>'); o[1]=o[1].replace("/>",">"); }
	return o.join("");
}

/* [MS-XLSB] 2.4.726 BrtRowHdr */
function parse_BrtRowHdr(data, length) {
	var z = ({});
	var tgt = data.l + length;
	z.r = data.read_shift(4);
	data.l += 4; // TODO: ixfe
	var miyRw = data.read_shift(2);
	data.l += 1; // TODO: top/bot padding
	var flags = data.read_shift(1);
	data.l = tgt;
	if(flags & 0x07) z.level = flags & 0x07;
	if(flags & 0x10) z.hidden = true;
	if(flags & 0x20) z.hpt = miyRw / 20;
	return z;
}
function write_BrtRowHdr(R, range, ws) {
	var o = new_buf(17+8*16);
	var row = (ws['!rows']||[])[R]||{};
	o.write_shift(4, R);

	o.write_shift(4, 0); /* TODO: ixfe */

	var miyRw = 0x0140;
	if(row.hpx) miyRw = px2pt(row.hpx) * 20;
	else if(row.hpt) miyRw = row.hpt * 20;
	o.write_shift(2, miyRw);

	o.write_shift(1, 0); /* top/bot padding */

	var flags = 0x0;
	if(row.level) flags |= row.level;
	if(row.hidden) flags |= 0x10;
	if(row.hpx || row.hpt) flags |= 0x20;
	o.write_shift(1, flags);

	o.write_shift(1, 0); /* phonetic guide */

	/* [MS-XLSB] 2.5.8 BrtColSpan explains the mechanism */
	var ncolspan = 0, lcs = o.l;
	o.l += 4;

	var caddr = {r:R, c:0};
	for(var i = 0; i < 16; ++i) {
		if((range.s.c > ((i+1) << 10)) || (range.e.c < (i << 10))) continue;
		var first = -1, last = -1;
		for(var j = (i<<10); j < ((i+1)<<10); ++j) {
			caddr.c = j;
			var cell = Array.isArray(ws) ? (ws[caddr.r]||[])[caddr.c] : ws[encode_cell(caddr)];
			if(cell) { if(first < 0) first = j; last = j; }
		}
		if(first < 0) continue;
		++ncolspan;
		o.write_shift(4, first);
		o.write_shift(4, last);
	}

	var l = o.l;
	o.l = lcs;
	o.write_shift(4, ncolspan);
	o.l = l;

	return o.length > o.l ? o.slice(0, o.l) : o;
}
function write_row_header(ba, ws, range, R) {
	var o = write_BrtRowHdr(R, range, ws);
	if((o.length > 17) || (ws['!rows']||[])[R]) write_record(ba, 'BrtRowHdr', o);
}

/* [MS-XLSB] 2.4.820 BrtWsDim */
var parse_BrtWsDim = parse_UncheckedRfX;
var write_BrtWsDim = write_UncheckedRfX;

/* [MS-XLSB] 2.4.821 BrtWsFmtInfo */
function parse_BrtWsFmtInfo() {
}
//function write_BrtWsFmtInfo(ws, o) { }

/* [MS-XLSB] 2.4.823 BrtWsProp */
function parse_BrtWsProp(data, length) {
	var z = {};
	/* TODO: pull flags */
	data.l += 19;
	z.name = parse_XLSBCodeName(data, length - 19);
	return z;
}
function write_BrtWsProp(str, o) {
	if(o == null) o = new_buf(84+4*str.length);
	for(var i = 0; i < 3; ++i) o.write_shift(1,0);
	write_BrtColor({auto:1}, o);
	o.write_shift(-4,-1);
	o.write_shift(-4,-1);
	write_XLSBCodeName(str, o);
	return o.slice(0, o.l);
}

/* [MS-XLSB] 2.4.306 BrtCellBlank */
function parse_BrtCellBlank(data) {
	var cell = parse_XLSBCell(data);
	return [cell];
}
function write_BrtCellBlank(cell, ncell, o) {
	if(o == null) o = new_buf(8);
	return write_XLSBCell(ncell, o);
}


/* [MS-XLSB] 2.4.307 BrtCellBool */
function parse_BrtCellBool(data) {
	var cell = parse_XLSBCell(data);
	var fBool = data.read_shift(1);
	return [cell, fBool, 'b'];
}
function write_BrtCellBool(cell, ncell, o) {
	if(o == null) o = new_buf(9);
	write_XLSBCell(ncell, o);
	o.write_shift(1, cell.v ? 1 : 0);
	return o;
}

/* [MS-XLSB] 2.4.308 BrtCellError */
function parse_BrtCellError(data) {
	var cell = parse_XLSBCell(data);
	var bError = data.read_shift(1);
	return [cell, bError, 'e'];
}

/* [MS-XLSB] 2.4.311 BrtCellIsst */
function parse_BrtCellIsst(data) {
	var cell = parse_XLSBCell(data);
	var isst = data.read_shift(4);
	return [cell, isst, 's'];
}
function write_BrtCellIsst(cell, ncell, o) {
	if(o == null) o = new_buf(12);
	write_XLSBCell(ncell, o);
	o.write_shift(4, ncell.v);
	return o;
}

/* [MS-XLSB] 2.4.313 BrtCellReal */
function parse_BrtCellReal(data) {
	var cell = parse_XLSBCell(data);
	var value = parse_Xnum(data);
	return [cell, value, 'n'];
}
function write_BrtCellReal(cell, ncell, o) {
	if(o == null) o = new_buf(16);
	write_XLSBCell(ncell, o);
	write_Xnum(cell.v, o);
	return o;
}

/* [MS-XLSB] 2.4.314 BrtCellRk */
function parse_BrtCellRk(data) {
	var cell = parse_XLSBCell(data);
	var value = parse_RkNumber(data);
	return [cell, value, 'n'];
}
function write_BrtCellRk(cell, ncell, o) {
	if(o == null) o = new_buf(12);
	write_XLSBCell(ncell, o);
	write_RkNumber(cell.v, o);
	return o;
}


/* [MS-XLSB] 2.4.317 BrtCellSt */
function parse_BrtCellSt(data) {
	var cell = parse_XLSBCell(data);
	var value = parse_XLWideString(data);
	return [cell, value, 'str'];
}
function write_BrtCellSt(cell, ncell, o) {
	if(o == null) o = new_buf(12 + 4 * cell.v.length);
	write_XLSBCell(ncell, o);
	write_XLWideString(cell.v, o);
	return o.length > o.l ? o.slice(0, o.l) : o;
}

/* [MS-XLSB] 2.4.653 BrtFmlaBool */
function parse_BrtFmlaBool(data, length, opts) {
	var end = data.l + length;
	var cell = parse_XLSBCell(data);
	cell.r = opts['!row'];
	var value = data.read_shift(1);
	var o = [cell, value, 'b'];
	if(opts.cellFormula) {
		data.l += 2;
		var formula = parse_XLSBCellParsedFormula(data, end - data.l, opts);
		o[3] = stringify_formula(formula, null/*range*/, cell, opts.supbooks, opts);/* TODO */
	}
	else data.l = end;
	return o;
}

/* [MS-XLSB] 2.4.654 BrtFmlaError */
function parse_BrtFmlaError(data, length, opts) {
	var end = data.l + length;
	var cell = parse_XLSBCell(data);
	cell.r = opts['!row'];
	var value = data.read_shift(1);
	var o = [cell, value, 'e'];
	if(opts.cellFormula) {
		data.l += 2;
		var formula = parse_XLSBCellParsedFormula(data, end - data.l, opts);
		o[3] = stringify_formula(formula, null/*range*/, cell, opts.supbooks, opts);/* TODO */
	}
	else data.l = end;
	return o;
}

/* [MS-XLSB] 2.4.655 BrtFmlaNum */
function parse_BrtFmlaNum(data, length, opts) {
	var end = data.l + length;
	var cell = parse_XLSBCell(data);
	cell.r = opts['!row'];
	var value = parse_Xnum(data);
	var o = [cell, value, 'n'];
	if(opts.cellFormula) {
		data.l += 2;
		var formula = parse_XLSBCellParsedFormula(data, end - data.l, opts);
		o[3] = stringify_formula(formula, null/*range*/, cell, opts.supbooks, opts);/* TODO */
	}
	else data.l = end;
	return o;
}

/* [MS-XLSB] 2.4.656 BrtFmlaString */
function parse_BrtFmlaString(data, length, opts) {
	var end = data.l + length;
	var cell = parse_XLSBCell(data);
	cell.r = opts['!row'];
	var value = parse_XLWideString(data);
	var o = [cell, value, 'str'];
	if(opts.cellFormula) {
		data.l += 2;
		var formula = parse_XLSBCellParsedFormula(data, end - data.l, opts);
		o[3] = stringify_formula(formula, null/*range*/, cell, opts.supbooks, opts);/* TODO */
	}
	else data.l = end;
	return o;
}

/* [MS-XLSB] 2.4.682 BrtMergeCell */
var parse_BrtMergeCell = parse_UncheckedRfX;
var write_BrtMergeCell = write_UncheckedRfX;
/* [MS-XLSB] 2.4.107 BrtBeginMergeCells */
function write_BrtBeginMergeCells(cnt, o) {
	if(o == null) o = new_buf(4);
	o.write_shift(4, cnt);
	return o;
}

/* [MS-XLSB] 2.4.662 BrtHLink */
function parse_BrtHLink(data, length) {
	var end = data.l + length;
	var rfx = parse_UncheckedRfX(data, 16);
	var relId = parse_XLNullableWideString(data);
	var loc = parse_XLWideString(data);
	var tooltip = parse_XLWideString(data);
	var display = parse_XLWideString(data);
	data.l = end;
	var o = ({rfx:rfx, relId:relId, loc:loc, display:display});
	if(tooltip) o.Tooltip = tooltip;
	return o;
}
function write_BrtHLink(l, rId) {
	var o = new_buf(50+4*(l[1].Target.length + (l[1].Tooltip || "").length));
	write_UncheckedRfX({s:decode_cell(l[0]), e:decode_cell(l[0])}, o);
	write_RelID("rId" + rId, o);
	var locidx = l[1].Target.indexOf("#");
	var loc = locidx == -1 ? "" : l[1].Target.slice(locidx+1);
	write_XLWideString(loc || "", o);
	write_XLWideString(l[1].Tooltip || "", o);
	write_XLWideString("", o);
	return o.slice(0, o.l);
}

/* [MS-XLSB] 2.4.6 BrtArrFmla */
function parse_BrtArrFmla(data, length, opts) {
	var end = data.l + length;
	var rfx = parse_RfX(data, 16);
	var fAlwaysCalc = data.read_shift(1);
	var o = [rfx]; o[2] = fAlwaysCalc;
	if(opts.cellFormula) {
		var formula = parse_XLSBArrayParsedFormula(data, end - data.l, opts);
		o[1] = formula;
	} else data.l = end;
	return o;
}

/* [MS-XLSB] 2.4.750 BrtShrFmla */
function parse_BrtShrFmla(data, length, opts) {
	var end = data.l + length;
	var rfx = parse_UncheckedRfX(data, 16);
	var o = [rfx];
	if(opts.cellFormula) {
		var formula = parse_XLSBSharedParsedFormula(data, end - data.l, opts);
		o[1] = formula;
		data.l = end;
	} else data.l = end;
	return o;
}

/* [MS-XLSB] 2.4.323 BrtColInfo */
/* TODO: once XLS ColInfo is set, combine the functions */
function write_BrtColInfo(C, col, o) {
	if(o == null) o = new_buf(18);
	var p = col_obj_w(C, col);
	o.write_shift(-4, C);
	o.write_shift(-4, C);
	o.write_shift(4, (p.width || 10) * 256);
	o.write_shift(4, 0/*ixfe*/); // style
	var flags = 0;
	if(col.hidden) flags |= 0x01;
	if(typeof p.width == 'number') flags |= 0x02;
	o.write_shift(1, flags); // bit flag
	o.write_shift(1, 0); // bit flag
	return o;
}

/* [MS-XLSB] 2.4.678 BrtMargins */
var BrtMarginKeys = ["left","right","top","bottom","header","footer"];
function parse_BrtMargins(data) {
	var margins = ({});
	BrtMarginKeys.forEach(function(k) { margins[k] = parse_Xnum(data, 8); });
	return margins;
}
function write_BrtMargins(margins, o) {
	if(o == null) o = new_buf(6*8);
	default_margins(margins);
	BrtMarginKeys.forEach(function(k) { write_Xnum((margins)[k], o); });
	return o;
}

/* [MS-XLSB] 2.4.299 BrtBeginWsView */
function parse_BrtBeginWsView(data) {
	var f = data.read_shift(2);
	data.l += 28;
	return { RTL: f & 0x20 };
}
function write_BrtBeginWsView(ws, Workbook, o) {
	if(o == null) o = new_buf(30);
	var f = 0x39c;
	if((((Workbook||{}).Views||[])[0]||{}).RTL) f |= 0x20;
	o.write_shift(2, f); // bit flag
	o.write_shift(4, 0);
	o.write_shift(4, 0); // view first row
	o.write_shift(4, 0); // view first col
	o.write_shift(1, 0); // gridline color ICV
	o.write_shift(1, 0);
	o.write_shift(2, 0);
	o.write_shift(2, 100); // zoom scale
	o.write_shift(2, 0);
	o.write_shift(2, 0);
	o.write_shift(2, 0);
	o.write_shift(4, 0); // workbook view id
	return o;
}

/* [MS-XLSB] 2.4.309 BrtCellIgnoreEC */
function write_BrtCellIgnoreEC(ref) {
	var o = new_buf(24);
	o.write_shift(4, 4);
	o.write_shift(4, 1);
	write_UncheckedRfX(ref, o);
	return o;
}

/* [MS-XLSB] 2.4.748 BrtSheetProtection */
function write_BrtSheetProtection(sp, o) {
	if(o == null) o = new_buf(16*4+2);
	o.write_shift(2, sp.password ? crypto_CreatePasswordVerifier_Method1(sp.password) : 0);
	o.write_shift(4, 1); // this record should not be written if no protection
	[
		["objects",             false], // fObjects
		["scenarios",           false], // fScenarios
		["formatCells",          true], // fFormatCells
		["formatColumns",        true], // fFormatColumns
		["formatRows",           true], // fFormatRows
		["insertColumns",        true], // fInsertColumns
		["insertRows",           true], // fInsertRows
		["insertHyperlinks",     true], // fInsertHyperlinks
		["deleteColumns",        true], // fDeleteColumns
		["deleteRows",           true], // fDeleteRows
		["selectLockedCells",   false], // fSelLockedCells
		["sort",                 true], // fSort
		["autoFilter",           true], // fAutoFilter
		["pivotTables",          true], // fPivotTables
		["selectUnlockedCells", false]  // fSelUnlockedCells
	].forEach(function(n) {
if(n[1]) o.write_shift(4, sp[n[0]] != null && !sp[n[0]] ? 1 : 0);
		else      o.write_shift(4, sp[n[0]] != null && sp[n[0]] ? 0 : 1);
	});
	return o;
}

/* [MS-XLSB] 2.1.7.61 Worksheet */
function parse_ws_bin(data, _opts, idx, rels, wb, themes, styles) {
	if(!data) return data;
	var opts = _opts || {};
	if(!rels) rels = {'!id':{}};
	if(DENSE != null && opts.dense == null) opts.dense = DENSE;
	var s = (opts.dense ? [] : {});

	var ref;
	var refguess = {s: {r:2000000, c:2000000}, e: {r:0, c:0} };

	var pass = false, end = false;
	var row, p, cf, R, C, addr, sstr, rr, cell;
	var merges = [];
	opts.biff = 12;
	opts['!row'] = 0;

	var ai = 0, af = false;

	var arrayf = [];
	var sharedf = {};
	var supbooks = opts.supbooks || ([[]]);
	supbooks.sharedf = sharedf;
	supbooks.arrayf = arrayf;
	supbooks.SheetNames = wb.SheetNames || wb.Sheets.map(function(x) { return x.name; });
	if(!opts.supbooks) {
		opts.supbooks = supbooks;
		if(wb.Names) for(var i = 0; i < wb.Names.length; ++i) supbooks[0][i+1] = wb.Names[i];
	}

	var colinfo = [], rowinfo = [];
	var seencol = false;

	recordhopper(data, function ws_parse(val, R_n, RT) {
		if(end) return;
		switch(RT) {
			case 0x0094: /* 'BrtWsDim' */
				ref = val; break;
			case 0x0000: /* 'BrtRowHdr' */
				row = val;
				if(opts.sheetRows && opts.sheetRows <= row.r) end=true;
				rr = encode_row(R = row.r);
				opts['!row'] = row.r;
				if(val.hidden || val.hpt || val.level != null) {
					if(val.hpt) val.hpx = pt2px(val.hpt);
					rowinfo[val.r] = val;
				}
				break;

			case 0x0002: /* 'BrtCellRk' */
			case 0x0003: /* 'BrtCellError' */
			case 0x0004: /* 'BrtCellBool' */
			case 0x0005: /* 'BrtCellReal' */
			case 0x0006: /* 'BrtCellSt' */
			case 0x0007: /* 'BrtCellIsst' */
			case 0x0008: /* 'BrtFmlaString' */
			case 0x0009: /* 'BrtFmlaNum' */
			case 0x000A: /* 'BrtFmlaBool' */
			case 0x000B: /* 'BrtFmlaError' */
				p = ({t:val[2]});
				switch(val[2]) {
					case 'n': p.v = val[1]; break;
					case 's': sstr = strs[val[1]]; p.v = sstr.t; p.r = sstr.r; break;
					case 'b': p.v = val[1] ? true : false; break;
					case 'e': p.v = val[1]; if(opts.cellText !== false) p.w = BErr[p.v]; break;
					case 'str': p.t = 's'; p.v = val[1]; break;
				}
				if((cf = styles.CellXf[val[0].iStyleRef])) safe_format(p,cf.numFmtId,null,opts, themes, styles);
				C = val[0].c;
				if(opts.dense) { if(!s[R]) s[R] = []; s[R][C] = p; }
				else s[encode_col(C) + rr] = p;
				if(opts.cellFormula) {
					af = false;
					for(ai = 0; ai < arrayf.length; ++ai) {
						var aii = arrayf[ai];
						if(row.r >= aii[0].s.r && row.r <= aii[0].e.r)
							if(C >= aii[0].s.c && C <= aii[0].e.c) {
								p.F = encode_range(aii[0]); af = true;
							}
					}
					if(!af && val.length > 3) p.f = val[3];
				}
				if(refguess.s.r > row.r) refguess.s.r = row.r;
				if(refguess.s.c > C) refguess.s.c = C;
				if(refguess.e.r < row.r) refguess.e.r = row.r;
				if(refguess.e.c < C) refguess.e.c = C;
				if(opts.cellDates && cf && p.t == 'n' && SSF.is_date(SSF._table[cf.numFmtId])) {
					var _d = SSF.parse_date_code(p.v); if(_d) { p.t = 'd'; p.v = new Date(_d.y, _d.m-1,_d.d,_d.H,_d.M,_d.S,_d.u); }
				}
				break;

			case 0x0001: /* 'BrtCellBlank' */
				if(!opts.sheetStubs || pass) break;
				p = ({t:'z',v:undefined});
				C = val[0].c;
				if(opts.dense) { if(!s[R]) s[R] = []; s[R][C] = p; }
				else s[encode_col(C) + rr] = p;
				if(refguess.s.r > row.r) refguess.s.r = row.r;
				if(refguess.s.c > C) refguess.s.c = C;
				if(refguess.e.r < row.r) refguess.e.r = row.r;
				if(refguess.e.c < C) refguess.e.c = C;
				break;

			case 0x00B0: /* 'BrtMergeCell' */
				merges.push(val); break;

			case 0x01EE: /* 'BrtHLink' */
				var rel = rels['!id'][val.relId];
				if(rel) {
					val.Target = rel.Target;
					if(val.loc) val.Target += "#"+val.loc;
					val.Rel = rel;
				} else if(val.relId == '') {
					val.Target = "#" + val.loc;
				}
				for(R=val.rfx.s.r;R<=val.rfx.e.r;++R) for(C=val.rfx.s.c;C<=val.rfx.e.c;++C) {
					if(opts.dense) {
						if(!s[R]) s[R] = [];
						if(!s[R][C]) s[R][C] = {t:'z',v:undefined};
						s[R][C].l = val;
					} else {
						addr = encode_cell({c:C,r:R});
						if(!s[addr]) s[addr] = {t:'z',v:undefined};
						s[addr].l = val;
					}
				}
				break;

			case 0x01AA: /* 'BrtArrFmla' */
				if(!opts.cellFormula) break;
				arrayf.push(val);
				cell = ((opts.dense ? s[R][C] : s[encode_col(C) + rr]));
				cell.f = stringify_formula(val[1], refguess, {r:row.r, c:C}, supbooks, opts);
				cell.F = encode_range(val[0]);
				break;
			case 0x01AB: /* 'BrtShrFmla' */
				if(!opts.cellFormula) break;
				sharedf[encode_cell(val[0].s)] = val[1];
				cell = (opts.dense ? s[R][C] : s[encode_col(C) + rr]);
				cell.f = stringify_formula(val[1], refguess, {r:row.r, c:C}, supbooks, opts);
				break;

			/* identical to 'ColInfo' in XLS */
			case 0x003C: /* 'BrtColInfo' */
				if(!opts.cellStyles) break;
				while(val.e >= val.s) {
					colinfo[val.e--] = { width: val.w/256, hidden: !!(val.flags & 0x01) };
					if(!seencol) { seencol = true; find_mdw_colw(val.w/256); }
					process_col(colinfo[val.e+1]);
				}
				break;

			case 0x00A1: /* 'BrtBeginAFilter' */
				s['!autofilter'] = { ref:encode_range(val) };
				break;

			case 0x01DC: /* 'BrtMargins' */
				s['!margins'] = val;
				break;

			case 0x0093: /* 'BrtWsProp' */
				if(!wb.Sheets[idx]) wb.Sheets[idx] = {};
				if(val.name) wb.Sheets[idx].CodeName = val.name;
				break;

			case 0x0089: /* 'BrtBeginWsView' */
				if(!wb.Views) wb.Views = [{}];
				if(!wb.Views[0]) wb.Views[0] = {};
				if(val.RTL) wb.Views[0].RTL = true;
				break;

			case 0x01E5: /* 'BrtWsFmtInfo' */
				break;
			case 0x00AF: /* 'BrtAFilterDateGroupItem' */
			case 0x0284: /* 'BrtActiveX' */
			case 0x0271: /* 'BrtBigName' */
			case 0x0232: /* 'BrtBkHim' */
			case 0x018C: /* 'BrtBrk' */
			case 0x0458: /* 'BrtCFIcon' */
			case 0x047A: /* 'BrtCFRuleExt' */
			case 0x01D7: /* 'BrtCFVO' */
			case 0x041A: /* 'BrtCFVO14' */
			case 0x0289: /* 'BrtCellIgnoreEC' */
			case 0x0451: /* 'BrtCellIgnoreEC14' */
			case 0x0031: /* 'BrtCellMeta' */
			case 0x024D: /* 'BrtCellSmartTagProperty' */
			case 0x025F: /* 'BrtCellWatch' */
			case 0x0234: /* 'BrtColor' */
			case 0x041F: /* 'BrtColor14' */
			case 0x00A8: /* 'BrtColorFilter' */
			case 0x00AE: /* 'BrtCustomFilter' */
			case 0x049C: /* 'BrtCustomFilter14' */
			case 0x01F3: /* 'BrtDRef' */
			case 0x0040: /* 'BrtDVal' */
			case 0x041D: /* 'BrtDVal14' */
			case 0x0226: /* 'BrtDrawing' */
			case 0x00AB: /* 'BrtDynamicFilter' */
			case 0x00A7: /* 'BrtFilter' */
			case 0x0499: /* 'BrtFilter14' */
			case 0x00A9: /* 'BrtIconFilter' */
			case 0x049D: /* 'BrtIconFilter14' */
			case 0x0227: /* 'BrtLegacyDrawing' */
			case 0x0228: /* 'BrtLegacyDrawingHF' */
			case 0x0295: /* 'BrtListPart' */
			case 0x027F: /* 'BrtOleObject' */
			case 0x01DE: /* 'BrtPageSetup' */
			case 0x0097: /* 'BrtPane' */
			case 0x0219: /* 'BrtPhoneticInfo' */
			case 0x01DD: /* 'BrtPrintOptions' */
			case 0x0218: /* 'BrtRangeProtection' */
			case 0x044F: /* 'BrtRangeProtection14' */
			case 0x02A8: /* 'BrtRangeProtectionIso' */
			case 0x0450: /* 'BrtRangeProtectionIso14' */
			case 0x0400: /* 'BrtRwDescent' */
			case 0x0098: /* 'BrtSel' */
			case 0x0297: /* 'BrtSheetCalcProp' */
			case 0x0217: /* 'BrtSheetProtection' */
			case 0x02A6: /* 'BrtSheetProtectionIso' */
			case 0x01F8: /* 'BrtSlc' */
			case 0x0413: /* 'BrtSparkline' */
			case 0x01AC: /* 'BrtTable' */
			case 0x00AA: /* 'BrtTop10Filter' */
			case 0x0C00: /* 'BrtUid' */
			case 0x0032: /* 'BrtValueMeta' */
			case 0x0816: /* 'BrtWebExtension' */
			case 0x0415: /* 'BrtWsFmtInfoEx14' */
				break;

			case 0x0023: /* 'BrtFRTBegin' */
				pass = true; break;
			case 0x0024: /* 'BrtFRTEnd' */
				pass = false; break;
			case 0x0025: /* 'BrtACBegin' */ break;
			case 0x0026: /* 'BrtACEnd' */ break;

			default:
				if((R_n||"").indexOf("Begin") > 0){/* empty */}
				else if((R_n||"").indexOf("End") > 0){/* empty */}
				else if(!pass || opts.WTF) throw new Error("Unexpected record " + RT + " " + R_n);
		}
	}, opts);

	delete opts.supbooks;
	delete opts['!row'];

	if(!s["!ref"] && (refguess.s.r < 2000000 || ref && (ref.e.r > 0 || ref.e.c > 0 || ref.s.r > 0 || ref.s.c > 0))) s["!ref"] = encode_range(ref || refguess);
	if(opts.sheetRows && s["!ref"]) {
		var tmpref = safe_decode_range(s["!ref"]);
		if(opts.sheetRows <= +tmpref.e.r) {
			tmpref.e.r = opts.sheetRows - 1;
			if(tmpref.e.r > refguess.e.r) tmpref.e.r = refguess.e.r;
			if(tmpref.e.r < tmpref.s.r) tmpref.s.r = tmpref.e.r;
			if(tmpref.e.c > refguess.e.c) tmpref.e.c = refguess.e.c;
			if(tmpref.e.c < tmpref.s.c) tmpref.s.c = tmpref.e.c;
			s["!fullref"] = s["!ref"];
			s["!ref"] = encode_range(tmpref);
		}
	}
	if(merges.length > 0) s["!merges"] = merges;
	if(colinfo.length > 0) s["!cols"] = colinfo;
	if(rowinfo.length > 0) s["!rows"] = rowinfo;
	return s;
}

/* TODO: something useful -- this is a stub */
function write_ws_bin_cell(ba, cell, R, C, opts, ws) {
	if(cell.v === undefined) return "";
	var vv = "";
	switch(cell.t) {
		case 'b': vv = cell.v ? "1" : "0"; break;
		case 'd': // no BrtCellDate :(
			cell = dup(cell);
			cell.z = cell.z || SSF._table[14];
			cell.v = datenum(parseDate(cell.v)); cell.t = 'n';
			break;
		/* falls through */
		case 'n': case 'e': vv = ''+cell.v; break;
		default: vv = cell.v; break;
	}
	var o = ({r:R, c:C});
	/* TODO: cell style */
	o.s = get_cell_style(opts.cellXfs, cell, opts);
	if(cell.l) ws['!links'].push([encode_cell(o), cell.l]);
	if(cell.c) ws['!comments'].push([encode_cell(o), cell.c]);
	switch(cell.t) {
		case 's': case 'str':
			if(opts.bookSST) {
				vv = get_sst_id(opts.Strings, (cell.v), opts.revStrings);
				o.t = "s"; o.v = vv;
				write_record(ba, "BrtCellIsst", write_BrtCellIsst(cell, o));
			} else {
				o.t = "str";
				write_record(ba, "BrtCellSt", write_BrtCellSt(cell, o));
			}
			return;
		case 'n':
			/* TODO: determine threshold for Real vs RK */
			if(cell.v == (cell.v | 0) && cell.v > -1000 && cell.v < 1000) write_record(ba, "BrtCellRk", write_BrtCellRk(cell, o));
			else write_record(ba, "BrtCellReal", write_BrtCellReal(cell, o));
			return;
		case 'b':
			o.t = "b";
			write_record(ba, "BrtCellBool", write_BrtCellBool(cell, o));
			return;
		case 'e': /* TODO: error */ o.t = "e"; break;
	}
	write_record(ba, "BrtCellBlank", write_BrtCellBlank(cell, o));
}

function write_CELLTABLE(ba, ws, idx, opts) {
	var range = safe_decode_range(ws['!ref'] || "A1"), ref, rr = "", cols = [];
	write_record(ba, 'BrtBeginSheetData');
	var dense = Array.isArray(ws);
	var cap = range.e.r;
	if(ws['!rows']) cap = Math.max(range.e.r, ws['!rows'].length - 1);
	for(var R = range.s.r; R <= cap; ++R) {
		rr = encode_row(R);
		/* [ACCELLTABLE] */
		/* BrtRowHdr */
		write_row_header(ba, ws, range, R);
		if(R <= range.e.r) for(var C = range.s.c; C <= range.e.c; ++C) {
			/* *16384CELL */
			if(R === range.s.r) cols[C] = encode_col(C);
			ref = cols[C] + rr;
			var cell = dense ? (ws[R]||[])[C] : ws[ref];
			if(!cell) continue;
			/* write cell */
			write_ws_bin_cell(ba, cell, R, C, opts, ws);
		}
	}
	write_record(ba, 'BrtEndSheetData');
}

function write_MERGECELLS(ba, ws) {
	if(!ws || !ws['!merges']) return;
	write_record(ba, 'BrtBeginMergeCells', write_BrtBeginMergeCells(ws['!merges'].length));
	ws['!merges'].forEach(function(m) { write_record(ba, 'BrtMergeCell', write_BrtMergeCell(m)); });
	write_record(ba, 'BrtEndMergeCells');
}

function write_COLINFOS(ba, ws) {
	if(!ws || !ws['!cols']) return;
	write_record(ba, 'BrtBeginColInfos');
	ws['!cols'].forEach(function(m, i) { if(m) write_record(ba, 'BrtColInfo', write_BrtColInfo(i, m)); });
	write_record(ba, 'BrtEndColInfos');
}

function write_IGNOREECS(ba, ws) {
	if(!ws || !ws['!ref']) return;
	write_record(ba, 'BrtBeginCellIgnoreECs');
	write_record(ba, 'BrtCellIgnoreEC', write_BrtCellIgnoreEC(safe_decode_range(ws['!ref'])));
	write_record(ba, 'BrtEndCellIgnoreECs');
}

function write_HLINKS(ba, ws, rels) {
	/* *BrtHLink */
	ws['!links'].forEach(function(l) {
		if(!l[1].Target) return;
		var rId = add_rels(rels, -1, l[1].Target.replace(/#.*$/, ""), RELS.HLINK);
		write_record(ba, "BrtHLink", write_BrtHLink(l, rId));
	});
	delete ws['!links'];
}
function write_LEGACYDRAWING(ba, ws, idx, rels) {
	/* [BrtLegacyDrawing] */
	if(ws['!comments'].length > 0) {
		var rId = add_rels(rels, -1, "../drawings/vmlDrawing" + (idx+1) + ".vml", RELS.VML);
		write_record(ba, "BrtLegacyDrawing", write_RelID("rId" + rId));
		ws['!legacy'] = rId;
	}
}

function write_AUTOFILTER(ba, ws) {
	if(!ws['!autofilter']) return;
	write_record(ba, "BrtBeginAFilter", write_UncheckedRfX(safe_decode_range(ws['!autofilter'].ref)));
	/* *FILTERCOLUMN */
	/* [SORTSTATE] */
	/* BrtEndAFilter */
	write_record(ba, "BrtEndAFilter");
}

function write_WSVIEWS2(ba, ws, Workbook) {
	write_record(ba, "BrtBeginWsViews");
	{ /* 1*WSVIEW2 */
		/* [ACUID] */
		write_record(ba, "BrtBeginWsView", write_BrtBeginWsView(ws, Workbook));
		/* [BrtPane] */
		/* *4BrtSel */
		/* *4SXSELECT */
		/* *FRT */
		write_record(ba, "BrtEndWsView");
	}
	/* *FRT */
	write_record(ba, "BrtEndWsViews");
}

function write_WSFMTINFO() {
	/* [ACWSFMTINFO] */
	//write_record(ba, "BrtWsFmtInfo", write_BrtWsFmtInfo(ws));
}

function write_SHEETPROTECT(ba, ws) {
	if(!ws['!protect']) return;
	/* [BrtSheetProtectionIso] */
	write_record(ba, "BrtSheetProtection", write_BrtSheetProtection(ws['!protect']));
}

function write_ws_bin(idx, opts, wb, rels) {
	var ba = buf_array();
	var s = wb.SheetNames[idx], ws = wb.Sheets[s] || {};
	var c = s; try { if(wb && wb.Workbook) c = wb.Workbook.Sheets[idx].CodeName || c; } catch(e) {}
	var r = safe_decode_range(ws['!ref'] || "A1");
	if(r.e.c > 0x3FFF || r.e.r > 0xFFFFF) {
		if(opts.WTF) throw new Error("Range " + (ws['!ref'] || "A1") + " exceeds format limit A1:XFD1048576");
		r.e.c = Math.min(r.e.c, 0x3FFF);
		r.e.r = Math.min(r.e.c, 0xFFFFF);
	}
	ws['!links'] = [];
	/* passed back to write_zip and removed there */
	ws['!comments'] = [];
	write_record(ba, "BrtBeginSheet");
	if(wb.vbaraw) write_record(ba, "BrtWsProp", write_BrtWsProp(c));
	write_record(ba, "BrtWsDim", write_BrtWsDim(r));
	write_WSVIEWS2(ba, ws, wb.Workbook);
	write_WSFMTINFO(ba, ws);
	write_COLINFOS(ba, ws, idx, opts, wb);
	write_CELLTABLE(ba, ws, idx, opts, wb);
	/* [BrtSheetCalcProp] */
	write_SHEETPROTECT(ba, ws);
	/* *([BrtRangeProtectionIso] BrtRangeProtection) */
	/* [SCENMAN] */
	write_AUTOFILTER(ba, ws);
	/* [SORTSTATE] */
	/* [DCON] */
	/* [USERSHVIEWS] */
	write_MERGECELLS(ba, ws);
	/* [BrtPhoneticInfo] */
	/* *CONDITIONALFORMATTING */
	/* [DVALS] */
	write_HLINKS(ba, ws, rels);
	/* [BrtPrintOptions] */
	if(ws['!margins']) write_record(ba, "BrtMargins", write_BrtMargins(ws['!margins']));
	/* [BrtPageSetup] */
	/* [HEADERFOOTER] */
	/* [RWBRK] */
	/* [COLBRK] */
	/* *BrtBigName */
	/* [CELLWATCHES] */
	if(!opts || opts.ignoreEC || (opts.ignoreEC == (void 0))) write_IGNOREECS(ba, ws);
	/* [SMARTTAGS] */
	/* [BrtDrawing] */
	write_LEGACYDRAWING(ba, ws, idx, rels);
	/* [BrtLegacyDrawingHF] */
	/* [BrtBkHim] */
	/* [OLEOBJECTS] */
	/* [ACTIVEXCONTROLS] */
	/* [WEBPUBITEMS] */
	/* [LISTPARTS] */
	/* FRTWORKSHEET */
	write_record(ba, "BrtEndSheet");
	return ba.end();
}
function parse_numCache(data) {
	var col = [];

	/* 21.2.2.150 pt CT_NumVal */
	(data.match(/<c:pt idx="(\d*)">(.*?)<\/c:pt>/mg)||[]).forEach(function(pt) {
		var q = pt.match(/<c:pt idx="(\d*?)"><c:v>(.*)<\/c:v><\/c:pt>/);
		if(!q) return;
		col[+q[1]] = +q[2];
	});

	/* 21.2.2.71 formatCode CT_Xstring */
	var nf = unescapexml((data.match(/<c:formatCode>([\s\S]*?)<\/c:formatCode>/) || ["","General"])[1]);

	return [col, nf];
}

/* 21.2 DrawingML - Charts */
function parse_chart(data, name, opts, rels, wb, csheet) {
	var cs = ((csheet || {"!type":"chart"}));
	if(!data) return csheet;
	/* 21.2.2.27 chart CT_Chart */

	var C = 0, R = 0, col = "A";
	var refguess = {s: {r:2000000, c:2000000}, e: {r:0, c:0} };

	/* 21.2.2.120 numCache CT_NumData */
	(data.match(/<c:numCache>[\s\S]*?<\/c:numCache>/gm)||[]).forEach(function(nc) {
		var cache = parse_numCache(nc);
		refguess.s.r = refguess.s.c = 0;
		refguess.e.c = C;
		col = encode_col(C);
		cache[0].forEach(function(n,i) {
			cs[col + encode_row(i)] = {t:'n', v:n, z:cache[1] };
			R = i;
		});
		if(refguess.e.r < R) refguess.e.r = R;
		++C;
	});
	if(C > 0) cs["!ref"] = encode_range(refguess);
	return cs;
}
RELS.CS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";

var CS_XML_ROOT = writextag('chartsheet', null, {
	'xmlns': XMLNS.main[0],
	'xmlns:r': XMLNS.r
});

/* 18.3 Worksheets also covers Chartsheets */
function parse_cs_xml(data, opts, idx, rels, wb) {
	if(!data) return data;
	/* 18.3.1.12 chartsheet CT_ChartSheet */
	if(!rels) rels = {'!id':{}};
	var s = {'!type':"chart", '!chart':null, '!rel':""};
	var m;

	/* 18.3.1.83 sheetPr CT_ChartsheetPr */
	var sheetPr = data.match(sheetprregex);
	if(sheetPr) parse_ws_xml_sheetpr(sheetPr[0], s, wb, idx);

	/* 18.3.1.36 drawing CT_Drawing */
	if((m = data.match(/drawing r:id="(.*?)"/))) s['!rel'] = m[1];

	if(rels['!id'][s['!rel']]) s['!chart'] = rels['!id'][s['!rel']];
	return s;
}
function write_cs_xml(idx, opts, wb, rels) {
	var o = [XML_HEADER, CS_XML_ROOT];
	o[o.length] = writextag("drawing", null, {"r:id": "rId1"});
	add_rels(rels, -1, "../drawings/drawing" + (idx+1) + ".xml", RELS.DRAW);
	if(o.length>2) { o[o.length] = ('</chartsheet>'); o[1]=o[1].replace("/>",">"); }
	return o.join("");
}

/* [MS-XLSB] 2.4.331 BrtCsProp */
function parse_BrtCsProp(data, length) {
	data.l += 10;
	var name = parse_XLWideString(data, length - 10);
	return { name: name };
}

/* [MS-XLSB] 2.1.7.7 Chart Sheet */
function parse_cs_bin(data, opts, idx, rels, wb) {
	if(!data) return data;
	if(!rels) rels = {'!id':{}};
	var s = {'!type':"chart", '!chart':null, '!rel':""};
	var state = [];
	var pass = false;
	recordhopper(data, function cs_parse(val, R_n, RT) {
		switch(RT) {

			case 0x0226: /* 'BrtDrawing' */
				s['!rel'] = val; break;

			case 0x028B: /* 'BrtCsProp' */
				if(!wb.Sheets[idx]) wb.Sheets[idx] = {};
				if(val.name) wb.Sheets[idx].CodeName = val.name;
				break;

			case 0x0232: /* 'BrtBkHim' */
			case 0x028C: /* 'BrtCsPageSetup' */
			case 0x029D: /* 'BrtCsProtection' */
			case 0x02A7: /* 'BrtCsProtectionIso' */
			case 0x0227: /* 'BrtLegacyDrawing' */
			case 0x0228: /* 'BrtLegacyDrawingHF' */
			case 0x01DC: /* 'BrtMargins' */
			case 0x0C00: /* 'BrtUid' */
				break;

			case 0x0023: /* 'BrtFRTBegin' */
				pass = true; break;
			case 0x0024: /* 'BrtFRTEnd' */
				pass = false; break;
			case 0x0025: /* 'BrtACBegin' */
				state.push(R_n); break;
			case 0x0026: /* 'BrtACEnd' */
				state.pop(); break;

			default:
				if((R_n||"").indexOf("Begin") > 0) state.push(R_n);
				else if((R_n||"").indexOf("End") > 0) state.pop();
				else if(!pass || opts.WTF) throw new Error("Unexpected record " + RT + " " + R_n);
		}
	}, opts);

	if(rels['!id'][s['!rel']]) s['!chart'] = rels['!id'][s['!rel']];
	return s;
}
function write_cs_bin() {
	var ba = buf_array();
	write_record(ba, "BrtBeginSheet");
	/* [BrtCsProp] */
	/* CSVIEWS */
	/* [[BrtCsProtectionIso] BrtCsProtection] */
	/* [USERCSVIEWS] */
	/* [BrtMargins] */
	/* [BrtCsPageSetup] */
	/* [HEADERFOOTER] */
	/* BrtDrawing */
	/* [BrtLegacyDrawing] */
	/* [BrtLegacyDrawingHF] */
	/* [BrtBkHim] */
	/* [WEBPUBITEMS] */
	/* FRTCHARTSHEET */
	write_record(ba, "BrtEndSheet");
	return ba.end();
}
/* 18.2.28 (CT_WorkbookProtection) Defaults */
var WBPropsDef = [
	['allowRefreshQuery',           false, "bool"],
	['autoCompressPictures',        true,  "bool"],
	['backupFile',                  false, "bool"],
	['checkCompatibility',          false, "bool"],
	['CodeName',                    ''],
	['date1904',                    false, "bool"],
	['defaultThemeVersion',         0,      "int"],
	['filterPrivacy',               false, "bool"],
	['hidePivotFieldList',          false, "bool"],
	['promptedSolutions',           false, "bool"],
	['publishItems',                false, "bool"],
	['refreshAllConnections',       false, "bool"],
	['saveExternalLinkValues',      true,  "bool"],
	['showBorderUnselectedTables',  true,  "bool"],
	['showInkAnnotation',           true,  "bool"],
	['showObjects',                 'all'],
	['showPivotChartFilter',        false, "bool"],
	['updateLinks', 'userSet']
];

/* 18.2.30 (CT_BookView) Defaults */
var WBViewDef = [
	['activeTab',                   0,      "int"],
	['autoFilterDateGrouping',      true,  "bool"],
	['firstSheet',                  0,      "int"],
	['minimized',                   false, "bool"],
	['showHorizontalScroll',        true,  "bool"],
	['showSheetTabs',               true,  "bool"],
	['showVerticalScroll',          true,  "bool"],
	['tabRatio',                    600,    "int"],
	['visibility',                  'visible']
	//window{Height,Width}, {x,y}Window
];

/* 18.2.19 (CT_Sheet) Defaults */
var SheetDef = [
	//['state', 'visible']
];

/* 18.2.2  (CT_CalcPr) Defaults */
var CalcPrDef = [
	['calcCompleted', 'true'],
	['calcMode', 'auto'],
	['calcOnSave', 'true'],
	['concurrentCalc', 'true'],
	['fullCalcOnLoad', 'false'],
	['fullPrecision', 'true'],
	['iterate', 'false'],
	['iterateCount', '100'],
	['iterateDelta', '0.001'],
	['refMode', 'A1']
];

/* 18.2.3 (CT_CustomWorkbookView) Defaults */
/*var CustomWBViewDef = [
	['autoUpdate', 'false'],
	['changesSavedWin', 'false'],
	['includeHiddenRowCol', 'true'],
	['includePrintSettings', 'true'],
	['maximized', 'false'],
	['minimized', 'false'],
	['onlySync', 'false'],
	['personalView', 'false'],
	['showComments', 'commIndicator'],
	['showFormulaBar', 'true'],
	['showHorizontalScroll', 'true'],
	['showObjects', 'all'],
	['showSheetTabs', 'true'],
	['showStatusbar', 'true'],
	['showVerticalScroll', 'true'],
	['tabRatio', '600'],
	['xWindow', '0'],
	['yWindow', '0']
];*/

function push_defaults_array(target, defaults) {
	for(var j = 0; j != target.length; ++j) { var w = target[j];
		for(var i=0; i != defaults.length; ++i) { var z = defaults[i];
			if(w[z[0]] == null) w[z[0]] = z[1];
			else switch(z[2]) {
			case "bool": if(typeof w[z[0]] == "string") w[z[0]] = parsexmlbool(w[z[0]]); break;
			case "int": if(typeof w[z[0]] == "string") w[z[0]] = parseInt(w[z[0]], 10); break;
			}
		}
	}
}
function push_defaults(target, defaults) {
	for(var i = 0; i != defaults.length; ++i) { var z = defaults[i];
		if(target[z[0]] == null) target[z[0]] = z[1];
		else switch(z[2]) {
			case "bool": if(typeof target[z[0]] == "string") target[z[0]] = parsexmlbool(target[z[0]]); break;
			case "int": if(typeof target[z[0]] == "string") target[z[0]] = parseInt(target[z[0]], 10); break;
		}
	}
}

function parse_wb_defaults(wb) {
	push_defaults(wb.WBProps, WBPropsDef);
	push_defaults(wb.CalcPr, CalcPrDef);

	push_defaults_array(wb.WBView, WBViewDef);
	push_defaults_array(wb.Sheets, SheetDef);

	_ssfopts.date1904 = parsexmlbool(wb.WBProps.date1904);
}

function safe1904(wb) {
	/* TODO: store date1904 somewhere else */
	if(!wb.Workbook) return "false";
	if(!wb.Workbook.WBProps) return "false";
	return parsexmlbool(wb.Workbook.WBProps.date1904) ? "true" : "false";
}

var badchars = "][*?\/\\".split("");
function check_ws_name(n, safe) {
	if(n.length > 31) { if(safe) return false; throw new Error("Sheet names cannot exceed 31 chars"); }
	var _good = true;
	badchars.forEach(function(c) {
		if(n.indexOf(c) == -1) return;
		if(!safe) throw new Error("Sheet name cannot contain : \\ / ? * [ ]");
		_good = false;
	});
	return _good;
}
function check_wb_names(N, S, codes) {
	N.forEach(function(n,i) {
		check_ws_name(n);
		for(var j = 0; j < i; ++j) if(n == N[j]) throw new Error("Duplicate Sheet Name: " + n);
		if(codes) {
			var cn = (S && S[i] && S[i].CodeName) || n;
			if(cn.charCodeAt(0) == 95 && cn.length > 22) throw new Error("Bad Code Name: Worksheet" + cn);
		}
	});
}
function check_wb(wb) {
	if(!wb || !wb.SheetNames || !wb.Sheets) throw new Error("Invalid Workbook");
	if(!wb.SheetNames.length) throw new Error("Workbook is empty");
	var Sheets = (wb.Workbook && wb.Workbook.Sheets) || [];
	check_wb_names(wb.SheetNames, Sheets, !!wb.vbaraw);
	for(var i = 0; i < wb.SheetNames.length; ++i) check_ws(wb.Sheets[wb.SheetNames[i]], wb.SheetNames[i], i);
	/* TODO: validate workbook */
}
/* 18.2 Workbook */
var wbnsregex = /<\w+:workbook/;
function parse_wb_xml(data, opts) {
	if(!data) throw new Error("Could not find file");
	var wb = { AppVersion:{}, WBProps:{}, WBView:[], Sheets:[], CalcPr:{}, Names:[], xmlns: "" };
	var pass = false, xmlns = "xmlns";
	var dname = {}, dnstart = 0;
	data.replace(tagregex, function xml_wb(x, idx) {
		var y = parsexmltag(x);
		switch(strip_ns(y[0])) {
			case '<?xml': break;

			/* 18.2.27 workbook CT_Workbook 1 */
			case '<workbook':
				if(x.match(wbnsregex)) xmlns = "xmlns" + x.match(/<(\w+):/)[1];
				wb.xmlns = y[xmlns];
				break;
			case '</workbook>': break;

			/* 18.2.13 fileVersion CT_FileVersion ? */
			case '<fileVersion': delete y[0]; wb.AppVersion = y; break;
			case '<fileVersion/>': case '</fileVersion>': break;

			/* 18.2.12 fileSharing CT_FileSharing ? */
			case '<fileSharing': case '<fileSharing/>': break;

			/* 18.2.28 workbookPr CT_WorkbookPr ? */
			case '<workbookPr':
			case '<workbookPr/>':
				WBPropsDef.forEach(function(w) {
					if(y[w[0]] == null) return;
					switch(w[2]) {
						case "bool": wb.WBProps[w[0]] = parsexmlbool(y[w[0]]); break;
						case "int": wb.WBProps[w[0]] = parseInt(y[w[0]], 10); break;
						default: wb.WBProps[w[0]] = y[w[0]];
					}
				});
				if(y.codeName) wb.WBProps.CodeName = y.codeName;
				break;
			case '</workbookPr>': break;

			/* 18.2.29 workbookProtection CT_WorkbookProtection ? */
			case '<workbookProtection': break;
			case '<workbookProtection/>': break;

			/* 18.2.1  bookViews CT_BookViews ? */
			case '<bookViews': case '<bookViews>': case '</bookViews>': break;
			/* 18.2.30   workbookView CT_BookView + */
			case '<workbookView': case '<workbookView/>': delete y[0]; wb.WBView.push(y); break;
			case '</workbookView>': break;

			/* 18.2.20 sheets CT_Sheets 1 */
			case '<sheets': case '<sheets>': case '</sheets>': break; // aggregate sheet
			/* 18.2.19   sheet CT_Sheet + */
			case '<sheet':
				switch(y.state) {
					case "hidden": y.Hidden = 1; break;
					case "veryHidden": y.Hidden = 2; break;
					default: y.Hidden = 0;
				}
				delete y.state;
				y.name = unescapexml(utf8read(y.name));
				delete y[0]; wb.Sheets.push(y); break;
			case '</sheet>': break;

			/* 18.2.15 functionGroups CT_FunctionGroups ? */
			case '<functionGroups': case '<functionGroups/>': break;
			/* 18.2.14   functionGroup CT_FunctionGroup + */
			case '<functionGroup': break;

			/* 18.2.9  externalReferences CT_ExternalReferences ? */
			case '<externalReferences': case '</externalReferences>': case '<externalReferences>': break;
			/* 18.2.8    externalReference CT_ExternalReference + */
			case '<externalReference': break;

			/* 18.2.6  definedNames CT_DefinedNames ? */
			case '<definedNames/>': break;
			case '<definedNames>': case '<definedNames': pass=true; break;
			case '</definedNames>': pass=false; break;
			/* 18.2.5    definedName CT_DefinedName + */
			case '<definedName': {
				dname = {};
				dname.Name = utf8read(y.name);
				if(y.comment) dname.Comment = y.comment;
				if(y.localSheetId) dname.Sheet = +y.localSheetId;
				if(parsexmlbool(y.hidden||"0")) dname.Hidden = true;
				dnstart = idx + x.length;
			}	break;
			case '</definedName>': {
				dname.Ref = unescapexml(utf8read(data.slice(dnstart, idx)));
				wb.Names.push(dname);
			} break;
			case '<definedName/>': break;

			/* 18.2.2  calcPr CT_CalcPr ? */
			case '<calcPr': delete y[0]; wb.CalcPr = y; break;
			case '<calcPr/>': delete y[0]; wb.CalcPr = y; break;
			case '</calcPr>': break;

			/* 18.2.16 oleSize CT_OleSize ? (ref required) */
			case '<oleSize': break;

			/* 18.2.4  customWorkbookViews CT_CustomWorkbookViews ? */
			case '<customWorkbookViews>': case '</customWorkbookViews>': case '<customWorkbookViews': break;
			/* 18.2.3    customWorkbookView CT_CustomWorkbookView + */
			case '<customWorkbookView': case '</customWorkbookView>': break;

			/* 18.2.18 pivotCaches CT_PivotCaches ? */
			case '<pivotCaches>': case '</pivotCaches>': case '<pivotCaches': break;
			/* 18.2.17 pivotCache CT_PivotCache ? */
			case '<pivotCache': break;

			/* 18.2.21 smartTagPr CT_SmartTagPr ? */
			case '<smartTagPr': case '<smartTagPr/>': break;

			/* 18.2.23 smartTagTypes CT_SmartTagTypes ? */
			case '<smartTagTypes': case '<smartTagTypes>': case '</smartTagTypes>': break;
			/* 18.2.22   smartTagType CT_SmartTagType ? */
			case '<smartTagType': break;

			/* 18.2.24 webPublishing CT_WebPublishing ? */
			case '<webPublishing': case '<webPublishing/>': break;

			/* 18.2.11 fileRecoveryPr CT_FileRecoveryPr ? */
			case '<fileRecoveryPr': case '<fileRecoveryPr/>': break;

			/* 18.2.26 webPublishObjects CT_WebPublishObjects ? */
			case '<webPublishObjects>': case '<webPublishObjects': case '</webPublishObjects>': break;
			/* 18.2.25 webPublishObject CT_WebPublishObject ? */
			case '<webPublishObject': break;

			/* 18.2.10 extLst CT_ExtensionList ? */
			case '<extLst': case '<extLst>': case '</extLst>': case '<extLst/>': break;
			/* 18.2.7    ext CT_Extension + */
			case '<ext': pass=true; break; //TODO: check with versions of excel
			case '</ext>': pass=false; break;

			/* Others */
			case '<ArchID': break;
			case '<AlternateContent':
			case '<AlternateContent>': pass=true; break;
			case '</AlternateContent>': pass=false; break;

			/* TODO */
			case '<revisionPtr': break;

			default: if(!pass && opts.WTF) throw new Error('unrecognized ' + y[0] + ' in workbook');
		}
		return x;
	});
	if(XMLNS.main.indexOf(wb.xmlns) === -1) throw new Error("Unknown Namespace: " + wb.xmlns);

	parse_wb_defaults(wb);

	return wb;
}

var WB_XML_ROOT = writextag('workbook', null, {
	'xmlns': XMLNS.main[0],
	//'xmlns:mx': XMLNS.mx,
	//'xmlns:s': XMLNS.main[0],
	'xmlns:r': XMLNS.r
});

function write_wb_xml(wb) {
	var o = [XML_HEADER];
	o[o.length] = WB_XML_ROOT;

	var write_names = (wb.Workbook && (wb.Workbook.Names||[]).length > 0);

	/* fileVersion */
	/* fileSharing */

	var workbookPr = ({codeName:"ThisWorkbook"});
	if(wb.Workbook && wb.Workbook.WBProps) {
		WBPropsDef.forEach(function(x) {
if((wb.Workbook.WBProps[x[0]]) == null) return;
			if((wb.Workbook.WBProps[x[0]]) == x[1]) return;
			workbookPr[x[0]] = (wb.Workbook.WBProps[x[0]]);
		});
if(wb.Workbook.WBProps.CodeName) { workbookPr.codeName = wb.Workbook.WBProps.CodeName; delete workbookPr.CodeName; }
	}
	o[o.length] = (writextag('workbookPr', null, workbookPr));

	/* workbookProtection */

	var sheets = wb.Workbook && wb.Workbook.Sheets || [];
	var i = 0;

	/* bookViews */

	o[o.length] = "<sheets>";
	for(i = 0; i != wb.SheetNames.length; ++i) {
		var sht = ({name:escapexml(wb.SheetNames[i].slice(0,31))});
		sht.sheetId = ""+(i+1);
		sht["r:id"] = "rId"+(i+1);
		if(sheets[i]) switch(sheets[i].Hidden) {
			case 1: sht.state = "hidden"; break;
			case 2: sht.state = "veryHidden"; break;
		}
		o[o.length] = (writextag('sheet',null,sht));
	}
	o[o.length] = "</sheets>";

	/* functionGroups */
	/* externalReferences */

	if(write_names) {
		o[o.length] = "<definedNames>";
		if(wb.Workbook && wb.Workbook.Names) wb.Workbook.Names.forEach(function(n) {
			var d = {name:n.Name};
			if(n.Comment) d.comment = n.Comment;
			if(n.Sheet != null) d.localSheetId = ""+n.Sheet;
			if(n.Hidden) d.hidden = "1";
			if(!n.Ref) return;
			o[o.length] = writextag('definedName', String(n.Ref).replace(/</g, "&lt;").replace(/>/g, "&gt;"), d);
		});
		o[o.length] = "</definedNames>";
	}

	/* calcPr */
	/* oleSize */
	/* customWorkbookViews */
	/* pivotCaches */
	/* smartTagPr */
	/* smartTagTypes */
	/* webPublishing */
	/* fileRecoveryPr */
	/* webPublishObjects */
	/* extLst */

	if(o.length>2){ o[o.length] = '</workbook>'; o[1]=o[1].replace("/>",">"); }
	return o.join("");
}
/* [MS-XLSB] 2.4.304 BrtBundleSh */
function parse_BrtBundleSh(data, length) {
	var z = {};
	z.Hidden = data.read_shift(4); //hsState ST_SheetState
	z.iTabID = data.read_shift(4);
	z.strRelID = parse_RelID(data,length-8);
	z.name = parse_XLWideString(data);
	return z;
}
function write_BrtBundleSh(data, o) {
	if(!o) o = new_buf(127);
	o.write_shift(4, data.Hidden);
	o.write_shift(4, data.iTabID);
	write_RelID(data.strRelID, o);
	write_XLWideString(data.name.slice(0,31), o);
	return o.length > o.l ? o.slice(0, o.l) : o;
}

/* [MS-XLSB] 2.4.815 BrtWbProp */
function parse_BrtWbProp(data, length) {
	var o = ({});
	var flags = data.read_shift(4);
	o.defaultThemeVersion = data.read_shift(4);
	var strName = (length > 8) ? parse_XLWideString(data) : "";
	if(strName.length > 0) o.CodeName = strName;
	o.autoCompressPictures = !!(flags & 0x10000);
	o.backupFile = !!(flags & 0x40);
	o.checkCompatibility = !!(flags & 0x1000);
	o.date1904 = !!(flags & 0x01);
	o.filterPrivacy = !!(flags & 0x08);
	o.hidePivotFieldList = !!(flags & 0x400);
	o.promptedSolutions = !!(flags & 0x10);
	o.publishItems = !!(flags & 0x800);
	o.refreshAllConnections = !!(flags & 0x40000);
	o.saveExternalLinkValues = !!(flags & 0x80);
	o.showBorderUnselectedTables = !!(flags & 0x04);
	o.showInkAnnotation = !!(flags & 0x20);
	o.showObjects = ["all", "placeholders", "none"][(flags >> 13) & 0x03];
	o.showPivotChartFilter = !!(flags & 0x8000);
	o.updateLinks = ["userSet", "never", "always"][(flags >> 8) & 0x03];
	return o;
}
function write_BrtWbProp(data, o) {
	if(!o) o = new_buf(72);
	var flags = 0;
	if(data) {
		/* TODO: mirror parse_BrtWbProp fields */
		if(data.filterPrivacy) flags |= 0x08;
	}
	o.write_shift(4, flags);
	o.write_shift(4, 0);
	write_XLSBCodeName(data && data.CodeName || "ThisWorkbook", o);
	return o.slice(0, o.l);
}

function parse_BrtFRTArchID$(data, length) {
	var o = {};
	data.read_shift(4);
	o.ArchID = data.read_shift(4);
	data.l += length - 8;
	return o;
}

/* [MS-XLSB] 2.4.687 BrtName */
function parse_BrtName(data, length, opts) {
	var end = data.l + length;
	data.l += 4; //var flags = data.read_shift(4);
	data.l += 1; //var chKey = data.read_shift(1);
	var itab = data.read_shift(4);
	var name = parse_XLNameWideString(data);
	var formula = parse_XLSBNameParsedFormula(data, 0, opts);
	var comment = parse_XLNullableWideString(data);
	//if(0 /* fProc */) {
		// unusedstring1: XLNullableWideString
		// description: XLNullableWideString
		// helpTopic: XLNullableWideString
		// unusedstring2: XLNullableWideString
	//}
	data.l = end;
	var out = ({Name:name, Ptg:formula});
	if(itab < 0xFFFFFFF) out.Sheet = itab;
	if(comment) out.Comment = comment;
	return out;
}

/* [MS-XLSB] 2.1.7.61 Workbook */
function parse_wb_bin(data, opts) {
	var wb = { AppVersion:{}, WBProps:{}, WBView:[], Sheets:[], CalcPr:{}, xmlns: "" };
	var state = [];
	var pass = false;

	if(!opts) opts = {};
	opts.biff = 12;

	var Names = [];
	var supbooks = ([[]]);
	supbooks.SheetNames = [];
	supbooks.XTI = [];

	recordhopper(data, function hopper_wb(val, R_n, RT) {
		switch(RT) {
			case 0x009C: /* 'BrtBundleSh' */
				supbooks.SheetNames.push(val.name);
				wb.Sheets.push(val); break;

			case 0x0099: /* 'BrtWbProp' */
				wb.WBProps = val; break;

			case 0x0027: /* 'BrtName' */
				if(val.Sheet != null) opts.SID = val.Sheet;
				val.Ref = stringify_formula(val.Ptg, null, null, supbooks, opts);
				delete opts.SID;
				delete val.Ptg;
				Names.push(val);
				break;
			case 0x040C: /* 'BrtNameExt' */ break;

			case 0x0165: /* 'BrtSupSelf' */
			case 0x0166: /* 'BrtSupSame' */
			case 0x0163: /* 'BrtSupBookSrc' */
			case 0x029B: /* 'BrtSupAddin' */
				if(!supbooks[0].length) supbooks[0] = [RT, val];
				else supbooks.push([RT, val]);
				supbooks[supbooks.length - 1].XTI = [];
				break;
			case 0x016A: /* 'BrtExternSheet' */
				if(supbooks.length === 0) { supbooks[0] = []; supbooks[0].XTI = []; }
				supbooks[supbooks.length - 1].XTI = supbooks[supbooks.length - 1].XTI.concat(val);
				supbooks.XTI = supbooks.XTI.concat(val);
				break;
			case 0x0169: /* 'BrtPlaceholderName' */
				break;

			/* case 'BrtModelTimeGroupingCalcCol' */
			case 0x0C00: /* 'BrtUid' */
			case 0x0C01: /* 'BrtRevisionPtr' */
			case 0x0817: /* 'BrtAbsPath15' */
			case 0x0216: /* 'BrtBookProtection' */
			case 0x02A5: /* 'BrtBookProtectionIso' */
			case 0x009E: /* 'BrtBookView' */
			case 0x009D: /* 'BrtCalcProp' */
			case 0x0262: /* 'BrtCrashRecErr' */
			case 0x0802: /* 'BrtDecoupledPivotCacheID' */
			case 0x009B: /* 'BrtFileRecover' */
			case 0x0224: /* 'BrtFileSharing' */
			case 0x02A4: /* 'BrtFileSharingIso' */
			case 0x0080: /* 'BrtFileVersion' */
			case 0x0299: /* 'BrtFnGroup' */
			case 0x0850: /* 'BrtModelRelationship' */
			case 0x084D: /* 'BrtModelTable' */
			case 0x0225: /* 'BrtOleSize' */
			case 0x0805: /* 'BrtPivotTableRef' */
			case 0x0254: /* 'BrtSmartTagType' */
			case 0x081C: /* 'BrtTableSlicerCacheID' */
			case 0x081B: /* 'BrtTableSlicerCacheIDs' */
			case 0x0822: /* 'BrtTimelineCachePivotCacheID' */
			case 0x018D: /* 'BrtUserBookView' */
			case 0x009A: /* 'BrtWbFactoid' */
			case 0x045D: /* 'BrtWbProp14' */
			case 0x0229: /* 'BrtWebOpt' */
			case 0x082B: /* 'BrtWorkBookPr15' */
				break;

			case 0x0023: /* 'BrtFRTBegin' */
				state.push(R_n); pass = true; break;
			case 0x0024: /* 'BrtFRTEnd' */
				state.pop(); pass = false; break;
			case 0x0025: /* 'BrtACBegin' */
				state.push(R_n); pass = true; break;
			case 0x0026: /* 'BrtACEnd' */
				state.pop(); pass = false; break;

			case 0x0010: /* 'BrtFRTArchID$' */ break;

			default:
				if((R_n||"").indexOf("Begin") > 0){/* empty */}
				else if((R_n||"").indexOf("End") > 0){/* empty */}
				else if(!pass || (opts.WTF && state[state.length-1] != "BrtACBegin" && state[state.length-1] != "BrtFRTBegin")) throw new Error("Unexpected record " + RT + " " + R_n);
		}
	}, opts);

	parse_wb_defaults(wb);

	// $FlowIgnore
	wb.Names = Names;

	(wb).supbooks = supbooks;
	return wb;
}

function write_BUNDLESHS(ba, wb) {
	write_record(ba, "BrtBeginBundleShs");
	for(var idx = 0; idx != wb.SheetNames.length; ++idx) {
		var viz = wb.Workbook && wb.Workbook.Sheets && wb.Workbook.Sheets[idx] && wb.Workbook.Sheets[idx].Hidden || 0;
		var d = { Hidden: viz, iTabID: idx+1, strRelID: 'rId' + (idx+1), name: wb.SheetNames[idx] };
		write_record(ba, "BrtBundleSh", write_BrtBundleSh(d));
	}
	write_record(ba, "BrtEndBundleShs");
}

/* [MS-XLSB] 2.4.649 BrtFileVersion */
function write_BrtFileVersion(data, o) {
	if(!o) o = new_buf(127);
	for(var i = 0; i != 4; ++i) o.write_shift(4, 0);
	write_XLWideString("SheetJS", o);
	write_XLWideString(XLSX.version, o);
	write_XLWideString(XLSX.version, o);
	write_XLWideString("7262", o);
	o.length = o.l;
	return o.length > o.l ? o.slice(0, o.l) : o;
}

/* [MS-XLSB] 2.4.301 BrtBookView */
function write_BrtBookView(idx, o) {
	if(!o) o = new_buf(29);
	o.write_shift(-4, 0);
	o.write_shift(-4, 460);
	o.write_shift(4,  28800);
	o.write_shift(4,  17600);
	o.write_shift(4,  500);
	o.write_shift(4,  idx);
	o.write_shift(4,  idx);
	var flags = 0x78;
	o.write_shift(1,  flags);
	return o.length > o.l ? o.slice(0, o.l) : o;
}

function write_BOOKVIEWS(ba, wb) {
	/* required if hidden tab appears before visible tab */
	if(!wb.Workbook || !wb.Workbook.Sheets) return;
	var sheets = wb.Workbook.Sheets;
	var i = 0, vistab = -1, hidden = -1;
	for(; i < sheets.length; ++i) {
		if(!sheets[i] || !sheets[i].Hidden && vistab == -1) vistab = i;
		else if(sheets[i].Hidden == 1 && hidden == -1) hidden = i;
	}
	if(hidden > vistab) return;
	write_record(ba, "BrtBeginBookViews");
	write_record(ba, "BrtBookView", write_BrtBookView(vistab));
	/* 1*(BrtBookView *FRT) */
	write_record(ba, "BrtEndBookViews");
}

/* [MS-XLSB] 2.4.305 BrtCalcProp */
/*function write_BrtCalcProp(data, o) {
	if(!o) o = new_buf(26);
	o.write_shift(4,0); // force recalc
	o.write_shift(4,1);
	o.write_shift(4,0);
	write_Xnum(0, o);
	o.write_shift(-4, 1023);
	o.write_shift(1, 0x33);
	o.write_shift(1, 0x00);
	return o;
}*/

/* [MS-XLSB] 2.4.646 BrtFileRecover */
/*function write_BrtFileRecover(data, o) {
	if(!o) o = new_buf(1);
	o.write_shift(1,0);
	return o;
}*/

/* [MS-XLSB] 2.1.7.61 Workbook */
function write_wb_bin(wb, opts) {
	var ba = buf_array();
	write_record(ba, "BrtBeginBook");
	write_record(ba, "BrtFileVersion", write_BrtFileVersion());
	/* [[BrtFileSharingIso] BrtFileSharing] */
	write_record(ba, "BrtWbProp", write_BrtWbProp(wb.Workbook && wb.Workbook.WBProps || null));
	/* [ACABSPATH] */
	/* [[BrtBookProtectionIso] BrtBookProtection] */
	write_BOOKVIEWS(ba, wb, opts);
	write_BUNDLESHS(ba, wb, opts);
	/* [FNGROUP] */
	/* [EXTERNALS] */
	/* *BrtName */
	/* write_record(ba, "BrtCalcProp", write_BrtCalcProp()); */
	/* [BrtOleSize] */
	/* *(BrtUserBookView *FRT) */
	/* [PIVOTCACHEIDS] */
	/* [BrtWbFactoid] */
	/* [SMARTTAGTYPES] */
	/* [BrtWebOpt] */
	/* write_record(ba, "BrtFileRecover", write_BrtFileRecover()); */
	/* [WEBPUBITEMS] */
	/* [CRERRS] */
	/* FRTWORKBOOK */
	write_record(ba, "BrtEndBook");

	return ba.end();
}
function parse_wb(data, name, opts) {
	if(name.slice(-4)===".bin") return parse_wb_bin((data), opts);
	return parse_wb_xml((data), opts);
}

function parse_ws(data, name, idx, opts, rels, wb, themes, styles) {
	if(name.slice(-4)===".bin") return parse_ws_bin((data), opts, idx, rels, wb, themes, styles);
	return parse_ws_xml((data), opts, idx, rels, wb, themes, styles);
}

function parse_cs(data, name, idx, opts, rels, wb, themes, styles) {
	if(name.slice(-4)===".bin") return parse_cs_bin((data), opts, idx, rels, wb, themes, styles);
	return parse_cs_xml((data), opts, idx, rels, wb, themes, styles);
}

function parse_ms(data, name, idx, opts, rels, wb, themes, styles) {
	if(name.slice(-4)===".bin") return parse_ms_bin((data), opts, idx, rels, wb, themes, styles);
	return parse_ms_xml((data), opts, idx, rels, wb, themes, styles);
}

function parse_ds(data, name, idx, opts, rels, wb, themes, styles) {
	if(name.slice(-4)===".bin") return parse_ds_bin((data), opts, idx, rels, wb, themes, styles);
	return parse_ds_xml((data), opts, idx, rels, wb, themes, styles);
}

function parse_sty(data, name, themes, opts) {
	if(name.slice(-4)===".bin") return parse_sty_bin((data), themes, opts);
	return parse_sty_xml((data), themes, opts);
}

function parse_theme(data, name, opts) {
	return parse_theme_xml(data, opts);
}

function parse_sst(data, name, opts) {
	if(name.slice(-4)===".bin") return parse_sst_bin((data), opts);
	return parse_sst_xml((data), opts);
}

function parse_cmnt(data, name, opts) {
	if(name.slice(-4)===".bin") return parse_comments_bin((data), opts);
	return parse_comments_xml((data), opts);
}

function parse_cc(data, name, opts) {
	if(name.slice(-4)===".bin") return parse_cc_bin((data), name, opts);
	return parse_cc_xml((data), name, opts);
}

function parse_xlink(data, name, opts) {
	if(name.slice(-4)===".bin") return parse_xlink_bin((data), name, opts);
	return parse_xlink_xml((data), name, opts);
}

function write_wb(wb, name, opts) {
	return (name.slice(-4)===".bin" ? write_wb_bin : write_wb_xml)(wb, opts);
}

function write_ws(data, name, opts, wb, rels) {
	return (name.slice(-4)===".bin" ? write_ws_bin : write_ws_xml)(data, opts, wb, rels);
}

// eslint-disable-next-line no-unused-vars
function write_cs(data, name, opts, wb, rels) {
	return (name.slice(-4)===".bin" ? write_cs_bin : write_cs_xml)(data, opts, wb, rels);
}

function write_sty(data, name, opts) {
	return (name.slice(-4)===".bin" ? write_sty_bin : write_sty_xml)(data, opts);
}

function write_sst(data, name, opts) {
	return (name.slice(-4)===".bin" ? write_sst_bin : write_sst_xml)(data, opts);
}

function write_cmnt(data, name, opts) {
	return (name.slice(-4)===".bin" ? write_comments_bin : write_comments_xml)(data, opts);
}
/*
function write_cc(data, name:string, opts) {
	return (name.slice(-4)===".bin" ? write_cc_bin : write_cc_xml)(data, opts);
}
*/
var attregexg2=/([\w:]+)=((?:")([^"]*)(?:")|(?:')([^']*)(?:'))/g;
var attregex2=/([\w:]+)=((?:")(?:[^"]*)(?:")|(?:')(?:[^']*)(?:'))/;
var _chr = function(c) { return String.fromCharCode(c); };
function xlml_parsexmltag(tag, skip_root) {
	var words = tag.split(/\s+/);
	var z = ([]); if(!skip_root) z[0] = words[0];
	if(words.length === 1) return z;
	var m = tag.match(attregexg2), y, j, w, i;
	if(m) for(i = 0; i != m.length; ++i) {
		y = m[i].match(attregex2);
if((j=y[1].indexOf(":")) === -1) z[y[1]] = y[2].slice(1,y[2].length-1);
		else {
			if(y[1].slice(0,6) === "xmlns:") w = "xmlns"+y[1].slice(6);
			else w = y[1].slice(j+1);
			z[w] = y[2].slice(1,y[2].length-1);
		}
	}
	return z;
}
function xlml_parsexmltagobj(tag) {
	var words = tag.split(/\s+/);
	var z = {};
	if(words.length === 1) return z;
	var m = tag.match(attregexg2), y, j, w, i;
	if(m) for(i = 0; i != m.length; ++i) {
		y = m[i].match(attregex2);
if((j=y[1].indexOf(":")) === -1) z[y[1]] = y[2].slice(1,y[2].length-1);
		else {
			if(y[1].slice(0,6) === "xmlns:") w = "xmlns"+y[1].slice(6);
			else w = y[1].slice(j+1);
			z[w] = y[2].slice(1,y[2].length-1);
		}
	}
	return z;
}

// ----

function xlml_format(format, value) {
	var fmt = XLMLFormatMap[format] || unescapexml(format);
	if(fmt === "General") return SSF._general(value);
	return SSF.format(fmt, value);
}

function xlml_set_custprop(Custprops, key, cp, val) {
	var oval = val;
	switch((cp[0].match(/dt:dt="([\w.]+)"/)||["",""])[1]) {
		case "boolean": oval = parsexmlbool(val); break;
		case "i2": case "int": oval = parseInt(val, 10); break;
		case "r4": case "float": oval = parseFloat(val); break;
		case "date": case "dateTime.tz": oval = parseDate(val); break;
		case "i8": case "string": case "fixed": case "uuid": case "bin.base64": break;
		default: throw new Error("bad custprop:" + cp[0]);
	}
	Custprops[unescapexml(key)] = oval;
}

function safe_format_xlml(cell, nf, o) {
	if(cell.t === 'z') return;
	if(!o || o.cellText !== false) try {
		if(cell.t === 'e') { cell.w = cell.w || BErr[cell.v]; }
		else if(nf === "General") {
			if(cell.t === 'n') {
				if((cell.v|0) === cell.v) cell.w = SSF._general_int(cell.v);
				else cell.w = SSF._general_num(cell.v);
			}
			else cell.w = SSF._general(cell.v);
		}
		else cell.w = xlml_format(nf||"General", cell.v);
	} catch(e) { if(o.WTF) throw e; }
	try {
		var z = XLMLFormatMap[nf]||nf||"General";
		if(o.cellNF) cell.z = z;
		if(o.cellDates && cell.t == 'n' && SSF.is_date(z)) {
			var _d = SSF.parse_date_code(cell.v); if(_d) { cell.t = 'd'; cell.v = new Date(_d.y, _d.m-1,_d.d,_d.H,_d.M,_d.S,_d.u); }
		}
	} catch(e) { if(o.WTF) throw e; }
}

function process_style_xlml(styles, stag, opts) {
	if(opts.cellStyles) {
		if(stag.Interior) {
			var I = stag.Interior;
			if(I.Pattern) I.patternType = XLMLPatternTypeMap[I.Pattern] || I.Pattern;
		}
	}
	styles[stag.ID] = stag;
}

/* TODO: there must exist some form of OSP-blessed spec */
function parse_xlml_data(xml, ss, data, cell, base, styles, csty, row, arrayf, o) {
	var nf = "General", sid = cell.StyleID, S = {}; o = o || {};
	var interiors = [];
	var i = 0;
	if(sid === undefined && row) sid = row.StyleID;
	if(sid === undefined && csty) sid = csty.StyleID;
	while(styles[sid] !== undefined) {
		if(styles[sid].nf) nf = styles[sid].nf;
		if(styles[sid].Interior) interiors.push(styles[sid].Interior);
		if(!styles[sid].Parent) break;
		sid = styles[sid].Parent;
	}
	switch(data.Type) {
		case 'Boolean':
			cell.t = 'b';
			cell.v = parsexmlbool(xml);
			break;
		case 'String':
			cell.t = 's'; cell.r = xlml_fixstr(unescapexml(xml));
			cell.v = xml.indexOf("<") > -1 ? unescapexml(ss) : cell.r;
			break;
		case 'DateTime':
			if(xml.slice(-1) != "Z") xml += "Z";
			cell.v = (parseDate(xml) - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
			if(cell.v !== cell.v) cell.v = unescapexml(xml);
			else if(cell.v<60) cell.v = cell.v -1;
			if(!nf || nf == "General") nf = "yyyy-mm-dd";
			/* falls through */
		case 'Number':
			if(cell.v === undefined) cell.v=+xml;
			if(!cell.t) cell.t = 'n';
			break;
		case 'Error': cell.t = 'e'; cell.v = RBErr[xml]; if(o.cellText !== false) cell.w = xml; break;
		default: cell.t = 's'; cell.v = xlml_fixstr(ss||xml); break;
	}
	safe_format_xlml(cell, nf, o);
	if(o.cellFormula !== false) {
		if(cell.Formula) {
			var fstr = unescapexml(cell.Formula);
			/* strictly speaking, the leading = is required but some writers omit */
			if(fstr.charCodeAt(0) == 61 /* = */) fstr = fstr.slice(1);
			cell.f = rc_to_a1(fstr, base);
			delete cell.Formula;
			if(cell.ArrayRange == "RC") cell.F = rc_to_a1("RC:RC", base);
			else if(cell.ArrayRange) {
				cell.F = rc_to_a1(cell.ArrayRange, base);
				arrayf.push([safe_decode_range(cell.F), cell.F]);
			}
		} else {
			for(i = 0; i < arrayf.length; ++i)
				if(base.r >= arrayf[i][0].s.r && base.r <= arrayf[i][0].e.r)
					if(base.c >= arrayf[i][0].s.c && base.c <= arrayf[i][0].e.c)
						cell.F = arrayf[i][1];
		}
	}
	if(o.cellStyles) {
		interiors.forEach(function(x) {
			if(!S.patternType && x.patternType) S.patternType = x.patternType;
		});
		cell.s = S;
	}
	if(cell.StyleID !== undefined) cell.ixfe = cell.StyleID;
}

function xlml_clean_comment(comment) {
	comment.t = comment.v || "";
	comment.t = comment.t.replace(/\r\n/g,"\n").replace(/\r/g,"\n");
	comment.v = comment.w = comment.ixfe = undefined;
}

function xlml_normalize(d) {
	if(has_buf && Buffer.isBuffer(d)) return d.toString('utf8');
	if(typeof d === 'string') return d;
	/* duktape */
	if(typeof Uint8Array !== 'undefined' && d instanceof Uint8Array) return utf8read(a2s(ab2a(d)));
	throw new Error("Bad input format: expected Buffer or string");
}

/* TODO: Everything */
/* UOS uses CJK in tags */
var xlmlregex = /<(\/?)([^\s?>!\/:]*:|)([^\s?>:\/]+)[^>]*>/mg;
//var xlmlregex = /<(\/?)([a-z0-9]*:|)(\w+)[^>]*>/mg;
function parse_xlml_xml(d, _opts) {
	var opts = _opts || {};
	make_ssf(SSF);
	var str = debom(xlml_normalize(d));
	if(opts.type == 'binary' || opts.type == 'array' || opts.type == 'base64') {
		if(typeof cptable !== 'undefined') str = cptable.utils.decode(65001, char_codes(str));
		else str = utf8read(str);
	}
	var opening = str.slice(0, 1024).toLowerCase(), ishtml = false;
	if(opening.indexOf("<?xml") == -1) ["html", "table", "head", "meta", "script", "style", "div"].forEach(function(tag) { if(opening.indexOf("<" + tag) >= 0) ishtml = true; });
	if(ishtml) return HTML_.to_workbook(str, opts);
	var Rn;
	var state = [], tmp;
	if(DENSE != null && opts.dense == null) opts.dense = DENSE;
	var sheets = {}, sheetnames = [], cursheet = (opts.dense ? [] : {}), sheetname = "";
	var table = {}, cell = ({}), row = {};// eslint-disable-line no-unused-vars
	var dtag = xlml_parsexmltag('<Data ss:Type="String">'), didx = 0;
	var c = 0, r = 0;
	var refguess = {s: {r:2000000, c:2000000}, e: {r:0, c:0} };
	var styles = {}, stag = {};
	var ss = "", fidx = 0;
	var merges = [];
	var Props = {}, Custprops = {}, pidx = 0, cp = [];
	var comments = [], comment = ({});
	var cstys = [], csty, seencol = false;
	var arrayf = [];
	var rowinfo = [], rowobj = {}, cc = 0, rr = 0;
	var Workbook = ({ Sheets:[], WBProps:{date1904:false} }), wsprops = {};
	xlmlregex.lastIndex = 0;
	str = str.replace(/<!--([\s\S]*?)-->/mg,"");
	while((Rn = xlmlregex.exec(str))) switch(Rn[3]) {
		case 'Data':
			if(state[state.length-1][1]) break;
			if(Rn[1]==='/') parse_xlml_data(str.slice(didx, Rn.index), ss, dtag, state[state.length-1][0]=="Comment"?comment:cell, {c:c,r:r}, styles, cstys[c], row, arrayf, opts);
			else { ss = ""; dtag = xlml_parsexmltag(Rn[0]); didx = Rn.index + Rn[0].length; }
			break;
		case 'Cell':
			if(Rn[1]==='/'){
				if(comments.length > 0) cell.c = comments;
				if((!opts.sheetRows || opts.sheetRows > r) && cell.v !== undefined) {
					if(opts.dense) {
						if(!cursheet[r]) cursheet[r] = [];
						cursheet[r][c] = cell;
					} else cursheet[encode_col(c) + encode_row(r)] = cell;
				}
				if(cell.HRef) {
					cell.l = ({Target:cell.HRef});
					if(cell.HRefScreenTip) cell.l.Tooltip = cell.HRefScreenTip;
					delete cell.HRef; delete cell.HRefScreenTip;
				}
				if(cell.MergeAcross || cell.MergeDown) {
					cc = c + (parseInt(cell.MergeAcross,10)|0);
					rr = r + (parseInt(cell.MergeDown,10)|0);
					merges.push({s:{c:c,r:r},e:{c:cc,r:rr}});
				}
				if(!opts.sheetStubs) { if(cell.MergeAcross) c = cc + 1; else ++c; }
				else if(cell.MergeAcross || cell.MergeDown) {
for(var cma = c; cma <= cc; ++cma) {
						for(var cmd = r; cmd <= rr; ++cmd) {
							if(cma > c || cmd > r) {
								if(opts.dense) {
									if(!cursheet[cmd]) cursheet[cmd] = [];
									cursheet[cmd][cma] = {t:'z'};
								} else cursheet[encode_col(cma) + encode_row(cmd)] = {t:'z'};
							}
						}
					}
					c = cc + 1;
				}
				else ++c;
			} else {
				cell = xlml_parsexmltagobj(Rn[0]);
				if(cell.Index) c = +cell.Index - 1;
				if(c < refguess.s.c) refguess.s.c = c;
				if(c > refguess.e.c) refguess.e.c = c;
				if(Rn[0].slice(-2) === "/>") ++c;
				comments = [];
			}
			break;
		case 'Row':
			if(Rn[1]==='/' || Rn[0].slice(-2) === "/>") {
				if(r < refguess.s.r) refguess.s.r = r;
				if(r > refguess.e.r) refguess.e.r = r;
				if(Rn[0].slice(-2) === "/>") {
					row = xlml_parsexmltag(Rn[0]);
					if(row.Index) r = +row.Index - 1;
				}
				c = 0; ++r;
			} else {
				row = xlml_parsexmltag(Rn[0]);
				if(row.Index) r = +row.Index - 1;
				rowobj = {};
				if(row.AutoFitHeight == "0" || row.Height) {
					rowobj.hpx = parseInt(row.Height, 10); rowobj.hpt = px2pt(rowobj.hpx);
					rowinfo[r] = rowobj;
				}
				if(row.Hidden == "1") { rowobj.hidden = true; rowinfo[r] = rowobj; }
			}
			break;
		case 'Worksheet': /* TODO: read range from FullRows/FullColumns */
			if(Rn[1]==='/'){
				if((tmp=state.pop())[0]!==Rn[3]) throw new Error("Bad state: "+tmp.join("|"));
				sheetnames.push(sheetname);
				if(refguess.s.r <= refguess.e.r && refguess.s.c <= refguess.e.c) {
					cursheet["!ref"] = encode_range(refguess);
					if(opts.sheetRows && opts.sheetRows <= refguess.e.r) {
						cursheet["!fullref"] = cursheet["!ref"];
						refguess.e.r = opts.sheetRows - 1;
						cursheet["!ref"] = encode_range(refguess);
					}
				}
				if(merges.length) cursheet["!merges"] = merges;
				if(cstys.length > 0) cursheet["!cols"] = cstys;
				if(rowinfo.length > 0) cursheet["!rows"] = rowinfo;
				sheets[sheetname] = cursheet;
			} else {
				refguess = {s: {r:2000000, c:2000000}, e: {r:0, c:0} };
				r = c = 0;
				state.push([Rn[3], false]);
				tmp = xlml_parsexmltag(Rn[0]);
				sheetname = unescapexml(tmp.Name);
				cursheet = (opts.dense ? [] : {});
				merges = [];
				arrayf = [];
				rowinfo = [];
				wsprops = {name:sheetname, Hidden:0};
				Workbook.Sheets.push(wsprops);
			}
			break;
		case 'Table':
			if(Rn[1]==='/'){if((tmp=state.pop())[0]!==Rn[3]) throw new Error("Bad state: "+tmp.join("|"));}
			else if(Rn[0].slice(-2) == "/>") break;
			else {
				table = xlml_parsexmltag(Rn[0]);
				state.push([Rn[3], false]);
				cstys = []; seencol = false;
			}
			break;

		case 'Style':
			if(Rn[1]==='/') process_style_xlml(styles, stag, opts);
			else stag = xlml_parsexmltag(Rn[0]);
			break;

		case 'NumberFormat':
			stag.nf = unescapexml(xlml_parsexmltag(Rn[0]).Format || "General");
			if(XLMLFormatMap[stag.nf]) stag.nf = XLMLFormatMap[stag.nf];
			for(var ssfidx = 0; ssfidx != 0x188; ++ssfidx) if(SSF._table[ssfidx] == stag.nf) break;
			if(ssfidx == 0x188) for(ssfidx = 0x39; ssfidx != 0x188; ++ssfidx) if(SSF._table[ssfidx] == null) { SSF.load(stag.nf, ssfidx); break; }
			break;

		case 'Column':
			if(state[state.length-1][0] !== 'Table') break;
			csty = xlml_parsexmltag(Rn[0]);
			if(csty.Hidden) { csty.hidden = true; delete csty.Hidden; }
			if(csty.Width) csty.wpx = parseInt(csty.Width, 10);
			if(!seencol && csty.wpx > 10) {
				seencol = true; MDW = DEF_MDW; //find_mdw_wpx(csty.wpx);
				for(var _col = 0; _col < cstys.length; ++_col) if(cstys[_col]) process_col(cstys[_col]);
			}
			if(seencol) process_col(csty);
			cstys[(csty.Index-1||cstys.length)] = csty;
			for(var i = 0; i < +csty.Span; ++i) cstys[cstys.length] = dup(csty);
			break;

		case 'NamedRange':
			if(!Workbook.Names) Workbook.Names = [];
			var _NamedRange = parsexmltag(Rn[0]);
			var _DefinedName = ({
				Name: _NamedRange.Name,
				Ref: rc_to_a1(_NamedRange.RefersTo.slice(1), {r:0, c:0})
			});
			if(Workbook.Sheets.length>0) _DefinedName.Sheet=Workbook.Sheets.length-1;
Workbook.Names.push(_DefinedName);
			break;

		case 'NamedCell': break;
		case 'B': break;
		case 'I': break;
		case 'U': break;
		case 'S': break;
		case 'Sub': break;
		case 'Sup': break;
		case 'Span': break;
		case 'Border': break;
		case 'Alignment': break;
		case 'Borders': break;
		case 'Font':
			if(Rn[0].slice(-2) === "/>") break;
			else if(Rn[1]==="/") ss += str.slice(fidx, Rn.index);
			else fidx = Rn.index + Rn[0].length;
			break;
		case 'Interior':
			if(!opts.cellStyles) break;
			stag.Interior = xlml_parsexmltag(Rn[0]);
			break;
		case 'Protection': break;

		case 'Author':
		case 'Title':
		case 'Description':
		case 'Created':
		case 'Keywords':
		case 'Subject':
		case 'Category':
		case 'Company':
		case 'LastAuthor':
		case 'LastSaved':
		case 'LastPrinted':
		case 'Version':
		case 'Revision':
		case 'TotalTime':
		case 'HyperlinkBase':
		case 'Manager':
		case 'ContentStatus':
		case 'Identifier':
		case 'Language':
		case 'AppName':
			if(Rn[0].slice(-2) === "/>") break;
			else if(Rn[1]==="/") xlml_set_prop(Props, Rn[3], str.slice(pidx, Rn.index));
			else pidx = Rn.index + Rn[0].length;
			break;
		case 'Paragraphs': break;

		case 'Styles':
		case 'Workbook':
			if(Rn[1]==='/'){if((tmp=state.pop())[0]!==Rn[3]) throw new Error("Bad state: "+tmp.join("|"));}
			else state.push([Rn[3], false]);
			break;

		case 'Comment':
			if(Rn[1]==='/'){
				if((tmp=state.pop())[0]!==Rn[3]) throw new Error("Bad state: "+tmp.join("|"));
				xlml_clean_comment(comment);
				comments.push(comment);
			} else {
				state.push([Rn[3], false]);
				tmp = xlml_parsexmltag(Rn[0]);
				comment = ({a:tmp.Author});
			}
			break;

		case 'AutoFilter':
			if(Rn[1]==='/'){if((tmp=state.pop())[0]!==Rn[3]) throw new Error("Bad state: "+tmp.join("|"));}
			else if(Rn[0].charAt(Rn[0].length-2) !== '/') {
				var AutoFilter = xlml_parsexmltag(Rn[0]);
				cursheet['!autofilter'] = { ref:rc_to_a1(AutoFilter.Range).replace(/\$/g,"") };
				state.push([Rn[3], true]);
			}
			break;

		case 'Name': break;

		case 'ComponentOptions':
		case 'DocumentProperties':
		case 'CustomDocumentProperties':
		case 'OfficeDocumentSettings':
		case 'PivotTable':
		case 'PivotCache':
		case 'Names':
		case 'MapInfo':
		case 'PageBreaks':
		case 'QueryTable':
		case 'DataValidation':
		case 'Sorting':
		case 'Schema':
		case 'data':
		case 'ConditionalFormatting':
		case 'SmartTagType':
		case 'SmartTags':
		case 'ExcelWorkbook':
		case 'WorkbookOptions':
		case 'WorksheetOptions':
			if(Rn[1]==='/'){if((tmp=state.pop())[0]!==Rn[3]) throw new Error("Bad state: "+tmp.join("|"));}
			else if(Rn[0].charAt(Rn[0].length-2) !== '/') state.push([Rn[3], true]);
			break;

		default:
			/* FODS file root is <office:document> */
			if(state.length == 0 && Rn[3] == "document") return parse_fods(str, opts);
			/* UOS file root is <uof:UOF> */
			if(state.length == 0 && Rn[3] == "UOF") return parse_fods(str, opts);

			var seen = true;
			switch(state[state.length-1][0]) {
				/* OfficeDocumentSettings */
				case 'OfficeDocumentSettings': switch(Rn[3]) {
					case 'AllowPNG': break;
					case 'RemovePersonalInformation': break;
					case 'DownloadComponents': break;
					case 'LocationOfComponents': break;
					case 'Colors': break;
					case 'Color': break;
					case 'Index': break;
					case 'RGB': break;
					case 'PixelsPerInch': break; // TODO: set PPI
					case 'TargetScreenSize': break;
					case 'ReadOnlyRecommended': break;
					default: seen = false;
				} break;

				/* ComponentOptions */
				case 'ComponentOptions': switch(Rn[3]) {
					case 'Toolbar': break;
					case 'HideOfficeLogo': break;
					case 'SpreadsheetAutoFit': break;
					case 'Label': break;
					case 'Caption': break;
					case 'MaxHeight': break;
					case 'MaxWidth': break;
					case 'NextSheetNumber': break;
					default: seen = false;
				} break;

				/* ExcelWorkbook */
				case 'ExcelWorkbook': switch(Rn[3]) {
					case 'Date1904':
Workbook.WBProps.date1904 = true;
						break;
					case 'WindowHeight': break;
					case 'WindowWidth': break;
					case 'WindowTopX': break;
					case 'WindowTopY': break;
					case 'TabRatio': break;
					case 'ProtectStructure': break;
					case 'ProtectWindows': break;
					case 'ActiveSheet': break;
					case 'DisplayInkNotes': break;
					case 'FirstVisibleSheet': break;
					case 'SupBook': break;
					case 'SheetName': break;
					case 'SheetIndex': break;
					case 'SheetIndexFirst': break;
					case 'SheetIndexLast': break;
					case 'Dll': break;
					case 'AcceptLabelsInFormulas': break;
					case 'DoNotSaveLinkValues': break;
					case 'Iteration': break;
					case 'MaxIterations': break;
					case 'MaxChange': break;
					case 'Path': break;
					case 'Xct': break;
					case 'Count': break;
					case 'SelectedSheets': break;
					case 'Calculation': break;
					case 'Uncalced': break;
					case 'StartupPrompt': break;
					case 'Crn': break;
					case 'ExternName': break;
					case 'Formula': break;
					case 'ColFirst': break;
					case 'ColLast': break;
					case 'WantAdvise': break;
					case 'Boolean': break;
					case 'Error': break;
					case 'Text': break;
					case 'OLE': break;
					case 'NoAutoRecover': break;
					case 'PublishObjects': break;
					case 'DoNotCalculateBeforeSave': break;
					case 'Number': break;
					case 'RefModeR1C1': break;
					case 'EmbedSaveSmartTags': break;
					default: seen = false;
				} break;

				/* WorkbookOptions */
				case 'WorkbookOptions': switch(Rn[3]) {
					case 'OWCVersion': break;
					case 'Height': break;
					case 'Width': break;
					default: seen = false;
				} break;

				/* WorksheetOptions */
				case 'WorksheetOptions': switch(Rn[3]) {
					case 'Visible':
						if(Rn[0].slice(-2) === "/>"){/* empty */}
						else if(Rn[1]==="/") switch(str.slice(pidx, Rn.index)) {
							case "SheetHidden": wsprops.Hidden = 1; break;
							case "SheetVeryHidden": wsprops.Hidden = 2; break;
						}
						else pidx = Rn.index + Rn[0].length;
						break;
					case 'Header':
						if(!cursheet['!margins']) default_margins(cursheet['!margins']={}, 'xlml');
						cursheet['!margins'].header = parsexmltag(Rn[0]).Margin;
						break;
					case 'Footer':
						if(!cursheet['!margins']) default_margins(cursheet['!margins']={}, 'xlml');
						cursheet['!margins'].footer = parsexmltag(Rn[0]).Margin;
						break;
					case 'PageMargins':
						var pagemargins = parsexmltag(Rn[0]);
						if(!cursheet['!margins']) default_margins(cursheet['!margins']={},'xlml');
						if(pagemargins.Top) cursheet['!margins'].top = pagemargins.Top;
						if(pagemargins.Left) cursheet['!margins'].left = pagemargins.Left;
						if(pagemargins.Right) cursheet['!margins'].right = pagemargins.Right;
						if(pagemargins.Bottom) cursheet['!margins'].bottom = pagemargins.Bottom;
						break;
					case 'DisplayRightToLeft':
						if(!Workbook.Views) Workbook.Views = [];
						if(!Workbook.Views[0]) Workbook.Views[0] = {};
						Workbook.Views[0].RTL = true;
						break;

					case 'Unsynced': break;
					case 'Print': break;
					case 'Panes': break;
					case 'Scale': break;
					case 'Pane': break;
					case 'Number': break;
					case 'Layout': break;
					case 'PageSetup': break;
					case 'Selected': break;
					case 'ProtectObjects': break;
					case 'EnableSelection': break;
					case 'ProtectScenarios': break;
					case 'ValidPrinterInfo': break;
					case 'HorizontalResolution': break;
					case 'VerticalResolution': break;
					case 'NumberofCopies': break;
					case 'ActiveRow': break;
					case 'ActiveCol': break;
					case 'ActivePane': break;
					case 'TopRowVisible': break;
					case 'TopRowBottomPane': break;
					case 'LeftColumnVisible': break;
					case 'LeftColumnRightPane': break;
					case 'FitToPage': break;
					case 'RangeSelection': break;
					case 'PaperSizeIndex': break;
					case 'PageLayoutZoom': break;
					case 'PageBreakZoom': break;
					case 'FilterOn': break;
					case 'DoNotDisplayGridlines': break;
					case 'SplitHorizontal': break;
					case 'SplitVertical': break;
					case 'FreezePanes': break;
					case 'FrozenNoSplit': break;
					case 'FitWidth': break;
					case 'FitHeight': break;
					case 'CommentsLayout': break;
					case 'Zoom': break;
					case 'LeftToRight': break;
					case 'Gridlines': break;
					case 'AllowSort': break;
					case 'AllowFilter': break;
					case 'AllowInsertRows': break;
					case 'AllowDeleteRows': break;
					case 'AllowInsertCols': break;
					case 'AllowDeleteCols': break;
					case 'AllowInsertHyperlinks': break;
					case 'AllowFormatCells': break;
					case 'AllowSizeCols': break;
					case 'AllowSizeRows': break;
					case 'NoSummaryRowsBelowDetail': break;
					case 'TabColorIndex': break;
					case 'DoNotDisplayHeadings': break;
					case 'ShowPageLayoutZoom': break;
					case 'NoSummaryColumnsRightDetail': break;
					case 'BlackAndWhite': break;
					case 'DoNotDisplayZeros': break;
					case 'DisplayPageBreak': break;
					case 'RowColHeadings': break;
					case 'DoNotDisplayOutline': break;
					case 'NoOrientation': break;
					case 'AllowUsePivotTables': break;
					case 'ZeroHeight': break;
					case 'ViewableRange': break;
					case 'Selection': break;
					case 'ProtectContents': break;
					default: seen = false;
				} break;

				/* PivotTable */
				case 'PivotTable': case 'PivotCache': switch(Rn[3]) {
					case 'ImmediateItemsOnDrop': break;
					case 'ShowPageMultipleItemLabel': break;
					case 'CompactRowIndent': break;
					case 'Location': break;
					case 'PivotField': break;
					case 'Orientation': break;
					case 'LayoutForm': break;
					case 'LayoutSubtotalLocation': break;
					case 'LayoutCompactRow': break;
					case 'Position': break;
					case 'PivotItem': break;
					case 'DataType': break;
					case 'DataField': break;
					case 'SourceName': break;
					case 'ParentField': break;
					case 'PTLineItems': break;
					case 'PTLineItem': break;
					case 'CountOfSameItems': break;
					case 'Item': break;
					case 'ItemType': break;
					case 'PTSource': break;
					case 'CacheIndex': break;
					case 'ConsolidationReference': break;
					case 'FileName': break;
					case 'Reference': break;
					case 'NoColumnGrand': break;
					case 'NoRowGrand': break;
					case 'BlankLineAfterItems': break;
					case 'Hidden': break;
					case 'Subtotal': break;
					case 'BaseField': break;
					case 'MapChildItems': break;
					case 'Function': break;
					case 'RefreshOnFileOpen': break;
					case 'PrintSetTitles': break;
					case 'MergeLabels': break;
					case 'DefaultVersion': break;
					case 'RefreshName': break;
					case 'RefreshDate': break;
					case 'RefreshDateCopy': break;
					case 'VersionLastRefresh': break;
					case 'VersionLastUpdate': break;
					case 'VersionUpdateableMin': break;
					case 'VersionRefreshableMin': break;
					case 'Calculation': break;
					default: seen = false;
				} break;

				/* PageBreaks */
				case 'PageBreaks': switch(Rn[3]) {
					case 'ColBreaks': break;
					case 'ColBreak': break;
					case 'RowBreaks': break;
					case 'RowBreak': break;
					case 'ColStart': break;
					case 'ColEnd': break;
					case 'RowEnd': break;
					default: seen = false;
				} break;

				/* AutoFilter */
				case 'AutoFilter': switch(Rn[3]) {
					case 'AutoFilterColumn': break;
					case 'AutoFilterCondition': break;
					case 'AutoFilterAnd': break;
					case 'AutoFilterOr': break;
					default: seen = false;
				} break;

				/* QueryTable */
				case 'QueryTable': switch(Rn[3]) {
					case 'Id': break;
					case 'AutoFormatFont': break;
					case 'AutoFormatPattern': break;
					case 'QuerySource': break;
					case 'QueryType': break;
					case 'EnableRedirections': break;
					case 'RefreshedInXl9': break;
					case 'URLString': break;
					case 'HTMLTables': break;
					case 'Connection': break;
					case 'CommandText': break;
					case 'RefreshInfo': break;
					case 'NoTitles': break;
					case 'NextId': break;
					case 'ColumnInfo': break;
					case 'OverwriteCells': break;
					case 'DoNotPromptForFile': break;
					case 'TextWizardSettings': break;
					case 'Source': break;
					case 'Number': break;
					case 'Decimal': break;
					case 'ThousandSeparator': break;
					case 'TrailingMinusNumbers': break;
					case 'FormatSettings': break;
					case 'FieldType': break;
					case 'Delimiters': break;
					case 'Tab': break;
					case 'Comma': break;
					case 'AutoFormatName': break;
					case 'VersionLastEdit': break;
					case 'VersionLastRefresh': break;
					default: seen = false;
				} break;

				case 'Sorting':
				case 'ConditionalFormatting':
				case 'DataValidation':
				switch(Rn[3]) {
					case 'Range': break;
					case 'Type': break;
					case 'Min': break;
					case 'Max': break;
					case 'Sort': break;
					case 'Descending': break;
					case 'Order': break;
					case 'CaseSensitive': break;
					case 'Value': break;
					case 'ErrorStyle': break;
					case 'ErrorMessage': break;
					case 'ErrorTitle': break;
					case 'CellRangeList': break;
					case 'InputMessage': break;
					case 'InputTitle': break;
					case 'ComboHide': break;
					case 'InputHide': break;
					case 'Condition': break;
					case 'Qualifier': break;
					case 'UseBlank': break;
					case 'Value1': break;
					case 'Value2': break;
					case 'Format': break;
					default: seen = false;
				} break;

				/* MapInfo (schema) */
				case 'MapInfo': case 'Schema': case 'data': switch(Rn[3]) {
					case 'Map': break;
					case 'Entry': break;
					case 'Range': break;
					case 'XPath': break;
					case 'Field': break;
					case 'XSDType': break;
					case 'FilterOn': break;
					case 'Aggregate': break;
					case 'ElementType': break;
					case 'AttributeType': break;
				/* These are from xsd (XML Schema Definition) */
					case 'schema':
					case 'element':
					case 'complexType':
					case 'datatype':
					case 'all':
					case 'attribute':
					case 'extends': break;

					case 'row': break;
					default: seen = false;
				} break;

				/* SmartTags (can be anything) */
				case 'SmartTags': break;

				default: seen = false; break;
			}
			if(seen) break;
			/* CustomDocumentProperties */
			if(!state[state.length-1][1]) throw 'Unrecognized tag: ' + Rn[3] + "|" + state.join("|");
			if(state[state.length-1][0]==='CustomDocumentProperties') {
				if(Rn[0].slice(-2) === "/>") break;
				else if(Rn[1]==="/") xlml_set_custprop(Custprops, Rn[3], cp, str.slice(pidx, Rn.index));
				else { cp = Rn; pidx = Rn.index + Rn[0].length; }
				break;
			}
			if(opts.WTF) throw 'Unrecognized tag: ' + Rn[3] + "|" + state.join("|");
	}
	var out = ({});
	if(!opts.bookSheets && !opts.bookProps) out.Sheets = sheets;
	out.SheetNames = sheetnames;
	out.Workbook = Workbook;
	out.SSF = SSF.get_table();
	out.Props = Props;
	out.Custprops = Custprops;
	return out;
}

function parse_xlml(data, opts) {
	fix_read_opts(opts=opts||{});
	switch(opts.type||"base64") {
		case "base64": return parse_xlml_xml(Base64.decode(data), opts);
		case "binary": case "buffer": case "file": return parse_xlml_xml(data, opts);
		case "array": return parse_xlml_xml(a2s(data), opts);
	}
}

/* TODO */
function write_props_xlml(wb, opts) {
	var o = [];
	/* DocumentProperties */
	if(wb.Props) o.push(xlml_write_docprops(wb.Props, opts));
	/* CustomDocumentProperties */
	if(wb.Custprops) o.push(xlml_write_custprops(wb.Props, wb.Custprops, opts));
	return o.join("");
}
/* TODO */
function write_wb_xlml() {
	/* OfficeDocumentSettings */
	/* ExcelWorkbook */
	return "";
}
/* TODO */
function write_sty_xlml(wb, opts) {
	/* Styles */
	var styles = ['<Style ss:ID="Default" ss:Name="Normal"><NumberFormat/></Style>'];
	opts.cellXfs.forEach(function(xf, id) {
		var payload = [];
		payload.push(writextag('NumberFormat', null, {"ss:Format": escapexml(SSF._table[xf.numFmtId])}));
		styles.push(writextag('Style', payload.join(""), {"ss:ID": "s" + (21+id)}));
	});
	return writextag("Styles", styles.join(""));
}
function write_name_xlml(n) { return writextag("NamedRange", null, {"ss:Name": n.Name, "ss:RefersTo":"=" + a1_to_rc(n.Ref, {r:0,c:0})}); }
function write_names_xlml(wb) {
	if(!((wb||{}).Workbook||{}).Names) return "";
var names = wb.Workbook.Names;
	var out = [];
	for(var i = 0; i < names.length; ++i) {
		var n = names[i];
		if(n.Sheet != null) continue;
		if(n.Name.match(/^_xlfn\./)) continue;
		out.push(write_name_xlml(n));
	}
	return writextag("Names", out.join(""));
}
function write_ws_xlml_names(ws, opts, idx, wb) {
	if(!ws) return "";
	if(!((wb||{}).Workbook||{}).Names) return "";
var names = wb.Workbook.Names;
	var out = [];
	for(var i = 0; i < names.length; ++i) {
		var n = names[i];
		if(n.Sheet != idx) continue;
		/*switch(n.Name) {
			case "_": continue;
		}*/
		if(n.Name.match(/^_xlfn\./)) continue;
		out.push(write_name_xlml(n));
	}
	return out.join("");
}
/* WorksheetOptions */
function write_ws_xlml_wsopts(ws, opts, idx, wb) {
	if(!ws) return "";
	var o = [];
	/* NOTE: spec technically allows any order, but stick with implied order */

	/* FitToPage */
	/* DoNotDisplayColHeaders */
	/* DoNotDisplayRowHeaders */
	/* ViewableRange */
	/* Selection */
	/* GridlineColor */
	/* Name */
	/* ExcelWorksheetType */
	/* IntlMacro */
	/* Unsynced */
	/* Selected */
	/* CodeName */

	if(ws['!margins']) {
		o.push("<PageSetup>");
		if(ws['!margins'].header) o.push(writextag("Header", null, {'x:Margin':ws['!margins'].header}));
		if(ws['!margins'].footer) o.push(writextag("Footer", null, {'x:Margin':ws['!margins'].footer}));
		o.push(writextag("PageMargins", null, {
			'x:Bottom': ws['!margins'].bottom || "0.75",
			'x:Left': ws['!margins'].left || "0.7",
			'x:Right': ws['!margins'].right || "0.7",
			'x:Top': ws['!margins'].top || "0.75"
		}));
		o.push("</PageSetup>");
	}

	/* PageSetup */
	/* DisplayPageBreak */
	/* TransitionExpressionEvaluation */
	/* TransitionFormulaEntry */
	/* Print */
	/* Zoom */
	/* PageLayoutZoom */
	/* PageBreakZoom */
	/* ShowPageBreakZoom */
	/* DefaultRowHeight */
	/* DefaultColumnWidth */
	/* StandardWidth */

	if(wb && wb.Workbook && wb.Workbook.Sheets && wb.Workbook.Sheets[idx]) {
		/* Visible */
		if(wb.Workbook.Sheets[idx].Hidden) o.push(writextag("Visible", (wb.Workbook.Sheets[idx].Hidden == 1 ? "SheetHidden" : "SheetVeryHidden"), {}));
		else {
			/* Selected */
			for(var i = 0; i < idx; ++i) if(wb.Workbook.Sheets[i] && !wb.Workbook.Sheets[i].Hidden) break;
			if(i == idx) o.push("<Selected/>");
		}
	}

	/* LeftColumnVisible */

	if(((((wb||{}).Workbook||{}).Views||[])[0]||{}).RTL) o.push("<DisplayRightToLeft/>");

	/* GridlineColorIndex */
	/* DisplayFormulas */
	/* DoNotDisplayGridlines */
	/* DoNotDisplayHeadings */
	/* DoNotDisplayOutline */
	/* ApplyAutomaticOutlineStyles */
	/* NoSummaryRowsBelowDetail */
	/* NoSummaryColumnsRightDetail */
	/* DoNotDisplayZeros */
	/* ActiveRow */
	/* ActiveColumn */
	/* FilterOn */
	/* RangeSelection */
	/* TopRowVisible */
	/* TopRowBottomPane */
	/* LeftColumnRightPane */
	/* ActivePane */
	/* SplitHorizontal */
	/* SplitVertical */
	/* FreezePanes */
	/* FrozenNoSplit */
	/* TabColorIndex */
	/* Panes */

	/* NOTE: Password not supported in XLML Format */
	if(ws['!protect']) {
		o.push(writetag("ProtectContents", "True"));
		if(ws['!protect'].objects) o.push(writetag("ProtectObjects", "True"));
		if(ws['!protect'].scenarios) o.push(writetag("ProtectScenarios", "True"));
		if(ws['!protect'].selectLockedCells != null && !ws['!protect'].selectLockedCells) o.push(writetag("EnableSelection", "NoSelection"));
		else if(ws['!protect'].selectUnlockedCells != null && !ws['!protect'].selectUnlockedCells) o.push(writetag("EnableSelection", "UnlockedCells"));
	[
		[ "formatCells", "AllowFormatCells" ],
		[ "formatColumns", "AllowSizeCols" ],
		[ "formatRows", "AllowSizeRows" ],
		[ "insertColumns", "AllowInsertCols" ],
		[ "insertRows", "AllowInsertRows" ],
		[ "insertHyperlinks", "AllowInsertHyperlinks" ],
		[ "deleteColumns", "AllowDeleteCols" ],
		[ "deleteRows", "AllowDeleteRows" ],
		[ "sort", "AllowSort" ],
		[ "autoFilter", "AllowFilter" ],
		[ "pivotTables", "AllowUsePivotTables" ]
	].forEach(function(x) { if(ws['!protect'][x[0]]) o.push("<"+x[1]+"/>"); });
	}

	if(o.length == 0) return "";
	return writextag("WorksheetOptions", o.join(""), {xmlns:XLMLNS.x});
}
function write_ws_xlml_comment(comments) {
	return comments.map(function(c) {
		// TODO: formatted text
		var t = xlml_unfixstr(c.t||"");
		var d =writextag("ss:Data", t, {"xmlns":"http://www.w3.org/TR/REC-html40"});
		return writextag("Comment", d, {"ss:Author":c.a});
	}).join("");
}
function write_ws_xlml_cell(cell, ref, ws, opts, idx, wb, addr){
	if(!cell || (cell.v == undefined && cell.f == undefined)) return "";

	var attr = {};
	if(cell.f) attr["ss:Formula"] = "=" + escapexml(a1_to_rc(cell.f, addr));
	if(cell.F && cell.F.slice(0, ref.length) == ref) {
		var end = decode_cell(cell.F.slice(ref.length + 1));
		attr["ss:ArrayRange"] = "RC:R" + (end.r == addr.r ? "" : "[" + (end.r - addr.r) + "]") + "C" + (end.c == addr.c ? "" : "[" + (end.c - addr.c) + "]");
	}

	if(cell.l && cell.l.Target) {
		attr["ss:HRef"] = escapexml(cell.l.Target);
		if(cell.l.Tooltip) attr["x:HRefScreenTip"] = escapexml(cell.l.Tooltip);
	}

	if(ws['!merges']) {
		var marr = ws['!merges'];
		for(var mi = 0; mi != marr.length; ++mi) {
			if(marr[mi].s.c != addr.c || marr[mi].s.r != addr.r) continue;
			if(marr[mi].e.c > marr[mi].s.c) attr['ss:MergeAcross'] = marr[mi].e.c - marr[mi].s.c;
			if(marr[mi].e.r > marr[mi].s.r) attr['ss:MergeDown'] = marr[mi].e.r - marr[mi].s.r;
		}
	}

	var t = "", p = "";
	switch(cell.t) {
		case 'z': return "";
		case 'n': t = 'Number'; p = String(cell.v); break;
		case 'b': t = 'Boolean'; p = (cell.v ? "1" : "0"); break;
		case 'e': t = 'Error'; p = BErr[cell.v]; break;
		case 'd': t = 'DateTime'; p = new Date(cell.v).toISOString(); if(cell.z == null) cell.z = cell.z || SSF._table[14]; break;
		case 's': t = 'String'; p = escapexlml(cell.v||""); break;
	}
	/* TODO: cell style */
	var os = get_cell_style(opts.cellXfs, cell, opts);
	attr["ss:StyleID"] = "s" + (21+os);
	attr["ss:Index"] = addr.c + 1;
	var _v = (cell.v != null ? p : "");
	var m = '<Data ss:Type="' + t + '">' + _v + '</Data>';

	if((cell.c||[]).length > 0) m += write_ws_xlml_comment(cell.c);

	return writextag("Cell", m, attr);
}
function write_ws_xlml_row(R, row) {
	var o = '<Row ss:Index="' + (R+1) + '"';
	if(row) {
		if(row.hpt && !row.hpx) row.hpx = pt2px(row.hpt);
		if(row.hpx) o += ' ss:AutoFitHeight="0" ss:Height="' + row.hpx + '"';
		if(row.hidden) o += ' ss:Hidden="1"';
	}
	return o + '>';
}
/* TODO */
function write_ws_xlml_table(ws, opts, idx, wb) {
	if(!ws['!ref']) return "";
	var range = safe_decode_range(ws['!ref']);
	var marr = ws['!merges'] || [], mi = 0;
	var o = [];
	if(ws['!cols']) ws['!cols'].forEach(function(n, i) {
		process_col(n);
		var w = !!n.width;
		var p = col_obj_w(i, n);
		var k = {"ss:Index":i+1};
		if(w) k['ss:Width'] = width2px(p.width);
		if(n.hidden) k['ss:Hidden']="1";
		o.push(writextag("Column",null,k));
	});
	var dense = Array.isArray(ws);
	for(var R = range.s.r; R <= range.e.r; ++R) {
		var row = [write_ws_xlml_row(R, (ws['!rows']||[])[R])];
		for(var C = range.s.c; C <= range.e.c; ++C) {
			var skip = false;
			for(mi = 0; mi != marr.length; ++mi) {
				if(marr[mi].s.c > C) continue;
				if(marr[mi].s.r > R) continue;
				if(marr[mi].e.c < C) continue;
				if(marr[mi].e.r < R) continue;
				if(marr[mi].s.c != C || marr[mi].s.r != R) skip = true;
				break;
			}
			if(skip) continue;
			var addr = {r:R,c:C};
			var ref = encode_cell(addr), cell = dense ? (ws[R]||[])[C] : ws[ref];
			row.push(write_ws_xlml_cell(cell, ref, ws, opts, idx, wb, addr));
		}
		row.push("</Row>");
		if(row.length > 2) o.push(row.join(""));
	}
	return o.join("");
}
function write_ws_xlml(idx, opts, wb) {
	var o = [];
	var s = wb.SheetNames[idx];
	var ws = wb.Sheets[s];

	var t = ws ? write_ws_xlml_names(ws, opts, idx, wb) : "";
	if(t.length > 0) o.push("<Names>" + t + "</Names>");

	/* Table */
	t = ws ? write_ws_xlml_table(ws, opts, idx, wb) : "";
	if(t.length > 0) o.push("<Table>" + t + "</Table>");

	/* WorksheetOptions */
	o.push(write_ws_xlml_wsopts(ws, opts, idx, wb));

	return o.join("");
}
function write_xlml(wb, opts) {
	if(!opts) opts = {};
	if(!wb.SSF) wb.SSF = SSF.get_table();
	if(wb.SSF) {
		make_ssf(SSF); SSF.load_table(wb.SSF);
		// $FlowIgnore
		opts.revssf = evert_num(wb.SSF); opts.revssf[wb.SSF[65535]] = 0;
		opts.ssf = wb.SSF;
		opts.cellXfs = [];
		get_cell_style(opts.cellXfs, {}, {revssf:{"General":0}});
	}
	var d = [];
	d.push(write_props_xlml(wb, opts));
	d.push(write_wb_xlml(wb, opts));
	d.push("");
	d.push("");
	for(var i = 0; i < wb.SheetNames.length; ++i)
		d.push(writextag("Worksheet", write_ws_xlml(i, opts, wb), {"ss:Name":escapexml(wb.SheetNames[i])}));
	d[2] = write_sty_xlml(wb, opts);
	d[3] = write_names_xlml(wb, opts);
	return XML_HEADER + writextag("Workbook", d.join(""), {
		'xmlns':      XLMLNS.ss,
		'xmlns:o':    XLMLNS.o,
		'xmlns:x':    XLMLNS.x,
		'xmlns:ss':   XLMLNS.ss,
		'xmlns:dt':   XLMLNS.dt,
		'xmlns:html': XLMLNS.html
	});
}
/* [MS-OLEDS] 2.3.8 CompObjStream */
function parse_compobj(obj) {
	var v = {};
	var o = obj.content;
/* [MS-OLEDS] 2.3.7 CompObjHeader -- All fields MUST be ignored */
	o.l = 28;

	v.AnsiUserType = o.read_shift(0, "lpstr-ansi");
	v.AnsiClipboardFormat = parse_ClipboardFormatOrAnsiString(o);

	if(o.length - o.l <= 4) return v;

	var m = o.read_shift(4);
	if(m == 0 || m > 40) return v;
	o.l-=4; v.Reserved1 = o.read_shift(0, "lpstr-ansi");

	if(o.length - o.l <= 4) return v;
	m = o.read_shift(4);
	if(m !== 0x71b239f4) return v;
	v.UnicodeClipboardFormat = parse_ClipboardFormatOrUnicodeString(o);

	m = o.read_shift(4);
	if(m == 0 || m > 40) return v;
	o.l-=4; v.Reserved2 = o.read_shift(0, "lpwstr");
}

/*
	Continue logic for:
	- 2.4.58 Continue
	- 2.4.59 ContinueBigName
	- 2.4.60 ContinueFrt
	- 2.4.61 ContinueFrt11
	- 2.4.62 ContinueFrt12
*/
function slurp(R, blob, length, opts) {
	var l = length;
	var bufs = [];
	var d = blob.slice(blob.l,blob.l+l);
	if(opts && opts.enc && opts.enc.insitu) switch(R.n) {
	case 'BOF': case 'FilePass': case 'FileLock': case 'InterfaceHdr': case 'RRDInfo': case 'RRDHead': case 'UsrExcl': break;
	default:
		if(d.length === 0) break;
		opts.enc.insitu(d);
	}
	bufs.push(d);
	blob.l += l;
	var next = (XLSRecordEnum[__readUInt16LE(blob,blob.l)]);
	var start = 0;
	while(next != null && next.n.slice(0,8) === 'Continue') {
		l = __readUInt16LE(blob,blob.l+2);
		start = blob.l + 4;
		if(next.n == 'ContinueFrt') start += 4;
		else if(next.n.slice(0,11) == 'ContinueFrt') start += 12;
		bufs.push(blob.slice(start,blob.l+4+l));
		blob.l += 4+l;
		next = (XLSRecordEnum[__readUInt16LE(blob, blob.l)]);
	}
	var b = (bconcat(bufs));
	prep_blob(b, 0);
	var ll = 0; b.lens = [];
	for(var j = 0; j < bufs.length; ++j) { b.lens.push(ll); ll += bufs[j].length; }
	return R.f(b, b.length, opts);
}

function safe_format_xf(p, opts, date1904) {
	if(p.t === 'z') return;
	if(!p.XF) return;
	var fmtid = 0;
	try {
		fmtid = p.z || p.XF.numFmtId || 0;
		if(opts.cellNF) p.z = SSF._table[fmtid];
	} catch(e) { if(opts.WTF) throw e; }
	if(!opts || opts.cellText !== false) try {
		if(p.t === 'e') { p.w = p.w || BErr[p.v]; }
		else if(fmtid === 0 || fmtid == "General") {
			if(p.t === 'n') {
				if((p.v|0) === p.v) p.w = SSF._general_int(p.v);
				else p.w = SSF._general_num(p.v);
			}
			else p.w = SSF._general(p.v);
		}
		else p.w = SSF.format(fmtid,p.v, {date1904:!!date1904});
	} catch(e) { if(opts.WTF) throw e; }
	if(opts.cellDates && fmtid && p.t == 'n' && SSF.is_date(SSF._table[fmtid] || String(fmtid))) {
		var _d = SSF.parse_date_code(p.v); if(_d) { p.t = 'd'; p.v = new Date(_d.y, _d.m-1,_d.d,_d.H,_d.M,_d.S,_d.u); }
	}
}

function make_cell(val, ixfe, t) {
	return ({v:val, ixfe:ixfe, t:t});
}

// 2.3.2
function parse_workbook(blob, options) {
	var wb = ({opts:{}});
	var Sheets = {};
	if(DENSE != null && options.dense == null) options.dense = DENSE;
	var out = ((options.dense ? [] : {}));
	var Directory = {};
	var range = ({});
	var last_formula = null;
	var sst = ([]);
	var cur_sheet = "";
	var Preamble = {};
	var lastcell, last_cell = "", cc, cmnt, rngC, rngR;
	var sharedf = {};
	var arrayf = [];
	var temp_val;
	var country;
	var cell_valid = true;
	var XFs = []; /* XF records */
	var palette = [];
	var Workbook = ({ Sheets:[], WBProps:{date1904:false}, Views:[{}] }), wsprops = {};
	var get_rgb = function getrgb(icv) {
		if(icv < 8) return XLSIcv[icv];
		if(icv < 64) return palette[icv-8] || XLSIcv[icv];
		return XLSIcv[icv];
	};
	var process_cell_style = function pcs(cell, line, options) {
		var xfd = line.XF.data;
		if(!xfd || !xfd.patternType || !options || !options.cellStyles) return;
		line.s = ({});
		line.s.patternType = xfd.patternType;
		var t;
		if((t = rgb2Hex(get_rgb(xfd.icvFore)))) { line.s.fgColor = {rgb:t}; }
		if((t = rgb2Hex(get_rgb(xfd.icvBack)))) { line.s.bgColor = {rgb:t}; }
	};
	var addcell = function addcell(cell, line, options) {
		if(file_depth > 1) return;
		if(options.sheetRows && cell.r >= options.sheetRows) cell_valid = false;
		if(!cell_valid) return;
		if(options.cellStyles && line.XF && line.XF.data) process_cell_style(cell, line, options);
		delete line.ixfe; delete line.XF;
		lastcell = cell;
		last_cell = encode_cell(cell);
		if(!range || !range.s || !range.e) range = {s:{r:0,c:0},e:{r:0,c:0}};
		if(cell.r < range.s.r) range.s.r = cell.r;
		if(cell.c < range.s.c) range.s.c = cell.c;
		if(cell.r + 1 > range.e.r) range.e.r = cell.r + 1;
		if(cell.c + 1 > range.e.c) range.e.c = cell.c + 1;
		if(options.cellFormula && line.f) {
			for(var afi = 0; afi < arrayf.length; ++afi) {
				if(arrayf[afi][0].s.c > cell.c || arrayf[afi][0].s.r > cell.r) continue;
				if(arrayf[afi][0].e.c < cell.c || arrayf[afi][0].e.r < cell.r) continue;
				line.F = encode_range(arrayf[afi][0]);
				if(arrayf[afi][0].s.c != cell.c || arrayf[afi][0].s.r != cell.r) delete line.f;
				if(line.f) line.f = "" + stringify_formula(arrayf[afi][1], range, cell, supbooks, opts);
				break;
			}
		}
		{
			if(options.dense) {
				if(!out[cell.r]) out[cell.r] = [];
				out[cell.r][cell.c] = line;
			} else out[last_cell] = line;
		}
	};
	var opts = ({
		enc: false, // encrypted
		sbcch: 0, // cch in the preceding SupBook
		snames: [], // sheetnames
		sharedf: sharedf, // shared formulae by address
		arrayf: arrayf, // array formulae array
		rrtabid: [], // RRTabId
		lastuser: "", // Last User from WriteAccess
		biff: 8, // BIFF version
		codepage: 0, // CP from CodePage record
		winlocked: 0, // fLockWn from WinProtect
		cellStyles: !!options && !!options.cellStyles,
		WTF: !!options && !!options.wtf
	});
	if(options.password) opts.password = options.password;
	var themes;
	var merges = [];
	var objects = [];
	var colinfo = [], rowinfo = [];
	// eslint-disable-next-line no-unused-vars
	var defwidth = 0, defheight = 0; // twips / MDW respectively
	var seencol = false;
	var supbooks = ([]); // 1-indexed, will hold extern names
	supbooks.SheetNames = opts.snames;
	supbooks.sharedf = opts.sharedf;
	supbooks.arrayf = opts.arrayf;
	supbooks.names = [];
	supbooks.XTI = [];
	var last_Rn = '';
	var file_depth = 0; /* TODO: make a real stack */
	var BIFF2Fmt = 0, BIFF2FmtTable = [];
	var FilterDatabases = []; /* TODO: sort out supbooks and process elsewhere */
	var last_lbl;

	/* explicit override for some broken writers */
	opts.codepage = 1200;
	set_cp(1200);
	var seen_codepage = false;
	while(blob.l < blob.length - 1) {
		var s = blob.l;
		var RecordType = blob.read_shift(2);
		if(RecordType === 0 && last_Rn === 'EOF') break;
		var length = (blob.l === blob.length ? 0 : blob.read_shift(2));
		var R = XLSRecordEnum[RecordType];
		//console.log(RecordType.toString(16), RecordType, R, blob.l, length, blob.length);
		//if(!R) console.log(blob.slice(blob.l, blob.l + length));
		if(R && R.f) {
			if(options.bookSheets) {
				if(last_Rn === 'BoundSheet8' && R.n !== 'BoundSheet8') break;
			}
			last_Rn = R.n;
			if(R.r === 2 || R.r == 12) {
				var rt = blob.read_shift(2); length -= 2;
				if(!opts.enc && rt !== RecordType && (((rt&0xFF)<<8)|(rt>>8)) !== RecordType) throw new Error("rt mismatch: " + rt + "!=" + RecordType);
				if(R.r == 12){ blob.l += 10; length -= 10; } // skip FRT
			}
			//console.error(R,blob.l,length,blob.length);
			var val;
			if(R.n === 'EOF') val = R.f(blob, length, opts);
			else val = slurp(R, blob, length, opts);
			var Rn = R.n;
			if(file_depth == 0 && Rn != 'BOF') continue;
			/* nested switch statements to workaround V8 128 limit */
			switch(Rn) {
				/* Workbook Options */
				case 'Date1904':
wb.opts.Date1904 = Workbook.WBProps.date1904 = val; break;
				case 'WriteProtect': wb.opts.WriteProtect = true; break;
				case 'FilePass':
					if(!opts.enc) blob.l = 0;
					opts.enc = val;
					if(!options.password) throw new Error("File is password-protected");
					if(val.valid == null) throw new Error("Encryption scheme unsupported");
					if(!val.valid) throw new Error("Password is incorrect");
					break;
				case 'WriteAccess': opts.lastuser = val; break;
				case 'FileSharing': break; //TODO
				case 'CodePage':
					/* overrides based on test cases */
					switch(val) {
						case 0x5212: val =  1200; break;
						case 0x8000: val = 10000; break;
						case 0x8001: val =  1252; break;
					}
					set_cp(opts.codepage = val);
					seen_codepage = true;
					break;
				case 'RRTabId': opts.rrtabid = val; break;
				case 'WinProtect': opts.winlocked = val; break;
				case 'Template': break; // TODO
				case 'BookBool': break; // TODO
				case 'UsesELFs': break;
				case 'MTRSettings': break;
				case 'RefreshAll':
				case 'CalcCount':
				case 'CalcDelta':
				case 'CalcIter':
				case 'CalcMode':
				case 'CalcPrecision':
				case 'CalcSaveRecalc':
					wb.opts[Rn] = val; break;
				case 'CalcRefMode': opts.CalcRefMode = val; break; // TODO: implement R1C1
				case 'Uncalced': break;
				case 'ForceFullCalculation': wb.opts.FullCalc = val; break;
				case 'WsBool':
					if(val.fDialog) out["!type"] = "dialog";
					break; // TODO
				case 'XF':
					XFs.push(val); break;
				case 'ExtSST': break; // TODO
				case 'BookExt': break; // TODO
				case 'RichTextStream': break;
				case 'BkHim': break;

				case 'SupBook':
					supbooks.push([val]);
					supbooks[supbooks.length-1].XTI = [];
					break;
				case 'ExternName':
					supbooks[supbooks.length-1].push(val);
					break;
				case 'Index': break; // TODO
				case 'Lbl':
					last_lbl = ({
						Name: val.Name,
						Ref: stringify_formula(val.rgce,range,null,supbooks,opts)
					});
					if(val.itab > 0) last_lbl.Sheet = val.itab - 1;
					supbooks.names.push(last_lbl);
					if(!supbooks[0]) { supbooks[0] = []; supbooks[0].XTI = []; }
					supbooks[supbooks.length-1].push(val);
					if(val.Name == "_xlnm._FilterDatabase" && val.itab > 0)
						if(val.rgce && val.rgce[0] && val.rgce[0][0] && val.rgce[0][0][0] == 'PtgArea3d')
							FilterDatabases[val.itab - 1] = { ref: encode_range(val.rgce[0][0][1][2]) };
					break;
				case 'ExternCount': opts.ExternCount = val; break;
				case 'ExternSheet':
					if(supbooks.length == 0) { supbooks[0] = []; supbooks[0].XTI = []; }
					supbooks[supbooks.length - 1].XTI = supbooks[supbooks.length - 1].XTI.concat(val); supbooks.XTI = supbooks.XTI.concat(val); break;
				case 'NameCmt':
					/* TODO: search for correct name */
					if(opts.biff < 8) break;
					if(last_lbl != null) last_lbl.Comment = val[1];
					break;

				case 'Protect': out["!protect"] = val; break; /* for sheet or book */
				case 'Password': if(val !== 0 && opts.WTF) console.error("Password verifier: " + val); break;
				case 'Prot4Rev': case 'Prot4RevPass': break; /*TODO: Revision Control*/

				case 'BoundSheet8': {
					Directory[val.pos] = val;
					opts.snames.push(val.name);
				} break;
				case 'EOF': {
					if(--file_depth) break;
					if(range.e) {
						if(range.e.r > 0 && range.e.c > 0) {
							range.e.r--; range.e.c--;
							out["!ref"] = encode_range(range);
							if(options.sheetRows && options.sheetRows <= range.e.r) {
								var tmpri = range.e.r;
								range.e.r = options.sheetRows - 1;
								out["!fullref"] = out["!ref"];
								out["!ref"] = encode_range(range);
								range.e.r = tmpri;
							}
							range.e.r++; range.e.c++;
						}
						if(merges.length > 0) out["!merges"] = merges;
						if(objects.length > 0) out["!objects"] = objects;
						if(colinfo.length > 0) out["!cols"] = colinfo;
						if(rowinfo.length > 0) out["!rows"] = rowinfo;
						Workbook.Sheets.push(wsprops);
					}
					if(cur_sheet === "") Preamble = out; else Sheets[cur_sheet] = out;
					out = ((options.dense ? [] : {}));
				} break;
				case 'BOF': {
					if(opts.biff === 8) opts.biff = {
0x0009:2,
0x0209:3,
0x0409:4
					}[RecordType] || {
0x0200:2,
0x0300:3,
0x0400:4,
0x0500:5,
0x0600:8,
0x0002:2,
0x0007:2
					}[val.BIFFVer] || 8;
					if(opts.biff == 8 && val.BIFFVer == 0 && val.dt == 16) opts.biff = 2;
					if(file_depth++) break;
					cell_valid = true;
					out = ((options.dense ? [] : {}));

					if(opts.biff < 8 && !seen_codepage) { seen_codepage = true; set_cp(opts.codepage = options.codepage || 1252); }
					if(opts.biff < 5) {
						if(cur_sheet === "") cur_sheet = "Sheet1";
						range = {s:{r:0,c:0},e:{r:0,c:0}};
						/* fake BoundSheet8 */
						var fakebs8 = {pos: blob.l - length, name:cur_sheet};
						Directory[fakebs8.pos] = fakebs8;
						opts.snames.push(cur_sheet);
					}
					else cur_sheet = (Directory[s] || {name:""}).name;
					if(val.dt == 0x20) out["!type"] = "chart";
					if(val.dt == 0x40) out["!type"] = "macro";
					merges = [];
					objects = [];
					opts.arrayf = arrayf = [];
					colinfo = []; rowinfo = [];
					defwidth = defheight = 0;
					seencol = false;
					wsprops = {Hidden:(Directory[s]||{hs:0}).hs, name:cur_sheet };
				} break;

				case 'Number': case 'BIFF2NUM': case 'BIFF2INT': {
					if(out["!type"] == "chart") if(options.dense ? (out[val.r]||[])[val.c]: out[encode_cell({c:val.c, r:val.r})]) ++val.c;
					temp_val = ({ixfe: val.ixfe, XF: XFs[val.ixfe]||{}, v:val.val, t:'n'});
					if(BIFF2Fmt > 0) temp_val.z = BIFF2FmtTable[(temp_val.ixfe>>8) & 0x1F];
					safe_format_xf(temp_val, options, wb.opts.Date1904);
					addcell({c:val.c, r:val.r}, temp_val, options);
				} break;
				case 'BoolErr': {
					temp_val = ({ixfe: val.ixfe, XF: XFs[val.ixfe], v:val.val, t:val.t});
					if(BIFF2Fmt > 0) temp_val.z = BIFF2FmtTable[(temp_val.ixfe>>8) & 0x1F];
					safe_format_xf(temp_val, options, wb.opts.Date1904);
					addcell({c:val.c, r:val.r}, temp_val, options);
				} break;
				case 'RK': {
					temp_val = ({ixfe: val.ixfe, XF: XFs[val.ixfe], v:val.rknum, t:'n'});
					if(BIFF2Fmt > 0) temp_val.z = BIFF2FmtTable[(temp_val.ixfe>>8) & 0x1F];
					safe_format_xf(temp_val, options, wb.opts.Date1904);
					addcell({c:val.c, r:val.r}, temp_val, options);
				} break;
				case 'MulRk': {
					for(var j = val.c; j <= val.C; ++j) {
						var ixfe = val.rkrec[j-val.c][0];
						temp_val= ({ixfe:ixfe, XF:XFs[ixfe], v:val.rkrec[j-val.c][1], t:'n'});
						if(BIFF2Fmt > 0) temp_val.z = BIFF2FmtTable[(temp_val.ixfe>>8) & 0x1F];
						safe_format_xf(temp_val, options, wb.opts.Date1904);
						addcell({c:j, r:val.r}, temp_val, options);
					}
				} break;
				case 'Formula': {
					if(val.val == 'String') { last_formula = val; break; }
					temp_val = make_cell(val.val, val.cell.ixfe, val.tt);
					temp_val.XF = XFs[temp_val.ixfe];
					if(options.cellFormula) {
						var _f = val.formula;
						if(_f && _f[0] && _f[0][0] && _f[0][0][0] == 'PtgExp') {
							var _fr = _f[0][0][1][0], _fc = _f[0][0][1][1];
							var _fe = encode_cell({r:_fr, c:_fc});
							if(sharedf[_fe]) temp_val.f = ""+stringify_formula(val.formula,range,val.cell,supbooks, opts);
							else temp_val.F = ((options.dense ? (out[_fr]||[])[_fc]: out[_fe]) || {}).F;
						} else temp_val.f = ""+stringify_formula(val.formula,range,val.cell,supbooks, opts);
					}
					if(BIFF2Fmt > 0) temp_val.z = BIFF2FmtTable[(temp_val.ixfe>>8) & 0x1F];
					safe_format_xf(temp_val, options, wb.opts.Date1904);
					addcell(val.cell, temp_val, options);
					last_formula = val;
				} break;
				case 'String': {
					if(last_formula) { /* technically always true */
						last_formula.val = val;
						temp_val = make_cell(val, last_formula.cell.ixfe, 's');
						temp_val.XF = XFs[temp_val.ixfe];
						if(options.cellFormula) {
							temp_val.f = ""+stringify_formula(last_formula.formula, range, last_formula.cell, supbooks, opts);
						}
						if(BIFF2Fmt > 0) temp_val.z = BIFF2FmtTable[(temp_val.ixfe>>8) & 0x1F];
						safe_format_xf(temp_val, options, wb.opts.Date1904);
						addcell(last_formula.cell, temp_val, options);
						last_formula = null;
					} else throw new Error("String record expects Formula");
				} break;
				case 'Array': {
					arrayf.push(val);
					var _arraystart = encode_cell(val[0].s);
					cc = options.dense ? (out[val[0].s.r]||[])[val[0].s.c] : out[_arraystart];
					if(options.cellFormula && cc) {
						if(!last_formula) break; /* technically unreachable */
						if(!_arraystart || !cc) break;
						cc.f = ""+stringify_formula(val[1], range, val[0], supbooks, opts);
						cc.F = encode_range(val[0]);
					}
				} break;
				case 'ShrFmla': {
					if(!cell_valid) break;
					if(!options.cellFormula) break;
					if(last_cell) {
						/* TODO: capture range */
						if(!last_formula) break; /* technically unreachable */
						sharedf[encode_cell(last_formula.cell)]= val[0];
						cc = options.dense ? (out[last_formula.cell.r]||[])[last_formula.cell.c] : out[encode_cell(last_formula.cell)];
						(cc||{}).f = ""+stringify_formula(val[0], range, lastcell, supbooks, opts);
					}
				} break;
				case 'LabelSst':
					temp_val=make_cell(sst[val.isst].t, val.ixfe, 's');
					temp_val.XF = XFs[temp_val.ixfe];
					if(BIFF2Fmt > 0) temp_val.z = BIFF2FmtTable[(temp_val.ixfe>>8) & 0x1F];
					safe_format_xf(temp_val, options, wb.opts.Date1904);
					addcell({c:val.c, r:val.r}, temp_val, options);
					break;
				case 'Blank': if(options.sheetStubs) {
					temp_val = ({ixfe: val.ixfe, XF: XFs[val.ixfe], t:'z'});
					if(BIFF2Fmt > 0) temp_val.z = BIFF2FmtTable[(temp_val.ixfe>>8) & 0x1F];
					safe_format_xf(temp_val, options, wb.opts.Date1904);
					addcell({c:val.c, r:val.r}, temp_val, options);
				} break;
				case 'MulBlank': if(options.sheetStubs) {
					for(var _j = val.c; _j <= val.C; ++_j) {
						var _ixfe = val.ixfe[_j-val.c];
						temp_val= ({ixfe:_ixfe, XF:XFs[_ixfe], t:'z'});
						if(BIFF2Fmt > 0) temp_val.z = BIFF2FmtTable[(temp_val.ixfe>>8) & 0x1F];
						safe_format_xf(temp_val, options, wb.opts.Date1904);
						addcell({c:_j, r:val.r}, temp_val, options);
					}
				} break;
				case 'RString':
				case 'Label': case 'BIFF2STR':
					temp_val=make_cell(val.val, val.ixfe, 's');
					temp_val.XF = XFs[temp_val.ixfe];
					if(BIFF2Fmt > 0) temp_val.z = BIFF2FmtTable[(temp_val.ixfe>>8) & 0x1F];
					safe_format_xf(temp_val, options, wb.opts.Date1904);
					addcell({c:val.c, r:val.r}, temp_val, options);
					break;

				case 'Dimensions': {
					if(file_depth === 1) range = val; /* TODO: stack */
				} break;
				case 'SST': {
					sst = val;
				} break;
				case 'Format': { /* val = [id, fmt] */
					if(opts.biff == 4) {
						BIFF2FmtTable[BIFF2Fmt++] = val[1];
						for(var b4idx = 0; b4idx < BIFF2Fmt + 163; ++b4idx) if(SSF._table[b4idx] == val[1]) break;
						if(b4idx >= 163) SSF.load(val[1], BIFF2Fmt + 163);
					}
					else SSF.load(val[1], val[0]);
				} break;
				case 'BIFF2FORMAT': {
					BIFF2FmtTable[BIFF2Fmt++] = val;
					for(var b2idx = 0; b2idx < BIFF2Fmt + 163; ++b2idx) if(SSF._table[b2idx] == val) break;
					if(b2idx >= 163) SSF.load(val, BIFF2Fmt + 163);
				} break;

				case 'MergeCells': merges = merges.concat(val); break;

				case 'Obj': objects[val.cmo[0]] = opts.lastobj = val; break;
				case 'TxO': opts.lastobj.TxO = val; break;
				case 'ImData': opts.lastobj.ImData = val; break;

				case 'HLink': {
					for(rngR = val[0].s.r; rngR <= val[0].e.r; ++rngR)
						for(rngC = val[0].s.c; rngC <= val[0].e.c; ++rngC) {
							cc = options.dense ? (out[rngR]||[])[rngC] : out[encode_cell({c:rngC,r:rngR})];
							if(cc) cc.l = val[1];
						}
				} break;
				case 'HLinkTooltip': {
					for(rngR = val[0].s.r; rngR <= val[0].e.r; ++rngR)
						for(rngC = val[0].s.c; rngC <= val[0].e.c; ++rngC) {
							cc = options.dense ? (out[rngR]||[])[rngC] : out[encode_cell({c:rngC,r:rngR})];
							if(cc && cc.l) cc.l.Tooltip = val[1];
							}
				} break;

				/* Comments */
				case 'Note': {
					if(opts.biff <= 5 && opts.biff >= 2) break; /* TODO: BIFF5 */
					cc = options.dense ? (out[val[0].r]||[])[val[0].c] : out[encode_cell(val[0])];
					var noteobj = objects[val[2]];
					if(!cc) break;
					if(!cc.c) cc.c = [];
					cmnt = {a:val[1],t:noteobj.TxO.t};
					cc.c.push(cmnt);
				} break;

				default: switch(R.n) { /* nested */
				case 'ClrtClient': break;
				case 'XFExt': update_xfext(XFs[val.ixfe], val.ext); break;

				case 'DefColWidth': defwidth = val; break;
				case 'DefaultRowHeight': defheight = val[1]; break; // TODO: flags

				case 'ColInfo': {
					if(!opts.cellStyles) break;
					while(val.e >= val.s) {
						colinfo[val.e--] = { width: val.w/256 };
						if(!seencol) { seencol = true; find_mdw_colw(val.w/256); }
						process_col(colinfo[val.e+1]);
					}
				} break;
				case 'Row': {
					var rowobj = {};
					if(val.level != null) { rowinfo[val.r] = rowobj; rowobj.level = val.level; }
					if(val.hidden) { rowinfo[val.r] = rowobj; rowobj.hidden = true; }
					if(val.hpt) {
						rowinfo[val.r] = rowobj;
						rowobj.hpt = val.hpt; rowobj.hpx = pt2px(val.hpt);
					}
				} break;

				case 'LeftMargin':
				case 'RightMargin':
				case 'TopMargin':
				case 'BottomMargin':
					if(!out['!margins']) default_margins(out['!margins'] = {});
					out['!margins'][Rn.slice(0,-6).toLowerCase()] = val;
					break;

				case 'Setup': // TODO
					if(!out['!margins']) default_margins(out['!margins'] = {});
					out['!margins'].header = val.header;
					out['!margins'].footer = val.footer;
					break;

				case 'Window2': // TODO
					// $FlowIgnore
					if(val.RTL) Workbook.Views[0].RTL = true;
					break;

				case 'Header': break; // TODO
				case 'Footer': break; // TODO
				case 'HCenter': break; // TODO
				case 'VCenter': break; // TODO
				case 'Pls': break; // TODO
				case 'GCW': break;
				case 'LHRecord': break;
				case 'DBCell': break; // TODO
				case 'EntExU2': break; // TODO
				case 'SxView': break; // TODO
				case 'Sxvd': break; // TODO
				case 'SXVI': break; // TODO
				case 'SXVDEx': break; // TODO
				case 'SxIvd': break; // TODO
				case 'SXString': break; // TODO
				case 'Sync': break;
				case 'Addin': break;
				case 'SXDI': break; // TODO
				case 'SXLI': break; // TODO
				case 'SXEx': break; // TODO
				case 'QsiSXTag': break; // TODO
				case 'Selection': break;
				case 'Feat': break;
				case 'FeatHdr': case 'FeatHdr11': break;
				case 'Feature11': case 'Feature12': case 'List12': break;
				case 'Country': country = val; break;
				case 'RecalcId': break;
				case 'DxGCol': break; // TODO: htmlify
				case 'Fbi': case 'Fbi2': case 'GelFrame': break;
				case 'Font': break; // TODO
				case 'XFCRC': break; // TODO
				case 'Style': break; // TODO
				case 'StyleExt': break; // TODO
				case 'Palette': palette = val; break;
				case 'Theme': themes = val; break;
				/* Protection */
				case 'ScenarioProtect': break;
				case 'ObjProtect': break;

				/* Conditional Formatting */
				case 'CondFmt12': break;

				/* Table */
				case 'Table': break; // TODO
				case 'TableStyles': break; // TODO
				case 'TableStyle': break; // TODO
				case 'TableStyleElement': break; // TODO

				/* PivotTable */
				case 'SXStreamID': break; // TODO
				case 'SXVS': break; // TODO
				case 'DConRef': break; // TODO
				case 'SXAddl': break; // TODO
				case 'DConBin': break; // TODO
				case 'DConName': break; // TODO
				case 'SXPI': break; // TODO
				case 'SxFormat': break; // TODO
				case 'SxSelect': break; // TODO
				case 'SxRule': break; // TODO
				case 'SxFilt': break; // TODO
				case 'SxItm': break; // TODO
				case 'SxDXF': break; // TODO

				/* Scenario Manager */
				case 'ScenMan': break;

				/* Data Consolidation */
				case 'DCon': break;

				/* Watched Cell */
				case 'CellWatch': break;

				/* Print Settings */
				case 'PrintRowCol': break;
				case 'PrintGrid': break;
				case 'PrintSize': break;

				case 'XCT': break;
				case 'CRN': break;

				case 'Scl': {
					//console.log("Zoom Level:", val[0]/val[1],val);
				} break;
				case 'SheetExt': {
					/* empty */
				} break;
				case 'SheetExtOptional': {
					/* empty */
				} break;

				/* VBA */
				case 'ObNoMacros': {
					/* empty */
				} break;
				case 'ObProj': {
					/* empty */
				} break;
				case 'CodeName': {
if(!cur_sheet) Workbook.WBProps.CodeName = val || "ThisWorkbook";
					else wsprops.CodeName = val || wsprops.name;
				} break;
				case 'GUIDTypeLib': {
					/* empty */
				} break;

				case 'WOpt': break; // TODO: WTF?
				case 'PhoneticInfo': break;

				case 'OleObjectSize': break;

				/* Differential Formatting */
				case 'DXF': case 'DXFN': case 'DXFN12': case 'DXFN12List': case 'DXFN12NoCB': break;

				/* Data Validation */
				case 'Dv': case 'DVal': break;

				/* Data Series */
				case 'BRAI': case 'Series': case 'SeriesText': break;

				/* Data Connection */
				case 'DConn': break;
				case 'DbOrParamQry': break;
				case 'DBQueryExt': break;

				case 'OleDbConn': break;
				case 'ExtString': break;

				/* Formatting */
				case 'IFmtRecord': break;
				case 'CondFmt': case 'CF': case 'CF12': case 'CFEx': break;

				/* Explicitly Ignored */
				case 'Excel9File': break;
				case 'Units': break;
				case 'InterfaceHdr': case 'Mms': case 'InterfaceEnd': case 'DSF': break;
				case 'BuiltInFnGroupCount': /* 2.4.30 0x0E or 0x10 but excel 2011 generates 0x11? */ break;
				/* View Stuff */
				case 'Window1': case 'HideObj': case 'GridSet': case 'Guts':
				case 'UserBView': case 'UserSViewBegin': case 'UserSViewEnd':
				case 'Pane': break;
				default: switch(R.n) { /* nested */
				/* Chart */
				case 'Dat':
				case 'Begin': case 'End':
				case 'StartBlock': case 'EndBlock':
				case 'Frame': case 'Area':
				case 'Axis': case 'AxisLine': case 'Tick': break;
				case 'AxesUsed':
				case 'CrtLayout12': case 'CrtLayout12A': case 'CrtLink': case 'CrtLine': case 'CrtMlFrt': case 'CrtMlFrtContinue': break;
				case 'LineFormat': case 'AreaFormat':
				case 'Chart': case 'Chart3d': case 'Chart3DBarShape': case 'ChartFormat': case 'ChartFrtInfo': break;
				case 'PlotArea': case 'PlotGrowth': break;
				case 'SeriesList': case 'SerParent': case 'SerAuxTrend': break;
				case 'DataFormat': case 'SerToCrt': case 'FontX': break;
				case 'CatSerRange': case 'AxcExt': case 'SerFmt': break;
				case 'ShtProps': break;
				case 'DefaultText': case 'Text': case 'CatLab': break;
				case 'DataLabExtContents': break;
				case 'Legend': case 'LegendException': break;
				case 'Pie': case 'Scatter': break;
				case 'PieFormat': case 'MarkerFormat': break;
				case 'StartObject': case 'EndObject': break;
				case 'AlRuns': case 'ObjectLink': break;
				case 'SIIndex': break;
				case 'AttachedLabel': case 'YMult': break;

				/* Chart Group */
				case 'Line': case 'Bar': break;
				case 'Surf': break;

				/* Axis Group */
				case 'AxisParent': break;
				case 'Pos': break;
				case 'ValueRange': break;

				/* Pivot Chart */
				case 'SXViewEx9': break; // TODO
				case 'SXViewLink': break;
				case 'PivotChartBits': break;
				case 'SBaseRef': break;
				case 'TextPropsStream': break;

				/* Chart Misc */
				case 'LnExt': break;
				case 'MkrExt': break;
				case 'CrtCoopt': break;

				/* Query Table */
				case 'Qsi': case 'Qsif': case 'Qsir': case 'QsiSXTag': break;
				case 'TxtQry': break;

				/* Filter */
				case 'FilterMode': break;
				case 'AutoFilter': case 'AutoFilterInfo': break;
				case 'AutoFilter12': break;
				case 'DropDownObjIds': break;
				case 'Sort': break;
				case 'SortData': break;

				/* Drawing */
				case 'ShapePropsStream': break;
				case 'MsoDrawing': case 'MsoDrawingGroup': case 'MsoDrawingSelection': break;
				/* Pub Stuff */
				case 'WebPub': case 'AutoWebPub': break;

				/* Print Stuff */
				case 'HeaderFooter': case 'HFPicture': case 'PLV':
				case 'HorizontalPageBreaks': case 'VerticalPageBreaks': break;
				/* Behavioral */
				case 'Backup': case 'CompressPictures': case 'Compat12': break;

				/* Should not Happen */
				case 'Continue': case 'ContinueFrt12': break;

				/* Future Records */
				case 'FrtFontList': case 'FrtWrapper': break;

				default: switch(R.n) { /* nested */
				/* BIFF5 records */
				case 'TabIdConf': case 'Radar': case 'RadarArea': case 'DropBar': case 'Intl': case 'CoordList': case 'SerAuxErrBar': break;

				/* BIFF2-4 records */
				case 'BIFF2FONTCLR': case 'BIFF2FMTCNT': case 'BIFF2FONTXTRA': break;
				case 'BIFF2XF': case 'BIFF3XF': case 'BIFF4XF': break;
				case 'BIFF4FMTCNT': case 'BIFF2ROW': case 'BIFF2WINDOW2': break;

				/* Miscellaneous */
				case 'SCENARIO': case 'DConBin': case 'PicF': case 'DataLabExt':
				case 'Lel': case 'BopPop': case 'BopPopCustom': case 'RealTimeData':
				case 'Name': break;
				case 'LHNGraph': case 'FnGroupName': case 'AddMenu': case 'LPr': break;
				case 'ListObj': case 'ListField': break;
				case 'RRSort': break;
				case 'BigName': break;
				case 'ToolbarHdr': case 'ToolbarEnd': break;
				case 'DDEObjName': break;
				case 'FRTArchId$': break;
				default: if(options.WTF) throw 'Unrecognized Record ' + R.n;
			}}}}
		} else blob.l += length;
	}
	wb.SheetNames=keys(Directory).sort(function(a,b) { return Number(a) - Number(b); }).map(function(x){return Directory[x].name;});
	if(!options.bookSheets) wb.Sheets=Sheets;
	if(wb.Sheets) FilterDatabases.forEach(function(r,i) { wb.Sheets[wb.SheetNames[i]]['!autofilter'] = r; });
	wb.Preamble=Preamble;
	wb.Strings = sst;
	wb.SSF = SSF.get_table();
	if(opts.enc) wb.Encryption = opts.enc;
	if(themes) wb.Themes = themes;
	wb.Metadata = {};
	if(country !== undefined) wb.Metadata.Country = country;
	if(supbooks.names.length > 0) Workbook.Names = supbooks.names;
	wb.Workbook = Workbook;
	return wb;
}

/* TODO: split props*/
var PSCLSID = {
	SI: "e0859ff2f94f6810ab9108002b27b3d9",
	DSI: "02d5cdd59c2e1b10939708002b2cf9ae",
	UDI: "05d5cdd59c2e1b10939708002b2cf9ae"
};
function parse_xls_props(cfb, props, o) {
	/* [MS-OSHARED] 2.3.3.2.2 Document Summary Information Property Set */
	var DSI = CFB.find(cfb, '!DocumentSummaryInformation');
	if(DSI && DSI.size > 0) try {
		var DocSummary = parse_PropertySetStream(DSI, DocSummaryPIDDSI, PSCLSID.DSI);
		for(var d in DocSummary) props[d] = DocSummary[d];
	} catch(e) {if(o.WTF) throw e;/* empty */}

	/* [MS-OSHARED] 2.3.3.2.1 Summary Information Property Set*/
	var SI = CFB.find(cfb, '!SummaryInformation');
	if(SI && SI.size > 0) try {
		var Summary = parse_PropertySetStream(SI, SummaryPIDSI, PSCLSID.SI);
		for(var s in Summary) if(props[s] == null) props[s] = Summary[s];
	} catch(e) {if(o.WTF) throw e;/* empty */}

	if(props.HeadingPairs && props.TitlesOfParts) {
		load_props_pairs(props.HeadingPairs, props.TitlesOfParts, props, o);
		delete props.HeadingPairs; delete props.TitlesOfParts;
	}
}
function write_xls_props(wb, cfb) {
	var DSEntries = [], SEntries = [], CEntries = [];
	var i = 0, Keys;
	if(wb.Props) {
		Keys = keys(wb.Props);
		// $FlowIgnore
		for(i = 0; i < Keys.length; ++i) (DocSummaryRE.hasOwnProperty(Keys[i]) ? DSEntries : SummaryRE.hasOwnProperty(Keys[i]) ? SEntries : CEntries).push([Keys[i], wb.Props[Keys[i]]]);
	}
	if(wb.Custprops) {
		Keys = keys(wb.Custprops);
		// $FlowIgnore
		for(i = 0; i < Keys.length; ++i) if(!(wb.Props||{}).hasOwnProperty(Keys[i])) (DocSummaryRE.hasOwnProperty(Keys[i]) ? DSEntries : SummaryRE.hasOwnProperty(Keys[i]) ? SEntries : CEntries).push([Keys[i], wb.Custprops[Keys[i]]]);
	}
	var CEntries2 = [];
	for(i = 0; i < CEntries.length; ++i) {
		if(XLSPSSkip.indexOf(CEntries[i][0]) > -1) continue;
		if(CEntries[i][1] == null) continue;
		CEntries2.push(CEntries[i]);
	}
	if(SEntries.length) CFB.utils.cfb_add(cfb, "/\u0005SummaryInformation", write_PropertySetStream(SEntries, PSCLSID.SI, SummaryRE, SummaryPIDSI));
	if(DSEntries.length || CEntries2.length) CFB.utils.cfb_add(cfb, "/\u0005DocumentSummaryInformation", write_PropertySetStream(DSEntries, PSCLSID.DSI, DocSummaryRE, DocSummaryPIDDSI, CEntries2.length ? CEntries2 : null, PSCLSID.UDI));
}

function parse_xlscfb(cfb, options) {
if(!options) options = {};
fix_read_opts(options);
reset_cp();
if(options.codepage) set_ansi(options.codepage);
var CompObj, WB;
if(cfb.FullPaths) {
	if(CFB.find(cfb, '/encryption')) throw new Error("File is password-protected");
	CompObj = CFB.find(cfb, '!CompObj');
	WB = CFB.find(cfb, '/Workbook') || CFB.find(cfb, '/Book');
} else {
	switch(options.type) {
		case 'base64': cfb = s2a(Base64.decode(cfb)); break;
		case 'binary': cfb = s2a(cfb); break;
		case 'buffer': break;
		case 'array': if(!Array.isArray(cfb)) cfb = Array.prototype.slice.call(cfb); break;
	}
	prep_blob(cfb, 0);
	WB = ({content: cfb});
}
var WorkbookP;

var _data;
if(CompObj) parse_compobj(CompObj);
if(options.bookProps && !options.bookSheets) WorkbookP = ({});
else {
	var T = has_buf ? 'buffer' : 'array';
	if(WB && WB.content) WorkbookP = parse_workbook(WB.content, options);
	/* Quattro Pro 7-8 */
	else if((_data=CFB.find(cfb, 'PerfectOffice_MAIN')) && _data.content) WorkbookP = WK_.to_workbook(_data.content, (options.type = T, options));
	/* Quattro Pro 9 */
	else if((_data=CFB.find(cfb, 'NativeContent_MAIN')) && _data.content) WorkbookP = WK_.to_workbook(_data.content, (options.type = T, options));
	else throw new Error("Cannot find Workbook stream");
	if(options.bookVBA && cfb.FullPaths && CFB.find(cfb, '/_VBA_PROJECT_CUR/VBA/dir')) WorkbookP.vbaraw = make_vba_xls(cfb);
}

var props = {};
if(cfb.FullPaths) parse_xls_props(cfb, props, options);

WorkbookP.Props = WorkbookP.Custprops = props; /* TODO: split up properties */
if(options.bookFiles) WorkbookP.cfb = cfb;
/*WorkbookP.CompObjP = CompObjP; // TODO: storage? */
return WorkbookP;
}


function write_xlscfb(wb, opts) {
	var o = opts || {};
	var cfb = CFB.utils.cfb_new({root:"R"});
	var wbpath = "/Workbook";
	switch(o.bookType || "xls") {
		case "xls": o.bookType = "biff8";
		/* falls through */
		case "xla": if(!o.bookType) o.bookType = "xla";
		/* falls through */
		case "biff8": wbpath = "/Workbook"; o.biff = 8; break;
		case "biff5": wbpath = "/Book"; o.biff = 5; break;
		default: throw new Error("invalid type " + o.bookType + " for XLS CFB");
	}
	CFB.utils.cfb_add(cfb, wbpath, write_biff_buf(wb, o));
	if(o.biff == 8 && (wb.Props || wb.Custprops)) write_xls_props(wb, cfb);
	// TODO: SI, DSI, CO
	if(o.biff == 8 && wb.vbaraw) fill_vba_xls(cfb, CFB.read(wb.vbaraw, {type: typeof wb.vbaraw == "string" ? "binary" : "buffer"}));
	return cfb;
}
/* [MS-XLSB] 2.3 Record Enumeration */
var XLSBRecordEnum = {
0x0000: { n:"BrtRowHdr", f:parse_BrtRowHdr },
0x0001: { n:"BrtCellBlank", f:parse_BrtCellBlank },
0x0002: { n:"BrtCellRk", f:parse_BrtCellRk },
0x0003: { n:"BrtCellError", f:parse_BrtCellError },
0x0004: { n:"BrtCellBool", f:parse_BrtCellBool },
0x0005: { n:"BrtCellReal", f:parse_BrtCellReal },
0x0006: { n:"BrtCellSt", f:parse_BrtCellSt },
0x0007: { n:"BrtCellIsst", f:parse_BrtCellIsst },
0x0008: { n:"BrtFmlaString", f:parse_BrtFmlaString },
0x0009: { n:"BrtFmlaNum", f:parse_BrtFmlaNum },
0x000A: { n:"BrtFmlaBool", f:parse_BrtFmlaBool },
0x000B: { n:"BrtFmlaError", f:parse_BrtFmlaError },
0x0010: { n:"BrtFRTArchID$", f:parse_BrtFRTArchID$ },
0x0013: { n:"BrtSSTItem", f:parse_RichStr },
0x0014: { n:"BrtPCDIMissing" },
0x0015: { n:"BrtPCDINumber" },
0x0016: { n:"BrtPCDIBoolean" },
0x0017: { n:"BrtPCDIError" },
0x0018: { n:"BrtPCDIString" },
0x0019: { n:"BrtPCDIDatetime" },
0x001A: { n:"BrtPCDIIndex" },
0x001B: { n:"BrtPCDIAMissing" },
0x001C: { n:"BrtPCDIANumber" },
0x001D: { n:"BrtPCDIABoolean" },
0x001E: { n:"BrtPCDIAError" },
0x001F: { n:"BrtPCDIAString" },
0x0020: { n:"BrtPCDIADatetime" },
0x0021: { n:"BrtPCRRecord" },
0x0022: { n:"BrtPCRRecordDt" },
0x0023: { n:"BrtFRTBegin" },
0x0024: { n:"BrtFRTEnd" },
0x0025: { n:"BrtACBegin" },
0x0026: { n:"BrtACEnd" },
0x0027: { n:"BrtName", f:parse_BrtName },
0x0028: { n:"BrtIndexRowBlock" },
0x002A: { n:"BrtIndexBlock" },
0x002B: { n:"BrtFont", f:parse_BrtFont },
0x002C: { n:"BrtFmt", f:parse_BrtFmt },
0x002D: { n:"BrtFill", f:parse_BrtFill },
0x002E: { n:"BrtBorder", f:parse_BrtBorder },
0x002F: { n:"BrtXF", f:parse_BrtXF },
0x0030: { n:"BrtStyle" },
0x0031: { n:"BrtCellMeta" },
0x0032: { n:"BrtValueMeta" },
0x0033: { n:"BrtMdb" },
0x0034: { n:"BrtBeginFmd" },
0x0035: { n:"BrtEndFmd" },
0x0036: { n:"BrtBeginMdx" },
0x0037: { n:"BrtEndMdx" },
0x0038: { n:"BrtBeginMdxTuple" },
0x0039: { n:"BrtEndMdxTuple" },
0x003A: { n:"BrtMdxMbrIstr" },
0x003B: { n:"BrtStr" },
0x003C: { n:"BrtColInfo", f:parse_ColInfo },
0x003E: { n:"BrtCellRString" },
0x003F: { n:"BrtCalcChainItem$", f:parse_BrtCalcChainItem$ },
0x0040: { n:"BrtDVal" },
0x0041: { n:"BrtSxvcellNum" },
0x0042: { n:"BrtSxvcellStr" },
0x0043: { n:"BrtSxvcellBool" },
0x0044: { n:"BrtSxvcellErr" },
0x0045: { n:"BrtSxvcellDate" },
0x0046: { n:"BrtSxvcellNil" },
0x0080: { n:"BrtFileVersion" },
0x0081: { n:"BrtBeginSheet" },
0x0082: { n:"BrtEndSheet" },
0x0083: { n:"BrtBeginBook", f:parsenoop, p:0 },
0x0084: { n:"BrtEndBook" },
0x0085: { n:"BrtBeginWsViews" },
0x0086: { n:"BrtEndWsViews" },
0x0087: { n:"BrtBeginBookViews" },
0x0088: { n:"BrtEndBookViews" },
0x0089: { n:"BrtBeginWsView", f:parse_BrtBeginWsView },
0x008A: { n:"BrtEndWsView" },
0x008B: { n:"BrtBeginCsViews" },
0x008C: { n:"BrtEndCsViews" },
0x008D: { n:"BrtBeginCsView" },
0x008E: { n:"BrtEndCsView" },
0x008F: { n:"BrtBeginBundleShs" },
0x0090: { n:"BrtEndBundleShs" },
0x0091: { n:"BrtBeginSheetData" },
0x0092: { n:"BrtEndSheetData" },
0x0093: { n:"BrtWsProp", f:parse_BrtWsProp },
0x0094: { n:"BrtWsDim", f:parse_BrtWsDim, p:16 },
0x0097: { n:"BrtPane" },
0x0098: { n:"BrtSel" },
0x0099: { n:"BrtWbProp", f:parse_BrtWbProp },
0x009A: { n:"BrtWbFactoid" },
0x009B: { n:"BrtFileRecover" },
0x009C: { n:"BrtBundleSh", f:parse_BrtBundleSh },
0x009D: { n:"BrtCalcProp" },
0x009E: { n:"BrtBookView" },
0x009F: { n:"BrtBeginSst", f:parse_BrtBeginSst },
0x00A0: { n:"BrtEndSst" },
0x00A1: { n:"BrtBeginAFilter", f:parse_UncheckedRfX },
0x00A2: { n:"BrtEndAFilter" },
0x00A3: { n:"BrtBeginFilterColumn" },
0x00A4: { n:"BrtEndFilterColumn" },
0x00A5: { n:"BrtBeginFilters" },
0x00A6: { n:"BrtEndFilters" },
0x00A7: { n:"BrtFilter" },
0x00A8: { n:"BrtColorFilter" },
0x00A9: { n:"BrtIconFilter" },
0x00AA: { n:"BrtTop10Filter" },
0x00AB: { n:"BrtDynamicFilter" },
0x00AC: { n:"BrtBeginCustomFilters" },
0x00AD: { n:"BrtEndCustomFilters" },
0x00AE: { n:"BrtCustomFilter" },
0x00AF: { n:"BrtAFilterDateGroupItem" },
0x00B0: { n:"BrtMergeCell", f:parse_BrtMergeCell },
0x00B1: { n:"BrtBeginMergeCells" },
0x00B2: { n:"BrtEndMergeCells" },
0x00B3: { n:"BrtBeginPivotCacheDef" },
0x00B4: { n:"BrtEndPivotCacheDef" },
0x00B5: { n:"BrtBeginPCDFields" },
0x00B6: { n:"BrtEndPCDFields" },
0x00B7: { n:"BrtBeginPCDField" },
0x00B8: { n:"BrtEndPCDField" },
0x00B9: { n:"BrtBeginPCDSource" },
0x00BA: { n:"BrtEndPCDSource" },
0x00BB: { n:"BrtBeginPCDSRange" },
0x00BC: { n:"BrtEndPCDSRange" },
0x00BD: { n:"BrtBeginPCDFAtbl" },
0x00BE: { n:"BrtEndPCDFAtbl" },
0x00BF: { n:"BrtBeginPCDIRun" },
0x00C0: { n:"BrtEndPCDIRun" },
0x00C1: { n:"BrtBeginPivotCacheRecords" },
0x00C2: { n:"BrtEndPivotCacheRecords" },
0x00C3: { n:"BrtBeginPCDHierarchies" },
0x00C4: { n:"BrtEndPCDHierarchies" },
0x00C5: { n:"BrtBeginPCDHierarchy" },
0x00C6: { n:"BrtEndPCDHierarchy" },
0x00C7: { n:"BrtBeginPCDHFieldsUsage" },
0x00C8: { n:"BrtEndPCDHFieldsUsage" },
0x00C9: { n:"BrtBeginExtConnection" },
0x00CA: { n:"BrtEndExtConnection" },
0x00CB: { n:"BrtBeginECDbProps" },
0x00CC: { n:"BrtEndECDbProps" },
0x00CD: { n:"BrtBeginECOlapProps" },
0x00CE: { n:"BrtEndECOlapProps" },
0x00CF: { n:"BrtBeginPCDSConsol" },
0x00D0: { n:"BrtEndPCDSConsol" },
0x00D1: { n:"BrtBeginPCDSCPages" },
0x00D2: { n:"BrtEndPCDSCPages" },
0x00D3: { n:"BrtBeginPCDSCPage" },
0x00D4: { n:"BrtEndPCDSCPage" },
0x00D5: { n:"BrtBeginPCDSCPItem" },
0x00D6: { n:"BrtEndPCDSCPItem" },
0x00D7: { n:"BrtBeginPCDSCSets" },
0x00D8: { n:"BrtEndPCDSCSets" },
0x00D9: { n:"BrtBeginPCDSCSet" },
0x00DA: { n:"BrtEndPCDSCSet" },
0x00DB: { n:"BrtBeginPCDFGroup" },
0x00DC: { n:"BrtEndPCDFGroup" },
0x00DD: { n:"BrtBeginPCDFGItems" },
0x00DE: { n:"BrtEndPCDFGItems" },
0x00DF: { n:"BrtBeginPCDFGRange" },
0x00E0: { n:"BrtEndPCDFGRange" },
0x00E1: { n:"BrtBeginPCDFGDiscrete" },
0x00E2: { n:"BrtEndPCDFGDiscrete" },
0x00E3: { n:"BrtBeginPCDSDTupleCache" },
0x00E4: { n:"BrtEndPCDSDTupleCache" },
0x00E5: { n:"BrtBeginPCDSDTCEntries" },
0x00E6: { n:"BrtEndPCDSDTCEntries" },
0x00E7: { n:"BrtBeginPCDSDTCEMembers" },
0x00E8: { n:"BrtEndPCDSDTCEMembers" },
0x00E9: { n:"BrtBeginPCDSDTCEMember" },
0x00EA: { n:"BrtEndPCDSDTCEMember" },
0x00EB: { n:"BrtBeginPCDSDTCQueries" },
0x00EC: { n:"BrtEndPCDSDTCQueries" },
0x00ED: { n:"BrtBeginPCDSDTCQuery" },
0x00EE: { n:"BrtEndPCDSDTCQuery" },
0x00EF: { n:"BrtBeginPCDSDTCSets" },
0x00F0: { n:"BrtEndPCDSDTCSets" },
0x00F1: { n:"BrtBeginPCDSDTCSet" },
0x00F2: { n:"BrtEndPCDSDTCSet" },
0x00F3: { n:"BrtBeginPCDCalcItems" },
0x00F4: { n:"BrtEndPCDCalcItems" },
0x00F5: { n:"BrtBeginPCDCalcItem" },
0x00F6: { n:"BrtEndPCDCalcItem" },
0x00F7: { n:"BrtBeginPRule" },
0x00F8: { n:"BrtEndPRule" },
0x00F9: { n:"BrtBeginPRFilters" },
0x00FA: { n:"BrtEndPRFilters" },
0x00FB: { n:"BrtBeginPRFilter" },
0x00FC: { n:"BrtEndPRFilter" },
0x00FD: { n:"BrtBeginPNames" },
0x00FE: { n:"BrtEndPNames" },
0x00FF: { n:"BrtBeginPName" },
0x0100: { n:"BrtEndPName" },
0x0101: { n:"BrtBeginPNPairs" },
0x0102: { n:"BrtEndPNPairs" },
0x0103: { n:"BrtBeginPNPair" },
0x0104: { n:"BrtEndPNPair" },
0x0105: { n:"BrtBeginECWebProps" },
0x0106: { n:"BrtEndECWebProps" },
0x0107: { n:"BrtBeginEcWpTables" },
0x0108: { n:"BrtEndECWPTables" },
0x0109: { n:"BrtBeginECParams" },
0x010A: { n:"BrtEndECParams" },
0x010B: { n:"BrtBeginECParam" },
0x010C: { n:"BrtEndECParam" },
0x010D: { n:"BrtBeginPCDKPIs" },
0x010E: { n:"BrtEndPCDKPIs" },
0x010F: { n:"BrtBeginPCDKPI" },
0x0110: { n:"BrtEndPCDKPI" },
0x0111: { n:"BrtBeginDims" },
0x0112: { n:"BrtEndDims" },
0x0113: { n:"BrtBeginDim" },
0x0114: { n:"BrtEndDim" },
0x0115: { n:"BrtIndexPartEnd" },
0x0116: { n:"BrtBeginStyleSheet" },
0x0117: { n:"BrtEndStyleSheet" },
0x0118: { n:"BrtBeginSXView" },
0x0119: { n:"BrtEndSXVI" },
0x011A: { n:"BrtBeginSXVI" },
0x011B: { n:"BrtBeginSXVIs" },
0x011C: { n:"BrtEndSXVIs" },
0x011D: { n:"BrtBeginSXVD" },
0x011E: { n:"BrtEndSXVD" },
0x011F: { n:"BrtBeginSXVDs" },
0x0120: { n:"BrtEndSXVDs" },
0x0121: { n:"BrtBeginSXPI" },
0x0122: { n:"BrtEndSXPI" },
0x0123: { n:"BrtBeginSXPIs" },
0x0124: { n:"BrtEndSXPIs" },
0x0125: { n:"BrtBeginSXDI" },
0x0126: { n:"BrtEndSXDI" },
0x0127: { n:"BrtBeginSXDIs" },
0x0128: { n:"BrtEndSXDIs" },
0x0129: { n:"BrtBeginSXLI" },
0x012A: { n:"BrtEndSXLI" },
0x012B: { n:"BrtBeginSXLIRws" },
0x012C: { n:"BrtEndSXLIRws" },
0x012D: { n:"BrtBeginSXLICols" },
0x012E: { n:"BrtEndSXLICols" },
0x012F: { n:"BrtBeginSXFormat" },
0x0130: { n:"BrtEndSXFormat" },
0x0131: { n:"BrtBeginSXFormats" },
0x0132: { n:"BrtEndSxFormats" },
0x0133: { n:"BrtBeginSxSelect" },
0x0134: { n:"BrtEndSxSelect" },
0x0135: { n:"BrtBeginISXVDRws" },
0x0136: { n:"BrtEndISXVDRws" },
0x0137: { n:"BrtBeginISXVDCols" },
0x0138: { n:"BrtEndISXVDCols" },
0x0139: { n:"BrtEndSXLocation" },
0x013A: { n:"BrtBeginSXLocation" },
0x013B: { n:"BrtEndSXView" },
0x013C: { n:"BrtBeginSXTHs" },
0x013D: { n:"BrtEndSXTHs" },
0x013E: { n:"BrtBeginSXTH" },
0x013F: { n:"BrtEndSXTH" },
0x0140: { n:"BrtBeginISXTHRws" },
0x0141: { n:"BrtEndISXTHRws" },
0x0142: { n:"BrtBeginISXTHCols" },
0x0143: { n:"BrtEndISXTHCols" },
0x0144: { n:"BrtBeginSXTDMPS" },
0x0145: { n:"BrtEndSXTDMPs" },
0x0146: { n:"BrtBeginSXTDMP" },
0x0147: { n:"BrtEndSXTDMP" },
0x0148: { n:"BrtBeginSXTHItems" },
0x0149: { n:"BrtEndSXTHItems" },
0x014A: { n:"BrtBeginSXTHItem" },
0x014B: { n:"BrtEndSXTHItem" },
0x014C: { n:"BrtBeginMetadata" },
0x014D: { n:"BrtEndMetadata" },
0x014E: { n:"BrtBeginEsmdtinfo" },
0x014F: { n:"BrtMdtinfo" },
0x0150: { n:"BrtEndEsmdtinfo" },
0x0151: { n:"BrtBeginEsmdb" },
0x0152: { n:"BrtEndEsmdb" },
0x0153: { n:"BrtBeginEsfmd" },
0x0154: { n:"BrtEndEsfmd" },
0x0155: { n:"BrtBeginSingleCells" },
0x0156: { n:"BrtEndSingleCells" },
0x0157: { n:"BrtBeginList" },
0x0158: { n:"BrtEndList" },
0x0159: { n:"BrtBeginListCols" },
0x015A: { n:"BrtEndListCols" },
0x015B: { n:"BrtBeginListCol" },
0x015C: { n:"BrtEndListCol" },
0x015D: { n:"BrtBeginListXmlCPr" },
0x015E: { n:"BrtEndListXmlCPr" },
0x015F: { n:"BrtListCCFmla" },
0x0160: { n:"BrtListTrFmla" },
0x0161: { n:"BrtBeginExternals" },
0x0162: { n:"BrtEndExternals" },
0x0163: { n:"BrtSupBookSrc", f:parse_RelID},
0x0165: { n:"BrtSupSelf" },
0x0166: { n:"BrtSupSame" },
0x0167: { n:"BrtSupTabs" },
0x0168: { n:"BrtBeginSupBook" },
0x0169: { n:"BrtPlaceholderName" },
0x016A: { n:"BrtExternSheet", f:parse_ExternSheet },
0x016B: { n:"BrtExternTableStart" },
0x016C: { n:"BrtExternTableEnd" },
0x016E: { n:"BrtExternRowHdr" },
0x016F: { n:"BrtExternCellBlank" },
0x0170: { n:"BrtExternCellReal" },
0x0171: { n:"BrtExternCellBool" },
0x0172: { n:"BrtExternCellError" },
0x0173: { n:"BrtExternCellString" },
0x0174: { n:"BrtBeginEsmdx" },
0x0175: { n:"BrtEndEsmdx" },
0x0176: { n:"BrtBeginMdxSet" },
0x0177: { n:"BrtEndMdxSet" },
0x0178: { n:"BrtBeginMdxMbrProp" },
0x0179: { n:"BrtEndMdxMbrProp" },
0x017A: { n:"BrtBeginMdxKPI" },
0x017B: { n:"BrtEndMdxKPI" },
0x017C: { n:"BrtBeginEsstr" },
0x017D: { n:"BrtEndEsstr" },
0x017E: { n:"BrtBeginPRFItem" },
0x017F: { n:"BrtEndPRFItem" },
0x0180: { n:"BrtBeginPivotCacheIDs" },
0x0181: { n:"BrtEndPivotCacheIDs" },
0x0182: { n:"BrtBeginPivotCacheID" },
0x0183: { n:"BrtEndPivotCacheID" },
0x0184: { n:"BrtBeginISXVIs" },
0x0185: { n:"BrtEndISXVIs" },
0x0186: { n:"BrtBeginColInfos" },
0x0187: { n:"BrtEndColInfos" },
0x0188: { n:"BrtBeginRwBrk" },
0x0189: { n:"BrtEndRwBrk" },
0x018A: { n:"BrtBeginColBrk" },
0x018B: { n:"BrtEndColBrk" },
0x018C: { n:"BrtBrk" },
0x018D: { n:"BrtUserBookView" },
0x018E: { n:"BrtInfo" },
0x018F: { n:"BrtCUsr" },
0x0190: { n:"BrtUsr" },
0x0191: { n:"BrtBeginUsers" },
0x0193: { n:"BrtEOF" },
0x0194: { n:"BrtUCR" },
0x0195: { n:"BrtRRInsDel" },
0x0196: { n:"BrtRREndInsDel" },
0x0197: { n:"BrtRRMove" },
0x0198: { n:"BrtRREndMove" },
0x0199: { n:"BrtRRChgCell" },
0x019A: { n:"BrtRREndChgCell" },
0x019B: { n:"BrtRRHeader" },
0x019C: { n:"BrtRRUserView" },
0x019D: { n:"BrtRRRenSheet" },
0x019E: { n:"BrtRRInsertSh" },
0x019F: { n:"BrtRRDefName" },
0x01A0: { n:"BrtRRNote" },
0x01A1: { n:"BrtRRConflict" },
0x01A2: { n:"BrtRRTQSIF" },
0x01A3: { n:"BrtRRFormat" },
0x01A4: { n:"BrtRREndFormat" },
0x01A5: { n:"BrtRRAutoFmt" },
0x01A6: { n:"BrtBeginUserShViews" },
0x01A7: { n:"BrtBeginUserShView" },
0x01A8: { n:"BrtEndUserShView" },
0x01A9: { n:"BrtEndUserShViews" },
0x01AA: { n:"BrtArrFmla", f:parse_BrtArrFmla },
0x01AB: { n:"BrtShrFmla", f:parse_BrtShrFmla },
0x01AC: { n:"BrtTable" },
0x01AD: { n:"BrtBeginExtConnections" },
0x01AE: { n:"BrtEndExtConnections" },
0x01AF: { n:"BrtBeginPCDCalcMems" },
0x01B0: { n:"BrtEndPCDCalcMems" },
0x01B1: { n:"BrtBeginPCDCalcMem" },
0x01B2: { n:"BrtEndPCDCalcMem" },
0x01B3: { n:"BrtBeginPCDHGLevels" },
0x01B4: { n:"BrtEndPCDHGLevels" },
0x01B5: { n:"BrtBeginPCDHGLevel" },
0x01B6: { n:"BrtEndPCDHGLevel" },
0x01B7: { n:"BrtBeginPCDHGLGroups" },
0x01B8: { n:"BrtEndPCDHGLGroups" },
0x01B9: { n:"BrtBeginPCDHGLGroup" },
0x01BA: { n:"BrtEndPCDHGLGroup" },
0x01BB: { n:"BrtBeginPCDHGLGMembers" },
0x01BC: { n:"BrtEndPCDHGLGMembers" },
0x01BD: { n:"BrtBeginPCDHGLGMember" },
0x01BE: { n:"BrtEndPCDHGLGMember" },
0x01BF: { n:"BrtBeginQSI" },
0x01C0: { n:"BrtEndQSI" },
0x01C1: { n:"BrtBeginQSIR" },
0x01C2: { n:"BrtEndQSIR" },
0x01C3: { n:"BrtBeginDeletedNames" },
0x01C4: { n:"BrtEndDeletedNames" },
0x01C5: { n:"BrtBeginDeletedName" },
0x01C6: { n:"BrtEndDeletedName" },
0x01C7: { n:"BrtBeginQSIFs" },
0x01C8: { n:"BrtEndQSIFs" },
0x01C9: { n:"BrtBeginQSIF" },
0x01CA: { n:"BrtEndQSIF" },
0x01CB: { n:"BrtBeginAutoSortScope" },
0x01CC: { n:"BrtEndAutoSortScope" },
0x01CD: { n:"BrtBeginConditionalFormatting" },
0x01CE: { n:"BrtEndConditionalFormatting" },
0x01CF: { n:"BrtBeginCFRule" },
0x01D0: { n:"BrtEndCFRule" },
0x01D1: { n:"BrtBeginIconSet" },
0x01D2: { n:"BrtEndIconSet" },
0x01D3: { n:"BrtBeginDatabar" },
0x01D4: { n:"BrtEndDatabar" },
0x01D5: { n:"BrtBeginColorScale" },
0x01D6: { n:"BrtEndColorScale" },
0x01D7: { n:"BrtCFVO" },
0x01D8: { n:"BrtExternValueMeta" },
0x01D9: { n:"BrtBeginColorPalette" },
0x01DA: { n:"BrtEndColorPalette" },
0x01DB: { n:"BrtIndexedColor" },
0x01DC: { n:"BrtMargins", f:parse_BrtMargins },
0x01DD: { n:"BrtPrintOptions" },
0x01DE: { n:"BrtPageSetup" },
0x01DF: { n:"BrtBeginHeaderFooter" },
0x01E0: { n:"BrtEndHeaderFooter" },
0x01E1: { n:"BrtBeginSXCrtFormat" },
0x01E2: { n:"BrtEndSXCrtFormat" },
0x01E3: { n:"BrtBeginSXCrtFormats" },
0x01E4: { n:"BrtEndSXCrtFormats" },
0x01E5: { n:"BrtWsFmtInfo", f:parse_BrtWsFmtInfo },
0x01E6: { n:"BrtBeginMgs" },
0x01E7: { n:"BrtEndMGs" },
0x01E8: { n:"BrtBeginMGMaps" },
0x01E9: { n:"BrtEndMGMaps" },
0x01EA: { n:"BrtBeginMG" },
0x01EB: { n:"BrtEndMG" },
0x01EC: { n:"BrtBeginMap" },
0x01ED: { n:"BrtEndMap" },
0x01EE: { n:"BrtHLink", f:parse_BrtHLink },
0x01EF: { n:"BrtBeginDCon" },
0x01F0: { n:"BrtEndDCon" },
0x01F1: { n:"BrtBeginDRefs" },
0x01F2: { n:"BrtEndDRefs" },
0x01F3: { n:"BrtDRef" },
0x01F4: { n:"BrtBeginScenMan" },
0x01F5: { n:"BrtEndScenMan" },
0x01F6: { n:"BrtBeginSct" },
0x01F7: { n:"BrtEndSct" },
0x01F8: { n:"BrtSlc" },
0x01F9: { n:"BrtBeginDXFs" },
0x01FA: { n:"BrtEndDXFs" },
0x01FB: { n:"BrtDXF" },
0x01FC: { n:"BrtBeginTableStyles" },
0x01FD: { n:"BrtEndTableStyles" },
0x01FE: { n:"BrtBeginTableStyle" },
0x01FF: { n:"BrtEndTableStyle" },
0x0200: { n:"BrtTableStyleElement" },
0x0201: { n:"BrtTableStyleClient" },
0x0202: { n:"BrtBeginVolDeps" },
0x0203: { n:"BrtEndVolDeps" },
0x0204: { n:"BrtBeginVolType" },
0x0205: { n:"BrtEndVolType" },
0x0206: { n:"BrtBeginVolMain" },
0x0207: { n:"BrtEndVolMain" },
0x0208: { n:"BrtBeginVolTopic" },
0x0209: { n:"BrtEndVolTopic" },
0x020A: { n:"BrtVolSubtopic" },
0x020B: { n:"BrtVolRef" },
0x020C: { n:"BrtVolNum" },
0x020D: { n:"BrtVolErr" },
0x020E: { n:"BrtVolStr" },
0x020F: { n:"BrtVolBool" },
0x0210: { n:"BrtBeginCalcChain$" },
0x0211: { n:"BrtEndCalcChain$" },
0x0212: { n:"BrtBeginSortState" },
0x0213: { n:"BrtEndSortState" },
0x0214: { n:"BrtBeginSortCond" },
0x0215: { n:"BrtEndSortCond" },
0x0216: { n:"BrtBookProtection" },
0x0217: { n:"BrtSheetProtection" },
0x0218: { n:"BrtRangeProtection" },
0x0219: { n:"BrtPhoneticInfo" },
0x021A: { n:"BrtBeginECTxtWiz" },
0x021B: { n:"BrtEndECTxtWiz" },
0x021C: { n:"BrtBeginECTWFldInfoLst" },
0x021D: { n:"BrtEndECTWFldInfoLst" },
0x021E: { n:"BrtBeginECTwFldInfo" },
0x0224: { n:"BrtFileSharing" },
0x0225: { n:"BrtOleSize" },
0x0226: { n:"BrtDrawing", f:parse_RelID },
0x0227: { n:"BrtLegacyDrawing" },
0x0228: { n:"BrtLegacyDrawingHF" },
0x0229: { n:"BrtWebOpt" },
0x022A: { n:"BrtBeginWebPubItems" },
0x022B: { n:"BrtEndWebPubItems" },
0x022C: { n:"BrtBeginWebPubItem" },
0x022D: { n:"BrtEndWebPubItem" },
0x022E: { n:"BrtBeginSXCondFmt" },
0x022F: { n:"BrtEndSXCondFmt" },
0x0230: { n:"BrtBeginSXCondFmts" },
0x0231: { n:"BrtEndSXCondFmts" },
0x0232: { n:"BrtBkHim" },
0x0234: { n:"BrtColor" },
0x0235: { n:"BrtBeginIndexedColors" },
0x0236: { n:"BrtEndIndexedColors" },
0x0239: { n:"BrtBeginMRUColors" },
0x023A: { n:"BrtEndMRUColors" },
0x023C: { n:"BrtMRUColor" },
0x023D: { n:"BrtBeginDVals" },
0x023E: { n:"BrtEndDVals" },
0x0241: { n:"BrtSupNameStart" },
0x0242: { n:"BrtSupNameValueStart" },
0x0243: { n:"BrtSupNameValueEnd" },
0x0244: { n:"BrtSupNameNum" },
0x0245: { n:"BrtSupNameErr" },
0x0246: { n:"BrtSupNameSt" },
0x0247: { n:"BrtSupNameNil" },
0x0248: { n:"BrtSupNameBool" },
0x0249: { n:"BrtSupNameFmla" },
0x024A: { n:"BrtSupNameBits" },
0x024B: { n:"BrtSupNameEnd" },
0x024C: { n:"BrtEndSupBook" },
0x024D: { n:"BrtCellSmartTagProperty" },
0x024E: { n:"BrtBeginCellSmartTag" },
0x024F: { n:"BrtEndCellSmartTag" },
0x0250: { n:"BrtBeginCellSmartTags" },
0x0251: { n:"BrtEndCellSmartTags" },
0x0252: { n:"BrtBeginSmartTags" },
0x0253: { n:"BrtEndSmartTags" },
0x0254: { n:"BrtSmartTagType" },
0x0255: { n:"BrtBeginSmartTagTypes" },
0x0256: { n:"BrtEndSmartTagTypes" },
0x0257: { n:"BrtBeginSXFilters" },
0x0258: { n:"BrtEndSXFilters" },
0x0259: { n:"BrtBeginSXFILTER" },
0x025A: { n:"BrtEndSXFilter" },
0x025B: { n:"BrtBeginFills" },
0x025C: { n:"BrtEndFills" },
0x025D: { n:"BrtBeginCellWatches" },
0x025E: { n:"BrtEndCellWatches" },
0x025F: { n:"BrtCellWatch" },
0x0260: { n:"BrtBeginCRErrs" },
0x0261: { n:"BrtEndCRErrs" },
0x0262: { n:"BrtCrashRecErr" },
0x0263: { n:"BrtBeginFonts" },
0x0264: { n:"BrtEndFonts" },
0x0265: { n:"BrtBeginBorders" },
0x0266: { n:"BrtEndBorders" },
0x0267: { n:"BrtBeginFmts" },
0x0268: { n:"BrtEndFmts" },
0x0269: { n:"BrtBeginCellXFs" },
0x026A: { n:"BrtEndCellXFs" },
0x026B: { n:"BrtBeginStyles" },
0x026C: { n:"BrtEndStyles" },
0x0271: { n:"BrtBigName" },
0x0272: { n:"BrtBeginCellStyleXFs" },
0x0273: { n:"BrtEndCellStyleXFs" },
0x0274: { n:"BrtBeginComments" },
0x0275: { n:"BrtEndComments" },
0x0276: { n:"BrtBeginCommentAuthors" },
0x0277: { n:"BrtEndCommentAuthors" },
0x0278: { n:"BrtCommentAuthor", f:parse_BrtCommentAuthor },
0x0279: { n:"BrtBeginCommentList" },
0x027A: { n:"BrtEndCommentList" },
0x027B: { n:"BrtBeginComment", f:parse_BrtBeginComment},
0x027C: { n:"BrtEndComment" },
0x027D: { n:"BrtCommentText", f:parse_BrtCommentText },
0x027E: { n:"BrtBeginOleObjects" },
0x027F: { n:"BrtOleObject" },
0x0280: { n:"BrtEndOleObjects" },
0x0281: { n:"BrtBeginSxrules" },
0x0282: { n:"BrtEndSxRules" },
0x0283: { n:"BrtBeginActiveXControls" },
0x0284: { n:"BrtActiveX" },
0x0285: { n:"BrtEndActiveXControls" },
0x0286: { n:"BrtBeginPCDSDTCEMembersSortBy" },
0x0288: { n:"BrtBeginCellIgnoreECs" },
0x0289: { n:"BrtCellIgnoreEC" },
0x028A: { n:"BrtEndCellIgnoreECs" },
0x028B: { n:"BrtCsProp", f:parse_BrtCsProp },
0x028C: { n:"BrtCsPageSetup" },
0x028D: { n:"BrtBeginUserCsViews" },
0x028E: { n:"BrtEndUserCsViews" },
0x028F: { n:"BrtBeginUserCsView" },
0x0290: { n:"BrtEndUserCsView" },
0x0291: { n:"BrtBeginPcdSFCIEntries" },
0x0292: { n:"BrtEndPCDSFCIEntries" },
0x0293: { n:"BrtPCDSFCIEntry" },
0x0294: { n:"BrtBeginListParts" },
0x0295: { n:"BrtListPart" },
0x0296: { n:"BrtEndListParts" },
0x0297: { n:"BrtSheetCalcProp" },
0x0298: { n:"BrtBeginFnGroup" },
0x0299: { n:"BrtFnGroup" },
0x029A: { n:"BrtEndFnGroup" },
0x029B: { n:"BrtSupAddin" },
0x029C: { n:"BrtSXTDMPOrder" },
0x029D: { n:"BrtCsProtection" },
0x029F: { n:"BrtBeginWsSortMap" },
0x02A0: { n:"BrtEndWsSortMap" },
0x02A1: { n:"BrtBeginRRSort" },
0x02A2: { n:"BrtEndRRSort" },
0x02A3: { n:"BrtRRSortItem" },
0x02A4: { n:"BrtFileSharingIso" },
0x02A5: { n:"BrtBookProtectionIso" },
0x02A6: { n:"BrtSheetProtectionIso" },
0x02A7: { n:"BrtCsProtectionIso" },
0x02A8: { n:"BrtRangeProtectionIso" },
0x0400: { n:"BrtRwDescent" },
0x0401: { n:"BrtKnownFonts" },
0x0402: { n:"BrtBeginSXTupleSet" },
0x0403: { n:"BrtEndSXTupleSet" },
0x0404: { n:"BrtBeginSXTupleSetHeader" },
0x0405: { n:"BrtEndSXTupleSetHeader" },
0x0406: { n:"BrtSXTupleSetHeaderItem" },
0x0407: { n:"BrtBeginSXTupleSetData" },
0x0408: { n:"BrtEndSXTupleSetData" },
0x0409: { n:"BrtBeginSXTupleSetRow" },
0x040A: { n:"BrtEndSXTupleSetRow" },
0x040B: { n:"BrtSXTupleSetRowItem" },
0x040C: { n:"BrtNameExt" },
0x040D: { n:"BrtPCDH14" },
0x040E: { n:"BrtBeginPCDCalcMem14" },
0x040F: { n:"BrtEndPCDCalcMem14" },
0x0410: { n:"BrtSXTH14" },
0x0411: { n:"BrtBeginSparklineGroup" },
0x0412: { n:"BrtEndSparklineGroup" },
0x0413: { n:"BrtSparkline" },
0x0414: { n:"BrtSXDI14" },
0x0415: { n:"BrtWsFmtInfoEx14" },
0x0416: { n:"BrtBeginConditionalFormatting14" },
0x0417: { n:"BrtEndConditionalFormatting14" },
0x0418: { n:"BrtBeginCFRule14" },
0x0419: { n:"BrtEndCFRule14" },
0x041A: { n:"BrtCFVO14" },
0x041B: { n:"BrtBeginDatabar14" },
0x041C: { n:"BrtBeginIconSet14" },
0x041D: { n:"BrtDVal14" },
0x041E: { n:"BrtBeginDVals14" },
0x041F: { n:"BrtColor14" },
0x0420: { n:"BrtBeginSparklines" },
0x0421: { n:"BrtEndSparklines" },
0x0422: { n:"BrtBeginSparklineGroups" },
0x0423: { n:"BrtEndSparklineGroups" },
0x0425: { n:"BrtSXVD14" },
0x0426: { n:"BrtBeginSXView14" },
0x0427: { n:"BrtEndSXView14" },
0x0428: { n:"BrtBeginSXView16" },
0x0429: { n:"BrtEndSXView16" },
0x042A: { n:"BrtBeginPCD14" },
0x042B: { n:"BrtEndPCD14" },
0x042C: { n:"BrtBeginExtConn14" },
0x042D: { n:"BrtEndExtConn14" },
0x042E: { n:"BrtBeginSlicerCacheIDs" },
0x042F: { n:"BrtEndSlicerCacheIDs" },
0x0430: { n:"BrtBeginSlicerCacheID" },
0x0431: { n:"BrtEndSlicerCacheID" },
0x0433: { n:"BrtBeginSlicerCache" },
0x0434: { n:"BrtEndSlicerCache" },
0x0435: { n:"BrtBeginSlicerCacheDef" },
0x0436: { n:"BrtEndSlicerCacheDef" },
0x0437: { n:"BrtBeginSlicersEx" },
0x0438: { n:"BrtEndSlicersEx" },
0x0439: { n:"BrtBeginSlicerEx" },
0x043A: { n:"BrtEndSlicerEx" },
0x043B: { n:"BrtBeginSlicer" },
0x043C: { n:"BrtEndSlicer" },
0x043D: { n:"BrtSlicerCachePivotTables" },
0x043E: { n:"BrtBeginSlicerCacheOlapImpl" },
0x043F: { n:"BrtEndSlicerCacheOlapImpl" },
0x0440: { n:"BrtBeginSlicerCacheLevelsData" },
0x0441: { n:"BrtEndSlicerCacheLevelsData" },
0x0442: { n:"BrtBeginSlicerCacheLevelData" },
0x0443: { n:"BrtEndSlicerCacheLevelData" },
0x0444: { n:"BrtBeginSlicerCacheSiRanges" },
0x0445: { n:"BrtEndSlicerCacheSiRanges" },
0x0446: { n:"BrtBeginSlicerCacheSiRange" },
0x0447: { n:"BrtEndSlicerCacheSiRange" },
0x0448: { n:"BrtSlicerCacheOlapItem" },
0x0449: { n:"BrtBeginSlicerCacheSelections" },
0x044A: { n:"BrtSlicerCacheSelection" },
0x044B: { n:"BrtEndSlicerCacheSelections" },
0x044C: { n:"BrtBeginSlicerCacheNative" },
0x044D: { n:"BrtEndSlicerCacheNative" },
0x044E: { n:"BrtSlicerCacheNativeItem" },
0x044F: { n:"BrtRangeProtection14" },
0x0450: { n:"BrtRangeProtectionIso14" },
0x0451: { n:"BrtCellIgnoreEC14" },
0x0457: { n:"BrtList14" },
0x0458: { n:"BrtCFIcon" },
0x0459: { n:"BrtBeginSlicerCachesPivotCacheIDs" },
0x045A: { n:"BrtEndSlicerCachesPivotCacheIDs" },
0x045B: { n:"BrtBeginSlicers" },
0x045C: { n:"BrtEndSlicers" },
0x045D: { n:"BrtWbProp14" },
0x045E: { n:"BrtBeginSXEdit" },
0x045F: { n:"BrtEndSXEdit" },
0x0460: { n:"BrtBeginSXEdits" },
0x0461: { n:"BrtEndSXEdits" },
0x0462: { n:"BrtBeginSXChange" },
0x0463: { n:"BrtEndSXChange" },
0x0464: { n:"BrtBeginSXChanges" },
0x0465: { n:"BrtEndSXChanges" },
0x0466: { n:"BrtSXTupleItems" },
0x0468: { n:"BrtBeginSlicerStyle" },
0x0469: { n:"BrtEndSlicerStyle" },
0x046A: { n:"BrtSlicerStyleElement" },
0x046B: { n:"BrtBeginStyleSheetExt14" },
0x046C: { n:"BrtEndStyleSheetExt14" },
0x046D: { n:"BrtBeginSlicerCachesPivotCacheID" },
0x046E: { n:"BrtEndSlicerCachesPivotCacheID" },
0x046F: { n:"BrtBeginConditionalFormattings" },
0x0470: { n:"BrtEndConditionalFormattings" },
0x0471: { n:"BrtBeginPCDCalcMemExt" },
0x0472: { n:"BrtEndPCDCalcMemExt" },
0x0473: { n:"BrtBeginPCDCalcMemsExt" },
0x0474: { n:"BrtEndPCDCalcMemsExt" },
0x0475: { n:"BrtPCDField14" },
0x0476: { n:"BrtBeginSlicerStyles" },
0x0477: { n:"BrtEndSlicerStyles" },
0x0478: { n:"BrtBeginSlicerStyleElements" },
0x0479: { n:"BrtEndSlicerStyleElements" },
0x047A: { n:"BrtCFRuleExt" },
0x047B: { n:"BrtBeginSXCondFmt14" },
0x047C: { n:"BrtEndSXCondFmt14" },
0x047D: { n:"BrtBeginSXCondFmts14" },
0x047E: { n:"BrtEndSXCondFmts14" },
0x0480: { n:"BrtBeginSortCond14" },
0x0481: { n:"BrtEndSortCond14" },
0x0482: { n:"BrtEndDVals14" },
0x0483: { n:"BrtEndIconSet14" },
0x0484: { n:"BrtEndDatabar14" },
0x0485: { n:"BrtBeginColorScale14" },
0x0486: { n:"BrtEndColorScale14" },
0x0487: { n:"BrtBeginSxrules14" },
0x0488: { n:"BrtEndSxrules14" },
0x0489: { n:"BrtBeginPRule14" },
0x048A: { n:"BrtEndPRule14" },
0x048B: { n:"BrtBeginPRFilters14" },
0x048C: { n:"BrtEndPRFilters14" },
0x048D: { n:"BrtBeginPRFilter14" },
0x048E: { n:"BrtEndPRFilter14" },
0x048F: { n:"BrtBeginPRFItem14" },
0x0490: { n:"BrtEndPRFItem14" },
0x0491: { n:"BrtBeginCellIgnoreECs14" },
0x0492: { n:"BrtEndCellIgnoreECs14" },
0x0493: { n:"BrtDxf14" },
0x0494: { n:"BrtBeginDxF14s" },
0x0495: { n:"BrtEndDxf14s" },
0x0499: { n:"BrtFilter14" },
0x049A: { n:"BrtBeginCustomFilters14" },
0x049C: { n:"BrtCustomFilter14" },
0x049D: { n:"BrtIconFilter14" },
0x049E: { n:"BrtPivotCacheConnectionName" },
0x0800: { n:"BrtBeginDecoupledPivotCacheIDs" },
0x0801: { n:"BrtEndDecoupledPivotCacheIDs" },
0x0802: { n:"BrtDecoupledPivotCacheID" },
0x0803: { n:"BrtBeginPivotTableRefs" },
0x0804: { n:"BrtEndPivotTableRefs" },
0x0805: { n:"BrtPivotTableRef" },
0x0806: { n:"BrtSlicerCacheBookPivotTables" },
0x0807: { n:"BrtBeginSxvcells" },
0x0808: { n:"BrtEndSxvcells" },
0x0809: { n:"BrtBeginSxRow" },
0x080A: { n:"BrtEndSxRow" },
0x080C: { n:"BrtPcdCalcMem15" },
0x0813: { n:"BrtQsi15" },
0x0814: { n:"BrtBeginWebExtensions" },
0x0815: { n:"BrtEndWebExtensions" },
0x0816: { n:"BrtWebExtension" },
0x0817: { n:"BrtAbsPath15" },
0x0818: { n:"BrtBeginPivotTableUISettings" },
0x0819: { n:"BrtEndPivotTableUISettings" },
0x081B: { n:"BrtTableSlicerCacheIDs" },
0x081C: { n:"BrtTableSlicerCacheID" },
0x081D: { n:"BrtBeginTableSlicerCache" },
0x081E: { n:"BrtEndTableSlicerCache" },
0x081F: { n:"BrtSxFilter15" },
0x0820: { n:"BrtBeginTimelineCachePivotCacheIDs" },
0x0821: { n:"BrtEndTimelineCachePivotCacheIDs" },
0x0822: { n:"BrtTimelineCachePivotCacheID" },
0x0823: { n:"BrtBeginTimelineCacheIDs" },
0x0824: { n:"BrtEndTimelineCacheIDs" },
0x0825: { n:"BrtBeginTimelineCacheID" },
0x0826: { n:"BrtEndTimelineCacheID" },
0x0827: { n:"BrtBeginTimelinesEx" },
0x0828: { n:"BrtEndTimelinesEx" },
0x0829: { n:"BrtBeginTimelineEx" },
0x082A: { n:"BrtEndTimelineEx" },
0x082B: { n:"BrtWorkBookPr15" },
0x082C: { n:"BrtPCDH15" },
0x082D: { n:"BrtBeginTimelineStyle" },
0x082E: { n:"BrtEndTimelineStyle" },
0x082F: { n:"BrtTimelineStyleElement" },
0x0830: { n:"BrtBeginTimelineStylesheetExt15" },
0x0831: { n:"BrtEndTimelineStylesheetExt15" },
0x0832: { n:"BrtBeginTimelineStyles" },
0x0833: { n:"BrtEndTimelineStyles" },
0x0834: { n:"BrtBeginTimelineStyleElements" },
0x0835: { n:"BrtEndTimelineStyleElements" },
0x0836: { n:"BrtDxf15" },
0x0837: { n:"BrtBeginDxfs15" },
0x0838: { n:"brtEndDxfs15" },
0x0839: { n:"BrtSlicerCacheHideItemsWithNoData" },
0x083A: { n:"BrtBeginItemUniqueNames" },
0x083B: { n:"BrtEndItemUniqueNames" },
0x083C: { n:"BrtItemUniqueName" },
0x083D: { n:"BrtBeginExtConn15" },
0x083E: { n:"BrtEndExtConn15" },
0x083F: { n:"BrtBeginOledbPr15" },
0x0840: { n:"BrtEndOledbPr15" },
0x0841: { n:"BrtBeginDataFeedPr15" },
0x0842: { n:"BrtEndDataFeedPr15" },
0x0843: { n:"BrtTextPr15" },
0x0844: { n:"BrtRangePr15" },
0x0845: { n:"BrtDbCommand15" },
0x0846: { n:"BrtBeginDbTables15" },
0x0847: { n:"BrtEndDbTables15" },
0x0848: { n:"BrtDbTable15" },
0x0849: { n:"BrtBeginDataModel" },
0x084A: { n:"BrtEndDataModel" },
0x084B: { n:"BrtBeginModelTables" },
0x084C: { n:"BrtEndModelTables" },
0x084D: { n:"BrtModelTable" },
0x084E: { n:"BrtBeginModelRelationships" },
0x084F: { n:"BrtEndModelRelationships" },
0x0850: { n:"BrtModelRelationship" },
0x0851: { n:"BrtBeginECTxtWiz15" },
0x0852: { n:"BrtEndECTxtWiz15" },
0x0853: { n:"BrtBeginECTWFldInfoLst15" },
0x0854: { n:"BrtEndECTWFldInfoLst15" },
0x0855: { n:"BrtBeginECTWFldInfo15" },
0x0856: { n:"BrtFieldListActiveItem" },
0x0857: { n:"BrtPivotCacheIdVersion" },
0x0858: { n:"BrtSXDI15" },
0x0859: { n:"BrtBeginModelTimeGroupings" },
0x085A: { n:"BrtEndModelTimeGroupings" },
0x085B: { n:"BrtBeginModelTimeGrouping" },
0x085C: { n:"BrtEndModelTimeGrouping" },
0x085D: { n:"BrtModelTimeGroupingCalcCol" },
0x0C00: { n:"BrtUid" },
0x0C01: { n:"BrtRevisionPtr" },
0x13e7: { n:"BrtBeginCalcFeatures" },
0x13e8: { n:"BrtEndCalcFeatures" },
0x13e9: { n:"BrtCalcFeature" },
0xFFFF: { n:"" }
};

var XLSBRE = evert_key(XLSBRecordEnum, 'n');

/* [MS-XLS] 2.3 Record Enumeration */
var XLSRecordEnum = {
0x0003: { n:"BIFF2NUM", f:parse_BIFF2NUM },
0x0004: { n:"BIFF2STR", f:parse_BIFF2STR },
0x0006: { n:"Formula", f:parse_Formula },
0x0009: { n:'BOF', f:parse_BOF },
0x000a: { n:'EOF', f:parsenoop2 },
0x000c: { n:"CalcCount", f:parseuint16 },
0x000d: { n:"CalcMode", f:parseuint16 },
0x000e: { n:"CalcPrecision", f:parsebool },
0x000f: { n:"CalcRefMode", f:parsebool },
0x0010: { n:"CalcDelta", f:parse_Xnum },
0x0011: { n:"CalcIter", f:parsebool },
0x0012: { n:"Protect", f:parsebool },
0x0013: { n:"Password", f:parseuint16 },
0x0014: { n:"Header", f:parse_XLHeaderFooter },
0x0015: { n:"Footer", f:parse_XLHeaderFooter },
0x0017: { n:"ExternSheet", f:parse_ExternSheet },
0x0018: { n:"Lbl", f:parse_Lbl },
0x0019: { n:"WinProtect", f:parsebool },
0x001a: { n:"VerticalPageBreaks" },
0x001b: { n:"HorizontalPageBreaks" },
0x001c: { n:"Note", f:parse_Note },
0x001d: { n:"Selection" },
0x0022: { n:"Date1904", f:parsebool },
0x0023: { n:"ExternName", f:parse_ExternName },
0x0026: { n:"LeftMargin", f:parse_Xnum },
0x0027: { n:"RightMargin", f:parse_Xnum },
0x0028: { n:"TopMargin", f:parse_Xnum },
0x0029: { n:"BottomMargin", f:parse_Xnum },
0x002a: { n:"PrintRowCol", f:parsebool },
0x002b: { n:"PrintGrid", f:parsebool },
0x002f: { n:"FilePass", f:parse_FilePass },
0x0031: { n:"Font", f:parse_Font },
0x0033: { n:"PrintSize", f:parseuint16 },
0x003c: { n:"Continue" },
0x003d: { n:"Window1", f:parse_Window1 },
0x0040: { n:"Backup", f:parsebool },
0x0041: { n:"Pane" },
0x0042: { n:'CodePage', f:parseuint16 },
0x004d: { n:"Pls" },
0x0050: { n:"DCon" },
0x0051: { n:"DConRef" },
0x0052: { n:"DConName" },
0x0055: { n:"DefColWidth", f:parseuint16 },
0x0059: { n:"XCT" },
0x005a: { n:"CRN" },
0x005b: { n:"FileSharing" },
0x005c: { n:'WriteAccess', f:parse_WriteAccess },
0x005d: { n:"Obj", f:parse_Obj },
0x005e: { n:"Uncalced" },
0x005f: { n:"CalcSaveRecalc", f:parsebool },
0x0060: { n:"Template" },
0x0061: { n:"Intl" },
0x0063: { n:"ObjProtect", f:parsebool },
0x007d: { n:"ColInfo", f:parse_ColInfo },
0x0080: { n:"Guts", f:parse_Guts },
0x0081: { n:"WsBool", f:parse_WsBool },
0x0082: { n:"GridSet", f:parseuint16 },
0x0083: { n:"HCenter", f:parsebool },
0x0084: { n:"VCenter", f:parsebool },
0x0085: { n:'BoundSheet8', f:parse_BoundSheet8 },
0x0086: { n:"WriteProtect" },
0x008c: { n:"Country", f:parse_Country },
0x008d: { n:"HideObj", f:parseuint16 },
0x0090: { n:"Sort" },
0x0092: { n:"Palette", f:parse_Palette },
0x0097: { n:"Sync" },
0x0098: { n:"LPr" },
0x0099: { n:"DxGCol" },
0x009a: { n:"FnGroupName" },
0x009b: { n:"FilterMode" },
0x009c: { n:"BuiltInFnGroupCount", f:parseuint16 },
0x009d: { n:"AutoFilterInfo" },
0x009e: { n:"AutoFilter" },
0x00a0: { n:"Scl", f:parse_Scl },
0x00a1: { n:"Setup", f:parse_Setup },
0x00ae: { n:"ScenMan" },
0x00af: { n:"SCENARIO" },
0x00b0: { n:"SxView" },
0x00b1: { n:"Sxvd" },
0x00b2: { n:"SXVI" },
0x00b4: { n:"SxIvd" },
0x00b5: { n:"SXLI" },
0x00b6: { n:"SXPI" },
0x00b8: { n:"DocRoute" },
0x00b9: { n:"RecipName" },
0x00bd: { n:"MulRk", f:parse_MulRk },
0x00be: { n:"MulBlank", f:parse_MulBlank },
0x00c1: { n:'Mms', f:parsenoop2 },
0x00c5: { n:"SXDI" },
0x00c6: { n:"SXDB" },
0x00c7: { n:"SXFDB" },
0x00c8: { n:"SXDBB" },
0x00c9: { n:"SXNum" },
0x00ca: { n:"SxBool", f:parsebool },
0x00cb: { n:"SxErr" },
0x00cc: { n:"SXInt" },
0x00cd: { n:"SXString" },
0x00ce: { n:"SXDtr" },
0x00cf: { n:"SxNil" },
0x00d0: { n:"SXTbl" },
0x00d1: { n:"SXTBRGIITM" },
0x00d2: { n:"SxTbpg" },
0x00d3: { n:"ObProj" },
0x00d5: { n:"SXStreamID" },
0x00d7: { n:"DBCell" },
0x00d8: { n:"SXRng" },
0x00d9: { n:"SxIsxoper" },
0x00da: { n:"BookBool", f:parseuint16 },
0x00dc: { n:"DbOrParamQry" },
0x00dd: { n:"ScenarioProtect", f:parsebool },
0x00de: { n:"OleObjectSize" },
0x00e0: { n:"XF", f:parse_XF },
0x00e1: { n:'InterfaceHdr', f:parse_InterfaceHdr },
0x00e2: { n:'InterfaceEnd', f:parsenoop2 },
0x00e3: { n:"SXVS" },
0x00e5: { n:"MergeCells", f:parse_MergeCells },
0x00e9: { n:"BkHim" },
0x00eb: { n:"MsoDrawingGroup" },
0x00ec: { n:"MsoDrawing" },
0x00ed: { n:"MsoDrawingSelection" },
0x00ef: { n:"PhoneticInfo" },
0x00f0: { n:"SxRule" },
0x00f1: { n:"SXEx" },
0x00f2: { n:"SxFilt" },
0x00f4: { n:"SxDXF" },
0x00f5: { n:"SxItm" },
0x00f6: { n:"SxName" },
0x00f7: { n:"SxSelect" },
0x00f8: { n:"SXPair" },
0x00f9: { n:"SxFmla" },
0x00fb: { n:"SxFormat" },
0x00fc: { n:"SST", f:parse_SST },
0x00fd: { n:"LabelSst", f:parse_LabelSst },
0x00ff: { n:"ExtSST", f:parse_ExtSST },
0x0100: { n:"SXVDEx" },
0x0103: { n:"SXFormula" },
0x0122: { n:"SXDBEx" },
0x0137: { n:"RRDInsDel" },
0x0138: { n:"RRDHead" },
0x013b: { n:"RRDChgCell" },
0x013d: { n:"RRTabId", f:parseuint16a },
0x013e: { n:"RRDRenSheet" },
0x013f: { n:"RRSort" },
0x0140: { n:"RRDMove" },
0x014a: { n:"RRFormat" },
0x014b: { n:"RRAutoFmt" },
0x014d: { n:"RRInsertSh" },
0x014e: { n:"RRDMoveBegin" },
0x014f: { n:"RRDMoveEnd" },
0x0150: { n:"RRDInsDelBegin" },
0x0151: { n:"RRDInsDelEnd" },
0x0152: { n:"RRDConflict" },
0x0153: { n:"RRDDefName" },
0x0154: { n:"RRDRstEtxp" },
0x015f: { n:"LRng" },
0x0160: { n:"UsesELFs", f:parsebool },
0x0161: { n:"DSF", f:parsenoop2 },
0x0191: { n:"CUsr" },
0x0192: { n:"CbUsr" },
0x0193: { n:"UsrInfo" },
0x0194: { n:"UsrExcl" },
0x0195: { n:"FileLock" },
0x0196: { n:"RRDInfo" },
0x0197: { n:"BCUsrs" },
0x0198: { n:"UsrChk" },
0x01a9: { n:"UserBView" },
0x01aa: { n:"UserSViewBegin" },
0x01ab: { n:"UserSViewEnd" },
0x01ac: { n:"RRDUserView" },
0x01ad: { n:"Qsi" },
0x01ae: { n:"SupBook", f:parse_SupBook },
0x01af: { n:"Prot4Rev", f:parsebool },
0x01b0: { n:"CondFmt" },
0x01b1: { n:"CF" },
0x01b2: { n:"DVal" },
0x01b5: { n:"DConBin" },
0x01b6: { n:"TxO", f:parse_TxO },
0x01b7: { n:"RefreshAll", f:parsebool },
0x01b8: { n:"HLink", f:parse_HLink },
0x01b9: { n:"Lel" },
0x01ba: { n:"CodeName", f:parse_XLUnicodeString },
0x01bb: { n:"SXFDBType" },
0x01bc: { n:"Prot4RevPass", f:parseuint16 },
0x01bd: { n:"ObNoMacros" },
0x01be: { n:"Dv" },
0x01c0: { n:"Excel9File", f:parsenoop2 },
0x01c1: { n:"RecalcId", f:parse_RecalcId, r:2},
0x01c2: { n:"EntExU2", f:parsenoop2 },
0x0200: { n:"Dimensions", f:parse_Dimensions },
0x0201: { n:"Blank", f:parse_Blank },
0x0203: { n:"Number", f:parse_Number },
0x0204: { n:"Label", f:parse_Label },
0x0205: { n:"BoolErr", f:parse_BoolErr },
0x0206: { n:"Formula", f:parse_Formula },
0x0207: { n:"String", f:parse_String },
0x0208: { n:'Row', f:parse_Row },
0x020b: { n:"Index" },
0x0221: { n:"Array", f:parse_Array },
0x0225: { n:"DefaultRowHeight", f:parse_DefaultRowHeight },
0x0236: { n:"Table" },
0x023e: { n:"Window2", f:parse_Window2 },
0x027e: { n:"RK", f:parse_RK },
0x0293: { n:"Style" },
0x0406: { n:"Formula", f:parse_Formula },
0x0418: { n:"BigName" },
0x041e: { n:"Format", f:parse_Format },
0x043c: { n:"ContinueBigName" },
0x04bc: { n:"ShrFmla", f:parse_ShrFmla },
0x0800: { n:"HLinkTooltip", f:parse_HLinkTooltip },
0x0801: { n:"WebPub" },
0x0802: { n:"QsiSXTag" },
0x0803: { n:"DBQueryExt" },
0x0804: { n:"ExtString" },
0x0805: { n:"TxtQry" },
0x0806: { n:"Qsir" },
0x0807: { n:"Qsif" },
0x0808: { n:"RRDTQSIF" },
0x0809: { n:'BOF', f:parse_BOF },
0x080a: { n:"OleDbConn" },
0x080b: { n:"WOpt" },
0x080c: { n:"SXViewEx" },
0x080d: { n:"SXTH" },
0x080e: { n:"SXPIEx" },
0x080f: { n:"SXVDTEx" },
0x0810: { n:"SXViewEx9" },
0x0812: { n:"ContinueFrt" },
0x0813: { n:"RealTimeData" },
0x0850: { n:"ChartFrtInfo" },
0x0851: { n:"FrtWrapper" },
0x0852: { n:"StartBlock" },
0x0853: { n:"EndBlock" },
0x0854: { n:"StartObject" },
0x0855: { n:"EndObject" },
0x0856: { n:"CatLab" },
0x0857: { n:"YMult" },
0x0858: { n:"SXViewLink" },
0x0859: { n:"PivotChartBits" },
0x085a: { n:"FrtFontList" },
0x0862: { n:"SheetExt" },
0x0863: { n:"BookExt", r:12},
0x0864: { n:"SXAddl" },
0x0865: { n:"CrErr" },
0x0866: { n:"HFPicture" },
0x0867: { n:'FeatHdr', f:parsenoop2 },
0x0868: { n:"Feat" },
0x086a: { n:"DataLabExt" },
0x086b: { n:"DataLabExtContents" },
0x086c: { n:"CellWatch" },
0x0871: { n:"FeatHdr11" },
0x0872: { n:"Feature11" },
0x0874: { n:"DropDownObjIds" },
0x0875: { n:"ContinueFrt11" },
0x0876: { n:"DConn" },
0x0877: { n:"List12" },
0x0878: { n:"Feature12" },
0x0879: { n:"CondFmt12" },
0x087a: { n:"CF12" },
0x087b: { n:"CFEx" },
0x087c: { n:"XFCRC", f:parse_XFCRC, r:12 },
0x087d: { n:"XFExt", f:parse_XFExt, r:12 },
0x087e: { n:"AutoFilter12" },
0x087f: { n:"ContinueFrt12" },
0x0884: { n:"MDTInfo" },
0x0885: { n:"MDXStr" },
0x0886: { n:"MDXTuple" },
0x0887: { n:"MDXSet" },
0x0888: { n:"MDXProp" },
0x0889: { n:"MDXKPI" },
0x088a: { n:"MDB" },
0x088b: { n:"PLV" },
0x088c: { n:"Compat12", f:parsebool, r:12 },
0x088d: { n:"DXF" },
0x088e: { n:"TableStyles", r:12 },
0x088f: { n:"TableStyle" },
0x0890: { n:"TableStyleElement" },
0x0892: { n:"StyleExt" },
0x0893: { n:"NamePublish" },
0x0894: { n:"NameCmt", f:parse_NameCmt, r:12 },
0x0895: { n:"SortData" },
0x0896: { n:"Theme", f:parse_Theme, r:12 },
0x0897: { n:"GUIDTypeLib" },
0x0898: { n:"FnGrp12" },
0x0899: { n:"NameFnGrp12" },
0x089a: { n:"MTRSettings", f:parse_MTRSettings, r:12 },
0x089b: { n:"CompressPictures", f:parsenoop2 },
0x089c: { n:"HeaderFooter" },
0x089d: { n:"CrtLayout12" },
0x089e: { n:"CrtMlFrt" },
0x089f: { n:"CrtMlFrtContinue" },
0x08a3: { n:"ForceFullCalculation", f:parse_ForceFullCalculation },
0x08a4: { n:"ShapePropsStream" },
0x08a5: { n:"TextPropsStream" },
0x08a6: { n:"RichTextStream" },
0x08a7: { n:"CrtLayout12A" },
0x1001: { n:"Units" },
0x1002: { n:"Chart" },
0x1003: { n:"Series" },
0x1006: { n:"DataFormat" },
0x1007: { n:"LineFormat" },
0x1009: { n:"MarkerFormat" },
0x100a: { n:"AreaFormat" },
0x100b: { n:"PieFormat" },
0x100c: { n:"AttachedLabel" },
0x100d: { n:"SeriesText" },
0x1014: { n:"ChartFormat" },
0x1015: { n:"Legend" },
0x1016: { n:"SeriesList" },
0x1017: { n:"Bar" },
0x1018: { n:"Line" },
0x1019: { n:"Pie" },
0x101a: { n:"Area" },
0x101b: { n:"Scatter" },
0x101c: { n:"CrtLine" },
0x101d: { n:"Axis" },
0x101e: { n:"Tick" },
0x101f: { n:"ValueRange" },
0x1020: { n:"CatSerRange" },
0x1021: { n:"AxisLine" },
0x1022: { n:"CrtLink" },
0x1024: { n:"DefaultText" },
0x1025: { n:"Text" },
0x1026: { n:"FontX", f:parseuint16 },
0x1027: { n:"ObjectLink" },
0x1032: { n:"Frame" },
0x1033: { n:"Begin" },
0x1034: { n:"End" },
0x1035: { n:"PlotArea" },
0x103a: { n:"Chart3d" },
0x103c: { n:"PicF" },
0x103d: { n:"DropBar" },
0x103e: { n:"Radar" },
0x103f: { n:"Surf" },
0x1040: { n:"RadarArea" },
0x1041: { n:"AxisParent" },
0x1043: { n:"LegendException" },
0x1044: { n:"ShtProps", f:parse_ShtProps },
0x1045: { n:"SerToCrt" },
0x1046: { n:"AxesUsed" },
0x1048: { n:"SBaseRef" },
0x104a: { n:"SerParent" },
0x104b: { n:"SerAuxTrend" },
0x104e: { n:"IFmtRecord" },
0x104f: { n:"Pos" },
0x1050: { n:"AlRuns" },
0x1051: { n:"BRAI" },
0x105b: { n:"SerAuxErrBar" },
0x105c: { n:"ClrtClient", f:parse_ClrtClient },
0x105d: { n:"SerFmt" },
0x105f: { n:"Chart3DBarShape" },
0x1060: { n:"Fbi" },
0x1061: { n:"BopPop" },
0x1062: { n:"AxcExt" },
0x1063: { n:"Dat" },
0x1064: { n:"PlotGrowth" },
0x1065: { n:"SIIndex" },
0x1066: { n:"GelFrame" },
0x1067: { n:"BopPopCustom" },
0x1068: { n:"Fbi2" },

0x0000: { n:"Dimensions", f:parse_Dimensions },
0x0002: { n:"BIFF2INT", f:parse_BIFF2INT },
0x0005: { n:"BoolErr", f:parse_BoolErr },
0x0007: { n:"String", f:parse_BIFF2STRING },
0x0008: { n:"BIFF2ROW" },
0x000b: { n:"Index" },
0x0016: { n:"ExternCount", f:parseuint16 },
0x001e: { n:"BIFF2FORMAT", f:parse_BIFF2Format },
0x001f: { n:"BIFF2FMTCNT" }, /* 16-bit cnt of BIFF2FORMAT records */
0x0020: { n:"BIFF2COLINFO" },
0x0021: { n:"Array", f:parse_Array },
0x0025: { n:"DefaultRowHeight", f:parse_DefaultRowHeight },
0x0032: { n:"BIFF2FONTXTRA", f:parse_BIFF2FONTXTRA },
0x0034: { n:"DDEObjName" },
0x003e: { n:"BIFF2WINDOW2" },
0x0043: { n:"BIFF2XF" },
0x0045: { n:"BIFF2FONTCLR" },
0x0056: { n:"BIFF4FMTCNT" }, /* 16-bit cnt, similar to BIFF2 */
0x007e: { n:"RK" }, /* Not necessarily same as 0x027e */
0x007f: { n:"ImData", f:parse_ImData },
0x0087: { n:"Addin" },
0x0088: { n:"Edg" },
0x0089: { n:"Pub" },
0x0091: { n:"Sub" },
0x0094: { n:"LHRecord" },
0x0095: { n:"LHNGraph" },
0x0096: { n:"Sound" },
0x00a9: { n:"CoordList" },
0x00ab: { n:"GCW" },
0x00bc: { n:"ShrFmla" }, /* Not necessarily same as 0x04bc */
0x00bf: { n:"ToolbarHdr" },
0x00c0: { n:"ToolbarEnd" },
0x00c2: { n:"AddMenu" },
0x00c3: { n:"DelMenu" },
0x00d6: { n:"RString", f:parse_RString },
0x00df: { n:"UDDesc" },
0x00ea: { n:"TabIdConf" },
0x0162: { n:"XL5Modify" },
0x01a5: { n:"FileSharing2" },
0x0209: { n:'BOF', f:parse_BOF },
0x0218: { n:"Lbl", f:parse_Lbl },
0x0223: { n:"ExternName", f:parse_ExternName },
0x0231: { n:"Font" },
0x0243: { n:"BIFF3XF" },
0x0409: { n:'BOF', f:parse_BOF },
0x0443: { n:"BIFF4XF" },
0x086d: { n:"FeatInfo" },
0x0873: { n:"FeatInfo11" },
0x0881: { n:"SXAddl12" },
0x08c0: { n:"AutoWebPub" },
0x08c1: { n:"ListObj" },
0x08c2: { n:"ListField" },
0x08c3: { n:"ListDV" },
0x08c4: { n:"ListCondFmt" },
0x08c5: { n:"ListCF" },
0x08c6: { n:"FMQry" },
0x08c7: { n:"FMSQry" },
0x08c8: { n:"PLV" },
0x08c9: { n:"LnExt" },
0x08ca: { n:"MkrExt" },
0x08cb: { n:"CrtCoopt" },
0x08d6: { n:"FRTArchId$", r:12 },

0x7262: {}
};

var XLSRE = evert_key(XLSRecordEnum, 'n');
function write_biff_rec(ba, type, payload, length) {
	var t = +type || +XLSRE[type];
	if(isNaN(t)) return;
	var len = length || (payload||[]).length || 0;
	var o = ba.next(4);
	o.write_shift(2, t);
	o.write_shift(2, len);
	if(len > 0 && is_buf(payload)) ba.push(payload);
}

function write_BIFF2Cell(out, r, c) {
	if(!out) out = new_buf(7);
	out.write_shift(2, r);
	out.write_shift(2, c);
	out.write_shift(2, 0);
	out.write_shift(1, 0);
	return out;
}

function write_BIFF2BERR(r, c, val, t) {
	var out = new_buf(9);
	write_BIFF2Cell(out, r, c);
	if(t == 'e') { out.write_shift(1, val); out.write_shift(1, 1); }
	else { out.write_shift(1, val?1:0); out.write_shift(1, 0); }
	return out;
}

/* TODO: codepage, large strings */
function write_BIFF2LABEL(r, c, val) {
	var out = new_buf(8 + 2*val.length);
	write_BIFF2Cell(out, r, c);
	out.write_shift(1, val.length);
	out.write_shift(val.length, val, 'sbcs');
	return out.l < out.length ? out.slice(0, out.l) : out;
}

function write_ws_biff2_cell(ba, cell, R, C) {
	if(cell.v != null) switch(cell.t) {
		case 'd': case 'n':
			var v = cell.t == 'd' ? datenum(parseDate(cell.v)) : cell.v;
			if((v == (v|0)) && (v >= 0) && (v < 65536))
				write_biff_rec(ba, 0x0002, write_BIFF2INT(R, C, v));
			else
				write_biff_rec(ba, 0x0003, write_BIFF2NUM(R,C, v));
			return;
		case 'b': case 'e': write_biff_rec(ba, 0x0005, write_BIFF2BERR(R, C, cell.v, cell.t)); return;
		/* TODO: codepage, sst */
		case 's': case 'str':
			write_biff_rec(ba, 0x0004, write_BIFF2LABEL(R, C, cell.v));
			return;
	}
	write_biff_rec(ba, 0x0001, write_BIFF2Cell(null, R, C));
}

function write_ws_biff2(ba, ws, idx, opts) {
	var dense = Array.isArray(ws);
	var range = safe_decode_range(ws['!ref'] || "A1"), ref, rr = "", cols = [];
	if(range.e.c > 0xFF || range.e.r > 0x3FFF) {
		if(opts.WTF) throw new Error("Range " + (ws['!ref'] || "A1") + " exceeds format limit A1:IV16384");
		range.e.c = Math.min(range.e.c, 0xFF);
		range.e.r = Math.min(range.e.c, 0x3FFF);
		ref = encode_range(range);
	}
	for(var R = range.s.r; R <= range.e.r; ++R) {
		rr = encode_row(R);
		for(var C = range.s.c; C <= range.e.c; ++C) {
			if(R === range.s.r) cols[C] = encode_col(C);
			ref = cols[C] + rr;
			var cell = dense ? (ws[R]||[])[C] : ws[ref];
			if(!cell) continue;
			/* write cell */
			write_ws_biff2_cell(ba, cell, R, C, opts);
		}
	}
}

/* Based on test files */
function write_biff2_buf(wb, opts) {
	var o = opts || {};
	if(DENSE != null && o.dense == null) o.dense = DENSE;
	var ba = buf_array();
	var idx = 0;
	for(var i=0;i<wb.SheetNames.length;++i) if(wb.SheetNames[i] == o.sheet) idx=i;
	if(idx == 0 && !!o.sheet && wb.SheetNames[0] != o.sheet) throw new Error("Sheet not found: " + o.sheet);
	write_biff_rec(ba, 0x0009, write_BOF(wb, 0x10, o));
	/* ... */
	write_ws_biff2(ba, wb.Sheets[wb.SheetNames[idx]], idx, o, wb);
	/* ... */
	write_biff_rec(ba, 0x000A);
	return ba.end();
}

function write_FONTS_biff8(ba, data, opts) {
	write_biff_rec(ba, "Font", write_Font({
		sz:12,
		color: {theme:1},
		name: "Arial",
		family: 2,
		scheme: "minor"
	}, opts));
}


function write_FMTS_biff8(ba, NF, opts) {
	if(!NF) return;
	[[5,8],[23,26],[41,44],[/*63*/50,/*66],[164,*/392]].forEach(function(r) {
for(var i = r[0]; i <= r[1]; ++i) if(NF[i] != null) write_biff_rec(ba, "Format", write_Format(i, NF[i], opts));
	});
}

function write_FEAT(ba, ws) {
	/* [MS-XLS] 2.4.112 */
	var o = new_buf(19);
	o.write_shift(4, 0x867); o.write_shift(4, 0); o.write_shift(4, 0);
	o.write_shift(2, 3); o.write_shift(1, 1); o.write_shift(4, 0);
	write_biff_rec(ba, "FeatHdr", o);
	/* [MS-XLS] 2.4.111 */
	o = new_buf(39);
	o.write_shift(4, 0x868); o.write_shift(4, 0); o.write_shift(4, 0);
	o.write_shift(2, 3); o.write_shift(1, 0); o.write_shift(4, 0);
	o.write_shift(2, 1); o.write_shift(4, 4); o.write_shift(2, 0);
	write_Ref8U(safe_decode_range(ws['!ref']||"A1"), o);
	o.write_shift(4, 4);
	write_biff_rec(ba, "Feat", o);
}

function write_CELLXFS_biff8(ba, opts) {
	for(var i = 0; i < 16; ++i) write_biff_rec(ba, "XF", write_XF({numFmtId:0, style:true}, 0, opts));
	opts.cellXfs.forEach(function(c) {
		write_biff_rec(ba, "XF", write_XF(c, 0, opts));
	});
}

function write_ws_biff8_hlinks(ba, ws) {
	for(var R=0; R<ws['!links'].length; ++R) {
		var HL = ws['!links'][R];
		write_biff_rec(ba, "HLink", write_HLink(HL));
		if(HL[1].Tooltip) write_biff_rec(ba, "HLinkTooltip", write_HLinkTooltip(HL));
	}
	delete ws['!links'];
}

function write_ws_biff8_cell(ba, cell, R, C, opts) {
	var os = 16 + get_cell_style(opts.cellXfs, cell, opts);
	if(cell.v != null) switch(cell.t) {
		case 'd': case 'n':
			var v = cell.t == 'd' ? datenum(parseDate(cell.v)) : cell.v;
			/* TODO: emit RK as appropriate */
			write_biff_rec(ba, "Number", write_Number(R, C, v, os, opts));
			return;
		case 'b': case 'e': write_biff_rec(ba, 0x0205, write_BoolErr(R, C, cell.v, os, opts, cell.t)); return;
		/* TODO: codepage, sst */
		case 's': case 'str':
			write_biff_rec(ba, "Label", write_Label(R, C, cell.v, os, opts));
			return;
	}
	write_biff_rec(ba, "Blank", write_XLSCell(R, C, os));
}

/* [MS-XLS] 2.1.7.20.5 */
function write_ws_biff8(idx, opts, wb) {
	var ba = buf_array();
	var s = wb.SheetNames[idx], ws = wb.Sheets[s] || {};
	var _WB = ((wb||{}).Workbook||{});
	var _sheet = ((_WB.Sheets||[])[idx]||{});
	var dense = Array.isArray(ws);
	var b8 = opts.biff == 8;
	var ref, rr = "", cols = [];
	var range = safe_decode_range(ws['!ref'] || "A1");
	var MAX_ROWS = b8 ? 65536 : 16384;
	if(range.e.c > 0xFF || range.e.r >= MAX_ROWS) {
		if(opts.WTF) throw new Error("Range " + (ws['!ref'] || "A1") + " exceeds format limit A1:IV16384");
		range.e.c = Math.min(range.e.c, 0xFF);
		range.e.r = Math.min(range.e.c, MAX_ROWS-1);
	}

	write_biff_rec(ba, 0x0809, write_BOF(wb, 0x10, opts));
	/* ... */
	write_biff_rec(ba, "CalcMode", writeuint16(1));
	write_biff_rec(ba, "CalcCount", writeuint16(100));
	write_biff_rec(ba, "CalcRefMode", writebool(true));
	write_biff_rec(ba, "CalcIter", writebool(false));
	write_biff_rec(ba, "CalcDelta", write_Xnum(0.001));
	write_biff_rec(ba, "CalcSaveRecalc", writebool(true));
	write_biff_rec(ba, "PrintRowCol", writebool(false));
	write_biff_rec(ba, "PrintGrid", writebool(false));
	write_biff_rec(ba, "GridSet", writeuint16(1));
	write_biff_rec(ba, "Guts", write_Guts([0,0]));
	/* ... */
	write_biff_rec(ba, "HCenter", writebool(false));
	write_biff_rec(ba, "VCenter", writebool(false));
	/* ... */
	write_biff_rec(ba, 0x200, write_Dimensions(range, opts));
	/* ... */

	if(b8) ws['!links'] = [];
	for(var R = range.s.r; R <= range.e.r; ++R) {
		rr = encode_row(R);
		for(var C = range.s.c; C <= range.e.c; ++C) {
			if(R === range.s.r) cols[C] = encode_col(C);
			ref = cols[C] + rr;
			var cell = dense ? (ws[R]||[])[C] : ws[ref];
			if(!cell) continue;
			/* write cell */
			write_ws_biff8_cell(ba, cell, R, C, opts);
			if(b8 && cell.l) ws['!links'].push([ref, cell.l]);
		}
	}
	var cname = _sheet.CodeName || _sheet.name || s;
	/* ... */
	if(b8 && _WB.Views) write_biff_rec(ba, "Window2", write_Window2(_WB.Views[0]));
	/* ... */
	if(b8 && (ws['!merges']||[]).length) write_biff_rec(ba, "MergeCells", write_MergeCells(ws['!merges']));
	/* ... */
	if(b8) write_ws_biff8_hlinks(ba, ws);
	/* ... */
	write_biff_rec(ba, "CodeName", write_XLUnicodeString(cname, opts));
	/* ... */
	if(b8) write_FEAT(ba, ws);
	/* ... */
	write_biff_rec(ba, "EOF");
	return ba.end();
}

/* [MS-XLS] 2.1.7.20.3 */
function write_biff8_global(wb, bufs, opts) {
	var A = buf_array();
	var _WB = ((wb||{}).Workbook||{});
	var _sheets = (_WB.Sheets||[]);
	var _wb = _WB.WBProps||{};
	var b8 = opts.biff == 8, b5 = opts.biff == 5;
	write_biff_rec(A, 0x0809, write_BOF(wb, 0x05, opts));
	if(opts.bookType == "xla") write_biff_rec(A, "Addin");
	write_biff_rec(A, "InterfaceHdr", b8 ? writeuint16(0x04b0) : null);
	write_biff_rec(A, "Mms", writezeroes(2));
	if(b5) write_biff_rec(A, "ToolbarHdr");
	if(b5) write_biff_rec(A, "ToolbarEnd");
	write_biff_rec(A, "InterfaceEnd");
	write_biff_rec(A, "WriteAccess", write_WriteAccess("SheetJS", opts));
	write_biff_rec(A, "CodePage", writeuint16(b8 ? 0x04b0 : 0x04E4));
	if(b8) write_biff_rec(A, "DSF", writeuint16(0));
	if(b8) write_biff_rec(A, "Excel9File");
	write_biff_rec(A, "RRTabId", write_RRTabId(wb.SheetNames.length));
	if(b8 && wb.vbaraw) {
		write_biff_rec(A, "ObProj");
		var cname = _wb.CodeName || "ThisWorkbook";
		write_biff_rec(A, "CodeName", write_XLUnicodeString(cname, opts));
	}
	write_biff_rec(A, "BuiltInFnGroupCount", writeuint16(0x11));
	write_biff_rec(A, "WinProtect", writebool(false));
	write_biff_rec(A, "Protect", writebool(false));
	write_biff_rec(A, "Password", writeuint16(0));
	if(b8) write_biff_rec(A, "Prot4Rev", writebool(false));
	if(b8) write_biff_rec(A, "Prot4RevPass", writeuint16(0));
	write_biff_rec(A, "Window1", write_Window1(opts));
	write_biff_rec(A, "Backup", writebool(false));
	write_biff_rec(A, "HideObj", writeuint16(0));
	write_biff_rec(A, "Date1904", writebool(safe1904(wb)=="true"));
	write_biff_rec(A, "CalcPrecision", writebool(true));
	if(b8) write_biff_rec(A, "RefreshAll", writebool(false));
	write_biff_rec(A, "BookBool", writeuint16(0));
	/* ... */
	write_FONTS_biff8(A, wb, opts);
	write_FMTS_biff8(A, wb.SSF, opts);
	write_CELLXFS_biff8(A, opts);
	/* ... */
	if(b8) write_biff_rec(A, "UsesELFs", writebool(false));
	var a = A.end();

	var C = buf_array();
	if(b8) write_biff_rec(C, "Country", write_Country());
	/* BIFF8: [SST *Continue] ExtSST */
	write_biff_rec(C, "EOF");
	var c = C.end();

	var B = buf_array();
	var blen = 0, j = 0;
	for(j = 0; j < wb.SheetNames.length; ++j) blen += (b8 ? 12 : 11) + (b8 ? 2 : 1) * wb.SheetNames[j].length;
	var start = a.length + blen + c.length;
	for(j = 0; j < wb.SheetNames.length; ++j) {
		var _sheet = _sheets[j] || ({});
		write_biff_rec(B, "BoundSheet8", write_BoundSheet8({pos:start, hs:_sheet.Hidden||0, dt:0, name:wb.SheetNames[j]}, opts));
		start += bufs[j].length;
	}
	/* 1*BoundSheet8 */
	var b = B.end();
	if(blen != b.length) throw new Error("BS8 " + blen + " != " + b.length);

	var out = [];
	if(a.length) out.push(a);
	if(b.length) out.push(b);
	if(c.length) out.push(c);
	return __toBuffer([out]);
}

/* [MS-XLS] 2.1.7.20 Workbook Stream */
function write_biff8_buf(wb, opts) {
	var o = opts || {};
	var bufs = [];

	if(wb && !wb.SSF) {
		wb.SSF = SSF.get_table();
	}
	if(wb && wb.SSF) {
		make_ssf(SSF); SSF.load_table(wb.SSF);
		// $FlowIgnore
		o.revssf = evert_num(wb.SSF); o.revssf[wb.SSF[65535]] = 0;
		o.ssf = wb.SSF;
	}
	o.cellXfs = [];
	o.Strings = []; o.Strings.Count = 0; o.Strings.Unique = 0;
	get_cell_style(o.cellXfs, {}, {revssf:{"General":0}});

	for(var i = 0; i < wb.SheetNames.length; ++i) bufs[bufs.length] = write_ws_biff8(i, o, wb);
	bufs.unshift(write_biff8_global(wb, bufs, o));
	return __toBuffer([bufs]);
}

function write_biff_buf(wb, opts) {
	var o = opts || {};
	switch(o.biff || 2) {
		case 8: case 5: return write_biff8_buf(wb, opts);
		case 4: case 3: case 2: return write_biff2_buf(wb, opts);
	}
	throw new Error("invalid type " + o.bookType + " for BIFF");
}
/* note: browser DOM element cannot see mso- style attrs, must parse */
var HTML_ = (function() {
	function html_to_sheet(str, _opts) {
		var opts = _opts || {};
		if(DENSE != null && opts.dense == null) opts.dense = DENSE;
		var ws = opts.dense ? ([]) : ({});
		var mtch = str.match(/<table/i);
		if(!mtch) throw new Error("Invalid HTML: could not find <table>");
		var mtch2 = str.match(/<\/table/i);
		var i = mtch.index, j = mtch2 && mtch2.index || str.length;
		var rows = split_regex(str.slice(i, j), /(:?<tr[^>]*>)/i, "<tr>");
		var R = -1, C = 0, RS = 0, CS = 0;
		var range = {s:{r:10000000, c:10000000},e:{r:0,c:0}};
		var merges = [];
		for(i = 0; i < rows.length; ++i) {
			var row = rows[i].trim();
			var hd = row.slice(0,3).toLowerCase();
			if(hd == "<tr") { ++R; if(opts.sheetRows && opts.sheetRows <= R) { --R; break; } C = 0; continue; }
			if(hd != "<td" && hd != "<th") continue;
			var cells = row.split(/<\/t[dh]>/i);
			for(j = 0; j < cells.length; ++j) {
				var cell = cells[j].trim();
				if(!cell.match(/<t[dh]/i)) continue;
				var m = cell, cc = 0;
				/* TODO: parse styles etc */
				while(m.charAt(0) == "<" && (cc = m.indexOf(">")) > -1) m = m.slice(cc+1);
				var tag = parsexmltag(cell.slice(0, cell.indexOf(">")));
				CS = tag.colspan ? +tag.colspan : 1;
				if((RS = +tag.rowspan)>1 || CS>1) merges.push({s:{r:R,c:C},e:{r:R + (RS||1) - 1, c:C + CS - 1}});
				var _t = tag.t || "";
				/* TODO: generate stub cells */
				if(!m.length) { C += CS; continue; }
				m = htmldecode(m);
				if(range.s.r > R) range.s.r = R; if(range.e.r < R) range.e.r = R;
				if(range.s.c > C) range.s.c = C; if(range.e.c < C) range.e.c = C;
				if(!m.length) continue;
				var o = {t:'s', v:m};
				if(opts.raw || !m.trim().length || _t == 's'){}
				else if(m === 'TRUE') o = {t:'b', v:true};
				else if(m === 'FALSE') o = {t:'b', v:false};
				else if(!isNaN(fuzzynum(m))) o = {t:'n', v:fuzzynum(m)};
				else if(!isNaN(fuzzydate(m).getDate())) {
					o = ({t:'d', v:parseDate(m)});
					if(!opts.cellDates) o = ({t:'n', v:datenum(o.v)});
					o.z = opts.dateNF || SSF._table[14];
				}
				if(opts.dense) { if(!ws[R]) ws[R] = []; ws[R][C] = o; }
				else ws[encode_cell({r:R, c:C})] = o;
				C += CS;
			}
		}
		ws['!ref'] = encode_range(range);
		return ws;
	}
	function html_to_book(str, opts) {
		return sheet_to_workbook(html_to_sheet(str, opts), opts);
	}
	function make_html_row(ws, r, R, o) {
		var M = (ws['!merges'] ||[]);
		var oo = [];
		for(var C = r.s.c; C <= r.e.c; ++C) {
			var RS = 0, CS = 0;
			for(var j = 0; j < M.length; ++j) {
				if(M[j].s.r > R || M[j].s.c > C) continue;
				if(M[j].e.r < R || M[j].e.c < C) continue;
				if(M[j].s.r < R || M[j].s.c < C) { RS = -1; break; }
				RS = M[j].e.r - M[j].s.r + 1; CS = M[j].e.c - M[j].s.c + 1; break;
			}
			if(RS < 0) continue;
			var coord = encode_cell({r:R,c:C});
			var cell = o.dense ? (ws[R]||[])[C] : ws[coord];
			var sp = {};
			if(RS > 1) sp.rowspan = RS;
			if(CS > 1) sp.colspan = CS;
			/* TODO: html entities */
			var w = (cell && cell.v != null) && (cell.h || escapehtml(cell.w || (format_cell(cell), cell.w) || "")) || "";
			sp.t = cell && cell.t || 'z';
			if(o.editable) w = '<span contenteditable="true">' + w + '</span>';
			sp.id = "sjs-" + coord;
			oo.push(writextag('td', w, sp));
		}
		var preamble = "<tr>";
		return preamble + oo.join("") + "</tr>";
	}
	function make_html_preamble(ws, R, o) {
		var out = [];
		return out.join("") + '<table' + (o && o.id ? ' id="' + o.id + '"' : "") + '>';
	}
	var _BEGIN = '<html><head><meta charset="utf-8"/><title>SheetJS Table Export</title></head><body>';
	var _END = '</body></html>';
	function sheet_to_html(ws, opts/*, wb:?Workbook*/) {
		var o = opts || {};
		var header = o.header != null ? o.header : _BEGIN;
		var footer = o.footer != null ? o.footer : _END;
		var out = [header];
		var r = decode_range(ws['!ref']);
		o.dense = Array.isArray(ws);
		out.push(make_html_preamble(ws, r, o));
		for(var R = r.s.r; R <= r.e.r; ++R) out.push(make_html_row(ws, r, R, o));
		out.push("</table>" + footer);
		return out.join("");
	}

	return {
		to_workbook: html_to_book,
		to_sheet: html_to_sheet,
		_row: make_html_row,
		BEGIN: _BEGIN,
		END: _END,
		_preamble: make_html_preamble,
		from_sheet: sheet_to_html
	};
})();

function parse_dom_table(table, _opts) {
	var opts = _opts || {};
	if(DENSE != null) opts.dense = DENSE;
	var ws = opts.dense ? ([]) : ({});
	var rows = table.getElementsByTagName('tr');
	var sheetRows = opts.sheetRows || 10000000;
	var range = {s:{r:0,c:0},e:{r:0,c:0}};
	var merges = [], midx = 0;
	var rowinfo = [];
	var _R = 0, R = 0, _C, C, RS, CS;
	for(; _R < rows.length && R < sheetRows; ++_R) {
		var row = rows[_R];
		if (is_dom_element_hidden(row)) {
			if (opts.display) continue;
			rowinfo[R] = {hidden: true};
		}
		var elts = (row.children);
		for(_C = C = 0; _C < elts.length; ++_C) {
			var elt = elts[_C];
			if (opts.display && is_dom_element_hidden(elt)) continue;
			var v = htmldecode(elt.innerHTML);
			for(midx = 0; midx < merges.length; ++midx) {
				var m = merges[midx];
				if(m.s.c == C && m.s.r <= R && R <= m.e.r) { C = m.e.c+1; midx = -1; }
			}
			/* TODO: figure out how to extract nonstandard mso- style */
			CS = +elt.getAttribute("colspan") || 1;
			if((RS = +elt.getAttribute("rowspan"))>0 || CS>1) merges.push({s:{r:R,c:C},e:{r:R + (RS||1) - 1, c:C + CS - 1}});
			var o = {t:'s', v:v};
			var _t = elt.getAttribute("t") || "";
			if(v != null) {
				if(v.length == 0) o.t = _t || 'z';
				else if(opts.raw || v.trim().length == 0 || _t == "s"){}
				else if(v === 'TRUE') o = {t:'b', v:true};
				else if(v === 'FALSE') o = {t:'b', v:false};
				else if(!isNaN(fuzzynum(v))) o = {t:'n', v:fuzzynum(v)};
				else if(!isNaN(fuzzydate(v).getDate())) {
					o = ({t:'d', v:parseDate(v)});
					if(!opts.cellDates) o = ({t:'n', v:datenum(o.v)});
					o.z = opts.dateNF || SSF._table[14];
				}
			}
			if(opts.dense) { if(!ws[R]) ws[R] = []; ws[R][C] = o; }
			else ws[encode_cell({c:C, r:R})] = o;
			if(range.e.c < C) range.e.c = C;
			C += CS;
		}
		++R;
	}
	if(merges.length) ws['!merges'] = merges;
	if(rowinfo.length) ws['!rows'] = rowinfo;
	range.e.r = R - 1;
	ws['!ref'] = encode_range(range);
	if(R >= sheetRows) ws['!fullref'] = encode_range((range.e.r = rows.length-_R+R-1,range)); // We can count the real number of rows to parse but we don't to improve the performance
	return ws;
}

function table_to_book(table, opts) {
	return sheet_to_workbook(parse_dom_table(table, opts), opts);
}

function is_dom_element_hidden(element) {
	var display = '';
	var get_computed_style = get_get_computed_style_function(element);
	if(get_computed_style) display = get_computed_style(element).getPropertyValue('display');
	if(!display) display = element.style.display; // Fallback for cases when getComputedStyle is not available (e.g. an old browser or some Node.js environments) or doesn't work (e.g. if the element is not inserted to a document)
	return display === 'none';
}

/* global getComputedStyle */
function get_get_computed_style_function(element) {
	// The proper getComputedStyle implementation is the one defined in the element window
	if(element.ownerDocument.defaultView && typeof element.ownerDocument.defaultView.getComputedStyle === 'function') return element.ownerDocument.defaultView.getComputedStyle;
	// If it is not available, try to get one from the global namespace
	if(typeof getComputedStyle === 'function') return getComputedStyle;
	return null;
}
/* OpenDocument */
var parse_content_xml = (function() {

	var parse_text_p = function(text) {
		/* 6.1.2 White Space Characters */
		var fixed = text
			.replace(/[\t\r\n]/g, " ").trim().replace(/ +/g, " ")
			.replace(/<text:s\/>/g," ")
			.replace(/<text:s text:c="(\d+)"\/>/g, function($$,$1) { return Array(parseInt($1,10)+1).join(" "); })
			.replace(/<text:tab[^>]*\/>/g,"\t")
			.replace(/<text:line-break\/>/g,"\n");
		var v = unescapexml(fixed.replace(/<[^>]*>/g,""));

		return [v];
	};

	var number_formats = {
		/* ods name: [short ssf fmt, long ssf fmt] */
		day:           ["d",   "dd"],
		month:         ["m",   "mm"],
		year:          ["y",   "yy"],
		hours:         ["h",   "hh"],
		minutes:       ["m",   "mm"],
		seconds:       ["s",   "ss"],
		"am-pm":       ["A/P", "AM/PM"],
		"day-of-week": ["ddd", "dddd"],
		era:           ["e",   "ee"],
		/* there is no native representation of LO "Q" format */
		quarter:       ["\\Qm", "m\\\"th quarter\""]
	};

	return function pcx(d, _opts) {
		var opts = _opts || {};
		if(DENSE != null && opts.dense == null) opts.dense = DENSE;
		var str = xlml_normalize(d);
		var state = [], tmp;
		var tag;
		var NFtag = {name:""}, NF = "", pidx = 0;
		var sheetag;
		var rowtag;
		var Sheets = {}, SheetNames = [];
		var ws = opts.dense ? ([]) : ({});
		var Rn, q;
		var ctag = ({value:""});
		var textp = "", textpidx = 0, textptag;
		var textR = [];
		var R = -1, C = -1, range = {s: {r:1000000,c:10000000}, e: {r:0, c:0}};
		var row_ol = 0;
		var number_format_map = {};
		var merges = [], mrange = {}, mR = 0, mC = 0;
		var rowinfo = [], rowpeat = 1, colpeat = 1;
		var arrayf = [];
		var WB = {Names:[]};
		var atag = ({});
		var _Ref = ["", ""];
		var comments = [], comment = ({});
		var creator = "", creatoridx = 0;
		var isstub = false, intable = false;
		var i = 0;
		xlmlregex.lastIndex = 0;
		str = str.replace(/<!--([\s\S]*?)-->/mg,"").replace(/<!DOCTYPE[^\[]*\[[^\]]*\]>/gm,"");
		while((Rn = xlmlregex.exec(str))) switch((Rn[3]=Rn[3].replace(/_.*$/,""))) {

			case 'table': case '工作表': // 9.1.2 <table:table>
				if(Rn[1]==='/') {
					if(range.e.c >= range.s.c && range.e.r >= range.s.r) ws['!ref'] = encode_range(range);
					if(opts.sheetRows > 0 && opts.sheetRows <= range.e.r) {
						ws['!fullref'] = ws['!ref'];
						range.e.r = opts.sheetRows - 1;
						ws['!ref'] = encode_range(range);
					}
					if(merges.length) ws['!merges'] = merges;
					if(rowinfo.length) ws["!rows"] = rowinfo;
					sheetag.name = sheetag['名称'] || sheetag.name;
					if(typeof JSON !== 'undefined') JSON.stringify(sheetag);
					SheetNames.push(sheetag.name);
					Sheets[sheetag.name] = ws;
					intable = false;
				}
				else if(Rn[0].charAt(Rn[0].length-2) !== '/') {
					sheetag = parsexmltag(Rn[0], false);
					R = C = -1;
					range.s.r = range.s.c = 10000000; range.e.r = range.e.c = 0;
					ws = opts.dense ? ([]) : ({}); merges = [];
					rowinfo = [];
					intable = true;
				}
				break;

			case 'table-row-group': // 9.1.9 <table:table-row-group>
				if(Rn[1] === "/") --row_ol; else ++row_ol;
				break;
			case 'table-row': case '行': // 9.1.3 <table:table-row>
				if(Rn[1] === '/') { R+=rowpeat; rowpeat = 1; break; }
				rowtag = parsexmltag(Rn[0], false);
				if(rowtag['行号']) R = rowtag['行号'] - 1; else if(R == -1) R = 0;
				rowpeat = +rowtag['number-rows-repeated'] || 1;
				/* TODO: remove magic */
				if(rowpeat < 10) for(i = 0; i < rowpeat; ++i) if(row_ol > 0) rowinfo[R + i] = {level: row_ol};
				C = -1; break;
			case 'covered-table-cell': // 9.1.5 <table:covered-table-cell>
				if(Rn[1] !== '/') ++C;
				if(opts.sheetStubs) {
					if(opts.dense) { if(!ws[R]) ws[R] = []; ws[R][C] = {t:'z'}; }
					else ws[encode_cell({r:R,c:C})] = {t:'z'};
				}
				textp = ""; textR = [];
				break; /* stub */
			case 'table-cell': case '数据':
				if(Rn[0].charAt(Rn[0].length-2) === '/') {
					++C;
					ctag = parsexmltag(Rn[0], false);
					colpeat = parseInt(ctag['number-columns-repeated']||"1", 10);
					q = ({t:'z', v:null});
					if(ctag.formula && opts.cellFormula != false) q.f = ods_to_csf_formula(unescapexml(ctag.formula));
					if((ctag['数据类型'] || ctag['value-type']) == "string") {
						q.t = "s"; q.v = unescapexml(ctag['string-value'] || "");
						if(opts.dense) {
							if(!ws[R]) ws[R] = [];
							ws[R][C] = q;
						} else {
							ws[encode_cell({r:R,c:C})] = q;
						}
					}
					C+= colpeat-1;
				} else if(Rn[1]!=='/') {
					++C;
					colpeat = 1;
					var rptR = rowpeat ? R + rowpeat - 1 : R;
					if(C > range.e.c) range.e.c = C;
					if(C < range.s.c) range.s.c = C;
					if(R < range.s.r) range.s.r = R;
					if(rptR > range.e.r) range.e.r = rptR;
					ctag = parsexmltag(Rn[0], false);
					comments = []; comment = ({});
					q = ({t:ctag['数据类型'] || ctag['value-type'], v:null});
					if(opts.cellFormula) {
						if(ctag.formula) ctag.formula = unescapexml(ctag.formula);
						if(ctag['number-matrix-columns-spanned'] && ctag['number-matrix-rows-spanned']) {
							mR = parseInt(ctag['number-matrix-rows-spanned'],10) || 0;
							mC = parseInt(ctag['number-matrix-columns-spanned'],10) || 0;
							mrange = {s: {r:R,c:C}, e:{r:R + mR-1,c:C + mC-1}};
							q.F = encode_range(mrange);
							arrayf.push([mrange, q.F]);
						}
						if(ctag.formula) q.f = ods_to_csf_formula(ctag.formula);
						else for(i = 0; i < arrayf.length; ++i)
							if(R >= arrayf[i][0].s.r && R <= arrayf[i][0].e.r)
								if(C >= arrayf[i][0].s.c && C <= arrayf[i][0].e.c)
									q.F = arrayf[i][1];
					}
					if(ctag['number-columns-spanned'] || ctag['number-rows-spanned']) {
						mR = parseInt(ctag['number-rows-spanned'],10) || 0;
						mC = parseInt(ctag['number-columns-spanned'],10) || 0;
						mrange = {s: {r:R,c:C}, e:{r:R + mR-1,c:C + mC-1}};
						merges.push(mrange);
					}

					/* 19.675.2 table:number-columns-repeated */
					if(ctag['number-columns-repeated']) colpeat = parseInt(ctag['number-columns-repeated'], 10);

					/* 19.385 office:value-type */
					switch(q.t) {
						case 'boolean': q.t = 'b'; q.v = parsexmlbool(ctag['boolean-value']); break;
						case 'float': q.t = 'n'; q.v = parseFloat(ctag.value); break;
						case 'percentage': q.t = 'n'; q.v = parseFloat(ctag.value); break;
						case 'currency': q.t = 'n'; q.v = parseFloat(ctag.value); break;
						case 'date': q.t = 'd'; q.v = parseDate(ctag['date-value']);
							if(!opts.cellDates) { q.t = 'n'; q.v = datenum(q.v); }
							q.z = 'm/d/yy'; break;
						case 'time': q.t = 'n'; q.v = parse_isodur(ctag['time-value'])/86400; break;
						case 'number': q.t = 'n'; q.v = parseFloat(ctag['数据数值']); break;
						default:
							if(q.t === 'string' || q.t === 'text' || !q.t) {
								q.t = 's';
								if(ctag['string-value'] != null) { textp = unescapexml(ctag['string-value']); textR = []; }
							} else throw new Error('Unsupported value type ' + q.t);
					}
				} else {
					isstub = false;
					if(q.t === 's') {
						q.v = textp || '';
						if(textR.length) q.R = textR;
						isstub = textpidx == 0;
					}
					if(atag.Target) q.l = atag;
					if(comments.length > 0) { q.c = comments; comments = []; }
					if(textp && opts.cellText !== false) q.w = textp;
					if(!isstub || opts.sheetStubs) {
						if(!(opts.sheetRows && opts.sheetRows <= R)) {
							for(var rpt = 0; rpt < rowpeat; ++rpt) {
								colpeat = parseInt(ctag['number-columns-repeated']||"1", 10);
								if(opts.dense) {
									if(!ws[R + rpt]) ws[R + rpt] = [];
									ws[R + rpt][C] = rpt == 0 ? q : dup(q);
									while(--colpeat > 0) ws[R + rpt][C + colpeat] = dup(q);
								} else {
									ws[encode_cell({r:R + rpt,c:C})] = q;
									while(--colpeat > 0) ws[encode_cell({r:R + rpt,c:C + colpeat})] = dup(q);
								}
								if(range.e.c <= C) range.e.c = C;
							}
						}
					}
					colpeat = parseInt(ctag['number-columns-repeated']||"1", 10);
					C += colpeat-1; colpeat = 0;
					q = {};
					textp = ""; textR = [];
				}
				atag = ({});
				break; // 9.1.4 <table:table-cell>

			/* pure state */
			case 'document': // TODO: <office:document> is the root for FODS
			case 'document-content': case '电子表格文档': // 3.1.3.2 <office:document-content>
			case 'spreadsheet': case '主体': // 3.7 <office:spreadsheet>
			case 'scripts': // 3.12 <office:scripts>
			case 'styles': // TODO <office:styles>
			case 'font-face-decls': // 3.14 <office:font-face-decls>
				if(Rn[1]==='/'){if((tmp=state.pop())[0]!==Rn[3]) throw "Bad state: "+tmp;}
				else if(Rn[0].charAt(Rn[0].length-2) !== '/') state.push([Rn[3], true]);
				break;

			case 'annotation': // 14.1 <office:annotation>
				if(Rn[1]==='/'){
					if((tmp=state.pop())[0]!==Rn[3]) throw "Bad state: "+tmp;
					comment.t = textp;
					if(textR.length) comment.R = textR;
					comment.a = creator;
					comments.push(comment);
				}
				else if(Rn[0].charAt(Rn[0].length-2) !== '/') {state.push([Rn[3], false]);}
				creator = ""; creatoridx = 0;
				textp = ""; textpidx = 0; textR = [];
				break;

			case 'creator': // 4.3.2.7 <dc:creator>
				if(Rn[1]==='/') { creator = str.slice(creatoridx,Rn.index); }
				else creatoridx = Rn.index + Rn[0].length;
				break;

			/* ignore state */
			case 'meta': case '元数据': // TODO: <office:meta> <uof:元数据> FODS/UOF
			case 'settings': // TODO: <office:settings>
			case 'config-item-set': // TODO: <office:config-item-set>
			case 'config-item-map-indexed': // TODO: <office:config-item-map-indexed>
			case 'config-item-map-entry': // TODO: <office:config-item-map-entry>
			case 'config-item-map-named': // TODO: <office:config-item-map-entry>
			case 'shapes': // 9.2.8 <table:shapes>
			case 'frame': // 10.4.2 <draw:frame>
			case 'text-box': // 10.4.3 <draw:text-box>
			case 'image': // 10.4.4 <draw:image>
			case 'data-pilot-tables': // 9.6.2 <table:data-pilot-tables>
			case 'list-style': // 16.30 <text:list-style>
			case 'form': // 13.13 <form:form>
			case 'dde-links': // 9.8 <table:dde-links>
			case 'event-listeners': // TODO
			case 'chart': // TODO
				if(Rn[1]==='/'){if((tmp=state.pop())[0]!==Rn[3]) throw "Bad state: "+tmp;}
				else if(Rn[0].charAt(Rn[0].length-2) !== '/') state.push([Rn[3], false]);
				textp = ""; textpidx = 0; textR = [];
				break;

			case 'scientific-number': // TODO: <number:scientific-number>
				break;
			case 'currency-symbol': // TODO: <number:currency-symbol>
				break;
			case 'currency-style': // TODO: <number:currency-style>
				break;
			case 'number-style': // 16.27.2 <number:number-style>
			case 'percentage-style': // 16.27.9 <number:percentage-style>
			case 'date-style': // 16.27.10 <number:date-style>
			case 'time-style': // 16.27.18 <number:time-style>
				if(Rn[1]==='/'){
					number_format_map[NFtag.name] = NF;
					if((tmp=state.pop())[0]!==Rn[3]) throw "Bad state: "+tmp;
				} else if(Rn[0].charAt(Rn[0].length-2) !== '/') {
					NF = "";
					NFtag = parsexmltag(Rn[0], false);
					state.push([Rn[3], true]);
				} break;

			case 'script': break; // 3.13 <office:script>
			case 'libraries': break; // TODO: <ooo:libraries>
			case 'automatic-styles': break; // 3.15.3 <office:automatic-styles>
			case 'master-styles': break; // TODO: <office:master-styles>

			case 'default-style': // TODO: <style:default-style>
			case 'page-layout': break; // TODO: <style:page-layout>
			case 'style': // 16.2 <style:style>
				break;
			case 'map': break; // 16.3 <style:map>
			case 'font-face': break; // 16.21 <style:font-face>

			case 'paragraph-properties': break; // 17.6 <style:paragraph-properties>
			case 'table-properties': break; // 17.15 <style:table-properties>
			case 'table-column-properties': break; // 17.16 <style:table-column-properties>
			case 'table-row-properties': break; // 17.17 <style:table-row-properties>
			case 'table-cell-properties': break; // 17.18 <style:table-cell-properties>

			case 'number': // 16.27.3 <number:number>
				switch(state[state.length-1][0]) {
					case 'time-style':
					case 'date-style':
						tag = parsexmltag(Rn[0], false);
						NF += number_formats[Rn[3]][tag.style==='long'?1:0]; break;
				} break;

			case 'fraction': break; // TODO 16.27.6 <number:fraction>

			case 'day': // 16.27.11 <number:day>
			case 'month': // 16.27.12 <number:month>
			case 'year': // 16.27.13 <number:year>
			case 'era': // 16.27.14 <number:era>
			case 'day-of-week': // 16.27.15 <number:day-of-week>
			case 'week-of-year': // 16.27.16 <number:week-of-year>
			case 'quarter': // 16.27.17 <number:quarter>
			case 'hours': // 16.27.19 <number:hours>
			case 'minutes': // 16.27.20 <number:minutes>
			case 'seconds': // 16.27.21 <number:seconds>
			case 'am-pm': // 16.27.22 <number:am-pm>
				switch(state[state.length-1][0]) {
					case 'time-style':
					case 'date-style':
						tag = parsexmltag(Rn[0], false);
						NF += number_formats[Rn[3]][tag.style==='long'?1:0]; break;
				} break;

			case 'boolean-style': break; // 16.27.23 <number:boolean-style>
			case 'boolean': break; // 16.27.24 <number:boolean>
			case 'text-style': break; // 16.27.25 <number:text-style>
			case 'text': // 16.27.26 <number:text>
				if(Rn[0].slice(-2) === "/>") break;
				else if(Rn[1]==="/") switch(state[state.length-1][0]) {
					case 'number-style':
					case 'date-style':
					case 'time-style':
						NF += str.slice(pidx, Rn.index);
						break;
				}
				else pidx = Rn.index + Rn[0].length;
				break;

			case 'named-range': // 9.4.12 <table:named-range>
				tag = parsexmltag(Rn[0], false);
				_Ref = ods_to_csf_3D(tag['cell-range-address']);
				var nrange = ({Name:tag.name, Ref:_Ref[0] + '!' + _Ref[1]});
				if(intable) nrange.Sheet = SheetNames.length;
				WB.Names.push(nrange);
				break;

			case 'text-content': break; // 16.27.27 <number:text-content>
			case 'text-properties': break; // 16.27.27 <style:text-properties>
			case 'embedded-text': break; // 16.27.4 <number:embedded-text>

			case 'body': case '电子表格': break; // 3.3 16.9.6 19.726.3

			case 'forms': break; // 12.25.2 13.2
			case 'table-column': break; // 9.1.6 <table:table-column>
			case 'table-header-rows': break; // 9.1.7 <table:table-header-rows>
			case 'table-rows': break; // 9.1.12 <table:table-rows>
			/* TODO: outline levels */
			case 'table-column-group': break; // 9.1.10 <table:table-column-group>
			case 'table-header-columns': break; // 9.1.11 <table:table-header-columns>
			case 'table-columns': break; // 9.1.12 <table:table-columns>

			case 'null-date': break; // 9.4.2 <table:null-date> TODO: date1904

			case 'graphic-properties': break; // 17.21 <style:graphic-properties>
			case 'calculation-settings': break; // 9.4.1 <table:calculation-settings>
			case 'named-expressions': break; // 9.4.11 <table:named-expressions>
			case 'label-range': break; // 9.4.9 <table:label-range>
			case 'label-ranges': break; // 9.4.10 <table:label-ranges>
			case 'named-expression': break; // 9.4.13 <table:named-expression>
			case 'sort': break; // 9.4.19 <table:sort>
			case 'sort-by': break; // 9.4.20 <table:sort-by>
			case 'sort-groups': break; // 9.4.22 <table:sort-groups>

			case 'tab': break; // 6.1.4 <text:tab>
			case 'line-break': break; // 6.1.5 <text:line-break>
			case 'span': break; // 6.1.7 <text:span>
			case 'p': case '文本串': // 5.1.3 <text:p>
				if(Rn[1]==='/' && (!ctag || !ctag['string-value'])) {
					var ptp = parse_text_p(str.slice(textpidx,Rn.index), textptag);
					textp = (textp.length > 0 ? textp + "\n" : "") + ptp[0];
				} else { textptag = parsexmltag(Rn[0], false); textpidx = Rn.index + Rn[0].length; }
				break; // <text:p>
			case 's': break; // <text:s>

			case 'database-range': // 9.4.15 <table:database-range>
				if(Rn[1]==='/') break;
				try {
					_Ref = ods_to_csf_3D(parsexmltag(Rn[0])['target-range-address']);
					Sheets[_Ref[0]]['!autofilter'] = { ref:_Ref[1] };
				} catch(e) {/* empty */}
				break;

			case 'date': break; // <*:date>

			case 'object': break; // 10.4.6.2 <draw:object>
			case 'title': case '标题': break; // <*:title> OR <uof:标题>
			case 'desc': break; // <*:desc>
			case 'binary-data': break; // 10.4.5 TODO: b64 blob

			/* 9.2 Advanced Tables */
			case 'table-source': break; // 9.2.6
			case 'scenario': break; // 9.2.6

			case 'iteration': break; // 9.4.3 <table:iteration>
			case 'content-validations': break; // 9.4.4 <table:
			case 'content-validation': break; // 9.4.5 <table:
			case 'help-message': break; // 9.4.6 <table:
			case 'error-message': break; // 9.4.7 <table:
			case 'database-ranges': break; // 9.4.14 <table:database-ranges>
			case 'filter': break; // 9.5.2 <table:filter>
			case 'filter-and': break; // 9.5.3 <table:filter-and>
			case 'filter-or': break; // 9.5.4 <table:filter-or>
			case 'filter-condition': break; // 9.5.5 <table:filter-condition>

			case 'list-level-style-bullet': break; // 16.31 <text:
			case 'list-level-style-number': break; // 16.32 <text:
			case 'list-level-properties': break; // 17.19 <style:

			/* 7.3 Document Fields */
			case 'sender-firstname': // 7.3.6.2
			case 'sender-lastname': // 7.3.6.3
			case 'sender-initials': // 7.3.6.4
			case 'sender-title': // 7.3.6.5
			case 'sender-position': // 7.3.6.6
			case 'sender-email': // 7.3.6.7
			case 'sender-phone-private': // 7.3.6.8
			case 'sender-fax': // 7.3.6.9
			case 'sender-company': // 7.3.6.10
			case 'sender-phone-work': // 7.3.6.11
			case 'sender-street': // 7.3.6.12
			case 'sender-city': // 7.3.6.13
			case 'sender-postal-code': // 7.3.6.14
			case 'sender-country': // 7.3.6.15
			case 'sender-state-or-province': // 7.3.6.16
			case 'author-name': // 7.3.7.1
			case 'author-initials': // 7.3.7.2
			case 'chapter': // 7.3.8
			case 'file-name': // 7.3.9
			case 'template-name': // 7.3.9
			case 'sheet-name': // 7.3.9
				break;

			case 'event-listener':
				break;
			/* TODO: FODS Properties */
			case 'initial-creator':
			case 'creation-date':
			case 'print-date':
			case 'generator':
			case 'document-statistic':
			case 'user-defined':
			case 'editing-duration':
			case 'editing-cycles':
				break;

			/* TODO: FODS Config */
			case 'config-item':
				break;

			/* TODO: style tokens */
			case 'page-number': break; // TODO <text:page-number>
			case 'page-count': break; // TODO <text:page-count>
			case 'time': break; // TODO <text:time>

			/* 9.3 Advanced Table Cells */
			case 'cell-range-source': break; // 9.3.1 <table:
			case 'detective': break; // 9.3.2 <table:
			case 'operation': break; // 9.3.3 <table:
			case 'highlighted-range': break; // 9.3.4 <table:

			/* 9.6 Data Pilot Tables <table: */
			case 'data-pilot-table': // 9.6.3
			case 'source-cell-range': // 9.6.5
			case 'source-service': // 9.6.6
			case 'data-pilot-field': // 9.6.7
			case 'data-pilot-level': // 9.6.8
			case 'data-pilot-subtotals': // 9.6.9
			case 'data-pilot-subtotal': // 9.6.10
			case 'data-pilot-members': // 9.6.11
			case 'data-pilot-member': // 9.6.12
			case 'data-pilot-display-info': // 9.6.13
			case 'data-pilot-sort-info': // 9.6.14
			case 'data-pilot-layout-info': // 9.6.15
			case 'data-pilot-field-reference': // 9.6.16
			case 'data-pilot-groups': // 9.6.17
			case 'data-pilot-group': // 9.6.18
			case 'data-pilot-group-member': // 9.6.19
				break;

			/* 10.3 Drawing Shapes */
			case 'rect': // 10.3.2
				break;

			/* 14.6 DDE Connections */
			case 'dde-connection-decls': // 14.6.2 <text:
			case 'dde-connection-decl': // 14.6.3 <text:
			case 'dde-link': // 14.6.4 <table:
			case 'dde-source': // 14.6.5 <office:
				break;

			case 'properties': break; // 13.7 <form:properties>
			case 'property': break; // 13.8 <form:property>

			case 'a': // 6.1.8 hyperlink
				if(Rn[1]!== '/') {
					atag = parsexmltag(Rn[0], false);
					if(!atag.href) break;
					atag.Target = atag.href; delete atag.href;
					if(atag.Target.charAt(0) == "#" && atag.Target.indexOf(".") > -1) {
						_Ref = ods_to_csf_3D(atag.Target.slice(1));
						atag.Target = "#" + _Ref[0] + "!" + _Ref[1];
					}
				}
				break;

			/* non-standard */
			case 'table-protection': break;
			case 'data-pilot-grand-total': break; // <table:
			case 'office-document-common-attrs': break; // bare
			default: switch(Rn[2]) {
				case 'dc:':       // TODO: properties
				case 'calcext:':  // ignore undocumented extensions
				case 'loext:':    // ignore undocumented extensions
				case 'ooo:':      // ignore undocumented extensions
				case 'chartooo:': // ignore undocumented extensions
				case 'draw:':     // TODO: drawing
				case 'style:':    // TODO: styles
				case 'chart:':    // TODO: charts
				case 'form:':     // TODO: forms
				case 'uof:':      // TODO: uof
				case '表:':       // TODO: uof
				case '字:':       // TODO: uof
					break;
				default: if(opts.WTF) throw new Error(Rn);
			}
		}
		var out = ({
			Sheets: Sheets,
			SheetNames: SheetNames,
			Workbook: WB
		});
		if(opts.bookSheets) delete out.Sheets;
		return out;
	};
})();

function parse_ods(zip, opts) {
	opts = opts || ({});
	var ods = !!safegetzipfile(zip, 'objectdata');
	if(ods) parse_manifest(getzipdata(zip, 'META-INF/manifest.xml'), opts);
	var content = getzipstr(zip, 'content.xml');
	if(!content) throw new Error("Missing content.xml in " + (ods ? "ODS" : "UOF")+ " file");
	var wb = parse_content_xml(ods ? content : utf8read(content), opts);
	if(safegetzipfile(zip, 'meta.xml')) wb.Props = parse_core_props(getzipdata(zip, 'meta.xml'));
	return wb;
}
function parse_fods(data, opts) {
	return parse_content_xml(data, opts);
}

/* OpenDocument */
var write_styles_ods = (function() {
	var payload = '<office:document-styles ' + wxt_helper({
		'xmlns:office':   "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
		'xmlns:table':    "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
		'xmlns:style':    "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
		'xmlns:text':     "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
		'xmlns:draw':     "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
		'xmlns:fo':       "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
		'xmlns:xlink':    "http://www.w3.org/1999/xlink",
		'xmlns:dc':       "http://purl.org/dc/elements/1.1/",
		'xmlns:number':   "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
		'xmlns:svg':      "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
		'xmlns:of':       "urn:oasis:names:tc:opendocument:xmlns:of:1.2",
		'office:version': "1.2"
	}) + '></office:document-styles>';
	return function wso() {
		return XML_HEADER + payload;
	};
})();
var write_content_ods = (function() {
	/* 6.1.2 White Space Characters */
	var write_text_p = function(text) {
		return escapexml(text)
			.replace(/  +/g, function($$){return '<text:s text:c="'+$$.length+'"/>';})
			.replace(/\t/g, "<text:tab/>")
			.replace(/\n/g, "<text:line-break/>")
			.replace(/^ /, "<text:s/>").replace(/ $/, "<text:s/>");
	};

	var null_cell_xml = '          <table:table-cell />\n';
	var covered_cell_xml = '          <table:covered-table-cell/>\n';
	var write_ws = function(ws, wb, i) {
		/* Section 9 Tables */
		var o = [];
		o.push('      <table:table table:name="' + escapexml(wb.SheetNames[i]) + '">\n');
		var R=0,C=0, range = decode_range(ws['!ref']);
		var marr = ws['!merges'] || [], mi = 0;
		var dense = Array.isArray(ws);
		for(R = 0; R < range.s.r; ++R) o.push('        <table:table-row></table:table-row>\n');
		for(; R <= range.e.r; ++R) {
			o.push('        <table:table-row>\n');
			for(C=0; C < range.s.c; ++C) o.push(null_cell_xml);
			for(; C <= range.e.c; ++C) {
				var skip = false, ct = {}, textp = "";
				for(mi = 0; mi != marr.length; ++mi) {
					if(marr[mi].s.c > C) continue;
					if(marr[mi].s.r > R) continue;
					if(marr[mi].e.c < C) continue;
					if(marr[mi].e.r < R) continue;
					if(marr[mi].s.c != C || marr[mi].s.r != R) skip = true;
					ct['table:number-columns-spanned'] = (marr[mi].e.c - marr[mi].s.c + 1);
					ct['table:number-rows-spanned'] =    (marr[mi].e.r - marr[mi].s.r + 1);
					break;
				}
				if(skip) { o.push(covered_cell_xml); continue; }
				var ref = encode_cell({r:R, c:C}), cell = dense ? (ws[R]||[])[C]: ws[ref];
				if(cell && cell.f) {
					ct['table:formula'] = escapexml(csf_to_ods_formula(cell.f));
					if(cell.F) {
						if(cell.F.slice(0, ref.length) == ref) {
							var _Fref = decode_range(cell.F);
							ct['table:number-matrix-columns-spanned'] = (_Fref.e.c - _Fref.s.c + 1);
							ct['table:number-matrix-rows-spanned'] =    (_Fref.e.r - _Fref.s.r + 1);
						}
					}
				}
				if(!cell) { o.push(null_cell_xml); continue; }
				switch(cell.t) {
					case 'b':
						textp = (cell.v ? 'TRUE' : 'FALSE');
						ct['office:value-type'] = "boolean";
						ct['office:boolean-value'] = (cell.v ? 'true' : 'false');
						break;
					case 'n':
						textp = (cell.w||String(cell.v||0));
						ct['office:value-type'] = "float";
						ct['office:value'] = (cell.v||0);
						break;
					case 's': case 'str':
						textp = cell.v;
						ct['office:value-type'] = "string";
						break;
					case 'd':
						textp = (cell.w||(parseDate(cell.v).toISOString()));
						ct['office:value-type'] = "date";
						ct['office:date-value'] = (parseDate(cell.v).toISOString());
						ct['table:style-name'] = "ce1";
						break;
					//case 'e':
					default: o.push(null_cell_xml); continue;
				}
				var text_p = write_text_p(textp);
				if(cell.l && cell.l.Target) {
					var _tgt = cell.l.Target; _tgt = _tgt.charAt(0) == "#" ? "#" + csf_to_ods_3D(_tgt.slice(1)) : _tgt;
					text_p = writextag('text:a', text_p, {'xlink:href': _tgt});
				}
				o.push('          ' + writextag('table:table-cell', writextag('text:p', text_p, {}), ct) + '\n');
			}
			o.push('        </table:table-row>\n');
		}
		o.push('      </table:table>\n');
		return o.join("");
	};

	var write_automatic_styles_ods = function(o) {
		o.push(' <office:automatic-styles>\n');
		o.push('  <number:date-style style:name="N37" number:automatic-order="true">\n');
		o.push('   <number:month number:style="long"/>\n');
		o.push('   <number:text>/</number:text>\n');
		o.push('   <number:day number:style="long"/>\n');
		o.push('   <number:text>/</number:text>\n');
		o.push('   <number:year/>\n');
		o.push('  </number:date-style>\n');
		o.push('  <style:style style:name="ce1" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N37"/>\n');
		o.push(' </office:automatic-styles>\n');
	};

	return function wcx(wb, opts) {
		var o = [XML_HEADER];
		/* 3.1.3.2 */
		var attr = wxt_helper({
			'xmlns:office':       "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
			'xmlns:table':        "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
			'xmlns:style':        "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
			'xmlns:text':         "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
			'xmlns:draw':         "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
			'xmlns:fo':           "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
			'xmlns:xlink':        "http://www.w3.org/1999/xlink",
			'xmlns:dc':           "http://purl.org/dc/elements/1.1/",
			'xmlns:meta':         "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
			'xmlns:number':       "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
			'xmlns:presentation': "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0",
			'xmlns:svg':          "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
			'xmlns:chart':        "urn:oasis:names:tc:opendocument:xmlns:chart:1.0",
			'xmlns:dr3d':         "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0",
			'xmlns:math':         "http://www.w3.org/1998/Math/MathML",
			'xmlns:form':         "urn:oasis:names:tc:opendocument:xmlns:form:1.0",
			'xmlns:script':       "urn:oasis:names:tc:opendocument:xmlns:script:1.0",
			'xmlns:ooo':          "http://openoffice.org/2004/office",
			'xmlns:ooow':         "http://openoffice.org/2004/writer",
			'xmlns:oooc':         "http://openoffice.org/2004/calc",
			'xmlns:dom':          "http://www.w3.org/2001/xml-events",
			'xmlns:xforms':       "http://www.w3.org/2002/xforms",
			'xmlns:xsd':          "http://www.w3.org/2001/XMLSchema",
			'xmlns:xsi':          "http://www.w3.org/2001/XMLSchema-instance",
			'xmlns:sheet':        "urn:oasis:names:tc:opendocument:sh33tjs:1.0",
			'xmlns:rpt':          "http://openoffice.org/2005/report",
			'xmlns:of':           "urn:oasis:names:tc:opendocument:xmlns:of:1.2",
			'xmlns:xhtml':        "http://www.w3.org/1999/xhtml",
			'xmlns:grddl':        "http://www.w3.org/2003/g/data-view#",
			'xmlns:tableooo':     "http://openoffice.org/2009/table",
			'xmlns:drawooo':      "http://openoffice.org/2010/draw",
			'xmlns:calcext':      "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
			'xmlns:loext':        "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0",
			'xmlns:field':        "urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0",
			'xmlns:formx':        "urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0",
			'xmlns:css3t':        "http://www.w3.org/TR/css3-text/",
			'office:version':     "1.2"
		});

		var fods = wxt_helper({
			'xmlns:config':    "urn:oasis:names:tc:opendocument:xmlns:config:1.0",
			'office:mimetype': "application/vnd.oasis.opendocument.spreadsheet"
		});

		if(opts.bookType == "fods") o.push('<office:document' + attr + fods + '>\n');
		else o.push('<office:document-content' + attr  + '>\n');
		write_automatic_styles_ods(o);
		o.push('  <office:body>\n');
		o.push('    <office:spreadsheet>\n');
		for(var i = 0; i != wb.SheetNames.length; ++i) o.push(write_ws(wb.Sheets[wb.SheetNames[i]], wb, i, opts));
		o.push('    </office:spreadsheet>\n');
		o.push('  </office:body>\n');
		if(opts.bookType == "fods") o.push('</office:document>');
		else o.push('</office:document-content>');
		return o.join("");
	};
})();

function write_ods(wb, opts) {
	if(opts.bookType == "fods") return write_content_ods(wb, opts);

var zip = new jszip();
	var f = "";

	var manifest = [];
	var rdf = [];

	/* Part 3 Section 3.3 MIME Media Type */
	f = "mimetype";
	zip.file(f, "application/vnd.oasis.opendocument.spreadsheet");

	/* Part 1 Section 2.2 Documents */
	f = "content.xml";
	zip.file(f, write_content_ods(wb, opts));
	manifest.push([f, "text/xml"]);
	rdf.push([f, "ContentFile"]);

	/* TODO: these are hard-coded styles to satiate excel */
	f = "styles.xml";
	zip.file(f, write_styles_ods(wb, opts));
	manifest.push([f, "text/xml"]);
	rdf.push([f, "StylesFile"]);

	/* TODO: this is hard-coded to satiate excel */
	f = "meta.xml";
	zip.file(f, write_meta_ods());
	manifest.push([f, "text/xml"]);
	rdf.push([f, "MetadataFile"]);

	/* Part 3 Section 6 Metadata Manifest File */
	f = "manifest.rdf";
	zip.file(f, write_rdf(rdf/*, opts*/));
	manifest.push([f, "application/rdf+xml"]);

	/* Part 3 Section 4 Manifest File */
	f = "META-INF/manifest.xml";
	zip.file(f, write_manifest(manifest/*, opts*/));

	return zip;
}

function write_sheet_index(wb, sheet) {
	if(!sheet) return 0;
	var idx = wb.SheetNames.indexOf(sheet);
	if(idx == -1) throw new Error("Sheet not found: " + sheet);
	return idx;
}

function write_obj_str(factory) {
	return function write_str(wb, o) {
		var idx = write_sheet_index(wb, o.sheet);
		return factory.from_sheet(wb.Sheets[wb.SheetNames[idx]], o, wb);
	};
}

var write_htm_str = write_obj_str(HTML_);
var write_csv_str = write_obj_str({from_sheet:sheet_to_csv});
var write_slk_str = write_obj_str(SYLK);
var write_dif_str = write_obj_str(DIF);
var write_prn_str = write_obj_str(PRN);
var write_rtf_str = write_obj_str(RTF);
var write_txt_str = write_obj_str({from_sheet:sheet_to_txt});
var write_dbf_buf = write_obj_str(DBF);
var write_eth_str = write_obj_str(ETH);

function fix_opts_func(defaults) {
	return function fix_opts(opts) {
		for(var i = 0; i != defaults.length; ++i) {
			var d = defaults[i];
			if(opts[d[0]] === undefined) opts[d[0]] = d[1];
			if(d[2] === 'n') opts[d[0]] = Number(opts[d[0]]);
		}
	};
}

var fix_read_opts = fix_opts_func([
	['cellNF', false], /* emit cell number format string as .z */
	['cellHTML', true], /* emit html string as .h */
	['cellFormula', true], /* emit formulae as .f */
	['cellStyles', false], /* emits style/theme as .s */
	['cellText', true], /* emit formatted text as .w */
	['cellDates', false], /* emit date cells with type `d` */

	['sheetStubs', false], /* emit empty cells */
	['sheetRows', 0, 'n'], /* read n rows (0 = read all rows) */

	['bookDeps', false], /* parse calculation chains */
	['bookSheets', false], /* only try to get sheet names (no Sheets) */
	['bookProps', false], /* only try to get properties (no Sheets) */
	['bookFiles', false], /* include raw file structure (keys, files, cfb) */
	['bookVBA', false], /* include vba raw data (vbaraw) */

	['password',''], /* password */
	['WTF', false] /* WTF mode (throws errors) */
]);


var fix_write_opts = fix_opts_func([
	['cellDates', false], /* write date cells with type `d` */

	['bookSST', false], /* Generate Shared String Table */

	['bookType', 'xlsx'], /* Type of workbook (xlsx/m/b) */

	['compression', false], /* Use file compression */

	['WTF', false] /* WTF mode (throws errors) */
]);
function get_sheet_type(n) {
	if(RELS.WS.indexOf(n) > -1) return "sheet";
	if(RELS.CS && n == RELS.CS) return "chart";
	if(RELS.DS && n == RELS.DS) return "dialog";
	if(RELS.MS && n == RELS.MS) return "macro";
	return (n && n.length) ? n : "sheet";
}
function safe_parse_wbrels(wbrels, sheets) {
	if(!wbrels) return 0;
	try {
		wbrels = sheets.map(function pwbr(w) { if(!w.id) w.id = w.strRelID; return [w.name, wbrels['!id'][w.id].Target, get_sheet_type(wbrels['!id'][w.id].Type)]; });
	} catch(e) { return null; }
	return !wbrels || wbrels.length === 0 ? null : wbrels;
}

function safe_parse_sheet(zip, path, relsPath, sheet, idx, sheetRels, sheets, stype, opts, wb, themes, styles) {
	try {
		sheetRels[sheet]=parse_rels(getzipstr(zip, relsPath, true), path);
		var data = getzipdata(zip, path);
		var _ws;
		switch(stype) {
			case 'sheet':  _ws = parse_ws(data, path, idx, opts, sheetRels[sheet], wb, themes, styles); break;
			case 'chart':  _ws = parse_cs(data, path, idx, opts, sheetRels[sheet], wb, themes, styles);
				if(!_ws || !_ws['!chart']) break;
				var dfile = resolve_path(_ws['!chart'].Target, path);
				var drelsp = get_rels_path(dfile);
				var draw = parse_drawing(getzipstr(zip, dfile, true), parse_rels(getzipstr(zip, drelsp, true), dfile));
				var chartp = resolve_path(draw, dfile);
				var crelsp = get_rels_path(chartp);
				_ws = parse_chart(getzipstr(zip, chartp, true), chartp, opts, parse_rels(getzipstr(zip, crelsp, true), chartp), wb, _ws);
				break;
			case 'macro':  _ws = parse_ms(data, path, idx, opts, sheetRels[sheet], wb, themes, styles); break;
			case 'dialog': _ws = parse_ds(data, path, idx, opts, sheetRels[sheet], wb, themes, styles); break;
		}
		sheets[sheet] = _ws;
	} catch(e) { if(opts.WTF) throw e; }
}

function strip_front_slash(x) { return x.charAt(0) == '/' ? x.slice(1) : x; }

function parse_zip(zip, opts) {
	make_ssf(SSF);
	opts = opts || {};
	fix_read_opts(opts);

	/* OpenDocument Part 3 Section 2.2.1 OpenDocument Package */
	if(safegetzipfile(zip, 'META-INF/manifest.xml')) return parse_ods(zip, opts);
	/* UOC */
	if(safegetzipfile(zip, 'objectdata.xml')) return parse_ods(zip, opts);
	/* Numbers */
	if(safegetzipfile(zip, 'Index/Document.iwa')) throw new Error('Unsupported NUMBERS file');

	var entries = zipentries(zip);
	var dir = parse_ct((getzipstr(zip, '[Content_Types].xml')));
	var xlsb = false;
	var sheets, binname;
	if(dir.workbooks.length === 0) {
		binname = "xl/workbook.xml";
		if(getzipdata(zip,binname, true)) dir.workbooks.push(binname);
	}
	if(dir.workbooks.length === 0) {
		binname = "xl/workbook.bin";
		if(!getzipdata(zip,binname,true)) throw new Error("Could not find workbook");
		dir.workbooks.push(binname);
		xlsb = true;
	}
	if(dir.workbooks[0].slice(-3) == "bin") xlsb = true;

	var themes = ({});
	var styles = ({});
	if(!opts.bookSheets && !opts.bookProps) {
		strs = [];
		if(dir.sst) try { strs=parse_sst(getzipdata(zip, strip_front_slash(dir.sst)), dir.sst, opts); } catch(e) { if(opts.WTF) throw e; }

		if(opts.cellStyles && dir.themes.length) themes = parse_theme(getzipstr(zip, dir.themes[0].replace(/^\//,''), true)||"",dir.themes[0], opts);

		if(dir.style) styles = parse_sty(getzipdata(zip, strip_front_slash(dir.style)), dir.style, themes, opts);
	}

	/*var externbooks = */dir.links.map(function(link) {
		return parse_xlink(getzipdata(zip, strip_front_slash(link)), link, opts);
	});

	var wb = parse_wb(getzipdata(zip, strip_front_slash(dir.workbooks[0])), dir.workbooks[0], opts);

	var props = {}, propdata = "";

	if(dir.coreprops.length) {
		propdata = getzipdata(zip, strip_front_slash(dir.coreprops[0]), true);
		if(propdata) props = parse_core_props(propdata);
		if(dir.extprops.length !== 0) {
			propdata = getzipdata(zip, strip_front_slash(dir.extprops[0]), true);
			if(propdata) parse_ext_props(propdata, props, opts);
		}
	}

	var custprops = {};
	if(!opts.bookSheets || opts.bookProps) {
		if (dir.custprops.length !== 0) {
			propdata = getzipstr(zip, strip_front_slash(dir.custprops[0]), true);
			if(propdata) custprops = parse_cust_props(propdata, opts);
		}
	}

	var out = ({});
	if(opts.bookSheets || opts.bookProps) {
		if(wb.Sheets) sheets = wb.Sheets.map(function pluck(x){ return x.name; });
		else if(props.Worksheets && props.SheetNames.length > 0) sheets=props.SheetNames;
		if(opts.bookProps) { out.Props = props; out.Custprops = custprops; }
		if(opts.bookSheets && typeof sheets !== 'undefined') out.SheetNames = sheets;
		if(opts.bookSheets ? out.SheetNames : opts.bookProps) return out;
	}
	sheets = {};

	var deps = {};
	if(opts.bookDeps && dir.calcchain) deps=parse_cc(getzipdata(zip, strip_front_slash(dir.calcchain)),dir.calcchain,opts);

	var i=0;
	var sheetRels = ({});
	var path, relsPath;

	{
		var wbsheets = wb.Sheets;
		props.Worksheets = wbsheets.length;
		props.SheetNames = [];
		for(var j = 0; j != wbsheets.length; ++j) {
			props.SheetNames[j] = wbsheets[j].name;
		}
	}

	var wbext = xlsb ? "bin" : "xml";
	var wbrelsi = dir.workbooks[0].lastIndexOf("/");
	var wbrelsfile = (dir.workbooks[0].slice(0, wbrelsi+1) + "_rels/" + dir.workbooks[0].slice(wbrelsi+1) + ".rels").replace(/^\//,"");
	if(!safegetzipfile(zip, wbrelsfile)) wbrelsfile = 'xl/_rels/workbook.' + wbext + '.rels';
	var wbrels = parse_rels(getzipstr(zip, wbrelsfile, true), wbrelsfile);
	if(wbrels) wbrels = safe_parse_wbrels(wbrels, wb.Sheets);

	/* Numbers iOS hack */
	var nmode = (getzipdata(zip,"xl/worksheets/sheet.xml",true))?1:0;
	for(i = 0; i != props.Worksheets; ++i) {
		var stype = "sheet";
		if(wbrels && wbrels[i]) {
			path = 'xl/' + (wbrels[i][1]).replace(/[\/]?xl\//, "");
			if(!safegetzipfile(zip, path)) path = wbrels[i][1];
			if(!safegetzipfile(zip, path)) path = wbrelsfile.replace(/_rels\/.*$/,"") + wbrels[i][1];
			stype = wbrels[i][2];
		} else {
			path = 'xl/worksheets/sheet'+(i+1-nmode)+"." + wbext;
			path = path.replace(/sheet0\./,"sheet.");
		}
		relsPath = path.replace(/^(.*)(\/)([^\/]*)$/, "$1/_rels/$3.rels");
		safe_parse_sheet(zip, path, relsPath, props.SheetNames[i], i, sheetRels, sheets, stype, opts, wb, themes, styles);
	}

	if(dir.comments) parse_comments(zip, dir.comments, sheets, sheetRels, opts);

	out = ({
		Directory: dir,
		Workbook: wb,
		Props: props,
		Custprops: custprops,
		Deps: deps,
		Sheets: sheets,
		SheetNames: props.SheetNames,
		Strings: strs,
		Styles: styles,
		Themes: themes,
		SSF: SSF.get_table()
	});
	if(opts.bookFiles) {
		out.keys = entries;
		out.files = zip.files;
	}
	if(opts.bookVBA) {
		if(dir.vba.length > 0) out.vbaraw = getzipdata(zip,strip_front_slash(dir.vba[0]),true);
		else if(dir.defaults && dir.defaults.bin === CT_VBA) out.vbaraw = getzipdata(zip, 'xl/vbaProject.bin',true);
	}
	return out;
}

/* [MS-OFFCRYPTO] 2.1.1 */
function parse_xlsxcfb(cfb, _opts) {
	var opts = _opts || {};
	var f = 'Workbook', data = CFB.find(cfb, f);
	try {
	f = '/!DataSpaces/Version';
	data = CFB.find(cfb, f); if(!data || !data.content) throw new Error("ECMA-376 Encrypted file missing " + f);
	/*var version = */parse_DataSpaceVersionInfo(data.content);

	/* 2.3.4.1 */
	f = '/!DataSpaces/DataSpaceMap';
	data = CFB.find(cfb, f); if(!data || !data.content) throw new Error("ECMA-376 Encrypted file missing " + f);
	var dsm = parse_DataSpaceMap(data.content);
	if(dsm.length !== 1 || dsm[0].comps.length !== 1 || dsm[0].comps[0].t !== 0 || dsm[0].name !== "StrongEncryptionDataSpace" || dsm[0].comps[0].v !== "EncryptedPackage")
		throw new Error("ECMA-376 Encrypted file bad " + f);

	/* 2.3.4.2 */
	f = '/!DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace';
	data = CFB.find(cfb, f); if(!data || !data.content) throw new Error("ECMA-376 Encrypted file missing " + f);
	var seds = parse_DataSpaceDefinition(data.content);
	if(seds.length != 1 || seds[0] != "StrongEncryptionTransform")
		throw new Error("ECMA-376 Encrypted file bad " + f);

	/* 2.3.4.3 */
	f = '/!DataSpaces/TransformInfo/StrongEncryptionTransform/!Primary';
	data = CFB.find(cfb, f); if(!data || !data.content) throw new Error("ECMA-376 Encrypted file missing " + f);
	/*var hdr = */parse_Primary(data.content);
	} catch(e) {}

	f = '/EncryptionInfo';
	data = CFB.find(cfb, f); if(!data || !data.content) throw new Error("ECMA-376 Encrypted file missing " + f);
	var einfo = parse_EncryptionInfo(data.content);

	/* 2.3.4.4 */
	f = '/EncryptedPackage';
	data = CFB.find(cfb, f); if(!data || !data.content) throw new Error("ECMA-376 Encrypted file missing " + f);

/*global decrypt_agile */
if(einfo[0] == 0x04 && typeof decrypt_agile !== 'undefined') return decrypt_agile(einfo[1], data.content, opts.password || "", opts);
/*global decrypt_std76 */
if(einfo[0] == 0x02 && typeof decrypt_std76 !== 'undefined') return decrypt_std76(einfo[1], data.content, opts.password || "", opts);
	throw new Error("File is password-protected");
}

function write_zip(wb, opts) {
	_shapeid = 1024;
	if(opts.bookType == "ods") return write_ods(wb, opts);
	if(wb && !wb.SSF) {
		wb.SSF = SSF.get_table();
	}
	if(wb && wb.SSF) {
		make_ssf(SSF); SSF.load_table(wb.SSF);
		// $FlowIgnore
		opts.revssf = evert_num(wb.SSF); opts.revssf[wb.SSF[65535]] = 0;
		opts.ssf = wb.SSF;
	}
	opts.rels = {}; opts.wbrels = {};
	opts.Strings = []; opts.Strings.Count = 0; opts.Strings.Unique = 0;
	if(browser_has_Map) opts.revStrings = new Map();
	else { opts.revStrings = {}; opts.revStrings.foo = []; delete opts.revStrings.foo; }
	var wbext = opts.bookType == "xlsb" ? "bin" : "xml";
	var vbafmt = VBAFMTS.indexOf(opts.bookType) > -1;
	var ct = new_ct();
	fix_write_opts(opts = opts || {});
var zip = new jszip();
	var f = "", rId = 0;

	opts.cellXfs = [];
	get_cell_style(opts.cellXfs, {}, {revssf:{"General":0}});

	if(!wb.Props) wb.Props = {};

	f = "docProps/core.xml";
	zip.file(f, write_core_props(wb.Props, opts));
	ct.coreprops.push(f);
	add_rels(opts.rels, 2, f, RELS.CORE_PROPS);

f = "docProps/app.xml";
	if(wb.Props && wb.Props.SheetNames){/* empty */}
	else if(!wb.Workbook || !wb.Workbook.Sheets) wb.Props.SheetNames = wb.SheetNames;
	else {
		var _sn = [];
		for(var _i = 0; _i < wb.SheetNames.length; ++_i)
			if((wb.Workbook.Sheets[_i]||{}).Hidden != 2) _sn.push(wb.SheetNames[_i]);
		wb.Props.SheetNames = _sn;
	}
	wb.Props.Worksheets = wb.Props.SheetNames.length;
	zip.file(f, write_ext_props(wb.Props, opts));
	ct.extprops.push(f);
	add_rels(opts.rels, 3, f, RELS.EXT_PROPS);

	if(wb.Custprops !== wb.Props && keys(wb.Custprops||{}).length > 0) {
		f = "docProps/custom.xml";
		zip.file(f, write_cust_props(wb.Custprops, opts));
		ct.custprops.push(f);
		add_rels(opts.rels, 4, f, RELS.CUST_PROPS);
	}

	for(rId=1;rId <= wb.SheetNames.length; ++rId) {
		var wsrels = {'!id':{}};
		var ws = wb.Sheets[wb.SheetNames[rId-1]];
		var _type = (ws || {})["!type"] || "sheet";
		switch(_type) {
		case "chart": /*
			f = "xl/chartsheets/sheet" + rId + "." + wbext;
			zip.file(f, write_cs(rId-1, f, opts, wb, wsrels));
			ct.charts.push(f);
			add_rels(wsrels, -1, "chartsheets/sheet" + rId + "." + wbext, RELS.CS);
			break; */
			/* falls through */
		default:
			f = "xl/worksheets/sheet" + rId + "." + wbext;
			zip.file(f, write_ws(rId-1, f, opts, wb, wsrels));
			ct.sheets.push(f);
			add_rels(opts.wbrels, -1, "worksheets/sheet" + rId + "." + wbext, RELS.WS[0]);
		}

		if(ws) {
			var comments = ws['!comments'];
			var need_vml = false;
			if(comments && comments.length > 0) {
				var cf = "xl/comments" + rId + "." + wbext;
				zip.file(cf, write_cmnt(comments, cf, opts));
				ct.comments.push(cf);
				add_rels(wsrels, -1, "../comments" + rId + "." + wbext, RELS.CMNT);
				need_vml = true;
			}
			if(ws['!legacy']) {
				if(need_vml) zip.file("xl/drawings/vmlDrawing" + (rId) + ".vml", write_comments_vml(rId, ws['!comments']));
			}
			delete ws['!comments'];
			delete ws['!legacy'];
		}

		if(wsrels['!id'].rId1) zip.file(get_rels_path(f), write_rels(wsrels));
	}

	if(opts.Strings != null && opts.Strings.length > 0) {
		f = "xl/sharedStrings." + wbext;
		zip.file(f, write_sst(opts.Strings, f, opts));
		ct.strs.push(f);
		add_rels(opts.wbrels, -1, "sharedStrings." + wbext, RELS.SST);
	}

	f = "xl/workbook." + wbext;
	zip.file(f, write_wb(wb, f, opts));
	ct.workbooks.push(f);
	add_rels(opts.rels, 1, f, RELS.WB);

	/* TODO: something more intelligent with themes */

	f = "xl/theme/theme1.xml";
	zip.file(f, write_theme(wb.Themes, opts));
	ct.themes.push(f);
	add_rels(opts.wbrels, -1, "theme/theme1.xml", RELS.THEME);

	/* TODO: something more intelligent with styles */

	f = "xl/styles." + wbext;
	zip.file(f, write_sty(wb, f, opts));
	ct.styles.push(f);
	add_rels(opts.wbrels, -1, "styles." + wbext, RELS.STY);

	if(wb.vbaraw && vbafmt) {
		f = "xl/vbaProject.bin";
		zip.file(f, wb.vbaraw);
		ct.vba.push(f);
		add_rels(opts.wbrels, -1, "vbaProject.bin", RELS.VBA);
	}

	zip.file("[Content_Types].xml", write_ct(ct, opts));
	zip.file('_rels/.rels', write_rels(opts.rels));
	zip.file('xl/_rels/workbook.' + wbext + '.rels', write_rels(opts.wbrels));

	delete opts.revssf; delete opts.ssf;
	return zip;
}
function firstbyte(f,o) {
	var x = "";
	switch((o||{}).type || "base64") {
		case 'buffer': return [f[0], f[1], f[2], f[3]];
		case 'base64': x = Base64.decode(f.slice(0,24)); break;
		case 'binary': x = f; break;
		case 'array':  return [f[0], f[1], f[2], f[3]];
		default: throw new Error("Unrecognized type " + (o && o.type || "undefined"));
	}
	return [x.charCodeAt(0), x.charCodeAt(1), x.charCodeAt(2), x.charCodeAt(3)];
}

function read_cfb(cfb, opts) {
	if(CFB.find(cfb, "EncryptedPackage")) return parse_xlsxcfb(cfb, opts);
	return parse_xlscfb(cfb, opts);
}

function read_zip(data, opts) {
var zip, d = data;
	var o = opts||{};
	if(!o.type) o.type = (has_buf && Buffer.isBuffer(data)) ? "buffer" : "base64";
	switch(o.type) {
		case "base64": zip = new jszip(d, { base64:true }); break;
		case "binary": case "array": zip = new jszip(d, { base64:false }); break;
		case "buffer": zip = new jszip(d); break;
		default: throw new Error("Unrecognized type " + o.type);
	}
	return parse_zip(zip, o);
}

function read_plaintext(data, o) {
	var i = 0;
	main: while(i < data.length) switch(data.charCodeAt(i)) {
		case 0x0A: case 0x0D: case 0x20: ++i; break;
		case 0x3C: return parse_xlml(data.slice(i),o);
		default: break main;
	}
	return PRN.to_workbook(data, o);
}

function read_plaintext_raw(data, o) {
	var str = "", bytes = firstbyte(data, o);
	switch(o.type) {
		case 'base64': str = Base64.decode(data); break;
		case 'binary': str = data; break;
		case 'buffer': str = data.toString('binary'); break;
		case 'array': str = cc2str(data); break;
		default: throw new Error("Unrecognized type " + o.type);
	}
	if(bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) str = utf8read(str);
	return read_plaintext(str, o);
}

function read_utf16(data, o) {
	var d = data;
	if(o.type == 'base64') d = Base64.decode(d);
	d = cptable.utils.decode(1200, d.slice(2), 'str');
	o.type = "binary";
	return read_plaintext(d, o);
}

function bstrify(data) {
	return !data.match(/[^\x00-\x7F]/) ? data : utf8write(data);
}

function read_prn(data, d, o, str) {
	if(str) { o.type = "string"; return PRN.to_workbook(data, o); }
	return PRN.to_workbook(d, o);
}

function readSync(data, opts) {
	reset_cp();
	if(typeof ArrayBuffer !== 'undefined' && data instanceof ArrayBuffer) return readSync(new Uint8Array(data), opts);
	var d = data, n = [0,0,0,0], str = false;
	var o = opts||{};
	_ssfopts = {};
	if(o.dateNF) _ssfopts.dateNF = o.dateNF;
	if(!o.type) o.type = (has_buf && Buffer.isBuffer(data)) ? "buffer" : "base64";
	if(o.type == "file") { o.type = has_buf ? "buffer" : "binary"; d = read_binary(data); }
	if(o.type == "string") { str = true; o.type = "binary"; o.codepage = 65001; d = bstrify(data); }
	if(o.type == 'array' && typeof Uint8Array !== 'undefined' && data instanceof Uint8Array && typeof ArrayBuffer !== 'undefined') {
		// $FlowIgnore
		var ab=new ArrayBuffer(3), vu=new Uint8Array(ab); vu.foo="bar";
		// $FlowIgnore
		if(!vu.foo) {o=dup(o); o.type='array'; return readSync(ab2a(d), o);}
	}
	switch((n = firstbyte(d, o))[0]) {
		case 0xD0: return read_cfb(CFB.read(d, o), o);
		case 0x09: return parse_xlscfb(d, o);
		case 0x3C: return parse_xlml(d, o);
		case 0x49: if(n[1] === 0x44) return read_wb_ID(d, o); break;
		case 0x54: if(n[1] === 0x41 && n[2] === 0x42 && n[3] === 0x4C) return DIF.to_workbook(d, o); break;
		case 0x50: return (n[1] === 0x4B && n[2] < 0x09 && n[3] < 0x09) ? read_zip(d, o) : read_prn(data, d, o, str);
		case 0xEF: return n[3] === 0x3C ? parse_xlml(d, o) : read_prn(data, d, o, str);
		case 0xFF: if(n[1] === 0xFE) { return read_utf16(d, o); } break;
		case 0x00: if(n[1] === 0x00 && n[2] >= 0x02 && n[3] === 0x00) return WK_.to_workbook(d, o); break;
		case 0x03: case 0x83: case 0x8B: case 0x8C: return DBF.to_workbook(d, o);
		case 0x7B: if(n[1] === 0x5C && n[2] === 0x72 && n[3] === 0x74) return RTF.to_workbook(d, o); break;
		case 0x0A: case 0x0D: case 0x20: return read_plaintext_raw(d, o);
	}
	if(n[2] <= 12 && n[3] <= 31) return DBF.to_workbook(d, o);
	return read_prn(data, d, o, str);
}

function readFileSync(filename, opts) {
	var o = opts||{}; o.type = 'file';
	return readSync(filename, o);
}
function write_cfb_ctr(cfb, o) {
	switch(o.type) {
		case "base64": case "binary": break;
		case "buffer": case "array": o.type = ""; break;
		case "file": return write_dl(o.file, CFB.write(cfb, {type:has_buf ? 'buffer' : ""}));
		case "string": throw new Error("'string' output type invalid for '" + o.bookType + "' files");
		default: throw new Error("Unrecognized type " + o.type);
	}
	return CFB.write(cfb, o);
}

/*global encrypt_agile */
function write_zip_type(wb, opts) {
	var o = opts||{};
	var z = write_zip(wb, o);
	var oopts = {};
	if(o.compression) oopts.compression = 'DEFLATE';
	if(o.password) oopts.type = has_buf ? "nodebuffer" : "string";
	else switch(o.type) {
		case "base64": oopts.type = "base64"; break;
		case "binary": oopts.type = "string"; break;
		case "string": throw new Error("'string' output type invalid for '" + o.bookType + "' files");
		case "buffer":
		case "file": oopts.type = has_buf ? "nodebuffer" : "string"; break;
		default: throw new Error("Unrecognized type " + o.type);
	}
	var out = z.generate(oopts);
	if(o.password && typeof encrypt_agile !== 'undefined') return write_cfb_ctr(encrypt_agile(out, o.password), o);
	if(o.type === "file") return write_dl(o.file, out);
	return o.type == "string" ? utf8read(out) : out;
}

function write_cfb_type(wb, opts) {
	var o = opts||{};
	var cfb = write_xlscfb(wb, o);
	return write_cfb_ctr(cfb, o);
}

function write_string_type(out, opts, bom) {
	if(!bom) bom = "";
	var o = bom + out;
	switch(opts.type) {
		case "base64": return Base64.encode(utf8write(o));
		case "binary": return utf8write(o);
		case "string": return out;
		case "file": return write_dl(opts.file, o, 'utf8');
		case "buffer": {
			// $FlowIgnore
			if(has_buf) return Buffer_from(o, 'utf8');
			else return write_string_type(o, {type:'binary'}).split("").map(function(c) { return c.charCodeAt(0); });
		}
	}
	throw new Error("Unrecognized type " + opts.type);
}

function write_stxt_type(out, opts) {
	switch(opts.type) {
		case "base64": return Base64.encode(out);
		case "binary": return out;
		case "string": return out; /* override in sheet_to_txt */
		case "file": return write_dl(opts.file, out, 'binary');
		case "buffer": {
			// $FlowIgnore
			if(has_buf) return Buffer_from(out, 'binary');
			else return out.split("").map(function(c) { return c.charCodeAt(0); });
		}
	}
	throw new Error("Unrecognized type " + opts.type);
}

/* TODO: test consistency */
function write_binary_type(out, opts) {
	switch(opts.type) {
		case "string":
		case "base64":
		case "binary":
			var bstr = "";
			// $FlowIgnore
			for(var i = 0; i < out.length; ++i) bstr += String.fromCharCode(out[i]);
			return opts.type == 'base64' ? Base64.encode(bstr) : opts.type == 'string' ? utf8read(bstr) : bstr;
		case "file": return write_dl(opts.file, out);
		case "buffer": return out;
		default: throw new Error("Unrecognized type " + opts.type);
	}
}

function writeSync(wb, opts) {
	check_wb(wb);
	var o = opts||{};
	if(o.type == "array") { o.type = "binary"; var out = (writeSync(wb, o)); o.type = "array"; return s2ab(out); }
	switch(o.bookType || 'xlsb') {
		case 'xml':
		case 'xlml': return write_string_type(write_xlml(wb, o), o);
		case 'slk':
		case 'sylk': return write_string_type(write_slk_str(wb, o), o);
		case 'htm':
		case 'html': return write_string_type(write_htm_str(wb, o), o);
		case 'txt': return write_stxt_type(write_txt_str(wb, o), o);
		case 'csv': return write_string_type(write_csv_str(wb, o), o, "\ufeff");
		case 'dif': return write_string_type(write_dif_str(wb, o), o);
		case 'dbf': return write_binary_type(write_dbf_buf(wb, o), o);
		case 'prn': return write_string_type(write_prn_str(wb, o), o);
		case 'rtf': return write_string_type(write_rtf_str(wb, o), o);
		case 'eth': return write_string_type(write_eth_str(wb, o), o);
		case 'fods': return write_string_type(write_ods(wb, o), o);
		case 'biff2': if(!o.biff) o.biff = 2; /* falls through */
		case 'biff3': if(!o.biff) o.biff = 3; /* falls through */
		case 'biff4': if(!o.biff) o.biff = 4; return write_binary_type(write_biff_buf(wb, o), o);
		case 'biff5': if(!o.biff) o.biff = 5; /* falls through */
		case 'biff8':
		case 'xla':
		case 'xls': if(!o.biff) o.biff = 8; return write_cfb_type(wb, o);
		case 'xlsx':
		case 'xlsm':
		case 'xlam':
		case 'xlsb':
		case 'ods': return write_zip_type(wb, o);
		default: throw new Error ("Unrecognized bookType |" + o.bookType + "|");
	}
}

function resolve_book_type(o) {
	if(o.bookType) return;
	var _BT = {
		"xls": "biff8",
		"htm": "html",
		"slk": "sylk",
		"socialcalc": "eth",
		"Sh33tJS": "WTF"
	};
	var ext = o.file.slice(o.file.lastIndexOf(".")).toLowerCase();
	if(ext.match(/^\.[a-z]+$/)) o.bookType = ext.slice(1);
	o.bookType = _BT[o.bookType] || o.bookType;
}

function writeFileSync(wb, filename, opts) {
	var o = opts||{}; o.type = 'file';
	o.file = filename;
	resolve_book_type(o);
	return writeSync(wb, o);
}

function writeFileAsync(filename, wb, opts, cb) {
	var o = opts||{}; o.type = 'file';
	o.file = filename;
	resolve_book_type(o);
	o.type = 'buffer';
	var _cb = cb; if(!(_cb instanceof Function)) _cb = (opts);
	return _fs.writeFile(filename, writeSync(wb, o), _cb);
}
function make_json_row(sheet, r, R, cols, header, hdr, dense, o) {
	var rr = encode_row(R);
	var defval = o.defval, raw = o.raw || !o.hasOwnProperty("raw");
	var isempty = true;
	var row = (header === 1) ? [] : {};
	if(header !== 1) {
		if(Object.defineProperty) try { Object.defineProperty(row, '__rowNum__', {value:R, enumerable:false}); } catch(e) { row.__rowNum__ = R; }
		else row.__rowNum__ = R;
	}
	if(!dense || sheet[R]) for (var C = r.s.c; C <= r.e.c; ++C) {
		var val = dense ? sheet[R][C] : sheet[cols[C] + rr];
		if(val === undefined || val.t === undefined) {
			if(defval === undefined) continue;
			if(hdr[C] != null) { row[hdr[C]] = defval; }
			continue;
		}
		var v = val.v;
		switch(val.t){
			case 'z': if(v == null) break; continue;
			case 'e': v = void 0; break;
			case 's': case 'd': case 'b': case 'n': break;
			default: throw new Error('unrecognized type ' + val.t);
		}
		if(hdr[C] != null) {
			if(v == null) {
				if(defval !== undefined) row[hdr[C]] = defval;
				else if(raw && v === null) row[hdr[C]] = null;
				else continue;
			} else {
				row[hdr[C]] = raw ? v : format_cell(val,v,o);
			}
			if(v != null) isempty = false;
		}
	}
	return { row: row, isempty: isempty };
}


function sheet_to_json(sheet, opts) {
	if(sheet == null || sheet["!ref"] == null) return [];
	var val = {t:'n',v:0}, header = 0, offset = 1, hdr = [], v=0, vv="";
	var r = {s:{r:0,c:0},e:{r:0,c:0}};
	var o = opts || {};
	var range = o.range != null ? o.range : sheet["!ref"];
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
	var cols = [];
	var out = [];
	var outi = 0, counter = 0;
	var dense = Array.isArray(sheet);
	var R = r.s.r, C = 0, CC = 0;
	if(dense && !sheet[R]) sheet[R] = [];
	for(C = r.s.c; C <= r.e.c; ++C) {
		cols[C] = encode_col(C);
		val = dense ? sheet[R][C] : sheet[cols[C] + rr];
		switch(header) {
			case 1: hdr[C] = C - r.s.c; break;
			case 2: hdr[C] = cols[C]; break;
			case 3: hdr[C] = o.header[C - r.s.c]; break;
			default:
				if(val == null) val = {w: "__EMPTY", t: "s"};
				vv = v = format_cell(val, null, o);
				counter = 0;
				for(CC = 0; CC < hdr.length; ++CC) if(hdr[CC] == vv) vv = v + "_" + (++counter);
				hdr[C] = vv;
		}
	}
	for (R = r.s.r + offset; R <= r.e.r; ++R) {
		var row = make_json_row(sheet, r, R, cols, header, hdr, dense, o);
		if((row.isempty === false) || (header === 1 ? o.blankrows !== false : !!o.blankrows)) out[outi++] = row.row;
	}
	out.length = outi;
	return out;
}

var qreg = /"/g;
function make_csv_row(sheet, r, R, cols, fs, rs, FS, o) {
	var isempty = true;
	var row = [], txt = "", rr = encode_row(R);
	for(var C = r.s.c; C <= r.e.c; ++C) {
		if (!cols[C]) continue;
		var val = o.dense ? (sheet[R]||[])[C]: sheet[cols[C] + rr];
		if(val == null) txt = "";
		else if(val.v != null) {
			isempty = false;
			txt = ''+format_cell(val, null, o);
			for(var i = 0, cc = 0; i !== txt.length; ++i) if((cc = txt.charCodeAt(i)) === fs || cc === rs || cc === 34) {txt = "\"" + txt.replace(qreg, '""') + "\""; break; }
			if(txt == "ID") txt = '"ID"';
		} else if(val.f != null && !val.F) {
			isempty = false;
			txt = '=' + val.f; if(txt.indexOf(",") >= 0) txt = '"' + txt.replace(qreg, '""') + '"';
		} else txt = "";
		/* NOTE: Excel CSV does not support array formulae */
		row.push(txt);
	}
	if(o.blankrows === false && isempty) return null;
	return row.join(FS);
}

function sheet_to_csv(sheet, opts) {
	var out = [];
	var o = opts == null ? {} : opts;
	if(sheet == null || sheet["!ref"] == null) return "";
	var r = safe_decode_range(sheet["!ref"]);
	var FS = o.FS !== undefined ? o.FS : ",", fs = FS.charCodeAt(0);
	var RS = o.RS !== undefined ? o.RS : "\n", rs = RS.charCodeAt(0);
	var endregex = new RegExp((FS=="|" ? "\\|" : FS)+"+$");
	var row = "", cols = [];
	o.dense = Array.isArray(sheet);
	var colinfo = o.skipHidden && sheet["!cols"] || [];
	var rowinfo = o.skipHidden && sheet["!rows"] || [];
	for(var C = r.s.c; C <= r.e.c; ++C) if (!((colinfo[C]||{}).hidden)) cols[C] = encode_col(C);
	for(var R = r.s.r; R <= r.e.r; ++R) {
		if ((rowinfo[R]||{}).hidden) continue;
		row = make_csv_row(sheet, r, R, cols, fs, rs, FS, o);
		if(row == null) { continue; }
		if(o.strip) row = row.replace(endregex,"");
		out.push(row + RS);
	}
	delete o.dense;
	return out.join("");
}

function sheet_to_txt(sheet, opts) {
	if(!opts) opts = {}; opts.FS = "\t"; opts.RS = "\n";
	var s = sheet_to_csv(sheet, opts);
	if(typeof cptable == 'undefined' || opts.type == 'string') return s;
	var o = cptable.utils.encode(1200, s, 'str');
	return String.fromCharCode(255) + String.fromCharCode(254) + o;
}

function sheet_to_formulae(sheet) {
	var y = "", x, val="";
	if(sheet == null || sheet["!ref"] == null) return [];
	var r = safe_decode_range(sheet['!ref']), rr = "", cols = [], C;
	var cmds = [];
	var dense = Array.isArray(sheet);
	for(C = r.s.c; C <= r.e.c; ++C) cols[C] = encode_col(C);
	for(var R = r.s.r; R <= r.e.r; ++R) {
		rr = encode_row(R);
		for(C = r.s.c; C <= r.e.c; ++C) {
			y = cols[C] + rr;
			x = dense ? (sheet[R]||[])[C] : sheet[y];
			val = "";
			if(x === undefined) continue;
			else if(x.F != null) {
				y = x.F;
				if(!x.f) continue;
				val = x.f;
				if(y.indexOf(":") == -1) y = y + ":" + y;
			}
			if(x.f != null) val = x.f;
			else if(x.t == 'z') continue;
			else if(x.t == 'n' && x.v != null) val = "" + x.v;
			else if(x.t == 'b') val = x.v ? "TRUE" : "FALSE";
			else if(x.w !== undefined) val = "'" + x.w;
			else if(x.v === undefined) continue;
			else if(x.t == 's') val = "'" + x.v;
			else val = ""+x.v;
			cmds[cmds.length] = y + "=" + val;
		}
	}
	return cmds;
}

function sheet_add_json(_ws, js, opts) {
	var o = opts || {};
	var offset = +!o.skipHeader;
	var ws = _ws || ({});
	var _R = 0, _C = 0;
	if(ws && o.origin != null) {
		if(typeof o.origin == 'number') _R = o.origin;
		else {
			var _origin = typeof o.origin == "string" ? decode_cell(o.origin) : o.origin;
			_R = _origin.r; _C = _origin.c;
		}
	}
	var cell;
	var range = ({s: {c:0, r:0}, e: {c:_C, r:_R + js.length - 1 + offset}});
	if(ws['!ref']) {
		var _range = safe_decode_range(ws['!ref']);
		range.e.c = Math.max(range.e.c, _range.e.c);
		range.e.r = Math.max(range.e.r, _range.e.r);
		if(_R == -1) { _R = range.e.r + 1; range.e.r = _R + js.length - 1 + offset; }
	}
	var hdr = o.header || [], C = 0;

	js.forEach(function (JS, R) {
		keys(JS).forEach(function(k) {
			if((C=hdr.indexOf(k)) == -1) hdr[C=hdr.length] = k;
			var v = JS[k];
			var t = 'z';
			var z = "";
			if(v && typeof v === 'object' && !(v instanceof Date)){
				ws[encode_cell({c:_C + C,r:_R + R + offset})] = v;
			} else {
				if(typeof v == 'number') t = 'n';
				else if(typeof v == 'boolean') t = 'b';
				else if(typeof v == 'string') t = 's';
				else if(v instanceof Date) {
					t = 'd';
					if(!o.cellDates) { t = 'n'; v = datenum(v); }
					z = o.dateNF || SSF._table[14];
				}
				ws[encode_cell({c:_C + C,r:_R + R + offset})] = cell = ({t:t, v:v});
				if(z) cell.z = z;
			}
		});
	});
	range.e.c = Math.max(range.e.c, _C + hdr.length - 1);
	var __R = encode_row(_R);
	if(offset) for(C = 0; C < hdr.length; ++C) ws[encode_col(C + _C) + __R] = {t:'s', v:hdr[C]};
	ws['!ref'] = encode_range(range);
	return ws;
}
function json_to_sheet(js, opts) { return sheet_add_json(null, js, opts); }

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
	sheet_add_aoa: sheet_add_aoa,
	sheet_add_json: sheet_add_json,
	aoa_to_sheet: aoa_to_sheet,
	json_to_sheet: json_to_sheet,
	table_to_sheet: parse_dom_table,
	table_to_book: table_to_book,
	sheet_to_csv: sheet_to_csv,
	sheet_to_txt: sheet_to_txt,
	sheet_to_json: sheet_to_json,
	sheet_to_html: HTML_.from_sheet,
	sheet_to_dif: DIF.from_sheet,
	sheet_to_slk: SYLK.from_sheet,
	sheet_to_eth: ETH.from_sheet,
	sheet_to_formulae: sheet_to_formulae,
	sheet_to_row_object_array: sheet_to_json
};

(function(utils) {
utils.consts = utils.consts || {};
function add_consts(R/*Array<any>*/) { R.forEach(function(a){ utils.consts[a[0]] = a[1]; }); }

function get_default(x, y, z) { return x[y] != null ? x[y] : (x[y] = z); }

/* get cell, creating a stub if necessary */
function ws_get_cell_stub(ws, R, C) {
	/* A1 cell address */
	if(typeof R == "string") return ws[R] || (ws[R] = {t:'z'});
	/* cell address object */
	if(typeof R != "number") return ws_get_cell_stub(ws, encode_cell(R));
	/* R and C are 0-based indices */
	return ws_get_cell_stub(ws, encode_cell({r:R,c:C||0}));
}

/* find sheet index for given name / validate index */
function wb_sheet_idx(wb, sh) {
	if(typeof sh == "number") {
		if(sh >= 0 && wb.SheetNames.length > sh) return sh;
		throw new Error("Cannot find sheet # " + sh);
	} else if(typeof sh == "string") {
		var idx = wb.SheetNames.indexOf(sh);
		if(idx > -1) return idx;
		throw new Error("Cannot find sheet name |" + sh + "|");
	} else throw new Error("Cannot find sheet |" + sh + "|");
}

/* simple blank workbook object */
utils.book_new = function() {
	return { SheetNames: [], Sheets: {} };
};

/* add a worksheet to the end of a given workbook */
utils.book_append_sheet = function(wb, ws, name) {
	if(!name) for(var i = 1; i <= 0xFFFF; ++i) if(wb.SheetNames.indexOf(name = "Sheet" + i) == -1) break;
	if(!name) throw new Error("Too many worksheets");
	check_ws_name(name);
	if(wb.SheetNames.indexOf(name) >= 0) throw new Error("Worksheet with name |" + name + "| already exists!");

	wb.SheetNames.push(name);
	wb.Sheets[name] = ws;
};

/* set sheet visibility (visible/hidden/very hidden) */
utils.book_set_sheet_visibility = function(wb, sh, vis) {
	get_default(wb,"Workbook",{});
	get_default(wb.Workbook,"Sheets",[]);

	var idx = wb_sheet_idx(wb, sh);
	// $FlowIgnore
	get_default(wb.Workbook.Sheets,idx, {});

	switch(vis) {
		case 0: case 1: case 2: break;
		default: throw new Error("Bad sheet visibility setting " + vis);
	}
	// $FlowIgnore
	wb.Workbook.Sheets[idx].Hidden = vis;
};
add_consts([
	["SHEET_VISIBLE", 0],
	["SHEET_HIDDEN", 1],
	["SHEET_VERY_HIDDEN", 2]
]);

/* set number format */
utils.cell_set_number_format = function(cell, fmt) {
	cell.z = fmt;
	return cell;
};

/* set cell hyperlink */
utils.cell_set_hyperlink = function(cell, target, tooltip) {
	if(!target) {
		delete cell.l;
	} else {
		cell.l = ({ Target: target });
		if(tooltip) cell.l.Tooltip = tooltip;
	}
	return cell;
};
utils.cell_set_internal_link = function(cell, range, tooltip) { return utils.cell_set_hyperlink(cell, "#" + range, tooltip); };

/* add to cell comments */
utils.cell_add_comment = function(cell, text, author) {
	if(!cell.c) cell.c = [];
	cell.c.push({t:text, a:author||"SheetJS"});
};

/* set array formula and flush related cells */
utils.sheet_set_array_formula = function(ws, range, formula) {
	var rng = typeof range != "string" ? range : safe_decode_range(range);
	var rngstr = typeof range == "string" ? range : encode_range(range);
	for(var R = rng.s.r; R <= rng.e.r; ++R) for(var C = rng.s.c; C <= rng.e.c; ++C) {
		var cell = ws_get_cell_stub(ws, R, C);
		cell.t = 'n';
		cell.F = rngstr;
		delete cell.v;
		if(R == rng.s.r && C == rng.s.c) cell.f = formula;
	}
	return ws;
};

return utils;
})(utils);

if(has_buf && typeof require != 'undefined') (function() {
	var Readable = {}.Readable;

	var write_csv_stream = function(sheet, opts) {
		var stream = Readable();
		var o = opts == null ? {} : opts;
		if(sheet == null || sheet["!ref"] == null) { stream.push(null); return stream; }
		var r = safe_decode_range(sheet["!ref"]);
		var FS = o.FS !== undefined ? o.FS : ",", fs = FS.charCodeAt(0);
		var RS = o.RS !== undefined ? o.RS : "\n", rs = RS.charCodeAt(0);
		var endregex = new RegExp((FS=="|" ? "\\|" : FS)+"+$");
		var row = "", cols = [];
		o.dense = Array.isArray(sheet);
		var colinfo = o.skipHidden && sheet["!cols"] || [];
		var rowinfo = o.skipHidden && sheet["!rows"] || [];
		for(var C = r.s.c; C <= r.e.c; ++C) if (!((colinfo[C]||{}).hidden)) cols[C] = encode_col(C);
		var R = r.s.r;
		var BOM = false;
		stream._read = function() {
			if(!BOM) { BOM = true; return stream.push("\uFEFF"); }
			if(R > r.e.r) return stream.push(null);
			while(R <= r.e.r) {
				++R;
				if ((rowinfo[R-1]||{}).hidden) continue;
				row = make_csv_row(sheet, r, R-1, cols, fs, rs, FS, o);
				if(row != null) {
					if(o.strip) row = row.replace(endregex,"");
					stream.push(row + RS);
					break;
				}
			}
		};
		return stream;
	};

	var write_html_stream = function(ws, opts) {
		var stream = Readable();

		var o = opts || {};
		var header = o.header != null ? o.header : HTML_.BEGIN;
		var footer = o.footer != null ? o.footer : HTML_.END;
		stream.push(header);
		var r = decode_range(ws['!ref']);
		o.dense = Array.isArray(ws);
		stream.push(HTML_._preamble(ws, r, o));
		var R = r.s.r;
		var end = false;
		stream._read = function() {
			if(R > r.e.r) {
				if(!end) { end = true; stream.push("</table>" + footer); }
				return stream.push(null);
			}
			while(R <= r.e.r) {
				stream.push(HTML_._row(ws, r, R, o));
				++R;
				break;
			}
		};
		return stream;
	};

	var write_json_stream = function(sheet, opts) {
		var stream = Readable({objectMode:true});

		if(sheet == null || sheet["!ref"] == null) { stream.push(null); return stream; }
		var val = {t:'n',v:0}, header = 0, offset = 1, hdr = [], v=0, vv="";
		var r = {s:{r:0,c:0},e:{r:0,c:0}};
		var o = opts || {};
		var range = o.range != null ? o.range : sheet["!ref"];
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
		var cols = [];
		var counter = 0;
		var dense = Array.isArray(sheet);
		var R = r.s.r, C = 0, CC = 0;
		if(dense && !sheet[R]) sheet[R] = [];
		for(C = r.s.c; C <= r.e.c; ++C) {
			cols[C] = encode_col(C);
			val = dense ? sheet[R][C] : sheet[cols[C] + rr];
			switch(header) {
				case 1: hdr[C] = C - r.s.c; break;
				case 2: hdr[C] = cols[C]; break;
				case 3: hdr[C] = o.header[C - r.s.c]; break;
				default:
					if(val == null) val = {w: "__EMPTY", t: "s"};
					vv = v = format_cell(val, null, o);
					counter = 0;
					for(CC = 0; CC < hdr.length; ++CC) if(hdr[CC] == vv) vv = v + "_" + (++counter);
					hdr[C] = vv;
			}
		}
		R = r.s.r + offset;
		stream._read = function() {
			if(R > r.e.r) return stream.push(null);
			while(R <= r.e.r) {
				//if ((rowinfo[R-1]||{}).hidden) continue;
				var row = make_json_row(sheet, r, R, cols, header, hdr, dense, o);
				++R;
				if((row.isempty === false) || (header === 1 ? o.blankrows !== false : !!o.blankrows)) {
					stream.push(row.row);
					break;
				}
			}
		};
		return stream;
	};

	XLSX.stream = {
		to_json: write_json_stream,
		to_html: write_html_stream,
		to_csv: write_csv_stream
	};
})();

XLSX.parse_xlscfb = parse_xlscfb;
XLSX.parse_ods = parse_ods;
XLSX.parse_fods = parse_fods;
XLSX.write_ods = write_ods;
XLSX.parse_zip = parse_zip;
XLSX.read = readSync; //xlsread
XLSX.readFile = readFileSync; //readFile
XLSX.readFileSync = readFileSync;
XLSX.write = writeSync;
XLSX.writeFile = writeFileSync;
XLSX.writeFileSync = writeFileSync;
XLSX.writeFileAsync = writeFileAsync;
XLSX.utils = utils;
XLSX.SSF = SSF;
XLSX.CFB = CFB;
}
/*global define */
if(typeof exports !== 'undefined') make_xlsx_lib(exports);
else if(typeof module !== 'undefined' && module.exports) make_xlsx_lib(module.exports);
else if(typeof define === 'function' && define.amd) define('xlsx', function() { if(!XLSX.version) make_xlsx_lib(XLSX); return XLSX; });
else make_xlsx_lib(XLSX);
/*exported XLS, ODS */
var XLS = XLSX, ODS = XLSX;
