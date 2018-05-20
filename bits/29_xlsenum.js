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
	/*::[*/0x01/*::]*/: { n: 'CodePage', t: VT_I2 },
	/*::[*/0x02/*::]*/: { n: 'Category', t: VT_STRING },
	/*::[*/0x03/*::]*/: { n: 'PresentationFormat', t: VT_STRING },
	/*::[*/0x04/*::]*/: { n: 'ByteCount', t: VT_I4 },
	/*::[*/0x05/*::]*/: { n: 'LineCount', t: VT_I4 },
	/*::[*/0x06/*::]*/: { n: 'ParagraphCount', t: VT_I4 },
	/*::[*/0x07/*::]*/: { n: 'SlideCount', t: VT_I4 },
	/*::[*/0x08/*::]*/: { n: 'NoteCount', t: VT_I4 },
	/*::[*/0x09/*::]*/: { n: 'HiddenCount', t: VT_I4 },
	/*::[*/0x0a/*::]*/: { n: 'MultimediaClipCount', t: VT_I4 },
	/*::[*/0x0b/*::]*/: { n: 'ScaleCrop', t: VT_BOOL },
	/*::[*/0x0c/*::]*/: { n: 'HeadingPairs', t: VT_VECTOR | VT_VARIANT },
	/*::[*/0x0d/*::]*/: { n: 'TitlesOfParts', t: VT_VECTOR | VT_LPSTR },
	/*::[*/0x0e/*::]*/: { n: 'Manager', t: VT_STRING },
	/*::[*/0x0f/*::]*/: { n: 'Company', t: VT_STRING },
	/*::[*/0x10/*::]*/: { n: 'LinksUpToDate', t: VT_BOOL },
	/*::[*/0x11/*::]*/: { n: 'CharacterCount', t: VT_I4 },
	/*::[*/0x13/*::]*/: { n: 'SharedDoc', t: VT_BOOL },
	/*::[*/0x16/*::]*/: { n: 'HyperlinksChanged', t: VT_BOOL },
	/*::[*/0x17/*::]*/: { n: 'AppVersion', t: VT_I4, p: 'version' },
	/*::[*/0x18/*::]*/: { n: 'DigSig', t: VT_BLOB },
	/*::[*/0x1A/*::]*/: { n: 'ContentType', t: VT_STRING },
	/*::[*/0x1B/*::]*/: { n: 'ContentStatus', t: VT_STRING },
	/*::[*/0x1C/*::]*/: { n: 'Language', t: VT_STRING },
	/*::[*/0x1D/*::]*/: { n: 'Version', t: VT_STRING },
	/*::[*/0xFF/*::]*/: {}
};

/* [MS-OSHARED] 2.3.3.2.1.1 Summary Information Property Set PIDSI */
var SummaryPIDSI = {
	/*::[*/0x01/*::]*/: { n: 'CodePage', t: VT_I2 },
	/*::[*/0x02/*::]*/: { n: 'Title', t: VT_STRING },
	/*::[*/0x03/*::]*/: { n: 'Subject', t: VT_STRING },
	/*::[*/0x04/*::]*/: { n: 'Author', t: VT_STRING },
	/*::[*/0x05/*::]*/: { n: 'Keywords', t: VT_STRING },
	/*::[*/0x06/*::]*/: { n: 'Comments', t: VT_STRING },
	/*::[*/0x07/*::]*/: { n: 'Template', t: VT_STRING },
	/*::[*/0x08/*::]*/: { n: 'LastAuthor', t: VT_STRING },
	/*::[*/0x09/*::]*/: { n: 'RevNumber', t: VT_STRING },
	/*::[*/0x0A/*::]*/: { n: 'EditTime', t: VT_FILETIME },
	/*::[*/0x0B/*::]*/: { n: 'LastPrinted', t: VT_FILETIME },
	/*::[*/0x0C/*::]*/: { n: 'CreatedDate', t: VT_FILETIME },
	/*::[*/0x0D/*::]*/: { n: 'ModifiedDate', t: VT_FILETIME },
	/*::[*/0x0E/*::]*/: { n: 'PageCount', t: VT_I4 },
	/*::[*/0x0F/*::]*/: { n: 'WordCount', t: VT_I4 },
	/*::[*/0x10/*::]*/: { n: 'CharCount', t: VT_I4 },
	/*::[*/0x11/*::]*/: { n: 'Thumbnail', t: VT_CF },
	/*::[*/0x12/*::]*/: { n: 'Application', t: VT_STRING },
	/*::[*/0x13/*::]*/: { n: 'DocSecurity', t: VT_I4 },
	/*::[*/0xFF/*::]*/: {}
};

/* [MS-OLEPS] 2.18 */
var SpecialProperties = {
	/*::[*/0x80000000/*::]*/: { n: 'Locale', t: VT_UI4 },
	/*::[*/0x80000003/*::]*/: { n: 'Behavior', t: VT_UI4 },
	/*::[*/0x72627262/*::]*/: {}
};

(function() {
	for(var y in SpecialProperties) if(SpecialProperties.hasOwnProperty(y))
	DocSummaryPIDDSI[y] = SummaryPIDSI[y] = SpecialProperties[y];
})();

var DocSummaryRE/*:{[key:string]:string}*/ = evert_key(DocSummaryPIDDSI, "n");
var SummaryRE/*:{[key:string]:string}*/ = evert_key(SummaryPIDSI, "n");

/* [MS-XLS] 2.4.63 Country/Region codes */
var CountryEnum = {
	/*::[*/0x0001/*::]*/: "US", // United States
	/*::[*/0x0002/*::]*/: "CA", // Canada
	/*::[*/0x0003/*::]*/: "", // Latin America (except Brazil)
	/*::[*/0x0007/*::]*/: "RU", // Russia
	/*::[*/0x0014/*::]*/: "EG", // Egypt
	/*::[*/0x001E/*::]*/: "GR", // Greece
	/*::[*/0x001F/*::]*/: "NL", // Netherlands
	/*::[*/0x0020/*::]*/: "BE", // Belgium
	/*::[*/0x0021/*::]*/: "FR", // France
	/*::[*/0x0022/*::]*/: "ES", // Spain
	/*::[*/0x0024/*::]*/: "HU", // Hungary
	/*::[*/0x0027/*::]*/: "IT", // Italy
	/*::[*/0x0029/*::]*/: "CH", // Switzerland
	/*::[*/0x002B/*::]*/: "AT", // Austria
	/*::[*/0x002C/*::]*/: "GB", // United Kingdom
	/*::[*/0x002D/*::]*/: "DK", // Denmark
	/*::[*/0x002E/*::]*/: "SE", // Sweden
	/*::[*/0x002F/*::]*/: "NO", // Norway
	/*::[*/0x0030/*::]*/: "PL", // Poland
	/*::[*/0x0031/*::]*/: "DE", // Germany
	/*::[*/0x0034/*::]*/: "MX", // Mexico
	/*::[*/0x0037/*::]*/: "BR", // Brazil
	/*::[*/0x003d/*::]*/: "AU", // Australia
	/*::[*/0x0040/*::]*/: "NZ", // New Zealand
	/*::[*/0x0042/*::]*/: "TH", // Thailand
	/*::[*/0x0051/*::]*/: "JP", // Japan
	/*::[*/0x0052/*::]*/: "KR", // Korea
	/*::[*/0x0054/*::]*/: "VN", // Viet Nam
	/*::[*/0x0056/*::]*/: "CN", // China
	/*::[*/0x005A/*::]*/: "TR", // Turkey
	/*::[*/0x0069/*::]*/: "JS", // Ramastan
	/*::[*/0x00D5/*::]*/: "DZ", // Algeria
	/*::[*/0x00D8/*::]*/: "MA", // Morocco
	/*::[*/0x00DA/*::]*/: "LY", // Libya
	/*::[*/0x015F/*::]*/: "PT", // Portugal
	/*::[*/0x0162/*::]*/: "IS", // Iceland
	/*::[*/0x0166/*::]*/: "FI", // Finland
	/*::[*/0x01A4/*::]*/: "CZ", // Czech Republic
	/*::[*/0x0376/*::]*/: "TW", // Taiwan
	/*::[*/0x03C1/*::]*/: "LB", // Lebanon
	/*::[*/0x03C2/*::]*/: "JO", // Jordan
	/*::[*/0x03C3/*::]*/: "SY", // Syria
	/*::[*/0x03C4/*::]*/: "IQ", // Iraq
	/*::[*/0x03C5/*::]*/: "KW", // Kuwait
	/*::[*/0x03C6/*::]*/: "SA", // Saudi Arabia
	/*::[*/0x03CB/*::]*/: "AE", // United Arab Emirates
	/*::[*/0x03CC/*::]*/: "IL", // Israel
	/*::[*/0x03CE/*::]*/: "QA", // Qatar
	/*::[*/0x03D5/*::]*/: "IR", // Iran
	/*::[*/0xFFFF/*::]*/: "US"  // United States
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

