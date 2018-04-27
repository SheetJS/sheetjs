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

function safe1904(wb/*:Workbook*/)/*:string*/ {
	/* TODO: store date1904 somewhere else */
	if(!wb.Workbook) return "false";
	if(!wb.Workbook.WBProps) return "false";
	return parsexmlbool(wb.Workbook.WBProps.date1904) ? "true" : "false";
}

var badchars = "][*?\/\\".split("");
function check_ws_name(n/*:string*/, safe/*:?boolean*/)/*:boolean*/ {
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
