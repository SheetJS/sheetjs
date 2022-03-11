/* Defaults determined by systematically testing in Excel 2019 */

/* These formats appear to default to other formats in the table */
var default_map/*:Array<number>*/ = [];
var defi = 0;

//  5 -> 37 ...  8 -> 40
for(defi = 5; defi <= 8; ++defi) default_map[defi] = 32 + defi;

// 23 ->  0 ... 26 ->  0
for(defi = 23; defi <= 26; ++defi) default_map[defi] = 0;

// 27 -> 14 ... 31 -> 14
for(defi = 27; defi <= 31; ++defi) default_map[defi] = 14;
// 50 -> 14 ... 58 -> 14
for(defi = 50; defi <= 58; ++defi) default_map[defi] = 14;

// 59 ->  1 ... 62 ->  4
for(defi = 59; defi <= 62; ++defi) default_map[defi] = defi - 58;
// 67 ->  9 ... 68 -> 10
for(defi = 67; defi <= 68; ++defi) default_map[defi] = defi - 58;
// 72 -> 14 ... 75 -> 17
for(defi = 72; defi <= 75; ++defi) default_map[defi] = defi - 58;

// 69 -> 12 ... 71 -> 14
for(defi = 67; defi <= 68; ++defi) default_map[defi] = defi - 57;

// 76 -> 20 ... 78 -> 22
for(defi = 76; defi <= 78; ++defi) default_map[defi] = defi - 56;

// 79 -> 45 ... 81 -> 47
for(defi = 79; defi <= 81; ++defi) default_map[defi] = defi - 34;

// 82 ->  0 ... 65536 -> 0 (omitted)

/* These formats technically refer to Accounting formats with no equivalent */
var default_str = {
	//  5 -- Currency,   0 decimal, black negative
	5:  '"$"#,##0_);\\("$"#,##0\\)',
	63: '"$"#,##0_);\\("$"#,##0\\)',

	//  6 -- Currency,   0 decimal, red   negative
	6:  '"$"#,##0_);[Red]\\("$"#,##0\\)',
	64: '"$"#,##0_);[Red]\\("$"#,##0\\)',

	//  7 -- Currency,   2 decimal, black negative
	7:  '"$"#,##0.00_);\\("$"#,##0.00\\)',
	65: '"$"#,##0.00_);\\("$"#,##0.00\\)',

	//  8 -- Currency,   2 decimal, red   negative
	8:  '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
	66: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',

	// 41 -- Accounting, 0 decimal, No Symbol
	41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',

	// 42 -- Accounting, 0 decimal, $  Symbol
	42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',

	// 43 -- Accounting, 2 decimal, No Symbol
	43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',

	// 44 -- Accounting, 2 decimal, $  Symbol
	44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
};

