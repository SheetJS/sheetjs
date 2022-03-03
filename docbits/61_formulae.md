#### Formulae

The A1-style formula string is stored in the `f` field.  Even though different
file formats store the formulae in different ways, the formats are translated.
Even though some formats store formulae with a leading equal sign, CSF formulae
do not start with `=`.

<details>
  <summary><b>Formulae File Format Support</b> (click to show)</summary>

| Storage Representation | Formats                  | Read  | Write |
|:-----------------------|:-------------------------|:-----:|:-----:|
| A1-style strings       | XLSX                     |   ✔   |   ✔   |
| RC-style strings       | XLML and plain text      |   ✔   |   ✔   |
| BIFF Parsed formulae   | XLSB and all XLS formats |   ✔   |       |
| OpenFormula formulae   | ODS/FODS/UOS             |   ✔   |   ✔   |
| Lotus Parsed formulae  | All Lotus WK_ formats    |   ✔   |       |

Since Excel prohibits named cells from colliding with names of A1 or RC style
cell references, a (not-so-simple) regex conversion is possible.  BIFF Parsed
formulae and Lotus Parsed formulae have to be explicitly unwound.  OpenFormula
formulae can be converted with regular expressions.

Shared formulae are decompressed and each cell has the formula corresponding to
its cell.  Writers generally do not attempt to generate shared formulae.
</details>

**Single-Cell Formulae**

For simple formulae, the `f` key of the desired cell can be set to the actual
formula text.  This worksheet represents `A1=1`, `A2=2`, and `A3=A1+A2`:

```js
var worksheet = {
  "!ref": "A1:A3",
  A1: { t:'n', v:1 },
  A2: { t:'n', v:2 },
  A3: { t:'n', v:3, f:'A1+A2' }
};
```

Utilities like `aoa_to_sheet` will accept cell objects in lieu of values:

```js
var worksheet = XLSX.utils.aoa_to_sheet([
  [ 1 ], // A1
  [ 2 ], // A2
  [ {t: "n", v: 3, f: "A1+A2"} ] // A3
]);
```

Cells with formula entries but no value will be serialized in a way that Excel
and other spreadsheet tools will recognize.  This library will not automatically
compute formula results!  For example, the following worksheet will include the
`BESSELJ` function but the result will not be available in JavaScript:

```js
var worksheet = XLSX.utils.aoa_to_sheet([
  [ 3.14159, 2 ], // Row "1"
  [ { t:'n', f:'BESSELJ(A1,B1)' } ] // Row "2" will be calculated on file open
}
```

If the actual results are needed in JS, [SheetJS Pro](https://sheetjs.com/pro)
offers a formula calculator component for evaluating expressions, updating
values and dependent cells, and refreshing entire workbooks.


**Array Formulae**

_Assign an array formula_

```js
XLSX.utils.sheet_set_array_formula(worksheet, range, formula);
```

Array formulae are stored in the top-left cell of the array block.  All cells
of an array formula have a `F` field corresponding to the range.  A single-cell
formula can be distinguished from a plain formula by the presence of `F` field.

For example, setting the cell `C1` to the array formula `{=SUM(A1:A3*B1:B3)}`:

```js
// API function
XLSX.utils.sheet_set_array_formula(worksheet, "C1", "SUM(A1:A3*B1:B3)");

// ... OR raw operations
worksheet['C1'] = { t:'n', f: "SUM(A1:A3*B1:B3)", F:"C1:C1" };
```

For a multi-cell array formula, every cell has the same array range but only the
first cell specifies the formula.  Consider `D1:D3=A1:A3*B1:B3`:

```js
// API function
XLSX.utils.sheet_set_array_formula(worksheet, "D1:D3", "A1:A3*B1:B3");

// ... OR raw operations
worksheet['D1'] = { t:'n', F:"D1:D3", f:"A1:A3*B1:B3" };
worksheet['D2'] = { t:'n', F:"D1:D3" };
worksheet['D3'] = { t:'n', F:"D1:D3" };
```

Utilities and writers are expected to check for the presence of a `F` field and
ignore any possible formula element `f` in cells other than the starting cell.
They are not expected to perform validation of the formulae!


**Dynamic Array Formulae**

_Assign a dynamic array formula_

```js
XLSX.utils.sheet_set_array_formula(worksheet, range, formula, true);
```

Released in 2020, Dynamic Array Formulae are supported in the XLSX/XLSM and XLSB
file formats.  They are represented like normal array formulae but have special
cell metadata indicating that the formula should be allowed to adjust the range.

An array formula can be marked as dynamic by setting the cell's `D` property to
true.  The `F` range is expected but can be the set to the current cell:

```js
// API function
XLSX.utils.sheet_set_array_formula(worksheet, "C1", "_xlfn.UNIQUE(A1:A3)", 1);

// ... OR raw operations
worksheet['C1'] = { t: "s", f: "_xlfn.UNIQUE(A1:A3)", F:"C1", D: 1 }; // dynamic
```

**Localization with Function Names**

SheetJS operates at the file level.  Excel stores formula expressions using the
English (United States) function names.  For non-English users, Excel uses a
localized set of function names.

For example, when the computer language and region is set to French (France),
Excel interprets `=SOMME(A1:C3)` as if `SOMME` is the `SUM` function.  However,
in the actual file, Excel stores `SUM(A1:C3)`.

**Prefixed "Future Functions"**

Functions introduced in newer versions of Excel are prefixed with `_xlfn.` when
stored in files.  When writing formula expressions using these functions, the
prefix is required for maximal compatibility:

```js
// Broadest compatibility
XLSX.utils.sheet_set_array_formula(worksheet, "C1", "_xlfn.UNIQUE(A1:A3)", 1);

// Can cause errors in spreadsheet software
XLSX.utils.sheet_set_array_formula(worksheet, "C1", "UNIQUE(A1:A3)", 1);
```

When reading a file, the `xlfn` option preserves the prefixes.

<details>
  <summary><b> Functions requiring `_xlfn.` prefix</b> (click to show)</summary>

This list is growing with each Excel release.

```
ACOT
ACOTH
AGGREGATE
ARABIC
BASE
BETA.DIST
BETA.INV
BINOM.DIST
BINOM.DIST.RANGE
BINOM.INV
BITAND
BITLSHIFT
BITOR
BITRSHIFT
BITXOR
BYCOL
BYROW
CEILING.MATH
CEILING.PRECISE
CHISQ.DIST
CHISQ.DIST.RT
CHISQ.INV
CHISQ.INV.RT
CHISQ.TEST
COMBINA
CONFIDENCE.NORM
CONFIDENCE.T
COT
COTH
COVARIANCE.P
COVARIANCE.S
CSC
CSCH
DAYS
DECIMAL
ERF.PRECISE
ERFC.PRECISE
EXPON.DIST
F.DIST
F.DIST.RT
F.INV
F.INV.RT
F.TEST
FIELDVALUE
FILTERXML
FLOOR.MATH
FLOOR.PRECISE
FORMULATEXT
GAMMA
GAMMA.DIST
GAMMA.INV
GAMMALN.PRECISE
GAUSS
HYPGEOM.DIST
IFNA
IMCOSH
IMCOT
IMCSC
IMCSCH
IMSEC
IMSECH
IMSINH
IMTAN
ISFORMULA
ISOMITTED
ISOWEEKNUM
LAMBDA
LET
LOGNORM.DIST
LOGNORM.INV
MAKEARRAY
MAP
MODE.MULT
MODE.SNGL
MUNIT
NEGBINOM.DIST
NORM.DIST
NORM.INV
NORM.S.DIST
NORM.S.INV
NUMBERVALUE
PDURATION
PERCENTILE.EXC
PERCENTILE.INC
PERCENTRANK.EXC
PERCENTRANK.INC
PERMUTATIONA
PHI
POISSON.DIST
QUARTILE.EXC
QUARTILE.INC
QUERYSTRING
RANDARRAY
RANK.AVG
RANK.EQ
REDUCE
RRI
SCAN
SEC
SECH
SEQUENCE
SHEET
SHEETS
SKEW.P
SORTBY
STDEV.P
STDEV.S
T.DIST
T.DIST.2T
T.DIST.RT
T.INV
T.INV.2T
T.TEST
UNICHAR
UNICODE
UNIQUE
VAR.P
VAR.S
WEBSERVICE
WEIBULL.DIST
XLOOKUP
XOR
Z.TEST
```

</details>

