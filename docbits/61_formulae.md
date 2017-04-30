#### Formulae

The A1-style formula string is stored in the `f` field.  Even though different
file formats store the formulae in different ways, the formats are translated.
Even though some formats store formulae with a leading equal sign, CSF formulae
do not start with `=`.

<details>
	<summary><b>Representation of A1=1, A2=2, A3=A1+A2</b> (click to show)</summary>

```js
{
	"!ref": "A1:A3",
	A1: { t:'n', v:1 },
	A2: { t:'n', v:2 },
	A3: { t:'n', v:3, f:'A1+A2' }
}
```
</details>

Shared formulae are decompressed and each cell has the formula corresponding to
its cell.  Writers generally do not attempt to generate shared formulae.

Cells with formula entries but no value will be serialized in a way that Excel
and other spreadsheet tools will recognize.  This library will not automatically
compute formula results!  For example, to compute `BESSELJ` in a worksheet:

<details>
	<summary><b>Formula without known value</b> (click to show)</summary>

```js
{
	"!ref": "A1:A3",
	A1: { t:'n', v:3.14159 },
	A2: { t:'n', v:2 },
	A3: { t:'n', f:'BESSELJ(A1,A2)' }
}
```
</details>

**Array Formulae**

Array formulae are stored in the top-left cell of the array block.  All cells
of an array formula have a `F` field corresponding to the range.  A single-cell
formula can be distinguished from a plain formula by the presence of `F` field.

<details>
	<summary><b>Array Formula examples</b> (click to show)</summary>

For example, setting the cell `C1` to the array formula `{=SUM(A1:A3*B1:B3)}`:

```js
worksheet['C1'] = { t:'n', f: "SUM(A1:A3*B1:B3)", F:"C1:C1" };
```

For a multi-cell array formula, every cell has the same array range but only the
first cell specifies the formula.  Consider `D1:D3=A1:A3*B1:B3`:

```js
worksheet['D1'] = { t:'n', F:"D1:D3", f:"A1:A3*B1:B3" };
worksheet['D2'] = { t:'n', F:"D1:D3" };
worksheet['D3'] = { t:'n', F:"D1:D3" };
```

</details>

Utilities and writers are expected to check for the presence of a `F` field and
ignore any possible formula element `f` in cells other than the starting cell.
They are not expected to perform validation of the formulae!

<details>
	<summary><b>Formula Output Utility Function</b> (click to show)</summary>

The `sheet_to_formulae` method generates one line per formula or array formula.
Array formulae are rendered in the form `range=formula` while plain cells are
rendered in the form `cell=formula or value`.  Note that string literals are
prefixed with an apostrophe `'`, consistent with Excel's formula bar display.
</details>

<details>
	<summary><b>Formulae File Format Details</b> (click to show)</summary>

| Storage Representation | Formats                  | Read  | Write |
|:-----------------------|:-------------------------|:-----:|:-----:|
| A1-style strings       | XLSX                     |  :o:  |  :o:  |
| RC-style strings       | XLML and plaintext       |  :o:  |  :o:  |
| BIFF Parsed formulae   | XLSB and all XLS formats |  :o:  |       |
| OpenFormula formulae   | ODS/FODS/UOS             |  :o:  |  :o:  |

Since Excel prohibits named cells from colliding with names of A1 or RC style
cell references, a (not-so-simple) regex conversion is possible.  BIFF Parsed
formulae have to be explicitly unwound.  OpenFormula formulae can be converted
with regexes for the most part.
</details>

