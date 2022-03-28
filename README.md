# xlsx-js-style

## ℹ️ About

SheetJS with Style! Create Excel spreadsheets with basic styling options.

This project is a fork of [SheetJS/sheetjs](https://github.com/sheetjs/sheetjs) combined with code from
[sheetjs-style](https://www.npmjs.com/package/sheetjs-style) (by [ShanaMaid](https://github.com/ShanaMaid/))
and [sheetjs-style-v2](https://www.npmjs.com/package/sheetjs-style-v2) (by [Raul Gonzalez](https://www.npmjs.com/~armandourbina)).

All projects are under the Apache 2.0 License

## 🔌 Installation

Install [npm](https://www.npmjs.org/package/xlsx-js-style):

```sh
npm install xlsx-js-style --save
```

Install browser:

```html
<script lang="javascript" src="dist/xlsx.bundle.js"></script>
```

## 🗒 Core API

Please refer to the [SheetJS](https://sheetjs.com/) documentation for core API reference.

## 🗒 Style API

### Cell Style Example

```js
ws["A1"].s = {
	font: {
		name: "Calibri",
		sz: 24,
		bold: true,
		color: { rgb: "FFFFAA00" },
	},
};
```

### Cell Style Properties

-   Cell styles are specified by a style object that roughly parallels the OpenXML structure.
-   Style properties currently supported are: `alignment`, `border`, `fill`, `font`, `numFmt`.

| Style Prop  | Sub Prop       | Default                    | Description/Values                                                                                |
| :---------- | :------------- | :------------------------- | ------------------------------------------------------------------------------------------------- |
| `alignment` | `vertical`     | TODO:                      | `"bottom"` or `"center"` or `"top"`                                                               |
|             | `horizontal`   | TODO:                      | `"left"` or `"center"` or `"right"`                                                               |
|             | `wrapText`     | TODO:                      | `true ` or ` false`                                                                               |
|             | `readingOrder` | `1`                        | text direction: `1` (LTR) or `2` (RTL)                                                            |
|             | `textRotation` | `0`                        | `0` to `180`, or `255` // `180` is rotated down 180 degrees, `255` is special, aligned vertically |
| `border`    | `top`          | TODO:                      | `{ style: BORDER_STYLE, color: COLOR_SPEC }`                                                      |
|             | `bottom`       |                            | `{ style: BORDER_STYLE, color: COLOR_SPEC }`                                                      |
|             | `left`         |                            | `{ style: BORDER_STYLE, color: COLOR_SPEC }`                                                      |
|             | `right`        |                            | `{ style: BORDER_STYLE, color: COLOR_SPEC }`                                                      |
|             | `diagonal`     |                            | `{ style: BORDER_STYLE, color: COLOR_SPEC }`                                                      |
|             | `diagonalUp`   | `false`                    | `true` or `false`                                                                                 |
|             | `diagonalDown` | `false`                    | `true` or `false`                                                                                 |
| `fill`      | `patternType`  | `"solid"`                  | `"solid"` or `"none"`                                                                             |
|             | `fgColor`      | TODO: (`{ indexed: 64 }`?) | foreground color: see `COLOR_SPEC`                                                                |
|             | `bgColor`      | TODO:                      | background color: see `COLOR_SPEC`                                                                |
| `font`      | `name`         | `"Calibri"`                |                                                                                                   |
|             | `sz`           | `"11"`                     | font size in points                                                                               |
|             | `bold`         | `false`                    | `true` or `false`                                                                                 |
|             | `color`        | TODO:                      | `COLOR_SPEC`                                                                                      |
|             | `italic`       | `false`                    | `true` or `false`                                                                                 |
|             | `outline`      | `false`                    | `true` or `false`                                                                                 |
|             | `shadow`       | `false`                    | `true` or `false`                                                                                 |
|             | `strike`       | `false`                    | `true` or `false`                                                                                 |
|             | `underline`    | `false`                    | `true` or `false`                                                                                 |
|             | `vertAlign`    | `false`                    | `true` or `false` (TODO:WTF is this)                                                              |
| `numFmt`    |                |                            | Ex: `"0"` // integer index to built in formats, see StyleBuilder.SSF property                     |
|             |                |                            | Ex: `"0.00%"` // string matching a built-in format, see StyleBuilder.SSF                          |
|             |                |                            | Ex: `"0.0%"` // string specifying a custom format                                                 |
|             |                |                            | Ex: `"0.00%;\\(0.00%\\);\\-;@"` // string specifying a custom format, escaping special characters |
|             |                |                            | Ex: `"m/dd/yy"` // string a date format using Excel's format notation                             |

#### `COLOR_STYLE` (object)

Colors for `border`, `fill`, `font` are specified as an name/value object - use one of the following:

| Color Prop | Description                                                      | Example                   |
| :--------- | ---------------------------------------------------------------- | ------------------------- |
| `auto`     | use automatic values                                             | `{auto: 1}`               |
| `indexed`  | use indexed value                                                | `{indexed: 64}`           |
| `rgb`      | use hex ARGB value                                               | `{rgb: "FFFFAA00"}`       |
| `theme`    | use theme color index (int) and a tint value (float) (default 0) | `{theme: 1, tint: -0.25}` |

#### `BORDER_STYLE` (string)

Border style is a string value which may take on one of the following values:

-   `dashDotDot`
-   `dashDot`
-   `dashed`
-   `dotted`
-   `hair`
-   `mediumDashDotDot`
-   `mediumDashDot`
-   `mediumDashed`
-   `medium`
-   `slantDashDot`
-   `thick`
-   `thin`

Borders for merged areas are specified for each cell within the merged area. So to apply a box border to a merged area of 3x3 cells, border styles would need to be specified for eight different cells:

-   left borders for the three cells on the left,
-   right borders for the cells on the right
-   top borders for the cells on the top
-   bottom borders for the cells on the left

## 🙏 Thanks

-   [sheetjs](https://github.com/SheetJS/sheetjs)
-   [js-xlsx](https://github.com/protobi/js-xlsx)
-   [sheetjs-style](https://www.npmjs.com/package/sheetjs-style)
-   [sheetjs-style-v2](https://www.npmjs.com/package/sheetjs-style-v2)

## 🔖 License

Please consult the attached LICENSE file for details. All rights not explicitly
granted by the Apache 2.0 License are reserved by the Original Author.
