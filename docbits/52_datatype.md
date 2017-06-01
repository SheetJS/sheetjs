#### Data Types

The raw value is stored in the `v` field, interpreted based on the `t` field.

Type `b` is the Boolean type.  `v` is interpreted according to JS truth tables.

Type `e` is the Error type. `v` holds the number and `w` holds the common name:

<details>
	<summary><b>Error values and interpretation</b> (click to show)</summary>

|  Value | Error Meaning   |
| -----: | :-------------- |
| `0x00` | `#NULL!`        |
| `0x07` | `#DIV/0!`       |
| `0x0F` | `#VALUE!`       |
| `0x17` | `#REF!`         |
| `0x1D` | `#NAME?`        |
| `0x24` | `#NUM!`         |
| `0x2A` | `#N/A`          |
| `0x2B` | `#GETTING_DATA` |

</details>

Type `n` is the Number type. This includes all forms of data that Excel stores
as numbers, such as dates/times and Boolean fields.  Excel exclusively uses data
that can be fit in an IEEE754 floating point number, just like JS Number, so the
`v` field holds the raw number.  The `w` field holds formatted text.  Dates are
stored as numbers by default and converted with `XLSX.SSF.parse_date_code`.

Type `d` is the Date type, generated only when the option `cellDates` is passed.
Since JSON does not have a natural Date type, parsers are generally expected to
store ISO 8601 Date strings like you would get from `date.toISOString()`.  On
the other hand, writers and exporters should be able to handle date strings and
JS Date objects.  Note that Excel disregards timezone modifiers and treats all
dates in the local timezone.  js-xlsx does not correct for this error.

Type `s` is the String type.  `v` should be explicitly stored as a string to
avoid possible confusion.

Type `z` represents blank stub cells.  These do not have any data or type, and
are not processed by any of the core library functions.  By default these cells
will not be generated; the parser `sheetStubs` option must be set to `true`.

#### Dates

<details>
	<summary><b>Excel Date Code details</b> (click to show)</summary>

By default, Excel stores dates as numbers with a format code that specifies date
processing.  For example, the date `19-Feb-17` is stored as the number `42785`
with a number format of `d-mmm-yy`.  The `SSF` module understands number formats
and performs the appropriate conversion.

XLSX also supports a special date type `d` where the data is an ISO 8601 date
string.  The formatter converts the date back to a number.

The default behavior for all parsers is to generate number cells.  Setting
`cellDates` to true will force the generators to store dates.

</details>

<details>
	<summary><b>Time Zones and Dates</b> (click to show)</summary>

Excel has no native concept of universal time.  All times are specified in the
local time zone.  Excel limitations prevent specifying true absolute dates.

Following Excel, this library treats all dates as relative to local time zone.

</details>

<details>
	<summary><b>Epochs: 1900 and 1904</b> (click to show)</summary>

Excel supports two epochs (January 1 1900 and January 1 1904), see
["1900 vs. 1904 Date System" article](http://support2.microsoft.com/kb/180162).
The workbook's epoch can be determined by examining the workbook's
`wb.Workbook.WBProps.date1904` property:

```js
!!(((wb.Workbook||{}).WBProps||{}).date1904)
```

</details>

