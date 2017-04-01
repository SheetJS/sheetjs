#### Worksheet Object

In addition to the base sheet keys, worksheets also add:

- `ws['!cols']`: array of column properties objects.  Column widths are actually
  stored in files in a normalized manner, measured in terms of the "Maximum
  Digit Width" (the largest width of the rendered digits 0-9, in pixels).  When
  parsed, the column objects store the pixel width in the `wpx` field, character
  width in the `wch` field, and the maximum digit width in the `MDW` field.

- `ws['!merges']`: array of range objects corresponding to the merged cells in
  the worksheet.  Plaintext utilities are unaware of merge cells.  CSV export
  will write all cells in the merge range if they exist, so be sure that only
  the first cell (upper-left) in the range is set.

#### Chartsheet Object

Chartsheets are represented as standard sheets.  They are distinguished with the
`!type` property set to `"chart"`.

The underlying data and `!ref` refer to the cached data in the chartsheet.  The
first row of the chartsheet is the underlying header.

