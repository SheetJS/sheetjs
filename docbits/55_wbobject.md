### Workbook Object

`workbook.SheetNames` is an ordered list of the sheets in the workbook

`wb.Sheets[sheetname]` returns an object representing the worksheet.

`wb.Props` is an object storing the standard properties.  `wb.Custprops` stores
custom properties.  Since the XLS standard properties deviate from the XLSX
standard, XLS parsing stores core properties in both places.

`wb.WBProps` includes more workbook-level properties:

- Excel supports two epochs (January 1 1900 and January 1 1904), see
  [1900 vs. 1904 Date System](http://support2.microsoft.com/kb/180162).
  The workbook's epoch can be determined by examining the workbook's
  `wb.WBProps.date1904` property.

