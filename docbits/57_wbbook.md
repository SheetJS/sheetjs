### Workbook-Level Attributes

`wb.Workbook` stores workbook level attributes.

#### Defined Names

`wb.Workbook.Names` is an array of defined name objects which have the keys:

| Key       | Description                                                      |
|:----------|:-----------------------------------------------------------------|
| `Sheet`   | Name scope.  Sheet Index (0 = first sheet) or `null` (Workbook)  |
| `Name`    | Case-sensitive name.  Standard rules apply **                    |
| `Ref`     | A1-style Reference (e.g. `"Sheet1!$A$1:$D$20"`)                  |
| `Comment` | Comment (only applicable for XLS/XLSX/XLSB)                      |

Excel allows two sheet-scoped defined names to share the same name.  However, a
sheet-scoped name cannot collide with a workbook-scope name.  Workbook writers
may not enforce this constraint.

