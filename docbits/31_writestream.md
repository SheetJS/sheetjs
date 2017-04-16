### Streaming Write

`XLSX.stream.to_csv` is the streaming version of `XLSX.utils.sheet_to_csv`.  It
takes the same arguments but returns a readable stream.

<https://github.com/sheetjs/sheetaki> pipes CSV write stream to nodejs response.
