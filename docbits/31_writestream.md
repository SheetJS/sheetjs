### Streaming Write

The streaming write functions are available in the `XLSX.stream` object.  They
take the same arguments as the normal write functions but return a readable
stream.  They are only exposed in node.

- `XLSX.stream.to_csv` is the streaming version of `XLSX.utils.sheet_to_csv`.
- `XLSX.stream.to_html` is the streaming version of the HTML output type.

<https://github.com/sheetjs/sheetaki> pipes write streams to nodejs response.

