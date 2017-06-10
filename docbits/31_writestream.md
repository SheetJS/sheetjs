### Streaming Write

The streaming write functions are available in the `XLSX.stream` object.  They
take the same arguments as the normal write functions but return a Readable
Stream.  They are only exposed in NodeJS.

- `XLSX.stream.to_csv` is the streaming version of `XLSX.utils.sheet_to_csv`.
- `XLSX.stream.to_html` is the streaming version of `XLSX.utils.sheet_to_html`.

<details>
	<summary><b>nodejs convert to CSV and write file</b> (click to show)</summary>

```js
var output_file_name = "out.csv";
var stream = XLSX.stream.to_csv(worksheet);
stream.pipe(fs.createWriteStream(output_file_name));
```

</details>

<https://github.com/sheetjs/sheetaki> pipes write streams to nodejs response.

