import React, { useState, ChangeEvent } from "react";
import DataGrid, { TextEditor } from "react-data-grid";
import { read, utils, WorkSheet, writeFile } from "xlsx";

import "../styles/App.css";

type Row = any[]; /*{
  [index: string]: string | number;
};*/

type Column = {
  key: string;
  name: string;
  editor: typeof TextEditor;
};

type DataSet = {
  [index: string]: WorkSheet;
};

function getRowsCols(
  data: DataSet,
  sheetName: string
): {
  rows: Row[];
  columns: Column[];
} {
  const rows: Row[] = utils.sheet_to_json(data[sheetName], {header:1});
  let columns: Column[] = [];

  for (let row of rows) {
    const keys: string[] = Object.keys(row);

    if (keys.length > columns.length) {
      columns = keys.map((key) => {
        return { key, name: utils.encode_col(+key), editor: TextEditor };
      });
    }
  }

  return { rows, columns };
}

export default function App() {
  const [rows, setRows] = useState<Row[]>([]);
  const [columns, setColumns] = useState<Column[]>([]);
  const [workBook, setWorkBook] = useState<DataSet>({} as DataSet);
  const [sheets, setSheets] = useState<string[]>([]);
  const [current, setCurrent] = useState<string>("");

  const exportTypes = ["xlsx", "xlsb", "csv", "html"];

  function selectSheet(name: string, reset = true) {
    if(reset) workBook[current] = utils.json_to_sheet(rows, {
      header: columns.map((col: Column) => col.key),
      skipHeader: true
    });

    const { rows: new_rows, columns: new_columns } = getRowsCols(workBook, name);

    setRows(new_rows);
    setColumns(new_columns);
    setCurrent(name);
  }

  async function handleFile(ev: ChangeEvent<HTMLInputElement>): Promise<void> {
    const file = await ev.target.files?.[0]?.arrayBuffer();
    const data = read(file);

    setWorkBook(data.Sheets);
    setSheets(data.SheetNames);
  }

  function saveFile(ext: string): void {
    const wb = utils.book_new();

    sheets.forEach((n) => {
      utils.book_append_sheet(wb, workBook[n], n);
    });

    writeFile(wb, "sheet." + ext);
  }

  return (
    <>
      <input type="file" onChange={handleFile} />
      <div className="flex-cont">
        {sheets.map((sheet) => (
          <button key={sheet} onClick={(e) => selectSheet(sheet)}>
            {sheet}
          </button>
        ))}
      </div>
      <div className="flex-cont">
        <b>Current Sheet: {current}</b>
      </div>
      <DataGrid columns={columns} rows={rows} onRowsChange={setRows} />
      <div className="flex-cont">
        {exportTypes.map((ext) => (
          <button key={ext} onClick={() => saveFile(ext)}>
            export [.{ext}]
          </button>
        ))}
      </div>
    </>
  );
}
