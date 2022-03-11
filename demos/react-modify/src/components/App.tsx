import React, { useState, ChangeEvent } from "react";
import DataGrid, { TextEditor } from "react-data-grid";
import { read, utils, writeFile, WorkBook } from "xlsx";

import "../styles/App.css";

type Row = {
  [index: string]: string | number;
};

type Column = {
  key: string;
  name: string;
  editor: typeof TextEditor;
};

function getRowsCols(
  data: WorkBook,
  sheetName: string
): {
  rows: Row[];
  columns: Column[];
} {
  const rows: Row[] = utils.sheet_to_json(data.Sheets[sheetName]);
  let columns: Column[] = [];

  for (let row of rows) {
    const keys: string[] = Object.keys(row);

    if (keys.length > columns.length) {
      columns = keys.map((key) => {
        return { key, name: key, editor: TextEditor };
      });
    }
  }

  return { rows, columns };
}

export default function App() {
  const [rows, setRows] = useState<Row[]>([]);
  const [columns, setColumns] = useState<Column[]>([]);
  const [workBook, setWorkBook] = useState<WorkBook>({} as WorkBook);
  const [sheets, setSheets] = useState<string[]>([]);

  const exportTypes = ["xlsx", "xlsb", "csv", "html"];

  function selectSheet(name: string) {
    const { rows, columns } = getRowsCols(workBook, name);

    setRows(rows);
    setColumns(columns);
  }

  async function handelFile(ev: ChangeEvent<HTMLInputElement>): Promise<void> {
    const file = await ev.target.files?.[0]?.arrayBuffer();
    const data = read(file);

    setWorkBook(data);
    setSheets(data.SheetNames);
  }

  function saveFile(ext: string): void {
    const wb = utils.book_new();
    const sheet = utils.json_to_sheet(rows, {
      header: columns.map((col: Column) => col.key),
    });

    utils.book_append_sheet(wb, sheet, "sheet");
    writeFile(wb, "sheet." + ext);
  }

  return (
    <>
      <input type="file" onChange={handelFile} />
      <div className="flex-cont">
        {sheets.map((sheet) => (
          <button key={sheet} onClick={(e) => selectSheet(sheet)}>
            {sheet}
          </button>
        ))}
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
