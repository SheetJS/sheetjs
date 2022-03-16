<script setup lang="ts">
import { ref } from "vue";
import { read, utils, writeFile, WorkBook } from "xlsx";

import VueTableLite from "vue3-table-lite/ts";

type DataSet = {
  [index: string]: WorkBook;
};

type Row = any[];

type Column = {
  field: string;
  label: string;
  display: (row: Row) => string;
};

const currFileName = ref<string>("");
const currSheet = ref<string>("");
const sheets = ref<string[]>([]);
const workBook = ref<DataSet>({} as DataSet);
const rows = ref<Row[]>([]);
const columns = ref<Column[]>([]);

const exportTypes: string[] = ["xlsx", "xlsb", "csv", "html"];

let cell = 0;

function resetCell() {
  cell = 0;
}

function display(col: number): (row: Row) => string {
  return function (row: Row) {
    return `<span
               style="user-select: none; display: block"
               position="${Math.floor(cell++ / columns.value.length)}.${col}"
               onblur="endEdit(event)"
               ondblclick="startEdit(event)"
               onkeydown="endEdit(event)">${row[col] ?? "&nbsp;"}</span>`;
  };
}

window.startEdit = function (ev) {
  ev.target.contentEditable = true;
  ev.target.focus();
};

window.endEdit = function (ev) {
  if (ev.key === undefined || ev.key === "Enter") {
    const pos = ev.target.getAttribute("position").split(".");

    ev.target.contentEditable = false;

    rows.value[pos[0]][pos[1]] = ev.target.innerText;

    workBook.value[currSheet.value] = utils.json_to_sheet(rows.value, {
      header: columns.value.map((col: Column) => col.field),
      skipHeader: true,
    });
  }
};

function getRowsCols(
  data: DataSet,
  sheetName: string
): {
  rows: Row[];
  cols: Column[];
} {
  const rows: Row[] = utils.sheet_to_json(data[sheetName], { header: 1 });
  let cols: Column[] = [];

  for (let row of rows) {
    const keys: string[] = Object.keys(row);

    if (keys.length > cols.length) {
      cols = keys.map((key) => {
        return {
          field: key,
          label: utils.encode_col(+key),
          display: display(key),
        };
      });
    }
  }

  return { rows, cols };
}

async function importFile(ev: ChangeEvent<HTMLInputElement>): Promise<void> {
  const file = ev.target.files[0];
  const data = read(await file.arrayBuffer());

  currFileName.value = file.name;
  currSheet.value = data.SheetNames?.[0];
  sheets.value = data.SheetNames;
  workBook.value = data.Sheets;

  selectSheet(currSheet.value);
}

function exportFile(type: string): void {
  const wb = utils.book_new();

  sheets.value.forEach((sheet) => {
    utils.book_append_sheet(wb, workBook.value[sheet], sheet);
  });

  writeFile(wb, `sheet.${type}`);
}

function selectSheet(sheet: string): void {
  const { rows: newRows, cols: newCols } = getRowsCols(workBook.value, sheet);

  resetCell();

  rows.value = newRows;
  columns.value = newCols;
  currSheet.value = sheet;
}
</script>

<template>
  <header class="imp-exp">
    <div class="import">
      <input type="file" id="import" @change="importFile" />
      <label for="import">import</label>
    </div>
    <span>{{ currFileName || "vue-modify demo" }}</span>
    <div class="export">
      <span>export</span>
      <ul>
        <li v-for="type in exportTypes" @click="exportFile(type)">
          {{ `.${type}` }}
        </li>
      </ul>
    </div>
  </header>
  <div class="sheets">
    <span
      v-for="sheet in sheets"
      @click="selectSheet(sheet)"
      :class="[currSheet === sheet ? 'selected' : '']"
    >
      {{ sheet }}
    </span>
  </div>
  <vue-table-lite
    :is-static-mode="true"
    :page-size="50"
    :columns="columns"
    :rows="rows"
  ></vue-table-lite>
</template>

<style>
.imp-exp {
  display: flex;
  justify-content: space-between;
  padding: 0.5rem;
  font-family: mono;
  color: #212529;
}

.import {
  font-size: medium;
}

.import input {
  position: absolute;
  opacity: 0;
  cursor: pointer;
}

.import label {
  background-color: white;
  border: 1px solid;
  padding: 0.3rem;
}

.export: hover {
  border-bottom: none;
}

.export:hover ul {
  display: block;
}

.export span {
  padding: 0.3rem;
  border: 1px solid;
  cursor: pointer;
}

.export ul {
  display: none;
  position: absolute;
  z-index: 5;
  background-color: white;
  list-style: none;
  padding: 0.3rem;
  border: 1px solid;
  margin-top: 0.3rem;
  border-top: none;
}

.export ul li {
  padding: 0.3rem;
  text-align: center;
}

.export ul li:hover {
  background-color: lightgray;
  cursor: pointer;
}

.sheets {
  display: flex;
  justify-content: center;
  margin: 0.3rem;
  color: #212529;
}

.sheets span {
  border: 1px solid;
  padding: 0.5rem;
  margin: 0.3rem;
}

.sheets span:hover:not(.selected) {
  background-color: lightgray;
  cursor: pointer;
}

.selected {
  background-color: #343a40;
  color: white;
}
</style>
