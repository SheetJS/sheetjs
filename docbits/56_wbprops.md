#### Workbook File Properties

The various file formats use different internal names for file properties.  The
workbook `Props` object normalizes the names:

<details>
  <summary><b>File Properties</b> (click to show)</summary>

| JS Name       | Excel Description              |
|:--------------|:-------------------------------|
| `Title`       | Summary tab "Title"            |
| `Subject`     | Summary tab "Subject"          |
| `Author`      | Summary tab "Author"           |
| `Manager`     | Summary tab "Manager"          |
| `Company`     | Summary tab "Company"          |
| `Category`    | Summary tab "Category"         |
| `Keywords`    | Summary tab "Keywords"         |
| `Comments`    | Summary tab "Comments"         |
| `LastAuthor`  | Statistics tab "Last saved by" |
| `CreatedDate` | Statistics tab "Created"       |

</details>

For example, to set the workbook title property:

```js
if(!wb.Props) wb.Props = {};
wb.Props.Title = "Insert Title Here";
```

Custom properties are added in the workbook `Custprops` object:

```js
if(!wb.Custprops) wb.Custprops = {};
wb.Custprops["Custom Property"] = "Custom Value";
```

Writers will process the `Props` key of the options object:

```js
/* force the Author to be "SheetJS" */
XLSX.write(wb, {Props:{Author:"SheetJS"}});
```

