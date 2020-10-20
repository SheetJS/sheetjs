import xlsx from "xlsx";
import drive from "drive-db";

(async () => {
  const data = await drive("1fvz34wY6phWDJsuIneqvOoZRPfo6CfJyPg1BYgHt59k");

  /* Create a new workbook */
  const workbook = xlsx.utils.book_new();

  /* make worksheet */
  const worksheet = xlsx.utils.json_to_sheet(data);

  /* Add the worksheet to the workbook */
  xlsx.utils.book_append_sheet(workbook, worksheet);

  xlsx.writeFile(workbook, "test.xlsx");
})();
