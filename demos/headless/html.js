#!/usr/bin/env node
/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
const puppeteer = require("puppeteer");
const path = require("path");
const fs = require("fs");

/* inf is the path to the html file -> url is a file URL */
let inf = process.argv[2] || "test.html";
let htmlpath = path.join(__dirname, inf);
if(!fs.existsSync(htmlpath)) htmlpath =  path.join(process.cwd(), inf);
if(!fs.existsSync(htmlpath)) htmlpath =  path.resolve(inf);
if(!fs.existsSync(htmlpath)) { console.error(`Could not find a valid file for \`${inf}\``); process.exit(4); }
console.error(`Reading from ${htmlpath}`);
const url = `file://${htmlpath}`;

/* get the standalone build source (e.g. node_modules/xlsx/dist/xlsx.full.min.js) */
// const websrc = fs.readFileSync(require.resolve("xlsx/dist/xlsx.full.min.js"), "utf8");
const get_lib = (jspath) => fs.readFileSync(path.resolve(__dirname, jspath)).toString();
const websrc = get_lib("xlsx.full.min.js");

(async() => {
  /* start browser and go to web page */
  const browser = await puppeteer.launch();
  const page = await browser.newPage();
  page.on("console", msg => console.log("PAGE LOG:", msg.text()));
  await page.setViewport({width: 1920, height: 1080});
  await page.goto(url, {waitUntil: "networkidle2"});

  /* inject library */
  await page.addScriptTag({content: websrc});

  /* this function `s5s` will be called by the script below, receiving the Base64-encoded file */
  await page.exposeFunction("s5s", async(b64) => {
    fs.writeFileSync("output.xlsb", b64, {encoding: "base64"});
  });

  /* generate XLSB file in webpage context and send back a Base64-encoded string */
  await page.addScriptTag({content: `
    /* call table_to_book on first table */
    var wb = XLSX.utils.table_to_book(document.getElementsByTagName("TABLE")[0]);

    /* generate XLSB file */
    var b64 = XLSX.write(wb, {type: "base64", bookType: "xlsb"});

    /* call "s5s" hook exposed from the node process */
    window.s5s(b64);
  `});

  /* cleanup */
  await browser.close();
})();
