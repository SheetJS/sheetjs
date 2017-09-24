/* xlsx.js (C) 2013-present  SheetJS -- http://sheetjs.com */
const puppeteer = require('puppeteer');

(async () => {

  const browser = await puppeteer.launch();
  const page = await browser.newPage();
  await page.goto('http://oss.sheetjs.com/js-xlsx/tests/', {waitUntil: 'load'});
	await page.waitFor(30*1000);
  await page.pdf({path: 'test.pdf', format: 'A4'});

  browser.close();
})();

