const puppeteer = require('puppeteer');

(async () => {

  const browser = await puppeteer.launch();
  const page = await browser.newPage();
  await page.goto('http://localhost:8000', {waitUntil: 'load'});
	await page.waitFor(30*1000);
  await page.pdf({path: 'test.pdf', format: 'A4'});

  browser.close();
})();

