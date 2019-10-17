const puppeteer = require("puppeteer");
const path = require("path");

const target = `file://${path.resolve(`${__dirname}/standalone.html`)}`;
console.log(target);
(async () => {
  const browser = await puppeteer.launch();
  const page = await browser.newPage();
  page.on('console', msg => console.log('PAGE LOG:', msg.text()));
  page.on('error', (err) => { console.error(err); process.exit(1); });
  page.on('pageerror', (err) => { console.error(err); process.exit(2); });
  await page.goto(target);
  await browser.close();
})();