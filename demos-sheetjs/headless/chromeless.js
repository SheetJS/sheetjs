const { Chromeless } = require('chromeless');
const TEST = 'http://localhost:8000', TIME = 30 * 1000;
(async() => {
  const browser = new Chromeless();
  const pth = await browser.goto(TEST).wait(TIME).screenshot();
  console.log(pth);
  await browser.end();
})().catch(e=>{ console.error(e); });

