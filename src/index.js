const puppeteer = require('puppeteer');
const async = require("async");
const readXlsxFile = require('read-excel-file/node');

const names = [];

function delay(time) {
  return new Promise(function(resolve) { 
      setTimeout(resolve, time)
  });
}

readXlsxFile('./Contacts.xlsx').then((rows) => {
  rows.forEach(el => names.push(el[0]));
  (async () => {
    const browser = await puppeteer.launch({ headless: false });
    const page = await browser.newPage();
    await page.goto("https://web.whatsapp.com/");
    const message = "Last Test yaraaab  😂‍";
  
    async.eachSeries(names, async name => {
      await page.waitForSelector("#side .selectable-text");
      const search = await page.$("#side .selectable-text");
      await page.$eval("#side .selectable-text", el => el.value = "");
      await search.type(name);
  
      await page.waitForSelector("span [title='"+ name +"']");
      const person = await page.$("span [title='"+ name +"']");
      await person.click();
  
      await page.waitForSelector("#main .selectable-text");
      const input = await page.$("#main .selectable-text");
      await input.type(message);
      await page.keyboard.press("Enter");
      await delay(5000);
    });
  
    
  
  })();
});

