const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');
const { url } = require('inspector');
const readline = require('readline').createInterface({
  input: process.stdin,
  output: process.stdout
});

let email = '';
let phone = '';

readline.question('URL: ', url =>
(async () => {
    const browser = await puppeteer.launch({
      headless: false // Launch in headless mode
    });
    const page = await browser.newPage();
      
      await page.goto(url);
  
      // Find the link element containing the desired text and click on it
      await page.evaluate(() => {
        const link = Array.from(document.querySelectorAll('a')).find(el => {
          const text = el.innerText.toLowerCase();
          return text.includes('контакты' || 'о компании' || 'о нас' || 'контактная информация');
        });
        link ? link.click() : console.log('no contact page')
      });
  
      // Wait for the new page to load
      await page.waitForNavigation();
  
      // Do something with the new page
      evaluateContacts(page)
      console.log(page.url());
  
    //await browser.close();
  })());

  async function evaluateContacts(page) {
    // Wait for the phone and email elements to load
  await page.waitForXPath(".//*[contains(text(), '+7') or contains(text(), 'tel:') or contains(@href, 'mailto:')]");

  // Find the email element on the page
  let emailElement = await page.$x(".//a[contains(@href, 'mailto:')]");
  if (emailElement.length === 0) {
    emailElement = await page.$x(".//*[contains(text(), '@')]");
  }

  if (emailElement.length > 0) {
    emailTest = await emailElement[0].evaluate((el) => el.textContent.trim());
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;
        if (emailTest !== '' && emailRegex.test(emailTest)) {
            email = emailTest;
        }
  }


    console.log(`Email: ${email}`);

  // Find the phone element on the page
  let phoneElement = await page.$x(".//*[contains(@href, 'tel:')]");
  if (phoneElement.length === 0) {
    phoneElement = await page.$x(".//*[contains(text(), '+7') or contains(text(), '+ 7')]");
  }

  // Get the text content of the phone element
  if (phoneElement.length > 0) {
    phone = await phoneElement[0].evaluate((el) => el.textContent.trim());
  }

  console.log(`Phone: ${phone}`);
}