const puppeteer = require('puppeteer');
const XLSX = require('xlsx');

const lastRow = 0;
const companies = [];
const companyEmails = [];
const companyPhones = [];

const workbook = XLSX.readFile('for_search.xlsx');
const worksheet1 = workbook.Sheets[workbook.SheetNames[0]];

async function search() {
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
  
    for (let i = 1; i < lastRow; i++) {
      try {
        const url = worksheet1[`D${i}`].v.toString();
        const name = worksheet1[`A${i}`].v.toString();
        console.log(`Fetching URL: ${url}`);
  
        await page.goto(url, { waitUntil: 'networkidle0' });
  
        let phoneFound = false; // Create a boolean variable to keep track of phone number
  
        const evaluatePromise = evaluateContacts(page, i);
  
        // Set a timer for 5 seconds to evaluate contacts
        const evaluateTimeout = new Promise((resolve) => {
          setTimeout(() => {
            resolve(false);
          }, 5000);
        });
  
        phoneFound = await Promise.race([evaluatePromise, evaluateTimeout]);
  
        companies[i] = {
          companyName: name,
          companyPhone: companyPhones[i],
          companyEmail: companyEmails[i]
        };
  
        console.log(`URL ${i} fetched successfully: ${url}\n`);
  
        if (phoneFound) {
          console.log('Phone number found. Stopping search.');
          break; // Stop searching for phone number
        }
      } catch (error) {
        console.error(`Error while fetching URL ${i + 1}: ${error.message}\n`);
      }
    }
  
    await browser.close();
  }

  async function evaluateContacts(page, i) {
    // Wait for the phone and email elements to load
    try {
      await page.waitForXPath(".//*[contains(text(), '+7') or contains(text(), 'tel:') or contains(@href, 'mailto:')]");
    
      // Find the email element on the page
      let emailElement = await page.$x(".//a[contains(@href, 'mailto:')]");
      if (emailElement.length === 0) {
        emailElement = await page.$x(".//*[contains(text(), '@')]");
      }
    
      if (emailElement.length > 0) {
        emailTest = await emailElement[0].evaluate((el) => el.textContent.trim());
        const formattedEmail = emailTest.slice(0, 30);
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;
        if (formattedEmail !== '' && emailRegex.test(formattedEmail)) {
            console.log(`${formattedEmail}`);
            companyEmails[i] = formattedEmail;
        }
      }
    
      // Find the phone element on the page
      let phoneElement = await page.$x(".//*[contains(@href, 'tel:')]");
      if (phoneElement.length === 0) {
        phoneElement = await page.$x(".//*[contains(text(), '+7') or contains(text(), '+ 7') or contains(text(), '8(') or contains(text(), '8 (') or contains(text(), '8-') or contains(text(), '8 ') or contains(text(), '88')]");
      }
    
      // Get the text content of the phone element
      if (phoneElement.length > 0) {
        const phoneTest = await phoneElement[0].evaluate((el) => el.textContent.trim());
        const formattedPhone = 'tel:' + phoneTest.slice(0, 30); // Cut phone number to maximum of 70 characters
        console.log(`tel:${formattedPhone}`);
        companyPhones[i] = formattedPhone;
      }
    } catch (error) {
      console.error(`Error while evaluating contacts: ${error.message}`);
    }
  }

  async function writeToExcel() {
    var newWorkbook = XLSX.utils.book_new();
    var newWorksheet = XLSX.utils.json_to_sheet(companies);
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Companies");
    XLSX.writeFile(newWorkbook, "готовый_без_телефона.xlsx");
}
async function main() {
    await search();
    await writeToExcel();
}

main();
