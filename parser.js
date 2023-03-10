const cheerio = require('cheerio');
const XLSX = require('xlsx');

// Used URLs
// https://hh.ru/search/vacancy?no_magic=true&L_save_area=true&text=%D0%93%D0%B5%D0%BD%D0%B5%D1%80%D0%B0%D0%BB%D1%8C%D0%BD%D1%8B%D0%B9+%D0%B4%D0%B8%D1%80%D0%B5%D0%BA%D1%82%D0%BE%D1%80&search_field=name&excluded_text=&salary=&currency_code=RUR&experience=doesNotMatter&label=not_from_agency&order_by=relevance&search_period=0&items_on_page=20
// https://hh.ru/search/vacancy?search_field=name&enable_snippets=true&text=%D0%B7%D0%B0%D0%BC%D0%B5%D1%81%D1%82%D0%B8%D1%82%D0%B5%D0%BB%D1%8C+%D0%B3%D0%B5%D0%BD%D0%B5%D1%80%D0%B0%D0%BB%D1%8C%D0%BD%D0%BE%D0%B3%D0%BE+%D0%B4%D0%B8%D1%80%D0%B5%D0%BA%D1%82%D0%BE%D1%80%D0%B0&no_magic=true&L_save_area=true&items_on_page=20
// https://hh.ru/search/vacancy?no_magic=true&L_save_area=true&text=%D0%BE%D0%BF%D0%B5%D1%80%D0%B0%D1%86%D0%B8%D0%BE%D0%BD%D0%BD%D1%8B%D0%B9+%D0%B4%D0%B8%D1%80%D0%B5%D0%BA%D1%82%D0%BE%D1%80&search_field=name&excluded_text=&salary=&currency_code=RUR&experience=doesNotMatter&label=not_from_agency&order_by=relevance&search_period=0&items_on_page=20
// https://hh.ru/search/vacancy?no_magic=true&L_save_area=true&text=%D0%97%D0%B0%D0%BC%D0%B5%D1%81%D1%82%D0%B8%D1%82%D0%B5%D0%BB%D1%8C+%D0%BE%D0%BF%D0%B5%D1%80%D0%B0%D1%86%D0%B8%D0%BE%D0%BD%D0%BD%D0%BE%D0%B3%D0%BE+%D0%B4%D0%B8%D1%80%D0%B5%D0%BA%D1%82%D0%BE%D1%80%D0%B0&search_field=name&excluded_text=&salary=&currency_code=RUR&experience=doesNotMatter&order_by=relevance&search_period=0&items_on_page=20

const url = 'https://hh.ru/search/vacancy?label=not_from_agency&search_field=description&enable_snippets=true&text=%D0%94%D0%B8%D1%80%D0%B5%D0%BA%D1%82%D0%BE%D1%80+%D0%BF%D0%BE+%D0%A1%D1%82%D1%80%D0%BE%D0%B8%D1%82%D0%B5%D0%BB%D1%8C%D1%81%D1%82%D0%B2%D1%83&page=0&disableBrowserCache=true&items_on_page=20';
const companies = [];
const companyNamesArray = [];
const vacancyNamesArray = [];
const vacancyLinksArray = [];
const companyLinksArray = [];
const companyVacaniesNumber = [];
const companyWebsites = [];
const companyPhones = [];
const companyEmails = [];
const XLSXcompanies = [];


// Get number of pages returned for request
async function findPagesNumber() {
    
    const response = await fetch(url);
    const html = await response.text();
    const $ = cheerio.load(html);

    const endPages = $(".pager a").last();
    const secondLastSpan = endPages.prev('span');
    const allPages = parseInt(secondLastSpan.find('a[data-qa="pager-page"]').text());
    
    console.log('Number of pages: ', allPages);

    return allPages;
}

// Fetch data from url
async function fetchData(url) {
    
    const response = await fetch(url);
    const html = await response.text();
    const $ = cheerio.load(html);

    const div = $('.vacancy-serp-content');

    const companyNames = div.find('.bloko-link_kind-tertiary');
    const vacancyNames =div.find('.serp-item__title');
    const vacancyLinks = $('a.serp-item__title');
    companyLinks = $('a.bloko-link_kind-tertiary');

    companyNames.each(function(i, title) {
      companyNamesArray.push($(title).text());
    })
    
    vacancyLinks.each(function(i, title) {
        vacancyLinksArray.push($(title).attr('href'));
      })

    companyLinks.each(function(i, title) {
        companyLinksArray.push('https://hh.ru' + $(title).attr('href'));
    })
    
    vacancyNames.each(function(i, title) {
        vacancyNamesArray.push( $(title).text());
    })

  }

  // Fetch single company profile page
  async function fetchCompanyData(url, i) {

    const response = await fetch(url);
    const html = await response.text();
    const $ = cheerio.load(html);
  
    // Get number of vacancies
    try {
        const span = $('a[data-qa="employer-page__employer-vacancies-link"] span');
        const text = span.text();
        const matches = text.match(/\d+/);
        const number = matches[0];

        companyVacaniesNumber[i] = number;
    } catch (error) {
        console.log("Vacancy number element not found.");
        console.log(error);

        companyVacaniesNumber[i] = null;
    }

    try {
        const siteEl = $('a[data-qa="sidebar-company-site"]');
        const site = siteEl.attr('href');

        companyWebsites[i] = site;
    } catch (error) {
        console.log("Company website element not found.");
        console.log(error);

        companyWebsites[i] = null;
    }
  
  }

  // Fetch company website and get phone number with email
  async function fetchCompanyWebsite(url, i) {

    try {
      const response = await fetch(url);
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }
  
      const html = await response.text();
      const $ = cheerio.load(html);
  
      const phoneLink = $('a[href^="tel:"]:first').attr('href');
      const emailLink = $('a[href^="mailto:"]:first').attr('href');

      const email = emailLink ? $('a[href^="mailto:"]:first').text() : null;
  
      companyPhones[i] = phoneLink;
      companyEmails[i] = email;
    } catch (error) {
      console.error(error);
    }

  }

  // Fetch all pages
  async function fetchAllUrls(allPages) {

    // Get first data about companies
    for (i = 75; i < 100; i++) {
        const pageUrl = (url + `&page=${i}`);
        console.log('Fetching page:', pageUrl);

        const data = await fetchData(pageUrl);
        console.log(`Fetched page ${i}, fetched data: `, companyNamesArray, vacancyLinksArray, companyLinksArray, vacancyNamesArray);

        await new Promise(resolve => setTimeout(resolve, 1000));
    }

    // Get number of vacancies and company website
    for (i = 0; i < companyLinksArray.length; i++) {
        const companyUrl = companyLinksArray[i];
        console.log('Fetching page:', companyUrl);

        const data = await fetchCompanyData(companyUrl, i);
        console.log(`Fetched company ${i}, fetched data: `, companyVacaniesNumber[i], companyWebsites[i]);

    }

    // Get  phone numbers and Emails
    for (i = 0; i < companyNamesArray.length; i ++) {
        const companyWebsite = companyWebsites[i];

        if (companyWebsite) {
            console.log('Fetching website:', companyWebsite);
            const data = await fetchCompanyWebsite(companyWebsite, i);
            console.log(`Fetched website ${i}, fetched data: `, companyPhones[i], companyEmails[i]);
        } else {
            console.log(`No website found for company ${i}.`);
        }

    }

    // Push company card object to companies
    for (i = 0; i < companyNamesArray.length; i++) {
        companies[i] = ({
            companyName: companyNamesArray[i],
            companyLink: companyLinksArray[i],
            vacanciesNumber: companyVacaniesNumber[i],
            companyWebsite: companyWebsites[i],
            companyPhone: companyPhones[i],
            companyEmail: companyEmails[i],
            vacancyName: vacancyNamesArray[i],
            vacancyLink: vacancyLinksArray[i]
        });
    }
      
    console.log('Company object example:', companies[5]);
  }

  function prepareForExcel() {

    const replacements = [
      { search: "?????????????????????? ????????????????", replacement: "?????????????????????? ????????????????" },
      { search: "???????????????????? ????????????????", replacement: "???????????????????? ????????????????" },
      { search: "?????????????????????? ???????????????????????? ??????????????????", replacement: "?????????????????????? ???????????????????????? ??????????????????" },
      { search: "?????????????????????? ?????????????????????? ??????????????????", replacement: "?????????????????????? ?????????????????????? ??????????????????" },
      { search: "???????????????? ?????????????????????????? ??????????", replacement: "???????????????????????? ????????????????" },
      { search: "???????????????????????? ????????????????", replacement: "???????????????????????? ????????????????" },
      { search: "?????????????????????? ?????????????????????????? ??????????????????", replacement: "?????????????????????? ?????????????????????????? ??????????????????" },
      { search: "???????????????????????? ??????????????????????", replacement: "???????????????????????? ????????????????" },
      { search: "??????????????????-???????????????????????? ????????????????", replacement: "???????????????????????? ????????????????" },
      { search: "???????????????????????????? ????????????????", replacement: "???????????????????????????? ????????????????" },
      { search: "?????????????????????? ?????????????????????????????? ??????????????????", replacement: "?????????????????????? ?????????????????????????????? ??????????????????" },
      { search: "?????????????????????? ????????????????????????/?????????????????????????????? ??????????????????", replacement: "?????????????????????? ?????????????????????????????? ??????????????????" },
      { search: "???????????????? ???? ?????????????????????????? ????????????????????", replacement: "???????????????? ???? ????????????????????" },
      { search: "???????????????????????? ????????????????", replacement: "???????????????????????? ????????????????" },
      { search: "?????????????????????? ?????????????????????????? ??????????????????", replacement: "?????????????????????? ?????????????????????????? ??????????????????" },
      { search: "???????????????? ???? ????????????????", replacement: "???????????????? ???? ????????????????" },
      { search: "???????????????????????? ???????????? ???? ????????????????", replacement: "???????????????? ???? ????????????????" },
      { search: "???????????????? ???? ????????????????????", replacement: "???????????????? ???? ????????????????????" },
      { search: "???????????????????????? ???????????? ????????????????????", replacement: "???????????????? ???? ????????????????????" },
      { search: "???????????????????????? ?? ?????????? ????????????????????", replacement: "???????????????? ???? ????????????????????" },
      { search: "???????????????????????? ????????????????????", replacement: "???????????????? ???? ????????????????????" },
      { search: "?????????????????? ???????????? ????????????????????", replacement: "???????????????? ???? ????????????????????" },
      { search: "???????????????????????? ???????????? ????????????????????", replacement: "???????????????? ???? ????????????????????" },
      { search: "???????????????????????? ???????????????????? ????????????????????", replacement: "???????????????? ???? ????????????????????" },
      { search: "?????????????????? ???????????? ????????????", replacement: "???????????????? ???? ????????????????????" },
      { search: "?????????????????????? ?????????????????? ???? ????????????????????", replacement: "?????????????????????? ?????????????????? ???? ????????????????????" },
      { search: "?????????????????????? ????????????????", replacement: "?????????????????????? ????????????????" },
      { search: "?????????????????????? ???????????????????????? ??????????????????", replacement: "?????????????????????? ???????????????????????? ??????????????????" },
      { search: "cto", replacement: "?????????????????????? ????????????????" },
      { search: "?????????????????????????????? ????????????????", replacement: "?????????????????????????????? ????????????????" },
      { search: "???????????????????????? ??????????????????????", replacement: "?????????????????????????????? ????????????????" },
      { search: "???????????????????????? ????????????????", replacement: "?????????????????????????????? ????????????????" },
      { search: "???????????????? ???? ????????????????????????", replacement: "?????????????????????? ????????????????" },
      { search: "???????????????????????? ???????????? ???? ???????????? ?? ????????????????????", replacement: "???????????????? ???? ??????????????????" },
      { search: "hr ????????????????", replacement: "???????????????? ???? ??????????????????" },
      { search: "hr-????????????????", replacement: "???????????????? ???? ??????????????????" },
      { search: "???????????????? ???? ??????????????????", replacement: "???????????????? ???? ??????????????????" },
      { search: "???????????????????????? ???? ??????????????????", replacement: "???????????????? ???? ??????????????????" },
      { search: "?????????????????????? ?????????????????? ???? ??????????????????", replacement: "?????????????????????? ?????????????????? ???? ??????????????????" },
      { search: "?????????????????? ???????????????????? ???? ???????????? ?? ????????????????????", replacement: "???????????????? ???? ??????????????????" },
      { search: "?????????????????? ???????????? ???? ???????????? ?? ????????????????????", replacement: "???????????????? ???? ??????????????????" },
      { search: "???????????????????????? ???????????? ??????????????????", replacement: "???????????????? ???? ??????????????????" },
      { search: "hrd", replacement: "???????????????? ???? ??????????????????" },
      { search: "???????????????? ???? ???????????????????? ????????????????????", replacement: "???????????????? ???? ??????????????????" },
      { search: "hr generalist", replacement: "???????????????? ???? ??????????????????" },
      { search: "???????????????? ???? ????????????????", replacement: "???????????????? ???? ????????????????" },
      { search: "?????????????????????? ?????????????????? ???? ????????????????", replacement: "?????????????????????? ?????????????????? ???? ????????????????" },
      { search: "???????????????????????? ???????????? ????????????", replacement: "???????????????? ???? ????????????????" },
      { search: "???????????????? ??????????????", replacement: "?????????????????????????????? ????????????????" },
      { search: "???????????????????????? ??????????????", replacement: "?????????????????????????????? ????????????????" },
      { search: "???????????????????????? ?????????????????????????? ??????????????????????????", replacement: "?????????????????????????????? ????????????????" },
      { search: "???????????????????????? ?????????????????????????? ??????????????", replacement: "?????????????????????????????? ????????????????" },
      { search: "???????????????????????? ???????????? ???????????????? ????????????????", replacement: "???????????????? ???? ????????????????" },
      { search: "???????????????? ???? ????????????????", replacement: "???????????????? ???? ????????????????" },
      { search: "?????????????????????? ?????????????????? ???? ????????????????", replacement: "?????????????????????? ?????????????????? ???? ????????????????" },
      { search: "?????????????????????? ???????????????????? ???????????????????? ?????????????????????? ?? ???????????????? ????????????????", replacement: "?????????????????????? ?????????????????? ???? ????????????????" },
      { search: "?????????????????????? ???????????????????????? ???????????? ???????????????? ????????????????", replacement: "?????????????????????? ?????????????????? ???? ????????????????" },
      { search: "???????????????????????? ???????????? ????????????????", replacement: "???????????????? ???? ????????????????" },
      { search: "?????????????????? ?????????????????????? ????????????????", replacement: "???????????????? ???? ????????????????" },
      { search: "?????????????????? ???????????? ????????????????", replacement: "???????????????? ???? ????????????????" },
      { search: "???????????????????????? ???????????? ???????????????? ??????????", replacement: "???????????????? ???? ????????????????" },
      { search: "???????????????????????? ???????????? ???????????????? ????????????????", replacement: "???????????????? ???? ????????????????" },
      { search: "?????????????????? ???????????? ????????????????", replacement: "???????????????? ???? ????????????????" },
      { search: "?????????????????? ???????????? ???????????????? ????????????????", replacement: "???????????????? ???? ????????????????" },
      { search: "???????????????????????? ????????????-?????????????????? ???? ????????????????", replacement: "???????????????? ???? ????????????????" },
      { search: "???????????????????????? ???????????????????? ?????????????????????? ????????????????", replacement: "???????????????? ???? ????????????????" },
      { search: "???????????????????????? ???????????? ?????????????? ?? ????????????????", replacement: "???????????????? ???? ????????????????" },
      { search: "?????????????????? ???????????? ????????????????????????????", replacement: "???????????????? ???? ????????????????" },
      { search: "?????????????????? ???????????? ???????????????????? ?? ????????????????", replacement: "???????????????? ???? ????????????????" },
      { search: "???????????????????????? ???????????? ??????????????????", replacement: "???????????????? ???? ??????????????????" },
      { search: "???????????????????????? ???????????? ???????????? ??????????????????", replacement: "???????????????? ???? ??????????????????" },
      { search: "???????????????????????? ???????????? ?????????????? ?? ??????????????????", replacement: "???????????????? ???? ??????????????????" },
      { search: "???????????????????????? ???? ?????????????????? ??????????????????", replacement: "???????????????? ???? ??????????????????" },
      { search: "???????????????????????? ???? ????????????????", replacement: "???????????????? ???? ??????????????????" },
      { search: "???????????????????????? ???????????? ??????????????????", replacement: "???????????????? ???? ??????????????????" },
      { search: "?????????????????? ???????????????????? ??????????????????", replacement: "???????????????? ???? ??????????????????" },
      { search: "???????????????? ???? ??????????????????", replacement: "???????????????? ???? ??????????????????" },
      { search: "???????????????? ???? ????????????", replacement: "???????????????? ???? ??????????????????" },
      { search: "???????????????? ???? ????????????????", replacement: "???????????????? ???? ??????????????????" },
      { search: "?????????????????????? ???????????????????????? ???????????? ??????????????", replacement: "?????????????????????? ?????????????????? ???? ??????????????????" },
      { search: "???????????????????????? ???????????? ???????????????????????? ??????????????????", replacement: "???????????????? ???? ??????????????????" },
      { search: "?????????????????? ???????????? ???????????????????????? ??????????????????", replacement: "???????????????? ???? ??????????????????" },
      { search: "???????????????????????? ???????????? ??????????????", replacement: "???????????????? ???? ??????????????????" },
      { search: "???????????????????????? ???????????? ??????????????????", replacement: "???????????????? ???? ??????????????????" },
      { search: "???????????????????????? ???????????????????????? ???????????????????????? ??????????????????", replacement: "???????????????? ???? ??????????????????" },
      { search: "?????????????????? ???????????????????? ?????????? ?? ??????????????????", replacement: "???????????????? ???? ??????????????????" },
      { search: "?????????????????????? ???????????????????????? ???????????? ??????????????", replacement: "?????????????????????? ?????????????????? ???? ??????????????????" },
      { search: "?????????????????????? ???????????????????????? ???????????? ??????????????????", replacement: "?????????????????????? ?????????????????? ???? ??????????????????" },
      { search: "?????????????????????? ?????????????????? ???? ??????????????????", replacement: "?????????????????????? ?????????????????? ???? ??????????????????" },
      { search: "???????????????????????? ?????????????????? ??????????????????", replacement: "???????????????? ???? ??????????????????" },
      { search: "???????????????????????? ????????????", replacement: "???????????????? ???? ??????????????????" },
      { search: "???????????????? ?????????????????? ???? ??????????????????", replacement: "?????????????????????? ?????????????????? ???? ??????????????????" },
      { search: "???????????????????????? ???????????? ??????????", replacement: "???????????????? ???? ??????????????????" },
      { search: "?????????????????????? ???????????????????????? ???????????? ??????????", replacement: "?????????????????????? ?????????????????? ???? ??????????????????" },
      { search: "???????????????? ???? ??????????????????", replacement: "???????????????? ???? ??????????????????" },
      { search: "?????????????????? ???????????? ???? ??????????????????", replacement: "???????????????? ???? ??????????????????" },
      { search: "???????????????? ???? ????????????????????????", replacement: "???????????????? ????????????????????????" },
      { search: "?????????????????????? ???????????????????????? ?????????????????? ???? ????????????????????????", replacement: "?????????????????????? ?????????????????? ???? ????????????????????????" },
      { search: "???????????????????????? ????????????????????????", replacement: "???????????????? ????????????????????????" },
      { search: "???????????????? ????????????????????????", replacement: "???????????????? ????????????????????????" },
      { search: "?????????????????????? ???????????????? ??????????????????????", replacement: "???????????????? ????????????????????????" },
      { search: "?????????????????? ????????????????????????", replacement: "???????????????? ????????????????????????" },
      { search: "???????????????? ???????????????????? ????????????????????????", replacement: "???????????????? ????????????????????????" },
      { search: "???????????????? ???????????????? ????????????????????????", replacement: "???????????????? ????????????????????????" },
      { search: "?????????????????????? ???????????????????? ????????????????????????", replacement: "???????????????? ????????????????????????" },
      { search: "?????????????????? ?????????????????????????? ????????????????????????", replacement: "???????????????? ????????????????????????" },
      { search: "?????????????????? ?????????????????????????????????????? ????????????????????????", replacement: "???????????????? ????????????????????????" },
      { search: "???????????????? ????????????", replacement: "???????????????? ????????????????????????" },
      { search: "???????????????????????? ???? ????????????????????????", replacement: "???????????????? ????????????????????????" },
      { search: "?????????????????????? ??????????????????????????", replacement: "???????????????? ????????????????????????" },
      { search: "?????????????????? ???????????????? ????????????????????????", replacement: "???????????????? ????????????????????????" },
      { search: "???????????????????????? ??????????????", replacement: "???????????????? ???? ??????????????????????????" },
      { search: "?????????????????? ??????????????", replacement: "???????????????? ???? ??????????????????????????" },
      { search: "???????????????? ???? ??????????????????????????", replacement: "???????????????? ???? ??????????????????????????" },
      { search: "?????????????????????? ???????????????????????? ?????????????????? ???? ??????????????????????????", replacement: "?????????????????????? ?????????????????? ???? ??????????????????????????" },
      { search: "?????????????????????? ???????????????? ????????????????????????", replacement: "???????????????? ???? ??????????????????????????" },
      { search: "???????????????????????? ???????????????? ??????????????????????????", replacement: "???????????????? ???? ??????????????????????????" },
      { search: "?????????????????????? ?????????????????? ???? ??????????????????????????", replacement: "???????????????? ???? ??????????????????????????" },
      { search: "???????????????? ???????????????????????? ?????????????? ???? ??????????????????????????", replacement: "?????????????????????? ?????????????????? ???? ??????????????????????????" },
      { search: "???????????????????????? ?????????????? ??????????????????????????", replacement: "???????????????? ???? ??????????????????????????" },
      { search: "???????????????? ???? ???????????????????????? ??????????????????????????", replacement: "???????????????? ???? ??????????????????????????" },
      { search: "???????????????????????? ????????????????", replacement: "???????????????? ???? ??????????????????????????" },
      { search: "???????????? ???????????????????????? ??????????????????????????", replacement: "???????????????? ???? ??????????????????????????" },
      { search: "?????????????? ?????????????????????????? ??????????????", replacement: "???????????????? ???? ??????????????????????????" },
      { search: "???????????????????????? ?????????????? ???? ??????????????????????????", replacement: "???????????????? ???? ??????????????????????????" },
      { search: "???????????????????????? ???????????????????????? ????????????????", replacement: "???????????????? ???? ??????????????????????????" },
      { search: "???????????????????????? ?????????????????????????? ??????????????", replacement: "???????????????? ???? ??????????????????????????" },
      { search: "???????????????????????? ???? ??????????????????", replacement: "???????????????? ???? ??????????????????????????" },
      { search: "?????????????????? ??????????????????????????", replacement: "???????????????? ???? ??????????????????????????" },
      { search: "???????????????? ???????????????????????? ????????????????", replacement: "???????????????? ???? ??????????????????????????" },
      { search: "???????????????????????? ????????", replacement: "???????????????? ???? ??????????????????????????" },
      { search: "?????????????????????? ???????????????????????? ????????????????", replacement: "?????????????????????? ?????????????????? ???? ??????????????????????????" },
      { search: "???????????????? ??????????", replacement: "???????????????? ???? ??????????????????????????" },
      { search: "???????????????????????? ??????????????????????????", replacement: "???????????????? ???? ??????????????????????????" },
      { search: "?????????????????????? ???????????????????????? ????????????", replacement: "?????????????????????? ?????????????????? ???? ??????????????????????????" },
      { search: "???????????????????????? ????????????????????", replacement: "???????????????? ???? ??????????????????????????" },
      { search: "?????????????????? ??????", replacement: "???????????????? ???? ??????????????????????????" }
    
    
      //{ search: "", replacement: "" },
    ];

    const clearCompaniesArray = [];
    companies.forEach(obj => {
    const index = clearCompaniesArray.findIndex(item => item.companyName === obj.companyName);
    if (index === -1) {
        clearCompaniesArray.push(obj);
    }
    });

    clearCompaniesArray.forEach(obj => {
    for (let i = 0; i < replacements.length; i++) {
        const { search, replacement } = replacements[i];
        if (obj.vacancyName.toLowerCase().includes(search.toLowerCase())) {
        // create a new object with the updated vacancyName
        const updatedObj = { ...obj, vacancyName: replacement };
        if (obj.companyWebsite) { // check if object contains companyWebsite property
            XLSXcompanies.push(updatedObj); // add updated object to new array
        }
        break; // exit loop after first replacement found
        }
    }
    });
  }

  function writeToExcel() {

    var workbook = XLSX.utils.book_new();
    var worksheet = XLSX.utils.json_to_sheet(XLSXcompanies);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Companies");

    for (var i = 0; i < XLSXcompanies.length; i++) {
        //XLSX.writeFile(workbook, "????????????_??????????????????????_??????????????????_????_????????????????????.xlsx");
        //XLSX.writeFile(workbook, "????????????_??????????????????????_????????????????.xlsx");
        XLSX.writeFile(workbook, "????????????_????????????????_??????????_4.xlsx");
        console.log(`Row ${i} has been added to xlsx.`);
    }
  }

  async function main() {
    const allPages = await findPagesNumber();
    console.log('Total Pages:', allPages);
    await fetchAllUrls(allPages);
    prepareForExcel();
    writeToExcel();
    
  }
  
  main().catch((error) => {
    console.error(error);
  });
