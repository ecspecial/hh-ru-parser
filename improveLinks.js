const cheerio = require('cheerio');
const XLSX = require('xlsx');
const path = require('path');

const allVacanciesLink = [];
const matchedCompaniesVacancies = [];
const newVacancyLinks = [];
const companies = [];
const rowNumber = 350;
const maxRow = 2;
const fileToReadFrom = 'Запрос_Директор_строй.xlsx';
const fileToWriteTo = "IMPROVED_Директор_по_Логистике.xlsx";


const replacements = [
  { search: "генеральный директор", replacement: "Генеральный директор" },
  { search: "финансовый директор", replacement: "Финансовый директор" },
  { search: "заместитель генерального директора", replacement: "Заместитель генерального директора" },
  { search: "заместитель финансового директора", replacement: "Заместитель финансового директора" },
  { search: "директор операционного офиса", replacement: "Операционный директор" },
  { search: "операционный директор", replacement: "Операционный директор" },
  { search: "заместитель операционного директора", replacement: "Заместитель операционного директора" },
  { search: "операционный управляющий", replacement: "Операционный директор" },
  { search: "финансист-операционный директор", replacement: "Операционный директор" },
  { search: "исполнительный директор", replacement: "Исполнительный директор" },
  { search: "заместитель исполнительного директора", replacement: "Заместитель исполнительного директора" },
  { search: "заместитель технического/исполнительного директора", replacement: "Заместитель исполнительного директора" },
  { search: "директор по операционному маркетингу", replacement: "Директор по маркетингу" },
  { search: "коммерческий директор", replacement: "Коммерческий директор" },
  { search: "заместитель коммерческого директора", replacement: "Заместитель коммерческого директора" },
  { search: "директор по развитию", replacement: "Директор по развитию" },
  { search: "руководитель отдела по развитию", replacement: "Директор по развитию" },
  { search: "директор по маркетингу", replacement: "Директор по маркетингу" },
  { search: "руководитель отдела маркетинга", replacement: "Директор по маркетингу" },
  { search: "руководитель в отдел маркетинга", replacement: "Директор по маркетингу" },
  { search: "руководитель маркетинга", replacement: "Директор по маркетингу" },
  { search: "начальник отдела маркетинга", replacement: "Директор по маркетингу" },
  { search: "руководитель службы маркетинга", replacement: "Директор по маркетингу" },
  { search: "руководитель удаленного маркетинга", replacement: "Директор по маркетингу" },
  { search: "начальник отдела продаж", replacement: "Директор по маркетингу" },
  { search: "заместитель директора по маркетингу", replacement: "Заместитель директора по маркетингу" },
  { search: "технический директор", replacement: "Технический директор" },
  { search: "заместитель технического директора", replacement: "Заместитель технического директора" },
  { search: "cto", replacement: "Технический директор" },
  { search: "территориальный директор", replacement: "Территориальный директор" },
  { search: "региональный управляющий", replacement: "Территориальный директор" },
  { search: "региональный директор", replacement: "Территориальный директор" },
  { search: "директор по производству", replacement: "Генеральный директор" },
  { search: "руководитель отдела по работе с персоналом", replacement: "Директор по персоналу" },
  { search: "hr директор", replacement: "Директор по персоналу" },
  { search: "hr-директор", replacement: "Директор по персоналу" },
  { search: "директор по персоналу", replacement: "Директор по персоналу" },
  { search: "руководитель по персоналу", replacement: "Директор по персоналу" },
  { search: "заместитель директора по персоналу", replacement: "Заместитель директора по персоналу" },
  { search: "начальник управления по работе с персоналом", replacement: "Директор по персоналу" },
  { search: "начальник отдела по работе с персоналом", replacement: "Директор по персоналу" },
  { search: "руководитель отдела персонала", replacement: "Директор по персоналу" },
  { search: "hrd", replacement: "Директор по персоналу" },
  { search: "директор по управлению персоналом", replacement: "Директор по персоналу" },
  { search: "hr generalist", replacement: "Директор по персоналу" },
  { search: "директор по продажам", replacement: "Директор по продажам" },
  { search: "заместитель директора по продажам", replacement: "Заместитель директора по продажам" },
  { search: "руководитель отдела продаж", replacement: "Директор по продажам" },
  { search: "директор филиала", replacement: "Территориальный директор" },
  { search: "руководитель филиала", replacement: "Территориальный директор" },
  { search: "руководитель обособленного подразделения", replacement: "Территориальный директор" },
  { search: "руководитель обособленного филиала", replacement: "Территориальный директор" },
  { search: "руководитель отдела контроля качества", replacement: "Директор по качеству" },
  { search: "директор по качеству", replacement: "Директор по качеству" },
  { search: "заместитель директора по качеству", replacement: "Заместитель директора по качеству" },
  { search: "заместитель начальника управления обеспечения и контроля качества", replacement: "Заместитель директора по качеству" },
  { search: "заместитель руководителя отдела контроля качества", replacement: "Заместитель директора по качеству" },
  { search: "руководитель службы качества", replacement: "Директор по качеству" },
  { search: "начальник обеспечения качества", replacement: "Директор по качеству" },
  { search: "начальник службы качества", replacement: "Директор по качеству" },
  { search: "руководитель группы качества услуг", replacement: "Директор по качеству" },
  { search: "руководитель группы контроля качества", replacement: "Директор по качеству" },
  { search: "начальник отдела качества", replacement: "Директор по качеству" },
  { search: "начальник отдела контроля качества", replacement: "Директор по качеству" },
  { search: "руководитель бизнес-процессов по качеству", replacement: "Директор по качеству" },
  { search: "руководитель управления обеспечения качества", replacement: "Директор по качеству" },
  { search: "руководитель отдела сервиса и качества", replacement: "Директор по качеству" },
  { search: "начальник отдела стандартизации", replacement: "Директор по качеству" },
  { search: "начальник отдела технологии и качества", replacement: "Директор по качеству" },
  { search: "руководитель отдела логистики", replacement: "Директор по логистике" },
  { search: "руководитель отдела продаж логистики", replacement: "Директор по логистике" },
  { search: "руководитель отдела закупок и логистики", replacement: "Директор по логистике" },
  { search: "руководитель по складской логистике", replacement: "Директор по логистике" },
  { search: "руководитель по закупкам", replacement: "Директор по логистике" },
  { search: "руководитель службы снабжения", replacement: "Директор по логистике" },
  { search: "начальник управления снабжения", replacement: "Директор по логистике" },
  { search: "директор по логистике", replacement: "Директор по логистике" },
  { search: "директор по складу", replacement: "Директор по логистике" },
  { search: "директор по закупкам", replacement: "Директор по логистике" },
  { search: "заместитель руководителя отдела закупок", replacement: "Заместитель директора по логистике" },
  { search: "руководитель отдела транспортной логистики", replacement: "Директор по логистике" },
  { search: "начальник отдела транспортной логистики", replacement: "Директор по логистике" },
  { search: "руководитель отдела закупок", replacement: "Директор по логистике" },
  { search: "руководитель службы логистики", replacement: "Директор по логистике" },
  { search: "руководитель департамента транспортной логистики", replacement: "Директор по логистике" },
  { search: "начальник управления сбыта и логистики", replacement: "Директор по логистике" },
  { search: "заместитель руководителя отдела закупок", replacement: "Заместитель директора по логистике" },
  { search: "заместитель руководителя отдела логистики", replacement: "Заместитель директора по логистике" },
  { search: "заместитель директора по логистике", replacement: "Заместитель директора по логистике" },
  { search: "руководитель складской логистике", replacement: "Директор по логистике" },
  { search: "руководитель склада", replacement: "Директор по логистике" },
  { search: "помощник директора по логистике", replacement: "Заместитель директора по логистике" },
  { search: "руководитель отдела склад", replacement: "Директор по логистике" },
  { search: "заместитель руководителя отдела транс", replacement: "Заместитель директора по логистике" },
  { search: "директор по снабжению", replacement: "Директор по логистике" },
  { search: "начальник отдела по снабжению", replacement: "Директор по логистике" },
  { search: "директор по производству", replacement: "Директор производства" },
  { search: "заместитель генерального директора по производству", replacement: "Заместитель директора по производству" },
  { search: "руководитель производства", replacement: "Директор производства" },
  { search: "директор производства", replacement: "Директор производства" },
  { search: "генеральный директор производств", replacement: "Директор производства" },
  { search: "начальник производства", replacement: "Директор производства" },
  { search: "директор мебельного производства", replacement: "Директор производства" },
  { search: "директор пищевого производства", replacement: "Директор производства" },
  { search: "заместитель начальника производства", replacement: "Директор производства" },
  { search: "начальник кондитерского производства", replacement: "Директор производства" },
  { search: "начальник машиностроительного производства", replacement: "Директор производства" },
  { search: "директор завода", replacement: "Директор производства" },
  { search: "руководитель по производству", replacement: "Директор производства" },
  { search: "управляющий производством", replacement: "Директор производства" },
  { search: "начальник швейного производства", replacement: "Директор производства" },
  { search: "руководитель проекта", replacement: "Директор по строительству" },
  { search: "начальник участка", replacement: "Директор по строительству" },
  { search: "директор по строительству", replacement: "Директор по строительству" },
  { search: "заместитель генерального директора по строительству", replacement: "Заместитель директора по строительству" },
  { search: "генеральный директор производство", replacement: "Директор по строительству" },
  { search: "руководитель проектов строительства", replacement: "Директор по строительству" },
  { search: "заместитель директора по строительству", replacement: "Директор по строительству" },
  { search: "помощник руководителя проекта по строительству", replacement: "Заместитель директора по строительству" },
  { search: "руководитель проекта строительства", replacement: "Директор по строительству" },
  { search: "директор по капитальному строительству", replacement: "Директор по строительству" },
  { search: "руководитель проектов", replacement: "Директор по строительству" },
  { search: "прораб малоэтажного строительства", replacement: "Директор по строительству" },
  { search: "куратор строительного проекта", replacement: "Директор по строительству" },
  { search: "руководитель проекта по строительству", replacement: "Директор по строительству" },
  { search: "руководитель строительных проектов", replacement: "Директор по строительству" },
  { search: "руководитель строительного проекта", replacement: "Директор по строительству" },
  { search: "руководитель по снабжению", replacement: "Директор по строительству" },
  { search: "начальник строительства", replacement: "Директор по строительству" },
  { search: "директор строительной компании", replacement: "Директор по строительству" },
  { search: "руководитель стро", replacement: "Директор по строительству" },
  { search: "заместитель руководителя проектов", replacement: "Заместитель директора по строительству" },
  { search: "директор строи", replacement: "Директор по строительству" },
  { search: "руководитель строительства", replacement: "Директор по строительству" },
  { search: "заместитель руководителя девело", replacement: "Заместитель директора по строительству" },
  { search: "руководитель девелоперс", replacement: "Директор по строительству" },
  { search: "Начальник пто", replacement: "Директор по строительству" }


  //{ search: "", replacement: "" },
];

  async function fetchXLSXurl() {
    const workbook = XLSX.readFile(fileToReadFrom);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  
    console.log('Starting...');
    // get range of cells in column B starting from row 2
    //const range = worksheet['B2:B' + worksheet['!ref'].split(':')[1]];
  
    try {
      for (let i = maxRow; i <= rowNumber; i++) {
        const link = worksheet[`B${i}`].v;
        const name = worksheet[`A${i}`].v;
  
        //console.log(link);
        //console.log(name);
  
        const response = await fetch(link);
        const html = await response.text();
        console.log(`-----------------------FETCHING LINK ${i}`);
        console.log(link);
        console.log(response.status);
        try {
            const $ = cheerio.load(html);
            // Find link to all vacancies
            await searchVacancy($, i);
            // Find number of pages for vacancy
            const allPages = await findPagesNumber(i);
            
            // Find the right vacancy title
            await findVacancyTitle(i, allPages);
  
            companies[i] = ({
                companyName: name,
                companyLink: link,
                companyVacancyLink: newVacancyLinks[i],
                newVacancyName: matchedCompaniesVacancies[i]
            });
        
            await writeToExcel(i, name, link);
  
            console.log('Current link to all vacancies: ' + allVacanciesLink[i]);
        } catch (error) {
            console.error('LINK TO VACANCIES NOT FOUND', error);
        }

      }
    } catch (error) {
      console.error('Error occurred during for loop iteration:', error);
    }
  }

  async function searchVacancy($, i) {
    const tag = $('a[data-qa="employer-page__employer-vacancies-link"]');
    const vacanciesLink = 'https://hh.ru' + tag.attr('href');
    allVacanciesLink[i] = vacanciesLink;
  }

  async function findVacancyTitle(i, allPages) {
    const pageUrl = allVacanciesLink[i];
    if (!isNaN(allPages)) {

        //for (i in pages)
        for (y = 0; y < 10; y++) {
            const link = (pageUrl + `&page=${y}`);
            const matchFound = await findVacancyByName(i, link);
            if (matchFound) {
                console.log(`Match found on page ${y}, breaking loop`);
                console.log(`Current text in ${i} is ${matchedCompaniesVacancies[i]}`);
                console.log(`Current link in ${i} is ${newVacancyLinks[i]}`);
                break;
            }
        }

        console.log(`Number of pages for : ${allVacanciesLink[i]}`, allPages);
    } else {
        console.log(`Error: allPages is not a number for link ${allVacanciesLink[i]}`);
        const singlePageMatchFound = await findVacancyByName(i, pageUrl);
        if (singlePageMatchFound) {
            console.log(`SINGLE LOOP SUCCESS. Match found on link ${i}, breaking loop`);
            console.log(`Current text in ${i} is ${matchedCompaniesVacancies[i]}`);
            console.log(`Current link in ${i} is ${newVacancyLinks[i]}`);
        }
        else {
            console.log(`ERROR IN SINGLE LOOP. No Match found on link ${i}, breaking loop`);
        }
    }
}

  async function findPagesNumber(i) {
    
    const response = await fetch(allVacanciesLink[i]);
    const html = await response.text();
    const $ = cheerio.load(html);

    const endPages = $(".pager a").last();
    const secondLastSpan = endPages.prev('span');
    const allPages = parseInt(secondLastSpan.find('a[data-qa="pager-page"]').text());

    return allPages;
}

async function findVacancyByName(i, link) {
    const response = await fetch(link);
    const html = await response.text();
    const $ = cheerio.load(html);
  
    const div = $('.vacancy-serp-content');
    const vacancyNames = div.find('.serp-item__title');
    const vacancyLinks = $('a.serp-item__title');
    //console.log('______________' + $(vacancyLinks[1]).attr('href'));
  
    for (let z = 0; z < vacancyNames.length; z++) {
      const title = vacancyNames[z];
      const link = vacancyLinks[z];
      for (let y = 0; y < replacements.length; y++) {
        const { search, replacement } = replacements[y];
        if ($(title).text().toLowerCase().includes(search.toLowerCase())) {
            //console.log('found a match');
            matchedCompaniesVacancies[i] = replacement;
            //console.log(`Found a match and put it into index  ${i}`);
            console.log('______________' + $(link).attr('href'));
            newVacancyLinks[i] = $(link).attr('href');
            return true; // return true if match is found

            // Write to xlsx
        }
      }
    }
  
    return false; // return false if no match is found
  }

  async function writeToExcel(i) {

    var workbook = XLSX.utils.book_new();
    var worksheet = XLSX.utils.json_to_sheet(companies);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Companies");

    XLSX.writeFile(workbook, fileToWriteTo);

    console.log(`LINE ${i} written to XLSX`);
  }

fetchXLSXurl();