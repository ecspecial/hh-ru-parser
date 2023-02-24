# hh-ru-parser

This is a demo version of hh.ru vacancy parser

parser.js is used to parse company info such as: 
- company name
- company vacancies number
- company vacancy link
- public website, phone number and Email

findContactPage.js is used to search website links for Contact pages and parse public contacts, returnes phone number and email

findContacts.js is used to parse main website link and search for phone number and Email on it

improveLinks.js is used to double check parsed links to vacancies and improve if not

comparing.js is used to compare improved vacancies links and primary parsed