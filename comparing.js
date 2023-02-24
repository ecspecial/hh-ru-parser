const XLSX = require('xlsx');

const clearRowNumber =  350;
const forComparingRowNumber = 247;
//const clearRowNumber = 100;
//const forComparingRowNumber = 100;

const company = [];

const workbook = XLSX.readFile('readfrom.xlsx');
const worksheet1 = workbook.Sheets[workbook.SheetNames[0]];
const worksheet2 = workbook.Sheets[workbook.SheetNames[1]];

async function checkValues() {
    for (let i = 2; i <= forComparingRowNumber; i++) {
        const nameWorksheet2 = worksheet2[`A${i}`].v.toString();
        for(z = 2; z <= clearRowNumber; z++) {
            const nameWorksheet1 = worksheet1[`A${z}`].v.toString();
            if (nameWorksheet2 === nameWorksheet1) {
    
                company[i] = ({
                    companyName: nameWorksheet2,
                    companyLink: worksheet2[`B${i}`].v.toString(),
                    vacanciesNumber: worksheet1[`C${z}`].v.toString(),
                    companyWebsite: worksheet1[`D${z}`].v.toString(),
                    companyPhone: worksheet1[`E${z}`].v.toString(),
                    companyEmail: worksheet1[`F${z}`].v.toString(),
                    vacancyName: worksheet2[`G${i}`].v.toString(),
                    vacancyLink: worksheet2[`H${i}`].v.toString()
                });
    
                console.log('Strings are same');
    
            console.log(nameWorksheet1, nameWorksheet2);
            } else {
                console.log('The strings are not equal.');
            }
        } 
        console.log(`ITERATION ${i} DONE`);
    } 
}


async function writeToExcel() {
    var newWorkbook = XLSX.utils.book_new();
    var newWorksheet = XLSX.utils.json_to_sheet(company);
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Companies");
    XLSX.writeFile(newWorkbook, "writeto.xlsx");
}
async function main() {
    await checkValues();
    await writeToExcel();
}

main();