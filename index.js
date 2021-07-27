const xlsx = require('xlsx');
const htmltoText = require('html-to-formatted-text');

const wb = xlsx.readFile("Book1.xlsx");
const ws = wb.Sheets["Sheet1"];
const data = xlsx.utils.sheet_to_json(ws);
const newData = data.map((row)=>{
    const txt = htmltoText(row["Item Description HTML"]);
    delete row["Item Description HTML"];
    const utf = Buffer.from(txt, 'utf-8').toString().trim();
    row["Item Description - Normal Text"] = utf;
    return row;
})

const newWb= xlsx.utils.book_new();
const newWs = xlsx.utils.json_to_sheet(newData);
xlsx.utils.book_append_sheet(newWb,newWs,"New Data");
xlsx.writeFile(newWb,"New file.xlsx");
console.log("Operation completed!!");