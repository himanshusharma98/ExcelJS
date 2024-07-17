const ExcelJS = require('exceljs');
async function  ExcelDemo() {
    
const workbook = new ExcelJS.Workbook();
await workbook.xlsx.readFile('C:\Users\sharma\Downloads\download.xlsx');
const worksheet = workbook.getWorksheet('Sheet1');
worksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell, colNumber) => {
        console.log(cell.value);
    })
})
}
ExcelDemo();