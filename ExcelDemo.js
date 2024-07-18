const ExcelJS = require('exceljs');

async function ExcelDemo() {
    try {
        let output = {row:-1,column:-1};
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile('D:\\ExcelJS\\ExcelFile.xlsx');
        const worksheet = workbook.getWorksheet('Sheet1');
        worksheet.eachRow((row, rowNumber) => {
            row.eachCell((cell, colNumber) => {
                console.log(cell.value);
                if(cell.value === "Banana") {
                    output.row = rowNumber;
                    output.column = colNumber;
            
                }
            })
        })

        const cell = worksheet.getCell(output.row, output.column); 
        cell.value = "Republic";  
        await workbook.xlsx.writeFile('D:\\ExcelJS\\ExcelFile.xlsx');
    } catch (error) {
        console.error(error);
    }
}

ExcelDemo();