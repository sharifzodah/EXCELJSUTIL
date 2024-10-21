const ExceJS = require('exceljs');

async function writeExcel(searchText, newText, changeCoord, filePath) 
{    
    const workbook = new ExceJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.getWorksheet('Sheet1');
    output = await readExcel(worksheet, searchText);
    const cell = worksheet.getCell(output.row, output.column + changeCoord.colChange);
    cell.value = newText;
    await workbook.xlsx.writeFile(filePath)
}

async function readExcel(worksheet, searchText)
{
    let output = {row:-1, column:-1};
    worksheet.eachRow((row, rowNumber) =>
        {
            row.eachCell((cell, colNumber) =>
                {
                    if(cell.value === searchText)
                    {
                        output.row = rowNumber;
                        output.column = colNumber;
                    }
                })

        })
        return output;    
}

const searchText = 'Republic';
const newText = 9000;
const filePath = 'D:\\Downloads - 2017\\Exceldownload.xlsx';
const changeCoord = {rowChange: 0, colChange: 2};
writeExcel(searchText, newText, changeCoord, filePath);

