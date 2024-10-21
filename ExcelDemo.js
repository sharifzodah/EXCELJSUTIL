const ExceJS = require('exceljs');

async function excelTest() 
{
    const workbook = new ExceJS.Workbook();
    await workbook.xlsx.readFile('D:\\Downloads - 2017\\Exceldownload.xlsx');
    const worksheet = workbook.getWorksheet('Sheet1');
    worksheet.eachRow((row, rowNumber) =>
        {
            row.eachCell((cell, colNumber) =>
                {
                    console.log(cell.value);
                })

        })
}

excelTest();

