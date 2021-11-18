var Excel = require('exceljs');

var wb = new Excel.Workbook();
var path = require('path');
var filePath = path.resolve(__dirname,'EXCEL_FILE.xlsx');

wb.xlsx.readFile(filePath).then(function(){

    var sh = wb.getWorksheet("Sheet1");

    sh.getRow(1).getCell(2).value;
    wb.xlsx.writeFile("sample2.xlsx");
   // console.log("Row-3 | Cell-2 - "+sh.getRow(3).getCell(2).value);

    console.log(sh.rowCount);
    //Get all the rows data [1st and 2nd column]
    //for(let x=0;x<3;x++)
    for (i = 1; i <= sh.rowCount; i++) {
        console.log(sh.getRow(i).getCell(1).value+" "+sh.getRow(i).getCell(2).value+" "+sh.getRow(i).getCell(3).value+" "+sh.getRow(i).getCell(4).value);
      
    }
});
