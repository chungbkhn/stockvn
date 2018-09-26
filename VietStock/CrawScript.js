var request = require('request');
const cheerio = require('cheerio');
var xl = require('excel4node');
// Create a new instance of a Workbook class
var wb = new xl.Workbook();
 
// Add Worksheets to the workbook
var ws = wb.addWorksheet('Sheet 1');

// Set value of cell B1 to 200 as a number type styled with paramaters of style
// ws.cell(1, 2)
//   .number(200)
//   .style(style);

// Create a reusable style
var style = wb.createStyle({
    font: {
      color: '#000000',
      size: 12,
    },
    numberFormat: '$#,##0.00; ($#,##0.00); -',
  });

request('http://finance.vietstock.vn/Controls/Report/Data/GetReport.ashx?rptType=CSTC&scode=DHA&bizType=1&rptUnit=1000000&rptTermTypeID=2&page=1', function (error, response, body) {
    if (!error && response.statusCode == 200) {
        var $ = cheerio.load(body, {
            xmlMode: true
        });

        var idxColumn = 2;
        var idxRow = 1;
        $('table thead tr#BR_rowHeader td.BR_colHeader_Time').each(function (i, element){
            const columnName = $(this).text();
            ws.cell(idxRow, idxColumn)
            .string(columnName)
            .style(style);
            idxColumn += 1;
        });

        idxRow = 2;
        $('table tbody tr.0.BR_tBody_rowName').each(function (i, elem){
            var tr = $(this);
            var nameRow = tr.find('td.BR_tBody_colName.Padding1').text();
            // console.log('Name:',nameRow);
            ws.cell(idxRow, 1)
            .string(nameRow)
            .style(style);

            idxColumn = 2;
            tr.find('td.BR_tBody_colValue').each(function (sub, subE){
                // console.log($(this).text());
                ws.cell(idxRow, idxColumn)
                .string($(this).text())
                .style(style);
                idxColumn += 1;
            });
            idxRow += 1;
        });

        wb.write('Excel.xlsx');
        console.log("successed!");
    }
});