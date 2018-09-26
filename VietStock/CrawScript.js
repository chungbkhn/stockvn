var util = require('./UtilVietStock.js');
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

function loadCSTC(code, pageNumber, endPageNumber, data, callback) {
    util.combineLoadCSTC(code, pageNumber, data, function(newData) {
        pageNumber ++;
        if (pageNumber <= endPageNumber) {
            loadCSTC(code, pageNumber, endPageNumber, newData, callback);
        } else {
            callback(newData);
        }
    })
}

var pageNumber = 1;
var endPageNumber = 10;
var code = 'dha';
var data = [];
loadCSTC(code, pageNumber, endPageNumber, data, function(newData) {
    for (let row = 0; row < newData.length; row++) {
        const item = newData[row];
        for (let column = 0; column < item.length; column++) {
            const value = item[column];
            ws.cell(row + 1, column + 1)
            .string(value)
            .style(style);
        }
    }

    wb.write('excel.xlsx');
    console.log('Write data successful!');
})