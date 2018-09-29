var xl = require('excel4node');

// Create a new instance of a Workbook class
var wb = new xl.Workbook();

// Add Worksheets to the workbook
// var ws = wb.addWorksheet('Sheet 1');
// Create a reusable style
var styleNumber = wb.createStyle({
    font: {
        color: '#000000',
        size: 12,
    },
    numberFormat: '#,##; (#,##); -',
});

var stylePercent = wb.createStyle({
    font: {
        color: '#000000',
        size: 12,
    },
    numberFormat: '#0.##%; (#0.##%); -'
});

var util_excel = {
    writeDataToExcel: function (name) {
        wb.write(name);
        console.log('Write data successful!');
    },
    addDataToExcel: function (data, sheetName) {
        var ws = wb.addWorksheet(sheetName);
        for (let row = 0; row < data.length; row++) {
            const item = data[row];
            var isPercentValue = false;
            for (let column = 0; column < item.length; column++) {
                const value = item[column];
                var number = Number(value);
                if (isNaN(number)) {
                    ws.cell(row + 1, column + 1)
                        .string(value)
                        .style(styleNumber);
                    if (value == '%') {
                        isPercentValue = true;
                    }
                } else {
                    var style = styleNumber;
                    if (isPercentValue) {
                        style = stylePercent;
                        number = number / 100.0;
                    }
                    ws.cell(row + 1, column + 1)
                        .number(number)
                        .style(style);
                }
            }
        }

        console.log('add data to sheet ' + sheetName + ' successful!');
    }
};

module.exports = util_excel;