var request = require('request');
const cheerio = require('cheerio');

var xl = require('excel4node');

// Create a new instance of a Workbook class
var wb = new xl.Workbook();

// Add Worksheets to the workbook
// var ws = wb.addWorksheet('Sheet 1');
// Create a reusable style
var style = wb.createStyle({
    font: {
        color: '#000000',
        size: 12,
    },
    numberFormat: '#,##0; (#,##0); -',
});
// Set value of cell B1 to 200 as a number type styled with paramaters of style
// ws.cell(1, 2)
//   .number(200)
//   .style(style);


function combineData(oldData, data, index) {
    var result = [];
    var idx = 0;
    while (idx < oldData.length) {
        var oldItem = oldData[idx];
        var newItem = data[idx];
        var item = [];
        newItem.forEach(element => {
            item.push(element);
        });

        for (var itemIdx = index; itemIdx < oldItem.length; itemIdx++) {
            item.push(oldItem[itemIdx]);
        }
        result.push(item);
        idx++;
    }
    return result;
}

var util = {
    loadCSTC: function (code, pageNumber, callback) {
        var link = 'http://finance.vietstock.vn/Controls/Report/Data/GetReport.ashx?rptType=CSTC&scode=' + code + '&bizType=1&rptUnit=1000000&rptTermTypeID=2&page=' + pageNumber;
        request(link, function (error, response, body) {
            if (!error && response.statusCode == 200) {
                var $ = cheerio.load(body, {
                    xmlMode: true
                });

                // new code
                var data = [];
                var item = [];
                var idx = 0;
                item[idx++] = "";

                $('table thead tr#BR_rowHeader td.BR_colHeader_Time').each(function (i, element) {
                    item[idx] = $(this).text();
                    idx += 1;
                });

                data.push(item);

                $('table tbody tr.BR_tBody_rowName').each(function (i, elem) {
                    var tr = $(this);
                    var nameRow = tr.find('td.BR_tBody_colName.Padding1').text();

                    idx = 0;
                    item = [];
                    item[idx++] = nameRow;

                    var unitRow = tr.find('td.FR_tBody_colUnit').first().text();
                    item[idx++] = unitRow;

                    var values = tr.find('span.rpt_chart').first().text().split(',');
                    for (var i in values) {
                        let value = values[i];
                        if (value == '_') {
                            value = '';
                        }
                        item[idx++] = value;
                    }

                    data.push(item);
                });

                // console.log(data);
                console.log("successed! Get: " + code + "for page: " + pageNumber);
                callback(data);
            }
        });
    },
    combineLoadCSTC: function (code, pageNumber, oldData, callback) {
        this.loadCSTC(code, pageNumber, function (data) {
            if (oldData.length == 0) { callback(data) }
            else {
                var result = combineData(oldData, data, 2);
                callback(result);
            }
        });
    },
    loadBCTC: function (code, pageNumber, type, callback) {
        var link = 'http://finance.vietstock.vn/Controls/Report/Data/GetReport.ashx?rptType=' + type + '&scode=' + code + '&bizType=1&rptUnit=1000000&rptTermTypeID=2&page=' + pageNumber;
        request(link, function (error, response, body) {
            if (!error && response.statusCode == 200) {
                var $ = cheerio.load(body, {
                    xmlMode: true
                });

                // new code
                var data = [];
                var item = [];
                var idx = 0;
                item[idx++] = "";

                $('table thead tr#BR_rowHeader td.BR_colHeader_Time').each(function (i, element) {
                    $(this).find('br').replaceWith('|')
                    var value = $(this).first().contents().filter(function () {
                        return this.type === 'text';
                    }).text();
                    value = value.split('|')[0];

                    item[idx] = value;
                    idx += 1;
                });

                data.push(item);

                $('table tbody tr.BR_tBody_rowName').each(function (i, elem) {
                    var tr = $(this);
                    var nameRow = tr.find('td.BR_tBody_colName').text();

                    idx = 0;
                    item = [];
                    item[idx++] = nameRow;

                    var values = tr.find('span.rpt_chart').first().text().split(',');
                    for (var i in values) {
                        let value = values[i];
                        if (value == '_') {
                            value = '';
                        }
                        item[idx++] = value;
                    }

                    data.push(item);
                });

                console.log("successed! Get: " + code + "for page: " + pageNumber);
                callback(data);
            }
        });
    },
    combineLoadCDKT: function (code, pageNumber, oldData, callback) {
        this.loadBCTC(code, pageNumber, 'CDKT', function (data) {
            if (oldData.length == 0) { callback(data) }
            else {
                var result = combineData(oldData, data, 1);
                callback(result);
            }
        });
    },
    combineLoadKQKD: function (code, pageNumber, oldData, callback) {
        this.loadBCTC(code, pageNumber, 'KQKD', function (data) {
            if (oldData.length == 0) { callback(data) }
            else {
                var result = combineData(oldData, data, 1);
                callback(result);
            }
        });
    },
    combineLoadLCTT: function (code, pageNumber, oldData, callback) {
        this.loadBCTC(code, pageNumber, 'LC', function (data) {
            if (oldData.length == 0) { callback(data) }
            else {
                var result = combineData(oldData, data, 1);
                callback(result);
            }
        });
    },
    writeDataToExcel: function(name) {
        wb.write(name);
        console.log('Write data successful!');
    },
    addDataToExcel: function(data, sheetName) {
        var ws = wb.addWorksheet(sheetName);
        for (let row = 0; row < data.length; row++) {
            const item = data[row];
            for (let column = 0; column < item.length; column++) {
                const value = item[column];
                let number = Number(value);
                if (isNaN(number)) {
                    ws.cell(row + 1, column + 1)
                        .string(value)
                        .style(style);
                } else {
                    ws.cell(row + 1, column + 1)
                        .number(number)
                        .style(style);
                }
            }
        }
    
        console.log('add data to sheet ' + sheetName + ' successful!');
    }
};

module.exports = util;