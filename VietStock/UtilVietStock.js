var request = require('request');
const cheerio = require('cheerio');

var util = {
    loadData: function (code, pageNumber, callback) {
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

                $('table tbody tr.0.BR_tBody_rowName').each(function (i, elem) {
                    var tr = $(this);
                    var nameRow = tr.find('td.BR_tBody_colName.Padding1').text();

                    idx = 0;
                    item = [];
                    item[idx++] = nameRow;

                    var unitRow = tr.find('td.FR_tBody_colUnit').first().text();
                    item[idx++] = unitRow;

                    tr.find('td.BR_tBody_colValue').each(function (sub, subE) {
                        item[idx++] = $(this).text();
                    });

                    data.push(item);
                });

                // console.log(data);
                console.log("successed! Get: " + code + "for page: " + pageNumber);
                callback(data);
            }
        });
    },
    combineLoadData: function (code, pageNumber, oldData, callback) {
        this.loadData(code, pageNumber, function (data) {
            if (oldData.length == 0) { callback(data) }
            else {
                var result = [];
                var idx = 0;
                while (idx < oldData.length) {
                    var oldItem = oldData[idx];
                    var newItem = data[idx];
                    var item = [];
                    newItem.forEach(element => {
                        item.push(element);
                    });

                    for (var itemIdx = 2; itemIdx < oldItem.length; itemIdx++) {
                        item.push(oldItem[itemIdx]);
                    }
                    result.push(item);
                    idx ++;
                }
                callback(result);
            }
        });
    }
};

module.exports = util;