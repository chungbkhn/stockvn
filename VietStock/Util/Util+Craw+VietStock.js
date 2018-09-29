var request = require('request');
const cheerio = require('cheerio');

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

var util_craw_vietstock = {
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
                console.log("successed! Get: " + code + " for page: " + pageNumber);
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

                if (type == 'LC') {
                    let trs = $('table tbody').children();
                    var startCalculate = false;
                    for (const tr in trs) {
                        if (trs.attr('id') == 'BR_rowHeader') {
                            startCalculate = true;
                            continue;
                        }

                        if (startCalculate) {
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
                        }
                    }
                } else {
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
                }

                console.log("successed! Get: " + code + " for page: " + pageNumber);
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
    calculateData: function (data, needSumQuater) {
        if (data.length == 0) { callback(data) };

        var listColQ4 = [];
        const listTitle = data[0];
        var newListTitle = [];

        var startUnusedData = 0;
        var endUnusedData = 0;
        for (let col = 0; col < listTitle.length; col++) {
            const item = listTitle[col];
            if (item.length > 0) { 
                startUnusedData = col;
                break;
            }
        }

        for (let col = 0; col < listTitle.length; col++) {
            const item = listTitle[col];
            if (item.indexOf('Quý 1') > -1) { 
                endUnusedData = col;
                break;
            }
        }

        for (let col = 0; col < listTitle.length; col++) {
            if (col >= startUnusedData && col < endUnusedData) { continue; }

            newListTitle.push(listTitle[col]);
            const items = listTitle[col].split('/');
            if (items.length != 2) { continue };

            const quater = items[0];
            const year = 'Năm ' + items[1];
            if (quater == 'Quý 4') {
                listColQ4.push(col);
                newListTitle.push(year);
            }
        }

        var newData = [];
        newData.push(newListTitle);

        for (let row = 1; row < data.length; row++) {
            var oldListItem = data[row];
            var newListItem = [];
            for (let col = 0; col < oldListItem.length; col++) {
                const item = oldListItem[col];

                newListItem.push(item);
                if (listColQ4.indexOf(col) > -1) {
                    if (needSumQuater) {
                        var value = oldListItem[col];
                        if (col >= 3 && listTitle[col - 3].indexOf('Quý 1') > -1) {
                            value += oldListItem[col - 3];
                        }
                        if (col >= 2 && listTitle[col - 2].indexOf('Quý 2') > -1) {
                            value += oldListItem[col - 2];
                        }
                        if (col >= 1 && listTitle[col - 1].indexOf('Quý 3') > -1) {
                            value += oldListItem[col - 1];
                        }
                        newListItem.push(value);
                    } else {
                        newListItem.push(item);
                    }
                }
            }
            newData.push(newListItem);
        }

        return newData;
    }
};

module.exports = util_craw_vietstock;