var util_craw_vietstock = require('./Util/Util+Craw+VietStock.js');
var util_excel = require('./Util/Util+Excel.js');

function loadCSTC(code, pageNumber, endPageNumber, data, callback) {
    util_craw_vietstock.combineLoadCSTC(code, pageNumber, data, function (newData) {
        pageNumber++;
        if (pageNumber <= endPageNumber) {
            loadCSTC(code, pageNumber, endPageNumber, newData, callback);
        } else {
            callback(newData);
        }
    })
}

function loadCDKT(code, pageNumber, endPageNumber, data, callback) {
    util_craw_vietstock.combineLoadCDKT(code, pageNumber, data, function (newData) {
        pageNumber++;
        if (pageNumber <= endPageNumber) {
            loadCDKT(code, pageNumber, endPageNumber, newData, callback);
        } else {
            callback(newData);
        }
    })
}

function loadKQKD(code, pageNumber, endPageNumber, data, callback) {
    util_craw_vietstock.combineLoadKQKD(code, pageNumber, data, function (newData) {
        pageNumber++;
        if (pageNumber <= endPageNumber) {
            loadKQKD(code, pageNumber, endPageNumber, newData, callback);
        } else {
            callback(newData);
        }
    })
}

function loadLCTT(code, pageNumber, endPageNumber, data, callback) {
    util_craw_vietstock.combineLoadLCTT(code, pageNumber, data, function (newData) {
        pageNumber++;
        if (pageNumber <= endPageNumber) {
            loadLCTT(code, pageNumber, endPageNumber, newData, callback);
        } else {
            callback(newData);
        }
    })
}

function loadPTBCTC(code) {
    var pageNumber = 1;
    var endPageNumber = 7;
    var data = [];

    loadCSTC(code, pageNumber, endPageNumber, data, function (dataCSTC) {
        util_excel.addDataToExcel(util_craw_vietstock.calculateData(dataCSTC, false), 'CSTC');
        data = [];
        loadCDKT(code, pageNumber, endPageNumber, data, function (dataCDKT) {
            util_excel.addDataToExcel(util_craw_vietstock.calculateData(dataCDKT, false), 'CDKT');
            data = [];
            loadKQKD(code, pageNumber, endPageNumber, data, function (dataKQKD) {
                util_excel.addDataToExcel(util_craw_vietstock.calculateData(dataKQKD, true), 'KQKD');
                data = [];
                loadLCTT(code, pageNumber, endPageNumber, data, function (dataLCTT) {
                    util_excel.addDataToExcel(util_craw_vietstock.calculateData(dataLCTT, false), 'LCTT');

                    util_excel.writeDataToExcel('./report/PTBCTC-' + code + '.xlsx');
                })
            })
        })
    })
}

loadPTBCTC('vne');