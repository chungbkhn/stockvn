var util = require('./UtilVietStock.js');

function loadCSTC(code, pageNumber, endPageNumber, data, callback) {
    util.combineLoadCSTC(code, pageNumber, data, function (newData) {
        pageNumber++;
        if (pageNumber <= endPageNumber) {
            loadCSTC(code, pageNumber, endPageNumber, newData, callback);
        } else {
            callback(newData);
        }
    })
}

function loadCDKT(code, pageNumber, endPageNumber, data, callback) {
    util.combineLoadCDKT(code, pageNumber, data, function (newData) {
        pageNumber++;
        if (pageNumber <= endPageNumber) {
            loadCDKT(code, pageNumber, endPageNumber, newData, callback);
        } else {
            callback(newData);
        }
    })
}

function loadKQKD(code, pageNumber, endPageNumber, data, callback) {
    util.combineLoadKQKD(code, pageNumber, data, function (newData) {
        pageNumber++;
        if (pageNumber <= endPageNumber) {
            loadKQKD(code, pageNumber, endPageNumber, newData, callback);
        } else {
            callback(newData);
        }
    })
}

function loadLCTT(code, pageNumber, endPageNumber, data, callback) {
    util.combineLoadLCTT(code, pageNumber, data, function (newData) {
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
        util.addDataToExcel(util.calculateData(dataCSTC, false), 'CSTC');
        data = [];
        loadCDKT(code, pageNumber, endPageNumber, data, function (dataCDKT) {
            util.addDataToExcel(util.calculateData(dataCDKT, false), 'CDKT');
            data = [];
            loadKQKD(code, pageNumber, endPageNumber, data, function (dataKQKD) {
                util.addDataToExcel(util.calculateData(dataKQKD, true), 'KQKD');
                data = [];
                loadLCTT(code, pageNumber, endPageNumber, data, function (dataLCTT) {
                    util.addDataToExcel(util.calculateData(dataLCTT, false), 'LCTT');

                    util.writeDataToExcel('./report/PTBCTC-' + code + '.xlsx');
                })
            })
        })
    })
}

loadPTBCTC('vne');