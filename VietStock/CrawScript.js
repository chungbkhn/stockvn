var util_craw_vietstock = require('./Util/Util+Craw+VietStock.js');
var util_excel = require('./Util/Util+Excel.js');

function loadCSTC(code, pageNumber, endPageNumber, termType, data, callback) {
    util_craw_vietstock.combineLoadCSTC(code, pageNumber, termType, data, function (newData) {
        pageNumber++;
        if (pageNumber <= endPageNumber) {
            loadCSTC(code, pageNumber, endPageNumber, termType, newData, callback);
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

function loadPriceHistory(code, callback) {
    util_craw_vietstock.loadPriceHistory(code, '01/01/13', '10/20/18', callback);
}

function loadPTBCTC(code) {
    var pageNumber = 1;
    var endPageNumber = 10;
    var data = [];

    loadCDKT(code, pageNumber, endPageNumber, data, function (dataCDKT) {
        const calculateDataCDKT = util_craw_vietstock.calculateData(dataCDKT, false);
        util_excel.addDataForPTCDKT(calculateDataCDKT);
        data = [];
        loadKQKD(code, pageNumber, endPageNumber, data, function (dataKQKD) {
            const calculateDataKQKD = util_craw_vietstock.calculateData(dataKQKD, true);
            util_excel.addDataForPTKQKD(calculateDataKQKD);
            data = [];
            loadLCTT(code, pageNumber, endPageNumber, data, function (dataLCTT) {
                const calculateDataLCTT = util_craw_vietstock.calculateData(dataLCTT, false);
                util_excel.addDataForPTLCTT(calculateDataLCTT);
                data = [];
                loadCSTC(code, pageNumber, endPageNumber, 1, data, function (dataCSTCYear) {
                    util_excel.addDataToExcel(dataCSTCYear, 'CSTC - Năm');
                    data = [];
                    loadCSTC(code, pageNumber, endPageNumber, 2, data, function (dataCSTCQuater) {
                        util_excel.addDataToExcel(dataCSTCQuater, 'CSTC - Quý');

                        loadPriceHistory(code, function (dataPrice) {
                            util_excel.addDataForPE(dataPrice);
                            util_excel.addDataForDLDT();

                            util_excel.writeToFileExcel('./report/PTBCTC-' + code + '.xlsx');
                        })
                    })
                })
            })
        })
    })
}

loadPTBCTC('HPG');