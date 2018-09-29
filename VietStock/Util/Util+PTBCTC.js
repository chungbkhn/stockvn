var util_excel = require('./Util/Util+Excel.js');

function rowIndexWithTitle(title, data) {
    for (let idx = 0; idx < data.length; idx++) {
        const rowData = data[idx];
        if (rowData[0] == title) {
            return idx;
        }
    }
    return -1;
}

function addRowDataWithStartColumn(rowData, idx, addRowData) {
    if (addRowData.length == 0) { return }

    for (let col = idx; col < rowData.length; col++) {
        rowData.push(addRowData[col]);
    }
}

var util_ptbctc = {
    calculateDataForCDKT: function (data) {
        if (data.length == 0) { return; }

        var startColumnData = -1;
        for (let col = 0; col < listTitle.length; col++) {
            const item = listTitle[col];
            if (item.indexOf('Quý 1') > -1) { 
                startColumnData = col;
                break;
            }
        }

        var newData = [];
        var rowData = [];
        rowData.push('');
        addRowDataWithStartColumn(rowData, startColumnData, data[0]);
        newData.push(rowData);


    }
};

module.exports = util_ptbctc;