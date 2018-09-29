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

function rowIndexWithTitle(title, data) {
    for (let idx = 0; idx < data.length; idx++) {
        const rowData = data[idx];
        if (rowData[0] == title) {
            return idx;
        }
    }
    return -1;
}

function rowIndexConstainTitle(title, data) {
    for (let idx = 0; idx < data.length; idx++) {
        const rowData = data[idx];
        if (rowData[0].indexOf(title) > -1) {
            return idx;
        }
    }
    return -1;
}

function addRow(title, dataTitle, startColumnData, data, row, ws) {
    addCell(title, row, 1, ws);
    var dataRowIndex = rowIndexWithTitle(dataTitle, data);
    if (dataRowIndex < 0) { 
        console.log('Can find row with Title: ' + dataTitle);
        return; 
    }
    addRowDataWithStartColumn(startColumnData, data[dataRowIndex], row, 2, ws);
}

function addRowConstainTitle(title, dataTitle, startColumnData, data, row, ws) {
    addCell(title, row, 1, ws);
    var dataRowIndex = rowIndexConstainTitle(dataTitle, data);
    addRowDataWithStartColumn(startColumnData, data[dataRowIndex], row, 2, ws);
}

function addRowDataWithStartColumn(startIndex, addRowData, row, startCol, ws) {
    if (startIndex >= addRowData.length || addRowData.length == 0) { return; }

    var col = startCol;
    for (let idx = startIndex; idx < addRowData.length; idx++) {
        const value = addRowData[idx];
        addCell(value, row, col, ws);
        col++;
    }
}

function addCell(value, row, col, ws) {
    var number = Number(value);
    if (isNaN(number)) {
        ws.cell(row, col)
            .string(value)
            .style(styleNumber);
    } else {
        ws.cell(row, col)
            .number(number)
            .style(styleNumber);
    }
}

var util_excel = {
    writeToFileExcel: function (name) {
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
    },
    addDataForPTCDKT: function (data, sheetName) {
        if (data.length == 0) { return; }

        var ws = wb.addWorksheet(sheetName);
        var rowIndex = 1;
        ws.cell(rowIndex++, 1)   // Write to cell A1
            .string('Sức khoẻ tài chính')
            .style(styleNumber);

        ws.cell(rowIndex++, 1)   // Write to cell A2
            .string('Đơn vị tính')
            .style(styleNumber);

        ws.cell(rowIndex++, 1)   // Write to cell A3
            .number(1000000)
            .style(styleNumber);

        var startColumnData = -1;
        for (let col = 0; col < data[0].length; col++) {
            const item = data[0][col];
            if (item.indexOf('Quý 1') > -1) {
                startColumnData = col;
                break;
            }
        }

        // Write Row Title
        rowIndex++;
        addCell('', rowIndex, 1, ws);
        addRowDataWithStartColumn(startColumnData, data[0], rowIndex, 2, ws);

        // VCSH / I. Vốn chủ sở hữu
        rowIndex++;
        addRow('VCSH', 'I. Vốn chủ sở hữu', startColumnData, data, rowIndex, ws);

        // Vốn đầu tư CSH / 1. Vốn góp của chủ sở hữu
        rowIndex++;
        addRow('Vốn đầu tư CSH', '1. Vốn góp của chủ sở hữu', startColumnData, data, rowIndex, ws);

        // Số lượng CP
        rowIndex++;

        // Nợ phải trả / A. NỢ PHẢI TRẢ
        rowIndex++;
        addRow('Nợ phải trả', 'A. NỢ PHẢI TRẢ', startColumnData, data, rowIndex, ws);

        // Nợ ngắn hạn / I. Nợ ngắn hạn
        rowIndex++;
        addRow('Nợ ngắn hạn', 'I. Nợ ngắn hạn', startColumnData, data, rowIndex, ws);

        // Nợ dài hạn / II. Nợ dài hạn 
        rowIndex++;
        addRowConstainTitle('Nợ dài hạn', 'II. Nợ dài hạn', startColumnData, data, rowIndex, ws);

        // Tổng nợ vay
        rowIndex++;

        // Nợ vay ngắn hạn / (Vay và nợ thuê tài chính ngắn hạn)
        rowIndex++;
        addRowConstainTitle('Nợ vay ngắn hạn', 'Vay và nợ thuê tài chính ngắn hạn', startColumnData, data, rowIndex, ws);

        // Nợ vay dài hạn / (Vay và nợ thuê tài chính dài hạn)
        rowIndex++;
        addRowConstainTitle('Nợ vay dài hạn', 'Vay và nợ thuê tài chính dài hạn', startColumnData, data, rowIndex, ws);

        // Tổng tài sản
        rowIndex++;

        // Chiếm dụng vốn

        // Tổng nợ/Tổng TS

        // Nợ vay/VCSH

        // Tỉ lệ chiếm dụng vốn

        // Cổ tức
    }
};

module.exports = util_excel;