var xl = require('excel4node');
const format = require('string-format');

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

function addRowFormula(formualaTemplate, row, startCol, numOfColumnData, ws, style = styleNumber) {
    for (let idx = startCol; idx < numOfColumnData + startCol; idx++) {
        const columnName = xl.getExcelAlpha(idx);
        const formula = format(formualaTemplate, columnName);
        ws.cell(row, idx)
        .formula(formula)
        .style(style);
    }
}

function addRowGFormula(formualaTemplate, row, startCol, numOfColumnData, ws, style = styleNumber) {
    for (let idx = startCol; idx < numOfColumnData + startCol; idx++) {
        if (idx < 7) {
            addCell(0, row, idx, ws);
            continue;
        }

        const columnName = xl.getExcelAlpha(idx);
        const columnNamePrevious = xl.getExcelAlpha(idx - 5);
        const formula = format(formualaTemplate, columnName, columnNamePrevious);
        ws.cell(row, idx)
        .formula(formula)
        .style(style);
    }
}

function addRowSum(title, dataTitle1, dataTitle2, startColumnData, data, row, ws, style = styleNumber) {
    addCell(title, row, 1, ws);
    var dataRowIndex1 = rowIndexConstainTitle(dataTitle1, data);
    var dataRowIndex2 = rowIndexConstainTitle(dataTitle2, data);

    var col = 2;
    for (let idx = startColumnData; idx < data[dataRowIndex1].length; idx++) {
        const value = Number(data[dataRowIndex1][idx]) + Number(data[dataRowIndex2][idx]);
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

        const numOfColumnData = data[0].length - startColumnData;

        // Write Row Title
        rowIndex++;
        addCell('', rowIndex, 1, ws);
        addRowDataWithStartColumn(startColumnData, data[0], rowIndex, 2, ws);

        // VCSH / I. Vốn chủ sở hữu
        rowIndex++;
        addRow('VCSH', 'I. Vốn chủ sở hữu', startColumnData, data, rowIndex, ws);

        // Số lượng CP
        rowIndex++;
        addCell('Số lượng CP', rowIndex, 1, ws);
        addRowFormula('{0}6*100',rowIndex, 2, numOfColumnData, ws);

        // Nợ phải trả / A. NỢ PHẢI TRẢ
        rowIndex++;
        addRow('Tổng nợ', 'A. NỢ PHẢI TRẢ', startColumnData, data, rowIndex, ws);

        // Tổng nợ vay
        rowIndex++;
        addRowSum('Tổng nợ vay', 'Vay và nợ thuê tài chính ngắn hạn', 'Vay và nợ thuê tài chính dài hạn', startColumnData, data, rowIndex, ws);

        // Tổng tài sản
        rowIndex++;
        addCell('Tổng tài sản', rowIndex, 1, ws);
        addRowFormula('{0}6+{0}8',rowIndex, 2, numOfColumnData, ws);

        // Chiếm dụng vốn
        rowIndex++;
        addCell('Chiếm dụng vốn', rowIndex, 1, ws);
        addRowFormula('{0}8-{0}9',rowIndex, 2, numOfColumnData, ws);

        // Bị chiếm dụng vốn = (Các khoản phải thu ngắn hạn) + (Các khoản phải thu dài hạn)
        rowIndex++;
        addRowSum('Bị chiếm dụng vốn', 'Các khoản phải thu ngắn hạn', 'Các khoản phải thu dài hạn', startColumnData, data, rowIndex, ws);

        // Tổng nợ/Tổng TS
        rowIndex++;
        addCell('Tổng nợ/Tổng TS', rowIndex, 1, ws);
        addRowFormula('{0}8/{0}10',rowIndex, 2, numOfColumnData, ws, stylePercent);

        // Nợ vay/VCSH
        rowIndex++;
        addCell('Nợ vay/VCSH', rowIndex, 1, ws);
        addRowFormula('{0}9/{0}6',rowIndex, 2, numOfColumnData, ws, stylePercent);

        // Tỉ lệ chiếm dụng vốn
        rowIndex++;
        addCell('Tỉ lệ chiếm dụng vốn', rowIndex, 1, ws);
        addRowFormula('{0}11/{0}12',rowIndex, 2, numOfColumnData, ws, stylePercent);

        // Cổ tức
        rowIndex++;
        addCell('Cổ tức', rowIndex, 1, ws);
    },
    addDataForPTKQKD: function (data, sheetName) {
        if (data.length == 0) { return; }

        var ws = wb.addWorksheet(sheetName);
        var rowIndex = 1;
        ws.cell(rowIndex++, 1)   // Write to cell A1
            .string('KQKD, HIỆU QUẢ VÀ KHẢ NĂNG SINH LỜI')
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

        const numOfColumnData = data[0].length - startColumnData;

        // Write Row Title
        rowIndex++;
        addCell('', rowIndex, 1, ws);
        addRowDataWithStartColumn(startColumnData, data[0], rowIndex, 2, ws);

        // Doanh thu thuần / 3. Doanh thu thuần về bán hàng và cung cấp dịch vụ
        rowIndex++;
        addRow('Doanh thu thuần', '3. Doanh thu thuần về bán hàng và cung cấp dịch vụ', startColumnData, data, rowIndex, ws);
        
        // LNG / 5. Lợi nhuận gộp về bán hàng và cung cấp dịch vụ
        rowIndex++;
        addRow('LNG', '5. Lợi nhuận gộp về bán hàng và cung cấp dịch vụ', startColumnData, data, rowIndex, ws);

        // LNR / 11. Lợi nhuận thuần từ hoạt động kinh doanh
        rowIndex++;
        addRow('LNR', '11. Lợi nhuận thuần từ hoạt động kinh doanh', startColumnData, data, rowIndex, ws);

        // LNTT / 15. Tổng lợi nhuận kế toán trước thuế
        rowIndex++;
        addRow('LNTT', '15. Tổng lợi nhuận kế toán trước thuế', startColumnData, data, rowIndex, ws);

        // LNST / Lợi nhuận sau thuế của cổ đông của Công ty mẹ
        rowIndex++;
        addRow('LNST', 'Lợi nhuận sau thuế của cổ đông của Công ty mẹ', startColumnData, data, rowIndex, ws);

        // EPS / 19. Lãi cơ bản trên cổ phiếu (*) (VNÐ)
        rowIndex++;
        addRow('EPS', '19. Lãi cơ bản trên cổ phiếu (*) (VNÐ)', startColumnData, data, rowIndex, ws);

        // Biên LNG
        rowIndex++;
        addCell('Biên LNG', rowIndex, 1, ws);
        addRowFormula('{0}7/{0}6',rowIndex, 2, numOfColumnData, ws, stylePercent);

        // Biên LNR
        rowIndex++;
        addCell('Biên LNR', rowIndex, 1, ws);
        addRowFormula('{0}8/{0}6',rowIndex, 2, numOfColumnData, ws, stylePercent);

        // G DT
        rowIndex++;
        addCell('G DT', rowIndex, 1, ws);
        addRowGFormula('{0}6/{1}6 - 1',rowIndex, 2, numOfColumnData, ws, stylePercent);

        // G LNG
        rowIndex++;
        addCell('G LNG', rowIndex, 1, ws);
        addRowGFormula('{0}7/{1}7 - 1',rowIndex, 2, numOfColumnData, ws, stylePercent);

        // G LNR
        rowIndex++;
        addCell('G LNR', rowIndex, 1, ws);
        addRowGFormula('{0}8/{1}8 - 1',rowIndex, 2, numOfColumnData, ws, stylePercent);
    }
};

module.exports = util_excel;