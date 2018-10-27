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

var styleDate = wb.createStyle({
    font: {
        color: '#000000',
        size: 12,
    },
    numberFormat: 'DD/MM/yyyy'
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

function addDLTC_KQKD(ws, quater, row, isYear, referenceSheetName, listRowTitles) {
    var rowIndex = row;
    var colIndex = 1;

    var title = 'Quý ' + quater;
    if (isYear) {
        title = 'Cả Năm'
    }
    ws.cell(rowIndex++, colIndex)   // Write to cell C2
        .string(title)
        .style(styleNumber);

    var startRow = rowIndex;
    ws.cell(rowIndex++, colIndex)   // Write to cell A5
        .string('Tên chỉ số')
        .style(styleNumber);

    for (let idx = 0; idx < listRowTitles.length; idx++) {
        const rowTitle = listRowTitles[idx][0];
        ws.cell(rowIndex++, colIndex)   // Write to cell A6
            .string(rowTitle)
            .style(styleNumber);
    }

    rowIndex = startRow;

    for (let addRowIndex = 0; addRowIndex < listRowTitles.length + 1; addRowIndex++) {
        colIndex = 2;
        for (let idx = 2013; idx < 2031; idx++) {
            if (addRowIndex == 0) {
                var headerTitle = 'Quý ' + quater + '/' + idx;
                if (isYear) {
                    headerTitle = 'Năm ' + idx
                }
                ws.cell(rowIndex, colIndex++)
                    .string(headerTitle)
                    .style(styleNumber);
            } else {
                const rowStyle = listRowTitles[addRowIndex - 1][1];
                const columnName = xl.getExcelAlpha(colIndex);
                ws.cell(rowIndex, colIndex++)
                    .formula('VLOOKUP($A' + rowIndex + ',' + referenceSheetName + '!$A$5:$BW$20,MATCH(' + columnName + '$' + startRow + ',' + referenceSheetName + '!$A$5:$BW$5,0),FALSE)')
                    .style(rowStyle);
            }
        }
        rowIndex++;
    }

    rowIndex++;
    return rowIndex;
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

        console.log('Add data to sheet ' + sheetName + ' successful!');
    },
    addDataForPTCDKT: function (data) {
        if (data.length == 0) { return; }

        const sheetName = 'SKTC';
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
        addRowFormula('{0}6*100', rowIndex, 2, numOfColumnData, ws);

        // Nợ phải trả / A. NỢ PHẢI TRẢ
        rowIndex++;
        addRow('Tổng nợ', 'A. NỢ PHẢI TRẢ', startColumnData, data, rowIndex, ws);

        // Tổng nợ vay
        rowIndex++;
        addRowSum('Tổng nợ vay', 'Vay và nợ thuê tài chính ngắn hạn', 'Vay và nợ thuê tài chính dài hạn', startColumnData, data, rowIndex, ws);

        // Tổng tài sản
        rowIndex++;
        addCell('Tổng tài sản', rowIndex, 1, ws);
        addRowFormula('{0}6+{0}8', rowIndex, 2, numOfColumnData, ws);

        // Chiếm dụng vốn
        rowIndex++;
        addCell('Chiếm dụng vốn', rowIndex, 1, ws);
        addRowFormula('{0}8-{0}9', rowIndex, 2, numOfColumnData, ws);

        // Bị chiếm dụng vốn = (Các khoản phải thu ngắn hạn) + (Các khoản phải thu dài hạn)
        rowIndex++;
        addRowSum('Bị chiếm dụng vốn', 'Các khoản phải thu ngắn hạn', 'Các khoản phải thu dài hạn', startColumnData, data, rowIndex, ws);

        // Tổng nợ/Tổng TS
        rowIndex++;
        addCell('Tổng nợ/Tổng TS', rowIndex, 1, ws);
        addRowFormula('{0}8/{0}10', rowIndex, 2, numOfColumnData, ws, stylePercent);

        // Nợ vay/VCSH
        rowIndex++;
        addCell('Nợ vay/VCSH', rowIndex, 1, ws);
        addRowFormula('{0}9/{0}6', rowIndex, 2, numOfColumnData, ws, stylePercent);

        // Tỉ lệ chiếm dụng vốn
        rowIndex++;
        addCell('Tỉ lệ chiếm dụng vốn', rowIndex, 1, ws);
        addRowFormula('{0}11/{0}12', rowIndex, 2, numOfColumnData, ws, stylePercent);

        // Cổ tức
        rowIndex++;
        addCell('Cổ tức', rowIndex, 1, ws);

        console.log('Add data to sheet ' + sheetName + ' successful!');
    },
    addDataForPTKQKD: function (data) {
        if (data.length == 0) { return; }

        const sheetName = 'KQKD';
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

        // DT thuần / 3. Doanh thu thuần về bán hàng và cung cấp dịch vụ
        rowIndex++;
        addRow('DT thuần', '3. Doanh thu thuần về bán hàng và cung cấp dịch vụ', startColumnData, data, rowIndex, ws);

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
        addRowFormula('{0}7/{0}6', rowIndex, 2, numOfColumnData, ws, stylePercent);

        // Biên LNR
        rowIndex++;
        addCell('Biên LNR', rowIndex, 1, ws);
        addRowFormula('{0}8/{0}6', rowIndex, 2, numOfColumnData, ws, stylePercent);

        // ROE
        rowIndex++;
        addCell('ROE', rowIndex, 1, ws);
        addRowGFormula('{0}8/\'SKTC\'!{1}6', rowIndex, 2, numOfColumnData, ws, stylePercent);

        // G DT
        rowIndex++;
        addCell('G DT', rowIndex, 1, ws);
        addRowGFormula('{0}6/{1}6 - 1', rowIndex, 2, numOfColumnData, ws, stylePercent);

        // G LNG
        rowIndex++;
        addCell('G LNG', rowIndex, 1, ws);
        addRowGFormula('{0}7/{1}7 - 1', rowIndex, 2, numOfColumnData, ws, stylePercent);

        // G LNR
        rowIndex++;
        addCell('G LNR', rowIndex, 1, ws);
        addRowGFormula('{0}8/{1}8 - 1', rowIndex, 2, numOfColumnData, ws, stylePercent);

        console.log('Add data to sheet ' + sheetName + ' successful!');
    },
    addDataForPTLCTT: function (data) {
        if (data.length == 0) { return; }

        const sheetName = 'LCTT';
        var ws = wb.addWorksheet(sheetName);
        var rowIndex = 1;
        ws.cell(rowIndex++, 1)   // Write to cell A1
            .string('Dòng tiền')
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
            if (item.indexOf('Năm') > -1) {
                startColumnData = col;
                break;
            }
        }

        const numOfColumnData = data[0].length - startColumnData;

        // Write Row Title
        rowIndex++;
        addCell('', rowIndex, 1, ws);
        addRowDataWithStartColumn(startColumnData, data[0], rowIndex, 2, ws);

        // HĐ SXKD / Lưu chuyển tiền thuần từ hoạt động kinh doanh
        rowIndex++;
        addRow('HĐ SXKD', 'Lưu chuyển tiền thuần từ hoạt động kinh doanh', startColumnData, data, rowIndex, ws);

        // HĐ ĐT / Lưu chuyển tiền thuần từ hoạt động đầu tư
        rowIndex++;
        addRow('HĐ ĐT', 'Lưu chuyển tiền thuần từ hoạt động đầu tư', startColumnData, data, rowIndex, ws);

        // HĐ TC / Lưu chuyển tiền thuần từ hoạt động tài chính
        rowIndex++;
        addRow('HĐ TC', 'Lưu chuyển tiền thuần từ hoạt động tài chính', startColumnData, data, rowIndex, ws);

        // Thuần / Lưu chuyển tiền thuần trong kỳ
        rowIndex++;
        addRow('Thuần', 'Lưu chuyển tiền thuần trong kỳ', startColumnData, data, rowIndex, ws);

        console.log('Add data to sheet ' + sheetName + ' successful!');
    },
    addDataForPE: function (data) {
        const sheetName = 'P-E';
        if (data.length == 0) { return; }

        var ws = wb.addWorksheet(sheetName);
        var rowIndex = 1;
        var colIndex = 1;
        var startRowData = 0;
        var endRowData = 0;
        ws.cell(rowIndex++, colIndex)   // Write to cell A1
            .string('P/E tổng hợp')
            .style(styleNumber);

        ws.cell(rowIndex++, colIndex)
            .string('Std Dev P/E')
            .style(styleNumber);

        ws.cell(rowIndex++, colIndex)
            .string('AVG P/E')
            .style(styleNumber);

        ws.cell(rowIndex, colIndex++)   // Write to cell A2
            .string('Ngày')
            .style(styleNumber);

        ws.cell(rowIndex, colIndex++)   // Write to cell B2
            .string('Năm')
            .style(styleNumber);

        ws.cell(rowIndex, colIndex++)   // Write to cell C2
            .string('VHTT')
            .style(styleNumber);

        ws.cell(rowIndex, colIndex++)   // Write to cell C2
            .string('LNST')
            .style(styleNumber);

        ws.cell(rowIndex, colIndex++)   // Write to cell C2
            .string('P/E')
            .style(styleNumber);

        ws.cell(rowIndex, colIndex++)   // Write to cell C2
            .string('Bottom')
            .style(styleNumber);

        ws.cell(rowIndex, colIndex++)   // Write to cell C2
            .string('Top')
            .style(styleNumber);

        ws.cell(rowIndex, colIndex++)   // Write to cell C2
            .string('AVG')
            .style(styleNumber);

        rowIndex++;
        startRowData = rowIndex;
        const headerRow = rowIndex - 1;
        // Write Row Title
        for (let idx = 0; idx < data.length; idx++) {
            const item = data[idx];

            colIndex = 1;
            ws.cell(rowIndex, colIndex++)
                .date(item[0])
                .style(styleDate);

            // const columnName = xl.getExcelAlpha(idx);
            ws.cell(rowIndex, colIndex++)
                .string('Năm ' + item[0].getFullYear())
                .style(styleNumber);

            ws.cell(rowIndex, colIndex++)
                .number(item[1])
                .style(styleNumber);

            // LNST
            ws.cell(rowIndex, colIndex++)
                .formula('VLOOKUP($D$' + headerRow + ',KQKD!$A$5:$BW$20,MATCH(B' + rowIndex + ',KQKD!$A$5:$BW$5,0),FALSE)')
                .style(styleNumber);

            const columnVHTT = xl.getExcelAlpha(colIndex - 2);
            const columnLNST = xl.getExcelAlpha(colIndex - 1);
            ws.cell(rowIndex, colIndex++)
                .formula(columnVHTT + rowIndex + '/' + columnLNST + rowIndex)
                .style(styleNumber);

            rowIndex++;
        }
        endRowData = rowIndex - 1;
        ws.cell(2, 2)
            .formula('STDEV($E$' + startRowData + ':$E$' + endRowData + ')')
            .style(styleNumber);
        ws.cell(3, 2)
            .formula('AVERAGE($E$' + startRowData + ':$E$' + endRowData + ')')
            .style(styleNumber);

            ws.cell(startRowData, 6)
            .formula('$B$3-$B$2')
            .style(styleNumber);
            ws.cell(endRowData, 6)
            .formula('$B$3-$B$2')
            .style(styleNumber);

            ws.cell(startRowData, 7)
            .formula('$B$3+$B$2')
            .style(styleNumber);
            ws.cell(endRowData, 7)
            .formula('$B$3+$B$2')
            .style(styleNumber);

            ws.cell(startRowData, 8)
            .formula('$B$3')
            .style(styleNumber);
            ws.cell(endRowData, 8)
            .formula('$B$3')
            .style(styleNumber);

        console.log('Add data to sheet ' + sheetName + ' successful!');
    },
    addDataForDLDT: function () {
        const sheetName = 'Dữ liệu đồ thị'
        var ws = wb.addWorksheet(sheetName);
        var rowIndex = 1;
        var colIndex = 1;
        ws.cell(rowIndex++, colIndex)   // Write to cell A1
            .string('Tổng hợp dữ liệu PT BCTC')
            .style(styleNumber);

        rowIndex++;
        ws.cell(rowIndex++, colIndex)   // Write to cell A3
            .string('Sức khoẻ tài chính')
            .style(styleNumber);

        rowIndex++;
        var listRowTitles = [['Tổng nợ/Tổng TS', stylePercent], ['Nợ vay/VCSH', stylePercent], ['Tỉ lệ chiếm dụng vốn', stylePercent]];
        var referenceSheetName = 'SKTC';
        rowIndex = addDLTC_KQKD(ws, '', rowIndex, true, referenceSheetName, listRowTitles);

        colIndex = 1;
        rowIndex = 15;
        var startRowPanel = rowIndex;
        const stepRowPanel = 15;
        ws.cell(rowIndex++, colIndex)   // Write to cell C2
            .string('Kết quả kinh doanh')
            .style(styleNumber);

        rowIndex++
        listRowTitles = [['DT thuần', styleNumber], ['LNG', styleNumber], ['EPS', styleNumber], ['Biên LNG', stylePercent], ['ROE', stylePercent], ['Biên LNR', stylePercent]];
        referenceSheetName = 'KQKD';
        rowIndex = addDLTC_KQKD(ws, '1', rowIndex, false, referenceSheetName, listRowTitles);

        rowIndex = startRowPanel + stepRowPanel * 1;
        rowIndex = addDLTC_KQKD(ws, '2', rowIndex, false, referenceSheetName, listRowTitles);

        rowIndex = startRowPanel + stepRowPanel * 2;
        rowIndex = addDLTC_KQKD(ws, '3', rowIndex, false, referenceSheetName, listRowTitles);

        rowIndex = startRowPanel + stepRowPanel * 3;
        rowIndex = addDLTC_KQKD(ws, '4', rowIndex, false, referenceSheetName, listRowTitles);

        rowIndex = startRowPanel + stepRowPanel * 4;
        rowIndex = addDLTC_KQKD(ws, '', rowIndex, true, referenceSheetName, listRowTitles);

        rowIndex = startRowPanel + stepRowPanel * 5;
        ws.cell(rowIndex++, colIndex)   // Write to cell C2
            .string('Dòng tiền')
            .style(styleNumber);

        rowIndex++
        listRowTitles = [['HĐ SXKD', styleNumber], ['HĐ ĐT', styleNumber], ['HĐ TC', styleNumber], ['Thuần', styleNumber]];
        referenceSheetName = 'LCTT';
        rowIndex = addDLTC_KQKD(ws, '', rowIndex, true, referenceSheetName, listRowTitles);

        console.log('Add data to sheet ' + sheetName + ' successful!');
    }
};

module.exports = util_excel;