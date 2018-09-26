var request = require('request');
const cheerio = require('cheerio');

request('http://finance.vietstock.vn/Controls/Report/Data/GetReport.ashx?rptType=CSTC&scode=DHA&bizType=1&rptUnit=1000000&rptTermTypeID=2&page=1', function (error, response, body) {
    if (!error && response.statusCode == 200) {
        var $ = cheerio.load(body, {
            xmlMode: true
        });
        $('table thead tr#BR_rowHeader td.BR_colHeader_Time').each(function (i, element){
            console.log($(this).text());
        });

        $('table tbody tr.0.BR_tBody_rowName').each(function (i, elem){
            var tr = $(this);
            var nameRow = tr.find('td.BR_tBody_colName.Padding1').text();
            console.log('Name:',nameRow);
            tr.find('td.BR_tBody_colValue').each(function (sub, subE){
                console.log($(this).text());
            });
        });
    }
});