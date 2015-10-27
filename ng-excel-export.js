(function (){
    'use strict';

    var app = angular.module('ngExcelExport');

    app.service('ExcelExport', ['Blob', 'FileSaver',
        function(Blob, FileSaver) {
            var XLSX = window.XLSX;
            function Workbook() {
                if (!(this instanceof Workbook)){
                    return new Workbook();
                }
                this.SheetNames = [];
                this.Sheets = {};
            }

            function datenum(v, date1904) {
                if (date1904){
                    v += 1462;
                }
                var epoch = Date.parse(v);
                return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
            }

            var s2ab = function (s) {
                var buf = new ArrayBuffer(s.length);
                var view = new Uint8Array(buf);
                for (var i = 0; i !== s.length; ++i){
                    view[i] = s.charCodeAt(i) & 0xFF;
                }
                return buf;
            };

            this.export = function (data, wsName, colsData, fileName) {
                var wsCols = [];
                colsData.forEach( function(cd){
                    var dataSize = (cd.size === undefined ? 8 : cd.size);
                    wsCols = wsCols.concat({
                        wch: dataSize
                    });
                });

                var wb = new Workbook(),
                    ws = data;
                wb.SheetNames.push(wsName);
                wb.Sheets[wsName] = ws;
                ws['!cols'] = wsCols;
                var wbOut = XLSX.write(wb, {
                    bookType: 'xlsx',
                    bookSST: false,
                    type: 'binary'
                });

                var excelObj = new Blob([s2ab(wbOut)], {
                    type: 'application/octet-stream'
                });

                var config = {
                    data: excelObj,
                    filename: fileName
                };

                FileSaver.saveAs(config);
            };
        }
    ]);
})();