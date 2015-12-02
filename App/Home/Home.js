/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#generate-month').click(generateMonth);
        });
    };

    function generateMonth() {
        var days = ["M", "T", "W", "T", "F", "S", "S"];
        var weekends = [];
        var doMonthStr = $('#select-month option:selected').text();
        var doMonth = parseInt($('#select-month option:selected').val());
        var nextMonth = doMonth + 1;
        var doYear = $('#select-year option:selected').val();
        var doStartDay = new Date(doYear, doMonth, 1);
        var doLastDay = new Date(doYear, nextMonth, 0);
        var monthDays = Math.round((doLastDay - doStartDay) / (1000 * 60 * 60 * 24));
        var cellsFormattingList = [];

        /* 
            Formatting help https://msdn.microsoft.com/en-us/library/office/dn535872.aspx
        */

        // Create a TableData object.
        var rowMonth = [];
        var rowDayOfWeek = [];
        var rowNames = [];
        var myTable = new Office.TableData();
        myTable.headers = [""];
        myTable.rows = [];
        
        rowMonth.push(doMonthStr.toString());   //Month name in first column
        rowDayOfWeek = [''];                    //empty first cell in days of week
        for (var i = 0; i <= monthDays; i++)
        {
            //month row
            rowMonth.push((i + 1).toString());

            //Day ow week row
            var tempDate = new Date(doYear, doMonth, i);
            rowDayOfWeek.push(days[tempDate.getDay()]);

            //list of cellstformatting
            var cellObject = {};
            var columnObject = {};
            var formatObject = {};
            columnObject['column'] = i + 1;
            formatObject['width'] = 3;
            if (tempDate.getDay() == 5 || tempDate.getDay() == 6) { //weekdays coloring
                columnObject['row'] = 0;
                formatObject['backgroundColor'] = '#7070F0';
            }
            cellObject['cells'] = columnObject;
            cellObject['format'] = formatObject;
            cellsFormattingList.push(cellObject);
        }
                
        myTable.headers = rowMonth;
        myTable.rows.push(rowDayOfWeek);

        var names = $('#names-list').val().split('\n');
        for (var n = 0; n < names.length; n++) {
            rowNames = [];
            rowNames.push(names[n]);
            for (var i = 0; i <= monthDays; i++) {
                rowNames.push('');
            }
            myTable.rows.push(rowNames);
        }

        cellsFormattingList.push({ cells: Office.Table.All, format: { borderStyle: "thin" } });

        // Set the myTable in the document.
        Office.context.document.setSelectedDataAsync(myTable,
          {
              coercionType: Office.CoercionType.Table,
              tableOptions: { filterButton: false, style: "None" },
              cellFormat: cellsFormattingList
          },
          function (asyncResult) {
              if (asyncResult.status == "failed") {
                  app.showNotification("Action failed with error: " + asyncResult.error.message);
              } else {
                  app.showNotification("Check out your fancy new table!");
              }
          }
        );
    }
})();