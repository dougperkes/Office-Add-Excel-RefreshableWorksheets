/// <reference path="../App.js" />
/// <reference path="../../Scripts/jquery-2.2.0.js" />

(function () {
    "use strict";
    var serviceRoot = "https://services.odata.org/V3/Northwind/Northwind.svc"
    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            // If not using Excel 2016, return
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                app.showNotification("Need Office 2016 or greater", "Sorry, this app only works with newer versions of Excel.");
                return;
            }
            loadEntityData();
        });
    };

    // Reads data from current document selection and displays a notification
    function loadEntityData() {
        $("#waitingMessage").html("Loading entities list");
        $.getJSON(serviceRoot + "?$format=json", function (data) {
            $("#entityList").empty();
            $.each(data.value, function (key, val) {
                var listItem = $("<li>" + val.name + "</li>");
                listItem.click(function () { loadDataSet(val.url); });
                $("#entityList").append(listItem);
            })


            $("#waitingContainer").hide("fast");
            $("#entityListContainer").show("fast");
        });
    }

    function loadDataSet(dataSetPath) {
        $("#entityListContainer").hide("fast");
        $("#waitingContainer").show("fast");
        $("#waitingMessage").html("Loading entity data");

        $.getJSON(serviceRoot + "/" + dataSetPath, function (data) {
            displayDataTable(dataSetPath, data);
        });
    }

    function displayDataTable(dataSetPath, data) {


        var cols = [], rows = new Array([]);
        for (var propertyName in data.value[0]) {
            cols.push(propertyName);
        }
        rows[0] = cols;
        for (var i = 0; i < data.value.length; i++) {
            var row = [];
            for (var propertyName in data.value[i]) {
                row.push(data.value[i][propertyName]);
            }
            rows.push(row);
        }

        Excel.run(function (ctx) {
            var sheetName = dataSetPath;
            ctx.workbook.worksheets.load("items");
            ctx.workbook.worksheets.load("name")
            return ctx.sync().then(function () {
                var sheet = findWorksheet(ctx.workbook.worksheets, sheetName);
                if (!sheet) {
                    sheet = ctx.workbook.worksheets.add(sheetName);
                }
                sheet.activate();
                sheet.getRange().clear();

                var firstCell = sheet.getCell(1, 0);
                var lastCell = sheet.getCell(rows.length, rows[0].length - 1);

                var range = firstCell.getBoundingRect(lastCell).insert('down');
                range.values = rows;
                range.load('address');

                return ctx.sync().then(function () {
                    //need to get the range address loaded before continuing
                    var table = sheet.tables.add(range.address, true);

                    // Queue commands to set the title in the sheet and format it
                    var pageTitle = sheet.getRange("A1:A1");
                    pageTitle.values = dataSetPath;
                    pageTitle.format.font.name = "Segoe UI Light";
                    pageTitle.format.font.bold = true;
                    pageTitle.format.font.size = 28;

                    $("#waitingContainer").hide("fast");
                    $("#entityListContainer").show("fast");
                }).catch(function (error) {
                    app.showNotification("Error", JSON.stringify(error))
                    console.log("Error: " + error);
                    if (error instanceof OfficeExtension.Error) {
                        console.log("Debug info: " + JSON.stringify(error.debugInfo));
                    }
                    $("#waitingContainer").hide("fast");
                    $("#entityListContainer").show("fast");
                });
            });
        });
    }


    function findWorksheet(worksheets, name) {
        var sheet = null;
        for (var i = 0; i < worksheets.items.length; i++) {
            if (worksheets.items[i].name === name)
                return worksheets.items[i];
        }
        return sheet;
    }
})();