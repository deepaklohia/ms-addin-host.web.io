(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        //$(document).ready(function () {
        // Initialize the notification mechanism and hide it
        var element = document.querySelector('.MessageBanner');
        messageBanner = new components.MessageBanner(element);
        messageBanner.hideBanner();

        // If not using Excel 2016, use fallback logic.
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {


            // Add a click event handler for the highlight button.
            /*
            $('#btnReset').on('click', cleardata);
            */

            return;
        }
        /*
        $("#toolHeader").text("Time Tracker tool to capture start and stop time of work.(click on reset to begin)");
        $('#btnResetText').text("Reset");
        */
        // Add a click event handler for the highlight button.
        $('#btnReset').on('click', clearDataConf);
        $('#btnClearData').on('click', clearData);
        $('#btnStart').on('click', start);
        $('#btnLap').on('click', lap);
        $('#btnStop').on('click', stop);
        $("#btnCus1").on('click', () => recordLast(5));       
        $('#btnCus2').on('click', () => recordLast(10));       
        $('#btnCus3').on('click', () => {
            let cusMin = document.getElementById("txtRowInput");
            recordLast(parseInt(cusMin.value));
        });       

       // });
    };

    function clearData() {
        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {
            // Create a proxy object for the active sheet
            var sh = ctx.workbook.worksheets.getActiveWorksheet();
            var lastRow_ = sh.getUsedRange().getLastRow();
            lastRow_.load("rowIndex");

            return ctx.sync()
                .then(function () {

                    let lastRow = lastRow_.rowIndex + 10;

                    //console.log("lastrow is:" + lastRow);
                    var rng = sh.getRange("A1:D" + lastRow);
                    rng.clear();
                    rng.format.fill.clear();
                    sh.getRange("A1").values = "Case Ref#";
                    sh.getRange("B1").values = "Start Time";
                    sh.getRange("C1").values = "Stop Time";
                    sh.getRange("D1").values = "Difference";
                    rng = sh.getRange("A1:D1");
                    rng.format.font.size = 12;
                    //rng.format.fill.color = "yellow";    
                    rng.format.fill.color = "cornflowerblue";
                    rng.format.font.bold = true;
                    sh.getRange("AZ1").values = 2;
                    rng.format.autofitColumns();


                });

            return ctx.sync();
        })
        .catch(errorHandler);
    }

    function clearDataConf() {
        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {
            // Create a proxy object for the active sheet
            var sh = ctx.workbook.worksheets.getActiveWorksheet();
            var lastRow_ = sh.getUsedRange().getLastRow();
            lastRow_.load("rowIndex");

            return ctx.sync()
                .then(function () {
                    let lastRow = lastRow_.rowIndex;

                    if (lastRow >= 1) {
                        var title = (Number(lastRow) + 1) + " - rows will be deleted";
                        $("#modClearData .modal-body").text(title);
                        $('#modClearData').modal('show');
                    }
                    else {
                        clearData();
                    }
                });

            return ctx.sync();
        })
            .catch(errorHandler);
    }

    function checkHeader() {
        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {
            var sh = ctx.workbook.worksheets.getActiveWorksheet();
            var RangeA1 = sh.getRange("A1").load("values");
            var RangeB1 = sh.getRange("B1").load("values");
            var RangeC1 = sh.getRange("C1").load("values");
            var RangeD1 = sh.getRange("D1").load("values");

            return ctx.sync()
                .then(function () {
                    var valA1 = "";
                    var valB1 = "";
                    var valC1 = "";
                    var valD1 = "";

                    valA1 = RangeA1.values[0];
                    valB1 = RangeB1.values[0];
                    valC1 = RangeC1.values[0];
                    valD1 = RangeD1.values[0];
                    var headerMissing = 0;

                    if (valA1 == "") {
                        sh.getRange("A1").values = "Case Ref#";
                        headerMissing = 1;
                    }
                    if (valB1 == "") {
                        sh.getRange("B1").values = "Start Time";
                        headerMissing = 1;
                    }
                    if (valC1 == "") {
                        sh.getRange("C1").values = "Stop Time";
                        headerMissing = 1;
                    }
                    if (valD1 == "") {
                        sh.getRange("D1").values = "Difference";
                        headerMissing = 1;
                    }

                    if (headerMissing == 1) {
                        var rng = sh.getRange("A1:D1");
                        rng.format.font.size = 12;
                        rng.format.fill.color = "cornflowerblue";
                        rng.format.font.bold = true;
                        sh.getRange("AZ1").values = 2;
                        rng.format.autofitColumns();
                    }

                })
                .then(ctx.sync);
        })
            .catch(errorHandler);
    }


    function start() {
        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {
            checkHeader();
            var sh = ctx.workbook.worksheets.getActiveWorksheet();
            //sh.getRange("AZ1").formulas = [["=COUNTA(C:C)+1"]];
            sh.getRange("AZ1").formulas = [["=COUNTA(B:B)+1"]];
            //sh.calculate;
            //when you are loading the values you have to use sync after that 
            //to ensure that the values has been uploaded
            var refRng = sh.getRange("AZ1").load("values");

            return ctx.sync()
                .then(function () {

                    //console.log("refrng" + refRng.values);

                    var refRow = 0;
                    refRow = refRng.values[0];

                    if (refRow <= 1) { refRow = 2; }
                    var desRng = sh.getRange("B" + refRow);
                    desRng.formulas = [[currentTime()]];
                    desRng.numberFormat = [["HH:MM:SS;@"]];
                })
                .then(ctx.sync);
        })
            .catch(errorHandler);
    }

    function currentTime() {
        /*
        var dt = new Date();
        var h = dt.getHours();
        var m = dt.getMinutes();
        var s = dt.getSeconds();
        return "=TIME(" + h + "," + m + "," + s + ")";
        */
        var dt = new Date();
        var yy = dt.getFullYear();
        var mm = dt.getMonth() + 1;
        var dd = dt.getDate();
        var hr = dt.getHours();
        var min = dt.getMinutes();
        var sec = dt.getSeconds();		
        return "=DATE(" + yy + "," + mm + "," + dd + ") + TIME(" + hr + "," + min + "," + sec + ")";
    }

    function currentTimeCustom(subMin) {
        var dt = new Date();
        var yy = dt.getFullYear();
        var mm = dt.getMonth() + 1;
        var dd = dt.getDate();
        var hr = dt.getHours();
        var min = dt.getMinutes();
        var sec = dt.getSeconds();
		
		//if its a next day
		if(min < subMin && hr == 0){
			dd = dd - 1;
			hr = 24;
		}
		
		return "=DATE(" + yy + "," + mm + "," + dd + ") + TIME(" + hr + "," + (min - subMin) + "," + sec + ")";	
    }


    function stop() {
        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {
            checkHeader();
            var sh = ctx.workbook.worksheets.getActiveWorksheet();
            sh.getRange("AZ1").formulas = [["=COUNTA(B:B)"]];
            var refRng = sh.getRange("AZ1").load("values");

            return ctx.sync()
                .then(function () {
                    var refRow = 0;
                    refRow = refRng.values[0];

                    if (refRow <= 1) {
                        refRow = 2;
                    }
                    var startRng = sh.getRange("B" + refRow).load("values");

                    return ctx.sync()
                        .then(function () {
                            var startVal = startRng.values[0];
                            if (startVal != "") {
                                var stopRng = sh.getRange("C" + refRow);
                                stopRng.formulas = [[currentTime()]];
                                stopRng.numberFormat = [["HH:MM:SS;@"]];
                                var diffRng = sh.getRange("D" + refRow);
                                diffRng.formulas = [["=C" + refRow + "-B" + refRow + ""]];
                                diffRng.numberFormat = [["HH:MM:SS;@"]];
                            }
                        })
                })
                .then(ctx.sync);
        })
            .catch(errorHandler);
    }


    function lap() {
        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {
            checkHeader();
            var sh = ctx.workbook.worksheets.getActiveWorksheet();
            sh.getRange("AZ1").formulas = [["=COUNTA(B:B)"]];
            //sh.calculate;
            var refRng = sh.getRange("AZ1").load("values");

            //RECORDING STOP TIME
            return ctx.sync()
                .then(function () {
                    var refRow = 0;
                    refRow = refRng.values[0];

                    if (refRow <= 1) {
                        refRow = 2;
                    }
                    var startRng = sh.getRange("B" + refRow).load("values");
                    return ctx.sync()
                        .then(function () {
                            var startVal = startRng.values[0];
                            if (startVal != "") {
                                var stopRng = sh.getRange("C" + refRow);
                                stopRng.formulas = [[currentTime()]];
                                stopRng.numberFormat = [["HH:MM:SS;@"]];
                                var diffRng = sh.getRange("D" + refRow);
                                diffRng.formulas = [["=C" + refRow + "-B" + refRow + ""]];
                                diffRng.numberFormat = [["HH:MM:SS;@"]];
                                sh.getRange("AZ1").formulas = [["= 1 + " + refRow]];

                                //NEW START
                                //sh.getRange("AZ1").formulas = [["=COUNTA(C:C)+1"]];
                                sh.getRange("AZ1").formulas = [["=COUNTA(B:B)+1"]];
                                refRng = sh.getRange("AZ1").load("values");

                                return ctx.sync()
                                    .then(function () {
                                        refRow = refRng.values[0];

                                        var desRng = sh.getRange("B" + refRow);
                                        desRng.formulas = [[currentTime()]];
                                        desRng.numberFormat = [["HH:MM:SS;@"]];
                                    })
                            }
                        })
                })
                .then(ctx.sync);
        })
            .catch(errorHandler);
    }

    function recordLast(subMin) {
            // Run a batch operation against the Excel object model
            Excel.run(function (ctx) {
                checkHeader();
                var sh = ctx.workbook.worksheets.getActiveWorksheet();
                sh.getRange("AZ1").formulas = [["=COUNTA(B:B)"]];
                var refRng = sh.getRange("AZ1").load("values");

                return ctx.sync()
                    .then(function () {
                        var refRow = 0;
                        refRow = refRng.values[0];

                        //console.log("ref row is:" + refRow);

                        if (refRow <= 1) {
                            refRow = 2;
                        }
                        else {
                            refRow = Number(refRow) + 1;
                        }
                        var startRng = sh.getRange("B" + refRow).load("values");

                        return ctx.sync()
                            .then(function () {
                                var startVal = startRng.values[0];
                                var stopRng = sh.getRange("C" + refRow);
                                stopRng.formulas = [[currentTime()]];
                                stopRng.numberFormat = [["HH:MM:SS;@"]];
                                var diffRng = sh.getRange("D" + refRow);
                                diffRng.formulas = [["=C" + refRow + "-B" + refRow + ""]];
                                diffRng.numberFormat = [["HH:MM:SS;@"]];

                                //NEW START.
                                var desRng = sh.getRange("B" + refRow);
                                desRng.formulas = [[currentTimeCustom(subMin)]];
                                desRng.numberFormat = [["HH:MM:SS;@"]];
                          
                                /*
                                return ctx.sync()
                                    .then(function () {      
                                    })
                                 */

                            })
                     
                    }).then(ctx.sync);

            }).catch(errorHandler);
        }
 

    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }

})();