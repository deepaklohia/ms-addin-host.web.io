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
        messageBanner.showBanner();
        messageBanner.hideBanner();

        // If not using Excel 2016, use fallback logic.

        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
            $("#btnCombineSheet").text("Outdated version.");
            $("#btnGetStarted").text("Outdated version.");
            //$('#button-text').text("Display!");
            //$('#button-desc').text("Display the selection");
            //$('#highlight-button').click(displaySelectedCells);
            return;
        }

        $('#btnCombineSheet').on('click', selectAndCombineConf);
        $('#btnCombineSheetYes').on('click', selectAndCombine);
        $('#btnRefresh').on('click', getSheetNames);
        $('#btnRemoveDups').on('click', funRemoveAnyDupsConf);
        $('#btnRemoveDupsYes').on('click', funRemoveAnyDups);

        $('#btnRemoveGenSheet').on('click', funcDeleteSheetConf);
        $('#btnDeleteSheet').on('click', funcDeleteSheet);
        $('#btnGetStarted').on('click', hideMain);
        $('#goToHomeLabel').on('click', hideMain);

        //var elem = document.getElementById("mainSection");
        $("#mainSection").hide();

        loadDefaults();
        //});
    };

    function loadDefaults() {
        getSheetNames(false);
        $('#shtTitle').text("Select Sheets");
        console.log("Sheets Loaded");
    }

    function getSheetNames(not = true) {
        var lst = document.getElementById('lsBxShNm');
        $("#lsBxShNm").empty();

        Excel.run(function (context) {
            var sheets = context.workbook.worksheets;
            sheets.load("name");

            return context.sync()
                .then(function () {
                    if (sheets.items.length > 1) {
                        //console.log(`There are ${sheets.items.length} worksheets in the workbook:`);
                        lst.disabled = false;

                    } else {
                        lst.disabled = true;
                        if (not) {
                            showNotification("Error", "Single Sheet can not be combined");
                        }
                    }
                    var i = 0;
                    sheets.items.forEach(function (sheet) {
                        //console.log(sheet.name);
                        var shName = sheet.name;

                        if (shName != "Combined" && !shName.includes("Remove_Dups")) {
                            i += i;
                            addOpt(shName, i);
                        }
                    });

                });
        }).catch(errorHandler);
    }

    function addOpt(itm, i) {
        var optSheet = document.getElementById('lsBxShNm');
        var opt = document.createElement("option");
        opt.name = "source";
        opt.id = "chkbx_" + i;
        opt.text = itm;
        opt.value = itm;
        optSheet.appendChild(opt);
    }

    function hideMain() {
        console.log("get started button clicked");
        var carasolBlock = document.getElementById("carouselSection");
        carasolBlock.style.display = "none";
        $("#carouselSection").hide();

        //var elem = document.getElementById("mainSection");
        $("#mainSection").show();
    }


    function selectAndCombineConf() {
        var selection = [];
        var shtCnt = 0;
        const desShNm = "Combined";
        var shFound = false;

        $('select > option:selected').each(function () {
            selection.push($(this).text());
            shtCnt += 1;
        });


        if (shtCnt <= 1) {
            showNotification("Error", "select atleast two sheets to merge");
            return;
        }

        var appendData = $('#ckbxAppend').is(':checked');

 
        if (!appendData) {

            Excel.run(function (context) {
                let sheets = context.workbook.worksheets;
                sheets.load("name");

                return context.sync()
                    .then(function () {
                        if (sheets.items.length <= 1) {
                            showNotification("Error", "atleast two sheets required to merge");
                            return;
                        }
                        sheets.items.forEach(function (sheet) {
                            var shNm = sheet.name;
                            if (shNm == desShNm) {
                                shFound = true;
                            }
                        });
                        return context.sync()
                            .then(function () {
                                //delete the data if sheet is found
                                if (shFound == true) {
                                    //we will confirm before deleting sheet
                                    var title = "You are about to delete 'Combined sheet' and create new one";
                                    $("#modCombineSheet .modal-body").text(title);
                                    $('#modCombineSheet').modal('show');
                                }
                                else {
                                    selectAndCombine();
                                }
                                //return context.sync()
                            });
                    });

            }).catch(errorHandler);
        }
        else {
            selectAndCombine();
        }

    };

    function selectAndCombine() {
        var selection = [];
        var shtCnt = 0;

        $('select > option:selected').each(function () {
            selection.push($(this).text());
            shtCnt += 1;
            //console.log($(this).text());
            //console.log($(this).text() + ' ' + $(this).val());
        });

        if (shtCnt <= 1) {
            showNotification("Error", "select atleast two sheets to merge");
            return;
        }
        combineSheets(selection, $('#ckboxIncludeHeader').is(':checked'),
            $('#ckbxAppend').is(':checked'),
            $('#ckbxRemoveDups').is(':checked'),
            $('#ckbxCopyFormat').is(':checked')
        );
    };


    function funcDeleteSheet() {

        Excel.run(function (context) {
            var shCnt = 0;
            let sheets = context.workbook.worksheets;
            sheets.load("name");

            return context.sync()
                .then(function () {
                    sheets.items.forEach(function (sheet) {
                        //console.log(sheet.name);
                        var shName = sheet.name;

                        if (shName.includes("Remove_Dups") || shName == "Combined") {
                            sheet.delete()
                            shCnt = shCnt + 1;
                        }

                        return context.sync()
                            .then(function () {
                                showNotification("Info", shCnt + " sheet(s) deleted.");
                            });
                    });

                });

        }).catch(errorHandler);
    }

    function funcDeleteSheetConf() {
        //var modTitle = document.getElementById("modDelSht");
        Excel.run(function (context) {
            var shCnt = 0;
            var shNms = "";
            let sheets = context.workbook.worksheets;
            sheets.load("name");

            return context.sync()
                .then(function () {
                    sheets.items.forEach(function (sheet) {
                        //console.log(sheet.name);
                        var shName = sheet.name;

                        if (shName.includes("Remove_Dups") || shName == "Combined") {
                            if (shNms == "") { shNms = "'" + shName + "'" ; }
                            else { shNms = shNms + ", '" + shName + "'"; }
                            shCnt = shCnt + 1;
                        }

                        return context.sync()
                            .then(function () {

                                if (shCnt <= 0) {
                                    showNotification("Info", "No Generated Sheets to Delete.");
                                }
                                else {
                                   // $("#myalertbox .modal-body").html('pass your html here');
                                    var title = "You are about to delete Sheets:  " + shNms + " (" + shCnt + ")";
                                    $("#modShtDel .modal-body").text(title);
                                    $('#modShtDel').modal('show');
                                }
                            });
                    });

                });

        }).catch(errorHandler);
    }

    function combineSheets(SheetNames, includeHeader, appendData, removeDups, copyFormat) {
        var shFound = false;
        const desShNm = "Combined";
        let firstRun = false;
        var shString;
        Excel.run(function (context) {
            let sheets = context.workbook.worksheets;
            sheets.load("name");

            return context.sync()
                .then(function () {
                    if (sheets.items.length <= 1) {
                        showNotification("Error", "atleast two sheets required to merge");
                        return;
                    }
                    sheets.items.forEach(function (sheet) {
                        var shNm = sheet.name;
                        shString = shString + "#" + shNm + "#";
                        //shNm = shNm.lower;
                        if (shNm == desShNm) {
                            shFound = true;
                        }
                    });
                    return context.sync()
                        .then(function () {
                            //delete the data if sheet is found
                            if (shFound == true && appendData == false) {
                                let delSheet = sheets.getItem(desShNm);;
                                delSheet.delete();
                                console.log("combined sheet deleted");
                                //sheets.getSheetByName(desShNm).deleteSheet(); 
                            }

                            return context.sync()
                                .then(function () {
                                    if (shFound == false || appendData == false) {
                                        sheets.add(desShNm);
                                    }

                                    return context.sync()
                                        .then(function () {
                                            SheetNames.forEach(function (srcShNm) {

                                                if (shString.includes("#" + srcShNm + "#")) {
                                                    feedData(srcShNm, desShNm, includeHeader, appendData, firstRun, copyFormat);
                                                }
                                                else { showNotification("Error", srcShNm + " is not found. Refresh and try again."); }
                                                firstRun = true;
                                            });
                                            //return context.sync()
                                        });
                                });
                        });
                });

        }).catch(errorHandler);
    }

    function feedData(srcShNm, desShNm, includeHeader_, appendData_, firstRun_, copyFormat_) {
        Excel.run(function (context) {
            //context.application.calculate(Excel.CalculationType.full);
            let sh = context.workbook.worksheets;
            var appendData = appendData_;
            var includeHeader = includeHeader_;
            var debug_mode = false;
            var tempVal = 0;

            //var docVal = document.getElementById("last_row");
            var docLastCol = document.getElementById("last_col");
            var docLastColAdd = document.getElementById("last_col_add");

            let srcSh = sh.getItem(srcShNm);
            srcSh.load("name");
            let desSh = sh.getItem(desShNm);
            desSh.load("name");


            return context.sync()
                .then(function () {
                    let srcLastCell_ = srcSh.getUsedRange().getLastCell();
                    srcLastCell_.load("address");
                    let srcLastRow_ = srcSh.getUsedRange().getLastRow();
                    srcLastRow_.load("rowIndex");

                    //actual range
                    var srcTempRng_ = srcSh.getUsedRange();
                    var srcTempRng = srcTempRng_.load("columnCount , address");

                    let desLastRow_ = desSh.getUsedRange().getLastRow();
                    let desLastAdd = desSh.getUsedRange().getLastCell();
                    desLastRow_.load("rowIndex");
                    desLastAdd.load("address");

                    return context.sync()
                        .then(function () {

                            let srcLastCell = cleanAdd(srcLastCell_.address);
                            var srcLastRow = srcLastRow_.rowIndex + 1;  //Source Data is counting from 0 so adding 1 to match excel row
                            var desLastRow = desLastRow_.rowIndex;
                            let tempDesLastRow = desLastRow;

                            //console.log("DES LAST ROW:>>>" + desLastRow);
                            //IF THERE IS NO DATA TO APPEND
                            if (desLastRow == 0) {
                                appendData = false;
                                desLastRow += 1;
                            } 

                            //first Time
                            if (!firstRun_) {

                                //IF THERE IS EXISTING DATA
                                var cellAdd = cleanAdd(desLastAdd.address);

                                //if A1 is last cell we assume there is no data
                                //isse was the rowIndex returns 0 in no data and 0 if there is data in first row
                                if (tempDesLastRow == 0 && cellAdd != 'A1' ) {
                                    desLastRow += 1;
                                } 
                                else if (tempDesLastRow > 0) {
                                    desLastRow += 2;
                                }

                                //docVal.value = desLastRow;
                                $('#last_row').val(desLastRow) ;

                                //BUILDING ADDRESS STRING FOR REMOVE DUPLICATES
                                docLastCol.value = srcTempRng.columnCount;
                                docLastColAdd.value = cleanAdd(srcTempRng.address);
                            }
                            else {
                                //desLastRow = parseInt(docVal.value);
                                desLastRow = parseInt($('#last_row').val());
                                
                                //BUILDING ADDRESS STRING FOR REMOVE DUPLICATES
                                var cTempVal = parseInt(docLastCol.value);
                                //if current value is more than last value. then put current address
                                if (cTempVal < srcTempRng.columnCount) {
                                    docLastCol.value = srcTempRng.columnCount;
                                    docLastColAdd.value = cleanAdd(srcTempRng.address);
                                }
                            }

                            //if no data start
                            if (srcLastCell != "A1") {

                                var srcRng;
                                if (includeHeader) { srcRng = "A1:" + srcLastCell; }
                                else { srcRng = "A2:" + srcLastCell; }

                                if (copyFormat_) {
                                    desSh.getRange("A" + desLastRow).copyFrom(srcSh.getRange(srcRng), Excel.RangeCopyType.all, false, false);
                                } else {
                                    desSh.getRange("A" + desLastRow).copyFrom(srcSh.getRange(srcRng), Excel.RangeCopyType.values, false, false);
                                }

                                if (!includeHeader) { srcLastRow -= 1; }
                                //docVal.value = parseInt(docVal.value) + srcLastRow;
                                $('#last_row').val(parseInt($('#last_row').val()) + srcLastRow);

                                if (debug_mode) {
                                    console.log("pasted row (" + srcLastRow + ") from source " + srcShNm + " at " + desLastRow);
                                }
                                
                            } else {
                                if (debug_mode) {
                                    console.log("not enough data in source " + srcShNm);
                                }
                            }
                            desSh.activate();
                            //return context.sync();
                        });
                });

        }).catch(errorHandler);
    }

    function addZero(str) {
        var val = str;
        if (val.length < 10) {
            val = 0 + val;
        }
        else if (val.length > 2000) {
            val = val.slice(-2);
        }
        return val;
    }


    function getDTime() {
        var dt = new Date();
        var dtStr = "";
        dtStr = addZero(dt.getDate()) + addZero(dt.getMonth()) + "" + "" + addZero(dt.getFullYear()) + "_" + dt.toLocaleTimeString();
        return dtStr;
    }

    function funRemoveAnyDupsConf() {
        Excel.run(function (context) {

            let sh = context.workbook.worksheets;
            let activeSh = sh.getActiveWorksheet();
            activeSh.load("name");

            return context.sync()
                .then(function () {
                    var title = "You are about to remove duplicates from '" + activeSh.name + "' sheet";
                    $("#modRemoveDups .modal-body").text(title);
                    $('#modRemoveDups').modal('show');
                    //return context.sync();
                });

        }).catch(errorHandler);
    }

    function funRemoveAnyDups() {
        Excel.run(function (context) {

            let sh = context.workbook.worksheets;
            let activeSh = sh.getActiveWorksheet();
            let activeShRng = activeSh.getUsedRange();

            let actLastCell_ = activeSh.getUsedRange().getLastCell();
            actLastCell_.load("rowIndex, address");

            return context.sync()
                .then(function () {


                    let actLastCell = cleanAdd(actLastCell_.address);
                    let actLastRow = actLastCell_.rowIndex;

                    if (actLastCell == "A1" || actLastRow <= 1) {
                        showNotification("Error", "Not Enough Data");
                        return;
                    }

                    //const newShNm = "Remove_Dups";
                    var newShNm = "Remove_Dups_" + getDTime();
                    newShNm = newShNm.replaceAll(":", "");
                    newShNm = newShNm.replace(" ", "");

                    let newSh = sh.add(newShNm);
                    //newSh.load("name");
                    let rng = "A1:" + actLastCell;

                    return context.sync()
                        .then(function () {
                            newSh.getRange(rng).copyFrom(activeSh.getRange(rng), Excel.RangeCopyType.all, false, false);
                            console.log("Dups Range" + rng);

                            let deleteResult = newSh.getRange(rng).removeDuplicates([0], true);
                            deleteResult.load();

                            newSh.activate();

                            return context.sync()
                                .then(function () {
                                    showNotification("Info", deleteResult.removed + " entries with duplicate names removed.");
                                    /*
                                    console.log(deleteResult.removed + " entries with duplicate names removed.");
                                    console.log(deleteResult.uniqueRemaining + " entries with unique names remain in the range.");
                                    */
                                });

                        });
                });

        }).catch(errorHandler);
    }

    //remove sheet name from cell address
    function cleanAdd(add) {
        var cStr = add;
        if (cStr.includes("!")) { cStr = cStr.split("!")[1]; }
        if (cStr.includes(":")) { cStr = cStr.split(":")[1]; }
        return cStr;
    }


    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
        //Please enable config.extendedErrorLogging to see full statements.
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        $("#notification-popup").slideDown("slow");
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
        setTimeout(() => {
            $("#notification-popup").slideUp("slow");
        }, "3500");
    }

})();


