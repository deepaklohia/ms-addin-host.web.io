(function () {
  "use strict";

  var cellToHighlight;
  var messageBanner;

  // The initialize function must be run each time a new page is loaded.
  Office.initialize = function (reason) {
      $(document).ready(function () {
          // Initialize the notification mechanism and hide it
          var element = document.querySelector('.MessageBanner');
          messageBanner = new components.MessageBanner(element);
          messageBanner.showBanner();
          messageBanner.hideBanner();

          // If not using Excel 2016, use fallback logic.
          /*
          if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
              //$("#btnFillHeader").text("Outdated version.");
              //$('#button-text').text("Display!");
              //$('#button-desc').text("Display the selection");

              //$('#highlight-button').click(displaySelectedCells);
              return;
          }
          */

          loadDefaults();

          $("#btnP").click(() => tryCatch(markAtt("P")));
          $("#btnPL").click(() => tryCatch(markAtt("PL")));
          $("#btnWO").click(() => tryCatch(markAtt("WO")));
          $("#btnHO").click(() => tryCatch(markAtt("HO")));
          $("#btnLate").click(() => tryCatch(markAtt("LATE")));
          $("#btnX").click(() => tryCatch(markAtt("X")));
          $("#btnAB").click(() => tryCatch(markAtt("AB")));
          $("#btnPAB").click(() => tryCatch(markAtt("PAB")));
          $("#btnUAB").click(() => tryCatch(markAtt("UAB")));
          
          $('#btnFillHeader').click(fillHeaderInfos);
          $('#btnAutoFillAtten').click(markAttenAuto);
          $('#ckbxSelOnly').click(autoHdrChg);
          
          $('#btnGetStarted').click(hideMain);
          $('#goToHomeLabel').click(hideMain);

          //$("attenSection").show();
          //$("attenSection").hide();
          

            var elem = document.getElementById("attenSection");
            $("#attenSection").hide();
            //elem.style.display = 'none'; // hide
            //elem.style.visibility = 'hidden';
            /*
            elem.style.visibility = 'hidden'; // hide, but lets the element keep its size
            elem.style.visibility = 'visible';
            */

      });
  };

  function hideMain(){
    console.log("get started button clicked");

    var carasolBlock = document.getElementById("carouselSection");
    carasolBlock.style.display = "none";
    $("#carouselSection").hide();

    //var carasolBlock = document.getElementById("attenSection");
    //carasolBlock.style.visibility = "visible";
    var elem = document.getElementById("attenSection");
    //elem.style.visibility = 'visible';
    $("#attenSection").show();
    
    //$("attenSection").show();
  }

  async function tryCatch(callback) {
      try {
          await callback;
      } catch (error) {
          // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
          showNotification("Error", "error");
          console.error(error);
      }
  }

  function autoHdrChg() {
      var selectedOnly = $("#ckbxSelOnly").is(":checked");
      var btn = document.getElementById("btnAutoFillAtten");
      var btnVal = "Auto-Fill";

      if (selectedOnly) {
          btnVal = "Auto-Fill Selection";
      } else {
          btnVal = "Auto-Fill All";
      }
      btn.value = btnVal;
      btn.innerHTML = btnVal;
  }

  function getWeekNameFrmRng(rngNm) {
      var cStr;
      //console.log("xxx" + abc.slice(0, 1));
      //console.log("xxx" + abc.substring(0, 1));
      //console.log("xxx" + abc.charAt(0));
      if (rngNm.includes("!")) {
          cStr = rngNm.split("!")[1];
          cStr = cStr.substring(0, 1);
          cStr = cStr + 2;
          return cStr;
      } else {
          return "";
      }
  }

  function markAttenAuto() {

      var replaceVal = $("#ckbxReplace").is(":checked");
      var selectedOnly = $("#ckbxSelOnly").is(":checked");

      Excel.run(function (ctx) {
          const sh = ctx.workbook.worksheets.getActiveWorksheet();
          const selRng = ctx.workbook.getSelectedRange();
          const usedRng = sh.getUsedRange();
          const rowIn = sh.getUsedRange().getLastRow().load("rowIndex");
          const colIn = sh.getUsedRange().getLastColumn().load("colIndex");
          const lastC = sh.getUsedRange().getLastCell();

          colIn.load("address");
          var lc = lastC.load("address");
          var sr = selRng.load("rowCount, columnCount , address");
          var ur = usedRng.load("rowCount, columnCount, address");
          var dtRng = sh.getRange("C2:AZ2").load("values");

          return ctx.sync().then(function () {
              const lastRow = rowIn.rowIndex + 1;

              //console.log ("sr.columnCount " +sr.columnCount ) ;
              //console.log ("ur.columnCount " +ur.columnCount ) ;

              if (lastRow <= 1) {
                  showNotification("No Data", "click on Fill header button first");
                  return
              } else if (lastRow <= 2) {
                  showNotification("No Data", "fill some names / ID in column A/B");
                  return;
              } else if ((sr.rowCount > 5000 || sr.columnCount > 5000) && selectedOnly) {
                  showNotification("invalid selection", "too many rows/column selected");
                  return;
              } else if (ur.rowCount > 5000) {
                  showNotification("Data Error", "too many rows present");
                  return;
              } else if (ur.columnCount > 35) {
                  showNotification("Data Error", "too many columns present. there should be max 31 days in column header");
                  return;
              }

              var dayText = "";
              var cusRng;
              var v;
              const lastCola = lc.address;

              var lastRangeCol = "AG10";
              if (lastCola.includes("!")) {
                  lastRangeCol = lastCola.split("!")[1];
              }
              //actual range
              var actualRange = sh.getRange("C3:" + lastRangeCol);
              var ar = actualRange.load("rowCount, columnCount , address");

              return ctx.sync().then(function () {
                  //building ACTUAL RANGE
                  const actualRngStr = [];
                  //preparing cell address of actual range
                  for (r = 0; r < ar.rowCount; r++) {
                      for (c = 0; c < ar.columnCount; c++) {
                          cel = ar.getCell(r, c);
                          cel.load("address", "values");
                          actualRngStr.push(cel);
                          //arRng.push(cel);
                      }
                  }

                  if (selectedOnly) {
                      //BUILDING SELECTED RANGE
                      const selStr = [];
                      for (var r = 0; r < sr.rowCount; r++) {
                          for (var c = 0; c < sr.columnCount; c++) {
                              var cel = sr.getCell(r, c);
                              cel.load("columnCount, address, values");
                              selStr.push(cel);
                          }
                      }

                      //getting actual values of the selected cells
                      return ctx.sync().then(function () {
                          var selectedStr;
                          var actualRngFullStr;

                          //getting actual range combined
                          for (var i = 0; i < actualRngStr.length; i++) {
                              actualRngFullStr = actualRngFullStr + "#" + actualRngStr[i].address + "#";
                          }
                          //CUSTOM BUILD RANGE FOR WEEKDAY NAMES
                          var mStr = [];
                          var dayVal;
                          for (i = 0; i < selStr.length; i++) {
                              var newRng = getWeekNameFrmRng(selStr[i].address);
                              if (newRng != "") {
                                  dayVal = sh.getRange(newRng);
                              } else {
                                  dayVal = sh.getRange("A1");
                              }
                              dayVal.load("address,values");
                              mStr.push(dayVal);
                          }

                          return ctx.sync().then(function () {
                              //marking values and getting actual values
                              for (i = 0; i < selStr.length; i++) {
                                  var dayText = mStr[i].values;
                                  v = setDayAtnObj(dayText);
                                  selectedStr = selectedStr + "#" + selStr[i].address + "#";
                                  if (actualRngFullStr.includes("#" + selStr[i].address + "#") && replaceVal == true) {
                                      selStr[i].values = v;
                                      fillRange(selStr[i], v);
                                  } else if (
                                      actualRngFullStr.includes("#" + selStr[i].address + "#") &&
                                      replaceVal == false &&
                                      selStr[i].values == ""
                                  ) {
                                      selStr[i].values = v;
                                      fillRange(selStr[i], v);
                                  }
                              }
                          }).then(ctx.sync);
                      });
                  } else if (selectedOnly == false && replaceVal == false) {
                      //building ACTUAL RANGE
                      const arRng = [];
                      //var arRng = [];
                      //preparing cell address of actual range
                      for (r = 0; r < ar.rowCount; r++) {
                          for (c = 0; c < ar.columnCount; c++) {
                              cel = ar.getCell(r, c);
                              cel.load("values ,address");
                              arRng.push(cel);
                          }
                      }

                      return ctx.sync().then(function () {
                          var mStr2 = [];
                          var dayVal2;
                          for (var i = 0; i < arRng.length; i++) {
                              var newRng = getWeekNameFrmRng(arRng[i].address);
                              if (newRng != "") {
                                  dayVal2 = sh.getRange(newRng);
                              } else {
                                  dayVal2 = sh.getRange("A1");
                              }
                              dayVal2.load("address,values");
                              mStr2.push(dayVal2);
                          }

                          return ctx.sync().then(function () {
                              var dayText2;
                              for (var i = 0; i < arRng.length; i++) {
                                  dayText2 = mStr2[i].values;
                                  v = setDayAtnObj(dayText2);
                                  if ((arRng[i].values == "" && replaceVal == false) || replaceVal == true) {
                                      arRng[i].values = v;
                                      fillRange(arRng[i], v);
                                  }
                              }
                          });
                      })
                      .then(ctx.sync);
                  }
                  //FILL ALL REPLCE ALL -- FASTEST
                  else if (selectedOnly == false && replaceVal == true) {
                      var col;
                      for (var j = 3; j <= ur.columnCount; j++) {
                          dayText = dtRng.values[0][j - 3];
                          col = getCharCol(j - 1);
                          v = setDayAtnStr(dayText);
                          //taking entire range
                          cusRng = sh.getRange(col + "3:" + col + lastRow);
                          sh.getRange(col + "3:" + col + lastRow).values = v;
                          fillRange(cusRng, v);
                      }
                  }
              });
          });
      });
  }

  function setDayAtnObj(dy) {
      //console.log("object type:"  +  typeof dy) ;
      if (typeof dy == "object") {
          // eslint-disable-next-line @typescript-eslint/no-unused-vars
          for (let [key, value] of Object.entries(dy)) {
              dy = value;
              break;
          }
      }
      var mo = document.getElementById("cmbxMon").value;
      var tu = document.getElementById("cmbxTue").value;
      var we = document.getElementById("cmbxWed").value;
      var th = document.getElementById("cmbxThu").value;
      var fr = document.getElementById("cmbxFri").value;
      var sa = document.getElementById("cmbxSat").value;
      var su = document.getElementById("cmbxSun").value;

      if (dy.includes("Mon")) {
          return mo;
      } else if (dy.includes("Tue")) {
          return tu;
      } else if (dy.includes("Wed")) {
          return we;
      } else if (dy.includes("Thu")) {
          return th;
      } else if (dy.includes("Fri")) {
          return fr;
      } else if (dy.includes("Sat")) {
          return sa;
      } else if (dy.includes("Sun")) {
          return su;
      } else {
          ("");
      }
  }

  function setDayAtnStr(dy) {
      var mo = document.getElementById("cmbxMon").value;
      var tu = document.getElementById("cmbxTue").value;
      var we = document.getElementById("cmbxWed").value;
      var th = document.getElementById("cmbxThu").value;
      var fr = document.getElementById("cmbxFri").value;
      var sa = document.getElementById("cmbxSat").value;
      var su = document.getElementById("cmbxSun").value;
      if (dy === "Mon") {
          return mo;
      } else if (dy.includes("Tue")) {
          return tu;
      } else if (dy.includes("Wed")) {
          return we;
      } else if (dy.includes("Thu")) {
          return th;
      } else if (dy.includes("Fri")) {
          return fr;
      } else if (dy.includes("Sat")) {
          return sa;
      } else if (dy.includes("Sun")) {
          return su;
      } else {
          ("");
      }
  }

  function fillRange(rng, aType) {
      var fillBg = $("#ckbxFill").is(":checked");
      if (fillBg == false) {
          return;
      }
      //rng.format.font.color = "white";
      var col = "";
      switch (aType) {
          case "P":
              col = "#d9ead3";
              break;
          case "PL":
              col = "#fff2cc";
              break;
          case "WO":
              col = "#cfe2f3";
              break;
          case "HO":
              col = "#ffd966";
              break;
          case "LATE":
              col = "#ffb366";
              break;
          case "AB":
              col = "#f4cccc";
              break;
          case "PAB":
              col = "#fc876e";
              break;
          case "UAB":
              col = "#ec6c51";
              break;
      }
      if (col == "") {
          rng.format.fill.clear();
      } else {
          rng.format.fill.color = col;
      }
      return;
  }

  function fillColor(aType) {
      //rng.format.font.color = "white";
      switch (aType) {
          case "P":
              return "#d9ead3";
          case "PL":
              return "#fff2cc";
          case "WO":
              return "#cfe2f3";
          case "HO":
              return "#ffd966";
          case "LATE":
              return "#ffb366";
          case "AB":
              return "#f4cccc";
          case "PAB":
              return "#f4cccc";
          case "UAB":
              return "#f4cccc";
          default:
              return " ";
      }
  }

  function markAtt(atnType) {
      console.log("entered");
      Excel.run(function (ctx) {
          const sh = ctx.workbook.worksheets.getActiveWorksheet();
          const selRng = ctx.workbook.getSelectedRange();
          const rowIn = sh.getUsedRange().getLastRow().load("rowIndex");
          const sr = selRng.load("rowCount, columnCount , address");
          const lastColNm = sh.getUsedRange().getLastCell().load("address");

          selRng.load("address");

          return ctx.sync().then(function () {
              var lastRow = rowIn.rowIndex + 1;

              //console.log(lastRow);
              if (lastRow <= 2) {
                  if (lastRow <= 1) {
                      showNotification("No Data", "click on Fill header button first");
                  } else {
                      showNotification("No Data", "fill names and id to begin");
                  }
                  return;
              } else if (sr.rowCount > 5000 || sr.columnCount > 5000) {
                  showNotification("invalid selection", "too many rows/columns selected");
                  return;
              }

              var fillBg = $("#ckbxFill").is(":checked");
              //selRng.format.fill.clear();

              var lastColA = lastColNm.address;
              var lastColCell = "AG10";
              if (lastColA.includes("!")) {
                  lastColCell = lastColA.split("!")[1];
              }

              //actual range
              var actualRange = sh.getRange("C3:" + lastColCell);
              var ar = actualRange.load("rowCount, columnCount , address");

              return ctx.sync().then(function () {
                  //building ACTUAL RANGE
                  var actualRngStr = [];
                  var cel;
                  //preparing cell address of actual range
                  for (var r = 0; r < ar.rowCount; r++) {
                      for (var c = 0; c < ar.columnCount; c++) {
                          cel = ar.getCell(r, c);
                          cel.load("address", "values");
                          actualRngStr.push(cel);
                      }
                  }
                  //BUILDING SELECTED RANGE
                  const selStr = [];
                  for (r = 0; r < sr.rowCount; r++) {
                      for (c = 0; c < sr.columnCount; c++) {
                          cel = sr.getCell(r, c);
                          cel.load("columnCount, address, values");
                          selStr.push(cel);
                      }
                  }

                  return ctx.sync().then(function () {
                      //var selectedStr;
                      var actualRngFullStr;

                      //getting actual range combined
                      for (var i = 0; i < actualRngStr.length; i++) {
                          actualRngFullStr = actualRngFullStr + "#" + actualRngStr[i].address + "#";
                      }
                      //we mark attendance by avoiding the header and the names
                      //console.log("full range" +  actualRngFullStr)
                      for (i = 0; i < selStr.length; i++) {
                          //if the range falls within the value
                          if (actualRngFullStr.includes("#" + selStr[i].address + "#")) {
                              selStr[i].format.fill.clear();
                              if (atnType == "X") {
                                  selStr[i].clear();
                              } else {
                                  selStr[i].values = atnType;
                                  if (fillBg) {
                                      selStr[i].format.fill.color = fillColor(atnType);
                                  }
                                  //selRng.values = atnType;
                              }
                          }
                      }
                      console.log("Marked at--" + atnType);
                  });
              });
          });
      }).catch(errorHandler);
  }

  function fillHeaderInfos() {
      Excel.run(function (ctx) {

       
          var cmbx = document.getElementById("listBoxMonthYear").options;
          var m = document.getElementById("listBoxMonthYear").value;
          var strDt = cmbx[cmbx.selectedIndex].text;
          /*
          console.log(m);
          console.log(strDt);
          */
          if (strDt == "") {
              return;
          }
          var sh = ctx.workbook.worksheets.getActiveWorksheet();
          var rng = sh.getRange("A1:AG2");
          rng.clear();
          rng.format.fill.clear();

          var usrMonth = m;
          //var ui = SpreadsheetApp.getUi();
          //var col = 2;

          var fullYear = strDt.split("-")[1];
          var startDate = new Date(fullYear, m, 1);

          //var userChoice = ui.alert(strDt + ': fill date header in row 1-2 ?', ui.ButtonSet.OK_CANCEL);
          //if (userChoice == ui.Button.CANCEL) { return; };

          //rename sheet name
          var shFound = false;
          var worksheets = ctx.workbook.worksheets;
          worksheets.load("items");

          return ctx.sync().then(function () {
              for (var i = 0; i < worksheets.items.length; i++) {
                  var ws = worksheets.items[i];
                  ws.load("name");
              }

              return ctx.sync().then(function () {
                  for (var i = 0; i < worksheets.items.length; i++) {
                      var wn = worksheets.items[i].name;
                      if (wn.toLowerCase() == strDt.toLowerCase()) {
                          shFound = true;
                          break;
                      }
                  }

                  if (shFound == false) {
                      sh.name = strDt;
                  }

                  sh.getRange("A2").values = "ID";
                  sh.getRange("B2").values = "Name";
                  var strMonth = getStrMonth(startDate.getMonth());
                  var strYear = startDate.getFullYear();

                  var lastColStr = 0;
                  for (i = 1; i <= 31; i++) {
                      var tempDate = new Date(startDate.getFullYear(), startDate.getMonth(), i);
                      //check if current month
                      if (tempDate.getMonth() == usrMonth) {
                          var colNm = getCharCol(i + 1);
                          /*
                          sh.getCell(0, i).values = i + "-" + strMonth + "-" + strYear;
                          sh.getCell(1, i).values = "=text(" + colNm + 1 + ', "ddd")';
                          */
                          sh.getRange(colNm + 1).values = i + "-" + strMonth + "-" + strYear;
                          sh.getRange(colNm + 2).values = "=text(" + colNm + 1 + ', "ddd")';
                          lastColStr = colNm;
                      }
                  }

                  var rng = sh.getRange("C1:" + lastColStr + 1);
                  var rng2 = sh.getRange("A2:" + lastColStr + 2);
                  /*
                      sh.getRange("C1:" + getCharCol(c) + 1)
                      .setBackground("#cfe2f3")
                      .setFontWeight("bold")
                      .setFontSize(12)
                      .setHorizontalAlignment("center");
        
                    sh.getRange("A2:" + getCharCol(c) + 2)
                      .setBackground("#cfe2f3")
                      .setFontWeight("regular")
                      .setFontStyle("italic")
                      .setFontSize(12)
                      .setHorizontalAlignment('center');
                    sh.getRange("A2:B2").setFontWeight('bold');
                    */
                  //sh.getRange("A1").style.font = { name: "Comic Sans MS" };

                  // Read the range address
                  rng.load("address");
                  rng2.load("address");

                  return ctx.sync()
                      .then(function () {
                      //rng.format.fill.color = "#cfe2f3";
                      rng.format.fill.color = "#E7E6E6";
                      rng2.format.fill.color = "#E7E6E6";
                      rng.format.autofitColumns();
                      console.log("Headers Filled");
                  });
              });
          }).then(ctx.sync);
      }).catch(errorHandler);
  }

  function getStrMonth(m) {
      var month = new Array();
      month[0] = "Jan";
      month[1] = "Feb";
      month[2] = "Mar";
      month[3] = "Apr";
      month[4] = "May";
      month[5] = "Jun";
      month[6] = "Jul";
      month[7] = "Aug";
      month[8] = "Sep";
      month[9] = "Oct";
      month[10] = "Nov";
      month[11] = "Dec";
      return month[m];
  }

  function getCharCol(col) {
      var str =
          "A;B;C;D;E;F;G;H;I;J;K;L;M;N;O;P;Q;R;S;T;U;V;W;X;Y;Z;AA;AB;AC;AD;AE;AF;AG;AH;AI;AJ;AK;AL;AM;AN;AO;AP;AQ;AR;AS;AT;AU;AV;AW;AX;AY;AZ;";
      return str.split(";")[col];
  }

  function loadDefaults() {
      var month = new Array();
      month[0] = "Jan";
      month[1] = "Feb";
      month[2] = "Mar";
      month[3] = "Apr";
      month[4] = "May";
      month[5] = "Jun";
      month[6] = "Jul";
      month[7] = "Aug";
      month[8] = "Sep";
      month[9] = "Oct";
      month[10] = "Nov";
      month[11] = "Dec";

      //console.log("Months Loaded");
      var dta = document.getElementById("listBoxMonthYear");
      //console.log("id identified");

      for (var i = -10; i <= 15; i++) {
          var currentDate = new Date();
          var customDate = new Date(currentDate.getFullYear(), currentDate.getMonth() + i, 1);
          var strMonth = month[customDate.getMonth()];
          var fullYear = customDate.getFullYear();
          var dropdownValue = strMonth + "-" + fullYear;
          dta.options[dta.options.length] = new Option(dropdownValue, customDate.getMonth());
      }

      dta.selectedIndex = 10;

      var monVal = document.getElementById("cmbxMon");
      monVal.value = "P";
      monVal.onfocus = function () {
          monVal.value = "";
      };
      var tueVal = document.getElementById("cmbxTue");
      tueVal.value = "P";
      tueVal.onfocus = function () {
          tueVal.value = "";
      };
      var wedVal = document.getElementById("cmbxWed");
      wedVal.value = "P";
      wedVal.onfocus = function () {
          wedVal.value = "";
      };
      var thuVal = document.getElementById("cmbxThu");
      thuVal.value = "P";
      thuVal.onfocus = function () {
          thuVal.value = "";
      };
      var friVal = document.getElementById("cmbxFri");
      friVal.value = "P";
      friVal.onfocus = function () {
          friVal.value = "";
      };
      var satVal = document.getElementById("cmbxSat");
      satVal.value = "WO";
      satVal.onfocus = function () {
          satVal.value = "";
      };
      var sunVal = document.getElementById("cmbxSun");
      sunVal.value = "WO";
      sunVal.onfocus = function () {
          sunVal.value = "";
      };
      console.log("Months Loaded");
  }


  function hightlightHighestValue() {
      // Run a batch operation against the Excel object model
      Excel.run(function (ctx) {
          // Create a proxy object for the selected range and load its properties
          var sourceRange = ctx.workbook.getSelectedRange().load("values, rowCount, columnCount");

          // Run the queued-up command, and return a promise to indicate task completion
          return ctx.sync()
              .then(function () {
                  var highestRow = 0;
                  var highestCol = 0;
                  var highestValue = sourceRange.values[0][0];

                  // Find the cell to highlight
                  for (var i = 0; i < sourceRange.rowCount; i++) {
                      for (var j = 0; j < sourceRange.columnCount; j++) {
                          if (!isNaN(sourceRange.values[i][j]) && sourceRange.values[i][j] > highestValue) {
                              highestRow = i;
                              highestCol = j;
                              highestValue = sourceRange.values[i][j];
                          }
                      }
                  }

                  cellToHighlight = sourceRange.getCell(highestRow, highestCol);
                  sourceRange.worksheet.getUsedRange().format.fill.clear();
                  sourceRange.worksheet.getUsedRange().format.font.bold = false;

                  // Highlight the cell
                  cellToHighlight.format.fill.color = "orange";
                  cellToHighlight.format.font.bold = true;
              })
              .then(ctx.sync);
      })
      .catch(errorHandler);
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