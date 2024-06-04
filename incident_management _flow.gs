const rawSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RAW")
const actionSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ACTION")
const progressSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PROGRESS")
const archiveSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ARCHIVE")

// create a UI for the process
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Incident Kezelés')
      //.addSubMenu(ui.createMenu('Lekérdezés')
      .addItem('Hibák frissítése', 'transferDataToAction')
      .addSeparator()
      .addItem('Folyamatban', 'copyToProgess')
      .addSeparator()
      //.addSubMenu(ui.createMenu('Törlés (tesztelési fázishoz)')
          //.addItem('ACTION sheet', 'clearRangeAction'))
      .addToUi();
}

//loading the data from the json which is load into RAW sheet - periodically
function transferDataToAction() {
  var rawData = rawSheet.getRange('A2:A').getValues().flat();
  var actionData = actionSheet.getRange('A:A').getValues().flat();
  var progressData = progressSheet.getRange('A:A').getValues().flat();

  //create a list of the already exisiting IDs
  var duplicateIDs = [];
  var uniqueRawData = [];
  
  //define which id goes which list (mentioned above)
  rawData.forEach(function(id, index) {
    if (id && (actionData.includes(id) || progressData.includes(id))) {
      duplicateIDs.push(id);
    } 
      else if (id) {
        uniqueRawData.push(index + 2);
      }; 
    });
  
  var lastRowAction = getLastRow(actionSheet); //from where, do not write over data in action sheet
  uniqueRawData.forEach(function(rowIndex) {
      var rangeToCopy = rawSheet.getRange(rowIndex, 1, 1, rawSheet.getLastColumn());
      rangeToCopy.copyTo(actionSheet.getRange(lastRowAction + 1, 1));
      lastRowAction++; 
  });

  copyToAction();
  archiveSolvedIncident();

  //alert for the user about the already exisiting IDs
  if (duplicateIDs.length > 0) {
    var ui = SpreadsheetApp.getUi();
    ui.alert('Az alábbi erőmű azonosítók még mindig hibát jeleznek:\n' + duplicateIDs.join('-'));
  }
}

//this funtion is define the last row on ACTION sheet
function getLastRow(actionSheet) {
  var lastRow = actionSheet.getLastRow();
  var range = actionSheet.getRange(lastRow, 1, 1, actionSheet.getLastColumn());
  var values = range.getValues().flat();

  while (values.join("") === "" && lastRow >0) {
    lastRow--;
    range = actionSheet.getRange(lastRow, 1, 1, actionSheet.getLastColumn());
    var values = range.getValues().flat();
  }

  return lastRow;
}

//this function is responsible to copy issue into PROGRESS sheet and prepare ACTION sheet for the next refresh
function copyToProgess() {
  var rangeAction = actionSheet.getDataRange(); //the whole data range in Action sheet
  var valuesAction = rangeAction.getValues();
  var lastRowAction = valuesAction.length;
  var dates = actionSheet.getRange('P2:P').getValues();
  var currentDate = new Date();

  //defining last row in column P in order to format the date correctly
  var lastNonEmptyRow = 1;
  for (var j = 0; j < dates.length; j++) {
    if (dates[j][0]) {
      lastNonEmptyRow = j + 2;
    }
  }

  //copy to progress sheet, but do not copy which is equal or less than TODAY
  for (var i = 1; i < lastRowAction; i++) {
    var hValue = valuesAction[i][13];
    //var dateValue = valuesAction[i][15];

    if (i < lastNonEmptyRow) {
      var sheetDate = new Date(dates[i - 1][0]);
      if (!isNaN(sheetDate.getTime())) {
        actionSheet.getRange('P' + (i + 1)).setValue(sheetDate).setNumberFormat('yyyy-MM-dd');
      }
    }

    if(hValue !== "" && sheetDate >= currentDate) {
      var rangeToCopy = actionSheet.getRange(i + 1, 1, 1, valuesAction[i].length);
      rangeToCopy.copyTo(progressSheet.getRange(progressSheet.getLastRow() + 1,1));
    }
  }
  //clear in order to prepare for the next refresh
  clearRangeAction();
}

//clear the ACTION sheet - only for rows which require actions
function clearRangeAction() {
  var lastRowAction = actionSheet.getLastRow();
  var dataRange = actionSheet.getRange("A2:Q" + lastRowAction);
  var data = dataRange.getValues();
  var today = new Date();

  //iterate backwards to delete safely
  for (var i = data.length - 1; i >= 0; i--) {
    var row = data[i];
    var dateCell = row[15]; // P column (index 15, starts from 0)
    var sheetDate = new Date(dateCell);

    if (!dateCell || sheetDate >= today) {
      actionSheet.deleteRow(i + 2);
      //actionSheet.getRange(i + 2, 1, 1, actionSheet.getLastColumn()).clearContent();
    }
  }
}

//format the date and copy the row which is equal or less than TODAY
function copyToAction() {
  var lastRowProgress = progressSheet.getRange("P:P").getValues().filter(String).length;
  var dates = progressSheet.getRange('P2:P' + lastRowProgress).getValues();
  var lastRowAction = actionSheet.getLastRow();
  var currentDate = new Date();

  for (var i = dates.length - 1; i >= 0; i--) {
    var sheetDate = new Date(dates[i][0]);
    progressSheet.getRange('P' + (i + 2)).setValue(sheetDate).setNumberFormat('yyyy-MM-dd');

    if (currentDate >= sheetDate) {
      var progressId = progressSheet.getRange('A' + (i +2)).getValue();
      var copyRow = true;
      var actionID = actionSheet.getRange('A2:A' + lastRowAction).getValues().flat();
      if (actionID.indexOf(progressId) !== -1) {
        copyRow = false;
      }
      if (copyRow) {
        var rangeToCopy = progressSheet.getRange(i + 2, 1, 1, progressSheet.getLastColumn());
        rangeToCopy.copyTo(actionSheet.getRange(lastRowAction + 1, 1));
        lastRowAction++;
      
        progressSheet.deleteRow(i + 2);
      }
    }
  }
}

//archive issue which has been solved
function archiveSolvedIncident() {
  // Get the data from the sheets
  var rawData = rawSheet.getDataRange().getValues();
  var actionData = actionSheet.getDataRange().getValues();
  var progressData = progressSheet.getDataRange().getValues();
  
  // Create a set for RAW IDs for fast lookup
  var rawIds = new Set(rawData.map(row => row[0])); // Assuming ID is in the first column
  
  // Initialize arrays to hold rows to be archived
  var actionRowsToArchive = [];
  var progressRowsToArchive = [];
  var rowsToDeleteFromAction = [];
  var rowsToDeleteFromProgress = [];
  
  // Iterate through the Action data starting from the second row (to skip headers)
  for (var i = 1; i < actionData.length; i++) {
    var id = actionData[i][0]; // Assuming ID is in the first column
    if (!rawIds.has(id)) {
      actionRowsToArchive.push(actionData[i]);
      rowsToDeleteFromAction.push(i + 1);
    }
  }
  
  // Iterate through the Progress data starting from the second row (to skip headers)
  for (var j = 1; j < progressData.length; j++) {
    var id = progressData[j][0]; // Assuming ID is in the first column
    if (!rawIds.has(id)) {
      progressRowsToArchive.push(progressData[j]);
      rowsToDeleteFromProgress.push(j + 1);
    }
  }

  // Append the rows to the ARCHIVE sheet if there are any
  if (actionRowsToArchive.length > 0) {
    archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, actionRowsToArchive.length, actionRowsToArchive[0].length).setValues(actionRowsToArchive);
    removeRows(actionSheet, rowsToDeleteFromAction);
  }
  
  if (progressRowsToArchive.length > 0) {
    archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, progressRowsToArchive.length, progressRowsToArchive[0].length).setValues(progressRowsToArchive);
    removeRows(progressSheet, rowsToDeleteFromProgress);
  }
}

function removeRows(sheet, rows) {
  // Sort rows in descending order
  rows.sort((a, b) => b - a);
  // Delete rows from the bottom to the top
  rows.forEach(row => sheet.deleteRow(row));
}
