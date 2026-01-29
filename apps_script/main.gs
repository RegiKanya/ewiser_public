const CONFIG = {
  SHEET_NAMES: {
    RAW: "RAW",
    ACTION: "ACTION",
    PROGRESS: "PROGRESS",
    ARCHIVE: "ARCHIVE"
  },
  HEADERS: {
    ID: "Erőmű azonosító",
    DATE_1: "Ellenőrzés dátuma  (yyyy.mm.dd)",
    DATE_2: "Hibaüzenet létrehozva (dátum)",
    CHECK_1: "Előzmény",
    CHECK_2: "Mit kell vele csinálni következőnek?",
    SYNC_COLUMNS: [
      "Totál Inverter",
      "Hibás Inverterek",    
      "Leállás",          
      "Referencia inverter státusz",
      "Inverter teljesítménykülönbség (kW)",
      "Inverter teljesítménykülönbség aránya (%)"  
    ]
  }
}



// create a UI for the process
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Incident Kezelés')
      .addItem('ACTION sheet mentése', 'remindMeForTheDates')
      .addSeparator()
      .addItem('Hibák frissítése', 'warningToSave')
      .addToUi();
}

function warningToSave() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
  'Művelet jóváhagyása', 
  'Ez a művelet frissíti a hibákat és törli az ACTION sheet műveleteit. A változtatások visszavonhatatlanok. \n Szeretnéd folytatni? \n\n Ha előbb szeretnéd elmenteni a munkád, kattints a "Nem" gombra.', ui.ButtonSet.YES_NO);

  var shouldContinue = true;
  while (shouldContinue) {
    if (response == ui.Button.YES) {
    transferDataToAction();
    shouldContinue = false;
    Logger.log('The user click Yes and run nthe transferDataAction() function.');
  } else if (response == ui.Button.NO) {
    break;
    }
  }
}

function remindMeForTheDates() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Ellenőrzés dátum', 'Beírtad az emlékeztető dátumát?', ui.ButtonSet.YES_NO);
  var shouldContinue = true;
  while (shouldContinue) {
    if (response == ui.Button.YES) {
      copyToProgress();
      shouldContinue = false;
    } else if (response == ui.Button.NO) {
      break;
    }
  }
}

function changeCellValues() {
  updateProgressColumns();
  updateActionColumns();
}

//loading the data from the json which is load into RAW sheet - periodically
function transferDataToAction() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.RAW);
  const actionSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ACTION);
  const progressSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PROGRESS);

  changeCellValues();

  const rawData = rawSheet.getDataRange().getValues();
  const rawHeaders = rawData.shift(); 
  const actionData = actionSheet.getDataRange().getValues();
  const actionHeaders = actionData.shift();
  const progressData = progressSheet.getDataRange().getValues();
  const progressHeaders = progressData.shift();

  const rawIdIdx = rawHeaders.indexOf(CONFIG.HEADERS.ID);
  const actionIdIdx = actionHeaders.indexOf(CONFIG.HEADERS.ID); 
  const progressIdIdx = progressHeaders.indexOf(CONFIG.HEADERS.ID);

  if (rawIdIdx === -1) {
    SpreadsheetApp.getUi().alert("Hiba: 'Erőmű azonosító' oszlop nem található a RAW sheeten!");
    return;
  }

  const existingIds = new Set();

  // Action ID-k begyűjtése
  if (actionIdIdx !== -1) {
    actionData.forEach(row => existingIds.add(String(row[actionIdIdx])));
  }
  // Progress ID-k begyűjtése
  if (progressIdIdx !== -1) {
    progressData.forEach(row => existingIds.add(String(row[progressIdIdx])));
  }

  const newRowsToAdd = [];
  const duplicateIdsList = [];

  rawData.forEach(row => {
    const id = String(row[rawIdIdx]);
    
    if (id && id !== "undefined" && id !== "") {
      if (existingIds.has(id)) {
        duplicateIdsList.push(id);
      } else {
        newRowsToAdd.push(row);
      }
    }
  });

  if (newRowsToAdd.length > 0) {
    const nextRow = actionSheet.getLastRow() + 1;
    actionSheet.getRange(nextRow, 1, newRowsToAdd.length, newRowsToAdd[0].length).setValues(newRowsToAdd);
    
    const dateIdx = rawHeaders.indexOf(CONFIG.HEADERS.DATE_1);
    if (dateIdx !== -1) {
       actionSheet.getRange(nextRow, dateIdx + 1, newRowsToAdd.length, 1).setNumberFormat('yyyy-MM-dd');
    }
  }

  archiveSolvedIncident(); 
  copyToAction();          

  if (duplicateIdsList.length > 0) {
    var ui = SpreadsheetApp.getUi();
    ui.alert('Az alábbi erőmű azonosítók még mindig hibát jeleznek:\n' + duplicateIdsList.join('-'));
  } 
}

//this function is responsible to copy issue into PROGRESS sheet and prepare ACTION sheet for the next refresh
function copyToProgress() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const actionSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ACTION);
  const progressSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PROGRESS);
  const actionData = actionSheet.getDataRange().getValues();
  const headers = actionData.shift(); 
  const idDate = headers.indexOf(CONFIG.HEADERS.DATE_1); 
  const idDate2 = headers.indexOf(CONFIG.HEADERS.DATE_2);
  const idCheck1 = headers.indexOf(CONFIG.HEADERS.CHECK_1); 
  const idCheck2 = headers.indexOf(CONFIG.HEADERS.CHECK_2);

  if (idDate === -1) { SpreadsheetApp.getUi().alert("Dátum oszlop nem található!"); return; }

  const rowsToMove = [];
  const rowsToKeep = [];
  
  actionData.forEach(row => {
    row[idDate] = parseMyDate(row[idDate]);
    if (idDate2 !== -1) row[idDate2] = parseMyDate(row[idDate2]);

    let sheetDate = row[idDate];    
    let isDateValid = (sheetDate instanceof Date && !isNaN(sheetDate.getTime()));
    let actionTaken = false; 

    if (isDateValid) {
       const isDateReached = sheetDate >= currentDate; 
       const hasIssue = (row[idCheck1] !== "" || row[idCheck2] !== "");
       
       if (hasIssue && isDateReached) {
         rowsToMove.push(row);
         actionTaken = true; 
       }
    }

    if (!actionTaken) {
       const isFuture = isDateValid && (sheetDate >= currentDate); 
       const isMissingDate = !isDateValid || row[idDate] === "" || row[idDate] === undefined;

       if (isMissingDate || isFuture) {
         actionTaken = true; 
       }
    }

    if (!actionTaken) {
      rowsToKeep.push(row);
    }
  });
  
  if (rowsToMove.length > 0) {
    const nextRow = progressSheet.getLastRow() + 1;
    progressSheet.getRange(nextRow, 1, rowsToMove.length, rowsToMove[0].length).setValues(rowsToMove);
    progressSheet.getRange(nextRow, idDate + 1, rowsToMove.length, 1).setNumberFormat('yyyy-MM-dd');
    if (idDate2 !== -1) {
      progressSheet.getRange(nextRow, idDate2 + 1, rowsToMove.length, 1).setNumberFormat('yyyy-MM-dd');
    }
  }

  if (actionData.length > 0) { 
    actionSheet.getRange(2, 1, actionSheet.getLastRow(), actionSheet.getLastColumn()).clearContent();
    
    if (rowsToKeep.length > 0) {
      actionSheet.getRange(2, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);
      actionSheet.getRange(2, idDate + 1, rowsToKeep.length, 1).setNumberFormat('yyyy-MM-dd');
      if (idDate2 !== -1) {
         actionSheet.getRange(2, idDate2 + 1, rowsToKeep.length, 1).setNumberFormat('yyyy-MM-dd');
      }
    }
  }
  sortProgressSheetByDate(); 
}

//format the date and copy the row which is equal or less than TODAY
function copyToAction() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const actionSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ACTION);
  const progressSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PROGRESS);
  const progressData = progressSheet.getDataRange().getValues();
  const progressHeaders = progressData.shift();
  const progressIdIndex = progressHeaders.indexOf(CONFIG.HEADERS.ID);
  const idDate = progressHeaders.indexOf(CONFIG.HEADERS.DATE_1);
  const actionData = actionSheet.getDataRange().getValues();
  const actionHeaders = actionData[0]; // Fejléc
  const actionIdIndex = actionHeaders.indexOf(CONFIG.HEADERS.ID);
  
  if (progressData.length === 0) return;
  if (idDate === -1 || progressIdIndex === -1) {
    SpreadsheetApp.getUi().alert("Hiba: 'Ellenőrzés dátuma' vagy 'Erőmű azonosítü' oszlop nem található a PROGRESS lapon!");
    return;
  }

  const actionIdsSet = new Set();
  if (actionIdIndex !== -1) {
    for (let i = 1; i < actionData.length; i++) {
      actionIdsSet.add(String(actionData[i][actionIdIndex]));
    }
  }

  const rowsToMoveBack = [];
  const rowsToKeepInProgress = [];

  progressData.forEach(row => {
    row[idDate] = parseMyDate(row[idDate]);

    let sheetDate = row[idDate];
    const id = String(row[actionIdIndex]);
    let shouldMoveBack = false;
  
    if (sheetDate instanceof Date && !isNaN(sheetDate.getTime())) {
      if (currentDate >= sheetDate) {
        if (!actionIdsSet.has(id)) {
          shouldMoveBack = true;
        }
      }
    }

    if (shouldMoveBack) {
      rowsToMoveBack.push(row);
      actionIdsSet.add(id); 
    } else {
      rowsToKeepInProgress.push(row);
    }
  });

  if (rowsToMoveBack.length > 0) {
    const nextRow = actionSheet.getLastRow() + 1;
    actionSheet.getRange(nextRow, 1, rowsToMoveBack.length, rowsToMoveBack[0].length).setValues(rowsToMoveBack);
    
    actionSheet.getRange(nextRow, idDate + 1, rowsToMoveBack.length, 1).setNumberFormat('yyyy-MM-dd');
  }

  if (progressData.length > 0) {
    progressSheet.getRange(2, 1, progressSheet.getLastRow(), progressSheet.getLastColumn()).clearContent();
    
    if (rowsToKeepInProgress.length > 0) {
      progressSheet.getRange(2, 1, rowsToKeepInProgress.length, rowsToKeepInProgress[0].length).setValues(rowsToKeepInProgress);
                   
      progressSheet.getRange(2, idDate + 1, rowsToKeepInProgress.length, 1).setNumberFormat('yyyy-MM-dd');
    }
  }
}

//archive issue which has been solved
function archiveSolvedIncident() {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const rawSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.RAW);
 const actionSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ACTION);
 const progressSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PROGRESS);
 const archiveSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ARCHIVE);
 const rawData = rawSheet.getDataRange().getValues();
 const rawHeaders = rawData.shift();
 const rawIdIndex = rawHeaders.indexOf(CONFIG.HEADERS.ID);

 if (rawIdIndex === -1) {
  SpreadsheetApp.getUi().alert("Hiba: ID oszlop nem található a RAW sheeten!");
  return;
 }

 const rawIds = new Set();
 rawData.forEach(row => {
  const ids = String(row[rawIdIndex]);
  if (ids) rawIds.add(ids);
 });

 processArchiveDataForSheets(actionSheet, archiveSheet, rawIds, "ACTION");
 processArchiveDataForSheets(progressSheet, archiveSheet, rawIds, "PROGRESS");
}
