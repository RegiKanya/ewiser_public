const currentDate = new Date();
currentDate.setHours(0,0,0,0); 
const currentYear = currentDate.getFullYear(); // Dinamikus év


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
  
//clear the ACTION sheet - only for rows which require actions
  function clearRangeAction() {
    var lastRowAction = actionSheet.getLastRow();
    var dataRange = actionSheet.getRange("A2:Q" + lastRowAction);
    var data = dataRange.getValues();
    var today = new Date();
  
    //delete unnecessary rows which are later than today
    for (var i = data.length - 1; i >= 0; i--) {
      var row = data[i];
      var dateCell = row[14]; // O column (index 14, starts from 0)
      var sheetDate = new Date(dateCell);
  
      if (!dateCell || sheetDate >= today) {
        actionSheet.deleteRows(i + 2); 
      }
    }
}

//simple remove rows
  function removeRows(sheet, rows) {
    // Sort rows in descending order
    rows.sort((a, b) => b - a);
    // Delete rows from the bottom to the top (POTENTIAL error - the worksheet can run out rows because this funtion is decreasing the rows)
    rows.forEach(row =>sheet.deleteRow(row));
}

//ordering PROGRESS sheet by asc date
function sortProgressSheetByDate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const progressSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PROGRESS);
  var lastRow = progressSheet.getLastRow();

  if (lastRow > 1) {
    var range = progressSheet.getRange(2, 1, lastRow - 1, progressSheet.getLastColumn());
    range.sort({column: 15, ascending: true});
  }
}

//-------------------------------------------

function updateProgressColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.RAW);
  const progressSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PROGRESS);

  const rawData = rawSheet.getDataRange().getValues();
  const rawHeaders = rawData.shift();
  const rawIdIndex = rawHeaders.indexOf(CONFIG.HEADERS.ID);

  const progressData = progressSheet.getDataRange().getValues();
  const progressHeaders = progressData.shift();
  const progressIdIndex = progressHeaders.indexOf(CONFIG.HEADERS.ID);

  if (rawIdIndex == -1 && progressIdIndex == -1) {
    SpreadsheetApp.getUi().alert("Hiba: Az 'Erőmű azonosító' oszlop nem található valamelyik sheeten!")
    return;
  }

  //keressük meg a változtatásra szoruló oszlopok indexeit mindekét sheeten
  const columnMap = CONFIG.HEADERS.SYNC_COLUMNS.map( colName => {
    return {
      name: colName,
      rawIndex: rawHeaders.indexOf(colName),
      progressIndex: progressHeaders.indexOf(colName)
    };
  }).filter( col => col.rawIndex !== -1 && col.progressIndex !== -1);

  if (columnMap.length === 0) {
    Logger.log("Nincs közös oszlop, amit frissíteni kéne a RAW és PROGRESS között.");
    return;
  }

  //RAW adatbázis felépítése (Map) a gyors kereséshez
  const rawMap = new Map();
  rawData.forEach( row => {
    const id = String(row[rawIdIndex]);
    if (id) rawMap.set(id, row);
  })

  let changesCount = 0;

  // Végigiterálunk a Progress adatokon (sorról sorra) és összehasonlítunk
  progressData.forEach((progressRow) => {
    const progressId = String(progressRow[progressIdIndex]);
    if (progressId && rawMap.has(progressId)) {
      const rawRow = rawMap.get(progressId);

      columnMap.forEach( colInfo => {
        const rawValue = rawRow[colInfo.rawIndex];
        const progressValue = progressRow[colInfo.progressIndex];

        if (rawValue !== progressValue) {
          progressRow[colInfo.progressIndex] = rawValue;
          changesCount++;
        }
      });
    }
  });

  //Tömeges visszaírás (Batch Write)
  // Csak akkor írunk, ha volt változás
  if (changesCount > 0) {
        progressSheet.getRange(2, 1, progressData.length, progressData[0].length).setValues(progressData);
  } else {
    Logger.log("Nem volt szükség frissítésre a RAW és PROGRESS között.")
  }
}

//--------------------------------------------

function updateActionColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.RAW);
  const actionSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ACTION);

  const rawData = rawSheet.getDataRange().getValues();
  const rawHeaders = rawData.shift();
  const rawIdIndex = rawHeaders.indexOf(CONFIG.HEADERS.ID);

  const actionData = actionSheet.getDataRange().getValues();
  const actionHeaders = actionData.shift();
  const actionIdIndex = actionHeaders.indexOf(CONFIG.HEADERS.ID);

  if (actionIdIndex == -1) {
    SpreadsheetApp.getUi().alert("Hiba: Az 'Erőmű azonosító' oszlop nem található valamelyik sheeten!")
    return;
  }

  //keressük meg a változtatásra szoruló oszlopok indexeit mindekét sheeten
  const columnMap = CONFIG.HEADERS.SYNC_COLUMNS.map( colName => {
    return {
      name: colName,
      rawIndex: rawHeaders.indexOf(colName),
      actionIndex: actionHeaders.indexOf(colName)
    };
  }).filter( col => col.rawIndex !== -1 && col.actionIndex !== -1);

  if (columnMap.length === 0) {
    Logger.log("Nincs közös oszlop, amit frissíteni kéne a RAW és ACTION között.");
    return;
  }

  //RAW adatbázis felépítése (Map) a gyors kereséshez
  const rawMap = new Map();
  rawData.forEach( row => {
    const id = String(row[rawIdIndex]);
    if (id) rawMap.set(id, row);
  })

  let changesCount = 0;

  // Végigiterálunk a action adatokon (sorról sorra) és összehasonlítunk
  actionData.forEach((actionRow) => {
    const actionId = String(actionRow[actionIdIndex]);
    if (actionId && rawMap.has(actionId)) {
      const rawRow = rawMap.get(actionId);

      columnMap.forEach( colInfo => {
        const rawValue = rawRow[colInfo.rawIndex];
        const actionValue = actionRow[colInfo.actionIndex];

        if (rawValue !== actionValue) {
          actionRow[colInfo.actionIndex] = rawValue;
          changesCount++;
        }
      });
    }
  });

  //Tömeges visszaírás (Batch Write)
  // Csak akkor írunk, ha volt változás
  if (changesCount > 0) {
        actionSheet.getRange(2, 1, actionData.length, actionData[0].length).setValues(actionData);
  } else {
    Logger.log("Nem volt szükség frissítésre a RAW és action között.")
  }
}

//Oszlop index megkeresése név alapján (0-tól indul)
function getColIndex(sheet, headerName) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const index = headers.indexOf(headerName);
  if (index === -1) throw new Error(`Nem található ilyen oszlop: ${headerName} a ${sheet.getName()} lapon.`);
  return index;
}

//Okosdátum konverzió, hogy minden esetleges elírást felismerjen
  function parseMyDate(value) {
    if (typeof value === 'string') {
      value = value.trim();

      // 1. ESET: Teljes dátum ponttal (pl. "2026.02.28" vagy "2026.02.28.")
      // A regex magyarázata: ^(4 számjegy).(2 számjegy).(2 számjegy)
      const fullDateMatch = value.match(/^(\d{4})\.(\d{2})\.(\d{2})/);
      if (fullDateMatch) {
        // A JS-ben a hónap 0-tól indul (0=Január, 11=December), ezért ki kell vonni 1-et
        return new Date(fullDateMatch[1], fullDateMatch[2] - 1, fullDateMatch[3]);
      }

      // 2. ESET: Rövid dátum (pl. "02.28" vagy "02.28.")
      if (/^\d{2}\.\d{2}\.?$/.test(value)) {
        let cleanVal = value.replace(/\.$/, ''); 
        return new Date(`${currentYear}-${cleanVal.replace('.', '-')}`);
      }
    }
    return value;
  }

function processArchiveDataForSheets(sourceSheet, archiveSheet, rawIds, sheetName) {
  if (sourceSheet.getLastRow < 2) return;

  const data = sourceSheet.getDataRange().getValues();
  const headers = data.shift();
  const idIndex = headers.indexOf(CONFIG.HEADERS.ID);

  if (idIndex === -1) {
    Logger.log(`Hiba: Nem található ID oszlop a ${sheetName} lapon.`);
    return;
  }

  const rowsToArchive = [];
  const rowsToKeep = [];

  data.forEach(row => {
    const id = String(row[idIndex]);
    
    if (id && !rawIds.has(id)) {
      rowsToArchive.push(row);
    } else {
      rowsToKeep.push(row);
    }
  });

  if (rowsToArchive.length > 0) {
    archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, rowsToArchive.length, rowsToArchive[0].length).setValues(rowsToArchive);
    Logger.log(`${rowsToArchive.length} sor archiválva a(z) ${sheetName} lapról.`);
  }

  if (rowsToArchive.length > 0) {
    sourceSheet.getRange(2, 1, sourceSheet.getLastRow(), sourceSheet.getLastColumn()).clearContent();
    
    if (rowsToKeep.length > 0) { 
      sourceSheet.getRange(2, 1, rowsToKeep.length, rowsToKeep[0].length).setValues(rowsToKeep);
      
      const dateIndex = headers.indexOf(CONFIG.HEADERS.DATE); // Vagy DATE_1
      if (dateIndex !== -1) {
         sourceSheet.getRange(2, dateIndex + 1, rowsToKeep.length, 1).setNumberFormat('yyyy-MM-dd');
      }
    }
  }
}




