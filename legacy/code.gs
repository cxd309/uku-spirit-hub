function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Spirit Data Merge')
      .addItem('Build Event Catalogue', 'buildEventCatalogue')
      .addItem('Merge Data', 'mergeSotgData')
      .addToUi();
}

function mergeSotgData() {
  var ms = SpreadsheetApp.getActiveSpreadsheet(); // ms = merge Spreadsheet
  var eventCatSheet = ms.getSheetByName("EventCatalogue");  
  var breakdownMergeSheet = ms.getSheetByName("BreakdownMerge");

  var eventDetails = eventCatSheet.getRange(2,1,eventCatSheet.getLastRow()-1,8).getValues();

  for(var i = 0 ; i < eventDetails.length ; i++) {
    if(eventDetails[i][7]=="Yes") {
      var fileID = eventDetails[i][4];
      breakdownSheetName = eventDetails[i][6];
      var ss = SpreadsheetApp.openById(fileID);
      var breakdownSheet = ss.getSheetByName(breakdownSheetName);

      // Get the SOTG breakdown data
      var newDataRange = breakdownSheet.getRange(2,1,breakdownSheet.getLastRow()-1,7);
      var newSotgData = newDataRange.getValues();

      // Create single column array of the appropriate length with the breakdownSheetName in each row
      var breakdownSheetNameColumn = [];
      for(var j = 0 ; j < newSotgData.length ; j++) {
        breakdownSheetNameColumn[j] = [];
        breakdownSheetNameColumn[j][0] = breakdownSheetName;
      }

      // Append to the merged data sheet 
      var appendRow = breakdownMergeSheet.getLastRow();
      breakdownMergeSheet.getRange(appendRow+1,1,newSotgData.length,1).setValues(breakdownSheetNameColumn);
      breakdownMergeSheet.getRange(appendRow+1,2,newSotgData.length,7).setValues(newSotgData);


    }

  }
  

}

function testOneFolderByName() {
  //var ss = SpreadsheetApp.getActiveSpreadsheet();
  //var file = DriveApp.getFileById(ss.getId());
  var folders = DriveApp.getFoldersByName("20220903 Nationals");

  while (folders.hasNext()) { 
    var folder = folders.next();
    var folderName = folder.getName();
    processFilesInFolder(folder);
  }
}

function buildEventCatalogue() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var file = DriveApp.getFileById(ss.getId());
  var parentFolders = file.getParents();
  while(parentFolders.hasNext()) {
    var parentFolder = parentFolders.next();
    // var folderName = parentFolder.getName();

    var folders = parentFolder.getFolders();
    while(folders.hasNext()) {
      var folder = folders.next();
      catalogueFilesInFolder(folder);
    }

  }
}

// This is an iterative function that runs itself on subfolders within a folder

function catalogueFilesInFolder(folder) {
  var ms = SpreadsheetApp.getActiveSpreadsheet(); // ms = merge Spreadsheet
  var eventCatSheet = ms.getSheetByName("EventCatalogue");

  // Do the iterative subfolders thing
  // First find the colletion of subfolders
  var folders = folder.getFolders();
  while (folders.hasNext()) {
    var childFolder = folders.next();
    catalogueFilesInFolder(childFolder);
  }

  var files = folder.getFiles();  // gets the iterator of files in the current folder
  var folderName = folder.getName();

  while (files.hasNext()) {
    var file = files.next();
    var fileURL = file.getUrl();  // probs don't need this
    var fileID = file.getId();
    var fileType = file.getMimeType();
    if(fileType=="application/vnd.google-apps.spreadsheet") {

      var ss = SpreadsheetApp.openById(fileID);
      var ssName = ss.getName();
      var appendRow = eventCatSheet.getLastRow()+1;

      // check if filename has the word "Breakdown"
      //str.toLowerCase().includes('Stark'.toLowerCase());
      if(ssName.toLowerCase().includes("breakdown")==true) {
        // Add file data to the Event Catalogue list
        eventCatSheet.getRange(appendRow,1).setValue(folder.getName());
        eventCatSheet.getRange(appendRow,2).setValue(folder.getId());
        eventCatSheet.getRange(appendRow,3).setValue(folder.getUrl());
        eventCatSheet.getRange(appendRow,4).setValue(ssName);
        eventCatSheet.getRange(appendRow,5).setValue(fileID);
        eventCatSheet.getRange(appendRow,6).setValue(fileURL);
      }
      
      // check if file has a Breakdown tab with correct naming convention
      var breakdownSheetName = ssName.replace("Spirit Results and ","");
      var breakdownSheet = ss.getSheetByName(breakdownSheetName);
      if(breakdownSheet != null) {
        // Add file data to the Event Catalogue list
        eventCatSheet.getRange(appendRow,7).setValue(breakdownSheetName);
      }

      // loop through sheets and find one with the word breakdown in the name
      var sheets = ss.getSheets();
      for (var i = 0 ; i < sheets.length ; i++) {
        var sheet = sheets[i];
        var sheetName = sheet.getName();
        if(sheetName.toLowerCase().includes("breakdown")==true) {
          eventCatSheet.getRange(appendRow,7).setValue(sheetName);
        }

      }

    }

    }

  }



function processFilesInFolder(folder) {
  var ms = SpreadsheetApp.getActiveSpreadsheet(); // ms = merge Spreadsheet
  var eventListSheet = ms.getSheetByName("EventList");
  var breakdownMergeSheet = ms.getSheetByName("BreakdownMerge");
  
  var files = folder.getFiles();  // gets the iterator of files in the current folder
  while (files.hasNext()) {
    var file = files.next();
    var fileURL = file.getUrl();  // probs don't need this
    var fileID = file.getId();
    var ss = SpreadsheetApp.openById(fileID);
    var ssName = ss.getName();

    // check if file has a Breakdown tab
    var breakdownSheetName = ssName.replace("Spirit Results and ","");
    var breakdownSheet = ss.getSheetByName(breakdownSheetName);
    if(breakdownSheet != null) {
      // Check if already processed this file


      // Get the SOTG breakdown data
      var newDataRange = breakdownSheet.getRange(2,1,breakdownSheet.getLastRow()-1,7);
      var newSotgData = newDataRange.getValues();

      // Append to the merged data sheet and add file ID (and URL for manual acccess) to the merge list
      var appendRow = breakdownMergeSheet.getLastRow();
      breakdownMergeSheet.getRange(appendRow+1,1,newSotgData.length,7).setValues(newSotgData);

    }
  }
}




function insertFilenameFromURL() {
  var ms = SpreadsheetApp.getActiveSpreadsheet(); // ms = merge Spreadsheet
  var sheet = ms.getSheetByName("EventList");
  var nextURL = "";
  var ssName = "";
  var tryNextURL = true;
  var newSotgData = [];

  // Loop down values in col D
  var row = 2;
  do {
    nextURL = sheet.getRange(row,4).getValue();
    if(nextURL.toString().search("https://docs.google.com/spreadsheets/")==0) {
      tryNextURL = true;
      var ss = SpreadsheetApp.openByUrl(nextURL);
      ssName = ss.getName();
      newSotgData = appendBreakdown(nextURL);
      var resultsSheet = ms.getSheetByName("Results2");
      var appendRow = resultsSheet.getLastRow();
      resultsSheet.getRange(appendRow+1,1,newSotgData.length,7).setValues(newSotgData);
      
      if(sheet.getRange(row,3).getValue()=="") {
        sheet.getRange(row,3).setValue(ssName);
      }
  }
  else {
    tryNextURL = false;
  }
  
  row++;

  } while(tryNextURL==true);
  
}

function appendBreakdown(ssURL) {
  var ss = SpreadsheetApp.openByUrl(ssURL);
  var ssName = ss.getName();
  var breakdownSheetName = ssName.replace("Spirit Results and ","");
  var breakdownSheet = ss.getSheetByName(breakdownSheetName);
  // if (sheet != null) { }
  // means you need to split getting the sheetname and getting the actual data into two functions 
  // so that you don't do any processing if the file doesn't have a breakdown sheet

  var newDataRange = breakdownSheet.getRange(2,1,breakdownSheet.getLastRow()-1,7);
  var newSotgData = newDataRange.getValues();

  return newSotgData;

}
