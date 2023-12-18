//Author: Sam Kasimalla;

var inputSpreadSheetId = ''
var mainModelSpreadSheetId = ''
var slide1Id = ''
var slide2Id = ''
var namesOfFilesLocal ={};


function main() {

  //open once and pass it around until function ends
  openedSpreadsSheet = SpreadsheetApp.openById(mainModelSpreadSheetId);
  openedPresentation = SlidesApp.openById(slide1Id);

  Logger.log('Main Model ID is ' + mainModelSpreadSheetId)
  Logger.log('Slide 1 Id is ' + slide1Id)
  Logger.log('inputSpreadSheetId Id is ' + inputSpreadSheetId)
  //updateRefInMainModel(openedSpreadsSheet, mainModelSpreadSheetId);
  //updateChartsInSlides(openedSpreadsSheet, openedPresentation);
  updateTablesInSlide(openedSpreadsSheet, openedPresentation);
}


function copyFolder(sourceFolderId, name, namesOfFiles) {
  namesOfFilesLocal = namesOfFiles;
  var targetFolder = createFolderInRoot_(name);
  var sourceFolder = DriveApp.getFolderById(sourceFolderId);

  // Start the folder copy
  copyFolderContents_(sourceFolder, targetFolder);
  var output = { "mainModelSpreadSheetId": mainModelSpreadSheetId, "slide1Id": slide1Id, "slide2Id": slide2Id, "inputSpreadSheetId":inputSpreadSheetId };
  return output;
}

function createFolderInRoot_(name) {
  if (!name)
    name = randomStr_(5);
  var newFolderName =  name;  // Change this to your desired folder name
  return DriveApp.createFolder(newFolderName);
}


function copyFolderContents_(source, target) {
  var folders = source.getFolders();
  var files = source.getFiles();

  // Copy folders and their contents
  while (folders.hasNext()) {
    var subFolder = folders.next();
    var folderCopy = target.createFolder(subFolder.getName());
    copyFolderContents_(subFolder, folderCopy); //Recursive
  }

  // Copy files
  while (files.hasNext()) {
    var file = files.next();
    var newFile = file.makeCopy(file.getName(), target);
    Logger.log('New file is ' + JSON.stringify(newFile.getName() + ' ' + newFile.getId()))

    assignNames(newFile);
  }
}

function listFilesInSameFolder() {

  var fileId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var file = DriveApp.getFileById(fileId);

  var parentFolders = file.getParents();

  // Assuming the file has at least one parent folder
  if (parentFolders.hasNext()) {
    var parentFolder = parentFolders.next();
    var files = parentFolder.getFiles();

    while (files.hasNext()) {
      var siblingFile = files.next();
      Logger.log(siblingFile.getName());
    }
  }
}

function assignNames(newFile) {

  if (newFile.getName().includes(namesOfFilesLocal.inputSpreadSheetId))
    inputSpreadSheetId = newFile.getId();  

  else if (newFile.getName().includes(namesOfFilesLocal.mainModelSpreadSheetId))
    mainModelSpreadSheetId = newFile.getId();

  else if (newFile.getName().includes(namesOfFilesLocal.slide1Id))
    slide1Id = newFile.getId();


}


function randomStr_(m) {
  var m = m || 15;
  s = '', r = 'abcdefghijklmnopqrstuvwxyz0123456789';
  for (var i = 0; i < m; i++) { s += r.charAt(Math.floor(Math.random() * r.length)); }
  Logger.log(s)
  return s;
};


// Sam Kasimalla;

//this updates all charts embedded from sheets on all slides of the specified slides id to a specified spreadsheet id
//Intended for use immediately after duplicating both the slide/presentation and the spreadsheet
//For now you must manually enter / cut and paste the ID of both files directly below in the two Id variables inside the quotes

function updateChartsInSlides(openedSpreadSheet, openedPresentation) {
  //sheet and slide id's - the charts in the slide id listed here will be linked to the the sheet id listed here:
  var slideId = slide1Id
  var spreadsheetId = mainModelSpreadSheetId

  //get all charts from all sheets of the spreadsheet copy
  var sheets2 = openedSpreadSheet.getSheets();
  var allCharts = [];//keep track of all charts on all sheets of the spreadsheet
  var allChartsIds = [];//keep track of the ids for all the charts
  var chartsnum = 0;
  for (var i = 0; i < sheets2.length; i++) {
    var curSheet2 = sheets2[i];
    var charts2 = curSheet2.getCharts();
    for (var j = 0; j < charts2.length; j++) {
      var curChart2 = charts2[j];
      allCharts[chartsnum] = curChart2;
      allChartsIds[chartsnum] = curChart2.getChartId();
      chartsnum++;
      Logger.log('Iterating charts in the sheetm id is ' + curChart2.getChartId());
    }
  }

  //total number of charts in all sheets of the spreadsheet
  var lengthAllCharts = allCharts.length;

  var slides2 = openedPresentation.getSlides();

  for (var i = 0; i < slides2.length; i++) {
    var curSlide2 = slides2[i];
    var charts2 = curSlide2.getSheetsCharts();
    for (var j = 0; j < charts2.length; j++) {
      var curChart2 = charts2[j];
      var chartHeight = curChart2.getHeight();
      var chartWidth = curChart2.getWidth();
      var chartLeft = curChart2.getLeft();
      var chartTop = curChart2.getTop();

      for (var k = 0; k < lengthAllCharts; k++) {
        if (curChart2.getChartId() == allChartsIds[k]) {
          var chart2 = allCharts[k];
          break;
        }
      }
      //Logger.log('\n chart2 ObjId is ' + curChart2.getObjectId());
      Logger.log('Iterating charts in the slides, id is ' + curChart2.getChartId());
      //Logger.log('\n chart2 Chart data ' + curChart2.getSpreadsheetId());
      curChart2.remove();
      curSlide2.insertSheetsChart(chart2, chartLeft, chartTop, chartWidth, chartHeight);
    }

  }


}




//Sam Kasimalla
function updateTablesInSlides(spreadsheet, presentation) {
  
  var sheet = spreadsheet.getSheetByName('Mapping');
  var range = sheet.getRange(2,1,100,6);
  var valuesMapping = range.getValues();

  var slides = presentation.getSlides();

  for (var iOuter = 0; iOuter < valuesMapping.length; iOuter++) {
    if (!valuesMapping[iOuter][0])
      break;
    
    //Logger.log('Mapping values from sheet are '+valuesMapping[iOuter][0]+valuesMapping[iOuter][1]+valuesMapping[iOuter][2]+valuesMapping[iOuter][3] )
    Logger.log('Mapping values from sheet are '+ JSON.stringify(valuesMapping[iOuter]) )
    
    var fromSheet = spreadsheet.getSheetByName(valuesMapping[iOuter][0]);

    //Now get the slide table as object
    Logger.log('(valuesMapping[iOuter][4]) is '+ (vMV = valuesMapping[iOuter][4]))
    tableNumber=valuesMapping[iOuter][5]
    var slide = slides[vMV-1];//Zero based index
    Logger.log('The number of tables is :'+slide.getTables().length)
    var table = slide.getTables()[tableNumber-1]; //Zero based index

    var numRows = table.getNumRows();
    var numCols = table.getNumColumns();

    //This depends on the accuracy provided in the mapping
    var range = fromSheet.getRange(valuesMapping[iOuter][1], valuesMapping[iOuter][2], numRows, numCols);
    var values = range.getDisplayValues();


    // Loop through rows and columns and print the table
    for (var i = 0; i < numRows; i++) {
      var rowString = "";
      var afterRowString = "";
      for (var j = 0; j < numCols; j++) {
        try {
          //props = table.getCell(i, j).ge
          rowString = table.getCell(i, j).getText().asString() + "\t";
          v = ''
          try {
            v = values[i][j]
          } catch (e) {
            Logger.log('Error retrieving' + e)
          }

          table.getCell(i, j).getText().setText(v);
          afterRowString = table.getCell(i, j).getText().asString() + "\t";

        } catch (e) {
          Logger.log(e)
        }

      }
      Logger.log('Before:'+rowString+ ' After' + afterRowString);
      

    }

  }

}


function editOneCell(spreadSheet,sheetName,range,val) {

var sheet = spreadSheet.getSheetByName(sheetName);
sheet.getRange(range).setValue(val);

}

//I am NOT using anything below this line

function updateChartsInSlide_() {
  var pr = SlidesApp.openById(new_slide);
  var slides = pr.getSlides();

  let pages = [];
  pages.push(slides[3].getObjectId());

  i = 0;

  var ss = SpreadsheetApp.openById(new_spreadsheet);
  var sh = ss.getSheetByName('A Accelerated software delivery');
  var rg = sh.getRange(2,2,5,5);
  var values = rg.getValues();

  Logger.log(l=sh.getCharts().length)
  chartId = sh.getCharts()[l-1].getChartId()

  old = old_spreadsheet
  newS = new_spreadsheet

  var r1 = createTable_(slides[0].getObjectId(), values);
  var r2 = replaceAllText_(newS, pages, old);
  var r3 = replaceAllShapesWithSheetsChart_('chart',newS, pages, chartId);

  var resp1 = Slides.Presentations.batchUpdate({ requests: [r3] }, pr.getId());
  Logger.log(JSON.stringify(resp1));
}



function replaceAllShapesWithSheetsChart_(textToReplace, newS, pages, chartId) {
  return {
    "replaceAllShapesWithSheetsChart": {
      "containsText": {
        "text": textToReplace
      },
      "spreadsheetId": newS,
      "linkingMode": "LINKED",
      "pageObjectIds": pages,
      "chartId": chartId,
    }
  }
}

function replaceAllText_(newS, pages, textToReplace) {
  return {
    "replaceAllText": {
      "containsText": {
        "text": textToReplace
      },
      "replaceText": newS,
      "pageObjectIds": pages,
    }
  }

}


function createTable_(slideObjectId, values) {

  return { "createTable": { "elementProperties": { "pageObjectId": slideObjectId }, "rows": values.length, "columns": values[0].length } }
}






