//Author: Sam Kasimalla;

const sourceFolderId = '1ktkBjOY42yTKq_EyeOXxc7IkBUwYTjww'
const prefix = "XYZBundle-"
const namesOfFiles = { "inputSpreadSheetId": "Input", "mainModelSpreadSheetId": "Model", "slide1Id":"Output Slides" };

function run() {
  
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var name = spreadSheet.getRange('D6').getValue();
  Logger.log('Value of D6 is'+name)
  name = prefix + name
  
  output = StandAloneValueEnggJAS.copyFolder(sourceFolderId, name, namesOfFiles );

  Logger.log('inputSpreadSheetId Id is ' + output.inputSpreadSheetId)
  Logger.log('Main Model ID is ' + output.mainModelSpreadSheetId)
  Logger.log('Slide 1 Id is ' + output.slide1Id)
  
  
  var mainSpreadSheet = SpreadsheetApp.openById(output.mainModelSpreadSheetId);
  
  StandAloneValueEnggJAS.editOneCell(mainSpreadSheet,'Inputs','B2',output.inputSpreadSheetId);
  StandAloneValueEnggJAS.editOneCell(mainSpreadSheet,'Inputs','B3',output.slide1Id);
  
  var presentation = SlidesApp.openById(output.slide1Id)

  StandAloneValueEnggJAS.updateChartsInSlides(mainSpreadSheet, presentation)

  //Browser.msgBox('Customer '+name+' folder created in your google drive root')
 
}

