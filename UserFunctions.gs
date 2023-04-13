function stepOne() {

  mainFunc();
 //Create folder in users drive to store files
  createDriveFolderMMD();
  //Ask the user if their doc has headers - if no it will create them
  askUserForHeaders();
}

function stepTwo() {
   askForTemplate();
  //Have the user select a document to use as a template
}

function stepThree() {
  copyAndReplace();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('HCS Mail Merge');
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showSidebar(html);
}
