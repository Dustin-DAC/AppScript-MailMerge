

function onOpen() {
  //Create the menu - currently designed for sheetNameing
  var ui = SpreadsheetApp.getUi()
    ui.createAddonMenu()
      .addItem('Sidebar', 'showSidebar')
      .addSeparator()
      .addSubMenu(ui.createMenu('Execution Steps')
        .addItem('Step 1', 'stepOne')
        .addItem('Step 2', 'stepTwo')
        .addItem('Step 3', 'stepThree'))
      .addToUi();
}


//--------------------------------Main---------------------------------------//

  function mainFunc() {
    //Set sheet locations to global variables for re-use
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var firstSheet = spreadsheet.getSheets()[0];
    var secondSheet = spreadsheet.getSheets()[1];
    var thirdSheet = spreadsheet.getSheets()[2];
    var mainData = firstSheet.getDataRange().getValues();
      //Main Headers
    var header = mainData[0]
    //Array of sheets names
    var sheetArr = [firstSheet, secondSheet, thirdSheet]
    //For validation of sheet names (Config)
    var sheetNameArray = []
    

    for (var i = 0; i < sheetArr.length; i++) {
      //Log names of the initial sheets and errors up to 3 sheets
      try {
          if (sheetArr[i].getName != null){
            console.log("Sheet Log " + i + " name : " + sheetArr[i].getName());
            sheetNameArray.push(sheetArr[i].getName())
          }
      } catch(err) {
        console.log("Itteration " + i + " : " + err.message);
        }
    }

    //Creates Configuration and Confirmation Sheets
    var checkConfig = sheetNameArray.includes("Configuration_Data");
    if (checkConfig == false){
            dataSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
            dataSheet.setName("Configuration_Data");
            SpreadsheetApp.getActive().getSheetByName("Configuration_Data").hideSheet();  
            console.log("Config Sheet did not exist");
    }
      var checkConfig = sheetNameArray.includes("Confirmation");
        if (checkConfig == false){
            dataSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
            dataSheet.setName("Confirmation");
              for (var i = 0; i < 3; i++){
              var confirmationHeaders = [ 'Draft_Created','Draft_Location','Email_Sent' ]  
              var currentCell = ['A1', 'B1', 'C1'];
              dataSheet.getRange(currentCell[i]).setValue(confirmationHeaders[i]);
              }
    }
    
    return firstSheet;
  }

//-----------------------------------------------------------------------//
//----------------------------Headers---------------------------------------//

  //User prompt asking for headers. If user doesn't have any headers it creates a standard set.
  function askUserForHeaders() { 

    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var firstSheet = spreadsheet.getSheets()[0];
    spreadsheet.setActiveSheet(firstSheet);
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('Do you already have headers in your document? Please be aware that certain headers need to be formatted properly for all aspects of the addon to work such as (Email, Last_Name).', ui.ButtonSet.YES_NO)
    if (response == ui.Button.YES) {
      validateHeaders()
    } 
    
    if (response == ui.Button.NO) {
      console.log("User does not have headers")
      // Add a line to the top of the sheet and create headers
      firstSheet.insertRowBefore(1) 
      var controlHeaders = [ 'First_Name', 'Last_Name', 'Email', 'Student_Name', 'Student_ID'] 
      var controlFields = ['A1', 'B1', 'C1', 'D1', 'E1']
      for (var i = 0; i < controlHeaders.length; i++) {
        SpreadsheetApp.getActive().getRange(controlFields[i]).setValue(controlHeaders[i])
      }
      ui.alert('We will add a set of sample headers to your sheet. Please adjust the data columns accordingly.')
    }

  }

  //Checks Header Values and adds columns if needed. 
  function validateHeaders() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var firstSheet = spreadsheet.getSheets()[0];
    //Validates the headers to standard values of the template sheet
    var sheet = SpreadsheetApp.getActive().getSheetByName(firstSheet.getName());
    var data = sheet.getDataRange().getValues();

    //Pre-defined headers 
    
    //Single out the header row from data var
    var headerValidation = data[0]
    var ui = SpreadsheetApp.getUi();
    var missingColumnsMessage = "The following Columns have been added to your sheet. If you already have these fields please match the format to what has been provided and remove any duplicate columns. ";
    var lastNameCheck = 0
    var emailCheck = 0

    for (var i = 0; i < headerValidation.length; i++){
        var controlHeaders = [ 'First_Name', 'Last_Name', 'Email', 'Student_Name', 'Student_ID'] 
        if (controlHeaders[i] == headerValidation[i]) {
          console.log("Control Item " + controlHeaders[i] + " is equal to " + headerValidation[i])
        } 
        if (lastNameCheck === 0){
        if (!(headerValidation.includes("Last_Name"))) {
          lastNameCheck = 1
          sheet.insertColumnBefore(1)
          sheet.getRange("A1").setValue("Last_Name")
          
        }}
        if (emailCheck === 0){
        if (!(headerValidation.includes("Email"))) {
          emailCheck = 1
          sheet.insertColumnBefore(1)
          sheet.getRange("A1").setValue("Email")
          
        }}
    }
    if (lastNameCheck === 1 || emailCheck === 1){
          missingColumnsMessage += " Missing :"
            if (lastNameCheck === 1) {missingColumnsMessage += " Last_Name"}
            if (emailCheck === 1) {missingColumnsMessage += " Email"}
          ui.alert(missingColumnsMessage);
    }
    if (lastNameCheck === 0 && emailCheck === 0){
        ui.alert("Your data is formatted correctly and you can move on to step two.")
    }
    
  }




//----------------------------Picker-------------------------------------//
   
   // Displays an HTML-service dialog in Google Sheets that contains client-side JavaScript code for the Google Picker API.
   //ID's for selected documents are stored in the config sheet for reference.
   

  function askForTemplate() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var firstSheet = spreadsheet.getSheets()[0];
    var mainData = firstSheet.getDataRange().getValues();
    var header = mainData[0]
    
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('Do you have a document you would like to use as a template?', ui.ButtonSet.YES_NO)
    if (response == ui.Button.YES) {
      showPicker();
    } else {
      var folderlocation = SpreadsheetApp.getActive().getRange("Configuration_Data!A2").getValue();
      var doc = DocumentApp.create("Empty_Template")
      DriveApp.getFileById(doc.getId()).moveTo(DriveApp.getFolderById(folderlocation))
      SpreadsheetApp.getActive().getRange("Configuration_Data!B1").setValue('Empty_Template');
      SpreadsheetApp.getActive().getRange("Configuration_Data!B2").setValue(doc.getId());
      Utilities.sleep(5000);
      addTemplateNotesNewDoc(SpreadsheetApp.getActive().getRange("Configuration_Data!B2").getValue(), header)
      }
  }

  function showPicker() {
    var html = HtmlService.createHtmlOutputFromFile('Picker.html')
      .setWidth(600)
      .setHeight(425)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getUi().showModalDialog(html, 'Select Your Template Doc - This window will close automatically. Please be patient as it does take a minute.');
  }

  function getOAuthToken() {
    DriveApp.getRootFolder();
    return ScriptApp.getOAuthToken();
  }

  function returnIdValue(value) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var firstSheet = spreadsheet.getSheets()[0];
    var header = firstSheet.getDataRange().getValues()
    if (value != null) {
    console.log("Picked ID Value : " + value);
    } else {
      console.log("The returned picker value is null" + value)
    }
      SpreadsheetApp.getActive().getRange("Configuration_Data!B1").setValue('Selected_Template');
      SpreadsheetApp.getActive().getRange("Configuration_Data!B2").setValue(value);
      Utilities.sleep(5000)
      addTemplateNotesNewDoc(SpreadsheetApp.getActive().getRange("Configuration_Data!B2").getValue(), header[0])
    
  }
  
//----------------------------Drive-------------------------------------//

  function createDriveFolderMMD() {

    var folderListing = DriveApp.getFolders();
    var folderListArray = []
    var folder 
    var folderID

    while (folderListing.hasNext()) {
      folder = folderListing.next();
      var folderNameQuery = folder.getName();
      folderListArray.push(folderNameQuery)
      if (folderNameQuery == "Mail Merge Drafts") { 
        folderID = folder.getId();
      }
    }
    if (folderListArray.includes("Mail Merge Drafts")) {
        SpreadsheetApp.getActive().getRange("Configuration_Data!A1").setValue('MMD_ID');
        SpreadsheetApp.getActive().getRange("Configuration_Data!A2").setValue(folderID);
      } else {
        //Create folder and dump info
        var newDriveFolder = DriveApp.createFolder("Mail Merge Drafts");
        SpreadsheetApp.getActive().getRange("Configuration_Data!A1").setValue('MMD_ID');
        SpreadsheetApp.getActive().getRange("Configuration_Data!A2").setValue(newDriveFolder.getId());
      }
  }

  function createNewFolder(name) {
    var folderListing = DriveApp.getFolders();
    var folderListArray = []
    var folder 
    var folderID

    while (folderListing.hasNext()) {
      folder = folderListing.next();
      var folderNameQuery = folder.getName();
      folderListArray.push(folderNameQuery)
      if (folderNameQuery == name) { 
        folderID = folder.getId();
      }
    }
    if (folderListArray.includes(name)) {
        SpreadsheetApp.getActive().getRange("Configuration_Data!A3").setValue(folderID);
      } else {
        //Create folder and dump info
        var newDriveFolder = DriveApp.createFolder(name);
        SpreadsheetApp.getActive().getRange("Configuration_Data!A3").setValue(newDriveFolder.getId());
      }
    newDriveFolder.moveTo(DriveApp.getFolderById(SpreadsheetApp.getActive().getRange("Configuration_Data!A2").getValue()))
  }

//----------------------------Docs--------------------------------------//
  //Functions used in 2.0 

  function selectTemplate() {
    // Prompt the user to select a template document using a file picker
    var template = DriveApp.getFileById(DriveApp.createFilePicker().pick());
    return template;
  }

  function copyAndReplace () {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var firstSheet = spreadsheet.getSheets()[0];
    //Copy and replace tags
    var data = SpreadsheetApp.getActive().getSheetByName(firstSheet.getName()).getDataRange().getValues();
    var head = data[0];
    //Email Questions for Run
    var ui = SpreadsheetApp.getUi()
    var subject
    var emailConfirmation = ui.alert("Would you like to send an email to all recipients?", ui.ButtonSet.YES_NO)
    if (emailConfirmation === ui.Button.YES){
      subject = SpreadsheetApp.getUi().prompt("Please enter the subject of the email")
      if (subject.getResponseText() === "") {SpreadsheetApp.getUi().alert("Your email subject is empty. This can cause issues with delivery so your email will not be sent.")}
      }


/// New Folder Creator for each iteration
    var folder = DriveApp.getFolderById(SpreadsheetApp.getActive().getRange("Configuration_Data!A2").getValue());
    var date = Utilities.formatDate(new Date(), "America/New_York", "MM-dd-yyyy--HH:mm:ss")

    if (emailConfirmation === ui.Button.YES) {
      if (subject.getResponseText() === "") {
        var folderName = "Blank" + " - " + date
      } else {var folderName = subject.getResponseText() + " - " + date}
            console.log(folderName)
      createNewFolder(folderName);
    } else { 
      var folderName = "Merge" + " - " + date
      console.log(folderName)
      createNewFolder(folderName);
    }
    var uniqueFolder = DriveApp.getFolderById(SpreadsheetApp.getActive().getRange("Configuration_Data!A3").getValue())

      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var rowData = {};
        for (var j = 0; j < head.length; j++) {
          rowData[head[j]] = row[j];
        }
        var documentCopy = DriveApp.getFileById(SpreadsheetApp.getActive().getRange("Configuration_Data!B2").getValue()).makeCopy();
        replaceTags(documentCopy, rowData);
        documentCopy.setName(rowData["Last_Name"] + " - Mail Merge");
        uniqueFolder.addFile(DriveApp.getFileById(documentCopy.getId()));
        //Adding Email Portion
          var body = DocumentApp.openById(documentCopy.getId()).getBody().getText()
          var email = rowData["Email"]
          if (emailConfirmation === ui.Button.YES && subject.getResponseText() != ""){
            GmailApp.sendEmail(email, subject.getResponseText(), body)
          }

        updateConfirmation(documentCopy.getName(), documentCopy.getUrl(), emailConfirmation)
      }
  }


  function replaceTags(documentCopy, rowData) {
    // Replace the tagged data in the document copy with information from the current row
    var doccopyid = documentCopy.getId();
    var body = DocumentApp.openById(doccopyid).getBody().editAsText();
    for (var header in rowData) {
      body.replaceText("<<" + header + ">>", rowData[header]);
    }
    console.log(body);
  }

  var templateDoc = {}; //Contains ID and Name key's via getOGtemplateDoc()
  var newFileTemplate = {}; //referenceable value for new file created from selected template


  function addTemplateNotesNewDoc(id, header) {
      var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      var firstSheet = spreadsheet.getSheets()[0];
      var docBody = DocumentApp.openById(id).getBody();
      var refresh = SpreadsheetApp.getActive().getSheetByName(firstSheet.getName()).getDataRange().getValues();
      var heads = refresh[0];

      if ((SpreadsheetApp.getActive().getRange("Configuration_Data!B1").getValue()) === "Empty_Template") {
        docBody.appendParagraph("Sample Template")
        docBody.appendParagraph("")
        docBody.appendParagraph("Dear Parent of <<Student_Name>>,")
        docBody.appendParagraph("")
        docBody.appendParagraph("They have been really great in class.")
        docBody.appendParagraph("")
        docBody.appendParagraph("Thank you, ")
        docBody.appendParagraph("Teacher")
        docBody.appendParagraph("")
      }
          docBody.appendParagraph("-----------------------------------------------------------")
      docBody.appendParagraph("The following key's will be used to dynamically fill the rest of the forms. Please keep the case and spacing (or lack of) the same as what is in this list. Once you are done please remove everything in this document between the lines. ");
      docBody.appendParagraph("")
          //Dynamic write key values into template document for reference
          function writeHeadersHash(head) {
            docBody.appendParagraph("<<"+head+">>")
          }
          heads.forEach(writeHeadersHash);
      docBody.appendParagraph("")
      docBody.appendParagraph("-----------------------------------------------------------")

      var folderlocation = SpreadsheetApp.getActive().getRange("Configuration_Data!A2").getValue();

      if (!(SpreadsheetApp.getActive().getRange("Configuration_Data!B2").getValue()) === "Empty_Template") {
      var newFile = DriveApp.getFileById(SpreadsheetApp.getActive().getRange("Configuration_Data!B2").getValue()).makeCopy(DriveApp.getFolderById(folderlocation))
      SpreadsheetApp.getActive().getRange("Configuration_Data!B2").setValue(newFile.getId());
      } else {DriveApp.getFileById(SpreadsheetApp.getActive().getRange("Configuration_Data!B2").getValue()).moveTo(DriveApp.getFolderById(folderlocation))}
      Utilities.sleep(3000)
      // Open the template in a new tab
        SpreadsheetApp.getUi().alert("Please allow all popup's. If they are currently blocked you will need to allow them for this site and re-run step two.");
        openUrl(DriveApp.getFileById(SpreadsheetApp.getActive().getRange("Configuration_Data!B2").getValue()).getUrl())
  }
//----------------------------Data Functions-----------------------------------------//

  function openUrl(url){
    var html = HtmlService.createHtmlOutput('<!DOCTYPE html><html><script>'
    +'window.close = function(){window.setTimeout(function(){google.script.host.close()},9)};'
    +'var a = document.createElement("a"); a.href="'+url+'"; a.target="_blank";'
    +'if(document.createEvent){'
    +'  var event=document.createEvent("MouseEvents");'
    +'  if(navigator.userAgent.toLowerCase().indexOf("firefox")>-1){window.document.body.append(a)}'                          
    +'  event.initEvent("click",true,true); a.dispatchEvent(event);'
    +'}else{ a.click() }'
    +'close();'
    +'</script>'
    // Offer URL as clickable link in case above code fails.
    +'<body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically.  Click below:<br/><a href="'+url+'" target="_blank" onclick="window.close()">Click here to proceed</a>.</body>'
    +'<script>google.script.host.setHeight(55);google.script.host.setWidth(410)</script>'
    +'</html>')
    .setWidth( 90 ).setHeight( 1 );
    SpreadsheetApp.getUi().showModalDialog( html, "Opening now, please be sure to allow popups..." );
  }

  //Updates the "Confirmation Sheet"
  function updateConfirmation(dc, dl, es) {
    sheet = SpreadsheetApp.getActive().getSheetByName("Confirmation")
    sheet.insertRowBefore(2)

    if (es != "YES"){es = "NO"}
    if (dc != "") {
    sheet.getRange("A2").setValue(dc)
    }
    if (dl != "") {
    sheet.getRange("B2").setValue(dl)
    }
    if (es != "") {
    sheet.getRange("C2").setValue(es)
    }
  }


//-----------------------------------------------------------------------//
