/*
* Do not allow sheet to be modified until surplus order has been taken care of.
*
*/

function activateSurplusMode() {
  
  //Define variables
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var surplusSheet = ss.getSheetByName("surplus");
  var date = Utilities.formatDate(new Date(), 'America/New_York', 'MM-dd-yyyy');
  var surplusNote = date + ": Surplus order placed. Please do not add anything to the surplus list until the pickup is complete.";
  var selectedRange = surplusSheet.getRange(1, 1);
  console.log(date + " surplus note: " + selectedRange.getValues().toString());
  
  //Print PDF
  printPDF(ss, date);
  
  //Set note
  selectedRange.setValue(surplusNote);
  
  //Protect sheet so editors are warned
  surplusSheet.protect().setWarningOnly(true);
}

/*
* Revert sheet to normal
*
*/

function deactivateSurplusMode() {
  
  //Define variables
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();    
  var surplusSheet = ss.getSheetByName("surplus");
  var surplusSheetRange = surplusSheet.getRange(1, 1);
  var lastOrderNote = surplusSheetRange.getValues()[0].toString();//.indexOf(":"); 
  var lastSurplusDate = lastOrderNote.substring(0, lastOrderNote.indexOf(":"));
  
  //Confirm date is valid
  var dateWrapper = new Date(lastSurplusDate);
  if (isNaN(dateWrapper.getDate())){
    lastSurplusDate = "";
  }
  
  //Add note for other consultants
  var surplusNote = '="Please add your surplus item to this list whenever it is placed in the Denny 200A surplus area." & CHAR(10) & ' +
    '"Place a sticky note on the computer with ""hard drive removed"" on it so we know it is good to be surplus\'d. Last Surplus: ' + lastSurplusDate + '"';
  
  //Set note
  surplusSheetRange.setValue(surplusNote);
  
  //Remove sheet protection
  surplusSheet.protect().remove();
  
  //Delete surplus range  
  var userPrompt = ui.alert('Delete Surplus Items?', 'Are you ready to delete these surplus items and clear the sheet? Please confirm they were exported first.', ui.ButtonSet.YES_NO_CANCEL);
  
  if (userPrompt == ui.Button.YES){    
    //Delete surplus items, add headers
    surplusSheet.getRange("A3:G").clearContent();
    surplusSheet.getRange("A3:A7").setFontWeights([["Bold"],[""],["Bold"],[""],["Bold"]]);
    surplusSheet.getRange("A3:A7").setValues([["Computers:"],[""],["Monitors:"],[""],["Other:"]]);
    
  } else {    
    //Do nothing
    return;
  }      
}

/*
* Export the sheet as a PDF
* @param {Spreadsheet} ss The current Spreadsheet object
* @param {Date} date The date of the surplus order
* 
*/

function printPDF(ss, date){
  
  //Get sheet info, initialize variables
  var outputSheet = ss.getSheetByName('surplus');
  var parentFolder = DriveApp.getFileById('0Bz0QxsRjYluvc1U4d1Bxbm1OV1U');
  var createPDFFile = "";
  var parentFolderFilesCheck = "";
  var PDFBlob = "";
  var PDFName = "";
  var token = "";
  var response = "";
  var PDFurl = "";
  var htmlApp = "";
  var surplusRange = 'A2:G';
  
  //Name PDF
  PDFName = "CLAS OAT Surplus " + date;
  
  //Build PDF URL
  PDFurl = 'https://docs.google.com/spreadsheets/d/' + ss.getId() 
  + '/export?exportFormat=pdf&format=pdf' // export as pdf
  + '&size=letter'                           // paper size legal / letter / A4
  + '&currentdate=true'                           // date, does not work
  + '&portrait=false'                     // orientation, false for landscape
  + '&fitw=true'                        // fit to page width, false for actual size
  + '&sheetnames=true&printtitle=true' // hide optional headers and footers
  + '&pagenum=CENTER&gridlines=false' // hide page numbers and gridlines
  + '&fzr=false'                         // do not repeat row headers (frozen rows) on each page
  //+ '&top_margin=.75&bottom_margin=.75&left_margin=.25&right_margin=.25' //Narrow margins
  + '&gid=' + outputSheet.getSheetId()    // the sheet's Id
  + '&range=' + surplusRange;
  
  Logger.log(PDFurl);
  
  //Authorize script
  token = ScriptApp.getOAuthToken();
  
  //Request export url
  try {
    response = UrlFetchApp.fetch(PDFurl, {
      headers: {
        'Authorization': 'Bearer ' +  token
      }
    });    
  } catch (e) {
    //options = {muteHttpExceptions: true};
    Utilities.sleep(10000);
    response = UrlFetchApp.fetch(PDFurl, {
      headers: {
        'Authorization': 'Bearer ' +  token
      }
    });
  }
  
  //Name PDF blob
  PDFBlob = response.getBlob().setName(PDFName + '.pdf');
  
  //Create PDF file from blob
  try {
    createPDFFile = DriveApp.getFolderById(parentFolder.getId()).createFile(PDFBlob);   
  } catch (e) {    
    //Delete PDF if it already exists
    parentFolderFilesCheck = DriveApp.getFolderById(parentFolder.getId()).getFilesByName(PDFName);
    while (parentFolderFilesCheck.hasNext())
    {
      parentFolderFilesCheck.next().setTrashed(true);
    }
    createPDFFile = DriveApp.getFolderById(parentFolder.getId()).createFile(PDFBlob);   
  }
  
  createPDFFile.setDescription("Generated by the OAT Secondhand Spreadsheet Google Apps Script.\n\n" +
                               "Contact: RyanMcCallum@uncc.edu, achapin1@uncc.edu");
  
  htmlApp = HtmlService
  .createHtmlOutput('<!DOCTYPE html> <html> <body> <div style="text-align:center">Your surplus report have been saved to your Google Drive ' +
                    '<a href="' + parentFolder.getUrl() + '" target="_top" rel="noopener noreferrer">here</a>. </div>' +
    '<div style="text-align:center">You can view it directly <a href="' + createPDFFile.getUrl() + '" target="_top" rel="noopener noreferrer">here</a>.</div> ' +
      '<div style="text-align:center"> Please download and submit this PDF to the <a href="https://workflowforms.uncc.edu/imaging/imaging-forms-department/reporting-fixed-assets"' +
        'target="_top" rel="noopener noreferrer">Fixed Assets Disposition and Change Form</a>.</div> </body> </html>')
        .setTitle("Success")
        .setHeight(150);
  
  SpreadsheetApp.getActiveSpreadsheet().show(htmlApp);
}



