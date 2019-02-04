//https://developers.google.com/apps-script/guides/triggers/#Simple
//https://webapps.stackexchange.com/questions/103976/how-to-add-a-note-containing-date-to-a-cell-in-column-x-when-it-is-edited
//https://stackoverflow.com/questions/12583187/google-spreadsheet-script-check-if-edited-cell-is-in-a-specific-range
//https://stackoverflow.com/questions/12995262/how-can-i-return-the-range-of-an-edited-cell-using-an-onedit-event

//create a menu option for script functions
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('OAT Functions')
  .addItem('Form Responses: Format request for Cherwell', 'copyToCherwell')
  .addItem('Add Link to Cherwell', 'addCherwellLinkOptimized')
  .addToUi();
}


function onEdit(e){
  //edited cell gets passed into function
  var range = e.range;
  var dateCell = e;
  
  //*** My code
  //Create a date variable that automatically updates
  var date = new Date();
  
  //Returns the number of the edited row and column
  var thisRow = e.range.getRow();
  var thisCol = e.range.getColumn();
  
  //Get sheet name
  var sheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getSheetName();
  
  //Returns how many columns are in the sheet
  var lastColumn = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getLastColumn();
  
  //tracks which step of script to initiate
  var updater = 0;
  var counter = 0;
  //************************************************************************************************************************
  //Adds Cherwell link to range of edited row's "Completed By/Work Order" column entry  
  if (thisRow > 2 && ((sheetName == 'computers' && thisCol >= 14 && thisCol <= 16) || (sheetName == 'Form Responses 1' && thisCol == 14) || (sheetName == 'monitors' && thisCol >= 8 && thisCol <= 9) || (sheetName == 'surplus' && thisCol == 7))) {     
    addCherwellLinkOptimized(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(thisRow, thisCol));
  }
  //************************************************************************************************************************
  //*** Automatically update Age column
  
  //Confirm we're in the 'computers' sheet
  //if (sheetName == 'computers' || sheetName == 'Reserved (<4 years old)'){      
  if (sheetName == 'computers'){      
    counter = 2;     
    //Make sure 'Age' column header is in the right place, don't want to overwrite important data because someone changed the headers
    if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange("B2").getDisplayValue() == 'Age'){           
      //Counter prevents header row and above from getting modified
      if (thisRow > counter){  
        if (thisRow == counter+1){
          //Don't drag down header row, drag formula from next row upwards 
          var age = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(thisRow+1, 2).copyTo(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(thisRow, 2));
        } else {
          var age = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(thisRow-1, 2).copyTo(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(thisRow, 2));
        }
        Logger.log("thisRow is" +thisRow);
      }
    } else {    
      SpreadsheetApp.getActiveSpreadsheet().toast('ERROR: Age not updated', '*Age* is not in header row at same place', 10);          
    }
  } else {    
    //Not in the right sheet
    Logger.log('Not in computers sheet, did not modify Age');
  }
  //************************************************************************************************************************
  //*** Automatically update Warranty column
  
  //Confirm we're in the 'computers' sheet
  //if (sheetName == 'computers' || sheetName == 'Reserved (<4 years old)'){      
  if (sheetName == 'computers'){      
    counter = 2;  
    //Make sure 'Warranty Check' column header is in the right place, don't want to overwrite important data because someone changed the headers
    if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange("M2").getDisplayValue() == 'Warranty Check'){
      //Counter prevents header row and above from getting modified
      if (thisRow > counter){   
        if (thisRow == counter+1){
          //Don't drag down header row, drag formula from next row upwards 
          var warranty = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(thisRow+1, 13).copyTo(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(thisRow, 13));
        } else {
          var warranty = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(thisRow-1, 13).copyTo(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(thisRow, 13));
        }
      }
    } else {    
      SpreadsheetApp.getActiveSpreadsheet().toast('ERROR: Warranty not updated', '*Warranty Check* is not in header row at same place', 10);          
    }
  } else {
    //Not in the right sheet
    Logger.log('Not in computers sheet, did not modify Warranty');
  }
  //************************************************************************************************************************
  //*** Automatically update Date column, append Note with specific changes to column
  
  //Returns range of edited row's "Last Modified Date" column entry  
  var newDateRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(thisRow, 10);
  
  //Get the current note in newDateRange
  var note = newDateRange.getNote();
  
  //Make sure 'Date Last Modified' column header is in the right place in both 'computers' and 'monitors' sheets, don't want to overwrite important data because someone changed the headers
  if ((SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange("J2").getDisplayValue() == 'Date Last Modified (or verified)') || (SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange("J1").getDisplayValue() == 'Date')){
    updater = 1;    
    //The Date header is in different rows on the computer/monitor sheets, this function accounts for that so we don't erase it
    if (sheetName == 'computers' || sheetName == 'Reserved (<4 years old)'){      
      counter = 2;
    } else if (sheetName == 'monitors'){      
      counter = 1; 
    }  else {
      //Not in the right sheet
      counter = -1;
      Logger.log('Not in computers or monitors sheet, did not modify Date');
      //SpreadsheetApp.getActiveSpreadsheet().toast('ERROR: Date not updated', '*Date Last Modified (or verified)* is not in header row at same place', 10);            
      
    }    
    //Counter prevents header row and above from getting modified
    if (thisRow > counter && counter > 0) {          
      //Set update date in row's 'Date Last Modified' cell
      newDateRange.setValue(date);      
      //Update note with user and date of last modification
      newDateRange.setNote(note + '\n--- Cell '+ SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange(thisRow, thisCol).getA1Notation() + ' Modified ' + date/* + ' by: ' + Session.getActiveUser()*/);     
    }                 
  } else {
    //Do not modify if 'Date Last Modified' isn't in the proper location
    //SpreadsheetApp.getActiveSpreadsheet().toast('ERROR: Date not updated', '*Date Last Modified (or verified)* is not in header row at same place', 10);            
  }  
}
//************************************************************************************************************************ 
//Format secondhand computer request to make it easier to copy to a Cherwell ticket
function copyToCherwell(){
  //Select whichever row/request user has clicked on
  var range = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = range.getActiveCell().getRow();
  var col = range.getActiveCell().getColumn();
  Logger.log("row " + row);
  Logger.log("col " + col);
  Logger.log(range.getName());
  
  //Make sure this only works on the Form Responses sheet and after header row
  if (range.getName()=="Form Responses 1" && row > 1){
    
    //Format text, show in pop-up box
    var formattedText = Browser.msgBox("Copy this to your Cherwell ticket",
                                       "Request for secondhand computers:\\n"+
                                       "Submitted on: " + range.getRange(row, 2).getDisplayValue()+ "\\n"+
      "Submitted by: " + range.getRange(row, 3).getDisplayValue()+ "\\n"+
        "Dept: " + range.getRange(row, 4).getDisplayValue()+ "\\n"+
          "Location: " + range.getRange(row, 5).getDisplayValue() + " " + range.getRange(row, 6).getDisplayValue() + "\\n"+
            "#: " + range.getRange(row, 7).getDisplayValue()+ "\\n"+
              "User(s): " + range.getRange(row, 8).getDisplayValue()+ "\\n"+
                "Purpose: " + range.getRange(row, 9).getDisplayValue()+ "\\n"+
                  "Type: " + range.getRange(row, 10).getDisplayValue()+ "\\n"+
                    "Comments: " + range.getRange(row, 11).getDisplayValue()+ "\\n"+
                      "\\n***TEMPLATE RESPONSE***"+
                        "\\nHello, we've received your request for an upgraded machine. It should be complete in the next day or two, "+
                          "after which we will contact you so we can coordinate delivery. The user does not necessarily have to be present, but we do like to make sure they can log in without errors."+
                            "\\n\\nIf there is a machine being replaced or sent to surplus, please make sure any personal files are backed up because we will destroy the hard drive and surplus the computer after approximately 2 weeks."+
                              "\\n\\nDo you also require any other peripherals? Mouse, keyboard, speaker bar, etc.",
                                Browser.Buttons.OK);
  } else if (row < 2){  
    //Don't allow for header row
    SpreadsheetApp.getActiveSpreadsheet().toast("Please select a valid row.");    
  } else {    
    //Don't allow outside of sheet
    SpreadsheetApp.getActiveSpreadsheet().toast("This function will only work on the Form Responses sheet.");
  }
  /*This is the sheets function for doing the same thing
  ="Request for secondhand computers:
  Dept: "&D206&"
  Location: "&E206&" "&F206&"
  #: "&G206&"
  User(s): "&H206&"
  Purpose: "&I206&"
  Type: "&J206&"
  Comment: "&K206&""*/
}

/**
* Add Cherwell hyperlink to work orders and usernames to directly link to Cherwell search
* @param range The address of the cell to update (optional, if not included the selected range will be the range)
*/

function addCherwellLinkOptimized(range){
  var selected = range || SpreadsheetApp.getActiveSheet().getActiveRange();
  var values = selected.getValues();
  var arr = [values];
  
  for (var i = 0; i < values.length; i++) {
    arr[i] = [values[i]];
    for (var j = 0; j < values[i].length; j++) {
      if (values[i][j]){
        arr[i][j] = '=HYPERLINK' + '("https://cherwell.uncc.edu/CherwellClient/Access/Command/Queries.GoToRecord?BusObID=Incident&PublicID='+values[i][j]+'","'+values[i][j]+'")';
      } else {
        arr[i][j] = "";
      }      
    }
  }  
  //Set hyperlink   
  selected.setValues(arr);      
}