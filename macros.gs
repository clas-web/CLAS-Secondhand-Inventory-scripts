function AllConditionalFormattingRules() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  var conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getRange('A1:AE272')])
  .whenFormulaSatisfied('=($K1="Ready to Image")')
  .setBackground('#D9EAD3')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getRange('A1:AE272')])
  .whenFormulaSatisfied('=($K1="Add to SCCM")')
  .setBackground('#B6D7A8')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getRange('A1:AE272')])
  .whenFormulaSatisfied('=($K1="Imaged")')
  .setBackground('#93C47D')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getRange('A1:AE272')])
  .whenFormulaSatisfied('=($K1="Hold")')
  .setBackground('#FCE8B2')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getRange('A1:AE272')])
  .whenFormulaSatisfied('=($K1="Needs Repair")')
  .setBackground('#FCE5CD')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getRange('A1:AE272')])
  .whenFormulaSatisfied('=($K1="Reassigned")')
  .setBackground('#F4C7C3')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getRange('A1:AE272')])
  .whenFormulaSatisfied('=($K1="Unknown Location")')
  .setBackground('#FFF2CC')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getRange('A1:AE272')])
  .whenFormulaSatisfied('=($K1="Surplus")')
  .setBackground('#DD7E6B')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getRange('A1:AE272')])
  .whenFormulaSatisfied('=($K1="Prepare for Surplus")')
  .setBackground('#F6B26B')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
};

function sortStatusAtoZ() {
  //automatically sort the Computer spreadsheet by Status
  var spreadsheet = SpreadsheetApp.getActive();
  var range = spreadsheet.getRange('computers!K2');
  range.activate();
  if (range.getDisplayValue() == "Status"){
    Logger.log("range.getDisplayValue() is " + range.getDisplayValue());
    spreadsheet.getActiveSheet().getFilter().sort(11, true);
  }
};

function sortMonitorStatusAtoZ() {
  //automatically sort the Monitor spreadsheet by Status
  var spreadsheet = SpreadsheetApp.getActive();
  var range = spreadsheet.getRange('monitors!G1');
  range.activate();
  if (range.getDisplayValue() == "Status"){
    Logger.log("range.getDisplayValue() is " + range.getDisplayValue());
    spreadsheet.getActiveSheet().getFilter().sort(7, false);
  }
};

function CondFormat() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('O209').activate();
  var conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getRange('A1:P265')])
  .whenFormulaSatisfied('=($O1="On Hold")')
  .setBackground('#FCE8B2')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.splice(conditionalFormatRules.length - 1, 1, SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getRange('A1:P265')])
  .whenFormulaSatisfied('=($O1="In Progress")')
  .setBackground('#F6B26B')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
};