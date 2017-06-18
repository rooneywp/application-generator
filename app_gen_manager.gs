function onOpen(){
  var ui = SpreadsheetApp.getUi()
  var menu = ui.createMenu('Management')
  menu.addItem('Run mail merge', 'runMailMerge')
  menu.addToUi()
}
 
function runMailMerge(){
  var settings = {
    SOURCE_TEMPLATE: "179iu1WPOvYVVNd_mCo8yVJhsxKUaUSQOATKeNdERw_o", // The template into which data will be merged.
    SOURCE_WORKBOOK: '1T7oi055_4hHWoUHyxt1l5m6SqKqi-mAF0UbfLnjr5Rw', // The Sheets workbook in which the Sheet with the data resides.
    SOURCE_SHEET: '1455047377', // The Sheet with the data (in SOURCE_WORKBOOK).
    TARGET_FOLDER: '0B0VywFz1LUz0a0ViQWFtWGszVnc' // Where the completed files go.
  }
  var result = GSATETalentApplicantDocMaker.generateForAllRecords(settings)
  var result_message = 'Created ' + result[1] + ' new applicant documents. Details: \n\n'
  for (r=0; r<result[0].length; r++){
    result_message = result_message + result[0][r] + '\n'
  }
  var ui = SpreadsheetApp.getUi().alert(result_message)
}
