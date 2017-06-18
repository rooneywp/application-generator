// Script for combining Sheets data with a Document template to generate a "mail merge".
// Written with TTS Talent's needs in mind.

// Based on / refactored from: 
// https://github.com/inviqa/SysAdmin/commit/36ab74d81dd067fc15cb465731d5e339af1b18d6#diff-7ca2d7db437dbd45a6e8ea332f888e0c

// Settings object requires following attributes:
// SOURCE_TEMPLATE = Google document id from the document template (get id from the URL).
// SOURCE_WORKBOOK = The workbook holding the sheet that holds the data.
// SOURCE_SHEET = The sheet that holds the data.
// TARGET_FOLDER = The folder in which the new documents should be created.
//
// See test() for example.

var NUMBER_FOR_TESTING = null // Set to > 1 to tests a subset of your records.
function test(){
  var settings = {
    SOURCE_TEMPLATE: "179iu1WPOvYVVNd_mCo8yVJhsxKUaUSQOATKeNdERw_o",
    SOURCE_WORKBOOK: '1ISRZG678ofXGI1Qx1Si2IDE5QHoLChvoZjUSfzcbraQ',
    SOURCE_SHEET: '1455047377',
    TARGET_FOLDER: '0B4JAQaDaWjUaNF9MaFdrZjZ0Q3c'
  }
  Logger.log(generateForAllRecords(settings))
}

function generateForAllRecords(settings){
  var sheet = GSATEUtilitiesLibrary.getSheetById(SpreadsheetApp.openById(settings.SOURCE_WORKBOOK), settings.SOURCE_SHEET)
  var data = sheet.getSheetValues(1,1,sheet.getMaxRows(), sheet.getMaxColumns())
  var result = []
  var created = 0
  data = GSATEUtilitiesLibrary.makeJSON(data)
  if (NUMBER_FOR_TESTING){
    var len = NUMBER_FOR_TESTING
    } else {
      var len = data.length
      }
  for (i=1; i<len; i++){
    var doc_name = data[i].gs_level_applied_to + " - " + data[i].full_name + " Application"
    if (!DriveApp.getFolderById(settings.TARGET_FOLDER).getFilesByName(doc_name).hasNext()){
      generateForSingleRecord(data[i], doc_name, settings)
      result.push(doc_name + " created!")
      created += 1
    } else {
      result.push(doc_name + " already exists, no action taken!")
    }
  }
  return [result, created]
}

function generateForSingleRecord(record, doc_name, settings) {
  var target = createDuplicateDocument(settings.SOURCE_TEMPLATE, doc_name, settings);
  var keys = Object.keys(record)
  for(var i=0; i<keys.length; i++) {
    // In template, fields to replaced are noted with leading and closing colons, as in ':key:'.
    replaceString(target, ":" + keys[i] + ":", record[keys[i]] || "");
  }
}

/**
 * Duplicates a Google Apps doc
 *
 * @return a new document with a given name from the orignal
 */
function createDuplicateDocument(sourceId, name, settings) {
    var source = DriveApp.getFileById(sourceId);
    var targetFolder = DriveApp.getFolderById(settings.TARGET_FOLDER);
    var newFile = source.makeCopy(name, targetFolder);
    return DocumentApp.openById(newFile.getId());
}

/**
 * Search a String in the document and replaces it with a newString.
 */
function replaceString(doc, String, newString) {
  var ps = doc.getParagraphs();
  for(var i=0; i<ps.length; i++) {
    var p = ps[i];
    var text = p.getText();
    if(text.indexOf(String) >= 0) {
      //look if the String is present in the current paragraph
      p.editAsText().replaceText(String, newString);
    }
  } 
  return doc
}
