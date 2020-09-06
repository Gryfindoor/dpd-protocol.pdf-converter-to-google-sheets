var FOLDER_ID = "";
var SS = "";
var SHEET = "";

function extractStudentIDsAndSectionToSheets(){
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var folder = DriveApp.getFolderById(FOLDER_ID);
  var files = folder.getFilesByType("application/pdf");
  while(files.hasNext()){
  var fileID = files.next().getId();
  var doc = getTextFromPDF(fileID);
  var studentIDs = extractStudentIDs(doc.text.toString()); 
  importToSpreadsheet(studentIDs);
  Drive.Files.remove(fileID);
  
  }
};

function getTextFromPDF(fileID) {
  var blob = DriveApp.getFileById(fileID).getBlob();
  var resource = {
    title: blob.getName(),
    mimeType: blob.getContentType()
  };
  var options = {
    ocr: true
  };
  var file = Drive.Files.insert(resource, blob, options);
  
  var doc = DocumentApp.openById(file.id);
  var text = doc.getBody().getText();
  var title = doc.getName();
  Drive.Files.remove(doc.getId());
  
  return {
    name:title,
    text:text
  };
}

function extractStudentIDs(text){
  var regexp = /\d{13}[A-Z]{1}/g;
  var array = [...text.matchAll(regexp)];
  var spliceArray = array.splice(0, Math.ceil(array.length /2));
  
  Logger.log(spliceArray);
  var companyArray = [];
  for(var y=0; y < spliceArray.length; y++){
    var posistion = text.indexOf(spliceArray[y])+19;
    var model = text.substring(posistion);
    for(var x = 0; x < 2; x++){
      var kkkgg = model.replace(/,/, '');
      model = kkkgg;
    }
    var split = model.slice(0, model.indexOf(",")-(model.length));
    companyArray.push(split);
  }
  return {code:spliceArray, company:companyArray};
};

function importToSpreadsheet(code){
  var sheet = SpreadsheetApp.openById(SS).getSheetByName(SHEET);
  for(var x = 0; x <= code.code.length-1; x++){
    sheet.appendRow([code.code[x][0],code.company[x]]);
  }
}
