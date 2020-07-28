const FOLDER_ID = "13IOecyfRgwkdphJr7lChQrDT2IeiErah"; //Folder ID of all PDFs
const SS = "13_euBdyshOplyxobP_v1geVwWg9GMfqO7ddYdmC-NoI";//The spreadsheet ID
const SHEET = "Arkusz1";//The sheet tab name

function extractStudentIDsAndSectionToSheets(){
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  //Get all PDF files:
  const folder = DriveApp.getFolderById(FOLDER_ID);
  //const files = folder.getFiles();
  const files = folder.getFilesByType("application/pdf");
  
  let allIDsAndCRNs = []
  //Iterate through each folderr
  let fileID = files.next().getId();
  Logger.log(fileID);
  
  const doc = getTextFromPDF(fileID);
  const studentIDs = extractStudentIDs(doc.text.toString()); 
  importToSpreadsheet(studentIDs);
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
  // Convert the pdf to a Google Doc with ocr.
  var file = Drive.Files.insert(resource, blob, options);

  // Get the texts from the newly created text.
  var doc = DocumentApp.openById(file.id);
  var text = doc.getBody().getText();
  var title = doc.getName();
  
  // Deleted the document once the text has been stored.
  Drive.Files.remove(doc.getId());
  
  return {
    name:title,
    text:text
  };
}

function extractStudentIDs(text){
  const regexp = /\d{13}[A-Z]{1}/g;
    let array = [...text.matchAll(regexp)];
    const spliceArray = array.splice(0, Math.ceil(array.length /2));

    Logger.log(spliceArray);
    let companyArray = [];
    for(let y=0; y < spliceArray.length; y++){
    const posistion = text.indexOf(spliceArray[y])+19;
    let model = text.substring(posistion);
    for(let x = 0; x < 2; x++){
      const kkkgg = model.replace(/,/, '');
      model = kkkgg;
    }
    const split = model.slice(0, model.indexOf(",")-(model.length));
    companyArray.push(split);;
    }
  return {code:spliceArray.toString(), company:companyArray.toString()};
};

function importToSpreadsheet(code){
  const sheet = SpreadsheetApp.openById(SS).getSheetByName(SHEET);
  const range = sheet.appendRow([code.code,code.company]);
  for(const x = 0; x <= code.length; x++){
    Logger.log(x);
  }
}
