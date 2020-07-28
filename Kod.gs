const FOLDER_ID = "13IOecyfRgwkdphJr7lChQrDT2IeiErah";
const SS = "13_euBdyshOplyxobP_v1geVwWg9GMfqO7ddYdmC-NoI";
const SHEET = "Arkusz1";

function extractStudentIDsAndSectionToSheets(){
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFilesByType("application/pdf");
  
  let allIDsAndCRNs = []
  let fileID = files.next().getId();
  Logger.log(fileID);
  
  const doc = getTextFromPDF(fileID);
  const studentIDs = extractStudentIDs(doc.text.toString()); 
  importToSpreadsheet(studentIDs);
};

function getTextFromPDF(fileID) {
  const blob = DriveApp.getFileById(fileID).getBlob();
  const resource = {
    title: blob.getName(),
    mimeType: blob.getContentType()
  };
  const options = {
    ocr: true
  };
  const file = Drive.Files.insert(resource, blob, options);
  
  const doc = DocumentApp.openById(file.id);
  const text = doc.getBody().getText();
  const title = doc.getName();
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
    companyArray.push(split);
  }
  return {code:spliceArray, company:companyArray};
};

function importToSpreadsheet(code){
  const sheet = SpreadsheetApp.openById(SS).getSheetByName(SHEET);
  for(let x = 0; x <= code.code.length-1; x++){
    Logger.log(code.code[x]);
    sheet.appendRow([code.code[x][0],code.company[x]]);
  }
}
