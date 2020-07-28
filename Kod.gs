//#####GLOBALS#####
const FOLDER_ID = "1CnTA5AQo4xzo4TJAWpGtEPVRhh9jdCHl"; //Folder ID of all PDFs
const SS = "1jJNhSkGq3SbqSnRXNRd4EhsYCjDwZ0LYZW6y3lkZVGw";//The spreadsheet ID
const SHEET = "Extracted";//The sheet tab name


/*########################################################
 * Main run file: extracts student IDs from PDFs and their 
 * section from the PDF name from multiple documents.
 *
 * Displays a list of students and sections in a Google Sheet. 
 *
 */
function extractStudentIDsAndSectionToSheets(){
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  //Get all PDF files:
  const folder = DriveApp.getFolderById(FOLDER_ID);
  //const files = folder.getFiles();
  const files = folder.getFilesByType("application/pdf");
  
  let allIDsAndCRNs = []
  //Iterate through each folderr
  while(files.hasNext()){
    let file = files.next();
    let fileID = file.getId();
    
    const doc = getTextFromPDF(fileID);
    const studentIDs = extractStudentIDs(doc.text);
    
    //Add ID to Section name
    const studentIDsWithCRN = studentIDs.map( ID => [ID,doc.name]);
    
    //Optional: Notify user of process. You can delete lines 33 to 38
    if(studentIDs[0] === "No items found") {
      ss.toast("No items found in " + doc.name, "Warning",2);
    }else{
      ss.toast(doc.name + " extracted");
    };    
    
    allIDsAndCRNs = allIDsAndCRNs.concat(studentIDsWithCRN);
    
  }
    importToSpreadsheet(allIDsAndCRNs);
};


/*########################################################
 * Extracts the text from a PDF and stores it in memory.
 * Also extracts the file name.
 *
 * param {string} : fileID : file ID of the PDF that the text will be extracted from.
 *
 * returns {array} : Contains the file name (section) and PDF text.
 *
 */
function getTextFromPDF(fileID) {
  var blob = DriveApp.getFileById(fileID).getBlob()
  var resource = {
    title: blob.getName(),
    mimeType: blob.getContentType()
  };
  var options = {
    ocr: true, 
    ocrLanguage: "en"
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

/*########################################################
 * Use the text extracted from PDF and extracts student id based on value parameters.
 * Also extracts the file name.
 *
 * param {string} : text : text of data from PDF.
 *
 * returns {array} : Of all student IDs found in text.
 *
 */
function extractStudentIDs(text){
  const regexp = /20\d{7}/g;
  try{
    let array = [...text.match(regexp)];
    return array;
  }catch(e){
    //Optional: If you want this info added to your Sheet data. Otherwise delete rows 98-99. 
    let array = ["Not items found"] 
    return array;
  }
};

/*########################################################
 * Takes the culminated list of IDs and sections and inserts them into 
 * a Google Sheet.
 *
 * param {array} : data : 2d array containing a list of ids and their associated sections.
 *
 */
function importToSpreadsheet(data){
  const sheet = SpreadsheetApp.openById(SS).getSheetByName(SHEET);
  
  const range = sheet.getRange(3,1,data.length,2);
  range.setValues(data);
  range.sort([2,1]);
}