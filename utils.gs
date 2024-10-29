/**
 * Creates a copy of an Excel file as a Google Spreadsheet.
 * Returns a Drive File object of the new Google Spreadsheet.
 * 
 * @param {DriveApp.File} excelFile The excel file to convert.
 * @param {DriveApp.Folder} destinationFolder Optional. Determines location of converted file. Defaults to parent folder of excelFile.
 * @param {Boolean} deleteExcelAfterConversion Optional. Determines deletion of the excel file. Defaults to true.
 * @return {DriveApp.File} The converted Google Spreadsheet.
 */
function convertExcelToGoogleSpreadsheet(excelFile,destinationFolder,deleteExcelAfterConversion) {

  // retrieve the parent folder if destination not specified
  if (!destinationFolder) {
    const parentFolders = excelFile.getParents();
    destinationFolder = parentFolders.next();
  }

  // retrieve excel file data
  const blob = excelFile.getBlob();

  // define the name and file type of the new google file to be created
  const resource = {
    title: excelFile.getName(),
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [{ id: destinationFolder.getId() }]
  };

  // create the new google sheet
  const convertedSpreadsheet = Drive.Files.insert(resource, blob, {
    convert: true
  });

  // move the google sheet to destination
  const cSpreadsheetId = convertedSpreadsheet.getId();
  const cSpreadsheet = DriveApp.getFileById(cSpreadsheetId);
  cSpreadsheet.moveTo(destinationFolder);

  // deletes the excel file if needed
  const deleteExcelFile = deleteExcelAfterConversion ?? true;
  excelFile.setTrashed(deleteExcelFile);

  return convertedSpreadsheet;
}

/** Returns the value of the cell within the specified spreadsheet and sheet.
 * 
 * @param {DriveApp.File} targetSpreadsheet The target spreadsheet.
 * @param {String} targetSheet The target sheet within the spreadsheet.
 * @param {String} targetCell The target cell within the sheet.
 * 
 * @return {String} The contents of the target cell.
 */
function findCellValueOfSheet(targetSpreadsheet,targetSheet,targetCell) {
  
  const activeSpreadsheet = SpreadsheetApp.openById(targetSpreadsheet.getId());
  const activeSheet = activeSpreadsheet.getSheetByName(targetSheet);
  const targetValue = String(activeSheet.getRange(targetCell).getValue());

  return targetValue;

}

/**
 * Determines the last updated file from a list of files.
 * Returns the name, id, or file object of the last updated file.
 * 
 * @param {DriveApp.FileIterator} filesToEvaluate The list of files to check last modified date.
 * @pararm {String} returnValue Optional. Determines the identifing value returned when the last updated file is found. Can equal name, id, or file. Defaults to name.
 * 
 * @return {String|DriveApp.File} The name, id, or file object of the last updated file.
 */
function findLastUpdatedFile(filesToEvaluate,returnValue) {

  // loop through files to evaluate to find the latest updated date
  var latestUpdateDate = 0;
  let latestUpdatedFile;

  while (filesToEvaluate.hasNext()) {
    var fileToEvaluate = filesToEvaluate.next();
    var fileUpdatedDate = fileToEvaluate.getLastUpdated().valueOf();
    if (fileUpdatedDate > latestUpdateDate) {
      latestUpdateDate = fileUpdatedDate;
      
      if (returnValue == "id") {
        latestUpdatedFile = fileToEvaluate.getId();
      } else if (returnValue == "file") {
        latestUpdatedFile = fileToEvaluate;
      } else {
        latestUpdatedFile = fileToEvaluate.getName();
      }
    }
  }

  return latestUpdatedFile;

}

/**
 * Looks for the specified folder. Either returns the existing folder or a new folder with the same name after creation.
 * If there are multiple folders with the same name in the specified location, the last matching folder found will be returned.
 * 
 * @param {String} folderName Name of the folder to find.
 * @param {DriveApp.Folder} folderLocation Optional. The parent folder to search. Defaults to the user's drive.
 * 
 * @return {DriveApp.Folder} The found or created folder.
 */
function findOrCreateFolderByName(folderName,folderLocation){

  // set the parent folder to the user's drive if not specified
  const parentFolder = folderLocation ?? DriveApp.getRootFolder();

  // search through all child folders with the specified name
  const matchingFolders = parentFolder.getFoldersByName(folderName);
  var folderCount = 0;
  let matchingFolder;

  while (matchingFolders.hasNext()) {
    matchingFolder = matchingFolders.next();
    folderCount += 1;
  }

  // if there are no matching folders, create one and return it, otherwise return the last matching folder found
  if (folderCount == 0) {
    const newFolder = parentFolder.createFolder(folderName);
    return newFolder;
  } else {
    return matchingFolder; 
  }
}

/**
 * Formats a sheet in a Google Spreadsheet as plain text.
 * No return value.
 * 
 * @param {DriveApp.File} targetSpreadsheet The target file. Must be a Google Spreadsheet file.
 * @param {String} targetSheetName Optional. The name of the sheet to retrieve data from. Defaults to the first sheet.
 * @param {Number} targetSheetIndex Optional. The zero-indexed position of the sheet to retreive the data from. Defaults to the first sheet.
 */
function formatSheetAsPlainText(targetSpreadsheet,targetSheetName,targetSheetIndex){

  // set destination spreadsheet index
  const sheetIndex = targetSheetIndex ?? 0
  let tSheet;

  // isolate the destination sheet
  if (!targetSheetName) {
    tSheet = targetSpreadsheet.getSheets()[sheetIndex];
  } else {
    tSheet = targetSpreadsheet.getSheetByName(targetSheetName);
  }

  tSheet.getDataRange.setNumberFormat('@STRING@')
}

/**
 * Retrieves data from a CSV, XLSX, or Google Spreadsheet file.
 * Returns data as a string.
 * 
 * @param {DriveApp.File} dataFile The target file. Must be a CSV, XLSX, or Google Spreadsheet file.
 * @param {String} dataFileSheetName Optional. The name of the sheet to retrieve data from if file is an XLSX or Google Spreadsheet file. Defaults to the first sheet.
 * @param {Number} dataFileSheetIndex Optional. The zero-indexed position of the sheet to retreive the data from if file is an XLSX or Google Spreadsheet file. Defaults to the first sheet.
 * 
 * @return {String} The file data.
 */
function getDataFromFile(dataFile,dataFileSheetName,dataFileSheetIndex){

  // retrieve data file name
  const dataFileName = dataFile.getName()

  // get data from different file types
  let sourceData;

  if (dataFileName.includes(".csv")) {
    sourceData = Utilities.parseCsv(dataFile.getBlob().getDataAsString());
  } else {
    // set data file sheet index
    const sheetIndex = dataFileSheetIndex ?? 0;
    const fileIsExcel = dataFileName.includes(".xlsx");

    // convert the file if it is an excel file to a readable format
    if (fileIsExcel) {
      dataFile = convertExcelToGoogleSpreadsheet(dataFile, false)
    }
    const dataFileId = dataFile.getId();
    let dataSheet;
    
    // if the data file sheet name wasn't specified, use the index, otherwise use the name
    if (!dataFileSheetName) {
      dataSheet = SpreadsheetApp.openById(dataFileId).getSheets()[sheetIndex];
    } else {
      dataSheet = SpreadsheetApp.openById(dataFileId).getSheetByName(dataFileSheetName);
    }
    
    const sourceRange = dataSheet.getDataRange();
    sourceData = sourceRange.getValues();

    // deletes the converted google spreadsheet if created
    DriveApp.getFileById(dataFileId).setTrashed(fileIsExcel);
  } 

  return sourceData;

}

/**
 * Creates or replaces a sheet in a Google Spreadsheet with data from a file.
 * Currently supports CSV, XLSX, and Google Spreadsheet files.
 * No return value.
 * 
 * @param {DriveApp.File} destinationSpreadsheet The destination spreadsheet.
 * @param {DriveApp.File} dataFile The data file to be copied to the destination spreadsheet. Must be a CSV, XLSX, or Google Spreadsheet file.
 * @param {String} dataFileSheetName Optional. The name of the sheet to retrieve data from if file is an XLSX or Google Spreadsheet file. Defaults to the first sheet.
 * @param {Number} dataFileSheetIndex Optional. The zero-indexed position of the sheet to copy if using an XLSX or Google Spreadsheet file. Defaults to the first sheet.
 * @param {String} targetSheetName Optional. The name of the sheet being created or replaced. Defaults to the name of the data file.
 * @param {Boolean} inferDataTypes Optional. Determines the destination spreadsheet inferring the data types of the new data. Defaults to false.
 * @param {Boolean} deleteDataFileAfterImport Optional. Determines deletion of the data file. Defaults to false.
 */
function upsertNewSheetFromFile(destinationSpreadsheet,dataFile,dataFileSheetName,dataFileSheetIndex,targetSheetName,inferDataTypes,deleteDataFileAfterImport){

  // determines optional parameters
  const sheetIndex = dataFileSheetIndex ?? 0;
  const dataFileName = dataFile.getName().split(".")[0];
  const tSheetName = targetSheetName || dataFileName;
  const inferTypes = inferDataTypes ?? false;
  const deleteDataFile = deleteDataFileAfterImport ?? false;

  // get the source data from file
  let sourceData;
  if (!dataFileSheetName) {
    sourceData = getDataFromFile(dataFile,undefined,sheetIndex);
  } else {
    sourceData = getDataFromFile(dataFile,dataFileSheetName);
  }

  // find or create the target sheet in the destination spreadsheet
  const activeSpreadsheet = SpreadsheetApp.openById(destinationSpreadsheet.getId());
  var activeSheet = activeSpreadsheet.getSheetByName(tSheetName);

  // if the sheet doesn't exist, create it, otherwise clear the data
  if (!activeSheet) {
    activeSheet = activeSpreadsheet.insertSheet(tSheetName);
  } else {
    activeSheet.clear();
  }

  // insert source data
  activeSheet.getRange(1,1, sourceData.length, sourceData[0].length).setValues(sourceData);

  if (!inferTypes) {
    activeSheet.getRange(1,1, sourceData.length, sourceData[0].length).setNumberFormat('@STRING@');
  }

  // deletes the data file if needed
  dataFile.setTrashed(deleteDataFile);

}
