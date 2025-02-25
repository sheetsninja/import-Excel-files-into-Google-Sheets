const FOLDER_ID = '1PvIhX3JHxBq0i8m-7wQaoJIRCbqphZAp';
const ARCHIVE_ID = '13U9WhI1Rmn-rBALxgdKThcNXteWGxn1r';
const TAB_NAME = "TAB_NAME"; // tab name in Sheets to append the data

function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Excel Script').addItem("Get Excel Data", "importExcel").addToUi();
}

function importExcel() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(TAB_NAME);
  let folder = DriveApp.getFolderById(FOLDER_ID);
  let archiveFolder = DriveApp.getFolderById(ARCHIVE_ID)
  let files = folder.getFilesByType(MimeType.MICROSOFT_EXCEL);

  while (files.hasNext()) {
    let excelFile = files.next();
    ss.toast(`Processing ${excelFile.getName()}`,"STATUS",-1);
    let excelFileId = excelFile.getId();
    let blob = excelFile.getBlob();
    let resource = {
      name: excelFile.getName().replace(/.xlsx?/, ""),
      mimeType: MimeType.GOOGLE_SHEETS
    };

    let newTempGS = Drive.Files.create(resource, blob);

    let ssNew = SpreadsheetApp.openById(newTempGS.id);
    let sheetNew = ssNew.getSheets()[0];
    let data = sheetNew.getDataRange().getValues();
    data.shift();
    sheet.getRange(sheet.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);

    DriveApp.getFileById(newTempGS.id).setTrashed(true);
    excelFile.moveTo(archiveFolder);
  }
  sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).sort({ column: 1, ascending: true });
  ss.toast("Completed successfully!","SUCCESS",5);
}
