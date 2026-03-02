function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Import Data")
    .addItem("Import Checkboxes Only", "importCheckboxes")
    .addItem("Import Quantities Only", "importQuantities")
    .addItem("Import Checkboxes and Quantities", "importBoth")
    .addToUi();
}

function importCheckboxes() {
  showFilePicker("checkboxes");
}

function importQuantities() {
  showFilePicker("quantities");
}

function importBoth() {
  showFilePicker("both");
}

function showFilePicker(mode) {
  const template = HtmlService.createTemplateFromFile("FilePicker");
  template.mode = mode;

  const html = template
    .evaluate()
    .setWidth(500)
    .setHeight(400)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);

  SpreadsheetApp.getUi().showModalDialog(html, "Select Old Spreadsheet");
}

function getSpreadsheetFiles() {
  const activeId = SpreadsheetApp.getActive().getId();
  const files = [];

  const iterator = DriveApp.searchFiles(
    "mimeType='application/vnd.google-apps.spreadsheet' " +
    "and trashed=false " +
    "and 'me' in owners"
  );

  while (iterator.hasNext()) {
    const file = iterator.next();

    if (file.getId() === activeId) continue;

    files.push({
      name: file.getName(),
      id: file.getId()
    });
  }

  return files;
}

function openProgressDialog(fileId, mode) {
  const template = HtmlService.createTemplateFromFile("Progress");
  template.fileId = fileId;
  template.mode = mode;

  const html = template
    .evaluate()
    .setWidth(400)
    .setHeight(150)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);

  SpreadsheetApp.getUi().showModalDialog(html, "Importing...");
}

function showProgressComplete() {
  SpreadsheetApp.getUi().alert("Import complete.");
}

function showProgressError(message) {
  SpreadsheetApp.getUi().alert("Error: " + message);
}

function importData(fileId, mode) {  
  const oldSpreadsheet = SpreadsheetApp.openById(fileId);
  const newSpreadsheet = SpreadsheetApp.getActive();

  const excludePages = ["Summary", "Card Collection", "Lookups"];
  let copiedSheets = [];

  oldSpreadsheet.getSheets().forEach(oldPage => {
    const pageName = oldPage.getName();
    if (excludePages.includes(pageName)) return;

    const newPage = newSpreadsheet.getSheetByName(pageName);
    if (!newPage) return;

    const lastRow = oldPage.getLastRow();
    if (lastRow <= 2) return;

    let values;

    let copied = false;

    Logger.log("Mode: " + mode);

    if (mode === "checkboxes" || mode === "both") {
      values = oldPage.getRange(3, 1, lastRow - 2, 1).getValues();
      newPage.getRange(3, 1, values.length, 1).setValues(values);

      copied = true;
    }
    
    if (mode === "quantities" || mode === "both") {
      const oldCol = findQuantityColumn(oldPage)
      const newCol = findQuantityColumn(newPage)

      if (!oldCol || !newCol) return;

      const values = oldPage
        .getRange(3, oldCol, lastRow - 2, 1)
        .getValues();

      newPage
        .getRange(3, newCol, values.length, 1)
        .setValues(values);

      copied = true;
    }

    if (copied) {
      copiedSheets.push(pageName);
    }
  });

  return copiedSheets;
}

function findQuantityColumn(sheet) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0];

  for (let col = 0; col < headers.length; col++) {
    if (headers[col] === "Obtained") {
      return col + 1;
    }
  }

  return null;
}

function showCompletionDialog(copiedSheets) {
  const template = HtmlService.createTemplateFromFile("Complete");
  template.copiedSheets = copiedSheets;

  const html = template
    .evaluate()
    .setWidth(400)
    .setHeight(300)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);

  SpreadsheetApp.getUi().showModalDialog(html, "Import Complete");
}
