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

function getHeaderMap(sheet) {
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(2, 1, 1, lastColumn).getValues()[0];
  const map = {};

  headers.forEach((h, i) => {
    if (!map[h]) map[h] = [];
    map[h].push(i + 1);
  });

  return map;
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

  const labels = {
    checkboxes: "Importing Checkboxes Only...",
    quantities: "Importing Quantities Only...",
    both: "Importing Checkboxes and Quantities..."
  };

  const title = labels[mode] || "Importing...";

  const html = template
    .evaluate()
    .setWidth(400)
    .setHeight(150)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);

  SpreadsheetApp.getUi().showModalDialog(html, title);
}

function showProgressComplete() {
  SpreadsheetApp.getUi().alert("Import complete.");
}

function showProgressError(message) {
  SpreadsheetApp.getUi().alert("Error: " + message);
}

function setProgress(sheetName) {
  PropertiesService.getScriptProperties().setProperty("IMPORT_PROGRESS", sheetName);
}

function getProgress() {
  return PropertiesService.getScriptProperties().getProperty("IMPORT_PROGRESS");
}


function importData(fileId, mode, startIndex, copiedSheets) {
  const startTime = Date.now();
  const maxRunTime = 60000;
  const oldSpreadsheet = SpreadsheetApp.openById(fileId);
  const newSpreadsheet = SpreadsheetApp.getActive();

  const needsQuantities = mode === "quantities" || mode === "both";
  const needsCheckboxes = mode === "checkboxes" || mode === "both";

  const excludePages = new Set(["Summary", "Card Collection", "Lookups"]);
  const oldPages = oldSpreadsheet.getSheets();
  const newSheets = {};

  newSpreadsheet.getSheets().forEach(sheet => {
    newSheets[sheet.getName()] = sheet;
  });

  startIndex = startIndex || 0;
  copiedSheets = copiedSheets || [];

  if (startIndex === 0) {
    PropertiesService.getScriptProperties().deleteProperty("IMPORT_PROGRESS");
  }

  for (let i = startIndex; i < oldPages.length; i++) {
    if (Date.now() - startTime >= maxRunTime) {
      return {
        done: false,
        nextIndex: i,
        copiedSheets: copiedSheets
      };
    }

    const oldPage = oldPages[i];
    const pageName = oldPage.getName();

    if (excludePages.has(pageName)) continue;

    setProgress(pageName);

    const newPage = newSheets[pageName];

    if (!newPage) continue;

    const lastRow = oldPage.getLastRow();

    if (lastRow <= 2) continue;

    let copied = false;
    let oldCol = null;
    let newCol = null;

    if (needsQuantities) {
      const oldMap = getHeaderMap(oldPage);
      const newMap = getHeaderMap(newPage);

      oldCol = oldMap["Obtained"]?.[0];
      newCol = newMap["Obtained"]?.[0];
    }

    if (needsCheckboxes) {
      const values = oldPage.getRange(3, 1, lastRow - 2, 1).getValues();
      newPage.getRange(3, 1, values.length, 1).setValues(values);
      copied = true;
    }

    if (needsQuantities && oldCol && newCol) {
      const values = oldPage.getRange(3, oldCol, lastRow - 2, 1).getValues();
      newPage.getRange(3, newCol, values.length, 1).setValues(values);
      copied = true;
    }

    if (copied) copiedSheets.push(pageName);
  }

  if (copiedSheets.length > 0) {
    showCompletionDialog(copiedSheets);
  }

  return {
    done: true,
    copiedSheets: copiedSheets
  };
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
