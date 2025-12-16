function getData(sheetKey, identifier) {
  try {
    const ss = connectToSheet(sheetKey);
    let sheet = ss.getSheetByName(identifier);
    if (!sheet) sheet = ss.getSheetByName(identifier.trim());

    // Fallback to Named Range
    if (!sheet) {
      const range = ss.getRangeByName(identifier);
      if(range) return processValues(range.getValues());
      throw new Error(`Tab "${identifier}" not found in ${ss.getName()}`);
    }
    // Get Data from Tab
    const range = sheet.getDataRange();
    return processValues(range.getValues());
  } catch (e) {
    Logger.log(e); throw new Error(e.message);
  }
}

function processValues(values) {
    if (values.length <= 1) return [];
    const headers = values[0];
    const data = [];
    for (let i = 1; i < values.length; i++) {
      let row = values[i];
      if (row.every(cell => cell === "")) continue;
      let obj = {};
      obj['_row'] = i + 1;
      for (let j = 0; j < headers.length; j++) {
        const headerName = headers[j].toString().trim().replace(/\s+/g, '_');
        let value = row[j];
        if (value instanceof Date) value = Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
        if (value === null || value === undefined) value = "";
        obj[headerName] = value;
      }
      data.push(obj);
    }
    return data;
}

function addData(sheetKey, tabName, dataObj) {
  try {
    const ss = connectToSheet(sheetKey);
    const sheet = ss.getSheetByName(tabName) || ss.getSheetByName(tabName.trim());
    if (!sheet) throw new Error(`Tab "${tabName}" missing.`);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const newRow = [];
    headers.forEach(header => {
      let key = header.toString().trim().replace(/\s+/g, '_');
      let val = dataObj[key];
      if (val === undefined) val = "";
      if (val instanceof Date) val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
      newRow.push(val);
    });
    sheet.appendRow(newRow);
    return "Success";
  } catch (e) { throw new Error("Save Failed: " + e.message); }
}

function updateData(sheetKey, tabName, rowNumber, dataObj) {
    try {
        const ss = connectToSheet(sheetKey);
        const sheet = ss.getSheetByName(tabName) || ss.getSheetByName(tabName.trim());
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const newRow = [];
        headers.forEach(header => {
            let key = header.toString().trim().replace(/\s+/g, '_');
            let val = dataObj[key];
            if (val === undefined) val = "";
            if (val instanceof Date) val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
            newRow.push(val);
        });
        sheet.getRange(rowNumber, 1, 1, newRow.length).setValues([newRow]);
        return "Success";
    } catch(e) { throw new Error("Update Failed: " + e.message); }
}

function deleteData(sheetKey, tabName, rowNumber) {
    try {
        const ss = connectToSheet(sheetKey);
        const sheet = ss.getSheetByName(tabName) || ss.getSheetByName(tabName.trim());
        sheet.deleteRow(Number(rowNumber));
        return "Success";
    } catch(e) { throw new Error("Delete Failed: " + e.message); }
}

function connectToSheet(key) {
  const IDS = {
     HQ: "1C9i-zOXJ5piESoxjyPsru6lIKQepSRDPzNjgjo3UxuA",
     LIMB_CENTRE: "1Ap057Z2Jh5-xuaDsEXQApZz87hgVDtPrTXyBmc41Jxw",
     DISPENSARY: "18Pegr9qRDaCE5PEnS-1KQLxrIbSaziRLt-Dh1H8VTsQ",
     RATION: "1sIZ6IXMKd74v2G0iEJKIK1jqLVmC4_8RqLIDifK67r0",
     CAMPS: "1sX74zsgFqb7s5bQGWnlFkcgIgcNeB2Tv57a5mm2haM8"
  };
  if (!IDS[key]) throw new Error("Invalid Key: " + key);
  return SpreadsheetApp.openById(IDS[key]);
}

function createPatientFolder(parentFolderId, folderName) {
  try {
    const parent = DriveApp.getFolderById(parentFolderId);
    const newFolder = parent.createFolder(folderName);
    return { id: newFolder.getId(), url: newFolder.getUrl() };
  } catch (e) { return { id: "", url: "#" }; }
}

function uploadFileToFolder(base64Data, filename, folderId) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const split = base64Data.split(',');
    const type = split[0].split(':')[1].split(';')[0];
    const blob = Utilities.newBlob(Utilities.base64Decode(split[1]), type, filename);
    return folder.createFile(blob).getUrl();
  } catch (e) { return "Error: " + e.toString(); }
}