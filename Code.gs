function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Dynamic Form')
    .addItem('📋 Dynamic Entry Form', 'showDynamicForm')
    .addToUi();
}

function showDynamicForm() {
  const html = HtmlService.createHtmlOutputFromFile('DynamicForm')
    .setTitle('Dynamic Data Entry Form');
  SpreadsheetApp.getUi().showSidebar(html);
  createDropdownSheet()
}

function createDropdownSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Dropdowns");

  // If the sheet already exists, do nothing
  if (sheet) return;

  // Otherwise, create it and set headers
  const newSheet = ss.insertSheet("Dropdowns");
  newSheet.getRange("A1").setValue("Dropdown");
  newSheet.getRange("B1").setValue("Options");
}



function getSheetInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const validations = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getDataValidations()[0];
  const dropdownsSheet = ss.getSheetByName('Dropdowns');
  const dropdownOptions = dropdownsSheet ? getDropdownOptions(dropdownsSheet) : {};

  return headers.map((header, index) => {
    const validation = validations[index];
    let type = 'text';
    let options = [];

    // Check data validation rules
    if (validation && validation.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      type = 'select';
      options = validation.getCriteriaValues();
    }
    // Override with Dropdowns sheet if available
    if (dropdownOptions[header]) {
      type = 'select';
      options = dropdownOptions[header];
    }
    // Set ID as read-only number
    if (header === 'ID') {
      type = 'number';
    }
    return {
      name: header,
      type: type,
      options: options,
      required: header !== 'ID', // ID is auto-generated, others required
      columnIndex: index + 1
    };
  });
}

function getDropdownOptions(dropdownsSheet) {
  const data = dropdownsSheet.getDataRange().getValues();
  const options = {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  for (let i = 1; i < data.length; i++) {
    const key = data[i][0];
    const value = data[i][1];

    if (key && value) {
      if (value.includes('!')) {
        // If the value is a range reference like "contacts!A:A"
        const [sheetName, colRange] = value.split('!');
        const sourceSheet = ss.getSheetByName(sheetName);
        if (sourceSheet) {
          const range = sourceSheet.getRange(colRange);
          const values = range.getValues().flat().filter(v => v !== '');
          options[key] = [...new Set(values)]; // remove duplicates
        } else {
          Logger.log(`Sheet ${sheetName} not found.`);
        }
      } else {
        // Comma-separated inline options
        options[key] = value.split(',').map(opt => opt.trim());
      }
    }
  }
  return options;
}


function getVisibleRecords() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const filter = sheet.getFilter();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const records = [];

  if (filter) {
    for (let i = 1; i < data.length; i++) {
      if (!sheet.isRowHiddenByFilter(i + 1)) {
        records.push(data[i]);
      }
    }
  } else {
    records.push(...data.slice(1));
  }
  return records.map(row => headers.reduce((obj, header, i) => {
    obj[header] = row[i];
    return obj;
  }, {}));
}

function addRecord(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const lastId = sheet.getLastRow() > 1 ? Number(sheet.getRange(sheet.getLastRow(), 1).getValue()) || 0 : 0;
  const newId = lastId + 1;
  const row = headers.map(header => header === 'ID' ? newId : formData[header] || '');
  sheet.appendRow(row);
  return { status: 'success', id: newId };
}

function updateRecord(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == formData.ID) {
      const row = headers.map(header => formData[header] || '');
      sheet.getRange(i + 1, 1, 1, headers.length).setValues([row]);
      return { status: 'success' };
    }
  }
  return { status: 'error', message: 'Record not found' };
}

function deleteRecord(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 1);
      return { status: 'success' };
    }
  }
  return { status: 'error', message: 'Record not found' };
}

function searchRecord(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id && !sheet.isRowHiddenByFilter(i + 1)) {
      return headers.reduce((obj, header, index) => {
        obj[header] = data[i][index];
        return obj;
      }, {});
    }
  }
  return null;
}
