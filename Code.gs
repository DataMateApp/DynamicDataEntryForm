// Code.gs

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Form Tool')
    .addItem('📋 Dynamic Data Entry Form', 'showDynamicForm')
    .addToUi();
}

function showDynamicForm() {
  const htmlContent = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: Arial, sans-serif; padding: 20px; }
        label { display: block; margin: 10px 0 5px; }
        input, select { width: 100%; padding: 8px; margin-bottom: 10px; }
        button { padding: 10px; margin: 5px; }
        #message { color: green; margin-top: 10px; }
        .error { color: red; }
        #spinner { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); }
        #spinner div { position: absolute; top: 50%; left: 50%; transform: translate(-50%,-50%); color: white; }
      </style>
    </head>
    <body>
      <form id="dynamicForm">
        <div id="formFields"></div>
        <button type="button" onclick="saveRecord()">Save</button>
        <button type="button" onclick="clearForm()">New</button>
        <button type="button" onclick="navigate('prev')">Previous</button>
        <button type="button" onclick="navigate('next')">Next</button>
      </form>
      <div id="message"></div>
      <div id="spinner"><div>Loading...</div></div>

      <script>
        let headers = [];
        let records = [];
        let currentIndex = -1;

        google.script.run.withSuccessHandler(populateForm).getSheetInfo();
        google.script.run.withSuccessHandler(loadRecords).getVisibleRecords();

        function populateForm(headerData) {
          headers = headerData;
          const formFields = document.getElementById('formFields');
          formFields.innerHTML = headers.map(header => {
            if (header.name === 'ID') {
              return \`<label for="\${header.name}">\${header.name}</label>
                      <input type="number" id="\${header.name}" readonly>\`;
            } else if (header.type === 'select') {
              return \`<label for="\${header.name}">\${header.name} \${header.required ? '*' : ''}</label>
                      <select id="\${header.name}" \${header.required ? 'required' : ''}>
                        <option value="">Select \${header.name}</option>
                        \${header.options.map(opt => \`<option value="\${opt}">\${opt}</option>\`).join('')}
                      </select>\`;
            } else {
              return \`<label for="\${header.name}">\${header.name} \${header.required ? '*' : ''}</label>
                      <input type="\${header.type}" id="\${header.name}" \${header.required ? 'required' : ''}>\`;
            }
          }).join('');
        }

        function loadRecords(data) {
          records = data;
          if (records.length > 0) {
            currentIndex = 0;
            displayRecord();
          }
        }

        function displayRecord() {
  isNewRecord = false; // edit mode
  if (currentIndex >= 0 && currentIndex < records.length) {
    headers.forEach(header => {
      const field = document.getElementById(header.name);
      const value = records[currentIndex][header.name] || '';

      if (field.tagName === 'SELECT') {
        // Check if value exists in dropdown
        let exists = Array.from(field.options).some(opt => opt.value === value);
        if (!exists && value) {
          // Add it temporarily so it can be selected
          const opt = document.createElement('option');
          opt.value = value;
          opt.textContent = value;
          field.appendChild(opt);
        }
        field.value = value;
      } else {
        field.value = value;
      }
    });
  }
}

        function saveRecord() {
  document.getElementById('spinner').style.display = 'block';
  const formData = {};
  headers.forEach(header => {
    formData[header.name] = document.getElementById(header.name).value.trim();
  });

  if (headers.some(header => header.required && !formData[header.name])) {
    showMessage('Please fill all required fields.', 'error');
    document.getElementById('spinner').style.display = 'none';
    return;
  }

  if (isNewRecord) {
    google.script.run.withSuccessHandler(result => onSave(result, null))
      .withFailureHandler(onError)
      .addRecord(formData);
  } else {
    // Send row number to server
    formData._rowNumber = currentIndex + 2; // +2 because currentIndex is 0-based and row 1 is headers
    google.script.run.withSuccessHandler(result => onSave(result, formData._rowNumber))
      .withFailureHandler(onError)
      .updateRecord(formData);
  }
}


        let isNewRecord = false; // track mode: add or edit

function clearForm(keepID = false) {
  isNewRecord = true; // new mode
  document.getElementById('dynamicForm').reset();
  headers.forEach(header => {
    if (keepID && header.name === 'ID') return;
    document.getElementById(header.name).value = '';
  });
  showMessage('Ready for new record.', '');
}

        function navigate(direction) {
          if (records.length === 0) return;
          if (direction === 'prev' && currentIndex > 0) {
            currentIndex--;
          } else if (direction === 'next' && currentIndex < records.length - 1) {
            currentIndex++;
          }
          displayRecord();
        }

        function searchRecord() {
          const id = prompt('Enter ID to search:');
          if (id) {
            document.getElementById('spinner').style.display = 'block';
            google.script.run.withSuccessHandler(record => {
              document.getElementById('spinner').style.display = 'none';
              if (record) {
                headers.forEach(header => {
                  document.getElementById(header.name).value = record[header.name] || '';
                });
                showMessage('Record found.', '');
              } else {
                showMessage('Record not found or filtered out.', 'error');
              }
            }).withFailureHandler(err => {
              document.getElementById('spinner').style.display = 'none';
              onError(err);
            }).searchRecord(id);
          }
        }

        function onSave(result, existingID) {
  document.getElementById('spinner').style.display = 'none';

  if (result.status === 'success') {
    showMessage('Record saved successfully.', '');
    google.script.run.withSuccessHandler(data => {
      records = data;

      if (existingID) {
        const idx = records.findIndex(r => String(r.ID) === String(existingID));
        if (idx >= 0) {
          currentIndex = idx;
          displayRecord();
        }
      } else {
        currentIndex = records.length - 1;
        displayRecord();
      }
    }).getVisibleRecords();
    isNewRecord = false;
  } else {
    showMessage(result.message || 'Error saving record.', 'error');
  }
}

        function onDelete(result) {
          document.getElementById('spinner').style.display = 'none';
          showMessage('Record deleted successfully.', '');
          google.script.run.withSuccessHandler(loadRecords).getVisibleRecords();
          clearForm();
        }

        function onError(error) {
          document.getElementById('spinner').style.display = 'none';
          showMessage('Error: ' + error.message, 'error');
        }

        function showMessage(message, className) {
          const msgDiv = document.getElementById('message');
          msgDiv.textContent = message;
          msgDiv.className = className;
          setTimeout(() => msgDiv.textContent = '', 3000);
        }
      </script>
    </body>
    </html>
  `;
  
  const html = HtmlService.createHtmlOutput(htmlContent)
    .setTitle('Dynamic Data Entry Form');
  SpreadsheetApp.getUi().showSidebar(html);
  createDropdownSheet();
}

function createDropdownSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Dropdowns");

  if (sheet) return;

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

    if (validation && validation.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      type = 'select';
      options = validation.getCriteriaValues();
    }
    if (dropdownOptions[header]) {
      type = 'select';
      options = dropdownOptions[header];
    }
    if (header === 'ID') {
      type = 'number';
    }
    return {
      name: header,
      type: type,
      options: options,
      required: header !== 'ID',
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
        const [sheetName, colRange] = value.split('!');
        const sourceSheet = ss.getSheetByName(sheetName);
        if (sourceSheet) {
          const range = sourceSheet.getRange(colRange);
          const values = range.getValues().flat().filter(v => v !== '');
          options[key] = [...new Set(values)];
        } else {
          Logger.log(`Sheet ${sheetName} not found.`);
        }
      } else {
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

  if (!formData._rowNumber || formData._rowNumber <= 1) {
    return { status: 'error', message: 'Invalid row number' };
  }

  // Get the existing row
  const existingRow = sheet.getRange(formData._rowNumber, 1, 1, headers.length).getValues()[0];

  // Merge changes (keep original if form left it blank)
  const updatedRow = headers.map((header, idx) => {
    return formData[header] !== '' && formData[header] !== undefined
      ? formData[header]
      : existingRow[idx];
  });

  // Write updated values back
  sheet.getRange(formData._rowNumber, 1, 1, headers.length).setValues([updatedRow]);

  return { status: 'success', row: formData._rowNumber };
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
