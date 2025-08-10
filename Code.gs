// Code.gs

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Form')
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
        let isNewRecord = false;

        // Load headers and records on sidebar open
        google.script.run.withSuccessHandler(populateForm).getSheetInfo();
        google.script.run.withSuccessHandler(loadRecords).getVisibleRecords();

        // Build form fields dynamically
        function populateForm(headerData) {
          headers = headerData;
          const formFields = document.getElementById('formFields');
          formFields.innerHTML = headers.map(header => {
            if (header.name === 'ID') {
              return \`<label for="\${header.name}">\${header.name}</label>
                      <input type="number" id="\${header.name}" readonly>\`;
            } else if (header.type === 'select') {
              return \`<label for="\${header.name}">\${header.name} \${header.required ? '*' : ''}</label>
                      <select id="\${header.name}" \${header.required ? 'required' : ''} onchange="onDropdownChange()">
                        <option value="">Select \${header.name}</option>
                        \${header.options.map(opt => \`<option value="\${opt}">\${opt}</option>\`).join('')}
                      </select>\`;
            } else {
              return \`<label for="\${header.name}">\${header.name} \${header.required ? '*' : ''}</label>
                      <input type="\${header.type}" id="\${header.name}" \${header.required ? 'required' : ''}>\`;
            }
          }).join('');
        }

        // Load all records from sheet
        function loadRecords(data) {
          records = data;
          if (records.length > 0) {
            currentIndex = 0;
            displayRecord();
          } else {
            clearForm();
          }
        }

        // Show the current record in the form
        function displayRecord() {
          isNewRecord = false; // we're editing existing record
          if (currentIndex >= 0 && currentIndex < records.length) {
            headers.forEach(header => {
              const field = document.getElementById(header.name);
              const value = records[currentIndex][header.name] || '';

              if (field.tagName === 'SELECT') {
                // Ensure dropdown includes current value (even if not in options)
                let exists = Array.from(field.options).some(opt => opt.value === value);
                if (!exists && value) {
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

        // Save the form data (add or update)
        function saveRecord() {
          document.getElementById('spinner').style.display = 'block';
          const formData = {};
          headers.forEach(header => {
            formData[header.name] = document.getElementById(header.name).value.trim();
          });

          // Validate required fields
          for (const header of headers) {
            if (header.required && !formData[header.name]) {
              showMessage('Please fill all required fields.', 'error');
              document.getElementById('spinner').style.display = 'none';
              return;
            }
          }

          if (isNewRecord) {
            google.script.run
              .withSuccessHandler(result => onSave(result, null))
              .withFailureHandler(onError)
              .addRecord(formData);
          } else {
            formData._rowNumber = currentIndex + 2; // Sheet row (header is row 1)
            google.script.run
              .withSuccessHandler(result => onSave(result, formData._rowNumber))
              .withFailureHandler(onError)
              .updateRecord(formData);
          }
        }

        // Clear form for new record entry
        function clearForm() {
          isNewRecord = true;
          document.getElementById('dynamicForm').reset();
          headers.forEach(header => {
            if (header.name === 'ID') return;
            document.getElementById(header.name).value = '';
          });
          showMessage('Ready for new record.', '');
        }

        // Navigate records prev/next
        function navigate(direction) {
          if (records.length === 0) return;
          if (direction === 'prev' && currentIndex > 0) {
            currentIndex--;
          } else if (direction === 'next' && currentIndex < records.length - 1) {
            currentIndex++;
          }
          displayRecord();
        }

        // Reload the current record from the sheet after dropdown change to get updated formulas
        function onDropdownChange() {
          if (isNewRecord) return; // no reload for new record, only existing

          const currentID = document.getElementById('ID').value;
          if (!currentID) return;

          document.getElementById('spinner').style.display = 'block';
          google.script.run
            .withSuccessHandler(record => {
              if (record) {
                headers.forEach(header => {
                  const field = document.getElementById(header.name);
                  const val = record[header.name] || '';

                  if (field.tagName === 'SELECT') {
                    // Add option if missing
                    let exists = Array.from(field.options).some(opt => opt.value === val);
                    if (!exists && val) {
                      const opt = document.createElement('option');
                      opt.value = val;
                      opt.textContent = val;
                      field.appendChild(opt);
                    }
                    field.value = val;
                  } else {
                    field.value = val;
                  }
                });
                showMessage('Record refreshed with formula updates.', '');
              } else {
                showMessage('Record not found on reload.', 'error');
              }
              document.getElementById('spinner').style.display = 'none';
            })
            .withFailureHandler(err => {
              showMessage('Error refreshing record: ' + err.message, 'error');
              document.getElementById('spinner').style.display = 'none';
            })
            .getRecordById(currentID);
        }

        // After save handler: reload records and display latest saved record with fresh formulas
        function onSave(result, existingRow) {
          document.getElementById('spinner').style.display = 'none';

          if (result.status === 'success') {
            showMessage('Record saved successfully.', '');
            // Reload all visible records
            google.script.run.withSuccessHandler(data => {
              records = data;

              if (existingRow) {
                // Find index of updated record by row number
                // We do not have row number in records, so find by ID
                const updatedID = document.getElementById('ID').value;
                const idx = records.findIndex(r => String(r.ID) === String(updatedID));
                if (idx >= 0) {
                  currentIndex = idx;
                  displayRecord();
                } else {
                  // fallback: show last record
                  currentIndex = records.length - 1;
                  displayRecord();
                }
              } else {
                // For new record, show last record added
                currentIndex = records.length - 1;
                displayRecord();
              }
            }).getVisibleRecords();
            isNewRecord = false;
          } else {
            showMessage(result.message || 'Error saving record.', 'error');
          }
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

// Protect all formula cells on active sheet
function protectAllFormulaCells() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getDataRange();
  const formulas = range.getFormulas();

  for (let r = 0; r < formulas.length; r++) {
    for (let c = 0; c < formulas[r].length; c++) {
      if (formulas[r][c]) {
        const cell = sheet.getRange(r + 1, c + 1);
        const protection = cell.protect();
        protection.setDescription('Formula cell - do not edit');
        protection.removeEditors(protection.getEditors());
      }
    }
  }
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'All formula cells have been protected.',
    'Done',
    3
  );
}

// Create dropdowns sheet if missing
function createDropdownSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName("Dropdowns")) {
    const newSheet = ss.insertSheet("Dropdowns");
    newSheet.getRange("A1").setValue("Dropdown");
    newSheet.getRange("B1").setValue("Options");
  }
}

// Get headers and dropdown info for form generation
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

// Get dropdown options from Dropdowns sheet
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
        }
      } else {
        options[key] = value.split(',').map(opt => opt.trim());
      }
    }
  }
  return options;
}

// Get all visible records (rows not filtered out)
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
  return records.map(row => {
    return headers.reduce((obj, header, i) => {
      obj[header] = row[i];
      return obj;
    }, {});
  });
}

// Add new record to sheet
function addRecord(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Generate new numeric ID (max existing + 1)
  const lastId = sheet.getLastRow() > 1 ? Number(sheet.getRange(sheet.getLastRow(), 1).getValue()) || 0 : 0;
  const newId = lastId + 1;

  const row = headers.map(header => header === 'ID' ? newId : formData[header] || '');
  sheet.appendRow(row);
  return { status: 'success', id: newId };
}

// Update existing record by row number
function updateRecord(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  if (!formData._rowNumber || formData._rowNumber <= 1) {
    return { status: 'error', message: 'Invalid row number' };
  }

  // Get existing values and formulas in the row
  const existingRowValues = sheet.getRange(formData._rowNumber, 1, 1, headers.length).getValues()[0];
  const existingRowFormulas = sheet.getRange(formData._rowNumber, 1, 1, headers.length).getFormulas()[0];

  // Build updated row, preserving formulas intact
  const updatedRow = headers.map((header, idx) => {
    if (existingRowFormulas[idx]) {
      // Preserve formula in this cell; do NOT overwrite with form data
      return existingRowFormulas[idx];
    } else {
      // No formula here; update with form data if present, else keep existing value
      return (formData[header] !== '' && formData[header] !== undefined)
        ? formData[header]
        : existingRowValues[idx];
    }
  });

  // Write updated row back (formulas intact, values updated)
  sheet.getRange(formData._rowNumber, 1, 1, headers.length).setValues([updatedRow]);

  return { status: 'success', row: formData._rowNumber };
}


// Delete record by ID (value in column A)
function deleteRecord(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { status: 'success' };
    }
  }
  return { status: 'error', message: 'Record not found' };
}

// Get a single record by ID from the sheet
function getRecordById(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      const record = {};
      headers.forEach((header, idx) => {
        record[header] = data[i][idx];
      });
      return record;
    }
  }
  return null;
}
