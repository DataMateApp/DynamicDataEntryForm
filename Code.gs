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
        #searchContainer { margin-bottom: 20px; }
      </style>
    </head>
    <body>
      <div id="searchContainer">
        <label for="searchId">Search by ID</label>
        <select id="searchId">
          <option value="">Select ID</option>
        </select>
        <button type="button" onclick="searchRecord()">Search</button>
      </div>
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
        let currentRowNumber = null;

        google.script.run
          .withSuccessHandler(populateForm)
          .withFailureHandler(err => {
            console.error('Error loading headers: ' + err.message);
            showMessage('Error loading form fields: ' + err.message, 'error');
          })
          .getSheetInfo();
        google.script.run
          .withSuccessHandler(loadRecords)
          .withFailureHandler(err => {
            console.error('Error loading records: ' + err.message);
            showMessage('Error loading records: ' + err.message, 'error');
          })
          .getVisibleRecords();
        google.script.run
          .withSuccessHandler(populateIdDropdown)
          .withFailureHandler(err => {
            console.error('Error loading ID dropdown: ' + err.message);
            showMessage('Error loading ID dropdown: ' + err.message, 'error');
          })
          .getColumnAValues();

        function populateForm(headerData) {
          if (!headerData || headerData.length === 0) {
            console.error('No headers received from getSheetInfo');
            showMessage('No fields found in the target sheet. Please check the sheet headers.', 'error');
            return;
          }
          headers = headerData;
          const formFields = document.getElementById('formFields');
          formFields.innerHTML = headers.map(header => {
            if (header.name === 'ID') {
              return '<label for="' + header.name + '">' + header.name + '</label>' +
                     '<input type="number" id="' + header.name + '" readonly>';
            } else if (header.type === 'select') {
              if (!header.options || header.options.length === 0) {
                console.warn(\`No options available for dropdown \${header.name}\`);
                showMessage(\`Warning: No options for \${header.name} dropdown. Check sheet configuration.\`, 'error');
                return '<label for="' + header.name + '">' + header.name + (header.required ? ' *' : '') + '</label>' +
                       '<select id="' + header.name + '" ' + (header.required ? 'required' : '') + ' onchange="onDropdownChange()">' +
                       '<option value="">Select ' + header.name + '</option>' +
                       '</select>';
              }
              console.log(\`Populating dropdown \${header.name} with options: \${header.options.join(', ')}\`);
              return '<label for="' + header.name + '">' + header.name + (header.required ? ' *' : '') + '</label>' +
                     '<select id="' + header.name + '" ' + (header.required ? 'required' : '') + ' onchange="onDropdownChange()">' +
                     '<option value="">Select ' + header.name + '</option>' +
                     header.options.map(opt => '<option value="' + opt + '">' + opt + '</option>').join('') +
                     '</select>';
            } else {
              return '<label for="' + header.name + '">' + header.name + (header.required ? ' *' : '') + '</label>' +
                     '<input type="' + header.type + '" id="' + header.name + '" ' + (header.required ? 'required' : '') + '>';
            }
          }).join('');
          console.log('Form fields populated with headers: ' + headers.map(h => h.name).join(', '));
          showMessage('Form fields loaded successfully.', '');
        }

        function populateIdDropdown(idValues) {
          const searchIdDropdown = document.getElementById('searchId');
          searchIdDropdown.innerHTML = '<option value="">Select ID</option>';
          if (idValues.length === 0) {
            console.warn('No IDs found in column A');
            showMessage('No IDs found in column A of the target sheet.', 'error');
          } else {
            idValues.forEach(id => {
              if (id !== '') {
                const option = document.createElement('option');
                option.value = id;
                option.textContent = id;
                searchIdDropdown.appendChild(option);
              }
            });
            console.log('ID dropdown populated with: ' + idValues.join(', '));
          }
        }

        function loadRecords(data) {
          records = data;
          if (records.length > 0) {
            currentIndex = 0;
            currentRowNumber = currentIndex + 2;
            isNewRecord = false;
            displayRecord();
          } else {
            clearForm();
            console.warn('No records loaded');
            showMessage('No records found in the target sheet.', 'error');
          }
          google.script.run.withSuccessHandler(populateIdDropdown).getColumnAValues();
        }

        function displayRecord() {
          isNewRecord = false;
          if (currentIndex >= 0 && currentIndex < records.length) {
            headers.forEach(header => {
              const field = document.getElementById(header.name);
              if (!field) {
                console.error(\`Field \${header.name} not found in form\`);
                showMessage(\`Error: Field \${header.name} not found in form.\`, 'error');
                return;
              }
              const value = records[currentIndex][header.name] != null ? records[currentIndex][header.name] : '';
              if (field.tagName === 'SELECT') {
                let exists = Array.from(field.options).some(opt => opt.value === String(value));
                if (!exists && value !== '') {
                  const opt = document.createElement('option');
                  opt.value = String(value);
                  opt.textContent = String(value);
                  field.appendChild(opt);
                  console.log(\`Added option \${value} to dropdown \${header.name}\`);
                }
                field.value = String(value);
              } else {
                field.value = value;
              }
            });
            const currentId = records[currentIndex]['ID'] != null ? String(records[currentIndex]['ID']) : '';
            document.getElementById('searchId').value = currentId;
            console.log(\`Displayed record at index \${currentIndex}: ID=\${currentId}, Row=\${currentRowNumber}, Data=\${JSON.stringify(records[currentIndex])}\`);
          } else {
            console.error(\`Invalid currentIndex: \${currentIndex}, records length: \${records.length}\`);
            showMessage('Error: No record to display.', 'error');
          }
        }

        function searchRecord() {
          const searchId = document.getElementById('searchId').value;
          if (!searchId) {
            console.warn('No ID selected for search');
            showMessage('Please select an ID to search.', 'error');
            return;
          }
          console.log(\`Initiating search for ID: \${searchId}\`);
          document.getElementById('spinner').style.display = 'block';
          google.script.run
            .withSuccessHandler(result => {
              document.getElementById('spinner').style.display = 'none';
              if (result && Object.keys(result.record).length > 0) {
                console.log(\`Record found for ID \${searchId}: \${JSON.stringify(result.record)}, Row=\${result.rowNumber}\`);
                records = records.filter(r => String(r.ID) !== String(searchId));
                records.push(result.record);
                currentIndex = records.length - 1;
                currentRowNumber = result.rowNumber;
                isNewRecord = false;
                displayRecord();
                showMessage('Record found and displayed.', '');
              } else {
                console.error(\`No record found for ID: \${searchId}\`);
                showMessage(\`No record found with ID: \${searchId}\`, 'error');
              }
            })
            .withFailureHandler(err => {
              document.getElementById('spinner').style.display = 'none';
              console.error(\`Search failed for ID \${searchId}: \${err.message}\`);
              showMessage('Error searching record: ' + err.message, 'error');
            })
            .getRecordById(searchId);
        }

        function saveRecord() {
          document.getElementById('spinner').style.display = 'block';
          const formData = {};
          headers.forEach(header => {
            const field = document.getElementById(header.name);
            if (!field) {
              console.error(\`Field \${header.name} not found during save\`);
              showMessage(\`Error: Field \${header.name} not found.\`, 'error');
              document.getElementById('spinner').style.display = 'none';
              return;
            }
            formData[header.name] = field.value.trim();
          });

          for (const header of headers) {
            if (header.required && !formData[header.name]) {
              console.warn(\`Required field \${header.name} is empty\`);
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
          } else if (currentRowNumber) {
            formData._rowNumber = currentRowNumber;
            google.script.run
              .withSuccessHandler(result => onSave(result, currentRowNumber))
              .withFailureHandler(onError)
              .updateRecord(formData);
          } else {
            console.error('No valid row number for updating record');
            showMessage('Error: Cannot update record, no row number available.', 'error');
            document.getElementById('spinner').style.display = 'none';
          }
        }

        function clearForm() {
          isNewRecord = true;
          currentRowNumber = null;
          document.getElementById('dynamicForm').reset();
          headers.forEach(header => {
            if (header.name === 'ID') return;
            const field = document.getElementById(header.name);
            if (field) field.value = '';
          });
          document.getElementById('searchId').value = '';
          console.log('Form cleared for new record');
          showMessage('Ready for new record.', '');
        }

        function navigate(direction) {
          if (records.length === 0) {
            console.warn('No records to navigate');
            return;
          }
          if (direction === 'prev' && currentIndex > 0) {
            currentIndex--;
            currentRowNumber = currentIndex + 2;
          } else if (direction === 'next' && currentIndex < records.length - 1) {
            currentIndex++;
            currentRowNumber = currentIndex + 2;
          }
          console.log(\`Navigating to \${direction}, new index: \${currentIndex}, row: \${currentRowNumber}\`);
          displayRecord();
        }

        function onDropdownChange() {
          if (isNewRecord) return;
          const currentID = document.getElementById('ID').value;
          if (!currentID) {
            console.warn('No ID for dropdown change refresh');
            return;
          }
          console.log(\`Refreshing record for ID: \${currentID} due to dropdown change\`);
          document.getElementById('spinner').style.display = 'block';
          google.script.run
            .withSuccessHandler(result => {
              document.getElementById('spinner').style.display = 'none';
              if (result && result.record) {
                headers.forEach(header => {
                  const field = document.getElementById(header.name);
                  if (!field) {
                    console.error(\`Field \${header.name} not found on dropdown change\`);
                    return;
                  }
                  const val = result.record[header.name] != null ? result.record[header.name] : '';
                  if (field.tagName === 'SELECT') {
                    let exists = Array.from(field.options).some(opt => opt.value === String(val));
                    if (!exists && val !== '') {
                      const opt = document.createElement('option');
                      opt.value = String(val);
                      opt.textContent = String(val);
                      field.appendChild(opt);
                      console.log(\`Added option \${val} to dropdown \${header.name}\`);
                    }
                    field.value = String(val);
                  } else {
                    field.value = val;
                  }
                });
                document.getElementById('searchId').value = result.record['ID'] != null ? String(result.record['ID']) : '';
                currentRowNumber = result.rowNumber;
                console.log(\`Refreshed record for ID \${currentID}: \${JSON.stringify(result.record)}, Row=\${result.rowNumber}\`);
                showMessage('Record refreshed with formula updates.', '');
              } else {
                console.error(\`No record found for ID \${currentID} on refresh\`);
                showMessage('Record not found on reload.', 'error');
              }
            })
            .withFailureHandler(err => {
              document.getElementById('spinner').style.display = 'none';
              console.error(\`Error refreshing record for ID \${currentID}: \${err.message}\`);
              showMessage('Error refreshing record: ' + err.message, 'error');
            })
            .getRecordById(currentID);
        }

        function onSave(result, existingRow) {
          document.getElementById('spinner').style.display = 'none';
          if (result.status === 'success') {
            console.log(\`Record saved: \${JSON.stringify(result)}\`);
            showMessage('Record saved successfully.', '');
            google.script.run
              .withSuccessHandler(data => {
                records = data;
                if (existingRow) {
                  const updatedID = document.getElementById('ID').value;
                  const idx = records.findIndex(r => String(r.ID) === String(updatedID));
                  if (idx >= 0) {
                    currentIndex = idx;
                    currentRowNumber = idx + 2;
                    displayRecord();
                  } else {
                    currentIndex = records.length - 1;
                    currentRowNumber = currentIndex + 2;
                    displayRecord();
                  }
                } else {
                  currentIndex = records.length - 1;
                  currentRowNumber = currentIndex + 2;
                  displayRecord();
                }
                google.script.run.withSuccessHandler(populateIdDropdown).getColumnAValues();
              })
              .withFailureHandler(err => {
                console.error('Error reloading records after save: ' + err.message);
                showMessage('Error reloading records: ' + err.message, 'error');
              })
              .getVisibleRecords();
            isNewRecord = false;
            document.getElementById('searchId').value = '';
          } else {
            console.error(\`Save failed: \${result.message}\`);
            showMessage(result.message || 'Error saving record.', 'error');
          }
        }

        function onError(error) {
          document.getElementById('spinner').style.display = 'none';
          console.error('Operation error: ' + error.message);
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

function doGet() {
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
        #searchContainer { margin-bottom: 20px; }
      </style>
    </head>
    <body>
      <div id="searchContainer">
        <label for="searchId">Search by ID</label>
        <select id="searchId">
          <option value="">Select ID</option>
        </select>
        <button type="button" onclick="searchRecord()">Search</button>
      </div>
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
        let currentRowNumber = null;

        google.script.run
          .withSuccessHandler(populateForm)
          .withFailureHandler(err => {
            console.error('Error loading headers: ' + err.message);
            showMessage('Error loading form fields: ' + err.message, 'error');
          })
          .getSheetInfo();
        google.script.run
          .withSuccessHandler(loadRecords)
          .withFailureHandler(err => {
            console.error('Error loading records: ' + err.message);
            showMessage('Error loading records: ' + err.message, 'error');
          })
          .getVisibleRecords();
        google.script.run
          .withSuccessHandler(populateIdDropdown)
          .withFailureHandler(err => {
            console.error('Error loading ID dropdown: ' + err.message);
            showMessage('Error loading ID dropdown: ' + err.message, 'error');
          })
          .getColumnAValues();

        function populateForm(headerData) {
          if (!headerData || headerData.length === 0) {
            console.error('No headers received from getSheetInfo');
            showMessage('No fields found in the target sheet. Please check the sheet headers.', 'error');
            return;
          }
          headers = headerData;
          const formFields = document.getElementById('formFields');
          formFields.innerHTML = headers.map(header => {
            if (header.name === 'ID') {
              return '<label for="' + header.name + '">' + header.name + '</label>' +
                     '<input type="number" id="' + header.name + '" readonly>';
            } else if (header.type === 'select') {
              if (!header.options || header.options.length === 0) {
                console.warn(\`No options available for dropdown \${header.name}\`);
                showMessage(\`Warning: No options for \${header.name} dropdown. Check sheet configuration.\`, 'error');
                return '<label for="' + header.name + '">' + header.name + (header.required ? ' *' : '') + '</label>' +
                       '<select id="' + header.name + '" ' + (header.required ? 'required' : '') + ' onchange="onDropdownChange()">' +
                       '<option value="">Select ' + header.name + '</option>' +
                       '</select>';
              }
              console.log(\`Populating dropdown \${header.name} with options: \${header.options.join(', ')}\`);
              return '<label for="' + header.name + '">' + header.name + (header.required ? ' *' : '') + '</label>' +
                     '<select id="' + header.name + '" ' + (header.required ? 'required' : '') + ' onchange="onDropdownChange()">' +
                     '<option value="">Select ' + header.name + '</option>' +
                     header.options.map(opt => '<option value="' + opt + '">' + opt + '</option>').join('') +
                     '</select>';
            } else {
              return '<label for="' + header.name + '">' + header.name + (header.required ? ' *' : '') + '</label>' +
                     '<input type="' + header.type + '" id="' + header.name + '" ' + (header.required ? 'required' : '') + '>';
            }
          }).join('');
          console.log('Form fields populated with headers: ' + headers.map(h => h.name).join(', '));
          showMessage('Form fields loaded successfully.', '');
        }

        function populateIdDropdown(idValues) {
          const searchIdDropdown = document.getElementById('searchId');
          searchIdDropdown.innerHTML = '<option value="">Select ID</option>';
          if (idValues.length === 0) {
            console.warn('No IDs found in column A');
            showMessage('No IDs found in column A of the target sheet.', 'error');
          } else {
            idValues.forEach(id => {
              if (id !== '') {
                const option = document.createElement('option');
                option.value = id;
                option.textContent = id;
                searchIdDropdown.appendChild(option);
              }
            });
            console.log('ID dropdown populated with: ' + idValues.join(', '));
          }
        }

        function loadRecords(data) {
          records = data;
          if (records.length > 0) {
            currentIndex = 0;
            currentRowNumber = currentIndex + 2;
            isNewRecord = false;
            displayRecord();
          } else {
            clearForm();
            console.warn('No records loaded');
            showMessage('No records found in the target sheet.', 'error');
          }
          google.script.run.withSuccessHandler(populateIdDropdown).getColumnAValues();
        }

        function displayRecord() {
          isNewRecord = false;
          if (currentIndex >= 0 && currentIndex < records.length) {
            headers.forEach(header => {
              const field = document.getElementById(header.name);
              if (!field) {
                console.error(\`Field \${header.name} not found in form\`);
                showMessage(\`Error: Field \${header.name} not found in form.\`, 'error');
                return;
              }
              const value = records[currentIndex][header.name] != null ? records[currentIndex][header.name] : '';
              if (field.tagName === 'SELECT') {
                let exists = Array.from(field.options).some(opt => opt.value === String(value));
                if (!exists && value !== '') {
                  const opt = document.createElement('option');
                  opt.value = String(value);
                  opt.textContent = String(value);
                  field.appendChild(opt);
                  console.log(\`Added option \${value} to dropdown \${header.name}\`);
                }
                field.value = String(value);
              } else {
                field.value = value;
              }
            });
            const currentId = records[currentIndex]['ID'] != null ? String(records[currentIndex]['ID']) : '';
            document.getElementById('searchId').value = currentId;
            console.log(\`Displayed record at index \${currentIndex}: ID=\${currentId}, Row=\${currentRowNumber}, Data=\${JSON.stringify(records[currentIndex])}\`);
          } else {
            console.error(\`Invalid currentIndex: \${currentIndex}, records length: \${records.length}\`);
            showMessage('Error: No record to display.', 'error');
          }
        }

        function searchRecord() {
          const searchId = document.getElementById('searchId').value;
          if (!searchId) {
            console.warn('No ID selected for search');
            showMessage('Please select an ID to search.', 'error');
            return;
          }
          console.log(\`Initiating search for ID: \${searchId}\`);
          document.getElementById('spinner').style.display = 'block';
          google.script.run
            .withSuccessHandler(result => {
              document.getElementById('spinner').style.display = 'none';
              if (result && result.record && Object.keys(result.record).length > 0) {
                console.log(\`Record found for ID \${searchId}: \${JSON.stringify(result.record)}, Row=\${result.rowNumber}\`);
                records = records.filter(r => String(r.ID) !== String(searchId));
                records.push(result.record);
                currentIndex = records.length - 1;
                currentRowNumber = result.rowNumber;
                isNewRecord = false;
                displayRecord();
                showMessage('Record found and displayed.', '');
              } else {
                console.error(\`No record found for ID: \${searchId}\`);
                showMessage(\`No record found with ID: \${searchId}\`, 'error');
              }
            })
            .withFailureHandler(err => {
              document.getElementById('spinner').style.display = 'none';
              console.error(\`Search failed for ID \${searchId}: \${err.message}\`);
              showMessage('Error searching record: ' + err.message, 'error');
            })
            .getRecordById(searchId);
        }

        function saveRecord() {
          document.getElementById('spinner').style.display = 'block';
          const formData = {};
          headers.forEach(header => {
            const field = document.getElementById(header.name);
            if (!field) {
              console.error(\`Field \${header.name} not found during save\`);
              showMessage(\`Error: Field \${header.name} not found.\`, 'error');
              document.getElementById('spinner').style.display = 'none';
              return;
            }
            formData[header.name] = field.value.trim();
          });

          for (const header of headers) {
            if (header.required && !formData[header.name]) {
              console.warn(\`Required field \${header.name} is empty\`);
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
          } else if (currentRowNumber) {
            formData._rowNumber = currentRowNumber;
            google.script.run
              .withSuccessHandler(result => onSave(result, currentRowNumber))
              .withFailureHandler(onError)
              .updateRecord(formData);
          } else {
            console.error('No valid row number for updating record');
            showMessage('Error: Cannot update record, no row number available.', 'error');
            document.getElementById('spinner').style.display = 'none';
          }
        }

        function clearForm() {
          isNewRecord = true;
          currentRowNumber = null;
          document.getElementById('dynamicForm').reset();
          headers.forEach(header => {
            if (header.name === 'ID') return;
            const field = document.getElementById(header.name);
            if (field) field.value = '';
          });
          document.getElementById('searchId').value = '';
          console.log('Form cleared for new record');
          showMessage('Ready for new record.', '');
        }

        function navigate(direction) {
          if (records.length === 0) {
            console.warn('No records to navigate');
            return;
          }
          if (direction === 'prev' && currentIndex > 0) {
            currentIndex--;
            currentRowNumber = currentIndex + 2;
          } else if (direction === 'next' && currentIndex < records.length - 1) {
            currentIndex++;
            currentRowNumber = currentIndex + 2;
          }
          console.log(\`Navigating to \${direction}, new index: \${currentIndex}, row: \${currentRowNumber}\`);
          displayRecord();
        }

        function onDropdownChange() {
          if (isNewRecord) return;
          const currentID = document.getElementById('ID').value;
          if (!currentID) {
            console.warn('No ID for dropdown change refresh');
            return;
          }
          console.log(\`Refreshing record for ID: \${currentID} due to dropdown change\`);
          document.getElementById('spinner').style.display = 'block';
          google.script.run
            .withSuccessHandler(result => {
              document.getElementById('spinner').style.display = 'none';
              if (result && result.record) {
                headers.forEach(header => {
                  const field = document.getElementById(header.name);
                  if (!field) {
                    console.error(\`Field \${header.name} not found on dropdown change\`);
                    return;
                  }
                  const val = result.record[header.name] != null ? result.record[header.name] : '';
                  if (field.tagName === 'SELECT') {
                    let exists = Array.from(field.options).some(opt => opt.value === String(val));
                    if (!exists && val !== '') {
                      const opt = document.createElement('option');
                      opt.value = String(val);
                      opt.textContent = String(val);
                      field.appendChild(opt);
                      console.log(\`Added option \${val} to dropdown \${header.name}\`);
                    }
                    field.value = String(val);
                  } else {
                    field.value = val;
                  }
                });
                document.getElementById('searchId').value = result.record['ID'] != null ? String(result.record['ID']) : '';
                currentRowNumber = result.rowNumber;
                console.log(\`Refreshed record for ID \${currentID}: \${JSON.stringify(result.record)}, Row=\${result.rowNumber}\`);
                showMessage('Record refreshed with formula updates.', '');
              } else {
                console.error(\`No record found for ID \${currentID} on refresh\`);
                showMessage('Record not found on reload.', 'error');
              }
            })
            .withFailureHandler(err => {
              document.getElementById('spinner').style.display = 'none';
              console.error(\`Error refreshing record for ID \${currentID}: \${err.message}\`);
              showMessage('Error refreshing record: ' + err.message, 'error');
            })
            .getRecordById(currentID);
        }

        function onSave(result, existingRow) {
          document.getElementById('spinner').style.display = 'none';
          if (result.status === 'success') {
            console.log(\`Record saved: \${JSON.stringify(result)}\`);
            showMessage('Record saved successfully.', '');
            google.script.run
              .withSuccessHandler(data => {
                records = data;
                if (existingRow) {
                  const updatedID = document.getElementById('ID').value;
                  const idx = records.findIndex(r => String(r.ID) === String(updatedID));
                  if (idx >= 0) {
                    currentIndex = idx;
                    currentRowNumber = idx + 2;
                    displayRecord();
                  } else {
                    currentIndex = records.length - 1;
                    currentRowNumber = currentIndex + 2;
                    displayRecord();
                  }
                } else {
                  currentIndex = records.length - 1;
                  currentRowNumber = currentIndex + 2;
                  displayRecord();
                }
                google.script.run.withSuccessHandler(populateIdDropdown).getColumnAValues();
              })
              .withFailureHandler(err => {
                console.error('Error reloading records after save: ' + err.message);
                showMessage('Error reloading records: ' + err.message, 'error');
              })
              .getVisibleRecords();
            isNewRecord = false;
            document.getElementById('searchId').value = '';
          } else {
            console.error(\`Save failed: \${result.message}\`);
            showMessage(result.message || 'Error saving record.', 'error');
          }
        }

        function onError(error) {
          document.getElementById('spinner').style.display = 'none';
          console.error('Operation error: ' + error.message);
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
  return HtmlService.createHtmlOutput(htmlContent)
    .setTitle('Dynamic Data Entry Form')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function createDropdownSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName("Dropdowns")) {
    const newSheet = ss.insertSheet("Dropdowns");
    newSheet.getRange("A1").setValue("Dropdown");
    newSheet.getRange("B1").setValue("Options");
    newSheet.getRange("C1").setValue("Target Sheet");
    newSheet.getRange("C2").setValue(ss.getActiveSheet().getName());
    Logger.log("Created Dropdowns sheet with default target: " + ss.getActiveSheet().getName());
  }
}

function getDropdownOptions(dropdownsSheet) {
  const data = dropdownsSheet.getDataRange().getValues();
  const options = {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log(`Processing Dropdowns sheet with ${data.length - 1} entries`);
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0]?.toString().trim();
    const value = data[i][1]?.toString().trim();
    if (!key || !value) {
      Logger.log(`Skipping invalid dropdown entry at row ${i + 1}: key=${key}, value=${value}`);
      continue;
    }
    Logger.log(`Processing dropdown: ${key}, options: ${value}`);
    if (value.includes('!')) {
      try {
        const [sheetName, colRange] = value.split('!');
        Logger.log(`Attempting to access sheet: ${sheetName}, range: ${colRange}`);
        const sourceSheet = ss.getSheetByName(sheetName);
        if (!sourceSheet) {
          Logger.log(`Error: Source sheet ${sheetName} for dropdown ${key} not found`);
          continue;
        }
        const lastRow = sourceSheet.getLastRow();
        if (lastRow < 1) {
          Logger.log(`Error: No data in sheet ${sheetName} for dropdown ${key}`);
          continue;
        }
        const range = sourceSheet.getRange(colRange);
        if (!range || range.isBlank()) {
          Logger.log(`Error: Range ${colRange} in sheet ${sheetName} is invalid or empty for dropdown ${key}`);
          continue;
        }
        const values = sourceSheet.getRange(1, range.getColumn(), lastRow, 1).getValues().flat().filter(v => v != null && v.toString().trim() !== '');
        if (values.length === 0) {
          Logger.log(`Warning: No valid values in ${sheetName}!${colRange} for dropdown ${key}`);
          continue;
        }
        options[key] = [...new Set(values.map(v => v.toString().trim()))];
        Logger.log(`Dropdown options for ${key} from ${sheetName}!${colRange}: ${JSON.stringify(options[key])}`);
      } catch (e) {
        Logger.log(`Error parsing dropdown options for ${key} from ${value}: ${e.message}`);
        continue;
      }
    } else {
      options[key] = value.split(',').map(opt => opt.trim()).filter(opt => opt !== '');
      Logger.log(`Dropdown options for ${key}: ${JSON.stringify(options[key])}`);
    }
  }
  Logger.log(`Final dropdown options: ${JSON.stringify(options)}`);
  return options;
}

function protectAllFormulaCells() {
  const sheet = getTargetSheet();
  const range = sheet.getDataRange();
  const formulas = range.getFormulas();

  for (let r = 0; r < formulas.length; r++) {
    for (let c = 0; c < formulas[r].length; c++) {
      if (formulas[r][c]) {
        const cell = sheet.getRange(r + 1, c + 1);
        const protection = cell.protect();
        protection.setDescription('Formula cell - do not edit');
        protection.removeEditors(protection.getEditors());
        Logger.log(`Protected formula cell at ${sheet.getName()}!${cell.getA1Notation()}`);
      }
    }
  }
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'All formula cells have been protected.',
    'Done',
    3
  );
}

function getTargetSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dropdownsSheet = ss.getSheetByName("Dropdowns");
  let targetSheetName = ss.getActiveSheet().getName();
  if (dropdownsSheet) {
    const targetCell = dropdownsSheet.getRange("C2").getValue()?.toString().trim();
    if (targetCell && ss.getSheetByName(targetCell)) {
      targetSheetName = targetCell;
      Logger.log(`Target sheet set to: ${targetSheetName}`);
    } else {
      Logger.log(`Invalid or missing sheet name in Dropdowns!C2: ${targetCell}. Falling back to active sheet: ${targetSheetName}`);
    }
  } else {
    Logger.log("Dropdowns sheet not found. Falling back to active sheet: " + targetSheetName);
  }
  const targetSheet = ss.getSheetByName(targetSheetName);
  if (!targetSheet) {
    Logger.log(`Error: Target sheet ${targetSheetName} does not exist`);
    throw new Error(`Target sheet ${targetSheetName} does not exist`);
  }
  return targetSheet;
}

function getSheetInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getTargetSheet();
  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) {
    Logger.log(`Error: No headers found in target sheet ${sheet.getName()}`);
    throw new Error(`No headers found in target sheet ${sheet.getName()}`);
  }
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].filter(h => h != null && h.toString().trim() !== '');
  const validations = lastColumn > 0 ? sheet.getRange(2, 1, 1, lastColumn).getDataValidations()[0] : [];
  const dropdownsSheet = ss.getSheetByName('Dropdowns');
  const dropdownOptions = dropdownsSheet ? getDropdownOptions(dropdownsSheet) : {};

  Logger.log(`Headers found in ${sheet.getName()}: ${JSON.stringify(headers)}`);
  Logger.log(`Dropdown options from Dropdowns sheet: ${JSON.stringify(dropdownOptions)}`);

  if (headers.length === 0) {
    Logger.log(`Error: No valid headers found in target sheet ${sheet.getName()}`);
    throw new Error(`No valid headers found in target sheet ${sheet.getName()}`);
  }

  return headers.map((header, index) => {
    const headerStr = header.toString().trim();
    const validation = validations[index] || null;
    let type = 'text';
    let options = [];

    if (validation && validation.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      type = 'select';
      options = validation.getCriteriaValues().filter(v => v != null && v.toString().trim() !== '');
      Logger.log(`Data validation for ${headerStr}: ${JSON.stringify(options)}`);
    }
    const dropdownKey = Object.keys(dropdownOptions).find(key => key.toLowerCase() === headerStr.toLowerCase());
    if (dropdownKey && dropdownOptions[dropdownKey]) {
      type = 'select';
      options = dropdownOptions[dropdownKey];
      Logger.log(`Dropdown options for ${headerStr} from Dropdowns sheet: ${JSON.stringify(options)}`);
    }
    if (headerStr === 'ID') {
      type = 'number';
      options = [];
    }
    Logger.log(`Header ${headerStr}: type=${type}, options=${JSON.stringify(options)}`);
    return {
      name: headerStr,
      type: type,
      options: options,
      required: headerStr !== 'ID',
      columnIndex: index + 1
    };
  });
}

function getColumnAValues() {
  const sheet = getTargetSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log(`No data rows in ${sheet.getName()}, returning empty array for column A`);
    return [];
  }
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const uniqueValues = [...new Set(values.filter(v => v != null && v.toString().trim() !== '').map(v => v.toString().trim()))];
  Logger.log(`Column A values in ${sheet.getName()}: ${JSON.stringify(uniqueValues)}`);
  return uniqueValues;
}

function getVisibleRecords() {
  const sheet = getTargetSheet();
  const filter = sheet.getFilter();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const records = [];

  if (headers.length === 0) {
    Logger.log(`No headers in ${sheet.getName()}, returning empty records`);
    return records;
  }

  if (filter) {
    for (let i = 1; i < data.length; i++) {
      if (!sheet.isRowHiddenByFilter(i + 1)) {
        records.push(data[i]);
      }
    }
  } else {
    records.push(...data.slice(1));
  }
  Logger.log(`Visible records in ${sheet.getName()}: ${records.length}`);
  return records.map((row, index) => {
    return headers.reduce((obj, header, i) => {
      obj[header] = row[i];
      return obj;
    }, {});
  });
}

function addRecord(formData) {
  const sheet = getTargetSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const lastId = sheet.getLastRow() > 1 ? Number(sheet.getRange(sheet.getLastRow(), 1).getValue()) || 0 : 0;
  const newId = lastId + 1;

  const row = headers.map(header => header === 'ID' ? newId : formData[header] || '');
  sheet.appendRow(row);
  Logger.log(`Added record with ID ${newId} to ${sheet.getName()}`);
  return { status: 'success', id: newId };
}

function updateRecord(formData) {
  const sheet = getTargetSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  if (!formData._rowNumber || formData._rowNumber <= 1) {
    Logger.log(`Error: Invalid row number ${formData._rowNumber} for update in ${sheet.getName()}`);
    return { status: 'error', message: 'Invalid row number' };
  }

  const existingRowValues = sheet.getRange(formData._rowNumber, 1, 1, headers.length).getValues()[0];
  const existingRowFormulas = sheet.getRange(formData._rowNumber, 1, 1, headers.length).getFormulas()[0];

  const updatedRow = headers.map((header, idx) => {
    if (existingRowFormulas[idx]) {
      return existingRowFormulas[idx];
    } else {
      return (formData[header] !== '' && formData[header] !== undefined)
        ? formData[header]
        : existingRowValues[idx];
    }
  });

  sheet.getRange(formData._rowNumber, 1, 1, headers.length).setValues([updatedRow]);
  Logger.log(`Updated record at row ${formData._rowNumber} in ${sheet.getName()}`);
  return { status: 'success', row: formData._rowNumber };
}

function deleteRecord(id) {
  const sheet = getTargetSheet();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      Logger.log(`Deleted record with ID ${id} from ${sheet.getName()}`);
      return { status: 'success' };
    }
  }
  Logger.log(`No record found to delete with ID ${id} in ${sheet.getName()}`);
  return { status: 'error', message: 'Record not found' };
}

function getRecordById(id) {
  const sheet = getTargetSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = sheet.getDataRange().getValues();

  Logger.log(`Searching for ID ${id} in ${sheet.getName()}`);
  for (let i = 1; i < data.length; i++) {
    const sheetId = String(data[i][0]).trim();
    const searchId = String(id).trim();
    if (sheetId === searchId) {
      const record = {};
      headers.forEach((header, idx) => {
        record[header] = data[i][idx];
      });
      Logger.log(`Record found for ID ${id}: ${JSON.stringify(record)}, Row=${i + 1}`);
      return { record: record, rowNumber: i + 1 };
    }
  }
  Logger.log(`No record found for ID ${id} in ${sheet.getName()}`);
  return { record: null, rowNumber: null };
}

function createDropdownSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName("Dropdowns")) {
    const newSheet = ss.insertSheet("Dropdowns");
    newSheet.getRange("A1").setValue("Dropdown");
    newSheet.getRange("B1").setValue("Options");
    newSheet.getRange("C1").setValue("Target Sheet");
    newSheet.getRange("C2").setValue(ss.getActiveSheet().getName());
  }
}
