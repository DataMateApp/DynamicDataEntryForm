Dynamic Form App (Sidebar Interface)
This dynamic form provides an alternative interface for data entry, editing, and navigation directly within the spreadsheet sidebar.

Launch the Form:
From the spreadsheet, go to Extensions > Apps Script and paste in the provided code.
Or if already installed, open from Dynamic Form > Open Form in the menu bar.
Initial Setup:
On first launch, a Dropdowns sheet will be created with headers A1: "Dropdown" and B1: "Options".
You may define dropdown options here using comma-separated values (e.g., Status,Active,Inactive).
You may also define dropdown options here using sheet and range (e.g., contacts!A:A).
Form Behavior:
The form dynamically reads the first row of your active sheet as field headers.
Data validation or entries in the Dropdowns sheet are used to create select menus.
The ID field is auto-generated and read-only.
Using the Form:
Fill out or edit entries in the form.
Click:
Save: Add new or update existing records based on ID.
New Record: Reset the form to blank state.
Previous / Next: Navigate between existing visible rows.
Form Data Handling:
Data is added directly to the active sheet as new rows.
Only visible rows are navigable—filtered rows are skipped.
Dropdowns configured either by in-sheet validation or the Dropdowns sheet are respected.
Error Handling:
Required fields (except ID) are validated before saving.
User-friendly messages show up in the sidebar (green for success, red for errors).
Watch the Dynamic Sidebar Form Video
