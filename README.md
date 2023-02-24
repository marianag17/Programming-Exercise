# Programming-Exercise
Programming tests for Genpact Backend Developer Position

# Python Programming Exercise

Solution Description:
Build a solution, which monitor a folder looking for new files.
1. Each time a file is found, it should verify if is an excel file (.xls* files). If is true, it should take each
sheet on it and consolidate it on a master workbook file (make a copy from each sheet to the
master file).
2. It should have an option to choose which folder to watch.
3. Every file found should be moved to 2 different folders depending if was or not a excel file
- Processed
- Not applicable.

# VBA Programming Exercise

Solution Description:
1. Create a sheet on excel which will be used as database. It should have five different columns:
Code, Name, date of birth, Email, Home Address. This sheet should be always hidden with no
option to unhide manually just programmatically.
2. Create a user form with the options below:
* New record: it should enable the user to create new records on the database sheet, consider the points below:
  - Code: mandatory field, locked, Auto-numeric/correlative. It should automatically populate the next code based on database data. Codes correlative increments by 1.
  - Name: mandatory field.
  - Date of birth: mandatory field, it should be validating its on date format (mm/dd/yyyy) and lower the current date.
  - Email: optional field, it should validate email format (ex: aa@aa.com) or blank.
  - Home address: optional
* Search: it should enable the user search elements on the database sheet. It should be able to use Code (exact match) or Name (Contains the keywords) as search key. If Name is used as search key all multiple coincidences found should be displayed. The user should be able to edit or delete the selected record found.
  - Edit: all rules from new records applies.
  - Edit/Delete: show a confirmation box before saving changes.
* Export data from database sheet to new excel book. It should have 2 different modes:
  - By code range: export all data between code ranges selected (ex: from 5 to 10)
  - By names Initials: export all data starting with the keyword selected.
