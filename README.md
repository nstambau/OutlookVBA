# OutlookVBA
Macro to increase efficiency in Outlook


# Overview
This script leaverages a data file where each row starts with a search string (email address) and contains 1 or more urls or email addresses associated with that search string. This data needs to be locally saved / updated.


# Installation Notes:

## Build Data Source
- Create Source Data information using template from "Sample.xls" file
- Save as an "Excel 97-2003 Workbook (*.xls)"
- Note the path to this file (ex:  fileName = "C:\Users\\####\VBA Script\Sample.xls"

## Copy and update code

Open Microsoft Outlook App

### Update Security

-	On Main Outlook window, open menu "File" > "Options" > "Trust Ceter"
-	Open "Trust Center Settings"
-	Select "Macro Settings" and select "Notifications for all macros"

### Copy Code
-	Press Alt-F11 to open Microsoft VBA window
-	Paste this code into "ThisOutlookSession"
-	Open Menu "Tools" > "References" and check box next to "Microsoft Excel Object Library"

### Modify Code
-	Locate "getInfo" function, and rename the "filename" variable to the above dsta source

## Add buttons to menu
- Right click on ribbon and "Customize..."
- Remove unused items to generally declutter
- Create new group (use any name and icon)
- In the panel on left, select "Macros..." from drop down, and add to new group
- Rename and choose icon for custom macros!

