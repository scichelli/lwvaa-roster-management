# lwvaa-roster-management
Volunteer support for LWVAA member roster data synchronization

## Overview

This repo shares scripts that are useful for comparing a data export from the National LWV member roster with a data export from a club's member roster to detect duplicates and discrepancies.

The solution uses Excel macros, which you can read, review, copy into your own local Excel workbook, and modify to suit your club's data export.

## Using Excel Macros for Data Comparison

Once you set up a Macro-Enabled Workbook with the `roster-compare` macro one time, you should be able to re-use that file to import and compare future data exports. After you follow these steps to wire up the button, you can compare new data exports by opening the workbook and clicking the button you created.


**Create a new Excel file.**

**Change its type so that it supports macros:**
1. Open the "File" menu.
2. Select "Save As".
3. Enter a name for the file.
4. In the dropdown list to specify the file type (It defaults to `Excel Workbook (*.xlsx)`.), select `Excel Macro-Enabled Workbook (*.xlsm)` instead.
5. Click the "Save" button.
<img width="1611" height="650" alt="Save_Excel_Macro-Enabled_Workbook" src="https://github.com/user-attachments/assets/4a324946-416f-46d3-88e6-2fe581a6c2f2" />


**Enable the Developer toolbar if not already visible:**
1. Open the "File" menu.
2. Select "Options" near the bottom left.
3. In the list of Options, select "Customize Ribbon".
4. In the right-hand "Customize the Ribbon" list of options for "Main Tabs", enable the checkbox for "Developer".
5. Click "OK" to close the Options dialog.
<img width="1032" height="845" alt="Add_Developer_toolbar_to_Ribbon" src="https://github.com/user-attachments/assets/e2f9b17d-b696-4bf1-ba4e-7aa3cf2ab58d" />


**Add a button to the first (probably only) worksheet:**
1. On the Developer tab, click "Insert" to open the list of Form Controls.
2. From the Form Controls, select "Button".
3. On the worksheet, click and drag to draw a button.
4. When you release the mouse after drawing the button, a dialog will pop up to assign a macro.
5. Click "New" to create a new macro.
6. The VBA editor will open, with a default subroutine definition.

**Replace the default macro with the code from this repository:**
1. Find the macro definition in this repository, in [macros/roster-compare.vba](./macros/roster-compare.vba).
2. Read the code to confirm that you are comfortable with the actions it is going to take.
3. Click the "Raw" button in GitHub to open a [plain-text view of the code](https://raw.githubusercontent.com/scichelli/lwvaa-roster-management/refs/heads/main/macros/roster-compare.vba) to make it easy to select only the code, not the rest of the GitHub website.
4. Select and copy the code.
5. Back in the VBA editor for Excel, replace the default subroutine and paste the `roster-compare.vba` code.
6. Click "Save" in the VBA editor.

**Connect the button to the macro:**
1. The spreadsheet is open in a different window.
2. Right-click on the button.
3. You can use "Edit Text" to give it a helpful label.
4. From the right-click menu, select "Assign Macro...".
5. From the list of available macros, choose `RunSynchronization`, which is the macro defined by the code in this repo.
6. Click "OK" to assign the macro to the button.
<img width="775" height="608" alt="Assign_Macro_to_Button" src="https://github.com/user-attachments/assets/d55cef02-a187-41e0-b0d0-15ba4895ebc4" />

**Use the macro to import and compare rosters:**
1. When you have new data exports to compare, you should be able to start here, reusing the Macro-Enabled Workbook you already set up.
2. If there is a new version of the macro code and you would like to take advantage of those improvements, then you can use the VBA editor to replace the macro with the new version. Alternatively, it is ok to follow the steps from the beginning to set up a new workbook, if that feels more comfortable.
3. Click the button and follow the prompts.
