Sub LoadNationalRoster()
    LoadExcelFileToNewSheet "National"
End Sub

Sub LoadClubRoster()
    LoadExcelFileToNewSheet "Club"
End Sub

Sub LoadExcelFileToNewSheet(rosterName As String)
    Dim filePath As String
    Dim sourceWB As Workbook
    Dim sourceWS As Worksheet
    Dim targetWS As Worksheet

    ' Prompt user to select an Excel file
    filePath = Application.GetOpenFilename("Excel Files (*.xlsx; *.xls), *.xlsx; *.xls", , "Select an Excel File")

    ' Exit if user cancels
    If filePath = "False" Then Exit Sub

    ' Open the selected workbook
    Set sourceWB = Workbooks.Open(filePath)
    Set sourceWS = sourceWB.Sheets(1) ' You can change this to a specific sheet name or index

    ' Add a new worksheet to the current workbook
    Set targetWS = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets("Dashboard"))
    targetWS.Name = rosterName & "_" & Format(Now, "yyyymmdd") & "_" & Format(Now, "hhmmss")

    ' Copy data from source worksheet to target worksheet
    sourceWS.UsedRange.Copy Destination:=targetWS.Range("A1")

    ' Close the source workbook without saving changes
    sourceWB.Close SaveChanges:=False

    MsgBox "File imported successfully into sheet: " & targetWS.Name
End Sub

