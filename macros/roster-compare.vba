Sub RunSynchronization()
    Dim nationalWorksheetName, clubWorksheetName As String
    
    ' Begin: load rosters into worksheets
    MsgBox "First we'll load the National roster into a worksheet"
    nationalWorksheetName = LoadNationalRoster()
    
    MsgBox "Next we'll load the Club roster into a worksheet"
    clubWorksheetName = LoadClubRoster()
    ' End: load rosters into worksheets
    
    If nationalWorksheetName = vbNullString Or clubWorksheetName = vbNullString Then
        MsgBox "Did not find two worksheets to compare"
        Exit Sub
    End If
    
    ' Begin: data cleanup
    MsgBox "Detecting duplicates..."
    ' to do
    ' End: data cleanup
End Sub

Function LoadNationalRoster()
    LoadNationalRoster = LoadExcelFileToNewSheet("National")
End Function

Function LoadClubRoster()
    LoadClubRoster = LoadExcelFileToNewSheet("Club")
End Function

Function LoadExcelFileToNewSheet(rosterName As String)
    Dim filePath As String
    Dim sourceWB As Workbook
    Dim sourceWS As Worksheet
    Dim targetWS As Worksheet

    ' Prompt user to select an Excel file
    filePath = Application.GetOpenFilename("Excel Files (*.xlsx; *.xls), *.xlsx; *.xls", , "Select the " & rosterName & " Excel File")

    ' Exit if user cancels
    If filePath = "False" Then Exit Function

    ' Open the selected workbook
    Set sourceWB = Workbooks.Open(filePath)
    Set sourceWS = sourceWB.Sheets(1) ' You can change this to a specific sheet name or index

    ' Add a new worksheet to the current workbook
    Set targetWS = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(1))
    targetWS.Name = rosterName & "_" & Format(Now, "yyyymmdd") & "_" & Format(Now, "hhmmss")

    ' Copy data from source worksheet to target worksheet
    sourceWS.UsedRange.Copy Destination:=targetWS.Range("A1")

    ' Close the source workbook without saving changes
    sourceWB.Close SaveChanges:=False

    MsgBox "File imported successfully into sheet: " & targetWS.Name
    
    LoadExcelFileToNewSheet = targetWS.Name
End Function

