Sub RunSynchronization()
    Dim nationalWorksheet As Worksheet
    Dim clubWorksheet As Worksheet
    Dim maxNationalRow, maxClubRow As Long
    
    ' Begin: load rosters into worksheets
    MsgBox "First we'll load the National roster into a worksheet"
    Set nationalWorksheet = LoadNationalRoster()
    
    MsgBox "Next we'll load the Club roster into a worksheet"
    Set clubWorksheet = LoadClubRoster()
    ' End: load rosters into worksheets
    
    If nationalWorksheet Is Nothing Or clubWorksheet Is Nothing Then
        MsgBox "Did not find two worksheets to compare"
        Exit Sub
    End If
    
    ' Begin: prep sheets
    ApplyHeaderRow nationalWorksheet
    ApplyHeaderRow clubWorksheet
    
    maxNationalRow = LastRowWithDataInColumn(nationalWorksheet, "Unique Contact Id")
    maxClubRow = LastRowWithDataInColumn(clubWorksheet, "member_number")
    
    If maxNationalRow = 0 Or maxClubRow = 0 Then
        MsgBox "Did not find data to compare"
        Exit Sub
    End If
    
    ' End: prep sheets
    
    ' Begin: data cleanup
    MsgBox "Detecting duplicates..."
    ' to do
    ' End: data cleanup
End Sub

Function LoadNationalRoster() As Worksheet
    Set LoadNationalRoster = LoadExcelFileToNewSheet("National")
End Function

Function LoadClubRoster() As Worksheet
    Set LoadClubRoster = LoadExcelFileToNewSheet("Club")
End Function

Function LoadExcelFileToNewSheet(rosterName As String) As Worksheet
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
    
    Set LoadExcelFileToNewSheet = targetWS
End Function

Sub ApplyHeaderRow(ByRef ws As Worksheet)
    ' Clear any existing filters
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    
    ' Apply AutoFilter to the first row
    ws.Rows(1).AutoFilter
End Sub

Function FindColumnByName(ByRef ws As Worksheet, ByVal columnName As String) As Long
    Dim foundCell As Range
    Set foundCell = ws.Rows(1).Find(What:=columnName, LookIn:=xlValues, LookAt:=xlWhole)
    
    If foundCell Is Nothing Then
        MsgBox "Could not find column " & columnName & " in worksheet " & ws.Name
        Exit Function
    End If
    
    FindColumnByName = foundCell.column
End Function

Function LastRowWithDataInColumn(ByRef ws As Worksheet, ByVal columnName As String) As Long
    Dim columnIndex As Long
    columnIndex = FindColumnByName(ws, columnName)
    
    If columnIndex = 0 Then
        Exit Function
    End If
    
    LastRowWithDataInColumn = ws.Cells(ws.Rows.Count, columnIndex).End(xlUp).Row ' Find last used row in specified column
End Function
