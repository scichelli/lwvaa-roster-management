' For column header constants,
'   If the roster starts to include new columns, no code change is needed, unless those columns could be useful in the logic.
'   If the roster stops including a column that is used in the logic, then a code change will be needed.
'   If the column gets a new name, then change the value in quotes, in the Const declarations below, to use the new name. The code will keep using the constant name.
'      For example, if "First Name" becomes "F Name", edit only the constant declaration at the start of this file:
'      Const N_FirstName As String = "First Name"
'      becomes
'      Const N_FirstName As String = "F Name"
'      The code can continue to refer to N_FirstName everywhere else.

' National roster column header names
Const N_FirstName As String = "First Name"
Const N_PreferredFirstName As String = "Preferred First Name"
Const N_MiddleName As String = "Middle Name"
Const N_LastName As String = "Last Name"
Const N_Phone As String = "Phone"
Const N_Email As String = "Email"
Const N_MailingStreet As String = "Mailing Street"
Const N_MailingCity As String = "Mailing City"
Const N_MailingState As String = "Mailing State"
Const N_MailingPostalCode As String = "Mailing Postal Code"
Const N_MailingCountry As String = "Mailing Country"
Const N_ExpirationDate As String = "Expiration Date"
Const N_LastLoginDate As String = "Last Login Date"
Const N_CurrentStatus As String = "Current Status"
Const N_UniqueContactId As String = "Unique Contact Id"
Const N_UniqueAccountId As String = "Unique Account Id"

' Club roster column header names
Const C_MemberNumber As String = "member_number"
Const C_LastName As String = "last_name"
Const C_FirstName As String = "first_name"
Const C_LoginName As String = "login_name"
Const C_Address1 As String = "address1"
Const C_Address2 As String = "address2"
Const C_City As String = "city"
Const C_State As String = "state"
Const C_Zip As String = "zip"
Const C_Phone As String = "phone"
Const C_CellPhone As String = "cell_phone"
Const C_PrimaryEmail As String = "primary_email"
Const C_MemberType_name As String = "member_type_name"
Const C_Level As String = "level"
Const C_ExpirationDate As String = "expiration_date"

' Internal operations column header names
Const I_SortableLastName As String = "Sortable Last Name"
Const I_CombinedName As String = "Combined Name"
Const I_DuplicateLastName As String = "Has Duplicate Last Name"
Const I_DuplicateCombinedName As String = "Has Duplicate Full Name"

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
    maxNationalRow = LastRowWithDataInColumn(nationalWorksheet, N_UniqueContactId)
    maxClubRow = LastRowWithDataInColumn(clubWorksheet, C_MemberNumber)
    
    If maxNationalRow = 0 Or maxClubRow = 0 Then
        MsgBox "Did not find data to compare"
        Exit Sub
    End If
    
    ' End: prep sheets
    
    ' Begin: identify duplicates
    SortByName nationalWorksheet, maxNationalRow, N_FirstName, N_LastName
    SortByName clubWorksheet, maxClubRow, C_FirstName, C_LastName
    
    HighlightDuplicateNames nationalWorksheet, maxNationalRow
    HighlightDuplicateNames clubWorksheet, maxClubRow
    
    ApplyHeaderRow nationalWorksheet
    ApplyHeaderRow clubWorksheet
    
    ' End: identify duplicates
    
    ' Begin: discrepancy report
    BuildDiscrepancyReport nationalWorksheet, clubWorksheet
    ' End: discrepancy report
End Sub

Sub StartDiscrepancyReport()
    ' Useful for testing, and for re-generating the report after modifying the roster worksheets
    ' If you've already loaded sheets and allowed the macro to prep them, a button connected to this macro can generate the report without repeating the prep.
    ' Use B11 and B12 for worksheet names
    
    Dim controlWS As Worksheet
    Dim nationalWS As Worksheet
    Dim clubWS As Worksheet
    Dim nationalWsName, clubWsName As String
    
    Set controlWS = ThisWorkbook.Sheets(1)
    nationalWsName = controlWS.Cells(11, 2).Value
    clubWsName = controlWS.Cells(12, 2).Value
    
    Set nationalWS = ThisWorkbook.Sheets(nationalWsName)
    Set clubWS = ThisWorkbook.Sheets(clubWsName)
    
    BuildDiscrepancyReport nationalWS, clubWS
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

Sub SortByName(ByRef ws As Worksheet, ByVal maxRow As Long, ByVal firstNameColumnName As String, ByVal lastNameColumnName As String)
    ' add column for last name, lower-cased, removing non-alphanumeric
    ' add column for last name + first name, lower-cased, removing non-alphanumeric
    ' sort by the last name column
    
    Dim maxColumn, sortableLastNameColumn, combinedNameColumn As Long
    Dim firstNameColumn, lastNameColumn As String
    maxColumn = LastColumnWithData(ws)
    sortableLastNameColumn = maxColumn + 1
    combinedNameColumn = maxColumn + 2
    firstNameColumn = FindColumnLetterByName(ws, firstNameColumnName)
    lastNameColumn = FindColumnLetterByName(ws, lastNameColumnName)
    
    ws.Cells(1, sortableLastNameColumn).Value = I_SortableLastName
    ws.Cells(1, combinedNameColumn).Value = I_CombinedName
    
    For i = 2 To maxRow
        ' Sortable Last Name is last name, lowercased, with all numbers and punctuation removed
        ws.Cells(i, sortableLastNameColumn).Formula = "=LowercaseLettersOnly(" & lastNameColumn & i & ")"
        ' Combined Name is last name and first name, to help with comparison
        ws.Cells(i, combinedNameColumn).Formula = "=CONCATENATE(LowercaseLettersOnly(" & lastNameColumn & i & "), LowercaseLettersOnly(" & firstNameColumn & i & "))"
    Next i

    ' Sort by the sortable last name column
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range(ColumnNumberToLetter(sortableLastNameColumn) & 2), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            ' ws.Range("M2"), if "M" is the sortable column, and "2" because row 1 is a header
        .SetRange ws.UsedRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

Sub HighlightDuplicateNames(ByRef ws As Worksheet, ByVal maxRow As Long)
    ' add column for duplicate SortableLastName
    ' add column for duplicate CombinedName
    ' add conditional formatting, gray if duplicate last name
    ' add conditional formatting, orange if duplicate combined name
    
    Dim maxColumn, hasDuplicateLastNameColumn, hasDuplicateFullNameColumn As Long
    Dim dupLN, dupFN As String ' Has Duplicate Last Name column letter and Has Duplicate Full Name column letter
    Dim ln, fn As String ' Last Name column letter and Full Name column letter
    maxColumn = LastColumnWithData(ws)
    hasDuplicateLastNameColumn = maxColumn + 1
    hasDuplicateFullNameColumn = maxColumn + 2
    ln = FindColumnLetterByName(ws, I_SortableLastName)
    fn = FindColumnLetterByName(ws, I_CombinedName)
    dupLN = ColumnNumberToLetter(hasDuplicateLastNameColumn)
    dupFN = ColumnNumberToLetter(hasDuplicateFullNameColumn)

    ws.Cells(1, hasDuplicateLastNameColumn).Value = I_DuplicateLastName
    ws.Cells(1, hasDuplicateFullNameColumn).Value = I_DuplicateCombinedName

    ' add boolean columns for detected duplicates
    For i = 2 To maxRow
        ' Find-duplicates formula example:
        ' =COUNTIF(M:M, M2) > 1
        ws.Cells(i, hasDuplicateLastNameColumn).Formula = "=COUNTIF(" & ln & ":" & ln & ", " & ln & i & ") > 1"
        ws.Cells(i, hasDuplicateFullNameColumn).Formula = "=COUNTIF(" & fn & ":" & fn & ", " & fn & i & ") > 1"
    Next i
    
    ' add conditional format rules to highlight rows with duplicates
    
    ' Clear existing conditional format rules
    ws.Cells.FormatConditions.Delete
    ' Define the range to apply formatting (entire rows from A to max column)
    Dim dataRange As Range
    Set dataRange = ws.Range("A2:" & dupFN & maxRow) ' A2 because headers are in row 1; dupFN because Has Duplicate Full Name is now the right-most column

    ' Orange formatting (Column Has Duplicate Full Name = TRUE) -- higher priority
    With dataRange.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=$" & dupFN & "2=TRUE")
        .Interior.Color = RGB(255, 192, 0) ' Orange
        .StopIfTrue = False
    End With

    ' Gray formatting (Column Has Duplicate Last Name = TRUE)
    With dataRange.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=$" & dupLN & "2=TRUE")
        .Interior.Color = RGB(200, 200, 200) ' Light gray
        .StopIfTrue = False
    End With
End Sub

Sub BuildDiscrepancyReport(ByRef nationalWS As Worksheet, ByRef clubWS As Worksheet)
    MsgBox "Found " & nationalWS.Name & " and " & clubWS.Name
End Sub

Function FindColumnLetterByName(ByRef ws As Worksheet, ByVal columnName As String) As String
    FindColumnLetterByName = ColumnNumberToLetter(FindColumnNumberByName(ws, columnName))
End Function

Function FindColumnNumberByName(ByRef ws As Worksheet, ByVal columnName As String) As Long
    Dim foundCell As Range
    Set foundCell = ws.Rows(1).Find(What:=columnName, LookIn:=xlValues, LookAt:=xlWhole)
    
    If foundCell Is Nothing Then
        MsgBox "Could not find column " & columnName & " in worksheet " & ws.Name
        Exit Function
    End If
    
    FindColumnNumberByName = foundCell.column
End Function

Function ColumnNumberToLetter(ByVal colNum As Long) As String
    ColumnNumberToLetter = Split(Cells(1, colNum).Address(True, False), "$")(0)
End Function

Function LastRowWithDataInColumn(ByRef ws As Worksheet, ByVal columnName As String) As Long
    Dim columnIndex As Long
    columnIndex = FindColumnNumberByName(ws, columnName)
    
    If columnIndex = 0 Then
        Exit Function
    End If
    
    LastRowWithDataInColumn = ws.Cells(ws.Rows.Count, columnIndex).End(xlUp).Row ' Find last used row in specified column
End Function

Function LastColumnWithData(ByRef ws As Worksheet) As Long
    LastColumnWithData = ws.Cells(1, ws.Columns.Count).End(xlToLeft).column
End Function

Function LowercaseLettersOnly(ByVal txt As String) As String
    Dim i As Integer
    Dim ch As String
    Dim result As String

    txt = LCase(txt)
    For i = 1 To Len(txt)
        ch = Mid(txt, i, 1)
        If ch >= "a" And ch <= "z" Then
            result = result & ch
        End If
    Next i

    LowercaseLettersOnly = result
End Function
