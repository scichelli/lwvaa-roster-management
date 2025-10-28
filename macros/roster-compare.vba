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
Const I_MissingFromOtherRoster As String = "Missing From Other Roster"

Sub RunSynchronization()
    Dim nationalWorksheet As Worksheet
    Dim clubWorksheet As Worksheet
    Dim maxNationalRow, maxClubRow As Long
    
    ' Begin: load rosters into worksheets
    MsgBox "First we'll load the National roster into a worksheet"
    Set nationalWorksheet = LoadNationalRoster()
    
    MsgBox "Next we'll load the Club roster into a worksheet"
    Set clubWorksheet = LoadClubRoster()
    
    If nationalWorksheet Is Nothing Or clubWorksheet Is Nothing Then
        MsgBox "Did not find two worksheets to compare"
        Exit Sub
    End If
    ' End: load rosters into worksheets
    
    ' Begin: identify data shape
    maxNationalRow = LastRowWithDataInColumn(nationalWorksheet, N_UniqueContactId)
    maxClubRow = LastRowWithDataInColumn(clubWorksheet, C_MemberNumber)
    
    If maxNationalRow = 0 Or maxClubRow = 0 Then
        MsgBox "Did not find data to compare"
        Exit Sub
    End If
    
    ' TODO: Verify required columns are present, emit an error sheet if not

    ' End: identify data shape
    
    ' Begin: identify duplicates
    SortByName nationalWorksheet, maxNationalRow, N_FirstName, N_LastName
    SortByName clubWorksheet, maxClubRow, C_FirstName, C_LastName
    
    HighlightDuplicateNames nationalWorksheet, maxNationalRow
    HighlightDuplicateNames clubWorksheet, maxClubRow
    ' End: identify duplicates
    
    ' Begin: identify missing from other roster
    HighlightNamesInFirstSheetMissingFromSecondSheet nationalWorksheet, maxNationalRow, clubWorksheet, maxClubRow
    HighlightNamesInFirstSheetMissingFromSecondSheet clubWorksheet, maxClubRow, nationalWorksheet, maxNationalRow
    ' End: identify missing from other roster
    
    ApplyHeaderRow nationalWorksheet
    ApplyHeaderRow clubWorksheet
    
    ' Begin: discrepancy report
    ' BuildDiscrepancyReport nationalWorksheet, clubWorksheet
    BuildSideBySideReport nationalWorksheet, clubWorksheet
    ' End: discrepancy report
End Sub

Sub StartDiscrepancyReport()
    ' Useful for testing, and for re-generating the report after modifying the roster worksheets
    ' If you've already loaded sheets and allowed the macro to prep them, a button connected to this macro can generate the report without repeating the prep.
    ' Use B11 and B12 for worksheet names
    
    Dim controlWS As Worksheet
    Dim nationalWorksheet As Worksheet
    Dim clubWorksheet As Worksheet
    Dim nationalWsName, clubWsName As String

    Set controlWS = ThisWorkbook.Sheets(1)
    nationalWsName = controlWS.Cells(11, 2).Value
    clubWsName = controlWS.Cells(12, 2).Value
    
    Set nationalWorksheet = ThisWorkbook.Sheets(nationalWsName)
    Set clubWorksheet = ThisWorkbook.Sheets(clubWsName)

    ' BuildDiscrepancyReport nationalWorksheet, clubWorksheet
    BuildSideBySideReport nationalWorksheet, clubWorksheet

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
    
    ' Define the range to apply formatting (entire rows from A to max column)
    Dim dataRange As Range
    Set dataRange = ws.Range("A2:" & dupFN & maxRow) ' A2 because headers are in row 1; dupFN because Has Duplicate Full Name is now the right-most column

    ' Orange formatting (Column Has Duplicate Full Name = TRUE) -- higher priority
    With dataRange.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=$" & dupFN & "2=TRUE")
        .Interior.Color = RGB(255, 204, 204) ' Orange
        .StopIfTrue = False
    End With

    ' Gray formatting (Column Has Duplicate Last Name = TRUE)
    With dataRange.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=$" & dupLN & "2=TRUE")
        .Interior.Color = RGB(221, 221, 221) ' Light gray
        .StopIfTrue = False
    End With
End Sub

Sub HighlightNamesInFirstSheetMissingFromSecondSheet(ByRef ws1 As Worksheet, ByVal maxRow1 As Long, ByRef ws2 As Worksheet, ByVal maxRow2 As Long)
    ' add column to ws1 for names missing from ws2
    ' add conditional formatting, red if missing

    Dim maxColumn, missingNameColumn As Long
    Dim mn As String ' Missing Name column letter
    Dim fn1, fn2 As String ' Full Name column letter from sheet 1 and sheet 2
    maxColumn = LastColumnWithData(ws1)
    missingNameColumn = maxColumn + 1
    fn1 = FindColumnLetterByName(ws1, I_CombinedName)
    fn2 = FindColumnLetterByName(ws2, I_CombinedName)
    mn = ColumnNumberToLetter(missingNameColumn)

    ws1.Cells(1, missingNameColumn).Value = I_MissingFromOtherRoster

    ' add boolean column for missing names
    For i = 2 To maxRow1
        ' Find-missing formula example:
        ' =ISNA(MATCH($B2, National!$G$2:$G$1000, 0))
        ws1.Cells(i, missingNameColumn).Formula = "=ISNA(MATCH($" & fn1 & i & ", " & ws2.Name & "!$" & fn2 & "$2:$" & fn2 & "$" & maxRow2 & ", 0))"
    Next i
    
    ' add conditional format rule to highlight rows with missing
    
    ' Define the range to apply formatting (entire rows from A to max column)
    Dim dataRange As Range
    Set dataRange = ws1.Range("A2:" & mn & maxRow1) ' A2 because headers are in row 1; mn because Missing Name is now the right-most column

    ' Red formatting (Column Missing Name = TRUE)
    With dataRange.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=$" & mn & "2=TRUE")
        .Interior.Color = RGB(255, 102, 102) ' Red
        .StopIfTrue = False
    End With
End Sub

Sub BuildSideBySideReport(ByRef nationalWS As Worksheet, ByRef clubWS As Worksheet)
    ' In a new worksheet, list each club row next to its matching national row if found, then list national rows that have no matching club
    Dim reportWS As Worksheet
    Dim lastRowNational As Long, lastRowClub As Long, outputRow As Long
    Dim lastColumnNational As Long, lastColumnClub As Long
    Dim nationalName As String, clubName As String
    Dim i As Long, j As Long
    Dim nDict As Object, cDict As Object
    Dim key As Variant
    
    ' Get column letters
    Set nDict = CreateObject("Scripting.Dictionary")
    Set cDict = CreateObject("Scripting.Dictionary")
    PopulateColumnsDictionaries nDict, nationalWS, cDict, clubWS
    
    Set reportWS = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(1))
    reportWS.Name = "SideBySide_" & Format(Now, "yyyymmdd") & "_" & Format(Now, "hhmmss")
    
    ' Get last rows and columns
    lastRowNational = LastRowWithDataInColumn(nationalWS, I_CombinedName)
    lastRowClub = LastRowWithDataInColumn(clubWS, I_CombinedName)
    lastColumnNational = LastColumnWithData(nationalWS)
    lastColumnClub = LastColumnWithData(clubWS)
    
    outputRow = 1
    
    ' Add headings
    With reportWS.Range(reportWS.Cells(outputRow, 1), reportWS.Cells(outputRow, lastColumnClub))
        .Merge
        .Value = "Club Roster " & Format(Now, "yyyymmdd")
        .Font.Bold = True
        .Font.Size = 14
    End With
    
    With reportWS.Range(reportWS.Cells(outputRow, lastColumnClub + 1), reportWS.Cells(outputRow, (lastColumnClub + 1) + lastColumnNational))
        .Merge
        .Value = "National Roster " & Format(Now, "yyyymmdd")
        .Font.Bold = True
        .Font.Size = 14
    End With
    outputRow = outputRow + 1
    
    ' Althought it would be nice to use "For Each key in cDict.Keys", that makes the report brittle to the inclusion of new fields in the export
    i = 1
    For Each key In cDict.Keys
        reportWS.Cells(outputRow, i).Value = key
        i = i + 1
    Next key
    For Each key In nDict.Keys
        reportWS.Cells(outputRow, i).Value = key
        i = i + 1
    Next key
    outputRow = outputRow + 1
    
    reportWS.Cells(outputRow, 1).Value = "ready for rows"

End Sub

Sub BuildDiscrepancyReport(ByRef nationalWS As Worksheet, ByRef clubWS As Worksheet)
    ' For rows that are present in both rosters, list discrepancies in a new worksheet
    
    Dim discrepancyWS As Worksheet
    Dim lastRowNational As Long, lastRowClub As Long, outputRow As Long
    Dim nationalName As String, clubName As String
    Dim nationalExpiration As Variant, clubExpiration As Variant
    Dim nationalEmail As String, clubEmail As String
    Dim i As Long, j As Long
    Dim nDict As Object, cDict As Object
    
    ' Get column letters
    Set nDict = CreateObject("Scripting.Dictionary")
    Set cDict = CreateObject("Scripting.Dictionary")
    PopulateColumnsDictionaries nDict, nationalWS, cDict, clubWS
    
    Set discrepancyWS = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(1))
    discrepancyWS.Name = "Discrepancies_" & Format(Now, "yyyymmdd") & "_" & Format(Now, "hhmmss")
    
    ' Get last rows
    lastRowNational = nationalWS.Cells(nationalWS.Rows.Count, nDict(I_CombinedName)).End(xlUp).row
    lastRowClub = clubWS.Cells(clubWS.Rows.Count, cDict(I_CombinedName)).End(xlUp).row

    outputRow = 1 ' Start writing to row 1

    ' Loop through national list
    For i = 2 To lastRowNational ' starting with 2 because row 1 is header
        nationalName = Trim(nationalWS.Cells(i, nDict(I_CombinedName)).Value)
        nationalExpiration = nationalWS.Cells(i, nDict(N_ExpirationDate)).Value
        nationalEmail = Trim(nationalWS.Cells(i, nDict(N_Email)).Value)

        ' Search for matching name in club list
        For j = 2 To lastRowClub
            clubName = Trim(clubWS.Cells(j, cDict(I_CombinedName)).Value)
            If StrComp(nationalName, clubName, vbTextCompare) = 0 Then
                clubExpiration = clubWS.Cells(j, cDict(C_ExpirationDate)).Value
                clubEmail = Trim(clubWS.Cells(j, cDict(C_PrimaryEmail)).Value)

                ' Compare expiration date and email
                If nationalExpiration <> clubExpiration Or nationalEmail <> clubEmail Then
                    ' Add national row to discrepancies sheet
                    discrepancyWS.Cells(outputRow, 1).Value = "National"
                    For col = 1 To 50 ' 50 is a placeholder for the max number of columns to copy; the count is exceeded when adding 1 to it
                        discrepancyWS.Cells(outputRow, col + 1).Value = nationalWS.Cells(i, col).Value
                    Next col
                    outputRow = outputRow + 1
                    
                    ' Add club row to discrepancies sheet
                    discrepancyWS.Cells(outputRow, 1).Value = "Club"
                    For col = 1 To 50 ' 50 is a placeholder, see above
                        discrepancyWS.Cells(outputRow, col + 1).Value = clubWS.Cells(j, col).Value
                    Next col
                    outputRow = outputRow + 1
                End If

                Exit For
            End If
        Next j
    Next i
    
    BuildCoverSheet nationalWS, clubWS, discrepancyWS
End Sub

Sub BuildCoverSheet(ByRef nationalWS As Worksheet, ByRef clubWS As Worksheet, ByRef discrepancyWS As Worksheet)
    Dim coverWS As Worksheet
    Dim row As Integer
    Set coverWS = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(1))
    coverWS.Name = "Report_" & Format(Now, "yyyymmdd") & "_" & Format(Now, "hhmmss")
    row = 1

    ' Heading
    With coverWS.Range(coverWS.Cells(row, 1), coverWS.Cells(row, 3))
        .Merge
        .Value = "Member Roster Report " & Format(Now, "yyyymmdd")
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = RGB(200, 200, 255)
        .VerticalAlignment = xlCenter
    End With
    row = row + 2
    
    ' Specify which worksheets we used
    With coverWS.Cells(row, 1)
        .Value = "Source Worksheets:"
        .Font.Bold = True
    End With
    row = row + 1
    coverWS.Cells(row, 1).Value = "National Roster"
    coverWS.Cells(row, 2).Value = nationalWS.Name
    row = row + 1
    coverWS.Cells(row, 1).Value = "Club Roster"
    coverWS.Cells(row, 2).Value = clubWS.Name
    row = row + 1
    coverWS.Cells(row, 1).Value = "Discrepancy Report"
    coverWS.Cells(row, 2).Value = discrepancyWS.Name
    row = row + 1
    
    ' TODO
    
    ' Verify all required columns are present
    ' List count of duplicates from National
    ' List count of duplicates from Club
    ' List count of National that are missing from Club
    ' List count of Club that are missing from National
    ' List count of discrepancies
    
    coverWS.Columns("A").AutoFit
End Sub

Function FindColumnLetterByName(ByRef ws As Worksheet, ByVal columnName As String) As String
    Dim columnNumber As Long
    columnNumber = FindColumnNumberByName(ws, columnName)
    If columnNumber = 0 Then
        Exit Function
    End If
    FindColumnLetterByName = ColumnNumberToLetter(columnNumber)
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
    
    LastRowWithDataInColumn = ws.Cells(ws.Rows.Count, columnIndex).End(xlUp).row ' Find last used row in specified column
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

Sub PopulateColumnsDictionaries(ByRef nDict As Object, ByRef nationalWS As Worksheet, ByRef cDict As Object, ByRef clubWS As Worksheet)
    nDict.Add N_FirstName, FindColumnLetterByName(nationalWS, N_FirstName)
    nDict.Add N_PreferredFirstName, FindColumnLetterByName(nationalWS, N_PreferredFirstName)
    nDict.Add N_LastName, FindColumnLetterByName(nationalWS, N_LastName)
    nDict.Add N_Phone, FindColumnLetterByName(nationalWS, N_Phone)
    nDict.Add N_Email, FindColumnLetterByName(nationalWS, N_Email)
    nDict.Add N_MailingStreet, FindColumnLetterByName(nationalWS, N_MailingStreet)
    nDict.Add N_MailingCity, FindColumnLetterByName(nationalWS, N_MailingCity)
    nDict.Add N_MailingState, FindColumnLetterByName(nationalWS, N_MailingState)
    nDict.Add N_MailingPostalCode, FindColumnLetterByName(nationalWS, N_MailingPostalCode)
    nDict.Add N_MailingCountry, FindColumnLetterByName(nationalWS, N_MailingCountry)
    nDict.Add N_ExpirationDate, FindColumnLetterByName(nationalWS, N_ExpirationDate)
    nDict.Add N_LastLoginDate, FindColumnLetterByName(nationalWS, N_LastLoginDate)
    nDict.Add N_UniqueContactId, FindColumnLetterByName(nationalWS, N_UniqueContactId)
    nDict.Add N_UniqueAccountId, FindColumnLetterByName(nationalWS, N_UniqueAccountId)
    nDict.Add I_SortableLastName, FindColumnLetterByName(nationalWS, I_SortableLastName)
    nDict.Add I_CombinedName, FindColumnLetterByName(nationalWS, I_CombinedName)
    nDict.Add I_DuplicateLastName, FindColumnLetterByName(nationalWS, I_DuplicateLastName)
    nDict.Add I_DuplicateCombinedName, FindColumnLetterByName(nationalWS, I_DuplicateCombinedName)
    nDict.Add I_MissingFromOtherRoster, FindColumnLetterByName(nationalWS, I_MissingFromOtherRoster)

    cDict.Add C_MemberNumber, FindColumnLetterByName(clubWS, C_MemberNumber)
    cDict.Add C_LastName, FindColumnLetterByName(clubWS, C_LastName)
    cDict.Add C_FirstName, FindColumnLetterByName(clubWS, C_FirstName)
    cDict.Add C_LoginName, FindColumnLetterByName(clubWS, C_LoginName)
    cDict.Add C_Address1, FindColumnLetterByName(clubWS, C_Address1)
    cDict.Add C_Address2, FindColumnLetterByName(clubWS, C_Address2)
    cDict.Add C_City, FindColumnLetterByName(clubWS, C_City)
    cDict.Add C_State, FindColumnLetterByName(clubWS, C_State)
    cDict.Add C_Zip, FindColumnLetterByName(clubWS, C_Zip)
    cDict.Add C_Phone, FindColumnLetterByName(clubWS, C_Phone)
    cDict.Add C_CellPhone, FindColumnLetterByName(clubWS, C_CellPhone)
    cDict.Add C_PrimaryEmail, FindColumnLetterByName(clubWS, C_PrimaryEmail)
    cDict.Add C_MemberType_name, FindColumnLetterByName(clubWS, C_MemberType_name)
    cDict.Add C_Level, FindColumnLetterByName(clubWS, C_Level)
    cDict.Add C_ExpirationDate, FindColumnLetterByName(clubWS, C_ExpirationDate)
    cDict.Add I_SortableLastName, FindColumnLetterByName(clubWS, I_SortableLastName)
    cDict.Add I_CombinedName, FindColumnLetterByName(clubWS, I_CombinedName)
    cDict.Add I_DuplicateLastName, FindColumnLetterByName(clubWS, I_DuplicateLastName)
    cDict.Add I_DuplicateCombinedName, FindColumnLetterByName(clubWS, I_DuplicateCombinedName)
    cDict.Add I_MissingFromOtherRoster, FindColumnLetterByName(clubWS, I_MissingFromOtherRoster)
End Sub
