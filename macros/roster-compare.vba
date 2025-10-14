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
    
    maxNationalRow = LastRowWithDataInColumn(nationalWorksheet, N_UniqueContactId)
    maxClubRow = LastRowWithDataInColumn(clubWorksheet, C_MemberNumber)
    
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
