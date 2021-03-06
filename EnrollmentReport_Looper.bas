Attribute VB_Name = "EnrollmentReport_Looper"
Sub LOOPER()

'04/26/19 - Automatically hide empty Created Date in In Progress
'12/07/18 - Updated SPED column names.
'12/01/18 - Adapted looper after Pathways SSR update. Transitioned from A to BDev.
'11/15/18 - Finalized for export and team distribution.
'10/12/18 - Completed localization mechanism. Generalized to main reporting use.
'08/30/18 - Completed working version of code to parse Gio Friday Reports
'07/29/2018 - Created duplicate from Loop_DIY to insert AD autoprocessing after weekly reporting on Friday. OBJ: loop through files in 'working' folder and generate new sheet with all AD entries.

Dim MyFolder As String 'Path collected from the folder picker dialog
Dim MyFile As String 'Filename obtained by DIR function
Dim wbk As Workbook 'Used to loop through each workbook
On Error Resume Next

Application.ScreenUpdating = False

'Opens the folder picker dialog to allow user selection

With Application.FileDialog(msoFileDialogFolderPicker)
.Title = "Please select a folder"
.Show
.AllowMultiSelect = False
    If .SelectedItems.Count = 0 Then    'If no folder is selected, abort
    MsgBox "You did not select a folder"
    Exit Sub
   End If
   
MyFolder = .SelectedItems(1) & "\" 'Assign selected folder to MyFolder
MyFile = Dir(MyFolder) 'DIR gets the first file of the folder
'Loop through all files in a folder until DIR cannot find anymore
End With

Do While MyFile <> ""
   'Opens the file and assigns to the wbk variable for future use
   Set wbk = Workbooks.Open(Filename:=MyFolder & MyFile)


'INSERT DESIRED CODE -------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------------------

'Backing up before parse ---------------
    Cells.Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Sheets(2).Select
    Sheets(2).Name = "backup"

'D - parse ---------------------
    Sheets(1).Select
    'Deletes every column unless it is listed in the keep columns string
    'Add or delete column in KeepCols as desired
    i = 1
    KeepCols = "students_lastname, students_firstname, students_local_id, status, gradelevel, zip, county, invited, created, imported, latest_note, first_completed_date, first_completed_time, current_completed_date, current_completed_time, special_ed, 504plandocuments, iepdocuments, iep504documents, 4iepdocuments,7504documents,7iepdocuments,6504plandocuments,7iep504documents"
RetestCol_1:  'Test replacement column
    'Checks to see if column is one of the columns to keep or not
    Do While Not Cells(1, i) = ""
        check = InStr(1, KeepCols, Cells(1, i).Value)
        If (check = 0) Then
            Columns(i).EntireColumn.Delete
            GoTo RetestCol_1
        End If
    i = i + 1
    Loop
    Range("A1").Select
    
'In Progress ====================================================
    
    'Returns to Sheet 1
    Sheets(1).Select
        If (ActiveSheet.AutoFilterMode And ActiveSheet.FilterMode) Or ActiveSheet.FilterMode Then
        ActiveSheet.ShowAllData
        End If
    'ActiveSheet.ShowAllData
    Range("A1").Select
    
    LRImport = Cells(Rows.Count, 1).End(xlUp).Row  'finds last row index
    LC = Cells(1, Columns.Count).End(xlToLeft).Column 'finds last column index
    LRPending = Cells(Rows.Count, 1).End(xlUp).Row  'finds last row index
    'Finds column indices for status, gradelevel, and students_local_id
    For i = 1 To LC
        If Cells(1, i).Value = "status" Then StatusIndex = i
        If Cells(1, i).Value = "gradelevel" Then GradeLevelIndex = i
        If Cells(1, i).Value = "students_local_id" Then LocalIDIndex = i
    Next i
    'Filters by status (Awaiting Data and Awaiting Import & Missing Parent)
    ActiveSheet.Range("$A$1:$S$15000").AutoFilter Field:=StatusIndex, Criteria1:=Array("=Awaiting Data", "=Awaiting Import", "=Missing Parent"), Operator:=xlFilterValues
    'Filters by blank students_local_id's
    ActiveSheet.Range("$A$1:$S$15000").AutoFilter Field:=LocalIDIndex, Criteria1:="="
    'Copies in progress students to new sheet
    Cells.Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    Sheets(2).Select
    Sheets(2).Name = "in progress"
    ActiveSheet.Paste
    Columns(StatusIndex).EntireColumn.AutoFit
    Cells(1, StatusIndex).Select
    Application.CutCopyMode = False
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("in progress").AutoFilter.Sort.SortFields.Clear
    
    'Finds status column and sorts by ascending status (Awaiting Data then Awaiting Import)
    ActiveWorkbook.Worksheets("in progress").AutoFilter.Sort.SortFields.Add Key:= _
        Range(Cells(2, StatusIndex), Cells(LRPending, StatusIndex)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    
    'Sorts by Gradelevel
    ActiveWorkbook.Worksheets("in progress").AutoFilter.Sort.SortFields.Add Key:= _
        Range(Cells(2, GradeLevelIndex), Cells(LRImport, GradeLevelIndex)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("in progress").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Hides blank Created Date
    ActiveSheet.Range("$A$1:$R$50000").AutoFilter Field:=9, Criteria1:="<>"
    
    'Clears filter from sheet 1
    Sheets(1).Select
    ActiveSheet.ShowAllData
    Range("A1").Select
    
'IMPORT AUDIT ===================================================================

'create import audit sheet
    Sheets("backup").Select
    Sheets("backup").Copy After:=Sheets(3)
    Sheets("backup (2)").Select
    Sheets("backup (2)").Name = "import-audit"

'create match formula and fill in until end of doc
    'Deletes every column unless it is listed in the keep columns string
    'Add or delete column in KeepCols as desired
    i = 1
    KeepCols = "students_lastname, students_firstname, students_local_id, status, gradelevel, zip, county, invited, created, imported, first_completed_date, first_completed_time, current_completed_date, current_completed_time, special_ed, 504plandocuments, iepdocuments, iep504documents, 4iepdocuments,7504documents,7iepdocuments,6504plandocuments,7iep504documents"
RetestCol_2: 'Test replacement column
    'Checks to see if column is one of the columns to keep or not
    Do While Not Cells(1, i) = ""
        check = InStr(1, KeepCols, Cells(1, i).Value)
        If (check = 0) Then
            Columns(i).EntireColumn.Delete
            GoTo RetestCol_2
        End If
    i = i + 1
    Loop

    LC = Cells(1, Columns.Count).End(xlToLeft).Column 'finds last column index
    NC = LC + 1 'finds new column index
    LRAudit = Cells(Rows.Count, 1).End(xlUp).Row 'finds last row index
    'Finds column indices to perform match and for sorting purposes
    For i = 1 To LC
        If Cells(1, i) = "first_completed_date" Then FirstDateIndex = i
        If Cells(1, i).Value = "current_completed_date" Then CurrentDateIndex = i
        If Cells(1, i).Value = "imported" Then ImportedIndex = i
        If Cells(1, i).Value = "gradelevel" Then GradeLevelIndex = i
        If Cells(1, i).Value = "created" Then CreatedIndex = i
        If Cells(1, i).Value = "county" Then CreatedIndex = i
    Next i
    Cells(1, NC).Value = "Match" 'set header for new column
    i = 2
    'Checks for match using first_completed_date and and current_completed_date
    Do While Cells(i, 1) <> ""
        If Cells(i, FirstDateIndex).Value = Cells(i, CurrentDateIndex).Value Then
            Cells(i, NC).Value = 1
        Else: Cells(i, NC).Value = False
        End If
        i = i + 1
    Loop
    Cells(1, NC).Select

'parses for easy counting - irrelevant columns hidden
    Range("A1:R20000").Select
    Selection.AutoFilter
    Cells.Select
    Cells.EntireColumn.AutoFit
    Selection.ColumnWidth = 14.29
    'Filters by latest weeks imports
    ActiveSheet.Range("$A$1:$R$20000").AutoFilter Field:=ImportedIndex, Criteria1:= _
        xlFilterThisWeek, Operator:=xlFilterDynamic
    ActiveWorkbook.Worksheets("import-audit").AutoFilter.Sort.SortFields.Clear
    'Sorts by gradelevel
    ActiveWorkbook.Worksheets("import-audit").AutoFilter.Sort.SortFields.Add Key _
        :=Range(Cells(2, GradeLevelIndex), Cells(LRAudit, GradeLevelIndex)), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("import-audit").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

'Hides the columns not listed in KeepCols - disabled 04/18/19
  '  i = 1 ' counter step
    'Hides every column unless it is listed in the keep columns string
    'Add or delete column in KeepCols as desired
        'KeepCols = "students_lastname, students_firstname, students_local_id, status,zip, county, gradelevel, Match"
    'Checks to see if column is one of the columns to keep or not
    'Do While Not Cells(1, i) = ""
       ' check = InStr(1, KeepCols, Cells(1, i).Value)
       ' If (check = 0) Then
       '     Columns(i).EntireColumn.Hidden = True
       ' End If
   ' i = i + 1
   ' Loop
   ' Range("A1").Select

'----------------------------------------- OWN CODE COMPLETED ----------------------------------------------------

'Done with DIY, save file as Excel File
ActiveWorkbook.SaveAs Filename:=Left((MyFolder & MyFile), Len(MyFolder & MyFile) - 4) & ".xlsx", FileFormat:=51
wbk.Close savechanges:=True

MyFile = Dir 'DIR gets the next file in the folder
Loop

Application.ScreenUpdating = True
MsgBox "Enrollment Report Looper Automation COMPLETE. This was fun!"

End Sub

