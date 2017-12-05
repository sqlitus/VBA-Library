Attribute VB_Name = "HomeLibrary"
Option Compare Text


Sub CiscoReport()


    ' find table start
    Dim tableStart As Range
    
    If Range("a1").Value <> "" Then
        Set tableStart = Range("a1")
    Else
        Set tableStart = Range("a1").End(xlDown)
    End If
    
    tableStart.Select
    
    
    ' select table range, columns within table, count rows & columns in table
    Dim myTable
    
    Set myTable = Range(tableStart, tableStart.End(xlToRight).End(xlDown))
    myTable.Select
    myTable.Columns(1).Select
    Range("A10").Value = myTable.Columns.Count
    
    
    ' loop through column headers of table range & replace string
    For i = 1 To myTable.Columns.Count
        myTable.Columns(i).Rows(1).Value = myTable.Columns(i).Rows(1).Value & " test"
        myTable.Columns(i).Rows(1).Value = Replace(myTable.Columns(i).Rows(1).Value, "dequeue", "voicemail")
    Next i
    
    
    ' replace string in column values
    For i = 1 To myTable.Rows.Count
        myTable.Columns(1).Rows(i).Value = Replace(myTable.Columns(1).Rows(i).Value, "opos_", "")
    Next i
    
    
    ' find & select column by name
    myTable.Rows(1).Find("abandoned").Select
    Range(myTable.Rows(1).Find("abandoned"), myTable.Rows(1).Find("abandoned").End(xlDown)).Select
    
    ' NOT WORKING: isolate cell by row (of value found) & column (by header found)
    myTable.Rows(1).Select
    myTable.Columns(1).Select
    
    myTable.Cells(myTable.Columns(1).Find("r10").Row, myTable.Rows(1).Find("abandoned").Column).Select
    
    
    
    
End Sub


Sub cleanColumn()
    
    Range("a1").Value = Range("a1").Column
    Range("b1").Value = Range("a1").Row
    
    
End Sub





Sub Find_First_Used_Cell()
'Finds the first non-blank cell in a worksheet.

Dim rFound As Range
    
    On Error Resume Next
    Set rFound = Cells.Find(what:="*", _
                    After:=Cells(Rows.Count, Columns.Count), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    searchorder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False)
    
    On Error GoTo 0
    With rFound
        Debug.Print .Value
        Debug.Print .Address
    End With
    
    
    rFound.Select
    If rFound Is Nothing Then
        MsgBox "All cells are blank."
    Else
        MsgBox "First Cell: " & rFound.Address
    End If
    
End Sub


Sub TestFind()
    Range("A1:D20").Find(what:="*", searchorder:=xlByColumns).Select
End Sub


Sub selectingCells()
    ' all the various [good] ways to select cells, ranges, tables, etc
    Dim tableStart As Range
    Dim table As Range
    Dim headers As Variant
    Dim headerCell As Variant
    Dim dataRange As Range
    
    Set tableStart = Cells.Find("*", Cells(Rows.Count, Columns.Count))
    ' best way to get table range: start cell (top left) to bottom right cell
    ' avoids problem of blank rows within table cell values. only requires first row & column contiguous
    ' also keeps same range if columns are deleted
    Set table = Range(tableStart, Cells(tableStart.End(xlDown).Row, tableStart.End(xlToRight).Column))
    Set headers = table.Rows(1)
    Set dataRange = Range(tableStart.Offset(1, 0), Cells(tableStart.End(xlDown).Row, tableStart.End(xlToRight).Column))
    
    
    tableStart.Select
    headers.Select
    table.Select                    ' dynamically updated if rows/cols deleted
    dataRange.Select
    table.Rows(1).Select            ' select first row (headers)
    table.Columns(1).Select     ' select first col
    table.Offset(1, 0).Select       ' select range 1 row down
    
    
        ' testing for each loop; working, but it skips when it shouldn'
'        For Each headerCell In headers.Columns
'            If headerCell Like "skill*" Then headerCell.EntireColumn.Delete
'        Next

    ' find a value in a column, and find its row/column index number
    On Error Resume Next
    table.Columns(1).Rows.Find("ff").Select
    Debug.Print table.Columns(1).Rows.Find("ff").Row
    Debug.Print table.Columns(1).Rows.Find("ff").Column
    On Error GoTo 0
    
    ' find column by name, etc
     On Error Resume Next
    table.Rows(1).Columns.Find("mose").Select
    On Error GoTo 0
    
    ' get matrix indices within table
    

End Sub



Sub RST_Report()
    Dim tableStart As Range
    Dim table As Range
    
    ' find first data cell in sheet, searching from top-left
    Set tableStart = Cells.Find("*", Cells(Rows.Count, Columns.Count))
    Set table = Range(tableStart, Cells(tableStart.End(xlDown).Row, tableStart.End(xlToRight).Column))
    
    ' deleting columns based on value; decreasing for-next-loop to avoid skipping issues; for-each loop didn't work
    For i = table.Rows(1).Columns.Count To 1 Step -1
        If table.Rows(1).Columns(i) Like "skill*" Then table.Rows(1).Columns(i).EntireColumn.Delete
    Next
    
    ' delete rows of table based on value
    For i = table.Columns(1).Rows.Count To 1 Step -1
        If table.Columns(1).Rows(i) Like "*SVR01*" Then table.Rows(i).EntireRow.Delete
    Next
    
    table.Select
    
    ' find a column header by value
    On Error Resume Next
    table.Columns(1).Rows.Find("ff").Select
    Debug.Print table.Columns(1).Rows.Find("ff").Row
    On Error GoTo 0
    
    ' find crosstab of row/column
    Dim check1 As Range
    Set check1 = Cells(table.Columns(1).Rows.Find("R10").Row, table.Rows(1).Columns.Find("asa").Column)
    
    ' check if crosstab has value, and highlight
    If check1.Value > 1 Then check1.Interior.Color = RGB(0, 255, 0) Else check1.Interior.Color = RGB(255, 0, 0)
      
    
    
    
End Sub


Private Sub FormattingAndDates()
    Debug.Print Date
    Debug.Print Format(Date, "mmmm") & "-" & Format(Date, "dd")
    Debug.Print DatePart("m", Date)
    Debug.Print Format(Date, "m")
    
    'folder name & date
    Debug.Print "CSQ " & Format(Date, "mmmm") & " - " & Format(Date, "yyyy")
    
End Sub



Sub Test_Folder_Exist_With_Dir()
'Updateby Extendoffice 20161109
    Dim sFolderPath As String
    sFolderPath = "C:\Users\Chris\Desktop\CSQ November - 2017"
    If Right(sFolderPath, 1) <> "\" Then
        sFolderPath = sFolderPath & "\"
    End If



End Sub

Sub CheckIfFolderExistsAndCreate()
'Updateby Extendoffice 20161109
    Dim fdObj As Object
    Dim dir As String
    Dim folderName As String
    
    ' Application.ScreenUpdating = False
    Set fdObj = CreateObject("Scripting.FileSystemObject")
    dir = "C:\Users\Chris\Desktop\"
    folderName = "CSQ " & Format(Date, "mmmm") & " - " & Format(Date, "yyyy")
    
    If fdObj.FolderExists(dir & folderName) Then
        MsgBox "Found it.", vbInformation, "Kutools for Excel"
    Else
        fdObj.CreateFolder (dir & folderName)
        MsgBox "It has been created.", vbInformation, "Kutools for Excel"
    End If
    ' Application.ScreenUpdating = True
End Sub
