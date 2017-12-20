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
    
    Debug.Print Environ("UserName")
    
    If Not fdObj.FolderExists(dir & folderName) Then
        fdObj.CreateFolder (dir & folderName)
    End If
    ' Application.ScreenUpdating = True
End Sub


' not working - chrome driver?
Sub seleniumtutorial()
    Dim driverName As New WebDriver
    driverName.start "chrome", "http://www.google.com"
    driverName.Quit
End Sub


Private Sub ArrayTest()
    Dim testArray() As Variant
    Dim rng As Range
    Dim start As Range: Set start = Range("A1")
    
    Set rng = Range(start, Range("A1").End(xlDown))
    rng.Select
    Debug.Print rng.Count
    Debug.Print rng.Cells.Count
    
    ReDim testArray(1 To rng.Count)
    
    Dim counter As Integer
    For counter = 1 To rng.Rows.Count
        testArray(counter) = Range("A1").Offset(counter - 1, 0)
    Next
    
    For counter = 1 To rng.Count
        Range("A1").Offset(counter - 1, 1).Value = testArray(counter) + 1
        start.Offset(1, 0).Select
    Next
    
    Erase testArray
End Sub


Private Sub matchColumnsAndPaste()
    ' copy + paste data columns based on matching column header
    ' CURRENTLY ONLY WORKS IF ALL DATA COLUMNS IN TABLE HAVE SAME LENGTH
    
    Dim headers As Range
    
    ' have to select sheet before setting the range equal to the range on that sheet
    Worksheets("data").Activate
    Set headers = Sheets("data").Range("A1", Range("A1").End(xlToRight))
    
    ' various selection methods
'    headers(2).Select
'    headers(2).Offset(1, 0).Select
'    Debug.Print headers(2).Column
'    Cells(Columns(2).Rows.Count, 2).End(xlUp).Select
'    Cells(5, headers(2).Column).Select
'    Cells(Rows.Count, headers(2).Column).Select     ' select bottom of column
    
    Dim dataTable As Range
    Dim rawData As Range
    Dim i As Integer
    
    For i = 1 To Application.Worksheets.Count
        Worksheets(i).Activate
        If Not ActiveSheet.Name Like "data" Then
            Set dataTable = Range("A1", Range("A1").End(xlDown).End(xlToRight))
            Set rawData = Range(Range("A1").Offset(1, 0), Range("A1").End(xlDown).End(xlToRight))

            ' loop through each data header vs proper header. paste at the end of the appropriate column on the data tab
            For j = 1 To dataTable.Rows(1).Columns.Count
                For k = 1 To headers.Count
                    Debug.Print dataTable.Rows(1).Columns(j)
                    Debug.Print headers(k)
                    If dataTable.Rows(1).Columns(j) Like headers(k) Then
                        rawData.Columns(j).Copy Worksheets("data").Cells(Rows.Count, headers(k).Column).End(xlUp).Offset(1, 0)
                    End If
                Next k
            Next j
            
            ' need to copy only the needed columns
            ' then need to make sure data input is incremented - not just append to existing columns
            ' (since it would fll in missing columns)
            
            ' something array put in proper dimension based on header order ...
        End If
    Next
End Sub

' copy a range to a diff worksheet works
Sub cp()
    Range("A1", Range("A1").End(xlDown)).Copy Worksheets("data").Range("A2")
End Sub







' function to return cross section of row / columns
Function GetTableCell(ByVal rowHeader As String, ByVal colHeader As String, ByVal table As Range)
    
    Set GetTableCell = Cells(table.Columns(1).Rows.Find(rowHeader, MatchCase:=False).Row, table.Rows(1).Columns.Find(colHeader, MatchCase:=False).Column)
    GetTableCell.Select
    If GetTableCell.Value > 1 Then GetTableCell.Interior.Color = RGB(100, 99, 99)
End Function

Sub testSelectFunction()
    Dim myrange: Set myrange = Range("a1:c4")
    Set table = myrange
    rowHeader = "r10"
    colHeader = "thing"
    
    ' myrange.Select
    Cells(table.Columns(1).Rows.Find(rowHeader, MatchCase:=False).Row, table.Rows(1).Columns.Find(colHeader, MatchCase:=False).Column).Select
    Call GetTableCell("r10", "queue", myrange)
    Call GetTableCell("r10", "thing", myrange)
    
    
End Sub

