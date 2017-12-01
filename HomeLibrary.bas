Attribute VB_Name = "HomeLibrary"
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
    
    If rFound Is Nothing Then
        MsgBox "All cells are blank."
    Else
        MsgBox "First Cell: " & rFound.Address
    End If
    
End Sub


Sub TestFind()
    Range("A1:D20").Find(what:="*", searchorder:=xlByColumns).Select
    
    
    
    
    
End Sub
