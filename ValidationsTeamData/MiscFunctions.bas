Attribute VB_Name = "MiscFunctions"



Sub CreateDataSheet()
        Dim newDataSheet As Worksheet
        If sheetExists("Data") Then
            MsgBox "Data Sheet already exists"
            Exit Sub
        End If
        
        ThisWorkbook.Worksheets.Add after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
        ActiveSheet.Name = "Data" ' & ThisWorkbook.Worksheets.Count
        
        Range("A1").Value = "Date"
        Range("B1").Value = "Validation Type"
        Range("C1").Value = "Machine Name"
        Range("D1").Value = "BU"
        Range("E1").Value = "Register"              ' or lane #
        Range("F1").Value = "Time Zone"
        Range("G1").Value = "POS Readiness Status"
        Range("H1").Value = "Assigned"
        Range("I1").Value = "Issue with device"     ' or "Status upon login"
        Range("J1").Value = "Resolution"
        Range("K1").Value = "Status"
        Range("L1").Value = "Follow-up needed"
        Range("M1").Value = "Software version verified"
        
        With Range("a1", Range("A1").End(xlToRight))
            .Font.Name = "Times New Roman"
            .Font.Size = 12
            .Interior.Color = RGB(47, 117, 181)
            .Font.Color = vbWhite
            .Font.Bold = True
            .EntireColumn.AutoFit
        End With
End Sub





Sub RemoveDuplicates()

    Dim myData As Range
    Dim before As Long
    Dim after As Long
    
    If Not sheetExists("Data") Then
        MsgBox "sheet does not exist"
        Exit Sub
    End If
    
    ThisWorkbook.Worksheets("Data").Activate
    Set myData = Range("A2", Cells(Cells(Rows.Count, 1).End(xlUp).Row, Range("A1").End(xlToRight).column))
    before = myData.Rows.Count
    
    myData.RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13), Header:=xlNo
    Set myData = Range("A2", Cells(Cells(Rows.Count, 1).End(xlUp).Row, Range("A1").End(xlToRight).column))
    after = myData.Rows.Count
    
    MsgBox before - after & " duplicate rows removed.", Title:="Remove Duplicates"
    
End Sub


Function sheetExists(sheetToFind As String) As Boolean
    sheetExists = False
    For Each Sheet In Worksheets
        If sheetToFind = Sheet.Name Then
            sheetExists = True
            Exit Function
        End If
    Next Sheet
End Function



Sub RemoveNonLanes()

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    Dim myData As Range
    Dim before As Long
    Dim after As Long
    
    If Not sheetExists("Data") Then
        MsgBox "sheet does not exist"
        Exit Sub
    End If
    
    ThisWorkbook.Worksheets("Data").Activate
    Set myData = Range("A2", Cells(Cells(Rows.Count, 1).End(xlUp).Row, Range("A1").End(xlToRight).column))
    before = myData.Rows.Count
    
    ' delete non "WFM" rows
    Dim r As Integer
        For r = myData.Rows.Count To 1 Step -1
            With myData
                If Not LCase(Left(Trim(.Cells(r, "C")), 3)) Like "wfm*" Then
                    .Rows(r).EntireRow.Delete
                End If
            End With
        Next
                
    Set myData = Range("A2", Cells(Cells(Rows.Count, 1).End(xlUp).Row, Range("A1").End(xlToRight).column))
    after = myData.Rows.Count
    
    MsgBox before - after & " rows without 'WFM..' removed.", Title:="Remove Non-Lane Rows"
    
ExitHandler:
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox Err.Description
    Resume ExitHandler
End Sub
