Attribute VB_Name = "ConsolidateValidationsTeamData"
Option Compare Text

Sub ConsolidateValidationsData()

    Dim FilesToOpen
    Dim x As Integer
    Dim lastrow As Variant

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    FilesToOpen = Application.GetOpenFilename _
      (MultiSelect:=True, Title:="Files to Merge")

    If TypeName(FilesToOpen) = "Boolean" Then
        MsgBox "No Files were selected"
        GoTo ExitHandler
    End If

    x = 1
    While x <= UBound(FilesToOpen)
    
    ' Open Workbook x
        Workbooks.Open Filename:=FilesToOpen(x), UpdateLinks:=False
            
    ' BEGIN DATA CLEAN+COPY SUBROUTINE
        Call CleanData

    ' Close source workbook without saving
        ActiveWorkbook.Close False
        
        x = x + 1
    Wend

ExitHandler:
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox Err.Description
    Resume ExitHandler
End Sub






Sub CleanData()

    ' loop through each (visible) data sheet in opened workbook
    For i = 1 To ActiveWorkbook.Worksheets.Count
        
        ' worksheets to skip
        If Worksheets(i).Visible = False Then GoTo NextIteration
        If (Worksheets(i).Name Like "*validators sched*" Or Worksheets(i).Name Like "*outstanding lanes*") Then GoTo NextIteration
        ActiveWorkbook.Worksheets(i).Activate
            
           
    ' delete rows with blanks & sub-header rows
            Dim datastart As Range
            Dim dataend As Range
            
            Set datastart = Range("A2")
            Set dataend = Cells(Rows.Count, 1).End(xlUp)
    
    ' delete rows with empty first column
            On Error Resume Next
            Range(datastart, dataend).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
'    On Error GoTo ErrHandler

    
    
    ' delete sub-header rows (SVR01, BU#s) and rows with random spaces
            Dim r As Integer
                For r = ActiveSheet.UsedRange.Rows.Count To 1 Step -1
                    If Left(Trim(Cells(r, "A")), 3) Like "*SVR*" Or Left(Trim(Cells(r, "A")), 4) Like "[0-9][0-9][0-9][0-9]" _
                    Or Trim(Cells(r, "A")) = "" Then
                        Rows(r).EntireRow.Delete
                    End If
                Next
  
    ' get table column range and row count
            Dim column As Range
            Dim rowCount As Variant

            Set column = Range(datastart, Cells(datastart.End(xlDown).Row, 1))
            rowCount = column.Rows.Count
    
    ' get date from name of workbook
            Dim wkbkname As String
            Dim prdposition As String
            Dim shavestr As String
            
            wkbkname = ActiveWorkbook.Name
            prdposition = InStr(wkbkname, ".") - 1
            shavestr = Left(wkbkname, prdposition)
            
            
    ' get validation type from name of sheet, insert value into new column
            Dim sheetName As String
            sheetName = ActiveSheet.Name
            
            Columns("A").Insert shift:=xlToRight, copyorigin:=xlFormatFromLeftOrAbove
            Set datastart = Range("A2")
            datastart.Value = "Validation Type"
            Range(datastart.Offset(1, 0).Address & ":" & datastart.Offset(rowCount - 1, 0).Address).Value = sheetName
            
            
    ' insert column
            Columns("A").Insert shift:=xlToRight, copyorigin:=xlFormatFromLeftOrAbove
            Set datastart = Range("A2")
            datastart.Value = "Date"
            
    ' paste date value from workbook name in new column range
            Range(datastart.Offset(1, 0).Address & ":" & datastart.Offset(rowCount - 1, 0).Address).Value = shavestr
               
    ' get table data range dimensions
            Dim tbl As Range
            Set tbl = Range(datastart.Offset(1, 0), Cells(datastart.End(xlDown).Row, datastart.End(xlToRight).column))
            
    ' find first blank in Data sheet, and paste visible rows
            Set lastrow = Workbooks("Validations Data Consolidator.xlsm").Worksheets("Data").Range("A1000000").End(xlUp).Offset(1, 0)
            ' tbl.Copy Destination:=lastrow
            tbl.SpecialCells(xlCellTypeVisible).Copy lastrow

'ExitHandler:
'    Application.ScreenUpdating = True
'    Exit Sub
'
'ErrHandler:
'    MsgBox Err.Description
'    Resume ExitHandler
           
NextIteration:
    Next i
    
End Sub





