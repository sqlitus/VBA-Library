Attribute VB_Name = "Sandbox"
Option Compare Text


' colors pie chart data series based on series name

Sub ColorPies()

        Dim cht As ChartObject
        Dim I As Integer
        Dim vntValues As Variant
        Dim s As String
        Dim myseries As Series
         
            For Each cht In ActiveSheet.ChartObjects
                For Each myseries In cht.Chart.SeriesCollection
         
                    If myseries.ChartType <> xlPie Then GoTo SkipNotPie
                    s = Split(myseries.Formula, ",")(2)
                    vntValues = myseries.Values
                   
                    For I = 1 To UBound(vntValues)
                        myseries.Points(I).Interior.Color = Range(s).Cells(I).Interior.Color
                    Next I
SkipNotPie:
                Next myseries
            Next cht
            
End Sub



' Various Report Automation Macros




Sub RST_Report()
Attribute RST_Report.VB_ProcData.VB_Invoke_Func = " \n14"
'
' RST_Report Macro - for all 4 teams present only
'

'
    Columns("B:B").Select
    Selection.Delete shift:=xlToLeft
    Selection.Delete shift:=xlToLeft
    Columns("L:L").Select
    Selection.Delete shift:=xlToLeft
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    Columns("P:P").Select
    Selection.Delete shift:=xlToLeft
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "Queue"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "Aloha"
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "NCR_Tech"
    Range("A5").Select
    ActiveCell.FormulaR1C1 = "Payment"
    Range("A6").Select
    ActiveCell.FormulaR1C1 = "R10"
    Range("A7").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "Max Time To voicemail"
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "Avg Time To voicemail"
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "Calls voicemail"
    Range("C6").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("I6").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
End Sub





' screen update
Sub screenupdate()
    Application.ScreenUpdating = True
End Sub


' loop sheets and get name
Sub LoopThroughSheets()
    For I = 1 To Worksheets.count
            Worksheets(I).Activate
            If (ActiveSheet.Name Like "*validation schedule*" Or ActiveSheet.Name Like "*other*") Then GoTo NextIteration
        Range("a5").Value = ActiveSheet.Name
NextIteration:
    Next I
End Sub




' delete rows if contains value
Sub Value()
Dim r As Integer
    For r = ActiveSheet.UsedRange.Rows.count To 1 Step -1
        If Cells(r, "A") Like "SVR01" Then
            Rows(r).EntireRow.Delete
        End If
    Next
End Sub


' delete columns if header contains value
Sub DeleteColumnsByName()

Dim headers As Range
Dim cell As Range
Set headers = Range("A2:f2")

Range("A1").Value = headers.count
    For Each cell In headers
        If cell Like "*token*" Then
        cell.EntireColumn.Delete
        End If
    Next cell
End Sub




' select data start and end, delete empty rows. then delete unnecessary header rows
Sub DeleteEmptyAndBlankRows()

    Dim datastart As Range
    Dim dataend As Range
    
    Set datastart = Range("A2")
    Set dataend = Cells(Rows.count, 1).End(xlUp)
    
    
    ' delete rows with empty first column
    Range(datastart, dataend).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
    ' delete sub-header rows
    Dim r As Integer
    
    For r = ActiveSheet.UsedRange.Rows.count To 1 Step -1
        If Cells(r, "A") Like "SVR01" Then
            Rows(r).EntireRow.Delete
        End If
    Next
End Sub



Sub CreateDataDumpSheet()
    ' create data dump sheet if doesn't exist
    Dim wsTest As Worksheet
    Const strSheetName As String = "FindMe"
     
    Set wsTest = Nothing
    On Error Resume Next
    Set wsTest = ActiveWorkbook.Worksheets(strSheetName)
    On Error GoTo 0
     
    If wsTest Is Nothing Then
        Worksheets.Add.Name = strSheetName
    End If

End Sub




' if cell starts with num
Sub testLeft()
If Left(ActiveCell.Value, 4) Like "[0-9][0-9][0-9][0-9]" Then Range("A1").Value = "is true"
End Sub



' count sheets, workbooks
Sub ActiveWorkbookWorksheetCount()
    
    Range("A1").Value = ActiveWorkbook.Worksheets.count
        

    ' count visible sheets
    Dim xSht As Variant
    Dim I As Long
    For Each xSht In ActiveWorkbook.Worksheets
        If xSht.Visible Then I = I + 1
    Next
    Range("A2").Value = I
    
    For count = 1 To ActiveWorkbook.Worksheets.count
        If Worksheets(count).Visible = False _
        Or 5 > 9 Then GoTo NextIteration
        If Worksheets(count).Visible Then MsgBox Sheets(count).Name
        
NextIteration:
    Next count

        
End Sub


' smart rst macro
Sub SmartRSTMacro()
    
    Dim tableStart As Variant
    tableStart = Range("A1").End(xlDown)
    
    Range("A1").Value = tableStart
    
    
End Sub



' find start of data
Sub FindDataStart()

Dim colchoice As String
colchoice = "A"


' find table start position with column of choice
        Dim searchcell As Range
        Dim tblstart As Variant
        Dim tblrow As Variant
        Dim tblcol As Variant
        
        
        With Columns(colchoice)
            Set searchcell = .Find(what:="*", LookIn:=xlValues)
        End With
        tblstart = searchcell.Address
        tblrow = CInt(Range(tblstart).Row)
        tblcol = CInt(Range(tblstart).column)
        
        Range("a10").Value = tblstart
        Range("a11").Value = tblrow
        Range("a12").Value = tblcol
                
End Sub



' visible sheets test

Sub VisibleSheetsCount()

    Dim numVisibleSheets As Integer: numVisibleSheets = 0
    
    For I = 1 To ActiveWorkbook.Worksheets.count
        If Worksheets(I).Visible = True Then numVisibleSheets = numVisibleSheets + 1
    Next I
    
    Sheets("Macro").Activate
    Range("A1").Value = numVisibleSheets
End Sub




' various ways of getting end of contiguous data
Sub LastRowOfData()
    Cells(Rows.count, 1).End(xlUp).Select
    Range("A1000000").End(xlUp).Offset(1, 0).Select
    Cells(Rows.count, 1).End(xlUp).Select
    Cells(Rows.count, 1).End(xlUp).Offset(1, 0).Select
End Sub


' for loop testing
Sub ForLoopTest()
    For I = 1 To 10
        Cells(I, 1).Value = I
        If I = 5 Then Exit For
        
    Next
End Sub

' copy and paste VISIBLE rows from a range
Sub CopyVisibleCells()
    Range("A2", Range("A2").End(xlDown).End(xlToRight)).SpecialCells(xlCellTypeVisible).Copy Worksheets("Sheet1").Range("a1")
End Sub


Sub RowsAndColumns()
    Range("a1").Value = Rows.count
    
    Cells(1, Columns.count).End(xlToLeft).Select
    ' scan down col 1. if found nothing, scan down col 2 [, ...n] until data is found
    
End Sub




Sub SelectTable()

    ' find first cell of table
    Dim rFound As Range
    
    On Error Resume Next
    Set rFound = Cells.Find(what:="*", _
                    After:=Cells(Rows.count, Columns.count), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    searchorder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=False)
    On Error GoTo 0
    

    ' get table range
    Dim myTable
    
    Set myTable = Range(rFound, rFound.End(xlDown).End(xlToRight))

    
    ' delete unneeded columns
    Dim headerCell As Range

    For Each headerCell In myTable.Columns
        If headerCell.Value Like "skills" Then headerCell.EntireColumn.Delete
    Next headerCell
End Sub



Sub CellForEachTest()
    Dim headerCell As Range
    Dim headers As Range
    
    Set myTable = Range("A2", Range("A2").End(xlDown).End(xlToRight))
    
    myTable.Select
    Set headers = myTable.Rows(1)
    

    For Each headerCell In headers
        Debug.Print headerCell.Address
        Debug.Print headers.Address
        
        If headerCell.Value Like "skills" Then headerCell.EntireColumn.Delete
    Next headerCell

End Sub



Sub CheckIfFolderFileExistsAndSave()
    ' check if directory subfolders exist, and create if necessary. Then save file w/ proper name

    Dim fdObj As Object
    Dim dir As String
    Dim yearFolder As String
    Dim monthFolder As String
    Dim reportDate As Date
    
    ' date of report - (yesterday for CSQAR)
    reportDate = DateAdd("m", 1, Date)
    
    ' invoke file system object reference library
    Set fdObj = CreateObject("Scripting.FileSystemObject")
    
    ' create year & month folders for report's date if they don't exist
    dir = "\\wfm-team\Team\Retail Support Team\Reporting\CSQ Activity Reports\"
    yearFolder = "CSQ Activity Reports - " & Format(reportDate, "yyyy") & "\"
    monthFolder = "CSQ Activity Reports - " & Format(reportDate, "mmmm") & " " & Format(reportDate, "yyyy") & "\"
    
    If Not fdObj.FolderExists(dir & yearFolder) Then fdObj.createfolder (dir & yearFolder)
    If Not fdObj.FolderExists(dir & yearFolder & monthFolder) Then fdObj.createfolder (dir & yearFolder & monthFolder)
    
    ' save file (without extras)
    saveFileName = "CSQ Activity Report - " & Format(reportDate, "mmddyyyy") & ".xlsx"
    If Not fdObj.FileExists(dir & yearFolder & monthFolder & saveFileName) Then
        ActiveWorkbook.SaveAs dir & yearFolder & monthFolder & saveFileName, xlWorkbookNormal
    End If
     
End Sub



Sub DeleteDuplicateRows()

    With ActiveSheet
        Set rng = Range("A3", Range("B3").End(xlDown))
        rng.Select
        
        rng.RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes
        Range("a1").RemoveDuplicates
    End With

End Sub


