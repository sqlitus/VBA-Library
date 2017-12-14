Attribute VB_Name = "Sandbox"
Option Compare Text


' colors pie chart data series based on series name

Private Sub ColorPies()

        Dim cht As ChartObject
        Dim i As Integer
        Dim vntValues As Variant
        Dim s As String
        Dim myseries As Series
         
            For Each cht In ActiveSheet.ChartObjects
                For Each myseries In cht.Chart.SeriesCollection
         
                    If myseries.ChartType <> xlPie Then GoTo SkipNotPie
                    s = Split(myseries.Formula, ",")(2)
                    vntValues = myseries.Values
                   
                    For i = 1 To UBound(vntValues)
                        myseries.Points(i).Interior.Color = Range(s).Cells(i).Interior.Color
                    Next i
SkipNotPie:
                Next myseries
            Next cht
            
End Sub



' Various Report Automation Macros




Private Sub OLD_RST_Report()
Attribute OLD_RST_Report.VB_ProcData.VB_Invoke_Func = " \n14"
'
' RST_Report Macro - for all 4 teams present only
'

'
    Columns("B:B").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("L:L").Select
    Selection.Delete Shift:=xlToLeft
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    Columns("P:P").Select
    Selection.Delete Shift:=xlToLeft
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
Private Sub screenupdate()
    Application.ScreenUpdating = True
End Sub


' loop sheets and get name
Private Sub LoopThroughSheets()
    For i = 1 To Worksheets.count
            Worksheets(i).Activate
            If (ActiveSheet.Name Like "*validation schedule*" Or ActiveSheet.Name Like "*other*") Then GoTo NextIteration
        Range("a5").Value = ActiveSheet.Name
NextIteration:
    Next i
End Sub




' delete rows if contains value
Private Sub Value()
Dim r As Integer
    For r = ActiveSheet.UsedRange.Rows.count To 1 Step -1
        If Cells(r, "A") Like "SVR01" Then
            Rows(r).EntireRow.Delete
        End If
    Next
End Sub


' delete columns if header contains value
Private Sub DeleteColumnsByName()

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
Private Sub DeleteEmptyAndBlankRows()

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



Private Sub CreateDataDumpSheet()
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
Private Sub testLeft()
If Left(ActiveCell.Value, 4) Like "[0-9][0-9][0-9][0-9]" Then Range("A1").Value = "is true"
End Sub



' count sheets, workbooks
Private Sub ActiveWorkbookWorksheetCount()
    
    Range("A1").Value = ActiveWorkbook.Worksheets.count
        

    ' count visible sheets
    Dim xSht As Variant
    Dim i As Long
    For Each xSht In ActiveWorkbook.Worksheets
        If xSht.Visible Then i = i + 1
    Next
    Range("A2").Value = i
    
    For count = 1 To ActiveWorkbook.Worksheets.count
        If Worksheets(count).Visible = False _
        Or 5 > 9 Then GoTo NextIteration
        If Worksheets(count).Visible Then MsgBox Sheets(count).Name
        
NextIteration:
    Next count

        
End Sub



' find start of data
Private Sub FindDataStart()

Dim colchoice As String
colchoice = "A"


' find table start position with column of choice
        Dim searchcell As Range
        Dim tblstart As Variant
        Dim tblrow As Variant
        Dim tblcol As Variant
        
        
        With Columns(colchoice)
            Set searchcell = .Find(What:="*", LookIn:=xlValues)
        End With
        tblstart = searchcell.Address
        tblrow = CInt(Range(tblstart).Row)
        tblcol = CInt(Range(tblstart).column)
        
        Range("a10").Value = tblstart
        Range("a11").Value = tblrow
        Range("a12").Value = tblcol
                
End Sub



' visible sheets test

Private Sub VisibleSheetsCount()

    Dim numVisibleSheets As Integer: numVisibleSheets = 0
    
    For i = 1 To ActiveWorkbook.Worksheets.count
        If Worksheets(i).Visible = True Then numVisibleSheets = numVisibleSheets + 1
    Next i
    
    Sheets("Macro").Activate
    Range("A1").Value = numVisibleSheets
End Sub




' various ways of getting end of contiguous data, selecting data range
Private Sub LastRowOfData()
    Cells(Rows.count, 1).End(xlUp).Select
    Range("A1000000").End(xlUp).Offset(1, 0).Select
    Cells(Rows.count, 1).End(xlUp).Select
    Cells(Rows.count, 1).End(xlUp).Offset(1, 0).Select
    Cells(1, Range("A1").End(xlToRight).column).Select
    
    ' bottom right of table
    Cells(Cells(Rows.count, 1).End(xlUp).Row, Range("A1").End(xlToRight).column).Select
    
End Sub


' for loop testing
Private Sub ForLoopTest()
    For i = 1 To 10
        Cells(i, 1).Value = i
        If i = 5 Then Exit For
        
    Next
End Sub

' copy and paste VISIBLE rows from a range
Private Sub CopyVisibleCells()
    Range("A2", Range("A2").End(xlDown).End(xlToRight)).SpecialCells(xlCellTypeVisible).Copy Worksheets("Sheet1").Range("a1")
End Sub


Private Sub RowsAndColumns()
    Range("a1").Value = Rows.count
    
    Cells(1, Columns.count).End(xlToLeft).Select
    ' scan down col 1. if found nothing, scan down col 2 [, ...n] until data is found
    
End Sub




Private Sub SelectTable()

    ' find first cell of table
    Dim rFound As Range
    
    On Error Resume Next
    Set rFound = Cells.Find(What:="*", _
                    after:=Cells(Rows.count, Columns.count), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
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



Private Sub CellForEachTest()
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




Private Sub DeleteDuplicateRows()

    With ActiveSheet
        Set rng = Range("A1", Range("a1").End(xlDown).End(xlToRight))
        rng.Select
        
        rng.RemoveDuplicates Columns:=Array(1, 2, 3, 4), Header:=xlNo
    End With
    
End Sub



Private Sub BreakLinks() ' not working?
'Updateby20140318
Dim wb As Workbook
Set wb = Application.ActiveWorkbook
If Not IsEmpty(wb.LinkSources(xlExcelLinks)) Then
    For Each link In wb.LinkSources(xlExcelLinks)
        wb.BreakLink link, xlLinkTypeExcelLinks
    Next link
End If
End Sub


' working with file system objects
Private Sub createFileSystemObject()
    ' method 1 - load Microsoft Scripting Library reference
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
            ' method 2 - create file system object without loading library (no intellisense)
            Dim fdObj As Object
            Set fdObj = CreateObject("Scripting.FileSystemObject") ' collection class
    
    
    ' get a file [object] & its properties
    Dim myFile As Scripting.File
    Set myFile = fso.GetFile("C:\Work\AgileTasks.xlsx")
    Debug.Print myFile.Attributes, myFile.DateCreated, myFile.DateLastAccessed, myFile.DateLastModified
    Debug.Print myFile.Drive, myFile.Name, myFile.ParentFolder, myFile.Path
    Debug.Print myFile.ShortName, myFile.ShortPath, myFile.Size, myFile.Type
    
    ' manipulate the file object variable directly
    myFile.Copy myFile.ParentFolder & "\" & "vba copy of " & myFile.Name
    
End Sub


Private Sub findCsqarDates()
    ' REQUIRES: REGEX LIBRARY Microsoft VBScript Regular Expressions 5.5
    
    Dim regEx As New RegExp
    Dim findCell As Range: Set findCell = Cells.Find("*filter*", Cells(Rows.count, Columns.count))
    
    Dim startDate As String: startDate = "Starting At"
    Dim endDate As String
    
    With regEx
        .Pattern = startDate
    End With
    
    If regEx.Test(findCell) Then
        MsgBox "yes"
    Else
        MsgBox "not matched"
    End If
    
End Sub

    
' subs to help consolidate validations data
' count num colums, copy & paste correct columns to masterworksheet

Public Function LastOccupiedColNum(Sheet As Worksheet) As Long
    Dim lng As Long
    If Application.WorksheetFunction.CountA(Sheet.Cells) <> 0 Then
        With Sheet
            lng = .Cells.Find(What:="*", _
                              after:=.Range("A1"), _
                              LookAt:=xlPart, _
                              LookIn:=xlFormulas, _
                              SearchOrder:=xlByColumns, _
                              SearchDirection:=xlPrevious, _
                              MatchCase:=False).column
        End With
    Else
        lng = 1
    End If
    LastOccupiedColNum = lng
End Function

Private Sub getCoulmns()
    LastOccupiedColNum (ActiveWorkbook.ActiveSheet)
    
End Sub


Private Sub CopyPasteWithoutLink()
'    Worksheets("Sheet1").Range("A1:A10").Copy
'    Worksheets("Sheet2").Range("A1").PasteSpecial xlPasteValuesAndNumberFormats

    Dim thiswkbk As Workbook: Set thiswkbk = ActiveWorkbook
    
    Workbooks("Validations Data Consolidator.xlsm").Activate
    Worksheets("Data").Activate
    Range("A1", Range("A1").End(xlDown).End(xlToRight)).Copy
    
    thiswkbk.Activate
    Worksheets("Sheet1").Activate
    Range("A1").PasteSpecial xlPasteFormats
    Range("A1").PasteSpecial xlPasteValues
    
End Sub


Private Sub CreateDataSheet()
        Dim newDataSheet As Worksheet
        
        ThisWorkbook.Worksheets.Add after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count)
        ActiveSheet.Name = "Data" & ThisWorkbook.Worksheets.count
        
        Range("A1").Value = "Date"
        Range("B1").Value = "Validation Type"
        Range("C1").Value = "Machine Name"
        Range("D1").Value = "BU"
        Range("E1").Value = "Register"
        Range("F1").Value = "Time Zone"
        Range("G1").Value = "POS Readiness Status"
        Range("H1").Value = "Assigned"
        Range("I1").Value = "Issue with device"
        Range("J1").Value = "Resolution"
        Range("K1").Value = "Stuatus"
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

Private Sub adjColWidth()
    Range("A1").EntireColumn.ColumnWidth = Range("A1").EntireColumn.ColumnWidth * 1.1
    With Range("A1").EntireColumn
        .ColumnWidth = .ColumnWidth * 1.2
    End With
    
    Set test1 = Range("A2")
    Debug.Print test1.Value
End Sub


Sub REORDER()

        Dim arrColOrder As Variant, ndx As Integer
        Dim Found As Range, counter As Integer
        Dim mycols() As Variant
        
        'Place the column headers in the end result order you want.
        arrColOrder = Array("COLUMN 2", "COLUMN 4", "COLUMN 6", "COLUMN 10", "COLUMN 1", _
                            "COLUMN 9", "COLUMN 3", "COLUMN 8", "COLUMN 7", "COLUMN 5")
        
        mycols = Range("A1", Range("A1").End(xlToRight))
        counter = 1
        
        For ndx = LBound(arrColOrder) To UBound(arrColOrder)
            Set Found = Rows("1:1").Find(arrColOrder(ndx), LookIn:=xlValues, LookAt:=xlWhole, _
                              SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
            If Not Found Is Nothing Then
                If Found.column <> counter Then
                    Found.EntireColumn.Cut
                    Columns(counter).Insert Shift:=xlToRight
                    Application.CutCopyMode = False
                End If
                counter = counter + 1
            End If
        Next ndx

End Sub

