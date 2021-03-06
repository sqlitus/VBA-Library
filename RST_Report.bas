Attribute VB_Name = "RST_Report"

Sub Run_RST_Report()
    Application.ScreenUpdating = False
    On Error GoTo ErrHandler
       
    Call RSTReport_CheckIfCSQARReport
    Call TrimData
    Call RSTReport_SaveToFolders
    Call RSTReport_TransformAndFormatData
    
ExitHandler:
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox Err.Description
    Resume ExitHandler
End Sub



Private Sub RSTReport_CheckIfCSQARReport()
    ' check if proper CSQAR Report to use the macro on
    
    Dim findCell As Range: Set findCell = Cells.Find("*by CCX*", Cells(Rows.count, Columns.count), MatchCase:=False)
    Dim findCell2 As Range: Set findCell2 = Cells.Find("*OPOS_R10*", Cells(Rows.count, Columns.count), MatchCase:=False)
    
    If findCell Is Nothing Or findCell2 Is Nothing Then
        Application.ScreenUpdating = True
        MsgBox "NOT R10 RST REPORT. MACRO NOT RUN"
        End
    End If
    
End Sub


Private Sub TrimData()

    ' find data table range
    Dim tableStart As Range
    Dim tableEnd As Range
    Dim table As Range
    Dim skills_column As Range
    Dim skills_data As Range
    
    Set tableStart = Cells.Find("*", Cells(Rows.count, Columns.count), MatchCase:=False)
    Set tableEnd = Cells.Find(What:="*", after:=Range("A1"), Lookat:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, _
                              SearchDirection:=xlPrevious, _
                              MatchCase:=False)
    Set table = Range(tableStart, tableEnd)
        
    ' unmerge
    table.Columns(1).UnMerge
    
    ' delete fluff rows identified by 'skills' column
    Set skills_column = table.Columns(table.Rows(1).Find("Skills").column)
    colstart = skills_column.Rows(1).Address
    colend = skills_column.Rows(skills_column.Rows.count).Offset(-1, 0).Address
    
    Set skills_data = Range(colstart, colend)
    
    On Error Resume Next
    skills_data.SpecialCells(xlCellTypeBlanks).Rows.EntireRow.Delete
    On Error GoTo 0
    
    
    '' bottom right cell of data
    
                              

End Sub




Private Sub RSTReport_SaveToFolders()
    ' check if directory subfolders exist, and create if necessary. Then save file w/ proper name

    Dim fdObj As Object
    Dim dir As String
    Dim yearFolder As String
    Dim monthFolder As String
    Dim reportDate As Date
    
    ' regex find & check dates ...most of this should be passed to a function as an argument.
    Dim regEx As Object: Set regEx = CreateObject("VBScript.RegExp")
    Dim findCell As Range: Set findCell = Cells.Find("*filter*", Cells(Rows.count, Columns.count))
    Dim startDate As Date
    Dim endDate As Date
    
    regEx.Pattern = "[\d][\d]/[\d][\d]/[\d][\d][\d][\d]"
    regEx.Global = True      ' global returns all matches. otherwise, just returns first
    
    ' check filter dates, and set report date
    If regEx.Test(findCell) Then
        startDate = regEx.Execute(findCell.Value)(0)
        endDate = regEx.Execute(findCell.Value)(1)
        
        If startDate = endDate Then
            reportDate = startDate
        Else
            MsgBox "START & END DATES NOT THE SAME. ENDING MACRO."
            End
        End If
    Else
        MsgBox "COULD NOT FIND START AND END DATES. ENDING MACRO."
        End
    End If
    
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
        ActiveWorkbook.SaveAs dir & yearFolder & monthFolder & saveFileName, xlWorkbookDefault
    End If
     
End Sub



Private Sub RSTReport_TransformAndFormatData()

    ' find data table range
    Dim tableStart As Range
    Dim table As Range
    
    Set tableStart = Cells.Find("*", Cells(Rows.count, Columns.count), MatchCase:=False)
    Set table = Range(tableStart, Cells(tableStart.End(xlDown).Row, tableStart.End(xlToRight).column))
    
    
    ' delete unneeded columns
    Dim i As Integer
    
    For i = table.Rows(1).Columns.count To 1 Step -1
        If LCase(table.Rows(1).Columns(i)) Like "csq id" _
            Or LCase(table.Rows(1).Columns(i)) Like "skills" _
            Or LCase(table.Rows(1).Columns(i)) Like "avg abandon per day" _
            Or LCase(table.Rows(1).Columns(i)) Like "calls handled by other" _
        Then table.Rows(1).Columns(i).EntireColumn.Delete
    Next
    
    ' rename columns & row values
    On Error Resume Next
        table.Rows(1).Columns.Find("csq name", MatchCase:=False).Value = "Queue"
        table.Rows(1).Columns.Replace "dequeued", "to Voicemail"
        table.Rows(1).Columns.Replace "dequeue", "Voicemail"
        table.Columns(1).Rows.Replace "OPOS_", ""
    On Error GoTo 0
    
    
    
    ' conditionally highlight cells in table
    Dim check1 As Range
    Dim check2 As Range
    Dim checkflag As Integer: checkflag = 0
    
    Set check1 = Cells(table.Columns(1).Rows.Find("r10", MatchCase:=False).Row, _
            table.Rows(1).Columns.Find("avg queue time", MatchCase:=False).column)
    
    With check1
        If Minute(.Value) < 2 Then
            .Interior.Color = RGB(146, 208, 80)
        Else
            .Interior.Color = RGB(255, 0, 0)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
        End If
    End With
    
    
    Set check1 = Cells(table.Columns(1).Rows.Find("r10", MatchCase:=False).Row, _
        table.Rows(1).Columns.Find("calls abandoned", MatchCase:=False).column)
    Set check2 = Cells(table.Columns(1).Rows.Find("r10", MatchCase:=False).Row, table.Rows(1).Columns.Find("calls presented", MatchCase:=False).column)
    
    With check1
        If .Value / check2.Value <= 0.1 Then
            .Interior.Color = RGB(146, 208, 80)
        Else
            .Interior.Color = RGB(255, 0, 0)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            checkflag = 1
        End If
    End With
    
    ' finish
    tableStart.CurrentRegion.Select
    Selection.Copy
    If checkflag = 1 Then MsgBox "Breach SLA: Mail to RST.L1.Leadership@wholefoods.com first", Title:="Mailto Steph"
    
End Sub






'' Functions to find last row & columns


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'INPUT       : Sheet, the worksheet we'll search to find the last row
'OUTPUT      : Long, the last occupied row
'SPECIAL CASE: if Sheet is empty, return 1
Public Function LastOccupiedRowNum(Sheet As Worksheet) As Long
    Dim lng As Long
    If Application.WorksheetFunction.CountA(Sheet.Cells) <> 0 Then
        With Sheet
            lng = .Cells.Find(What:="*", _
                              after:=.Range("A1"), _
                              Lookat:=xlPart, _
                              LookIn:=xlFormulas, _
                              SearchOrder:=xlByRows, _
                              SearchDirection:=xlPrevious, _
                              MatchCase:=False).Row
        End With
    Else
        lng = 1
    End If
    LastOccupiedRowNum = lng
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'INPUT       : Sheet, the worksheet we'll search to find the last column
'OUTPUT      : Long, the last occupied column
'SPECIAL CASE: if Sheet is empty, return 1
Public Function LastOccupiedColNum(Sheet As Worksheet) As Long
    Dim lng As Long
    If Application.WorksheetFunction.CountA(Sheet.Cells) <> 0 Then
        With Sheet
            lng = .Cells.Find(What:="*", _
                              after:=.Range("A1"), _
                              Lookat:=xlPart, _
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


