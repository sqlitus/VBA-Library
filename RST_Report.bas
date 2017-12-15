Attribute VB_Name = "RST_Report"
Option Compare Text



Sub Run_RST_Report()
    Application.ScreenUpdating = False
    On Error GoTo ErrHandler
       
    Call RSTReport_CheckIfCSQARReport
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
    
    Dim findCell As Range: Set findCell = Cells.Find("*by CCX*", Cells(Rows.count, Columns.count))
    
    If findCell Is Nothing Then
        Application.ScreenUpdating = True
        MsgBox "NOT CSQAR REPORT. MACRO NOT RUN"
        End
    End If
    
End Sub



Private Sub RSTReport_SaveToFolders()
    ' check if directory subfolders exist, and create if necessary. Then save file w/ proper name

    Dim fdObj As Object
    Dim dir As String
    Dim yearFolder As String
    Dim monthFolder As String
    Dim reportDate As Date
    
    ' date of report - (yesterday for CSQAR) ************ test change *************
    reportDate = DateAdd("d", -1, Date)
    
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
    
    Set tableStart = Cells.Find("*", Cells(Rows.count, Columns.count))
    Set table = Range(tableStart, Cells(tableStart.End(xlDown).Row, tableStart.End(xlToRight).column))
    
    ' delete unneeded columns
    Dim i As Integer
    
    For i = table.Rows(1).Columns.count To 1 Step -1
        If table.Rows(1).Columns(i) Like "csq id" _
            Or table.Rows(1).Columns(i) Like "skills" _
            Or table.Rows(1).Columns(i) Like "avg abandon per day" _
            Or table.Rows(1).Columns(i) Like "calls handled by other" _
        Then table.Rows(1).Columns(i).EntireColumn.Delete
    Next
    
    ' rename columns & row values
    On Error Resume Next
        table.Rows(1).Columns.Find("csq name").Value = "Queue"
        table.Rows(1).Columns.Replace "dequeued", "to Voicemail"
        table.Rows(1).Columns.Replace "dequeue", "Voicemail"
        table.Columns(1).Rows.Replace "OPOS_", ""
    On Error GoTo 0
    
    
    
    ' conditionally highlight cells in table
    Dim check1 As Range
    Dim check2 As Range
    Dim checkflag As Integer: checkflag = 0
    
    Set check1 = Cells(table.Columns(1).Rows.Find("r10").Row, table.Rows(1).Columns.Find("avg queue time").column)
    
    With check1
        If Minute(.Value) < 2 Then
            .Interior.Color = RGB(146, 208, 80)
        Else
            .Interior.Color = RGB(255, 0, 0)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
        End If
    End With
    
    
    Set check1 = Cells(table.Columns(1).Rows.Find("r10").Row, table.Rows(1).Columns.Find("calls abandoned").column)
    Set check2 = Cells(table.Columns(1).Rows.Find("r10").Row, table.Rows(1).Columns.Find("calls presented").column)
    
    With check1
        If .Value / check2.Value < 0.1 Then
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
    If checkflag = 1 Then MsgBox "Breach SLA: Mail to RST.L1.Leadership@wholefoods.com first", Title:="Mailto Steph"
    
End Sub
