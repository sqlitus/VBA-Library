Attribute VB_Name = "SandboxPP"
Sub Macro1()
    MsgBox "Hello from Macro1"
End Sub


Sub PositionObject()
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then Exit Sub
    With ActiveWindow.Selection.ShapeRange(1)
        .Top = 1.2 * 72     ' inches * pixels per inch
        .Left = (ActivePresentation.PageSetup.SlideWidth - .Width) / 2
        .ZOrder msoSendToBack
    End With
End Sub

