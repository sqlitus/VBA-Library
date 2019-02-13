Attribute VB_Name = "websiteLogin"
Dim HTMLDoc As HTMLDocument
Dim MyBrowser As InternetExplorer

Sub MyGmail()

    Dim MyHTML_Element As IHTMLElement
    Dim MyURL As String
    On Error GoTo Err_Clear
    
    MyURL = "https://10.2.89.122:8444/cuic/Main.htmx"
    Set MyBrowser = New InternetExplorer
    ' MyBrowser.Silent = True
    MyBrowser.Navigate MyURL
    MyBrowser.Visible = True
    Do
    Loop Until MyBrowser.ReadyState = READYSTATE_COMPLETE
    
    Set HTMLDoc = MyBrowser.Document
    HTMLDoc.all.Email.Value = "teststestestsettest" 'Enter your email id here
    HTMLDoc.all.passwd.Value = "abc+123" 'Enter your password here
    For Each MyHTML_Element In HTMLDoc.getElementsByTagName("input")
        If MyHTML_Element.Type = "submit" Then MyHTML_Element.Click: Exit For
    Next
Err_Clear:
    If Err <> 0 Then
    Err.Clear
    Resume Next
    End If
End Sub



' wiseowl browsing websites example
Sub BrowseToSite()
    
    ' early bind "Microsoft Website Controls"
    Dim IE As New SHDocVw.InternetExplorer ' waits until var is used, if it references something, automatically creates instance
    IE.Visible = True
    IE.Navigate "www.google.com"
    
    ' wait until page loads
    Do While IE.ReadyState <> READYSTATE_COMPLETE
        ' don't need to write any code here
        Application.Wait Now + TimeValue("00:00:01")
        
    Loop
    
    Debug.Print IE.LocationName, IE.LocationURL, IE.Name, IE.AddressBar, IE.Application, IE.Busy
    Debug.Print IE.FullName, IE.FullScreen, IE.Height
End Sub

