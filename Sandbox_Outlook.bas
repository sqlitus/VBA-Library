Attribute VB_Name = "Sandbox_Outlook"
'https://www.techonthenet.com/excel/formulas/index_vba.php
' SAVE MACROS
' note: display puts signature



Sub CreateRSTReport()
    Dim myItem As Outlook.MailItem
    Dim myRecipient As Outlook.Recipient
    
    Set myItem = Application.CreateItem(olMailItem)
    Set myRecipient = myItem.Recipients.Add("CE.CEN.Retail.Support.Daily.Reports@wholefoods.com")
    myItem.Subject = "RST - " & Format(DateAdd("d", -1, Date), "dddd mm/dd")
    myItem.Display
End Sub


Sub CreateAlohaKeywordsReport()
    Dim myItem As Outlook.MailItem
    Dim myRecipient As Outlook.Recipient
    
    Set myItem = Application.CreateItem(olMailItem)
    Set myRecipient = myItem.Recipients.Add("Stephanie.Stonebraker@wholefoods.com")
    myItem.Subject = "Aloha Keywords Report " & _
        Format(DateAdd("d", 2 - Weekday(Date), DateAdd("d", -7, Date)), "m/d") & " - " & _
        Format(DateAdd("d", 1 - Weekday(Date), Date), "m/d")
    myItem.Display
End Sub


Sub CreateL1HeatmapReport()
    Dim myItem As Outlook.MailItem
    Dim myRecipient As Outlook.Recipient
    
    Set myItem = Application.CreateItem(olMailItem)
    Set myRecipient = myItem.Recipients.Add("CECENRetailSupportL1@wholefoods.com")
    myItem.CC = "Stephanie.Stonebraker@wholefoods.com; Retail.Support.L1.Shift.Leads@wholefoods.com"
    myItem.Subject = "L1 Analyst Metrics Report for " & _
        Format(DateAdd("d", 2 - Weekday(Date), DateAdd("d", -7, Date)), "mm/dd") & " - " & _
        Format(DateAdd("d", 1 - Weekday(Date), Date), "mm/dd")
    myItem.Display
End Sub


Sub CreateGHDRSTReport()
    Dim myItem As Outlook.MailItem
'    Dim myRecipient As Outlook.Recipient
    
    Set myItem = Application.CreateItem(olMailItem)
    With myItem
        .Recipients.Add ("Brian.Rees@wholefoods.com; Paulette.Lindgens@wholefoods.com")
        .CC = "Chris.Talley@wholefoods.com; Paul.Flores@wholefoods.com; Philip.Norman@wholefoods.com"
    End With
    
    myItem.Subject = "GHD CSQ Report - " & Format(DateAdd("d", -1, Date), "dddd mm/dd")
    myItem.Display
End Sub






'''

Sub test()
    Dim myItem As Outlook.MailItem
    Dim myRecipient As Outlook.Recipient
    
    Set myItem = Application.CreateItem(olMailItem)
    Set myRecipient = myItem.Recipients.Add("chris.jabr@wholefoods.com")
    
'    myItem.Body = "something something"
    myItem.Subject = "L1 Analyst Metrics Report for " & _
        Format(DateAdd("d", 2 - Weekday(Date), DateAdd("d", -7, Date)), "mm/dd") & " - " & _
        Format(DateAdd("d", 1 - Weekday(Date), Date), "mm/dd")
    myItem.Display
    
    ' myItem.Body = "something else"    - overrides signature
End Sub


