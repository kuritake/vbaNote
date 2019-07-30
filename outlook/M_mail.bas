Attribute VB_Name = "M_mail"
Option Explicit
Sub mail()

Dim objOutlook          As Outlook.Application
Dim objMail             As Outlook.MailItem

Dim bodyWs              As Worksheet
Dim i                   As Long
Dim mailbody            As String


Set bodyWs = ThisWorkbook.Worksheets("body")

Set objOutlook = CreateObject("Outlook.Application")
Set objMail = objOutlook.CreateItem(olMailItem)
    With objMail
        .To = "enterYourMailaddress@XXXXXX.XXXXXX"
        .Subject = "This is a test mail."
            
        mailbody = ""
            'get data from "A" column
            For i = 1 To bodyWs.Cells(Rows.Count, 1).End(xlUp).Row
                If bodyWs.Cells(i, 1).Value = "" Then
                    mailbody = mailbody & vbCrLf
                Else
                    mailbody = mailbody & bodyWs.Cells(i, 1).Value & vbCrLf
                End If
            Next
        .ReadReceiptRequested = True
        .Body = mailbody
        .BodyFormat = 1
        .Save '
        .Close 0
    End With
Set objMail = Nothing
Set objOutlook = Nothing

End Sub
