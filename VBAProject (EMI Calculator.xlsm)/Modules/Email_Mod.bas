Attribute VB_Name = "Email_Mod"
Sub EmailReport()

'   Outlook must be configured

'   Refer to library "Microsoft Office 14.0 Object Library", if not
'   Tools -> References

'   vbCrLf -> Use for Next Line or Carriage Return


'   If error, then go to ErrHandler Label
    On Error GoTo ErrHandler
    
    Call SaveReport_Mod.SaveReport
    
    Dim outlookApp As Object
'   Creating object of type Outlook.Application to connect with Microsoft Outlook
    Set outlookApp = CreateObject("Outlook.application")

'   If there is a default profile set, then execute next
    If Not outlookApp.DefaultProfileName = "" Then
    
'       If there is at least one account configured in Microsoft Outlook
        If outlookApp.Session.Accounts.Count > 0 Then

            Dim oMail As Object
'           Creating an object to handle mailing
            Set oMail = outlookApp.CreateItem(olMailItem)

'           Setting the contents of mail options
            oMail.to = Mail_Ufm.mailto_tbx.Value
            oMail.Subject = Mail_Ufm.subject_tbx.Value
            oMail.Body = Mail_Ufm.body_tbx.Value
            
'           Adding attachment
'           ThisWorkbook.Path returns the path of current file
            oMail.Attachments.Add (ThisWorkbook.Path & "\Report_" & VBA.Format(Now, "ddmmyyyy-hhmm") & ".pdf")
            oMail.Send
        Else
            MsgBox ("Microsoft Outlook is not configured.")

        End If
    Else
        MsgBox ("Microsoft Outlook is not configured.")
    End If

'   Setting objects to Nothing to avoid errors in the next run of this code
    Set outlookApp = Nothing:   Set oMail = Nothing
    
'As there is nothing present in ErrHandler, it will work same as "On Error GoTo 0"
ErrHandler:
    
End Sub
