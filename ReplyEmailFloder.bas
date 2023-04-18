Attribute VB_Name = "Module1"
Sub ReplyAllAndAddText()

    Dim olApp As Outlook.Application
    Dim olNs As Outlook.NameSpace
    Dim olFolder As Outlook.MAPIFolder
    Dim olMail As Outlook.MailItem
    Dim olReply As Outlook.MailItem
    Dim olRecip As Outlook.Recipient
    Dim objMail As Object
    Dim objFolder As MAPIFolder
    Dim i As Integer
    Dim strText As String
    Dim strSubject As String
    Dim strBodyText As String
    Dim strNewBodyText As String
    Dim intResponse As Integer
    Dim blnShowDetails As Boolean
    Dim dictRepliedEmails As Object
    
    Set dictRepliedEmails = CreateObject("Scripting.Dictionary")
    
    'Get the body text and subject from user input
    strSubject = InputBox("Enter the subject for the reply emails")
    strBodyText = InputBox("Enter the body text to be added to the reply emails")
    
    'Select the folder to reply to emails in
    Set olApp = Outlook.Application
    Set olNs = olApp.GetNamespace("MAPI")
    Set olFolder = olNs.PickFolder
    
    'Check if a folder was selected
    If Not olFolder Is Nothing Then
        'Prompt the user before replying to the emails
        intResponse = MsgBox("Do you want to reply to all emails in the folder " & olFolder.Name & "?", vbYesNo + vbQuestion, "Reply to Emails")
        
        If intResponse = vbYes Then
            'Loop through each email in the folder
            For Each objMail In olFolder.Items
                If TypeOf objMail Is MailItem Then
                    'Check if the email has already been replied to
                    Set olMail = objMail
                    If Not dictRepliedEmails.exists(olMail.EntryID) Then
                        'Reply to the email
                        Set olReply = olMail.ReplyAll
                        
                        'Preserve the "RE:" prefix in the subject line
                        If InStr(1, olMail.Subject, "RE:", vbTextCompare) = 1 Then
                            olReply.Subject = olMail.Subject
                        Else
                            olReply.Subject = "RE: " & olMail.Subject
                        End If
                        
                        'Add the specified body text to the existing body
                        strNewBodyText = vbCrLf & strBodyText & vbCrLf & olMail.Body
                        olReply.Body = strNewBodyText
                        
                        'Send the reply
                        If blnShowDetails Then
                            olReply.Display
                        Else
                            olReply.Send
                        End If
                        
                        'Add the email to the dictionary of replied emails
                        dictRepliedEmails.Add olMail.EntryID, True
                    End If
                End If
            Next objMail
        End If
    End If
    
    'Clear the objects
    Set olFolder = Nothing
    Set olNs = Nothing
    Set olApp = Nothing

End Sub

 



