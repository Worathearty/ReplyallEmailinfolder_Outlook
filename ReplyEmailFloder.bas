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
                    'Reply to the email
                    Set olMail = objMail
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
                    
                    
                    If intResponse = vbYes Then
                        'Send the reply
                        If blnShowDetails Then
                            olReply.Display
                        Else
                            olReply.Send
                        End If
                    ElseIf intResponse = vbCancel Then
                        Exit Sub
                    End If
                    
                    'Clear the objects
                    Set olMail = Nothing
                    Set olReply = Nothing
                End If
            Next objMail
        End If
    End If
    
    'Clear the objects
    Set olFolder = Nothing
    Set olNs = Nothing
    Set olApp = Nothing
    
End Sub




