Option Explicit

'Duplicate this routine to create different trigger subjects.
Sub ExampleFlowStart()
    'Modify this line to create a different trigger subject.
    Dim triggerSubject As String: triggerSubject = "Loungers data erasure request - mn"
    
    'Leave everything below this line as is.
    Dim myOlExp As Outlook.Explorer: Set myOlExp = Application.ActiveExplorer
    Dim myOlSel As Outlook.Selection: Set myOlSel = myOlExp.Selection
    Dim olItem As Outlook.MailItem
    Dim x As Integer
    For x = 1 To myOlSel.Count
        Set olItem = myOlSel.Item(x)
        If olItem.Class = 43 Then
            Call triggerFlowForSelectedMessages(olItem, triggerSubject)
        End If
    Next x
End Sub

'Main routine that sends the email to fire the Power Automate trigger
Sub triggerFlowForSelectedMessages(olItem As Outlook.MailItem, triggerSubject As String)

    Dim strFolder As String, JSON As String, sender As String
    Dim messageFolder As Outlook.Folder
    Dim oPA As Outlook.PropertyAccessor: Set oPA = olItem.PropertyAccessor
    
    Const PR_INTERNET_MESSAGE_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"

    'Find and convert the folder path
    Set messageFolder = olItem.Parent
    strFolder = Replace(Mid(messageFolder.FolderPath, (InStr(Mid(messageFolder.FolderPath, 3), "\") + 3)), "\", "/")
     
    'Definition of JSON that will be sent to trigger the flow
    JSON = "{" & vbLf & _
    "'from': '{{from}}'," & vbLf & _
    "'subject': '{{subject}}'," & vbLf & _
    "'internetMessageID': '{{internetMessageID}}'," & vbLf & _
    "'folder': '{{messageFolder}}'," & vbLf & _
    "'Name': '{{from}}'" & vbLf & _
    "}"
        
    'Replace JSON values
    JSON = Replace(JSON, "{{from}}", GetSenderSMTPAddress(olItem))
    JSON = Replace(JSON, "{{subject}}", olItem.Subject)
    JSON = Replace(JSON, "{{internetMessageID}}", oPA.GetProperty(PR_INTERNET_MESSAGE_ID))
    JSON = Replace(JSON, "{{messageFolder}}", messageFolder)
    JSON = Replace(JSON, "'", Chr(34))
    'Debug.Print (JSON)

     
    'Send the message that Triggers the flow
    Dim objMsg As MailItem
    Set objMsg = Application.CreateItem(olMailItem)
        
    With objMsg
        .To = "dpl@loungers.co"
        .Subject = triggerSubject
        .BodyFormat = olFormatPlain
        .Body = JSON
        .Send
    End With
     
End Sub

'Function to get the sender address of an email. Required for exchange accounts.
Function GetSenderSMTPAddress(mail As Outlook.MailItem) As String
    Const PR_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
    If mail Is Nothing Then
        GetSenderSMTPAddress = vbNullString
        Exit Function
    End If
    If mail.SenderEmailType = "EX" Then
        Dim sender As Outlook.AddressEntry
        Set sender = mail.sender
        If Not sender Is Nothing Then
            'Now we have an AddressEntry representing the Sender
            If sender.AddressEntryUserType = _
                    Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry Or _
                    sender.AddressEntryUserType = _
                    Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry Then
                'Use the ExchangeUser object PrimarySMTPAddress
                Dim exchUser As Outlook.ExchangeUser
                Set exchUser = sender.GetExchangeUser()
                If Not exchUser Is Nothing Then
                     GetSenderSMTPAddress = exchUser.PrimarySmtpAddress
                Else
                    GetSenderSMTPAddress = vbNullString
                End If
            Else
                 GetSenderSMTPAddress = sender.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
            End If
        Else
            GetSenderSMTPAddress = vbNullString
        End If
    Else
        GetSenderSMTPAddress = mail.SenderEmailAddress
    End If
End Function

'Function to get the Primary email address of the current user
Function GetUserEmailAddress()
    Dim outApp As Outlook.Application, outSession As Object, curr
    Set outApp = CreateObject("Outlook.Application")
    Set outSession = outApp.Session.CurrentUser
    Set outApp = Nothing
    GetUserEmailAddress = outSession.AddressEntry.GetExchangeUser().PrimarySmtpAddress
End Function

