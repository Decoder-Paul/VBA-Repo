Attribute VB_Name = "OutlookOpenDL"
'========================================================================================================
' MyMacroThatUseOutlook
' -------------------------------------------------------------------------------------------------------
' Purpose   :   To open Outlook Application And passing outlook object to
'               'ExtractFirstUnreadEmailDetails' procedure
'
' Author    :   Subhankar Paul, 24th February, 2017
' Notes     :   N/A
'
' Parameter :   N/A
' Returns   :   N/A
' ---------------------------------------------------------------
' Revision History
'========================================================================================================
#Const LateBind = True

Const olMinimized As Long = 1
Const olMaximized As Long = 2
Const olFolderInbox As Long = 6

#If LateBind Then

Public Function OutlookApp( _
    Optional WindowState As Long = olMinimized, _
    Optional ReleaseIt As Boolean = False _
    ) As Object
    Static o As Object
#Else
Public Function OutlookApp( _
    Optional WindowState As Outlook.OlWindowState = olMinimized, _
    Optional ReleaseIt As Boolean _
) As Outlook.Application
    Static o As Outlook.Application
#End If
On Error GoTo ErrHandler
 
    Select Case True
        Case o Is Nothing, Len(o.Name) = 0
            Set o = GetObject(, "Outlook.Application")
            If o.Explorers.Count = 0 Then
InitOutlook:
                'Open inbox to prevent errors with security prompts
                o.Session.GetDefaultFolder(olFolderInbox).Display
                o.ActiveExplorer.WindowState = WindowState
            End If
        Case ReleaseIt
            Set o = Nothing
    End Select
    Set OutlookApp = o
 
ExitProc:
    Exit Function
ErrHandler:
    Select Case Err.Number
        Case -2147352567
            'User cancelled setup, silently exit
            Set o = Nothing
        Case 429, 462
            Set o = GetOutlookApp()
            If o Is Nothing Then
                Err.Raise 429, "OutlookApp", "Outlook Application does not appear to be installed."
            Else
                Resume InitOutlook
            End If
        Case Else
            MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Unexpected error"
    End Select
    Resume ExitProc
    Resume
End Function

#If LateBind Then
Private Function GetOutlookApp() As Object
#Else
Private Function GetOutlookApp() As Outlook.Application
#End If
On Error GoTo ErrHandler
    
    Set GetOutlookApp = CreateObject("Outlook.Application")
    
ExitProc:
    Exit Function
ErrHandler:
    Select Case Err.Number
        Case Else
            'Do not raise any errors
            Set GetOutlookApp = Nothing
    End Select
    Resume ExitProc
    Resume
End Function

Sub MyMacroThatUseOutlook()
    Dim OutApp  As Object
    Set OutApp = OutlookApp()
    'Automate OutApp as desired
    'Application.Wait (Now + TimeValue("0:00:10"))
    Call ExtractFirstUnreadEmailDetails(OutApp)
End Sub

Sub ExtractFirstUnreadEmailDetails(OutlookObj As Object)
    Dim oOlAp As Object, oOlns As Object, oOlInb As Object
    Dim oOlItm As Object
    Dim AttachmentPath As String
    AttachmentPath = ThisWorkbook.Path & "\Master\"

    '~~> New File Name for the attachment
    Dim NewFileName As String
    NewFileName = AttachmentPath & Format(Date, "DD-MM-YYYY") & "-"
    '~~> Outlook Variables for email
    Dim eSender As String, dtRecvd As String, dtSent As String
    Dim sSubj As String, sMsg As String
    
    
    '~~> Get Outlook instance
    'Set oOlAp = GetObject(, "Outlook.application")
    Set oOlAp = OutlookObj
    Set oOlns = oOlAp.GetNamespace("MAPI")
    Set oOlInb = oOlns.GetDefaultFolder(olFolderInbox)

    '~~> Check if there are any actual unread emails
    If oOlInb.Items.Restrict("[UnRead] = True").Count = 0 Then
        MsgBox "NO Unread Email In Inbox"
        Exit Sub
    Else
        Debug.Print oOlInb.Items.Restrict("[UnRead] = True").Count
    End If
    '~~> Store the relevant info in the variables
    For Each oOlItm In oOlInb.Items.Restrict("[Unread] = true")
        eSender = oOlItm.SenderEmailAddress
        dtRecvd = oOlItm.ReceivedTime
        dtSent = oOlItm.CreationTime
        sSubj = oOlItm.Subject
        sMsg = oOlItm.Body
        sEmailType = oOlItm.SenderEmailType
        'oOlItm.UnRead = False
        'Debug.Print eSender
        Debug.Print dtRecvd
        'Debug.Print dtSent
        Debug.Print sSubj
        'Debug.Print sMsg
        Debug.Print oOlInb.Items.Restrict("[UnRead] = True").Count
        eSender = GetSenderSMTPAddress(oOlItm)
        Debug.Print eSender
        
        If eSender = "Subhankar.Paul@trianz.com" Then
        
            If oOlItm.Attachments.Count <> 0 Then
                For Each oOlAtch In oOlItm.Attachments
                    '~~> Download the attachment
                    oOlAtch.SaveAsFile NewFileName & oOlAtch.Filename
                Next
            Else
                MsgBox "No attachment Found"
            End If
        End If
        Debug.Print "Unread Count" & oOlInb.Items.Restrict("[UnRead] = True").Count
    Next
    
End Sub
Private Function GetSenderSMTPAddress(mail As Outlook.MailItem) As String

If mail Is Nothing Then
    GetSenderSMTPAddress = vbNullString
    Exit Function
End If
If mail.SenderEmailType = "EX" Then
    Dim sender As Outlook.AddressEntry
    Set sender = mail.sender
    If Not sender Is Nothing Then
        'Now we have an AddressEntry representing the Sender
        If sender.AddressEntryUserType = Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry Or sender.AddressEntryUserType = Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry Then
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


