Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Module Summary
'''   A late bound object made to utilize Outlook functionalities from Excel.
'''   Creates a new Outlook session.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Const AppPID As String = "Outlook.Application"

''' Summary 
'''     Constants used by Outlook (Outlook.OlActionResponseStyle)
Public Enum OutlookActionResponseType
    olSave = 0
    olSend = 1
    olPrompt = 2
End Enum

''' Summary 
'''     Constants used by Outlook (Outlook.OlDefaultFolders)
Public Enum OutlookFolderType
    olFolderDeletedItems = 3                'The Deleted Items folder.
    olFolderOutbox = 4                      'The Outbox folder.
    olFolderSentMail = 5                    'The Sent Mail folder.
    olFolderInbox = 6                       'The Inbox folder.
    olFolderCalendar = 9                    'The Calendar folder.
    olFolderContacts = 10                   'The Contacts folder.
    olFolderJournal = 11                    'The Journal folder.
    olFolderNotes = 12                      'The Notes folder.
    olFolderTasks = 13                      'The Tasks folder.
    olFolderDrafts = 16                     'The Drafts folder.
    olPublicFoldersAllPublicFolders = 18    'The All Public Folders folder in the Exchange Public Folders store. Only available for an Exchange account.
    olFolderConflicts = 19                  'The Conflicts folder (subfolder of the Sync Issues folder). Only available for an Exchange account.
    olFolderSyncIssues = 20                 'The Sync Issues folder. Only available for an Exchange account.
    olFolderLocalFailures = 21              'The Local Failures folder (subfolder of the Sync Issues folder). Only available for an Exchange account.
    olFolderServerFailures = 22             'The Server Failures folder (subfolder of the Sync Issues folder). Only available for an Exchange account.
    olFolderJunk = 23                       'The Junk E-Mail folder.
    olFolderRssFeeds = 25                   'The RSS Feeds folder.
    olFolderToDo = 28                       'The To Do folder.
    olFolderManagedEmail = 29               'The top-level folder in the Managed Folders group. For more information on Managed Folders, see the Help in Microsoft Outlook. Only available for an Exchange account.
    olFolderSuggestedContacts = 30          'The Suggested Contacts folder.
End Enum

''' Summary 
'''     Constants used by Outlook (Outlook.OlItemType)
Public Enum OutlookItemType
    olMailItem = 0               'A MailItem object.
    olAppointmentItem = 1        'An AppointmentItem object.
    olContactItem = 2            'A ContactItem object.
    olTaskItem = 3               'A TaskItem object.
    olJournalItem = 4            'A JournalItem object.
    olNoteItem = 5               'A NoteItem object.
    olPostItem = 6               'A PostItem object.
    olDistributionListItem = 7   'A DistListItem object.
End Enum

''' Summary 
'''     A type to reflect the object's self.
Private Type TOutlookApp
    App             As Object
    IsInstalled     As Boolean
    EmailAddress    As String
    Account         As Object
    Accounts        As Variant
    EmailSignature  As String

    UseSignature    As Boolean

End Type

Private this As TOutlookApp

''' Summary:
'''     Initializes the class object
Private Sub Class_Initialize()

    On Error GoTo Err
    ' Try to create an Outlook session
    '   This will error out is Outlook isn't installed
    Set this.App = CreateObject(AppPID)

    ' Set default values if a session was able to be created
    this.IsInstalled = True
    this.UseSignature = True

    ' Store the user's Accounts in this.Accounts
    SetOutlookAccounts

    GoTo Out
Err:
    this.IsInstalled = False
    MsgBox "Outlook isn't installed or is inaccessable.", vbExclamation, "Error"
    
Out:
    On Error GoTo 0
End Sub


''' Summary:
'''     Clears objects from memory
Private Sub Class_Terminate()
    Dim i As Long

    If Not this.App Is Nothing Then Set this.App = Nothing
    If Not this.Account Is Nothing Then Set this.Account = Nothing

    If Not IsEmpty(this.Accounts) Then
        For i = LBound(this.Accounts) To UBound(this.Accounts)
            Set this.Accounts(i) = Nothing
        Next
    End If

End Sub

''' Summary:
'''     If Outlook is installed (read-only).
Public Property Get IsInstalled() As Boolean
    IsInstalled = this.IsInstalled
End Property

''' Summary:
'''     The email address to use.
Public Property Get EmailAddress() As String
    EmailAddress = this.EmailAddress
End Property

Public Property Let EmailAddress(ByVal Value As String)
    this.EmailAddress = Value
    EmailChanged this.EmailAddress
End Property

''' Summary:
'''     The selected email account object (read-only).
'''     This is set when the EmailAddress property is changed.
'''     The default account will be the first account if the email address provided doesn't match an account.
Public Property Get Account() As Object
    Set Account = this.Account
End Property

''' Summary:
'''     All email accounts on the user's Outlook (read-only).
'''     This can be used to populate controls (i.e. a drop down) with the user's accounts.
Public Property Get Accounts() As Variant
    Accounts = this.Accounts
End Property

''' Summary:
'''     The email account's default signature for new emails (read-only).
Public Property Get EmailSignature() As String
    EmailSignature = this.EmailSignature
End Property


''' Summary:
'''     Whether or not to use the signature in the email body.
Public Property Get UseSignature() As Boolean
    UseSignature = this.UseSignature
End Property

Public Property Let UseSignature(ByVal Value As Boolean)
    this.UseSignature = Value
End Property


''' Summary:
'''     Gets all the accounts on the user's Outlook.
Private Sub SetOutlookAccounts()
    Dim i As Long
    
    ReDim this.Accounts(1 To this.App.Session.Accounts.Count)
    For i = 1 To this.App.Session.Accounts.Count
        Set this.Accounts(i) = this.App.Session.Accounts.Item(i)
    Next i
End Sub


''' Summary:
'''     Gets a single Outlook account that's added on the user's Outlook application.
''' Parameters:
'''     ByRef olEmailFrom As Variant: The email address or the Outlook account.
''' Returns:
'''     An Object; The Outlook account.
Public Function GetAccount(ByVal olEmailFrom As String) As Object
    If CheckInstall = False Then Exit Function

    Dim olAcc As Object

    For Each olAcc In this.App.Session.Accounts
        If olAcc = olEmailFrom Then Set GetAccount = olAcc: Exit Function
    Next

    If olAcc Is Nothing Then Set GetAccount = this.App.Session.Accounts.Item(1)
End Function

''' Summary:
'''     Gets the default email signature for a new email from an account.
'''     Has to create and display a draft to extract the signature from the draft's HTML body.
'''     Images are not imported correctly due to the <img>'s src so they are removed.
''' Parameters:
'''     ByRef olEmailFrom As Variant: The email address or the Outlook account.
''' Returns:
'''     A String; The account's email signature as HTML.
Public Function GetSignature(ByRef olEmailFrom As Variant) As String
    If CheckInstall = False Then Exit Function

    Dim olAcc        As Object
    Dim olEmail      As Object
    Dim olEmailBody  As String
    Dim html         As String
    Dim sIndex       As Long: sIndex = 0
    Dim eIndex       As Long: eIndex = 0

    On Error GoTo Err

    ' Get the Outlook account
    '   If an email address was provided the get the account for the email address
    '   If an Outlook account was provided the use that account
    '   Otherwise return an error message
    If VarType(olEmailFrom) = vbString Then
        ' Check if the email address is an empty string
        If Trim(CStr(olEmailFrom)) = vbNullString Then GoTo Err
        Set olAcc = GetAccount(olEmailFrom)
    ElseIf IsObject(olEmailFrom) Then
        Set olAcc = olEmailFrom
    Else
        GoTo Err
    End If

    '   Get HTML email body
    Set olEmail = this.App.CreateItem(olMailItem)
    
    With olEmail
        Set .SendUsingAccount = olAcc
        .Display
        olEmailBody = .HTMLBody
        .Delete
    End With

    '   Remove images from the HTML body
    Do Until InStr(olEmailBody, "img") = 0
        If InStr(olEmailBody, "img") > 0 Then
            sIndex = InStrRev(Left$(olEmailBody, InStr(olEmailBody, "img")), "<p")
            eIndex = InStr(Mid$(olEmailBody, sIndex, Len(olEmailBody)), "</p") + 3

            html = Mid$(olEmailBody, sIndex, eIndex)
            olEmailBody = Replace(olEmailBody, html, vbNullString)
        End If
    Loop
         
    GetSignature = olEmailBody
    GoTo Out

Err:
    GetSignature = ""
    MsgBox "Error: Invalid Email Address Provided", vbExclamation, "Error"

Out:
    Set olAcc = Nothing
    Set olEmail = Nothing

    On Error GoTo 0

End Function

''' Summary:
'''     Sets the Account whenever the Email Address is changed.
Private Sub EmailChanged(ByVal olEmailFrom As String)
    Dim olAcc As Variant
    
    ' Set the Outlook Account
    For Each olAcc In this.Accounts
        If olAcc = olEmailFrom Then
            Set this.Account = olAcc
            Exit For
        End If
    Next
    
    If olAcc Is Nothing Then Set this.Account = this.Accounts(1)

    ' Set the Email Signature for the account
    this.EmailSignature = GetSignature(this.Account)
End Sub


''' Summary:
'''     Creates an Outlook draft.
''' Parameters:
'''     ByVal olEmailBody As String: Content for the HTML email body. Could just be text provided.
'''     ByVal olSubject As String: The email subject.
'''     ByVal olEmailTo As String: The email TO recipients.
'''         Each recipients is delimited by a semi-colon (;).
'''     ByVal olEmailCc As String: The email CC recipients.
'''         Each recipients is delimited by a semi-colon (;).
'''     Optional ByRef olAttachments As Variant: An array of the attachments to include.
'''         Each item in the array is a full file path.
Public Sub CreateDraft(ByVal olEmailBody As String, ByVal olSubject As String, ByVal olEmailTo As String, _
    Optional ByVal olEmailCc As String, Optional ByRef olAttachments As Variant)
    
    If CheckInstall = False Then Exit Sub
    
    Dim olEmail       As Object
    Dim errMsg        As String
    Dim x             As Long

    On Error GoTo Err

    ' Check if the email address set is an empty string
    If Trim(this.EmailAddress) = vbNullString Then
        errMsg = "A valid email address to send from must be set."
        GoTo Err
    End If

    ' Check if an account was successfully found for the email address provided
    If this.Account Is Nothing Then
        errMsg = "The is not an Outlook account with the email address: " & this.EmailAddress
        GoTo Err
    End If
    
    Set olEmail = this.App.CreateItem(olMailItem)
    Set olEmail.SendUsingAccount = this.Account
    
    With olEmail
        .To = olEmailTo

        If Trim(olEmailCc) <> vbNullString Then .CC = olEmailCc
        If Trim(olEmailBody) <> vbNullString Then .HTMLBody = olEmailBody
        If this.UseSignature Then .HTMLBody = .HTMLBody & vbNewLine & this.EmailSignature
    
        .Subject = olSubject
        
        ' Add the attachments
        If Not IsMissing(olAttachments) Then
            For x = LBound(olAttachments) To UBound(olAttachments)
                .Attachments.Add olAttachments(x)
            Next
        End If

        .Recipients.ResolveAll
        .Save
    End With

    GoTo Cleanup

Err:
    MsgBox errMsg, vbExclamation, "Error"

Cleanup:
    Set olEmail = Nothing
    On Error GoTo 0

End Sub


''' Summary:
'''     Checks if Outlook is installed by reference the IsInstalled property.
'''     Displays an error message if Outlook is not installed.
Private Function CheckInstall() As Boolean
    CheckInstall = this.IsInstalled

    If this.IsInstalled = False Then 
        MsgBox "Outlook isn't installed or is inaccessable.", vbExclamation, "Error"
    End If
End Function


