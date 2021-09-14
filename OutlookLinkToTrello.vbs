'------------------------------------------------------------------------------
' Title:    Outlook Link With Trello
' Desc:     Uploads the selected mail item from Outlook to a designated Trello
'           board, creating a backlink that allows the user to open the
'           original email item directly from the Trello card.
' Language: VBA [Outlook]
'------------------------------------------------------------------------------

'------STRUCTURES------
' consider using XML? https://stackoverflow.com/questions/11305/how-to-parse-xml-using-vba 

' Mail data structure
Type EmailData

    hyperlink As String
    sender As String
    subject As String
    conversationID As String
    receivedTime As String
    ' TODO: Add file attachment support in the future: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/file-object

End Type

' To package all necessary data before making HTTP requests to create Trello Card
Type CardPayload

    boardID As String
    listID AS String
    sender As String
    subject As String
    mailUID As String
    conversationID As String

    token As String
    key As String

End Type

' Structure for token and key cache
Type TrelloCredentialCache

    token As String
    key AS String
    username As String

End Type

' ' Error Codes
' Type ErrorCodes

'     e0 As String
'     e1 As String
'     e2 As String
'     e3 As String
'     e4 As String
'     e5 As String
'     e6 As String
'     e7 As String
'     e8 As String
'     e9 As String

' End Type


' '------GLOBAL VARIABLES------

' ' Define Error Structure
' Dim errorCodes As ErrorCodes
' errorCodes.e0 = ""
' errorCodes.e1 = ""
' errorCodes.e2 = ""
' errorCodes.e3 = ""
' errorCodes.e4 = ""
' errorCodes.e5 = ""
' errorCodes.e6 = ""
' errorCodes.e7 = ""
' errorCodes.e8 = ""
' errorCodes.e9 = ""


'------SUBS------

' TODO: Finalize this method as the "main" method
Sub outlookLinkToTrello()

    'TODO: First Run Setup logic
    ' firstRunSetup()

    ' VARIABLE DECLARATION
    Dim objMail As Outlook.MailItem ' Create new Outlook MailItem object
    Dim email As EmailData          ' Instantiate new email structure

    'One and ONLY one message muse be selected
    If Application.ActiveExplorer.Selection.Count <> 1 Then
        MsgBox ("Select one and ONLY one message.") ' TODO: Enable grouping and/or batch processing
        Exit Sub
    End If

    Set objMail = Application.ActiveExplorer.Selection.Item(1)
    
    ' transfer selected-mail data into email object
    email.hyperlink = "outlook:" + objMail.EntryID
    email.sender = objMail.Sender
    email.subject = objMail.Subject
    email.conversationID = objMail.ConversationID
    email.receivedTime = Format(objMail.ReceivedTime, "yyyymmddhhnn")

    trelloCreateCard(email)

End Sub

'------FUNCTIONS------

' TODO: Create new Card
Function trelloCreateCard(email as EmailData) as Boolean
    ' Use this method to create a card with custom fields and attachments, and to provide
    ' useful feedback in the event of an operation failure.
    ' https://docs.microsoft.com/en-us/windows/win32/winhttp/winhttprequest?redirectedfrom=MSDN



End Function

' TODO: [V2.0] Setup function to configure all necessary parameters
Function firstRunSetup() as Boolean
    ' Start by checking the registry to see if the system is configured to open "Outlook:" hyperlinks in outlook
    If Not checkRegistryKeysForOutlookHyperlinking()
        addRegistryKeysForOutlookHyperlinking()
    
    ' trelloCacheBoardID(trelloFindBoardID()) ' Walk user through finding BoardID, then cache it
    ' trelloCacheListID(trelloFindListID()) ' Walk user through finding ListID, then cache it

End Function

' TODO: [V2.0] Get Cached Info
Function getCacheInfo()
    ' Return the cached info requested
End Function

' TODO: [V2.0] Registry content verification: Ensure handling of "outlook:" hyperlinks
Function checkRegistryKeysForOutlookHyperlinking()
    ' need to check for the right keys and return the status as a Boolean
End Function

' TODO: [V2.0] Registry key addition to enable proper handling of "outlook:" hyperlinks
Function addRegistryKeysForOutlookHyperlinking()
    ' need to add for the right keys and return the status as a Boolean
End Function

' TODO: [V2.0] Determine if Trello API Key and Token have been collected
Function trelloCheckForCredentials()
    ' if not, provide quick dialog box to collect it
End Function

' TODO: [V2.0] Cache Trello User ID
Function trelloCacheUserID()
    ' save the user ID for API purposes
    ' https://stackoverflow.com/questions/12428293/best-way-to-cache-a-password-in-an-excel-vba-function
End Function

' TODO: [V2.0] Cache Trello Credentials
Function trelloCacheCredentials()
    ' https://docs.microsoft.com/en-us/office/vba/word/concepts/miscellaneous/storing-values-when-a-macro-ends
End Function

' TODO: [V2.0] Find Board ID in Trello
Function trelloFindBoardID() As String
    ' https://stackoverflow.com/questions/26552278/trello-api-getting-boards-lists-cards-information/50908600
End Function

' TODO: [V2.0] Cache Trello Board ID
Function trelloCacheBoardID(boardID As String)
    ' https://docs.microsoft.com/en-us/office/vba/word/concepts/miscellaneous/storing-values-when-a-macro-ends
End Function

' TODO: [V2.0] Retrieve Trello Board ID from cache
Function trelloGetBoardID() As String
    ' https://docs.microsoft.com/en-us/office/vba/word/concepts/miscellaneous/storing-values-when-a-macro-ends
End Function

' TODO: [V2.0] Find List ID in Trello
Function trelloFindListID()
    ' https://stackoverflow.com/questions/26552278/trello-api-getting-boards-lists-cards-information/50908600
End Function

' TODO: [V2.0] Cache Trello List ID
Function trelloCacheListID(boardID As String)
    ' https://docs.microsoft.com/en-us/office/vba/word/concepts/miscellaneous/storing-values-when-a-macro-ends
End Function

' TODO: [V2.0] Retrieve Trello List ID from cache
Function trelloGetListID() As String
    ' https://docs.microsoft.com/en-us/office/vba/word/concepts/miscellaneous/storing-values-when-a-macro-ends
End Function

' TODO: [V2.0] Delete all cached data pertaining to Trello
Function trelloClearCache()
    ' https://docs.microsoft.com/en-us/office/vba/word/concepts/miscellaneous/storing-values-when-a-macro-ends
End Function