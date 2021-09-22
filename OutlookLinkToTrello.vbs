'------------------------------------------------------------------------------
' Title:    Outlook Link With Trello
' Desc:     Uploads the selected mail item from Outlook to a designated Trello
'           board, creating a backlink that allows the user to open the
'           original email item directly from the Trello card.
' Language: VBA [Outlook]
'------------------------------------------------------------------------------

'------STRUCTURES------

' To package all necessary data before making HTTP requests to create Trello Card
Type CardPayload

    ' boardID As String ' This is not necessary for card creation
    listID As String
    cardID As String
    cardName As String
    sender As String
    subject As String
    mailUID As String
    conversationID As String
    receivedTime As String
    ' TODO: Add file attachment support in the future: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/file-object

    token As String
    key As String

End Type

' Structure for token and key cache
Type TrelloCredentialCache

    token As String
    key AS String
    username As String

End Type

'------CONSTANTS------
    Public Const LIST_ID_LENGTH As Integer = 24

'------SUBS------

' TODO: Finalize this method as the "main" method
Sub outlookLinkToTrello()

    'TODO: First Run Setup logic
    ' firstRunSetup()

    ' VARIABLE DECLARATION
    Dim objMail As Outlook.MailItem ' Create new Outlook MailItem object
    Dim cardPayload As CardPayload  ' Instantiate new email structure
    Dim responseText As String      ' this variable to contain the text returned from the server after any HTTP request

    'One and ONLY one message muse be selected
    If Application.ActiveExplorer.Selection.Count <> 1 Then
        MsgBox ("Select one and ONLY one message.") ' TODO: Enable grouping and/or batch processing
        Exit Sub
    End If

    Set objMail = Application.ActiveExplorer.Selection.Item(1)
    
    ' transfer selected-mail data into email object
    cardPayload.mailUID = "outlook:" + objMail.EntryID
    cardPayload.sender = objMail.Sender
    cardPayload.subject = objMail.Subject
    cardPayload.conversationID = objMail.ConversationID
    cardPayload.receivedTime = Format(objMail.ReceivedTime, "yyyymmddhhnn")
    cardPayload.listID = getCachedListID()
    cardPayload.key = getCachedKey()
    cardPayload.token = getCachedToken()

    ' Collect Card Name from InputBox
    ' TODO: add error checking
    cardPayload.cardName = InputBox("Please enter Card name here:")

    responseText = trelloCreateCard(cardPayload)
    cardPayload.cardID = extractCardID(responseText)

    MsgBox cardPayload.cardID

End Sub

'------FUNCTIONS------

' Create new Card with CardPayload object as input
Function trelloCreateCard(cardPayload As CardPayload) As String
    ' Use this method to create a card with custom fields and attachments, and to provide
    ' useful feedback in the event of an operation failure.
    ' https://developer.atlassian.com/cloud/trello/rest/api-group-actions/
    ' https://stackoverflow.com/questions/158633/how-can-i-send-an-http-post-request-to-a-server-from-excel-using-vba
    ' https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms762278(v=vs.85)?redirectedfrom=MSDN

    ' Data needed before creating card
    Dim cardApiUrl As String
    Dim responseText As String ' to be returned from POST request
    Dim payload As String ' variable (in JSON format) to contain all of the parts required for the POST request

    cardApiUrl = "https://api.trello.com/1/cards" ' URL for Trello API calls for Cards

    ' Generate the payload
    payload = "{""name"":""" & cardPayload.cardName & """, ""idList"":""" & cardPayload.listID & """, ""key"":""" & cardPayload.key & """, ""token"":""" & cardPayload.token & """, ""pos"":""top""}" 

    ' Initiate HTTP interface object
    ' https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms754586(v=vs.85)
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")

    objHTTP.open "POST", cardApiUrl, False ' stage a new POST request
    objHTTP.SetRequestHeader "Content-type", "application/json" ' tell server what format to expect payload (JSON)
    objHTTP.send payload ' send POST request to create card

    responseText = objHTTP.responseText ' Server will return the newly created cardID, which will come in handy later

    ' return server response
    trelloCreateCard = responseText

End Function

' Extract CardID from server response upon card creation
Function extractCardID(responseText As String) As String

    ' VARIABLE DECLARATION
    Dim listIdOffset As Integer

    listIdOffset = (InStr(responseText, """id"":*""") + 8) ' Determine where the List ID is contained in the string

    ' return new CardID
    extractCardID = Mid(responseText, listIdOffset, LIST_ID_LENGTH)

End Function

' TODO: Retrieve ListID from cache
Function getCachedListID() As String
    getCachedListID = ""
End Function

' TODO: Retrieve Key from cache
Function getCachedKey() As String
    getCachedKey = ""
End Function

' TODO: Retrieve Token from cache
Function getCachedToken() As String
    getCachedToken = ""
End Function

' TODO: [V2.0] Setup function to configure all necessary parameters
Function firstRunSetup() as Boolean
    ' Start by checking the registry to see if the system is configured to open "Outlook:" hyperlinks in outlook
    ' If Not checkRegistryKeysForOutlookHyperlinking()
    '     addRegistryKeysForOutlookHyperlinking()
    
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