'------------------------------------------------------------------------------
' Title:    Outlook-to-Trello
' Desc:     Uploads the selected mail item from Outlook to a designated Trello
'           board, creating a backlink that allows the user to open the
'           original email item directly from the Trello card.
' Language: VBA [Outlook for Windows]
'------------------------------------------------------------------------------

'------CONSTANTS------
Public Const LIST_ID_LENGTH As Integer = 24
Public Const MAX_LOOP_ITERATIONS As Integer = 500
Public Const HKEY_CLASSES_ROOT  = &H80000000
Public Const HKEY_LOCAL_MACHINE = &H80000002

' var for config file path
Public configFilePath As String

'------SUBS------

' TODO: Finalize this method as the "main" method
Sub OutlookToTrello()

    ' initialization
    ' ensure the config path variable is correct
    configFilePath = getConfigFilePath()

    ' ensures that all necessary parameters are known before running the following code
    firstRunSetup

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

    ' handle "cancel" condition, exit macro
    If cardPayload.cardName = "" Then
        Exit Sub
    End If

    trelloCreateCard cardPayload ' Create new card based on the data in cardPayload

End Sub

' Sub to put mail hyperlink in Clipboard
Sub emailUrlToClipboard()

    ' VARIABLE DECLARATION
    Dim clipboard As MSForms.DataObject
    Set clipboard = New MSForms.DataObject

    Dim objMail As Outlook.MailItem ' Create new Outlook MailItem object
    Dim mailUID As string ' create string to hold the URL

    'One and ONLY one message muse be selected
    If Application.ActiveExplorer.Selection.Count <> 1 Then
        MsgBox ("Select one and ONLY one message.")
        Exit Sub
    End If

    Set objMail = Application.ActiveExplorer.Selection.Item(1)
    
    ' put URL in clipboard
    clipboard.SetText ("outlook:" + objMail.EntryID)
    clipboard.PutInClipboard
End Sub

'------FUNCTIONS------

' Extract CardID from server response upon card creation
Function extractCardID(responseText As String) As String

    ' VARIABLE DECLARATION
    Dim listIdOffset As Integer

    listIdOffset = (InStr(responseText, """id"":*""") + 8) ' Determine where the List ID is contained in the string

    ' return new CardID
    extractCardID = Mid(responseText, listIdOffset, LIST_ID_LENGTH)

End Function

