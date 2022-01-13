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
Public Const HKEY_CLASSES_ROOT  = &H80000000
Public Const HKEY_LOCAL_MACHINE = &H80000002

'------SUBS------

' TODO: Finalize this method as the "main" method
Sub OutlookToTrello()

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

' TODO: [V2.0] Get Cached Info
Function getCacheInfo()
    ' Return the cached info requested
End Function

' Check Registry for Key, return as Boolean
Function checkRegistryForKey(ByVal regItem As RegistryItem) As Boolean

    Dim oReg: Set oReg = GetObject("winmgmts:!root/default:StdRegProv")

    If oReg.EnumKey(HKEY_CLASSES_ROOT, regItem.path, "", "") = 0 Then
        checkRegistryForKey = True
    Else
        checkRegistryForKey = False
    End If

End Function

' Find Data in Registry Key
Function getRegistryKeyData(ByVal Key As String, ByVal KeyPath As String) As String

End Function

' Check to see if outlook hyperlinking is enabled in the registry
Sub checkOutlookHyperlinkingStatus()
    ' To modify the registry, see this article: https://docs.microsoft.com/en-us/windows/win32/wmisdk/obtaining-registry-data

    ' more registry notes and code: https://docs.microsoft.com/en-us/office/vba/word/concepts/miscellaneous/storing-values-when-a-macro-ends
    ' Sub GetRegistryInfo() 
    ' Dim strSection As String 
    ' Dim strPgmDir As String 
    ' strSection = "HKEY_CURRENT_USER\Software\Microsoft" _ 
    ' & "\Office\12.0\Word\Options" 
    ' strPgmDir = System.PrivateProfileString(FileName:="", _ 
    ' Section:=strSection, Key:="PROGRAMDIR") 
    ' MsgBox "The directory for Word is - " & strPgmDir 
    ' End Sub

    ' Outlook's backend enables hyperlinking as a legacy feature, if the correct keys exist in the registry
    ' The necessary registry structure is as follows:
    ' - HKEY_CLASSES_ROOT\outlook
    ' -- (Default) : "URL:Outlook Folders"
    ' -- URL Protocol : ""
    ' - HKEY_CLASSES_ROOT\outlook\DefaultIcon
    ' -- (Default) : """C:\Program Files\Microsoft Office\root\Office16\1033\OUTLLIBR.DLL"", -9403"
    ' - HKEY_CLASSES_ROOT\outlook\shell
    ' -- (Default) : (value not set)
    ' - HKEY_CLASSES_ROOT\outlook\shell\open
    ' -- (Default) : ""
    ' - HKEY_CLASSES_ROOT\outlook\shell\open\command
    ' -- (Default) : """C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"" /select ""%1"""

    Dim outlookExeRegistryPath As String

    outlookExeRegistryPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE"

End Sub

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