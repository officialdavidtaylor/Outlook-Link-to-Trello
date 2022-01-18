'------------------------------------------------------------------------------
' Title:    ott-first-run-setup
' Desc:     checks to see if setup has been run, if not, collects all necessary
'           data for app to function correctly (api stuff, user creds, etc)
' Language: VBA [Outlook for Windows]
'------------------------------------------------------------------------------

Public Function firstRunSetup()
' desc: if app hasn't been run yet, collect all of the necessary parameters from the user

' TODO: ensure user is using Outlook for Windows

  If firstRunComplete Then
    firstRunSetup = True
  Else
    MsgBox("Welcome, and thanks for using the Outlook-to-Trello app." & vbCrLf _
    & "For help completing this setup, please visit the following guide by copying this hyperlink into your browser search bar:" & vbCrLf _
    & "https://github.com/officialdavidtaylor/Outlook-Link-to-Trello/blob/main/SOP.md")

    ' collect user input
    trelloApiKey = InputBox("Please enter your API key:")
    trelloApiToken = InputBox("Please enter your API token:")
    trelloListId = InputBox("Please enter the List ID of the Trello List for us to use when creating new cards:") ' TODO: Collect user credentials instead and just search through the user's available boards and lists
    'trelloBoardId = InputBox("Please enter Board ID:") ' not needed in MVP

    ' TODO: sanitize user input

    ' write trello data to cache
    trelloCacheApiKey(trelloApiKey)
    trelloCacheApiToken(trelloApiToken)
    trelloCacheListId(trelloListId)
    'trelloCacheBoardID(trelloBoardId)

    ' TODO: determine if Outlook is configured to open hyperlinks

    ' TODO: edit registry to enable Outlook to open message hyperlinks

    firstRunSetup = WriteIni("app", "first-run-complete", "true")
  End If
End Function