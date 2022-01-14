'------------------------------------------------------------------------------
' Title:    ott-first-run-setup
' Desc:     checks to see if setup has been run, if not, collects all necessary
'           data for app to function correctly (api stuff, user creds, etc)
' Language: VBA [Outlook for Windows]
'------------------------------------------------------------------------------

Public Function firstRunSetup()
' desc: if app hasn't been run yet, collect all of the necessary parameters from the user

  If firstRunComplete Then
    firstRunSetup = True
  Else
    MsgBox "Welcome, and thanks for using the Outlook-to-Trello app.\nFor help completing this setup, please visit the following guide by copying this hyperlink into your browser search bar:\nhttps://github.com/officialdavidtaylor/Outlook-Link-to-Trello/blob/main/SOP.md"

    ' TODO: Sanitize user inputs
    trelloApiKey = InputBox("Please enter your API key:")
    trelloApiToken = InputBox("Please enter your API token:")
    trelloListId = InputBox("Please enter the List ID of the Trello List for us to use when creating new cards:")

    firstRunSetup = WriteIniFileString("app", "first-run-complete", "true")
  End If
End Function