'------------------------------------------------------------------------------
' Title:    ott-first-run
' Desc:     help user configure upon first run
' Language: VBA [Outlook for Windows]
'------------------------------------------------------------------------------

' TODO: [V2.0] Setup function to configure all necessary parameters
Public Function firstRunSetup() as Boolean
    ' Start by checking the registry to see if the system is configured to open "Outlook:" hyperlinks in outlook
    ' If Not checkRegistryKeysForOutlookHyperlinking()
    '     addRegistryKeysForOutlookHyperlinking()
    
    ' trelloCacheBoardID(trelloFindBoardID()) ' Walk user through finding BoardID, then cache it
    ' trelloCacheListID(trelloFindListID()) ' Walk user through finding ListID, then cache it

End Function