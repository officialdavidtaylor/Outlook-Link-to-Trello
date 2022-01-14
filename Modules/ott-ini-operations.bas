'------------------------------------------------------------------------------
' Title:    ott-ini-operations
' Desc:     uses Win API to read/write config file
' Language: VBA [Outlook for Windows]
' Note:     Special thanks to https://stackoverflow.com/users/11485/birger
'------------------------------------------------------------------------------

Public Function getCachedListID() As String
  getCachedListID = ReadIniFileString("trello", "list_id")
End Function

Public Function getCachedBoardID() As String
  getCachedKey = ReadIniFileString("trello", "board_id")
End Function

Public Function getCachedKey() As String
  getCachedKey = ReadIniFileString("trello", "api_key")
End Function

Public Function getCachedToken() As String
  getCachedToken = ReadIniFileString("trello", "api_token")
End Function

Public Function trelloCacheUserID()
' TODO: [V2.0] Cache Trello User ID
  ' save the user ID for API purposes
  ' https://stackoverflow.com/questions/12428293/best-way-to-cache-a-password-in-an-excel-vba-function
End Function

Public Function trelloCacheCredentials()
' TODO: [V2.0] Cache Trello Credentials

End Function

Public Function trelloCacheListID(listID As String) As String
' cache Trello list ID in INI file
  trelloCacheListID = WriteIniFileString("trello", "list_id", listID)
End Function

Public Function trelloCacheBoardID(boardID As String) As String
' cache Trello board ID in INI file
  trelloCacheBoardID = WriteIniFileString("trello", "board_id", boardID)
End Function

Public Function trelloCacheApiKey(apiKey As String) As String
' cache Trello API Key in INI file
  trelloCacheApiKey = WriteIniFileString("trello", "api_key", apiKey)
End Function

Public Function trelloCacheApiToken(apiToken As String) As String
' cache Trello API Token in INI file
  trelloCacheApiToken = WriteIniFileString("trello", "api_token", apiToken)
End Function

Public Function clearCache()
' delete INI file

  fileToDelete = getConfigFilePath()
  ' ensure file is not read-only
  SetAttr fileToDelete, vbNormal
  ' delete the file
  Kill fileToDelete

  If Dir(fileToDelete) = "" Then
    MsgBox "Cache cleared successfully", vbInformation
  Else
    MsgBox "Error clearing cache. Please contact IT to help resolve this error.", vbCritical, "Cache Error"
  End If
End Function

Public Function firstRunComplete()
' return boolean with whether or not the first run has been completed
' TODO: implement error checking
  firstRunComplete = CBool(ReadIniFileString("app", "first-run-complete"))
End Function