'------------------------------------------------------------------------------
' Title:    ott-ini-operations
' Desc:     uses Win API to read/write config file
' Language: VBA [Outlook for Windows]
' Note:     Special thanks to https://stackoverflow.com/users/11485/birger
'------------------------------------------------------------------------------

Public Function getCachedUserID() As String
  getCachedUserID = ReadIni("trello", "user_id")
End Function

Public Function getCachedListID() As String
  getCachedListID = ReadIni("trello", "list_id")
End Function

Public Function getCachedBoardID() As String
  getCachedKey = ReadIni("trello", "board_id")
End Function

Public Function getCachedKey() As String
  getCachedKey = ReadIni("trello", "api_key")
End Function

Public Function getCachedToken() As String
  getCachedToken = ReadIni("trello", "api_token")
End Function


Public Function trelloCacheCredentials()
' TODO: [V2.0] Cache Trello Credentials

End Function

Public Function trelloCacheUserID(userID As String)
  trelloCacheUserID = WriteIni("trello", "user_id", userID)
End Function

Public Function trelloCacheListID(listID As String) As String
' cache Trello list ID in INI file
  trelloCacheListID = WriteIni("trello", "list_id", listID)
End Function

Public Function trelloCacheBoardID(boardID As String) As String
' cache Trello board ID in INI file
  trelloCacheBoardID = WriteIni("trello", "board_id", boardID)
End Function

Public Function trelloCacheApiKey(apiKey As String) As String
' cache Trello API Key in INI file
  trelloCacheApiKey = WriteIni("trello", "api_key", apiKey)
End Function

Public Function trelloCacheApiToken(apiToken As String) As String
' cache Trello API Token in INI file
  trelloCacheApiToken = WriteIni("trello", "api_token", apiToken)
End Function

Public Function clearCache()
' delete INI file

  fileToDelete = configFilePath
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
  firstRunComplete = CBool(ReadIni("app", "first-run-complete"))
End Function