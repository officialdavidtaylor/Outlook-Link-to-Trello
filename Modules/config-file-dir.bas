'------------------------------------------------------------------------------
' Title:    config-file-dir
' Desc:     returns the location of the config.ini file.
' Language: VBA [Outlook for Windows]
'------------------------------------------------------------------------------

'------DECLARATIONS------

Public Enum eSpecialFolders
' use to determine location of AppData folder, based on Windows
  SpecialFolder_AppData = &H1A        'for the current Windows user, on any computer on the network [Windows 98 or later]
  SpecialFolder_CommonAppData = &H23  'for all Windows users on this computer [Windows 2000 or later]
  SpecialFolder_LocalAppData = &H1C   'for the current Windows user, on this computer only [Windows 2000 or later]
  SpecialFolder_Documents = &H5       'the Documents folder for the current Windows user
End Enum

'------CONSTANTS------

Public Const CONFIG_PATH As String = "\Outlook-to-Trello\config.ini"

'------FUNCTIONS------

Function SpecialFolder(pFolder As eSpecialFolders) As String
' returns the path to the specified special folder (AppData etc)

  Dim objShell  As Object
  Dim objFolder As Object

  Set objShell = CreateObject("Shell.Application")
  Set objFolder = objShell.Namespace(CLng(pFolder))

  If (Not objFolder Is Nothing) Then SpecialFolder = objFolder.Self.Path

  Set objFolder = Nothing
  Set objShell = Nothing

  If SpecialFolder = "" Then Err.Raise 513, "SpecialFolder", "The folder path could not be detected"

End Function

Public Function getConfigPath()
' return the path of config file in the directory specified in the CONFIG_PATH const
  getConfigPath = SpecialFolder(SpecialFolder_AppData) & CONFIG_PATH
End Sub