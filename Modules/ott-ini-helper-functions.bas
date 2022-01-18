'------------------------------------------------------------------------------
' Title:    ini-helper-functions
' Desc:     uses Win API to read/write config file
' Language: VBA [Outlook for Windows]
' Note:     Special thanks to @dee-u https://www.vbforums.com/showthread.php?349993.html
'------------------------------------------------------------------------------

'------DECLARATIONS------

'declarations for working with Ini files
Private Declare PtrSafe Function GetPrivateProfileSection Lib "kernel32" Alias _
    "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, _
    ByVal nSize As Long, ByVal lpFileName As String) As Long
 
Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias _
    "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
    ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, _
    ByVal lpFileName As String) As Long
 
Private Declare PtrSafe Function WritePrivateProfileSection Lib "kernel32" Alias _
    "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, _
    ByVal lpFileName As String) As Long
 
Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias _
    "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
    ByVal lpString As Any, ByVal lpFileName As String) As Long
 
'// INI CONTROLLING PROCEDURES
'reads an Ini string
Public Function ReadIni(Section As String, Key As String) As String
  Dim RetVal As String * 255, v As Long
  v = GetPrivateProfileString(Section, Key, "", RetVal, 255, configFilePath)
  If v Then
    ReadIni = Left(RetVal, v)
  Else
    ReadIni = 0
  End If
End Function
 
'reads an Ini section
Public Function ReadIniSection(Section As String) As String
  Dim RetVal As String * 255, v As Long
  v = GetPrivateProfileSection(Section, RetVal, 255, configFilePath)
  If v Then
    ReadIniSection = Left(RetVal, v)
  Else
    ReadIni = 0
  End If
End Function
 
'writes an Ini string
Public Function WriteIni(Section As String, Key As String, Value As String)
  WriteIni = WritePrivateProfileString(Section, Key, Value, configFilePath)
End Function
 
'writes an Ini section
Public Function WriteIniSection(Section As String, Value As String)
  WriteIniSection = WritePrivateProfileSection(Section, Value, configFilePath)
End Function
