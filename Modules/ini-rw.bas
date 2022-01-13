'------------------------------------------------------------------------------
' Title:    ini-rw
' Desc:     uses Win API to read/write config file
' Language: VBA [Outlook for Windows]
' Note:     Special thanks to https://stackoverflow.com/users/11485/birger
'------------------------------------------------------------------------------

'------DECLARATIONS------

' use for reading/writing INI file
#If VBA7 Then
Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
#Else
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
#End If

'------FUNCTIONS------

Function IniFileName() As String
' relative path based on users environment, targeting AppData
  IniFileName = getConfigFilePath()
End Function

Private Function ReadIniFileString(ByVal Sect As String, ByVal Keyname As String) As String
' read ini file basedon 
  Dim Worked As Long
  Dim RetStr As String * 128
  Dim StrSize As Long

  iNoOfCharInIni = 0
  sIniString = ""
  If Sect = "" Or Keyname = "" Then
    MsgBox "Section Or Key To Read Not Specified !!!", vbExclamation, "INI"
  Else
    sProfileString = ""
    RetStr = Space(128)
    StrSize = Len(RetStr)
    Worked = GetPrivateProfileString(Sect, Keyname, "", RetStr, StrSize, IniFileName)
    If Worked Then
      iNoOfCharInIni = Worked
      sIniString = Left$(RetStr, Worked)
    End If
  End If
  ReadIniFileString = sIniString
End Function

Private Function WriteIniFileString(ByVal Sect As String, ByVal Keyname As String, ByVal Wstr As String) As String
' write to ini file based on Section, Keyname, and Value
  Dim Worked As Long

  iNoOfCharInIni = 0
  sIniString = ""
  If Sect = "" Or Keyname = "" Then
    MsgBox "Section Or Key To Write Not Specified !!!", vbExclamation, "INI"
  Else
    Worked = WritePrivateProfileString(Sect, Keyname, Wstr, IniFileName)
    If Worked Then
      iNoOfCharInIni = Worked
      sIniString = Wstr
    End If
    WriteIniFileString = sIniString
  End If
End Function