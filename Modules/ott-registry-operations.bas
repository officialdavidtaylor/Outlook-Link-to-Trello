'------------------------------------------------------------------------------
' Title:    ott-registry-operations
' Desc:     functions to edit registry to enable message hyperlink handling
' Language: VBA [Outlook for Windows]
'------------------------------------------------------------------------------

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
Function checkOutlookHyperlinkingStatus() As String
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

' TODO: [V2.0] Find Board ID in Trello
Function trelloFindBoardID() As String
    ' https://stackoverflow.com/questions/26552278/trello-api-getting-boards-lists-cards-information/50908600
End Function

' TODO: [V2.0] Find List ID in Trello
Function trelloFindListID()
    ' https://stackoverflow.com/questions/26552278/trello-api-getting-boards-lists-cards-information/50908600
End Function