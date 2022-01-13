'------------------------------------------------------------------------------
' Title:    type-declarations
' Desc:     declaration file for all custom types used in Outlook-to-Trello
' Language: VBA [Outlook for Windows]
'------------------------------------------------------------------------------

'------STRUCTURES------

Public Type CardPayload
' To package all necessary data before making HTTP requests to create Trello Card

  listID As String            ' Trello List ID for card to be entered
  cardID As String            ' Trello Card ID
  cardName As String          ' Trello Card name
  sender As String            ' Sender of email selected
  subject As String           ' Subject of email selected
  mailUID As String           ' UID of email selected
  conversationID As String    ' conversationID of email selected
  receivedTime As String      ' Time received of email selected
  ' TODO: Add file attachment support in the future: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/file-object

  cardCreated As Boolean      ' status variable for checking if card creation was successful

  ' goal is to remove these from this object for security purposes
  token As String             ' Trello API credentials: token
  key As String               ' Trello API credentials: key

End Type

Public Type RegistryItem
' path and key name

  path As String
  key As String

End Type

Public Type TrelloCredentialCache
' Structure for token and key cache

  token As String     ' API access token from Atlassian/Trello
  key AS String       ' API access key from Atlassian/Trello
  username As String  ' Trello username

End Type