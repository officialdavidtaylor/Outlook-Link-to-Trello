'------------------------------------------------------------------------------
' Title:    ott-create-trello-card
' Desc:     use Trello API to create card from email
' Language: VBA [Outlook for Windows]
'------------------------------------------------------------------------------

' Create new Card with CardPayload object as input
Public Function trelloCreateCard(ByRef cardPayload As CardPayload)
  ' Use this method to create a card with custom fields and attachments, and to provide
  ' useful feedback in the event of an operation failure.
  ' https://developer.atlassian.com/cloud/trello/rest/api-group-actions/
  ' https://stackoverflow.com/questions/158633/how-can-i-send-an-http-post-request-to-a-server-from-excel-using-vba
  ' https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms762278(v=vs.85)?redirectedfrom=MSDN

  ' VARIABLES
  Dim cardApiUrl As String        ' URL needed to create a card
  Dim attachmentApiUrl As String  ' URL needed to add an attachment to a card
  Dim responseText As String      ' to be returned from POST request
  Dim cardPayloadString As String ' variable (in JSON format) to contain all of the parts required for the POST request
  Dim attachmentPayload As String ' To contain required info to add attachment to card
  Dim counter As Integer          ' used to ensure loops are ended eventually

  cardApiUrl = "https://api.trello.com/1/cards" ' URL for Trello API calls for Cards

  ' Generate the payload
  cardPayloadString = "{""name"":""" & cardPayload.cardName & """, ""idList"":""" & cardPayload.listID & """, ""key"":""" & getCachedKey() & """, ""token"":""" & getCachedToken() & """, ""pos"":""top""}" 

  ' CREATE TRELLO CARD
  ' Initiate HTTP interface object: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms754586(v=vs.85)
  Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
  ' Prepare HTTP request
  objHTTP.open "POST", cardApiUrl, False ' stage a new POST request
  objHTTP.SetRequestHeader "Content-type", "application/json" ' tell server what format to expect payload (JSON)
  ' Send POST request
  objHTTP.send cardPayloadString
  ' save and process server response
  responseText = objHTTP.responseText
  cardPayload.cardID = extractCardID(responseText)

  ' check cardID length to ensure it has been saved correctly
  Do
      If (Len(cardPayload.cardID) < 23) Then
          counter = counter + 1
          If (counter > MAX_LOOP_ITERATIONS) Then
              MsgBox "Error: CardID fumbled, terminating hyperlink addition operation"
              Exit Sub ' Cancel the hyperlink operation by ending the Sub
          End If
      End If
      Exit Do ' If the Length is correct, proceed with the operation
  Loop

  ' ADD BACKLINK TO OUTLOOK AS ATTACHMENT TO TRELLO CARD
  ' Construct API hyperlink with appropriate CardID
  attachmentApiUrl = (cardApiUrl & "/" & cardPayload.cardID & "/attachments")
  ' construct payload for HTTP request
  attachmentPayload = "{""id"":""" & cardPayload.cardID & """, ""key"":""" & getCachedKey() & """, ""token"":""" & getCachedToken() & """, ""name"":""Email Link"", ""url"":""" & cardPayload.mailUID & """}" 
  ' prepare HTTP request
  objHTTP.open "POST", attachmentApiUrl, False ' stage a new POST request
  objHTTP.SetRequestHeader "Content-type", "application/json" ' tell server what format to expect payload (JSON)
  ' send POST request
  objHTTP.send attachmentPayload

  trelloCreateCard = objHTTP.responseText

End Function