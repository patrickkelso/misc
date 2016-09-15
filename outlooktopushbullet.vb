'You need to paste it into the code file called "ThisOutlookSession".
'IMPORTANT You'll need to change the text YOUR_ACCESS_TOKEN to your Access Token that you
'you got from https://www.pushbullet.com/account

Const vbDoubleQuote As String = """" 'represents 1 double quote (")
Const vbSingleQuote As String = "'" 'represents 1 single quote (')
 
Private Sub Application_Reminder(ByVal Item As Object)
   Dim Title As String
   Dim Body As String
   Dim Message As String
  If Item.MessageClass <> "IPM.Appointment" Then
    Exit Sub
  End If

  Title = Item.Location & " at " & Format(Item.Start, "Short Time") & " " & Format(Item.Start, "Short Date")

  Body = Item.Subject

  TargetURL = "https://api.pushbullet.com/v2/pushes"
  Set HTTPReq = CreateObject("WinHttp.WinHttpRequest.5.1")
  HTTPReq.Option(4) = 13056 '
  HTTPReq.Open "POST", TargetURL, False
  HTTPReq.SetCredentials "user", "password", 0

  HTTPReq.setRequestHeader "Authorization", "Bearer YOUR_ACCESS_TOKEN"
  HTTPReq.setRequestHeader "Content-Type", "application/json"

  Message = "{""type"": ""note"", ""title"": " & _
      vbDoubleQuote & Title & vbDoubleQuote & ", ""body"": " & _
      vbDoubleQuote & Body & vbDoubleQuote & "}"
  HTTPReq.Send (Message)

  ' delete this line once you are happy it works. Or put in some proper error handling!
  ' MsgBox (HTTPReq.responseText)

End Sub
