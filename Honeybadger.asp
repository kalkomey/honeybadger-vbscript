<%
'********************************************************************
' Name: Honeybadger.asp
'
' Purpose: Provide a simple function to log an exception in Honeybadger.
'********************************************************************

Const HONEYBADGER_URL = "https://api.honeybadger.io/v1/notices"

Function EscapeQuotationMarks(strText)
  EscapeQuotationMarks = Replace(strText, """", "\""")
End Function

Function EscapeBackSlashes(strText)
  EscapeBackSlashes = Replace(strText, "\", "\\")
End Function

Function GetServerVariablesAsJSON()
  Dim strServerVariables, strVar

  For Each strVar in Request.ServerVariables
    If strVar <> "ALL_HTTP" And strVar <> "ALL_RAW" Then
      strServerVariables = strServerVariables & """" & strVar & """: """ & _
        Server.HTMLEncode(EscapeBackSlashes(EscapeQuotationMarks(Request(strVar)))) & """, "
    End If
  Next

  ' Remove additional, trailing comma.
  GetServerVariablesAsJSON = Left(strServerVariables, Len(strServerVariables) - 2)
End Function

Function GetFullURL()
  Dim strProtocol, strUrl

  strProtocol = "http"
  If LCase(Request("HTTPS")) <> "off" Then
     strProtocol = "https"
  End If

  strUrl = strProtocol & "://" & Request("HTTP_HOST") & Request("URL")
  If Len(Request("QUERY_STRING")) > 0 Then
    strUrl = strUrl & "?" & Request("QUERY_STRING")
  End If

  GetFullURL = strUrl
End Function

Sub PostToHoneybadger(strApiKey, strContext, strServer)
  On Error Resume Next

  Dim objXmlHttpMain, strUrl, strServerVariables, objASPError, strJSONToSend
  Set objASPError = Server.GetLastError()
  strServerVariables = GetServerVariablesAsJSON()

  strJSONToSend = "{" & _
    """notifier"": { " & _
      """name"": ""Honeybadger VBScript Error Notifications""" & _
    "}, " & _
    """error"": { " & _
      """class"": """ & Server.HTMLEncode(objASPError.Description) & " " & Server.HTMLEncode(" (0x" & Hex(objASPError.Number) & ")" ) & """, " & _
      """message"": """ & Server.HTMLEncode(objASPError.Category) & ": " & Server.HTMLEncode(objASPError.Description) & " " & Server.HTMLEncode(" (0x" & Hex(objASPError.Number) & ")" ) & """, " & _
      """backtrace"": [ " & _
         "{" & _
         """number"": """ & Server.HTMLEncode(objASPError.Line) & """, " & _
          """file"": """ & Server.HTMLEncode(objASPError.File) & """" & _
        "}" & _
      "]" & _
    "}, " & _
    """server"": " & strServer & ", " & _
    """request"": { " & _
      """context"": " & strContext & ", " & _
      """cgi_data"": {" & strServerVariables & "}, " & _
      """url"": """ & GetFullURL() & """" & _
    "}" & _
    "}"

  Set objXmlHttpMain = CreateObject("MSXML2.ServerXMLHTTP")
  objXmlHttpMain.open "POST", HONEYBADGER_URL, False
  objXmlHttpMain.setRequestHeader "X-API-Key", strApiKey
  objXmlHttpMain.setRequestHeader "Content-Type", "application/json"
  objXmlHttpMain.setRequestHeader "Accept", "application/json"
  objXmlHttpMain.setRequestHeader "User-Agent", "MSXML2.ServerXMLHTTP; VBScript; " & Request("SERVER_SOFTWARE")
  objXmlHttpMain.send(strJSONToSend)
End Sub

%>
