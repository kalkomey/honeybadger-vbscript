<%
'********************************************************************
' Name: ErrorLogger.asp
'
' Purpose: Creates an error logger object that will notify
'          Honeybadger in the event of an error.
'********************************************************************
%>

<!--#include virtual="vendor/honeybadger-vbscript/Honeybadger.asp" -->

<%

Class ErrorLogger
  Private Function GetLastPathItem(strPath)
    Dim arrSegments
    arrSegments = Split(strPath, "\")
    GetLastPathItem = arrSegments(UBound(arrSegments))
  End Function

  Private Function GetServerName()
    Dim objNetwork
    Set objNetwork = Server.CreateObject("WScript.Network")
    GetServerName = objNetwork.ComputerName
  End Function

  Private Function GetCurrentEnvironment()
    ' This returns the current environment.
    ' E.g. development, qa, staging, or production.
    ' This example uses an Application variable from the Global.asa file.
    GetCurrentEnvironment = Application("current_environment")
  End Function

  Private Function GetHoneybadgerApiKey()
    ' This returns the Honeybadger API key.
    ' This example uses an Application variable from the Global.asa file.
    GetHoneybadgerApiKey = Application("honeybadger_api_key")
  End Function

  Private Function GetApplicationRoot()
    ' This returns the filesystem path to the root of your application.
    ' This example uses an Application variable from the Global.asa file.
    GetApplicationRoot = Application("application_root")
  End Function

  Private Function GetUserId()
    ' This returns the logged-in users's ID.
    GetUserId = Session("UserId")
  End Function

  Public Sub Class_Terminate
    On Error Resume Next

    If Server.GetLastError().Number <> 0 Then
      If GetCurrentEnvironment() <> "development" And GetHoneybadgerApiKey() <> "" Then
        Dim strContext, strServer

        strContext = "{""user_id"": """ & GetUserId() & """}"
        strServer = "{ " & _
            """project_root"": """ & Replace(GetApplicationRoot(), "\", "\\") & """, " & _
            """revision"": """ & GetLastPathItem(GetApplicationRoot()) & """, " & _
            """hostname"": """ & GetServerName() & """, " & _
            """environment_name"": """ & GetCurrentEnvironment() & """" & _
          "}"

        PostToHoneybadger GetHoneybadgerApiKey(), strContext, strServer
      End If
    End If
  End Sub
End Class

Dim objErrorLogger
Set objErrorLogger = New ErrorLogger

%>
