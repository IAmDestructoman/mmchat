VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Plugin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Function GetName() As String
    GetName = "DarkNomads Plugin By Dirocus " & App.Major & "." & App.Minor
End Function
Public Function GetCommands() As String
    GetCommands = ",Riddle,0,Combs,0,ClanTown,0,"
End Function
Public Function init(othervars) As String

End Function
Public Function main(dadata As String, othervars As String) As String
Form1.Winsock1.Close
Dim params() As String
params() = Split(dadata, " ")
If UBound(params) < 1 Then
    main = "0Invalid Params" & Chr(10)
    Exit Function
End If

dadata = Mid(dadata, Len(params(0)) + 2, Len(dadata))

Form1.Winsock1.Connect "www.darknomads.com", 80
DoEvents
x = Timer
Do: DoEvents
If Timer > x + 3 Then Exit Do
If Form1.Winsock1.State = 7 Then Exit Do
Loop
If Form1.Winsock1.State <> 7 Then
    main = "0Connection failure to Darknomads.com"
    Exit Function
End If
Select Case LCase(params(0))
Case "riddle"
    tts = "GET /riddles/searchraw?name=" & dadata & " HTTP/1.1"
Case "combs"
    If InStr(dadata, "|") > 0 Then
        var1 = Mid(dadata, 1, InStr(dadata, "|") - 1)
        var2 = Mid(dadata, InStr(dadata, "|") + 1, Len(dadata))
    Else
        var1 = dadata
        var2 = ""
    End If
    tts = "GET /combs/searchraw?zone=" & var1 & "&room=" & var2 & " HTTP/1.1"
Case "clantown"
    tts = "GET /clantowns/searchraw?name=" & dadata & " HTTP/1.1"
End Select
tts = tts & vbCrLf & "Host: www.darknomads.com" & vbCrLf
tts = tts & "User -Agent: Microsoft -ATL - Native / 7#" & vbCrLf

Form1.Winsock1.Tag = ""
Form1.Winsock1.SendData tts & vbCrLf & vbCrLf
DoEvents
x = Timer
Do: DoEvents
If Timer > x + 2 Then Exit Do
If Form1.Winsock1.Tag <> "" Then Exit Do
Loop
If Form1.Winsock1.Tag = "" Then
main = "0Darknomads.com did not respond in a timly fashon" & Chr(10)
Else
x = Timer
Do: DoEvents
Loop While x + 0.2 > Timer
'main = Form1.Winsock1.Tag
main = "0" & Mid(Form1.Winsock1.Tag, InStr(Form1.Winsock1.Tag, Chr(10) & Chr(27)), Len(Form1.Winsock1.Tag))
End If
End Function



'GET /riddles/searchraw?name=test HTTP/1.1
'Host: www.darknomads.com
'User -Agent: Microsoft -ATL - Native / 7#

'GET /combs/searchraw?zone=test&room= HTTP/1.1
'Host: www.darknomads.com
'User -Agent: Microsoft -ATL - Native / 7#

'GET /clantowns/searchraw?name=test HTTP/1.1
'Host: www.darknomads.com
'User -Agent: Microsoft -ATL - Native / 7#
