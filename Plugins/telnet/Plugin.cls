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
'main = result
'there can be multipule results seperated by ~
'the first char of each result tells if you should:
'0 chatper
'1 chatall not per
'2 cchatall not per
'3 chatall all
'4 cchatall all
'


Public Function GetName() As String
    
    
    GetName = "Telnet connector Plugin By Dirocus " & App.Major & "." & App.Minor
    If Form1.Winsock1(0).State <> 2 Then
        Form1.Winsock1(0).Close
        Form1.Winsock1(0).LocalPort = Val(GetFromINI("Main", "Port", "4050", App.Path & "\..\mmchat.ini")) + 1
        Form1.Winsock1(0).Listen
    
    End If
End Function

Public Function GetCommands() As String

    GetCommands = ",,"

End Function

Function getvar(varname As String, othervars As String) As String
If InStr(othervars, varname & "=") = 0 Then
    getvar = vbNullString
Else

    daloc = InStr(othervars, varname & "=")
    daloc2 = InStr(daloc, othervars, ",")
    getvar = Mid(othervars, daloc + Len(varname) + 1, daloc2 - daloc - Len(varname) - 1)
End If
End Function

Public Function main(dadata As String, othervars As String) As String
End Function

