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
    GetName = "Holofinder Plugin By Dirocus " & App.Major & "." & App.Minor
End Function
Public Function GetCommands() As String
    GetCommands = ",HoloList,0,"
End Function
Public Function init(othervars) As String
    Load Form1
    'Form1.Visible = True
    
    
End Function

Public Function main(dadata As String, othervars As String) As String
On Error GoTo errored
If dadata = "hololist" Then
    Form1.sendupdate
End If
    

Exit Function
errored:
main = "0Error occured"

End Function


Public Function timer() As String
timer = vbNullString
If daupdate <> vbNullString Then
    timer = daupdate
    daupdate = vbNullString
End If
End Function

