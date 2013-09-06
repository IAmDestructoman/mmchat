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
    
    GetName = "Faeball Fly Plugin By Dirocus " & App.Major & "." & App.Minor
End Function
Public Function GetCommands() As String
    GetCommands = ",FBFly,0,"
End Function
Public Function init(othervars) As String

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
On Error GoTo errored
Dim params() As String
params() = Split(dadata, " ")
If UBound(params()) <> 2 Then
    main = "0Usage: FBFly distance direction"
    Exit Function
End If

Dim PI#

PI# = 4 * Atn(1#)
dadir = Val(params(2)) - (90 * (Fix(Val(params(2)) / 90)))
a3 = 90
a2 = 90 - dadir
a1 = dadir
s3 = Val(params(1))
s1 = Round((s3 * Sin(a1 * PI# / 180)) / Sin(a3 * PI# / 180) / 10, 0)
s2 = Round((s3 * Sin(a2 * PI# / 180)) / Sin(a3 * PI# / 180) / 10, 0)
Select Case Val(params(2))
    Case Is < 90
    dirs = "E" & s1 & " N" & s2
    Case Is < 180
    dirs = "E" & s2 & " S" & s1
    Case Is < 270
    dirs = "W" & s1 & " S" & s2
    Case Is < 360
    dirs = "W" & s2 & " N" & s1
End Select
main = "0Fly " & dirs
Exit Function
errored:
main = "0Error occured"

End Function

