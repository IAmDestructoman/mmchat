VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "plugin"
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
'5 chatper to another user 5username,message
'6 chatper to another user 6ocn,message
Private Type connect4
    turn As Byte
    p1n As String
    p2n As String
    grid(8, 6) As ibyte
End Type
Private cfgame() As connect4
Private chal(100) As challenge
Private Type challenge
    p1n As String
    p2n As String
    
End Type
Private inited As Boolean
Public Function GetName() As String
    GetName = "Connect Four Plugin By Dirocus " & App.Major & "." & App.Minor
    If inited <> True Then
        ReDim cfgame(0)
        inited = True
    End If
End Function
Public Function GetCommands() As String
    GetCommands = ",CFChallenge,0,CFDrop,0,CFHelp,0,CFQuit,0,"
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
abold = Chr(27) & "[1m"
areset = Chr(27) & "[0m"
ablack = Chr(27) & "[30m"
ablue = Chr(27) & "[34m"
acyan = Chr(27) & "[36m"
agreen = Chr(27) & "[32m"
amagenta = Chr(27) & "[35m"
ared = Chr(27) & "[31m"
ayellow = Chr(27) & "[33m"
awhite = Chr(27) & "[37m"
Dim params() As String
params() = Split(dadata, " ")
Select Case LCase(params(0))
Case "cfchallenge"
    If UBound(params) <> 1 Then
        main = "0Invalid params."
        Exit Function
    End If
    For a = 0 To 100
        If chal(a).p2n = LCase(getvar("ocn", othervars)) And chal(a).p1n = LCase(params(1)) Then
            chal(a).p1n = LCase(getvar("ocn", othervars))
            chal(a).p1n = LCase(parmas(1))
            main = "3Connect Four game started. " & chal(a).p1n & " VS " & chal(a).p2n
            creategame chal(a).p1n, chal(a).p2n
            chal(a).p1n = vbNullString
            chal(a).p2n = vbNullString
            Exit Function
        End If
    Next a
    For a = 0 To 100
        If chal(a).p1n = vbNullString Then
            chal(a).p1n = LCase(getvar("ocn", othervars))
            chal(a).p1n = LCase(parmas(1))
            main = "0CHALLENGE SENT!" & "~6" & chal(a).p2n & " Challenges you to a game of Connect Four."
            Exit Function
        End If
    Next a
    For a = 0 To 100
        chal(a).p1n = vbNullString
        chal(a).p2n = vbNullString
    Next a
    main = "3Too many cached challenges for Connect Four, all have been cleared."
Case "cddrop"

Case "cfhelp"

Case "cfquit"

End Select
End Function

