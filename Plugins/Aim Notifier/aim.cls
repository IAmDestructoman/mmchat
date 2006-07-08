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
f 'main = result
'there can be multipule results seperated by ~
'the first char of each result tells if you should:
'0 chatper
'1 chatall not per
'2 cchatall not per
'3 chatall all
'4 cchatall all

Public dacount As Integer


Public Function GetName() As String
    
    
    GetName = "Aim Notifier Plugin By Dirocus " & App.Major & "." & App.Minor

End Function

Public Function GetCommands() As String

    GetCommands = ",AIMNotify,4,AIMAdd,1,AIMList,4,"

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
main = "0TIMMY"
Dim daparams() As String
daparams() = Split(dadata, " ")
Select Case LCase(daparams(0))
Case "aimnotify"
    'load all settings, make array of users with aim
    
    Form1.Text1.Text = GetFromINI("main", "aimuser", "none", getvar("path", othervars) & "\plugins\aim.ini")
    Form1.Text2.Text = GetFromINI("main", "aimpass", "none", getvar("path", othervars) & "\plugins\aim.ini")
    If Form1.Text1.Text = "none" Then
        main = "0No aim account to log in with!"
        Exit Function
    End If
    
    If Form1.Text2.Text = "none" Then
        main = "0No aim account to log in with!"
        Exit Function
    End If
    
    Form1.List1.Clear
    
    dapath = getvar("path", othervars)
    Form1.File1.Path = dapath & "\users"
    Form1.File1.Pattern = "*.usr"
    Form1.File1.Refresh
    For a = 0 To Form1.File1.ListCount - 1
        daname = GetFromINI("main", "aim", "", dapath & "\users\" & Form1.File1.List(a))
        If daname <> "" Then Form1.List1.AddItem daname
    Next a
    If Form1.List1.ListCount = 0 Then
        main = "0There are no aim accounts. users must use aimadd to do this"
        Exit Function
    End If
    dacount = 3
    Form1.Winsock1.Close
    Form1.Winsock1.Connect "toc.oscar.aol.com", 23
    x = Timer
    Do: DoEvents
    If x + 5 < Timer Then
        main = "0Timeout connecting to AIM.(1)" & Chr(10)
        Exit Function
    End If
    Loop While Form1.Winsock1.State <> 7
    Form1.Winsock1.Tag = "flapon"
    Form1.Winsock1.SendData "FLAPON" & vbCrLf & vbCrLf
    x = Timer
    Do: DoEvents
    If x + 5 < Timer Then
    main = "0Timeout connecting to AIM. (Username / password correct??)" & Chr(10)
        Exit Function
    End If
    Loop While Form1.Winsock1.Tag = "flapon"
    
    Dim damessage As String
    If Len(dadata) > 10 Then
        damessage = "Automated message: " & getvar("ocn", othervars) & " says: " & Mid(dadata, 11, Len(dadata))
    Else
        damessage = "Automated message: " & getvar("ocn", othervars) & " has requested you join the " & getvar("servername", othervars) & " chat."
    End If
    
    'loop through users
    
    For a = 0 To Form1.List1.ListCount - 1
    timeout 3
    Call sendim(Form1.List1.List(a), damessage)
    
    Next a
    
    timeout 2
    Form1.Winsock1.Close
    main = "0Aim notification sent." & Chr(10)
    Exit Function
Case "aimadd"
    Select Case UBound(daparams)
    Case 1
        addaimtofile getvar("path", othervars) & "\users\" & getvar("ocn", othervars) & ".usr", daparams(1)
        main = "0Your aim screenname has been set to " & daparams(1) & "."
    Case 2
        If getvar("perm", othervars) < 4 Then
            main = "0You do not have the required permissions to do this."
            Exit Function
        End If
        If Dir(getvar("path", othervars) & "\users\" & daparams(1) & ".usr", vbNormal) = "" Then
            main = "0User " & daparams(1) & " not found in the user database."
            Exit Function
        End If
        addaimtofile getvar("path", othervars) & "\users\" & daparams(1) & ".usr", daparams(2)
        main = "0Aim screen name succesfully updated to " & daparams(2) & " for user " & daparams(1)
    Case Else
        main = "0Incorrect parameters"
    End Select
Case "aimlist"
    dapath = getvar("path", othervars)
    Form1.File1.Path = dapath & "\users"
    Form1.File1.Pattern = "*.usr"
    Form1.File1.Refresh
    main = "0Current known aim accounts:" & Chr(10)
    For a = 0 To Form1.File1.ListCount - 1
        daname = GetFromINI("main", "aim", "", dapath & "\users\" & Form1.File1.List(a))
        If daname <> "" Then main = main & Form1.File1.List(a) & " - " & daname & Chr(10)
    Next a
End Select

End Function
Function addaimtofile(filename As String, aimname As String)
AddToINI "main", "aim", aimname, filename
End Function
Public Function sendim(dauser As String, damessage As String)
therest = "toc2_send_im " & dauser & " " & Chr(34) & damessage & Chr(34) & Chr(0)
Form1.Winsock1.SendData Chr(42) & Chr(2) & Chr(0) & Chr(dacount) & Chr(0) & Chr(Len(therest)) & therest

dacount = dacount + 1
End Function
Function timeout(xx)
x = Timer
Do: DoEvents
Loop While x + xx > Timer
End Function






