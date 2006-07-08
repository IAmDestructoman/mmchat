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
    
    
    GetName = "Message Board Plugin By Dirocus " & App.Major & "." & App.Minor
End Function
Public Function GetCommands() As String
    GetCommands = ",Read,0,Take,1,Post,1,"
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

On Error GoTo errored
Dim params() As String
params() = Split(dadata, " ")

username = getvar("ocn", othervars)
dapath = getvar("path", othervars)
If Dir(dapath & "\messages", vbDirectory) = "" Then MkDir (dapath & "\messages")
Form1.File1.Path = dapath & "\messages"

Form1.File1.Refresh
Select Case LCase(params(0))

Case "read"
    
    If UBound(params) > 0 Then
        num = Val(params(1))
        
        tts = Chr(10) & abold & awhite & "                            " & getvar("servername", othervars) & " BULLETIN BOARD"
        tts = tts & Chr(10) & ablue & "[ " & awhite & "#" & ablue & " ][" & awhite & "   Date   " & ablue & "][" & awhite & "       Author" & ablue & "][" & awhite & "Subject                                      " & ablue & "]" & Chr(10) & awhite & " ---  ----------  -------------  ---------------------------------------------" & Chr(10)
        
        
        ffile = FreeFile
        Open dapath & "\messages\" & Form1.File1.List(num - 1) For Input As ffile
        dafile = Input(LOF(ffile), ffile)
        Close ffile
        Dim hrmm3() As String
        hrmm3() = Split(dafile, ",")
        spaces = Space(13 - Len(hrmm3(0)))
        spaces2 = Space(45 - Len(hrmm3(2)))
        danum = num
        Do While Len(danum) <> 3
        danum = "0" & danum
        Loop
        tts = tts & "[" & ayellow & danum & awhite & "][" & acyan & hrmm3(1) & awhite & "][" & agreen & spaces & hrmm3(0) & awhite & "][" & acyan & hrmm3(2) & spaces2 & awhite & "]" & Chr(10)
        damessage = ""
        For ttt = 3 To UBound(hrmm3)
        damessage = damessage & hrmm3(ttt)
        Next ttt
        tts = tts & Chr(10) & areset & damessage
        main = "0" & tts
        Exit Function
        
    Else

        tts = Chr(10) & abold & awhite & "                            " & getvar("servername", othervars) & " BULLETIN BOARD"
        tts = tts & Chr(10) & ayellow & "            (read)(read #)(post subject,message)(take #)" & Chr(10)
        tts = tts & Chr(10) & ablue & "[ " & awhite & "#" & ablue & " ][" & awhite & "   Date   " & ablue & "][" & awhite & "       Author" & ablue & "][" & awhite & "Subject                                      " & ablue & "]" & Chr(10) & awhite & " ---  ----------  -------------  ---------------------------------------------" & Chr(10)
        For num = 0 To Form1.File1.ListCount - 1
            ffile = FreeFile
            Open dapath & "\messages\" & Form1.File1.List(num) For Input As ffile
            dafile = Input(LOF(ffile), ffile)
            Dim hrmm() As String
            hrmm() = Split(dafile, ",")
            spaces = Space(13 - Len(hrmm(0)))
            spaces2 = Space(45 - Len(hrmm(2)))
            danum = num + 1
            Do While Len(danum) <> 3
            danum = "0" & danum
            Loop
            tts = tts & "[" & ayellow & danum & awhite & "][" & acyan & hrmm(1) & awhite & "][" & agreen & spaces & hrmm(0) & awhite & "][" & acyan & hrmm(2) & spaces2 & awhite & "]" & Chr(10)
            Close ffile
        Next num
        main = "0" & tts & "~1" & username & " glances at the board."
    End If
    Exit Function
Case "take"

    If UBound(params) <> 1 Then
        main = "0Invalid Params" & Chr(10)
        Exit Function
    End If
    ffile = FreeFile
    Open dapath & "\messages\" & Form1.File1.List(Val(params(1)) - 1) For Input As ffile
    dafile = Input(LOF(ffile), ffile)
    Close ffile
    Dim hrmm2() As String
    hrmm2() = Split(dafile, ",")
    If LCase(hrmm2(0)) = LCase(username) Then
    Kill dapath & "\messages\" & Form1.File1.List(params(1) - 1)
    main = "0You take the paper from the board~1" & username & " takes a message from the board."
    Else
    main = "0Thats not your message to take!"
    
    End If
    
    
    Form1.File1.Refresh
Case "post"
    
    ffile = FreeFile
    If InStr(dadata, ",") = 0 Then
        main = "0You must serperate the subject and the message with a comma"
        Exit Function
    End If
    If InStr(dadata, ",") >= 45 Then
        main = "0Subject too long."
        Exit Function
    End If
    Open dapath & "\messages\" & "10.msg" For Output As ffile
    Print #ffile, username & "," & Format(Date, "yyyy-mm-dd"); "," & Mid(dadata, 6, Len(dadata))
    Close ffile
    
    
    If Form1.File1.ListCount > 24 Then
        Kill (dapath & "\messages\" & "35.msg")
    End If
    destfile = 35
    Form1.File1.Refresh
    For temp = Form1.File1.ListCount - 1 To 0 Step -1
        fil1 = Form1.File1.List(temp)
        dafile1 = dapath & "\messages\" & fil1
        dafile2 = dapath & "\messages\" & destfile & ".msg"
        If dafile1 <> dafile2 Then Name dafile1 As dafile2
        destfile = destfile - 1
    Next temp
    main = "0You tack the paper to the board.~1" & username & " tacks a piece of paper to the board."
Case Else
errored:
    main = "0ERROR " & Err.Description
End Select

End Function
