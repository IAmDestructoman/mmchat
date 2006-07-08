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
'6 chatper to another user 5ocn,message

Public Type bsloc
    id As String
    '0 open water
    '1 some ship
    '2 some ship
    '3 etc
    hit As Boolean
End Type
Public Type bsgamevars
    p1n As String
    p2n As String
    p1grid(1 To 10, 1 To 10) As bsloc
    p2grid(1 To 10, 1 To 10) As bsloc
    turn As Boolean
    p1r As Boolean
    p2r As Boolean
    p1sunken As Byte
    p2sunken As Byte
    
    
End Type
Private BSGAME() As bsgamevars


Private Sub CreateGame(p1 As String, p2 As String)
x = -1
For A = UBound(BSGAME) To 1 Step -1
    If BSGAME(A).p1n = vbNullString Then x = A
Next A
If x = -1 Then
ReDim Preserve BSGAME(UBound(BSGAME) + 1)
x = UBound(BSGAME)
End If

With BSGAME(x)
    .p1n = p1
    .p2n = p2
    .turn = True
    .p1r = False
    .p2r = False
    .p1sunken = 0
    .p2sunken = 0
End With
For A = 1 To 10
    For aa = 1 To 10
        BSGAME(x).p1grid(A, aa).id = "O"
        BSGAME(x).p1grid(A, aa).hit = False
    Next aa
Next A
For A = 1 To 10
    For aa = 1 To 10
        BSGAME(x).p2grid(A, aa).id = "O"
        BSGAME(x).p2grid(A, aa).hit = False
    Next aa
Next A
End Sub
Public Function GetName() As String
    GetName = "BattleShip Plugin By Dirocus " & App.Major & "." & App.Minor & " Inspired by OMA's Battleship. "
    ReDim BSGAME(0)
End Function

Public Function GetCommands() As String

    GetCommands = ",BSHelp,0,BSChallenge,0,BSFire,0,BSShow,0,BSReady,0,BSPlace,0,BSRemove,0,BSQuit,0"

End Function

Function getvar(varname As String, othervars As String) As String
If InStr(othervars, varname & "=") = 0 Then
    getvar = vbNullString
Else
    Dim daloc As Integer
    Dim daloc2 As Integer
    daloc = InStr(othervars, varname & "=")
    daloc2 = InStr(daloc, othervars, ",")
    getvar = Mid(othervars, daloc + Len(varname) + 1, daloc2 - daloc - Len(varname) - 1)
End If
End Function

Public Function main(dadata As String, othervars As String) As String
abold = Chr(27) & "[1m"
ared = Chr(27) & "[31m"
areset = Chr(27) & "[0m"
ablack = Chr(27) & "[30m"
ablue = Chr(27) & "[34m"
acyan = Chr(27) & "[36m"
agreen = Chr(27) & "[32m"
amagenta = Chr(27) & "[35m"
ayellow = Chr(27) & "[33m"
awhite = Chr(27) & "[37m"
Dim dauser As String
Dim temp As Integer
Dim playernum As Byte
playernum = 0
dauser = getvar("ocn", othervars)
dagame = -1

For temp = 1 To UBound(BSGAME)

If LCase(BSGAME(temp).p1n) = LCase(dauser) Then
    playernum = 1
    dagame = temp
End If
If LCase(BSGAME(temp).p2n) = LCase(dauser) Then
    playernum = 2
    dagame = temp
End If
Next temp

Dim params() As String
params = Split(dadata, " ")
Select Case LCase(params(0))
Case "bshelp"
    'help file
    main = "0help file here."
    Exit Function
Case "bschallenge"
    'challenge a user to BATTLESHIP
    If playernum > 0 Then
        main = "0Your already in a game! BSQuit to end your current game."
        Exit Function
    End If
    If UBound(params) <> 1 Then
        main = "0You Must specify someone to challenge"
        Exit Function
    End If

    For temp = 0 To Form1.challenge.ListCount
        If LCase(Form1.challenge.List(temp)) = LCase(params(1) & "~" & getvar("ocn", othervars)) Then
            Dim dauser2 As String
            dauser2 = Mid(Form1.challenge.List(temp), 1, InStr(Form1.challenge.List(temp), "~") - 1)
            Form1.challenge.RemoveItem temp
            CreateGame dauser2, dauser
            main = "3BattleShip War Started!  " & Chr(10) & dauser2 & " Wins(0) Losses(0)" & Chr(10) & dauser & " Wins(0) Losses(0)"
            main = main & "~0" & DrawGrid(dauser)
            main = main & "~6" & dauser2 & "," & DrawGrid(dauser2)
            Exit Function
        End If
    Next temp
    Form1.challenge.AddItem getvar("ocn", othervars) & "~" & params(1)
    main = "0Challenge SENT!!!!!!~6" & params(1) & "," & dauser & " CHALLENGES you to a game of BATTLESHIP!"
    Exit Function
End Select

'cant go past here if no game is active!
If playernum = 0 Then
    main = "0No Game Started.."
    Exit Function
End If

Select Case LCase(params(0))
Case "bsfire"
'FIRE!
Case "bsshow"
'redraw the board
Case "bsready"
'finsished placing ships, start the game
Case "bsplace"
'place a ship
Case "bsremove"
'remove a ship
Case "bsquit"
'quit/forefit

End Select
End Function
Function DrawGrid(dauser As String) As String

abold = Chr(27) & "[1m"
ared = Chr(27) & "[31m"
areset = Chr(27) & "[0m"
ablack = Chr(27) & "[30m"
ablue = Chr(27) & "[34m"
acyan = Chr(27) & "[36m"
agreen = Chr(27) & "[32m"
amagenta = Chr(27) & "[35m"
ayellow = Chr(27) & "[33m"
awhite = Chr(27) & "[37m"



'              null                             YOU
'   A  B  C  D  E  F  G  H  I  J     A  B  C  D  E  F  G  H  I  J
' 1 +--+--+--+--+--+--+--+--+--+   1 +--+--+--+--+--+--+--+--+--+
'   |  |  |  |  |  |  |  |  |  |     |  |  |  |  |  |  |  |  |  |
' 2 +--+--+--+--+--+--+--+--+--+   2 +--+--+--+--+--+--+--+--+--+
'   |  |  |  |  |  |  |  |  |  |     |  |  |  |  |  |  |  |  |  |
' 3 +--+--+--+--+--+--+--+--+--+   3 +--+--+--+--+--+--+--+--+--+
'   |  |  |  |  |  |  |  |  |  |     |  |  |  |  |  |  |  |  |  |
' 4 +--+--+--+--+--+--+--+--+--+   4 +--+--+--+--+--+--+--+--+--+
'   |  |  |  |  |  |  |  |  |  |     |  |  |  |  |  |  |  |  |  |
' 5 +--+--+--+--+--+--+--+--+--+   5 +--+--+--+--+--+--+--+--+--+
'   |  |  |  |  |  |  |  |  |  |     |  |  |  |  |  |  |  |  |  |
' 6 +--+--+--+--+--+--+--+--+--+   6 +--+--+--+--+--+--+--+--+--+
'   |  |  |  |  |  |  |  |  |  |     |  |  |  |  |  |  |  |  |  |
' 7 +--+--+--+--+--+--+--+--+--+   7 +--+--+--+--+--+--+--+--+--+
'   |  |  |  |  |  |  |  |  |  |     |  |  |  |  |  |  |  |  |  |
' 8 +--+--+--+--+--+--+--+--+--+   8 +--+--+--+--+--+--+--+--+--+
'   |  |  |  |  |  |  |  |  |  |     |  |  |  |  |  |  |  |  |  |
' 9 +--+--+--+--+--+--+--+--+--+   9 +--+--+--+--+--+--+--+--+--+
'   |  |  |  |  |  |  |  |  |  |     |  |  |  |  |  |  |  |  |  |
'10 +--+--+--+--+--+--+--+--+--+  10 +--+--+--+--+--+--+--+--+--+
'   (c)Oma's ToysForTards

For temp = 1 To UBound(BSGAME)
    If LCase(BSGAME(temp).p1n) = LCase(dauser) Then
        playernum = 1
        dagame = temp
    End If
    If LCase(BSGAME(temp).p2n) = LCase(dauser) Then
        playernum = 2
        dagame = temp
    End If
Next temp

If playernum = 1 Then
    topname = BSGAME(dagame).p2n
Else
    topname = BSGAME(dagame).p1n
End If
spaces1 = Space(Fix((28 - Len(topname)) / 2))
spaces2 = Space(45 - (Fix((28 - Len(topname)) / 2)) - Len(topname))
DrawGrid = abold & ared & "   " & spaces1 & topname & spaces2 & agreen & "YOU" & Chr(10)

DrawGrid = DrawGrid & ayellow & "   A  B  C  D  E  F  G  H  I  J     A  B  C  D  E  F  G  H  I  J" & Chr(10)
For temp = 1 To 10
    If temp < 10 Then
        DrawGrid = DrawGrid & " " & ayellow & temp & " "
    Else
        DrawGrid = DrawGrid & ayellow & temp & " "
    End If
    For temp2 = 1 To 10
        If playernum = 1 Then
            If BSGAME(dagame).p2grid(temp, temp2).hit = True Then
                
                If BSGAME(dagame).p2grid(temp, temp2).id = "O" Then
                DrawGrid = DrawGrid & awhite & "¤"
                Else
                DrawGrid = DrawGrid & ared & BSGAME(dagame).p2grid(temp, temp2).id
                End If
                
            Else
                DrawGrid = DrawGrid & ablue & "+"
            End If
        Else
            If BSGAME(dagame).p1grid(temp, temp2).hit = True Then
                If BSGAME(dagame).p1grid(temp, temp2).id = "O" Then
                DrawGrid = DrawGrid & awhite & "¤"
                Else
                DrawGrid = DrawGrid & ared & BSGAME(dagame).p1grid(temp, temp2).id
                End If
            Else
                DrawGrid = DrawGrid & ablue & "+"
            End If
        End If
        If temp2 < 10 Then DrawGrid = DrawGrid & ablue & "--"
    Next temp2
    If temp < 10 Then
        DrawGrid = DrawGrid & "   " & ayellow & temp & " "
    Else
        DrawGrid = DrawGrid & "  " & ayellow & temp & " "
    End If
    For temp2 = 1 To 10
        DrawGrid = DrawGrid & ablue
        If playernum = 1 Then
            
            If BSGAME(dagame).p1grid(temp, temp2).hit = True Then
                If BSGAME(dagame).p1grid(temp, temp2).id = "O" Then
                    DrawGrid = DrawGrid & awhite
                Else
                    DrawGrid = DrawGrid & ared
                End If
            End If
            
            If BSGAME(dagame).p1grid(temp, temp2).id = "O" Then
                If BSGAME(dagame).p1grid(temp, temp2).hit = True Then
                    DrawGrid = DrawGrid & "¤"
                Else
                    DrawGrid = DrawGrid & "+"
                End If
            Else
                DrawGrid = DrawGrid & BSGAME(dagame).p1grid(temp, temp2).id
            End If
            
        Else
            
            If BSGAME(dagame).p2grid(temp, temp2).hit = True Then
                If BSGAME(dagame).p2grid(temp, temp2).id = "O" Then
                    DrawGrid = DrawGrid & awhite
                Else
                    DrawGrid = DrawGrid & ared
                End If
            End If
            If BSGAME(dagame).p2grid(temp, temp2).id = "O" Then
                If BSGAME(dagame).p2grid(temp, temp2).hit = True Then
                    DrawGrid = DrawGrid & "¤"
                Else
                    DrawGrid = DrawGrid & "+"
                End If
            Else
                DrawGrid = DrawGrid & BSGAME(dagame).p2grid(temp, temp2).id
            End If
            
        End If
        If temp2 < 10 Then DrawGrid = DrawGrid & ablue & "--"
        
    Next temp2
    DrawGrid = DrawGrid & Chr(10)
    If temp < 10 Then DrawGrid = DrawGrid & "   |  |  |  |  |  |  |  |  |  |     |  |  |  |  |  |  |  |  |  |" & Chr(10)
    
Next temp
DrawGrid = DrawGrid & areset & awhite & "           ToysForTards                     ToysForTards"

End Function

















