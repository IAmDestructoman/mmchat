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
    'O open water
    'A Aircraft Carrier
    'B BattleShip
    '3 etc
    hit As Boolean
End Type
Public Type bsgamevars
    p1n As String
    p2n As String
    p1grid(1 To 10, 1 To 10) As bsloc
    p2grid(1 To 10, 1 To 10) As bsloc
    turn As Integer
    p1r As Boolean
    p2r As Boolean
    p1sunken As Byte
    p2sunken As Byte
    
    
End Type
Private inited  As Boolean
Private BSGAME() As bsgamevars
Private Function issunken(dagame As Integer, playernum As Integer, datype As String) As Boolean
issunken = True
For temp = 1 To 10
    For temp2 = 1 To 10
        If playernum = 1 Then
            If LCase(BSGAME(dagame).p1grid(temp, temp2).id) = datype Then
                If BSGAME(dagame).p1grid(temp, temp2).hit = False Then issunken = False
            End If
        Else
            If LCase(BSGAME(dagame).p2grid(temp, temp2).id) = datype Then
                If BSGAME(dagame).p2grid(temp, temp2).hit = False Then issunken = False
            End If
        End If
    Next temp2
Next temp
If issunken = False Then Exit Function
If playernum = 1 Then
    BSGAME(dagame).p1sunken = BSGAME(dagame).p1sunken + 1
Else
    BSGAME(dagame).p2sunken = BSGAME(dagame).p2sunken + 1
End If
End Function

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
    .turn = 0
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
    If inited <> True Then
    ReDim BSGAME(0)
    inited = True
    End If
End Function

Public Function GetCommands() As String

    GetCommands = ",BSHelp,0,BSChallenge,0,BSFire,0,BSShow,0,BSReady,0,BSPlace,0,BSRemove,0,BSQuit,0"

End Function
Public Function init(othervars) As String

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
Dim dagame As Integer
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
    main = "0" & abold
main = main & acyan & "*******--------******" & Chr(10)
main = main & ayellow & "BATTLESHIP By Dirocus" & Chr(10)
main = main & acyan & "*******--------******" & Chr(10) & Chr(10)

main = main & ablue & "Basic Commands and Procedures:" & Chr(10)
main = main & ared & "BSHelp             " & acyan & "|This help file." & Chr(10)
main = main & ared & "BSChallenge <name> " & acyan & "|Sends a game request to an opponent" & Chr(10)
main = main & ared & "                   " & acyan & "|NOTE: Use the person's Orig. Name as listed on Who" & Chr(10)
main = main & ared & "BSShow             " & acyan & "|Show the maps of both your ships and your targets." & Chr(10)
main = main & ared & "BSPlace <type> <a1-j10> <r|d>" & Chr(10)
main = main & ared & "                   " & acyan & "|Use to set up your ships before the game.  <type> is" & Chr(10)
main = main & ared & "                   " & acyan & "|f, b, c, d, s, or p, respectively for each ship.           " & Chr(10)
main = main & ared & "                   " & acyan & "|<a1-j10> is the point you wish to place at, and <r|d>" & Chr(10)
main = main & ared & "                   " & acyan & "|determines whether the ship will stretch right or down" & Chr(10)
main = main & ared & "                   " & acyan & "|from your original input point." & Chr(10)
main = main & ared & "BSRemove           " & acyan & "|Used during placement.  Removes a ship from the board." & Chr(10)
main = main & ared & "BSReady            " & acyan & "|End ship placement." & Chr(10)
main = main & ared & "BSFire <a-j><1-10> " & acyan & "|Use when it's your turn to fire your shot." & Chr(10)
main = main & ared & "BSQuit             " & acyan & "|Quits and forfeits the game." & Chr(10) & Chr(10)

main = main & ared & "Ship Information:" & Chr(10) & Chr(10)

main = main & "-=Class A:Aircraft Carrier=-" & Chr(10)
main = main & "Size: Large (6)" & Chr(10) & Chr(10)

main = main & "-=Class B:Battleship=-" & Chr(10)
main = main & "Size: Large (5)" & Chr(10) & Chr(10)

main = main & "-=Class C:Cruiser=-" & Chr(10)
main = main & "Size: Medium-Large (4)" & Chr(10) & Chr(10)

main = main & "-=Class D:Destroyer=-" & Chr(10)
main = main & "Size: Medium (3)" & Chr(10) & Chr(10)

main = main & "-=Class S:Submarine=-" & Chr(10)
main = main & "Size: Medium (3)" & Chr(10) & Chr(10)

main = main & "-=Class P:Patrol Boat=-" & Chr(10)
main = main & "Size: Small (2)" & Chr(10)


    
    
    
    
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
            main = "3" & abold & ayellow & "BattleShip War Started!  " & Chr(10) & dauser2 & " Wins(0) Losses(0)" & Chr(10) & dauser & " Wins(0) Losses(0)"
            '& "~0" & abold & ayellow & "BattleShip War Started!  " & Chr(10) & dauser2 & " Wins(0) Losses(0)" & Chr(10) & dauser & " Wins(0) Losses(0)" & "~6" & dauser2 & "," & abold & ayellow & "BattleShip War Started!  " & Chr(10) & dauser2 & " Wins(0) Losses(0)" & Chr(10) & dauser & " Wins(0) Losses(0)"
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
    If BSGAME(dagame).turn = 0 Then
        main = "0Both players must use BSReady to start the game."
        Exit Function
    End If
    If BSGAME(dagame).turn <> playernum Then
        main = "0Its not your turn!."
        Exit Function
    End If
    syntax = "0" & Chr(10) & ared & "BSFire <a-j><1-10> " & acyan & "|Use when it's your turn to fire your shot." & Chr(10)
    If UBound(params) <> 1 Then
        main = syntax
        Exit Function
    End If
    
    slet = LCase(Mid(params(1), 1, 1))
    snum = Mid(params(1), 2, Len(params(1)))
    valid = True
    
    If Asc(slet) < Asc("a") Or Asc(slet) > Asc("j") Then valid = False
    If Val(snum) < 1 Or Val(snum) > 10 Then valid = False
    
    If valid = False Then
        main = syntax
        Exit Function
    End If

    slet = Asc(slet) - 96

    If playernum = 1 Then
        If BSGAME(dagame).p2grid(Val(snum), slet).hit = True Then
            main = "You've already shot there! pick again."
            Exit Function
        Else
            BSGAME(dagame).p2grid(Val(snum), slet).hit = True
            main = "6" & BSGAME(dagame).p2n & "," & BSGAME(dagame).p1n & " Shoots at " & UCase(params(1)) & " and..." & Chr(10) & DrawGrid(BSGAME(dagame).p2n)
            main = main & "~0 You shoot and..." & Chr(10) & DrawGrid(BSGAME(dagame).p1n)
            
            Select Case LCase(BSGAME(dagame).p2grid(Val(snum), slet).id)
            
            Case "a"
                If issunken(dagame, 2, "a") Then
                    main = main & "~0SHAAaaaPOW! You 've SUNK " & BSGAME(dagame).p2n & "'s Aircraft Carrier!"
                    main = main & "~6" & BSGAME(dagame).p2n & "," & BSGAME(dagame).p1n & " has SUNK your Aircraft Carrier!"
                Else
                    main = main & "~0SHAAaaaPOW! You 've hit " & BSGAME(dagame).p2n & "'s Aircraft Carrier!"
                    main = main & "~6" & BSGAME(dagame).p2n & "," & BSGAME(dagame).p1n & " has hit your Aircraft Carrier!"
                End If
            Case "b"
                If issunken(dagame, 2, "b") Then
                    main = main & "~0SHAAaaaPOW! You 've SUNK " & BSGAME(dagame).p2n & "'s Battleship!"
                    main = main & "~6" & BSGAME(dagame).p2n & "," & BSGAME(dagame).p1n & " has SUNK your Battleship!"
                Else
                    main = main & "~0SHAAaaaPOW! You 've hit " & BSGAME(dagame).p2n & "'s Battleship!"
                    main = main & "~6" & BSGAME(dagame).p2n & "," & BSGAME(dagame).p1n & " has hit your Battleship!"
                End If
            Case "c"
                If issunken(dagame, 2, "c") Then
                    main = main & "~0SHAAaaaPOW! You 've SUNK " & BSGAME(dagame).p2n & "'s Cruiser!"
                    main = main & "~6" & BSGAME(dagame).p2n & "," & BSGAME(dagame).p1n & " has SUNK your Destroyer!"
                Else
                    main = main & "~0SHAAaaaPOW! You 've hit " & BSGAME(dagame).p2n & "'s Cruiser!"
                    main = main & "~6" & BSGAME(dagame).p2n & "," & BSGAME(dagame).p1n & " has hit your Destroyer!"
                End If
            Case "d"
                If issunken(dagame, 2, "d") Then
                    main = main & "~0SHAAaaaPOW! You 've SUNK " & BSGAME(dagame).p2n & "'s Destroyer!"
                    main = main & "~6" & BSGAME(dagame).p2n & "," & BSGAME(dagame).p1n & " has SUNK your Destroyer!"
                Else
                    main = main & "~0SHAAaaaPOW! You 've hit " & BSGAME(dagame).p2n & "'s Destroyer!"
                    main = main & "~6" & BSGAME(dagame).p2n & "," & BSGAME(dagame).p1n & " has hit your Destroyer!"
                End If
            Case "p"
                If issunken(dagame, 2, "p") Then
                    main = main & "~0SHAAaaaPOW! You 've SUNK " & BSGAME(dagame).p2n & "'s PT Boat!"
                    main = main & "~6" & BSGAME(dagame).p2n & "," & BSGAME(dagame).p1n & " has SUNK your PT Boat!"
                Else
                    main = main & "~0SHAAaaaPOW! You 've hit " & BSGAME(dagame).p2n & "'s PT Boat!"
                    main = main & "~6" & BSGAME(dagame).p2n & "," & BSGAME(dagame).p1n & " has hit your PT Boat!"
                End If
            Case "s"
                If issunken(dagame, 2, "s") Then
                    main = main & "~0SHAAaaaPOW! You 've SUNK " & BSGAME(dagame).p2n & "'s Submarine!"
                    main = main & "~6" & BSGAME(dagame).p2n & "," & BSGAME(dagame).p1n & " has SUNK your Submarine!"
                Else
                    main = main & "~0SHAAaaaPOW! You 've hit " & BSGAME(dagame).p2n & "'s Submarine!"
                    main = main & "~6" & BSGAME(dagame).p2n & "," & BSGAME(dagame).p1n & " has hit your Submarine!"
                End If
            Case Else
                main = main & "~0Neener neener! You Missed!"
                main = main & "~6" & BSGAME(dagame).p2n & "," & BSGAME(dagame).p1n & " MISSED!"
                
            End Select
            BSGAME(dagame).turn = 2
            If BSGAME(dagame).p2sunken = 6 Then
                'GAME OVER!!
                main = main & "~3" & ayellow & "BATTLE SHIP WAR OVER.  WINNER:" & BSGAME(dagame).p1n
                BSGAME(dagame).p1n = vbNullString
                BSGAME(dagame).p2n = vbNullString
                    
            End If
        End If
        
        
    Else
        If BSGAME(dagame).p1grid(Val(snum), slet).hit = True Then
            main = "You've already shot there! pick again."
            Exit Function
        Else
            BSGAME(dagame).p1grid(Val(snum), slet).hit = True
            main = "6" & BSGAME(dagame).p1n & "," & BSGAME(dagame).p2n & " Shoots at " & UCase(params(1)) & " and..." & Chr(10) & DrawGrid(BSGAME(dagame).p1n)
            main = main & "~0 You shoot and..." & Chr(10) & DrawGrid(BSGAME(dagame).p2n)
            
            Select Case LCase(BSGAME(dagame).p1grid(Val(snum), slet).id)
            
            Case "a"
                If issunken(dagame, 1, "a") Then
                    main = main & "~0SHAAaaaPOW! You 've SUNK " & BSGAME(dagame).p1n & "'s Aircraft Carrier!"
                    main = main & "~6" & BSGAME(dagame).p1n & "," & BSGAME(dagame).p2n & " has SUNK your Aircraft Carrier!"
                Else
                    main = main & "~0SHAAaaaPOW! You 've hit " & BSGAME(dagame).p1n & "'s Aircraft Carrier!"
                    main = main & "~6" & BSGAME(dagame).p1n & "," & BSGAME(dagame).p2n & " has hit your Aircraft Carrier!"
                End If
            Case "b"
                If issunken(dagame, 1, "b") Then
                    main = main & "~0SHAAaaaPOW! You 've SUNK " & BSGAME(dagame).p1n & "'s Battleship!"
                    main = main & "~6" & BSGAME(dagame).p1n & "," & BSGAME(dagame).p2n & " has SUNK your Battleship!"
                Else
                    main = main & "~0SHAAaaaPOW! You 've hit " & BSGAME(dagame).p1n & "'s Battleship!"
                    main = main & "~6" & BSGAME(dagame).p1n & "," & BSGAME(dagame).p2n & " has hit your Battleship!"
                End If
            Case "c"
                If issunken(dagame, 1, "c") Then
                    main = main & "~0SHAAaaaPOW! You 've SUNK " & BSGAME(dagame).p1n & "'s Cruiser!"
                    main = main & "~6" & BSGAME(dagame).p1n & "," & BSGAME(dagame).p2n & " has SUNK your Destroyer!"
                Else
                    main = main & "~0SHAAaaaPOW! You 've hit " & BSGAME(dagame).p1n & "'s Cruiser!"
                    main = main & "~6" & BSGAME(dagame).p1n & "," & BSGAME(dagame).p2n & " has hit your Destroyer!"
                End If
            Case "d"
                If issunken(dagame, 1, "d") Then
                    main = main & "~0SHAAaaaPOW! You 've SUNK " & BSGAME(dagame).p1n & "'s Destroyer!"
                    main = main & "~6" & BSGAME(dagame).p1n & "," & BSGAME(dagame).p2n & " has SUNK your Destroyer!"
                Else
                    main = main & "~0SHAAaaaPOW! You 've hit " & BSGAME(dagame).p1n & "'s Destroyer!"
                    main = main & "~6" & BSGAME(dagame).p1n & "," & BSGAME(dagame).p2n & " has hit your Destroyer!"
                End If
            Case "p"
                If issunken(dagame, 1, "p") Then
                    main = main & "~0SHAAaaaPOW! You 've SUNK " & BSGAME(dagame).p1n & "'s PT Boat!"
                    main = main & "~6" & BSGAME(dagame).p1n & "," & BSGAME(dagame).p2n & " has SUNK your PT Boat!"
                Else
                    main = main & "~0SHAAaaaPOW! You 've hit " & BSGAME(dagame).p1n & "'s PT Boat!"
                    main = main & "~6" & BSGAME(dagame).p1n & "," & BSGAME(dagame).p2n & " has hit your PT Boat!"
                End If
            Case "s"
                If issunken(dagame, 1, "s") Then
                    main = main & "~0SHAAaaaPOW! You 've SUNK " & BSGAME(dagame).p1n & "'s Submarine!"
                    main = main & "~6" & BSGAME(dagame).p1n & "," & BSGAME(dagame).p2n & " has SUNK your Submarine!"
                Else
                    main = main & "~0SHAAaaaPOW! You 've hit " & BSGAME(dagame).p1n & "'s Submarine!"
                    main = main & "~6" & BSGAME(dagame).p1n & "," & BSGAME(dagame).p2n & " has hit your Submarine!"
                End If
            Case Else
                main = main & "~0Neener neener! You Missed!"
                main = main & "~6" & BSGAME(dagame).p1n & "," & BSGAME(dagame).p2n & " MISSED!"
                
            End Select
            BSGAME(dagame).turn = 1
            If BSGAME(dagame).p1sunken = 6 Then
                'GAME OVER!!
                main = main & "~3BATTLE SHIP WAR OVER.  WINNER:" & BSGAME(dagame).p2n
                BSGAME(dagame).p2n = vbNullString
                BSGAME(dagame).p1n = vbNullString
                    
            End If
        End If
    
    End If
    

Case "bsshow"
main = "0" & DrawGrid(getvar("ocn", othervars))

Case "bsready"
    If playernum = 1 Then
        If BSGAME(dagame).p1r = True Then
            main = "0You already finished placing ships, can not alter now!"
            Exit Function
        End If
    Else
        If BSGAME(dagame).p2r = True Then
            main = "0You already finished placing ships, can not alter now!"
            Exit Function
        End If
    End If
'finsished placing ships, start the game
    If playernum = 1 Then
        allpres = True
        For bsship = 1 To 6
            Select Case bsship
            Case 1
                shiptype = "A"
            Case 2
                shiptype = "B"
            Case 3
                shiptype = "C"
            Case 4
                shiptype = "D"
            Case 5
                shiptype = "P"
            Case 6
                shiptype = "S"
            End Select
            shippres = False
            For temp = 1 To 10
                For temp2 = 1 To 10
                    If BSGAME(dagame).p1grid(temp, temp2).id = shiptype Then shippres = True
                Next temp2
            Next temp
            If shippres = False Then
                allpres = False
            End If
            
        Next bsship
        If allpres = False Then
            main = "0All your ships arnt placed yet."
            Exit Function
        End If
    Else
        allpres = True
        For bsship = 1 To 6
            Select Case bsship
            Case 1
                shiptype = "A"
            Case 2
                shiptype = "B"
            Case 3
                shiptype = "C"
            Case 4
                shiptype = "D"
            Case 5
                shiptype = "P"
            Case 6
                shiptype = "S"
            End Select
            shippres = False
            For temp = 1 To 10
                For temp2 = 1 To 10
                    If BSGAME(dagame).p2grid(temp, temp2).id = shiptype Then shippres = True
                Next temp2
            Next temp
            If shippres = False Then
                allpres = False
            End If
            
        Next bsship
        If allpres = False Then
            main = "0All your ships arnt placed yet."
            Exit Function
        End If
    End If
    If playernum = 1 Then
        If BSGAME(dagame).p2r = True Then
            main = "0" & BSGAME(dagame).p2n & " was finished placing his ships first, he will start" & Chr(10) & "ALL SHIPS PLACED!"
            main = main & "~6" & BSGAME(dagame).p2n & ",ALL SHIPS PLACED, ITS YOUR TURN TO FIRE!"
            BSGAME(dagame).turn = 2
        Else
            main = "0You were first to finish placing your ships, so you will have the starting shot!"
            main = main & "~6" & BSGAME(dagame).p2n & "," & BSGAME(dagame).p1n & " Has finished placing his ships"
        End If
        BSGAME(dagame).p1r = True
    Else
        If BSGAME(dagame).p1r = True Then
            main = "0" & BSGAME(dagame).p1n & " was finished placing his ships first, he will start" & Chr(10) & "ALL SHIPS PLACED!"
            main = main & "~6" & BSGAME(dagame).p1n & ",ALL SHIPS PLACED, ITS YOUR TURN TO FIRE!"
            BSGAME(dagame).turn = 1
        Else
            main = "0You were first to finish placing your ships, so you will have the starting shot!"
            main = main & "~6" & BSGAME(dagame).p1n & "," & BSGAME(dagame).p2n & " Has finished placing his ships"
        End If
        BSGAME(dagame).p2r = True
    End If

Case "bsplace"
    'place a ship
    If playernum = 1 Then
        If BSGAME(dagame).p1r = True Then
            main = "0You already finished placing ships, can not alter now!"
            Exit Function
        End If
    Else
        If BSGAME(dagame).p2r = True Then
            main = "0You already finished placing ships, can not alter now!"
            Exit Function
        End If
    End If
    
    syntax = "0" & Chr(10)
    syntax = syntax & ared & "BSPlace <type> <a1-j10> <r|d>" & Chr(10)
    syntax = syntax & ared & "                   " & acyan & "|Use to set up your ships before the game.  <type> is" & Chr(10)
    syntax = syntax & ared & "                   " & acyan & "|f, b, c, d, s, or p, respectively for each ship.           " & Chr(10)
    syntax = syntax & ared & "                   " & acyan & "|<a1-j10> is the point you wish to place at, and <r|d>" & Chr(10)
    syntax = syntax & ared & "                   " & acyan & "|determines whether the ship will stretch right or down" & Chr(10)
    syntax = syntax & ared & "                   " & acyan & "|from your original input point." & Chr(10)

    
    
    If UBound(params) <> 3 Then
        main = syntax
        Exit Function
    End If
   
    stype = LCase(params(1))
    slet = LCase(Mid(params(2), 1, 1))
    snum = Mid(params(2), 2, Len(params(2)))
    sdir = LCase(params(3))
    valid = True
    Select Case stype
    Case "a"
    shiplen = 6
    Case "b"
    shiplen = 5
    Case "c"
    shiplen = 4
    Case "d"
    shiplen = 3
    Case "s"
    shiplen = 3
    Case "p"
    shiplen = 2
    Case Else
        valid = False
    End Select
    If Asc(slet) < Asc("a") Then valid = False
    If Asc(slet) > Asc("j") Then valid = False
    If Val(snum) < 1 Or Val(snum) > 10 Then valid = False
    If sdir <> "r" And sdir <> "d" Then valid = False
    If valid = False Then
        main = syntax
        Exit Function
    End If
    'all valid input, check fit
    
    If playernum = 1 Then
        For temp = 1 To 10
            For temp2 = 1 To 10
                If LCase(BSGAME(dagame).p1grid(temp, temp2).id) = stype Then BSGAME(dagame).p1grid(temp, temp2).id = "O"
            Next temp2
        Next temp
    Else
        For temp = 1 To 10
            For temp2 = 1 To 10
                If LCase(BSGAME(dagame).p2grid(temp, temp2).id) = stype Then BSGAME(dagame).p2grid(temp, temp2).id = "O"
            Next temp2
        Next temp
    End If
    
      
    slet = Asc(slet) - 96
    If sdir = "r" Then
        If shiplen + slet - 1 > 10 Then
            main = "0That ship wont fit there."
            Exit Function
        End If
        
        For temp = slet To slet + shiplen - 1
            If playernum = 1 Then
                If BSGAME(dagame).p1grid(snum, temp).id <> "O" Then
                    main = "0Another Ship is in the way."
                    Exit Function
                End If
            Else
                If BSGAME(dagame).p2grid(snum, temp).id <> "O" Then
                    main = "0Another Ship is in the way."
                    Exit Function
                End If
            End If
        Next temp
        
    Else
        If shiplen + snum - 1 > 10 Then
            main = "0That ship wont fit there."
            Exit Function
        End If
        
        For temp = snum To snum + shiplen - 1
            If playernum = 1 Then
                If BSGAME(dagame).p1grid(temp, slet).id <> "O" Then
                    main = "0Another Ship is in the way."
                    Exit Function
                End If
            Else
                If BSGAME(dagame).p2grid(temp, slet).id <> "O" Then
                    main = "0Another Ship is in the way."
                    Exit Function
                End If
            End If
        Next temp
        
    End If
    'it will fit no obstructions, actually do it and bsshow
    If sdir = "r" Then
        For temp = slet To slet + shiplen - 1
            If playernum = 1 Then
                BSGAME(dagame).p1grid(snum, temp).id = UCase(stype)
            Else
                BSGAME(dagame).p2grid(snum, temp).id = UCase(stype)
            End If
        Next temp
    Else
        For temp = snum To snum + shiplen - 1
            If playernum = 1 Then
                BSGAME(dagame).p1grid(temp, slet).id = UCase(stype)
            Else
                BSGAME(dagame).p2grid(temp, slet).id = UCase(stype)
            End If
        Next temp
    End If
    
    main = "0Ship placed." & Chr(10) & DrawGrid(getvar("ocn", othervars))

Case "bsremove"
    If playernum = 1 Then
        If BSGAME(dagame).p1r = True Then
            main = "0You already finished placing ships, can not alter now!"
            Exit Function
        End If
    Else
        If BSGAME(dagame).p2r = True Then
            main = "0You already finished placing ships, can not alter now!"
            Exit Function
        End If
    End If
    If UBound(params) <> 1 Then
        main = "0Yes but what ship do you want to remove?"
        Exit Function
    End If
    params(1) = LCase(params(1))
    Select Case params(1)
    Case "a"
    Case "b"
    Case "c"
    Case "d"
    Case "p"
    Case "s"
    Case Else
        main = "0What ship was that?"
        Exit Function
    End Select
    If playernum = 1 Then
        For temp = 1 To 10
            For temp2 = 1 To 10
                If LCase(BSGAME(dagame).p1grid(temp, temp2).id) = params(1) Then BSGAME(dagame).p1grid(temp, temp2).id = "O"
            Next temp2
        Next temp
    Else
        For temp = 1 To 10
            For temp2 = 1 To 10
                If LCase(BSGAME(dagame).p2grid(temp, temp2).id) = params(1) Then BSGAME(dagame).p2grid(temp, temp2).id = "O"
            Next temp2
        Next temp
    End If
    main = "0" & DrawGrid(getvar("ocn", othervars))
Case "bsquit"
'quit/forefit
main = "0QUITER!! GAME OVER!"
If playernum = 1 Then
main = main & "~6" & BSGAME(dagame).p2n & "," & getvar("ocn", othervars) & " HAS QUIT!! GAME OVER!"
Else
main = main & "~6" & BSGAME(dagame).p2n & "," & getvar("ocn", othervars) & " HAS QUIT!! GAME OVER!"
End If
BSGAME(dagame).p1n = vbNullString
BSGAME(dagame).p2n = vbNullString
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
                DrawGrid = DrawGrid & awhite & Chr(15)
                Else
                DrawGrid = DrawGrid & ared & BSGAME(dagame).p2grid(temp, temp2).id
                End If
                
            Else
                DrawGrid = DrawGrid & ablue & "+"
            End If
        Else
            If BSGAME(dagame).p1grid(temp, temp2).hit = True Then
                If BSGAME(dagame).p1grid(temp, temp2).id = "O" Then
                DrawGrid = DrawGrid & awhite & Chr(15)
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
        DrawGrid = DrawGrid & agreen
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
                    DrawGrid = DrawGrid & Chr(15)
                Else
                    DrawGrid = DrawGrid & ablue & "+"
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
                    DrawGrid = DrawGrid & Chr(15)
                Else
                    DrawGrid = DrawGrid & ablue & "+"
                End If
            Else
                DrawGrid = DrawGrid & BSGAME(dagame).p2grid(temp, temp2).id
            End If
            
        End If
        If temp2 < 10 Then DrawGrid = DrawGrid & ablue & "--"
        
    Next temp2
    DrawGrid = DrawGrid & Chr(10)
    If temp < 10 Then DrawGrid = DrawGrid & ablue & "   |  |  |  |  |  |  |  |  |  |     |  |  |  |  |  |  |  |  |  |" & Chr(10)
    
Next temp
DrawGrid = DrawGrid & areset & awhite & "           ToysForTards                     ToysForTards"

End Function

