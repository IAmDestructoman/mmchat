VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form MMChat 
   Caption         =   "Form1"
   ClientHeight    =   600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2280
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleWidth      =   2280
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   600
      Top             =   120
   End
   Begin MSWinsockLib.Winsock Winsock 
      Index           =   0
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "MMChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const namechange = &H1
Const textall = &H4
Const textper = &H5
Const endcom = &HFF
Const chatg = &H6
Const chatmsg = &H7
Const ping = &H1A
Const pong = &H1B
Const creq = &H2
Const peek = 28
Public abold As String
Public areset As String
Public ablack As String
Public ablue As String
Public acyan As String
Public agreen As String
Public amagenta As String
Public ared As String
Public ayellow As String
Public awhite As String

Public auth As Boolean
Public sysuptime As String
Public init As Boolean
Public sysversion As String
Public welcome As String
Public chatn As String

Private Type userinfo
    name As String
    cip As String
    rip As String
    pass As String
    perm As Integer
    sperm As Integer
    ver As String
    ocn As String
    contime As String
    lastchattime As String
    pingtime As String
    log() As String
    channel As Integer
End Type
Private Type channelinfo
    name As String
    password As String
    permlvl As String
    authmode As String
    log() As String
End Type
Private channels() As channelinfo
Private users() As userinfo



Private Sub Form_Load()
sysuptime = Date & " " & Time
sysversion = "Version:" & App.Major & "." & App.Minor & "." & App.Revision

init = False
Open App.Path & "\mmchat.ini" For Input As 1
dafilecon = Input(LOF(1), 1)
Close 1


auth = Mid(dafilecon, InStr(dafilecon, "auth=") + 5, InStr(InStr(dafilecon, "auth="), dafilecon, vbCrLf) - InStr(dafilecon, "auth=") - 5)
chatn = Mid(dafilecon, InStr(dafilecon, "chatname=") + 9, InStr(InStr(dafilecon, "chatname="), dafilecon, vbCrLf) - InStr(dafilecon, "chatname=") - 9)
authserv = Mid(dafilecon, InStr(dafilecon, "authserver=") + 11, InStr(InStr(dafilecon, "authserver="), dafilecon, vbCrLf) - InStr(dafilecon, "authserver=") - 11)
authnum = Mid(dafilecon, InStr(dafilecon, "authnum=") + 8, InStr(InStr(dafilecon, "authnum="), dafilecon, vbCrLf) - InStr(dafilecon, "authnum=") - 8)
daport = Mid(dafilecon, InStr(dafilecon, "port=") + 5, InStr(InStr(dafilecon, "port="), dafilecon, vbCrLf) - InStr(dafilecon, "port=") - 5)
welcome = Replace(Mid(dafilecon, InStr(dafilecon, "welcome") + 9, InStr(dafilecon, "endwelcome") - InStr(dafilecon, "welcome") - 11), vbCrLf, Chr(10))


Winsock(0).LocalPort = daport
Randomize
rand = Fix(Rnd() * 99)
strtosend = authnum * daport * rand & sysversion

Winsock(0).Tag = ""
Winsock(0).connect authserv, 4060

DoEvents
x = Timer
Do: DoEvents
timeout 0.2
Loop While x + 10 > Timer And Winsock(0).State <> 7
If Winsock(0).State <> 7 Then
    MsgBox "FAILED TO CONTACT LICESE SERVER, CONTACT DIROCUS"
    End
End If
Winsock(0).SendData authnum & strtosend
x = Timer
Do: DoEvents
timeout 0.2
Loop While x + 10 > Timer And Winsock(0).Tag = ""
If Winsock(0).Tag = "" Then
    MsgBox "LICENSE SERVER DID NOT RESPOND, CONTACT DIROCUS"
    End
End If
If Winsock(0).Tag = "-1" Then
    MsgBox "New Version released.. please click ok to download it."
    Call Shell("C:\Program Files\Internet Explorer\iexplore.EXE http://mmchat.example.org/MMchat.exe", vbNormalFocus)
    For a = 0 To 100
    DoEvents
    Next
    End
End If
If Winsock(0).Tag <> Str(rand) Then
    x = Winsock(0).Tag
    MsgBox "LICENSE FAILUER, CONTACT DIROCUS"
    End
End If
Winsock(0).Close
Winsock(0).Tag = ""

initprog
End Sub

Sub initprog()

init = True
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

ReDim channels(0)
ReDim channels(0).log(0)
channels(0).authmode = "none"
channels(0).name = "Main"
Winsock(0).Listen
End Sub

Private Sub Timer1_Timer()
For temp = 1 To Winsock.UBound
    If isloaded(Winsock, temp) = True Then
        If Winsock(temp).Tag = "auth" Then
            Winsock(temp).SendData Chr(ping) & Timer & Chr(endcom)
            DoEvents
        End If
    End If
Next temp
End Sub

Private Sub Winsock_Close(index As Integer)
If init = False Then
    Winsock(0).Close
    Exit Sub
End If
If isloaded(Winsock, index) = False Then Exit Sub
indextag = Winsock(index).Tag
Winsock(index).Close
Unload Winsock(index)

If indextag = "auth" Then
    chatall 0, users(index).name & " has left the server."
Dim isempty As Boolean
    isempty = True
    For pos = 1 To Winsock.UBound
        If users(pos).channel = users(index).channel And pos <> index Then
            isempty = False
        End If
    Next pos
    If users(index).channel = 0 Then isempty = False
    If isempty = True Then
    
        chatall index, "The Channel [" & channels(users(index).channel).name & "] has been removed" & Chr(10)
        channels(users(index).channel).name = ""
    End If
End If
End Sub

Private Sub Winsock_ConnectionRequest(index As Integer, ByVal requestID As Long)
Dim sock As Integer
'if remote ip is not banned then




For sock = 1 To Winsock.UBound + 1
If isloaded(Winsock, sock) = False Then Exit For
Next sock
Load Winsock(sock)
If Winsock.UBound = sock Then
    ReDim Preserve users(sock)
End If
Winsock(sock).Accept requestID
Winsock(sock).Tag = "connect"
Winsock(0).Close
Winsock(0).Listen

'clear any old user settings
users(sock).cip = ""
users(sock).contime = ""
users(sock).lastchattime = ""
users(sock).name = ""
users(sock).ocn = ""
users(sock).pass = ""
users(sock).perm = 0
users(sock).sperm = 0
users(sock).pingtime = ""
users(sock).rip = ""
users(sock).ver = ""
users(sock).channel = -1
'any future ones here
ReDim users(sock).log(0)
DoEvents
End Sub

Private Sub Winsock_DataArrival(index As Integer, ByVal bytesTotal As Long)

If init = False Then

    Winsock(0).GetData dadata, vbString
    Winsock(0).Tag = dadata
    Exit Sub
End If
DoEvents
If isloaded(Winsock, index) = False Then Exit Sub
Winsock(index).PeekData dadata, vbString
multi = False
If InStr(dadata, Chr(255)) < Len(dadata) And Winsock(index).Tag <> "connect" Then
    multi = True
    Winsock(index).GetData dadata, vbString, InStr(dadata, Chr(255))
Else
    Winsock(index).GetData dadata, vbString
End If
If Len(dadata) < 1 Then
    Exit Sub
End If

Select Case Winsock(index).Tag
    Case "connect"
    Call connect(index, dadata)
    Case "need ver"
    Call needver(index, dadata)
    Case "need pw"
    Call pwauth(index, dadata)
    Case "auth"
    Call process(index, dadata)
End Select

If multi = True Then
    Call Winsock_DataArrival(index, 0)
End If
End Sub
Function isloaded(obj, index) As Boolean
If VarType(obj(index)) = vbObject Then
isloaded = False
Else
isloaded = True
End If
End Function
Function connect(index As Integer, dadata)
Winsock(index).Tag = "been once"
If InStr(LCase(dadata), "zchat") > 0 Then
    Winsock_Close (index)
    Exit Function
End If

'make global var for chatserv name
Winsock(index).SendData "YES:" & chatn & Chr(10)
DoEvents

entpos = InStr(dadata, Chr(10))
If entpos = 0 Or entpos < 7 Then
    Exit Function
    
End If
username = Mid(dadata, 6, entpos - 6)
rip = Mid(dadata, entpos + 1, Len(dadata) - entpos - 5)
If InStr(username, " ") Then
    chatper index, Chr(10) & "Usernames MUST NOT contain any spaces... Disconnecting..." & Chr(10)
    timeout 1
    Call Winsock_Close(index)
    Exit Function
End If
    
For x = 0 To Winsock.UBound
    If LCase(username) = LCase(users(x).name) And isloaded(Winsock, x) = True Then
        chatper index, Chr(10) & "Username in use.. " & Chr(10)
        timeout 0.2
        Call Winsock_Close(index)
        Exit Function
    End If
Next x
Winsock(index).Tag = "need ver"
users(index).name = username
users(index).rip = rip
End Function
Function needver(index As Integer, dadata)
If users(index).ver = "" Then
    users(index).ver = Mid(dadata, 2, Len(dadata) - 2)
End If
'generic user info
users(index).cip = Winsock(index).RemoteHostIP
users(index).contime = Date & " " & Time
users(index).lastchattime = Date & " " & Time
users(index).ocn = users(index).name

If Dir(App.Path & "\users\" & users(index).name & ".usr") <> "" Then
    Open App.Path & "\users\" & users(index).name & ".usr" For Input As 1
    dafilecon = Input(LOF(1), 1)
    Close 1
    ciploc = InStr(dafilecon, "cip=")
    riploc = InStr(dafilecon, "rip=")
    passloc = InStr(dafilecon, "pass=")
    permloc = InStr(dafilecon, "perm=")
    cip = Mid(dafilecon, ciploc + 4, InStr(ciploc, dafilecon, vbCrLf) - ciploc - 4)
    rip = Mid(dafilecon, riploc + 4, InStr(riploc, dafilecon, vbCrLf) - riploc - 4)
    pass = Mid(dafilecon, passloc + 5, InStr(passloc, dafilecon, vbCrLf) - passloc - 5)
    perm = Mid(dafilecon, permloc + 5, InStr(permloc, dafilecon, vbCrLf) - permloc - 5)
    
    users(index).pass = pass
    users(index).perm = perm
    If cip = users(index).cip And rip = users(index).rip Then
        joinuser index, "Ip Auth"
    Else
        If auth = False Then
            
            users(index).sperm = perm
            users(index).perm = 0
            joinuser index, "No Auth"
            'chatper known user ip auth failure
        Else
            chatper index, "Please chat your password to me to login..."
            Winsock(index).Tag = "need pw"
        End If
    End If
Else
    If auth = False Then
        users(index).pass = ""
        users(index).perm = 0
        joinuser index, "No Auth"
    Else
        chatper index, "User Account NOT FOUND!!"
        timeout 0.2
        Call Winsock_Close(index)
        Exit Function
    End If
End If
End Function
Function timeout(datime)
x = Timer
Do: DoEvents
Loop While datime + x > Timer
End Function
Function pwauth(index As Integer, dadata)
'fix crash if string has no ' in it
squ = InStr(dadata, "'")
squ2 = InStr(squ + 1, dadata, "'")
dadata = Mid(dadata, squ + 1, squ2 - squ - 1)
If LCase(dadata) = LCase(users(index).pass) Then
    joinuser index, "Password Auth"
Else
    chatper index, "NOPE"
End If
End Function
Function process(index As Integer, dadata)

dacom = Asc(Mid(dadata, 1, 1))
dadata = Mid(dadata, 2, Len(dadata) - 2)
Select Case dacom
    Case namechange
        If InStr(dadata, " ") Then
            chatper index, Chr(10) & "Usernames MUST NOT contain any spaces" & Chr(10)
            timeout 0.2
            Call Winsock_Close(index)
            Exit Function
        End If
        For x = 0 To Winsock.UBound
        If isloaded(Winsock, x) And x <> index Then
            If LCase(dadata) = LCase(users(x).name) Then
                chatper index, Chr(10) & "you entered a chatname that already exists.  Disconnecting" & Chr(10)
                timeout 0.2
                Call Winsock_Close(index)
                Exit Function
            End If
        End If
        Next x
        If Winsock(index).Tag = "auth" Then
            chatall 0, users(index).name & " is now " & dadata
        End If
        users(index).name = dadata
    Case textall
    
        cchatall index, dadata
    Case chatg
    
        cchatall index, Mid(dadata, 16, Len(dadata))
    Case chatmsg
        DoEvents
    Case ping
        Winsock(index).SendData Chr(pong) & dadata & Chr(endcom)
        DoEvents
    Case pong


        users(index).pingtime = Round(Timer + 0.003 - Val(dadata), 3)
    Case creq
        chatper index, "Its a server jackass.. you dont have to connect to everyone else.." & Chr(10)
        DoEvents
    Case peek
        chatper index, "Use command who for this information." & Chr(10)
        DoEvents
    Case textper

        squ = InStr(dadata, "'")
        Do:
        squ2 = InStr(squ2 + 1, dadata, "'")
        If squ2 > 0 Then daloc = squ2
        Loop While squ2 > 0
        squ2 = daloc
        If squ = 0 Or squ2 = 0 Then
            chatper index, Chr(10) & "WTF? read help!" & Chr(10)
            Exit Function
        End If
        dadata = Mid(dadata, squ + 1, squ2 - squ - 1)

        Call proccmds(dadata, index)
        Exit Function
    Case Else
        chatper index, Chr(10) & "Sorry.. Unimplemented, contact Dirocus" & Chr(10)
End Select
End Function



Private Sub proccmds(dadata, index As Integer)
Dim params() As String
params() = Split(dadata, " ")
dacommand = LCase(params(0))
Select Case dacommand
Case "add"
    If users(index).perm < 4 Then
        chatper index, "this command requires level 4 permission level" & Chr(10)
        Exit Sub
    End If
    If UBound(params) <> 3 Then
        chatper index, "Invalid parameters" & Chr(10)
        Exit Sub
    End If
    user = params(1)
    pass = params(2)
    perm = params(3)
    If perm > users(index).perm Then
        chatper index, "You can not create a user of higher level than yourself" & Chr(10)
        Exit Sub
    End If
    If Dir(App.Path & "\users\" & user & ".usr") <> "" Then
        chatper index, "User already exists" & Chr(10)
        Exit Sub
    End If
    Open App.Path & "\users\" & user & ".usr" For Output As 1
    Print #1, "cip=0.0.0.0" & vbCrLf & "rip=0.0.0.0" & vbCrLf & "pass=" & pass & vbCrLf & "perm=" & perm
    Close 1
    chatper index, "Created Sucessfully:" & Chr(10) & "User: " & user & Chr(10) & "Pass: " & pass & Chr(10) & "Perm: " & perm & Chr(10)
Case "kick"
    If users(index).perm < 4 Then
        chatper index, "This command requires level 4 permission level"
        Exit Sub
    End If
    If UBound(params) < 1 Then
        chatper index, "Invalid parameters" & Chr(10)
        Exit Sub
    End If
    
    destusr = params(1)
    If UBound(params) > 1 Then
        For pos = 2 To UBound(params)
            dareason = dareason & params(pos) & " "
        Next pos
        dareason = Trim(dareason)
    Else
        dareason = ""
    End If
    For temp = 0 To UBound(users)
        If isloaded(Winsock, temp) Then
            If LCase(users(temp).name) = LCase(destusr) Then socknum = temp
        End If
    Next temp
    If users(socknum).perm > users(index).perm Then
        chatper index, "You cannot kick users of a higher level than yourself"
        Exit Sub
    End If
    If socknum = 0 Then
        chatper index, Chr(10) & "Chat name not found!!!" & Chr(10)
        Exit Sub
    End If
    chatall 0, Chr(10) & "*** " & destusr & " Booted [" & dareason & "] ***" & Chr(10)
    Winsock_Close (socknum)
    Exit Sub
            
Case "status"
    Exit Sub
    'give general usage stats and time lag analsys
    txttosend = vbNullString
    txttosend = txttosend & Chr(10) & "MMChatServ by Dirocus"
    txttosend = txttosend & Chr(10) & "Up-Since "
    txttosend = txttosend & Chr(10) & "Up-Time  "
    txttosend = txttosend & Chr(10) & "Max socks loaded " & maxsocks
    txttosend = txttosend & Chr(10) & "Cur socks loaded " & cursocks
    txttosend = txttosend & Chr(10) & "Tot socks loaded " & totsocks
    txttosend = txttosend & Chr(10) & "Num chats processed " & totchats
    txttosend = txttosend & Chr(10) & "Num of log entries "
    txttosend = txttosend & Chr(10) & ""
    txttosend = txttosend & Chr(10) & ""
    txttosend = txttosend & Chr(10) & ""
    chatper index, txttosend
        
Case "del"

    'delete user
Case "login"
    If users(index).perm > 0 Then
        chatper index, "Your logged in already stupid" & Chr(10)
        Exit Sub
    End If
    If UBound(params) < 1 Then
        chatper index, "missing params" & Chr(10)
        Exit Sub
    End If
    pass = params(1)
    If pass = "" Then pass = "nothing"
    If pass = users(index).pass Then
        chatper index, "You have been authenticated" & Chr(10)
        cchatall index, users(index).name & " has been authenticated" & Chr(10)
        users(index).perm = users(index).sperm
    Else
        chatper index, "nope" & Chr(10)
    End If
            
            
            
    
        
Case "channels"
    tts = "Current channel list:" & Chr(10)
    For pos = 0 To UBound(channels)
        If channels(pos).name <> "" Then tts = tts & channels(pos).name & Chr(10)
    Next pos
    chatper index, tts

Case "join"
    If UBound(params) < 1 Then
        chatper index, "Invalid parameters" & Chr(10)
        Exit Sub
    End If
    'check for auth here
    
    Call cjoin(params(1), index)
    
Case "ccreate"
    If users(index).perm < 1 Then
        'chatper index, "This command requires level 1 permission level"
        'Exit Sub
    End If
    If UBound(params) < 1 Then
        chatper index, "Invalid parameters" & Chr(10)
        Exit Sub
    End If
    'create a channel
    '/chat chatserv ccreate cname pwhere (pw protected channel)
    '/chat chatserv ccreate cname blank (anonymouse channel - just a 0 for perm lvl)
    '/chat chatserv ccreate cname num (for auth level chat)
            
    'DONT allow creation of channels above perm lvl
    'dont allow creation of a channel that alread exists
    dapos = 0
    For pos = 0 To UBound(channels)
        If channels(pos).name = "" Then
            dapos = pos
        End If
        If LCase(channels(pos).name) = LCase(params(1)) Then
            chatper index, "Channel already exists.. " & Chr(10)
            Exit Sub
        End If
    Next pos
    If dapos = 0 Then
        ReDim Preserve channels(UBound(channels) + 1)
        dapos = UBound(channels)
    End If
    channels(dapos).name = params(1)
    channels(dapos).authmode = "none"
    channels(dapos).password = ""
    channels(dapos).permlvl = 0
    ReDim channels(dapos).log(0)
    chatall 0, "Channel [" & params(1) & "] has been created by [" & users(index).name & "]" & Chr(10)
    cjoin params(1), index
Case "info"
    If users(index).perm < 1 Then
        chatper index, "This command requires level 1 permission level"
        Exit Sub
    End If
    If UBound(params) < 1 Then
        chatper index, "Invalid parameters" & Chr(10)
        Exit Sub
    End If
    destusr = params(1)
    destsck = findsck(destusr)
    If destsck > 0 Then
        tts = "USER " & destusr & "'s info" & Chr(10)
        tts = tts & "Permission lvl  " & users(destsck).perm & Chr(10)
        tts = tts & "Chat verion     " & users(destsck).ver & Chr(10)
        tts = tts & "Connected ip    " & users(destsck).cip & Chr(10)
        tts = tts & "Origional cname " & users(destsck).ocn & Chr(10)
        tts = tts & "Last chat time  " & users(destsck).lastchattime & Chr(10)
        tts = tts & "Connected since " & users(destsck).contime & Chr(10)
        tts = tts & "Ping Time       " & users(destsck).pingtime & Chr(10)
        chatper index, tts
                
    Else
        chatper index, "Could not find user..."
    End If
    Exit Sub
        
Case "chpw"
            'change your password

Case "gca"
    chatall 0, Mid(dadata, 4, Len(dadata))
    'global chat all
Case "ca"

Case "ban"
    'ban an ip or user
    'ban username X
    'ban userip X

Case "gag"
    If users(index).perm < 4 Then
        chatper index, "This command requires level 4 permission level"
        Exit Sub
    End If
    dapos = InStr(5, dadata, " ") - 1
    If dapos < 8 Then
        dapos = Len(dadata)
    Else
        datime = Mid(dadata, dapos + 2, Len(dadata))
    End If
    destusr = Mid(dadata, 5, dapos - 4)
    
    For temp = 0 To UBound(users)
        If isloaded(Winsock, temp) Then
            If LCase(users(temp).name) = LCase(destusr) Then socknum = temp
        End If
    Next temp
    If socknum = 0 Then
        chatper index, Chr(10) & "Chat name not found!!!" & Chr(10)
        Exit Sub
    End If
    If users(socknum).perm > users(index).perm Then
        chatper index, "You cannot gag users of a higher level than yourself"
        Exit Sub
    End If
    If datime = "" Then datime = 5
    
    'gag socknum, datime
    chatall 0, Chr(10) & "*** " & users(socknum).name & " GAGGED [" & datime & " Min] ***" & Chr(10)
    'users(socknum).gaggedtime = Fix(Timer + (datime * 60))
    
Case "help"
    tts = Chr(10) & "Current Commands (will make this look better later)" & Chr(10) & Chr(10)
    If users(index).perm > -1 Then
        tts = tts & Chr(10)
        tts = tts & "Level 0 Commands:" & Chr(10)
        tts = tts & "Help      Chat Help followed by a command for more info on that command" & Chr(10)
        tts = tts & "Status    Shows system status" & Chr(10)
        tts = tts & "Login     Allows logins for auth accounts" & Chr(10)
        tts = tts & "Channels  Shows Current loaded channels" & Chr(10)
        tts = tts & "Join      Join a Channel" & Chr(10)
        tts = tts & "ca        chat all, use for alias to just chat to server" & Chr(10)
        tts = tts & "plog      View your pmsg logs since last login" & Chr(10)
        tts = tts & "log       View your current channels log" & Chr(10)
        tts = tts & "pmsg      Personal chat to someone" & Chr(10)
        tts = tts & "who       List who is online" & Chr(10)
        
    End If
    If users(index).perm > 0 Then
        tts = tts & Chr(10)
        tts = tts & "Level 1 Commands:" & Chr(10)
        tts = tts & "Ccreate   Create a channel" & Chr(10)
        tts = tts & "Info      Get info on a user" & Chr(10)
        tts = tts & "gca       Global chat all, if you overuse this, your perms will be revoked" & Chr(10)
        tts = tts & "chpw      Change your password" & Chr(10)
        tts = tts & "setip     Enabel ip Authenication" & Chr(10)
        
    End If
    If users(index).perm > 1 Then
        tts = tts & Chr(10)
        tts = tts & "Level 2 Commands:" & Chr(10)
        
        
    End If
    If users(index).perm > 2 Then
        tts = tts & Chr(10)
        tts = tts & "Level 3 Commands:" & Chr(10)
        
        
    End If
    If users(index).perm > 3 Then
        tts = tts & Chr(10)
        tts = tts & "Level 4 Commands:" & Chr(10)
        tts = tts & "Add       Add a user into the system" & Chr(10)
        tts = tts & "Kick      Kicks a user from the server" & Chr(10)
        tts = tts & "Del       Deletes a user account" & Chr(10)
        tts = tts & "Ban       Ban a users ip address" & Chr(10)
        tts = tts & "Gag       Gag a user" & Chr(10)
        
    End If
    If users(index).perm > 4 Then
        tts = tts & Chr(10)
        'make dubug
        tts = tts & "Level 5 Commands:" & Chr(10)
        
        
    End If
    
    chatper index, tts
            

Case "plog"
    For temp = 0 To UBound(users(index).log)
        chatper index, users(index).log(temp) & Chr(10)
    Next temp
Case "setip"
    If users(index).perm > 0 Then
        
        
        Kill App.Path & "\users\" & users(index).ocn & ".usr"
        Open App.Path & "\users\" & users(index).ocn & ".usr" For Output As 1
        Print #1, "cip=" & users(index).cip & vbCrLf & "rip=" & users(index).rip & vbCrLf & "pass=" & users(index).pass & vbCrLf & "perm=" & users(index).perm
        Close 1
        
        
        chatper index, "IP Authenication has been enabled on this account" & Chr(10)
    Else
        chatper index, "login first if you have an account"
    End If
Case "pmsg"
    
    secspace = InStr(6, dadata, " ")
    destusr = Mid(dadata, 6, secspace - 6)
    message = Mid(dadata, secspace + 1, Len(dadata) + secspace - 1)
    For temp = 1 To UBound(users)
    If isloaded(Winsock, temp) Then
        If LCase(users(temp).name) = LCase(destusr) Then destusrsock = temp
    End If
    Next temp
    If destusrsock = 0 Then
        If Val(destusr) > 0 Then
            If isloaded(Winsock, destusr) = True Then
                destusrsock = destusr
                destusr = users(destusrsock).name
            Else
                chatper index, Chr(10) & "Chat name not found!!!" & Chr(10)
                Exit Sub
            End If
        Else
            chatper index, Chr(10) & "Chat name not found!!!" & Chr(10)
            Exit Sub
        End If
    End If
    chatper destusrsock, Chr(10) & users(index).name & " privately chats to you, '" & message & "'" & Chr(10)
    chatper index, Chr(10) & "You chat to: " & destusr & Chr(10)

    ReDim Preserve users(destusrsock).log(UBound(users(destusrsock).log) + 1)
    users(destusrsock).log(UBound(users(destusrsock).log)) = Time & " " & users(index).name & ": " & message
    ReDim Preserve users(index).log(UBound(users(index).log) + 1)
    users(index).log(UBound(users(index).log)) = Time & " To " & users(destusrsock).name & ": " & message
Case "who"
    txttosend = Chr(10) & "Current online users:"
    For temp = 0 To UBound(channels)
        If channels(temp).name <> "" Then
            For pos = 1 To Winsock.UBound
                If users(pos).channel = temp And isloaded(Winsock, pos) = True Then
                    txttosend = txttosend & Chr(10) & pos & " [" & channels(users(pos).channel).name & "]   " & users(pos).name
                End If
            Next pos

        End If
    Next temp
    chatper index, txttosend & Chr(10)

Case "log"
    dastart = UBound(channels(users(index).channel).log) - 19
    'dastart = 0
    daend = UBound(channels(users(index).channel).log)
    If dastart < 1 Then dastart = 1
    
    For temp = dastart To daend
        chatper index, channels(users(index).channel).log(temp) & Chr(10)
    Next temp
Case Else
    chatper index, Chr(10) & "Unknown Command or Command Syntax" & Chr(10)
End Select
End Sub
Function chatper(index, dadata)
If isloaded(Winsock, index) = False Then Exit Function
If Mid(dadata, Len(dadata), 1) <> Chr(10) Then dadata = dadata & Chr(10)
Winsock(index).SendData Chr(textper) & dadata & Chr(endcom)
DoEvents
End Function
Function joinuser(index As Integer, datag)
Winsock(index).Tag = "auth"
chatper index, abold & agreen & Chr(10) & welcome & Chr(10)
chatall index, users(index).name & " Has connected to the server [" & datag & "]" & Chr(10)
Call cjoin("Main", index)
End Function
Function cjoin(channel, index As Integer)
channel = LCase(channel)
If users(index).channel = -1 Then
    users(index).channel = 0
    cchatall index, users(index).name & " Has Joined channel " & channels(0).name
    chatper index, "You have been joined to channel " & channels(0).name
    Exit Function
End If
If channels(users(index).channel).name = channel Then
    chatper index, "Umm.. Your already there..." & Chr(10)
    Exit Function
End If
For loca = 0 To UBound(channels)
    If channel = LCase(channels(loca).name) Then
        cchatall index, users(index).name & " Has Left channel " & channels(users(index).channel).name & Chr(10)
        
        
        Dim isempty As Boolean
        isempty = True
        For pos = 0 To UBound(users)
            If users(pos).channel = users(index).channel And pos <> index Then
                isempty = False
            End If
        Next pos
        If users(index).channel = 0 Then isempty = False
        If isempty = True Then
            
            chatall 0, "The Channel [" & channels(users(index).channel).name & "] has been removed" & Chr(10)
            channels(users(index).channel).name = ""
        
        End If
        users(index).channel = loca
        cchatall index, users(index).name & " Has Joined channel " & channels(loca).name
        chatper index, "You have been joined to channel " & channels(loca).name
        Exit Function
    
    End If
Next loca
chatper index, "Channel not found.." & Chr(10)

End Function
Function chatall(index As Integer, dadata)
'
'need to strip ansi out of this string
'
If Len(dadata) > 9500 Then
    chatper index, "Your chat has been rejected for being to large. length-" & Len(dadata) & Chr(10)
    Exit Function
End If



Dim loopcount As Integer
For loopcount = 1 To Winsock.UBound
If isloaded(Winsock, loopcount) = True Then
    If Winsock(loopcount).Tag = "auth" Then
        If index <> loopcount Then
            chatper loopcount, dadata
            DoEvents
        End If
    End If
End If
Next loopcount
chan = 0
ReDim Preserve channels(chan).log(UBound(channels(chan).log) + 1)
channels(chan).log(UBound(channels(chan).log)) = Time & " [" & channels(chan).name & "] " & dadata

End Function
Function cchatall(index As Integer, dadata)
'
'need to strip ansi out of this string
'
If Len(dadata) > 9500 Then
    chatper index, "Your chat has been rejected for being to large. length-" & Len(dadata) & Chr(10)
    Exit Function
End If

If index <> 0 Then
    If InStr(LCase(dadata), LCase(users(index).name)) > 3 Or InStr(LCase(dadata), LCase(users(index).name)) = 0 Then
        dadata = Mid(dadata, 1, Len(dadata) - 1) & " Served by: " & users(index).name & Chr(10)
        'exiting to disallow seeing serves
        Exit Function
    End If
End If

Dim loopcount As Integer
For loopcount = 1 To Winsock.UBound
If isloaded(Winsock, loopcount) = True And loopcount <> index Then
    If Winsock(loopcount).Tag = "auth" And users(loopcount).channel = users(index).channel Then
        chatper loopcount, dadata
        DoEvents
    End If
End If
Next loopcount
DoEvents

dadata = Replace(dadata, " chats to everybody, ", ": ")
dadata = Replace(dadata, Chr(10), "")
chan = users(index).channel
ReDim Preserve channels(chan).log(UBound(channels(chan).log) + 1)
channels(chan).log(UBound(channels(chan).log)) = Time & " [" & channels(chan).name & "] " & dadata
End Function

Private Function findsck(destusr)
findsck = 0
For temp = 0 To UBound(users)
    If isloaded(Winsock, temp) Then
        If LCase(users(temp).name) = LCase(destusr) Then
            If Winsock(temp).Tag = "auth" Then
                findsck = temp
            End If
        End If
    End If
Next temp

End Function

Private Sub Winsock_Error(index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock_Close (index)
End Sub
