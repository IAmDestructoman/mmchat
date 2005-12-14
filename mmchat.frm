VERSION 5.00
Begin VB.Form MMChat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "o"
   ClientHeight    =   4050
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox pending 
      Height          =   450
      Left            =   2640
      TabIndex        =   27
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   960
      TabIndex        =   26
      Text            =   ".5"
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   960
      TabIndex        =   25
      Text            =   "20"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   960
      TabIndex        =   24
      Text            =   ".1"
      Top             =   3000
      Width           =   495
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Out"
      Height          =   255
      Left            =   2040
      TabIndex        =   19
      Top             =   3720
      Width           =   615
   End
   Begin VB.CheckBox Check3 
      Caption         =   "In"
      Height          =   255
      Left            =   1560
      TabIndex        =   18
      Top             =   3720
      Width           =   495
   End
   Begin VB.ListBox List3 
      Enabled         =   0   'False
      Height          =   1425
      Left            =   3960
      TabIndex        =   16
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply"
      Height          =   255
      Left            =   2640
      TabIndex        =   15
      Top             =   3720
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   960
      TabIndex        =   10
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      TabIndex        =   6
      Text            =   "4050"
      Top             =   2280
      Width           =   1575
   End
   Begin VB.ListBox List2 
      Height          =   645
      Left            =   2640
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Text            =   "ChatServ"
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Auth"
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   600
      Top             =   120
   End
   Begin Project1.Socket winsock 
      Index           =   0
      Left            =   2880
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      LocalPort       =   0
      RemoteHost      =   ""
      RemotePort      =   0
      Protocol        =   0
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "Gag point"
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Gag Time"
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Gag Interval"
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Allows Serves"
      Height          =   255
      Left            =   1560
      TabIndex        =   20
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Online"
      Height          =   255
      Left            =   3960
      TabIndex        =   17
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label inpacks 
      Caption         =   "0"
      Height          =   255
      Left            =   7320
      TabIndex        =   14
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "In:"
      Height          =   255
      Left            =   6960
      TabIndex        =   13
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Out:"
      Height          =   255
      Left            =   6960
      TabIndex        =   12
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Packets"
      Height          =   255
      Left            =   6840
      TabIndex        =   11
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Main PW:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label outpacks 
      Caption         =   "0"
      Height          =   255
      Left            =   7320
      TabIndex        =   8
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Port:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Banned"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Chatname:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   855
   End
   Begin VB.Menu filemenu 
      Caption         =   "File"
      Begin VB.Menu fileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu filhelp 
      Caption         =   "Help"
      Begin VB.Menu filehelp 
         Caption         =   "Server Help"
      End
      Begin VB.Menu menuabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "MMChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nid As NOTIFYICONDATA ' trayicon variable

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
Public datime As Variant

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
    gtimer As String
    gaggedtime As String
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
Function chatper(Index, dadata)
If Mid(dadata, 1, 1) <> Chr(10) Then dadata = Chr(10) & dadata
If Mid(dadata, Len(dadata), 1) <> Chr(10) Then dadata = dadata & Chr(10)
If isloaded(winsock, Index) = False Then Exit Function
MMChat.winsock(Index).SendData Chr(textper) & dadata & Chr(endcom)
'doevents
End Function


Private Sub Winsock_onDataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error GoTo errored
If init = False Then
    winsock(0).GetData dadata, vbString
    winsock(0).Tag = dadata
    Exit Sub
End If
multi = False
winsock(Index).PeekData dadata, vbString
dalen = InStr(dadata, Chr(255))
If dalen < Len(dadata) Then multi = True
If winsock(Index).Tag <> "auth" Then
    multi = False
    dalen = Len(dadata)
End If
winsock(Index).GetData dadata, vbString, dalen

If isloaded(winsock, Index) = False Then Exit Sub
If Len(dadata) < 1 Then
    Exit Sub
End If
List1.AddItem stripansi(dadata)
List1.ListIndex = List1.ListCount - 1
Do While List1.ListCount > 9
List1.RemoveItem 0
Loop

Select Case winsock(Index).Tag
    Case "connect"
    Call Connect(Index, dadata)
    Case "need ver"
    Call needver(Index, dadata)
    Case "need pw"
    Call pwauth(Index, dadata)
    Case "auth"

    Call process(Index, dadata)
End Select

If multi = True Then

    Call Winsock_onDataArrival(Index, 0)
End If
Exit Sub
errored:
chatper Index, "You have caused an error.." & Chr(10) & "Tech Data:" & Err.Description
timeout 0.5
Winsock_onClose (Index)
End Sub
Sub minimize_to_tray()
Me.Hide
nid.cbSize = Len(nid)
nid.hwnd = Me.hwnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Me.Icon ' the icon will be your mmchat project icon
nid.szTip = "MM Chat Server By:Dirocus" & vbNullChar
Shell_NotifyIcon NIM_ADD, nid
End Sub



Sub timeout(X)
start = Timer
Do: DoEvents
Loop While start + X > Timer
End Sub
Function isloaded(obj, Index) As Boolean
isloaded = False
On Error GoTo errored
If MMChat.winsock(Index).Tag <> "asfd" Then isloaded = True
Exit Function
errored:
isloaded = False
End Function
Function process(Index As Integer, dadata)
'check for main pw
If users(Index).channel = -1 Then
    squ = InStr(dadata, "'")
    Do:
    squ2 = InStr(squ2 + 1, dadata, "'")
    If squ2 > 0 Then daloc = squ2
    Loop While squ2 > 0
    squ2 = daloc
    If squ = 0 Or squ = 0 Then Exit Function
    dadata = Mid(dadata, squ + 1, squ2 - squ - 1)
    dadata = Replace(LCase(dadata), "login ", "")
    If dadata <> Text3 Then
        chatper Index, "Incorrect"
        timeout 0.5
        Winsock_onClose (Index)
        Exit Function
    End If
    users(Index).channel = 0
    cchatall 0, users(Index).name & " Has Joined channel " & channels(0).name
    For a = 0 To pending.ListCount - 1
        daind = Mid(pending.List(a), 1, InStr(pending.List(a), "~") - 1)
        If daind = Index Then
            pending.RemoveItem a
            Exit For
        End If
        
        
    Next a

    Exit Function
End If

isgagged = gag(Index)
dacom = Asc(Mid(dadata, 1, 1))
dadata = Mid(dadata, 2, Len(dadata) - 2)
Select Case dacom
    Case namechange
        If InStr(dadata, " ") Then
            chatper Index, Chr(10) & "Usernames MUST NOT contain any spaces" & Chr(10)
            timeout 0.5
            Call MMChat.Winsock_onClose(Index)
            Exit Function
        End If
        For X = 0 To winsock.UBound
        If isloaded(winsock, X) And X <> Index Then
            If LCase(dadata) = LCase(users(X).name) Then
                chatper Index, Chr(10) & "you entered a chatname that already exists.  Disconnecting" & Chr(10)
                timeout 0.5
                Call MMChat.Winsock_onClose(Index)
                Exit Function
            End If
        End If
        Next X
        If MMChat.winsock(Index).Tag = "auth" Then
            If isgagged = flase Then chatall 0, users(Index).name & " is now " & dadata
        End If
        users(Index).name = dadata
    Case textall
    
        If isgagged = flase Then cchatall Index, dadata
    Case chatg
    
        If isgagged = flase Then cchatall Index, Mid(dadata, 16, Len(dadata))
    Case chatmsg
        'doevents
    Case ping
        MMChat.winsock(Index).SendData Chr(pong) & dadata & Chr(endcom)
        'doevents
    Case pong


        users(Index).pingtime = Round(Timer + 0.003 - Val(dadata), 3)
    Case creq
        If isgagged = flase Then chatper Index, "Its a server jackass.. you dont have to connect to everyone else.." & Chr(10)
        'doevents
    Case peek
        If isgagged = flase Then chatper Index, "Use command who for this information." & Chr(10)
        'doevents
    Case textper

        squ = InStr(dadata, "'")
        Do:
        squ2 = InStr(squ2 + 1, dadata, "'")
        If squ2 > 0 Then daloc = squ2
        Loop While squ2 > 0
        squ2 = daloc
        If squ = 0 Or squ2 = 0 Then
            chatper Index, Chr(10) & "WTF? read help!" & Chr(10)
            Exit Function
        End If
        dadata = Mid(dadata, squ + 1, squ2 - squ - 1)

        If isgagged = flase Then Call proccmds(dadata, Index)
        Exit Function
    Case Else
        chatper Index, Chr(10) & "Sorry.. Unimplemented, contact Dirocus" & Chr(10)
End Select
End Function
Function Connect(Index As Integer, dadata)
MMChat.winsock(Index).Tag = "been once"
If InStr(LCase(dadata), "zchat") > 0 Then
    MMChat.Winsock_onClose (Index)
    Exit Function
End If

'make global var for chatserv name
MMChat.winsock(Index).SendData "YES:" & MMChat.Text1.Text & Chr(10)


entpos = InStr(dadata, Chr(10))
If entpos = 0 Or entpos < 7 Then
    Exit Function
    
End If
username = Mid(dadata, 6, entpos - 6)
rip = Mid(dadata, entpos + 1, Len(dadata) - entpos - 5)
If InStr(username, " ") Then
    chatper Index, Chr(10) & "Usernames MUST NOT contain any spaces... Disconnecting..." & Chr(10)
    timeout 0.5
    Call MMChat.Winsock_onClose(Index)
    Exit Function
End If
    
For X = 0 To MMChat.winsock.UBound
    If LCase(username) = LCase(users(X).name) And isloaded(winsock, X) = True Then
        chatper Index, Chr(10) & "Username in use.. " & Chr(10)
        timeout 0.5
        Call MMChat.Winsock_onClose(Index)
        Exit Function
    End If
Next X
MMChat.winsock(Index).Tag = "need ver"
users(Index).name = username
users(Index).rip = rip
'doevents
End Function
Function needver(Index As Integer, dadata)
If users(Index).ver = "" Then
    users(Index).ver = Mid(dadata, 2, Len(dadata) - 2)
End If
'generic user info
users(Index).cip = MMChat.winsock(Index).RemoteHostIP
users(Index).contime = Date & " " & Time
users(Index).lastchattime = Date & " " & Time
users(Index).ocn = users(Index).name
users(Index).gaggedtime = 0
users(Index).gtimer = 0

If Dir(App.Path & "\users\" & users(Index).name & ".usr") <> "" Then
    Open App.Path & "\users\" & users(Index).name & ".usr" For Input As 1
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
    
    users(Index).pass = pass
    users(Index).perm = perm
    If cip = users(Index).cip And rip = users(Index).rip Then
        joinuser Index, "Ip Auth"
    Else
        If Check1.Value = 0 Then
            
            users(Index).sperm = perm
            users(Index).perm = 0
            joinuser Index, "No Auth"
            'chatper known user ip auth failure
        Else
            chatper Index, "You have 15 seconds to chat your password to me."
            pending.AddItem Index & "~" & Fix(Timer)
            MMChat.winsock(Index).Tag = "need pw"
        End If
    End If
Else
    If Check1.Value = 0 Then
        users(Index).pass = ""
        users(Index).perm = 0
        joinuser Index, "No Auth"
    Else
        chatper Index, "User Account NOT FOUND!!"
        timeout 0.5
        Call MMChat.Winsock_onClose(Index)
        Exit Function
    End If
End If
End Function

Function pwauth(Index As Integer, dadata)
'fix crash if string has no ' in it
squ = InStr(dadata, "'")
squ2 = InStr(squ + 1, dadata, "'")
If squ = 0 Or squ2 = 0 Then
    Exit Function
Else
    dadata = Mid(dadata, squ + 1, squ2 - squ - 1)
End If
dadata = Replace(LCase(dadata), "login ", "")
If LCase(dadata) = LCase(users(Index).pass) Then
    
    For a = 0 To pending.ListCount - 1
        daind = Mid(pending.List(a), 1, InStr(pending.List(a), "~") - 1)
        If daind = Index Then
            pending.RemoveItem a
            Exit For
        End If
        
    Next a
    joinuser Index, "Password Auth"
Else
    chatper Index, "Incorrect..."
    timeout 0.5
    winsock(Index).CloseSocket
End If
End Function

Private Sub proccmds(ByVal dadata, Index As Integer)
dadata = stripansi(dadata)
Dim params() As String
params() = Split(dadata, " ")
dacommand = LCase(params(0))
Select Case dacommand
Case "auth"
    If users(Index).perm < 3 Then
        chatper Index, "this command requires level 3 permission level" & Chr(10)
        Exit Sub
    End If
    If Check1.Value = 0 Then
        chatper Index, "Auth Enabled"
        Check1.Value = 1
    Else
        chatper Index, "Auth Disabled"
        Check1.Value = 0
    End If
    
Case "add"
    If users(Index).perm < 4 Then
        chatper Index, "this command requires level 4 permission level" & Chr(10)
        Exit Sub
    End If
    If UBound(params) <> 3 Then
        chatper Index, "Invalid parameters" & Chr(10)
        Exit Sub
    End If
    user = params(1)
    pass = params(2)
    perm = params(3)
    If perm > users(Index).perm Then
        chatper Index, "You can not create a user of higher level than yourself" & Chr(10)
        Exit Sub
    End If
    If Dir(App.Path & "\users\" & user & ".usr") <> "" Then
        chatper Index, "User already exists" & Chr(10)
        Exit Sub
    End If
    Open App.Path & "\users\" & user & ".usr" For Output As 1
    Print #1, "cip=0.0.0.0" & vbCrLf & "rip=0.0.0.0" & vbCrLf & "pass=" & pass & vbCrLf & "perm=" & perm
    Close 1
    chatper Index, "Created Sucessfully:" & Chr(10) & "User: " & user & Chr(10) & "Pass: " & pass & Chr(10) & "Perm: " & perm & Chr(10)
Case "kick"
    If users(Index).perm < 4 Then
        chatper Index, "This command requires level 4 permission level"
        Exit Sub
    End If
    If UBound(params) < 1 Then
        chatper Index, "Invalid parameters" & Chr(10)
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
        If isloaded(winsock, temp) Then
            If LCase(users(temp).name) = LCase(destusr) Then socknum = temp
        End If
    Next temp
    If users(socknum).perm > users(Index).perm Then
        chatper Index, "You cannot kick users of a higher level than yourself"
        Exit Sub
    End If
    If socknum = 0 Then
        chatper Index, Chr(10) & "Chat name not found!!!" & Chr(10)
        Exit Sub
    End If
    chatall 0, Chr(10) & "*** " & destusr & " Booted [" & dareason & "] ***" & Chr(10)
    MMChat.Winsock_onClose (socknum)
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
    chatper Index, txttosend
        
Case "del"
    'delete user
    
    If users(Index).perm < 4 Then
        chatper Index, "this command requires level 4 permission level" & Chr(10)
        Exit Sub
    End If
    If UBound(params) <> 2 Then
        chatper Index, "Invalid parameters" & Chr(10)
        Exit Sub
    End If
    If Dir(App.Path & "\users\" & params(1) & ".usr", vbNormal) <> "" Then
        Kill App.Path & "\users\" & params(1) & ".usr"
        chatper Index, "Account DELETED!"
    Else
        chatper Index, "account not found"
    End If
    
Case "login"
    If users(Index).perm > 0 Then
        chatper Index, "Your logged in already stupid" & Chr(10)
        Exit Sub
    End If
    If UBound(params) < 1 Then
        chatper Index, "missing params" & Chr(10)
        Exit Sub
    End If
    pass = params(1)
    If pass = "" Then pass = "nothing"
    If pass = users(Index).pass Then
        chatper Index, "You have been authenticated" & Chr(10)
        cchatall Index, users(Index).name & " has been authenticated" & Chr(10)
        users(Index).perm = users(Index).sperm
    Else
        chatper Index, "Afraid not..." & Chr(10)
        timeout 0.5
        Winsock_onClose (Index)
        Exit Sub
    End If
            
            
            
    
        
Case "channels"
    tts = "Current channel list:" & Chr(10)
    For pos = 0 To UBound(channels)
        If channels(pos).name <> "" Then tts = tts & channels(pos).name & Chr(10)
    Next pos
    chatper Index, tts

Case "join"
    If UBound(params) < 1 Then
        chatper Index, "Invalid parameters" & Chr(10)
        Exit Sub
    End If
    'check for auth here
    
    Call cjoin(params(1), Index)
    
Case "ccreate"

    If UBound(params) < 1 Then
        chatper Index, "Invalid parameters" & Chr(10)
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
            chatper Index, "Channel already exists.. " & Chr(10)
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
    chatall 0, "Channel [" & params(1) & "] has been created by [" & users(Index).name & "]" & Chr(10)
    cjoin params(1), Index
Case "info"
    If users(Index).perm < 1 Then
        chatper Index, "This command requires level 1 permission level"
        Exit Sub
    End If
    If UBound(params) < 1 Then
        chatper Index, "Invalid parameters" & Chr(10)
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
        chatper Index, tts
                
    Else
        chatper Index, "Could not find user..."
    End If
    Exit Sub
        
Case "chpw"
            'change your password

Case "ca"

Case "ban"
    'ban an ip or user
    'ban username X
    'ban userip X

Case "gag"

    If users(Index).perm < 4 Then
        chatper Index, "This command requires level 4 permission level"
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
        If isloaded(winsock, temp) Then
            If LCase(users(temp).name) = LCase(destusr) Then socknum = temp
        End If
    Next temp
    If socknum = 0 Then
        chatper Index, Chr(10) & "Chat name not found!!!" & Chr(10)
        Exit Sub
    End If
    If users(socknum).perm > users(Index).perm Then
        chatper Index, "You cannot gag users of a higher level than yourself"
        Exit Sub
    End If
    If datime = "" Then datime = 0.5
    
    
    chatall 0, Chr(10) & "*** " & users(socknum).name & " GAGGED [" & datime & " Min] ***" & Chr(10)
    users(socknum).gaggedtime = Fix(Timer + (datime * 60))
    
Case "help"
    tts = Chr(10) & "Current Commands (will make this look better later)" & Chr(10) & Chr(10)
    If users(Index).perm > -1 Then
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
    If users(Index).perm > 0 Then
        tts = tts & Chr(10)
        tts = tts & "Level 1 Commands:" & Chr(10)
        tts = tts & "Ccreate   Create a channel" & Chr(10)
        tts = tts & "Info      Get info on a user" & Chr(10)
        tts = tts & "gca       Global chat all, if you overuse this, your perms will be revoked" & Chr(10)
        tts = tts & "chpw      Change your password" & Chr(10)
        tts = tts & "setip     Enabel ip Authenication" & Chr(10)
        
    End If
    If users(Index).perm > 1 Then
        tts = tts & Chr(10)
        tts = tts & "Level 2 Commands:" & Chr(10)
        
        
    End If
    If users(Index).perm > 2 Then
        tts = tts & Chr(10)
        tts = tts & "Level 3 Commands:" & Chr(10)
        tts = tts & "auth      Toggle requiring logins" & Chr(10)
        
    End If
    If users(Index).perm > 3 Then
        tts = tts & Chr(10)
        tts = tts & "Level 4 Commands:" & Chr(10)
        tts = tts & "Add       Add a user into the system" & Chr(10)
        tts = tts & "Kick      Kicks a user from the server" & Chr(10)
        tts = tts & "Del       Deletes a user account" & Chr(10)
        tts = tts & "Ban       Ban a users ip address" & Chr(10)
        tts = tts & "Gag       Gag a user" & Chr(10)
        
    End If
    If users(Index).perm > 4 Then
        tts = tts & Chr(10)
        'make dubug
        tts = tts & "Level 5 Commands:" & Chr(10)
        
        
    End If
    
    chatper Index, tts
            


Case "setip"
    If users(Index).perm > 0 Then
        
        
        Kill App.Path & "\users\" & users(Index).ocn & ".usr"
        Open App.Path & "\users\" & users(Index).ocn & ".usr" For Output As 1
        Print #1, "cip=" & users(Index).cip & vbCrLf & "rip=" & users(Index).rip & vbCrLf & "pass=" & users(Index).pass & vbCrLf & "perm=" & users(Index).perm
        Close 1
        
        
        chatper Index, "IP Authenication has been enabled on this account" & Chr(10)
    Else
        chatper Index, "login first if you have an account"
    End If
Case "pmsg"
    
    secspace = InStr(6, dadata, " ")
    destusr = Mid(dadata, 6, secspace - 6)
    message = Mid(dadata, secspace + 1, Len(dadata) + secspace - 1)
    For temp = 1 To UBound(users)
    If isloaded(winsock, temp) Then
        If LCase(users(temp).name) = LCase(destusr) Then destusrsock = temp
    End If
    Next temp
    If destusrsock = 0 Then
        If Val(destusr) > 0 Then
            If isloaded(winsock, destusr) = True Then
                destusrsock = destusr
                destusr = users(destusrsock).name
            Else
                chatper Index, Chr(10) & "Chat name not found!!!" & Chr(10)
                Exit Sub
            End If
        Else
            chatper Index, Chr(10) & "Chat name not found!!!" & Chr(10)
            Exit Sub
        End If
    End If
    chatper destusrsock, Chr(10) & users(Index).name & " privately chats to you, '" & message & "'" & Chr(10)
    chatper Index, Chr(10) & "You chat to: " & destusr & Chr(10)

    ReDim Preserve users(destusrsock).log(UBound(users(destusrsock).log) + 1)
    users(destusrsock).log(UBound(users(destusrsock).log)) = Time & " " & users(Index).name & ": " & message
    ReDim Preserve users(Index).log(UBound(users(Index).log) + 1)
    users(Index).log(UBound(users(Index).log)) = Time & " To " & users(destusrsock).name & ": " & message
Case "who"
    damid = String(53, Chr(205))
    txttosend = Chr(10) & abold & awhite & Chr(201) & damid & Chr(187) & Chr(10)
    txttosend = txttosend & Chr(186) & " " & ared & "Num  Channel   Name          Orig Name         Perm " & awhite & Chr(186) & Chr(10)
    For temp = 0 To UBound(channels)
        If channels(temp).name <> "" Then
            For pos = 1 To MMChat.winsock.UBound
                If users(pos).channel = temp And isloaded(winsock, pos) = True Then
                    danum = Str(pos) & Space(5 - Len(pos))
                    dachan = channels(users(pos).channel).name & Space(10 - Len(channels(users(pos).channel).name))
                    daname = users(pos).name
                    If Len(daname) > 13 Then daname = Mid(daname, 1, 10) & "..."
                    daname = daname & Space(14 - Len(daname))
                    daoname = users(pos).ocn
                    If Len(daoname) > 17 Then daoname = Mid(daoname, 1, 14) & "..."
                    daoname = daoname & Space(18 - Len(daoname))
                    daperm = users(pos).perm & Space(6 - Len(users(pos).perm))
                    
                    txttosend = txttosend & Chr(186) & agreen & danum & dachan & daname & daoname & daperm & awhite & Chr(186) & Chr(10)
                End If
            Next pos

        End If
    Next temp
    txttosend = txttosend & Chr(200) & damid & Chr(188) & Chr(10)
    chatper Index, txttosend & Chr(10)
Case "plog"
    txtstr = ""
    start = UBound(users(Index).log) - 9
    If start < 0 Then start = 0
    If UBound(params) > 0 Then
        For temp = 1 To UBound(params) - 1
            txtstr = txtstr & params(temp) & " "
        Next temp
        If Val(params(UBound(params))) > 0 Then
            start = UBound(users(Index).log) - Val(params(UBound(params))) + 1
            If start < 1 Then start = 0
        Else
            txtstr = txtstr & params(UBound(params))
        End If
            
            
    End If
    txtstr = Trim(txtstr)
    
    If txtstr <> "" Then
        desnum = UBound(users(Index).log) - start + 1
        dacount = 0
        start = 0
        For temp = UBound(users(Index).log) To 0 Step -1
            If InStr(LCase(users(Index).log(temp)), LCase(txtstr)) > 0 Then
                dacount = dacount + 1
            End If
            If dacount = desnum Then start = temp
        Next temp
        
        
    End If
    For temp = start To UBound(users(Index).log)
        If InStr(LCase(users(Index).log(temp)), LCase(txtstr)) > 0 Then
            chatper Index, users(Index).log(temp) & Chr(10)
        End If
    Next temp
    
    
Case "log"

    txtstr = ""
    start = UBound(channels(users(Index).channel).log) - 9
    If start < 0 Then start = 0
    If UBound(params) > 0 Then
        For temp = 1 To UBound(params) - 1
            txtstr = txtstr & params(temp) & " "
        Next temp
        If Val(params(UBound(params))) > 0 Then
            start = UBound(channels(users(Index).channel).log) - Val(params(UBound(params))) + 1
            If start < 1 Then start = 0
        Else
            txtstr = txtstr & params(UBound(params))
        End If
            
            
    End If
    txtstr = Trim(txtstr)
    
    If txtstr <> "" Then
        desnum = UBound(channels(users(Index).channel).log) - start + 1
        dacount = 0
        start = 0
        For temp = UBound(channels(users(Index).channel).log) To 0 Step -1
            If InStr(LCase(channels(users(Index).channel).log(temp)), LCase(txtstr)) > 0 Then
                dacount = dacount + 1
            End If
            If dacount = desnum Then start = temp
        Next temp
        
        
    End If
    For temp = start To UBound(channels(users(Index).channel).log)
        If InStr(LCase(channels(users(Index).channel).log(temp)), LCase(txtstr)) > 0 Then
            chatper Index, channels(users(Index).channel).log(temp) & Chr(10)
        End If
    Next temp


Case Else
    chatper Index, Chr(10) & "Unknown Command or Command Syntax" & Chr(10)
End Select
End Sub

Private Function findsck(destusr)
findsck = 0
For temp = 0 To UBound(users)
    If isloaded(winsock, temp) Then
        If LCase(users(temp).name) = LCase(destusr) Then
            If MMChat.winsock(temp).Tag = "auth" Then
                findsck = temp
            End If
        End If
    End If
Next temp

End Function
Function joinuser(Index As Integer, datag)
MMChat.winsock(Index).Tag = "auth"
chatper Index, abold & agreen & Chr(10) & welcome & Chr(10)
chatall 0, users(Index).name & " Has connected to the server [" & datag & "]" & Chr(10)
Call cjoin("Main", Index)
End Function
Function cjoin(channel, Index As Integer)
channel = LCase(channel)
If users(Index).channel = -1 Then
    If Text3 <> "" Then
        chatper Index, "Please enter password..."
        
        pending.AddItem Index & "~" & Fix(Timer)
        Exit Function
    End If
    users(Index).channel = 0
    cchatall 0, users(Index).name & " Has Joined channel " & channels(0).name
    Exit Function
End If
If LCase(channels(users(Index).channel).name) = LCase(channel) Then
    chatper Index, "Umm.. Your already there..." & Chr(10)
    Exit Function
End If
For loca = 0 To UBound(channels)
    If LCase(channel) = LCase(channels(loca).name) Then
        cchatall 0, users(Index).name & " Has Left channel " & channels(users(Index).channel).name & Chr(10)
        
        
        Dim isempty As Boolean
        isempty = True
        For pos = 0 To UBound(users)
            If users(pos).channel = users(Index).channel And pos <> Index Then
                isempty = False
            End If
        Next pos
        If users(Index).channel = 0 Then isempty = False
        If isempty = True Then
            
            chatall 0, "The Channel [" & channels(users(Index).channel).name & "] has been removed" & Chr(10)
            channels(users(Index).channel).name = ""
        
        End If
        users(Index).channel = loca
        cchatall 0, users(Index).name & " Has Joined channel " & channels(loca).name
        Exit Function
    
    End If
Next loca
chatper Index, "Channel not found.." & Chr(10)

End Function
Function chatall(Index As Integer, dadata)
'
'need to strip ansi out of this string
'
If Len(dadata) > 9500 Then
    'chatper index, "Your chat has been rejected for being to large. length-" & Len(dadata) & Chr(10)
    'Exit Function
End If

If Mid(dadata, 1, 1) = Chr(10) Then dadata = Mid(dadata, 2, Len(dadata))
If Index = 0 Then dadata = "[" & Text1 & "] " & dadata
dadata = Chr(10) & dadata
If Mid(dadata, Len(dadata), 1) <> Chr(10) Then dadata = dadata & Chr(10)


Dim loopcount As Integer
For loopcount = 1 To MMChat.winsock.UBound
If isloaded(winsock, loopcount) = True Then
    If MMChat.winsock(loopcount).Tag = "auth" Then
        If Index <> loopcount Then
            chatper loopcount, dadata
            'doevents
        End If
    End If
End If
Next loopcount
chan = 0
ReDim Preserve channels(chan).log(UBound(channels(chan).log) + 1)
channels(chan).log(UBound(channels(chan).log)) = Time & " [" & channels(chan).name & "] " & dadata

End Function
Function cchatall(Index As Integer, dadata)

If Len(stripansi(dadata)) > 2000 Then
    'chatper index, "Your chat has been rejected for being to large. length-" & Len(dadata) & Chr(10)
    'Exit Function
End If

If Mid(dadata, 1, 1) = Chr(10) Then dadata = Mid(dadata, 2, Len(dadata))
If Index = 0 Then dadata = "[" & Text1 & "] " & dadata
dadata = Chr(10) & dadata
If Mid(dadata, Len(dadata), 1) <> Chr(10) Then dadata = dadata & Chr(10)


If Index <> 0 Then
    
    If InStr(LCase(stripansi(dadata)), LCase(users(Index).name)) > 3 Or InStr(LCase(stripansi(dadata)), LCase(users(Index).name)) = 0 Then
        dadata = Mid(dadata, 1, Len(dadata) - 1) & " Served by: " & users(Index).name & Chr(10)

    End If
End If

Dim loopcount As Integer
For loopcount = 1 To MMChat.winsock.UBound
If isloaded(winsock, loopcount) = True And loopcount <> Index Then
    If MMChat.winsock(loopcount).Tag = "auth" And users(loopcount).channel = users(Index).channel Then
        chatper loopcount, dadata
        'doevents
    End If
End If
Next loopcount
'doevents

dadata = Replace(dadata, " chats to everybody, ", ": ")
dadata = Replace(dadata, Chr(10), "")
chan = users(Index).channel
ReDim Preserve channels(chan).log(UBound(channels(chan).log) + 1)
channels(chan).log(UBound(channels(chan).log)) = Time & " [" & channels(chan).name & "] " & dadata
End Function



Private Function gag(Index) As Boolean
'text4 interval
'text5 point
'text6 time
gag = False
If Index = 0 Then Exit Function

'check for time change
If datime > Timer Then
    For X = 1 To UBound(users)
        users(X).gtimer = Timer
    Next X
End If
datime = Timer



If users(Index).gtimer < Timer Then users(Index).gtimer = Timer
users(Index).gtimer = users(Index).gtimer + Val(Text4)
per = (((users(Index).gtimer - Timer) * 100) / Text5)

If per > 500 Then
    Winsock_onClose (Index)
    gag = True
    Exit Function
End If
If users(Index).gaggedtime > Timer Then
    chatper Index, "You are gagged Biotch.. " & Fix(users(Index).gaggedtime - Timer) & " sec left." & Chr(10)
    gag = True
    Exit Function
End If
If per > 100 Then
    users(Index).gaggedtime = Timer + (Text6 * 60)
    cchatall 0, users(Index).name & " Auto-Gagged by System " & Text6 & " Min" & Chr(10)
    gag = True
    Exit Function
End If
If per > 70 Then
    chatper Index, "*****SPAM WARNING " & Fix(per) & "*****" & Chr(10)
End If

End Function


Private Sub Command1_Click()
Command1.Enabled = False
'close listening socket
'update all settings
'write mmchat.ini
'relisten

End Sub




Private Sub fileExit_Click()
On Error Resume Next
MMChat.Caption = "CLOSING SOCKETS..."
chatall 0, "SERVER SHUTING DOWN"
timeout 0.5
If Dir(App.Path & "\log.log") <> "" Then
    Open App.Path & "\log.log" For Input As 1
    dafilecon = Input(LOF(1), 1)
    Close 1
    Kill App.Path & "\log.log"
End If
dafilecon = dafilecon & "Program closed at USER request " & Date & " " & Time
Open App.Path & "\log.log" For Output As 1
Print #1, dafilecon
Close 1
winsock(0).CloseSocket
For a = 1 To winsock.UBound
If isloaded(winsock, a) Then
    If isloaded(winsock, a) = True Then
        
        Winsock_onClose (a)
        timeout 0.2
        Unload winsock(a)
    End If
End If
Next a
timeout 0.2
End
End Sub

Private Sub filehelp_Click()
'if help.txt isnt there then msgbox where to dl
Call Shell("notepad.exe " & App.Path & "\help.txt", vbNormalFocus)
End Sub


Private Function stripansi(ByVal bob)
stripit = InStr(bob, Chr(27) & "[")
Do While stripit > 0
findm = InStr(stripit, bob, "m")
endstrip = Mid(bob, findm + 1, Len(bob))
beginstrip = Mid(bob, 1, stripit - 1)
bob = beginstrip & endstrip
stripit = InStr(bob, Chr(27) & "[")

Loop
stripansi = bob

End Function





Private Sub Form_Load()
sysuptime = Date & " " & Time
sysversion = "Version:" & App.Major & "." & App.Minor & "." & App.Revision

init = False
If Dir(App.Path & "\banned.txt", vbNormal) = "" Then
    Open App.Path & "\banned.txt" For Output As 1
    Close 1
End If
    
Open App.Path & "\banned.txt" For Input As 1
dafilecon = Input(LOF(1), 1)
Close 1
Dim banned() As String
banned() = Split(Replace(dafilecon, vbCrLf & vbCrLf, vbCrLf), vbCrLf)
For aaa = 0 To UBound(banned)
List2.AddItem banned(aaa)
Next aaa


'if mmchat.ini doesnt exist make it prompt for all values
'write values
Open App.Path & "\mmchat.ini" For Input As 1
dafilecon = Input(LOF(1), 1)
Close 1
'load all values to default

'load ini settings
'make error safe for if a value doesnt exist
auth = Mid(dafilecon, InStr(dafilecon, "auth=") + 5, InStr(InStr(dafilecon, "auth="), dafilecon, vbCrLf) - InStr(dafilecon, "auth=") - 5)
Text1 = Mid(dafilecon, InStr(dafilecon, "chatname=") + 9, InStr(InStr(dafilecon, "chatname="), dafilecon, vbCrLf) - InStr(dafilecon, "chatname=") - 9)
authserv = Mid(dafilecon, InStr(dafilecon, "authserver=") + 11, InStr(InStr(dafilecon, "authserver="), dafilecon, vbCrLf) - InStr(dafilecon, "authserver=") - 11)
authnum = Mid(dafilecon, InStr(dafilecon, "authnum=") + 8, InStr(InStr(dafilecon, "authnum="), dafilecon, vbCrLf) - InStr(dafilecon, "authnum=") - 8)
Text2 = Mid(dafilecon, InStr(dafilecon, "port=") + 5, InStr(InStr(dafilecon, "port="), dafilecon, vbCrLf) - InStr(dafilecon, "port=") - 5)
welcome = Replace(Mid(dafilecon, InStr(dafilecon, "welcome") + 9, InStr(dafilecon, "endwelcome") - InStr(dafilecon, "welcome") - 11), vbCrLf, Chr(10))

'set objects to ini values
chatn = Text1

Text3 = ""
Check1.Value = 0
If auth = True Then Check1.Value = 1



Randomize
rand = Fix(Rnd() * 9990) + 5
rand2 = Fix(Rnd() * 9990) + 5
strtosend = Val(Mid(authnum, 1, 2)) * rand & "," & Val(Mid(authnum, 3, 2)) * rand2 & "," & Text2 & "," & sysversion

winsock(0).Tag = ""
winsock(0).LocalPort = Fix(Rnd() * 500) + 3000
winsock(0).Connect authserv, 4060


X = Timer
Do: 'doevents
timeout 0.2
Loop While X + 10 > Timer And winsock(0).State <> 7
If winsock(0).State <> 7 Then
    MsgBox "FAILED TO CONTACT LICESE SERVER, CONTACT DIROCUS"
    End
End If
winsock(0).SendData strtosend
outpacks.Caption = outpacks.Caption + 1
X = Timer
Do: 'doevents
timeout 2
Loop While X + 10 > Timer And winsock(0).Tag = ""
If winsock(0).Tag = "" Then
    MsgBox "LICENSE SERVER DID NOT RESPOND, CONTACT DIROCUS"
    End
End If
If InStr(winsock(0).Tag, "http") = 1 Then
    MsgBox "New Version released.. please click ok to download it."
    Call Shell("C:\Program Files\Internet Explorer\iexplore.EXE " & winsock(0).Tag, vbNormalFocus)
    For a = 0 To 100
    'doevents
    Next
    End
End If
If winsock(0).Tag <> Str(rand + rand2 + authnum) Then
    X = winsock(0).Tag
    MsgBox "LICENSE FAILUER, CONTACT DIROCUS"
    End
End If
winsock(0).CloseSocket
Do: DoEvents
Loop While winsock(0).State <> sckClosed
winsock(0).Tag = ""
winsock(0).LocalPort = Val(Text2)

'Winsock(0).LocalPort = 4051


MMChat.Caption = chatn


'initilize
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
'doevents
winsock(0).Listen
Command1.Enabled = False
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim Msg As Long
Dim sFilter As String
Msg = X / Screen.TwipsPerPixelX
Select Case Msg
Case WM_LBUTTONDOWN
Me.Show ' show form
Shell_NotifyIcon NIM_DELETE, nid ' del tray icon
Case WM_LBUTTONUP
Case WM_LBUTTONDBLCLK
Case WM_RBUTTONDOWN
Case WM_RBUTTONUP
Me.Show
Shell_NotifyIcon NIM_DELETE, nid
Case WM_RBUTTONDBLCLK
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call minimize_to_tray
'Shell_NotifyIcon NIM_DELETE, nid ' del tray icon
'End
Cancel = 1
End Sub



Private Sub List2_Click()
If MsgBox("Remove " & List2.List(List2.ListIndex) & " from banned list?", vbYesNo, "Remove?") = vbYes Then

    dafilecon = ""
    For temp = 0 To UBound(banned)
        If Len(banned(temp)) > 0 Then dafilecon = dafilecon & banned(temp) & vbCrLf
    Next temp
    dafilecon = Replace(dafilecon, List2.List(List2.ListIndex), "")
    Kill App.Path & "\banned.txt"
    Open App.Path & "\banned.txt" For Output As 1
    Print #1, dafilecon
    Close 1
    banned() = Split(dafilecon, vbCrLf)
    List2.Clear
    For temp = 0 To UBound(banned)
    List2.AddItem banned(temp)
    Next temp
    
End If
End Sub



Private Sub menuabout_Click()
MsgBox "Mud-Master / Zmud Chat server" & vbCrLf & sysversion & vbCrLf & "By Dirocus", vbOKOnly, "MM Chat"
End Sub

Private Sub Text1_Change()
'Command1.Enabled = True
End Sub

Private Sub Text2_Change()
'Command1.Enabled = True
End Sub

Private Sub Text3_Change()
'Command1.Enabled = True
End Sub

Private Sub Timer1_Timer()
'check pending
MsgBox outpacks.Caption
For a = 0 To pending.ListCount - 1
    datime = Mid(pending.List(a), InStr(pending.List(a), "~") + 1, Len(pending.List(a)))
    If Val(datime) < Timer - 15 Or Val(datime) > Timer Then
        Index = Mid(pending.List(a), 1, InStr(pending.List(a), "~") - 1)
        chatper Index, "Timeout.."
        timeout 1
        Winsock_onClose (Index)
        Exit For
    End If
Next a

mins = Mid(Time, InStr(Time, ":") + 4, 2)
If mins = "00" Or mins = "30" Then
    For temp = 1 To winsock.UBound
        If isloaded(winsock, temp) = True Then
            If winsock(temp).Tag = "auth" Then
                X = winsock(temp).State
                winsock(temp).SendData Chr(ping) & Timer & Chr(endcom)
                outpacks.Caption = outpacks.Caption + 1
                'doevents
            End If
        End If
    Next temp
End If
End Sub


Public Sub Winsock_onClose(Index As Integer)
If init = False Then
    winsock(0).CloseSocket
    Exit Sub
End If

a = 0
Do
If a < pending.ListCount Then
    If Mid(pending.List(a), 1, InStr(pending.List(a), "~") - 1) = Index Then
        pending.RemoveItem a
        a = a - 1
    End If
Else
    Exit Do
End If
a = a + 1
Loop

If isloaded(winsock, Index) = False Then Exit Sub

indextag = winsock(Index).Tag

winsock(Index).CloseSocket
Unload winsock(Index)

If indextag = "auth" Then
    chatall 0, users(Index).name & " has left the server."
    'remove user from user list if there
    
    
    For temp = 0 To List3.ListCount - 1
        If LCase(users(Index).name) = LCase(List3.List(temp)) Then
            List3.RemoveItem temp
            temp = temp - 1
        End If
    Next temp
    
    
Dim isempty As Boolean
    isempty = True
    For pos = 1 To winsock.UBound
        If users(pos).channel = users(Index).channel And pos <> Index Then
            isempty = False
        End If
    Next pos
    If users(Index).channel < 1 Then isempty = False
    If isempty = True Then
    
        chatall Index, "The Channel [" & channels(users(Index).channel).name & "] has been removed" & Chr(10)
        channels(users(Index).channel).name = ""
    End If
End If
End Sub


Private Sub Winsock_onConnectionRequest(Index As Integer, ByVal requestId As Long)



For a = 0 To List2.ListCount - 1
If winsock(0).RemoteHostIP = List2.List(a) Then
    winsock(0).CloseSocket
    winsock(0).Listen
    For aa = 1 To winsock.UBound
        If users(aa).perm > 3 Then
            If isloaded(winsock, aa) Then chatper aa, Chr(10) & "*** banned ip attempted to connect***"
        End If
    Next aa
    Exit Sub
End If
Next a


Dim sock As Integer
For sock = 1 To winsock.UBound + 1
If isloaded(winsock, sock) = False Then Exit For
Next sock

Load winsock(sock)
If winsock.UBound = sock Then
    ReDim Preserve users(sock)
End If
winsock(sock).Tag = "connect"

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
users(sock).gaggedtime = 0
users(sock).gtimer = Timer
'any future ones here
ReDim users(sock).log(0)
winsock(Index).CloseSocket
Do: DoEvents
Loop While winsock(Index).State <> sckClosed
winsock(Index).Listen
Do: DoEvents
Loop While winsock(Index).State <> sckListening
winsock(sock).Accept requestId
End Sub

Private Sub Winsock_onError(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock_onClose (Index)
End Sub

