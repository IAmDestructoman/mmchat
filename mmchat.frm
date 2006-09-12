VERSION 5.00
Begin VB.Form MMChat 
   Caption         =   "o"
   ClientHeight    =   4815
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox replist 
      Height          =   4740
      Left            =   5280
      TabIndex        =   25
      Top             =   0
      Width           =   3855
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Main rep filter"
      Height          =   375
      Left            =   1320
      TabIndex        =   23
      Top             =   4200
      Width           =   855
   End
   Begin VB.ListBox List4 
      Height          =   645
      Left            =   840
      TabIndex        =   21
      Top             =   1920
      Width           =   4335
   End
   Begin VB.ListBox pending 
      Height          =   450
      Left            =   0
      TabIndex        =   20
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   3240
      TabIndex        =   19
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3240
      TabIndex        =   18
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   3240
      TabIndex        =   17
      Top             =   3720
      Width           =   495
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Out"
      Height          =   255
      Left            =   600
      TabIndex        =   13
      Top             =   3960
      Width           =   615
   End
   Begin VB.CheckBox Check3 
      Caption         =   "In"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3960
      Width           =   495
   End
   Begin VB.ListBox List3 
      Height          =   1815
      Left            =   3840
      TabIndex        =   10
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   960
      TabIndex        =   9
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   645
      Left            =   2280
      TabIndex        =   5
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2640
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
      Left            =   9600
      Top             =   720
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Auth"
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Allow serving:"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3720
      Width           =   1095
   End
   Begin ChatServer.Socket Socket1 
      Left            =   2640
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      LocalPort       =   0
      RemoteHost      =   ""
      RemotePort      =   0
      Protocol        =   0
   End
   Begin ChatServer.Socket Winsock 
      Index           =   0
      Left            =   6960
      Top             =   4200
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
      Left            =   2280
      TabIndex        =   15
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Gag Time"
      Height          =   255
      Left            =   2280
      TabIndex        =   16
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Gag Interval"
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Online"
      Height          =   255
      Left            =   3840
      TabIndex        =   11
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Main PW:"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Port:"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Banned"
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Chatname:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "Plugins"
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   1920
      Width           =   855
   End
   Begin VB.Menu filemenu 
      Caption         =   "File"
      Begin VB.Menu fadduser 
         Caption         =   "Add User"
      End
      Begin VB.Menu rreps 
         Caption         =   "Reload Reps"
      End
      Begin VB.Menu fsetup 
         Caption         =   "Setup"
      End
      Begin VB.Menu fsave 
         Caption         =   "Save settings"
      End
      Begin VB.Menu fileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu fplugins 
      Caption         =   "Plugins"
      Begin VB.Menu funloadplugins 
         Caption         =   "Unload all"
      End
      Begin VB.Menu floadplugins 
         Caption         =   "Load all"
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
Public Sub LoadPlugin(strFileName)
    ReDim Preserve plugins(UBound(plugins) + 1)
    Set plugins(UBound(plugins)) = CreateObject(MakeObjectFromPath(strFileName))
     List4.AddItem plugins(UBound(plugins)).getname
    'MsgBox plugins(UBound(plugins)).getname
End Sub
Public Sub FindPlugIns(ByVal PlugInPath As String)
    Dim strFileName As String
    strFileName = Dir(PlugInPath & "\*.dll", vbNormal)
    Do While strFileName <> ""
        If PlugInExists(strFileName) Then
            LoadPlugin (strFileName)
        Else
            If PlugInExists(strFileName) Then
                LoadPlugin (strFileName)
            End If
        End If
        strFileName = Dir
    Loop
End Sub
Public Function MakeObjectFromPath(ByVal strFileName As String) As String
    MakeObjectFromPath = Replace(strFileName, ".dll", ".Plugin")
End Function
Private Function PlugInExists(strFileName As String) As Boolean
    Dim strCreateString As String
    Dim vDummy As Object
    On Error Resume Next
    strCreateString = MakeObjectFromPath(strFileName)
    Set vDummy = Nothing
    Set vDummy = CreateObject(strCreateString)

    If vDummy Is Nothing Then
        PlugInExists = False
        Call Shell("regsvr32 /s /u " & App.Path & "\plugins\" & strFileName)
        Call Shell("regsvr32 /s " & App.Path & "\plugins\" & strFileName)
    Else
        PlugInExists = True
        Set vDummy = Nothing
    End If
End Function

Function chatper(Index, dadata)
If isloaded(Winsock, Index) = False Then Exit Function
If Mid(dadata, Len(dadata), 1) <> Chr(10) Then dadata = dadata & Chr(10)
If Mid(dadata, 1, 1) <> Chr(10) Then dadata = Chr(10) & dadata
'Do: DoEvents
'Loop While users(Index).sending = True
users(Index).sending = True
MMChat.Winsock(Index).SendData Chr(textper) & dadata & Chr(endcom)



End Function
Function chatpera(Index, dadata)
If isloaded(Winsock, Index) = False Then Exit Function
If Mid(dadata, Len(dadata), 1) <> Chr(10) Then dadata = dadata & Chr(10)
If Mid(dadata, 1, 1) <> Chr(10) Then dadata = Chr(10) & dadata
'Do: DoEvents
'Loop While users(Index).sending = True
users(Index).sending = True

MMChat.Winsock(Index).SendData Chr(textall) & dadata & Chr(endcom)



End Function
Private Sub updateonline()
List3.Clear
For a = 1 To UBound(users)
    If isloaded(Winsock, a) = True Then
        If Winsock(a).Tag = "auth" Then
        List3.AddItem users(a).name
        End If
    End If
Next a
End Sub



Private Sub fadduser_Click()
Dim dauser As String
Dim dapass As String
Dim daperm As String
dauser:
dauser = InputBox("What username would you like to add?", "Add user to database")
If dauser = "" Then Exit Sub
If InStr(dauser, " ") Then
    MsgBox "Usernames cannot contain spaces"
    GoTo dauser:
    
End If
dapass:
dapass = InputBox("What password would you like to assign this user?", "Add user to database")
If dapass = "" Then Exit Sub
If InStr(dapass, " ") Then
    MsgBox "Passwords cannot contain spaces"
    GoTo dapass:
End If
daperm:
daperm = InputBox("What permission level should this user be? (0 - 4)", "Add user to database")
If daperm = "" Then Exit Sub
If Val(daperm) > 4 Or Val(daperm) < 0 Then
    MsgBox "Valid levels are 0-4"
    GoTo daperm
End If
AddToINI "Main", "pass", dapass, App.Path & "\users\" & dauser & ".usr"
AddToINI "Main", "perm", daperm, App.Path & "\users\" & dauser & ".usr"
End Sub

Private Sub floadplugins_Click()
Call funloadplugins_Click
FindPlugIns (App.Path & "\plugins")
End Sub


Private Sub funloadplugins_Click()
List4.Clear
For temp = 1 To (UBound(plugins))
Set plugins(temp) = Nothing
Next temp
ReDim plugins(0)
End Sub
Private Sub fsave_Click()
savetoini
End Sub

Private Sub fsetup_Click()
txt1 = Text1
Text1 = InputBox("Enter the Server's Chatname", "Chatname", GetFromINI("Main", "Chatname", "ChatServ", App.Path & "\mmchat.ini"))
txt2 = Text2
Text2 = InputBox("Enter the port you would like the server to listen on", "Port", GetFromINI("Main", "Port", "4050", App.Path & "\mmchat.ini"))
welcome = Replace(GetFromINI("Main", "Welcome", "'Welcome to MMChat by Dirocus/nChat to the server 'help' for usage information'", App.Path & "\mmchat.ini"), "/n", vbCrLf)
Form1.Enabled = True
Form1.Visible = True
Form1.Tag = ""
timeout 0.05
Form1.Label1.Caption = "Enter the Server welcome message"
Form1.Caption = "Welcome"
tempwel = welcome

welcome = Replace(welcome, abold, "@Bold")
welcome = Replace(welcome, areset, "@Reset")
welcome = Replace(welcome, ablack, "@Black")
welcome = Replace(welcome, ablue, "@Blue")
welcome = Replace(welcome, acyan, "@Cyan")
welcome = Replace(welcome, agreen, "@Green")
welcome = Replace(welcome, amagenta, "@Magenta")
welcome = Replace(welcome, ared, "@Red")
welcome = Replace(welcome, ayellow, "@Yellow")
welcome = Replace(welcome, awhite, "@White")

Form1.Text1 = welcome
welcome = tempwel
Do: timeout 0.1
Loop While Form1.Tag = ""
welcome = Replace(Form1.Text1, vbCrLf, Chr(10))
welcome = Replace(welcome, "@Bold", abold)
welcome = Replace(welcome, "@Reset", areset)
welcome = Replace(welcome, "@Black", ablack)
welcome = Replace(welcome, "@Blue", ablue)
welcome = Replace(welcome, "@Cyan", acyan)
welcome = Replace(welcome, "@Green", agreen)
welcome = Replace(welcome, "@Magenta", amagenta)
welcome = Replace(welcome, "@Red", ared)
welcome = Replace(welcome, "@Yellow", ayellow)
welcome = Replace(welcome, "@White", awhite)

Form1.Enabled = False
Form1.Visible = False

On Error Resume Next
If txt1 <> Text1 Then
    For a = 1 To UBound(users)
        If isloaded(Winsock, a) Then
            Winsock(a).SendData Chr(namechange) & Text1 & Chr(endcom)
            
        End If
    Next a
End If
If txt2 <> Text2 Then
    'change listening port
    Winsock(0).CloseSocket
    Winsock(0).LocalPort = Text2
    Winsock(0).Listen
End If
savetoini

End Sub
Private Sub savetoini()
AddToINI "Main", "Auth", Check1.Value, App.Path & "\mmchat.ini"
AddToINI "Main", "Chatname", Text1, App.Path & "\mmchat.ini"
AddToINI "Main", "Authserver", authserv, App.Path & "\mmchat.ini"
AddToINI "Main", "Authnum", authnum, App.Path & "\mmchat.ini"
AddToINI "Main", "Port", Text2, App.Path & "\mmchat.ini"
AddToINI "Main", "Welcome", "'" & Replace(welcome, Chr(10), "/n") & "'", App.Path & "\mmchat.ini"
AddToINI "Main", "asin", Check3.Value, App.Path & "\mmchat.ini"
AddToINI "Main", "asout", Check4.Value, App.Path & "\mmchat.ini"
AddToINI "Main", "Mainpw", Text3, App.Path & "\mmchat.ini"
AddToINI "Main", "gagi", Text4, App.Path & "\mmchat.ini"
AddToINI "Main", "gagp", Text5, App.Path & "\mmchat.ini"
AddToINI "Main", "gagt", Text6, App.Path & "\mmchat.ini"
AddToINI "Main", "spamon", Check2.Value, App.Path & "\mmchat.ini"



End Sub



Private Sub rreps_Click()
replist.Clear
If Dir(App.Path & "\replist.txt", vbNormal) = "" Then
Open App.Path & "\replist.txt" For Output As 4
Close 4
End If
Dim reps() As String
Open App.Path & "\replist.txt" For Input As 4
reps = Split(Input(LOF(4), 4), vbCrLf)
Close 4
For temp = 0 To UBound(reps)
If reps(temp) <> "" Then replist.AddItem reps(temp)
Next temp

End Sub

Private Sub Winsock_onDataArrival(Index As Integer, ByVal bytesTotal As Long)

On Error GoTo errored
If init = False Then
    Winsock(0).GetData dadata, vbString
    Winsock(0).Tag = dadata
    Exit Sub
End If
multi = False
Winsock(Index).PeekData dadata, vbString


dalen = InStr(dadata, Chr(255))

If dalen < Len(dadata) Then multi = True
If Winsock(Index).Tag <> "auth" Then
    multi = False
    dalen = Len(dadata)
End If
If dalen = 0 Then
    If Len(dadata) > 10000 Then GoTo errored
    Exit Sub
End If
Winsock(Index).GetData dadata, vbString, dalen

If isloaded(Winsock, Index) = False Then Exit Sub
If Len(dadata) < 1 Then
    GoTo errored
    'Exit Sub
End If

If pong <> Asc(Mid(dadata, 1, 1)) Then List1.AddItem Winsock(Index).RemoteHost & "-" & Index & "-" & stripansi(dadata)

List1.ListIndex = List1.ListCount - 1
Do While List1.ListCount > 19
List1.RemoveItem 0
Loop
If Mid(dadata, 1, 1) = Chr(19) And Mid(dadata, Len(dadata), 1) = Chr(255) Then
    users(Index).ver = Mid(dadata, 2, Len(dadata) - 2)
    Exit Sub
End If
If Asc(Mid(dadata, 1, 1)) = namechange Then
    GoTo errored
    Exit Sub
End If
Select Case Winsock(Index).Tag
    Case "connect"
        Call Connect(Index, dadata)

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
If MMChat.Winsock(Index).Tag <> "asfd" Then isloaded = True
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
            chatper Index, Chr(10) & "Usernames MUST NOT contain any spaces"
            timeout 0.5
            Call MMChat.Winsock_onClose(Index)
            Exit Function
        End If
        For X = 0 To Winsock.UBound
        If isloaded(Winsock, X) And X <> Index Then
            If LCase(dadata) = LCase(users(X).name) Then
                chatper Index, Chr(10) & "you entered a chatname that already exists.  Disconnecting" & Chr(10)
                timeout 0.5
                Call MMChat.Winsock_onClose(Index)
                Exit Function
            End If
        End If
        Next X
        If MMChat.Winsock(Index).Tag = "auth" Then
            If isgagged = flase Then chatall 0, users(Index).name & " is now " & dadata
        End If
        users(Index).name = dadata
    Case textall
        users(Index).lastchattime = Date & " " & Time
        If isgagged = flase Then cchatall Index, dadata
    Case chatg
         
        users(Index).lastchattime = Date & " " & Time
        If isgagged = flase Then cchatall Index, Mid(dadata, 16, Len(dadata))
    Case chatmsg
        'doevents
    Case ping
        MMChat.Winsock(Index).SendData Chr(pong) & dadata & Chr(endcom)
        
        
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
Winsock(Index).Tag = "Connecting"
If InStr(LCase(dadata), "zchat") > 0 Then
    MMChat.Winsock_onClose (Index)
    Exit Function
End If

entpos = InStr(dadata, Chr(10))
If entpos = 0 Or entpos < 7 Then
    Winsock_onClose (Index)
    Exit Function
    
End If
username = Mid(dadata, 6, entpos - 6)

For temp = 0 To List2.ListCount - 1
    If LCase(List2.List(temp)) = LCase(username) Then
        Winsock_onClose (Index)
        For aa = 1 To Winsock.UBound
            If users(aa).perm > 1 Then
                If isloaded(Winsock, aa) Then chatper aa, Chr(10) & "*** banned user [" & username & "] attempted to connect***"
            End If
        Next aa
        Exit Function
    End If
Next temp
For temp = 0 To List2.ListCount - 1
    If LCase(List2.List(temp)) = Winsock(Index).RemoteHostIP Then
        daip = Winsock(Index).RemoteHostIP
        Winsock_onClose (Index)
        For aa = 1 To Winsock.UBound
            If users(aa).perm > 1 Then
                If isloaded(Winsock, aa) Then chatper aa, Chr(10) & "*** banned ip [" & daip & "] attempted to connect***"
            End If
        Next aa
        Exit Function
    End If
Next temp

MMChat.Winsock(Index).SendData "YES:" & MMChat.Text1.Text & Chr(10)
rip = Mid(dadata, entpos + 1, Len(dadata) - entpos - 5)
If InStr(username, " ") Then
    chatper Index, Chr(10) & "Usernames MUST NOT contain any spaces... Disconnecting..." & Chr(10)
    timeout 0.5
    Call MMChat.Winsock_onClose(Index)
    Exit Function
End If
    

users(Index).name = username & "~"
users(Index).rip = rip
users(Index).cip = MMChat.Winsock(Index).RemoteHostIP
users(Index).contime = Date & " " & Time
users(Index).lastchattime = Date & " " & Time
users(Index).ocn = username
users(Index).gaggedtime = 0
users(Index).gtimer = 0

If Dir(App.Path & "\users\" & username & ".usr") <> "" Then
    
    cip = GetFromINI("main", "cip", "0.0.0.0", App.Path & "\users\" & username & ".usr")
    rip = GetFromINI("main", "rip", "0.0.0.0", App.Path & "\users\" & username & ".usr")
    pass = GetFromINI("main", "pass", "SHOULD_NEVER_SEE_THIS!", App.Path & "\users\" & username & ".usr")
    perm = GetFromINI("main", "perm", "0", App.Path & "\users\" & username & ".usr")
    
    
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
            MMChat.Winsock(Index).Tag = "need pw"
            If cip = "0.0.0.0" And rip = "0.0.0.0" Then
                users(Index).cip = "setip"
            End If
            
        End If
    End If
Else
    If Check1.Value = 0 Then
        users(Index).pass = ""
        users(Index).perm = 0
        joinuser Index, "No Auth"
    Else
        chatper Index, "User Account " & username & " NOT FOUND!!"
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
    If users(Index).cip = "setip" Then
        users(Index).cip = Winsock(Index).RemoteHostIP
        
        AddToINI "Main", "cip", users(Index).cip, App.Path & "\users\" & users(Index).ocn & ".usr"
        AddToINI "Main", "rip", users(Index).rip, App.Path & "\users\" & users(Index).ocn & ".usr"
        
    End If
Else
    chatper Index, "Incorrect..."
    timeout 0.5
    Winsock_onClose (Index)
End If
End Function

Private Sub proccmds(ByVal dadata As String, Index As Integer)
Dim daorgdata As String
daorgdata = dadata
dadata = stripansi(dadata)
Dim params() As String
params() = Split(dadata, " ")
dacommand = LCase(params(0))
For temp = 1 To UBound(commands)
    If dacommand = LCase(commands(temp).name) Then
        If users(Index).perm < commands(temp).permlvl Then
            chatper Index, "this command requires level " & commands(temp).permlvl & " permissions."
            Exit Sub
        End If
        If commands(temp).sparams = 1 Then
            If UBound(params) < commands(temp).nparams Then
                chatper Index, "Invalid params"
                Exit Sub
            End If
        End If
        If commands(temp).sparams = 2 Then
            If UBound(params) <> commands(temp).nparams Then
                chatper Index, "Invalid params"
                Exit Sub
            End If
        End If
    End If
Next temp
Select Case dacommand
Case "auth"
    If Check1.Value = 0 Then
        chatper Index, "Auth is not currently on."
        
    Else
        chatall 0, "Authentication Disabled for 2 minutes by " & users(Index).name
        authdis = 0
        Check1.Value = 0
    End If
Case "add"

    Dim user2 As String
    Dim pass2 As String
    Dim perm2 As String
    user2 = params(1)
    pass2 = params(2)
    perm2 = params(3)
    
    If perm2 > users(Index).perm Then
        chatper Index, "You can not create a user of higher level than yourself" & Chr(10)
        Exit Sub
    End If
    If Dir(App.Path & "\users\" & user2 & ".usr") <> "" Then
        chatper Index, "User already exists" & Chr(10)
        Exit Sub
    End If
 
    AddToINI "Main", "pass", pass2, App.Path & "\users\" & user2 & ".usr"
    AddToINI "Main", "perm", perm2, App.Path & "\users\" & user2 & ".usr"
    
    
    chatper Index, "Created Sucessfully:" & Chr(10) & "User: " & user2 & Chr(10) & "Pass: " & pass2 & Chr(10) & "Perm: " & perm2 & Chr(10)
Case "kick"
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
    If users(socknum).perm > users(Index).perm Then
        chatper Index, "You cannot kick users of a higher level than yourself"
        Exit Sub
    End If
    If socknum = 0 Then
        chatper Index, Chr(10) & "Chat name not found!!!" & Chr(10)
        Exit Sub
    End If
    chatall 0, Chr(10) & "*** " & destusr & " Booted [" & dareason & "] by [" & users(Index).name & "] ***" & Chr(10)
    MMChat.Winsock_onClose (socknum)
    Exit Sub
            

        
Case "del"
    'delete user
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
        If channels(pos).name <> "" Then tts = tts & channels(pos).name & "(" & pos & ")" & " " & UBound(channels(pos).log) & " log entries" & Chr(10)
    Next pos
    chatper Index, tts

Case "join"
    'check for auth here
    For a = 0 To UBound(channels)
        If LCase(channels(a).name) = LCase(params(1)) Then
            Select Case channels(a).authmode
            Case "none"
                Call cjoin(params(1), Index)
                Exit Sub
            Case "pl"
                If Val(users(Index).perm) >= Val(channels(a).permlvl) Then
                    Call cjoin(params(1), Index)
                    Exit Sub
                Else
                    chatper Index, "You do not have a high enough permission level to join this channel" & Chr(10)
                    Exit Sub
                End If
            Case "pw"
                If UBound(params) < 2 Then
                    chatper Index, "This channel is password protected." & Chr(10)
                    Exit Sub
                End If
                If LCase(params(2)) = LCase(channels(a).password) Then
                    Call cjoin(params(1), Index)
                    Exit Sub
                Else
                    chatper Index, "Incorrect password.." & Chr(10)
                    Exit Sub
                End If
            End Select
        End If
    Next a
    chatper Index, "Channel NOT found..." & Chr(10)
    Exit Sub
    'Call cjoin(params(1), index)
    
Case "ccreate"

    If Len(params(1)) > 9 Then
        chatper Index, "Too many characters" & Chr(10)
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
    authm = "none"
    authpw = ""
    authpl = 0
    If UBound(params) > 1 Then
        If Val(params(2)) > 0 Then
            authm = "pl"
            If Val(params(2)) > users(Index).perm Then
                chatper Index, "Cannot create a channel with a Permision Level greater than yours" & Chr(10)
                Exit Sub
            End If
            
            authpl = params(2)
        Else
            authm = "pw"
            authpw = params(2)
        End If
    End If
    channels(dapos).name = params(1)
    channels(dapos).authmode = authm
    channels(dapos).password = authpw
    channels(dapos).permlvl = authpl
    channels(dapos).sil = False
    ReDim channels(dapos).log(0)
    chatall 0, "Channel [" & params(1) & "] has been created by [" & users(Index).name & "]" & Chr(10)
    cjoin params(1), Index
Case "info"
    destusr = params(1)
    destsck = findsck(destusr)
    If destsck > 0 Then
        tts = "USER " & destusr & "'s info" & Chr(10)
        tts = tts & "Current channel " & channels(users(destsck).channel).name & Chr(10)
        tts = tts & "Permission lvl  " & users(destsck).perm & Chr(10)
        tts = tts & "Chat verion     " & users(destsck).ver & Chr(10)
        tts = tts & "Connected ip    " & users(destsck).cip & Chr(10)
        tts = tts & "Reported ip     " & users(destsck).rip & Chr(10)
        
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
    AddToINI "Main", "pass", params(1), App.Path & "\users\" & users(Index).ocn & ".usr"
    chatper Index, "Password Changed."
Case "ca"
    dacaloc = InStr(daorgdata, "'ca ")
    dachat = Mid(daorgdata, dacaloc + 4, Len(daorgdata))
    cchatall Index, users(Index).name & " " & dachat
Case "gca"
    If channels(users(Index).channel).sil = True Then
        If users(Index).perm < 2 Then
            chatper Index, "You can not use the Global Chat feature right now as the channel you are in is currently in silent mode."
            Exit Sub
        End If
    End If
    dacaloc = InStr(daorgdata, "'gca ")
    dachat = Mid(daorgdata, dacaloc + 5, Len(daorgdata))
    
    chatall 0, "[GLOBAL] [" & users(Index).name & "] " & dachat

Case "ban"
    List2.AddItem params(1)
    Kill App.Path & "\banned.txt"
    dafilecon = ""
    For temp = 0 To List2.ListCount - 1
        dafilecon = dafilecon & vbCrLf & List2.List(temp)
    Next temp
    Filenum = FreeFile
    Open App.Path & "\banned.txt" For Output As Filenum
    Print #Filenum, dafilecon;
    Close Filenum
    chatper Index, params(1) & " added to banned user/ip list"
    
    
Case "gag"

    
    datime = 1
    destusr = params(1)
    If UBound(params) > 1 Then datime = Val(params(2))
       
    For temp = 0 To UBound(users)
        If isloaded(Winsock, temp) Then
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
    
    For temp = 1 To 50
        If gaglist(temp, 1) = vbNull Then Exit For
    Next temp
    gaglist(temp, 1) = Now + TimeSerial(0, datime, 0)
    gaglist(temp, 2) = Winsock(socknum).RemoteHostIP
    chatall 0, Chr(10) & "*** " & users(socknum).name & " GAGGED [" & datime & " Min] by [" & users(Index).name & "] ***" & Chr(10)
    users(socknum).gaggedtime = Now + TimeSerial(0, datime, 0)

Case "fjoin"
    dauser = -1
    For a = 0 To UBound(users)
        If LCase(users(a).name) = LCase(params(1)) Then dauser = a
    Next a
    If dauser = -1 Then
        chatper Index, "User not found" & Chr(10)
        Exit Sub
    End If
    cjoin channels(users(Index).channel).name, Val(dauser)
    chatper dauser, users(Index).name & " has joined you to channel " & channels(users(dauser).channel).name & Chr(10)
Case "silent"

    If channels(users(Index).channel).sil = True Then
        channels(users(Index).channel).sil = False
        chatall 0, "Silent mode DISABLED for channel " & channels(users(Index).channel).name & " by " & users(Index).name & Chr(10)
    Else
        channels(users(Index).channel).sil = True
        chatall 0, "Silent mode ENABLED for channel " & channels(users(Index).channel).name & " by " & users(Index).name & Chr(10)
    End If
    
Case "remove"

    If LCase(params(1)) = "main" Then Exit Sub
    For temp = 1 To UBound(channels)
        If LCase(channels(temp).name) = LCase(params(1)) Then
            
            'Dim a As Integer
            For a = 1 To UBound(users)
                If users(a).channel = temp Then
                    chatper a, "The channel you were is was manualy removed, you have been automatically joined to channel main" & Chr(10)
                    users(a).channel = 0
                End If
            Next a
            chatall 0, "The channel [" & channels(temp).name & "] was manualy removed by [" & users(Index).name & "]" & Chr(10)
            channels(temp).name = ""
            Exit Sub
        End If
    Next temp
    chatper Index, "Channel not found" & Chr(10)
  

Case "help"
    If UBound(params) > 0 Then
        If Dir(App.Path & "\help\" & params(1) & ".txt", vbNormal) <> "" Then
            Filenum = FreeFile
            Open App.Path & "\help\" & params(1) & ".txt" For Input As Filenum
            dafilecon = Input(LOF(Filenum), Filenum)
            Close Filenum
            chatper Index, Replace(dafilecon, vbCrLf, Chr(10))
        Else
            chatper Index, "No help file for that command."
        End If
        Exit Sub
    End If
    tts = Chr(10) & "MMChat by Dirocus.  " & sysversion & Chr(10) & "Current Commands: Use help sub-command for more info" & Chr(10)
    For temp = 0 To users(Index).perm
        tts = tts & Chr(10) & "Level " & temp & " commands:" & Chr(10)
        For temp2 = 1 To UBound(commands)
            If commands(temp2).permlvl = temp Then
                spaces = Space(10 - Len(commands(temp2).name))
                tts = tts & commands(temp2).name & spaces & commands(temp2).help & Chr(10)
            End If
        Next temp2
    Next temp
    chatper Index, tts
            


Case "setip"
    If users(Index).perm > 0 Then
        
       
        AddToINI "Main", "cip", users(Index).cip, App.Path & "\users\" & users(Index).ocn & ".usr"
        AddToINI "Main", "rip", users(Index).rip, App.Path & "\users\" & users(Index).ocn & ".usr"
        
        chatper Index, "IP Authenication has been enabled on this account" & Chr(10)
    Else
        chatper Index, "login first if you have an account"
    End If
Case "pmsg"
    
    secspace = InStr(6, dadata, " ")
    destusr = Mid(dadata, 6, secspace - 6)
    '
    '   UNTESTED CODE
    '
    message = Mid(daorgdata, InStr(daorgdata, destuser & " ") + Len(destusr) + 2, Len(daorgdata))
    msssage = Mid(message, 1, Len(message) - 1)
    '
    '
    '
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
            For pos = 1 To MMChat.Winsock.UBound
                If users(pos).channel = temp And isloaded(Winsock, pos) = True Then
                    danum = Str(pos) & Space(5 - Len(pos))
                    dachan = channels(users(pos).channel).name & Space(10 - Len(channels(users(pos).channel).name))
                    daname = users(pos).name
                    If Len(daname) > 13 Then daname = Mid(daname, 1, 10) & "..."
                    daname = daname & Space(14 - Len(daname))
                    daoname = users(pos).ocn
                    If Len(daoname) > 17 Then daoname = Mid(daoname, 1, 14) & "..."
                    daoname = daoname & Space(18 - Len(daoname))
                    daperm = users(pos).perm & Space(6 - Len(users(pos).perm))
                    If LCase(users(pos).ocn) <> LCase(users(pos).name) Then daname = ayellow & daname
                    txttosend = txttosend & Chr(186) & agreen & danum & dachan & daname & agreen & daoname & daperm & awhite & Chr(186) & Chr(10)
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
            If dacount = desnum Then
                start = temp
                temp = 0
            End If
        Next temp
        
        
    End If
    For temp = start To UBound(users(Index).log)
        If InStr(LCase(users(Index).log(temp)), LCase(txtstr)) > 0 Then
            chatper Index, users(Index).log(temp) & Chr(10)
            timeout 0.05
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
            If dacount = desnum Then
                start = temp
                temp = 0
            End If
        Next temp
        
        
    End If
    For temp = start To UBound(channels(users(Index).channel).log)
        If InStr(LCase(channels(users(Index).channel).log(temp)), LCase(txtstr)) > 0 Then
            chatper Index, channels(users(Index).channel).log(temp) & Chr(10)
            timeout 0.01
        End If
    Next temp

    
Case "plugins"
    tts = "Current Enabled plugins and commands:" & Chr(10) & Chr(10)
    For temp = 1 To UBound(plugins)
        tts = tts & plugins(temp).getname & Chr(10)
        tts = tts & "Command        Perm Lvl" & Chr(10)
        Dim commands2() As String
        commands2 = Split(plugins(temp).getcommands, ",")
        For temp2 = 1 To UBound(commands2) - 1 Step 2
        commands2(temp2) = commands2(temp2) & Space(15 - Len(commands2(temp2)))
        tts = tts & commands2(temp2) & commands2(temp2 + 1) & Chr(10)
        Next temp2
        tts = tts & Chr(10)
        
    Next temp
    chatper Index, tts
Case Else
    'check plugins for command
    For temp = 1 To UBound(plugins)
        plugincommands = LCase(plugins(temp).getcommands)
        If InStr(plugincommands, "," & LCase(params(0)) & ",") Then
            comloc = InStr(plugincommands, "," & LCase(params(0)) & ",")
            reqperm = Mid(plugincommands, comloc + Len(params(0)) + 2, 1)
            If reqperm < users(Index).perm + 1 Then
                If procplugin = True Then
                    chatper Index, "Unavailable, try again in a few seconds." & Chr(10)
                    Exit Sub
                End If
                procplugin = True
                
                'CREATE OTHERVARS VAR WITH ANY OTHER VARS A PLUGIN MAY NEED
                'VAR=DATA FORMAT
                Dim othervars As String
                othervars = ""
                othervars = othervars & "path=" & App.Path & ","
                othervars = othervars & "servername=" & Text1 & ","
                othervars = othervars & "username=" & users(Index).name & ","
                othervars = othervars & "ocn=" & users(Index).ocn & ","
                othervars = othervars & "perm=" & users(Index).perm & ","
                
                retresult = plugins(temp).main(daorgdata, othervars)
'there can be multipule results seperated by ~
'the first char of each result tells if you should:
'0 chatper
'1 chatall not per
'2 cchatall not per
'3 chatall all
'4 cchatall all
'5 chatper to another user 5username,message
'6 chatper to another user 5ocn,message
                
                If InStr(retresult, "~") = 0 Then
                    'only 1 result
                    datxt = Mid(retresult, 2, Len(retresult)) & Chr(10)
                    Select Case Val(Mid(retresult, 1, 1))
                    Case 0
                        chatper Index, datxt
                    Case 1
                        chatall Index, datxt
                    Case 2
                        cchatall Index, datxt
                    Case 3
                        chatall 0, datxt
                    Case 4
                        cchatall 0, datxt
                    Case 5
                        commaloc = InStr(datxt, ",")
                        dest = Left(datxt, commaloc - 1)
                        destm = Right(datxt, Len(datxt) - commaloc)
                        sentit = 0
                        For xx = 1 To UBound(users)
                            If LCase(users(xx).name) = LCase(dest) Then
                                  chatper xx, destm
                                  sendit = 1
                            End If
                        Next xx
                        If sendit = 0 Then chatper Index, "A Message intendend for " & dest & "was not delivered as the user was not found."
                    Case 6
                        commaloc = InStr(datxt, ",")
                        dest = Left(datxt, commaloc - 1)
                        destm = Right(datxt, Len(datxt) - commaloc)
                        sendit = 0
                        For xx = 1 To UBound(users)
                            If LCase(users(xx).ocn) = LCase(dest) Then
                                  chatper xx, destm
                                  sendit = 1
                            End If
                        Next xx
                        If sendit = 0 Then chatper Index, "A Message intendend for " & dest & "was not delivered as the user was not found."
                    End Select
                Else
                    'multiresult
                    
                    Dim retr() As String
                    retr = Split(retresult, "~")
                    For temp2 = 0 To UBound(retr)
                    datxt = Mid(retr(temp2), 2, Len(retr(temp2))) & Chr(10)
                    Select Case Val(Mid(retr(temp2), 1, 1))
                    Case 0
                        chatper Index, datxt
                    Case 1
                        chatall Index, datxt
                    Case 2
                        cchatall Index, datxt
                    Case 3
                        chatall 0, datxt
                    Case 4
                        cchatall 0, datxt
                    Case 5
                        commaloc = InStr(datxt, ",")
                        dest = Left(datxt, commaloc - 1)
                        destm = Right(datxt, Len(datxt) - commaloc)
                        sendit = 0
                        For xx = 1 To UBound(users)
                            If LCase(users(xx).name) = LCase(dest) Then
                                  chatper xx, destm
                                  sendit = 1
                            End If
                        Next xx
                        If sendit = 0 Then chatper Index, "A Message intendend for " & dest & "was not delivered as the user was not found."
                    Case 6
                        commaloc = InStr(datxt, ",")
                        dest = Left(datxt, commaloc - 1)
                        destm = Right(datxt, Len(datxt) - commaloc)
                        sendit = 0
                        For xx = 1 To UBound(users)
                            If LCase(users(xx).ocn) = LCase(dest) Then
                                  chatper xx, destm
                                  sendit = 1
                            End If
                        Next xx
                        If sendit = 0 Then chatper Index, "A Message intendend for " & dest & "was not delivered as the user was not found."
                    End Select
                    Next temp2
                End If
                
                procplugin = False
            Else
                chatper Index, "This command requires level " & reqperm & " permission level"
            End If
            Exit Sub
        End If
    Next temp
    
    
    chatper Index, Chr(10) & "Unknown Command or Command Syntax" & Chr(10)
End Select
End Sub

Private Function findsck(destusr)
findsck = 0
For temp = 0 To UBound(users)
    If isloaded(Winsock, temp) Then
        If LCase(users(temp).name) = LCase(destusr) Then
            If MMChat.Winsock(temp).Tag = "auth" Then
                findsck = temp
            End If
        End If
    End If
Next temp

End Function
Function joinuser(Index As Integer, datag)
Dim X As Integer
For X = 0 To MMChat.Winsock.UBound
    If LCase(Mid(users(Index).name, 1, Len(users(Index).name) - 1)) = LCase(users(X).name) And isloaded(Winsock, X) = True Then
        
        If X <> Index Then
            If LCase(datag) = "no auth" Then
                If users(Index).pass <> "" Then
                    chatper Index, "You have 15 seconds to chat your password to me."
                    pending.AddItem Index & "~" & Fix(Timer)
                    MMChat.Winsock(Index).Tag = "need pw"
                    Exit Function
                Else
                    chatper Index, "Username in use.. " & Chr(10)
                    timeout 0.1
                    Winsock_onClose (Index)
                    Exit Function
                End If
            End If
            
            
            chatper X, "You have logged on from another location.. Disconnecting.." & Chr(10)
            timeout 0.1
            Winsock_onClose X
        End If
    End If
Next X

users(Index).name = Mid(users(Index).name, 1, Len(users(Index).name) - 1)


MMChat.Winsock(Index).Tag = "auth"

Winsock(Index).SendData Chr(19) & "MMChat by Dirocus.  " & sysversion & Chr(255)
chatper Index, Chr(10) & welcome & Chr(10)

chatall 0, users(Index).name & " Has connected to the server [" & datag & "]" & Chr(10)
updateonline
Call cjoin("Main", Index)
If authnum = "0000" Then chatall 0, "UNREGISTERED VERSION, CONTACT DIROCUS FOR FULL VERSION INFORMATION"

For temp = 1 To 50
    If Now > gaglist(temp, 1) Then
        gaglist(temp, 1) = vbNull
        gaglist(temp, 2) = vbNull
    End If
    If gaglist(temp, 2) = Winsock(Index).RemoteHostIP Then
        users(Index).gaggedtime = gaglist(temp, 1)
    
    End If
Next temp
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
    cchatall Index, users(Index).name & " Has Joined channel " & channels(0).name
    Exit Function
End If
If LCase(channels(users(Index).channel).name) = LCase(channel) Then
    chatper Index, "Umm.. Your already there..." & Chr(10)
    Exit Function
End If
For loca = 0 To UBound(channels)
    If LCase(channel) = LCase(channels(loca).name) Then
        cchatall Index, users(Index).name & " Has Left channel " & channels(users(Index).channel).name & Chr(10)
        
        
        Dim isempty As Boolean
        isempty = True
        For pos = 0 To UBound(users)
            If users(pos).channel = users(Index).channel And pos <> Index Then
                If isloaded(Winsock, pos) = True Then isempty = False
            End If
        Next pos
        If users(Index).channel = 0 Then isempty = False
        If isempty = True Then
            
            chatall 0, "The Channel [" & channels(users(Index).channel).name & "] has been removed" & Chr(10)
            channels(users(Index).channel).name = ""
        
        End If
        users(Index).channel = loca
        cchatall Index, users(Index).name & " Has Joined channel " & channels(loca).name
        chatper Index, "You have joined channel " & channels(loca).name
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

If Mid(dadata, Len(dadata), 1) <> Chr(10) Then dadata = dadata & Chr(10)


Dim loopcount As Integer
For loopcount = 1 To MMChat.Winsock.UBound
If isloaded(Winsock, loopcount) = True Then
    If MMChat.Winsock(loopcount).Tag = "auth" Then
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
If Mid(dadata, Len(dadata), 1) <> Chr(10) Then dadata = dadata & Chr(10)


If Index <> 0 Then
    
    If users(Index).channel = 0 And Check2.Value = 1 Then
        'check database for REP SPAM
        Dim daspam() As String
        For temp = 0 To replist.ListCount - 1
            
            If InStr(replist.List(temp), ",") > 0 Then
                daspam = Split(replist.List(temp), ",")
            Else
                ReDim daspam(0)
                daspam(0) = replist.List(temp)
            End If
                
                isspam = True
                For temp2 = 0 To UBound(daspam)
                    If InStr(stripansi(dadata), daspam(temp2)) = 0 Then isspam = False
                Next temp2
                If isspam = True Then
                    chatper Index, "Your chat has been denied because it matches the spam database."
                    Exit Function
                End If
        Next temp
    End If
    
    If channels(users(Index).channel).sil = True Then
        If users(Index).perm < 2 Then
            chatper Index, "This channel is in SILENT mode only level 2+ can speak" & Chr(10)
            Exit Function
        End If
    End If
    
    If InStr(LCase(stripansi(dadata)), LCase(users(Index).name)) > 3 Or InStr(LCase(stripansi(dadata)), LCase(users(Index).name)) = 0 Then
        If Check3.Value = 0 Then Exit Function
        norname = 1
        For a = 2 To 5
            If Asc(Mid(dadata, a, 1)) < 97 Or Asc(Mid(dadata, a, 1)) > 122 Then
                If Asc(Mid(dadata, a, 1)) < 65 Or Asc(Mid(dadata, a, 1)) > 90 Then
                    norname = 0
                End If
            End If
        Next a
        If norname = 0 Then
            dadata = Mid(dadata, 8, Len(dadata))
        End If
        dadata = Mid(dadata, 1, Len(dadata) - 1) & " Served by: " & users(Index).name & Chr(10)
        

    End If
End If

Dim loopcount As Integer
For loopcount = 1 To MMChat.Winsock.UBound
If isloaded(Winsock, loopcount) = True And loopcount <> Index Then
    If MMChat.Winsock(loopcount).Tag = "auth" And users(loopcount).channel = users(Index).channel Then
        If Check4.Value = 1 Then
            chatpera loopcount, dadata
        Else
            chatper loopcount, dadata
        End If
    End If
End If
Next loopcount
'doevents

dadata = Replace(dadata, " chats to everybody, ", ": ")
dadata = Replace(dadata, Chr(10), "")
chan = users(Index).channel
If UBound(channels(chan).log) > 1500 Then
    Dim templog(1000) As String
    temp1 = 1000
    For temp = UBound(channels(chan).log) To UBound(channels(chan).log) - 1000 Step -1
        templog(temp1) = channels(chan).log(temp)
        temp1 = temp1 - 1
    Next temp
    ReDim channels(chan).log(1000)
    For temp = 0 To 1000
        channels(chan).log(temp) = templog(temp)
    Next temp
End If

ReDim Preserve channels(chan).log(UBound(channels(chan).log) + 1)
channels(chan).log(UBound(channels(chan).log)) = Time & " [" & channels(chan).name & "] " & dadata
End Function



Private Function gag(Index) As Boolean

'text4 interval
'text5 point
'text6 time

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
If users(Index).gaggedtime > Now Then
    chatper Index, "You are Gagged until " & users(Index).gaggedtime & " EST."
    gag = True
    Exit Function
End If
If per > 100 Then
    For temp = 1 To 50
    If gaglist(temp, 1) = vbNull Then Exit For
    Next temp
    gaglist(temp, 1) = Now + TimeSerial(0, Text6, 0)
    gaglist(temp, 2) = Winsock(Index).RemoteHostIP
    users(Index).gaggedtime = Now + TimeSerial(0, Text6, 0)
    cchatall 0, users(Index).name & " Auto-Gagged by System " & Text6 & " Min" & Chr(10)
    gag = True
    Exit Function
End If
If per > 70 Then
    chatper Index, "*****SPAM WARNING " & Fix(per) & "*****" & Chr(10)
End If

End Function







Private Sub fileExit_Click()
savetoini
On Error Resume Next
MMChat.Caption = "CLOSING SOCKETS..."
chatall 0, "SERVER SHUTING DOWN"
timeout 0.5
If Dir(App.Path & "\log.log") <> "" Then
    Filenum = FreeFile
    Open App.Path & "\log.log" For Input As Filenum
    dafilecon = Input(LOF(Filenum), Filenum)
    Close Filenum
    Kill App.Path & "\log.log"
End If
dafilecon = dafilecon & "Program closed at USER request " & Date & " " & Time
Filenum = FreeFile
Open App.Path & "\log.log" For Output As Filenum
Print #Filenum, dafilecon
Close Filenum
Winsock(0).CloseSocket
For a = 1 To Winsock.UBound
If isloaded(Winsock, a) Then
    If isloaded(Winsock, a) = True Then
        
        Winsock_onClose (a)
        timeout 0.2
        Unload Winsock(a)
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
MMChat.Enabled = False
MMChat.Show
MMChat.Caption = "Loading"
If App.EXEName = "mmchat.new" Then
    timeout 10
    Kill App.Path & "\mmchat.exe"
    Name App.Path & "\mmchat.prg" As App.Path & "\mmchat.exe"
    Shell App.Path & "\mmchat.exe", vbNormalFocus
    End
End If
timeout 1
If Dir(App.Path & "\mmchat.new.exe", vbNormal) <> "" Then
    Kill App.Path & "\mmchat.new.exe"
End If


MMChat.Enabled = False
MMChat.Show
MMChat.Caption = "Loading"
timeout (Rnd() * 2) + 1
sysuptime = Date & " " & Time
sysversion = "Version:" & App.Major & "." & App.Minor & "." & App.Revision

init = False
If Dir(App.Path & "\banned.txt", vbNormal) = "" Then
    Open App.Path & "\banned.txt" For Output As 1
    Close 1
End If
If Dir(App.Path & "\plugins", vbDirectory) = "" Then
    MkDir (App.Path & "\plugins")
End If
If Dir(App.Path & "\users", vbDirectory) = "" Then
    MkDir (App.Path & "\users")
End If
    
Open App.Path & "\banned.txt" For Input As 1
dafilecon = Input(LOF(1), 1)
Close 1
Do
dafilecon = Replace(dafilecon, vbCrLf & vbCrLf, vbCrLf)
Loop While InStr(dafilecon, vbCrLf & vbCrLf) > 0
Dim banned() As String
banned() = Split(dafilecon, vbCrLf)
For aaa = 0 To UBound(banned)
List2.AddItem banned(aaa)
Next aaa

loadsettings


Randomize
rand = Fix(Rnd() * 9990) + 5
rand2 = Fix(Rnd() * 9990) + 5
strtosend = Val(Mid(authnum, 1, 2)) * rand & "," & Val(Mid(authnum, 3, 2)) * rand2 & "," & Text2 & "," & sysversion

'to force demo mode
'authnum = "0000"


If authnum = "0000" Then strtosend = "DEMO," & sysversion
Winsock(0).Tag = ""
Winsock(0).LocalPort = Fix(Rnd() * 500) + 3000


'bypass auth
'GoTo bypassit


Winsock(0).Connect authserv, 4060


X = Timer
Do: 'doevents
timeout 0.2
Loop While X + 10 > Timer And Winsock(0).State <> 7
If Winsock(0).State <> 7 Then
    MsgBox "FAILED TO CONTACT LICESE SERVER, CONTACT DIROCUS"
    End
End If
Winsock(0).SendData strtosend


X = Timer
Do: 'doevents
timeout 2
Loop While X + 10 > Timer And Winsock(0).Tag = ""
If Winsock(0).Tag = "" Then
    MsgBox "LICENSE SERVER DID NOT RESPOND, CONTACT DIROCUS"
    End
End If
If InStr(Winsock(0).Tag, "http") = 1 Then
    
    Call Load(Form2)
    MMChat.Caption = "Updating... please wait..."
    Exit Sub
End If
If authnum = "0000" Then
    If Winsock(0).Tag <> "OK" Then
        MsgBox "Demo use rejected"
        End
    End If
Else
    If Winsock(0).Tag <> Str(rand + rand2 + authnum) Then
        X = Winsock(0).Tag
        MsgBox "LICENSE FAILUER, CONTACT DIROCUS"
        End
    End If
End If
Winsock(0).CloseSocket
Do: DoEvents
Loop While Winsock(0).State <> sckClosed
bypassit:
'

ReDim plugins(0)
MMChat.Caption = "Searching for Plugins"
FindPlugIns (App.Path & "\plugins")
rreps_Click

'
Winsock(0).Tag = ""
Winsock(0).LocalPort = Val(Text2)

MMChat.Caption = "MMChat by Dirocus - " & sysversion
MMChat.Enabled = True

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
channels(0).sil = False

'doevents

Winsock(0).Listen
procplugin = False
End Sub

Private Sub loadsettings()
Check1.Value = GetFromINI("Main", "Auth", "0", App.Path & "\mmchat.ini")
Text1 = GetFromINI("Main", "Chatname", "ChatServ", App.Path & "\mmchat.ini")
authserv = GetFromINI("Main", "Authserver", "ns1.mmchat.example.org", App.Path & "\mmchat.ini")
authnum = GetFromINI("Main", "Authnum", "0000", App.Path & "\mmchat.ini")
Text2 = GetFromINI("Main", "Port", "4050", App.Path & "\mmchat.ini")
welcome = Replace(GetFromINI("Main", "Welcome", "'Welcome to MMChat by Dirocus/nChat to the server 'help' for usage information'", App.Path & "\mmchat.ini"), "/n", Chr(10))
Check3.Value = GetFromINI("Main", "asin", "1", App.Path & "\mmchat.ini")
Check4.Value = GetFromINI("Main", "asout", "1", App.Path & "\mmchat.ini")
Text3 = GetFromINI("Main", "Mainpw", "", App.Path & "\mmchat.ini")
Text4 = GetFromINI("Main", "gagi", "1", App.Path & "\mmchat.ini")
Text5 = GetFromINI("Main", "gagp", "8", App.Path & "\mmchat.ini")
Text6 = GetFromINI("Main", "gagt", ".5", App.Path & "\mmchat.ini")
Check2.Value = GetFromINI("Main", "spamon", "1", App.Path & "\mmchat.ini")

commands(1).name = "auth"
commands(1).nparams = 0
commands(1).sparams = 2
commands(1).permlvl = GetFromINI("Perms", "auth", 3, App.Path & "\mmchat.ini")
commands(1).help = "Disable requiring auth for 2 minutes"


commands(2).name = "add"
commands(2).nparams = 3
commands(2).sparams = 2
commands(2).permlvl = GetFromINI("Perms", "add", 4, App.Path & "\mmchat.ini")
commands(2).help = "Adds a user account to the system"


commands(3).name = "kick"
commands(3).nparams = 1
commands(3).sparams = 1
commands(3).permlvl = GetFromINI("Perms", "kick", 4, App.Path & "\mmchat.ini")
commands(3).help = "Kicks a user from the server"


commands(4).name = "del"
commands(4).nparams = 1
commands(4).sparams = 2
commands(4).permlvl = GetFromINI("Perms", "del", 4, App.Path & "\mmchat.ini")
commands(4).help = "Delete a user account"


commands(5).name = "login"
commands(5).nparams = 1
commands(5).sparams = 2
commands(5).permlvl = GetFromINI("Perms", "login", 0, App.Path & "\mmchat.ini")
commands(5).help = "Used to authenicate on an anonymous server"

commands(6).name = "channels"
commands(6).nparams = 0
commands(6).sparams = 0
commands(6).permlvl = GetFromINI("Perms", "channels", 0, App.Path & "\mmchat.ini")
commands(6).help = "Shows current loaded channels"

commands(7).name = "join"
commands(7).nparams = 1
commands(7).sparams = 1
commands(7).permlvl = GetFromINI("Perms", "join", 0, App.Path & "\mmchat.ini")
commands(7).help = "Join a channel"

commands(8).name = "ccreate"
commands(8).nparams = 1
commands(8).sparams = 1
commands(8).permlvl = GetFromINI("Perms", "ccreate", 1, App.Path & "\mmchat.ini")
commands(8).help = "Create a channel"

commands(9).name = "info"
commands(9).nparams = 1
commands(9).sparams = 2
commands(9).permlvl = GetFromINI("Perms", "info", 1, App.Path & "\mmchat.ini")
commands(9).help = "Get detailed info on a connected user"

commands(10).name = "chpw"
commands(10).nparams = 1
commands(10).sparams = 2
commands(10).permlvl = GetFromINI("Perms", "chpw", 1, App.Path & "\mmchat.ini")
commands(10).help = "Change your password"

commands(11).name = "ca"
commands(11).nparams = 1
commands(11).sparams = 1
commands(11).permlvl = GetFromINI("Perms", "ca", 0, App.Path & "\mmchat.ini")
commands(11).help = "Chat all just to server"

commands(12).name = "gca"
commands(12).nparams = 1
commands(12).sparams = 1
commands(12).permlvl = GetFromINI("Perms", "gca", 1, App.Path & "\mmchat.ini")
commands(12).help = "Global chat all.  Chat across all channels"

commands(13).name = "ban"
commands(13).nparams = 1
commands(13).sparams = 2
commands(13).permlvl = GetFromINI("Perms", "ban", 4, App.Path & "\mmchat.ini")
commands(13).help = "Ban a username or IP address"

commands(14).name = "gag"
commands(14).nparams = 1
commands(14).sparams = 1
commands(14).permlvl = GetFromINI("Perms", "gag", 4, App.Path & "\mmchat.ini")
commands(14).help = "Gag a user for specified time"

commands(15).name = "fjoin"
commands(15).nparams = 1
commands(15).sparams = 2
commands(15).permlvl = GetFromINI("Perms", "fjoin", 4, App.Path & "\mmchat.ini")
commands(15).help = "Force a user to join your current channel"

commands(16).name = "silent"
commands(16).nparams = 0
commands(16).sparams = 2
commands(16).permlvl = GetFromINI("Perms", "silent", 3, App.Path & "\mmchat.ini")
commands(16).help = "Toggle silent mode for your current channel"

commands(17).name = "remove"
commands(17).nparams = 1
commands(17).sparams = 2
commands(17).permlvl = GetFromINI("Perms", "remove", 3, App.Path & "\mmchat.ini")
commands(17).help = "Force removal of a channel"

commands(18).name = "help"
commands(18).nparams = 0
commands(18).sparams = 0
commands(18).permlvl = GetFromINI("Perms", "help", 0, App.Path & "\mmchat.ini")
commands(18).help = "Use help sub-command for more info"

commands(19).name = "setip"
commands(19).nparams = 0
commands(19).sparams = 0
commands(19).permlvl = GetFromINI("Perms", "setip", 1, App.Path & "\mmchat.ini")
commands(19).help = "Enable/reset ipauth feature"

commands(20).name = "pmsg"
commands(20).nparams = 1
commands(20).sparams = 1
commands(20).permlvl = GetFromINI("Perms", "pmsg", 0, App.Path & "\mmchat.ini")
commands(20).help = "Private chat to another user"

commands(21).name = "who"
commands(21).nparams = 0
commands(21).sparams = 1
commands(21).permlvl = GetFromINI("Perms", "who", 0, App.Path & "\mmchat.ini")
commands(21).help = "Lists who is online and thier channel and perm lvl"

commands(22).name = "plog"
commands(22).nparams = 0
commands(22).sparams = 1
commands(22).permlvl = GetFromINI("Perms", "plog", 0, App.Path & "\mmchat.ini")
commands(22).help = "Shows your pmsg log"

commands(23).name = "log"
commands(23).nparams = 0
commands(23).sparams = 1
commands(23).permlvl = GetFromINI("Perms", "log", 0, App.Path & "\mmchat.ini")
commands(23).help = "Shows you currnet channels log"

commands(24).name = "plugins"
commands(24).nparams = 0
commands(24).sparams = 2
commands(24).permlvl = GetFromINI("Perms", "plugins", 0, App.Path & "\mmchat.ini")
commands(24).help = "List loaded plugins and their commands"


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

    List2.RemoveItem (List2.ListIndex)
    Kill App.Path & "\banned.txt"
    dafilecon = ""
    For temp = 0 To List2.ListCount - 1
        dafilecon = dafilecon & vbCrLf & List2.List(temp)
    Next temp
    Filenum = FreeFile
    Open App.Path & "\banned.txt" For Output As Filenum
    Print #Filenum, dafilecon;
    Close Filenum
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



Private Sub Timer1_Timer()

For a = 0 To pending.ListCount - 1
    datime = Mid(pending.List(a), InStr(pending.List(a), "~") + 1, Len(pending.List(a)))
    If Val(datime) < Timer - 15 Or Val(datime) > Timer Then
        Index = Mid(pending.List(a), 1, InStr(pending.List(a), "~") - 1)
        chatper Index, "Timeout.."
        timeout 0.1
        Winsock_onClose (Index)
        Exit For
    End If
Next a
mins = Mid(Time, InStr(Time, ":") + 1, 5)
If mins = "00:00" Or mins = "30:00" Then
    If authnum = "0000" Then
        chatall 0, Chr(10) & "UNREGISTERED VERSION, CONTACT DIROCUS FOR FULL VERSION INFORMATION"
        timeout 1
        chatall 0, Chr(10) & "UNREGISTERED VERSION, CONTACT DIROCUS FOR FULL VERSION INFORMATION"
        timeout 1
        chatall 0, Chr(10) & "UNREGISTERED VERSION, CONTACT DIROCUS FOR FULL VERSION INFORMATION"
        If Fix(Rnd() * 5) = 3 Then
            For temp = 1 To Winsock.UBound
                
                Winsock_onClose (temp)
            Next temp
            Exit Sub
        End If
    End If
End If
mins = Mid(Time, InStr(Time, ":") + 4, 2)
If mins = "00" Or mins = "30" Then

    If authdis <> "" Then
        authdis = authdis + 1
        If authdis > 3 Then
            authdis = ""
            Check1.Value = 1
            chatall 0, "Authenication enabled by System" & Chr(10)
        End If
    End If



    For temp = 1 To Winsock.UBound
        If isloaded(Winsock, temp) = True Then
            If Winsock(temp).Tag = "auth" Then
                X = Winsock(temp).State
                Winsock(temp).SendData Chr(ping) & Timer & Chr(endcom)
                
                
                'doevents
            End If
        End If
    Next temp
End If
End Sub


Public Sub Winsock_onClose(Index As Integer)
procplugin = False
If init = False Then
    Winsock(0).CloseSocket
    Exit Sub
End If

If Index = 0 Then
    MsgBox "Listing socket failed.  Port already in use."
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

If isloaded(Winsock, Index) = False Then Exit Sub

indextag = Winsock(Index).Tag

Winsock(Index).CloseSocket
Unload Winsock(Index)
updateonline
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
    For pos = 1 To Winsock.UBound
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
clearusers (sock)
End Sub
Private Sub Winsock_onConnectionRequest(Index As Integer, ByVal requestId As Long)
Dim sock As Integer
For sock = 1 To Winsock.UBound + 1
If isloaded(Winsock, sock) = False Then Exit For
Next sock
Load Winsock(sock)
If Winsock.UBound = sock Then
    ReDim Preserve users(sock)
End If
Winsock(sock).Tag = "connect"
clearusers (sock)
timeout 0.1
Winsock(sock).Accept requestId
DoEvents
End Sub
Sub clearusers(sock)
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
users(sock).sending = False
'any future ones here
ReDim users(sock).log(0)
End Sub
Private Sub Winsock_onError(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock_onClose (Index)
End Sub
Private Sub winsock_OnSendComplete(Index As Integer)
If Index = 0 Then Exit Sub
DoEvents
users(Index).sending = False
End Sub
