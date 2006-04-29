VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   ScaleHeight     =   1785
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   2160
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   4060
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Winsock1(0).LocalPort = 4060
Winsock1(0).Listen
End Sub
Private Sub Winsock1_Close(Index As Integer)
On Error Resume Next
Winsock1(Index).Close
If Index = 0 Then
    Winsock1(0).Listen
Else
    Unload Winsock1(Index)
End If
End Sub
Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
num = Winsock1.UBound + 1
Load Winsock1(num)
Winsock1(num).Accept requestID
DoEvents
End Sub
Private Sub timeout(datime)
x = Timer
Do: DoEvents
Loop While x + datime > Timer
End Sub
Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error GoTo daerror
Winsock1(Index).GetData dadata, vbString
'Version:3.4.2
'update=http://ns3.mmchat.example.org/mmchat.exe
'192.168.1.1 4854 Diro-Chat
'203.0.113.10 8724 Town-100
'203.0.113.11 9797 Town-97
dafile = FreeFile
Open App.Path & "\auth.ini" For Input As dafile
dafilecon = Input(LOF(dafile), dafile)
Close dafile
daver = Mid(dafilecon, 1, InStr(dafilecon, vbCrLf) - 1)
udurl = Mid(dafilecon, InStr(dafilecon, "update=") + 7, InStr(InStr(dafilecon, "update="), dafilecon, vbCrLf) - InStr(dafilecon, "update=") - 7)
daloc = InStr(dafilecon, Winsock1(Index).RemoteHostIP)
If InStr(dadata, daver) = 0 Then
    Winsock1(Index).SendData udurl
    timeout 2
    List1.AddItem Winsock1(Index).RemoteHostIP & " Old Version  " & Time & " " & Date
    List1.ListIndex = List1.ListCount - 1
    Winsock1_Close (Index)
    Exit Sub
End If

If daloc = 0 Then
    If InStr(dadata, "DEMO") Then
        List1.AddItem Winsock1(Index).RemoteHostIP & " DEMO VERSION " & Time & " " & Date
        Winsock1(Index).SendData "OK"
        timeout 1
        Winsock1_Close (Index)
    Else
        List1.AddItem Winsock1(Index).RemoteHostIP & " UNKNOWN USER!!! " & Time & " " & Date
        List1.ListIndex = List1.ListCount - 1
        Winsock1_Close (Index)
        Exit Sub
    End If
Else
    dainfo = Mid(dafilecon, daloc, InStr(daloc, dafilecon, vbCrLf) - daloc)
    Dim params() As String
    params() = Split(dainfo, " ")
    who = params(2)
    auth = params(1)
    Dim daparams() As String
    daparams = Split(dadata, ",")
    num1 = Val(daparams(0))
    num2 = Val(daparams(1))
    List1.AddItem Winsock1(Index).RemoteHostIP & " " & daparams(2) & " " & who & " " & Time & " " & Date
    List1.ListIndex = List1.ListCount - 1
    num1 = num1 / Val(Mid(auth, 1, 2))
    num2 = num2 / Val(Mid(auth, 3, 2))
    Winsock1(Index).SendData Str(auth + num1 + num2)
    timeout 2
    Winsock1_Close (Index)
End If
Exit Sub
daerror:
Winsock1_Close (Index)
Exit Sub
End Sub


Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock1_Close (Index)
End Sub
