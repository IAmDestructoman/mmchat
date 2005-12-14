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
Winsock1.LocalPort = 4060
Winsock1.Listen
End Sub

Private Sub Winsock1_Close()
Winsock1.Close
Winsock1.Listen
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Winsock1.Close
Winsock1.Accept requestID
End Sub
Private Sub timeout(datime)
x = Timer
Do: DoEvents
Loop While x + datime > Timer
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.GetData dadata, vbString
'Version:3.4.2
'update=http://ns3.mmchat.example.org/mmchat.exe
'192.168.1.1 4854 Diro-Chat
'203.0.113.10 8724 Town-100
'203.0.113.11 9797 Town-97
Open App.Path & "\auth.ini" For Input As 1
dafilecon = Input(LOF(1), 1)
Close 1
daver = Mid(dafilecon, 1, InStr(dafilecon, vbCrLf) - 1)
udurl = Mid(dafilecon, InStr(dafilecon, "update=") + 7, InStr(InStr(dafilecon, "update="), dafilecon, vbCrLf) - InStr(dafilecon, "update=") - 7)
daloc = InStr(dafilecon, Winsock1.RemoteHostIP)
If daloc = 0 Then
    List1.AddItem Winsock1.RemoteHostIP & " UNKNOWN USER!!! " & Time & " " & Date
    List1.ListIndex = List1.ListCount - 1
    Winsock1.Close
    Winsock1.Listen
    Exit Sub
Else
    dainfo = Mid(dafilecon, daloc, InStr(daloc, dafilecon, vbCrLf) - daloc)
    Dim params() As String
    params() = Split(dainfo, " ")
    who = params(2)
    auth = params(1)
End If
If InStr(dadata, daver) = 0 Then
    Winsock1.SendData udurl
    timeout 2
    List1.AddItem Winsock1.RemoteHostIP & " Old Version  " & who & " " & Time & " " & Date
    List1.ListIndex = List1.ListCount - 1
    Winsock1.Close
    Winsock1.Listen
    Exit Sub
End If
Dim daparams() As String
daparams = Split(dadata, ",")
num1 = Val(daparams(0))
num2 = Val(daparams(1))
List1.AddItem Winsock1.RemoteHostIP & " " & daparams(2) & " " & who & " " & Time & " " & Date
List1.ListIndex = List1.ListCount - 1
num1 = num1 / Val(Mid(auth, 1, 2))
num2 = num2 / Val(Mid(auth, 3, 2))
Winsock1.SendData Str(auth + num1 + num2)
timeout 2
Winsock1.Close
Winsock1.Listen
Exit Sub



End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock1_Close
End Sub
