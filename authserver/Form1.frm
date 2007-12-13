VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   ScaleHeight     =   1410
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
   End
   Begin Project1.DNS DNS1 
      Left            =   720
      Top             =   600
      _ExtentX        =   1931
      _ExtentY        =   873
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   3360
      Top             =   1920
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

Winsock1(0).LocalPort = 4062
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
On Error GoTo errored
Winsock1(Index).GetData dadata, vbString
'ver,dns,port,authnum,checknum
params() = Split(dadata, ",")
daver = params(0)
dahost = params(1)
daport = params(2)
checknum = params(3)

'check ver
If daver <> GetFromINI("Main", "Version", "", App.Path & "\auth.ini") Then
    Winsock1(Index).SendData "Update"
    timeout 1
    Winsock1_Close (Index)
    Exit Sub
End If

If dahost = "IP" Then
    dahost = Winsock1(Index).RemoteHostIP
Else
    If Winsock1(Index).RemoteHostIP <> DNS1.NameToAddress(dahost) Then
        Winsock1(Index).SendData "IP-Mismatch"
        timeout 1
        Winsock1_Close (Index)
        Exit Sub
    End If
End If

iniauthnum = GetFromINI(dahost, "authnum", "", App.Path & "\auth.ini")
If iniauthnum = "" Then
    Winsock1(Index).SendData "Unknown-Host"
    timeout 1
    Winsock1_Close (Index)
    Exit Sub
End If



errored:
Winsock1_Close (Index)
End Sub


Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock1_Close (Index)
End Sub
