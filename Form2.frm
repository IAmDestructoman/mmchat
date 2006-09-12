VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1695
   LinkTopic       =   "Form2"
   ScaleHeight     =   255
   ScaleWidth      =   1695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "Updating... Please wait."
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
   Begin ChatServer.Socket Socket1 
      Left            =   1920
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      LocalPort       =   0
      RemoteHost      =   ""
      RemotePort      =   0
      Protocol        =   0
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private dafile
Private Sub Form_Load()
Socket1.Connect GetFromINI("Main", "Authserver", "ns1.mmchat.example.org", App.Path & "\mmchat.ini"), 80

End Sub

Private Sub Socket1_onClose()
Open App.Path & "\mmchat.new.exe" For Output As 9
crlfpos = InStr(dafile, vbCrLf & vbCrLf) + 4
dafile = Mid(dafile, crlfpos, Len(dafile))
Print #9, dafile;
Close 9
Socket1.CloseSocket
FileCopy App.Path & "\mmchat.new.exe", App.Path & "\mmchat.prg"
Shell App.Path & "\mmchat.new.exe"

End

End Sub

Private Sub Socket1_onConnect()

Socket1.SendData "GET /mmchat/mmchat.exe HTTP/1.1" & vbCrLf & "Host: mmchat.example.org" & vbCrLf & vbCrLf



End Sub

Private Sub Socket1_onDataArrival(ByVal bytesTotal As Long)
Socket1.GetData dadata, vbString
dafile = dafile & dadata

End Sub

Private Sub Socket1_onError(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
End
End Sub


