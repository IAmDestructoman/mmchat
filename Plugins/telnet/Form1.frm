VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock2 
      Index           =   0
      Left            =   1560
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   720
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
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

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next
Load Winsock1(Winsock1.UBound + 1)
Load Winsock2(Winsock1.UBound)
Winsock1(Winsock1.UBound).Accept requestID
Winsock1(Winsock1.UBound).SendData "Welcome to telnet-2-mmchat emulator plugin by Dirocus" & vbCrLf & "Please read help.  Type .help to do so"
End Sub
Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
Winsock1(Index).PeekData dadata, vbString
If InStr(dadata, vbCrLf) = 0 Then Exit Sub
Winsock1(Index).GetData dadata, vbString
dadata = Trim(dadata)
If Mid(dadata, 1, 1) = "." Then
    If Len(dadata) < 5 Then
        Exit Sub
    End If
    Select Case Mid(dadata, 2, 4)
    Case "setn"
        dadata = Mid(dadata, 1, Len(dadata) - 2)
        Winsock1(Index).Tag = Mid(dadata, 7, Len(dadata))
        Winsock1(Index).SendData "Chatname set to " & Winsock1(Index).Tag & vbCrLf
        Winsock1(Index).SendData "Tring to connect.." & vbCrLf
        Winsock2(Index).Connect Winsock1(0).LocalIP, Val(GetFromINI("Main", "Port", "4050", App.Path & "\..\mmchat.ini"))
    Case "help"
        Winsock1(Index).SendData "Use the following commands.." & vbCrLf & ".setn to set your chat name" & vbCrLf & ".chat to chat to the server"
    Case "chat"
        Winsock2(Index).SendData Chr(textper) & Winsock1(Index).Tag & " chats to you, '" & Mid(dadata, 7, Len(dadata) - 8) & "'" & Chr(endcom)
    Case Else
        Winsock1(Index).SendData "Not understood" & vbCrLf
    End Select
    Exit Sub
End If
If Winsock1(Index).Tag = "" Then
    Winsock1(Index).SendData "Welcome to telnet-2-mmchat emulator plugin by Dirocus" & vbCrLf & "Please read help.  Type .help to do so" & vbCrLf
    Exit Sub
End If
Winsock2(Index).SendData Chr(textall) & Winsock1(Index).Tag & " chats to everybody, '" & Replace(dadata, vbCrLf, "") & "'" & Chr(10) & Chr(endcom)
Winsock1(Index).SendData "You chat to everybody, '" & Replace(dadata, vbCrLf, "") & "'" & vbCrLf
End Sub

Private Sub Winsock2_Connect(Index As Integer)
On Error Resume Next
Winsock2(Index).SendData "CHAT:" & Winsock1(Index).Tag & Chr(10) & Winsock1(Index).RemoteHostIP & "00000"
End Sub

Private Sub Winsock2_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
Winsock2(Index).GetData dadata, vbString
dadata = Replace(dadata, Chr(10), vbCrLf)
If Mid(dadata, 1, 1) = Chr(ping) Then Exit Sub
Winsock1(Index).SendData Mid(dadata, 2, Len(dadata) - 2)


End Sub
Private Sub closesocket(Index)
On Error Resume Next
Winsock1(Index).Close
Winsock2(Index).Close
Unload Winsock1(Index)
Unload Winsock2(Index)

End Sub
Private Sub Winsock1_Close(Index As Integer)
closesocket (Index)
End Sub
Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
closesocket (Index)
End Sub
Private Sub Winsock2_Close(Index As Integer)
closesocket (Index)
End Sub
Private Sub Winsock2_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
closesocket (Index)
End Sub
