VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   1680
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   855
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   2040
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









Private Sub Timer1_Timer()
Form1.Caption = Winsock1.State

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim Text As String
   
    Winsock1.GetData Text, vbString
    
        If Asc(Mid(Text, 2, 1)) = 1 Then
            'Send the log on information
            Winsock1.SendData Chr(42) & Chr(1) & Chr(0) & Chr(0) & Chr(0) & Chr(8 + Len(Text1)) & Chr(0) & Chr(0) & Chr(0) & Chr(1) & Chr(0) & Chr(1) & Chr(0) & Chr(Len(Text1)) & Trim(Text1)
            DoEvents
            Winsock1.SendData Chr(42) & Chr(2) & Chr(0) & Chr(1) & Chr(0) & Chr(Len(Text1) + Len(AIM_EncryptPW(Trim(Text2))) + 90) & "toc2_signon login.oscar.aol.com 29999 " & Trim(Text1) & " " & AIM_EncryptPW(Trim(Text2)) & " english-US " & Chr(34) & "TIC:\$Revision: 1.83 \$" & Chr(34) & " 160 " & AIM_Algorithum(Text1, Text2) & Chr(0)
            DoEvents
            

        End If
        If Asc(Mid(Text, 2, 1)) = 2 And Winsock1.Tag <> "online" Then
            DoEvents
            Winsock1.SendData Chr(42) & Chr(2) & Chr(0) & Chr(2) & Chr(0) & Chr(14) & "toc_init_done" & Chr(0)
            DoEvents
         
            Winsock1.Tag = "online"
        End If


End Sub
Function timeout(xx)
x = Timer
Do: DoEvents
Loop While x + xx > Timer
End Function
Public Function AIM_Algorithum(ByVal sUser As String, ByVal sPass As String) As String
    'This is the code that generates the 8 o
    '     r 9 digit number on the end of
    'the logon packet. Uses the Screen Name
    '     and Password to make it
    Dim sUserChar As Long, sVar As Long
    sUser = Left(LCase(sUser), 1)
    sUserChar = Int(Asc(sUser) - 96)
    sVar = Int(sUserChar * 7696) + 738816
    sBase = Int(sUserChar * 746512)
    sVal = Int(Asc(Left(LCase(sPass), 1)) - 96) * sVar
    AIM_Algorithum = Int(Int(sVal) - sVar) + Int(sBase + 71665152)
End Function


Public Function AIM_EncryptPW(ByVal sPass As String) As String
    'This will take the password, and encryp
    '     t it using the word "Tic/Toc"
    Dim vTable() As Variant, sString As String
    Dim sLoop As Long, sHex As String
    vTable = Array("84", "105", "99", "47", "84", "111", "99")
    sString = "0x"


    For sLoop = 0 To Len(sPass) - 1
        sHex = Hex(Asc(Mid(sPass, sLoop + 1, 1)) Xor CLng(vTable(sLoop Mod 7)))


        If CLng("&H" & sHex) < 16 Then
            sString = sString & "0"
        End If
        sString = sString & sHex
    Next
    AIM_EncryptPW = LCase(sString)
End Function
