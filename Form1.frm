VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   9645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Help"
      Height          =   375
      Left            =   7440
      TabIndex        =   3
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   9615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   8520
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Tag = "submitted"
End Sub

Private Sub Command2_Click()
damessage = damessage & "Color codes:" & vbCrLf
damessage = damessage & "@Bold    - Ansi Bold" & vbCrLf
damessage = damessage & "@Reset   - Ansi Reset" & vbCrLf
damessage = damessage & "@Black   - Ansi Black" & vbCrLf
damessage = damessage & "@Blue    - Ansi Blue" & vbCrLf
damessage = damessage & "@Cyan    - Ansi Cyan" & vbCrLf
damessage = damessage & "@Green   - Ansi Green" & vbCrLf
damessage = damessage & "@Magenta - Ansi Magenta" & vbCrLf
damessage = damessage & "@Red     - Ansi Red" & vbCrLf
damessage = damessage & "@Yellow  - Ansi Yellow" & vbCrLf
damessage = damessage & "@White   - Ansi White" & vbCrLf
MsgBox damessage, vbOKOnly, "MMchat"
End Sub

Private Sub Form_Load()
Form1.Tag = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Tag = ""
End Sub

