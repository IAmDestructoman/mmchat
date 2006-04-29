VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   8295
   StartUpPosition =   3  'Windows Default
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
      Height          =   2895
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Form1.frx":0000
      Top             =   480
      Width           =   8295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5895
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

Private Sub Form_Load()
Form1.Tag = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Tag = ""
End Sub

