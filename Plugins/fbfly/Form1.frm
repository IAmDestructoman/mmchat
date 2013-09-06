VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6135
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox s3 
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Text            =   "Text5"
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox s2 
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox s1 
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox a3 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "Text5"
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox a2 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox a1 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "distance"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "heading"
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo errored

Dim PI#
PI# = 4 * Atn(1#)
dadir = Text2 - (90 * (Fix(Text2 / 90)))
a3 = 90
a2 = 90 - dadir
a1 = dadir
s3 = Text1
s1 = Round((s3 * Sin(a1 * PI# / 180)) / Sin(a3 * PI# / 180) / 10, 0)
s2 = Round((s3 * Sin(a2 * PI# / 180)) / Sin(a3 * PI# / 180) / 10, 0)
Select Case Val(Text2)
    Case Is < 90
    Form1.Caption = "E" & s1 & " N" & s2
    Case Is < 180
    Form1.Caption = "E" & s2 & " S" & s1
    Case Is < 270
    Form1.Caption = "W" & s1 & " S" & s2
    Case Is < 360
    Form1.Caption = "W" & s2 & " N" & s1
End Select
errored:
Exit Sub
End Sub
