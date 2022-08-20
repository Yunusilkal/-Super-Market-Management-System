VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   9315
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13620
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   13620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "LOG_OUT"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   3
      Left            =   11760
      OLEDropMode     =   1  'Manual
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "GENERATE_BILL"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6120
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "CUSTOMER_DETAILS"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   3375
      Left            =   2760
      TabIndex        =   0
      Top             =   3840
      Width           =   8175
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FFFF&
         Caption         =   "STAFF_DETAILS"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   2895
      End
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   11400
      Picture         =   "Form3.frx":0000
      Stretch         =   -1  'True
      Top             =   960
      Width           =   345
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "SUPER MARKET MANGEMENT "
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   13695
   End
   Begin VB.Image Image1 
      Height          =   9285
      Left            =   0
      Picture         =   "Form3.frx":1820
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13680
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Form4.Show
Form3.Hide


End Sub

Private Sub Command2_Click(Index As Integer)
Form6.Show
Form3.Hide

End Sub

Private Sub Command3_Click(Index As Integer)
Form7.Show
Form3.Hide

End Sub

Private Sub Command4_Click(Index As Integer)
Form1.Show
Form3.Hide

End Sub

Private Sub Command5_Click(Index As Integer)
Form7.Show
Form3.Hide
End Sub
