VERSION 5.00
Begin VB.Form Form10 
   Caption         =   "Form10"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9540
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   9540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "LOGOUT -->"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "CUSTOMER_DETAILS"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "STAFF_DETAILS"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   7440
      Picture         =   "Form10.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   0
      Picture         =   "Form10.frx":1820
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9540
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form5.Show
Form10.Hide

End Sub

Private Sub Command2_Click()
Form9.Show
Form10.Hide
End Sub

Private Sub Command6_Click()
Form1.Show
Form10.Hide
End Sub
