VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   9585
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   LinkTopic       =   "Form9"
   ScaleHeight     =   9585
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000B&
      Caption         =   "<--BACK"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   240
      Width           =   1215
   End
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
      Left            =   18840
      TabIndex        =   8
      Top             =   240
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\project\db\customer.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\project\db\customer.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM CUSTOMER"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Caption         =   "Operations"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   1215
      Left            =   3720
      TabIndex        =   1
      Top             =   960
      Width           =   12975
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   2775
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000FF00&
         Caption         =   "SEARCH BY NAME / BIIL_NO"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFF80&
         Caption         =   "UPDATE"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H000000FF&
         Caption         =   "DELETE"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H0080FFFF&
         Caption         =   "REFRESH"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10320
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   2295
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form9.frx":0000
      Height          =   8775
      Left            =   0
      TabIndex        =   7
      Top             =   2280
      Width           =   20535
      _ExtentX        =   36221
      _ExtentY        =   15478
      _Version        =   393216
      BackColor       =   12640511
      ForeColor       =   16711680
      HeadLines       =   1
      RowHeight       =   27
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   120
      Picture         =   "Form9.frx":0015
      Stretch         =   -1  'True
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   18360
      Picture         =   "Form9.frx":E943
      Stretch         =   -1  'True
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "CUSTOMER DETAILS"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20535
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form10.Show
Form9.Hide
End Sub

Private Sub Command2_Click()
Adodc1.RecordSource = "select * from CUSTOMER where BILL_NO like'" & Text1.Text & "' or CUSTOMER_NAME like'" & Text1.Text & "' "
Adodc1.Refresh
If Text1.Text = "" Then


MsgBox "Plese Enter The Name/ID Properlly"
End If
End Sub

Private Sub Command6_Click()
Form1.Show
Form9.Hide
End Sub

Private Sub Command7_Click()
Adodc1.Recordset.Update
MsgBox "UPDATED SUCCESSFULL", vbInformation
Adodc1.RecordSource = "select * from CUSTOMER"
Adodc1.Refresh
End Sub

Private Sub Command8_Click()
conform = MsgBox("Do You Want To Delete The Record", vbYesNo + vbExclamation, "warning message")
If conform = vbYes Then
Adodc1.Recordset.Delete
MsgBox "Record Deleted", vbInformation, "delete record conformation"
Else
MsgBox "Record not deleted", vbInformation, "not deleted"
End If
End Sub

Private Sub Command9_Click()
Adodc1.RecordSource = "select * from CUSTOMER"
Adodc1.Refresh
End Sub
