VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   LinkTopic       =   "Form5"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00008000&
      Caption         =   "Search Betwwen Dates"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1575
      Left            =   5160
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   10335
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   960
         Width           =   3735
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   7200
         TabIndex        =   11
         Top             =   480
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   8454143
         CalendarForeColor=   0
         CalendarTitleBackColor=   16711680
         CalendarTitleForeColor=   33023
         CalendarTrailingForeColor=   16776960
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy"
         DateIsNull      =   -1  'True
         Format          =   119603203
         CurrentDate     =   44569
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2520
         TabIndex        =   10
         Top             =   480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   8454143
         CalendarTitleBackColor=   16711680
         CalendarTitleForeColor=   33023
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy"
         DateIsNull      =   -1  'True
         Format          =   119603203
         CurrentDate     =   44569
         MaxDate         =   43100
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "To Date"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   13
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "From Date"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   1935
      End
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
      Left            =   18720
      TabIndex        =   3
      Top             =   240
      Width           =   1215
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
      TabIndex        =   2
      Top             =   1080
      Width           =   12975
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
         TabIndex        =   8
         Top             =   360
         Width           =   2295
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
         TabIndex        =   7
         Top             =   360
         Width           =   2055
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
         TabIndex        =   6
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000FF00&
         Caption         =   "SEARCH BY NAME / ID"
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
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
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
         TabIndex        =   4
         Top             =   480
         Width           =   2775
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form5.frx":0000
      Height          =   4455
      Left            =   1560
      TabIndex        =   0
      Top             =   4680
      Width           =   17415
      _ExtentX        =   30718
      _ExtentY        =   7858
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   12720
      Top             =   10320
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\project\db\register.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\project\db\register.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from regdetails"
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
   Begin VB.Image Image4 
      Height          =   375
      Left            =   0
      Picture         =   "Form5.frx":0015
      Stretch         =   -1  'True
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   18240
      Picture         =   "Form5.frx":E943
      Stretch         =   -1  'True
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   " STAFF  DETAILS"
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
      TabIndex        =   1
      Top             =   0
      Width           =   20535
   End
   Begin VB.Image Image1 
      Height          =   11880
      Left            =   0
      Picture         =   "Form5.frx":10163
      Stretch         =   -1  'True
      Top             =   240
      Width           =   20745
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conform As Integer
Dim date1 As String
Dim date2 As String





Private Sub Command1_Click()
date1 = Format(DTPickar1.Value, "mm/dd/yyyy")
date2 = Format(DTPickar2.Value, "mm/dd/yyyy")
If date2 < date1 Then
MsgBox "please select the correct date:End date cannot be lesser than start date", vbCritical, "warning message"
Else
Adodc1.RecordSource = "Select * from regdetails where start_date between # " & date1 & " # and # " & date2 & " # "
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
MsgBox "record not found,", vbCritical, "Warring Message"
Else
Adodc1.Caption = Adodc1.RecordSource
End If
End If




End Sub

Private Sub Command2_Click()
Adodc1.RecordSource = "select * from regdetails where ID like'" & Text1.Text & "' or username like'" & Text1.Text & "' "
Adodc1.Refresh
If Text1.Text = "" Then


MsgBox "Plese Enter The Name/ID Properlly"

End If





End Sub


Private Sub Command3_Click()
Form3.Show
Form5.Hide
End Sub

Private Sub Command6_Click()
Form1.Show
Form5.Hide
End Sub

Private Sub Command7_Click()
Adodc1.Recordset.Update
MsgBox "UPDATED SUCCESSFULL", vbInformation
Adodc1.RecordSource = "select * from regdetails"
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
Adodc1.RecordSource = "select * from regdetails"
Adodc1.Refresh
End Sub
