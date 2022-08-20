VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   8865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11595
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11595
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   12600
      Top             =   9000
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
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\project\db\printout.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\project\db\printout.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "print"
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
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   13560
      Top             =   8280
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000080FF&
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   2040
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   1200
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Get Data"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   480
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   7215
      Begin VB.Line Line18 
         X1              =   3240
         X2              =   7200
         Y1              =   8880
         Y2              =   8880
      End
      Begin VB.Line Line17 
         X1              =   0
         X2              =   7200
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Line Line16 
         X1              =   0
         X2              =   7200
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Line Line15 
         X1              =   0
         X2              =   7200
         Y1              =   4800
         Y2              =   4800
      End
      Begin VB.Line Line14 
         X1              =   0
         X2              =   7200
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Line Line13 
         X1              =   0
         X2              =   7200
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Image Image2 
         Height          =   945
         Left            =   120
         Picture         =   "Form8.frx":0000
         Stretch         =   -1  'True
         Top             =   7200
         Width           =   3015
      End
      Begin VB.Line Line12 
         X1              =   3240
         X2              =   7200
         Y1              =   7920
         Y2              =   7920
      End
      Begin VB.Line Line11 
         X1              =   3240
         X2              =   7200
         Y1              =   7320
         Y2              =   7320
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         DataField       =   "INC_TAX"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   45
         Top             =   8160
         Width           =   1335
      End
      Begin VB.Label Label44 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   "Including Tax"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   44
         Top             =   8160
         Width           =   2175
      End
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         DataField       =   "TAX"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   43
         Top             =   7440
         Width           =   1335
      End
      Begin VB.Label Label42 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Caption         =   "Tax"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   42
         Top             =   7440
         Width           =   1095
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         DataField       =   "TOTAL_QTY"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   41
         Top             =   6840
         Width           =   735
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         DataField       =   "AMOUNT"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   40
         Top             =   6840
         Width           =   1335
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   39
         Top             =   6840
         Width           =   1095
      End
      Begin VB.Line Line10 
         X1              =   0
         X2              =   7200
         Y1              =   6720
         Y2              =   6720
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         DataField       =   "M_TOTAL"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   38
         Top             =   6240
         Width           =   1335
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         DataField       =   "C_TOTAL"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   37
         Top             =   5520
         Width           =   1335
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         DataField       =   "T_TOTAL"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   36
         Top             =   4920
         Width           =   1335
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         DataField       =   "G_TOTAL"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   35
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         DataField       =   "S_TOTAL"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   34
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         DataField       =   "D_TOTAL"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   33
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Line Line9 
         X1              =   5640
         X2              =   5640
         Y1              =   2400
         Y2              =   8880
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         DataField       =   "M_PRICE"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   32
         Top             =   6240
         Width           =   1095
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         DataField       =   "C_PRICE"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   31
         Top             =   5520
         Width           =   1095
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         DataField       =   "T_PRICE"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   30
         Top             =   4920
         Width           =   1095
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         DataField       =   "G_PRICE"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   29
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         DataField       =   "S_PRICE"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   28
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         DataField       =   "D_PRICE"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   27
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Line Line8 
         X1              =   4320
         X2              =   4320
         Y1              =   2520
         Y2              =   7320
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         DataField       =   "M_QTY"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   26
         Top             =   6240
         Width           =   735
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         DataField       =   "C_QTY"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   25
         Top             =   5520
         Width           =   735
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         DataField       =   "T_QTY"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   24
         Top             =   4920
         Width           =   735
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         DataField       =   "G_QTY"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   23
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         DataField       =   "S_QTY"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   22
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         DataField       =   "D_QTY"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   21
         Top             =   3120
         Width           =   735
      End
      Begin VB.Line Line7 
         X1              =   3240
         X2              =   3240
         Y1              =   2400
         Y2              =   8880
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         DataField       =   "MISCELLANEOUS"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   6240
         Width           =   3015
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         DataField       =   "COOKING OIL"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   5520
         Width           =   3015
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         DataField       =   "TOOTH_PASTE"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   4920
         Width           =   3015
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         DataField       =   "GRAINS"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   4320
         Width           =   3015
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         DataField       =   "SOAP"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   3720
         Width           =   3015
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         DataField       =   "DETERGENT"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   3120
         Width           =   3015
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5280
         TabIndex        =   14
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4200
         TabIndex        =   13
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Qty"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3480
         TabIndex        =   12
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   120
         TabIndex        =   11
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         X1              =   0
         X2              =   7200
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   0
         X2              =   7200
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line4 
         BorderWidth     =   3
         X1              =   0
         X2              =   3480
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         DataField       =   "Customer_Name"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2280
         TabIndex        =   10
         Top             =   1920
         Width           =   4335
      End
      Begin VB.Label Label4 
         BackColor       =   &H000080FF&
         Caption         =   "Customer_Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   0
         X2              =   7200
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   8
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Line Line2 
         BorderWidth     =   3
         X1              =   3480
         X2              =   7080
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line1 
         BorderStyle     =   6  'Inside Solid
         BorderWidth     =   3
         X1              =   3480
         X2              =   3480
         Y1              =   1080
         Y2              =   1800
      End
      Begin VB.Label Label9 
         BackColor       =   &H0000FF00&
         Caption         =   "GST-NO :- AATY34576Q23Q"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Label date 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         DataField       =   "DATE"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   6
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   5
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   4
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         DataField       =   "Bill_No"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1080
         TabIndex        =   3
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H000080FF&
         Caption         =   "Bill_No"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   2
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000040&
         Caption         =   "Big Supermarket"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   855
         Left            =   1680
         TabIndex        =   1
         Top             =   120
         Width           =   5535
      End
      Begin VB.Image Image1 
         Height          =   1095
         Left            =   0
         Picture         =   "Form8.frx":B613
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   1680
      End
   End
   Begin VB.Image Image3 
      Height          =   8880
      Left            =   7200
      Picture         =   "Form8.frx":14812
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6060
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Form8.Caption = Form7.DTPicker1
date.Caption = Form7.DTPicker1

Form8.Caption = Form7.Label33 'bill
Label3.Caption = Form7.Label33

Form8.Caption = Form7.Text1.Text 'cust_name
Label5.Caption = Form7.Text1.Text

Form8.Caption = Form7.Combo1.Text 'p1
Label15.Caption = Form7.Combo1.Text

Form8.Caption = Form7.Combo2.Text 'p2
Label16.Caption = Form7.Combo2.Text

Form8.Caption = Form7.Combo3.Text 'p3
Label17.Caption = Form7.Combo3.Text

Form8.Caption = Form7.Combo4.Text 'p4
Label18.Caption = Form7.Combo4.Text

Form8.Caption = Form7.Combo5.Text 'p5
Label19.Caption = Form7.Combo5.Text

Form8.Caption = Form7.Combo6.Text 'p6
Label20.Caption = Form7.Combo6.Text

Form8.Caption = Form7.Text3.Text 'q1
Label21.Caption = Form7.Text3.Text

Form8.Caption = Form7.Text9.Text 'q2
Label22.Caption = Form7.Text9.Text

Form8.Caption = Form7.Text12.Text 'q3
Label23.Caption = Form7.Text12.Text

Form8.Caption = Form7.Text15.Text 'q4
Label24.Caption = Form7.Text15.Text

Form8.Caption = Form7.Text18.Text 'q5
Label25.Caption = Form7.Text18.Text

Form8.Caption = Form7.Text21.Text 'q6
Label26.Caption = Form7.Text21.Text

Form8.Caption = Form7.Text4.Text 'pr1
Label27.Caption = Form7.Text4.Text

Form8.Caption = Form7.Text8.Text 'pr2
Label28.Caption = Form7.Text8.Text

Form8.Caption = Form7.Text11.Text 'pr3
Label29.Caption = Form7.Text11.Text

Form8.Caption = Form7.Text14.Text 'pr4
Label30.Caption = Form7.Text14.Text

Form8.Caption = Form7.Text17.Text 'pr5
Label31.Caption = Form7.Text17.Text

Form8.Caption = Form7.Text20.Text 'pr6
Label32.Caption = Form7.Text20.Text

Form8.Caption = Form7.Label36 't1
Label33.Caption = Form7.Label36

Form8.Caption = Form7.Label37 't2
Label34.Caption = Form7.Label37

Form8.Caption = Form7.Label38 't3
Label35.Caption = Form7.Label38

Form8.Caption = Form7.Label39 't4
Label36.Caption = Form7.Label39

Form8.Caption = Form7.Label40 't5
Label37.Caption = Form7.Label40

Form8.Caption = Form7.Label41 't6
Label38.Caption = Form7.Label41

Form8.Caption = Form7.Label35
Label41.Caption = Form7.Label35

Form8.Caption = Form7.Label30
Label40.Caption = Form7.Label30

Form8.Caption = Form7.Label31
Label43.Caption = Form7.Label31

Form8.Caption = Form7.Label32
Label45.Caption = Form7.Label32



Adodc1.Recordset.Fields("Bill_No") = Label3
Adodc1.Recordset.Fields("Customer_name") = Label5

Adodc1.Recordset.Fields("DETERGENT") = Label15
Adodc1.Recordset.Fields("D_QTY") = Label21
Adodc1.Recordset.Fields("D_PRICE") = Label27
Adodc1.Recordset.Fields("D_TOTAL") = Label33

Adodc1.Recordset.Fields("SOAP") = Label16
Adodc1.Recordset.Fields("S_QTY") = Label22
Adodc1.Recordset.Fields("S_PRICE") = Label28
Adodc1.Recordset.Fields("S_TOTAL") = Label34

Adodc1.Recordset.Fields("GRAINS") = Label17
Adodc1.Recordset.Fields("G_QTY") = Label23
Adodc1.Recordset.Fields("G_PRICE") = Label29
Adodc1.Recordset.Fields("G_TOTAL") = Label35

Adodc1.Recordset.Fields("TOOTH_PASTE") = Label18
Adodc1.Recordset.Fields("T_QTY") = Label24
Adodc1.Recordset.Fields("T_PRICE") = Label30
Adodc1.Recordset.Fields("T_TOTAL") = Label36

Adodc1.Recordset.Fields("COOKING OIL") = Label19
Adodc1.Recordset.Fields("C_QTY") = Label25
Adodc1.Recordset.Fields("C_PRICE") = Label31
Adodc1.Recordset.Fields("C_TOTAL") = Label37

Adodc1.Recordset.Fields("MISCELLANEOUS") = Label20
Adodc1.Recordset.Fields("M_QTY") = Label26
Adodc1.Recordset.Fields("M_PRICE") = Label32
Adodc1.Recordset.Fields("M_TOTAL") = Label38

Adodc1.Recordset.Fields("AMOUNT") = Label40
Adodc1.Recordset.Fields("TAX") = Label43
Adodc1.Recordset.Fields("INC_TAX") = Label45

Adodc1.Recordset.Fields("TOTAL_QTY") = Label41
Adodc1.Recordset.Fields("DATE") = date


Adodc1.Recordset.Update
MsgBox "Added Successful", vbInformation

End Sub

Private Sub Command2_Click()
Form7.Show
Form8.Hide
Adodc1.RecordSource = "Delete from print where ID like'" & Label3 & "'"



End Sub

Private Sub Command3_Click()
If DataEnvironment2.rsCommand1.State = 1 Then DataEnvironment2.rsCommand1.Close
DataEnvironment2.rsCommand1.Filter = "BILL_NO = " & Val(Label3)
DataReport2.Show
Adodc1.Recordset.Delete
End Sub

Private Sub Form_Load()
Adodc1.Recordset.AddNew
lblTime.Caption = Time


End Sub


