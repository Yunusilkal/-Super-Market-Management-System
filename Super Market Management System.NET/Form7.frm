VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   20250
   LinkTopic       =   "Form7"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker DTPicker1 
      DataField       =   "DATE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   17520
      TabIndex        =   60
      Top             =   1200
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   14737632
      CalendarForeColor=   0
      CalendarTitleBackColor=   12582912
      CalendarTitleForeColor=   33023
      CalendarTrailingForeColor=   65535
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   119078913
      CurrentDate     =   44569
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   -960
      Top             =   7680
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\project\db\customer.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\project\db\customer.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "CUSTOMER"
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
   Begin VB.CommandButton Command7 
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
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Generate_Bill_No"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080FFFF&
      Caption         =   "RESET"
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
      Left            =   17760
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   8640
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FFFF&
      Caption         =   "Generate-Bill"
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
      Left            =   17760
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   10200
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
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
      TabIndex        =   53
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
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
      Height          =   405
      Left            =   7560
      TabIndex        =   51
      Top             =   1200
      Width           =   3975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "SAVE"
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
      Left            =   17760
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   9480
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   9000
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Height          =   6015
      Left            =   1080
      TabIndex        =   3
      Top             =   1920
      Width           =   18255
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "TOOTH_PASTE"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   9360
         TabIndex        =   64
         Top             =   480
         Width           =   8415
         Begin VB.TextBox Text14 
            BackColor       =   &H00FFFF80&
            DataField       =   "T_PRICE"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5280
            TabIndex        =   67
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox Text15 
            BackColor       =   &H00FFFF80&
            DataField       =   "T_QTY"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   4080
            TabIndex        =   66
            Top             =   840
            Width           =   975
         End
         Begin VB.ComboBox Combo4 
            BackColor       =   &H00FFFF80&
            DataField       =   "TOOTH_PASTE"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   120
            TabIndex        =   65
            Text            =   "SELECT"
            Top             =   840
            Width           =   3615
         End
         Begin VB.Label Label39 
            BackColor       =   &H00FFFF80&
            DataField       =   "T_TOTAL"
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
            Height          =   375
            Left            =   6720
            TabIndex        =   72
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label16 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6840
            TabIndex        =   71
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label17 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Price"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5400
            TabIndex        =   70
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label18 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Quantiy"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4080
            TabIndex        =   69
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label19 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Select Items"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   68
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "MISCELLANEOUS"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   9360
         TabIndex        =   36
         Top             =   4080
         Width           =   8415
         Begin VB.ComboBox Combo6 
            BackColor       =   &H00FFFF80&
            DataField       =   "MISCELLANEOUS"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   120
            TabIndex        =   39
            Text            =   "SELECT"
            Top             =   840
            Width           =   3615
         End
         Begin VB.TextBox Text21 
            BackColor       =   &H00FFFF80&
            DataField       =   "M_QTY"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   4080
            TabIndex        =   38
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox Text20 
            BackColor       =   &H00FFFF80&
            DataField       =   "M_PRICE"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5280
            TabIndex        =   37
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label41 
            BackColor       =   &H00FFFF80&
            DataField       =   "M_TOTAL"
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
            Height          =   375
            Left            =   6720
            TabIndex        =   74
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label27 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Select Items"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   43
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label26 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Quantiy"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4080
            TabIndex        =   42
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label25 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Price"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5400
            TabIndex        =   41
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label24 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6840
            TabIndex        =   40
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "COOKING OIL"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   9360
         TabIndex        =   28
         Top             =   2280
         Width           =   8415
         Begin VB.ComboBox Combo5 
            BackColor       =   &H00FFFF80&
            DataField       =   "COOKING OIL"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   120
            TabIndex        =   31
            Text            =   "SELECT"
            Top             =   840
            Width           =   3615
         End
         Begin VB.TextBox Text18 
            BackColor       =   &H00FFFF80&
            DataField       =   "C_QTY"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   4080
            TabIndex        =   30
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox Text17 
            BackColor       =   &H00FFFF80&
            DataField       =   "C_PRICE"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5280
            TabIndex        =   29
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label40 
            BackColor       =   &H00FFFF80&
            DataField       =   "C_TOTAL"
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
            Height          =   375
            Left            =   6720
            TabIndex        =   73
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label23 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Select Items"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   35
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label22 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Quantiy"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4080
            TabIndex        =   34
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label21 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Price"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5400
            TabIndex        =   33
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label20 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6840
            TabIndex        =   32
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "GRAINS"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   360
         TabIndex        =   20
         Top             =   4080
         Width           =   8415
         Begin VB.ComboBox Combo3 
            BackColor       =   &H00FFFF80&
            DataField       =   "GRAINS"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   120
            TabIndex        =   23
            Text            =   "SELECT"
            Top             =   840
            Width           =   3615
         End
         Begin VB.TextBox Text12 
            BackColor       =   &H00FFFF80&
            DataField       =   "G_QTY"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   4080
            TabIndex        =   22
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox Text11 
            BackColor       =   &H00FFFF80&
            DataField       =   "G_PRICE"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5280
            TabIndex        =   21
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label38 
            BackColor       =   &H00FFFF80&
            DataField       =   "G_TOTAL"
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
            Height          =   375
            Left            =   6720
            TabIndex        =   63
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label15 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Select Items"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   27
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label14 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Quantiy"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4080
            TabIndex        =   26
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label13 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Price"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5400
            TabIndex        =   25
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label12 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6840
            TabIndex        =   24
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "SOAP"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   360
         TabIndex        =   12
         Top             =   2280
         Width           =   8415
         Begin VB.ComboBox Combo2 
            BackColor       =   &H00FFFF80&
            DataField       =   "SOAP"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   120
            TabIndex        =   15
            Text            =   "SELECT"
            Top             =   840
            Width           =   3615
         End
         Begin VB.TextBox Text9 
            BackColor       =   &H00FFFF80&
            DataField       =   "S_QTY"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   4080
            TabIndex        =   14
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox Text8 
            BackColor       =   &H00FFFF80&
            DataField       =   "S_PRICE"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5280
            TabIndex        =   13
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label37 
            BackColor       =   &H00FFFF80&
            DataField       =   "S_TOTAL"
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
            Height          =   375
            Left            =   6720
            TabIndex        =   62
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label11 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Select Items"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   19
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label10 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Quantiy"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4080
            TabIndex        =   18
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label9 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Price"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5400
            TabIndex        =   17
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label8 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6840
            TabIndex        =   16
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "DETERGENT"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   8415
         Begin VB.TextBox Text4 
            BackColor       =   &H00FFFF80&
            DataField       =   "D_PRICE"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5280
            TabIndex        =   8
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H00FFFF80&
            DataField       =   "D_QTY"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   4080
            TabIndex        =   7
            Top             =   840
            Width           =   975
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00FFFF80&
            DataField       =   "DETERGENT"
            DataSource      =   "Adodc1"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   120
            TabIndex        =   5
            Text            =   "SELECT"
            Top             =   840
            Width           =   3615
         End
         Begin VB.Label Label36 
            BackColor       =   &H00FFFF80&
            DataField       =   "D_TOTAL"
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
            Height          =   375
            Left            =   6720
            TabIndex        =   61
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label7 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6840
            TabIndex        =   11
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label6 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Price"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5400
            TabIndex        =   10
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Quantiy"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4080
            TabIndex        =   9
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Select Items"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   6
            Top             =   480
            Width           =   1455
         End
      End
   End
   Begin VB.Label Label35 
      BackColor       =   &H0080C0FF&
      DataField       =   "TOTAL_QTY"
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
      Height          =   375
      Left            =   2400
      TabIndex        =   59
      Top             =   8400
      Width           =   1815
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Total Quantity"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   58
      Top             =   8400
      Width           =   2295
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   360
      Picture         =   "Form7.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   18360
      Picture         =   "Form7.frx":E92E
      Stretch         =   -1  'True
      Top             =   360
      Width           =   480
   End
   Begin VB.Label Label33 
      BackColor       =   &H00E0E0E0&
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
      Height          =   375
      Left            =   12960
      TabIndex        =   52
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label32 
      BackColor       =   &H0080C0FF&
      DataField       =   "INC_TAX"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   50
      Top             =   10200
      Width           =   1695
   End
   Begin VB.Label Label31 
      BackColor       =   &H0080C0FF&
      DataField       =   "TAX"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   49
      Top             =   9600
      Width           =   2655
   End
   Begin VB.Label Label30 
      BackColor       =   &H0080C0FF&
      DataField       =   "AMOUNT"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   48
      Top             =   9000
      Width           =   1815
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "INCLUDING_TAX"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   46
      Top             =   10200
      Width           =   2295
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "TAX 5%"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   45
      Top             =   9600
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Bill_No"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11880
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Customer_Name"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      Caption         =   "BIG SUPER MARKET "
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20535
   End
   Begin VB.Image Image1 
      Height          =   11760
      Left            =   -240
      Picture         =   "Form7.frx":1014E
      Stretch         =   -1  'True
      Top             =   840
      Width           =   20610
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub Command1_Click()
Dim com1 As Double
Dim com2 As Double
Dim com3 As Double
Dim com4 As Double
Dim com5 As Double
Dim com6 As Double

Dim com7 As Double
Dim com8 As Double
Dim com9 As Double
Dim com10 As Double
Dim com11 As Double
Dim com12 As Double

Dim total As Double
Dim qty As Double


com1 = Label36
com2 = Label37
com3 = Label38
com4 = Label39
com5 = Label40
com6 = Label41

com7 = Text3.Text
com8 = Text9.Text
com9 = Text12.Text
com10 = Text15.Text
com11 = Text18.Text
com12 = Text21.Text

total = (com1 + com2 + com3 + com4 + com5 + com6)

qty = (com7 + com8 + com9 + com10 + com11 + com12)

Label30 = total
Label35 = qty
tax = (5 / 100) * total
Label31 = tax


Dim gtax As Double
Dim gtotal As Double

'------------------------'
gtax = Label31
gtotal = Label30

gtotal = (gtax + gtotal)
Label32 = gtotal




End Sub

Private Sub Command2_Click()
Adodc1.Recordset.Fields("Customer_Name") = Text1.Text

Adodc1.Recordset.Fields("DETERGENT") = Combo1.Text
Adodc1.Recordset.Fields("SOAP") = Combo2.Text
Adodc1.Recordset.Fields("GRAINS") = Combo3.Text
Adodc1.Recordset.Fields("TOOTH_PASTE") = Combo4.Text
Adodc1.Recordset.Fields("COOKING OIL") = Combo5.Text
Adodc1.Recordset.Fields("MISCELLANEOUS") = Combo6.Text

Adodc1.Recordset.Fields("D_QTY") = Text3.Text
Adodc1.Recordset.Fields("S_QTY") = Text9.Text
Adodc1.Recordset.Fields("G_QTY") = Text12.Text
Adodc1.Recordset.Fields("T_QTY") = Text15.Text
Adodc1.Recordset.Fields("C_QTY") = Text18.Text
Adodc1.Recordset.Fields("M_QTY") = Text21.Text

Adodc1.Recordset.Fields("D_PRICE") = Text4.Text
Adodc1.Recordset.Fields("S_PRICE") = Text8.Text
Adodc1.Recordset.Fields("G_PRICE") = Text11.Text
Adodc1.Recordset.Fields("T_PRICE") = Text14.Text
Adodc1.Recordset.Fields("C_PRICE") = Text17.Text
Adodc1.Recordset.Fields("M_PRICE") = Text20.Text

Adodc1.Recordset.Fields("D_TOTAL") = Label36
Adodc1.Recordset.Fields("S_TOTAL") = Label37
Adodc1.Recordset.Fields("G_TOTAL") = Label38
Adodc1.Recordset.Fields("T_TOTAL") = Label39
Adodc1.Recordset.Fields("C_TOTAL") = Label40
Adodc1.Recordset.Fields("M_TOTAL") = Label41

Adodc1.Recordset.Fields("TOTAL_QTY") = Label35
Adodc1.Recordset.Fields("AMOUNT") = Label30
Adodc1.Recordset.Fields("TAX") = Label31
Adodc1.Recordset.Fields("INC_TAX") = Label32
Adodc1.Recordset.Fields("DATE") = DTPicker1

Adodc1.Recordset.Update
MsgBox "Saved Successful", vbInformation

End Sub

Private Sub Command3_Click()
Form1.Show
Form7.Hide
Unload Form7
End Sub

Private Sub Command4_Click()
Form8.Show
Form7.Hide

End Sub

Private Sub Command5_Click()
Text3.Text = ""
Text4.Text = ""
Label36 = ""
Text1.Text = ""
Label37 = ""
Text8.Text = ""
Text9.Text = ""
Label38 = ""
Text11.Text = ""
Text12.Text = ""
Label39 = ""
Text14.Text = ""
Text15.Text = ""
Label40 = ""
Text17.Text = ""
Text18.Text = ""
Label41 = ""
Text20.Text = ""
Text21.Text = ""

Combo1 = ""
Combo2 = ""
Combo3 = ""
Combo4 = ""
Combo5 = ""
Combo6 = ""

Label31 = ""
Label32 = ""
Label30 = ""
Label35 = ""
End Sub

Private Sub Command6_Click()
Dim newid As Integer
Adodc1.Recordset.MoveLast
newid = Val(Adodc1.Recordset.Fields(0)) + 1
Adodc1.Recordset.AddNew
Label33.Caption = newid
End Sub

Private Sub Command8_Click()
Form8.Show

End Sub

Private Sub Command7_Click()
Form3.Show
Form7.Hide
End Sub

Private Sub Form_Load()
Combo1.AddItem "Klia Fresh"
Combo1.AddItem "Patanjali"
Combo1.AddItem "SurfExcel "
Combo1.AddItem "Ariel"
Combo1.AddItem "Nirma  "
Combo1.AddItem "Ghadi "
Combo1.AddItem "Wheel "
Combo1.AddItem "Tide"
Combo1.AddItem "Rin  "
Combo1.AddItem "Fena  "
Combo1.AddItem "Henko"
Combo1.AddItem "Sunlight Detergent Powder  "

Combo2.AddItem "Lifebuoy"
Combo2.AddItem "Cinthol"
Combo2.AddItem "Dettol"
Combo2.AddItem "Lux"
Combo2.AddItem "Liril"
Combo2.AddItem "Dove"
Combo2.AddItem "Pears"
Combo2.AddItem "Medimix"
Combo2.AddItem "Patanjali Soap"
Combo2.AddItem "Himalaya"
Combo2.AddItem "Godrej No1 Soap"
Combo2.AddItem "Venus Soap"
Combo2.AddItem "Santoor Soap"
Combo2.AddItem "Vivel Bar"


Combo3.AddItem "Amaranth"
Combo3.AddItem "Barley"
Combo3.AddItem "Buckwheat"
Combo3.AddItem "Bulgur"
Combo3.AddItem "Corn"
Combo3.AddItem "Einkorn"
Combo3.AddItem "Farro"
Combo3.AddItem "Freekeh"
Combo3.AddItem "Wheat"
Combo3.AddItem "Wild Rice"
Combo3.AddItem "Oats"
Combo3.AddItem "Quinoa "
Combo3.AddItem "Brown Rice "
Combo3.AddItem "Rye"

Combo4.AddItem "Colgate"
Combo4.AddItem "Close Up"
Combo4.AddItem "Pepsodent"
Combo4.AddItem "Patanjali Dant Kanti"
Combo4.AddItem "Meswak"
Combo4.AddItem "Dabur Red Paste"
Combo4.AddItem "Vicco Vajradanti"
Combo4.AddItem "Sensodyne"
Combo4.AddItem "Amway Glister"
Combo4.AddItem "Himalaya Herbals"
Combo4.AddItem "Lever Ayush"
Combo4.AddItem "Stomatol "
Combo4.AddItem "Zendium "
Combo4.AddItem "Ultra Brite"

Combo5.AddItem "Sesame Oil"
Combo5.AddItem "Mustard Oil"
Combo5.AddItem "Coconut Oil"
Combo5.AddItem "Groundnut Oil "
Combo5.AddItem "Sunflower Seeds Oil"
Combo5.AddItem "Soybean Oil"
Combo5.AddItem "Olive Oil"
Combo5.AddItem "Rice Bran Oil"
Combo5.AddItem "Safflower Oil"
Combo5.AddItem "Linseed Oil"
Combo5.AddItem "Corn Oil"
Combo5.AddItem "Cottonseed Oil"
Combo5.AddItem "Palm Oil"
Combo5.AddItem "Niger Seed Oil"

Combo6.AddItem "Others"

End Sub

Private Sub Text11_Change()
Dim a, b As Integer
a = Val(Text12.Text)
b = Val(Text11.Text)
Label38 = a * b
End Sub

Private Sub Text14_Change()
Dim a, b As Integer
a = Val(Text15.Text)
b = Val(Text14.Text)
Label39 = a * b
End Sub

Private Sub Text17_Change()
Dim a, b As Integer
a = Val(Text18.Text)
b = Val(Text17.Text)
Label40 = a * b
End Sub

Private Sub Text20_Change()
Dim a, b As Integer
a = Val(Text21.Text)
b = Val(Text20.Text)
Label41 = a * b
End Sub

Private Sub Text4_Change()
Dim a, b As Integer
a = Val(Text3.Text)
b = Val(Text4.Text)
Label36 = a * b
End Sub


Private Sub Text8_Change()
Dim a, b As Integer
a = Val(Text9.Text)
b = Val(Text8.Text)
Label37 = a * b
End Sub
