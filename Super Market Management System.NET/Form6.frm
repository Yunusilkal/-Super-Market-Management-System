VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form6 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Customer Details"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13665
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form6.frx":0000
   ScaleHeight     =   12968.47
   ScaleMode       =   0  'User
   ScaleWidth      =   13665
   StartUpPosition =   3  'Windows Default
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Customer Details"
      Height          =   10215
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   13695
      Begin VB.TextBox Text1 
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
         Height          =   375
         Left            =   1920
         TabIndex        =   44
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0000FF00&
         Caption         =   "Find"
         BeginProperty Font 
            Name            =   "Castellar"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Display in Grid"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   8520
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H000080FF&
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   8520
         Width           =   2295
      End
      Begin VB.Line Line14 
         BorderWidth     =   2
         X1              =   7320
         X2              =   7320
         Y1              =   7560
         Y2              =   9720
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "Billed-Date"
         BeginProperty Font 
            Name            =   "Perpetua"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9360
         TabIndex        =   47
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         DataField       =   "DATE"
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
         Left            =   11280
         TabIndex        =   45
         Top             =   120
         Width           =   1935
      End
      Begin VB.Line Line13 
         BorderWidth     =   2
         X1              =   11640
         X2              =   1680
         Y1              =   9720
         Y2              =   9720
      End
      Begin VB.Line Line11 
         BorderWidth     =   2
         X1              =   7320
         X2              =   11640
         Y1              =   9000
         Y2              =   9000
      End
      Begin VB.Line Line12 
         BorderWidth     =   2
         X1              =   1680
         X2              =   11640
         Y1              =   8280
         Y2              =   8280
      End
      Begin VB.Label Label53 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         DataField       =   "INC_TAX"
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   9840
         TabIndex        =   40
         Top             =   9120
         Width           =   1575
      End
      Begin VB.Label Label52 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         DataField       =   "TAX"
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   9480
         TabIndex        =   39
         Top             =   8400
         Width           =   1935
      End
      Begin VB.Label Label51 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         DataField       =   "AMOUNT"
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   9480
         TabIndex        =   38
         Top             =   7680
         Width           =   1935
      End
      Begin VB.Label Label50 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   5640
         TabIndex        =   37
         Top             =   7680
         Width           =   1575
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Total_QTY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   36
         Top             =   7680
         Width           =   1575
      End
      Begin VB.Label Label48 
         BackColor       =   &H0000FFFF&
         Caption         =   "Including_Tax"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7560
         TabIndex        =   35
         Top             =   9120
         Width           =   2055
      End
      Begin VB.Label Label47 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7560
         TabIndex        =   34
         Top             =   7680
         Width           =   1575
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "TAX"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7560
         TabIndex        =   33
         Top             =   8400
         Width           =   1575
      End
      Begin VB.Line Line10 
         BorderWidth     =   2
         X1              =   0
         X2              =   13680
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line9 
         BorderWidth     =   2
         X1              =   1680
         X2              =   11640
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Line Line8 
         BorderWidth     =   2
         X1              =   11640
         X2              =   11640
         Y1              =   2280
         Y2              =   9720
      End
      Begin VB.Line Line7 
         BorderWidth     =   2
         X1              =   9360
         X2              =   9360
         Y1              =   2280
         Y2              =   9000
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         X1              =   7560
         X2              =   7560
         Y1              =   2280
         Y2              =   7560
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   1680
         X2              =   11640
         Y1              =   7560
         Y2              =   7560
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   5640
         X2              =   5640
         Y1              =   2280
         Y2              =   7560
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   1680
         X2              =   1680
         Y1              =   2280
         Y2              =   9720
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         DataField       =   "MISCELLANEOUS"
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1920
         TabIndex        =   32
         Top             =   6960
         Width           =   3735
      End
      Begin VB.Label Label44 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         DataField       =   "COOKING OIL"
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1920
         TabIndex        =   31
         Top             =   6240
         Width           =   3735
      End
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         DataField       =   "TOOTH_PASTE"
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1920
         TabIndex        =   30
         Top             =   5520
         Width           =   3735
      End
      Begin VB.Label Label40 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         DataField       =   "GRAINS"
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1920
         TabIndex        =   29
         Top             =   4800
         Width           =   3735
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         DataField       =   "SOAP"
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1920
         TabIndex        =   28
         Top             =   4080
         Width           =   3735
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         DataField       =   "DETERGENT"
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1920
         TabIndex        =   27
         Top             =   3360
         Width           =   3735
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   9720
         TabIndex        =   26
         Top             =   6240
         Width           =   1335
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   9720
         TabIndex        =   25
         Top             =   6960
         Width           =   1335
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         DataField       =   "S_QTY"
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   9720
         TabIndex        =   24
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   9720
         TabIndex        =   23
         Top             =   4800
         Width           =   1335
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   9720
         TabIndex        =   22
         Top             =   5520
         Width           =   1335
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   9720
         TabIndex        =   21
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         DataField       =   "C_PRICE"
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   7560
         TabIndex        =   20
         Top             =   6240
         Width           =   2175
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         DataField       =   "M_PRICE"
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   7560
         TabIndex        =   19
         Top             =   6960
         Width           =   2175
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         DataField       =   "D_PRICE"
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   7560
         TabIndex        =   18
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         DataField       =   "S_PRICE"
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   7560
         TabIndex        =   17
         Top             =   4080
         Width           =   2175
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         DataField       =   "G_PRICE"
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   7560
         TabIndex        =   16
         Top             =   4800
         Width           =   2175
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         DataField       =   "T_PRICE"
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   7560
         TabIndex        =   15
         Top             =   5520
         Width           =   2175
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         DataField       =   "C_QTY"
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   5640
         TabIndex        =   14
         Top             =   6240
         Width           =   1935
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         DataField       =   "M_QTY"
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   5640
         TabIndex        =   13
         Top             =   6960
         Width           =   1935
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         DataField       =   "S_QTY"
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   5640
         TabIndex        =   12
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         DataField       =   "G_QTY"
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   5640
         TabIndex        =   11
         Top             =   4800
         Width           =   1935
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         DataField       =   "T_QTY"
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   5640
         TabIndex        =   10
         Top             =   5520
         Width           =   1935
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   1680
         X2              =   11640
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   13680
         X2              =   17160
         Y1              =   2400
         Y2              =   2280
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   "PRICE"
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
         Height          =   495
         Left            =   7800
         TabIndex        =   9
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "TOTAL"
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
         Height          =   495
         Left            =   9480
         TabIndex        =   8
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         DataField       =   "D_QTY"
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   5640
         TabIndex        =   7
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "PRODUCT"
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
         Left            =   1920
         TabIndex        =   6
         Top             =   2520
         Width           =   3255
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "QTY"
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
         Height          =   495
         Left            =   6000
         TabIndex        =   5
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label5 
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
         Height          =   375
         Left            =   2760
         TabIndex        =   4
         Top             =   840
         Width           =   4455
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
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
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "Bill_NO"
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
         Left            =   360
         TabIndex        =   2
         Top             =   120
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   10695
         Left            =   0
         Picture         =   "Form6.frx":1BC6B
         Stretch         =   -1  'True
         Top             =   0
         Width           =   13890
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   10320
      Top             =   1560
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin VB.Image Image4 
      Height          =   375
      Left            =   0
      Picture         =   "Form6.frx":378D6
      Stretch         =   -1  'True
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      DataField       =   "DETERGENT"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   14760
      TabIndex        =   46
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "CUSTOMER DETAILS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13695
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If DataEnvironment1.rsCommand1.State = 1 Then DataEnvironment1.rsCommand1.Close
DataEnvironment1.rsCommand1.Filter = "BILL_NO = " & Val(Text1.Text)
DataReport1.Show
End Sub

Private Sub Command2_Click()
Form4.Show
Form6.Hide
End Sub

Private Sub Command3_Click()
Adodc1.RecordSource = "select * from CUSTOMER where BILL_NO like'" & Text1.Text & "' "
Adodc1.Refresh
If Text1.Text = "" Then


MsgBox "Plese Enter The Bill_no Properlly", vbCritical
End If
End Sub

Private Sub Command7_Click()
 Form3.Show
Form6.Hide
End Sub
