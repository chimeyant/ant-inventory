VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmsrtibca 
   Caption         =   "SURAT INSTRUKSI BCA"
   ClientHeight    =   6585
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9675
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   9675
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   8640
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   5280
      Width           =   855
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   5760
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   5280
      Width           =   2775
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   495
      Left            =   8160
      TabIndex        =   15
      Top             =   5880
      Width           =   1335
      _Version        =   851970
      _ExtentX        =   2355
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Close"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   6720
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   6720
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2400
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2400
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   960
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   480
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   0
      Width           =   2775
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   2655
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   495
      Left            =   6720
      TabIndex        =   16
      Top             =   5880
      Width           =   1335
      _Version        =   851970
      _ExtentX        =   2355
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Print"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label Label7 
      Caption         =   "Ekuivalen"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "Jumlah"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Atas Nama"
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Nomor Giro Rupiah"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Untuk"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Kantor Cabang"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Tanggal Surat Pernyataan"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmsrtibca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
