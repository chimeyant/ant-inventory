VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmBA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Berita Acara"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9030
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtnobkt 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1065
      TabIndex        =   1
      Top             =   645
      Width           =   2565
   End
   Begin Crystal.CrystalReport Crystal 
      Left            =   195
      Top             =   2385
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox txtket 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   1365
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1125
      Width           =   7560
   End
   Begin VB.TextBox txtnama 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1065
      TabIndex        =   0
      Top             =   195
      Width           =   2565
   End
   Begin XtremeSuiteControls.PushButton btnclose 
      Height          =   435
      Left            =   8085
      TabIndex        =   4
      Top             =   2325
      Width           =   900
      _Version        =   851970
      _ExtentX        =   1587
      _ExtentY        =   767
      _StockProps     =   79
      Caption         =   "Close"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnview 
      Height          =   435
      Left            =   7140
      TabIndex        =   3
      Top             =   2325
      Width           =   900
      _Version        =   851970
      _ExtentX        =   1587
      _ExtentY        =   767
      _StockProps     =   79
      Caption         =   "View"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label Label5 
      Caption         =   "* Max. 4 Baris"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "1 2 3 4 "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   855
      Left            =   1170
      TabIndex        =   8
      Top             =   1170
      Width           =   150
   End
   Begin VB.Label Label3 
      Caption         =   "No. Bukti"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   7
      Top             =   690
      Width           =   900
   End
   Begin VB.Label Label2 
      Caption         =   "Keterangan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   135
      TabIndex        =   6
      Top             =   1155
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "Nama"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   135
      TabIndex        =   5
      Top             =   225
      Width           =   900
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1230
      Left            =   1065
      Top             =   1080
      Width           =   7905
   End
End
Attribute VB_Name = "frmBA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim baris As Integer

Private Sub btnclose_Click()
    Unload Me
End Sub

Private Sub btnview_Click()
    crystal.reset
    crystal.WindowState = crptMaximized
    crystal.WindowShowCloseBtn = True
    crystal.WindowShowPrintSetupBtn = True
    crystal.WindowShowSearchBtn = True
    crystal.Connect = dsnreport
    crystal.ReportFileName = App.Path & "\reports\Forms\berita_acaraLot.rpt"
    crystal.ParameterFields(0) = "@tgl;" & Format(Date, "yyyy/MM/dd") & ";true"
    crystal.ParameterFields(1) = "@nama;" & txtnama & ";true"
    crystal.ParameterFields(2) = "@ket;" & txtket & ";True"
    crystal.ParameterFields(3) = "@nobkt;" & txtnobkt & ";True"
    crystal.RetrieveDataFiles
    crystal.Action = 1
End Sub

Private Sub Form_Load()
    baris = 1
End Sub

Private Sub txtket_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        baris = baris + 1
        If baris > 4 Then MsgBox "Maximum 4 Baris !", vbExclamation, AppName
    End If
End Sub
