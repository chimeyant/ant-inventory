VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmgirolist 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Laporan Giro"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option3 
      Caption         =   "Giro Belum Cair"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Giro Tolak"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Giro Cair"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   975
   End
   Begin Crystal.CrystalReport Crystal 
      Left            =   120
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   1800
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmgirolist.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdview 
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   1800
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Preview"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmgirolist.frx":031A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   300
      Left            =   1560
      TabIndex        =   3
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   111017987
      CurrentDate     =   38679
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   300
      Left            =   1560
      TabIndex        =   4
      Top             =   1320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   111017987
      CurrentDate     =   38679
   End
   Begin VB.Label Label1 
      Caption         =   "S/D tanggal cair"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1350
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Dari tanggal cair"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   990
      Width           =   1335
   End
End
Attribute VB_Name = "frmgirolist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdview_Click()
    If date1 > Date Then
        MsgBox "From Date Greather To Date.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    Crystal.Reset
    Crystal.WindowState = crptMaximized
    Crystal.WindowShowCloseBtn = True
    Crystal.WindowShowPrintSetupBtn = True
    Crystal.WindowShowSearchBtn = True
    Crystal.WindowShowRefreshBtn = True
    Crystal.Connect = dsnreport
    Crystal.ReportFileName = AppPath & "\reports\finance\purc\girolist.rpt"
    Crystal.DataFiles(0) = "Proc(am_girolistsupp)"
    Crystal.ParameterFields(0) = "@tanggal1;" + Format(date1, "yyyyMMdd") + ";true"
    Crystal.ParameterFields(1) = "@tanggal2;" + Format(date2, "yyyyMMdd") + ";true"
    If Option1.Value = True Then Crystal.ParameterFields(2) = "@pilih;a;true"
    If Option2.Value = True Then Crystal.ParameterFields(2) = "@pilih;b;true"
    If Option3.Value = True Then Crystal.ParameterFields(2) = "@pilih;c;true"
    Crystal.ParameterFields(3) = "@namauser;" + nmuser + ";true"
    Crystal.RetrieveDataFiles
    Crystal.Action = 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    date1 = Date
    date2 = Date
End Sub

Private Sub Option1_Click()
    Label2 = "Dari tanggal cair"
    Label1 = "S/D tanggal cair"
End Sub

Private Sub Option2_Click()
    Label2 = "Dari tanggal tolak"
    Label1 = "S/D tanggal tolak"
End Sub

Private Sub Option3_Click()
    Label2 = "Dari tanggal J/T"
    Label1 = "S/D tanggal J/T"
End Sub
