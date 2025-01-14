VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmterimalist 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List Mutasi Barang"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmterimalist.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option15 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Daftar Terima Tuangan Lem"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   3480
      Width           =   2415
   End
   Begin VB.OptionButton Option14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Daftar Keluar Barang Bocor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   3240
      Width           =   2295
   End
   Begin VB.OptionButton Option13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Daftar Kembali Pinjaman dr Customer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   3000
      Width           =   3015
   End
   Begin VB.OptionButton Option12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Daftar Terima Pinjaman dr Customer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2760
      Width           =   3015
   End
   Begin VB.OptionButton Option11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Daftar Retur Ke Pabrik"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2520
      Width           =   1935
   End
   Begin VB.OptionButton Option10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Daftar Terima Dari Pabrik "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2280
      Width           =   2175
   End
   Begin VB.OptionButton Option9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Daftar Produksi Harian Karet"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.OptionButton Option8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Daftar Barang Bonus"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   1815
   End
   Begin VB.OptionButton Option7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Daftar Keluar Sampel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   1935
   End
   Begin VB.OptionButton Option6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Daftar Barang Rusak"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Daftar Produksi Harian Lem"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   2295
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Daftar Terima Over Zak"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Daftar Keluar Over Zak"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Daftar Terima Retur Customer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   2535
   End
   Begin VB.OptionButton Option5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Daftar Dari Bahan Baku"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1440
      TabIndex        =   15
      Top             =   4080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
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
      Format          =   103677955
      CurrentDate     =   37464
   End
   Begin Crystal.CrystalReport crystal 
      Left            =   0
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   1440
      TabIndex        =   16
      Top             =   4440
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
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
      Format          =   103677955
      CurrentDate     =   37464
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   2400
      TabIndex        =   18
      Top             =   4920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
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
      MICON           =   "frmterimalist.frx":2372
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
      Left            =   1440
      TabIndex        =   17
      Top             =   4920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
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
      MICON           =   "frmterimalist.frx":268C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   3855
      Left            =   -240
      TabIndex        =   21
      Top             =   0
      Width           =   3735
   End
   Begin VB.Label Label4 
      Caption         =   "To Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   4470
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "From Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   4110
      Width           =   1215
   End
End
Attribute VB_Name = "frmterimalist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim str1 As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdview_Click()
    If date2 < date1 Then
        MsgBox "To Date Can Not Smaller Then From Date.", vbExclamation, "Warning"
        date2.SetFocus
        Exit Sub
    End If
    
    If Option1.Value = True Then str1 = "01a"
    If Option9.Value = True Then str1 = "01b"
    If Option2.Value = True Then str1 = "02"
    If Option3.Value = True Then str1 = "03"
    If Option4.Value = True Then str1 = "04"
    If Option6.Value = True Then str1 = "05"
    If Option7.Value = True Then str1 = "06"
    If Option8.Value = True Then str1 = "07"
    If Option5.Value = True Then str1 = "08"
    If Option10.Value = True Then str1 = "09"
    If Option11.Value = True Then str1 = "10"
    If Option12.Value = True Then str1 = "11"
    If Option13.Value = True Then str1 = "12"
    If Option14.Value = True Then str1 = "13"
    If Option15.Value = True Then str1 = "14"
    
    Crystal.Reset
    Crystal.WindowState = crptMaximized
    Crystal.WindowShowCloseBtn = True
    Crystal.WindowShowPrintSetupBtn = True
    Crystal.WindowShowSearchBtn = True
    Crystal.Connect = dsnreport
    Crystal.DataFiles(0) = "Proc(am_daftarbpb1)"
    Crystal.ReportFileName = AppPath & "\reports\sale\mut\daftarbpb.rpt"
    Crystal.ParameterFields(0) = "@tanggal1;" & Format(date1, "yyyyMMdd") & ";true"
    Crystal.ParameterFields(1) = "@tanggal2;" & Format(date2, "yyyyMMdd") & ";true"
    Crystal.ParameterFields(2) = "@namauser ;" + nmuser + ";true"
    Crystal.ParameterFields(3) = "@pilih ;" + str1 + ";true"
    Crystal.RetrieveDataFiles
    Crystal.Action = 1
End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='105' and b.kodeuser = '1" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If RST.EOF Then
    '        MsgBox "User Rights Denied !!" & vbCrLf & _
   '         "Please contact your Administrator.", vbCritical, "User Rights"
   '         OBJ.Close
   '         Unload Me
    '        Exit Sub
    '    End If
    '    OBJ.Close
    'End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
   
    
    date1 = Date
    date2 = Date
End Sub
