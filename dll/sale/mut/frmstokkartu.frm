VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmstokkartu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kartu Stok"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5280
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   5280
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtkode 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   7
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtgudang 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   1920
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
      MICON           =   "frmstokkartu.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Crystal.CrystalReport Crystal 
      Left            =   120
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Chameleon.chameleonButton cmdview 
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   1920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "View"
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
      MICON           =   "frmstokkartu.frx":031A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdkode 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Kode"
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
      MICON           =   "frmstokkartu.frx":0634
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdgudang 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Gudang"
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
      MICON           =   "frmstokkartu.frx":094E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   315
      Left            =   1155
      TabIndex        =   2
      Top             =   1200
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
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
      Format          =   134807553
      CurrentDate     =   42039
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   315
      Left            =   2985
      TabIndex        =   3
      Top             =   1200
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
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
      Format          =   134807553
      CurrentDate     =   42039
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      TabIndex        =   11
      Top             =   1260
      Width           =   465
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2640
      TabIndex        =   10
      Top             =   1260
      Width           =   315
   End
   Begin VB.Label lblbarang 
      BackColor       =   &H80000014&
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
      Left            =   2280
      TabIndex        =   9
      Top             =   360
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   615
      Left            =   0
      Top             =   1800
      Width           =   5415
   End
   Begin VB.Label lblgudang 
      BackColor       =   &H80000014&
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
      Left            =   2280
      TabIndex        =   8
      Top             =   720
      Width           =   2895
   End
End
Attribute VB_Name = "frmstokkartu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdgudang_Click()
    carisql1 = "select kodegudang, namagudang from am_gudang"
    namatabel = "Gudang"
    frmsearch.Show vbModal
End Sub

Private Sub cmdgudang_GotFocus()
    If hasil = "" Then Exit Sub
    txtgudang = hasil
    lblgudang = hasil1
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdkode_Click()
    carisql1 = "select kodebarang, namabarang from am_itemmst"
    namatabel = "Item"
    frmsearch.Show vbModal
End Sub

Private Sub cmdkode_GotFocus()
    If hasil = "" Then Exit Sub
    txtkode = hasil
    lblbarang = hasil1
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdview_Click()
    Dim strakhir As String
    If txtkode = "" Then Exit Sub
    
    If date2 < date1 Then
            MsgBox "To Date Can Not Smaller Then From Date.", vbExclamation, "Warning"
            date2.SetFocus
            Exit Sub
        End If
    
    If txtgudang = "" Then
        MsgBox "Gudang column cannot be empty", vbExclamation, AppName
        Exit Sub
    End If
    strakhir = Right(date1, 4) & "-12-31"
    Crystal.Reset
    Crystal.WindowState = crptMaximized
    Crystal.WindowShowCloseBtn = True
    Crystal.WindowShowPrintSetupBtn = True
    Crystal.WindowShowSearchBtn = True
    Crystal.WindowShowExportBtn = True
    Crystal.WindowShowRefreshBtn = True
    Crystal.Connect = dsnreport
    Crystal.DataFiles(0) = "Proc(am_lapstok_posisi)"
    Crystal.ReportFileName = AppPath & "\reports\sale\mut\daftarlot_detail.rpt"
    Crystal.ParameterFields(0) = "@kode;" & txtkode & ";true"
    Crystal.ParameterFields(1) = "@tgl1 ;" + Format(date1, "yyyymmdd") + ";true"
    Crystal.ParameterFields(2) = "@tgl2 ;" + Format(date2, "yyyymmdd") + ";true"
    Crystal.ParameterFields(3) = "@gudang;" & txtgudang & ";true"
    Crystal.ParameterFields(4) = "@namauser;" + nmuser + ";true"
    Crystal.ParameterFields(5) = "@thn;" + strakhir + ";true"
    Crystal.RetrieveDataFiles
    Crystal.Action = 1
    
End Sub

Private Sub Form_Load()
    date1 = Date
    date2 = Date
End Sub
