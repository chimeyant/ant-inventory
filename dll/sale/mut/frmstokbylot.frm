VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmstokbylot 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan Persediaan Barang By Lot"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   5310
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2280
      ScaleHeight     =   375
      ScaleWidth      =   2655
      TabIndex        =   12
      Top             =   1800
      Visible         =   0   'False
      Width           =   2655
      Begin MSComCtl2.DTPicker date1 
         Height          =   285
         Left            =   0
         TabIndex        =   13
         Top             =   0
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
         Format          =   134479875
         CurrentDate     =   37426
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "By. Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   11
      Top             =   1800
      Width           =   1095
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
      TabIndex        =   8
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtkode2 
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
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtkode1 
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
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   2400
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
      MICON           =   "frmstokbylot.frx":0000
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
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Chameleon.chameleonButton cmdview 
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   2400
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
      MICON           =   "frmstokbylot.frx":031A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdkode1 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "From Kode"
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
      MICON           =   "frmstokbylot.frx":0634
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdkode2 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "To Kode"
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
      MICON           =   "frmstokbylot.frx":094E
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
      TabIndex        =   9
      Top             =   1200
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
      MICON           =   "frmstokbylot.frx":0C68
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000A&
      X1              =   240
      X2              =   5160
      Y1              =   1080
      Y2              =   1080
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
      TabIndex        =   10
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   615
      Left            =   0
      Top             =   2280
      Width           =   5415
   End
   Begin VB.Label lblbarang2 
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
      TabIndex        =   7
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label lblbarang1 
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
      TabIndex        =   4
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "frmstokbylot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private Sub Check1_Click()
    If Check1.Value = Checked Then
        Picture1.Visible = True
    ElseIf Check1.Value = Unchecked Then
        Picture1.Visible = False
    End If
End Sub

Private Sub cmdClose_Click()
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

Private Sub cmdkode1_Click()
    carisql1 = "select kodebarang, namabarang from am_itemmst"
    namatabel = "Item"
    frmsearch.Show vbModal
End Sub

Private Sub cmdkode1_GotFocus()
    If hasil = "" Then Exit Sub
    txtkode1 = hasil
    lblbarang1 = hasil1
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdkode2_Click()
    carisql1 = "select kodebarang, namabarang from am_itemmst"
    namatabel = "Item"
    frmsearch.Show vbModal
End Sub

Private Sub cmdkode2_GotFocus()
    If hasil = "" Then Exit Sub
    txtkode2 = hasil
    lblbarang2 = hasil1
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdview_Click()
    Dim akses As Boolean
    If txtkode1 = "" Or txtkode2 = "" Then Exit Sub
    
    If txtkode2 < txtkode1 Then
        MsgBox "To Kode Can Not Smaller Then From Kode.", vbExclamation, "Warning"
        txtkode1.SetFocus
        Exit Sub
    End If
    
    If txtgudang = "" Then
        MsgBox "Gudang column cannot be empty", vbExclamation, AppName
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "Select * From LIST_USERS Where username = '" & nmuser & "'"
    Set RST = OBJ.Execute(SQL)
    
    If Not RST.EOF Then
        If RST!gl = "1" Then
            akses = True
        Else
            akses = False
            'If nmuser = "bina" Then akses = True
        End If
    Else
        If nmuser = "Creator" Then akses = True
    End If
    OBJ.Close

    Crystal.Reset
    Crystal.WindowState = crptMaximized
    Crystal.WindowShowCloseBtn = True
    Crystal.WindowShowPrintSetupBtn = True
    Crystal.WindowShowSearchBtn = True
    Crystal.WindowShowExportBtn = True
    Crystal.WindowShowRefreshBtn = True
    Crystal.Connect = dsnreport
    If Check1.Value = 0 Then
        Crystal.DataFiles(0) = "Proc(am_lapstok_bylot)"
        If akses = True Then Crystal.ReportFileName = AppPath & "\reports\sale\mut\daftarlothpp.rpt"
        If akses = False Then Crystal.ReportFileName = AppPath & "\reports\sale\mut\daftarlot.rpt"
    ElseIf Check1.Value = 1 Then
        Crystal.DataFiles(0) = "Proc(am_lapstok_bydate)"
        If akses = True Then Crystal.ReportFileName = AppPath & "\reports\sale\mut\persediaanbydate.rpt"
        If akses = False Then Crystal.ReportFileName = AppPath & "\reports\sale\mut\daftarlot.rpt"
    End If
    Crystal.ParameterFields(0) = "@kode1;" & txtkode1 & ";true"
    Crystal.ParameterFields(1) = "@kode2;" & txtkode2 & ";true"
    Crystal.ParameterFields(2) = "@gudang;" & txtgudang & ";true"
    Crystal.ParameterFields(3) = "@namauser;" & nmuser & ";true"
    
    If Check1.Value = 1 Then
        Crystal.ParameterFields(4) = "@tgl;" & Format(date1, "yyyy/MM/dd") & ";true"
    End If
    
    Crystal.RetrieveDataFiles
    Crystal.Action = 1
End Sub

Private Sub Form_Load()
    date1 = Date
End Sub
