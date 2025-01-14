VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmpackadd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Confirm Request"
   ClientHeight    =   5985
   ClientLeft      =   -15
   ClientTop       =   330
   ClientWidth     =   12105
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   12105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picblok 
      BackColor       =   &H000000FF&
      Height          =   5295
      Left            =   120
      ScaleHeight     =   5235
      ScaleWidth      =   2595
      TabIndex        =   22
      Top             =   600
      Width           =   2655
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "EDIT COMFIRM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   0
         TabIndex        =   23
         Top             =   2280
         Width           =   2655
      End
   End
   Begin VB.ComboBox cmbstatus 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10680
      TabIndex        =   21
      Text            =   "Combo1"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Edit (cetak ke)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6960
      TabIndex        =   19
      Top             =   960
      Width           =   1335
   End
   Begin VB.ComboBox cmbkode 
      Height          =   315
      Left            =   8400
      TabIndex        =   17
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   2760
      Top             =   120
   End
   Begin TDBNumber6Ctl.TDBNumber txthasil 
      Height          =   255
      Left            =   9120
      TabIndex        =   13
      Top             =   4920
      Visible         =   0   'False
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   450
      Calculator      =   "frmpackadd.frx":0000
      Caption         =   "frmpackadd.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpackadd.frx":008C
      Keys            =   "frmpackadd.frx":00AA
      Spin            =   "frmpackadd.frx":00EC
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00)"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,###,##0.00;(###,###,###,##0.00)"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999999999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   1
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.PictureBox blank 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9720
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   12
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox check 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10200
      Picture         =   "frmpackadd.frx":0114
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox uncheck 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9960
      Picture         =   "frmpackadd.frx":03F6
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   10
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtpetugas 
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
      Left            =   4200
      Locked          =   -1  'True
      MaxLength       =   17
      TabIndex        =   8
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox txtnolot 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      MaxLength       =   17
      TabIndex        =   3
      Top             =   840
      Width           =   2535
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   5340
      ItemData        =   "frmpackadd.frx":0744
      Left            =   120
      List            =   "frmpackadd.frx":0746
      Sorted          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Klik Nomor Lot untuk confirm"
      Top             =   600
      Width           =   2655
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   11040
      TabIndex        =   0
      Top             =   5400
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
      MICON           =   "frmpackadd.frx":0748
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   3615
      Left            =   3120
      TabIndex        =   4
      Top             =   1710
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   6376
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColorBkg    =   -2147483642
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
   End
   Begin Chameleon.chameleonButton cmdsave 
      Height          =   375
      Left            =   10080
      TabIndex        =   14
      Top             =   5400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Confirm"
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
      MICON           =   "frmpackadd.frx":0A62
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdclear 
      Height          =   375
      Left            =   9120
      TabIndex        =   15
      Top             =   5400
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Clear"
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
      MICON           =   "frmpackadd.frx":0D7C
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
      Height          =   315
      Left            =   9000
      TabIndex        =   16
      Top             =   1320
      Visible         =   0   'False
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
      Format          =   125370369
      CurrentDate     =   42039
   End
   Begin Chameleon.chameleonButton cmdupdate 
      Height          =   375
      Left            =   8160
      TabIndex        =   18
      Top             =   5400
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Update"
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
      MICON           =   "frmpackadd.frx":1096
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
      Left            =   6960
      TabIndex        =   20
      Top             =   1200
      Visible         =   0   'False
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
      MICON           =   "frmpackadd.frx":13B0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdchange 
      Height          =   375
      Left            =   3120
      TabIndex        =   24
      Top             =   5400
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Change Item Code"
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
      MICON           =   "frmpackadd.frx":16CA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Petugas"
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
      Left            =   3240
      TabIndex        =   9
      Top             =   1230
      Width           =   1335
   End
   Begin VB.Label lblTgl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
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
      Left            =   9360
      TabIndex        =   7
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "No. LOT "
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
      Left            =   3240
      TabIndex        =   6
      Top             =   870
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Confirm"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   240
      Width           =   9015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nomor LOT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   2655
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H80000014&
      FillStyle       =   0  'Solid
      Height          =   5175
      Left            =   3000
      Top             =   720
      Width           =   9015
   End
End
Attribute VB_Name = "frmpackadd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private Sub Check1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Check1.Value = Unchecked Then
        cmbkode.Visible = True
        cmdview.Visible = True
        cmdupdate.Visible = True
        cmdsave.Enabled = False
        picblok.Visible = True
        LoadCetak
    End If
End Sub

Private Sub cmdchange_Click()
    frmgantikode.Show vbModal
End Sub

Private Sub cmdclear_Click()
    hapusgrid
    txtnolot = ""
    txtpetugas = ""
    lblTgl = ""
    date1 = Date
    cmdview.Visible = False
    cmbkode.Visible = False
    cmdupdate.Visible = False
    cmdsave.Enabled = True
    picblok.Visible = False
    Check1.Value = Unchecked
    List1.Clear
    OBJ.Open dsn
    SQL = "SELECT distinct nolot FROM am_gudang_permintaan WHERE status = '0'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        List1.AddItem RST!nolot
        RST.MoveNext
    Loop
    OBJ.Close
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmbstatus_Click()
    grid.TextMatrix(grid.Row, 5) = cmbstatus
    cmbstatus.Visible = False
End Sub

Private Sub cmbstatus_LostFocus()
    cmbstatus.Visible = False
End Sub
Private Sub LoadCetak()
    cmbkode.Clear
    OBJ.Open dsn
    SQL = "Select distinct cetak_ke From am_gudang_permintaan Where nolot = '" & txtnolot & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        Do While Not RST.EOF
            cmbkode.AddItem RST!cetak_ke
            
            RST.MoveNext
        Loop
    End If
    OBJ.Close
End Sub

Private Sub cmdsave_Click()
    Dim request, used As Double
    If txtnolot = "" Then Exit Sub
    If grid.TextMatrix(1, 1) = "" Then
        MsgBox "There is no items on grid", vbExclamation, AppName
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "Select * From am_gudang_permintaan Where nolot = '" & txtnolot & "'"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        SQL = "Update am_gudang_permintaan "
        'TAMBAH ITEM
        If grid.TextMatrix(grid.Row, 5) = "Tambah Item" Or grid.TextMatrix(grid.Row, 5) = "Request" Then
            SQL = SQL + "set status = '1',flag = '1',"
            SQL = SQL + " qty_confirmed=convert(money,'" & Format(grid.TextMatrix(grid.Row, 4), "general number") & "')*-1,"
            SQL = SQL + " user_confirmed='" & nmuser & "',"
            SQL = SQL + " tgl_confirmed=convert(datetime,'" & tanggalconfirm & "')"
        'RUSAK
        ElseIf grid.TextMatrix(grid.Row, 5) = "Rusak" Or grid.TextMatrix(grid.Row, 5) = "Rusak " Then
            SQL = SQL + "set status = '5',flag = '1',"
            SQL = SQL + " qty_confirmed=convert(money,'" & Format(grid.TextMatrix(grid.Row, 4), "general number") & "'),"
            SQL = SQL + " user_confirmed='" & nmuser & "',"
            SQL = SQL + " tgl_confirmed=convert(datetime,'" & tanggalconfirm & "')"
        'HILANG
        ElseIf grid.TextMatrix(grid.Row, 5) = "Hilang" Then
            SQL = SQL + "set status = '6',flag = '1',"
            SQL = SQL + " qty_confirmed=convert(money,'" & Format(grid.TextMatrix(grid.Row, 4), "general number") & "'),"
            SQL = SQL + " user_confirmed='" & nmuser & "',"
            SQL = SQL + " tgl_confirmed=convert(datetime,'" & tanggalconfirm & "')"
        'RETUR
        ElseIf grid.TextMatrix(grid.Row, 5) = "Return Gudang" Then
            SQL = SQL + "set status = '7',flag = '1',"
            SQL = SQL + " qty_confirmed=convert(money,'" & Format(grid.TextMatrix(grid.Row, 4), "general number") & "'),"
            SQL = SQL + " user_confirmed='" & nmuser & "',"
            SQL = SQL + " tgl_confirmed=convert(datetime,'" & tanggalconfirm & "')"
        End If

        SQL = SQL + " Where nolot = '" & txtnolot & "'"
        SQL = SQL + " and kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' And keterangan='" & grid.TextMatrix(grid.Row, 5) & "' and flag='0'"
        Set RST = OBJ.Execute(SQL)
        grid.Row = grid.Row + 1
    Loop
    
    SQL = "Select (SUM(qty) + SUM(qty_add))'ambil' From am_gudang_permintaan"
    SQL = SQL + " Where nolot = '" & txtnolot & "' and flag = '1' group by nolot"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        request = RST!ambil
    End If
    
    SQL = "Select SUM(qty_bahan)'pakai' From list_produksi_kemasan"
    SQL = SQL + " Where nolot = '" & txtnolot & "' group by nolot"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        used = RST!pakai
    End If
    
    If request = used Then
        SQL = "Insert into am_gudang_perclose(nolot,tgl)"
        SQL = SQL + " Values('" & txtnolot & "',convert(datetime,'" & tanggalconfirm & "'))"
        Set RST = OBJ.Execute(SQL)
        
        'Close Permintaan
        SQL = "update am_gudang_permintaan set flag = '2' Where nolot = '" & txtnolot & "'"
        Set RST = OBJ.Execute(SQL)
    End If
    
    OBJ.Close
    MsgBox "Berhasil Disimpan", vbInformation, AppName
    cmdclear_Click
End Sub

Private Sub cmdUpdate_Click()
    Dim request, used As Double
    If txtnolot = "" Then Exit Sub
    If grid.TextMatrix(1, 1) = "" Then
        MsgBox "There is no items on grid", vbExclamation, AppName
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "Select * From am_gudang_permintaan Where nolot = '" & txtnolot & "'"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        SQL = "Update am_gudang_permintaan "
        'TAMBAH ITEM
        If grid.TextMatrix(grid.Row, 5) = "Tambah Item" Or grid.TextMatrix(grid.Row, 5) = "Request" Then
            SQL = SQL + "set status = '1',flag = '1',"
            SQL = SQL + " qty_confirmed=convert(money,'" & Format(grid.TextMatrix(grid.Row, 4), "general number") & "')*-1,"
            SQL = SQL + " user_confirmed='" & nmuser & "',"
            SQL = SQL + " tgl_confirmed=convert(datetime,'" & tanggalconfirm & "'),"
            SQL = SQL + " keterangan='" & grid.TextMatrix(grid.Row, 5) & "'"
        'RUSAK
        ElseIf grid.TextMatrix(grid.Row, 5) = "Rusak" Or grid.TextMatrix(grid.Row, 5) = "Rusak " Then
            SQL = SQL + "set status = '5',flag = '1',"
            SQL = SQL + " qty_confirmed=convert(money,'" & Format(grid.TextMatrix(grid.Row, 4), "general number") & "'),"
            SQL = SQL + " user_confirmed='" & nmuser & "',"
            SQL = SQL + " tgl_confirmed=convert(datetime,'" & tanggalconfirm & "'),"
            SQL = SQL + " keterangan='" & grid.TextMatrix(grid.Row, 5) & "'"
        'HILANG
        ElseIf grid.TextMatrix(grid.Row, 5) = "Hilang" Then
            SQL = SQL + "set status = '6',flag = '1',"
            SQL = SQL + " qty_confirmed=convert(money,'" & Format(grid.TextMatrix(grid.Row, 4), "general number") & "'),"
            SQL = SQL + " user_confirmed='" & nmuser & "',"
            SQL = SQL + " tgl_confirmed=convert(datetime,'" & tanggalconfirm & "'),"
            SQL = SQL + " keterangan='" & grid.TextMatrix(grid.Row, 5) & "'"
        'RETUR
        ElseIf grid.TextMatrix(grid.Row, 5) = "Return Gudang" Then
            SQL = SQL + "set status = '7',flag = '1',"
            SQL = SQL + " qty_confirmed=convert(money,'" & Format(grid.TextMatrix(grid.Row, 4), "general number") & "'),"
            SQL = SQL + " user_confirmed='" & nmuser & "',"
            SQL = SQL + " tgl_confirmed=convert(datetime,'" & tanggalconfirm & "'),"
            SQL = SQL + " keterangan='" & grid.TextMatrix(grid.Row, 5) & "'"
        End If
        
        If grid.TextMatrix(grid.Row, 5) = "Request" Then
            SQL = SQL + " Where nolot = '" & txtnolot & "'"
            SQL = SQL + " and kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and cetak_ke='" & cmbkode.text & "'"
            SQL = SQL + " and qty = '" & grid.TextMatrix(grid.Row, 3) & "'"
        Else
            SQL = SQL + " Where nolot = '" & txtnolot & "'"
            SQL = SQL + " and kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and cetak_ke='" & cmbkode.text & "'"
            SQL = SQL + " and qty_add = '" & grid.TextMatrix(grid.Row, 3) & "'"
            SQL = SQL + " and keterangan='" & grid.TextMatrix(grid.Row, 5) & "'"
        End If
        Set RST = OBJ.Execute(SQL)
        grid.Row = grid.Row + 1
    Loop
    
    SQL = "Select (SUM(qty) + SUM(qty_add))'ambil' From am_gudang_permintaan"
    SQL = SQL + " Where nolot = '" & txtnolot & "' and flag = '1' group by nolot"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        request = RST!ambil
    End If
    
    SQL = "Select SUM(qty_bahan)'pakai' From list_produksi_kemasan"
    SQL = SQL + " Where nolot = '" & txtnolot & "' group by nolot"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        used = RST!pakai
    End If
    
    If request = used Then
        SQL = "Insert into am_gudang_perclose(nolot,tgl)"
        SQL = SQL + " Values('" & txtnolot & "',convert(datetime,'" & tanggalconfirm & "'))"
        Set RST = OBJ.Execute(SQL)
        
        'Close Permintaan
        SQL = "update am_gudang_permintaan set flag = '2' Where nolot = '" & txtnolot & "'"
        Set RST = OBJ.Execute(SQL)
    End If
    
    OBJ.Close
    MsgBox "Berhasil Disimpan", vbInformation, AppName
    cmdclear_Click
End Sub

Private Sub cmdview_Click()
    Dim j As Integer
    j = 0
    
    If txtnolot = "" Then Exit Sub
    If Check1.Value = Unchecked Then Exit Sub
    If cmbkode.text = "" Then Exit Sub
        
    OBJ.Open dsn
    SQL = "Select a.*,b.NamaBarang,c.NamaSatuan From am_gudang_permintaan a"
    SQL = SQL + " inner join am_apitemmst b on a.kodebarang = b.KodeBarang"
    SQL = SQL + " inner join am_apunit c on b.KodeSatuan = c.KodeSatuan"
    SQL = SQL + " Where a.nolot = '" & txtnolot & "' and cetak_ke = '" & cmbkode & "'"
    Set RST = OBJ.Execute(SQL)
        
    If RST.EOF Then
        OBJ.Close
        MsgBox "No Lot not found", vbCritical, AppName
        Exit Sub
    End If
        
    hapusgrid
    lblTgl = "Tanggal : " & Format(RST!tgl, "dd MMMM yyyy")
    txtpetugas = RST!petugas
    Do Until RST.EOF
        grid.Col = 0
        Set grid.CellPicture = uncheck
        grid.TextMatrix(grid.Row, 1) = RST!kodebarang
        grid.TextMatrix(grid.Row, 2) = RST!namabarang
        If RST!keterangan = "Request" Then
            grid.TextMatrix(grid.Row, 3) = Format(RST!qty, "###,###,##0.00")
        Else
            grid.TextMatrix(grid.Row, 3) = Format(RST!qty_add, "###,###,##0.00")
        End If
        If RST!qty_confirmed < 0 Then
            grid.TextMatrix(grid.Row, 4) = Format(RST!qty_confirmed * -1, "###,###,##0.00")
        Else
            grid.TextMatrix(grid.Row, 4) = Format(RST!qty_confirmed, "###,###,##0.00")
        End If
        grid.TextMatrix(grid.Row, 5) = RST!keterangan
        If RST!keterangan = "Tambah Item" Then
            For j = 0 To grid.Cols - 1
                grid.Col = j
                grid.CellBackColor = &HE0E0E0
            Next
        Else
            For j = 0 To grid.Cols - 1
                grid.Col = j
                grid.CellBackColor = &HFFFFFF
            Next
        End If
        grid.Rows = grid.Rows + 1
        grid.Row = grid.Row + 1
        RST.MoveNext
    Loop
    OBJ.Close
End Sub

Private Sub Form_Activate()
    Timer1.Enabled = True
End Sub

Private Sub Form_Load()
    grid.TextMatrix(0, 0) = "X"
    grid.TextMatrix(0, 1) = "Kode"
    grid.TextMatrix(0, 2) = "Nama Barang"
    grid.TextMatrix(0, 3) = "Produksi"
    grid.TextMatrix(0, 4) = "Gudang"
    grid.TextMatrix(0, 5) = "Keterangan"
    grid.ColWidth(0) = 300
    grid.ColWidth(1) = 1100
    grid.ColWidth(2) = 3000
    grid.ColWidth(3) = 1000
    grid.ColWidth(4) = 1000
    grid.ColWidth(5) = 1600
    
    picblok.Visible = False
    List1.Clear
    OBJ.Open dsn
    SQL = "SELECT distinct nolot FROM am_gudang_permintaan WHERE status <> '1' and flag = '0'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        List1.AddItem RST!nolot
        RST.MoveNext
    Loop
    OBJ.Close
    cmbstatus.AddItem "Request"
    cmbstatus.AddItem "Tambah Item"
    cmbstatus.AddItem "Rusak"
    cmbstatus.AddItem "Hilang"
    cmbstatus.AddItem "Return Gudang"
    date1 = Date
End Sub

Private Sub hapusgrid()
    Dim j As Integer
    j = 0
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        grid.TextMatrix(grid.Row, 0) = ""
        grid.TextMatrix(grid.Row, 1) = ""
        grid.TextMatrix(grid.Row, 2) = ""
        grid.TextMatrix(grid.Row, 3) = ""
        grid.TextMatrix(grid.Row, 4) = ""
        grid.TextMatrix(grid.Row, 5) = ""
        grid.Col = 0
        Set grid.CellPicture = blank
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
    grid.ColWidth(0) = 300
    grid.ColWidth(1) = 1100
    grid.ColWidth(2) = 3000
    grid.ColWidth(3) = 1000
    grid.ColWidth(4) = 1000
    grid.ColWidth(5) = 1600
    For j = 0 To grid.Cols - 1
        grid.Col = j
        grid.CellBackColor = &HFFFFFF
    Next
End Sub

Private Sub hapusrow()
    grid.TextMatrix(grid.Row, 1) = ""
    grid.TextMatrix(grid.Row, 2) = ""
    grid.TextMatrix(grid.Row, 3) = ""
    grid.TextMatrix(grid.Row, 4) = ""
    grid.TextMatrix(grid.Row, 4) = ""
    Do While True
        If grid.TextMatrix(grid.Row + 1, 1) = "" Then
            grid.TextMatrix(grid.Row, 1) = ""
            grid.TextMatrix(grid.Row, 2) = ""
            grid.TextMatrix(grid.Row, 3) = ""
            grid.TextMatrix(grid.Row, 4) = ""
            grid.TextMatrix(grid.Row, 4) = ""
            Exit Do
        End If
        grid.TextMatrix(grid.Row, 1) = grid.TextMatrix(grid.Row + 1, 1)
        grid.TextMatrix(grid.Row, 2) = grid.TextMatrix(grid.Row + 1, 2)
        grid.TextMatrix(grid.Row, 3) = grid.TextMatrix(grid.Row + 1, 3)
        grid.TextMatrix(grid.Row, 4) = grid.TextMatrix(grid.Row + 1, 4)
        grid.TextMatrix(grid.Row, 5) = grid.TextMatrix(grid.Row + 1, 5)
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = grid.Rows - 1
    grid.Col = 0
    Set grid.CellPicture = blank
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    Select Case grid.Col
        Case 0:
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
                If grid.CellPicture = uncheck Then
                    Set grid.CellPicture = check
                'If MsgBox("Delete that Row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                '    Set grid.CellPicture = uncheck
                '    hapusrow
                '    Exit Sub
                'End If
                'Set grid.CellPicture = uncheck
                Else
                    Set grid.CellPicture = uncheck
                End If
        Case 4:
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            txthasil.Width = grid.ColWidth(grid.Col) - 40
            txthasil = grid.TextMatrix(grid.Row, grid.Col)
            txthasil.Left = grid.Left + grid.CellLeft
            txthasil.Top = grid.Top + grid.CellTop + 20
            txthasil.Visible = True
            txthasil.SetFocus
        Case 5:
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            If grid.TextMatrix(grid.Row, 5) = "Request" Then Exit Sub
            'If grid.TextMatrix(grid.Row, 8) = "1" Then Exit Sub
            cmbstatus.Width = grid.ColWidth(grid.Col) - 40
            cmbstatus = grid.TextMatrix(grid.Row, grid.Col)
            cmbstatus.Left = grid.Left + grid.CellLeft
            cmbstatus.Top = grid.Top + grid.CellTop + 20
            cmbstatus.Visible = True
            cmbstatus.SetFocus
    End Select
End Sub

Private Sub grid_EnterCell()
    If grid.MouseRow = 0 Then Exit Sub
    Select Case grid.Col
        Case 0:
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
                If grid.CellPicture = uncheck Then
                    Set grid.CellPicture = check
                'If MsgBox("Delete that Row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                '    Set grid.CellPicture = uncheck
                '    hapusrow
                '    Exit Sub
                'End If
                'Set grid.CellPicture = uncheck
                Else
                    Set grid.CellPicture = uncheck
                End If
        Case 4:
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            txthasil.Width = grid.ColWidth(grid.Col) - 40
            txthasil = grid.TextMatrix(grid.Row, grid.Col)
            txthasil.Left = grid.Left + grid.CellLeft
            txthasil.Top = grid.Top + grid.CellTop + 20
            txthasil.Visible = True
            txthasil.SetFocus
        Case 5:
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            If grid.TextMatrix(grid.Row, 5) = "Request" Then Exit Sub
            cmbstatus.Width = grid.ColWidth(grid.Col) - 40
            cmbstatus = grid.TextMatrix(grid.Row, grid.Col)
            cmbstatus.Left = grid.Left + grid.CellLeft
            cmbstatus.Top = grid.Top + grid.CellTop + 20
            cmbstatus.Visible = True
            cmbstatus.SetFocus
    End Select
End Sub

Private Sub List1_Click()
    Dim j As Integer
    j = 0
    If List1.text = "" Then Exit Sub
    txtnolot = List1.text
    
    Check1.Value = Unchecked
    cmbkode.Visible = False
    cmdview.Visible = False
    cmdupdate.Visible = False
    cmdsave.Enabled = True
    hapusgrid
    
    OBJ.Open dsn
    SQL = "SELECT a.*,b.NamaBarang,c.NamaSatuan FROM am_gudang_permintaan a"
    SQL = SQL + " inner join am_apitemmst b on a.kodebarang = b.KodeBarang "
    SQL = SQL + " inner join am_apunit c on b.KodeSatuan = c.KodeSatuan"
    SQL = SQL + " Where a.nolot='" & txtnolot & "' and a.status<>'1' and flag='0'"
    Set RST = OBJ.Execute(SQL)
    
    If Not RST.EOF Then
    
        lblTgl = "Tanggal : " & Format(RST!tgl, "dd MMMM yyyy")
        txtpetugas = RST!petugas
        Do Until RST.EOF
            grid.Col = 0
            Set grid.CellPicture = uncheck
            grid.TextMatrix(grid.Row, 1) = RST!kodebarang
            grid.TextMatrix(grid.Row, 2) = RST!namabarang
            If RST!keterangan = "Request" Then
                grid.TextMatrix(grid.Row, 3) = Format(RST!qty, "###,###,##0.00")
            Else
                grid.TextMatrix(grid.Row, 3) = Format(RST!qty_add, "###,###,##0.00")
            End If
            grid.TextMatrix(grid.Row, 4) = RST!qty_confirmed
            grid.TextMatrix(grid.Row, 5) = RST!keterangan
            If RST!keterangan = "Tambah Item" Then
                For j = 0 To grid.Cols - 1
                    grid.Col = j
                    grid.CellBackColor = &HE0E0E0
                Next
            Else
                For j = 0 To grid.Cols - 1
                    grid.Col = j
                    grid.CellBackColor = &HFFFFFF
                Next
            End If
            grid.Rows = grid.Rows + 1
            grid.Row = grid.Row + 1
            RST.MoveNext
        Loop
    Else
        MsgBox "Data not Found.", vbExclamation, AppName
        txtnolot = ""
    End If
    OBJ.Close
End Sub

Private Sub Timer1_Timer()
    List1.Clear
    OBJ.Open dsn
    SQL = "SELECT distinct nolot FROM am_gudang_permintaan WHERE status <> '1' and flag = '0'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        List1.AddItem RST!nolot
        RST.MoveNext
    Loop
    OBJ.Close
End Sub

Private Sub txthasil_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grid.TextMatrix(grid.Row, 4) = txthasil.text
        grid.SetFocus
    End If
End Sub

Private Sub txthasil_LostFocus()
    txthasil.Visible = False
End Sub

Function tanggalconfirm()
    tanggalconfirm = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Private Sub txtnolot_Change()
    LoadCetak
End Sub
