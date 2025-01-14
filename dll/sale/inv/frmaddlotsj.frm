VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmaddlotsj 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tambah Lot Surat Jalan"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9585
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   9585
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PicSJ 
      BackColor       =   &H00C0C0C0&
      Height          =   1095
      Left            =   7440
      ScaleHeight     =   1035
      ScaleWidth      =   1995
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
      Begin VB.TextBox txtOpensj 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   720
         TabIndex        =   21
         ToolTipText     =   "Input nomor SJ lalu tekan tombol Enter"
         Top             =   120
         Width           =   1215
      End
      Begin Chameleon.chameleonButton cmdCancel 
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Cancel"
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
         MICON           =   "frmaddlotsj.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "No SJ"
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
         Left            =   120
         TabIndex        =   20
         Top             =   195
         Width           =   1095
      End
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   840
      TabIndex        =   14
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   450
      Calculator      =   "frmaddlotsj.frx":031A
      Caption         =   "frmaddlotsj.frx":033A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmaddlotsj.frx":03A6
      Keys            =   "frmaddlotsj.frx":03C4
      Spin            =   "frmaddlotsj.frx":0406
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00);0"
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
      MinValue        =   -999999999999999
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
   Begin TDBText6Ctl.TDBText txtket 
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   450
      Caption         =   "frmaddlotsj.frx":042E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmaddlotsj.frx":049A
      Key             =   "frmaddlotsj.frx":04B8
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   0
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   0
      BorderStyle     =   0
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   50
      LengthAsByte    =   0
      Text            =   ""
      Furigana        =   0
      HighlightText   =   -1
      IMEMode         =   0
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
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
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   7
      Top             =   1200
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
      Left            =   600
      Picture         =   "frmaddlotsj.frx":04F4
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   6
      Top             =   1200
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
      Left            =   360
      Picture         =   "frmaddlotsj.frx":07D6
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtsj 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1080
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   8640
      TabIndex        =   0
      Top             =   4200
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
      MICON           =   "frmaddlotsj.frx":0B24
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdfind 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "No SJ"
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
      MICON           =   "frmaddlotsj.frx":0E3E
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
      Height          =   3495
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   6165
      _Version        =   393216
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin Chameleon.chameleonButton cmdsave 
      Height          =   375
      Left            =   7680
      TabIndex        =   9
      Top             =   4200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Save"
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
      MICON           =   "frmaddlotsj.frx":1158
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid2 
      Height          =   2535
      Left            =   120
      TabIndex        =   15
      Top             =   4800
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4471
      _Version        =   393216
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin Chameleon.chameleonButton cmdproses 
      Height          =   375
      Left            =   8640
      TabIndex        =   16
      Top             =   7440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Compare"
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
      MICON           =   "frmaddlotsj.frx":1472
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
      Left            =   6720
      TabIndex        =   17
      Top             =   4200
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
      MICON           =   "frmaddlotsj.frx":178C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdOpen 
      Height          =   375
      Left            =   8640
      TabIndex        =   18
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Open SJ"
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
      MICON           =   "frmaddlotsj.frx":1AA6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lbljml 
      Caption         =   "0"
      Height          =   255
      Left            =   2400
      TabIndex        =   13
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblkdgudang 
      Height          =   255
      Left            =   2640
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Nomor lot yang dipakai :"
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
      Left            =   120
      TabIndex        =   11
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label lbllot 
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
      Left            =   1920
      TabIndex        =   10
      Top             =   4200
      Width           =   4575
   End
   Begin VB.Label lblcust 
      Alignment       =   1  'Right Justify
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
      Left            =   2400
      TabIndex        =   4
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmaddlotsj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim OBJ1 As New ADODB.Connection
Dim RST1 As New ADODB.Recordset
Dim SQL1 As String

Dim OBJ2 As New ADODB.Connection
Dim RST2 As New ADODB.Recordset
Dim SQL2 As String

Dim posrow As String
Dim poscol As String
Dim kodelot As String
Dim jml As Integer
Dim tqty, Qmatch As Double
Dim reopenSJ As Boolean

Private Sub cmdCancel_Click()
    PicSJ.Visible = False
End Sub

Private Sub cmdclear_Click()
    hapusgrid
    txtsj = ""
    txtket = ""
    lblcust = ""
    lbljml = "0"
    lbllot = ""
    hapusgrid2
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdfind_Click()
    carisql1 = "Select a.nosj,b.namacust,a.tglsj,a.kodegudang from am_sjhdr a"
    carisql1 = carisql1 + " inner join am_customer b on a.kodecust=b.kodecust Where (a.Via2 ='0' or via2='1')"
    namatabel = "Surat Jalan."
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdfind_GotFocus()
    If hasil = "" Then Exit Sub
    txtsj = hasil
    lblcust = hasil1
    lblkdgudang = hasil2
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    carisj
End Sub

Private Sub carisj()
    hapusgrid2
    hapusgrid
    OBJ.Open dsn
    SQL = "Select COUNT(kodebarang)'jml' from am_sjlin Where nosj='" & txtsj & "'"
    Set RST = OBJ.Execute(SQL)
    jml = RST!jml
    
    grid.Row = 1
    SQL = "select a.kodebarang,c.namabarang,a.qty,a.kodesatuan,b.satuan,b.nolot,b.kg,b.hpp from am_sjlin a"
    SQL = SQL + " left join am_sjlot b on a.nosj=b.nosj and a.kodebarang=b.kodebarang"
    SQL = SQL + " inner join am_itemmst c on a.kodebarang=c.kodebarang"
    SQL = SQL + " where a.nosj = '" & txtsj & "' Order By a.lineitem asc"
    Set RST = OBJ.Execute(SQL)
    Do Until RST.EOF
        grid2.Col = 1
        grid2.CellAlignment = 1
        grid.Col = 1
        grid.CellAlignment = 1
        grid.TextMatrix(grid.Row, 1) = RST!kodebarang: grid2.TextMatrix(grid2.Row, 1) = RST!kodebarang
        'grid.Col = 2
        'grid.CellAlignment = 1
        grid.TextMatrix(grid.Row, 2) = RST!namabarang: grid2.TextMatrix(grid2.Row, 2) = RST!namabarang
        grid.TextMatrix(grid.Row, 3) = Format(RST!qty, "###,###,##0.00"): grid2.TextMatrix(grid2.Row, 3) = Format(RST!qty, "###,###,##0.00")
        grid.TextMatrix(grid.Row, 4) = RST!kodesatuan
        If IsNull(RST!satuan) Then
            OBJ1.Open dsn
            SQL1 = "Select namasatuan from am_unit Where kodesatuan='" & RST!kodesatuan & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            grid.TextMatrix(grid.Row, 5) = RST1!namasatuan
            grid.TextMatrix(grid.Row, 6) = ""
            grid.TextMatrix(grid.Row, 7) = ""
            grid.TextMatrix(grid.Row, 8) = ""
            grid.TextMatrix(grid.Row, 9) = ""
            OBJ1.Close
        Else
            grid.TextMatrix(grid.Row, 5) = RST!satuan
            grid.TextMatrix(grid.Row, 6) = RST!nolot
            grid.TextMatrix(grid.Row, 7) = Format(RST!kg, "###,##0.00")
            grid.TextMatrix(grid.Row, 8) = Format(RST!hpp, "###,###,##0.00")
            grid.TextMatrix(grid.Row, 9) = Format(RST!kg * RST!qty, "###,###,##0.00")
            'Edit Mode
        End If
        SetRow grid.Row, True
            
        grid.Rows = grid.Rows + 1: grid2.Rows = grid2.Rows + 1
        grid.Row = grid.Row + 1: grid2.Row = grid2.Row + 1
        lotgrid
        RST.MoveNext
    Loop
    OBJ.Close
End Sub
Private Sub hapusrow4()
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
        grid2.TextMatrix(grid2.Row, 4) = ""
        grid2.Row = grid2.Row + 1
    Loop
End Sub
Private Sub hapusgrid()
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        grid.TextMatrix(grid.Row, 1) = ""
        grid.TextMatrix(grid.Row, 2) = ""
        grid.TextMatrix(grid.Row, 3) = ""
        grid.TextMatrix(grid.Row, 4) = ""
        grid.TextMatrix(grid.Row, 5) = ""
        grid.TextMatrix(grid.Row, 6) = ""
        grid.TextMatrix(grid.Row, 7) = ""
        grid.TextMatrix(grid.Row, 8) = ""
        grid.TextMatrix(grid.Row, 9) = ""
        grid.Col = 0
        Set grid.CellPicture = blank
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1200
    grid.ColWidth(2) = 2500
    grid.ColWidth(3) = 1000
    grid.ColWidth(4) = 0
    grid.ColWidth(5) = 1000
    grid.ColWidth(6) = 1500
    grid.ColWidth(7) = 0
    grid.ColWidth(8) = 0
    grid.ColWidth(9) = 0
End Sub
Private Sub hapusgrid2()
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
        grid2.TextMatrix(grid2.Row, 1) = ""
        grid2.TextMatrix(grid2.Row, 2) = ""
        grid2.TextMatrix(grid2.Row, 3) = ""
        grid2.TextMatrix(grid2.Row, 4) = ""
        grid2.Col = 0
        Set grid2.CellPicture = blank
        grid2.Row = grid2.Row + 1
    Loop
    grid2.Rows = 2
End Sub
Private Sub hapusrow()
    grid.TextMatrix(grid.Row, 1) = ""
    grid.TextMatrix(grid.Row, 2) = ""
    grid.TextMatrix(grid.Row, 3) = ""
    grid.TextMatrix(grid.Row, 4) = ""
    grid.TextMatrix(grid.Row, 5) = ""
    grid.TextMatrix(grid.Row, 6) = ""
    grid.TextMatrix(grid.Row, 7) = ""
    grid.TextMatrix(grid.Row, 8) = ""
    grid.TextMatrix(grid.Row, 9) = ""
    
    Do While True
        If grid.TextMatrix(grid.Row + 1, 1) = "" Then
            grid.TextMatrix(grid.Row, 1) = ""
            grid.TextMatrix(grid.Row, 2) = ""
            grid.TextMatrix(grid.Row, 3) = ""
            grid.TextMatrix(grid.Row, 4) = ""
            grid.TextMatrix(grid.Row, 5) = ""
            grid.TextMatrix(grid.Row, 6) = ""
            grid.TextMatrix(grid.Row, 7) = ""
            grid.TextMatrix(grid.Row, 8) = ""
            grid.TextMatrix(grid.Row, 9) = ""
            Exit Do
        End If
        grid.TextMatrix(grid.Row, 1) = grid.TextMatrix(grid.Row + 1, 1)
        grid.TextMatrix(grid.Row, 2) = grid.TextMatrix(grid.Row + 1, 2)
        grid.TextMatrix(grid.Row, 3) = grid.TextMatrix(grid.Row + 1, 3)
        grid.TextMatrix(grid.Row, 4) = grid.TextMatrix(grid.Row + 1, 4)
        grid.TextMatrix(grid.Row, 5) = grid.TextMatrix(grid.Row + 1, 5)
        grid.TextMatrix(grid.Row, 6) = grid.TextMatrix(grid.Row + 1, 6)
        grid.TextMatrix(grid.Row, 7) = grid.TextMatrix(grid.Row + 1, 7)
        grid.TextMatrix(grid.Row, 8) = grid.TextMatrix(grid.Row + 1, 8)
        grid.TextMatrix(grid.Row, 9) = grid.TextMatrix(grid.Row + 1, 9)
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = grid.Rows - 1
    grid.Col = 0
    Set grid.CellPicture = blank
End Sub

Private Sub cmdOpen_Click()
    PicSJ.Visible = True
End Sub

Private Sub cmdproses_Click()
On Error Resume Next
    Dim Qlot, Qg As Double
    hapusrow4
    Qlot = 0
    Qmatch = 0
    grid2.Row = 1
    Do While True
        DoEvents
        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
        grid.Row = 1
        Do While True
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
            If grid.TextMatrix(grid.Row, 1) = grid2.TextMatrix(grid2.Row, 1) Then
                Qlot = Qlot + grid.TextMatrix(grid.Row, 3)
                grid2.TextMatrix(grid2.Row, 4) = Format(Qlot, "###,##0.00")
            End If
            grid.Row = grid.Row + 1
            If grid2.TextMatrix(grid2.Row, 3) = grid2.TextMatrix(grid2.Row, 4) Then
                grid2.TextMatrix(grid2.Row, 0) = "0"
            Else
                grid2.TextMatrix(grid2.Row, 0) = "1"
            End If
        Loop
        Qlot = 0
        Qmatch = Qmatch + grid2.TextMatrix(grid2.Row, 0)
        grid2.Row = grid2.Row + 1
    Loop
End Sub

Private Sub cmdsave_Click()
    If grid.TextMatrix(1, 1) = "" Then Exit Sub
    If txtsj = "" Then Exit Sub
    cmdproses_Click
    
    If lbljml < jml Then
        MsgBox "Data item tidak sesuai dengan SJ" & vbCrLf & "Mohon periksa kembali data pada SJ", vbCritical, AppName
        Exit Sub
    End If
    
    If Qmatch <> 0 Then
        MsgBox "Qty SJ tidak sesuai" & vbCrLf & "Mohon periksa kembali data pada SJ", vbCritical, AppName
        Exit Sub
    End If
    
    If MsgBox("Apakah Anda yakin ingin menyimpan data ini", vbYesNo + vbQuestion, AppName) = vbNo Then Exit Sub

    OBJ.Open dsn
    'periksa nosj di am_sjlot
    SQL = "Select * From am_sjlot Where nosj = '" & txtsj & "'"
    Set RST = OBJ.Execute(SQL)
    
    If Not RST.EOF Then
        SQL = "Delete From am_sjlot Where nosj ='" & txtsj & "'"
        Set RST = OBJ.Execute(SQL)
    End If
    OBJ.Close
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        OBJ1.Open dsn
        SQL1 = "Insert into am_sjlot ("
        SQL1 = SQL1 + "nosj,"
        SQL1 = SQL1 + "kodebarang,"
        SQL1 = SQL1 + "qty,"
        SQL1 = SQL1 + "satuan,"
        SQL1 = SQL1 + "kg,"
        SQL1 = SQL1 + "nolot,"
        SQL1 = SQL1 + "hpp,"
        SQL1 = SQL1 + "flag,"
        SQL1 = SQL1 + "kode,"
        SQL1 = SQL1 + "keterangan)"
        SQL1 = SQL1 + " Values ('" & txtsj & "',"
        SQL1 = SQL1 + "'" & grid.TextMatrix(grid.Row, 1) & "',"
        SQL1 = SQL1 + "'" & grid.TextMatrix(grid.Row, 3) & "',"
        SQL1 = SQL1 + "'" & grid.TextMatrix(grid.Row, 5) & "',"
        SQL1 = SQL1 + "'" & grid.TextMatrix(grid.Row, 7) & "',"
        SQL1 = SQL1 + "'" & grid.TextMatrix(grid.Row, 6) & "',"
        SQL1 = SQL1 + "'" & grid.TextMatrix(grid.Row, 8) & "',"
        SQL1 = SQL1 + "'0','" & kodelot & "',"
        SQL1 = SQL1 + "'" & grid.TextMatrix(grid.Row, 6) & "')"
        Set RST1 = OBJ1.Execute(SQL1)
        OBJ1.Close
             'menandai lot yang sudah kosong (Habis)
            'If grid.TextMatrix(grid.Row, 8) = "0.00" Then
                'SQL = "Update am_sjlot set flag='1' Where nolot='" & grid.TextMatrix(grid.Row, 3) & "'"
                'SQL = SQL + " And kodebarang='" & grid.TextMatrix(grid.Row, 1) & "'"
                'Set RST = OBJ.Execute(SQL)
            'End If
        grid.Row = grid.Row + 1
    Loop
    
    'simpan am_stokgudang
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        OBJ1.Open dsn
        SQL1 = "Insert into am_stokgudang ("
        SQL1 = SQL1 + "nolot,"
        SQL1 = SQL1 + "palet,"
        SQL1 = SQL1 + "tanggal,"
        SQL1 = SQL1 + "ref,"
        SQL1 = SQL1 + "keterangan,"
        SQL1 = SQL1 + "kodebarang,"
        SQL1 = SQL1 + "namabarang,"
        SQL1 = SQL1 + "kg,"
        SQL1 = SQL1 + "kgperpalet,"
        SQL1 = SQL1 + "hppperkg,"
        SQL1 = SQL1 + "qin,"
        SQL1 = SQL1 + "qout,"
        SQL1 = SQL1 + "kdsatuan,"
        SQL1 = SQL1 + "satuan,"
        SQL1 = SQL1 + "gudang,"
        SQL1 = SQL1 + "username,"
        SQL1 = SQL1 + "flag)"
        SQL1 = SQL1 + " Values ('" & grid.TextMatrix(grid.Row, 6) & "',"
        SQL1 = SQL1 + "'01' + '" & grid.TextMatrix(grid.Row, 6) & "',"
        SQL1 = SQL1 + "convert(datetime,'" & tanggalsj & "'),"
        SQL1 = SQL1 + "'" & txtsj & "',"
        SQL1 = SQL1 + "'SJ',"
        SQL1 = SQL1 + "'" & grid.TextMatrix(grid.Row, 1) & "',"
        SQL1 = SQL1 + "'" & grid.TextMatrix(grid.Row, 2) & "',"
        SQL1 = SQL1 + "'" & grid.TextMatrix(grid.Row, 7) & "',"
        SQL1 = SQL1 + "'" & grid.TextMatrix(grid.Row, 9) & "',"
        SQL1 = SQL1 + "'" & grid.TextMatrix(grid.Row, 8) & "',"
        SQL1 = SQL1 + "'0.00',"
        SQL1 = SQL1 + "'" & grid.TextMatrix(grid.Row, 3) & "',"
        SQL1 = SQL1 + "'" & grid.TextMatrix(grid.Row, 4) & "',"
        SQL1 = SQL1 + "'" & grid.TextMatrix(grid.Row, 5) & "',"
        SQL1 = SQL1 + "'" & lblkdgudang & "',"
        SQL1 = SQL1 + "'" & nmuser & "','1')"
        Set RST1 = OBJ1.Execute(SQL1)
        OBJ1.Close
        grid.Row = grid.Row + 1
    Loop
    
    If reopenSJ = True Then
        reopenSJ = False
        OBJ.Open dsn
        SQL = "Update am_sjhdr set via2='2' Where nosj='" & txtsj.text & "'"
        Set RST = OBJ.Execute(SQL)
        OBJ.Close
    End If
    MsgBox "Data berhasil disimpan", vbInformation, AppName
    cmdclear_Click
    
End Sub

Private Sub Form_Load()
    grid.Cols = 10
    grid.TextMatrix(0, 1) = "Kode Barang"
    grid.TextMatrix(0, 2) = "Nama Barang"
    grid.TextMatrix(0, 3) = "Quantity"
    grid.TextMatrix(0, 4) = "K/Satuan"
    grid.TextMatrix(0, 5) = "N/Satuan"
    grid.TextMatrix(0, 6) = "No Lot"
    grid.TextMatrix(0, 7) = "kg"
    grid.TextMatrix(0, 8) = "hpp/kg"
    grid.TextMatrix(0, 9) = "kg/palet"
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1200
    grid.ColWidth(2) = 2500
    grid.ColWidth(3) = 1000
    grid.ColWidth(4) = 0
    grid.ColWidth(5) = 1000
    grid.ColWidth(6) = 1800
    grid.ColWidth(7) = 0
    grid.ColWidth(8) = 0
    grid.ColWidth(9) = 0
    grid.RowHeightMin = 300
    kodelot = getkode
    
    grid2.Cols = 5
    grid2.TextMatrix(0, 1) = "Kode Barang"
    grid2.TextMatrix(0, 2) = "Nama Barang"
    grid2.TextMatrix(0, 3) = "Qty SJ"
    grid2.TextMatrix(0, 4) = "Qty Lot"
    grid2.ColWidth(0) = 250
    grid2.ColWidth(1) = 1200
    grid2.ColWidth(2) = 2500
    grid2.ColWidth(3) = 1000
    grid2.ColWidth(4) = 1000
    grid2.RowHeightMin = 300
    PicSJ.Visible = False
    reopenSJ = False
End Sub
Private Sub SetRow(ByVal idx As Integer, ByVal hapus As String)
    grid.Row = idx
    grid.Col = 0
    If hapus Then
        Set grid.CellPicture = uncheck.Picture
    End If
    grid.Col = 1
End Sub

Private Sub grid_Click()
On Error Resume Next
    If grid.MouseRow = 0 Then Exit Sub
    Select Case grid.Col
        Case 0:
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            If grid.CellPicture = uncheck Then
                Set grid.CellPicture = check
                If MsgBox("Delete that Row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                    Set grid.CellPicture = uncheck
                    hapusrow
                    lotgrid
                    Exit Sub
                End If
                Set grid.CellPicture = uncheck
            End If
        Case 1:
            carisql1 = "select a.kodebarang,c.namabarang,a.qty,a.kodesatuan,d.namasatuan,b.nolot,b.kg,b.hpp from am_sjlin a"
            carisql1 = carisql1 + " left join am_sjlot b on a.nosj=b.nosj and a.kodebarang=b.kodebarang"
            carisql1 = carisql1 + " inner join am_itemmst c on a.kodebarang=c.kodebarang"
            carisql1 = carisql1 + " inner join am_unit d on a.kodesatuan= d.kodesatuan"
            carisql1 = carisql1 + " where a.nosj = '" & txtsj & "'"
            namatabel = "SJ Item"
            frmsearch.Show vbModal
        Case 3:
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            txtnilai.Width = grid.ColWidth(grid.Col) - 40
            txtnilai = grid.TextMatrix(grid.Row, grid.Col)
            txtnilai.Left = grid.Left + grid.CellLeft
            txtnilai.Top = grid.Top + grid.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
        Case 6:
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            posrow = grid.Row
            poscol = grid.Row
            txtket.Width = grid.ColWidth(grid.Col) - 40
            txtket = grid.TextMatrix(grid.Row, grid.Col)
            txtket.Left = grid.Left + grid.CellLeft
            txtket.Top = grid.Top + grid.CellTop + 20
            txtket.Visible = True
            txtket.SetFocus
    End Select
End Sub

Private Sub grid_EnterCell()
    If grid.MouseRow = 0 Then Exit Sub
    Select Case grid.Col
        Case 0:
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            If grid.CellPicture = uncheck Then
                Set grid.CellPicture = check
                If MsgBox("Delete that Row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                    Set grid.CellPicture = uncheck
                    hapusrow
                    lotgrid
                    Exit Sub
                End If
                Set grid.CellPicture = uncheck
            End If
        Case 1:
            carisql1 = "select a.kodebarang,c.namabarang,a.qty,a.kodesatuan,d.namasatuan,b.nolot,b.kg,b.hpp from am_sjlin a"
            carisql1 = carisql1 + " left join am_sjlot b on a.nosj=b.nosj and a.kodebarang=b.kodebarang"
            carisql1 = carisql1 + " inner join am_itemmst c on a.kodebarang=c.kodebarang"
            carisql1 = carisql1 + " inner join am_unit d on a.kodesatuan= d.kodesatuan"
            carisql1 = carisql1 + " where a.nosj = '" & txtsj & "'"
            namatabel = "SJ Item"
            frmsearch.Show vbModal
        Case 3:
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            txtnilai.Width = grid.ColWidth(grid.Col) - 40
            txtnilai = grid.TextMatrix(grid.Row, grid.Col)
            txtnilai.Left = grid.Left + grid.CellLeft
            txtnilai.Top = grid.Top + grid.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
        Case 6:
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            posrow = grid.Row
            poscol = grid.Row
            txtket.Width = grid.ColWidth(grid.Col) - 40
            txtket = grid.TextMatrix(grid.Row, grid.Col)
            txtket.Left = grid.Left + grid.CellLeft
            txtket.Top = grid.Top + grid.CellTop + 20
            txtket.Visible = True
            txtket.SetFocus
    End Select
End Sub

Private Sub grid_GotFocus()
    If grid.MouseRow = 0 Then Exit Sub
    Select Case grid.Col
        Case 1:
            If hasil = "" Then Exit Sub
            'periksa total qty jika > dari Qty SJ maka batalkan
            tqty = 0
            grid.Row = 1
            Do While True
                DoEvents
                If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
                OBJ1.Open dsn
                SQL1 = "Select * From am_sjlin Where nosj = '" & txtsj & "' and kodebarang='" & hasil & "'"
                SQL1 = SQL1 + " and kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then tqty = tqty + grid.TextMatrix(grid.Row, 3)
                OBJ1.Close
                grid.Row = grid.Row + 1
            Loop

            If hasil = "" Then Exit Sub
            OBJ1.Open dsn
            SQL1 = "Select * From am_sjlin Where nosj = '" & txtsj & "' and kodebarang='" & hasil & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If RST1!qty <= tqty Then
                MsgBox "Jumlah Qty telah mencukupi", vbCritical, AppName
                'MsgBox RST!qty & " <= " & tqty
                OBJ1.Close
                hasil = "": hasil1 = "": hasil2 = "": hasil3 = "": hasil4 = "": hasil5 = "": hasil6 = ""
                Exit Sub
            End If
            Dim row3 As Double
            row3 = RST1!qty - tqty
            OBJ1.Close
            grid.TextMatrix(grid.Row, 1) = hasil
            grid.TextMatrix(grid.Row, 2) = hasil1
            grid.TextMatrix(grid.Row, 3) = Format(row3, "###,##0.00")
            grid.TextMatrix(grid.Row, 4) = hasil3
            grid.TextMatrix(grid.Row, 5) = hasil4
            grid.TextMatrix(grid.Row, 6) = hasil5
            grid.TextMatrix(grid.Row, 7) = Format(hasil6, "###,##0.00")
            SetRow grid.Row, True
            grid.Rows = grid.Rows + 1
            grid.Row = grid.Row + 1
            lotgrid
    End Select
End Sub

Private Sub txtket_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    If KeyAscii = 13 Then
        Select Case grid.Col
            Case 6
        'periksa lot
                OBJ.Open dsn
                SQL = "Select * From list_hpp_produksi Where nolot='" & txtket & "' and kodebarang='" & grid.TextMatrix(posrow, 1) & "'"
                'SQL = "Select * From list_hpp_produksi Where nolot='" & txtket & "'"   '(sementara buat item ganti baju)
                Set RST = OBJ.Execute(SQL)

                If Not RST.EOF Then
                    
                    grid.Col = 6
                    grid.CellAlignment = 1
                    grid.TextMatrix(posrow, 6) = txtket
                    grid.TextMatrix(grid.Row, 7) = Format(RST!kg, "###,##0.00")
                    grid.TextMatrix(grid.Row, 8) = Format(RST!hppperkg, "###,###,##0.00")
                    grid.TextMatrix(grid.Row, 9) = Format(RST!kg * grid.TextMatrix(posrow, 3), "###,##0.00")
                    txtket.Visible = False
                    grid.SetFocus
                Else
                    'hitung hpp lot lama, ambil data langsung dari produksi
                    OBJ1.Open dsn
                    SQL1 = "Select a.noref,a.tanggal,b.kodebarang,b.kg,isnull(c.pack,0)'pack',e.thppbahan,e.perkilo,isnull(g.thpppack,0)'thpppack',g.thasil,(a.qty_bahan*b.kg)'hasil',"
                    SQL1 = SQL1 + " e.thppbahan +isnull(g.thpppack,0)'brutto',(g.thasil*e.perkilo)+isnull(g.thpppack,0)'tjadi',"
                    SQL1 = SQL1 + " (e.thppbahan +isnull(g.thpppack,0))-((g.thasil*e.perkilo)+isnull(g.thpppack,0))'loss',"
                    SQL1 = SQL1 + " (((e.thppbahan +isnull(g.thpppack,0))-((g.thasil*e.perkilo)+isnull(g.thpppack,0)))/g.thasil)'lossperkg',"
                    SQL1 = SQL1 + " (isnull(c.pack,0)/(a.qty_bahan*b.kg))'packperkg',"
                    SQL1 = SQL1 + " (e.perkilo + (isnull(c.pack,0)/(a.qty_bahan*b.kg))+(((e.thppbahan +isnull(g.thpppack,0))-((g.thasil*e.perkilo)+isnull(g.thpppack,0)))/g.thasil))'hppperkg'"
                    SQL1 = SQL1 + " From list_produksi_hasil a"
                    SQL1 = SQL1 + " inner join (select kodebarang,kodesatuan,tahun,case when SUBSTRING('" & txtket & "',3,1)='A' then kg1"
                    SQL1 = SQL1 + " when SUBSTRING('" & txtket & "',3,1)='B' then kg2"
                    SQL1 = SQL1 + " when SUBSTRING('" & txtket & "' ,3,1)='C' then kg3"
                    SQL1 = SQL1 + " when SUBSTRING('" & txtket & "' ,3,1)='D' then kg4"
                    SQL1 = SQL1 + " when SUBSTRING('" & txtket & "' ,3,1)='E' then kg5"
                    SQL1 = SQL1 + " when SUBSTRING('" & txtket & "' ,3,1)='F' then kg6"
                    SQL1 = SQL1 + " when SUBSTRING('" & txtket & "' ,3,1)='G' then kg7"
                    SQL1 = SQL1 + " when SUBSTRING('" & txtket & "' ,3,1)='H' then kg8"
                    SQL1 = SQL1 + " when SUBSTRING('" & txtket & "' ,3,1)='J' then kg9"
                    SQL1 = SQL1 + " when SUBSTRING('" & txtket & "' ,3,1)='K' then kg10"
                    SQL1 = SQL1 + " when SUBSTRING('" & txtket & "' ,3,1)='L' then kg11"
                    SQL1 = SQL1 + " when SUBSTRING('" & txtket & "' ,3,1)='M' then kg12 End as kg From am_itemkg)"
                    SQL1 = SQL1 + " b on a.kode_bahan = b.kodebarang and a.kode_satuan = b.kodesatuan"
                    SQL1 = SQL1 + " left join (select noref,isnull(SUM(hpp),0)'pack' from list_produksi_kemasan where nolot = '" & txtket & "' group by noref) c on a.noref = c.noref"
                    SQL1 = SQL1 + " left join list_produksi_child d on a.nolot = d.nolot"
                    SQL1 = SQL1 + " inner join (Select x.nolot,y.noref,SUM(x.hpp)'thppbahan',SUM(x.hpp)/SUM(x.qty_bahan)'perkilo'"
                    SQL1 = SQL1 + " from list_produksi_child x left join list_produksi_hasil y on x.nolot = y.nolot where x.nolot ='" & txtket & "'"
                    SQL1 = SQL1 + " group by x.nolot,y.noref) e on a.noref = e.noref"
                    SQL1 = SQL1 + " left join (Select m.nolot,o.thpppack,SUM(m.qty_bahan * n.kg)'thasil' From list_produksi_hasil m"
                    SQL1 = SQL1 + " inner join (select kodebarang,kodesatuan,tahun,case when SUBSTRING('" & txtket & "',3,1)='A' then kg1"
                    SQL1 = SQL1 + " when SUBSTRING('" & txtket & "',3,1)='B' then kg2"
                    SQL1 = SQL1 + " when SUBSTRING('" & txtket & "' ,3,1)='C' then kg3"
                    SQL1 = SQL1 + " when SUBSTRING('" & txtket & "' ,3,1)='D' then kg4"
                    SQL1 = SQL1 + " when SUBSTRING('" & txtket & "' ,3,1)='E' then kg5"
                    SQL1 = SQL1 + " when SUBSTRING('" & txtket & "' ,3,1)='F' then kg6"
                    SQL1 = SQL1 + " when SUBSTRING('" & txtket & "' ,3,1)='G' then kg7"
                    SQL1 = SQL1 + " when SUBSTRING('" & txtket & "' ,3,1)='H' then kg8"
                    SQL1 = SQL1 + " when SUBSTRING('" & txtket & "' ,3,1)='J' then kg9"
                    SQL1 = SQL1 + " when SUBSTRING('" & txtket & "' ,3,1)='K' then kg10"
                    SQL1 = SQL1 + " when SUBSTRING('" & txtket & "' ,3,1)='L' then kg11"
                    SQL1 = SQL1 + " when SUBSTRING('" & txtket & "' ,3,1)='M' then kg12 End as kg From am_itemkg)"
                    SQL1 = SQL1 + " n on m.kode_bahan = n.kodebarang and m.kode_satuan = n.kodesatuan"
                    SQL1 = SQL1 + " left join (Select nolot,isnull(SUM(hpp),0)'thpppack' from list_produksi_kemasan Where nolot = '" & txtket & "' group by nolot)"
                    SQL1 = SQL1 + " o on m.nolot=o.nolot"
                    SQL1 = SQL1 + " Where m.nolot = '" & txtket & "' and n.tahun = '20' + LEFT('" & txtket & "',2) group by m.nolot,o.thpppack) g on a.nolot = g.nolot"
                    SQL1 = SQL1 + " Where a.nolot = '" & txtket & "' and b.tahun = '20' + LEFT('" & txtket & "',2)"
                    SQL1 = SQL1 + " and a.kode_bahan= '" & grid.TextMatrix(grid.Row, 1) & "'"   'query tambahan
                    SQL1 = SQL1 + " group by a.noref,a.tanggal,b.kodebarang,b.kg,c.pack,e.thppbahan,e.perkilo,g.thpppack,g.thasil,a.qty_bahan order by a.noref asc"
                    Set RST1 = OBJ1.Execute(SQL1)
                    If Not RST1.EOF Then
                        grid.TextMatrix(posrow, 6) = txtket
                        grid.TextMatrix(grid.Row, 7) = Format(RST1!kg, "###,##0.00")
                        grid.TextMatrix(grid.Row, 8) = Format(RST1!hppperkg, "###,###,##0.00")
                        grid.TextMatrix(grid.Row, 9) = Format(RST1!kg * grid.TextMatrix(posrow, 3), "###,##0.00")
                        txtket.Visible = False
                        grid.SetFocus
                        MsgBox "Data diambil dari lot produksi", vbInformation, AppName
                    Else
                        'periksa overzak
                        OBJ2.Open dsn
                        SQL2 = "Select * From am_stokgudang Where nolot='" & txtket & "' and kodebarang='" & grid.TextMatrix(grid.Row, 1) & "'"
                        Set RST2 = OBJ2.Execute(SQL2)
                        If Not RST2.EOF Then
                            grid.TextMatrix(posrow, 6) = txtket
                            grid.TextMatrix(grid.Row, 7) = Format(RST2!kg, "###,##0.00")
                            grid.TextMatrix(grid.Row, 8) = Format(RST2!hppperkg, "###,###,##0.00")
                            grid.TextMatrix(grid.Row, 9) = Format(RST2!kgperpalet, "###,##0.00")
                            txtket.Visible = False
                            grid.SetFocus
                            MsgBox "Data diambil dari lot over zak", vbInformation, AppName
                        Else
                            MsgBox "Lot number not Found" & vbCrLf & "Please check your lot and item code", vbCritical, AppName
                            txtket = ""
                        End If
                        OBJ2.Close
                    End If
                    OBJ1.Close
                End If
                lotgrid
                OBJ.Close
        End Select
    End If
End Sub

Private Sub lotgrid()
On Error Resume Next
    Dim lot As String
    grid.Row = 1
    lot = ""
    Do While True
        DoEvents
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        lot = lot & grid.TextMatrix(grid.Row, 6) & ","
        lbljml = grid.Row
        grid.Row = grid.Row + 1
    Loop
    lbllot = lot
End Sub

Function getkode() As String    '230323001'
    On Error GoTo Err_handler:
    Dim SQL As String
    Dim strnumber As String
    Dim tempkode As String
    Dim kd As Long
    
    strnumber = Format(Date, "yymmdd")
    
    Set OBJ = New ADODB.Connection
    OBJ.Open dsn
    SQL = "select max(kode)as kr from am_sjlot where kode like '" + strnumber + "%'"
    Set RST = OBJ.Execute(SQL)

    If IsNull(RST!kr) = True Or RST!kr = "" Then
        getkode = strnumber + "001"
    Else
        kd = CLng(Mid(RST!kr, 7, 3)) + 1
        
        If (Len(Trim(Str(kd))) = 1) Then
            tempkode = strnumber + "00" + Trim(Str(kd))
        End If
        If (Len(Trim(Str(kd))) = 2) Then
            tempkode = strnumber + "0" + Trim(Str(kd))
        End If
        If (Len(Trim(Str(kd))) = 3) Then
            tempkode = strnumber + Trim(Str(kd))
        End If
        getkode = tempkode
    End If
    OBJ.Close
    Exit Function
Err_handler:
    getkode = strnumber + "001"
End Function
Function tanggalsj()
    tanggalsj = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function

Private Sub txtket_LostFocus()
    txtket.Visible = False
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grid.TextMatrix(grid.Row, 3) = Format(txtnilai, "#,##0.00")
        grid.SetFocus
    End If
End Sub

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
End Sub

Private Sub txtOpensj_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        OBJ.Open dsn
        SQL = "Select * From am_sjhdr Where nosj='" & txtOpensj.text & "'"
        Set RST = OBJ.Execute(SQL)
        
        If Not RST.EOF Then
            If RST!via2 = "2" Then
                If MsgBox("Buka kembali SJ close ?", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then
                    OBJ.Close
                    PicSJ.Visible = False
                    Exit Sub
                Else
                    OBJ1.Open dsn
                    SQL1 = "Update am_sjhdr set via2='1' Where nosj='" & txtOpensj & "'"
                    Set RST1 = OBJ1.Execute(SQL1)
                    MsgBox "No SJ berhasil dibuka", vbInformation, AppName
                    txtOpensj = ""
                    PicSJ.Visible = False
                    reopenSJ = True
                    OBJ1.Close
                End If
            Else
                MsgBox "Silahkan input nomor Lot SJ", vbInformation, AppName
                txtOpensj = ""
                PicSJ.Visible = False
            End If
        Else
            MsgBox "Nomor SJ tidak ditemukan", vbCritical, AppName
            txtOpensj = ""
        End If
        OBJ.Close
    End If
End Sub
