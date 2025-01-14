VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmpermint_list 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Update Lot Permintaan"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Left            =   8280
      TabIndex        =   17
      Text            =   "Combo1"
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
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
      Height          =   4860
      ItemData        =   "frmpermint_list.frx":0000
      Left            =   120
      List            =   "frmpermint_list.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   6
      ToolTipText     =   "Klik Nomor Lot untuk confirm"
      Top             =   480
      Width           =   2655
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
      TabIndex        =   5
      Top             =   720
      Width           =   2535
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
      MaxLength       =   17
      TabIndex        =   4
      Top             =   1080
      Width           =   2535
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
      Picture         =   "frmpermint_list.frx":0004
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   240
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
      Picture         =   "frmpermint_list.frx":0352
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   255
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
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   9120
      TabIndex        =   0
      Top             =   4680
      Visible         =   0   'False
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   450
      Calculator      =   "frmpermint_list.frx":0634
      Caption         =   "frmpermint_list.frx":0654
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpermint_list.frx":06C0
      Keys            =   "frmpermint_list.frx":06DE
      Spin            =   "frmpermint_list.frx":0720
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
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   10080
      TabIndex        =   7
      Top             =   5280
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
      MICON           =   "frmpermint_list.frx":0748
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
      TabIndex        =   8
      Top             =   1590
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   6376
      _Version        =   393216
      Cols            =   10
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
      _Band(0).Cols   =   10
   End
   Begin Chameleon.chameleonButton cmdsave 
      Height          =   375
      Left            =   9120
      TabIndex        =   9
      Top             =   5280
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Add"
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
      MICON           =   "frmpermint_list.frx":0A62
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
      Left            =   8160
      TabIndex        =   10
      Top             =   5280
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
      MICON           =   "frmpermint_list.frx":0D7C
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
      Left            =   8400
      TabIndex        =   11
      Top             =   720
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
      Format          =   59637761
      CurrentDate     =   42039
   End
   Begin Chameleon.chameleonButton cmdrefresh 
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   5400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Refresh"
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
      MICON           =   "frmpermint_list.frx":1096
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Crystal.CrystalReport crystal 
      Left            =   0
      Top             =   5520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Chameleon.chameleonButton cmdreprint 
      Height          =   375
      Left            =   1560
      TabIndex        =   21
      Top             =   5400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "RePrint"
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
      MICON           =   "frmpermint_list.frx":13B0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdeditnolot 
      Height          =   375
      Left            =   6840
      TabIndex        =   22
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Update No Lot"
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
      MICON           =   "frmpermint_list.frx":16CA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Qty"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7560
      TabIndex        =   20
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "*   Semua kerusakan dan kehilangan barang (Packaging) harus di input di form ini berdasarkan Lot permintaan/Pengambilan barang"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   3120
      TabIndex        =   18
      Top             =   5280
      Width           =   4815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lot Confirmed"
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
      TabIndex        =   16
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Add/Edit"
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
      TabIndex        =   15
      Top             =   120
      Width           =   8295
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
      TabIndex        =   14
      Top             =   750
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
      Left            =   8400
      TabIndex        =   13
      Top             =   720
      Width           =   2535
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
      TabIndex        =   12
      Top             =   1110
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H80000014&
      FillStyle       =   0  'Solid
      Height          =   5175
      Left            =   3000
      Top             =   600
      Width           =   8295
   End
End
Attribute VB_Name = "frmpermint_list"
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
Dim lotedit As String

Private Sub cmbstatus_Click()
    grid.TextMatrix(grid.Row, 7) = cmbstatus
    cmbstatus.Visible = False
    If grid.TextMatrix(grid.Row, 7) = "Tambah Item" Then
        grid.TextMatrix(grid.Row, 8) = "0"
    ElseIf grid.TextMatrix(grid.Row, 7) = "Rusak" Then
        grid.TextMatrix(grid.Row, 8) = "5"
    ElseIf grid.TextMatrix(grid.Row, 7) = "Hilang" Then
        grid.TextMatrix(grid.Row, 8) = "6"
    ElseIf grid.TextMatrix(grid.Row, 7) = "Return Gudang" Then
        grid.TextMatrix(grid.Row, 8) = "7"
    End If
End Sub

Private Sub cmbstatus_LostFocus()
    cmbstatus.Visible = False
End Sub

Private Sub cmdclear_Click()
    txtnolot = ""
    lotedit = ""
    txtpetugas = ""
    date1 = Date
    cmbstatus.Visible = False
    txtnilai.Visible = False
    cmdsave.Caption = "Add"
    hapusgrid
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdeditnolot_Click()
    If lotedit = "" Then Exit Sub
    If MsgBox("Are you sure, you want to update this data ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    OBJ.Open dsn
    SQL = "UPDATE am_gudang_permintaan SET nolot = '" & txtnolot & "' Where nolot = '" & lotedit & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    MsgBox "No Lot berhasil diupdate", vbInformation, AppName
    cmdclear_Click
End Sub

Private Sub cmdrefresh_Click()
    List1.Clear
    OBJ.Open dsn
    SQL = "SELECT distinct nolot FROM am_gudang_permintaan WHERE status <>'0' and flag='1'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        List1.AddItem RST!nolot
        RST.MoveNext
    Loop
    OBJ.Close
End Sub

Private Sub cmdreprint_Click()
    frmpermint_reprint.Show vbModal
End Sub

Private Sub cmdsave_Click()
    Dim i As Integer
    If grid.TextMatrix(grid.Row, 8) = "0" And grid.TextMatrix(grid.Row, 7) = "" Then
        MsgBox "Keterangan pada baris penambahan item tidak boleh kosong", vbExclamation, AppName
        Exit Sub
    End If
    If grid.TextMatrix(grid.Row, 8) = "" Then
        MsgBox "Keterangan pada baris penambahan item tidak boleh kosong", vbExclamation, AppName
        Exit Sub
    End If
    If grid.TextMatrix(grid.Row, 3) = "0" And grid.TextMatrix(grid.Row, 4) = "" Then
        MsgBox "Qty item tidak boleh kosong", vbExclamation, AppName
        Exit Sub
    End If
 
    If MsgBox("Are you sure, you want to update this data ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
        
    OBJ.Open dsn
    SQL = "Select MAX(cetak_ke)'no' from am_gudang_permintaan Where nolot = '" & txtnolot & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        i = RST!no + 1
    Else
        i = 1
    End If
    
    
        
    If cmdsave.Caption = "Add" Then
        SQL = "Select * From am_gudang_permintaan Where 0=1"
        Set RST = New ADODB.Recordset
        RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
        grid.Row = 1
        Do While True
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
            If grid.TextMatrix(grid.Row, 9) = "1" Then GoTo lewati:
            RST.AddNew
            RST!nolot = txtnolot
            RST!tgl = date1
            RST!petugas = txtpetugas
            RST!kodebarang = grid.TextMatrix(grid.Row, 1)
            RST!qty = "0"
            RST!Status = grid.TextMatrix(grid.Row, 8)
            If grid.TextMatrix(grid.Row, 7) = "Tambah Item" Then
                RST!qty_add = grid.TextMatrix(grid.Row, 3)
            ElseIf grid.TextMatrix(grid.Row, 7) = "Rusak" Then
                RST!qty_add = grid.TextMatrix(grid.Row, 3) * -1
            ElseIf grid.TextMatrix(grid.Row, 7) = "Hilang" Then
                RST!qty_add = grid.TextMatrix(grid.Row, 3) * -1
            ElseIf grid.TextMatrix(grid.Row, 7) = "Return Gudang" Then
                RST!qty_add = grid.TextMatrix(grid.Row, 3) * -1
            End If
            RST!keterangan = grid.TextMatrix(grid.Row, 7)
            RST!flag = "0"
            RST!qty_confirmed = "0"
            RST!cetak_ke = i
            RST.Update
lewati:
            grid.Row = grid.Row + 1
        Loop
    
    ElseIf cmdsave.Caption = "Update" Then
        SQL = "Delete From am_gudang_permintaan"
        SQL = SQL + " Where user_confirmed IS NULL and nolot = '" & txtnolot & "' and flag = '0'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "Select * From am_gudang_permintaan Where 0=1"
        Set RST = New ADODB.Recordset
        RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
        grid.Row = 1
        Do While True
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
            If grid.TextMatrix(grid.Row, 9) = "1" Then GoTo lewati2:
            RST.AddNew
            RST!nolot = txtnolot
            RST!tgl = date1
            RST!petugas = txtpetugas
            RST!kodebarang = grid.TextMatrix(grid.Row, 1)
            RST!qty = "0"
            RST!Status = grid.TextMatrix(grid.Row, 8)
            If grid.TextMatrix(grid.Row, 7) = "Tambah Item" Then
                RST!qty_add = grid.TextMatrix(grid.Row, 3)
            ElseIf grid.TextMatrix(grid.Row, 7) = "Rusak" Then
                RST!qty_add = grid.TextMatrix(grid.Row, 3) * -1
            ElseIf grid.TextMatrix(grid.Row, 7) = "Hilang" Then
                RST!qty_add = grid.TextMatrix(grid.Row, 3) * -1
            ElseIf grid.TextMatrix(grid.Row, 7) = "Return Gudang" Then
                RST!qty_add = grid.TextMatrix(grid.Row, 3) * -1
            End If
            RST!keterangan = grid.TextMatrix(grid.Row, 7)
            RST!flag = "0"
            RST!qty_confirmed = "0"
            RST!cetak_ke = i - 1
            RST.Update
lewati2:
            grid.Row = grid.Row + 1
        Loop
    End If
    OBJ.Close
    MsgBox "Berhasil Disimpan", vbInformation, AppName
    cetakreport
    cmdclear_Click
End Sub

Private Sub cetakreport()
    crystal.Reset
    crystal.WindowState = crptMaximized
    crystal.WindowShowCloseBtn = True
    crystal.WindowShowPrintSetupBtn = True
    crystal.WindowShowSearchBtn = True
    crystal.Connect = dsnreport
    crystal.DataFiles(0) = "Proc(am_cetake_repack)"
    crystal.ReportFileName = AppPath & "\reports\produksi\take_repack.rpt"
    crystal.ParameterFields(0) = "@nolot;" & txtnolot.text & ";true"
    'crystal.ParameterFields(1) = "@username;" & nmuser & ";true"
    crystal.RetrieveDataFiles
    crystal.Action = 1
End Sub

Private Sub Form_Load()
    grid.TextMatrix(0, 0) = "X"
    grid.TextMatrix(0, 1) = "Kode"
    grid.TextMatrix(0, 2) = "Nama Barang"
    grid.TextMatrix(0, 3) = "Produksi"
    grid.TextMatrix(0, 4) = "Gudang"
    grid.TextMatrix(0, 5) = "Kode"
    grid.TextMatrix(0, 6) = "Satuan"
    grid.TextMatrix(0, 7) = "Keterangan"
    grid.TextMatrix(0, 8) = "Status"
    grid.TextMatrix(0, 9) = "Fg"
    grid.ColWidth(0) = 300
    grid.ColWidth(1) = 1100
    grid.ColWidth(2) = 3000
    grid.ColWidth(3) = 1000
    grid.ColWidth(4) = 1000
    grid.ColWidth(5) = 0
    grid.ColWidth(6) = 0
    grid.ColWidth(7) = 1500
    grid.ColWidth(8) = 0 '300
    grid.ColWidth(9) = 0 '300
    grid.ColAlignmentFixed(3) = flexAlignCenterCenter
    grid.ColAlignmentFixed(4) = flexAlignCenterCenter

    List1.Clear
    OBJ.Open dsn
    SQL = "SELECT distinct nolot FROM am_gudang_permintaan WHERE status >'0' and flag='1'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        List1.AddItem RST!nolot
        RST.MoveNext
    Loop
    OBJ.Close
    cmbstatus.AddItem "Tambah Item"
    cmbstatus.AddItem "Rusak"
    cmbstatus.AddItem "Hilang"
    cmbstatus.AddItem "Return Gudang"
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    Select Case grid.Col
        Case 0:
                If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
                If grid.TextMatrix(grid.Row, 8) = "1" Then Exit Sub
                If grid.CellPicture = uncheck Then
                Set grid.CellPicture = check
                If MsgBox("Delete Row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                    Set grid.CellPicture = uncheck
                    hapusrow
                    Exit Sub
                End If
                Set grid.CellPicture = uncheck
                End If
        Case 1:
                If txtnolot = "" Then Exit Sub
                If grid.TextMatrix(grid.Row, 8) = "1" Then Exit Sub
                carisql1 = "select kodebarang, namabarang from am_apitemmst where KodeProduk in('KTN/L','ETK/L','KLG/L','U/SP')"
                namatabel = "Kemasan"
                frmsearch.Show vbModal
        Case 3:
                If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
                If grid.TextMatrix(grid.Row, 8) = "1" Then Exit Sub
                txtnilai.Width = grid.ColWidth(grid.Col) - 40
                txtnilai = grid.TextMatrix(grid.Row, grid.Col)
                txtnilai.Left = grid.Left + grid.CellLeft
                txtnilai.Top = grid.Top + grid.CellTop + 20
                txtnilai.Visible = True
                txtnilai.SetFocus
        Case 7:
                If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
                If grid.TextMatrix(grid.Row, 8) = "1" Then Exit Sub
                cmbstatus.Width = grid.ColWidth(grid.Col) - 40
                cmbstatus = grid.TextMatrix(grid.Row, grid.Col)
                cmbstatus.Left = grid.Left + grid.CellLeft
                cmbstatus.Top = grid.Top + grid.CellTop + 20
                cmbstatus.Visible = True
                cmbstatus.SetFocus
    End Select
End Sub

Private Sub grid_GotFocus()
    If hasil = "" Then Exit Sub
    
    Select Case grid.Col
        Case 1:
            grid.TextMatrix(grid.Row, 1) = hasil
            grid.TextMatrix(grid.Row, 2) = hasil1
            'cari satuan
            SQL = "select kodesatuan from am_apitemmst  "
            SQL = SQL + "where kodebarang='" & hasil & "'"
            OBJ.Open dsn
            Set RST = New ADODB.Recordset
            RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
            grid.TextMatrix(grid.Row, 3) = "0.00"
            grid.TextMatrix(grid.Row, 5) = RST!kodesatuan
                    
            'cari nama satuan
            SQL = "select * from am_apunit where kodesatuan ='" & grid.TextMatrix(grid.Row, 5) & "'"
            Set RST = OBJ.Execute(SQL)
            grid.TextMatrix(grid.Row, 6) = RST!namasatuan

            OBJ.Close
                    
            grid.Col = 0
            Set grid.CellPicture = uncheck

            If grid.Rows = grid.Row + 1 Then
                grid.Rows = grid.Rows + 1
                grid.Row = grid.Row + 1
            Else
            End If
            hasil = ""
            hasil1 = ""
            namatabel = ""
            carisql1 = ""
    End Select
End Sub

Private Sub List1_Click()
    If List1.text = "" Then Exit Sub
    txtnolot = List1.text
    lotedit = List1.text
    hapusgrid
    
    OBJ.Open dsn
    SQL = "SELECT a.*,b.NamaBarang,c.NamaSatuan FROM am_gudang_permintaan a"
    SQL = SQL + " inner join am_apitemmst b on a.kodebarang = b.KodeBarang "
    SQL = SQL + " inner join am_apunit c on b.KodeSatuan = c.KodeSatuan"
    SQL = SQL + " Where a.nolot='" & txtnolot & "' and a.status<>'0' and a.flag='1'"
    Set RST = OBJ.Execute(SQL)
    
    If Not RST.EOF Then
    
        lblTgl = "Tanggal : " & Format(RST!tgl, "dd MMMM yyyy")
        date1 = RST!tgl
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
            grid.TextMatrix(grid.Row, 4) = Format(RST!qty_confirmed, "###,###,##0.00")
            grid.TextMatrix(grid.Row, 7) = RST!keterangan
            grid.TextMatrix(grid.Row, 8) = RST!Status
            grid.TextMatrix(grid.Row, 9) = RST!flag
            grid.Rows = grid.Rows + 1
            grid.Row = grid.Row + 1
            RST.MoveNext
        Loop
    Else
        MsgBox "Data not Found.", vbExclamation, AppName
        txtnolot = ""
        lotedit = ""
    End If
    OBJ.Close
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    Dim stokbahan As Double
    Dim hppbahan As Double
    If KeyAscii = 13 Then
        grid.TextMatrix(grid.Row, 3) = txtnilai.text
        'grid.TextMatrix(grid.Row, 4) = "Pcs"
        grid.SetFocus
    End If
End Sub

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
End Sub

Private Sub hapusgrid()
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        grid.TextMatrix(grid.Row, 0) = ""
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
    grid.ColWidth(0) = 300
    grid.ColWidth(1) = 1100
    grid.ColWidth(2) = 3000
    grid.ColWidth(3) = 1000
    grid.ColWidth(4) = 1000
    grid.ColWidth(5) = 0
    grid.ColWidth(6) = 0
    grid.ColWidth(7) = 1500
    grid.ColWidth(8) = 0 '300
    grid.ColWidth(9) = 0 '300
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

Private Sub txtnolot_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtnolot = "" Then Exit Sub
        OBJ1.Open dsn
        SQL1 = "Select a.*,b.NamaBarang,c.NamaSatuan From am_gudang_permintaan a"
        SQL1 = SQL1 + " inner join am_apitemmst b on a.kodebarang = b.KodeBarang"
        SQL1 = SQL1 + " inner join am_apunit c on b.KodeSatuan = c.KodeSatuan"
        SQL1 = SQL1 + " Where a.user_confirmed IS NULL and a.nolot = '" & txtnolot & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        
        If RST1.EOF Then
            OBJ1.Close
            MsgBox "Lot Request not found", vbCritical, AppName
            Exit Sub
        End If
        lotedit = txtnolot
        hapusgrid
        lblTgl = "Tanggal : " & Format(RST1!tgl, "dd MMMM yyyy")
        date1 = RST1!tgl
        txtpetugas = RST1!petugas
        Do Until RST1.EOF
            grid.Col = 0
            Set grid.CellPicture = uncheck
            grid.TextMatrix(grid.Row, 1) = RST1!kodebarang
            grid.TextMatrix(grid.Row, 2) = RST1!namabarang
            If RST1!keterangan = "Request" Then
                grid.TextMatrix(grid.Row, 3) = Format(RST1!qty, "###,###,##0.00")
            Else
                If RST1!Status = "0" Then
                    grid.TextMatrix(grid.Row, 3) = Format(RST1!qty_add, "###,###,##0.00")
                Else
                    grid.TextMatrix(grid.Row, 3) = Format(RST1!qty_add, "###,###,##0.00") * -1
                End If
            End If
            grid.TextMatrix(grid.Row, 4) = Format(RST1!qty_confirmed, "###,###,##0.00")
            grid.TextMatrix(grid.Row, 7) = RST1!keterangan
            grid.TextMatrix(grid.Row, 8) = RST1!Status
            grid.TextMatrix(grid.Row, 9) = RST1!flag
            grid.Rows = grid.Rows + 1
            grid.Row = grid.Row + 1
            RST1.MoveNext
        Loop
        cmdsave.Caption = "Update"
        OBJ1.Close
    End If
End Sub
