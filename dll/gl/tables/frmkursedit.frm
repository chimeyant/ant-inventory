VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Begin VB.Form frmkursedit 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Currency"
   ClientHeight    =   5400
   ClientLeft      =   5715
   ClientTop       =   5565
   ClientWidth     =   5985
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmkursedit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkbase 
      Caption         =   "Base Currency"
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
      Left            =   1800
      TabIndex        =   11
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox txtsymbol 
      Appearance      =   0  'Flat
      DataField       =   "NamaArea"
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
      Left            =   1800
      MaxLength       =   5
      TabIndex        =   2
      Top             =   1920
      Width           =   3495
   End
   Begin VB.TextBox txtKode 
      Appearance      =   0  'Flat
      DataField       =   "KodeArea"
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
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   0
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txtNama 
      Appearance      =   0  'Flat
      DataField       =   "NamaArea"
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
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1560
      Width           =   3495
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   450
      Calculator      =   "frmkursedit.frx":2372
      Caption         =   "frmkursedit.frx":2392
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmkursedit.frx":23FE
      Keys            =   "frmkursedit.frx":241C
      Spin            =   "frmkursedit.frx":245E
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
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   2011758597
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   2175
      Left            =   360
      TabIndex        =   4
      Top             =   2640
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   3836
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   4800
      TabIndex        =   8
      Top             =   4920
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
      MICON           =   "frmkursedit.frx":2486
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
      Left            =   3840
      TabIndex        =   7
      Top             =   4920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
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
      MICON           =   "frmkursedit.frx":27A0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdupdate 
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   4920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
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
      MICON           =   "frmkursedit.frx":2ABA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdelete 
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   4920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Delete"
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
      MICON           =   "frmkursedit.frx":2DD4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch 
      Height          =   285
      Left            =   360
      TabIndex        =   15
      Top             =   1200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Kode Currency"
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
      MICON           =   "frmkursedit.frx":30EE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Currency"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Updating"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "Nama Currency"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   1590
      Width           =   1455
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      Caption         =   "Symbol Currency"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1950
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   10335
   End
End
Attribute VB_Name = "frmkursedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim posrow As String

Private Sub cmdclear_Click()
    txtKode.Enabled = True
    cmdsearch.Enabled = True
    txtKode = ""
    txtNama = ""
    txtsymbol = ""
    grid.TextMatrix(1, 1) = "0.00"
    grid.TextMatrix(2, 1) = "0.00"
    grid.TextMatrix(3, 1) = "0.00"
    grid.TextMatrix(4, 1) = "0.00"
    grid.TextMatrix(5, 1) = "0.00"
    grid.TextMatrix(6, 1) = "0.00"
    grid.TextMatrix(1, 4) = "0.00"
    grid.TextMatrix(2, 4) = "0.00"
    grid.TextMatrix(3, 4) = "0.00"
    grid.TextMatrix(4, 4) = "0.00"
    grid.TextMatrix(5, 4) = "0.00"
    grid.TextMatrix(6, 4) = "0.00"
    chkbase.Value = 0
    chkbase.Enabled = True
    txtKode.SetFocus
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdelete_Click()
    If txtKode = "" Or txtNama = "" Or txtsymbol = "" Then
        MsgBox "Data Entry Not Complite", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If MsgBox("Are You Sure Want To Delete ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from gl_transaksi WHERE currtrx = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then GoTo jump1
    
    SQL = "DELETE FROM gl_kurs WHERE Kdkurs = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    
    MsgBox "Data Is Deleted, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
    OBJ.Close
    Exit Sub
    
jump1:
    OBJ.Close
    MsgBox "Can Not Delete, Record Still In Use.", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdsearch_Click()
    carisql1 = "select kdkurs, nmkurs from gl_kurs"
    namatabel = "Currency"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    txtKode = hasil
    caritype
    txtNama.SetFocus
    hasil = ""
End Sub

Private Sub cmdupdate_click()
    If txtKode = "" Or txtNama = "" Or txtsymbol = "" Then
        MsgBox "Data Entry Not Complite", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If MsgBox("Are You Sure Want To Update ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "UPDATE gl_kurs SET "
    SQL = SQL + "Nmkurs = '" & txtNama & "',"
    SQL = SQL + "symkurs = '" & txtsymbol & "',"
    SQL = SQL + "base = '" & chkbase.Value & "',"
    SQL = SQL + "kurs1 = convert(money,'" & Format(grid.TextMatrix(1, 1), "general number") & "'),"
    SQL = SQL + "kurs2 = convert(money,'" & Format(grid.TextMatrix(2, 1), "general number") & "'),"
    SQL = SQL + "kurs3 = convert(money,'" & Format(grid.TextMatrix(3, 1), "general number") & "'),"
    SQL = SQL + "kurs4 = convert(money,'" & Format(grid.TextMatrix(4, 1), "general number") & "'),"
    SQL = SQL + "kurs5 = convert(money,'" & Format(grid.TextMatrix(5, 1), "general number") & "'),"
    SQL = SQL + "kurs6 = convert(money,'" & Format(grid.TextMatrix(6, 1), "general number") & "'),"
    SQL = SQL + "kurs7 = convert(money,'" & Format(grid.TextMatrix(1, 4), "general number") & "'),"
    SQL = SQL + "kurs8 = convert(money,'" & Format(grid.TextMatrix(2, 4), "general number") & "'),"
    SQL = SQL + "kurs9 = convert(money,'" & Format(grid.TextMatrix(3, 4), "general number") & "'),"
    SQL = SQL + "kurs10 = convert(money,'" & Format(grid.TextMatrix(4, 4), "general number") & "'),"
    SQL = SQL + "kurs11 = convert(money,'" & Format(grid.TextMatrix(5, 4), "general number") & "'),"
    SQL = SQL + "kurs12 = convert(money,'" & Format(grid.TextMatrix(6, 4), "general number") & "'),"
    SQL = SQL + "idUpdate = '" & kuser & "',"
    SQL = SQL + "DateUpdate = convert(datetime,'" & tanggalsekarang & "')"
    SQL = SQL + "WHERE Kdkurs =  '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Data Is Updated, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()

    
    grid.TextMatrix(0, 0) = "Bulan"
    grid.TextMatrix(0, 1) = "Nilai Kurs"
    grid.TextMatrix(0, 3) = "Bulan"
    grid.TextMatrix(0, 4) = "Nilai Kurs"
    grid.ColWidth(0) = 1000
    grid.ColWidth(1) = 1500
    grid.ColWidth(2) = 200
    grid.ColWidth(3) = 1000
    grid.ColWidth(4) = 1500
    
    grid.RowHeightMin = 300
    grid.Rows = 7
    grid.TextMatrix(1, 0) = "Januari"
    grid.TextMatrix(1, 1) = "0.00"
    grid.TextMatrix(1, 3) = "Juli"
    grid.TextMatrix(1, 4) = "0.00"
    grid.TextMatrix(2, 0) = "Februari"
    grid.TextMatrix(2, 1) = "0.00"
    grid.TextMatrix(2, 3) = "Agustus"
    grid.TextMatrix(2, 4) = "0.00"
    grid.TextMatrix(3, 0) = "Maret"
    grid.TextMatrix(3, 1) = "0.00"
    grid.TextMatrix(3, 3) = "September"
    grid.TextMatrix(3, 4) = "0.00"
    grid.TextMatrix(4, 0) = "April"
    grid.TextMatrix(4, 1) = "0.00"
    grid.TextMatrix(4, 3) = "Oktober"
    grid.TextMatrix(4, 4) = "0.00"
    grid.TextMatrix(5, 0) = "Mei"
    grid.TextMatrix(5, 1) = "0.00"
    grid.TextMatrix(5, 3) = "November"
    grid.TextMatrix(5, 4) = "0.00"
    grid.TextMatrix(6, 0) = "Juni"
    grid.TextMatrix(6, 1) = "0.00"
    grid.TextMatrix(6, 3) = "Desember"
    grid.TextMatrix(6, 4) = "0.00"
End Sub

Private Sub chkbase_Click()
    If chkbase.Value = 1 Then
        grid.TextMatrix(1, 1) = "1.00"
        grid.TextMatrix(2, 1) = "1.00"
        grid.TextMatrix(3, 1) = "1.00"
        grid.TextMatrix(4, 1) = "1.00"
        grid.TextMatrix(5, 1) = "1.00"
        grid.TextMatrix(6, 1) = "1.00"
        grid.TextMatrix(1, 4) = "1.00"
        grid.TextMatrix(2, 4) = "1.00"
        grid.TextMatrix(3, 4) = "1.00"
        grid.TextMatrix(4, 4) = "1.00"
        grid.TextMatrix(5, 4) = "1.00"
        grid.TextMatrix(6, 4) = "1.00"
    Else
        grid.TextMatrix(1, 1) = "0.00"
        grid.TextMatrix(2, 1) = "0.00"
        grid.TextMatrix(3, 1) = "0.00"
        grid.TextMatrix(4, 1) = "0.00"
        grid.TextMatrix(5, 1) = "0.00"
        grid.TextMatrix(6, 1) = "0.00"
        grid.TextMatrix(1, 4) = "0.00"
        grid.TextMatrix(2, 4) = "0.00"
        grid.TextMatrix(3, 4) = "0.00"
        grid.TextMatrix(4, 4) = "0.00"
        grid.TextMatrix(5, 4) = "0.00"
        grid.TextMatrix(6, 4) = "0.00"
    End If
End Sub

Private Sub caritype()
    If txtKode = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from gl_kurs where base = '1'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then chkbase.Enabled = False
    If RST.EOF Then chkbase.Enabled = True
    
    SQL = "select * from gl_kurs where kdkurs = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtNama = RST!nmkurs
        txtsymbol = RST!symkurs
        chkbase.Value = RST!base
        If chkbase.Value = 1 Then chkbase.Enabled = True
        grid.TextMatrix(1, 1) = Format(RST!kurs1, "###,###,##0.00")
        grid.TextMatrix(2, 1) = Format(RST!kurs2, "###,###,##0.00")
        grid.TextMatrix(3, 1) = Format(RST!kurs3, "###,###,##0.00")
        grid.TextMatrix(4, 1) = Format(RST!kurs4, "###,###,##0.00")
        grid.TextMatrix(5, 1) = Format(RST!kurs5, "###,###,##0.00")
        grid.TextMatrix(6, 1) = Format(RST!kurs6, "###,###,##0.00")
        grid.TextMatrix(1, 4) = Format(RST!kurs7, "###,###,##0.00")
        grid.TextMatrix(2, 4) = Format(RST!kurs8, "###,###,##0.00")
        grid.TextMatrix(3, 4) = Format(RST!kurs9, "###,###,##0.00")
        grid.TextMatrix(4, 4) = Format(RST!kurs10, "###,###,##0.00")
        grid.TextMatrix(5, 4) = Format(RST!kurs11, "###,###,##0.00")
        grid.TextMatrix(6, 4) = Format(RST!kurs12, "###,###,##0.00")
        
        txtKode.Enabled = False
        cmdsearch.Enabled = False
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    MsgBox "Type " & txtKode & " Not Found.", vbInformation, "Information"
    cmdclear_Click
End Sub

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    If txtKode = "" Or chkbase.Value = 1 Then Exit Sub
    posrow = grid.Row
    Select Case grid.Col
        Case 1, 4
            If txtnilai.Visible = True Then Exit Sub
            
            txtnilai.Width = grid.ColWidth(grid.Col) - 40
            txtnilai = grid.TextMatrix(grid.Row, grid.Col)
            txtnilai.Left = grid.Left + grid.CellLeft
            txtnilai.Top = grid.Top + grid.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
    End Select
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grid.TextMatrix(grid.Row, grid.Col) = Format(txtnilai, "###,###,##0.00")
        txtnilai = 0
        grid.SetFocus
        grid.Row = posrow
    End If
End Sub

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
End Sub

Private Sub grid_EnterCell()
    Select Case grid.Col
    Case 1, 4
        If chkbase.Value = 1 Then Exit Sub
            
        posrow = grid.Row
        txtnilai.Width = grid.ColWidth(grid.Col) - 40
        txtnilai = grid.TextMatrix(grid.Row, grid.Col)
        txtnilai.Left = grid.Left + grid.CellLeft
        txtnilai.Top = grid.Top + grid.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
    End Select
End Sub

Private Sub txtKode_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtNama.SetFocus
End Sub

Private Sub txtKode_LostFocus()
    caritype
End Sub

Private Sub txtNama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtsymbol.SetFocus
End Sub
