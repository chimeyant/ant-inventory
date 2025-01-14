VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Begin VB.Form frmkurs 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Currency"
   ClientHeight    =   5370
   ClientLeft      =   5715
   ClientTop       =   5565
   ClientWidth     =   6015
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
   Icon            =   "frmkurs.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Calculator      =   "frmkurs.frx":2372
      Caption         =   "frmkurs.frx":2392
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmkurs.frx":23FE
      Keys            =   "frmkurs.frx":241C
      Spin            =   "frmkurs.frx":245E
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
      ValueVT         =   103350277
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
   Begin VB.CheckBox chkbase 
      Caption         =   "Base Currency"
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox txtsymbol 
      Appearance      =   0  'Flat
      DataField       =   "NamaArea"
      Height          =   285
      Left            =   1800
      MaxLength       =   5
      TabIndex        =   2
      Top             =   1920
      Width           =   3495
   End
   Begin TDBText6Ctl.TDBText txtkode 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   1200
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   503
      Caption         =   "frmkurs.frx":2486
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmkurs.frx":24F2
      Key             =   "frmkurs.frx":2510
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
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   0
      ErrorBeep       =   0
      MaxLength       =   4
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
   Begin VB.TextBox txtNama 
      Appearance      =   0  'Flat
      DataField       =   "NamaArea"
      Height          =   285
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1560
      Width           =   3495
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   4800
      TabIndex        =   7
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
      MICON           =   "frmkurs.frx":254C
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
      TabIndex        =   6
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
      MICON           =   "frmkurs.frx":2866
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdadd 
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   4920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
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
      MICON           =   "frmkurs.frx":2B80
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Adding"
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
      TabIndex        =   13
      Top             =   0
      Width           =   2655
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
      TabIndex        =   12
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "Symbol Currency"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   1950
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Nama Currency"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1590
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Kode Currency"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1230
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   10335
   End
End
Attribute VB_Name = "frmkurs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim posrow As String

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

Private Sub cmdadd_Click()
    If txtKode = "" Or txtNama = "" Or txtsymbol = "" Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If Len(Trim(txtKode)) = 0 Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    txtKode = Trim(txtKode)

    OBJ.Open dsn
    SQL = "select * from gl_kurs where kdkurs = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Can't Add, Type " & txtKode & " Already Exsist.", vbInformation, "Information"
        cmdclear_Click
        OBJ.Close
        Exit Sub
    End If
    
    SQL = "insert into gl_kurs"
    SQL = SQL + "(Kdkurs"
    SQL = SQL + ",Nmkurs"
    SQL = SQL + ",symkurs"
    SQL = SQL + ",kurs1"
    SQL = SQL + ",kurs2"
    SQL = SQL + ",kurs3"
    SQL = SQL + ",kurs4"
    SQL = SQL + ",kurs5"
    SQL = SQL + ",kurs6"
    SQL = SQL + ",kurs7"
    SQL = SQL + ",kurs8"
    SQL = SQL + ",kurs9"
    SQL = SQL + ",kurs10"
    SQL = SQL + ",kurs11"
    SQL = SQL + ",kurs12"
    SQL = SQL + ",base"
    SQL = SQL + ",idupdate"
    SQL = SQL + ",dateupdate"
    SQL = SQL + ",identry"
    SQL = SQL + ",Dateentry)"
    
    SQL = SQL + "VALUES"
    SQL = SQL + "('" & txtKode & "'"
    SQL = SQL + ", '" & txtNama & "'"
    SQL = SQL + ", '" & txtsymbol & "'"
    SQL = SQL + ", convert(money,'" & Format(grid.TextMatrix(1, 1), "general number") & "')"
    SQL = SQL + ", convert(money,'" & Format(grid.TextMatrix(2, 1), "general number") & "')"
    SQL = SQL + ", convert(money,'" & Format(grid.TextMatrix(3, 1), "general number") & "')"
    SQL = SQL + ", convert(money,'" & Format(grid.TextMatrix(4, 1), "general number") & "')"
    SQL = SQL + ", convert(money,'" & Format(grid.TextMatrix(5, 1), "general number") & "')"
    SQL = SQL + ", convert(money,'" & Format(grid.TextMatrix(6, 1), "general number") & "')"
    SQL = SQL + ", convert(money,'" & Format(grid.TextMatrix(1, 4), "general number") & "')"
    SQL = SQL + ", convert(money,'" & Format(grid.TextMatrix(2, 4), "general number") & "')"
    SQL = SQL + ", convert(money,'" & Format(grid.TextMatrix(3, 4), "general number") & "')"
    SQL = SQL + ", convert(money,'" & Format(grid.TextMatrix(4, 4), "general number") & "')"
    SQL = SQL + ", convert(money,'" & Format(grid.TextMatrix(5, 4), "general number") & "')"
    SQL = SQL + ", convert(money,'" & Format(grid.TextMatrix(6, 4), "general number") & "')"
    SQL = SQL + ", '" & chkbase.Value & "'"
    SQL = SQL + ", ' '"
    SQL = SQL + ", convert(datetime,'" & tanggalsekarang & "')"
    SQL = SQL + ", '" & kuser & "'"
    SQL = SQL + ", convert(datetime,'" & tanggalsekarang & "'))"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdclear_Click()
    txtKode = ""
    txtNama = ""
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
    txtsymbol = ""
    chkbase.Enabled = True
    chkbase.Value = 0
    txtKode.SetFocus
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function

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

Private Sub txtKode_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtNama.SetFocus
End Sub

Private Sub txtKode_LostFocus()
    If txtKode = "" Then Exit Sub
    If txtKode.SelLength <> 0 Then Exit Sub
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
        
        MsgBox "Type " & txtKode & " Already Exsist.", vbInformation, "Information"
        txtKode.SetFocus
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
        
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
    txtNama = ""
    txtsymbol = ""
    chkbase.Value = 0
    txtNama.SetFocus
End Sub

Private Sub txtNama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtsymbol.SetFocus
End Sub
