VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmmodul 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Modul Aplikasi"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7065
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrDept 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   840
      Top             =   1755
   End
   Begin VB.Timer tmrSelected 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   345
      Top             =   1770
   End
   Begin VB.Frame Frame1 
      Caption         =   "Form Modul"
      Height          =   5865
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   6780
      Begin VB.ComboBox cmbDept 
         Height          =   315
         Left            =   1455
         TabIndex        =   11
         Top             =   300
         Width           =   3105
      End
      Begin TrueOleDBGrid70.TDBGrid DBGrid 
         Height          =   3480
         Left            =   105
         TabIndex        =   9
         Top             =   2235
         Width           =   6510
         _ExtentX        =   11483
         _ExtentY        =   6138
         _LayoutType     =   0
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   953
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   15790320
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   15790320
         RowDividerColor =   15790320
         RowSubDividerColor=   15790320
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=12,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Named:id=33:Normal"
         _StyleDefs(45)  =   ":id=33,.parent=0"
         _StyleDefs(46)  =   "Named:id=34:Heading"
         _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(48)  =   ":id=34,.wraptext=-1"
         _StyleDefs(49)  =   "Named:id=35:Footing"
         _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(51)  =   "Named:id=36:Selected"
         _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(53)  =   "Named:id=37:Caption"
         _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(55)  =   "Named:id=38:HighlightRow"
         _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=39:EvenRow"
         _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HC0C0C0&"
         _StyleDefs(59)  =   "Named:id=40:OddRow"
         _StyleDefs(60)  =   ":id=40,.parent=33"
         _StyleDefs(61)  =   "Named:id=41:RecordSelector"
         _StyleDefs(62)  =   ":id=41,.parent=34"
         _StyleDefs(63)  =   "Named:id=42:FilterBar"
         _StyleDefs(64)  =   ":id=42,.parent=33"
      End
      Begin VB.Frame Frame2 
         Height          =   750
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   6525
         Begin VB.CommandButton cmdSimpan 
            Caption         =   "Simpan"
            Height          =   345
            Left            =   2940
            TabIndex        =   8
            Top             =   270
            Width           =   1110
         End
         Begin VB.CommandButton cmdHapus 
            Caption         =   "Hapus"
            Height          =   345
            Left            =   4110
            TabIndex        =   7
            Top             =   255
            Width           =   1110
         End
         Begin VB.CommandButton cmdSelesai 
            Caption         =   "Selesai"
            Height          =   345
            Left            =   5295
            TabIndex        =   6
            Top             =   240
            Width           =   1110
         End
      End
      Begin VB.TextBox txtNamaModul 
         Height          =   345
         Left            =   1440
         TabIndex        =   4
         Top             =   1020
         Width           =   3105
      End
      Begin VB.TextBox txtKodeModul 
         Height          =   345
         Left            =   1440
         TabIndex        =   3
         Top             =   645
         Width           =   3105
      End
      Begin VB.Label Label3 
         Caption         =   "Departemen"
         Height          =   210
         Left            =   165
         TabIndex        =   10
         Top             =   330
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "Nama Modul"
         Height          =   210
         Left            =   195
         TabIndex        =   2
         Top             =   1140
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Kode Modul "
         Height          =   210
         Left            =   180
         TabIndex        =   1
         Top             =   735
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmModul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program Name   : Exclusive Inventory Technology System
'Alias          : EI-Tech System
'Copyright      : 2012
'Company        : SPARTA PRIMA
'Programmer     : U. Selamat Raharja & Adnan

Private RS As ADODB.Recordset
Private kode_dept As String

Private Sub cmbDept_Click()
    tmrDept.Enabled = True
End Sub

Private Sub cmdHapus_Click()
    Dim knf As String
    knf = MsgBox("Apakah anda yakin akan menghapus modul : " + txtNamaModul.text, vbOKCancel, AppName)
    
    If knf = vbOK Then
        DoHapus
    Else
        DoClearForm
    End If
End Sub

Private Sub cmdSelesai_Click()
    Unload Me
End Sub

Private Sub OpenMSSQLDB()
    On Error GoTo err_handler
    OpenDB
    Exit Sub
err_handler:
    MsgBox "Tidak dapat terhubung dengan server..!. " + Err.Description, vbCritical, "Warning"
End Sub

Private Sub CloseMSSQLDB()
    If ConSQL.State <> 0 Then
        CloseSQLDB
    End If
End Sub

Private Sub DoLoadDepartemen()
   ' On Error GoTo err_handler
    Dim SQL As String
    
    SQL = "select * from am_apdepartemen"
    OpenMSSQLDB
    Set RS = ConSQL.Execute(SQL)
    cmbDept.Clear
    Do While Not RS.EOF
        cmbDept.AddItem RS!dept
        RS.MoveNext
    Loop
    CloseMSSQLDB
    cmbDept.text = "MASTER"
    Exit Sub
err_handler:
    MsgBox "Gagal membuka data departemen..!. " + Err.Description, vbCritical, "Warning"
End Sub


Private Sub DoLoadModul()
    On Error GoTo err_handler
    Dim SQL As String
    SQL = "select * from am_apmodul where dept ='" + cmbDept.text + "'"
    
    OpenMSSQLDB
    Set RS = New ADODB.Recordset
    RS.Open SQL, ConSQL, adOpenDynamic, adLockReadOnly
    Set DBGrid.DataSource = Display_Data(RS)
    CloseMSSQLDB
    DoSetDBGrid
    Exit Sub
err_handler:
    MsgBox "Gagal membuka data modul...!." + Err.Description, vbCritical, "Warning"
End Sub

Private Sub DoSetDBGrid()
    With DBGrid
        .Columns(0).HeadAlignment = dbgCenter
        .Columns(0).Caption = "KD.Dept"
        .Columns(0).Width = 500
        .Columns(1).HeadAlignment = dbgCenter
        .Columns(1).Caption = "KD.Mod"
        .Columns(1).Width = 500
        .Columns(2).HeadAlignment = dbgCenter
        .Columns(2).Caption = "Departemen"
        .Columns(2).Width = 1500
        .Columns(3).HeadAlignment = dbgCenter
        .Columns(3).Caption = "Modul"
        .Columns(3).Width = 4000
    End With
    DBGrid.MoveLast
    DBGrid.Scroll 0, DBGrid.Row
End Sub


Private Sub cmdSimpan_Click()
    If txtKodeModul.text <> "" And txtNamaModul.text <> "" Then
        DoSimpan
    Else
        Exit Sub
    End If
End Sub

Private Sub DBGrid_Click()
    tmrSelected.Enabled = True
End Sub

Private Sub Form_Load()
    DoLoadDepartemen
    DoLoadModul
End Sub

Private Sub DoSimpan()
    On Error GoTo err_msg
    Dim SQL As String
    SQL = "select kode_dept from am_apdepartemen where dept ='" + Trim(cmbDept.text) + "'"
    OpenMSSQLDB
    Set RS = ConSQL.Execute(SQL)
    kode_dept = RS!kode_dept
    
    SQL = "insert into am_apmodul values('" + kode_dept + "','" + txtKodeModul.text + "','" + cmbDept.text + "','" + txtNamaModul.text + "')"
    ConSQL.Execute SQL
    CloseMSSQLDB
    DoLoadModul
    DoClearForm
    Exit Sub
err_msg:
    MsgBox "Gagal Menyimpan kemungkinan data telah ada...!." + Err.Description, vbCritical, "Warning"
End Sub

Private Sub DoClearForm()
    txtKodeModul.text = ""
    txtNamaModul.text = ""
    txtKodeModul.SetFocus
End Sub

Private Sub tmrDept_Timer()
    tmrDept.Enabled = False
    DoLoadModul
End Sub

Private Sub tmrSelected_Timer()
    On Error GoTo err_handler
    tmrSelected.Enabled = False
    Dim SQL As String
    Dim dept As String
    
    SQL = "select dept from am_apdepartemen where kode_dept='" + DBGrid.Columns(0).Value + "'"
    OpenMSSQLDB
    Set RS = ConSQL.Execute(SQL)
    dept = RS!dept
    CloseMSSQLDB
    
    cmbDept.text = dept
    kode_dept = DBGrid.Columns(0).Value
    txtKodeModul.text = DBGrid.Columns(1).Value
    txtNamaModul.text = DBGrid.Columns(3).Value
    Exit Sub
err_handler:
    MsgBox "Gagal membuka data departemen...!. " + Err.Description, vbCritical, "Warning"
End Sub

Private Sub DoHapus()
    On Error GoTo err_handler
    Dim SQL As String
    SQL = "delete from am_apmodul where kode_dept='" + kode_dept + "' and kode_modul='" + txtKodeModul.text + "'"
    
    OpenMSSQLDB
    ConSQL.Execute SQL
    CloseMSSQLDB
    
    DoLoadModul
    DoClearForm
    Exit Sub
err_handler:
    MsgBox "Data tidak berhasil dihapus....!. " + Err.Description
End Sub

Private Sub txtKodeModul_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaModul.SetFocus
End Sub

Private Sub txtNamaModul_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan_Click
End Sub
