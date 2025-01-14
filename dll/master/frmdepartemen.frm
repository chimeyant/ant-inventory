VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmdepartemen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Departemen"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6750
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6750
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Form Departemen"
      Height          =   3900
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   6600
      Begin VB.Timer tmrDBGrid 
         Enabled         =   0   'False
         Interval        =   20
         Left            =   5760
         Top             =   360
      End
      Begin VB.Frame Frame3 
         Height          =   630
         Left            =   105
         TabIndex        =   6
         Top             =   1020
         Width           =   6255
         Begin VB.CommandButton cmdSimpan 
            Caption         =   "Simpan"
            Height          =   330
            Left            =   3075
            TabIndex        =   10
            Top             =   180
            Width           =   960
         End
         Begin VB.CommandButton cmdHapus 
            Caption         =   "Hapus"
            Height          =   330
            Left            =   4110
            TabIndex        =   9
            Top             =   180
            Width           =   960
         End
         Begin VB.CommandButton cmdSelesai 
            Caption         =   "Selesai"
            Height          =   330
            Left            =   5175
            TabIndex        =   8
            Top             =   180
            Width           =   960
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "List Departemen"
         Height          =   2145
         Left            =   105
         TabIndex        =   5
         Top             =   1650
         Width           =   6270
         Begin TrueOleDBGrid70.TDBGrid DBGrid 
            Height          =   1815
            Left            =   75
            TabIndex        =   7
            Top             =   210
            Width           =   6105
            _ExtentX        =   10769
            _ExtentY        =   3201
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
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
      End
      Begin VB.TextBox txtNamaDept 
         Height          =   300
         Left            =   1665
         TabIndex        =   4
         Top             =   690
         Width           =   4050
      End
      Begin VB.TextBox txtKodeDept 
         Height          =   300
         Left            =   1665
         TabIndex        =   3
         Top             =   375
         Width           =   4050
      End
      Begin VB.Label Label2 
         Caption         =   "Nama Departemen"
         Height          =   270
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "Kode Departemen "
         Height          =   270
         Left            =   135
         TabIndex        =   1
         Top             =   390
         Width           =   1545
      End
   End
End
Attribute VB_Name = "frmDepartemen"
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

Private Sub cmdHapus_Click()
    Dim knf As Integer
    If txtKodeDept.text <> "" Then
        knf = MsgBox("Apakah anda yakin akan menghapus Departemen : " + txtKodeDept.text, vbOKCancel, AppName)
        If knf = vbOK Then
            DoHapus
        End If
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
    MsgBox "Tidak dapat terkoneksi dengan server...!", vbCritical, "Warning"
End Sub

Private Sub CloseMSSQLDB()
    On Error Resume Next
    If ConSQL.State <> 0 Then
        CloseSQLDB
    End If
End Sub

Private Sub DoLoadDepartemen()
    On Error Resume Next
    Dim SQL As String
    SQL = "select * from am_apdepartemen"
    
    OpenMSSQLDB
    Set RS = ConSQL.Execute(SQL)
    Set DBGrid.DataSource = Display_Data(RS)
    CloseMSSQLDB
    DoSetDBGrid
End Sub

Private Sub DoSetDBGrid()
    With DBGrid
        .Columns(0).HeadAlignment = dbgCenter
        .Columns(0).Caption = "Kode Departemen"
        .Columns(0).Width = 2000
        .Columns(1).HeadAlignment = dbgCenter
        .Columns(1).Caption = "Departemen"
        .Columns(1).Width = 2000
    End With
End Sub

Private Sub cmdSimpan_Click()
    If txtKodeDept.text <> "" And txtNamaDept.text <> "" Then
        DoSimpan
    Else
        MsgBox "Kode dan Nama Departemen Tidak Boleh Kosong...!", vbCritical, AppName
        txtKodeDept.text = ""
        txtNamaDept.text = ""
        txtKodeDept.SetFocus
    End If
End Sub

Private Sub DBGrid_Click()
    tmrDBGrid.Enabled = True
End Sub

Private Sub Form_Load()
    DoLoadDepartemen
End Sub

Private Sub DoSimpan()
    On Error GoTo err_process
    Dim SQL As String
    
    SQL = "insert into am_apdepartemen values('" + txtKodeDept.text + "','" + txtNamaDept.text + "')"
    OpenMSSQLDB
    ConSQL.Execute (SQL)
    CloseMSSQLDB
    DoLoadDepartemen
    DoClearForm
    Exit Sub
err_process:
    MsgBox "Gagal Simpan, kemungkinan data telah ada..!", vbCritical, AppName
End Sub

Private Sub DoClearForm()
    txtKodeDept.text = ""
    txtNamaDept.text = ""
    txtKodeDept.SetFocus
End Sub

Private Sub tmrDBGrid_Timer()
    tmrDBGrid.Enabled = False
    txtKodeDept.text = DBGrid.Columns(0).Value
    txtNamaDept.text = DBGrid.Columns(1).Value
End Sub

Private Sub DoHapus()
    On Error GoTo err_handler
    Dim SQL As String
    SQL = "delete from am_apdepartemen where kode_dept ='" + txtKodeDept.text + "'"
    OpenMSSQLDB
    ConSQL.Execute SQL
    CloseMSSQLDB
    DoLoadDepartemen
    DoClearForm
    Exit Sub
err_handler:
    MsgBox "Data tidak berhasil dihapus..!", vbCritical, "Warning"
End Sub
