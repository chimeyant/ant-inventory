VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmuser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage User"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Form Manage User"
      Height          =   6465
      Left            =   60
      TabIndex        =   0
      Top             =   210
      Width           =   7080
      Begin VB.CheckBox Check1 
         Caption         =   "Show"
         Height          =   375
         Left            =   5640
         TabIndex        =   28
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame Frame5 
         Caption         =   "Departemen"
         Height          =   945
         Left            =   75
         TabIndex        =   19
         Top             =   2385
         Width           =   6900
         Begin VB.CheckBox chksale 
            Caption         =   "Sale"
            Height          =   240
            Left            =   120
            TabIndex        =   27
            Top             =   600
            Width           =   915
         End
         Begin VB.CheckBox chkledger 
            Caption         =   "Ledger"
            Height          =   240
            Left            =   2295
            TabIndex        =   26
            Top             =   615
            Width           =   1215
         End
         Begin VB.CheckBox chkfinance 
            Caption         =   "Finance"
            Height          =   240
            Left            =   1275
            TabIndex        =   25
            Top             =   615
            Width           =   1215
         End
         Begin VB.CheckBox chkwarehouse 
            Caption         =   "Warehouse"
            Height          =   240
            Left            =   5235
            TabIndex        =   24
            Top             =   285
            Width           =   1215
         End
         Begin VB.CheckBox chkprod 
            Caption         =   "Production"
            Height          =   240
            Left            =   3600
            TabIndex        =   23
            Top             =   285
            Width           =   1215
         End
         Begin VB.CheckBox chkpurc 
            Caption         =   "Purchasing"
            Height          =   240
            Left            =   2295
            TabIndex        =   22
            Top             =   300
            Width           =   1215
         End
         Begin VB.CheckBox chkhrd 
            Caption         =   "HRD"
            Height          =   240
            Left            =   1275
            TabIndex        =   21
            Top             =   300
            Width           =   705
         End
         Begin VB.CheckBox chkmaster 
            Caption         =   "Master"
            Height          =   240
            Left            =   120
            TabIndex        =   20
            Top             =   300
            Width           =   1215
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Level User"
         Height          =   630
         Left            =   75
         TabIndex        =   15
         Top             =   1770
         Width           =   6900
         Begin VB.OptionButton optoperator 
            Caption         =   "Operator"
            Height          =   300
            Left            =   3015
            TabIndex        =   18
            Top             =   255
            Value           =   -1  'True
            Width           =   1260
         End
         Begin VB.OptionButton optsupervisor 
            Caption         =   "Supervisor"
            Height          =   300
            Left            =   1680
            TabIndex        =   17
            Top             =   240
            Width           =   1260
         End
         Begin VB.OptionButton optAdm 
            Caption         =   "Administrator"
            Height          =   300
            Left            =   120
            TabIndex        =   16
            Top             =   225
            Width           =   1260
         End
      End
      Begin VB.TextBox txtPasswordB 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1950
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1350
         Width           =   3615
      End
      Begin VB.Frame Frame3 
         Height          =   720
         Left            =   75
         TabIndex        =   9
         Top             =   3300
         Width           =   6900
         Begin VB.CommandButton btnSimpan 
            Caption         =   "Simpan"
            Height          =   495
            Left            =   3870
            TabIndex        =   12
            Top             =   135
            Width           =   945
         End
         Begin VB.CommandButton cmdHapus 
            Caption         =   "Hapus"
            Height          =   495
            Left            =   4815
            TabIndex        =   11
            Top             =   135
            Width           =   975
         End
         Begin VB.CommandButton cmdSelesai 
            Caption         =   "Selesai"
            Height          =   495
            Left            =   5805
            TabIndex        =   10
            Top             =   135
            Width           =   1005
         End
      End
      Begin VB.TextBox txtPasswordA 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1950
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   975
         Width           =   3615
      End
      Begin VB.TextBox txtFullname 
         Height          =   330
         Left            =   1950
         TabIndex        =   2
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox txtUsername 
         Height          =   330
         Left            =   1950
         TabIndex        =   1
         Top             =   255
         Width           =   3615
      End
      Begin VB.Frame Frame2 
         Caption         =   "User List"
         Height          =   2340
         Left            =   75
         TabIndex        =   8
         Top             =   3960
         Width           =   6870
         Begin TrueOleDBGrid70.TDBGrid DBGrid 
            Height          =   1965
            Left            =   90
            TabIndex        =   13
            Top             =   270
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   3466
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
      Begin VB.Label Label4 
         Caption         =   "Ulang Password/Sandi"
         Height          =   240
         Left            =   240
         TabIndex        =   14
         Top             =   1395
         Width           =   1680
      End
      Begin VB.Label Label3 
         Caption         =   "Password/Sandi"
         Height          =   240
         Left            =   240
         TabIndex        =   7
         Top             =   990
         Width           =   1680
      End
      Begin VB.Label Label2 
         Caption         =   "Nama Lengkap"
         Height          =   240
         Left            =   240
         TabIndex        =   6
         Top             =   660
         Width           =   1260
      End
      Begin VB.Label Label1 
         Caption         =   "Nama User"
         Height          =   240
         Left            =   240
         TabIndex        =   5
         Top             =   315
         Width           =   1680
      End
   End
End
Attribute VB_Name = "frmuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program Name   : Ant Inventory System
'Alias          : Ant Inventory System
'Copyright      : 2012 - 2013
'Company        : Antsoft Media
'Programmer1    : U. Selamat Raharja
'Programmer2    : Chandra Kirana

'declare variable database
Private RS As ADODB.Recordset
Private OBJ As New ADODB.Connection
Private SQL As String

Private Sub btnSimpan_Click()
    'On Error GoTo err_msg
    
    'declare variable
    Dim var_administrator, var_supervisor, var_operator As String
    Dim var_master, var_hrd, var_purc, var_prod, var_warehouse, var_sale, var_finance, var_ledger As String
    
    'initial variable
    var_administrator = "0"
    var_supervisor = "0"
    var_operator = "0"

    If optAdm = True Then
        var_administrator = "1"
    End If
    
    If optsupervisor = True Then
        var_supervisor = "1"
    End If
    
    If optoperator = True Then
        var_operator = "1"
    End If
    
    var_master = "0"
    var_hrd = "0"
    var_purc = "0"
    var_prod = "0"
    var_warehouse = "0"
    var_sale = "0"
    var_finance = "0"
    var_ledger = "0"
    
    If chkmaster.Value = 1 Then
        var_master = "1"
    End If
    If chkhrd.Value = 1 Then
        var_hrd = "1"
    End If
    If chkpurc.Value = 1 Then
        var_purc = "1"
    End If
    If chkprod.Value = 1 Then
        var_prod = "1"
    End If
    If chkwarehouse.Value = 1 Then
        var_warehouse = "1"
    End If
    If chksale.Value = 1 Then
        var_sale = "1"
    End If
    If chkfinance.Value = 1 Then
        var_finance = "1"
    End If
    If chkledger.Value = 1 Then
        var_ledger = "1"
    End If
    
    If txtPasswordA <> txtPasswordB Then
        MsgBox "Password Not Match...!", vbCritical, AppName
        Exit Sub
    End If
    
    SQL = "insert into userlist (username,realname,pass,su,ad,sv,op,master,hrd,purch,prod,warehouse,sale,finance,gl) "
    SQL = SQL & " Values('"
    SQL = SQL & txtUsername & "','"
    SQL = SQL & txtFullname & "','"
    SQL = SQL & Cheap_Encrypt(txtPasswordA) & "','0','"
    'SQL = SQL & "'0','"
    SQL = SQL & var_administrator & "','"
    SQL = SQL & var_supervisor & "','"
    SQL = SQL & var_operator & "','"
    SQL = SQL & var_master & "','"
    SQL = SQL & var_hrd & "','"
    SQL = SQL & var_purc & "','"
    SQL = SQL & var_prod & "','"
    SQL = SQL & var_warehouse & "','"
    SQL = SQL & var_sale & "','"
    SQL = SQL & var_finance & "','"
    SQL = SQL & var_ledger & "')"
    
    OpenDB
    ConSQL.Execute SQL
    'MsgBox "Proses Simpan berhasil...!", vbInformation, AppName
    'ClearForm
    CloseSQLDB
    
    SQL = "insert into LIST_USERS (username,realname,pass,kode_role,master,hrd,purch,prod,warehouse,sale,finance,gl,flag_aktip) "
    SQL = SQL & " Values('"
    SQL = SQL & txtUsername & "','"
    SQL = SQL & txtFullname & "','"
    SQL = SQL & Cheap_Encrypt(txtPasswordA) & "','"
    If optAdm.Value = True Then
        SQL = SQL & "03" & "','"
    ElseIf optsupervisor.Value = True Then
        SQL = SQL & "02" & "','"
    ElseIf optoperator.Value = True Then
        SQL = SQL & "01" & "','"
    End If
    SQL = SQL & var_master & "','"
    SQL = SQL & var_hrd & "','"
    SQL = SQL & var_purc & "','"
    SQL = SQL & var_prod & "','"
    SQL = SQL & var_warehouse & "','"
    SQL = SQL & var_sale & "','"
    SQL = SQL & var_finance & "','"
    SQL = SQL & var_ledger & "','0')"
    'SQL = SQL & "'0')"
    OpenDB
    ConSQL.Execute SQL
    MsgBox "Proses Simpan berhasil...!", vbInformation, AppName
    ClearForm
    CloseSQLDB
    DoLoadUserList
    Exit Sub
err_msg:
    MsgBox Err.Description, vbCritical, AppName
End Sub

Private Sub Check1_Click()
    If Check1.Value = Checked Then
        txtPasswordA.PasswordChar = ""
    ElseIf Check1.Value = Unchecked Then
        txtPasswordA.PasswordChar = "*"
    End If
End Sub

Private Sub cmdHapus_Click()
    Dim knf As Integer
    knf = MsgBox("Apakah anda yakin akan menhapus user : " + txtUsername, vbOKCancel, AppName)
    If knf = vbOK Then
        DoHapus
    Else
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    If nmuser = "Creator" Then Check1.Visible = True
    DoLoadUserList
End Sub

Private Sub cmdSelesai_Click()
    Unload Me
End Sub

Private Sub OpenMSSQLDB()
    OpenDB
End Sub

Private Sub CloseMSSQLDB()
    If ConSQL.State <> 0 Then
        CloseSQLDB
    End If
End Sub

Private Function DoLoadUserList()
    On Error GoTo err_handler
    'declare variable
    Dim SQL As String
    
    'ini variale
    OpenMSSQLDB
    SQL = "select * from userlist order by username"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, ConSQL, adOpenDynamic, adLockOptimistic
    
    Set DBGrid.DataSource = Display_Data(RS)
    
    'Load Set DBGrid
    DoSetDBGrid
    CloseMSSQLDB
    Exit Function
err_handler:
    MsgBox "Gagal melakukan pengambilan data user.....!. " + Err.Description, vbCritical, "Warning"
End Function

Private Function DoSetDBGrid()
    With DBGrid
        .Columns(0).HeadAlignment = dbgCenter
        .Columns(0).Caption = "UserName"
        .Columns(0).Width = 1500
        .Columns(1).HeadAlignment = dbgCenter
        .Columns(1).Caption = "Fullname"
        .Columns(1).Width = 3000
        .Columns(2).HeadAlignment = dbgCenter
        .Columns(2).Caption = "Password"
        .Columns(2).Width = 1500
    End With
End Function
    


Private Sub DoClearForm()
    txtUsername.text = ""
    txtFullname.text = ""
    txtPasswordA.text = ""
    txtPasswordB.text = ""
    txtUsername.SetFocus
End Sub
    

Private Sub txtFullname_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPasswordA.SetFocus
End Sub

Private Sub txtPasswordA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPasswordB.SetFocus
End Sub



Private Sub txtUsername_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then DoCekUser (txtUsername.text)
End Sub

'cek user visible
Private Function DoCekUser(ByVal username As String)
    On Error GoTo err_process
    Dim SQL As String
    SQL = "select * from list_users where username='" + username + "'"
    
    OpenMSSQLDB
    Set RS = New ADODB.Recordset
    RS.Open SQL, ConSQL, adOpenDynamic, adLockReadOnly
    
    txtFullname.text = RS!realname
    txtPasswordA.text = Cheap_Decrypt(RS!pass)
    txtPasswordB.text = txtPasswordA.text
    CloseSQLDB
    Exit Function
err_process:
    txtFullname.text = ""
    txtPasswordA.text = ""
    txtPasswordB.text = ""
    txtFullname.SetFocus
End Function

Private Sub DoHapus()
    On Error GoTo err_handler
    Dim SQL As String
    SQL = "delete from userlist where username='" + txtUsername.text + "'"
    OpenMSSQLDB
    ConSQL.Execute SQL
    CloseMSSQLDB
    DoClearForm
    DoLoadUserList
    Exit Sub
err_handler:
    MsgBox "Gagal melakukan penghapusan data....!. " + Err.Description, vbCritical, "Warning"
End Sub

Private Sub ClearForm()
    txtUsername = ""
    txtFullname = ""
    txtPasswordA = ""
    txtPasswordB = ""
    optAdm.Value = False
    optsupervisor.Value = False
    optoperator.Value = True
    chkmaster.Value = 0
    chkhrd.Value = 0
    chkpurc.Value = 0
    chkprod.Value = 0
    chkwarehouse.Value = 0
    chksale.Value = 0
    chkfinance.Value = 0
    chkledger.Value = 0
End Sub
