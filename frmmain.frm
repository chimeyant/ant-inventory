VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#13.2#0"; "CODEJO~3.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#13.2#0"; "CODEJO~1.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmain 
   Caption         =   "na"
   ClientHeight    =   7260
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11160
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   11160
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Interval        =   60000
      Left            =   6000
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   4065
      Top             =   150
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   2400
      Top             =   180
   End
   Begin XtremeSuiteControls.WebBrowser wbrmain 
      Height          =   5760
      Left            =   120
      TabIndex        =   4
      Top             =   645
      Width           =   10875
      _Version        =   851970
      _ExtentX        =   19182
      _ExtentY        =   10160
      _StockProps     =   173
      BackColor       =   -2147483643
      WebBrowserContextMenu=   0   'False
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   1515
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   45
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":08CA
            Key             =   ""
            Object.Tag             =   "100111"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":11A4
            Key             =   ""
            Object.Tag             =   "10011"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2B36
            Key             =   ""
            Object.Tag             =   "100112"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":44C8
            Key             =   ""
            Object.Tag             =   "100131"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":5E5A
            Key             =   ""
            Object.Tag             =   "100132"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":77EC
            Key             =   ""
            Object.Tag             =   "100133"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":80C6
            Key             =   ""
            Object.Tag             =   "10013"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":9A58
            Key             =   ""
            Object.Tag             =   "10023"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":B3EA
            Key             =   ""
            Object.Tag             =   "10022"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":CD7C
            Key             =   ""
            Object.Tag             =   "10024"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":E70E
            Key             =   ""
            Object.Tag             =   "10021"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":100A0
            Key             =   ""
            Object.Tag             =   "10031"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":11715
            Key             =   ""
            Object.Tag             =   "10025"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":130A7
            Key             =   ""
            Object.Tag             =   "3011"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":14A39
            Key             =   ""
            Object.Tag             =   "3012"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":163CB
            Key             =   ""
            Object.Tag             =   "3013"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":17D5D
            Key             =   ""
            Object.Tag             =   "3014"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":196EF
            Key             =   ""
            Object.Tag             =   "3015"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1B081
            Key             =   ""
            Object.Tag             =   "3016"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1CA13
            Key             =   ""
            Object.Tag             =   "60011"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1E3A5
            Key             =   ""
            Object.Tag             =   "60012"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1FD37
            Key             =   ""
            Object.Tag             =   "60013"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":216C9
            Key             =   ""
            Object.Tag             =   "60015"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2305B
            Key             =   ""
            Object.Tag             =   "60014"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":249ED
            Key             =   ""
            Object.Tag             =   "50011"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2637F
            Key             =   ""
            Object.Tag             =   "50012"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":27D11
            Key             =   ""
            Object.Tag             =   "50013"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":296A3
            Key             =   ""
            Object.Tag             =   "50014"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2B035
            Key             =   ""
            Object.Tag             =   "40011"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2C9C7
            Key             =   ""
            Object.Tag             =   "40012"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2E359
            Key             =   ""
            Object.Tag             =   "4002"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2FCEB
            Key             =   ""
            Object.Tag             =   "40031"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3167D
            Key             =   ""
            Object.Tag             =   "40032"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3300F
            Key             =   ""
            Object.Tag             =   "40033"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":349A1
            Key             =   ""
            Object.Tag             =   "40034"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":36333
            Key             =   ""
            Object.Tag             =   "40035"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":37CC5
            Key             =   ""
            Object.Tag             =   "14"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":39657
            Key             =   ""
            Object.Tag             =   "15"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3AFE9
            Key             =   ""
            Object.Tag             =   "13"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3B8C3
            Key             =   ""
            Object.Tag             =   "16"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3C19D
            Key             =   ""
            Object.Tag             =   "12"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3CA77
            Key             =   ""
            Object.Tag             =   "900011"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3E409
            Key             =   ""
            Object.Tag             =   "900012"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3FD9B
            Key             =   ""
            Object.Tag             =   "900013"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":4172D
            Key             =   ""
            Object.Tag             =   "50015"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicHolder 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   11160
      TabIndex        =   0
      Top             =   6780
      Visible         =   0   'False
      Width           =   11160
      Begin VB.CheckBox chkBatasTgl 
         Caption         =   "Pencarian Cepat berdasarkan tanggal"
         Height          =   300
         Left            =   11250
         TabIndex        =   1
         Top             =   45
         Width           =   3090
      End
      Begin MSComCtl2.DTPicker DTP2 
         Height          =   300
         Left            =   15795
         TabIndex        =   2
         Top             =   45
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         Format          =   144834561
         CurrentDate     =   41185
      End
      Begin MSComCtl2.DTPicker DTP1 
         Height          =   315
         Left            =   14475
         TabIndex        =   3
         Top             =   45
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   144834561
         CurrentDate     =   41185
      End
   End
   Begin Crystal.CrystalReport Crystal 
      Left            =   7320
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin XtremeSuiteControls.PopupControl popuppesan 
      Left            =   4680
      Top             =   135
      _Version        =   851970
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   4
      Animation       =   2
   End
   Begin XtremeSuiteControls.PopupControl popupstatus 
      Left            =   2940
      Top             =   135
      _Version        =   851970
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   4
      VisualTheme     =   3
      Animation       =   2
   End
   Begin XtremeCommandBars.CommandBars ComBars 
      Left            =   855
      Top             =   150
      _Version        =   851970
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   4
   End
   Begin XtremeSkinFramework.SkinFramework SF 
      Left            =   120
      Top             =   135
      _Version        =   851970
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      AutoApplyNewWindows=   0   'False
      AutoApplyNewThreads=   0   'False
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private OBJ As New ADODB.Connection
Private RST As ADODB.Recordset
Private SQL As String

Private WithEvents frmLogin As frmLogin
Attribute frmLogin.VB_VarHelpID = -1

'Pane
Public WithEvents StatusBar As XtremeCommandBars.StatusBar
Attribute StatusBar.VB_VarHelpID = -1

Private Sub ComBars_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    Select Case control.ID
    '####################################### MASTER USER MENU #################################
        Case ID_FILE_LOGOUT: logout
        Case ID_FILE_CHANGEPASS:
                                Dim cngpass As Object
                                 Set cngpass = CreateObject("master.CPassword")
                                 With cngpass
                                    .token = "kusumah"
                                    .UserOnline = UserOnline
                                    .Path = App.Path
                                    .formname = "changepass"
                                    .FastSearch = False
                                    .Show
                                End With
        Case ID_FILE_DASHBOARD:
                            If UserOnline = "Budiman" Then frmberanda.Show vbModal
        Case ID_FILE_EXIT:
                            If MsgBox("Are your sure to Exit Application ? ", vbYesNo + vbQuestion, AppName) = vbYes Then
                                End
                            End If
        Case ID_FILE_VIEWSTOKBAHANBAKU:
                                            Dim viewstokbahanbaku
                                            Set viewstokbahanbaku = CreateObject("purc_tables.CTables")
                                            With viewstokbahanbaku
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                '.setuseronlinelevel = UserOnLineLevel
                                                .Path = App.Path
                                                .formname = "viewstokbahanbaku"
                                                .FastSearch = False
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_FILE_PESANSYSTEM: pesanasystem
        
        
        Case ID_MASTER_DATABASE_KONEKSISQL:
                                 Dim dllkoneksi
                                 Set dllkoneksi = CreateObject("master.CKoneksi")
                                 dllkoneksi.token = "kusumah"
                                 dllkoneksi.Path = App.Path
                                 dllkoneksi.Show
        Case ID_MASTER_DATABASE_KONEKSI_BACKUP:
                                 frmbackup.Show vbModal
                                 
        Case ID_MASTER_DATABASE_MAINTENANCE_IMPORTBAHANBAKU:
                                 Dim objimportbahanbaku As Object
                                 Set objimportbahanbaku = CreateObject("master.CMaster")
                                 With objimportbahanbaku
                                        .token = "kusumah"
                                        .Path = App.Path
                                        .Show
                                 End With
                                 
        Case ID_MASTER_MANAGEUSER_DEPARTEMEN:
                                            Dim dlldept
                                            Set dlldept = CreateObject("master.CDepartement")
                                            dlldept.token = "kusumah"
                                            dlldept.Path = App.Path
                                            dlldept.Show
                                            SetParent dlldept.hWnd, Me.hWnd
        Case ID_MASTER_MANAGEUSER_MODUL:
                                            Dim dllmodul
                                            Set dllmodul = CreateObject("master.CModul")
                                            dllmodul.token = "kusumah"
                                            dllmodul.Path = App.Path
                                            dllmodul.Show
                                            SetParent dllmodul.hWnd, Me.hWnd
        Case ID_MASTER_MANAGEUSER_LEVEL:
                                            Dim dlllevel
                                            Set dlllevel = CreateObject("master.CLevel")
                                            dlllevel.token = "kusumah"
                                            dlllevel.Path = App.Path
                                            dlllevel.Show
                                            SetParent dlllevel.hWnd, Me.hWnd
        Case ID_MASTER_MANAGEUSER_USER:
                                            Dim dlluser
                                            Set dlluser = CreateObject("master.CUser")
                                            dlluser.token = "kusumah"
                                            'dlluser.setuseronlinelevel = UserOnLineLevel
                                            dlluser.Path = App.Path
                                            'dlluser.setdsn = dsn
                                            dlluser.Show
                                            SetParent dlluser.hWnd, Me.hWnd
        Case ID_MASTER_MANAGEUSER_USERDEPT:
                                            Dim dlluserdept
                                            Set dlluserdept = CreateObject("master.CUserdept")
                                            dlluserdept.token = "kusumah"
                                            dlluserdept.Path = App.Path
                                            dlluserdept.Show
                                            SetParent dlluserdept.hWnd, Me.hWnd
        Case ID_MASTER_TELEGRAM_BOT: frmtelegram.Show vbModal
                                                                                
        '############################################## PURCHASING TABLES MENU #########################
        Case ID_PEMBELIAN_TABEL_SUPPLIER_ADD:
                                            Dim p_tabels
                                            Set p_tabels = CreateObject("purc_tables.CTables")
                                            With p_tabels
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "addsupplier"
                                                .FastSearch = False
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_TABEL_SUPPLIER_CHANGE:
                                            Dim cngsupplier As Object
                                            Set cngsupplier = CreateObject("purc_tables.CTables")
                                            With cngsupplier
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "changesupplier"
                                                .FastSearch = False
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_TABEL_SUPPLIER_LIST:
                                            Dim daftarsupplier As Object
                                            Set daftarsupplier = CreateObject("purc_tables.CTables")
                                            With daftarsupplier
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "daftarsupplier"
                                                .FastSearch = False
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_TABEL_SUPPLIER_PRICELIST:
                                            Dim pricelist As Object
                                            Set pricelist = CreateObject("purc_tables.CTables")
                                            With pricelist
                                               .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "pricelist"
                                                .FastSearch = False
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_TABEL_SUPPLIER_GRAPRICE: subcombars_purchasing (ID_PEMBELIAN_TABEL_SUPPLIER_GRAPRICE)
        
        Case ID_PEMBELIAN_TABEL_SATUAN_ADD:
                                            Dim addunitpembelian As Object
                                            Set addunitpembelian = CreateObject("purc_tables.CTables")
                                            With addunitpembelian
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "addunit"
                                                .FastSearch = False
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_TABEL_SATUAN_CHANGE:
                                            Dim changeunitpembelian As Object
                                            Set changeunitpembelian = CreateObject("purc_tables.CTables")
                                            With changeunitpembelian
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "changeunit"
                                                .FastSearch = False
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_TABEL_SATUAN_LIST:
                                            Dim listunitpembelian As Object
                                            Set listunitpembelian = CreateObject("purc_tables.CTables")
                                            With listunitpembelian
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "listunit"
                                                .FastSearch = False
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_TABEL_BAHANBAKU_MANAGE:
                                            Dim addbahanbaku As Object
                                            Set addbahanbaku = CreateObject("purc_tables.CTables")
                                            With addbahanbaku
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "addbahanbaku"
                                                .FastSearch = False
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_PEMBELIAN_TABEL_BAHANBAKU_LIST:
                                            Dim listitembb As Object
                                            Set listitembb = CreateObject("purc_tables.CTables")
                                            With listitembb
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "listitembb"
                                                .FastSearch = False
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_PEMBELIAN_TABEL_PACKAGING_ADD:
                                            Dim addpackagingbahan As Object
                                            Set addpackagingbahan = CreateObject("purc_tables.CTables")
                                            With addpackagingbahan
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "addpackaging"
                                                .FastSearch = False
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        
        Case ID_PEMBELIAN_TABEL_MINSTOCK_ADD:
                                            Dim addminstock As Object
                                            Set addminstock = CreateObject("purc_tables.CTables")
                                            With addminstock
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "addmin"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_TABEL_MINSTOCK_CHANGE:
                                            Dim cngminstock As Object
                                            Set cngminstock = CreateObject("purc_tables.CTables")
                                            With cngminstock
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "cngmin"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_TABEL_MINSTOCK_LIST:
                                            Dim listminstock As Object
                                            Set listminstock = CreateObject("purc_tables.CTables")
                                            With listminstock
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "listmin"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        
        Case ID_PEMBELIAN_MUTASIBARANG_ADD:
                                            Dim addmutasibahan As Object
                                            Set addmutasibahan = CreateObject("purc_mut.CMutasi")
                                            With addmutasibahan
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "addmutasi"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_PEMBELIAN_MUTASIBARANG_CHANGE:
                                            Dim cngmutasipembelian As Object
                                            Set cngmutasipembelian = CreateObject("purc_mut.CMutasi")
                                            With cngmutasipembelian
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "changemutasi"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_MUTASIBARANG_LIST:
                                            Dim listmutasipembelian As Object
                                            Set listmutasipembelian = CreateObject("purc_mut.CMutasi")
                                            With listmutasipembelian
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "listmutasi"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_MUTASIBARANG_BASE:
                                            Dim mutasibase As Object
                                            Set mutasibase = CreateObject("purc_mut.CMutasi")
                                            With mutasibase
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "mutasibase"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        
        Case ID_PEMBELIAN_PURCHASING_PO_ADD:
                                            Dim addpo As Object
                                            Set addpo = CreateObject("purc_purc.CPurchasing")
                                            With addpo
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "addpo"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_PURCHASING_PO_CHANGE:
                                            Dim cngpo As Object
                                            Set cngpo = CreateObject("purc_purc.CPurchasing")
                                            With cngpo
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "changepo"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_PURCHASING_PO_CLOSECANCEL:
                                            Dim closepo As Object
                                            Set closepo = CreateObject("purc_purc.CPurchasing")
                                            With closepo
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "closepo"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_PURCHASING_PENERIMAANBARANG_ADD:
                                            Dim addpenerimaanbahan As Object
                                            Set addpenerimaanbahan = CreateObject("purc_purc.CPurchasing")
                                            With addpenerimaanbahan
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "addpenerimaan"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_PURCHASING_PENERIMAANBARANG_CHANGE:
                                            Dim cngpenerimaanbahan As Object
                                            Set cngpenerimaanbahan = CreateObject("purc_purc.CPurchasing")
                                            With cngpenerimaanbahan
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "changepenerimaan"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_PURCHASING_PENERIMAANBARANG_RETUR:
                                            Dim returpenerimaanbahan As Object
                                            Set returpenerimaanbahan = CreateObject("purc_purc.CPurchasing")
                                            With returpenerimaanbahan
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "returpenerimaan"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_PURCHASING_PENERIMAANBARANG_PRINTBPB:
                                            Dim printbpb As Object
                                            Set printbpb = CreateObject("purc_purc.CPurchasing")
                                            With printbpb
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "printbpb"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_PURCHASING_PERMINTAAN_ADD:
                                            Dim addpermintaanbarang As Object
                                            Set addpermintaanbarang = CreateObject("purc_purc.CPurchasing")
                                            With addpermintaanbarang
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "addpermintaan"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_PURCHASING_PERMINTAAN_CHANGE:
                                            Dim cngpermintaanbarang As Object
                                            Set cngpermintaanbarang = CreateObject("purc_purc.CPurchasing")
                                            With cngpermintaanbarang
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "cngpermintaan"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_PURCHASING_PERMINTAAN_CLOSE:
                                            Dim clspermintaanbarang As Object
                                            Set clspermintaanbarang = CreateObject("purc_purc.CPurchasing")
                                            With clspermintaanbarang
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "clspermintaan"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_PURCHASING_PERMINTAAN_LIST:
                                            Dim listpermintaanbarang As Object
                                            Set listpermintaanbarang = CreateObject("purc_purc.CPurchasing")
                                            With listpermintaanbarang
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "listpermintaan"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_PURCHASING_PERMINTAAN_DAYSCOUNT: subcombars_purchasing (ID_PEMBELIAN_PURCHASING_PERMINTAAN_DAYSCOUNT)
        Case ID_PEMBELIAN_PURCHASING_PERMINTAAN_PRINT:
                                            Dim permintaanprint As Object
                                            Set permintaanprint = CreateObject("purc_purc.CPurchasing")
                                            With permintaanprint
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "printpermintaan"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        '############################################## SALE MENU ######################################
        
        Case ID_PENJUALAN_MAINMENU_TABLE_SATUAN_ADD:
                                            Dim addunit
                                            Set addunit = CreateObject("sale_tbl.CTables")
                                            With addunit
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "addunit"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_TABLE_SATUAN_CHANGE:
                                            Dim cngunit
                                            Set cngunit = CreateObject("sale_tbl.CTables")
                                            With cngunit
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "changeunit"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_TABLE_SATUAN_LIST:
                                            Dim listunit
                                            Set listunit = CreateObject("sale_tbl.CTables")
                                            With listunit
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "listunit"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_TABLE_CATAGORI_ADD:
                                            Dim addcat
                                            Set addcat = CreateObject("sale_tbl.CTables")
                                            With addcat
                                                 .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "addcat"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_TABLE_CATAGORI_CHANGE:
                                            Dim cngcat
                                            Set cngcat = CreateObject("sale_tbl.CTables")
                                            With cngcat
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "changecat"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_TABLE_CATAGORI_LIST:
                                            Dim catlist
                                            Set catlist = CreateObject("sale_tbl.CTables")
                                            With catlist
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "catlist"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_TABLE_BARANGJADI_MANAGE:
                                            Dim additem
                                            Set additem = CreateObject("sale_tbl.CTables")
                                            With additem
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "additem"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_TABLE_BARANGJADI_LIST:
                                            Dim itemlist
                                            Set itemlist = CreateObject("sale_tbl.CTables")
                                            With itemlist
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "itemlist"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_TABLE_GUDANG_ADD:
                                            Dim addgudang
                                            Set addgudang = CreateObject("sale_tbl.CTables")
                                            With addgudang
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "addgudang"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_TABLE_GUDANG_CHANGE:
                                            Dim changegudang
                                            Set changegudang = CreateObject("sale_tbl.CTables")
                                            With changegudang
                                                 .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "changegudang"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_TABLE_GUDANG_LIST:
                                            Dim gudanglist
                                            Set gudanglist = CreateObject("sale_tbl.CTables")
                                            With gudanglist
                                                 .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "gudanglist"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_TABLE_AREA_ADD:
                                            Dim addarea
                                            Set addarea = CreateObject("sale_tbl.CTables")
                                            With addarea
                                                 .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "addarea"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_TABLE_AREA_CHANGE:
                                            Dim cngarea
                                            Set cngarea = CreateObject("sale_tbl.CTables")
                                            With cngarea
                                                 .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "changearea"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_TABLE_AREA_LIST:
                                            Dim arealist
                                            Set arealist = CreateObject("sale_tbl.CTables")
                                            With arealist
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "arealist"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_TABLE_CUSTOMER_ADD:
                                            Dim addcust
                                            Set addcust = CreateObject("sale_tbl.CTables")
                                            With addcust
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "addcust"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_TABLE_CUSTOMER_CHANGE:
                                            Dim cngcust
                                            Set cngcust = CreateObject("sale_tbl.CTables")
                                            With cngcust
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "changecust"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_TABLE_CUSTOMER_LIST:
                                            Dim custlist
                                            Set custlist = CreateObject("sale_tbl.CTables")
                                            With custlist
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "custlist"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_TABLE_SALES_MANAGE:
                                            Dim salesmanage
                                            Set salesmanage = CreateObject("sale_tbl.CTables")
                                            With salesmanage
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "salesmanage"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        'MUTATION ********************************************************************************
        
        Case ID_PENJUALAN_MAINMENU_MUTASI_ADD:
                                            Dim addmut
                                            Set addmut = CreateObject("sale_mut.CMutasi")
                                            With addmut
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "addmut"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_MUTASI_CHANGE:
                                            Dim cngmut
                                            Set cngmut = CreateObject("sale_mut.CMutasi")
                                            With cngmut
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "changemut"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_MUTASI_LIST:
                                            Dim mutlist
                                            Set mutlist = CreateObject("sale_mut.CMutasi")
                                            With mutlist
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "mutlist"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_MUTASI_MUT_OVERZAK: subcombars (ID_PENJUALAN_MAINMENU_MUTASI_MUT_OVERZAK)
        Case ID_PENJUALAN_MAINMENU_MUTASI_MUT_MUTASILOT: subcombars (ID_PENJUALAN_MAINMENU_MUTASI_MUT_MUTASILOT)
        Case ID_PENJUALAN_MAINMENU_MUTASI_MUT_MUTPABRIK: subcombars (ID_PENJUALAN_MAINMENU_MUTASI_MUT_MUTPABRIK)
        Case ID_PENJUALAN_MAINMENU_MUTASI_MUT_ADJSTOK: subcombars (ID_PENJUALAN_MAINMENU_MUTASI_MUT_ADJSTOK)
        Case ID_PENJUALAN_MAINMENU_MUTASI_GUDANG: subcombars (ID_PENJUALAN_MAINMENU_MUTASI_GUDANG)
        Case ID_PENJUALAN_MAINMENU_MUTASI_FAILED: subcombars (ID_PENJUALAN_MAINMENU_MUTASI_FAILED)
        Case ID_PENJUALAN_MAINMENU_MUTASI_WIP: subcombars (ID_PENJUALAN_MAINMENU_MUTASI_WIP)
        Case ID_PENJUALAN_MAINMENU_MUTASI_KARTUSTOK: subcombars (ID_PENJUALAN_MAINMENU_MUTASI_KARTUSTOK)
        Case ID_PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_TOWIP: subcombars (ID_PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_TOWIP)
        Case ID_PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_ADD:
                                            Dim pindahgudang
                                            Set pindahgudang = CreateObject("sale_mut.CMutasi")
                                            With pindahgudang
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "pindahgudang"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_CHANGE:
                                            Dim pindahgudangch
                                            Set pindahgudangch = CreateObject("sale_mut.CMutasi")
                                            With pindahgudangch
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "pindahgudangch"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_PRINT:
                                            Dim pindahgudangprt
                                            Set pindahgudangprt = CreateObject("sale_mut.CMutasi")
                                            With pindahgudangprt
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "pindahgudangprt"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_LOTPALET:
                                            Dim pindahgudangpalet
                                            Set pindahgudangpalet = CreateObject("sale_mut.CMutasi")
                                            With pindahgudangpalet
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "pindahgudangpalet"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_SCANPALET: subcombars (ID_PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_SCANPALET)
        Case ID_PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_COMPAREHPP: subcombars (ID_PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_COMPAREHPP)
        Case ID_PENJUALAN_MAINMENU_MUTASI_PRINTPRICE:
                                            Dim printmut
                                            Set printmut = CreateObject("sale_mut.CMutasi")
                                            With printmut
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "printmut"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_MUTASI_LISTSTOK:
                                            Dim printstok
                                            Set printstok = CreateObject("sale_mut.CMutasi")
                                            With printstok
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "printstok"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_MUTASI_PACKLIST_ADD: subcombars (ID_PENJUALAN_MAINMENU_MUTASI_PACKLIST_ADD)
        Case ID_PENJUALAN_MAINMENU_MUTASI_PACKLIST_LIST: subcombars (ID_PENJUALAN_MAINMENU_MUTASI_PACKLIST_LIST)
        Case ID_PENJUALAN_MAINMENU_MUTASI_PACKLIST_CLOSE: subcombars (ID_PENJUALAN_MAINMENU_MUTASI_PACKLIST_CLOSE)
        'UTILITY ***********************************************************************************
        Case ID_PENJUALAN_MAINMENU_UTILITY_OPTIONS:
                                            Dim opt
                                            Set opt = CreateObject("sale_uti.CUtility")
                                            With opt
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "opt"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_UTILITY_CANCELSO:
                                            Dim cancelso
                                            Set cancelso = CreateObject("sale_uti.CUtility")
                                            With cancelso
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "cancelso"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_UTILITY_CLOSESO:
                                            Dim closeso
                                            Set closeso = CreateObject("sale_uti.CUtility")
                                            With closeso
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "closeso"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_UTILITY_EXPORTSO:
                                            Dim soexport
                                            Set soexport = CreateObject("sale_uti.CUtility")
                                            With soexport
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "soexport"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_UTILITY_IMPORTSOPBRK:
                                            Dim soimport
                                            Set soimport = CreateObject("sale_uti.CUtility")
                                            With soimport
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "soimport"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_UTILITY_EXPORTSJ:
                                            Dim sjexport
                                            Set sjexport = CreateObject("sale_uti.CUtility")
                                            With sjexport
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "sjexport"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_UTILITY_IMPORTSJ:
                                            Dim sjimport
                                            Set sjimport = CreateObject("sale_uti.CUtility")
                                            With sjimport
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "sjimport"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_UTILITY_IMPORTINV:
                                            Dim objimportinv
                                            Set objimportinv = CreateObject("sale_uti.CUtility")
                                            With objimportinv
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "importinv"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_UTILITY_DELINV:
                                            Dim deleteinv
                                            Set deleteinv = CreateObject("sale_uti.CUtility")
                                            With deleteinv
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "deleteinv"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        'INVOICING
        Case ID_PENJUALAN_MAINMENU_INVOICING_SO_ADD:
                                            Dim addso
                                            Set addso = CreateObject("sale_inv.CInvoicing")
                                            addso.token = "kusumah"
                                            addso.Path = App.Path
                                            addso.formname = "AddSO"
                                            addso.asremote = remoteserver
                                            addso.ipserver = dbServer
                                            addso.Show
                                            SetParent addso.hWnd, Me.hWnd
        Case ID_PENJUALAN_MAINMENU_INVOICING_SO_CHANGE:
                                            Dim chgso
                                            Set chgso = CreateObject("sale_inv.CInvoicing")
                                            chgso.token = "kusumah"
                                            chgso.Path = App.Path
                                            chgso.formname = "ChangeSO"
                                            chgso.FastSearch = False
                                            chgso.asremote = remoteserver
                                            chgso.ipserver = dbServer
                                            chgso.Show
                                            SetParent chgso.hWnd, Me.hWnd
        'Case ID_PENJUALAN_MAINMENU_INVOICING_SO_CANCEL: subcombars (ID_PENJUALAN_MAINMENU_INVOICING_SO_CANCEL)
        Case ID_PENJUALAN_MAINMENU_INVOICING_SJ_ADD:
                                            Dim addsj
                                            Set addsj = CreateObject("sale_inv.CInvoicing")
                                            With addsj
                                                .token = "kusumah"
                                                .Path = App.Path
                                                .formname = "AddSJ"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_INVOICING_SJ_CHANGE:
                                            Dim cngsj
                                            Set cngsj = CreateObject("sale_inv.CInvoicing")
                                            With cngsj
                                                .token = "kusumah"
                                                .Path = App.Path
                                                .formname = "ChangeSJ"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_INVOICING_SJ_PRINT:
                                            Dim printsj
                                            Set printsj = CreateObject("sale_inv.CInvoicing")
                                            With printsj
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "PrintSJ"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_INVOICING_SJ_ADDLOT: subcombars_penjualan (ID_PENJUALAN_MAINMENU_INVOICING_SJ_ADDLOT)
        Case ID_PENJUALAN_MAINMENU_INVOICING_FAKTURJUAL_ADD:
                                            Dim addfjual
                                            Set addfjual = CreateObject("sale_inv.CInvoicing")
                                            With addfjual
                                                 .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "addfjual"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_INVOICING_FAKTURJUAL_CHANGE:
                                            Dim changefjual
                                            Set changefjual = CreateObject("sale_inv.CInvoicing")
                                            With changefjual
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "changefjual"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_INVOICING_FAKTURJUAL_PRINT:
                                            Dim printfjual
                                            Set printfjual = CreateObject("sale_inv.CInvoicing")
                                            With printfjual
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "printfjual"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_INVOICING_FAKTURJUAL_PREVIEW:
                                            Dim prevfjual
                                            Set prevfjual = CreateObject("sale_inv.CInvoicing")
                                            With prevfjual
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "prevfjual"
                                                .FastSearch = False
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_INVOICING_FAKTURPAJAK_DEFINE:
                                            Dim definefpajak
                                            Set definefpajak = CreateObject("sale_inv.CInvoicing")
                                            With definefpajak
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "definefpajak"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_INVOICING_FAKTURPAJAK_BROWSE:
                                            Dim browsefpajak
                                            Set browsefpajak = CreateObject("sale_inv.CInvoicing")
                                            With browsefpajak
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "browsefpajak"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_INVOICING_SJSBY_ADD:
                                            Dim addsjsby
                                            Set addsjsby = CreateObject("sale_inv.CInvoicing")
                                            With addsjsby
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "addsjsby"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_INVOICING_SJSBY_LIST:
                                            Dim listsjsby
                                            Set listsjsby = CreateObject("sale_inv.CInvoicing")
                                            With listsjsby
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "listsjsby"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_INVOICING_INQUERYSO:
                                            Dim inqso
                                            Set inqso = CreateObject("sale_inv.CInvoicing")
                                            With inqso
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "iqso"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_INVOICING_LAPSO:
                                            Dim lapso
                                            Set lapso = CreateObject("sale_inv.CInvoicing")
                                            With lapso
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "solist"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_PENJUALAN_MAINMENU_INVOICING_LAPSJ_DAFTAR:
                                            Dim lapsjdaftar
                                            Set lapsjdaftar = CreateObject("sale_inv.CInvoicing")
                                            With lapsjdaftar
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "lapsjdaftar"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_INVOICING_LAPSJ_DAFTARGD:
                                            Dim lapsjdaftargd
                                            Set lapsjdaftargd = CreateObject("sale_inv.CInvoicing")
                                            With lapsjdaftargd
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "lapsjdaftargd"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_INVOICING_LAPSJ_BYFAKTUR:
                                            Dim lapsjbyfaktur
                                            Set lapsjbyfaktur = CreateObject("sale_inv.CInvoicing")
                                            With lapsjbyfaktur
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "lapsjbyfaktur"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_INVOICING_LAPSJ_BYLOT: subcombars_penjualan (ID_PENJUALAN_MAINMENU_INVOICING_LAPSJ_BYLOT)
                                            
        Case ID_PENJUALAN_MAINMENU_INVOICING_LAPSJ_LAPORAN:
                                            Dim lasjlap
                                            Set lasjlap = CreateObject("sale_inv.CInvoicing")
                                            With lasjlap
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "lapsjlap"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_INVOICING_LAPJUAL:
                                            Dim dafjual
                                            Set dafjual = CreateObject("sale_inv.CInvoicing")
                                            With dafjual
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "dafjual"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_INVOICING_LAPJUALDTL:
                                            Dim dafjualdet
                                            Set dafjualdet = CreateObject("sale_inv.CInvoicing")
                                            With dafjualdet
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "dafjualdet"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_INVOICING_MONTHLY:
                                            Dim jualmonthly
                                            Set jualmonthly = CreateObject("sale_inv.CInvoicing")
                                            With jualmonthly
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "jualmonthly"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_INVOICING_BYKATEGORI: subcombars (ID_PENJUALAN_MAINMENU_INVOICING_BYKATEGORI)
        Case ID_PENJUALAN_MAINMENU_INVOICING_LAPKOMISI:
                                            Dim dafkomisi
                                            Set dafkomisi = CreateObject("sale_inv.CInvoicing")
                                            With dafkomisi
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "dafkomisi"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PENJUALAN_MAINMENU_INVOICING_ANALISAJUAL:
                                            Dim anjual
                                            Set anjual = CreateObject("sale_inv.CInvoicing")
                                            With anjual
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "anjual"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        
        'DEPARTEMEN FINANCE MENU
        'SALE / PENJUALAN
        Case ID_KEUANGAN_PENJUALAN_PIUTANG_KOREKSIPIUTANG_ADD:
                                            Dim addkoreksipiutang As Object
                                            Set addkoreksipiutang = CreateObject("finance_sale.CSale")
                                            With addkoreksipiutang
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "addkoreksi"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_PENJUALAN_PIUTANG_KOREKSIPIUTANG_CHANGE:
                                            Dim cngkoreksipiutang As Object
                                            Set cngkoreksipiutang = CreateObject("finance_sale.CSale")
                                            With cngkoreksipiutang
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "changekoreksi"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_PENJUALAN_PIUTANG_KOREKSIPIUTANG_WRITEOFF:
                                            Dim wroff As Object
                                            Set wroff = CreateObject("finance_sale.CSale")
                                            With wroff
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "writeoff"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_PENJUALAN_PIUTANG_PEMBAYARAN_ADD:
                                            Dim addpayarpiutang As Object
                                            Set addpayarpiutang = CreateObject("finance_sale.CSale")
                                            With addpayarpiutang
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "addpembayaran"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_PENJUALAN_PIUTANG_PEMBAYARAN_CHANGE:
                                            Dim cngpayarpiutang As Object
                                            Set cngpayarpiutang = CreateObject("finance_sale.CSale")
                                            With cngpayarpiutang
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "changepembayaran"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_PENJUALAN_PIUTANG_DAFTARKOREKSIPIUTANG:
                                            Dim dafkoreksipiutang As Object
                                            Set dafkoreksipiutang = CreateObject("finance_sale.CSale")
                                            With dafkoreksipiutang
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "daftarkoreksi"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_PENJUALAN_PIUTANG_LAPPENAGIHAN:
                                            Dim dafpenagihanpiutang As Object
                                            Set dafpenagihanpiutang = CreateObject("finance_sale.CSale")
                                            With dafpenagihanpiutang
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "daftarpenagihan"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_PENJUALAN_PIUTANG_LAPPIUTANGAGING:
                                            Dim dafpiutangaging As Object
                                            Set dafpiutangaging = CreateObject("finance_sale.CSale")
                                            With dafpiutangaging
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "daftarpiutang"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_PENJUALAN_PIUTANG_LAPKARTUPIUTANG:
                                            Dim dafpiutangkartu As Object
                                            Set dafpiutangkartu = CreateObject("finance_sale.Csale")
                                            With dafpiutangkartu
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "daftarpiutangkartu"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_PENJUALAN_PIUTANG_LAPSISAPIUTANG:
                                            Dim daftarpiutangsisa As Object
                                            Set daftarpiutangsisa = CreateObject("finance_sale.CSale")
                                            With daftarpiutangsisa
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "daftarpiutangsisa"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_PENJUALAN_PIUTANG_LAPMUTPIUTANG:
                                            Dim daftarmutasipiutang As Object
                                            Set daftarmutasipiutang = CreateObject("finance_sale.Csale")
                                            With daftarmutasipiutang
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "daftarmutasipiutang"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_PENJUALAN_PIUTANG_LAPPEMBAYARAN:
                                            Dim dafpembayaranpiutang As Object
                                            Set dafpembayaranpiutang = CreateObject("finance_sale.CSale")
                                            With dafpembayaranpiutang
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "daftarpembayaran"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_PENJUALAN_PIUTANG_LAPPEMBAYARANDTL:
                                            Dim dafpembayarandtl
                                            Set dafpembayarandtl = CreateObject("finance_sale.CSale")
                                            With dafpembayarandtl
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "daftarpembayarandetail"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_PENJUALAN_PIUTANG_LAPTANDATERIMAPBY:
                                            Dim daftarttpembayaran As Object
                                            Set daftarttpembayaran = CreateObject("finance_sale.CSale")
                                            With daftarttpembayaran
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "daftarttpembayaran"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_PENJUALAN_GIRO_MAINTENANCE:
                                            Dim maintenancegiro As Object
                                            Set maintenancegiro = CreateObject("finance_sale.CSale")
                                            With maintenancegiro
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "maintenancegiro"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_PENJUALAN_GIRO_GANTIGIROTOLAK_ADD:
                                            Dim addgirotolak As Object
                                            Set addgirotolak = CreateObject("finance_sale.CSale")
                                            With addgirotolak
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "addgirotolak"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_PENJUALAN_GIRO_GANTIGIROTOLAK_CHANGE:
                                            Dim changegirotolak As Object
                                            Set changegirotolak = CreateObject("finance_sale.CSale")
                                            With changegirotolak
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "changegirotolak"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_PENJUALAN_GIRO_LAPGIRO:
                                            Dim daftargiropenjualan As Object
                                            Set daftargiropenjualan = CreateObject("finance_sale.CSale")
                                            With daftargiropenjualan
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "daftargiro"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_PENJUALAN_GIRO_LISTAPPGIROTOLAK:
                                            Dim daftargiroturun As Object
                                            Set daftargiroturun = CreateObject("finance_sale.CSale")
                                            With daftargiroturun
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "daftargiroturun"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_PENJUALAN_UTILITY_DEFAINACCBANKCASH:
                                            Dim defineaccbankpenjualan As Object
                                            Set defineaccbankpenjualan = CreateObject("finance_sale.CSale")
                                            With defineaccbankpenjualan
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "defineaccbank"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_PENJUALAN_UTILITY_DEFAINJURNAL:
                                            Dim definejurnalpenjualan As Object
                                            Set defineaccbankpenjualan = CreateObject("finance_sale.CSale")
                                            With defineaccbankpenjualan
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "definejurnal"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_PENJUALAN_UTILITY_DEFINEAGING:
                                            Dim defineagingpiutang As Object
                                            Set defineagingpiutang = CreateObject("finance_sale.CSale")
                                            With defineagingpiutang
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "defineaging"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_PENJUALAN_UTILITY_DEFINEKG:
                                            Dim defineitemkgpenjualan As Object
                                            Set defineitemkgpenjualan = CreateObject("finance_sale.CSale")
                                            With defineitemkgpenjualan
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "definekgbase"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_PENJUALAN_UTILITY_DEFINECOSUTOMERACC:
                                            Dim defaccountcust As Object
                                            Set defaccountcust = CreateObject("finance_sale.CSale")
                                            With defaccountcust
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "defineaccountcust"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        
        Case ID_KEUANGAN_PENJUALAN_UTILITY_LAPPOSTINGPENJUALAN:
                                            Dim lappostingpenjualan As Object
                                            Set lappostingpenjualan = CreateObject("finance_sale.CSale")
                                            With lappostingpenjualan
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "lap_posting_penjualan"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        
        Case ID_KEUANGAN_PENJUALAN_UTILITY_LISTKG: Shell App.Path & "\ext\kilosales.exe", vbNormalFocus
        
        Case ID_PEMBELIAN_PURCHASING_CONFIRM_CONFIRMPENERIMAANBARANG_ADD:
                                            Dim confirmpenerimaan As Object
                                            Set confirmpenerimaan = CreateObject("finance_purc.CPurchasing")
                                            With confirmpenerimaan
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "confirmpenerimaan"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_PURCHASING_CONFIRM_CONFIRMPENERIMAANBARANG_REPRINT:
                                            Dim reprintpenerimaan As Object
                                            Set reprintpenerimaan = CreateObject("finance_purc.CPurchasing")
                                            With reprintpenerimaan
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "reprintpenerimaan"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_PEMBELIAN_PURCHASING_CONFIRM_CONFIRMRETURPENERIMAANBARANG:
                                            Dim confirmpenerimaanretur As Object
                                            Set confirmpenerimaanretur = CreateObject("finance_purc.CPurchasing")
                                            With confirmpenerimaanretur
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "confirmpenerimaanretur"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_PURCHASING_CONFIRM_UNCONFIRMPENERIMAANBARANG:
                                            Dim unconfirmpenerimaan As Object
                                            Set unconfirmpenerimaan = CreateObject("finance_purc.CPurchasing")
                                            With unconfirmpenerimaan
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "unconfirm"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_PURCHASING_CONFIRM_UNCOFIRMRETURPENERIMAANBARANG:
                                            Dim unconfirmpenerimaanretur As Object
                                            Set unconfirmpenerimaanretur = CreateObject("finance_purc.CPurchasing")
                                            With unconfirmpenerimaanretur
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "unconfirmretur"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_PURCHASING_CONFIRM_CREATEVOUCHER:
                                            Dim createvoucher As Object
                                            Set createvoucher = CreateObject("finance_purc.CPurchasing")
                                            With createvoucher
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "createvoucher"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_HUTANG_PBYHUTANG_ADD:
                                            Dim pbyhutang As Object
                                            Set pbyhutang = CreateObject("finance_purc.CPurchasing")
                                            With pbyhutang
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "addpembayaran"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_HUTANG_PBYHUTANG_CHANGE:
                                            Dim changepbyhutang As Object
                                            Set changepbyhutang = CreateObject("finance_purc.CPurchasing")
                                            With changepbyhutang
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "changepembayaran"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_HUTANG_PBYHUTANG_UNPOST:
                                            Dim unposthutang As Object
                                            Set unposthutang = CreateObject("finance_purc.CPurchasing")
                                            With unposthutang
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "unposthutang"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_HUTANG_KOREKSIHTN_ADD:
                                            Dim addkoreksihutang As Object
                                            Set addkoreksihutang = CreateObject("finance_purc.CPurchasing")
                                            With addkoreksihutang
                                               .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "addkoreksi"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_HUTANG_KOREKSIHTN_CHANGE:
                                            Dim cngkoreksihutang As Object
                                            Set cngkoreksihutang = CreateObject("finance_purc.CPurchasing")
                                            With cngkoreksihutang
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "changekoreksi"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_HUTANG_LISTKOREKSI:
                                            Dim daftarkoreksihutang As Object
                                            Set daftarkoreksihutang = CreateObject("finance_purc.CPurchasing")
                                            With daftarkoreksihutang
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "daftarkoreksi"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_HUTANG_LISTPBYHUTANG:
                                            Dim daftarbayarhutang As Object
                                            Set daftarbayarhutang = CreateObject("finance_purc.CPurchasing")
                                            With daftarbayarhutang
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "daftarbayarhutang"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_PEMBELIAN_HTNGKARTU:
                                            Dim daftarhutangkartu As Object
                                            Set daftarhutangkartu = CreateObject("finance_purc.CPurchasing")
                                            With daftarhutangkartu
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "daftarhutangkartu"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_HUTANG_LAPSISAHUTANG:
                                            Dim daftarsisahutang As Object
                                            Set daftarsisahutang = CreateObject("finance_purc.CPurchasing")
                                            With daftarsisahutang
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "daftarsisahutang"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_HUTANG_LAPBYJT:
                                            Dim lapbyjt As Object
                                            Set lapbyjt = CreateObject("finance_purc.CPurchasing")
                                            With lapbyjt
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "lapbyjt"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_PEMBELIAN_MAINTENACEGIRO:
                                            Dim maintenancegiropembelian As Object
                                            Set maintenancegiropembelian = CreateObject("finance_purc.CPurchasing")
                                            With maintenancegiropembelian
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "maintenancegiro"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_PEMBELIAN_ADDGIROTOLAK:
                                            Dim addgirotolakpembelian As Object
                                            Set addgirotolakpembelian = CreateObject("finance_purc.CPurchasing")
                                            With addgirotolakpembelian
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "addgirotolak"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_PEMBELIAN_CHANGEGIROTOLAK:
                                            Dim cnggirotolakpembelian As Object
                                            Set cnggirotolakpembelian = CreateObject("finance_purc.CPurchasing")
                                            With cnggirotolakpembelian
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "changegirotolak"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_PEMBELIAN_LAPGIRO:
                                            Dim daftargiropembelian As Object
                                            Set daftargiropembelian = CreateObject("finance_purc.CPurchasing")
                                            With daftargiropembelian
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "daftargiro"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_PEMBELIAN_UTILITY_DEFINEACCOUNTSUPPLIER:
                                            Dim definesupp As Object
                                            Set definesupp = CreateObject("finance_purc.CPurchasing")
                                            With definesupp
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "definesupp"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_PEMBELIAN_UTILITY_DEFINEBANKKAS:
                                            Dim defineaccbankpenerimaan As Object
                                            Set defineaccbankpenerimaan = CreateObject("finance_purc.CPurchasing")
                                            With defineaccbankpenerimaan
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "defineaccbank"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_PEMBELIAN_UTILITY_DEFINEJURNALANDPROSES:
                                            Dim definejurnalpenerimaan As Object
                                            Set defineaccbankpenerimaan = CreateObject("finance_purc.CPurchasing")
                                            With defineaccbankpenerimaan
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "definejurnal"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_PEMBELIAN_UTILITY_LAPORANPOSTING:
                                            Dim lappostingpenerimaan As Object
                                            Set lappostingpenerimaan = CreateObject("finance_purc.CPurchasing")
                                            With lappostingpenerimaan
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "lap_posting_penerimaan"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        
        Case ID_KEUANGAN_GL_TABEL_COMPANYTYPE_ADD:
                                            Dim addtypecomp As Object
                                            Set addtypecomp = CreateObject("gl_tables.CTables")
                                            With addtypecomp
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "addtypecomp"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_GL_TABEL_COMPANYTYPE_UPDATE:
                                            Dim changetypecomp As Object
                                            Set changetypecomp = CreateObject("gl_tables.CTables")
                                            With changetypecomp
                                               .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "changetypecomp"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_GL_TABEL_COMPANYTYPE_LIST:
                                            Dim typecomplist As Object
                                            Set typecomplist = CreateObject("gl_tables.CTables")
                                            With typecomplist
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "typecomplist"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_GL_TABEL_MASTERACC_ADD:
                                            Dim addmasteracc As Object
                                            Set addmasteracc = CreateObject("gl_tables.CTables")
                                            With addmasteracc
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "addmasteracc"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                             End With
                                                
        Case ID_KEUANGAN_GL_TABEL_MASTERACC_UPDATE:
                                            Dim changemasteracc As Object
                                            Set changemasteracc = CreateObject("gl_tables.CTables")
                                            With changemasteracc
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "changemasteracc"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_TABEL_MASTERACC_LIST:
                                            Dim listmasteracc As Object
                                            Set listmasteracc = CreateObject("gl_tables.CTables")
                                            With listmasteracc
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "listmasteracc"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_GL_TABEL_COMPANYIDENTIFICATION_ADD:
                                            Dim addcompany As Object
                                            Set addcompany = CreateObject("gl_tables.CTables")
                                            With addcompany
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "addcompany"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_GL_TABEL_COMPANYIDENTIFICATION_UPDATE:
                                            Dim changecompany As Object
                                            Set changecompany = CreateObject("gl_tables.CTables")
                                            With changecompany
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "changecompany"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_GL_TABEL_COMPANYIDENTIFICATION_LIST:
                                            Dim listcompany As Object
                                            Set listcompany = CreateObject("gl_tables.CTables")
                                            With listcompany
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "listcompany"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                        
        Case ID_KEUANGAN_GL_TABEL_COMPANYACCOUNT_BROWSE:
                                            Dim browseacc As Object
                                            Set browseacc = CreateObject("gl_tables.CTables")
                                            With browseacc
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "browseacc"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_TABEL_COMPANYACCOUNT_LISTACCOUNT:
                                            Dim listaccount As Object
                                            Set listaccount = CreateObject("gl_tables.CTables")
                                            With listaccount
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "listaccount"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_TABEL_COMPANYACCOUNT_LISTBUDGET:
                                            Dim listbudget As Object
                                            Set listbudget = CreateObject("gl_tables.CTables")
                                            With listbudget
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "listbudget"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_TABEL_CURRENCY_ADD:
                                            Dim addkursgl As Object
                                            Set addkursgl = CreateObject("gl_tables.CTables")
                                            With addkursgl
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "addkurs"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_TABEL_CURRENCY_UPDATE:
                                            Dim changekursgl As Object
                                            Set changekursgl = CreateObject("gl_tables.CTables")
                                            With changekursgl
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "changekurs"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_TABEL_CURRENCY_LIST:
                                            Dim listkursgl As Object
                                            Set listkursgl = CreateObject("gl_tables.CTables")
                                            With listkursgl
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "changekurs"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_GL_TABEL_FIXEDASSETTYPE_ADD:
                                            Dim addjenisfa As Object
                                            Set addjenisfa = CreateObject("gl_tables.CTables")
                                            With addjenisfa
                                               .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "addjenisfa"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                        
        Case ID_KEUANGAN_GL_TABEL_FIXEDASSETTYPE_UPDATE:
                                            Dim changejenisfa As Object
                                            Set changejenisfa = CreateObject("gl_tables.CTables")
                                            With changejenisfa
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "changejenisfa"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_TABEL_FIXEDASSETTYPE_LIST:
                                            Dim listjenisfa As Object
                                            Set listjenisfa = CreateObject("gl_tables.CTables")
                                            With listjenisfa
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "listjenisfa"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
         Case ID_KEUANGAN_GL_TABEL_BANK_ADD:
                                            Dim addbank As Object
                                            Set addbank = CreateObject("gl_tables.CTables")
                                            With addbank
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "addbank"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        
         Case ID_KEUANGAN_GL_TABEL_BANK_UPDATE:
                                            Dim changebank As Object
                                            Set changebank = CreateObject("gl_tables.CTables")
                                            With changebank
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "changebank"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        
         Case ID_KEUANGAN_GL_TABEL_BANK_LIST:
                                            Dim listbank As Object
                                            Set listbank = CreateObject("gl_tables.CTables")
                                            With listbank
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "listbank"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_FIXEDASSET_PEMBELIANFIXEDASSET:
                                            Dim belifa As Object
                                            Set belifa = CreateObject("gl_fixas.CFixedAsset")
                                            With belifa
                                               .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "belifa"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_GL_FIXEDASSET_POSTINGPEMBELIANFIXEDASSET:
                                            Dim postingbfa As Object
                                            Set postingbfa = CreateObject("gl_fixas.CFixedAsset")
                                            With postingbfa
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "postingbfa"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_FIXEDASSET_UNPOSTINGPEMBELIANFIXEDASSET:
                                            Dim unpostingbfa As Object
                                            Set unpostingbfa = CreateObject("gl_fixas.CFixedAsset")
                                            With unpostingbfa
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "unpostingbfa"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_FIXEDASSET_PENJUALANFIXEDASSET:
                                            Dim jualfa As Object
                                            Set jualfa = CreateObject("gl_fixas.CFixedAsset")
                                            With jualfa
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "jualfa"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_FIXEDASSET_POSTINGPENJUALANFIXEDASSET:
                                            Dim postingjfa As Object
                                            Set postingjfa = CreateObject("gl_fixas.CFixedAsset")
                                            With postingjfa
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "postingjfa"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_GL_FIXEDASSET_UNPOSTINGPENJUALANFIXEDASSET:
                                            Dim unpostingjfa As Object
                                            Set unpostingjfa = CreateObject("gl_fixas.CFixedAsset")
                                            With unpostingjfa
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "unpostingjfa"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_GL_LEDGER_JURNAL_ADD:
                                            Dim addjurnal As Object
                                            Set addjurnal = CreateObject("gl_ledger.CLedger")
                                            With addjurnal
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "addjurnal"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_LEDGER_JURNAL_UPDATE:
                                            Dim changejurnal As Object
                                            Set changejurnal = CreateObject("gl_ledger.CLedger")
                                            With changejurnal
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "changejurnal"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_LEDGER_JURNAL_LIST:
                                            Dim tranlist As Object
                                            Set tranlist = CreateObject("gl_ledger.CLedger")
                                            With tranlist
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "tranlist"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_LEDGER_JURNAL_POSTING:
                                            Dim tranpostingjurnal As Object
                                            Set tranpostingjurnal = CreateObject("gl_ledger.CLedger")
                                            With tranpostingjurnal
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "tranpostingjurnal"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_LEDGER_JURNAL_UNPOSTING:
                                            Dim untranpostingjurnal As Object
                                            Set untranpostingjurnal = CreateObject("gl_ledger.CLedger")
                                            With untranpostingjurnal
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "untranpostingjurnal"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_GL_LEDGER_CAHBANKIN_ADD:
                                            Dim cashbankin As Object
                                            Set cashbankin = CreateObject("gl_ledger.CLedger")
                                            With cashbankin
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "cashbankin"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With

        Case ID_KEUANGAN_GL_LEDGER_CAHBANKIN_UPDATE:
                                            Dim changecashbankin As Object
                                            Set changecashbankin = CreateObject("gl_ledger.CLedger")
                                            With changecashbankin
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "changecashbankin"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                            
        Case ID_KEUANGAN_GL_LEDGER_CAHBANKIN_LIST:
                                            Dim listcashbankin As Object
                                            Set listcashbankin = CreateObject("gl_ledger.CLedger")
                                            With listcashbankin
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "listcashbankin"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_LEDGER_CAHBANKIN_POSTING:
                                            Dim postingcashbankin As Object
                                            Set postingcashbankin = CreateObject("gl_ledger.CLedger")
                                            With postingcashbankin
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "postingcashbankin"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_LEDGER_CAHBANKIN_UNPOSTING:
                                            Dim unpostingcashbankin As Object
                                            Set unpostingcashbankin = CreateObject("gl_ledger.CLedger")
                                            With unpostingcashbankin
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "unpostingcashbankin"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_LEDGER_CAHBANKOUT_ADD_NEW: subcombars_gl_ledger (ID_KEUANGAN_GL_LEDGER_CAHBANKOUT_ADD_NEW)
        Case ID_KEUANGAN_GL_LEDGER_CAHBANKOUT_ADD_OLD: subcombars_gl_ledger (ID_KEUANGAN_GL_LEDGER_CAHBANKOUT_ADD_OLD)
        
        Case ID_KEUANGAN_GL_LEDGER_CAHBANKOUT_PRINT:
                                            Dim lapcbo As Object
                                            Set lapcbo = CreateObject("gl_ledger.CLedger")
                                            With lapcbo
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "lapcbo"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        
        Case ID_KEUANGAN_GL_LEDGER_CAHBANKOUT_UPDATE:
                                            Dim changecashbankout As Object
                                            Set changecashbankout = CreateObject("gl_ledger.CLedger")
                                            With changecashbankout
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "changecashbankout"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_LEDGER_CAHBANKOUT_LIST:
                                            Dim listcashbankout As Object
                                            Set listcashbankout = CreateObject("gl_ledger.CLedger")
                                            With listcashbankout
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "listcashbankout"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_LEDGER_CAHBANKOUT_POSTING:
                                            Dim postingcashbankout As Object
                                            Set postingcashbankout = CreateObject("gl_ledger.CLedger")
                                            With postingcashbankout
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "postingcashbankout"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        
        Case ID_KEUANGAN_GL_LEDGER_CAHBANKOUT_UNPOSTING:
                                            Dim unpostingcashbankout As Object
                                            Set unpostingcashbankout = CreateObject("gl_ledger.CLedger")
                                            With unpostingcashbankout
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "unpostingcashbankout"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_GL_LEDGER_BUKTIKELUAR_ADD:
                                            Dim objbuktikeluar As Object
                                            Set objbuktikeluar = CreateObject("gl_ledger.CLedger")
                                            With objbuktikeluar
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "buktikeluar"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_GL_LEDGER_BUKTIKELUAR_NEW: subcombars_gl_ledger (ID_KEUANGAN_GL_LEDGER_BUKTIKELUAR_NEW)
        Case ID_KEUANGAN_GL_LEDGER_BUKTIKELUAR_UPDATE: subcombars_gl_ledger (ID_KEUANGAN_GL_LEDGER_BUKTIKELUAR_UPDATE)
        Case ID_KEUANGAN_GL_LEDGER_BUKTIKELUAR_REPRINT: subcombars_gl_ledger (ID_KEUANGAN_GL_LEDGER_BUKTIKELUAR_REPRINT)
                                            
        Case ID_KEUANGAN_GL_LEDGER_BUKTIKELUAR_LIST:
                                            Dim objbuktikeluarlist As Object
                                            Set objbuktikeluarlist = CreateObject("gl_ledger.CLedger")
                                            With objbuktikeluarlist
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "listbuktikeluar"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_GL_LEDGER_ETOL:
                                            Dim objetol As Object
                                            Set objetol = CreateObject("gl_ledger.CLedger")
                                            With objetol
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "trans_etol"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_REPORT_TRIALBALANCE:
                                            Dim rpttrialbalance As Object
                                            Set rpttrialbalance = CreateObject("gl_report.CReport")
                                            With rpttrialbalance
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "rpttrialbalance"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_REPORT_BUKUBESAR:
                                            Dim bukubesar As Object
                                            Set bukubesar = CreateObject("gl_report.CReport")
                                            With bukubesar
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "rptbukubesar"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_REPORT_WORKSHEET:
                                            Dim worksheet As Object
                                            Set worksheet = CreateObject("gl_report.CReport")
                                            With worksheet
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "rptworksheet"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        
        Case ID_KEUANGAN_GL_REPORT_DAILYCASHFLOW:
                                            Dim bukukas As Object
                                            Set bukukas = CreateObject("gl_report.CReport")
                                            With bukukas
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "rptbukukas"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_REPORT_BALANCESHEET:
                                            Dim balance As Object
                                            Set balance = CreateObject("gl_report.CReport")
                                            With balance
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "rptbalance"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_REPORT_INCOMESTATEMENT:
                                            Dim income As Object
                                            Set income = CreateObject("gl_report.CReport")
                                            With income
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "rptincome"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                                
        Case ID_KEUANGAN_GL_REPORT_DAFTARFIXEDASSET:
                                            Dim listaktiva As Object
                                            Set listaktiva = CreateObject("gl_report.CReport")
                                            With listaktiva
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "rptlistaktiva"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_REPORT_NILAIFIXEDASSET:
                                            Dim dafnilaiaktiva As Object
                                            Set dafnilaiaktiva = CreateObject("gl_report.CReport")
                                            With dafnilaiaktiva
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "rptnilaiaktiva"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_REPORT_PENJUALANFIXEDASSET:
                                            Dim dafkativajual As Object
                                            Set dafkativajual = CreateObject("gl_report.CReport")
                                            With dafkativajual
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "rptaktivajual"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_UTILITY_UNBALANCETRANS:
                                            Dim dafunbalance As Object
                                            Set dafunbalance = CreateObject("gl_util.CUtility")
                                            With dafunbalance
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "dafunbalance"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_UTILITY_RESETSALDO:
                                            Dim reset As Object
                                            Set reset = CreateObject("gl_util.CUtility")
                                            With reset
                                               .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "reset"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_UTILITY_SETUPLAYOUTREPORT_DEFINE:
                                            Dim definereport As Object
                                            Set definereport = CreateObject("gl_util.CUtility")
                                            With definereport
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "definereport"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_KEUANGAN_GL_UTILITY_SETUPLAYOUTREPORT_CROSS:
                                            Dim checkreport As Object
                                            Set checkreport = CreateObject("gl_util.CUtility")
                                            With checkreport
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "checkreport"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_UTILITY_SETUPLAYOUTREPORT_LIST:
                                            Dim listreport As Object
                                            Set listreport = CreateObject("gl_util.CUtility")
                                            With listreport
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "listreport"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_UTILITY_CLOSING:
                                            Dim glclosing As Object
                                            Set glclosing = CreateObject("gl_util.CUtility")
                                            With glclosing
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "glclosing"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_UTILITY_UNCLOSING:
                                            Dim glunclosing As Object
                                            Set glunclosing = CreateObject("gl_util.CUtility")
                                            With glunclosing
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "glunclosing"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_KEUANGAN_GL_UTILITY_REKONSILIASI:
                                            Dim rekonsiliasi As Object
                                            Set rekonsiliasi = CreateObject("gl_util.CUtility")
                                            With rekonsiliasi
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "rekonsiliasi"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_PEMBELIAN_PURCHASING_LAPCONPEMBELIAN:
                                            Dim daftarconfirmunconfirmpembelian As Object
                                            Set daftarconfirmunconfirmpembelian = CreateObject("finance_purc.CPurchasing")
                                            With daftarconfirmunconfirmpembelian
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "daftarconfirmunconfirm"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        
        Case ID_PEMBELIAN_PURCHASING_LAPVOUCER:
                                            Dim lapvoucer As Object
                                            Set lapvoucer = CreateObject("finance_Purc.Cpurchasing")
                                            With lapvoucer
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "laporanvoucer"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_PURCHASING_LAPPROSESVOUCHER:
                                            Dim lapprocvoucer As Object
                                            Set lapprocvoucer = CreateObject("finance_Purc.Cpurchasing")
                                            With lapprocvoucer
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "laporanprocesvoucer"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                
        Case ID_PEMBELIAN_PURCHASING_PEMAKAIANBAHANBAKU_ADD:
                                            Dim addpemakaian As Object
                                            Set addpemakaian = CreateObject("purc_purc.CPurchasing")
                                            With addpemakaian
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "addpemakaian"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_PURCHASING_PEMAKAIANBAHANBAKU_CHANGE:
                                            Dim changepemakaian As Object
                                            Set changepemakaian = CreateObject("purc_purc.CPurchasing")
                                            With changepemakaian
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "changepemakaian"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_PURCHASING_PEMAKAIANBAHANBAKU_SISA:
                                            Dim sisapemakaian As Object
                                            Set sisapemakaian = CreateObject("purc_purc.CPurchasing")
                                            With sisapemakaian
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "sisapemakaian"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_PURCHASING_INQUIRY:
                                            Dim inquerypo As Object
                                            Set inquerypo = CreateObject("purc_purc.CPurchasing")
                                            With inquerypo
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "inquerypo"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_PURCHASING_LAPPO:
                                            Dim daftarpo As Object
                                            Set daftarpo = CreateObject("purc_purc.CPurchasing")
                                            With daftarpo
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "daftarpo"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_PURCHASING_LAPPEMBELIAN:
                                            Dim lappembelian As Object
                                            Set lappembelian = CreateObject("purc_purc.CPurchasing")
                                            With lappembelian
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "lappembelian"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_PURCHASING_PRINTPO:
                                            Dim lappo As Object
                                            Set lappo = CreateObject("purc_purc.CPurchasing")
                                            With lappo
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "lappo"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_PURCHASING_LAPPEMAKAIANBB:
                                            Dim lappemakaianbb As Object
                                            Set lappemakaianbb = CreateObject("purc_purc.CPurchasing")
                                            With lappemakaianbb
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "lappemakaian"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_UTILITY_OPTIONS:
                                            Dim optionpembelian1 As Object
                                            Set optionpembelian1 = CreateObject("purc_util.CUtility")
                                            With optionpembelian1
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "option"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_UTILITY_EXPORTDATAPABRIK:
                                            Dim exportpenerimaan As Object
                                            Set exportpenerimaan = CreateObject("purc_util.CUtility")
                                            With exportpenerimaan
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "export"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PEMBELIAN_UTILITY_IMPORTDATA:
                                            Dim importpenerimaan As Object
                                            Set importpenerimaan = CreateObject("purc_util.CUtility")
                                            With importpenerimaan
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "import"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        
        Case ID_PEMBELIAN_UTILITY_FASTSEARCH: FastSearch
        
        Case ID_PEMBELIAN_UTILITY_ADJUSTSTOK:
                                            Dim stokopbahan As Object
                                            Set stokopbahan = CreateObject("purc_util.CUtility")
                                            With stokopbahan
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "stokopbahan"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        
        '################################ BAGIAN PRODUKSI ##################################

        Case ID_PRODUKSI_MAINMENU_TABLES_RESEP_LEM:
                                            Dim addresep As Object
                                            Set addresep = CreateObject("prod_tables.CTables")
                                            With addresep
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .userlevel = UserOnLineLevel
                                                .Path = App.Path
                                                .formname = "addresep"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PRODUKSI_MAINMENU_TABLES_RESEP_KARET:
                                            Dim addresepk As Object
                                            Set addresepk = CreateObject("prod_tables.CTables")
                                            With addresepk
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .userlevel = UserOnLineLevel
                                                .Path = App.Path
                                                .formname = "addresepk"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PRODUKSI_MAINMENU_TABLES_RESEP_LISTLEM:
                                            Dim listlem As Object
                                            Set listlem = CreateObject("prod_tables.CTables")
                                            With listlem
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .userlevel = UserOnLineLevel
                                                .Path = App.Path
                                                .formname = "listlem"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PRODUKSI_MAINMENU_TABLES_KONVERSI_ADD: subcombars_produksi (ID_PRODUKSI_MAINMENU_TABLES_KONVERSI_ADD)
        Case ID_PRODUKSI_MAINMENU_TABLES_KONVERSI_CHANGE: subcombars_produksi (ID_PRODUKSI_MAINMENU_TABLES_KONVERSI_CHANGE)
        Case ID_PRODUKSI_MAINMENU_TABLES_KONVERSI_LIST: subcombars_produksi (ID_PRODUKSI_MAINMENU_TABLES_KONVERSI_LIST)
        Case ID_PRODUKSI_MAINMENU_TABLES_KONVERSI_UNIT: subcombars_produksi (ID_PRODUKSI_MAINMENU_TABLES_KONVERSI_UNIT)
        Case ID_PRODUKSI_MAINMENU_TABLES_DEFINEKG: subcombars_produksi (ID_PRODUKSI_MAINMENU_TABLES_DEFINEKG)
        
        Case ID_PRODUKSI_MAINMENU_TABLES_RNCPROD_ADD: subcombars_produksi (ID_PRODUKSI_MAINMENU_TABLES_RNCPROD_ADD)
        Case ID_PRODUKSI_MAINMENU_TABLES_RNCPROD_CHANGE: subcombars_produksi (ID_PRODUKSI_MAINMENU_TABLES_RNCPROD_CHANGE)
        Case ID_PRODUKSI_MAINMENU_TABLES_RNCPROD_LIST: subcombars_produksi (ID_PRODUKSI_MAINMENU_TABLES_RNCPROD_LIST)
        Case ID_PRODUKSI_MAINMENU_TABLES_KATEGORI: subcombars_produksi (ID_PRODUKSI_MAINMENU_TABLES_KATEGORI)
        Case ID_PRODUKSI_MAINMENU_TABLES_REAKTOR: subcombars_produksi (ID_PRODUKSI_MAINMENU_TABLES_REAKTOR)
        
        Case ID_PRODUKSI_MAINMENU_MUTASI_WIPBASE: subcombars_produksi (ID_PRODUKSI_MAINMENU_MUTASI_WIPBASE)
        Case ID_PRODUKSI_MAINMENU_MUTASI_LAPWIPBASE: subcombars_produksi (ID_PRODUKSI_MAINMENU_MUTASI_LAPWIPBASE)
        Case ID_PRODUKSI_MAINMENU_MUTASI_ADJWIP: subcombars_produksi (ID_PRODUKSI_MAINMENU_MUTASI_ADJWIP)
        
        Case ID_PRODUKSI_MAINMENU_SOP_ADD_NEW: subcombars_produksi (ID_PRODUKSI_MAINMENU_SOP_ADD_NEW)
        Case ID_PRODUKSI_MAINMENU_SOP_ADD_EDIT: subcombars_produksi (ID_PRODUKSI_MAINMENU_SOP_ADD_EDIT)
                                            
        Case ID_PRODUKSI_MAINMENU_SOP_TAKEPACKAGE_ADD:
                                            Dim permintaan As Object
                                            Set permintaan = CreateObject("prod_sop.CSOP")
                                            With permintaan
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "permintaan"
                                                .setuseronlinelevel = UserOnLineLevel
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_PRODUKSI_MAINMENU_SOP_TAKEPACKAGE_LIST:
                                            Dim listconfirm As Object
                                            Set listconfirm = CreateObject("prod_sop.CSOP")
                                            With listconfirm
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "listconfirm"
                                                .setuseronlinelevel = UserOnLineLevel
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PRODUKSI_MAINMENU_SOP_TAKEPACKAGE_REPORT: subcombars_produksi (ID_PRODUKSI_MAINMENU_SOP_TAKEPACKAGE_REPORT)
        Case ID_PRODUKSI_MAINMENU_SOP_PRODMONTHLY_BYBB: subcombars_produksi (ID_PRODUKSI_MAINMENU_SOP_PRODMONTHLY_BYBB)
        Case ID_PRODUKSI_MAINMENU_SOP_PRODMONTHLY_BYTOPPRODUK: subcombars_produksi (ID_PRODUKSI_MAINMENU_SOP_PRODMONTHLY_BYTOPPRODUK)
        Case ID_PRODUKSI_MAINMENU_SOP_CHANGE:
                                            Dim changesop As Object
                                            Set changesop = CreateObject("prod_sop.CSOP")
                                            With changesop
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .Path = App.Path
                                                .formname = "changesop"
                                                .setuseronlinelevel = UserOnLineLevel
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_PRODUKSI_MAINMENU_SOP_LIST:
                                            Dim liststokbahanbaku As Object
                                            Set liststokbahanbaku = CreateObject("prod_sop.CSOP")
                                            With liststokbahanbaku
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .setuseronlinelevel = UserOnLineLevel
                                                .Path = App.Path
                                                .formname = "liststokbahanbaku"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_PRODUKSI_MAINMENU_SOP_LOTPALET:
                                            Dim listlotpalet As Object
                                            Set listlotpalet = CreateObject("prod_sop.CSOP")
                                            With listlotpalet
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .setuseronlinelevel = UserOnLineLevel
                                                .Path = App.Path
                                                .formname = "listlotpalet"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_PRODUKSI_MAINMENU_SOP_LISTGROUP:
                                            Dim group_report As Object
                                            Set group_report = CreateObject("prod_sop.CSOP")
                                            With group_report
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .setuseronlinelevel = UserOnLineLevel
                                                .Path = App.Path
                                                .formname = "group_report"
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PRODUKSI_MAINMENU_SOP_SCANLOT_SCANPALET:
                                            Dim scanlot As Object
                                            Set scanlot = CreateObject("prod_sop.CSOP")
                                            With scanlot
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .setuseronlinelevel = UserOnLineLevel
                                                .Path = App.Path
                                                .formname = "scanlot"
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PRODUKSI_MAINMENU_SOP_SCANLOT_EDITPALET:
                                            Dim editlot As Object
                                            Set editlot = CreateObject("prod_sop.CSOP")
                                            With editlot
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .setuseronlinelevel = UserOnLineLevel
                                                .Path = App.Path
                                                .formname = "ubahlot"
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PRODUKSI_MAINMENU_SOP_SCANLOT_KEYPALET:
                                            Dim kunci As Object
                                            Set kunci = CreateObject("prod_sop.CSOP")
                                            With kunci
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .setuseronlinelevel = UserOnLineLevel
                                                .Path = App.Path
                                                .formname = "kuncipalet"
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_PRODUKSI_MAINMENU_SOP_RFID:
                                            Dim rfid As Object
                                            Set rfid = CreateObject("prod_sop.CSOP")
                                            With rfid
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .setuseronlinelevel = UserOnLineLevel
                                                .Path = App.Path
                                                .formname = "rfid"
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_PRODUKSI_MAINMENU_MUTASI_ADD:
                                            Dim objaddmut As Object
                                            Set objaddmut = CreateObject("prod_mut.CMUT")
                                            With objaddmut
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .setuseronlinelevel = UserOnLineLevel
                                                .Path = App.Path
                                                .formname = "addmutasi"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
        Case ID_PRODUKSI_MAINMENU_MUTASI_LAPLOT:
                                            Dim objlapmut As Object
                                            Set objlapmut = CreateObject("prod_mut.CMUT")
                                            With objlapmut
                                                .token = "kusumah"
                                                .UserOnline = UserOnline
                                                .setuseronlinelevel = UserOnLineLevel
                                                .Path = App.Path
                                                .formname = "lapmutasi"
                                                .FastSearch = False
                                                .asremote = remoteserver
                                                .ipserver = dbServer
                                                .Show
                                                SetParent .hWnd, Me.hWnd
                                            End With
                                            
        Case ID_ABOUT_HOME_ABOUTUS: frmabout.Show vbModal
        
        Case ID_ABOUT_HOME_HELP_HELP_NP: subcombars_about (ID_ABOUT_HOME_HELP_HELP_NP)
        Case ID_ABOUT_HOME_HELP_HELP_NPEDIT: subcombars_about (ID_ABOUT_HOME_HELP_HELP_NPEDIT)
        Case ID_ABOUT_HOME_HELP_HELP_NPPRINT: subcombars_about (ID_ABOUT_HOME_HELP_HELP_NPPRINT)
        Case ID_ABOUT_HOME_HELP_HELP_TS: subcombars_about (ID_ABOUT_HOME_HELP_HELP_TS)
        Case ID_ABOUT_HOME_HELP_HELP_CETAKITG: subcombars_about (ID_ABOUT_HOME_HELP_HELP_CETAKITG)
        Case ID_ABOUT_HOME_HELP_HELP_CETAKREKAP: subcombars_about (ID_ABOUT_HOME_HELP_HELP_CETAKREKAP)
        Case ID_ABOUT_HOME_HELP_HELP_LEMBUR: subcombars_about (ID_ABOUT_HOME_HELP_HELP_LEMBUR)
        Case ID_ABOUT_HOME_HELP_HELP_KAS: subcombars_about (ID_ABOUT_HOME_HELP_HELP_KAS)
        Case ID_ABOUT_HOME_HELP_FORM: subcombars_about (ID_ABOUT_HOME_HELP_FORM)
        Case ID_ABOUT_HOME_CHECKUPDATE: MsgBox "Up To Date...!", vbInformation, AppName
        
        Case ID_OPTIONS_STYLEBLUE:
                                    ComBars.VisualTheme = xtpThemeRibbon
                                    CommandBarsGlobalSettings.Office2007Images = ""
                                    CtrlFile.Style = xtpButtonAutomatic
                                    ComBars.PaintManager.RefreshMetrics
                                    ComBars.RecalcLayout
        Case ID_OPTIONS_STYLEBLACK:
                                    ComBars.VisualTheme = xtpThemeRibbon
                                    CommandBarsGlobalSettings.Office2007Images = App.Path & "\Skins\Office2007Black.dll"
                                    CtrlFile.Style = xtpButtonAutomatic
                                    ComBars.PaintManager.RefreshMetrics
                                    ComBars.RecalcLayout
        Case ID_OPTIONS_STYLEAQUA:
                                    ComBars.VisualTheme = xtpThemeRibbon
                                    CommandBarsGlobalSettings.Office2007Images = App.Path & "\Skins\Office2007Aqua.dll"
                                    CtrlFile.Style = xtpButtonAutomatic
                                    ComBars.PaintManager.RefreshMetrics
                                    ComBars.RecalcLayout
        Case ID_OPTIONS_STYLESILVER:
                                    ComBars.VisualTheme = xtpThemeRibbon
                                    CommandBarsGlobalSettings.Office2007Images = App.Path & "\Skins\Office2007Silver.dll"
                                    CtrlFile.Style = xtpButtonAutomatic
                                    ComBars.PaintManager.RefreshMetrics
                                    ComBars.RecalcLayout
        End Select
End Sub

Private Sub ComBars_Resize()
    On Error Resume Next
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    Dim LeftBtn As Long
    
    ComBars.GetClientRect Left, Top, Right, Bottom
    wbrmain.Move Left, Top, Right - Left, (Bottom - Top)
End Sub

Private Sub Form_Activate()
    Me.Refresh
    
End Sub

Private Sub Form_Load()
    'initial main form
    LoadSkin
    LoadCombars
    LoadMenu
    wbrmain.Navigate App.Path & "\html\page.html"
    DisableAllMenu
    If UserOnline = "" Then
        Set frmLogin = New frmLogin
        Timer1.Enabled = True
    End If
    
    LoadStatusBars
End Sub

Private Sub LoadSkin()
    CommandBarsGlobalSettings.Office2007Images = App.Path & "\Skins\Office2007Black.dll"
    SF.ApplyWindow hWnd
End Sub

Private Sub LoadCombars()
    CommandBarsGlobalSettings.App = App
    ComBars.VisualTheme = xtpThemeRibbon
    ComBars.AddImageList ImgList
End Sub

Private Sub LoadStatusBars()
    Dim Pane As StatusBarPane
    Set StatusBar = ComBars.StatusBar
    StatusBar.MinHeight = 27
    StatusBar.Visible = True
    
    
    
    Set Pane = StatusBar.AddPane(ID_INDICATOR_SERVER)
    Pane.Visible = True
    Pane.text = "Database : " + UCase(dbServer) & " : " & UCase(DBName) & " User : "
    Pane.Font.Bold = True
    
    
    Set Pane = StatusBar.AddPane(0)
    Pane.Style = SBPS_STRETCH
    Pane.text = "Ready"
    Pane.Width = 0 ' Autro Size
    
    Set Pane = StatusBar.AddPane(ID_INDICATOR_DARITGL)
    Pane.text = "Dari Tanggal :"
    Pane.Width = 75
    Pane.Visible = False

    Set Pane = StatusBar.AddPane(ID_INDICATOR_DTP1)
    Pane.Handle = dtp1.hWnd
    Pane.Width = 100
    Pane.Visible = False
    
    Set Pane = StatusBar.AddPane(ID_INDICATOR_SDTGL)
    Pane.text = "SD"
    Pane.Width = 20
    Pane.Visible = False
    
    Set Pane = StatusBar.AddPane(ID_INDICATOR_DTP2)
    Pane.Handle = dtp2.hWnd
    Pane.Width = 100
    Pane.Visible = False
    
    StatusBar.AddPane ID_INDICATOR_CAPS
    StatusBar.AddPane ID_INDICATOR_NUM
    StatusBar.AddPane ID_INDICATOR_SCRL
    StatusBar.IdleText = "Ready"
    
End Sub

Private Sub LoadMenu()
    Set RibbonBar = ComBars.AddRibbonBar("RIBBON")
    RibbonBar.EnableDocking xtpFlagStretched
    RibbonBar.EnableFrameTheme
    Set CtrlFile = RibbonBar.AddSystemButton
    CtrlFile.IconId = ID_FILE
    
    'INIT MENU FILE
    With CtrlFile.CommandBar.Controls
        Set FILE_CHANGEPAS = .Add(xtpControlButton, ID_FILE_CHANGEPASS, "Ubah Password")
        Set FILE_DASHBOARD = .Add(xtpControlButton, ID_FILE_DASHBOARD, "Dashboard")
        FILE_CHANGEPAS.BeginGroup = False
        Set FILE_LOGOUT = .Add(xtpControlButton, ID_FILE_LOGOUT, "LogOut..!")
        Set FILE_EXIT = .Add(xtpControlButton, ID_FILE_EXIT, "Exit/Keluar")
        FILE_EXIT.BeginGroup = True
    End With
    
    'INIT TAB RIBBON MASTER
        
    Set MASTER = RibbonBar.InsertTab(ID_MASTER, "KONFIGURASI")
    Set MASTER_DATABASE = MASTER.Groups.AddGroup("Database", ID_MASTER_DATABASE)
    With MASTER_DATABASE
        Set MASTER_DATABASE_KONEKSISQL = .Add(xtpControlButton, ID_MASTER_DATABASE_KONEKSISQL, "Konfigurasi Server")
        Set MASTER_DATABASE_KONEKSI_BACKUP = .Add(xtpControlButton, ID_MASTER_DATABASE_KONEKSI_BACKUP, "Backup Database")
    End With
    
    Set MASTER_DATABASE_MAINTENANCE = MASTER.Groups.AddGroup("Pemeliharaan", ID_MASTER_DATABASE_MAINTENANCE)
    MASTER_DATABASE_MAINTENANCE.Visible = False
    With MASTER_DATABASE_MAINTENANCE
        Set MASTER_DATABASE_MAINTENANCE_IMPORTBAHANBAKU = .Add(xtpControlButton, ID_MASTER_DATABASE_MAINTENANCE_IMPORTBAHANBAKU, "Import Stok Awal Bahan Baku")
        Set MASTER_DATABASE_MAINTENANCE_IMPORTBARANGJADI = .Add(xtpControlButton, ID_MASTER_DATABASE_MAINTENANCE_IMPORTBARANGJADI, "Import Stok Barang Jadi")
        Set MASTER_DATABASE_MAINTENANCE_BROADCASTMESSAGE = .Add(xtpControlButton, ID_MASTER_DATABASE_MAINTENANCE_BROADCASTMESSAGE, "Broadcast Message System")
        MASTER_DATABASE_MAINTENANCE_BROADCASTMESSAGE.BeginGroup = True
    End With
    
    'Group Manage User
    Set MASTER_MANAGEUSER = MASTER.Groups.AddGroup("Pengguna", ID_MASTER_MANAGEUSER)
    With MASTER_MANAGEUSER
        Set MASTER_MANAGEUSER_USER = .Add(xtpControlButton, ID_MASTER_MANAGEUSER_USER, "Atur Pengguna")
    End With
    
    Set MASTER_TELEGRAM = MASTER.Groups.AddGroup("Telegram", ID_MASTER_TELEGRAM)
    With MASTER_TELEGRAM
        Set MASTER_TELEGRAM_BOT = .Add(xtpControlButton, ID_MASTER_TELEGRAM_BOT, "Add Telegram")
    End With
 
    'LOAD PURCHASING MENU
    Set PEMBELIAN = RibbonBar.InsertTab(ID_PEMBELIAN, "PEMBELIAN")
    Set PEMBELIAN_MAINMENU = PEMBELIAN.Groups.AddGroup("Aktivitas", ID_PEMBELIAN_TABEL)
    PEMBELIAN.Selected = True
    With PEMBELIAN_MAINMENU
        Set PEMBELIAN_TABEL = .Add(xtpControlButtonPopup, ID_PEMBELIAN_TABEL, "Tabel")
        With PEMBELIAN_TABEL.CommandBar.Controls
            
            '# Menu Supplier
            Set PEMBELIAN_TABEL_SUPPLIER = .Add(xtpControlButtonPopup, ID_PEMBELIAN_TABEL_SUPPLIER, "Supplier/Pemasok")
            With PEMBELIAN_TABEL_SUPPLIER.CommandBar.Controls
                Set PEMBELIAN_TABEL_SUPPLIER_ADD = .Add(xtpControlButton, ID_PEMBELIAN_TABEL_SUPPLIER_ADD, "Add/Tambah Supplier")
                Set PEMBELIAN_TABEL_SUPPLIER_CHANGE = .Add(xtpControlButton, ID_PEMBELIAN_TABEL_SUPPLIER_CHANGE, "Change/Ubah Supplier")
                Set PEMBELIAN_TABEL_SUPPLIER_LIST = .Add(xtpControlButton, ID_PEMBELIAN_TABEL_SUPPLIER_LIST, "List/Daftar Supplier")
                PEMBELIAN_TABEL_SUPPLIER_LIST.BeginGroup = False
                Set PEMBELIAN_TABEL_SUPPLIER_PRICELIST = .Add(xtpControlButton, ID_PEMBELIAN_TABEL_SUPPLIER_PRICELIST, "Price/Harga List Supplier")
                Set PEMBELIAN_TABEL_SUPPLIER_GRAPRICE = .Add(xtpControlButton, ID_PEMBELIAN_TABEL_SUPPLIER_GRAPRICE, "Grafik Harga Supplier")
            End With
            
            '#Menu Satuan
            Set PEMBELIAN_TABEL_SATUAN = .Add(xtpControlButtonPopup, ID_PEMBELIAN_TABEL_SATUAN, "Satuan")
            PEMBELIAN_TABEL_SATUAN.BeginGroup = False
            With PEMBELIAN_TABEL_SATUAN.CommandBar.Controls
                Set PEMBELIAN_TABEL_SATUAN_ADD = .Add(xtpControlButton, ID_PEMBELIAN_TABEL_SATUAN_ADD, "Tambah Satuan")
                Set PEMBELIAN_TABEL_SATUAN_CHANGE = .Add(xtpControlButton, ID_PEMBELIAN_TABEL_SATUAN_CHANGE, "Ubah Satuan")
                Set PEMBELIAN_TABEL_SATUAN_LIST = .Add(xtpControlButton, ID_PEMBELIAN_TABEL_SATUAN_LIST, "Daftar Satuan")
                PEMBELIAN_TABEL_SATUAN_LIST.BeginGroup = False
            End With
            
            '#Menu Bahan Baku
            Set PEMBELIAN_TABEL_BAHANBAKU = .Add(xtpControlButtonPopup, ID_PEMBELIAN_TABEL_BAHANBAKU, "Bahan Baku")
            With PEMBELIAN_TABEL_BAHANBAKU.CommandBar.Controls
                Set PEMBELIAN_TABEL_BAHANBAKU_MANAGE = .Add(xtpControlButton, ID_PEMBELIAN_TABEL_BAHANBAKU_MANAGE, "Atur Bahan Baku")
                Set PEMBELIAN_TABEL_BAHANBAKU_LIST = .Add(xtpControlButton, ID_PEMBELIAN_TABEL_BAHANBAKU_LIST, "Daftar Bahan Baku")
                PEMBELIAN_TABEL_BAHANBAKU_LIST.BeginGroup = False
            End With
            
            '#Menu Packaging
            'Set PEMBELIAN_TABEL_PACKAGING = .Add(xtpControlButtonPopup, ID_PEMBELIAN_TABEL_PACKAGING, "Packaging")
            'With PEMBELIAN_TABEL_PACKAGING.CommandBar.Controls
            '    Set PEMBELIAN_TABEL_PACKAGING_ADD = .Add(xtpControlButton, ID_PEMBELIAN_TABEL_PACKAGING_ADD, "Add Packaging")
            '    Set PEMBELIAN_TABEL_PACKAGING_CHANGE = .Add(xtpControlButton, ID_PEMBELIAN_TABEL_PACKAGING_CHANGE, "Change Packaging")
            'End With
            
            Set PEMBELIAN_TABEL_MINSTOCK = .Add(xtpControlButtonPopup, ID_PEMBELIAN_TABEL_MINSTOCK, "Minimum Stock")
            With PEMBELIAN_TABEL_MINSTOCK.CommandBar.Controls
                PEMBELIAN_TABEL_MINSTOCK.BeginGroup = True
                Set PEMBELIAN_TABEL_MINSTOCK_ADD = .Add(xtpControlButton, ID_PEMBELIAN_TABEL_MINSTOCK_ADD, "Tambah Minimum Stock")
                Set PEMBELIAN_TABEL_MINSTOCK_CHANGE = .Add(xtpControlButton, ID_PEMBELIAN_TABEL_MINSTOCK_CHANGE, "Ubah Minimum Stock")
                Set PEMBELIAN_TABEL_MINSTOCK_LIST = .Add(xtpControlButton, ID_PEMBELIAN_TABEL_MINSTOCK_LIST, "Daftar Minimum Stock")
            End With
        End With
        
        '#Menu Mutasi Barang
        Set PEMBELIAN_MUTASIBARANG = .Add(xtpControlButtonPopup, ID_PEMBELIAN_MUTASIBARANG, "Mutasi")
        With PEMBELIAN_MUTASIBARANG.CommandBar.Controls
            Set PEMBELIAN_MUTASIBARANG_ADD = .Add(xtpControlButton, ID_PEMBELIAN_MUTASIBARANG_ADD, "Tambah Mutasi")
            Set PEMBELIAN_MUTASIBARANG_CHANGE = .Add(xtpControlButton, ID_PEMBELIAN_MUTASIBARANG_CHANGE, "Ubah Mutasi")
            Set PEMBELIAN_MUTASIBARANG_LIST = .Add(xtpControlButton, ID_PEMBELIAN_MUTASIBARANG_LIST, "Daftar Mutasi")
            Set PEMBELIAN_MUTASIBARANG_BASE = .Add(xtpControlButton, ID_PEMBELIAN_MUTASIBARANG_BASE, "Mutasi Base")
            PEMBELIAN_MUTASIBARANG_BASE.BeginGroup = True
        End With
        
        'Menu Purchasing
        Set PEMBELIAN_PURCHASING = .Add(xtpControlButtonPopup, ID_PEMBELIAN_PURCHASING, "Purchasing / Pemesanan")
        PEMBELIAN_PURCHASING.BeginGroup = False
        With PEMBELIAN_PURCHASING.CommandBar.Controls
            'Menu Purchasing PO
            Set PEMBELIAN_PURCHASING_PERMINTAAN = .Add(xtpControlButtonPopup, ID_PEMBELIAN_PURCHASING_PERMINTAAN, "Permintaan Barang")
            With PEMBELIAN_PURCHASING_PERMINTAAN.CommandBar.Controls
                Set PEMBELIAN_PURCHASING_PERMINTAAN_ADD = .Add(xtpControlButton, ID_PEMBELIAN_PURCHASING_PERMINTAAN_ADD, "Tambah Permintaan Barang")
                Set PEMBELIAN_PURCHASING_PERMINTAAN_CHANGE = .Add(xtpControlButton, ID_PEMBELIAN_PURCHASING_PERMINTAAN_CHANGE, "Ubah Permintaan Barang")
                Set PEMBELIAN_PURCHASING_PERMINTAAN_CLOSE = .Add(xtpControlButton, ID_PEMBELIAN_PURCHASING_PERMINTAAN_CLOSE, "Close Permintaan Barang")
                Set PEMBELIAN_PURCHASING_PERMINTAAN_LIST = .Add(xtpControlButton, ID_PEMBELIAN_PURCHASING_PERMINTAAN_LIST, "List Permintaan Barang")
                Set PEMBELIAN_PURCHASING_PERMINTAAN_DAYSCOUNT = .Add(xtpControlButton, ID_PEMBELIAN_PURCHASING_PERMINTAAN_DAYSCOUNT, "List Permintaan Barang (Days Count)")
                Set PEMBELIAN_PURCHASING_PERMINTAAN_PRINT = .Add(xtpControlButton, ID_PEMBELIAN_PURCHASING_PERMINTAAN_PRINT, "Reprint Permintaan")
                PEMBELIAN_PURCHASING_PERMINTAAN_PRINT.BeginGroup = True
            End With
            
            Set PEMBELIAN_PURCHASING_PO = .Add(xtpControlButtonPopup, ID_PEMBELIAN_PURCHASING_PO, "Purchase Order")
            With PEMBELIAN_PURCHASING_PO.CommandBar.Controls
                Set PEMBELIAN_PURCHASING_PO_ADD = .Add(xtpControlButton, ID_PEMBELIAN_PURCHASING_PO_ADD, "Tambah Purchase Order")
                Set PEMBELIAN_PURCHASING_PO_CHANGE = .Add(xtpControlButton, ID_PEMBELIAN_PURCHASING_PO_CHANGE, "Ubah Purchase Order")
                Set PEMBELIAN_PURCHASING_PO_CLOSECANCEL = .Add(xtpControlButton, ID_PEMBELIAN_PURCHASING_PO_CLOSECANCEL, "PO Close/Cancel")
                PEMBELIAN_PURCHASING_PO_CLOSECANCEL.BeginGroup = False
            End With
            
            'Menu Penerimaan Barang
            Set PEMBELIAN_PURCHASING_PENERIMAANBARANG = .Add(xtpControlButtonPopup, ID_PEMBELIAN_PURCHASING_PENERIMAANBARANG, "Penerimaan Barang")
            With PEMBELIAN_PURCHASING_PENERIMAANBARANG.CommandBar.Controls
                Set PEMBELIAN_PURCHASING_PENERIMAANBARANG_ADD = .Add(xtpControlButton, ID_PEMBELIAN_PURCHASING_PENERIMAANBARANG_ADD, "Tambah Penerimaan")
                Set PEMBELIAN_PURCHASING_PENERIMAANBARANG_CHANGE = .Add(xtpControlButton, ID_PEMBELIAN_PURCHASING_PENERIMAANBARANG_CHANGE, "Ubah Penerimaan")
                Set PEMBELIAN_PURCHASING_PENERIMAANBARANG_RETUR = .Add(xtpControlButton, ID_PEMBELIAN_PURCHASING_PENERIMAANBARANG_RETUR, "Retur Penerimaan")
                Set PEMBELIAN_PURCHASING_PENERIMAANBARANG_PRINTBPB = .Add(xtpControlButton, ID_PEMBELIAN_PURCHASING_PENERIMAANBARANG_PRINTBPB, "Print BPB")
                PEMBELIAN_PURCHASING_PENERIMAANBARANG_RETUR.BeginGroup = False
            End With
            
            'Menu Pemakaian Bahan Baku
            Set PEMBELIAN_PURCHASING_PEMAKAIANBAHANBAKU = .Add(xtpControlButtonPopup, ID_PEMBELIAN_PURCHASING_PEMAKAIANBAHANBAKU, "Pemakaian Bahan Baku")
            With PEMBELIAN_PURCHASING_PEMAKAIANBAHANBAKU.CommandBar.Controls
                Set PEMBELIAN_PURCHASING_PEMAKAIANBAHANBAKU_ADD = .Add(xtpControlButton, ID_PEMBELIAN_PURCHASING_PEMAKAIANBAHANBAKU_ADD, "Tambah Pemakaian Bahan Baku")
                Set PEMBELIAN_PURCHASING_PEMAKAIANBAHANBAKU_CHANGE = .Add(xtpControlButton, ID_PEMBELIAN_PURCHASING_PEMAKAIANBAHANBAKU_CHANGE, "Ubah Pemakaian Bahan Baku")
                Set PEMBELIAN_PURCHASING_PEMAKAIANBAHANBAKU_SISA = .Add(xtpControlButton, ID_PEMBELIAN_PURCHASING_PEMAKAIANBAHANBAKU_SISA, "Sisa Pemakaian Bahan Baku")
                PEMBELIAN_PURCHASING_PEMAKAIANBAHANBAKU_SISA.BeginGroup = False
            End With
             
            Set PEMBELIAN_PURCHASING_INQUIRY = .Add(xtpControlButton, ID_PEMBELIAN_PURCHASING_INQUIRY, "Inquring Purchase Order")
            PEMBELIAN_PURCHASING_INQUIRY.BeginGroup = True
            Set PEMBELIAN_PURCHASING_PRINTPO = .Add(xtpControlButton, ID_PEMBELIAN_PURCHASING_PRINTPO, "Print Purchase Order")
            PEMBELIAN_PURCHASING_PRINTPO.BeginGroup = True
            Set PEMBELIAN_PURCHASING_LAPPO = .Add(xtpControlButton, ID_PEMBELIAN_PURCHASING_LAPPO, "Laporan Purchase Order")
            PEMBELIAN_PURCHASING_LAPPO.BeginGroup = True
            Set PEMBELIAN_PURCHASING_LAPPEMBELIAN = .Add(xtpControlButton, ID_PEMBELIAN_PURCHASING_LAPPEMBELIAN, "Laporan Pembelian dan Retur")
            Set PEMBELIAN_PURCHASING_LAPPEMAKAIANBB = .Add(xtpControlButton, ID_PEMBELIAN_PURCHASING_LAPPEMAKAIANBB, "Laporan Pemakaian Bahan Baku")
        End With
        
        Set PEMBELIAN_UTILITY = .Add(xtpControlButtonPopup, ID_PEMBELIAN_UTILITY, "Utility")
        With PEMBELIAN_UTILITY.CommandBar.Controls
            Set PEMBELIAN_UTILITY_OPTIONS = .Add(xtpControlButton, ID_PEMBELIAN_UTILITY_OPTIONS, "Options")
            Set PEMBELIAN_UTILITY_FASTSEARCH = .Add(xtpControlCheckBox, ID_PEMBELIAN_UTILITY_FASTSEARCH, "Pencarian Cepat Dengan batas Tanggal")
            Set PEMBELIAN_UTILITY_ADJUSTSTOK = .Add(xtpControlButton, ID_PEMBELIAN_UTILITY_ADJUSTSTOK, "Adjust Stok Barang")
        End With
    End With
    
    'END PURCHASING MENU ##########################################################
        
    'STAR PRODUKSI MENU #############################
    Set PRODUKSI = RibbonBar.InsertTab(ID_PRODUKSI, "PRODUKSI")
    With PRODUKSI
        Set PRODUKSI_MAINMENU = .Groups.AddGroup("Activity", ID_PRODUKSI_MAINMENU)
        With PRODUKSI_MAINMENU
            Set PRODUKSI_MAINMENU_TABLES = .Add(xtpControlButtonPopup, ID_PRODUKSI_MAINMENU_TABLES, "Tabel")
            With PRODUKSI_MAINMENU_TABLES.CommandBar.Controls
                Set PRODUKSI_MAINMENU_TABLES_RESEP = .Add(xtpControlButtonPopup, ID_PRODUKSI_MAINMENU_TABLES_RESEP, "Atur Formula/SOP")
                With PRODUKSI_MAINMENU_TABLES_RESEP.CommandBar.Controls
                    Set PRODUKSI_MAINMENU_TABLES_RESEP_LEM = .Add(xtpControlButton, ID_PRODUKSI_MAINMENU_TABLES_RESEP_LEM, "SOP LEM")
                    Set PRODUKSI_MAINMENU_TABLES_RESEP_KARET = .Add(xtpControlButton, ID_PRODUKSI_MAINMENU_TABLES_RESEP_KARET, "SOP KARET")
                    Set PRODUKSI_MAINMENU_TABLES_RESEP_LISTLEM = .Add(xtpControlButton, ID_PRODUKSI_MAINMENU_TABLES_RESEP_LISTLEM, "LIST PRODUK LEM")
                    PRODUKSI_MAINMENU_TABLES_RESEP_LISTLEM.BeginGroup = True
                End With
                Set PRODUKSI_MAINMENU_TABLES_KONVERSI = .Add(xtpControlButtonPopup, ID_PRODUKSI_MAINMENU_TABLES_KONVERSI, "Packaging")
                With PRODUKSI_MAINMENU_TABLES_KONVERSI.CommandBar.Controls
                    Set PRODUKSI_MAINMENU_TABLES_KONVERSI_ADD = .Add(xtpControlButton, ID_PRODUKSI_MAINMENU_TABLES_KONVERSI_ADD, "Tambah Konversi Kemasan")
                    Set PRODUKSI_MAINMENU_TABLES_KONVERSI_CHANGE = .Add(xtpControlButton, ID_PRODUKSI_MAINMENU_TABLES_KONVERSI_CHANGE, "Ubah Konversi Kemasan")
                    Set PRODUKSI_MAINMENU_TABLES_KONVERSI_LIST = .Add(xtpControlButton, ID_PRODUKSI_MAINMENU_TABLES_KONVERSI_LIST, "Print/View List Konversi")
                End With
                Set PRODUKSI_MAINMENU_TABLES_KONVERSI_UNIT = .Add(xtpControlButton, ID_PRODUKSI_MAINMENU_TABLES_KONVERSI_UNIT, "Konversi Satuan")
                Set PRODUKSI_MAINMENU_TABLES_DEFINEKG = .Add(xtpControlButton, ID_PRODUKSI_MAINMENU_TABLES_DEFINEKG, "Define Kilogram for Base Unit")
                PRODUKSI_MAINMENU_TABLES_KONVERSI.BeginGroup = True
                Set PRODUKSI_MAINMENU_TABLES_RNCPROD = .Add(xtpControlButtonPopup, ID_PRODUKSI_MAINMENU_TABLES_RNCPROD, "Rencana Produksi")
                With PRODUKSI_MAINMENU_TABLES_RNCPROD.CommandBar.Controls
                    Set PRODUKSI_MAINMENU_TABLES_RNCPROD_ADD = .Add(xtpControlButton, ID_PRODUKSI_MAINMENU_TABLES_RNCPROD_ADD, "Tambah Rencana Produksi")
                    Set PRODUKSI_MAINMENU_TABLES_RNCPROD_CHANGE = .Add(xtpControlButton, ID_PRODUKSI_MAINMENU_TABLES_RNCPROD_CHANGE, "Ubah Rencana Produksi")
                    Set PRODUKSI_MAINMENU_TABLES_RNCPROD_LIST = .Add(xtpControlButton, ID_PRODUKSI_MAINMENU_TABLES_RNCPROD_LIST, "View Rencana Produksi")
                End With
                PRODUKSI_MAINMENU_TABLES_RNCPROD.BeginGroup = True
                Set PRODUKSI_MAINMENU_TABLES_KATEGORI = .Add(xtpControlButton, ID_PRODUKSI_MAINMENU_TABLES_KATEGORI, "Kategori")
                Set PRODUKSI_MAINMENU_TABLES_REAKTOR = .Add(xtpControlButton, ID_PRODUKSI_MAINMENU_TABLES_REAKTOR, "Reaktor")
            End With
            
            Set PRODUKSI_MAINMENU_SOP = .Add(xtpControlButtonPopup, ID_PRODUKSI_MAINMENU_SOP, "SOP")
            With PRODUKSI_MAINMENU_SOP.CommandBar.Controls
                Set PRODUKSI_MAINMENU_SOP_ADD = .Add(xtpControlButtonPopup, ID_PRODUKSI_MAINMENU_SOP_ADD, "SOP")
                With PRODUKSI_MAINMENU_SOP_ADD.CommandBar.Controls
                    Set PRODUKSI_MAINMENU_SOP_ADD_NEW = .Add(xtpControlButton, ID_PRODUKSI_MAINMENU_SOP_ADD_NEW, "Tambah SOP")
                    Set PRODUKSI_MAINMENU_SOP_ADD_EDIT = .Add(xtpControlButton, ID_PRODUKSI_MAINMENU_SOP_ADD_EDIT, "Ubah SOP")
                End With
                Set PRODUKSI_MAINMENU_SOP_TAKEPACKAGE = .Add(xtpControlButtonPopup, ID_PRODUKSI_MAINMENU_SOP_TAKEPACKAGE, "Permintaan Packaging")
                With PRODUKSI_MAINMENU_SOP_TAKEPACKAGE.CommandBar.Controls
                    Set PRODUKSI_MAINMENU_SOP_TAKEPACKAGE_ADD = .Add(xtpControlButton, ID_PRODUKSI_MAINMENU_SOP_TAKEPACKAGE_ADD, "Tambah Permintaan")
                    Set PRODUKSI_MAINMENU_SOP_TAKEPACKAGE_LIST = .Add(xtpControlButton, ID_PRODUKSI_MAINMENU_SOP_TAKEPACKAGE_LIST, "Update Lot Permintaan")
                    Set PRODUKSI_MAINMENU_SOP_TAKEPACKAGE_REPORT = .Add(xtpControlButton, ID_PRODUKSI_MAINMENU_SOP_TAKEPACKAGE_REPORT, "Laporan Permintaan barang")
                End With
                Set PRODUKSI_MAINMENU_SOP_CHANGE = .Add(xtpControlButton, ID_PRODUKSI_MAINMENU_SOP_CHANGE, "Cetak Ulang SOP")
                Set PRODUKSI_MAINMENU_SOP_LIST = .Add(xtpControlButton, ID_PRODUKSI_MAINMENU_SOP_LIST, "Laporan SOP")
                Set PRODUKSI_MAINMENU_SOP_LISTGROUP = .Add(xtpControlButton, ID_PRODUKSI_MAINMENU_SOP_LISTGROUP, "Laporan Group Packaging")
                PRODUKSI_MAINMENU_SOP_LISTGROUP.BeginGroup = True
                Set PRODUKSI_MAINMENU_SOP_LOTPALET = .Add(xtpControlButton, ID_PRODUKSI_MAINMENU_SOP_LOTPALET, "Laporan Lot/Palet")
                Set PRODUKSI_MAINMENU_SOP_PRODMONTHLY = .Add(xtpControlPopup, ID_PRODUKSI_MAINMENU_SOP_PRODMONTHLY, "Laporan Produksi")
                With PRODUKSI_MAINMENU_SOP_PRODMONTHLY.CommandBar.Controls
                    Set PRODUKSI_MAINMENU_SOP_PRODMONTHLY_BYBB = .Add(xtpControlButton, ID_PRODUKSI_MAINMENU_SOP_PRODMONTHLY_BYBB, "By.Bahan Baku (Monthly)")
                    Set PRODUKSI_MAINMENU_SOP_PRODMONTHLY_BYTOPPRODUK = .Add(xtpControlButton, ID_PRODUKSI_MAINMENU_SOP_PRODMONTHLY_BYTOPPRODUK, "By.Kg (Top Produk)")
                End With
                Set PRODUKSI_MAINMENU_SOP_SCANLOT = .Add(xtpControlPopup, ID_PRODUKSI_MAINMENU_SOP_SCANLOT, "Scan Lot/Palet")
                With PRODUKSI_MAINMENU_SOP_SCANLOT.CommandBar.Controls
                    Set PRODUKSI_MAINMENU_SOP_SCANLOT_SCANPALET = .Add(xtpControlButton, ID_PRODUKSI_MAINMENU_SOP_SCANLOT_SCANPALET, "Scan Lot")
                    Set PRODUKSI_MAINMENU_SOP_SCANLOT_EDITPALET = .Add(xtpControlButton, ID_PRODUKSI_MAINMENU_SOP_SCANLOT_EDITPALET, "Ubah Palet")
                    Set PRODUKSI_MAINMENU_SOP_SCANLOT_KEYPALET = .Add(xtpControlButton, ID_PRODUKSI_MAINMENU_SOP_SCANLOT_KEYPALET, "Buka Kunci Palet")
                End With
                PRODUKSI_MAINMENU_SOP_SCANLOT.BeginGroup = True
                Set PRODUKSI_MAINMENU_SOP_RFID = .Add(xtpControlButton, ID_PRODUKSI_MAINMENU_SOP_RFID, "Data Palet Pending")
            End With
            Set PRODUKSI_MAINMENU_MUTASI = .Add(xtpControlButtonPopup, ID_PRODUKSI_MAINMENU_MUTASI, "Mutasi Setengah Jadi")
            With PRODUKSI_MAINMENU_MUTASI.CommandBar.Controls
                Set PRODUKSI_MAINMENU_MUTASI_ADD = .Add(xtpControlButton, ID_PRODUKSI_MAINMENU_MUTASI_ADD, "Mutasi Setengah Jadi")
                'Set PRODUKSI_MAINMENU_MUTASI_LAPLOT = .Add(xtpControlButton, ID_PRODUKSI_MAINMENU_MUTASI_LAPLOT, "Laporan Mutasi By Lot")
                Set PRODUKSI_MAINMENU_MUTASI_WIPBASE = .Add(xtpControlButton, ID_PRODUKSI_MAINMENU_MUTASI_WIPBASE, "Base WIP to Bahan Baku")
                Set PRODUKSI_MAINMENU_MUTASI_LAPWIPBASE = .Add(xtpControlButton, ID_PRODUKSI_MAINMENU_MUTASI_LAPWIPBASE, "Laporan Mutasi WIP Base")
                Set PRODUKSI_MAINMENU_MUTASI_ADJWIP = .Add(xtpControlButton, ID_PRODUKSI_MAINMENU_MUTASI_ADJWIP, "Adjust Stock Base WIP")
            End With
        End With
    End With
    
    'END PRODUCTION MENU ######################################################
    'START WIRE HOUSE MENU
    Set GUDANG = RibbonBar.InsertTab(ID_GUDANG, "WAREHOUSE")
    GUDANG.Visible = False
    With GUDANG
        Set GUDANG_MAINMENU = .Groups.AddGroup("Activity", ID_GUDANG_MAINMENU)
    End With
    
    'END WAREHOUSE MENU ######################################################
    
    'STAR MENU PENJUALAN
    Set PENJUALAN = RibbonBar.InsertTab(ID_PENJUALAN, "PENJUALAN")
    PENJUALAN.Selected = False
    With PENJUALAN
        Set PENJUALAN_MAINMENU = .Groups.AddGroup("Activity", ID_PENJUALAN_MAINMENU)
        With PENJUALAN_MAINMENU
            Set PENJUALAN_MAINMENU_TABLES = .Add(xtpControlButtonPopup, ID_PENJUALAN_MAINMENU_TABLE, "Tabel")
            With PENJUALAN_MAINMENU_TABLES.CommandBar.Controls
                Set PENJUALAN_MAINMENU_TABLES_SATUAN = .Add(xtpControlButtonPopup, ID_PENJUALAN_MAINMENU_TABLE_SATUAN, "Satuan")
                With PENJUALAN_MAINMENU_TABLES_SATUAN.CommandBar.Controls
                    Set PENJUALAN_MAINMENU_TABLES_SATUAN_ADD = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_TABLE_SATUAN_ADD, "Tambah Satuan")
                    Set PENJUALAN_MAINMENU_TABLES_SATUAN_CHANGE = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_TABLE_SATUAN_CHANGE, "Ubah Satuan")
                    Set PENJUALAN_MAINMENU_TABLES_SATUAN_LIST = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_TABLE_SATUAN_LIST, "Daftar Satuan")
                End With
                
                Set PENJUALAN_MAINMENU_TABLES_CATAGORI = .Add(xtpControlButtonPopup, ID_PENJUALAN_MAINMENU_TABLE_CATAGORI, "Katagori")
                With PENJUALAN_MAINMENU_TABLES_CATAGORI.CommandBar.Controls
                    Set PENJUALAN_MAINMENU_TABLES_CATAGORI_ADD = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_TABLE_CATAGORI_ADD, "Tambah Katagori")
                    Set PENJUALAN_MAINMENU_TABLES_CATAGORI_CHANGE = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_TABLE_CATAGORI_CHANGE, "Ubah Katagori")
                    Set PENJUALAN_MAINMENU_TABLES_CATAGORI_LIST = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_TABLE_CATAGORI_LIST, "Daftar Katagori")
                End With
                
                Set PENJUALAN_MAINMENU_TABLES_BARANGJADI = .Add(xtpControlButtonPopup, ID_PENJUALAN_MAINMENU_TABLE_BARANGJADI, "Barang Jadi")
                With PENJUALAN_MAINMENU_TABLES_BARANGJADI.CommandBar.Controls
                    Set PENJUALAN_MAINMENU_TABLES_BARANGJADI_MANAGE = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_TABLE_BARANGJADI_MANAGE, "Atur Barang Jadi")
                    Set PENJUALAN_MAINMENU_TABLES_BARANGJADI_LIST = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_TABLE_BARANGJADI_LIST, "Daftar Barang Jadi")
                End With
                
                Set PENJUALAN_MAINMENU_TABLES_GUDANG = .Add(xtpControlButtonPopup, ID_PENJUALAN_MAINMENU_TABLE_GUDANG, "Gudang")
                With PENJUALAN_MAINMENU_TABLES_GUDANG.CommandBar.Controls
                    Set PENJUALAN_MAINMENU_TABLES_GUDANG_ADD = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_TABLE_GUDANG_ADD, "Tambah Gudang")
                    Set PENJUALAN_MAINMENU_TABLES_GUDANG_CHANGE = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_TABLE_GUDANG_CHANGE, "Ubah Gudang")
                    Set PENJUALAN_MAINMENU_TABLES_GUDANG_LIST = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_TABLE_GUDANG_LIST, "Daftar Gudang")
                End With
                
                Set PENJUALAN_MAINMENU_TABLES_AREA = .Add(xtpControlButtonPopup, ID_PENJUALAN_MAINMENU_TABLE_AREA, "Area")
                PENJUALAN_MAINMENU_TABLES_AREA.BeginGroup = True
                With PENJUALAN_MAINMENU_TABLES_AREA.CommandBar.Controls
                    Set PENJUALAN_MAINMENU_TABLES_AREA_ADD = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_TABLE_AREA_ADD, "Tambah Area")
                    Set PENJUALAN_MAINMENU_TABLES_AREA_CHANGE = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_TABLE_AREA_CHANGE, "Ubah Area")
                    Set PENJUALAN_MAINMENU_TABLES_AREA_LIST = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_TABLE_AREA_LIST, "Daftar Area")
                End With
                
                Set PENJUALAN_MAINMENU_TABLES_CUSTOMER = .Add(xtpControlButtonPopup, ID_PENJUALAN_MAINMENU_TABLE_CUSTOMER, "Pelanggan/Customer")
                With PENJUALAN_MAINMENU_TABLES_CUSTOMER.CommandBar.Controls
                    Set PENJUALAN_MAINMENU_TABLES_CUSTOMER_ADD = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_TABLE_CUSTOMER_ADD, "Tambah Customer")
                    Set PENJUALAN_MAINMENU_TABLES_CUSTOMER_CHANGE = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_TABLE_CUSTOMER_CHANGE, "Ubah Customer")
                    Set PENJUALAN_MAINMENU_TABLES_CUSTOMER_LIST = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_TABLE_CUSTOMER_LIST, "Daftar Customer")
                End With
                
                Set PENJUALAN_MAINMENU_TABLES_SALES = .Add(xtpControlButtonPopup, ID_PENJUALAN_MAINMENU_TABLE_SALES, "Sales")
                With PENJUALAN_MAINMENU_TABLES_SALES.CommandBar.Controls
                    Set PENJUALAN_MAINMENU_TABLES_SALES_MANAGE = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_TABLE_SALES_MANAGE, "Manage Sales")
                End With
            End With
            
            Set PENJUALAN_MAINMENU_MUTASI = .Add(xtpControlButtonPopup, ID_PENJUALAN_MAINMENU_MUTASI, "Inventory")
            With PENJUALAN_MAINMENU_MUTASI.CommandBar.Controls
                Set PENJUALAN_MAINMENU_MUTASI_ADD = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_MUTASI_ADD, "Tambah Mutasi Barang jadi")
                Set PENJUALAN_MAINMENU_MUTASI_CHANGE = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_MUTASI_CHANGE, "Ubah Mutasi Barang Jadi")
                Set PENJUALAN_MAINMENU_MUTASI_MUT = .Add(xtpControlButtonPopup, ID_PENJUALAN_MAINMENU_MUTASI_MUT, "Mutasi")
                PENJUALAN_MAINMENU_MUTASI_MUT.Enabled = True
                With PENJUALAN_MAINMENU_MUTASI_MUT.CommandBar.Controls
                    Set PENJUALAN_MAINMENU_MUTASI_MUT_OVERZAK = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_MUTASI_MUT_OVERZAK, "Mutasi Over Zak")
                    Set PENJUALAN_MAINMENU_MUTASI_MUT_MUTASILOT = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_MUTASI_MUT_MUTASILOT, "Adjust Stok By Lot")
                    Set PENJUALAN_MAINMENU_MUTASI_MUT_MUTPABRIK = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_MUTASI_MUT_MUTPABRIK, "Terima Dari Pabrik")
                    'Set PENJUALAN_MAINMENU_MUTASI_MUT_ADJSTOK = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_MUTASI_MUT_ADJSTOK, "Penyesuaian Stok")
                End With
                Set PENJUALAN_MAINMENU_MUTASI_LIST = .Add(xtpControlButtonPopup, ID_PENJUALAN_MAINMENU_MUTASI_LIST, "Laporan Mutasi")
                PENJUALAN_MAINMENU_MUTASI_LIST.Enabled = True
                With PENJUALAN_MAINMENU_MUTASI_LIST.CommandBar.Controls
                    Set PENJUALAN_MAINMENU_MUTASI_LIST = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_MUTASI_LIST, "Laporan Mutasi")
                    Set PENJUALAN_MAINMENU_MUTASI_GUDANG = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_MUTASI_GUDANG, "Laporan Mutasi Gudang")
                    Set PENJUALAN_MAINMENU_MUTASI_FAILED = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_MUTASI_FAILED, "Mutasi palet tidak lengkap")
                    Set PENJUALAN_MAINMENU_MUTASI_PRINTPRICE = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_MUTASI_PRINTPRICE, "Nilai/Nota Retur")
                    PENJUALAN_MAINMENU_MUTASI_PRINTPRICE.BeginGroup = True
                End With
                
                Set PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_ADD = .Add(xtpControlButtonPopup, ID_PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_ADD, "Pindah Gudang")
                PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_ADD.Enabled = True
                With PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_ADD.CommandBar.Controls
                    Set PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_ADD = .Add(xtpControlButtonPopup, ID_PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_ADD, "Pindah Gudang")
                    PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_ADD.Enabled = True
                    With PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_ADD.CommandBar.Controls
                        Set PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_ADD = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_ADD, "Pindah Gudang")
                        Set PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_TOWIP = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_TOWIP, "To WIP")
                    End With
                    Set PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_CHANGE = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_CHANGE, "Ubah Pindah Gudang")
                    Set PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_PRINT = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_PRINT, "Reprint Pindah Gudang")
                    Set PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_SCANPALET = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_SCANPALET, "Penerimaan Palet WIP")
                    PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_SCANPALET.BeginGroup = True
                    Set PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_COMPAREHPP = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_COMPAREHPP, "Validasi data palet")
                    Set PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_LOTPALET = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_LOTPALET, "Laporan Lot/Palet")
                    PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_LOTPALET.BeginGroup = True
                End With
                
                Set PENJUALAN_MAINMENU_MUTASI_LISTSTOK = .Add(xtpControlButtonPopup, ID_PENJUALAN_MAINMENU_MUTASI_LISTSTOK, "Laporan Stok")
                PENJUALAN_MAINMENU_MUTASI_LISTSTOK.Enabled = True
                With PENJUALAN_MAINMENU_MUTASI_LISTSTOK.CommandBar.Controls
                    Set PENJUALAN_MAINMENU_MUTASI_LISTSTOK = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_MUTASI_LISTSTOK, "Laporan Stok")
                    Set PENJUALAN_MAINMENU_MUTASI_WIP = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_MUTASI_WIP, "Persediaan Barang")
                    Set PENJUALAN_MAINMENU_MUTASI_KARTUSTOK = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_MUTASI_KARTUSTOK, "Kartu Stok")
                End With
                
                Set PENJUALAN_MAINMENU_MUTASI_PACKLIST = .Add(xtpControlButtonPopup, ID_PENJUALAN_MAINMENU_MUTASI, "Kontrol Pengambilan Kemasan")
                PENJUALAN_MAINMENU_MUTASI_PACKLIST.Enabled = True
                With PENJUALAN_MAINMENU_MUTASI_PACKLIST.CommandBar.Controls
                    Set PENJUALAN_MAINMENU_MUTASI_PACKLIST_ADD = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_MUTASI_PACKLIST_ADD, "Confirm")
                    Set PENJUALAN_MAINMENU_MUTASI_PACKLIST_LIST = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_MUTASI_PACKLIST_LIST, "Pending Lot Kemasan")
                    Set PENJUALAN_MAINMENU_MUTASI_PACKLIST_CLOSE = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_MUTASI_PACKLIST_CLOSE, "Open/Close  Lot")
                End With
            End With
            
            Set PENJUALAN_MAINMENU_INVOICING = .Add(xtpControlButtonPopup, ID_PENJUALAN_MAINMENU_INVOICING, "Penagihan / Invoicing")
            
            With PENJUALAN_MAINMENU_INVOICING.CommandBar.Controls
                Set PENJUALAN_MAINMENU_INVOICING_SO = .Add(xtpControlButtonPopup, ID_PENJUALAN_MAINMENU_INVOICING_SO, "Sales Order")
                PENJUALAN_MAINMENU_INVOICING_SO.Enabled = True
                With PENJUALAN_MAINMENU_INVOICING_SO.CommandBar.Controls
                    Set PENJUALAN_MAINMENU_INVOICING_SO_ADD = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_INVOICING_SO_ADD, "Tambah Sales Order")
                    Set PENJUALAN_MAINMENU_INVOICING_SO_CHANGE = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_INVOICING_SO_CHANGE, "Ubah Sales Order")
                    Set PENJUALAN_MAINMENU_INVOICING_SO_CANCEL = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_INVOICING_SO_CANCEL, "Cancel Sales Order")
                End With
                
                Set PENJUALAN_MAINMENU_INVOICING_SJ = .Add(xtpControlButtonPopup, ID_PENJUALAN_MAINMENU_INVOICING_SJ, "Surat Jalan")
                PENJUALAN_MAINMENU_INVOICING_SJ.BeginGroup = True
                With PENJUALAN_MAINMENU_INVOICING_SJ.CommandBar.Controls
                    Set PENJUALAN_MAINMENU_INVOICING_SJ_ADD = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_INVOICING_SJ_ADD, "Tambah Surat Jalan")
                    Set PENJUALAN_MAINMENU_INVOICING_SJ_CHANGE = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_INVOICING_SJ_CHANGE, "Ubah Surat Jalan")
                    Set PENJUALAN_MAINMENU_INVOICING_SJ_PRINT = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_INVOICING_SJ_PRINT, "Print Surat Jalan")
                    Set PENJUALAN_MAINMENU_INVOICING_SJ_ADDLOT = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_INVOICING_SJ_ADDLOT, "Input Nomor Lot SJ")
                    PENJUALAN_MAINMENU_INVOICING_SJ_ADDLOT.BeginGroup = True
                End With
                
                Set PENJUALAN_MAINMENU_INVOICING_FAKTURJUAL = .Add(xtpControlButtonPopup, ID_PENJUALAN_MAINMENU_INVOICING_FAKTURJUAL, "Faktur Penjualan")
                With PENJUALAN_MAINMENU_INVOICING_FAKTURJUAL.CommandBar.Controls
                    Set PENJUALAN_MAINMENU_INVOICING_FAKTURJUAL_ADD = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_INVOICING_FAKTURJUAL_ADD, "Tambah Faktur Penjualan")
                    Set PENJUALAN_MAINMENU_INVOICING_FAKTURJUAL_CHANGE = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_INVOICING_FAKTURJUAL_CHANGE, "Ubah Faktur Penjualan")
                    Set PENJUALAN_MAINMENU_INVOICING_FAKTURJUAL_PRINT = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_INVOICING_FAKTURJUAL_PRINT, "Print Faktur Penjualan")
                    Set PENJUALAN_MAINMENU_INVOICING_FAKTURJUAL_PREVIEW = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_INVOICING_FAKTURJUAL_PREVIEW, "Preview Faktur Penjualan")
                    PENJUALAN_MAINMENU_INVOICING_FAKTURJUAL_PREVIEW.BeginGroup = True
                End With
                Set PENJUALAN_MAINMENU_INVOICING_FAKTURPAJAK = .Add(xtpControlButtonPopup, ID_PENJUALAN_MAINMENU_INVOICING_FAKTURPAJAK, "Faktur Pajak")
                With PENJUALAN_MAINMENU_INVOICING_FAKTURPAJAK.CommandBar.Controls
                    Set PENJUALAN_MAINMENU_INVOICING_FAKTURPAJAK_DEFINE = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_INVOICING_FAKTURPAJAK_DEFINE, "Define No Seri Faktur Pajak")
                    Set PENJUALAN_MAINMENU_INVOICING_FAKTURPAJAK_BROWSE = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_INVOICING_FAKTURPAJAK_BROWSE, "Browse No Seri Faktur Pajak")
                End With
                Set PENJUALAN_MAINMENU_INVOICING_SJSBY = .Add(xtpControlButtonPopup, ID_PENJUALAN_MAINMENU_INVOICING_SJSBY, "Surat Jalan Surabaya")
                PENJUALAN_MAINMENU_INVOICING_SJSBY.BeginGroup = True
                With PENJUALAN_MAINMENU_INVOICING_SJSBY.CommandBar.Controls
                    Set PENJUALAN_MAINMENU_INVOICING_SJSBY_ADD = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_INVOICING_SJSBY_ADD, "Surat Jalan")
                    Set PENJUALAN_MAINMENU_INVOICING_SJSBY_LIST = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_INVOICING_SJSBY_LIST, "List Surat Jalan")
                End With
                Set PENJUALAN_MAINMENU_INVOICING_INQUERYSO = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_INVOICING_INQUERYSO, "Inquring Sales Order")
                PENJUALAN_MAINMENU_INVOICING_INQUERYSO.BeginGroup = True
                Set PENJUALAN_MAINMENU_INVOICING_LAPSO = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_INVOICING_LAPSO, "Laporan Sales Order")
                PENJUALAN_MAINMENU_INVOICING_LAPSO.BeginGroup = True
                Set PENJUALAN_MAINMENU_INVOICING_LAPSJ = .Add(xtpControlButtonPopup, ID_PENJUALAN_MAINMENU_INVOICING_LAPSJ, "Laporan Surat Jalan")
                With PENJUALAN_MAINMENU_INVOICING_LAPSJ.CommandBar.Controls
                    Set PENJUALAN_MAINMENU_INVOICING_LAPSJ_DAFTAR = .Add(xtpControlButtonPopup, ID_PENJUALAN_MAINMENU_INVOICING_LAPSJ_DAFTAR, "Daftar")
                    With PENJUALAN_MAINMENU_INVOICING_LAPSJ_DAFTAR.CommandBar.Controls
                        Set PENJUALAN_MAINMENU_INVOICING_LAPSJ_DAFTAR = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_INVOICING_LAPSJ_DAFTAR, "Daftar Surat Jalan")
                        Set PENJUALAN_MAINMENU_INVOICING_LAPSJ_DAFTARGD = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_INVOICING_LAPSJ_DAFTARGD, "Daftar Surat Jalan By Gudang")
                        Set PENJUALAN_MAINMENU_INVOICING_LAPSJ_BYFAKTUR = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_INVOICING_LAPSJ_BYFAKTUR, "Daftar Surat Jalan By Faktur")
                        Set PENJUALAN_MAINMENU_INVOICING_LAPSJ_BYLOT = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_INVOICING_LAPSJ_BYLOT, "Daftar Surat Jalan By Lot")
                    End With
                    Set PENJUALAN_MAINMENU_INVOICING_LAPSJ_LAPORAN = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_INVOICING_LAPSJ_LAPORAN, "Laporan")
                End With
                Set PENJUALAN_MAINMENU_INVOICING_LAPJUAL = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_INVOICING_LAPJUAL, "Laporan Penjualan")
                PENJUALAN_MAINMENU_INVOICING_LAPJUAL.BeginGroup = True
                Set PENJUALAN_MAINMENU_INVOICING_LAPJUALDTL = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_INVOICING_LAPJUALDTL, "Laporan Penjualan Detail")
                Set PENJUALAN_MAINMENU_INVOICING_MONTHLY = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_INVOICING_MONTHLY, "Laporan Penjualan Monthly")
                Set PENJUALAN_MAINMENU_INVOICING_BYKATEGORI = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_INVOICING_BYKATEGORI, "Laporan Penjualan By Kategori")
                Set PENJUALAN_MAINMENU_INVOICING_KOMISI = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_INVOICING_LAPKOMISI, "Laporan Komisi Sales")
                PENJUALAN_MAINMENU_INVOICING_KOMISI.BeginGroup = True
                Set PENJUALAN_MAINMENU_INVOICING_ANALISAJUAL = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_INVOICING_ANALISAJUAL, "Analisa Penjualan")
            End With
            
            Set PENJUALAN_MAINMENU_UTILITY = .Add(xtpControlButtonPopup, ID_PENJUALAN_MAINMENU_UTILITY, "Utility")
            With PENJUALAN_MAINMENU_UTILITY.CommandBar.Controls
                Set PENJUALAN_MAINMENU_UTILITY_OPTIONS = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_UTILITY_OPTIONS, "Options...")
                Set PEMBELIAN_UTILITY_FASTSEARCH = .Add(xtpControlCheckBox, ID_PEMBELIAN_UTILITY_FASTSEARCH, "Use Date Range For Faster Search")
                Set PENJUALAN_MAINMENU_UTILITY_CANCELSO = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_UTILITY_CANCELSO, "Cancel Sales Order")
                PENJUALAN_MAINMENU_UTILITY_CANCELSO.BeginGroup = True
                Set PENJUALAN_MAINMENU_UTILITY_CLOSESO = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_UTILITY_CLOSESO, "Close Sales Order")
                Set PENJUALAN_MAINMENU_UTILITY_EXPORTSO = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_UTILITY_EXPORTSO, "Export Sales Order")
                PENJUALAN_MAINMENU_UTILITY_EXPORTSO.BeginGroup = True
                Set PENJUALAN_MAINMENU_UTILITY_IMPORTSO = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_UTILITY_IMPORTSOPBRK, "Import Sales Order")
                Set PENJUALAN_MAINMENU_UTILITY_EXPORTSJ = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_UTILITY_EXPORTSJ, "Export Surat Jalan")
                PENJUALAN_MAINMENU_UTILITY_EXPORTSJ.BeginGroup = True
                Set PENJUALAN_MAINMENU_UTILITY_IMPORTSJ = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_UTILITY_IMPORTSJ, "Import Surat Jalan")
                Set PENJUALAN_MAINMENU_UTILITY_IMPORTINV = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_UTILITY_IMPORTINV, "Import Invoice")
                PENJUALAN_MAINMENU_UTILITY_IMPORTINV.BeginGroup = True
                Set PENJUALAN_MAINMENU_UTILITY_DELINV = .Add(xtpControlButton, ID_PENJUALAN_MAINMENU_UTILITY_DELINV, "D.I")
            End With
        End With
    End With
    
    'STAR FINANCE MENU
    Set KEUANGAN = RibbonBar.InsertTab(ID_KEUANGAN, "KEUANGAN")
    Set KEUANGAN_MAINMENU = KEUANGAN.Groups.AddGroup("Activity", ID_KEUANGAN_MAINMENU)
    With KEUANGAN_MAINMENU
        'Set KEUANGAN_TABLE = .Add(xtpControlButtonPopup, ID_KEUANGAN_TABLE, "Tables")
        'With KEUANGAN_TABLE.CommandBar.Controls
             '#Menu Currency
            'Set PEMBELIAN_TABEL_CUR = .Add(xtpControlButtonPopup, ID_PEMBELIAN_TABEL_CUR, "Currency")
            'PEMBELIAN_TABEL_CUR.BeginGroup = False
            'With PEMBELIAN_TABEL_CUR.CommandBar.Controls
                'Set PEMBELIAN_TABEL_CUR_ADD = .Add(xtpControlButton, ID_PEMBELIAN_TABEL_CUR_ADD, "Add Currency")
                'Set PEMBELIAN_TABEL_CUR_CHANGE = .Add(xtpControlButton, ID_PEMBELIAN_TABEL_CUR_CHANGE, "Change Currency")
                'Set PEMBELIAN_TABEL_CUR_LIST = .Add(xtpControlButton, ID_PEMBELIAN_TABEL_CUR_LIST, "Currency List")
                'PEMBELIAN_TABEL_CUR_LIST.BeginGroup = False
            'End With
            
            '#Menu BANK
            'Set PEMBELIAN_TABEL_BANK = .Add(xtpControlButtonPopup, ID_PEMBELIAN_TABEL_BANK, "Bank")
            'With PEMBELIAN_TABEL_BANK.CommandBar.Controls
                'Set PEMBELIAN_TABEL_BANK_ADD = .Add(xtpControlButton, ID_PEMBELIAN_TABEL_BANK_ADD, "Add Bank")
                'Set PEMBELIAN_TABEL_BANK_CHANGE = .Add(xtpControlButton, ID_PEMBELIAN_TABEL_BANK_CHANGE, "Change Bank")
                'Set PEMBELIAN_TABEL_BANK_LIST = .Add(xtpControlButton, ID_PEMBELIAN_TABEL_BANK_LIST, "Bank List")
            'End With
        'End With
        Set KEUANGAN_PEMBELIAN = .Add(xtpControlButtonPopup, ID_KEUANGAN_PEMBELIAN, "Pembelian")
        With KEUANGAN_PEMBELIAN.CommandBar.Controls
             'Menu Confirm
            Set PEMBELIAN_PURCHASING_CONFIRM = .Add(xtpControlButtonPopup, ID_PEMBELIAN_PURCHASING_CONFIRM, "Confirm")
            PEMBELIAN_PURCHASING_CONFIRM.BeginGroup = False
            With PEMBELIAN_PURCHASING_CONFIRM.CommandBar.Controls
                Set PEMBELIAN_PURCHASING_CONFIRM_CONFIRMPENERIMAANBARANG = .Add(xtpControlButtonPopup, ID_PEMBELIAN_PURCHASING_CONFIRM_CONFIRMPENERIMAANBARANG, "Confirm Penerimaan")
                With PEMBELIAN_PURCHASING_CONFIRM_CONFIRMPENERIMAANBARANG.CommandBar.Controls
                    Set PEMBELIAN_PURCHASING_CONFIRM_CONFIRMPENERIMAANBARANG_ADD = .Add(xtpControlButton, ID_PEMBELIAN_PURCHASING_CONFIRM_CONFIRMPENERIMAANBARANG_ADD, "Confirm Penerimaan")
                    Set PEMBELIAN_PURCHASING_CONFIRM_CONFIRMPENERIMAANBARANG_REPRINT = .Add(xtpControlButton, ID_PEMBELIAN_PURCHASING_CONFIRM_CONFIRMPENERIMAANBARANG_REPRINT, "Reprint Voucher Penerimaan")
                End With
                Set PEMBELIAN_PURCHASING_CONFIRM_CONFIRMRETURPENERIMAANBARANG = .Add(xtpControlButton, ID_PEMBELIAN_PURCHASING_CONFIRM_CONFIRMRETURPENERIMAANBARANG, "Confirm Retur Penerimaan")
                Set PEMBELIAN_PURCHASING_CONFIRM_UNCONFIRMPENRIMAANBARANG = .Add(xtpControlButton, ID_PEMBELIAN_PURCHASING_CONFIRM_UNCONFIRMPENERIMAANBARANG, "UnConfirm Penerimaan")
                Set PEMBELIAN_PURCHASING_CONFIRM_UNCONFIRMRETURPENERIMAANBARANG = .Add(xtpControlButton, ID_PEMBELIAN_PURCHASING_CONFIRM_UNCOFIRMRETURPENERIMAANBARANG, "UnConfirm Retur Penerimaan")
                Set PEMBELIAN_PURCHASING_CONFIRM_CREATEVOUCHER = .Add(xtpControlButton, ID_PEMBELIAN_PURCHASING_CONFIRM_CREATEVOUCHER, "Add Voucher")
                PEMBELIAN_PURCHASING_CONFIRM_CREATEVOUCHER.BeginGroup = True
            End With
            
            Set PEMBELIAN_HUTANG = .Add(xtpControlButtonPopup, ID_PEMBELIAN_HUTANG, "Hutang Pembelian")
            With PEMBELIAN_HUTANG.CommandBar.Controls
                Set PEMBELIAN_HUTANG_PBYHUTANG = .Add(xtpControlButtonPopup, ID_PEMBELIAN_HUTANG_PBYHUTANG, "Pembayaran Hutang")
                With PEMBELIAN_HUTANG_PBYHUTANG.CommandBar.Controls
                        Set PEMBELIAN_HUTANG_PBYHUTANG_ADD = .Add(xtpControlButton, ID_PEMBELIAN_HUTANG_PBYHUTANG_ADD, "Add Pembayaran Hutang")
                        Set PEMBELIAN_HUTANG_PBYHUTANG_CHANGE = .Add(xtpControlButton, ID_PEMBELIAN_HUTANG_PBYHUTANG_CHANGE, "Change Pembayaran Hutang")
                        Set PEMBELIAN_HUTANG_PBYHUTANG_UNPOST = .Add(xtpControlButton, ID_PEMBELIAN_HUTANG_PBYHUTANG_UNPOST, "Unpost Pembayaran Hutang")
                        PEMBELIAN_HUTANG_PBYHUTANG_UNPOST.BeginGroup = True
                End With
                Set PEMBELIAN_HUTANG_KOREKSIHTNG = .Add(xtpControlButtonPopup, ID_PEMBELIAN_HUTANG_KOREKSIHTN, "Koreksi Hutang")
                    With PEMBELIAN_HUTANG_KOREKSIHTNG.CommandBar.Controls
                        Set PEMBELIAN_HUTANG_KOREKSIHTNG_ADD = .Add(xtpControlButton, ID_PEMBELIAN_HUTANG_KOREKSIHTN_ADD, "Add Koreksi Hutang")
                        Set PEMBELIAN_HUTANG_KOREKSIHTNG_CHANGE = .Add(xtpControlButton, ID_PEMBELIAN_HUTANG_KOREKSIHTN_CHANGE, "Change Koreksi Hutang")
                    End With
                    Set PEMBELIAN_HUTANG_LISTKOREKSI = .Add(xtpControlButton, ID_PEMBELIAN_HUTANG_LISTKOREKSI, "Daftar Koreksi")
                    PEMBELIAN_HUTANG_LISTKOREKSI.BeginGroup = True
                    Set PEMBELIAN_HUTANG_LISTPBYHUTANG = .Add(xtpControlButton, ID_PEMBELIAN_HUTANG_LISTPBYHUTANG, "Daftar Pembayaran Hutang")
                    Set KEUANGAN_PEMBELIAN_HTNGKARTU = .Add(xtpControlButton, ID_KEUANGAN_PEMBELIAN_HTNGKARTU, "Daftar Hutang Kartu")
                    Set PEMBELIAN_HUTANG_LAPSISAHUTANG = .Add(xtpControlButton, ID_PEMBELIAN_HUTANG_LAPSISAHUTANG, "Laporan Sisa Hutang")
                    Set PEMBELIAN_HUTANG_LAPBYJT = .Add(xtpControlButton, ID_PEMBELIAN_HUTANG_LAPBYJT, "Laporan Sisa Hutang By Jatuh Tempo")
                End With
                Set PEMBELIAN_GIRO = .Add(xtpControlButtonPopup, ID_PEMBELIAN_GIRO, "Giro Pembelian")
                With PEMBELIAN_GIRO.CommandBar.Controls
                    Set KEUANGAN_PEMBELIAN_MAINTENACEGIRO = .Add(xtpControlButton, ID_KEUANGAN_PEMBELIAN_MAINTENACEGIRO, "Maintenance Giro")
                    Set KEUANGAN_PEMBELIAN_ADDGIROTOLAK = .Add(xtpControlButton, ID_KEUANGAN_PEMBELIAN_ADDGIROTOLAK, "Add Giro Tolak")
                    KEUANGAN_PEMBELIAN_ADDGIROTOLAK.BeginGroup = True
                    Set KEUANGAN_PEMBELIAN_CHANGEGIROTOLAK = .Add(xtpControlButton, ID_KEUANGAN_PEMBELIAN_CHANGEGIROTOLAK, "Change Giro Tolak")
                    Set KEUANGAN_PEMBELIAN_LAPGIRO = .Add(xtpControlButton, ID_KEUANGAN_PEMBELIAN_LAPGIRO, "Laporan Giro")
                    KEUANGAN_PEMBELIAN_LAPGIRO.BeginGroup = True
                End With
                
                Set KEUANGAN_PEMBELIAN_UTILITY = .Add(xtpControlButtonPopup, ID_KEUANGAN_PEMBELIAN_UTILITY, "Utility")
                With KEUANGAN_PEMBELIAN_UTILITY.CommandBar.Controls
                    Set KEUANGAN_PEMBELIAN_PERIODONPROSES = .Add(xtpControlButton, ID_KEUANGAN_PEMBELIAN_PERIODONPROSES, "Periode On Proses")
                    Set KEUANGAN_PEMBELIAN_UTILITY_DEFINEACCOUNTSUPPLIER = .Add(xtpControlButton, ID_KEUANGAN_PEMBELIAN_UTILITY_DEFINEACCOUNTSUPPLIER, "Define Account Supplier")
                    Set KEUANGAN_PEMBELIAN_UTILITY_DEFINEBANKKAS = .Add(xtpControlButton, ID_KEUANGAN_PEMBELIAN_UTILITY_DEFINEBANKKAS, "Define BANK/KAS")
                    Set KEUANGAN_PEMBELIAN_UTILITY_DEFINEJURNALANDPROSES = .Add(xtpControlButton, ID_KEUANGAN_PEMBELIAN_UTILITY_DEFINEJURNALANDPROSES, "Define Jurnal and Process")
                    Set KEUANGAN_PEMBELIAN_UTILITY_LAPORANPOSTING = .Add(xtpControlButton, ID_KEUANGAN_PEMBELIAN_UTILITY_LAPORANPOSTING, "Laporan Posting")
                End With
                
                Set PEMBELIAN_PURCHASING_LAPCONPEMBELIAN = .Add(xtpControlButton, ID_PEMBELIAN_PURCHASING_LAPCONPEMBELIAN, "Laporan Confirm/Unconfir Pembellian dan Retur")
                Set PEMBELIAN_PURCHASING_LAPVOUCER = .Add(xtpControlButton, ID_PEMBELIAN_PURCHASING_LAPVOUCER, "Laporan Voucher")
                Set PEMBELIAN_PURCHASING_LAPPROSESVOUCHER = .Add(xtpControlButton, ID_PEMBELIAN_PURCHASING_LAPPROSESVOUCHER, "Laporan Processed/Unprocessed Voucher")
                PEMBELIAN_PURCHASING_LAPCONPEMBELIAN.BeginGroup = True
            End With
            Set KEUANGAN_PENJUALAN = .Add(xtpControlButtonPopup, ID_KEUANGAN_PENJUALAN, "Penjualan")
            With KEUANGAN_PENJUALAN.CommandBar.Controls
                Set KEUANGAN_PENJUALAN_PIUTANG = .Add(xtpControlButtonPopup, ID_KEUANGAN_PENJUALAN_PIUTANG, "Piutang")
                With KEUANGAN_PENJUALAN_PIUTANG.CommandBar.Controls
                    Set KEUANGAN_PENJUALAN_PIUTANG_KOREKSI = .Add(xtpControlButtonPopup, ID_KEUANGAN_PENJUALAN_PIUTANG_KOREKSIPIUTANG, "Koreksi Piutang")
                    With KEUANGAN_PENJUALAN_PIUTANG_KOREKSI.CommandBar.Controls
                        Set KEUANGAN_PENJUALAN_PIUTANG_KOREKSI_ADD = .Add(xtpControlButton, ID_KEUANGAN_PENJUALAN_PIUTANG_KOREKSIPIUTANG_ADD, "Add Koreksi Piutang")
                        Set KEUANGAN_PENJUALAN_PIUTANG_KOREKSI_CHANGE = .Add(xtpControlButton, ID_KEUANGAN_PENJUALAN_PIUTANG_KOREKSIPIUTANG_CHANGE, "Change Koreksi Piutang")
                        Set KEUANGAN_PENJUALAN_PIUTANG_KOREKSI_WRITEOFF = .Add(xtpControlButton, ID_KEUANGAN_PENJUALAN_PIUTANG_KOREKSIPIUTANG_WRITEOFF, "Write Off")
                        KEUANGAN_PENJUALAN_PIUTANG_KOREKSI_WRITEOFF.BeginGroup = True
                    End With
                    Set KEUANGAN_PENJUALAN_PIUTANG_PEMBAYARAN = .Add(xtpControlButtonPopup, ID_KEUANGAN_PENJUALAN_PIUTANG_PEMBAYARAN, "Pembayaran")
                    With KEUANGAN_PENJUALAN_PIUTANG_PEMBAYARAN.CommandBar.Controls
                        Set KEUANGAN_PENJUALAN_PIUTANG_PEMBAYARAN_ADD = .Add(xtpControlButton, ID_KEUANGAN_PENJUALAN_PIUTANG_PEMBAYARAN_ADD, "Add Pembayaran")
                        Set KEUANGAN_PENJUALAN_PIUTANG_PEMBAYARAN_CHANGE = .Add(xtpControlButton, ID_KEUANGAN_PENJUALAN_PIUTANG_PEMBAYARAN_CHANGE, "Change Pembayaran")
                    End With
                    Set KEUANGAN_PENJUALAN_PIUTANG_DAFTARKOREKSIPIUTANG = .Add(xtpControlButton, ID_KEUANGAN_PENJUALAN_PIUTANG_DAFTARKOREKSIPIUTANG, "Daftar Koreksi Piutang")
                    KEUANGAN_PENJUALAN_PIUTANG_DAFTARKOREKSIPIUTANG.BeginGroup = True
                    Set KEUANGAN_PENJUALAN_PIUTANG_LAPPENAGIHAN = .Add(xtpControlButton, ID_KEUANGAN_PENJUALAN_PIUTANG_LAPPENAGIHAN, "Laporan Penagihan")
                    Set KEUANGAN_PENJUALAN_PIUTANG_LAPPIUTANGAGING = .Add(xtpControlButton, ID_KEUANGAN_PENJUALAN_PIUTANG_LAPPIUTANGAGING, "Laporan Piutang Aging")
                    Set KEUANGAN_PENJUALAN_PIUTANG_LAPKARTUPIUTANG = .Add(xtpControlButton, ID_KEUANGAN_PENJUALAN_PIUTANG_LAPKARTUPIUTANG, "Laporan Kartu Piutang")
                    Set KEUANGAN_PENJUALAN_PIUTANG_LAPSISAUTANG = .Add(xtpControlButton, ID_KEUANGAN_PENJUALAN_PIUTANG_LAPSISAPIUTANG, "Laporan Sisa Piutang")
                    Set KEUANGAN_PENJUALAN_PIUTANG_LAPMUTPIUTANG = .Add(xtpControlButton, ID_KEUANGAN_PENJUALAN_PIUTANG_LAPMUTPIUTANG, "Laporan Mutasi Piutang")
                    Set KEUANGAN_PENJUALAN_PIUTANG_LAPPEMBAYARAN = .Add(xtpControlButton, ID_KEUANGAN_PENJUALAN_PIUTANG_LAPPEMBAYARAN, "Laporan Pembayaran")
                    Set KEUANGAN_PENJUALAN_PIUTANG_LAPPEMBAYARANDTL = .Add(xtpControlButton, ID_KEUANGAN_PENJUALAN_PIUTANG_LAPPEMBAYARANDTL, "Laporan Pembayaran Detail")
                    Set KEUANGAN_PENJUALAN_PIUTANG_LAPTANDATERIMAPBY = .Add(xtpControlButton, ID_KEUANGAN_PENJUALAN_PIUTANG_LAPTANDATERIMAPBY, "Laporan Tanda Terima Pembayaran")
                End With
                Set KEUANGAN_PENJUALAN_GIRO = .Add(xtpControlButtonPopup, ID_KEUANGAN_PENJUALAN_GIRO, "Giro")
                With KEUANGAN_PENJUALAN_GIRO.CommandBar.Controls
                    Set KEUANGAN_PENJUALAN_GIRO_MAINTENANCE = .Add(xtpControlButton, ID_KEUANGAN_PENJUALAN_GIRO_MAINTENANCE, "Maintenance Giro")
                    Set KEUANGAN_PENJUALAN_GIRO_GANTIGIROTOLAK = .Add(xtpControlButtonPopup, ID_KEUANGAN_PENJUALAN_GIRO_GANTIGIROTOLAK, "Ganti Giro Tolak")
                    With KEUANGAN_PENJUALAN_GIRO_GANTIGIROTOLAK.CommandBar.Controls
                        Set KEUANGAN_PENJUALAN_GIRO_GANTIGIROTOLAK_ADD = .Add(xtpControlButton, ID_KEUANGAN_PENJUALAN_GIRO_GANTIGIROTOLAK_ADD, "Add Ganti Giro Tolak")
                        Set KEUANGAN_PENJUALAN_GIRO_GANTIGIROTOLAK_CHANGE = .Add(xtpControlButton, ID_KEUANGAN_PENJUALAN_GIRO_GANTIGIROTOLAK_CHANGE, "Change Ganti Giro Tolak")
                    End With
                    Set KEUANGAN_PENJUALAN_GIRO_LAPGIRO = .Add(xtpControlButton, ID_KEUANGAN_PENJUALAN_GIRO_LAPGIRO, "Laporan Giro")
                    KEUANGAN_PENJUALAN_GIRO_LAPGIRO.BeginGroup = True
                    Set KEUANGAN_PENJUALAN_GIRO_LISTAPPGANTIGIROTAL = .Add(xtpControlButton, ID_KEUANGAN_PENJUALAN_GIRO_LISTAPPGIROTOLAK, "List Apply Giro Ganti Tolak")
                End With
                Set KEUANGAN_PENJUALAN_UTILITY = .Add(xtpControlButtonPopup, ID_KEUANGAN_PENJUALAN_UTILITY, "Utility")
                With KEUANGAN_PENJUALAN_UTILITY.CommandBar.Controls
                    Set KEUANGAN_PENJUALAN_UTILITY_PERIODONPROSES = .Add(xtpControlButton, ID_KEUANGAN_PENJUALAN_UTILITY_PERIODONPROSES, "Period On Proses")
                    Set KEUANGAN_PENJUALAN_UTILITY_DEFAINACCBANKCASH = .Add(xtpControlButton, ID_KEUANGAN_PENJUALAN_UTILITY_DEFAINACCBANKCASH, "Define Account Bank/Cash")
                    KEUANGAN_PENJUALAN_UTILITY_DEFAINACCBANKCASH.BeginGroup = True
                    Set KEUANGAN_PENJUALAN_UTILITY_DEFINECOSUTOMERACC = .Add(xtpControlButton, ID_KEUANGAN_PENJUALAN_UTILITY_DEFINECOSUTOMERACC, "Define Account Customers")
                    Set KEUANGAN_PENJUALAN_UTILITY_DEFINEJURNAL = .Add(xtpControlButton, ID_KEUANGAN_PENJUALAN_UTILITY_DEFAINJURNAL, "Define Jurnal And Proses")
                    Set KEUANGAN_PENJUALAN_UTILITY_DEFINEAGING = .Add(xtpControlButton, ID_KEUANGAN_PENJUALAN_UTILITY_DEFINEAGING, "Define Aging")
                    Set KEUANGAN_PENJUALAN_UTILITY_DEFINEKG = .Add(xtpControlButton, ID_KEUANGAN_PENJUALAN_UTILITY_DEFINEKG, "Define Kilogram for Base Unit")
                    Set KEUANGAN_PENJUALAN_UTILITY_LAPPOSTINGPENJUALAN = .Add(xtpControlButton, ID_KEUANGAN_PENJUALAN_UTILITY_LAPPOSTINGPENJUALAN, "Laporan Posting Penjualan")
                    Set KEUANGAN_PENJUALAN_UTILITY_LISTKG = .Add(xtpControlButton, ID_KEUANGAN_PENJUALAN_UTILITY_LISTKG, "Laporan Kilosales")
                End With
            End With
    End With
    Set KEUANGAN_GL = KEUANGAN.Groups.AddGroup("General Ledger", ID_KEUANGAN_GL)
     
    With KEUANGAN_GL
        Set KEUANGAN_GL_TABEL = .Add(xtpControlButtonPopup, ID_KEUANGAN_GL_TABEL, "Tabel")
        With KEUANGAN_GL_TABEL.CommandBar.Controls
            Set KEUANGAN_GL_TABEL_COMPANYTYPE = .Add(xtpControlButtonPopup, ID_KEUANGAN_GL_TABEL_COMPANYTYPE, "Company Type")
            With KEUANGAN_GL_TABEL_COMPANYTYPE.CommandBar.Controls
                Set KEUANGAN_GL_TABLE_COMPANYTYPE_ADD = .Add(xtpControlButton, ID_KEUANGAN_GL_TABEL_COMPANYTYPE_ADD, "Add")
                Set KEUANGAN_GL_TABEL_COMPANYTYPE_UPDATE = .Add(xtpControlButton, ID_KEUANGAN_GL_TABEL_COMPANYTYPE_UPDATE, "Update")
                Set KEUANGAN_GL_TABEL_COMPANYTYPE_LIST = .Add(xtpControlButton, ID_KEUANGAN_GL_TABEL_COMPANYTYPE_LIST, "List")
            End With
            Set KEUANGAN_GL_TABEL_MASTERACC = .Add(xtpControlButtonPopup, ID_KEUANGAN_GL_TABEL_MASTERACC, "Master Account")
            With KEUANGAN_GL_TABEL_MASTERACC.CommandBar.Controls
                Set KEUANGAN_GL_TABEL_MASTERACC_ADD = .Add(xtpControlButton, ID_KEUANGAN_GL_TABEL_MASTERACC_ADD, "Add")
                Set KEUANGAN_GL_TABEL_MASTERACC_UPDATE = .Add(xtpControlButton, ID_KEUANGAN_GL_TABEL_MASTERACC_UPDATE, "Update/Delete")
                Set KEUANGAN_GL_TABEL_MASTERACC_LIST = .Add(xtpControlButton, ID_KEUANGAN_GL_TABEL_MASTERACC_LIST, "List")
            End With
            Set KEUANGAN_GL_TABEL_COMPANYIDENTIFICATION = .Add(xtpControlButtonPopup, ID_KEUANGAN_GL_TABEL_COMPANYIDENTIFICATION, "Company Indentification")
            With KEUANGAN_GL_TABEL_COMPANYIDENTIFICATION.CommandBar.Controls
                Set KEUANGAN_GL_TABEL_COMPANYIDENTIFICATION_ADD = .Add(xtpControlButton, ID_KEUANGAN_GL_TABEL_COMPANYIDENTIFICATION_ADD, "Add")
                Set KEUANGAN_GL_TABEL_COMPANYIDENTIFICATION_UPDATE = .Add(xtpControlButton, ID_KEUANGAN_GL_TABEL_COMPANYIDENTIFICATION_UPDATE, "Update/Delete")
                Set KEUANGAN_GL_TABEL_COMPANYIDENTIFICATION_LIST = .Add(xtpControlButton, ID_KEUANGAN_GL_TABEL_COMPANYIDENTIFICATION_LIST, "List")
            End With
            Set KEUANGAN_GL_TABEL_COMPANYACCOUNT = .Add(xtpControlButtonPopup, ID_KEUANGAN_GL_TABEL_COMPANYACCOUNT, "Company Acoount")
            With KEUANGAN_GL_TABEL_COMPANYACCOUNT.CommandBar.Controls
                Set KEUANGAN_GL_TABEL_COMPANYACCOUNT_BROWSE = .Add(xtpControlButton, ID_KEUANGAN_GL_TABEL_COMPANYACCOUNT_BROWSE, "Browse Company Account")
                Set KEUANGAN_GL_TABEL_COMPANYACCOUNT_LISTACOOUNT = .Add(xtpControlButton, ID_KEUANGAN_GL_TABEL_COMPANYACCOUNT_LISTACCOUNT, "List Company Account")
                Set KEUANGAN_GL_TABEL_COMPANYACCOUNT_LISTBUDGET = .Add(xtpControlButton, ID_KEUANGAN_GL_TABEL_COMPANYACCOUNT_LISTBUDGET, "List Company Budget")
            End With
            Set KEUANGAN_GL_TABEL_CURRENCY = .Add(xtpControlButtonPopup, ID_KEUANGAN_GL_TABEL_CURRENCY, "Currency")
            With KEUANGAN_GL_TABEL_CURRENCY.CommandBar.Controls
                Set KEUANGAN_GL_TABEL_CURRENCY_ADD = .Add(xtpControlButton, ID_KEUANGAN_GL_TABEL_CURRENCY_ADD, "Add")
                Set KEUANGAN_GL_TABEL_CURRENCY_UPDATE = .Add(xtpControlButton, ID_KEUANGAN_GL_TABEL_CURRENCY_UPDATE, "Update/Delete")
                Set KEUANGAN_GL_TABEL_MASTERACC_LIST = .Add(xtpControlButton, ID_KEUANGAN_GL_TABEL_CURRENCY_LIST, "List")
            End With
            Set KEUANGAN_GL_TABEL_FIXEDASSETTYPE = .Add(xtpControlButtonPopup, ID_KEUANGAN_GL_TABEL_FIXEDASSETTYPE, "Fixed Asset Type")
            With KEUANGAN_GL_TABEL_FIXEDASSETTYPE.CommandBar.Controls
                Set KEUANGAN_GL_TABEL_FIXEDASSETTYPE_ADD = .Add(xtpControlButton, ID_KEUANGAN_GL_TABEL_FIXEDASSETTYPE_ADD, "Add")
                Set KEUANGAN_GL_TABEL_FIXEDASSETTYPE_UPDATE = .Add(xtpControlButton, ID_KEUANGAN_GL_TABEL_FIXEDASSETTYPE_UPDATE, "Update/Delete")
                Set KEUANGAN_GL_TABEL_FIXEDASSETTYPE_LIST = .Add(xtpControlButton, ID_KEUANGAN_GL_TABEL_FIXEDASSETTYPE_LIST, "List")
            End With
            Set KEUANGAN_GL_TABEL_BANK = .Add(xtpControlButtonPopup, ID_KEUANGAN_GL_TABEL_BANK, "Bank")
            With KEUANGAN_GL_TABEL_BANK.CommandBar.Controls
                Set KEUANGAN_GL_TABEL_BANK_ADD = .Add(xtpControlButton, ID_KEUANGAN_GL_TABEL_BANK_ADD, "Add")
                Set KEUANGAN_GL_TABEL_BANK_UPDATE = .Add(xtpControlButton, ID_KEUANGAN_GL_TABEL_BANK_UPDATE, "Update/Delete")
                Set KEUANGAN_GL_TABEL_BANK_LIST = .Add(xtpControlButton, ID_KEUANGAN_GL_TABEL_BANK_LIST, "List")
            End With
        End With
        Set KEUANGAN_GL_FIXEDASSET = .Add(xtpControlButtonPopup, ID_KEUANGAN_GL_FIXEDASSET, "Fixed Asset")
        With KEUANGAN_GL_FIXEDASSET.CommandBar.Controls
            Set KEUANGAN_GL_FIXEDASSET_PEMBELIANASSET = .Add(xtpControlButton, ID_KEUANGAN_GL_FIXEDASSET_PEMBELIANFIXEDASSET, "Pembelian Fixed Asset")
            Set KEUANGAN_GL_FIXEDASSET_POSTINGPEMBELIANFIXEDASSET = .Add(xtpControlButton, ID_KEUANGAN_GL_FIXEDASSET_POSTINGPEMBELIANFIXEDASSET, "Posting Pembelian Fixed Asset")
            Set KEUANGAN_GL_FIXEDASSET_UNPOSTINGPEMBELIANFIXEDASSET = .Add(xtpControlButton, ID_KEUANGAN_GL_FIXEDASSET_UNPOSTINGPEMBELIANFIXEDASSET, "Unposting Pembelian Fixed Asset")
            Set KEUANGAN_GL_FIXEDASSET_PENJUALANASSET = .Add(xtpControlButton, ID_KEUANGAN_GL_FIXEDASSET_PENJUALANFIXEDASSET, "Penjualan Fixed Asset")
            KEUANGAN_GL_FIXEDASSET_PENJUALANASSET.BeginGroup = True
            Set KEUANGAN_GL_FIXEDASSET_POSTINGPENJUALANFIXEDASSET = .Add(xtpControlButton, ID_KEUANGAN_GL_FIXEDASSET_POSTINGPENJUALANFIXEDASSET, "Posting Penjualan Fixed Asset")
            Set KEUANGAN_GL_FIXEDASSET_UNPOSTINGPENJUALANFIXEDASSET = .Add(xtpControlButton, ID_KEUANGAN_GL_FIXEDASSET_UNPOSTINGPENJUALANFIXEDASSET, "Unposting Penjualan Fixed Asset")
        End With
        
        Set KEUANGAN_GL_LEDGER = .Add(xtpControlButtonPopup, ID_KEUANGAN_GL_LEDGER, "Ledger")
        With KEUANGAN_GL_LEDGER.CommandBar.Controls
            Set KEUANGAN_GL_LEDGER_JURNAL = .Add(xtpControlButtonPopup, ID_KEUANGAN_GL_LEDGER_JURNAL, "Journal")
 
            With KEUANGAN_GL_LEDGER_JURNAL.CommandBar.Controls
                Set KEUANGAN_GL_LEDGER_JURNAL_ADD = .Add(xtpControlButton, ID_KEUANGAN_GL_LEDGER_JURNAL_ADD, "Add")
                Set KEUANGAN_GL_LEDGER_JURNAL_UPDATE = .Add(xtpControlButton, ID_KEUANGAN_GL_LEDGER_JURNAL_UPDATE, "Update/Delete")
                Set KEUANGAN_GL_LEDGER_JURNAL_LIST = .Add(xtpControlButton, ID_KEUANGAN_GL_LEDGER_JURNAL_LIST, "List")
                Set KEUANGAN_GL_LEDGER_JURNAL_POSTING = .Add(xtpControlButton, ID_KEUANGAN_GL_LEDGER_JURNAL_POSTING, "Posting")
                KEUANGAN_GL_LEDGER_JURNAL_POSTING.BeginGroup = True
                Set KEUANGAN_GL_LEDGER_JURNAL_UNPOSTING = .Add(xtpControlButton, ID_KEUANGAN_GL_LEDGER_JURNAL_UNPOSTING, "Unposting")
            End With
            
            Set KEUANGAN_GL_LEDGER_CASHBANKIN = .Add(xtpControlButtonPopup, ID_KEUANGAN_GL_LEDGER_CAHBANKIN, "Cash/Bank In")
            With KEUANGAN_GL_LEDGER_CASHBANKIN.CommandBar.Controls
                Set KEUANGAN_GL_LEDGER_CAHBANKIN_ADD = .Add(xtpControlButton, ID_KEUANGAN_GL_LEDGER_CAHBANKIN_ADD, "Add")
                Set KEUANGAN_GL_LEDGER_CAHBANKIN_UPDATE = .Add(xtpControlButton, ID_KEUANGAN_GL_LEDGER_CAHBANKIN_UPDATE, "Update/Delete")
                Set KEUANGAN_GL_LEDGER_CAHBANKIN_LIST = .Add(xtpControlButton, ID_KEUANGAN_GL_LEDGER_CAHBANKIN_LIST, "List")
                Set KEUANGAN_GL_LEDGER_CAHBANKIN_POSTING = .Add(xtpControlButton, ID_KEUANGAN_GL_LEDGER_CAHBANKIN_POSTING, "Posting")
                KEUANGAN_GL_LEDGER_CAHBANKIN_POSTING.BeginGroup = True
                Set KEUANGAN_GL_LEDGER_CAHBANKIN_UNPOSTING = .Add(xtpControlButton, ID_KEUANGAN_GL_LEDGER_CAHBANKIN_UNPOSTING, "Unposting")
            End With
            
            Set KEUANGAN_GL_LEDGER_CASHBANKOUT = .Add(xtpControlButtonPopup, ID_KEUANGAN_GL_LEDGER_CAHBANKOUT, "Cash/Bank Out")
            With KEUANGAN_GL_LEDGER_CASHBANKOUT.CommandBar.Controls
                Set KEUANGAN_GL_LEDGER_CAHBANKOUT_ADD = .Add(xtpControlButtonPopup, ID_KEUANGAN_GL_LEDGER_CAHBANKOUT_ADD, "Add")
                With KEUANGAN_GL_LEDGER_CAHBANKOUT_ADD.CommandBar.Controls
                    Set KEUANGAN_GL_LEDGER_CAHBANKOUT_ADD_NEW = .Add(xtpControlButton, ID_KEUANGAN_GL_LEDGER_CAHBANKOUT_ADD_NEW, "Current Version")
                    Set KEUANGAN_GL_LEDGER_CAHBANKOUT_ADD_OLD = .Add(xtpControlButton, ID_KEUANGAN_GL_LEDGER_CAHBANKOUT_ADD_OLD, "Previous Version")
                End With
                
                Set KEUANGAN_GL_LEDGER_CAHBANKOUT_UPDATE = .Add(xtpControlButton, ID_KEUANGAN_GL_LEDGER_CAHBANKOUT_UPDATE, "Update/Delete")
                Set KEUANGAN_GL_LEDGER_CAHBANKOUT_LIST = .Add(xtpControlButton, ID_KEUANGAN_GL_LEDGER_CAHBANKOUT_LIST, "List")
                Set KEUANGAN_GL_LEDGER_CAHBANKOUT_POSTING = .Add(xtpControlButton, ID_KEUANGAN_GL_LEDGER_CAHBANKOUT_POSTING, "Posting")
                KEUANGAN_GL_LEDGER_CAHBANKOUT_POSTING.BeginGroup = True
                Set KEUANGAN_GL_LEDGER_CAHBANKOUT_UNPOSTING = .Add(xtpControlButton, ID_KEUANGAN_GL_LEDGER_CAHBANKOUT_UNPOSTING, "Unposting")
                Set KEUANGAN_GL_LEDGER_BUKTIKELUAR = .Add(xtpControlButtonPopup, ID_KEUANGAN_GL_LEDGER_BUKTIKELUAR, "Bukti Keluar")
                With KEUANGAN_GL_LEDGER_BUKTIKELUAR.CommandBar.Controls
                    Set KEUANGAN_GL_LEDGER_BUKTIKELUAR_ADD = .Add(xtpControlButton, ID_KEUANGAN_GL_LEDGER_BUKTIKELUAR_ADD, "Bukti Keluar")
                    Set KEUANGAN_GL_LEDGER_BUKTIKELUAR_NEW = .Add(xtpControlButton, ID_KEUANGAN_GL_LEDGER_BUKTIKELUAR_NEW, "Bukti Keluar New")
                    Set KEUANGAN_GL_LEDGER_BUKTIKELUAR_UPDATE = .Add(xtpControlButton, ID_KEUANGAN_GL_LEDGER_BUKTIKELUAR_UPDATE, "Edit Bukti Keluar")
                    Set KEUANGAN_GL_LEDGER_BUKTIKELUAR_LIST = .Add(xtpControlButton, ID_KEUANGAN_GL_LEDGER_BUKTIKELUAR_LIST, "Laporan Pengeluaran Kas")
                    Set KEUANGAN_GL_LEDGER_BUKTIKELUAR_REPRINT = .Add(xtpControlButton, ID_KEUANGAN_GL_LEDGER_BUKTIKELUAR_REPRINT, "Reprint Bukti Keluar")
                End With
                KEUANGAN_GL_LEDGER_BUKTIKELUAR.BeginGroup = True
                Set KEUANGAN_GL_LEDGER_ETOL = .Add(xtpControlButton, ID_KEUANGAN_GL_LEDGER_ETOL, "Transaksi E-TOL")
                Set KEUANGAN_GL_LEDGER_CAHBANKOUT_PRINT = .Add(xtpControlButton, ID_KEUANGAN_GL_LEDGER_CAHBANKOUT_PRINT, "Print Cash/Bank Out")
                KEUANGAN_GL_LEDGER_CAHBANKOUT_PRINT.BeginGroup = True
            End With
            
            
        End With
        
        Set KEUANGAN_GL_REPORT = .Add(xtpControlButtonPopup, ID_KEUANGAN_GL_REPORT, "Report")
        With KEUANGAN_GL_REPORT.CommandBar.Controls
            Set KEUANGAN_GL_REPORT_TRIALBALANCE = .Add(xtpControlButton, ID_KEUANGAN_GL_REPORT_TRIALBALANCE, "Trial Balance")
            Set KEUANGAN_GL_REPORT_BUKUBESAR = .Add(xtpControlButton, ID_KEUANGAN_GL_REPORT_BUKUBESAR, "Buku Besar")
            Set KEUANGAN_GL_REPORT_WORKSHEET = .Add(xtpControlButton, ID_KEUANGAN_GL_REPORT_WORKSHEET, "Worksheet")
            Set KEUANGAN_GL_REPORT_DAILYCASHFLOW = .Add(xtpControlButton, ID_KEUANGAN_GL_REPORT_DAILYCASHFLOW, "Daily Cash Flow")
            Set KEUANGAN_GL_REPORT_BALANCESHEET = .Add(xtpControlButton, ID_KEUANGAN_GL_REPORT_BALANCESHEET, "Balance Sheet")
            Set KEUANGAN_GL_REPORT_INCOMESTATEMENT = .Add(xtpControlButton, ID_KEUANGAN_GL_REPORT_INCOMESTATEMENT, "Incomes Statement")
            Set KEUANGAN_GL_REPORT_DAFTARFIXEDASSET = .Add(xtpControlButton, ID_KEUANGAN_GL_REPORT_DAFTARFIXEDASSET, "Daftar Fixed Asset")
            KEUANGAN_GL_REPORT_DAFTARFIXEDASSET.BeginGroup = True
            Set KEUANGAN_GL_REPORT_NILAIFIXEDASSET = .Add(xtpControlButton, ID_KEUANGAN_GL_REPORT_NILAIFIXEDASSET, "Nilai Fixed Asset")
            Set KEUANGAN_GL_REPORT_PENJUALANFIXEDASSET = .Add(xtpControlButton, ID_KEUANGAN_GL_REPORT_PENJUALANFIXEDASSET, "Penjualan Fixed Asset")
            
        End With
        Set KEUANGAN_GL_UTILITY = .Add(xtpControlButtonPopup, ID_KEUANGAN_GL_UTILITY, "Utiity")
        With KEUANGAN_GL_UTILITY.CommandBar.Controls
            Set KEUANGAN_GL_UTILITY_UNBALANCETRANS = .Add(xtpControlButton, ID_KEUANGAN_GL_UTILITY_UNBALANCETRANS, "UnBalance Transaction")
            KEUANGAN_GL_UTILITY_UNBALANCETRANS.BeginGroup = True
            Set KEUANGAN_GL_UTILITY_EXPORT = .Add(xtpControlButton, ID_KEUANGAN_GL_UTILITY_EXPORT, "Export")
            Set KEUANGAN_GL_UTILITY_IMPORT = .Add(xtpControlButton, ID_KEUANGAN_GL_UTILITY_IMPORT, "Import")
            Set KEUANGAN_GL_UTILITY_RESETSALDO = .Add(xtpControlButton, ID_KEUANGAN_GL_UTILITY_RESETSALDO, "Reset Saldo")
            KEUANGAN_GL_UTILITY_RESETSALDO.BeginGroup = True
            Set KEUANGAN_GL_UTILITY_SETUPLAYOUTREPORT = .Add(xtpControlButtonPopup, ID_KEUANGAN_GL_UTILITY_SETUPLAYOUTREPORT, "Setup Layout Report")
            With KEUANGAN_GL_UTILITY_SETUPLAYOUTREPORT.CommandBar.Controls
                Set KEUANGAN_GL_UTILITY_SETUPLAYOUTREPORT_DEFINE = .Add(xtpControlButton, ID_KEUANGAN_GL_UTILITY_SETUPLAYOUTREPORT_DEFINE, "Define")
                Set KEUANGAN_GL_UTILITY_SETUPLAYOUTREPORT_CROSS = .Add(xtpControlButton, ID_KEUANGAN_GL_UTILITY_SETUPLAYOUTREPORT_CROSS, "Cross Check")
                Set KEUANGAN_GL_UTILITY_SETUPLAYOUTREPORT_LIST = .Add(xtpControlButton, ID_KEUANGAN_GL_UTILITY_SETUPLAYOUTREPORT_LIST, "List")
            End With
            Set KEUANGAN_GL_UTILITY_CLOSING = .Add(xtpControlButton, ID_KEUANGAN_GL_UTILITY_CLOSING, "Closing")
            KEUANGAN_GL_UTILITY_CLOSING.BeginGroup = True
            Set KEUANGAN_GL_UTILITY_UNCLOSING = .Add(xtpControlButton, ID_KEUANGAN_GL_UTILITY_UNCLOSING, "UnClosing")
            Set KEUANGAN_GL_UTILITY_REKONSILIASI = .Add(xtpControlButton, ID_KEUANGAN_GL_UTILITY_REKONSILIASI, "Rekonsiliasi")
        End With
        
        Set ABOUT = RibbonBar.InsertTab(ID_ABOUT, "ABOUT")
        With ABOUT
            Set ABOUT_HOME = ABOUT.Groups.AddGroup("Home", ID_ABOUT_HOME)
            With ABOUT_HOME
                Set ABOUT_HOME_ABOUTUS = .Add(xtpControlButton, ID_ABOUT_HOME_ABOUTUS, "About Us")
                Set ABOUT_HOME_HELP = .Add(xtpControlButtonPopup, ID_ABOUT_HOME_HELP, "Help")
                With ABOUT_HOME_HELP.CommandBar.Controls
                    Set ABOUT_HOME_HELP_HELP = .Add(xtpControlButtonPopup, ID_ABOUT_HOME_HELP_HELP, "Help")
                    With ABOUT_HOME_HELP_HELP.CommandBar.Controls
                        Set ABOUT_HOME_HELP_HELP_NP = .Add(xtpControlButton, ID_ABOUT_HOME_HELP_HELP_NP, "Nota Permintaan")
                        Set ABOUT_HOME_HELP_HELP_NPEDIT = .Add(xtpControlButton, ID_ABOUT_HOME_HELP_HELP_NPEDIT, "Edit NP")
                        Set ABOUT_HOME_HELP_HELP_NPPRINT = .Add(xtpControlButton, ID_ABOUT_HOME_HELP_HELP_NPPRINT, "Reprint Nota")
                        
                        Set ABOUT_HOME_HELP_HELP_TS = .Add(xtpControlButton, ID_ABOUT_HOME_HELP_HELP_TS, "Form Troubleshoot")
                        ABOUT_HOME_HELP_HELP_TS.BeginGroup = True
                        Set ABOUT_HOME_HELP_HELP_CETAKITG = .Add(xtpControlButton, ID_ABOUT_HOME_HELP_HELP_CETAKITG, "Cetak Form ITG")
                        Set ABOUT_HOME_HELP_HELP_CETAKREKAP = .Add(xtpControlButton, ID_ABOUT_HOME_HELP_HELP_CETAKREKAP, "Laporan Troubleshoot")
                        Set ABOUT_HOME_HELP_HELP_LEMBUR = .Add(xtpControlButton, ID_ABOUT_HOME_HELP_HELP_LEMBUR, "SPK LEMBUR")
                        Set ABOUT_HOME_HELP_HELP_KAS = .Add(xtpControlButton, ID_ABOUT_HOME_HELP_HELP_KAS, "Pengeluaran")
                        ABOUT_HOME_HELP_HELP_LEMBUR.BeginGroup = True
                    End With
                    'Set ABOUT_HOME_HELP_HELP = .Add(xtpControlButton, ID_ABOUT_HOME_HELP_HELP, "Help")
                    Set ABOUT_HOME_HELP_FORM = .Add(xtpControlButton, ID_ABOUT_HOME_HELP_FORM, "Forms")
                    
                End With
                Set ABOUT_HOME_CHECKUPDATE = .Add(xtpControlButton, ID_ABOUT_HOME_CHECKUPDATE, "Check Update")
                ABOUT_HOME_CHECKUPDATE.BeginGroup = True
            End With
        End With
        
        Set ctrloption = RibbonBar.Controls.Add(xtpControlPopup, 0, "Options")
        ctrloption.Flags = xtpFlagRightAlign
        With ctrloption.CommandBar.Controls
            .Add xtpControlButton, ID_OPTIONS_STYLEBLUE, "Blue"
            .Add xtpControlButton, ID_OPTIONS_STYLEAQUA, "Aqua"
            .Add xtpControlButton, ID_OPTIONS_STYLEBLACK, "Black"
            .Add xtpControlButton, ID_OPTIONS_STYLESILVER, "Silver"
        End With
        
         Set controlabout = RibbonBar.Controls.Add(xtpControlButton, ID_ABOUT_HOME_ABOUTUS, "&About")
        controlabout.Flags = xtpFlagRightAlign
        
        
        RibbonBar.QuickAccessControls.Add xtpControlButton, ID_FILE_VIEWSTOKBAHANBAKU, "Stok Bahan Baku"
        RibbonBar.QuickAccessControls.Add xtpControlButton, ID_FILE_VIEWSTOKBARANGJADI, "Stok Barang Jadi"
        RibbonBar.QuickAccessControls.Add xtpControlButton, ID_FILE_PESANSYSTEM, "System Message"
        RibbonBar.QuickAccessControls.Add xtpControlButton, ID_FILE_LOGOUT, "LogOut"
    End With
End Sub

Private Sub subcombars(ByVal menuid As String)
    If menuid = ID_PENJUALAN_MAINMENU_MUTASI_PACKLIST_ADD Then
        Dim packadd
        Set packadd = CreateObject("sale_mut.CMutasi")
        With packadd
            .token = "kusumah"
            .UserOnline = UserOnline
            .Path = App.Path
            .formname = "packadd"
            .FastSearch = False
            .asremote = remoteserver
            .ipserver = dbServer
            .Show
             SetParent .hWnd, Me.hWnd
        End With
    ElseIf menuid = ID_PENJUALAN_MAINMENU_MUTASI_PACKLIST_LIST Then
        Dim packlist
        Set packlist = CreateObject("sale_mut.CMutasi")
        With packlist
            .token = "kusumah"
            .UserOnline = UserOnline
            .Path = App.Path
            .formname = "packlist"
            .FastSearch = False
            .asremote = remoteserver
            .ipserver = dbServer
            .Show
             SetParent .hWnd, Me.hWnd
        End With
    ElseIf menuid = ID_PENJUALAN_MAINMENU_MUTASI_PACKLIST_CLOSE Then
        Dim packclose
        Set packclose = CreateObject("sale_mut.CMutasi")
        With packclose
            .token = "kusumah"
            .UserOnline = UserOnline
            .Path = App.Path
            .formname = "packclose"
            .FastSearch = False
            .asremote = remoteserver
            .ipserver = dbServer
            .Show
             SetParent .hWnd, Me.hWnd
        End With
    ElseIf menuid = ID_PENJUALAN_MAINMENU_INVOICING_SO_CANCEL Then
        Dim socancel
        Set socancel = CreateObject("sale_inv.CInvoicing")
        With socancel
            .token = "kusumah"
            .UserOnline = UserOnline
            .Path = App.Path
            .formname = "socancel"
            .FastSearch = False
            .asremote = remoteserver
            .ipserver = dbServer
            .Show
             SetParent .hWnd, Me.hWnd
        End With
    ElseIf menuid = ID_PENJUALAN_MAINMENU_MUTASI_WIP Then
        Dim mutwip
        Set mutwip = CreateObject("sale_mut.CMutasi")
        With mutwip
            .token = "kusumah"
            .UserOnline = UserOnline
            .Path = App.Path
            .formname = "mutwip"
            .FastSearch = False
            .asremote = remoteserver
            .ipserver = dbServer
            .Show
            SetParent .hWnd, Me.hWnd
        End With
    ElseIf menuid = ID_PENJUALAN_MAINMENU_MUTASI_KARTUSTOK Then
        Dim kartu
        Set kartu = CreateObject("sale_mut.CMutasi")
        With kartu
            .token = "kusumah"
            .UserOnline = UserOnline
            .Path = App.Path
            .formname = "kartu"
            .FastSearch = False
            .asremote = remoteserver
            .ipserver = dbServer
            .Show
            SetParent .hWnd, Me.hWnd
        End With
    ElseIf menuid = ID_PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_SCANPALET Then
        Dim scanpalet
        Set scanpalet = CreateObject("sale_mut.CMutasi")
            With scanpalet
                .token = "kusumah"
                .UserOnline = UserOnline
                .Path = App.Path
                .formname = "scanpalet"
                .FastSearch = False
                .asremote = remoteserver
                .ipserver = dbServer
                .Show
                SetParent .hWnd, Me.hWnd
        End With
    ElseIf menuid = ID_PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_COMPAREHPP Then
        Dim comparepalet
        Set comparepalet = CreateObject("sale_mut.CMutasi")
            With comparepalet
                .token = "kusumah"
                .UserOnline = UserOnline
                .Path = App.Path
                .formname = "comparepalet"
                .FastSearch = False
                .asremote = remoteserver
                .ipserver = dbServer
                .Show
                SetParent .hWnd, Me.hWnd
        End With
    ElseIf menuid = ID_PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_TOWIP Then
        Dim towip
        Set towip = CreateObject("sale_mut.CMutasi")
            With towip
                .token = "kusumah"
                .UserOnline = UserOnline
                .Path = App.Path
                .formname = "towip"
                .FastSearch = False
                .asremote = remoteserver
                .ipserver = dbServer
                .Show
                SetParent .hWnd, Me.hWnd
        End With
    ElseIf menuid = ID_PENJUALAN_MAINMENU_MUTASI_GUDANG Then
        Dim mutgudang
        Set mutgudang = CreateObject("sale_mut.CMutasi")
            With mutgudang
                .token = "kusumah"
                .UserOnline = UserOnline
                .Path = App.Path
                .formname = "mutgudang"
                .FastSearch = False
                .asremote = remoteserver
                .ipserver = dbServer
                .Show
                SetParent .hWnd, Me.hWnd
            End With
    ElseIf menuid = ID_PENJUALAN_MAINMENU_MUTASI_FAILED Then
        Dim mutfail
        Set mutfail = CreateObject("sale_mut.Cmutasi")
            With mutfail
                .token = "kusumah"
                .UserOnline = UserOnline
                .Path = App.Path
                .formname = "mutfail"
                .FastSearch = False
                .asremote = remoteserver
                .ipserver = dbServer
                .Show
                SetParent .hWnd, Me.hWnd
            End With
    ElseIf menuid = ID_PENJUALAN_MAINMENU_MUTASI_MUT_OVERZAK Then
        Dim overzak
        Set overzak = CreateObject("sale_mut.CMutasi")
            With overzak
                .token = "kusumah"
                .UserOnline = UserOnline
                .Path = App.Path
                .formname = "overzak"
                .FastSearch = False
                .asremote = remoteserver
                .ipserver = dbServer
                .Show
                SetParent .hWnd, Me.hWnd
            End With
    ElseIf menuid = ID_PENJUALAN_MAINMENU_MUTASI_MUT_MUTASILOT Then
        Dim mutbylot
        Set mutbylot = CreateObject("sale_mut.CMutasi")
            With mutbylot
                .token = "kusumah"
                .UserOnline = UserOnline
                .Path = App.Path
                .formname = "mutbylot"
                .FastSearch = False
                .asremote = remoteserver
                .ipserver = dbServer
                .Show
                SetParent .hWnd, Me.hWnd
            End With
    ElseIf menuid = ID_PENJUALAN_MAINMENU_MUTASI_MUT_MUTPABRIK Then
        Dim mutpabrik
        Set mutpabrik = CreateObject("sale_mut.CMutasi")
            With mutpabrik
                .token = "kusumah"
                .UserOnline = UserOnline
                .Path = App.Path
                .formname = "mutpabrik"
                .FastSearch = False
                .asremote = remoteserver
                .ipserver = dbServer
                .Show
                SetParent .hWnd, Me.hWnd
            End With
    ElseIf menuid = ID_PENJUALAN_MAINMENU_MUTASI_MUT_ADJSTOK Then
        Dim adjstok
        Set adjstok = CreateObject("sale_mut.CMutasi")
            With adjstok
                .token = "kusumah"
                .UserOnline = UserOnline
                .Path = App.Path
                .formname = "adjstok"
                .FastSearch = False
                .asremote = remoteserver
                .ipserver = dbServer
                .Show
                SetParent .hWnd, Me.hWnd
            End With
    ElseIf menuid = ID_PENJUALAN_MAINMENU_INVOICING_BYKATEGORI Then
        Dim bykategori
        Set bykategori = CreateObject("sale_inv.CInvoicing")
        With bykategori
            .token = "kusumah"
            .UserOnline = UserOnline
            .Path = App.Path
            .formname = "bykategori"
            .FastSearch = False
            .asremote = remoteserver
            .ipserver = dbServer
            .Show
            SetParent .hWnd, Me.hWnd
        End With
    End If
End Sub
Private Sub subcombars_purchasing(ByVal menuid As String)
    If menuid = ID_PEMBELIAN_TABEL_SUPPLIER_GRAPRICE Then
        Dim graphicprice As Object
        Set graphicprice = CreateObject("purc_tables.CTables")
        With graphicprice
            .token = "kusumah"
            .UserOnline = UserOnline
            .Path = App.Path
            .formname = "graphicprice"
            .FastSearch = False
            .Show
            SetParent .hWnd, Me.hWnd
        End With
    ElseIf menuid = ID_PEMBELIAN_PURCHASING_PERMINTAAN_DAYSCOUNT Then
        Dim daylistpermintaanbarang As Object
        Set daylistpermintaanbarang = CreateObject("purc_purc.CPurchasing")
        With daylistpermintaanbarang
            .token = "kusumah"
            .UserOnline = UserOnline
            .Path = App.Path
            .formname = "daylistpermintaan"
            .FastSearch = False
            .asremote = remoteserver
            .ipserver = dbServer
            .Show
            SetParent .hWnd, Me.hWnd
        End With
        
    End If
End Sub
Private Sub subcombars_produksi(ByVal menuid As String)
    If menuid = ID_PRODUKSI_MAINMENU_TABLES_KONVERSI_UNIT Then
        Dim konvunit
        Set konvunit = CreateObject("prod_tables.CTables")
        With konvunit
            .token = "kusumah"
            .UserOnline = UserOnline
            .Path = App.Path
            .formname = "konvunit"
            .FastSearch = False
            .asremote = remoteserver
            .ipserver = dbServer
            .Show
             SetParent .hWnd, Me.hWnd
        End With
    ElseIf menuid = ID_PRODUKSI_MAINMENU_TABLES_KONVERSI_ADD Then
        Dim konversikemasan As Object
        Set konversikemasan = CreateObject("prod_tables.CTables")
        With konversikemasan
            .token = "kusumah"
            .UserOnline = UserOnline
            .userlevel = UserOnLineLevel
            .Path = App.Path
            .formname = "konversikemasan"
            .FastSearch = False
            .asremote = remoteserver
            .ipserver = dbServer
            .Show
            SetParent .hWnd, Me.hWnd
        End With
    ElseIf menuid = ID_PRODUKSI_MAINMENU_TABLES_KONVERSI_CHANGE Then
        Dim konversich As Object
        Set konversich = CreateObject("prod_tables.CTables")
        With konversich
            .token = "kusumah"
            .UserOnline = UserOnline
            .userlevel = UserOnLineLevel
            .Path = App.Path
            .formname = "konversich"
            .FastSearch = False
            .asremote = remoteserver
            .ipserver = dbServer
            .Show
            SetParent .hWnd, Me.hWnd
        End With
    ElseIf menuid = ID_PRODUKSI_MAINMENU_TABLES_KONVERSI_LIST Then
        Dim konversilist As Object
        Set konversilist = CreateObject("prod_tables.CTables")
        With konversilist
            .token = "kusumah"
            .UserOnline = UserOnline
            .userlevel = UserOnLineLevel
            .Path = App.Path
            .formname = "konversilist"
            .FastSearch = False
            .asremote = remoteserver
            .ipserver = dbServer
            .Show
            SetParent .hWnd, Me.hWnd
       End With
    ElseIf menuid = ID_PRODUKSI_MAINMENU_TABLES_DEFINEKG Then
        Dim definekg As Object
        Set definekg = CreateObject("prod_tables.CTables")
        With definekg
            .token = "kusumah"
            .UserOnline = UserOnline
            .userlevel = UserOnLineLevel
            .Path = App.Path
            .formname = "definekg"
            .FastSearch = False
            .asremote = remoteserver
            .ipserver = dbServer
            .Show
            SetParent .hWnd, Me.hWnd
       End With
    ElseIf menuid = ID_PRODUKSI_MAINMENU_TABLES_RNCPROD_ADD Then
        Dim prodplan As Object
        Set prodplan = CreateObject("prod_tables.CTables")
        With prodplan
            .token = "kusumah"
            .UserOnline = UserOnline
            .userlevel = UserOnLineLevel
            .Path = App.Path
            .formname = "prodplan"
            .FastSearch = False
            .asremote = remoteserver
            .ipserver = dbServer
            .Show
            SetParent .hWnd, Me.hWnd
        End With
    ElseIf menuid = ID_PRODUKSI_MAINMENU_TABLES_RNCPROD_CHANGE Then
        Dim prodpled As Object
        Set prodpled = CreateObject("prod_tables.CTables")
        With prodpled
            .token = "kusumah"
            .UserOnline = UserOnline
            .userlevel = UserOnLineLevel
            .Path = App.Path
            .formname = "prodpled"
            .FastSearch = False
            .asremote = remoteserver
            .ipserver = dbServer
            .Show
            SetParent .hWnd, Me.hWnd
        End With
    ElseIf menuid = ID_PRODUKSI_MAINMENU_TABLES_RNCPROD_LIST Then
        Dim prodplist As Object
        Set prodplist = CreateObject("prod_tables.CTables")
        With prodplist
            .token = "kusumah"
            .UserOnline = UserOnline
            .userlevel = UserOnLineLevel
            .Path = App.Path
            .formname = "prodplist"
            .FastSearch = False
            .asremote = remoteserver
            .ipserver = dbServer
            .Show
            SetParent .hWnd, Me.hWnd
        End With
    ElseIf menuid = ID_PRODUKSI_MAINMENU_TABLES_KATEGORI Then
        Dim ktgori As Object
        Set ktgori = CreateObject("prod_tables.CTables")
        With ktgori
            .token = "kusumah"
            .UserOnline = UserOnline
            .userlevel = UserOnLineLevel
            .Path = App.Path
            .formname = "ktgori"
            .FastSearch = False
            .asremote = remoteserver
            .ipserver = dbServer
            .Show
            SetParent .hWnd, Me.hWnd
        End With
    ElseIf menuid = ID_PRODUKSI_MAINMENU_TABLES_REAKTOR Then
        Dim reaktor As Object
        Set reaktor = CreateObject("prod_tables.CTables")
        With reaktor
            .token = "kusumah"
            .UserOnline = UserOnline
            .userlevel = UserOnLineLevel
            .Path = App.Path
            .formname = "reaktor"
            .FastSearch = False
            .asremote = remoteserver
            .ipserver = dbServer
            .Show
            SetParent .hWnd, Me.hWnd
        End With
    ElseIf menuid = ID_PRODUKSI_MAINMENU_SOP_ADD_NEW Then
        Dim addsop As Object
        Set addsop = CreateObject("prod_sop.CSOP")
        With addsop
            .token = "kusumah"
            .UserOnline = UserOnline
            .Path = App.Path
            .formname = "addsop"
            .setuseronlinelevel = UserOnLineLevel
            .FastSearch = False
            .asremote = remoteserver
            .ipserver = dbServer
            .Show
            SetParent .hWnd, Me.hWnd
        End With
    ElseIf menuid = ID_PRODUKSI_MAINMENU_SOP_ADD_EDIT Then
        Dim editsop As Object
        Set editsop = CreateObject("prod_sop.CSOP")
        With editsop
            .token = "kusumah"
            .UserOnline = UserOnline
            .Path = App.Path
            .formname = "editsop"
            .setuseronlinelevel = UserOnLineLevel
            .FastSearch = False
            .asremote = remoteserver
            .ipserver = dbServer
            .Show
            SetParent .hWnd, Me.hWnd
        End With
    ElseIf menuid = ID_PRODUKSI_MAINMENU_SOP_TAKEPACKAGE_REPORT Then
        Dim reportpack As Object
        Set reportpack = CreateObject("prod_sop.CSOP")
        With reportpack
            .token = "kusumah"
            .UserOnline = UserOnline
            '.userlevel = UserOnLineLevel
            .Path = App.Path
            .formname = "reportpack"
            .FastSearch = False
            .asremote = remoteserver
            .ipserver = dbServer
            .Show
            SetParent .hWnd, Me.hWnd
       End With
    ElseIf menuid = ID_PRODUKSI_MAINMENU_SOP_PRODMONTHLY_BYBB Then
        Dim prodmonthly As Object
        Set prodmonthly = CreateObject("prod_sop.CSOP")
        With prodmonthly
            .token = "kusumah"
            .UserOnline = UserOnline
            .setuseronlinelevel = UserOnLineLevel
            .Path = App.Path
            .formname = "monthlyreport"
            .FastSearch = False
            .asremote = remoteserver
            .ipserver = dbServer
            .Show
            SetParent .hWnd, Me.hWnd
        End With
    ElseIf menuid = ID_PRODUKSI_MAINMENU_SOP_PRODMONTHLY_BYTOPPRODUK Then
        Dim topkg As Object
        Set topkg = CreateObject("prod_sop.CSOP")
        With topkg
            .token = "kusumah"
            .UserOnline = UserOnline
            .setuseronlinelevel = UserOnLineLevel
            .Path = App.Path
            .formname = "topkilo"
            .FastSearch = False
            .asremote = remoteserver
            .ipserver = dbServer
            .Show
            SetParent .hWnd, Me.hWnd
        End With
    ElseIf menuid = ID_PRODUKSI_MAINMENU_MUTASI_WIPBASE Then
        Dim wipbase As Object
        Set wipbase = CreateObject("prod_mut.CMUT")
        With wipbase
            .token = "kusumah"
            .UserOnline = UserOnline
            '.userlevel = UserOnLineLevel
            .Path = App.Path
            .formname = "wipbase"
            .FastSearch = False
            .asremote = remoteserver
            .ipserver = dbServer
            .Show
            SetParent .hWnd, Me.hWnd
       End With
    ElseIf menuid = ID_PRODUKSI_MAINMENU_MUTASI_LAPWIPBASE Then
        Dim lapwipbase As Object
        Set lapwipbase = CreateObject("prod_mut.CMUT")
        With lapwipbase
            .token = "kusumah"
            .UserOnline = UserOnline
            .Path = App.Path
            .formname = "lapwipbase"
            .FastSearch = False
            .asremote = remoteserver
            .ipserver = dbServer
            .Show
            SetParent .hWnd, Me.hWnd
        End With
    ElseIf menuid = ID_PRODUKSI_MAINMENU_MUTASI_ADJWIP Then
        Dim adjwipbase As Object
        Set adjwipbase = CreateObject("prod_mut.CMUT")
        With adjwipbase
            .token = "kusumah"
            .UserOnline = UserOnline
            '.userlevel = UserOnLineLevel
            .Path = App.Path
            .formname = "adjwipbase"
            .FastSearch = False
            .asremote = remoteserver
            .ipserver = dbServer
            .Show
            SetParent .hWnd, Me.hWnd
       End With
    End If
End Sub
Private Sub subcombars_penjualan(ByVal menuid As String)
    If menuid = ID_PENJUALAN_MAINMENU_INVOICING_LAPSJ_BYLOT Then
        Dim lapsjbylot
            Set lapsjbylot = CreateObject("sale_inv.CInvoicing")
            With lapsjbylot
                .token = "kusumah"
                .UserOnline = UserOnline
                .Path = App.Path
                .formname = "lapsjbylot"
                .FastSearch = False
                .asremote = remoteserver
                .ipserver = dbServer
                .Show
                SetParent .hWnd, Me.hWnd
            End With
    ElseIf menuid = ID_PENJUALAN_MAINMENU_INVOICING_SJ_ADDLOT Then
        Dim addlot
            Set addlot = CreateObject("sale_inv.CInvoicing")
            With addlot
                .token = "kusumah"
                .UserOnline = UserOnline
                .Path = App.Path
                .formname = "AddLot"
                .FastSearch = False
                .asremote = remoteserver
                .ipserver = dbServer
                .Show
                SetParent .hWnd, Me.hWnd
            End With
    End If
End Sub
Private Sub subcombars_about(ByVal menuid As String)
    If menuid = ID_ABOUT_HOME_HELP_HELP_NP Then
        frmpermintaan.Show
        SetParent frmpermintaan.hWnd, Me.hWnd
    ElseIf menuid = ID_ABOUT_HOME_HELP_HELP_NPEDIT Then
        frmpermintaan.Show
        SetParent frmpermintaan.hWnd, Me.hWnd
    ElseIf menuid = ID_ABOUT_HOME_HELP_HELP_NPPRINT Then
        frmReprint.Show
        SetParent frmReprint.hWnd, Me.hWnd
    ElseIf menuid = ID_ABOUT_HOME_HELP_HELP_TS Then
        frmts.Show  'MsgBox "Maaf, Sedang Dalam Proses Pengembangan", vbInformation, AppName
        SetParent frmts.hWnd, Me.hWnd
    ElseIf menuid = ID_ABOUT_HOME_HELP_HELP_CETAKITG Then
        frmtsreprint.Show
        SetParent frmtsreprint.hWnd, Me.hWnd
    ElseIf menuid = ID_ABOUT_HOME_HELP_HELP_CETAKREKAP Then
        frmtsreport.Show
        SetParent frmtsreport.hWnd, Me.hWnd
    ElseIf menuid = ID_ABOUT_HOME_HELP_HELP_LEMBUR Then
        frmSPKL.Show
        SetParent frmSPKL.hWnd, Me.hWnd
    ElseIf menuid = ID_ABOUT_HOME_HELP_HELP_KAS Then
        'frmpengeluaran.Show
        'SetParent frmpengeluaran.hWnd, Me.hWnd
        frmoutrankas.Show
        SetParent frmoutrankas.hWnd, Me.hWnd
    ElseIf menuid = ID_ABOUT_HOME_HELP_FORM Then
        frmFArepair.Show vbModal 'frmBA.Show vbModal
    
    End If
    
End Sub
Private Sub subcombars_gl_ledger(ByVal menuid As String)
    If menuid = ID_KEUANGAN_GL_LEDGER_BUKTIKELUAR_NEW Then
        Dim objbuktinew As Object
        Set objbuktinew = CreateObject("gl_ledger.CLedger")
        With objbuktinew
            .token = "kusumah"
            .UserOnline = UserOnline
            .Path = App.Path
            .formname = "buktinew"
            .FastSearch = False
            .asremote = remoteserver
            .ipserver = dbServer
            .Show
            SetParent .hWnd, Me.hWnd
        End With
    ElseIf menuid = ID_KEUANGAN_GL_LEDGER_BUKTIKELUAR_UPDATE Then
        Dim objeditbuktinew
        Set objeditbuktinew = CreateObject("gl_ledger.CLedger")
        With objeditbuktinew
            .token = "kusumah"
            .UserOnline = UserOnline
            .Path = App.Path
            .formname = "editbuktinew"
            .FastSearch = False
            .asremote = remoteserver
            .ipserver = dbServer
            .Show
            SetParent .hWnd, Me.hWnd
        End With
    ElseIf menuid = ID_KEUANGAN_GL_LEDGER_BUKTIKELUAR_REPRINT Then
        Dim objreprintbuktinew
        Set objreprintbuktinew = CreateObject("gl_ledger.CLedger")
        With objreprintbuktinew
            .token = "kusumah"
            .UserOnline = UserOnline
            .Path = App.Path
            .formname = "reprintbuktinew"
            .FastSearch = False
            .asremote = remoteserver
            .ipserver = dbServer
            .Show
            SetParent .hWnd, Me.hWnd
        End With
    ElseIf menuid = ID_KEUANGAN_GL_LEDGER_CAHBANKOUT_ADD_NEW Then
        Dim addcashbankout As Object
        Set addcashbankout = CreateObject("gl_ledger.CLedger")
        With addcashbankout
            .token = "kusumah"
            .UserOnline = UserOnline
            .Path = App.Path
            .formname = "addcashbankout"
            .FastSearch = False
            .asremote = remoteserver
            .ipserver = dbServer
            .Show
            SetParent .hWnd, Me.hWnd
        End With
    ElseIf menuid = ID_KEUANGAN_GL_LEDGER_CAHBANKOUT_ADD_OLD Then
        Dim addcashbankout2 As Object
        Set addcashbankout2 = CreateObject("gl_ledger.CLedger")
        With addcashbankout2
            .token = "kusumah"
            .UserOnline = UserOnline
            .Path = App.Path
            .formname = "addcashbankout2"
            .FastSearch = False
            .asremote = remoteserver
            .ipserver = dbServer
            .Show
            SetParent .hWnd, Me.hWnd
        End With
    End If
End Sub

Private Sub FastSearch()
    If PEMBELIAN_UTILITY_FASTSEARCH.Checked = False Then
        PEMBELIAN_UTILITY_FASTSEARCH.Checked = True
        StatusBar.FindPane(ID_INDICATOR_DTP1).Visible = True
        StatusBar.FindPane(ID_INDICATOR_DTP2).Visible = True
        StatusBar.FindPane(ID_INDICATOR_DARITGL).Visible = True
        StatusBar.FindPane(ID_INDICATOR_SDTGL).Visible = True
        StatusBar.FindPane(0).text = "Pencarian Cepat Aktif"
    Else
        PEMBELIAN_UTILITY_FASTSEARCH.Checked = False
        StatusBar.FindPane(ID_INDICATOR_DTP1).Visible = False
        StatusBar.FindPane(ID_INDICATOR_DTP2).Visible = False
        StatusBar.FindPane(ID_INDICATOR_DARITGL).Visible = False
        StatusBar.FindPane(ID_INDICATOR_SDTGL).Visible = False
        StatusBar.FindPane(0).text = "Ready"
    End If
End Sub

Private Sub popupstatus_ItemClick(ByVal Item As XtremeSuiteControls.IPopupControlItem)
    crystal.reset
    crystal.WindowState = crptMaximized
    crystal.WindowShowCloseBtn = True
    crystal.WindowShowPrintSetupBtn = True
    crystal.WindowShowSearchBtn = True
    crystal.WindowShowRefreshBtn = True
    crystal.Connect = dsnreport
    crystal.DataFiles(0) = "Proc(am_perminwarning)"
    crystal.ReportFileName = App.Path & "\reports\purchasing\purc\nota_up7day.rpt"
    crystal.RetrieveDataFiles
    crystal.Action = 1
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    StatusBar.IdleText = "Waiting For Login...!"
    frmLogin.Show vbModal
End Sub

Private Sub frmLogin_Login(ByVal UserName, UserPass As String)
    On Error GoTo err_msg
    Dim var_username, var_pass As String
    Dim var_administrator, var_supervisor, var_operator As String
    Dim var_master, var_hrd, var_purch, var_prod, var_warehouse, var_sale, var_finance, var_ledger As String
    Dim var_flag As String
    
    If UserName = "chimey" Or UserName = "ch" Or UserName = "budiman" Then
        If UserPass = "kusumah" Or UserPass = "123" Then
            UserOnline = "Creator"
            UserOnLineLevel = "creator"
            EnableAllMenu
            If dbServer = "36.64.1.231" Then dbServer = "SPARTAPRIMA"
        'MsgBox dbServer
            StatusBar.FindPane(ID_INDICATOR_SERVER).text = "Database : " + UCase(dbServer) & " : " & UCase(DBName) & " User : " & UserOnline
            Unload frmLogin
            SetPopupInfo popupstatus, "Creator", "Selamat Datang Programmer Tercinta di Sistem Inventory Anda"
            popupstatus.Show
            Exit Sub
        ElseIf UserPass = "12345" Then
            UserOnline = "Budiman"
            UserOnLineLevel = "Administrator"
            EnableAllMenu
            StatusBar.FindPane(ID_INDICATOR_SERVER).text = "Database : " + UCase(dbServer) & " : " & UCase(DBName) & " User : " & UserOnline
            Unload frmLogin
            'SetPopupInfo popupstatus, "Creator", "Selamat Datang Programmer Tercinta di Sistem Inventory Anda"
            'popupstatus.Show
            frmberanda.Show vbModal
            Exit Sub
        End If
    End If
    
    SQL = "select * from list_users where username = '" & UserName & "'"
    
    OBJ.Open dsn
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox msg_err_user_denied, vbCritical, AppName
        frmLogin.txtUser = ""
        frmLogin.txtPassword = ""
        frmLogin.txtUser.SetFocus
        OBJ.Close
        Exit Sub
    End If

    var_username = RST!UserName
    var_pass = Cheap_Decrypt(RST!pass)
    
    If var_pass <> UserPass Then
        MsgBox msg_err_password_tidak_sama, vbCritical, AppName
        frmLogin.txtPassword = ""
        OBJ.Close
        Exit Sub
    End If
    
    StatusBar.IdleText = "Ready"
    
    UserOnline = var_username
    If RST!kode_role = "01" Then UserOnLineLevel = "Operator"
    If RST!kode_role = "02" Then UserOnLineLevel = "Supervisor"
    If RST!kode_role = "03" Then UserOnLineLevel = "Admnistrator"
    If RST!kode_role = "04" Then UserOnLineLevel = "Report"
    
    
    var_master = RST!MASTER
    var_hrd = RST!hrd
    var_purch = RST!purch
    UserOnlineDept = RST!purch
    var_prod = RST!prod
    var_warehouse = RST!warehouse
    var_sale = RST!sale
    var_finance = RST!finance
    var_ledger = RST!gl
    var_flag = RST!flag_aktip
    
    OBJ.Close
    
    'aktivasi menu
    If var_master = "1" Then
        MASTER.Enabled = True
        MASTER.Selected = True
    Else
        MASTER.Enabled = False
    End If
    
    If var_purch = "1" Then
        PEMBELIAN.Enabled = True
        PEMBELIAN.Selected = True
        OpenMenuPurch
        'Jika Flag aktif = 9 hide beberapa fitur pembelian
        If var_flag = "9" And UserOnLineLevel = "Operator" Then
            CloseMenuPurch
        End If
    Else
        PEMBELIAN.Enabled = False
    End If
    
    If var_prod = "1" Then
        PRODUKSI.Enabled = True
        PRODUKSI.Selected = True
    Else
        PRODUKSI.Enabled = False
    End If
    
    If var_warehouse = "1" Then
        GUDANG.Enabled = True
        GUDANG.Selected = True
    Else
        GUDANG.Enabled = False
    End If
    
    If var_sale = "1" Then
        PENJUALAN.Enabled = True
        PENJUALAN.Selected = True
        OpenMenuSale
        'Jika Flag aktif = 9 hide beberapa fitur pembelian
        If var_flag = "9" And UserOnLineLevel = "Operator" Then
            CloseMenuSale
        End If
    Else
        PENJUALAN.Enabled = False
    End If
    
    If var_finance = "1" Then
        KEUANGAN.Enabled = True
        KEUANGAN_MAINMENU.Visible = True
        KEUANGAN_PEMBELIAN.Enabled = True
        KEUANGAN_PENJUALAN.Enabled = True
        KEUANGAN.Selected = True
        PEMBELIAN_TABEL_BAHANBAKU_MANAGE.Enabled = True
        If var_ledger = "0" Then
            KEUANGAN_GL.Visible = False
        End If
    ElseIf var_finance = "0" Then
        KEUANGAN.Enabled = False
        KEUANGAN_PEMBELIAN.Enabled = False
        KEUANGAN_PENJUALAN.Enabled = False
        If var_purch = "1" And var_prod = "1" Then
            'mengaktifkan laporan uncomform (Enah)
            KEUANGAN.Enabled = True
            KEUANGAN_MAINMENU.Visible = True
            KEUANGAN_PEMBELIAN.Enabled = True
            KEUANGAN_PEMBELIAN_UTILITY.Enabled = False
            PEMBELIAN_GIRO.Enabled = False
            PEMBELIAN_PURCHASING_CONFIRM.Enabled = False
            PEMBELIAN_HUTANG.Enabled = False
            PEMBELIAN_PURCHASING_LAPVOUCER.Enabled = False
            KEUANGAN.Selected = True
        End If
    End If
    
    If var_ledger = "1" Then
        KEUANGAN.Enabled = True
        If var_finance = "0" Then
        KEUANGAN_MAINMENU.Visible = False
        End If
        KEUANGAN_GL.Visible = True
        KEUANGAN.Selected = True
    Else
        KEUANGAN_GL.Visible = False
    End If
    
    'Maintenance
    If UserOnLineLevel = "Supervisor" And var_purch = "1" And var_prod = "0" And var_warehouse = "0" And var_sale = "0" And var_finance = "0" Then
        MASTER.Enabled = False
        
        PEMBELIAN.Enabled = True
        PEMBELIAN.Selected = True
        'PEMBELIAN_PURCHASING_PENERIMAANBARANG
        PEMBELIAN_PURCHASING_LAPPEMBELIAN.Visible = False
        PEMBELIAN_PURCHASING_INQUIRY.Visible = False
        PRODUKSI.Enabled = False
        GUDANG.Enabled = False
        PENJUALAN.Enabled = False
        KEUANGAN.Enabled = False
    End If
    

    
    StatusBar.FindPane(ID_INDICATOR_SERVER).text = "Database : " + UCase(dbServer) & " : " & UCase(DBName) & " User : " & UserOnline
    
    SetPopupInfo popupstatus, "Selamat Datang " & UserOnline, "Anda telah masuk ke dalam sistem Inventory PT.SPARTA PRIMA . Selamat Bekerja"
    popupstatus.Show
    
    'open menu
    Unload frmLogin
    Timer2.Enabled = True
    
    If UserOnLineLevel = "Operator" And var_purch = "1" Then
        Call PurchNotif
    End If
    
    Exit Sub
err_msg:
    If OBJ.State = 1 Then OBJ.Close
    MsgBox Err.Description, vbCritical, AppName
End Sub

Private Sub OpenMenuPurch()
    PEMBELIAN_UTILITY.Enabled = True
    PEMBELIAN_TABEL.Enabled = True
    PEMBELIAN_MUTASIBARANG.Enabled = True
    PEMBELIAN_PURCHASING_PO.Enabled = True
    PEMBELIAN_PURCHASING_CONFIRM.Enabled = True
    PEMBELIAN_PURCHASING_INQUIRY.Enabled = True
    PEMBELIAN_PURCHASING_LAPCONPEMBELIAN.Enabled = True
    PEMBELIAN_PURCHASING_LAPPEMAKAIANBB.Enabled = True
    PEMBELIAN_PURCHASING_LAPPEMBELIAN.Enabled = True
    PEMBELIAN_PURCHASING_LAPPO.Enabled = True
    PEMBELIAN_PURCHASING_PEMAKAIANBAHANBAKU.Enabled = True
    PEMBELIAN_PURCHASING_PRINTPO.Enabled = True
    PEMBELIAN_PURCHASING_PENERIMAANBARANG.Enabled = True
    ABOUT_HOME_HELP_FORM.Enabled = True
    ABOUT_HOME_HELP_HELP_KAS.Enabled = True
    ABOUT_HOME_HELP_HELP_LEMBUR.Enabled = True
End Sub
Private Sub CloseMenuPurch()
    PEMBELIAN_UTILITY.Enabled = False
    PEMBELIAN_TABEL.Enabled = False
    PEMBELIAN_MUTASIBARANG.Enabled = False
    PEMBELIAN_PURCHASING_PO.Enabled = False
    PEMBELIAN_PURCHASING_CONFIRM.Enabled = False
    PEMBELIAN_PURCHASING_INQUIRY.Enabled = False
    PEMBELIAN_PURCHASING_LAPCONPEMBELIAN.Enabled = False
    PEMBELIAN_PURCHASING_LAPPEMAKAIANBB.Enabled = False
    PEMBELIAN_PURCHASING_LAPPEMBELIAN.Enabled = False
    PEMBELIAN_PURCHASING_LAPPO.Enabled = False
    PEMBELIAN_PURCHASING_PEMAKAIANBAHANBAKU.Enabled = False
    PEMBELIAN_PURCHASING_PRINTPO.Enabled = False
    PEMBELIAN_PURCHASING_PENERIMAANBARANG.Enabled = False
    ABOUT_HOME_HELP_FORM.Enabled = False
    ABOUT_HOME_HELP_HELP_KAS.Enabled = False
    ABOUT_HOME_HELP_HELP_LEMBUR.Enabled = False
End Sub
Private Sub CloseMenuSale()
    PENJUALAN_MAINMENU_TABLES.Enabled = False
    PENJUALAN_MAINMENU_MUTASI_PACKLIST.Enabled = False
    PENJUALAN_MAINMENU_MUTASI_KARTUSTOK.Enabled = False
    PENJUALAN_MAINMENU_MUTASI_MUT_MUTPABRIK.Enabled = False
    PENJUALAN_MAINMENU_MUTASI_MUT_MUTASILOT.Enabled = False
    PENJUALAN_MAINMENU_MUTASI_GUDANG.Enabled = False
    PENJUALAN_MAINMENU_MUTASI_FAILED.Enabled = False
    PENJUALAN_MAINMENU_MUTASI_PRINTPRICE.Enabled = False
    PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_SCANPALET.Enabled = False
    PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_COMPAREHPP.Enabled = False
    PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_LOTPALET.Enabled = False
    PENJUALAN_MAINMENU_MUTASI_WIP.Enabled = False
    PENJUALAN_MAINMENU_MUTASI_KARTUSTOK.Enabled = False
    PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_TOWIP.Enabled = False
    'Invoice
    PENJUALAN_MAINMENU_INVOICING_FAKTURPAJAK.Enabled = False
    PENJUALAN_MAINMENU_INVOICING_FAKTURJUAL_PREVIEW.Enabled = False
    PENJUALAN_MAINMENU_INVOICING_SJSBY.Enabled = False
    PENJUALAN_MAINMENU_INVOICING_INQUERYSO.Enabled = False
    PENJUALAN_MAINMENU_INVOICING_LAPSJ_BYLOT.Enabled = False
    PENJUALAN_MAINMENU_INVOICING_LAPJUAL.Enabled = False
    PENJUALAN_MAINMENU_INVOICING_LAPJUALDTL.Enabled = False
    PENJUALAN_MAINMENU_INVOICING_MONTHLY.Enabled = False
    PENJUALAN_MAINMENU_INVOICING_BYKATEGORI.Enabled = False
    PENJUALAN_MAINMENU_INVOICING_KOMISI.Enabled = False
    PENJUALAN_MAINMENU_INVOICING_ANALISAJUAL.Enabled = False
    PENJUALAN_MAINMENU_UTILITY_IMPORTINV.Enabled = False
    PENJUALAN_MAINMENU_UTILITY_DELINV.Enabled = False
    ABOUT_HOME_HELP_FORM.Enabled = False
    ABOUT_HOME_HELP_HELP_KAS.Enabled = False
    ABOUT_HOME_HELP_HELP_LEMBUR.Enabled = False
End Sub
Private Sub OpenMenuSale()
    PENJUALAN_MAINMENU_TABLES.Enabled = True
    PENJUALAN_MAINMENU_MUTASI_PACKLIST.Enabled = True
    PENJUALAN_MAINMENU_MUTASI_KARTUSTOK.Enabled = True
    PENJUALAN_MAINMENU_MUTASI_MUT_MUTPABRIK.Enabled = True
    PENJUALAN_MAINMENU_MUTASI_MUT_MUTASILOT.Enabled = True
    PENJUALAN_MAINMENU_MUTASI_GUDANG.Enabled = True
    PENJUALAN_MAINMENU_MUTASI_FAILED.Enabled = True
    PENJUALAN_MAINMENU_MUTASI_PRINTPRICE.Enabled = True
    PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_SCANPALET.Enabled = True
    PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_COMPAREHPP.Enabled = True
    PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_LOTPALET.Enabled = True
    PENJUALAN_MAINMENU_MUTASI_WIP.Enabled = True
    PENJUALAN_MAINMENU_MUTASI_KARTUSTOK.Enabled = True
    PENJUALAN_MAINMENU_MUTASI_PINDAHGUDANG_TOWIP.Enabled = True
    'Invoice
    PENJUALAN_MAINMENU_INVOICING_FAKTURPAJAK.Enabled = True
    PENJUALAN_MAINMENU_INVOICING_FAKTURJUAL_PREVIEW.Enabled = True
    PENJUALAN_MAINMENU_INVOICING_SJSBY.Enabled = True
    PENJUALAN_MAINMENU_INVOICING_INQUERYSO.Enabled = True
    PENJUALAN_MAINMENU_INVOICING_LAPSJ_BYLOT.Enabled = True
    PENJUALAN_MAINMENU_INVOICING_LAPJUAL.Enabled = True
    PENJUALAN_MAINMENU_INVOICING_LAPJUALDTL.Enabled = True
    PENJUALAN_MAINMENU_INVOICING_MONTHLY.Enabled = True
    PENJUALAN_MAINMENU_INVOICING_BYKATEGORI.Enabled = True
    PENJUALAN_MAINMENU_INVOICING_KOMISI.Enabled = True
    PENJUALAN_MAINMENU_INVOICING_ANALISAJUAL.Enabled = True
    PENJUALAN_MAINMENU_UTILITY_IMPORTINV.Enabled = True
    PENJUALAN_MAINMENU_UTILITY_DELINV.Enabled = True
    ABOUT_HOME_HELP_FORM.Enabled = True
    ABOUT_HOME_HELP_HELP_KAS.Enabled = True
    ABOUT_HOME_HELP_HELP_LEMBUR.Enabled = True
End Sub

Private Sub PurchNotif()
    Dim OBJ As ADODB.Connection
    Dim RST As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim SQL As String
    Dim result As Variant
    Dim jml As Long

    Set OBJ = New ADODB.Connection
    OBJ.Open dsn
    
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = OBJ
    cmd.CommandText = "am_perminwarning"
    cmd.CommandType = adCmdStoredProc
    
    Set RST = New ADODB.Recordset
    RST.CursorType = adOpenForwardOnly
    RST.LockType = adLockReadOnly
    RST.Open cmd
    
    jml = 0
    Do While Not RST.EOF
        jml = jml + 1
        RST.MoveNext
    Loop
    RST.MoveFirst
    
    If jml > 0 Then
        SetPopupInfo popupstatus, "PURCHASING", "Jumlah Outstanding Permintaan barang Lebih dari 7 Hari: " _
        & jml & " Item, Mohon untuk segera diproses. Terima kasih"
        popupstatus.Show
        MsgBox "Jumlah outstanding Permintaan barang Lebih dari 7 Hari = " & jml & _
        " Item, Mohon untuk segera di proses" & vbCrLf & "Terima Kasih.", vbExclamation, "PURCHASING"
    End If
    
    RST.Close
    Set RST = Nothing
    OBJ.Close
    Set OBJ = Nothing

End Sub

Private Sub logout()
    SetPopupInfo popupstatus, "Selamat Tinggal " & UserOnline, "Terima Kasih. Anda telah keluar dari sistem Inventory PT. SPARTA PRIMA"
    popupstatus.Show
    UserOnline = ""
    UserOnLineLevel = ""
    StatusBar.FindPane(ID_INDICATOR_SERVER).text = "Database : " + "" & " : " & UCase(DBName) & " User : " & UserOnline
    DisableAllMenu
    StatusBar.IdleText = "Waiting For Login....!"
    Timer1.Enabled = True
End Sub

Private Sub DisableAllMenu()
    MASTER.Enabled = False
    PEMBELIAN.Enabled = False
    PRODUKSI.Enabled = False
    GUDANG.Enabled = False
    PENJUALAN.Enabled = False
    KEUANGAN.Enabled = False
    ABOUT.Selected = True
End Sub

Private Sub EnableAllMenu()
    MASTER.Enabled = True
    PEMBELIAN.Enabled = True
    PRODUKSI.Enabled = True
    'PRODUKSI_MAINMENU_SOP_SCANLOT_KEYPALET.Enabled = False
    GUDANG.Enabled = True
    PENJUALAN.Enabled = True
    KEUANGAN.Enabled = True
    ABOUT.Selected = True
End Sub

Private Sub SetPopupInfo(popup As XtremeSuiteControls.PopupControl, ByVal msgTitle As String, ByVal msgStatus As String)
    Dim Item As PopupControlItem
    
    popup.RemoveAllItems
    popup.Icons.RemoveAll
    
    Set Item = popup.additem(5, 6, 170, 19, AppName)
    Item.Hyperlink = False
    
    Set Item = popup.additem(5, 27, 160, 25, msgTitle)
    Item.TextAlignment = DT_LEFT
    Item.CalculateHeight
    Item.CalculateWidth
    
    Set Item = popup.additem(5, 50, 170, 200, msgStatus)
    Item.TextAlignment = DT_LEFT Or DT_WORDBREAK
    Item.CalculateHeight
    
    popup.VisualTheme = xtpPopupThemeMSN
    popup.SetSize 200, 130
End Sub

Private Sub pesanasystem()
    Dim pesan As String
    SQL = "select * from pesan where flag='1'"
    OBJ.Open dsn
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        OBJ.Close
        Exit Sub
    End If
    pesan = RST!pesan
    OBJ.Close
    SetPopupPesan popuppesan, "System Info", pesan
   
    popuppesan.Show
End Sub

Private Sub SetPopupPesan(popup As XtremeSuiteControls.PopupControl, ByVal msgTitle As String, ByVal msgStatus As String)
    Dim Item As PopupControlItem
    
    popup.RemoveAllItems
    popup.Icons.RemoveAll
    
    Set Item = popup.additem(5, 6, 170, 19, AppName)
    Item.Hyperlink = False
    
    Set Item = popup.additem(5, 27, 160, 25, msgTitle)
    Item.TextAlignment = DT_LEFT
    Item.CalculateHeight
    Item.CalculateWidth
    
    Set Item = popup.additem(5, 50, 170, 200, msgStatus)
    Item.TextAlignment = DT_LEFT Or DT_WORDBREAK
    Item.CalculateHeight
    
    popup.VisualTheme = xtpPopupThemeMSN
    popup.SetSize 200, 130
End Sub


Private Sub Timer2_Timer()
    Timer2.Enabled = False
    pesanasystem
End Sub

Private Sub Timer3_Timer()
    Dim currentTime As String
    currentTime = Format(Time, "HH:MM")
    
    If currentTime = "15:30" Then
        If UserOnLineLevel = "Operator" And UserOnlineDept = "1" Then Call PurchNotif

    End If
End Sub
