VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmkonvbrg 
   Caption         =   "Konversi Kemasan"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11325
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8745
   ScaleWidth      =   11325
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox added 
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
      Left            =   7380
      Picture         =   "frmkonvbrg.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   18
      Top             =   165
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtproduk 
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
      Height          =   315
      Left            =   2145
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   105
      Width           =   3060
   End
   Begin VB.TextBox txtkodeproduk 
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
      Height          =   315
      Left            =   1065
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   105
      Width           =   1050
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
      Left            =   6465
      Picture         =   "frmkonvbrg.frx":03B6
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   165
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
      Left            =   6720
      Picture         =   "frmkonvbrg.frx":0704
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   165
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
      Left            =   7080
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   165
      Visible         =   0   'False
      Width           =   255
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   5085
      Left            =   90
      TabIndex        =   0
      Top             =   3120
      Width           =   11190
      _Version        =   851970
      _ExtentX        =   19738
      _ExtentY        =   8969
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Color           =   8
      ItemCount       =   1
      Item(0).Caption =   "Konversi Karton"
      Item(0).ControlCount=   7
      Item(0).Control(0)=   "grid2"
      Item(0).Control(1)=   "txtnilai"
      Item(0).Control(2)=   "grid3"
      Item(0).Control(3)=   "grid4"
      Item(0).Control(4)=   "Label1"
      Item(0).Control(5)=   "Label2"
      Item(0).Control(6)=   "Label3"
      Begin TDBNumber6Ctl.TDBNumber txtnilai 
         Height          =   255
         Left            =   3960
         TabIndex        =   13
         Top             =   45
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   450
         Calculator      =   "frmkonvbrg.frx":09E6
         Caption         =   "frmkonvbrg.frx":0A06
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmkonvbrg.frx":0A72
         Keys            =   "frmkonvbrg.frx":0A90
         Spin            =   "frmkonvbrg.frx":0AD2
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   0
         BackColor       =   -2147483643
         BorderStyle     =   0
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   "."
         DisplayFormat   =   "###,###,###,##0;(###,###,###,##0);0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "###,###,###,##0;(###,###,###,##0)"
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid3 
         Height          =   1380
         Left            =   45
         TabIndex        =   14
         Top             =   2085
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   2434
         _Version        =   393216
         BackColor       =   -2147483628
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorFixed  =   -2147483642
         ForeColorFixed  =   -2147483633
         BackColorBkg    =   12632256
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
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid2 
         Height          =   1440
         Left            =   45
         TabIndex        =   12
         Top             =   360
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   2540
         _Version        =   393216
         BackColor       =   -2147483628
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorFixed  =   -2147483642
         ForeColorFixed  =   -2147483633
         BackColorBkg    =   12632256
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
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid4 
         Height          =   1245
         Left            =   45
         TabIndex        =   15
         Top             =   3765
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   2196
         _Version        =   393216
         BackColor       =   -2147483628
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorFixed  =   -2147483642
         ForeColorFixed  =   -2147483633
         BackColorBkg    =   12632256
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
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "PACKAGING"
         BeginProperty Font 
            Name            =   "News706 BT"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   8205
         TabIndex        =   19
         Top             =   -60
         Width           =   2880
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Konversi Etiket"
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
         Left            =   120
         TabIndex        =   17
         Top             =   3495
         Width           =   2025
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Konversi Kaleng"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   16
         Top             =   1815
         Width           =   2025
      End
   End
   Begin XtremeSuiteControls.PushButton cmdproduksi 
      Height          =   240
      Left            =   30
      TabIndex        =   6
      Top             =   120
      Width           =   990
      _Version        =   851970
      _ExtentX        =   1746
      _ExtentY        =   423
      _StockProps     =   79
      Caption         =   "PRODUK :"
      BackColor       =   -2147483644
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      TextAlignment   =   1
      Appearance      =   6
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   2595
      Left            =   105
      TabIndex        =   7
      ToolTipText     =   "Pilih item untuk konversi kemasan"
      Top             =   480
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4577
      _Version        =   393216
      BackColor       =   -2147483628
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   12632256
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
      _Band(0).Cols   =   2
   End
   Begin XtremeSuiteControls.PushButton btnClose 
      Height          =   420
      Left            =   9885
      TabIndex        =   8
      Top             =   8250
      Width           =   1290
      _Version        =   851970
      _ExtentX        =   2275
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "CLOSE"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnsave 
      Height          =   420
      Left            =   8520
      TabIndex        =   9
      Top             =   8235
      Width           =   1290
      _Version        =   851970
      _ExtentX        =   2275
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "SAVE"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnclear 
      Height          =   420
      Left            =   7140
      TabIndex        =   10
      Top             =   8250
      Width           =   1290
      _Version        =   851970
      _ExtentX        =   2275
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "CLEAR"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnview 
      Height          =   420
      Left            =   105
      TabIndex        =   11
      Top             =   8265
      Width           =   1845
      _Version        =   851970
      _ExtentX        =   3254
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "View List Konversi"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin Crystal.CrystalReport crystal 
      Left            =   2055
      Top             =   8280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   1665
      Left            =   7590
      Picture         =   "frmkonvbrg.frx":0AFA
      Stretch         =   -1  'True
      Top             =   1050
      Width           =   2385
   End
End
Attribute VB_Name = "frmkonvbrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private OBJ As New ADODB.Connection
Private RST As ADODB.Recordset
Private OBJ1 As New ADODB.Connection
Private RST1 As ADODB.Recordset

Private SQL As String
Private SQL1 As String
Private poscol As Integer
Private posrow As Integer
Private ongrid As String
Dim i, j, k As Integer

Private Sub btnClear_Click()
    txtkodeproduk = ""
    txtproduk = ""
    hapusgrid
    hapusgrid2
    hapusgrid3
    hapusgrid4
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnSave_Click()
    If txtkodeproduk = "" Then Exit Sub
    If grid2.TextMatrix(1, 1) = "" Then
        MsgBox "Item konversi belum terisi...", vbCritical, "Warning"
        Exit Sub
    End If
    
    
    OBJ.Open dsn
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
        'CEK KEMASAN DI GRID DAN DATABASE
        SQL = "Select * From list_konversibrg Where kodebarang = '" & grid2.TextMatrix(grid2.Row, 1) & "'"
        SQL = SQL + " and lineitem = '" & grid2.Row & "'"
        Set RST = New ADODB.Recordset
        RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
        
        If Not RST.EOF Then
            'UPDATE KONVERSI KEMASAN
            If grid2.TextMatrix(grid2.Row, 6) = "" Then
                RST!konv_1 = "0"
            Else
                RST!konv_1 = grid2.TextMatrix(grid2.Row, 6)
            End If
            RST!kodekemasan_1 = grid2.TextMatrix(grid2.Row, 7)
            RST!kemasan_1 = grid2.TextMatrix(grid2.Row, 8)
            RST.Update
            'UPDATE AM_ITEMDTL
            updatekonv

        ElseIf RST.EOF Then
                RST.AddNew
                RST!KodeBarang = grid2.TextMatrix(grid2.Row, 1)
                RST!namabarang = grid2.TextMatrix(grid2.Row, 2)
                RST!kodesatuan = grid2.TextMatrix(grid2.Row, 4)
                
                If grid2.TextMatrix(grid2.Row, 6) = "" Then
                    RST!konv_1 = "0"
                Else
                    RST!konv_1 = grid2.TextMatrix(grid2.Row, 6)
                End If
                RST!kodekemasan_1 = grid2.TextMatrix(grid2.Row, 7)
                RST!kemasan_1 = grid2.TextMatrix(grid2.Row, 8)
                RST!lineitem = grid2.Row
                RST.Update
        End If
        grid2.Row = grid2.Row + 1
    Loop
    OBJ.Close
    
    OBJ.Open dsn
    grid3.Row = 1
    Do While True
        If grid3.TextMatrix(grid3.Row, 1) = "" Then Exit Do
        
        SQL = "Select * From list_konversibrg Where kodebarang = '" & grid3.TextMatrix(grid3.Row, 1) & "'"
        SQL = SQL + " and lineitem='" & grid3.Row & "'"
        Set RST = New ADODB.Recordset
        RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
        
        If grid3.TextMatrix(grid3.Row, 7) = "" Then
            RST!konv_2 = "0"
            RST!kodekemasan_2 = ""
            RST!kemasan_2 = ""
            RST.Update
            GoTo pass1:
        End If
        If Not RST.EOF Then
            If grid3.TextMatrix(grid3.Row, 1) = "" Then Exit Do
            If grid3.TextMatrix(grid3.Row, 6) = "" Then
                RST!konv_2 = "0"
            Else
                RST!konv_2 = grid3.TextMatrix(grid3.Row, 6)
            End If
            RST!kodekemasan_2 = grid3.TextMatrix(grid3.Row, 7)
            RST!kemasan_2 = grid3.TextMatrix(grid3.Row, 8)
            RST.Update
        End If
pass1:
        grid3.Row = grid3.Row + 1
    Loop
    
    grid4.Row = 1
    Do While True
        If grid4.TextMatrix(grid4.Row, 1) = "" Then Exit Do
        SQL = "Select * From list_konversibrg Where kodebarang = '" & grid4.TextMatrix(grid4.Row, 1) & "'"
        SQL = SQL + " and lineitem='" & grid4.Row & "'"
        Set RST = New ADODB.Recordset
        RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
        
        If grid4.TextMatrix(grid4.Row, 7) = "" Then
            RST!konv_3 = "0"
            RST!kodekemasan_3 = ""
            RST!kemasan_3 = ""
            RST.Update
            GoTo pass2:
        End If
        
        If Not RST.EOF Then
            If grid4.TextMatrix(grid4.Row, 1) = "" Then Exit Do
            If grid4.TextMatrix(grid4.Row, 6) = "" Then
                RST!konv_3 = "0"
            Else
                RST!konv_3 = grid4.TextMatrix(grid4.Row, 6)
            End If
            RST!kodekemasan_3 = grid4.TextMatrix(grid4.Row, 7)
            RST!kemasan_3 = grid4.TextMatrix(grid4.Row, 8)
            RST.Update
        End If
pass2:
        grid4.Row = grid4.Row + 1
    Loop

    OBJ.Close
    MsgBox "Data berhasil disimpan", vbInformation, AppName
    btnClear_Click
End Sub
Private Sub updatekonv()
On Error Resume Next
    SQL = "Select * From am_itemdtl Where kodebarang = '" & grid3.TextMatrix(grid3.Row, 1) & "' "
    SQL = SQL + "and kodesatuan ='" & grid3.TextMatrix(grid3.Row, 4) & "'"
    
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
        
    If Not RST.EOF Then
        RST!konversi = Format(grid3.TextMatrix(grid3.Row, 6), "general number")
        RST.Update
    End If
End Sub

Private Sub btnview_Click()
    crystal.Reset
    crystal.WindowState = crptMaximized
    crystal.WindowShowCloseBtn = True
    crystal.WindowShowPrintSetupBtn = True
    crystal.WindowShowSearchBtn = True
    crystal.Connect = dsnreport
    crystal.ReportFileName = AppPath & "\reports\produksi\list_konversibrg.rpt"
    crystal.RetrieveDataFiles
    crystal.Action = 1
End Sub

Private Sub cmdproduksi_Click()
    If UserOnLineLevel = "creator" Then GoTo proses:
    If UserOnLineLevel = "Administrator" Then GoTo proses:
    If UserOnLineLevel <> "Supervisor" Then
        MsgBox "Anda tidak memiliki akses..! ", vbCritical, AppName
        Exit Sub
    End If
proses:
    namatabel = "produk"
    carisql1 = "select kode_produk,nama_produk from list_produk_master"
    frmsearch.Show vbModal
End Sub

Private Sub cmdproduksi_GotFocus()
    If hasil = "" Then Exit Sub
    txtkodeproduk = hasil
    txtproduk = hasil1
    hasil = ""
    hasil1 = ""
    carisql1 = ""
    findbrgjadi
End Sub

Private Sub findbrgjadi()
    SQL = "select a.kodebarang,a.namabarang,b.kode_satuan,c.namasatuan "
    SQL = SQL + "from am_itemmst a inner join list_produk_hasil b "
    SQL = SQL + "on a.kodebarang=b.kode_barang_jadi inner join am_unit c "
    SQL = SQL + "on b.kode_satuan=c.kodesatuan "
    SQL = SQL + "and b.kode_produk='" & txtkodeproduk & "' "
    
    OBJ.Open dsn
    Set RST = OBJ.Execute(SQL)
    hapusgrid
    hapusgrid2
    grid.Row = 1
    Do While Not RST.EOF
        grid.TextMatrix(grid.Row, 0) = grid.Row
        grid.TextMatrix(grid.Row, 1) = RST!KodeBarang
        grid.TextMatrix(grid.Row, 2) = RST!namabarang
        grid.TextMatrix(grid.Row, 3) = RST!kode_satuan
        grid.TextMatrix(grid.Row, 4) = RST!namasatuan
        grid.Rows = grid.Rows + 1
        grid.Row = grid.Row + 1
        RST.MoveNext
    Loop
    OBJ.Close
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

        grid.Col = 0
        Set grid.CellPicture = blank
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
    SetGrid
    For j = 0 To grid.Cols - 1
        grid.Col = j
        grid.CellBackColor = &HFFFFFF
    Next
End Sub
Private Sub hapusgrid2()
    grid2.Row = 1
    Do While True
        If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Do
        grid2.TextMatrix(grid2.Row, 0) = ""
        grid2.TextMatrix(grid2.Row, 1) = ""
        grid2.TextMatrix(grid2.Row, 2) = ""
        grid2.TextMatrix(grid2.Row, 3) = ""
        grid2.TextMatrix(grid2.Row, 4) = ""
        grid2.TextMatrix(grid2.Row, 5) = ""
        grid2.TextMatrix(grid2.Row, 6) = ""
        grid2.TextMatrix(grid2.Row, 7) = ""
        grid2.TextMatrix(grid2.Row, 8) = ""
        
        grid2.Col = 0
        Set grid2.CellPicture = blank
        grid2.Row = grid2.Row + 1
    Loop
    grid2.Rows = 2
    SetGrid2
End Sub
Private Sub hapusgrid3()
    grid3.Row = 1
    Do While True
        If grid3.TextMatrix(grid3.Row, 1) = "" Then Exit Do
        grid3.TextMatrix(grid3.Row, 0) = ""
        grid3.TextMatrix(grid3.Row, 1) = ""
        grid3.TextMatrix(grid3.Row, 2) = ""
        grid3.TextMatrix(grid3.Row, 3) = ""
        grid3.TextMatrix(grid3.Row, 4) = ""
        grid3.TextMatrix(grid3.Row, 5) = ""
        grid3.TextMatrix(grid3.Row, 6) = ""
        grid3.TextMatrix(grid3.Row, 7) = ""
        grid3.TextMatrix(grid3.Row, 8) = ""
        
        grid3.Col = 0
        Set grid3.CellPicture = blank
        grid3.Row = grid3.Row + 1
    Loop
    grid3.Rows = 2
    SetGrid3
End Sub
Private Sub hapusgrid4()
    grid4.Row = 1
    Do While True
        If grid4.TextMatrix(grid4.Row, 1) = "" Then Exit Do
        grid4.TextMatrix(grid4.Row, 0) = ""
        grid4.TextMatrix(grid4.Row, 1) = ""
        grid4.TextMatrix(grid4.Row, 2) = ""
        grid4.TextMatrix(grid4.Row, 3) = ""
        grid4.TextMatrix(grid4.Row, 4) = ""
        grid4.TextMatrix(grid4.Row, 5) = ""
        grid4.TextMatrix(grid4.Row, 6) = ""
        grid4.TextMatrix(grid4.Row, 7) = ""
        grid4.TextMatrix(grid4.Row, 8) = ""
        
        grid4.Col = 0
        Set grid4.CellPicture = blank
        grid4.Row = grid4.Row + 1
    Loop
    grid4.Rows = 2
    SetGrid4
End Sub

Private Sub SetGrid()
    With grid
        .ColWidth(0) = 300
        .ColWidth(1) = 1200
        .ColWidth(2) = 3000
        .ColWidth(3) = 0
        .ColWidth(4) = 1000
    End With
End Sub
Private Sub SetGrid2()
    With grid2
        .ColWidth(0) = 300
        .ColWidth(1) = 1200
        .ColWidth(2) = 3000
        .ColWidth(3) = 400
        .ColWidth(4) = 0
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        .ColWidth(7) = 1000
        .ColWidth(8) = 3000
    End With
End Sub
Private Sub SetGrid3()
    With grid3
        .ColWidth(0) = 300
        .ColWidth(1) = 1200
        .ColWidth(2) = 3000
        .ColWidth(3) = 400
        .ColWidth(4) = 0
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        .ColWidth(7) = 1000
        .ColWidth(8) = 3000
    End With
End Sub
Private Sub SetGrid4()
    With grid4
        .ColWidth(0) = 300
        .ColWidth(1) = 1200
        .ColWidth(2) = 3000
        .ColWidth(3) = 400
        .ColWidth(4) = 0
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        .ColWidth(7) = 1000
        .ColWidth(8) = 3000
    End With
End Sub

Private Sub initGrid()
    poscol = grid2.Col
    posrow = grid2.Row
    With grid
        .Cols = 5
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "KODE"
        .TextMatrix(0, 2) = "NAMA BARANG"
        .TextMatrix(0, 3) = "KODE"
        .TextMatrix(0, 4) = "SATUAN"
    End With
    With grid2
        .Cols = 9
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "KODE"
        .TextMatrix(0, 2) = "NAMA BARANG"
        .TextMatrix(0, 3) = ""
        .TextMatrix(0, 4) = "KODE"
        .TextMatrix(0, 5) = "SATUAN"
        .TextMatrix(0, 6) = "KONVERSI"
        .TextMatrix(0, 7) = "KODE"
        .TextMatrix(0, 8) = "KEMASAN"

        For i = 6 To 8
            grid2.Col = i
            grid2.CellBackColor = &HE0E0E0
        Next
    End With
    With grid3
        .Cols = 9
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "KODE"
        .TextMatrix(0, 2) = "NAMA BARANG"
        .TextMatrix(0, 3) = ""
        .TextMatrix(0, 4) = "KODE"
        .TextMatrix(0, 5) = "SATUAN"
        .TextMatrix(0, 6) = "KONVERSI"
        .TextMatrix(0, 7) = "KODE"
        .TextMatrix(0, 8) = "KEMASAN"

        For i = 6 To 8
            grid3.Col = i
            grid3.CellBackColor = &HE0E0E0
        Next
    End With
    With grid4
        .Cols = 9
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "KODE"
        .TextMatrix(0, 2) = "NAMA BARANG"
        .TextMatrix(0, 3) = ""
        .TextMatrix(0, 4) = "KODE"
        .TextMatrix(0, 5) = "SATUAN"
        .TextMatrix(0, 6) = "KONVERSI"
        .TextMatrix(0, 7) = "KODE"
        .TextMatrix(0, 8) = "KEMASAN"

        For i = 6 To 8
            grid4.Col = i
            grid4.CellBackColor = &HE0E0E0
        Next
    End With
End Sub

Private Sub Form_Load()
    initGrid
    SetGrid
    SetGrid2
    SetGrid3
    SetGrid4
End Sub

Private Function SetAlternatingGrid(ByVal i As Integer)
    Dim j, k As Integer
    j = 0
    k = 0
    For k = 1 To grid.Rows - 1
        For j = 0 To grid.Cols - 1
        grid.Col = j
        grid.CellBackColor = &HFFFFFF
        Next
    Next k
End Function

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    hapusgrid2
    hapusgrid3
    hapusgrid4
    
    grid2.Col = 1
    grid2.Row = 1
    'PERIKSA GRID2
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
        With grid2
            If .TextMatrix(.Row, 1) = "" Then GoTo isigrid2:
            If .TextMatrix(.Row, 1) = grid.TextMatrix(grid.Row, 1) Then Exit Sub
                .Rows = .Rows + 1
        End With
    Loop
    
    
isigrid2:
    caribrgjadi
    grid2.Row = grid2.Rows - 1

    For i = 6 To 8
        grid2.Col = i
        grid2.CellBackColor = &HE0E0E0
    Next
'------------------------------------------------GRID 3
    grid3.Col = 1
    grid3.Row = 1
    'PERIKSA GRID2
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
        With grid3
            If .TextMatrix(.Row, 1) = "" Then GoTo isigrid3:
            If .TextMatrix(.Row, 1) = grid.TextMatrix(grid.Row, 1) Then Exit Sub
                .Rows = .Rows + 1
                .Col = 0
                .Row = .Rows - 1
                Set .CellPicture = added
        End With
    Loop
    
isigrid3:
    caribrgjadi2
    For i = 6 To 8
        grid3.Col = i
        grid3.CellBackColor = &HE0E0E0
    Next
'------------------------------------------------GRID 4
    grid4.Col = 1
    grid4.Row = 1
    'PERIKSA grid4
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
        With grid4
            If .TextMatrix(.Row, 1) = "" Then GoTo isigrid4:
            If .TextMatrix(.Row, 1) = grid.TextMatrix(grid.Row, 1) Then Exit Sub
                .Rows = .Rows + 1
                .Col = 0
                .Row = .Rows - 1
                Set .CellPicture = added
        End With
    Loop
    
isigrid4:
    caribrgjadi3
    For i = 6 To 8
        grid4.Col = i
        grid4.CellBackColor = &HE0E0E0
    Next
End Sub

Private Sub caribrgjadi()
    OBJ.Open dsn
    SQL = "select * from am_itemdtl where kodebarang='" & grid.TextMatrix(grid.Row, 1) & "' and kodesatuan ='" & grid.TextMatrix(grid.Row, 3) & "'"
    Set RST = OBJ.Execute(SQL)
    grid2.TextMatrix(grid2.Row, 6) = RST!konversi
    
    SQL = "Select * From list_konversibrg Where kodebarang='" & grid.TextMatrix(grid.Row, 1) & "'"
    Set RST = OBJ.Execute(SQL)
    
    If RST.EOF Then
        With grid2
            grid2.TextMatrix(grid2.Row, 1) = grid.TextMatrix(grid.Row, 1)
            grid2.TextMatrix(grid2.Row, 2) = grid.TextMatrix(grid.Row, 2)
            grid2.TextMatrix(grid2.Row, 3) = "1"
            grid2.TextMatrix(grid2.Row, 4) = grid.TextMatrix(grid.Row, 3)
            grid2.TextMatrix(grid2.Row, 5) = grid.TextMatrix(grid.Row, 4)
            grid2.TextMatrix(grid2.Row, 6) = ""
            grid2.TextMatrix(grid2.Row, 7) = ""
            grid2.TextMatrix(grid2.Row, 8) = ""
            .Rows = .Rows + 1
            .Row = .Row + 1
            .TextMatrix(.Row, 1) = .TextMatrix(1, 1)
            .TextMatrix(.Row, 2) = .TextMatrix(1, 2)
            .TextMatrix(.Row, 3) = .TextMatrix(1, 3)
            .TextMatrix(.Row, 4) = .TextMatrix(1, 4)
            .TextMatrix(.Row, 5) = .TextMatrix(1, 5)
            .Rows = .Rows + 1
        End With
    End If
    
    Do While Not RST.EOF
        grid2.Col = 0
        Set grid2.CellPicture = uncheck
        grid2.TextMatrix(grid2.Row, 1) = grid.TextMatrix(grid.Row, 1)
        grid2.TextMatrix(grid2.Row, 2) = grid.TextMatrix(grid.Row, 2)
        grid2.TextMatrix(grid2.Row, 3) = "1"
        grid2.TextMatrix(grid2.Row, 4) = grid.TextMatrix(grid.Row, 3)
        grid2.TextMatrix(grid2.Row, 5) = grid.TextMatrix(grid.Row, 4)
        grid2.TextMatrix(grid2.Row, 6) = RST!konv_1
        grid2.TextMatrix(grid2.Row, 7) = RST!kodekemasan_1
        grid2.TextMatrix(grid2.Row, 8) = RST!kemasan_1
        grid2.Rows = grid2.Rows + 1
        grid2.Row = grid2.Row + 1
        RST.MoveNext
    Loop

    
    OBJ.Close
End Sub
Private Sub caribrgjadi2()
    OBJ.Open dsn
    
    SQL = "Select * From list_konversibrg Where kodebarang='" & grid.TextMatrix(grid.Row, 1) & "'"
    Set RST = OBJ.Execute(SQL)
    
    If RST.EOF Then
        grid3.Col = 0
        Set grid3.CellPicture = uncheck
        grid3.TextMatrix(grid3.Row, 1) = grid.TextMatrix(grid.Row, 1)
        grid3.TextMatrix(grid3.Row, 2) = grid.TextMatrix(grid.Row, 2)
        grid3.TextMatrix(grid3.Row, 3) = "1"
        grid3.TextMatrix(grid3.Row, 4) = grid.TextMatrix(grid.Row, 3)
        grid3.TextMatrix(grid3.Row, 5) = grid.TextMatrix(grid.Row, 4)
        grid3.Rows = grid3.Rows + 1
        grid3.Row = grid3.Row + 1
        grid3.TextMatrix(grid3.Row, 6) = ""
        grid3.TextMatrix(grid3.Row, 7) = ""
        grid3.TextMatrix(grid3.Row, 8) = ""
        grid3.Col = 0
        grid3.Row = grid3.Rows - 1
    End If
    
    Do While Not RST.EOF
        If IsNull(RST!konv_2) Or RST!konv_2 = "0" And RST!lineitem = "2" Then Exit Do
        If IsNull(RST!konv_2) Or RST!konv_2 = "" And RST!kodekemasan_2 = "" Or IsNull(RST!kodekemasan_2) Then RST.MoveNext
        grid3.Col = 0
        Set grid3.CellPicture = uncheck
        grid3.TextMatrix(grid3.Row, 1) = grid.TextMatrix(grid.Row, 1)
        grid3.TextMatrix(grid3.Row, 2) = grid.TextMatrix(grid.Row, 2)
        grid3.TextMatrix(grid3.Row, 3) = "1"
        grid3.TextMatrix(grid3.Row, 4) = grid.TextMatrix(grid.Row, 3)
        grid3.TextMatrix(grid3.Row, 5) = grid.TextMatrix(grid.Row, 4)
        grid3.TextMatrix(grid3.Row, 6) = RST!konv_2
        grid3.TextMatrix(grid3.Row, 7) = RST!kodekemasan_2
        grid3.TextMatrix(grid3.Row, 8) = RST!kemasan_2
        grid3.Rows = grid3.Rows + 1
        grid3.Row = grid3.Row + 1
        RST.MoveNext
    Loop
    Set grid3.CellPicture = added
    OBJ.Close
End Sub
Private Sub caribrgjadi3()
    OBJ.Open dsn
    
    SQL = "Select * From list_konversibrg Where kodebarang='" & grid.TextMatrix(grid.Row, 1) & "'"
    Set RST = OBJ.Execute(SQL)
    
    If RST.EOF Then
        grid4.Col = 0
        Set grid4.CellPicture = uncheck
        grid4.TextMatrix(grid4.Row, 1) = grid.TextMatrix(grid.Row, 1)
        grid4.TextMatrix(grid4.Row, 2) = grid.TextMatrix(grid.Row, 2)
        grid4.TextMatrix(grid4.Row, 3) = "1"
        grid4.TextMatrix(grid4.Row, 4) = grid.TextMatrix(grid.Row, 3)
        grid4.TextMatrix(grid4.Row, 5) = grid.TextMatrix(grid.Row, 4)
        grid4.Rows = grid4.Rows + 1
        grid4.Row = grid4.Row + 1
        grid4.TextMatrix(grid4.Row, 6) = ""
        grid4.TextMatrix(grid4.Row, 7) = ""
        grid4.TextMatrix(grid4.Row, 8) = ""
        grid4.Col = 0
        grid4.Row = grid4.Rows - 1
    End If
    
    Do While Not RST.EOF
        If IsNull(RST!konv_3) Or RST!konv_3 = "0" And RST!lineitem = "2" Then Exit Do
        If IsNull(RST!konv_3) Or RST!konv_3 = "" And RST!kodekemasan_3 = "" Or IsNull(RST!kodekemasan_3) Then RST.MoveNext
        grid4.Col = 0
        Set grid4.CellPicture = uncheck
        grid4.TextMatrix(grid4.Row, 1) = grid.TextMatrix(grid.Row, 1)
        grid4.TextMatrix(grid4.Row, 2) = grid.TextMatrix(grid.Row, 2)
        grid4.TextMatrix(grid4.Row, 3) = "1"
        grid4.TextMatrix(grid4.Row, 4) = grid.TextMatrix(grid.Row, 3)
        grid4.TextMatrix(grid4.Row, 5) = grid.TextMatrix(grid.Row, 4)
        grid4.TextMatrix(grid4.Row, 6) = RST!konv_3
        grid4.TextMatrix(grid4.Row, 7) = RST!kodekemasan_3
        grid4.TextMatrix(grid4.Row, 8) = RST!kemasan_3
        grid4.Rows = grid4.Rows + 1
        grid4.Row = grid4.Row + 1
        RST.MoveNext
    Loop
    grid4.Col = 0
    Set grid4.CellPicture = added
    OBJ.Close
End Sub

Private Sub grid2_Click()
    If grid2.MouseRow = 0 Then Exit Sub
    If grid2.TextMatrix(1, 1) = "" Then Exit Sub
    poscol = grid2.Col
    posrow = grid2.Row

    Select Case grid2.Col
        Case 0:
            'If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Sub
            If grid2.CellPicture = uncheck Then
                Set grid2.CellPicture = check
                If MsgBox("Delete this row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                    Set grid2.CellPicture = uncheck
                    deletegrid2
                    hapusrow2
                    Set grid2.CellPicture = uncheck
                    With grid2
                        .TextMatrix(.Row, 1) = .TextMatrix(1, 1)
                        .TextMatrix(.Row, 2) = .TextMatrix(1, 2)
                        .TextMatrix(.Row, 3) = .TextMatrix(1, 3)
                        .TextMatrix(.Row, 4) = .TextMatrix(1, 4)
                        .TextMatrix(.Row, 5) = .TextMatrix(1, 5)
                        .Rows = .Rows + 1
                    End With
                    Exit Sub
                End If
                Set grid2.CellPicture = uncheck
            ElseIf grid2.CellPicture = added Then
                If grid2.TextMatrix(1, 7) = "" Then Exit Sub
                Set grid2.CellPicture = uncheck
                With grid2
                    .TextMatrix(.Row, 1) = .TextMatrix(1, 1)
                    .TextMatrix(.Row, 2) = .TextMatrix(1, 2)
                    .TextMatrix(.Row, 3) = .TextMatrix(1, 3)
                    .TextMatrix(.Row, 4) = .TextMatrix(1, 4)
                    .TextMatrix(.Row, 5) = .TextMatrix(1, 5)
                    .Rows = .Rows + 1
                End With
            End If
        Case 6:
            ongrid = "grid2"
            If grid2.TextMatrix(grid2.Row, 1) = "" Then Exit Sub
            txtnilai.Width = grid2.ColWidth(grid2.Col) - 40
            txtnilai = grid2.TextMatrix(grid2.Row, grid2.Col)
            txtnilai.Left = grid2.Left + grid2.CellLeft
            txtnilai.Top = grid2.Top + grid2.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
        Case 7:
            carisql1 = "select kodebarang, namabarang from am_apitemmst"
            namatabel = "Bahan Baku"
            frmsearch.Show vbModal
    End Select
End Sub

Private Sub grid2_GotFocus()
    If hasil = "" Then Exit Sub
    Select Case grid2.Col
        Case 7:
            grid2.TextMatrix(grid2.Row, 7) = hasil
            grid2.TextMatrix(grid2.Row, 8) = hasil1
            hasil = ""
            hasil1 = ""
            hasil2 = ""
            hasil3 = ""
            namatabel = ""
            carisql1 = ""
    End Select
End Sub

Private Sub grid3_Click()
    If grid3.MouseRow = 0 Then Exit Sub
    If grid3.TextMatrix(1, 1) = "" Then Exit Sub
    poscol = grid3.Col
    posrow = grid3.Row
    
    Select Case grid3.Col
        Case 0:
            'If grid3.TextMatrix(grid3.Row, 1) = "" Then Exit Sub
            If grid3.CellPicture = uncheck Then
                Set grid3.CellPicture = check
                If MsgBox("Delete this row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                    Set grid3.CellPicture = uncheck
                    deletegrid3
                    hapusrow3
                    Set grid3.CellPicture = uncheck
                    With grid3
                        .TextMatrix(.Row, 1) = .TextMatrix(1, 1)
                        .TextMatrix(.Row, 2) = .TextMatrix(1, 2)
                        .TextMatrix(.Row, 3) = .TextMatrix(1, 3)
                        .TextMatrix(.Row, 4) = .TextMatrix(1, 4)
                        .TextMatrix(.Row, 5) = .TextMatrix(1, 5)
                        .Rows = .Rows + 1
                    End With
                    Exit Sub
                End If
                Set grid3.CellPicture = uncheck
            ElseIf grid3.CellPicture = added Then
                If grid3.TextMatrix(1, 7) = "" Then Exit Sub
                Set grid3.CellPicture = uncheck
                With grid3
                    .TextMatrix(.Row, 1) = .TextMatrix(1, 1)
                    .TextMatrix(.Row, 2) = .TextMatrix(1, 2)
                    .TextMatrix(.Row, 3) = .TextMatrix(1, 3)
                    .TextMatrix(.Row, 4) = .TextMatrix(1, 4)
                    .TextMatrix(.Row, 5) = .TextMatrix(1, 5)
                    .Rows = .Rows + 1
                End With
            End If
        Case 6:
            ongrid = "grid3"
            If grid3.TextMatrix(grid3.Row, 1) = "" Then Exit Sub
            txtnilai.Width = grid3.ColWidth(grid3.Col) - 40
            txtnilai = grid3.TextMatrix(grid3.Row, grid3.Col)
            txtnilai.Left = grid3.Left + grid3.CellLeft
            txtnilai.Top = grid3.Top + grid3.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
        Case 7:
            carisql1 = "select kodebarang, namabarang from am_apitemmst"
            namatabel = "Bahan Baku"
            frmsearch.Show vbModal
    End Select
End Sub

Private Sub grid3_GotFocus()
    If hasil = "" Then Exit Sub
    Select Case grid3.Col
        Case 7:
            grid3.TextMatrix(grid3.Row, 7) = hasil
            grid3.TextMatrix(grid3.Row, 8) = hasil1
            hasil = ""
            hasil1 = ""
            hasil2 = ""
            hasil3 = ""
            namatabel = ""
            carisql1 = ""
    End Select
End Sub

Private Sub grid4_Click()
    If grid4.MouseRow = 0 Then Exit Sub
    If grid4.TextMatrix(1, 1) = "" Then Exit Sub
    poscol = grid4.Col
    posrow = grid4.Row
    
    Select Case grid4.Col
        Case 0:
            'If grid4.TextMatrix(grid4.Row, 1) = "" Then Exit Sub
            If grid4.CellPicture = uncheck Then
                Set grid4.CellPicture = check
                If MsgBox("Delete this row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                    Set grid4.CellPicture = uncheck
                    deletegrid4
                    hapusrow4
                    Set grid4.CellPicture = uncheck
                    With grid4
                        .TextMatrix(.Row, 1) = .TextMatrix(1, 1)
                        .TextMatrix(.Row, 2) = .TextMatrix(1, 2)
                        .TextMatrix(.Row, 3) = .TextMatrix(1, 3)
                        .TextMatrix(.Row, 4) = .TextMatrix(1, 4)
                        .TextMatrix(.Row, 5) = .TextMatrix(1, 5)
                        .Rows = .Rows + 1
                    End With
                    Exit Sub
                End If
                Set grid4.CellPicture = uncheck
            ElseIf grid4.CellPicture = added Then
                If grid4.TextMatrix(1, 7) = "" Then Exit Sub
                Set grid4.CellPicture = uncheck
                With grid4
                    .TextMatrix(.Row, 1) = .TextMatrix(1, 1)
                    .TextMatrix(.Row, 2) = .TextMatrix(1, 2)
                    .TextMatrix(.Row, 3) = .TextMatrix(1, 3)
                    .TextMatrix(.Row, 4) = .TextMatrix(1, 4)
                    .TextMatrix(.Row, 5) = .TextMatrix(1, 5)
                    .Rows = .Rows + 1
                End With
            End If
        Case 6:
            ongrid = "grid4"
            If grid4.TextMatrix(grid4.Row, 1) = "" Then Exit Sub
            txtnilai.Width = grid4.ColWidth(grid4.Col) - 40
            txtnilai = grid4.TextMatrix(grid4.Row, grid4.Col)
            txtnilai.Left = grid4.Left + grid4.CellLeft
            txtnilai.Top = grid4.Top + grid4.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
        Case 7:
            carisql1 = "select kodebarang, namabarang from am_apitemmst"
            namatabel = "Bahan Baku"
            frmsearch.Show vbModal
    End Select
End Sub

Private Sub grid4_GotFocus()
    If hasil = "" Then Exit Sub
    Select Case grid4.Col
        Case 7:
            grid4.TextMatrix(grid4.Row, 7) = hasil
            grid4.TextMatrix(grid4.Row, 8) = hasil1
            hasil = ""
            hasil1 = ""
            hasil2 = ""
            hasil3 = ""
            namatabel = ""
            carisql1 = ""
    End Select
End Sub
Private Sub deletegrid2()
    OBJ.Open dsn
    SQL = "Update list_konversibrg set konv_1='0',kodekemasan_1='',kemasan_1=''"
    SQL = SQL + " Where kodebarang='" & grid2.TextMatrix(grid2.Row, 1) & "'"
    SQL = SQL + " And kodekemasan_1='" & grid2.TextMatrix(grid2.Row, 7) & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
End Sub

Private Sub deletegrid3()
    OBJ.Open dsn
    SQL = "Update list_konversibrg set konv_2='0',kodekemasan_2='',kemasan_2=''"
    SQL = SQL + " Where kodebarang='" & grid3.TextMatrix(grid3.Row, 1) & "'"
    SQL = SQL + " And kodekemasan_2='" & grid3.TextMatrix(grid3.Row, 7) & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
End Sub
Private Sub deletegrid4()
    OBJ.Open dsn
    SQL = "Update list_konversibrg set konv_3='0',kodekemasan_3='',kemasan_3=''"
    SQL = SQL + " Where kodebarang='" & grid4.TextMatrix(grid4.Row, 1) & "'"
    SQL = SQL + " And kodekemasan_3='" & grid4.TextMatrix(grid4.Row, 7) & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If ongrid = "grid2" Then
            grid2.TextMatrix(grid2.Row, 6) = txtnilai.text
            grid2.SetFocus
        ElseIf ongrid = "grid3" Then
            grid3.TextMatrix(grid3.Row, 6) = txtnilai.text
            grid3.SetFocus
        ElseIf ongrid = "grid4" Then
            grid4.TextMatrix(grid4.Row, 6) = txtnilai.text
            grid4.SetFocus
        End If
        
    End If
End Sub

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
End Sub

Private Sub hapusrow2()
    grid2.TextMatrix(grid2.Row, 1) = ""
    grid2.TextMatrix(grid2.Row, 2) = ""
    grid2.TextMatrix(grid2.Row, 3) = ""
    grid2.TextMatrix(grid2.Row, 4) = ""
    grid2.TextMatrix(grid2.Row, 5) = ""
    grid2.TextMatrix(grid2.Row, 6) = ""
    grid2.TextMatrix(grid2.Row, 7) = ""
    grid2.TextMatrix(grid2.Row, 8) = ""
    
    Do While True
        If grid2.TextMatrix(grid2.Row + 1, 1) = "" Then
            grid2.TextMatrix(grid2.Row, 1) = ""
            grid2.TextMatrix(grid2.Row, 2) = ""
            grid2.TextMatrix(grid2.Row, 3) = ""
            grid2.TextMatrix(grid2.Row, 4) = ""
            grid2.TextMatrix(grid2.Row, 5) = ""
            grid2.TextMatrix(grid2.Row, 6) = ""
            grid2.TextMatrix(grid2.Row, 7) = ""
            grid2.TextMatrix(grid2.Row, 8) = ""
            Exit Do
        End If
        grid2.TextMatrix(grid2.Row, 1) = grid2.TextMatrix(grid2.Row + 1, 1)
        grid2.TextMatrix(grid2.Row, 2) = grid2.TextMatrix(grid2.Row + 1, 2)
        grid2.TextMatrix(grid2.Row, 3) = grid2.TextMatrix(grid2.Row + 1, 3)
        grid2.TextMatrix(grid2.Row, 4) = grid2.TextMatrix(grid2.Row + 1, 4)
        grid2.TextMatrix(grid2.Row, 5) = grid2.TextMatrix(grid2.Row + 1, 5)
        grid2.TextMatrix(grid2.Row, 6) = grid2.TextMatrix(grid2.Row + 1, 6)
        grid2.TextMatrix(grid2.Row, 7) = grid2.TextMatrix(grid2.Row + 1, 7)
        grid2.TextMatrix(grid2.Row, 8) = grid2.TextMatrix(grid2.Row + 1, 8)
        grid2.Row = grid2.Row + 1
    Loop
    grid2.Rows = grid2.Rows - 1
    grid2.Col = 0
    Set grid2.CellPicture = added
End Sub

Private Sub hapusrow3()
    grid3.TextMatrix(grid3.Row, 1) = ""
    grid3.TextMatrix(grid3.Row, 2) = ""
    grid3.TextMatrix(grid3.Row, 3) = ""
    grid3.TextMatrix(grid3.Row, 4) = ""
    grid3.TextMatrix(grid3.Row, 5) = ""
    grid3.TextMatrix(grid3.Row, 6) = ""
    grid3.TextMatrix(grid3.Row, 7) = ""
    grid3.TextMatrix(grid3.Row, 8) = ""
    
    Do While True
        If grid3.TextMatrix(grid3.Row + 1, 1) = "" Then
            grid3.TextMatrix(grid3.Row, 1) = ""
            grid3.TextMatrix(grid3.Row, 2) = ""
            grid3.TextMatrix(grid3.Row, 3) = ""
            grid3.TextMatrix(grid3.Row, 4) = ""
            grid3.TextMatrix(grid3.Row, 5) = ""
            grid3.TextMatrix(grid3.Row, 6) = ""
            grid3.TextMatrix(grid3.Row, 7) = ""
            grid3.TextMatrix(grid3.Row, 8) = ""
            Exit Do
        End If
        grid3.TextMatrix(grid3.Row, 1) = grid3.TextMatrix(grid3.Row + 1, 1)
        grid3.TextMatrix(grid3.Row, 2) = grid3.TextMatrix(grid3.Row + 1, 2)
        grid3.TextMatrix(grid3.Row, 3) = grid3.TextMatrix(grid3.Row + 1, 3)
        grid3.TextMatrix(grid3.Row, 4) = grid3.TextMatrix(grid3.Row + 1, 4)
        grid3.TextMatrix(grid3.Row, 5) = grid3.TextMatrix(grid3.Row + 1, 5)
        grid3.TextMatrix(grid3.Row, 6) = grid3.TextMatrix(grid3.Row + 1, 6)
        grid3.TextMatrix(grid3.Row, 7) = grid3.TextMatrix(grid3.Row + 1, 7)
        grid3.TextMatrix(grid3.Row, 8) = grid3.TextMatrix(grid3.Row + 1, 8)
        grid3.Row = grid3.Row + 1
    Loop
    grid3.Rows = grid3.Rows - 1
    grid3.Col = 0
    Set grid3.CellPicture = added
End Sub
Private Sub hapusrow4()
    grid4.TextMatrix(grid4.Row, 1) = ""
    grid4.TextMatrix(grid4.Row, 2) = ""
    grid4.TextMatrix(grid4.Row, 3) = ""
    grid4.TextMatrix(grid4.Row, 4) = ""
    grid4.TextMatrix(grid4.Row, 5) = ""
    grid4.TextMatrix(grid4.Row, 6) = ""
    grid4.TextMatrix(grid4.Row, 7) = ""
    grid4.TextMatrix(grid4.Row, 8) = ""
    
    Do While True
        If grid4.TextMatrix(grid4.Row + 1, 1) = "" Then
            grid4.TextMatrix(grid4.Row, 1) = ""
            grid4.TextMatrix(grid4.Row, 2) = ""
            grid4.TextMatrix(grid4.Row, 3) = ""
            grid4.TextMatrix(grid4.Row, 4) = ""
            grid4.TextMatrix(grid4.Row, 5) = ""
            grid4.TextMatrix(grid4.Row, 6) = ""
            grid4.TextMatrix(grid4.Row, 7) = ""
            grid4.TextMatrix(grid4.Row, 8) = ""
            Exit Do
        End If
        grid4.TextMatrix(grid4.Row, 1) = grid4.TextMatrix(grid4.Row + 1, 1)
        grid4.TextMatrix(grid4.Row, 2) = grid4.TextMatrix(grid4.Row + 1, 2)
        grid4.TextMatrix(grid4.Row, 3) = grid4.TextMatrix(grid4.Row + 1, 3)
        grid4.TextMatrix(grid4.Row, 4) = grid4.TextMatrix(grid4.Row + 1, 4)
        grid4.TextMatrix(grid4.Row, 5) = grid4.TextMatrix(grid4.Row + 1, 5)
        grid4.TextMatrix(grid4.Row, 6) = grid4.TextMatrix(grid4.Row + 1, 6)
        grid4.TextMatrix(grid4.Row, 7) = grid4.TextMatrix(grid4.Row + 1, 7)
        grid4.TextMatrix(grid4.Row, 8) = grid4.TextMatrix(grid4.Row + 1, 8)
        grid4.Row = grid4.Row + 1
    Loop
    grid4.Rows = grid4.Rows - 1
    grid4.Col = 0
    Set grid4.CellPicture = added
End Sub
