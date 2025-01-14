VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmterimawip 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Terima Barang WIP (Scan)"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4485
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   4485
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtkdsatuan 
      Enabled         =   0   'False
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
      Left            =   120
      TabIndex        =   23
      Top             =   4800
      Width           =   1695
   End
   Begin VB.TextBox txtnolot 
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   4320
      Width           =   3135
   End
   Begin VB.TextBox txtnobpb 
      Enabled         =   0   'False
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
      Left            =   1200
      TabIndex        =   13
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox txtkode 
      Enabled         =   0   'False
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
      Left            =   1200
      TabIndex        =   10
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtnamabrg 
      Enabled         =   0   'False
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
      Left            =   1200
      TabIndex        =   7
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox txtsatuan 
      Enabled         =   0   'False
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
      TabIndex        =   5
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox txtqty 
      Enabled         =   0   'False
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
      Left            =   1200
      TabIndex        =   3
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtpalet 
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
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   3255
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5741
      _Version        =   393216
      Cols            =   5
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
      _Band(0).Cols   =   5
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Top             =   3000
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
      MICON           =   "frmterimawip.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdconfirm 
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   3000
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
      MICON           =   "frmterimawip.frx":031A
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
      Left            =   1200
      TabIndex        =   12
      Top             =   1560
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
      Format          =   146472961
      CurrentDate     =   42039
   End
   Begin Chameleon.chameleonButton cmdclear 
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   3000
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
      MICON           =   "frmterimawip.frx":0634
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdRepair 
      Height          =   375
      Left            =   1080
      TabIndex        =   16
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Repair Data"
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
      MICON           =   "frmterimawip.frx":094E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtkg 
      Height          =   225
      Left            =   120
      TabIndex        =   18
      Top             =   3600
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   397
      Calculator      =   "frmterimawip.frx":0C68
      Caption         =   "frmterimawip.frx":0C88
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmterimawip.frx":0CF4
      Keys            =   "frmterimawip.frx":0D12
      Spin            =   "frmterimawip.frx":0D54
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##,###,##0.00;(##,###,##0.00);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##,###,##0.00;(##,###,##0.00)"
      HighlightText   =   0
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
      ReadOnly        =   1
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtkgperpalet 
      Height          =   225
      Left            =   120
      TabIndex        =   19
      Top             =   3840
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   397
      Calculator      =   "frmterimawip.frx":0D7C
      Caption         =   "frmterimawip.frx":0D9C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmterimawip.frx":0E08
      Keys            =   "frmterimawip.frx":0E26
      Spin            =   "frmterimawip.frx":0E68
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##,###,##0.00;(##,###,##0.00);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##,###,##0.00;(##,###,##0.00)"
      HighlightText   =   0
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
      ReadOnly        =   1
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txthppperkg 
      Height          =   225
      Left            =   120
      TabIndex        =   20
      Top             =   4080
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   397
      Calculator      =   "frmterimawip.frx":0E90
      Caption         =   "frmterimawip.frx":0EB0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmterimawip.frx":0F1C
      Keys            =   "frmterimawip.frx":0F3A
      Spin            =   "frmterimawip.frx":0F7C
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##,###,##0.00;(##,###,##0.00);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##,###,##0.00;(##,###,##0.00)"
      HighlightText   =   0
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
      ReadOnly        =   1
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   -65531
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin Chameleon.chameleonButton ccmdtes 
      Height          =   375
      Left            =   3480
      TabIndex        =   22
      Top             =   3840
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "simpan"
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
      MICON           =   "frmterimawip.frx":0FA4
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
      Caption         =   "* WIP Jadi yang masih diproses jangan discan disini karena barang masuk ke stok gudang pusat"
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   21
      Top             =   2280
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   0
      Top             =   2880
      Width           =   4455
   End
   Begin VB.Label Label5 
      Caption         =   "No. Bukti"
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
      Left            =   240
      TabIndex        =   14
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Barang"
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
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Tgl. Confirm"
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
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Qty"
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
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Palet"
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
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmterimawip"
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
Dim str99 As String
Dim kode_stok As String

Private Sub ccmdtes_Click()
    OBJ.Open dsn
    SQL = "Update am_stokgudang set kg='" & txtkg & "',kgperpalet='" & txtkgperpalet & "',"
    SQL = SQL + "hppperkg='" & txthppperkg & "' Where palet='" & txtpalet & "'"
    Set RST = OBJ.Execute(SQL)
    
    'SQL = "Select * From am_stokgudang Where 0=1"
    'Set RST = New ADODB.Recordset
    'RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    'With RST
        '.AddNew
        '!nolot = txtnolot
        '!palet = txtpalet
        '!tanggal = Format(Date, "yyyy/MM/dd")
        '!ref = txtnobpb
        'If Left(txtkode, 1) = "L" Then
            '!keterangan = "Produksi Lem"
        'ElseIf Left(txtkode, 1) = "K" Then
            '!keterangan = "Produksi Karpet"
        'End If
        '!kodebarang = txtkode
        '!NamaBarang = txtnamabrg
        '!kg = txtkg
        '!kgperpalet = txtkgperpalet
        '!hppperkg = txthppperkg
        '!qin = txtqty
        '!qout = "0.00"
        '!kdsatuan = txtkdsatuan
        '!satuan = txtsatuan
        '!gudang = "G1"
        '!UserName = nmuser
        '!flag = "0"
        '.Update
    'End With
    OBJ.Close
    MsgBox "berhasil simpan", vbInformation, AppName
End Sub

Private Sub cmdclear_Click()
    Call clearform
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdconfirm_Click()
    Dim strformat As String
    If txtpalet = "" Then Exit Sub
    If txtqty = "" Then Exit Sub
    If txtkode = "" Then Exit Sub
    'cek palet sudah discan atau belum !
    OBJ.Open dsn
    SQL = "Select * From list_mutasi_produksi_header Where kode_palet = '" & txtpalet & "'"
    Set RST = OBJ.Execute(SQL)
    If RST!Status = "1" Then
        'MsgBox "Pallet has been received", vbCritical, AppName
        'OBJ.Close
        'Exit Sub
    End If
    OBJ.Close
    
    'ambil kodestok
    kode_stok = GetKdStok
    
OBJ1.Open dsn
    SQL1 = "Select * From list_produksi_hasil Where noref = '" & txtpalet & "'"
    RST1.Open SQL1, OBJ1, adOpenDynamic, adLockOptimistic
    
    If RST1.EOF Then
        MsgBox "Kode palet not found", vbCritical, AppName
        OBJ1.Close
        Exit Sub
    End If
    
    'ambil kode pindah gudang baru
    strformat = Format(Date1, "yymm")
    
    If Left(txtkode, 1) = "L" Then
        OBJ.Open dsn
        SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'PHL0-' + '" + strformat + "%' order by nobpb desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            str99 = Right(RST!nobpb, 4)
        Else
            str99 = 0
        End If
        OBJ.Close
        str99 = str99 + 1
        If Len(str99) = 1 Then txtnobpb = "PHL0-" & strformat & "000" & str99
        If Len(str99) = 2 Then txtnobpb = "PHL0-" & strformat & "00" & str99
        If Len(str99) = 3 Then txtnobpb = "PHL0-" & strformat & "0" & str99
        If Len(str99) = 4 Then txtnobpb = "PHL0-" & strformat & str99
        
    End If
    
    If Left(txtkode, 1) = "K" Then
        OBJ.Open dsn
        SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'PHK0-' + '" + strformat + "%' order by nobpb desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            str99 = Right(RST!nobpb, 4)
        Else
            str99 = 0
        End If
        OBJ.Close
        str99 = str99 + 1
        If Len(str99) = 1 Then txtnobpb = "PHK0-" & strformat & "000" & str99
        If Len(str99) = 2 Then txtnobpb = "PHK0-" & strformat & "00" & str99
        If Len(str99) = 3 Then txtnobpb = "PHK0-" & strformat & "0" & str99
        If Len(str99) = 4 Then txtnobpb = "PHK0-" & strformat & str99
    End If
    
    'simpan ke tabel am_bpbhdr type 88 & 99
    OBJ.Open dsn
'GoTo stokgudang: 'simpan ke am_stokgudang aja karena ketabel lain sudah masuk
    SQL = "Select * From am_bpbhdr Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    With RST
        .AddNew
        !Type = "88"    'in gudang
        !nobpb = txtnobpb
        !tglbpb = Format(Date1, "yyyy/MM/dd")
        !kodegudang = "G1"
        !keterangan = txtpalet
        !noref = txtpalet
        !identry = nmuser
        !dateentry = Format(Date, "yyyy/MM/dd")
        !idupdate = nmuser
        !dateupdate = Format(Date, "yyyy/MM/dd")
        .Update
    End With
    
    SQL = "Select * From am_bpbhdr Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    With RST
        .AddNew
        !Type = "99"    'out wip
        !nobpb = txtnobpb
        !tglbpb = Format(Date1, "yyyy/MM/dd")
        !kodegudang = "G3"
        !keterangan = txtpalet
        !noref = txtpalet
        !identry = nmuser
        !dateentry = Format(Date, "yyyy/MM/dd")
        !idupdate = nmuser
        !dateupdate = Format(Date, "yyyy/MM/dd")
        .Update
    End With

    'simpan ke tabel am_bpblin type 88 & 99
    SQL = "Select * From am_bpblin Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    With RST
        .AddNew
        !Type = "88"    'in gudang
        !nobpb = txtnobpb
        !tglbpb = Format(Date1, "yyyy/MM/dd")
        !kodebarang = txtkode
        !qty = txtqty
        !keterangan = txtpalet
        !kodesatuan = RST1!kode_satuan
        !lineitem = "1"
        .Update
    End With
    
    SQL = "Select * From am_bpblin Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    With RST
        .AddNew
        !Type = "99"    'out wip
        !nobpb = txtnobpb
        !tglbpb = Format(Date1, "yyyy/MM/dd")
        !kodebarang = txtkode
        !qty = txtqty * -1
        !keterangan = txtpalet
        !kodesatuan = RST1!kode_satuan
        !lineitem = "1"
        .Update
    End With
    
    'simpan ke tabel am_stok (G3 out & G1 in); flag = 1 artinya sudah diterima gudang
    SQL = "Select * From am_stok Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    With RST
        .AddNew
        !kdstok = kode_stok
        !tanggal = Format(Date1, "yyyy/MM/dd")
        !nolot = RST1!nolot
        !palet = txtpalet
        !type_trans = "SC"
        !noref = txtnobpb
        !gudang = "G1"  'in gudang (88)
        !kodebarang = txtkode
        !NamaBarang = txtnamabrg
        !kodeproduk = grid.TextMatrix(1, 0)
        !kodesatuan = RST1!kode_satuan
        !awal = "0.00"
        !qtyin = txtqty
        !qtyout = "0.00"
        !isi = grid.TextMatrix(1, 1)
        !kg = grid.TextMatrix(1, 2)
        !hpp = grid.TextMatrix(1, 3)
        !nosj = ""
        !kodecust = ""
        !UserName = nmuser
        !useredit = ""
        !tgledit = Format(Date, "yyyy/MM/dd")
        !baris = "1"
        !keterangan = "Confirm Gudang"
        !flag = "3"
        !hpp_totpack = grid.TextMatrix(1, 4)
        .Update
    End With
    
    SQL = "Select * From am_stok Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    With RST
        .AddNew
        !kdstok = kode_stok
        !tanggal = Format(Date1, "yyyy/MM/dd")
        !nolot = RST1!nolot
        !palet = txtpalet
        !type_trans = "SC"
        !noref = txtnobpb
        !gudang = "G3"  'out wip (99)
        !kodebarang = txtkode
        !NamaBarang = txtnamabrg
        !kodeproduk = grid.TextMatrix(1, 0)
        !kodesatuan = RST1!kode_satuan
        !awal = "0.00"
        !qtyin = "0.00"
        !qtyout = txtqty
        !isi = grid.TextMatrix(1, 1)
        !kg = grid.TextMatrix(1, 2)
        !hpp = grid.TextMatrix(1, 3)
        !nosj = ""
        !kodecust = ""
        !UserName = nmuser
        !useredit = ""
        !tgledit = Format(Date, "yyyy/MM/dd")
        !baris = "2"
        !keterangan = "Confirm Gudang"
        !flag = "1"
        !hpp_totpack = grid.TextMatrix(1, 4)
        .Update
    End With
stokgudang:
    SQL = "Select * From am_stokgudang Where palet = '" & txtpalet & "' and keterangan <> 'SJ'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        SQL = "Update am_stokgudang set kg='" & txtkg & "',kgperpalet='" & txtkgperpalet & "',"
        SQL = SQL + "hppperkg='" & txthppperkg & "' Where palet='" & txtpalet & "'"
        Set RST = OBJ.Execute(SQL)
    Else
    'Tabel stok gudang
        SQL = "Select * From am_stokgudang Where 0=1"
        Set RST = New ADODB.Recordset
        RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
        With RST
            .AddNew
            !nolot = txtnolot
            !palet = txtpalet
            !tanggal = Format(Date1, "yyyy/MM/dd")
            !ref = txtnobpb
            If Left(txtkode, 1) = "L" Then
                !keterangan = "Produksi Lem"
            ElseIf Left(txtkode, 1) = "K" Then
                !keterangan = "Produksi Karpet"
            End If
            !kodebarang = txtkode
            !NamaBarang = txtnamabrg
            !kg = txtkg
            !kgperpalet = txtkgperpalet
            !hppperkg = txthppperkg
            !qin = txtqty
            !qout = "0.00"
            !kdsatuan = txtkdsatuan
            !satuan = txtsatuan
            !gudang = "G1"
            !UserName = nmuser
            !flag = "0"
            .Update
        End With
    End If
OBJ1.Close

    'update status = 1 sudah diterima gudang
    SQL = "Update List_mutasi_produksi_header set status='1' Where kode_palet='" & txtpalet & "'"
    Set RST = OBJ.Execute(SQL)
    
    'close pending lot
    SQL = "Update list_hpp_produksi set flag='2' Where palet='" & txtpalet & "'"
    Set RST = OBJ.Execute(SQL)
    
    'close SOP
    SQL = "Update list_produksi_master set flagprint='4' Where nolot='" & txtnolot & "'"
    Set RST = OBJ.Execute(SQL)
    
    OBJ.Close
    MsgBox "Data saved successfully", vbInformation, AppName
    Call clearform
End Sub

Private Sub clearform()
    Dim strformat As String
    Date1 = Date
    strformat = Format(Date1, "yymm")
    
    txtpalet = ""
    txtkode = ""
    txtnamabrg = ""
    txtqty = "0"
    txtsatuan = ""
    txtkdsatuan = ""
    txtkg = "0.00"
    txtkgperpalet = "0.00"
    txthppperkg = "0.00"
    txtnolot = ""
    txtnobpb = ""
    Date1 = Date
    hapusgrid
    
    'OBJ.Open dsn
    'SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'PG0-' + '" + strformat + "%' order by nobpb desc"
    'Set RST = OBJ.Execute(SQL)
    'If Not RST.EOF Then
        'str99 = Right(RST!nobpb, 4)
    'Else
        'str99 = 0
    'End If
    'OBJ.Close
        
    'str99 = str99 + 1
    
    'If Len(str99) = 1 Then txtnobpb = "PG0-" & strformat & "000" & str99
    'If Len(str99) = 2 Then txtnobpb = "PG0-" & strformat & "00" & str99
    'If Len(str99) = 3 Then txtnobpb = "PG0-" & strformat & "0" & str99
    'If Len(str99) = 4 Then txtnobpb = "PG0-" & strformat & str99
End Sub

Private Sub cmdRepair_Click()
    frmterimawiprepair.Show
End Sub

Private Sub date1_Change()
    Dim strformat As String
    strformat = Format(Date1, "yymm")
    
    OBJ.Open dsn
    SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'PG0-' + '" + strformat + "%' order by nobpb desc"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        str99 = Right(RST!nobpb, 4)
    Else
        str99 = 0
    End If
    OBJ.Close
        
    str99 = str99 + 1
    
    If Len(str99) = 1 Then txtnobpb = "PG0-" & strformat & "000" & str99
    If Len(str99) = 2 Then txtnobpb = "PG0-" & strformat & "00" & str99
    If Len(str99) = 3 Then txtnobpb = "PG0-" & strformat & "0" & str99
    If Len(str99) = 4 Then txtnobpb = "PG0-" & strformat & str99
End Sub

Private Sub Form_Load()
    'Dim strformat As String
    Date1 = Date
    'strformat = Format(Date1, "yymm")
    
    'OBJ.Open dsn
    'SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'PG0-' + '" + strformat + "%' order by nobpb desc"
    'Set RST = OBJ.Execute(SQL)
    'If Not RST.EOF Then
        'str99 = Right(RST!nobpb, 4)
    'Else
        'str99 = 0
    'End If
    'OBJ.Close
        
    'str99 = str99 + 1
    
    'If Len(str99) = 1 Then txtnobpb = "PG0-" & strformat & "000" & str99
    'If Len(str99) = 2 Then txtnobpb = "PG0-" & strformat & "00" & str99
    'If Len(str99) = 3 Then txtnobpb = "PG0-" & strformat & "0" & str99
    'If Len(str99) = 4 Then txtnobpb = "PG0-" & strformat & str99
    
    grid.TextMatrix(0, 0) = "Kode Produk"
    grid.TextMatrix(0, 1) = "isi"
    grid.TextMatrix(0, 2) = "kg"
    grid.TextMatrix(0, 3) = "hpp jadi/kg"
    grid.TextMatrix(0, 4) = "Total hpp pack"
    
    If nmuser = "Creator" Then cmdRepair.Visible = True
End Sub

Private Sub hapusgrid()
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 0) = "" Then Exit Do
        grid.TextMatrix(grid.Row, 0) = ""
        grid.TextMatrix(grid.Row, 1) = ""
        grid.TextMatrix(grid.Row, 2) = ""
        grid.TextMatrix(grid.Row, 3) = ""
        grid.TextMatrix(grid.Row, 4) = ""
        grid.Rows = grid.Row + 1
    Loop
    grid.Rows = 2
End Sub


Private Sub txtpalet_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Dim strformat As String
    If KeyAscii = 13 Then
        OBJ.Open dsn
        
        SQL = "select * From list_hpp_produksi Where palet='" & txtpalet & "' and flag = '0'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Palet belum divalidasi", vbCritical, AppName
            OBJ.Close
            txtpalet = ""
            Exit Sub
        End If
        
        SQL = "select * From list_hpp_produksi Where palet='" & txtpalet & "' and flag='2'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            'MsgBox "Palet sudah diterima", vbCritical, AppName
            'OBJ.Close
            'txtpalet = ""
            'Exit Sub
        End If
        
        SQL = "Select * From am_stokgudang where palet='" & txtpalet & "' and keterangan <> 'SJ'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            'MsgBox "Palet sudah diterima.", vbCritical, AppName
            'OBJ.Close
            'txtpalet = ""
            'Exit Sub
        End If
        
        SQL = "Select a.*,b.NamaSatuan From am_stok a inner join am_unit b on a.kodesatuan = b.KodeSatuan"
        SQL = SQL + " Where a.palet='" & txtpalet & "'"
        Set RST = OBJ.Execute(SQL)
        If RST.EOF Then
            MsgBox "Kode Palet tidak ditemukan", vbCritical, AppName
            OBJ.Close
            txtpalet = ""
            Exit Sub
        End If
        
        txtkode = RST!kodebarang
        txtnamabrg = RST!NamaBarang
        txtqty = RST!qtyin
        txtsatuan = RST!namasatuan
        txtkdsatuan = RST!kodesatuan
        txtnolot = RST!noref
        
        grid.Row = 1
        With grid
            .TextMatrix(.Row, 0) = RST!kodeproduk
            .TextMatrix(.Row, 1) = RST!isi
            .TextMatrix(.Row, 2) = Format(RST!kg, "##,##0.00")
            .TextMatrix(.Row, 3) = Format(RST!hpp, "##,###,##0.00")
            .TextMatrix(.Row, 4) = Format(RST!hpp_totpack, "##,###,##0.00")
        End With
        
        SQL = "select * From list_hpp_produksi Where palet='" & txtpalet & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            txtkg = RST!kg
            txtkgperpalet = RST!kgperpalet
            txthppperkg = RST!hppperkg
        End If
        OBJ.Close
           
        strformat = Format(Date1, "yymm")
    
        If Left(txtkode, 1) = "L" Then
            OBJ.Open dsn
            SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'PHL0-' + '" + strformat + "%' order by nobpb desc"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                str99 = Right(RST!nobpb, 4)
            Else
                str99 = 0
            End If
            OBJ.Close
            str99 = str99 + 1
            If Len(str99) = 1 Then txtnobpb = "PHL0-" & strformat & "000" & str99
            If Len(str99) = 2 Then txtnobpb = "PHL0-" & strformat & "00" & str99
            If Len(str99) = 3 Then txtnobpb = "PHL0-" & strformat & "0" & str99
            If Len(str99) = 4 Then txtnobpb = "PHL0-" & strformat & str99
            
        End If
        
        If Left(txtkode, 1) = "K" Then
            OBJ.Open dsn
            SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'PHK0-' + '" + strformat + "%' order by nobpb desc"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                str99 = Right(RST!nobpb, 4)
            Else
                str99 = 0
            End If
            OBJ.Close
            str99 = str99 + 1
            If Len(str99) = 1 Then txtnobpb = "PHK0-" & strformat & "000" & str99
            If Len(str99) = 2 Then txtnobpb = "PHK0-" & strformat & "00" & str99
            If Len(str99) = 3 Then txtnobpb = "PHK0-" & strformat & "0" & str99
            If Len(str99) = 4 Then txtnobpb = "PHK0-" & strformat & str99
        End If
    End If
End Sub
