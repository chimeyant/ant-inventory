VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Begin VB.Form frmAddmut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Mutasi Barang Jadi Ke Bahan Baku"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8790
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   8790
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtbaseBB 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Left            =   5280
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   13
      Top             =   3375
      Width           =   2655
   End
   Begin VB.TextBox txtqtybb 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5280
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   12
      Text            =   "0"
      Top             =   3915
      Width           =   975
   End
   Begin VB.TextBox txtnmbjadi 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   9
      Top             =   3375
      Width           =   2655
   End
   Begin VB.TextBox txtqty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   8
      Text            =   "0"
      Top             =   3900
      Width           =   735
   End
   Begin MSComCtl2.DTPicker Date1 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   3000
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
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
      CustomFormat    =   "ddMMMMyyyy"
      Format          =   143327235
      CurrentDate     =   42052
   End
   Begin VB.TextBox txtnobukti 
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
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   0
      Top             =   2400
      Width           =   2175
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   2055
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   3625
      _Version        =   393216
      Cols            =   7
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
      _Band(0).Cols   =   7
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   7800
      TabIndex        =   4
      Top             =   4560
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
      MICON           =   "frmAddmut.frx":0000
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
      Left            =   6840
      TabIndex        =   3
      Top             =   4560
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
      MICON           =   "frmAddmut.frx":031A
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
      Left            =   5880
      TabIndex        =   2
      Top             =   4560
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
      MICON           =   "frmAddmut.frx":0634
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdBB 
      Height          =   285
      Left            =   5160
      TabIndex        =   14
      Top             =   3000
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "Base Bahan Baku"
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
      MICON           =   "frmAddmut.frx":094E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Hpp/Kg"
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
      Left            =   360
      TabIndex        =   21
      Top             =   4560
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000A&
      BorderWidth     =   2
      X1              =   4560
      X2              =   4560
      Y1              =   3000
      Y2              =   4320
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Kg"
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
      Left            =   6480
      TabIndex        =   20
      Top             =   3915
      Width           =   375
   End
   Begin VB.Label lblhpp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   1440
      TabIndex        =   19
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label lblkdsatuan 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Left            =   4560
      TabIndex        =   18
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label lblsatuan 
      BackStyle       =   0  'Transparent
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
      Left            =   2520
      TabIndex        =   17
      Top             =   3900
      Width           =   1695
   End
   Begin VB.Label lblkdbrgjadi 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3360
      TabIndex        =   16
      Top             =   3000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblkdbasebb 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   7440
      TabIndex        =   15
      Top             =   3840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   360
      Left            =   5160
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   360
      Left            =   5160
      Shape           =   4  'Rounded Rectangle
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Barang"
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
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Left            =   360
      TabIndex        =   10
      Top             =   3960
      Width           =   975
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   360
      Left            =   1440
      Shape           =   4  'Rounded Rectangle
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   360
      Left            =   1440
      Shape           =   4  'Rounded Rectangle
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Mutasi"
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
      Left            =   0
      TabIndex        =   6
      Top             =   3030
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "No Mutasi"
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
      TabIndex        =   5
      Top             =   2430
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   0
      Top             =   2880
      Width           =   8775
   End
End
Attribute VB_Name = "frmAddmut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim strnomut, strlot As String

Private Sub cmdadd_Click()
    Dim kode_stok As String
    
    If txtnobukti = "" Then
        MsgBox "Please fill in the No Mutasi column", vbExclamation, "Warning"
        txtnobukti.SetFocus
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "Select * From am_muthdr Where nomut = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        MsgBox "No Mutasi is already exist" + vbCrLf + "Save Abort, Click OK To Continue ...", vbExclamation, "Information"
        txtnobukti.SetFocus
        Exit Sub
    End If
    OBJ.Close
    
    If txtbaseBB = "" Then
        MsgBox "Please select Base Bahan Baku first", vbExclamation, "Warning"
        cmdBB.SetFocus
        Exit Sub
    End If
    
    strnomut = getnomut
    
    OBJ.Open dsn
    SQL = "Select * From am_muthdr Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    With RST
        .AddNew
        !nomut = txtnobukti
        !tglmut = Date1
        !Type = "01"
        !keterangan = strnomut
        .Update
    End With
    
    SQL = "Select * From am_mutlin Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    
    With RST
        .AddNew
        !nomut = txtnobukti
        !Type = "01"
        !kodebarang = lblkdbasebb
        !qty = txtqtybb
        !kodesatuan = "002"
        !lineitem = "1"
        .Update
    End With
    
    'ambil kodestok
    kode_stok = GetNoStok
    
    'SIMPAN KE AM_STOKBARANG
    SQL = "Select * From am_stokbarang Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    
    RST.AddNew
    RST!kode_stok = kode_stok
    RST!tanggal = Format(Date1, "yyyy/MM/dd")
    RST!type_transaksi = "M1"
    RST!no_transaksi = txtnobukti
    RST!ref = strlot
    RST!KODE_SUPORCUST = ""
    RST!TYPE_BARANG = "BAHAN BAKU"
    RST!GROUP_BARANG = ""
    RST!kode_barang = lblkdbasebb 'kode bahan baku base
    RST!LOT_NUMBER = strlot
    RST!kode_satuan = "002"
    RST!QTY_AWAL = 0
    RST!QTY_MASUK = txtqtybb
    RST!QTY_KELUAR = 0
    RST!NO_ACC = ""
    RST!KODE_CUR = ""
    RST!NILAI_CUR = 0
    RST!HARGA_AWAL = 0
    RST!HARGA_MASUK = 0
    RST!HARGA_KELUAR = 0
    RST!keterangan = "Mutasi Base WIP Ke Base Bahan Baku"
    RST!ON_PO = "0"
    RST!ON_SO = "0"
    RST!ON_DELV = "0"
    RST!ON_USE = "0"
    RST!ON_CLOSE = "0"
    RST!flag = "0"
    RST!BARIS = "1"
    RST!UserName = nmuser
    RST.Update
    
     'SIMPAN KE AM_STOKLOT
    SQL = "Select * From am_stoklot Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    
    RST.AddNew
    RST!lotstok = strnomut
    RST!nolot = strlot
    RST!kodebahan = lblkdbasebb
    RST!qtybahan = txtqtybb
    RST!kodesatuan = "002"
    RST!hpp = lblhpp * txtqtybb
    RST!flag = "0"
    RST.Update
    
    'SIMPAN KE AM_STOKWIP
    SQL = "Select * from am_stokwip Where nolot='" & strlot & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If Format(RST!qin, "#,##0.00") = Format(txtqty, "#,##0.00") Then
            SQL = "Update am_stokwip set qout = '" & txtqty & "',flag= '1' Where nolot = '" & strlot & "'"
            Set RST = OBJ.Execute(SQL)
        Else
            SQL = "Update am_stokwip set qout = '" & txtqty & "' Where nolot = '" & strlot & "'"
            Set RST = OBJ.Execute(SQL)
        End If
    End If
    
    OBJ.Close
    
    MsgBox "Data Is Save, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdBB_Click()
    If lblkdbrgjadi = "" Then Exit Sub
    namatabel = "Base Bahan"
    carisql1 = "Select kodebarang,namabarang From am_apitemmst"
    carisql1 = carisql1 + " Where (KodeBarang like 'L10.%' or KodeBarang like 'L11.%' or KodeBarang like 'L12.%' or KodeBarang like 'L13.%' or KodeBarang like 'K05.%')"
    frmsearch.Show vbModal
End Sub

Private Sub cmdBB_GotFocus()
    If hasil = "" Then Exit Sub
    txtbaseBB = hasil1
    lblkdbasebb = hasil
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    carisql1 = ""
End Sub

Private Sub cmdclear_Click()
On Error GoTo Err_handler:
    hapusgrid
    OBJ.Open dsn
    SQL = "Select nolot,kodebarang,namabarang,(qin)- (qout)'qty',kdsatuan,satuan From am_stokwip"
    SQL = SQL + " Where flag='0'" ' group by nolot,kodebarang,namabarang,kdsatuan,satuan"
    Set RST = OBJ.Execute(SQL)
    
    grid.Row = 1
    Do While Not RST.EOF
        grid.TextMatrix(grid.Row, 0) = grid.Row
        grid.TextMatrix(grid.Row, 1) = RST!nolot
        grid.TextMatrix(grid.Row, 2) = RST!kodebarang
        grid.TextMatrix(grid.Row, 3) = RST!namabarang
        grid.TextMatrix(grid.Row, 4) = Format(RST!qty, "##,###,##0.00")
        grid.TextMatrix(grid.Row, 5) = RST!kdsatuan
        grid.TextMatrix(grid.Row, 6) = RST!satuan
        grid.Rows = grid.Rows + 1
        grid.Row = grid.Row + 1
        RST.MoveNext
    Loop
    OBJ.Close
    
    txtnobukti = ""
    txtnmbjadi = ""
    lblkdbrgjadi = ""
    txtqty = "0"
    txtbaseBB = ""
    lblkdbasebb = ""
    txtqtybb = "0"
    lblhpp = "0"
    
    Exit Sub
Err_handler:
    If OBJ.State = 1 Then OBJ.Close
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    grid.TextMatrix(0, 1) = "Nolot"
    grid.TextMatrix(0, 2) = "Kode"
    grid.TextMatrix(0, 3) = "Item"
    grid.TextMatrix(0, 4) = "Qty"
    grid.TextMatrix(0, 5) = "K/Sat."
    grid.TextMatrix(0, 6) = "Satuan"
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1500
    grid.ColWidth(2) = 1000
    grid.ColWidth(3) = 2500
    grid.ColWidth(4) = 800
    grid.ColWidth(5) = 0
    grid.ColWidth(6) = 1000
    
    grid.RowHeightMin = 300
    
    Date1.Value = Date
    
    hapusgrid
    OBJ.Open dsn
    SQL = "Select nolot,kodebarang,namabarang,(qin)- (qout)'qty',kdsatuan,satuan From am_stokwip"
    SQL = SQL + " Where flag='0'" ' group by nolot,kodebarang,namabarang,kdsatuan,satuan"
    Set RST = OBJ.Execute(SQL)
    
    grid.Row = 1
    Do While Not RST.EOF
        grid.TextMatrix(grid.Row, 0) = grid.Row
        grid.TextMatrix(grid.Row, 1) = RST!nolot
        grid.TextMatrix(grid.Row, 2) = RST!kodebarang
        grid.TextMatrix(grid.Row, 3) = RST!namabarang
        grid.TextMatrix(grid.Row, 4) = Format(RST!qty, "##,###,##0.00")
        grid.TextMatrix(grid.Row, 5) = RST!kdsatuan
        grid.TextMatrix(grid.Row, 6) = RST!satuan
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
        grid.TextMatrix(grid.Row, 1) = ""
        grid.TextMatrix(grid.Row, 2) = ""
        grid.TextMatrix(grid.Row, 3) = ""
        grid.TextMatrix(grid.Row, 4) = ""
        grid.TextMatrix(grid.Row, 5) = ""
        grid.TextMatrix(grid.Row, 6) = ""
        'grid.TextMatrix(grid.Row, 7) = ""

        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
End Sub

Private Sub hapusrow()
    grid.TextMatrix(grid.Row, 1) = ""
    grid.TextMatrix(grid.Row, 2) = ""
    grid.TextMatrix(grid.Row, 3) = ""
    grid.TextMatrix(grid.Row, 4) = ""
    grid.TextMatrix(grid.Row, 5) = ""
    grid.TextMatrix(grid.Row, 6) = ""
    grid.TextMatrix(grid.Row, 7) = ""
    Do While True
        If grid.TextMatrix(grid.Row + 1, 1) = "" Then
            grid.TextMatrix(grid.Row, 1) = ""
            grid.TextMatrix(grid.Row, 2) = ""
            grid.TextMatrix(grid.Row, 3) = ""
            grid.TextMatrix(grid.Row, 4) = ""
            grid.TextMatrix(grid.Row, 5) = ""
            grid.TextMatrix(grid.Row, 6) = ""
            grid.TextMatrix(grid.Row, 7) = ""
            Exit Do
        End If
        grid.TextMatrix(grid.Row, 1) = grid.TextMatrix(grid.Row + 1, 1)
        grid.TextMatrix(grid.Row, 2) = grid.TextMatrix(grid.Row + 1, 2)
        grid.TextMatrix(grid.Row, 3) = grid.TextMatrix(grid.Row + 1, 3)
        grid.TextMatrix(grid.Row, 4) = grid.TextMatrix(grid.Row + 1, 4)
        grid.TextMatrix(grid.Row, 5) = grid.TextMatrix(grid.Row + 1, 5)
        grid.TextMatrix(grid.Row, 6) = grid.TextMatrix(grid.Row + 1, 6)
        grid.TextMatrix(grid.Row, 7) = grid.TextMatrix(grid.Row + 1, 7)
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = grid.Rows - 1
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    
    Select Case grid.Col

        Case 0, 1, 2, 3, 4, 6
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            lblkdbrgjadi = grid.TextMatrix(grid.Row, 2)
            txtnmbjadi = grid.TextMatrix(grid.Row, 3)
            txtqty = grid.TextMatrix(grid.Row, 4)
            lblsatuan = grid.TextMatrix(grid.Row, 6)
            
            OBJ.Open dsn
            SQL = "Select * From am_stokwip Where nolot='" & grid.TextMatrix(grid.Row, 1) & "'"
            SQL = SQL + " and kodebarang='" & grid.TextMatrix(grid.Row, 2) & "' and qin='" & txtqty & "'"
            Set RST = OBJ.Execute(SQL)
            
            If Not RST.EOF Then
                txtqtybb = Format(RST!kg * RST!qin, "#,##0.00")
                lblhpp = RST!hppperkg
                strlot = RST!nolot
            End If
            OBJ.Close
    End Select
End Sub

Function getnomut() As String    '2016060001'
    On Error GoTo Err_handler:
    Dim SQL As String
    Dim strnumber As String
    Dim tempkode As String
    Dim kode As Long
    
    strnumber = Format(Date, "yyyymm")
    
    Set OBJ = New ADODB.Connection
    OBJ.Open dsn
    SQL = "select max(lotstok)as kr from am_stoklot where lotstok like '" + strnumber + "%'"
    Set RST = OBJ.Execute(SQL)

    If IsNull(RST!kr) = True Or RST!kr = "" Then
        getnomut = strnumber + "0001"
    Else
        kode = CLng(Mid(RST!kr, 7, 4)) + 1
        
        If (Len(Trim(Str(kode))) = 1) Then
            tempkode = strnumber + "000" + Trim(Str(kode))
        End If
        If (Len(Trim(Str(kode))) = 2) Then
            tempkode = strnumber + "00" + Trim(Str(kode))
        End If
        If (Len(Trim(Str(kode))) = 3) Then
            tempkode = strnumber + "0" + Trim(Str(kode))
        End If
        If (Len(Trim(Str(kode))) = 4) Then
            tempkode = strnumber + Trim(Str(kode))
        End If
        getnomut = tempkode
    End If
    OBJ.Close
    Exit Function
Err_handler:
    getnomut = strnumber + "0001"
End Function
