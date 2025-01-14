VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Begin VB.Form frmWiptoBb 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mutasi WIP ke Bahan Baku"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optjadi 
      Caption         =   "WIP karpet / WIP Jadi Lem"
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
      Left            =   2880
      TabIndex        =   23
      Top             =   0
      Width           =   2295
   End
   Begin VB.OptionButton optbase 
      Caption         =   "Base WIP Lem"
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
      Left            =   1320
      TabIndex        =   22
      Top             =   0
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.TextBox txtqtybb 
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
      Left            =   4920
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   10
      Top             =   2535
      Width           =   1815
   End
   Begin VB.TextBox txtqty 
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
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   9
      Top             =   2535
      Width           =   1815
   End
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
      Left            =   4920
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   7
      Top             =   2055
      Width           =   2655
   End
   Begin VB.TextBox txtnmbase 
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
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   5
      Top             =   2055
      Width           =   2655
   End
   Begin VB.TextBox txtnolot 
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
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   3
      Top             =   360
      Width           =   2175
   End
   Begin VB.TextBox txtnomut 
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
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   6960
      TabIndex        =   0
      Top             =   3120
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
      MICON           =   "frmWiptoBb.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdWIP 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   1680
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "Base WIP"
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
      MICON           =   "frmWiptoBb.frx":031A
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
      Left            =   4800
      TabIndex        =   8
      Top             =   1680
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
      MICON           =   "frmWiptoBb.frx":0634
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsave 
      Height          =   375
      Left            =   6000
      TabIndex        =   11
      Top             =   3120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Proses"
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
      MICON           =   "frmWiptoBb.frx":094E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdlot 
      Height          =   285
      Left            =   240
      TabIndex        =   12
      Top             =   360
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "No.Lot"
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
      MICON           =   "frmWiptoBb.frx":0C68
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker Date1 
      Height          =   285
      Left            =   5880
      TabIndex        =   20
      Top             =   360
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
      Format          =   132448259
      CurrentDate     =   42052
   End
   Begin VB.Label Label7 
      Caption         =   "Tgl. Mutasi"
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
      Left            =   4800
      TabIndex        =   21
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "* Nilai Hpp /Kg = Total Hpp Bahan Baku : Total Qty Hasil"
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
      TabIndex        =   19
      Top             =   3360
      Width           =   4455
   End
   Begin VB.Label Label5 
      Caption         =   "u"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   21.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4305
      TabIndex        =   18
      Top             =   1920
      Width           =   495
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7800
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MUTASI WIP KE BAHAN BAKU"
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
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Width           =   7695
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   360
      Left            =   4800
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   360
      Left            =   4800
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   360
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   360
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
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
      TabIndex        =   16
      Top             =   2535
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      Left            =   240
      TabIndex        =   15
      Top             =   2040
      Width           =   975
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
      Left            =   6840
      TabIndex        =   14
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblkodebase 
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
      Left            =   0
      TabIndex        =   13
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000006&
      FillColor       =   &H00404040&
      Height          =   1935
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   7695
   End
   Begin VB.Label lblhpp_kilo 
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
      Top             =   3120
      Width           =   3495
   End
   Begin VB.Label Label2 
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
      Left            =   240
      TabIndex        =   2
      Top             =   750
      Width           =   975
   End
End
Attribute VB_Name = "frmWiptoBb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String
Dim OBJ1 As New ADODB.Connection
Dim RST1 As New ADODB.Recordset
Dim SQL1 As String

Dim kode_stok As String
Dim hppkg As Double
Dim strpalet, strnomut As String

    Dim totbahan As Double
    Dim kilowip As Double
    Dim kgprice As Double
    Dim qtywip As Double
    Dim losswip As Double

Private Sub cmdBB_Click()
    If lblkodebase = "" Then Exit Sub
    namatabel = "Base Bahan"
    carisql1 = "Select kodebarang,namabarang From am_apitemmst"
    carisql1 = carisql1 + " Where (KodeBarang like 'L10.%' or KodeBarang like 'L11.%' or KodeBarang like 'L12.%' or KodeBarang like 'L13.%' or KodeBarang like 'K05.%')"
    frmsearch.Show vbModal
End Sub

Private Sub cmdBB_GotFocus()
    If hasil = "" Then Exit Sub
    txtbaseBB = hasil1
    lblkdbasebb = hasil
    txtqtybb = txtqty
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    carisql1 = ""
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdlot_Click()
    If cmdWIP.Caption = "Base WIP" Then
        namatabel = "nolot"
        carisql1 = "Select a.kode_produk,a.nama_produk,b.nolot from list_produk_master a"
        carisql1 = carisql1 + " inner join list_produksi_master b on a.kode_produk=b.kode_produk"
        carisql1 = carisql1 + " inner join am_sopbase c on b.nolot = c.nolot"
        carisql1 = carisql1 + " where b.flagprint <= '4'"
        frmsearch.Show vbModal
    ElseIf cmdWIP.Caption = "WIP Jadi" Then
        namatabel = "LotWIP Jadi"
        carisql1 = "Select distinct a.kode_produk,c.nama_produk,a.nolot from list_produksi_hasil a"
        carisql1 = carisql1 + " inner join list_produksi_master b on a.nolot = b.nolot"
        carisql1 = carisql1 + " inner join list_produk_master c on a.kode_produk = c.kode_produk"
        frmsearch.Show vbModal
    End If
End Sub

Private Sub cmdlot_GotFocus()
    If hasil = "" Then Exit Sub
    txtnolot = hasil2
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    carisql1 = ""
    openhpp
    txtnomut = getnomut
End Sub

Private Sub openhpp()
    OBJ.Open dsn
    If cmdWIP.Caption = "Base WIP" Then
        SQL = "Select a.hpp/b.hasil'HPP' from"
        SQL = SQL + " (Select nolot,Nullif(SUM(hpp),0)'hpp' from list_produksi_child where nolot= '" & txtnolot & "' group by nolot)a inner join"
        SQL = SQL + " (Select nolot,Nullif(SUM(qty_bahan),0)'hasil' from list_produksi_hasil Where nolot = '" & txtnolot & "' group by nolot)b "
        SQL = SQL + " on a.nolot=b.nolot"
        Set RST = OBJ.Execute(SQL)
        
        If Not RST.EOF Then
            lblhpp_kilo = "Hpp perkilo : " & Format(RST!hpp, "###,##0.00")
            
            If RST!hpp = "0" Or IsNull(RST!hpp) Then
                OBJ.Close
                MsgBox "Stok lot ini telah habis", vbExclamation, AppName
                txtnolot = ""
                txtnomut = ""
                Exit Sub
            Else
                hppkg = Format(RST!hpp, "###,##0.00")
            End If
        Else
            lblhpp_kilo = "0"
            hppkg = "0"
        End If
        
    ElseIf cmdWIP.Caption = "WIP Jadi" Then
        'periksa stok lot wip
        SQL = "Select a.hpp/b.hasil'HPP' from"
        SQL = SQL + " (Select nolot,Nullif(SUM(hpp),0)'hpp' from list_produksi_child where nolot= '" & txtnolot & "' group by nolot)a inner join"
        SQL = SQL + " (Select nolot,Nullif(SUM(qty_bahan),0)'hasil' from list_produksi_hasil Where nolot = '" & txtnolot & "' group by nolot)b "
        SQL = SQL + " on a.nolot=b.nolot"
        Set RST = OBJ.Execute(SQL)
        
        If Not RST.EOF Then
            lblhpp_kilo = "Hpp perkilo : " & Format(RST!hpp, "###,##0.00")
            
            If RST!hpp = "0" Or IsNull(RST!hpp) Then
                OBJ.Close
                MsgBox "Stok lot ini telah habis", vbExclamation, AppName
                txtnolot = ""
                txtnomut = ""
                Exit Sub
            Else
                SQL = "Select nolot,Nullif(SUM(hpp),0)'hpp',Nullif(SUM(hpp)/SUM(qty_bahan),0)'perkg'"
                SQL = SQL + " from list_produksi_child where nolot= '" & txtnolot & "' group by nolot"
                Set RST = OBJ.Execute(SQL)
                If Not RST.EOF Then
                    kgprice = RST!perkg        'Hpp Bahan 1Kg/item (kg base unit) jika karet ini dikali kg per lembar
                    totbahan = RST!hpp
                End If
                
                SQL = "Select a.nolot,Nullif(SUM(a.qty_bahan*b.kg),0)'hasil' from list_produksi_hasil a"
                SQL = SQL + " inner join (Select kodebarang,kodesatuan,tahun,case when SUBSTRING('" & txtnolot & "' ,3,1) ='A' or SUBSTRING('" & txtnolot & "' ,3,2) ='01' Then kg1"
                SQL = SQL + " when SUBSTRING('" & txtnolot & "' ,3,1)='B' or SUBSTRING('" & txtnolot & "' ,3,2) ='02' Then kg2"
                SQL = SQL + " when SUBSTRING('" & txtnolot & "' ,3,1)='C' or SUBSTRING('" & txtnolot & "' ,3,2) ='03' Then kg3"
                SQL = SQL + " when SUBSTRING('" & txtnolot & "' ,3,1)='D' or SUBSTRING('" & txtnolot & "' ,3,2) ='04' Then kg4"
                SQL = SQL + " when SUBSTRING('" & txtnolot & "' ,3,1)='E' or SUBSTRING('" & txtnolot & "' ,3,2) ='05' Then kg5"
                SQL = SQL + " when SUBSTRING('" & txtnolot & "' ,3,1)='F' or SUBSTRING('" & txtnolot & "' ,3,2) ='06' Then kg6"
                SQL = SQL + " when SUBSTRING('" & txtnolot & "' ,3,1)='G' or SUBSTRING('" & txtnolot & "' ,3,2) ='07' Then kg7"
                SQL = SQL + " when SUBSTRING('" & txtnolot & "' ,3,1)='H' or SUBSTRING('" & txtnolot & "' ,3,2) ='08' Then kg8"
                SQL = SQL + " when SUBSTRING('" & txtnolot & "' ,3,1)='J' or SUBSTRING('" & txtnolot & "' ,3,2) ='09' Then kg9"
                SQL = SQL + " when SUBSTRING('" & txtnolot & "' ,3,1)='K' or SUBSTRING('" & txtnolot & "' ,3,2) ='10' Then kg10"
                SQL = SQL + " when SUBSTRING('" & txtnolot & "' ,3,1)='L' or SUBSTRING('" & txtnolot & "' ,3,2) ='11' Then kg11"
                SQL = SQL + " when SUBSTRING('" & txtnolot & "' ,3,1)='M' or SUBSTRING('" & txtnolot & "' ,3,2) ='12' Then kg12"
                SQL = SQL + " End as kg from am_itemkg) b on a.kode_bahan = b.kodebarang and a.kode_satuan = b.kodesatuan"
                SQL = SQL + " Where a.nolot = '" & txtnolot & "' and a.proses_ke = '2' and b.tahun = '20' + LEFT('" & txtnolot & "',2) group by nolot"
                Set RST = OBJ.Execute(SQL)
                If Not RST.EOF Then
                    kilowip = RST!hasil
                End If
                'hpp loss perkilo
                lblhpp_kilo = "Hpp perkilo : " & Format((totbahan - (kilowip * kgprice)) / kilowip, "###,##0.00")
                losswip = Format((totbahan - (kilowip * kgprice)) / kilowip, "###,##0.00")
            End If
        Else
            lblhpp_kilo = "0"
            hppkg = "0"
        End If
    End If
    
    OBJ.Close
End Sub

Private Sub cmdsave_Click()
    Dim satuan As String
    If txtnolot = "" Then Exit Sub
    If lblkodebase = "" Or lblkdbasebb = "" Then
        MsgBox "Data tidak lengkap", vbCritical, AppName
        Exit Sub
    End If
    'Cek SOP lengkap atau belum, kalau belum tidak boleh dimutasi supaya HPP terhitung semua(all palet)
    OBJ.Open dsn
    SQL = "Select * From list_produksi_master where nolot ='" & txtnolot & "' and flagprint > '3'"
    Set RS = OBJ.Execute(SQL)
    If RS.EOF Then
        MsgBox "Maaf, SOP Lot " & txtnolot & " tidak bisa diproses" & vbCrLf & _
        "Mohon tunggu hingga semua hasil produksi selesai", vbCritical, AppName
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close

    strnomut = getnomut
    
    'SIMPAN KE LIST_PRODUKSI_HASIL
    'Keluarkan Base dari WIP
    OBJ1.Open dsn
    SQL1 = "Select * From list_produksi_hasil Where nolot='" & txtnolot & "'"
    Set RST1 = OBJ1.Execute(SQL1)
    If Not RST1.EOF Then satuan = RST1!kode_satuan
    
    OBJ.Open dsn
    SQL = "Select * From list_produksi_hasil Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic

    RST.AddNew
    RST!kode_produk = RST1!kode_produk
    RST!nolot = txtnolot
    RST!kode_bahan = lblkodebase 'kode barang jadi base
    RST!Lot_bahan = txtnomut
    RST!qty_bahan = txtqtybb * -1
    RST!kode_satuan = RST1!kode_satuan
    RST!flag_tambahan = "1"
    RST!tanggal = Format(Date1, "yyyy/MM/dd")
    RST!noref = txtnomut
    RST!proses_ke = "5"
    RST.Update

    'SIMPAN KE LIST_MUTASI_PRODUKSI_DETAIL
    Dim palet As Integer
'    OBJ1.Open dsn
    SQL1 = "Select kode_satuan,max(LEFT(kode_palet,2))'kode' From list_mutasi_produksi_details Where kode_palet like '%" & txtnolot & "'"
    SQL1 = SQL1 + " Group by kode_satuan"
    Set RST1 = OBJ1.Execute(SQL1)
    If Not RST1.EOF Then
        palet = RST1!kode + 1
        If palet >= "10" Then
            strpalet = palet
        Else
            strpalet = "0" & palet
        End If
    End If
    
    SQL = "Select * From list_mutasi_produksi_details Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic

    RST.AddNew
    RST!kode_palet = strpalet & txtnolot
    RST!kode_barang = lblkodebase 'kode barang jadi base
    RST!kode_satuan = satuan
    RST!qty = txtqtybb * -1
    RST!BARIS = "2"
    RST.Update
    OBJ1.Close
    
    'note : WIP Base tidak masuk ke Gudang, jadi tidak perlu di insert ke am_bpbhdr dan am_bpblin
    'masukkan base ke stok bahan baku
    'SIMPAN KE AM_MUTHDR
    'CEK NOMUT SUDAH ADA ATAU BELUM
    Dim DoubleLot As Boolean
    Dim TripleLot As Boolean
    Dim QuardLot As Boolean
    Dim strlot As String
    Dim i As Integer
    SQL = "Select * From am_muthdr Where nomut='" & txtnolot & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        DoubleLot = True
    End If
    strlot = txtnolot & "/1"
    SQL = "Select * From am_muthdr Where nomut='" & strlot & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        TripleLot = True
    End If
    strlot = txtnolot & "/2"
    SQL = "Select * From am_muthdr Where nomut='" & strlot & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        QuardLot = True
    End If
    
    SQL = "Select * From am_muthdr Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    
    RST.AddNew
    If DoubleLot = True Then
        If TripleLot = True Then
            RST!nomut = txtnolot & "/2"
        ElseIf QuardLot = True Then
            RST!nomut = txtnolot & "/3"
        Else
            RST!nomut = txtnolot & "/1"
        End If
    Else
        RST!nomut = txtnolot
    End If
    RST!tglmut = Format(Date1, "yyyy/MM/dd")
    RST!Type = "01"
    RST!keterangan = txtnomut
    RST.Update
    
    'SIMPAN KE AM_MUTLIN
    SQL = "Select * From am_mutlin Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    
    RST.AddNew
    If DoubleLot = True Then
        If TripleLot = True Then
            RST!nomut = txtnolot & "/2"
        ElseIf QuardLot = True Then
            RST!nomut = txtnolot & "/3"
        Else
            RST!nomut = txtnolot & "/1"
        End If
    Else
        RST!nomut = txtnolot
    End If
    RST!Type = "01"
    RST!kodebarang = lblkdbasebb 'kode bahan baku base
    RST!qty = txtqtybb
    RST!kodesatuan = "002"
    RST!lineitem = "1"
    RST.Update
    
    DoubleLot = False
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
    RST!no_transaksi = txtnomut
    RST!ref = txtnolot
    RST!KODE_SUPORCUST = ""
    RST!TYPE_BARANG = "BAHAN BAKU"
    RST!GROUP_BARANG = ""
    RST!kode_barang = lblkdbasebb 'kode bahan baku base
    RST!LOT_NUMBER = txtnolot
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
    RST!nolot = txtnolot
    RST!kodebahan = lblkdbasebb 'kode bahan baku base
    RST!qtybahan = txtqtybb
    RST!kodesatuan = "002"
    RST!hpp = hppkg * txtqtybb
    RST!flag = "0"
    RST.Update
    
    'Update base wip /wip jadi agar tidak masuk gudang pusat
    SQL = "Update list_hpp_produksi set flag = '2' Where nolot='" & txtnolot & "' and kodebarang= '" & lblkodebase & "'"
    Set RST = OBJ.Execute(SQL)

    OBJ.Close
    
    MsgBox "Data berhasil disimpan.", vbInformation, AppName
    clearform
End Sub

Private Sub clearform()
    txtnolot = ""
    lblkodebase = ""
    lblkdbasebb = ""
    txtnmbase = ""
    txtbaseBB = ""
    txtqty = ""
    txtqtybb = ""
    lblhpp_kilo = "0"
    txtnomut = ""
    Date1 = Date
End Sub

Private Sub cmdWIP_Click()
    If txtnolot = "" Then Exit Sub
    If optbase.Value = True Then
        namatabel = "Base"
        carisql1 = "Select distinct a.kode_bahan,b.NamaBarang,SUM(a.qty_bahan)'qty' from list_produksi_hasil a"
        carisql1 = carisql1 + " inner join am_itemmst b on a.kode_bahan = b.KodeBarang"
        carisql1 = carisql1 + " Where a.nolot= '" & txtnolot & "' "
        frmsearch.Show vbModal
    ElseIf optjadi.Value = True Then
        namatabel = "WIP Jadi"
        carisql1 = "Select distinct a.kode_bahan,b.NamaBarang,a.qty_bahan'qty',a.noref from list_produksi_hasil a"
        carisql1 = carisql1 + " inner join am_itemmst b on a.kode_bahan = b.KodeBarang"
        carisql1 = carisql1 + " Where a.nolot= '" & txtnolot & "' and (a.kode_bahan like 'L98%' or a.kode_bahan like 'K98%') and a.proses_ke = '2'"
        frmsearch.Show vbModal
    End If
End Sub

Private Sub cmdWIP_GotFocus()
    If hasil = "" Then Exit Sub
    lblkodebase = hasil
    txtnmbase = hasil1
    txtqty = Format(hasil2, "##0.00")
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    carisql1 = ""
    openhppjadi
End Sub
Private Sub openhppjadi()
    If optjadi.Value = True Then
        OBJ.Open dsn
        SQL = "Select a.nolot,a.kode_bahan,b.kg,SUM(a.qty_bahan*b.kg)'kilo' From list_produksi_hasil a"
        SQL = SQL + " inner join (Select kodebarang,kodesatuan,tahun,case when SUBSTRING('" & txtnolot & "' ,3,1) ='A' or SUBSTRING('" & txtnolot & "' ,3,2) ='01' Then kg1"
        SQL = SQL + " when SUBSTRING('" & txtnolot & "' ,3,1)='B' or SUBSTRING('" & txtnolot & "' ,3,2) ='02' Then kg2"
        SQL = SQL + " when SUBSTRING('" & txtnolot & "' ,3,1)='C' or SUBSTRING('" & txtnolot & "' ,3,2) ='03' Then kg3"
        SQL = SQL + " when SUBSTRING('" & txtnolot & "' ,3,1)='D' or SUBSTRING('" & txtnolot & "' ,3,2) ='04' Then kg4"
        SQL = SQL + " when SUBSTRING('" & txtnolot & "' ,3,1)='E' or SUBSTRING('" & txtnolot & "' ,3,2) ='05' Then kg5"
        SQL = SQL + " when SUBSTRING('" & txtnolot & "' ,3,1)='F' or SUBSTRING('" & txtnolot & "' ,3,2) ='06' Then kg6"
        SQL = SQL + " when SUBSTRING('" & txtnolot & "' ,3,1)='G' or SUBSTRING('" & txtnolot & "' ,3,2) ='07' Then kg7"
        SQL = SQL + " when SUBSTRING('" & txtnolot & "' ,3,1)='H' or SUBSTRING('" & txtnolot & "' ,3,2) ='08' Then kg8"
        SQL = SQL + " when SUBSTRING('" & txtnolot & "' ,3,1)='J' or SUBSTRING('" & txtnolot & "' ,3,2) ='09' Then kg9"
        SQL = SQL + " when SUBSTRING('" & txtnolot & "' ,3,1)='K' or SUBSTRING('" & txtnolot & "' ,3,2) ='10' Then kg10"
        SQL = SQL + " when SUBSTRING('" & txtnolot & "' ,3,1)='L' or SUBSTRING('" & txtnolot & "' ,3,2) ='11' Then kg11"
        SQL = SQL + " when SUBSTRING('" & txtnolot & "' ,3,1)='M' or SUBSTRING('" & txtnolot & "' ,3,2) ='12' Then kg12"
        SQL = SQL + " End as kg from am_itemkg) b"
        SQL = SQL + " on a.kode_bahan =b.kodebarang and a.kode_satuan = b.kodesatuan"
        SQL = SQL + " Where a.nolot = '" & txtnolot & "' and a.proses_ke = '2' and b.tahun = '20' + LEFT('" & txtnolot & "',2)"
        SQL = SQL + " and a.kode_bahan='" & lblkodebase & "' group by a.nolot,a.kode_bahan,b.kg"
        Set RS = OBJ.Execute(SQL)
        If Not RS.EOF Then
            lblhpp_kilo = "hpp perkilo : " & Format((kgprice * RS!kg) + (losswip * RS!kg), "###,##0.00")
            hppkg = Format((kgprice * RS!kg) + (losswip * RS!kg), "###,##0.00")
        End If
        OBJ.Close
    End If
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

Private Sub Form_Load()
    Date1 = Date
End Sub

Private Sub optbase_Click()
    If optbase.Value = True Then
        cmdWIP.Caption = "Base WIP"
    End If
End Sub

Private Sub optjadi_Click()
    If optjadi.Value = True Then
        cmdWIP.Caption = "WIP Jadi"
    End If
End Sub


