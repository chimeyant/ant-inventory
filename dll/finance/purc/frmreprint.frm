VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmreprint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reprint Voucher Penerimaan"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3615
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtalamat 
      Height          =   345
      Left            =   3810
      TabIndex        =   7
      Top             =   990
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox txtcurr 
      Height          =   345
      Left            =   3795
      TabIndex        =   6
      Top             =   585
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox txtnobukti 
      Height          =   345
      Left            =   3795
      TabIndex        =   4
      Top             =   180
      Width           =   2895
   End
   Begin VB.TextBox txtvoucher 
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
      Left            =   1845
      TabIndex        =   3
      Top             =   225
      Width           =   1710
   End
   Begin XtremeSuiteControls.PushButton cmdclose 
      Height          =   420
      Left            =   2565
      TabIndex        =   0
      Top             =   1395
      Width           =   975
      _Version        =   851970
      _ExtentX        =   1720
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Close"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   5
   End
   Begin XtremeSuiteControls.PushButton cmdview 
      Height          =   420
      Left            =   1530
      TabIndex        =   1
      Top             =   1395
      Width           =   975
      _Version        =   851970
      _ExtentX        =   1720
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "View"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   5
   End
   Begin Chameleon.chameleonButton cmdsearch0 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Voucher"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmreprint.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   3795
      TabIndex        =   8
      Top             =   1365
      Visible         =   0   'False
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
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   134479875
      CurrentDate     =   37694
   End
   Begin VB.TextBox txtsup 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Height          =   495
      Left            =   195
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   3270
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   645
      Left            =   90
      Top             =   645
      Width           =   3465
   End
End
Attribute VB_Name = "frmreprint"
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

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdsearch0_Click()
    'carisql1 = "Select distinct Ref1'Voucher',NoBeli'NoLPB',Ref2'NoApply' From am_beliapp Where Ref1 <> ''"
    
    carisql1 = "Select distinct a.Ref1'Voucher',a.Ref2'NoApply',a.NoBeli'NoBPB',b.NamaSupp from am_beliapp a inner join am_supplier b on a.Kodesupp = b.KodeSupp"
    carisql1 = carisql1 + " Where Ref1 <>''"
    namatabel = "Voucher Penerimaan"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch0_GotFocus()
    If hasil = "" Then Exit Sub
    txtnobukti = hasil1
    txtvoucher = hasil
    txtsup = hasil2
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    opendata
End Sub

Private Sub opendata()
    OBJ.Open dsn
    SQL = "Select * From am_voucherhdr Where novoucher = '" & txtvoucher & "'"
    Set RST = OBJ.Execute(SQL)
    
    If Not RST.EOF Then
        txtsup = RST!kepada
        txtcurr = RST!kdkurs
        txtalamat = RST!alamat
        date2 = RST!tgl
    End If
    OBJ.Close
End Sub
Private Sub cmdview_Click()
    Dim kode_kurs As String
    Dim nilai_kurs As Long
    Dim nilai_jumlah As Long
    Dim nilai_ppn As Long
    Dim nilai_potongan As Long
    Dim nilai_hutang As Long
    
    SQL = "SELECT NoApply,nilaikurs,Amount,Selisih,potongan,(PPN * nilaikurs) AS nilaippn,kodecur, TransType, Amount - Potongan + PPN AS jumlah"
    SQL = SQL + " From am_apopnfil"
    SQL = SQL + " Where NoBeli='" + txtnobukti + "'"

    OBJ1.Open dsn
    Set RST = OBJ1.Execute(SQL)

    Do While Not RST.EOF
        kode_kurs = RST!kodecur
        nilai_kurs = RST!nilaikurs
        nilai_jumlah = RST!amount
        nilai_ppn = RST!nilaippn
        nilai_potongan = RST!potongan
        nilai_hutang = RST!jumlah
        RST.MoveNext
    Loop
    'cek novoucher jika ada keluarkan nomor payment
    SQL = "Select distinct a.nilaikurs From am_apopnfil a left join am_beliapp b on a.NoApply = b.ref2"
    SQL = SQL + " Where a.TransType = 'PM' and b.Ref1 = '" & txtvoucher & "'"
    Set RST = OBJ1.Execute(SQL)
    If Not RST.EOF Then
        nilai_kurs = RST!nilaikurs
    End If
    OBJ1.Close

    SQL = "Select  a.* ,b.namabarang,a.qty , d.namasatuan ,(a.qty * a.price) as jumlah,c.noapply "
    SQL = SQL + " from am_beliapp   as a inner join am_apitemmst as b on b.kodebarang= a.kodebarang "
    SQL = SQL + " inner join am_apopnfil c on c.nobeli= a.nobeli"
    SQL = SQL + " inner join am_apunit as d on d.kodesatuan=a.kodesatuan"
    SQL = SQL + " Where a.nobeli = '" + txtnobukti + "'"

    With rptvoucher
        .lblsupp = txtsup
        .lblnpwp = ""
        .lblkurs = kode_kurs
        .lblnilaikurs = Format(nilai_kurs, "###,###,##0.00")
        .lbljumlah = Format(nilai_jumlah, "###,###,##0.00")
        .lblppn = Format(nilai_ppn, "###,###,##0.00")
        .lblpotongan = Format(nilai_potongan, "###,###,##0.00")
        If txtcurr = "IDR" Then
            .lblhutang = Format(nilai_jumlah + nilai_ppn - nilai_potongan, "###,###,##0.00")
        Else
            .lblhutang = .lbljumlah
        End If
        
        .lblalamat = txtalamat
        .lblnovoucher = ": " + txtvoucher
        .lbltanggal = ": " + Format(date2, "dd/MM/yyyy")
        .DataControl1.Source = SQL
        .DataControl1.ConnectionString = dsn
        .WindowState = 2
        .Show
    End With
    Clearform
End Sub
Sub Clearform()
    txtvoucher = ""
    txtnobukti = ""
    txtsup = ""
    txtalamat = ""
    txtcurr = ""
    date2 = Date
End Sub
