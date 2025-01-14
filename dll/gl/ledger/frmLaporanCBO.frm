VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmLaporanCBO 
   Caption         =   "Print Cash Bank Out"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkppn 
      Caption         =   "Ppn"
      Height          =   255
      Left            =   5640
      TabIndex        =   13
      Top             =   2280
      Value           =   1  'Checked
      Width           =   975
   End
   Begin XtremeSuiteControls.CheckBox ChkBB 
      Height          =   255
      Left            =   5640
      TabIndex        =   11
      Top             =   480
      Width           =   1215
      _Version        =   851970
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Bahan Baku"
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
   Begin VB.ListBox List2 
      Height          =   3660
      ItemData        =   "frmLaporanCBO.frx":0000
      Left            =   120
      List            =   "frmLaporanCBO.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   510
      Width           =   5415
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Preview"
      Height          =   255
      Left            =   5640
      TabIndex        =   1
      Top             =   2670
      Value           =   1  'Checked
      Width           =   975
   End
   Begin Chameleon.chameleonButton cmdnotprint 
      Height          =   495
      Left            =   5640
      TabIndex        =   0
      Top             =   1710
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Show"
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
      MICON           =   "frmLaporanCBO.frx":0004
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Crystal.CrystalReport Crystal 
      Left            =   6675
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   5640
      TabIndex        =   3
      Top             =   3750
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
      MICON           =   "frmLaporanCBO.frx":031E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdprint 
      Height          =   615
      Left            =   5640
      TabIndex        =   4
      Top             =   3030
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "Preview Print"
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
      MICON           =   "frmLaporanCBO.frx":0638
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
      Height          =   285
      Left            =   2400
      TabIndex        =   5
      Top             =   150
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
      Format          =   143851523
      CurrentDate     =   37426
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   4560
      TabIndex        =   6
      Top             =   150
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
      Format          =   143851523
      CurrentDate     =   37426
   End
   Begin Chameleon.chameleonButton cmdsubmit 
      Height          =   495
      Left            =   5640
      TabIndex        =   7
      Top             =   1080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Show"
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
      MICON           =   "frmLaporanCBO.frx":0952
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtnil1 
      Height          =   255
      Left            =   5655
      TabIndex        =   8
      Top             =   990
      Visible         =   0   'False
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   450
      Calculator      =   "frmLaporanCBO.frx":0C6C
      Caption         =   "frmLaporanCBO.frx":0C8C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmLaporanCBO.frx":0CF8
      Keys            =   "frmLaporanCBO.frx":0D16
      Spin            =   "frmLaporanCBO.frx":0D58
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##,###,##0.0000;(##,###,##0.0000);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##,###,###,##0.0000;(##,###,###,##0.0000)"
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
   Begin TDBNumber6Ctl.TDBNumber txtnil2 
      Height          =   255
      Left            =   5640
      TabIndex        =   9
      Top             =   1230
      Visible         =   0   'False
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   450
      Calculator      =   "frmLaporanCBO.frx":0D80
      Caption         =   "frmLaporanCBO.frx":0DA0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmLaporanCBO.frx":0E0C
      Keys            =   "frmLaporanCBO.frx":0E2A
      Spin            =   "frmLaporanCBO.frx":0E6C
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "#,###,###,##0.0000;(#,###,###,##0.0000);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "#,###,###,##0.0000;(#,###,###,##0.0000)"
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
   Begin XtremeSuiteControls.CheckBox ChkNonBB 
      Height          =   255
      Left            =   5640
      TabIndex        =   12
      Top             =   720
      Width           =   1455
      _Version        =   851970
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Non Bahan Baku"
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
   Begin VB.Label Label1 
      Caption         =   "Display Cash Bank Out from                                              to"
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
      TabIndex        =   10
      Top             =   195
      Width           =   4335
   End
End
Attribute VB_Name = "frmLaporanCBO"
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

Dim i As Integer

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdnotprint_Click()
    List2.Clear
    If date1 > date2 Then
        MsgBox "Invalid date range, search abort.", vbExclamation, "Error"
        Exit Sub
    End If
    OBJ.Open dsn
    SQL = "Select distinct a.*,b.tgltrx From no_bank_payment a inner join gl_transaksi b "
    SQL = SQL + " on a.notrx=b.notrx Where a.ref = 'B' and b.tgltrx >= '" & tanggal1 & "' "
    SQL = SQL + "and b.tgltrx <= '" & tanggal2 & "' and b.dbkrtrx ='K' Order By a.no_voucher"
    
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        List2.AddItem RST!no_voucher
        
        RST.MoveNext
    Loop
    OBJ.Close
    
End Sub

Private Sub cmdprint_Click()
    Dim nilai_rupiah As Double
    Dim nilai_ppn As Double
    Dim nilai_hutang As Double
    Dim kode_kurs As String
    Dim nilai_kurs As Long
    Dim nilai_jumlah As Long
    Dim nilai_potongan As Long
    Dim total As Long 'As Double
    Dim str1, str2, str3, str4, str5, str6, str7, str8, str9, str10 As String
'------------------------------------------------
    If ChkBB = xtpChecked And ChkNonBB = xtpUnchecked Then
    'CETAK BAHAN BAKU
        If chkppn.Value = Unchecked Then
            For i = 0 To List2.ListCount - 1
            If List2.Selected(i) = True Then
                OBJ.Open dsn
                
                SQL = "SELECT a.NoApply,a.nilaikurs,a.Amount,a.Selisih,a.potongan,(a.PPN * a.nilaikurs) AS nilaippn,a.kodecur, a.TransType, a.Amount - a.Potongan + a.PPN AS jumlah "
                SQL = SQL + "From am_apopnfil a inner join am_beliapp b on b.NoBeli = a.NoBeli "
                SQL = SQL + "Where b.ref1 = '" + List2.List(i) + "'"
                
                Set RST = OBJ.Execute(SQL)
                Do While Not RST.EOF
                    kode_kurs = RST!kodecur
                    nilai_kurs = RST!nilaikurs
                    nilai_jumlah = RST!amount
                    nilai_ppn = RST!nilaippn
                    nilai_potongan = RST!potongan
                    nilai_hutang = RST!jumlah
                    RST.MoveNext
                Loop
    
                SQL = "Select SUM(Qty * Price * nilaikurs) as Jml  From am_beliapp Where Ref1 = '" + List2.List(i) + "'"
                Set RST = OBJ.Execute(SQL)
                If Not RST.EOF Then
                    total = Format(RST!jml, "###,##0.00")
                End If
                'OBJ.Close
     
                SQL = "Select distinct a.*,b.Kodecur,b.Nilaikurs,c.tglbkt from no_bank_payment a inner join am_beliapp b on a.no_voucher = b.Ref1 "
                SQL = SQL + "inner join am_apcashhdr c on a.no_payment = c.NoBkt Where a.no_voucher = '" + List2.List(i) + "'"
                Set RST = OBJ.Execute(SQL)
                        str1 = RST!kpd
                       ' str2 = RST!cekbg
                        str3 = Format(RST!tgljt, "dd/MM/yyyy")
                        str4 = RST!no_payment
                        str5 = RST!no_voucher
                        str6 = Format(RST!tglbkt, "dd/MM/yyyy")
     
                SQL = "Select  a.*, b.namabarang ,d.namasatuan ,SUM(a.qty * a.price) as jumlah,c.noapply"
                SQL = SQL + " from am_beliapp as a inner join am_apitemmst as b on b.kodebarang= a.kodebarang"
                SQL = SQL + " inner join am_apopnfil c on c.nobeli=a.nobeli"
                SQL = SQL + " inner join am_apunit as d on d.kodesatuan=a.kodesatuan"
                SQL = SQL + " Where a.ref1 = '" + List2.List(i) + "'"
                SQL = SQL + " GROUP BY a.NoBeli,a.TglBeli,a.NoPO,a.Ref1,a.Ref2,a.Kodesupp,a.Kodecur,a.Nilaikurs,"
                SQL = SQL + " a.KodeBarang,a.Qty,a.Price,a.KodeSatuan,a.ppn,a.keterangan,a.keterangan2,a.keterangan3,"
                SQL = SQL + " a.keterangan4,a.LineItem,a.flag1,a.flag2,a.amount,a.id,b.NamaBarang,d.NamaSatuan,c.NoApply"
                
                With rptBB
                    .Field17 = str1
                    .Field18 = str2
                    .Field22 = date2
                    .Field19 = str4
                    .Field20 = str5
                    .Field21 = date1
                    '.Field31 = txtcash
                    '.Field26 = txtketcash
                    .Field32 = total
                    .Field23 = total
                    .Field27 = total
                    '.Field28 = total
                    .lblkurs = kode_kurs
                    .lblnilaikurs = Format(nilai_kurs, "###,###,##0.00")
                    .DataControl1.Source = SQL
                    .DataControl1.ConnectionString = dsn
                    .Show vbModal
                End With
                OBJ.Close
            End If
            Next i
        Else
        'BAHAN BAKU + PPN
            For i = 0 To List2.ListCount - 1
            If List2.Selected(i) = True Then
                OBJ.Open dsn
                
                SQL = "SELECT a.NoApply,a.nilaikurs,a.Amount,a.Selisih,a.potongan,(a.PPN * a.nilaikurs) AS nilaippn,a.kodecur, a.TransType, a.Amount - a.Potongan + a.PPN AS jumlah "
                SQL = SQL + "From am_apopnfil a inner join am_beliapp b on b.NoBeli = a.NoBeli "
                SQL = SQL + "Where b.ref1 = '" + List2.List(i) + "'"
                
                Set RST = OBJ.Execute(SQL)
                Do While Not RST.EOF
                    kode_kurs = RST!kodecur
                    'nilai_kurs = RST!nilaikurs
                    nilai_jumlah = RST!amount
                    nilai_ppn = RST!nilaippn
                    nilai_potongan = RST!potongan
                    nilai_hutang = RST!jumlah '+ RST!nilaippn   (1. Berbeda dengan yg langsung print)
                    str7 = RST!noapply
                    RST.MoveNext
                Loop
    
                SQL = "Select SUM(Qty * Price) as Jml  From am_beliapp Where Ref1 = '" + List2.List(i) + "'"
                Set RST = OBJ.Execute(SQL)
                If Not RST.EOF Then
                    total = RST!jml
                End If
                
                SQL = "Select nilaikurs From am_apopnfil Where NoApply = '" & str7 & "' and TransType = 'PM'"
                Set RST = OBJ.Execute(SQL)
                If Not RST.EOF Then
                    nilai_kurs = RST!nilaikurs
                End If
                
                SQL = "Select distinct a.*,b.Kodecur,b.Nilaikurs,c.tglbkt from no_bank_payment a inner join am_beliapp b on a.no_voucher = b.Ref1 "
                SQL = SQL + "inner join am_apcashhdr c on a.no_payment = c.NoBkt Where a.no_voucher = '" + List2.List(i) + "'"
                Set RST = OBJ.Execute(SQL)
      
                        str1 = RST!kpd
    '                    str2 = RST!cekbg
                        str3 = Format(RST!tgljt, "dd/MM/yyyy")
                        str4 = RST!no_payment
                        str5 = RST!no_voucher
                        str6 = Format(RST!tglbkt, "dd/MM/yyyy")
                OBJ.Close
                SQL = "Select  a.*, b.namabarang ,d.namasatuan ,(SUM((a.qty) * a.price)) as jumlah,c.noapply"
                SQL = SQL + " from am_beliapp as a inner join am_apitemmst as b on b.kodebarang= a.kodebarang"
                SQL = SQL + " inner join am_apopnfil c on c.nobeli=a.nobeli"
                SQL = SQL + " inner join am_apunit as d on d.kodesatuan=a.kodesatuan"
                SQL = SQL + " Where a.ref1 = '" + List2.List(i) + "'"
                SQL = SQL + " GROUP BY a.NoBeli,a.TglBeli,a.NoPO,a.Ref1,a.Ref2,a.Kodesupp,a.Kodecur,a.Nilaikurs,"
                SQL = SQL + " a.KodeBarang,a.Qty,a.Price,a.KodeSatuan,a.ppn,a.keterangan,a.keterangan2,a.keterangan3,"
                SQL = SQL + " a.keterangan4,a.LineItem,a.flag1,a.flag2,a.amount,a.id,b.NamaBarang,d.NamaSatuan,c.NoApply"
                
                With rptBBPPn
                    .Field17 = str1
                    .Field18 = str2
                    .Field22 = str3
                    .Field19 = str4
                    .Field20 = str5
                    .Field21 = str6
                    '.Field31 = frmoutran.txtcash
                    '.Field26 = frmoutran.txtketcash
                    .Field27 = total
                    .lbljumlah = Format(total, "###,###,##0.00")
                    .lblppn = Format(nilai_ppn, "###,###,##0.00")
                    .lblpotongan = Format(nilai_potongan, "###,###,##0.00")
                    .lblhutang = Format((total + nilai_ppn - nilai_potongan), "###,###,##0.00")
                    .lblkurs = kode_kurs
                    .lblnilaikurs = Format(nilai_kurs, "###,###,##0.00")
                    .DataControl1.Source = SQL
                    .DataControl1.ConnectionString = dsn
                    .Show vbModal
                End With
            'OBJ.Close
            End If
            Next i
        End If
    ElseIf ChkNonBB = xtpChecked And ChkBB = xtpUnchecked Then
    'CETAK NON BAHAN BAKU
        For i = 0 To List2.ListCount - 1
        If List2.Selected(i) = True Then
            OBJ.Open dsn
                SQL = "select a.ppn,a.nilai,sum(b.jumlah) as jml "
                SQL = SQL + "from am_voucherhdr a inner join am_voucherin b "
                SQL = SQL + "On a.novoucher =b.novoucher Where a.novoucher='" + List2.List(i) + "' "
                SQL = SQL + "group By a.ppn,a.nilai"
                Set RST = OBJ.Execute(SQL)
                    
                    nilai_rupiah = RST!jml * RST!nilai
                    If RST!ppn > 0 Then
                        nilai_ppn = nilai_rupiah * (RST!ppn / 100)
                    End If
                    nilai_hutang = nilai_rupiah + nilai_ppn
            
                SQL = "Select distinct a.*,b.* From gl_transaksi a inner join no_bank_payment b on a.notrx = b.notrx "
                SQL = SQL + "Where b.no_voucher='" + List2.List(i) + "' and a.dbkrtrx = 'K'"
                Set RST = OBJ.Execute(SQL)
                    
                    str1 = RST!kpd
                    str2 = RST!cekbg
                    str3 = Format(RST!tgljt, "dd/MM/yyyy")
                    str4 = RST!no_payment
                    str5 = RST!no_voucher
                    str6 = Format(RST!dateentry, "dd/MM/yyyy")
                    str7 = RST!desctrx
                    str8 = RST!noactrx
                    str9 = RST!currtrx
                    str10 = RST!kurs
                
                SQL = "Select a.*,b.tgl,b.nilai,b.ppn,(a.jumlah * nilai) as jml From am_voucherin a inner join am_voucherhdr b "
                SQL = SQL + "On a.novoucher=b.novoucher Where a.novoucher='" + List2.List(i) + "'"
                With rptNONBB
                    .Field17 = str1
                    .Field18 = str2
                    .Field22 = str3
                    .Field19 = str4
                    .Field20 = str5
                    .Field21 = str6
                    .Field26 = str7
                    .Field27 = Format(nilai_rupiah, "###,###,##0.00")
                    .Field31 = str8
                    .lblkurs = str9
                    .lblnilaikurs = str10
                    .lblppn = Format(nilai_ppn, "###,###,##0.00")
                    .lbljumlah = Format(nilai_rupiah, "###,###,##0.00")
                    .lblhutang = Format(nilai_hutang, "###,###,##0.00")
                    .DataControl1.Source = SQL
                    .DataControl1.ConnectionString = dsn
                    .Show vbModal
                End With
              OBJ.Close
        End If
        Next i
    End If
End Sub

Private Sub cmdsubmit_Click()
    List2.Clear
    If date1 > date2 Then
        MsgBox "Invalid date range, search abort.", vbExclamation, "Error"
        Exit Sub
    End If
    If ChkBB = xtpChecked And ChkNonBB = xtpUnchecked Then
'TAMPILKAN VOUCHER BAHAN BAKU
            OBJ.Open dsn
            SQL = "Select distinct a.*,c.tglbkt from no_bank_payment a inner join am_beliapp b on a.no_voucher = b.Ref1 "
            SQL = SQL + "inner join am_apcashhdr c on a.no_payment = c.NoBkt Where a.ref = 'P' and a.flag = '0' and c.TglBkt >= '" & tanggal1 & "' "
            SQL = SQL + "and c.TglBkt <= '" & tanggal2 & "' order by a.no_voucher"
            
            Set RST = OBJ.Execute(SQL)
                Do While Not RST.EOF
                    List2.AddItem RST!no_voucher
                    RST.MoveNext
                Loop
            OBJ.Close
    ElseIf ChkNonBB = xtpChecked And ChkBB = xtpUnchecked Then
'TAMPILKAN VOUCHER NON BAHAN BAKU
            OBJ.Open dsn
            SQL = "Select distinct a.*,b.tgltrx From no_bank_payment a inner join gl_transaksi b "
            SQL = SQL + " on a.notrx=b.notrx Where a.ref = 'P' and a.flag = '0' and b.tgltrx >= '" & tanggal1 & "' "
            SQL = SQL + "and b.tgltrx <= '" & tanggal2 & "' and b.dbkrtrx ='K' Order By a.no_voucher"
            
            Set RST = OBJ.Execute(SQL)
                Do While Not RST.EOF
                    List2.AddItem RST!no_voucher
                    RST.MoveNext
                Loop
            OBJ.Close
    End If
End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='115' and b.kodeuser = '2" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If RST.EOF Then
    '        MsgBox "User Rights Denied !!" & vbCrLf & _
    '        "Please contact your Administrator.", vbCritical, "User Rights"
    '        OBJ.Close
    '        Unload Me
    '        Exit Sub
    '    End If
    '    OBJ.Close
    'End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    date1 = Date
    date2 = Date
End Sub

Function tanggal1()
    tanggal1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggal2()
    tanggal2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function

