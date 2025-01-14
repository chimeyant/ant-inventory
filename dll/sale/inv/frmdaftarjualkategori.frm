VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmdaftarjualkategori 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan Penjualan By Kategori"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtkategori 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   600
      MaxLength       =   20
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox txtkode2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3240
      MaxLength       =   20
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtkode1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3240
      MaxLength       =   20
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtsales2 
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
      MaxLength       =   20
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtsales1 
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
      MaxLength       =   20
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   2520
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
      MICON           =   "frmdaftarjualkategori.frx":0000
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
      Left            =   1320
      TabIndex        =   3
      Top             =   1560
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
      Format          =   134807555
      CurrentDate     =   37464
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   1920
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
      Format          =   134807555
      CurrentDate     =   37464
   End
   Begin Chameleon.chameleonButton cmdsales1 
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "From Sales"
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmdaftarjualkategori.frx":031A
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsales2 
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "To Sales"
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmdaftarjualkategori.frx":0634
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdview 
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      Top             =   2520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "View"
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
      MICON           =   "frmdaftarjualkategori.frx":094E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Crystal.CrystalReport crystal 
      Left            =   240
      Top             =   2520
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
   Begin MSForms.ComboBox cmbtypex 
      Height          =   285
      Left            =   1320
      TabIndex        =   10
      Top             =   1200
      Width           =   1815
      VariousPropertyBits=   746608667
      MaxLength       =   2
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3201;503"
      ListRows        =   11
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      Caption         =   "Ketegori"
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
      TabIndex        =   9
      Top             =   1230
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "From Date"
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
      TabIndex        =   8
      Top             =   1590
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "To Date"
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
      TabIndex        =   7
      Top             =   1950
      Width           =   855
   End
End
Attribute VB_Name = "frmdaftarjualkategori"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim str1 As String

Private Sub cmbtypex_Change()

    OBJ.Open dsn
    SQL = "Select TOP 1 * from list_kategoriProduk Where kdkatproduk='" & Left(cmbtypex, 3) & "' order by kodebarang asc"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtkode1 = RST!kodebarang
        txtkategori = RST!kategoriproduk
    End If
    
    SQL = "Select TOP 1 * from list_kategoriProduk Where kdkatproduk='" & Left(cmbtypex, 3) & "' order by kodebarang desc"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtkode2 = RST!kodebarang
    End If
    
    OBJ.Close
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdsales1_Click()
    carisql1 = "select kodesales, namasales,idupdate from AM_salesman"
    namatabel = "Sales"
    frmsearch.Show vbModal
End Sub

Private Sub cmdsales1_GotFocus()
    If hasil = "" Then Exit Sub
    txtsales1 = hasil
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdsales2_Click()
    carisql1 = "select kodesales, namasales,idupdate from AM_salesman"
    namatabel = "Sales"
    frmsearch.Show vbModal
End Sub

Private Sub cmdsales2_GotFocus()
    If hasil = "" Then Exit Sub
    txtsales2 = hasil
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdview_Click()
    If txtsales1 = "" Or txtsales2 = "" Then Exit Sub
    If txtsales1 > txtsales2 Then
        MsgBox "To statement Can Not Smaller Then From statement.", vbOKOnly + vbExclamation, "Warning"
        txtsales1 = ""
        txtsales1.SetFocus
        Exit Sub
    End If
    If cmbtypex = "" Then
        MsgBox "Kategori column is empty", vbExclamation, "Warning"
        cmbtypex.SetFocus
        Exit Sub
    End If
    If date2 < date1 Then
        MsgBox "To Date Can Not Smaller Then From Date.", vbExclamation, "Warning"
        date2.SetFocus
        Exit Sub
    End If
    
    crystal.Reset
    crystal.WindowState = crptMaximized
    crystal.WindowShowCloseBtn = True
    crystal.WindowShowPrintSetupBtn = True
    crystal.WindowShowSearchBtn = True
    crystal.WindowShowRefreshBtn = True
    crystal.Connect = dsnreport
    crystal.DataFiles(0) = "Proc(am_daftarjualkategori)"
    crystal.ReportFileName = AppPath & "\reports\sale\inv\jsales_kategori.rpt"
    
    crystal.ParameterFields(0) = "@tanggal1;" & Format(date1, "yyyyMMdd") & ";true"
    crystal.ParameterFields(1) = "@tanggal2;" & Format(date2, "yyyyMMdd") & ";true"
    crystal.ParameterFields(2) = "@sales1 ;" + txtsales1 + ";true"
    crystal.ParameterFields(3) = "@sales2 ;" + txtsales2 + ";true"
    crystal.ParameterFields(4) = "@pilih ;" + str1 + ";true"
    crystal.ParameterFields(5) = "@kode1 ;" + txtkode1 + ";true"
    crystal.ParameterFields(6) = "@kode2 ;" + txtkode2 + ";true"
    crystal.ParameterFields(7) = "@kategori ;" + txtkategori + ";True"
    crystal.ParameterFields(8) = "@namauser ;" + nmuser + ";true"
    crystal.RetrieveDataFiles
    crystal.Action = 1
End Sub

Private Sub Form_Load()
    date1 = Date
    date2 = Date
    
    str1 = "dela"
    cmbtypex.Clear
    cmbtypex.ColumnCount = 1
    
    OBJ.Open dsn
    SQL = "Select distinct kdkatproduk,kategoriproduk From list_kategoriProduk"
    Set RST = OBJ.Execute(SQL)

    If Not RST.EOF Then
        Do While Not RST.EOF
            cmbtypex.AddItem RST!kdkatproduk & " - " & RST!kategoriproduk
            RST.MoveNext
        Loop
    End If
    OBJ.Close
End Sub
