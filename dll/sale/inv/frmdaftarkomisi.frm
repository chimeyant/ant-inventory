VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmdaftarkomisi 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Laporan Komisi Sales"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmdaftarkomisi.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtkode2 
      Appearance      =   0  'Flat
      DataField       =   "KodeArea"
      Height          =   285
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   9
      Top             =   1110
      Width           =   720
   End
   Begin VB.CommandButton cmdsearch2 
      Caption         =   "To Sales"
      Height          =   270
      Left            =   195
      TabIndex        =   8
      Top             =   1140
      Width           =   945
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "From Sales"
      Height          =   270
      Left            =   195
      TabIndex        =   7
      Top             =   825
      Width           =   945
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   360
      Left            =   3465
      TabIndex        =   6
      Top             =   1980
      Width           =   870
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "Preview"
      Height          =   360
      Left            =   2565
      TabIndex        =   5
      Top             =   1980
      Width           =   870
   End
   Begin VB.TextBox txtKode 
      Appearance      =   0  'Flat
      DataField       =   "KodeArea"
      Height          =   285
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   0
      Top             =   810
      Width           =   720
   End
   Begin Crystal.CrystalReport Crystal 
      Left            =   105
      Top             =   1845
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1455
      TabIndex        =   1
      Top             =   165
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
      Format          =   59572227
      CurrentDate     =   37426
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   1455
      TabIndex        =   2
      Top             =   495
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
      Format          =   59572227
      CurrentDate     =   37426
   End
   Begin TDBNumber6Ctl.TDBNumber txthari 
      Height          =   285
      Left            =   660
      TabIndex        =   11
      Top             =   1530
      Visible         =   0   'False
      Width           =   600
      _Version        =   65536
      _ExtentX        =   1058
      _ExtentY        =   503
      Calculator      =   "frmdaftarkomisi.frx":2372
      Caption         =   "frmdaftarkomisi.frx":2392
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmdaftarkomisi.frx":23FE
      Keys            =   "frmdaftarkomisi.frx":241C
      Spin            =   "frmdaftarkomisi.frx":245E
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "[###,##0.00];[-###,##0.00];0"
      EditMode        =   1
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "####0"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999999999
      MinValue        =   0
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
   Begin TDBNumber6Ctl.TDBNumber txtkomisi 
      Height          =   285
      Left            =   2400
      TabIndex        =   12
      Top             =   1530
      Visible         =   0   'False
      Width           =   765
      _Version        =   65536
      _ExtentX        =   1349
      _ExtentY        =   503
      Calculator      =   "frmdaftarkomisi.frx":2486
      Caption         =   "frmdaftarkomisi.frx":24A6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmdaftarkomisi.frx":2512
      Keys            =   "frmdaftarkomisi.frx":2530
      Spin            =   "frmdaftarkomisi.frx":2572
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,##0.00;;0"
      EditMode        =   1
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,##0.00"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999999999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   -65531
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.Label lblsales2 
      Caption         =   "Label7"
      Height          =   255
      Left            =   2280
      TabIndex        =   17
      Top             =   1200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblsales 
      Caption         =   "Label6"
      Height          =   255
      Left            =   2280
      TabIndex        =   16
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "% dari omzet"
      Height          =   255
      Left            =   3225
      TabIndex        =   15
      Top             =   1560
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label4 
      Caption         =   "Komisi "
      Height          =   255
      Left            =   1920
      TabIndex        =   14
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Hari"
      Height          =   255
      Left            =   1320
      TabIndex        =   13
      Top             =   1575
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "TOP"
      Height          =   255
      Left            =   195
      TabIndex        =   10
      Top             =   1575
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Start Date"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   180
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "To Date"
      Height          =   255
      Left            =   210
      TabIndex        =   4
      Top             =   510
      Width           =   975
   End
End
Attribute VB_Name = "frmdaftarkomisi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As ADODB.Recordset
Dim CMD As ADODB.Command
Dim param As ADODB.Parameter
Dim SQL As String

Private Sub cmdsearch_Click()
    carisql1 = "select kodesales, namasales,idupdate from am_salesman"
    namatabel = "Sales"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    txtKode = hasil
    lblsales = hasil1
    hasil = ""
    hasil1 = ""
    carisales
    date1.SetFocus
End Sub

Private Sub cmdadd_Click()
    If txtKode = "" Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
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
    crystal.DataFiles(0) = "Proc(am_daftarkomisi2)"
    crystal.ReportFileName = AppPath & "\reports\sale\inv\daftarkomisi2.rpt"
    crystal.ParameterFields(0) = "@namauser;" + nmuser + ";true"
    crystal.ParameterFields(1) = "@tanggal1 ;" + Format(date1, "yyyymmdd") + ";true"
    crystal.ParameterFields(2) = "@tanggal2 ;" + Format(date2, "yyyymmdd") + ";true"
    crystal.ParameterFields(3) = "@kode1;" & "0" & ";true"
    crystal.ParameterFields(4) = "@kode2;" & txtkomisi.Value & ";true"
    crystal.ParameterFields(5) = "@kode3;" & txtKode & ";true"
    crystal.ParameterFields(6) = "@kode4;" & txtkode2 & ";true"
    crystal.RetrieveDataFiles
    crystal.Action = 1
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdsearch2_Click()
    carisql1 = "select kodesales, namasales, idupdate from am_salesman"
    namatabel = "Sales"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch2_GotFocus()
    If hasil = "" Then Exit Sub
    txtkode2 = hasil
    lblsales2 = hasil1
    hasil = ""
    hasil1 = ""
    carisales2
    date1.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    date1.Value = Date
    date2.Value = Date
End Sub

Private Sub txtkode_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then date1.SetFocus
End Sub

Private Sub txtkode_LostFocus()
    If txtKode = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from am_salesman where kodesales = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Data not found.", vbExclamation, "Warning"
        txtKode = ""
        txtKode.SetFocus
    End If
    OBJ.Close
End Sub
Private Sub carisales()
    If txtKode = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "SELECT * FROM AM_Salesman WHERE KodeSales = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
        '- 0 = sales non aktif -
    If RST!idupdate = "0" Then
        MsgBox "Salesman " & lblsales & " is not active !", vbExclamation, "Warning"
        txtKode = ""
        lblsales = ""
    End If
    OBJ.Close
End Sub

Private Sub carisales2()
    If txtkode2 = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "SELECT * FROM AM_Salesman WHERE KodeSales = '" & txtkode2 & "'"
    Set RST = OBJ.Execute(SQL)
        '- 0 = sales non aktif -
    If RST!idupdate = "0" Then
        MsgBox "Salesman " & lblsales2 & " is not active !", vbExclamation, "Warning"
        txtkode2 = ""
        lblsales2 = ""
    End If
    OBJ.Close
End Sub
