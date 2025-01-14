VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmmutasipiutang 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan Mutasi Piutang"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5610
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   5610
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "by Area"
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
      Left            =   270
      TabIndex        =   11
      Top             =   480
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "by Customer"
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
      Left            =   285
      TabIndex        =   10
      Top             =   180
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox txtcust1 
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
      Left            =   1395
      MaxLength       =   10
      TabIndex        =   1
      Top             =   1425
      Width           =   1215
   End
   Begin VB.TextBox txtcust 
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
      Left            =   1395
      MaxLength       =   10
      TabIndex        =   0
      Top             =   1065
      Width           =   1215
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   2355
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
      MICON           =   "frmmutasipiutang.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdview 
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   2355
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Preview"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmmutasipiutang.frx":031A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch 
      Height          =   285
      Left            =   210
      TabIndex        =   4
      Top             =   1065
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "From Code"
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
      MICON           =   "frmmutasipiutang.frx":0634
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch1 
      Height          =   285
      Left            =   225
      TabIndex        =   5
      Top             =   1440
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "To Code"
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
      MICON           =   "frmmutasipiutang.frx":094E
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1380
      TabIndex        =   6
      Top             =   1785
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
      Format          =   93847555
      CurrentDate     =   37845
   End
   Begin Crystal.CrystalReport Crystal 
      Left            =   315
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lblnmcust1 
      BackColor       =   &H8000000E&
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
      Left            =   2700
      TabIndex        =   9
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label lblnmcust 
      BackColor       =   &H8000000E&
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
      Left            =   2715
      TabIndex        =   8
      Top             =   1080
      Width           =   2640
   End
   Begin VB.Label Label3 
      Caption         =   "To  D a t e"
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
      Left            =   270
      TabIndex        =   7
      Top             =   1815
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      FillColor       =   &H8000000F&
      Height          =   1290
      Left            =   105
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   5415
   End
End
Attribute VB_Name = "frmmutasipiutang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdsearch_Click()
    If Option1.Value = True And Option2.Value = False Then
        carisql1 = "select kodecust, namacust, alamatcust, kota from am_customer"
        namatabel = "Customer"
    ElseIf Option1.Value = False And Option2.Value = True Then
        carisql1 = "select kode, nama from am_area"
        namatabel = "Area"
    End If
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    txtcust = hasil
    lblnmcust = hasil1
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdsearch1_Click()
    If Option1.Value = True And Option2.Value = False Then
        carisql1 = "select kodecust, namacust, alamatcust, kota from am_customer"
        namatabel = "Customer"
    ElseIf Option1.Value = False And Option2.Value = True Then
        carisql1 = "select kode, nama from am_area"
        namatabel = "Area"
    End If
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtcust1 = hasil
    lblnmcust1 = hasil1
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdview_Click()
    If txtcust = "" Or txtcust1 = "" Then Exit Sub
    If txtcust1 < txtcust Then
        MsgBox "To statement Can Not Smaller Then From statement.", vbOKOnly + vbExclamation, "Warning"
        txtcust1 = ""
        txtcust = ""
        txtcust.SetFocus
        Exit Sub
    End If

    Crystal.Reset
    Crystal.WindowState = crptMaximized
    Crystal.WindowShowCloseBtn = True
    Crystal.WindowShowPrintSetupBtn = True
    Crystal.WindowShowPrintBtn = True
    Crystal.WindowShowSearchBtn = True
    Crystal.WindowShowRefreshBtn = True
    Crystal.Connect = dsnreport
    
    If Option1.Value = True Then
        Crystal.ReportFileName = AppPath & "\reports\finance\sale\piutang_mutasi.rpt"
        Crystal.DataFiles(0) = "Proc(am_piutang_mutasi)"
    ElseIf Option2.Value = True Then
        Crystal.ReportFileName = AppPath & "\reports\finance\sale\piutang_mutasi_area.rpt"
        Crystal.DataFiles(0) = "Proc(am_piutang_mutasi_area)"
    End If
    
    Crystal.ParameterFields(0) = "@namauser;" + nmuser + ";true"
    Crystal.ParameterFields(1) = "@tanggal1 ;" + Format(date1, "yyyy-mm-dd") + ";true"
    Crystal.ParameterFields(2) = "@kode1;" & txtcust & ";true"
    Crystal.ParameterFields(3) = "@kode2;" & txtcust1 & ";true"
    
    Crystal.RetrieveDataFiles
    Crystal.Action = 1
End Sub

Private Sub Form_Load()
    date1 = Date
End Sub

Private Sub Option1_Click()
    txtcust = "": txtcust1 = ""
    lblnmcust = "": lblnmcust1 = ""
End Sub

Private Sub Option2_Click()
    txtcust = "": txtcust1 = ""
    lblnmcust = "": lblnmcust1 = ""
End Sub

Private Sub txtcust_Change()
    If Option1.Value = True And Option2.Value = False And Len(txtcust) = 7 Then
        OBJ.Open dsn
        SQL = "select namacust from am_customer Where KodeCust='" & txtcust & "'"
        Set RST = OBJ.Execute(SQL)
        If RST.EOF Then OBJ.Close: lblnmcust = "": Exit Sub
        lblnmcust = RST!namacust
        OBJ.Close
    ElseIf Option2.Value = True And Option1.Value = False And Len(txtcust) = 2 Then
        OBJ.Open dsn
        SQL = "select nama from am_area Where kode='" & txtcust & "'"
        Set RST = OBJ.Execute(SQL)
        If RST.EOF Then OBJ.Close: lblnmcust = "": Exit Sub
        lblnmcust = RST!nama
        OBJ.Close
    End If
End Sub

Private Sub txtcust1_Change()
    If Option1.Value = True And Option2.Value = False And Len(txtcust1) = 7 Then
        OBJ.Open dsn
        SQL = "select namacust from am_customer Where KodeCust='" & txtcust1 & "'"
        Set RST = OBJ.Execute(SQL)
        If RST.EOF Then OBJ.Close: lblnmcust1 = "": Exit Sub
        lblnmcust1 = RST!namacust
        OBJ.Close
    ElseIf Option2.Value = True And Option1.Value = False And Len(txtcust1) = 2 Then
        OBJ.Open dsn
        SQL = "select nama from am_area Where kode='" & txtcust1 & "'"
        Set RST = OBJ.Execute(SQL)
        If RST.EOF Then OBJ.Close: lblnmcust1 = "": Exit Sub
        lblnmcust1 = RST!nama
        OBJ.Close
    End If
End Sub
