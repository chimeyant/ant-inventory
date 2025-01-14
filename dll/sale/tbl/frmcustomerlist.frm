VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmcustomerlist 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List Customer"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Undefine Account Customer"
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
      TabIndex        =   4
      Top             =   960
      Width           =   2895
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Defined Account Customer"
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
      TabIndex        =   3
      Top             =   720
      Width           =   2775
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "With NPWP Only"
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
      Left            =   1920
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
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
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   975
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
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   1455
   End
   Begin Crystal.CrystalReport Crystal 
      Left            =   240
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox txtkode1 
      Appearance      =   0  'Flat
      DataField       =   "KodeArea"
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
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   6
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txtKode 
      Appearance      =   0  'Flat
      DataField       =   "KodeArea"
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
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   5
      Top             =   1440
      Width           =   1575
   End
   Begin Chameleon.chameleonButton cmdsearch 
      Height          =   285
      Left            =   240
      TabIndex        =   9
      Top             =   1440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "From Customer"
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
      MICON           =   "frmcustomerlist.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   2280
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
      MICON           =   "frmcustomerlist.frx":031A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdpreview 
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   2280
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
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmcustomerlist.frx":0634
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch1 
      Height          =   285
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "To Customer"
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
      MICON           =   "frmcustomerlist.frx":094E
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   -360
      TabIndex        =   11
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "frmcustomerlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdpreview_Click()
    If Option1.Value Or Option2.Value Then
        If txtKode = "" Then txtKode = "0"
        If txtkode1 = "" Then txtkode1 = "z"
            
        If txtkode1 < txtKode Then
            MsgBox "To... can not Smaller Then From...", vbExclamation, "Warning"
            txtkode1 = ""
            txtkode1.SetFocus
            Exit Sub
        End If
    End If
    
    Crystal.Reset
    Crystal.WindowState = crptMaximized
    Crystal.WindowShowCloseBtn = True
    Crystal.WindowShowPrintSetupBtn = True
    Crystal.WindowShowSearchBtn = True
    Crystal.WindowShowRefreshBtn = True
    Crystal.Connect = dsnreport
    If Option1.Value Or Option2.Value Then
        If Check1.Value = 0 Then
            If Option1.Value = True Then Crystal.DataFiles(0) = "Proc(am_listcustomer)"
            If Option1.Value = True Then Crystal.ReportFileName = AppPath & "\reports\sale\tbl\listcust.rpt"
            If Option2.Value = True Then Crystal.DataFiles(0) = "Proc(am_listcustomerx)"
            If Option2.Value = True Then Crystal.ReportFileName = AppPath & "\reports\sale\tbl\listcustx.rpt"
        Else
            If Option1.Value = True Then Crystal.DataFiles(0) = "Proc(am_listcustomer)"
            If Option1.Value = True Then Crystal.ReportFileName = AppPath & "\reports\sale\tbl\listcustnpwp.rpt"
            If Option2.Value = True Then Crystal.DataFiles(0) = "Proc(am_listcustomerx)"
            If Option2.Value = True Then Crystal.ReportFileName = AppPath & "\reports\sale\tbl\listcustxnpwp.rpt"
        End If
        Crystal.ParameterFields(0) = "@kode1;" + txtKode + ";true"
        Crystal.ParameterFields(1) = "@kode2;" + txtkode1 + ";true"
        Crystal.ParameterFields(2) = "@namauser;" + nmuser + ";true"
    Else
        Crystal.DataFiles(0) = "Proc(am_listcustomer_def)"
        Crystal.ReportFileName = AppPath & "\reports\sale\tbl\listcust_def.rpt"
        If Option3.Value Then Crystal.ParameterFields(0) = "@kode1;a;true"
        If Option4.Value Then Crystal.ParameterFields(0) = "@kode1;b;true"
    End If
    Crystal.RetrieveDataFiles
    Crystal.Action = 1
End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='65' and b.kodeuser = '1" & kuser & "'"
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

Private Sub Option1_Click()
    cmdsearch.Enabled = True
    cmdsearch1.Enabled = True
    txtKode.Enabled = True
    txtkode1.Enabled = True
    cmdsearch = "From Customer"
    cmdsearch1 = "To Customer"
    txtKode = ""
    txtkode1 = ""
End Sub

Private Sub Option2_Click()
    cmdsearch.Enabled = True
    cmdsearch1.Enabled = True
    txtKode.Enabled = True
    txtkode1.Enabled = True
    cmdsearch = "From Area"
    cmdsearch1 = "To Area"
    txtKode = ""
    txtkode1 = ""
End Sub

Private Sub Option3_Click()
    cmdsearch = "From Customer"
    cmdsearch1 = "To Customer"
    txtKode = ""
    txtkode1 = ""
    cmdsearch.Enabled = False
    cmdsearch1.Enabled = False
    txtKode.Enabled = False
    txtkode1.Enabled = False
End Sub

Private Sub Option4_Click()
    cmdsearch = "From Customer"
    cmdsearch1 = "To Customer"
    txtKode = ""
    txtkode1 = ""
    cmdsearch.Enabled = False
    cmdsearch1.Enabled = False
    txtKode.Enabled = False
    txtkode1.Enabled = False
End Sub

Private Sub txtkode_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtkode_LostFocus
End Sub

Private Sub txtkode_LostFocus()
     cariunit
End Sub

Private Sub txtKode1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtKode1_LostFocus
End Sub

Private Sub txtKode1_LostFocus()
     cariunit1
End Sub

Private Sub cariunit()
    If txtKode = "" Then Exit Sub
    OBJ.Open dsn
    If Option1.Value = True Then SQL = "select kodecust from am_customer where kodecust = '" & txtKode & "'"
    If Option2.Value = True Then SQL = "select kode from am_area where kode = '" & txtKode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    MsgBox "Data not found.", vbInformation, "Information"
    txtKode = ""
    txtKode.SetFocus
End Sub

Private Sub cariunit1()
    If txtkode1 = "" Then Exit Sub
    OBJ.Open dsn
    If Option1.Value = True Then SQL = "select * from am_customer where kodecust = '" & txtkode1 & "'"
    If Option2.Value = True Then SQL = "select kode from am_area where kode = '" & txtkode1 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    MsgBox "Data not found.", vbInformation, "Information"
    txtkode1 = ""
    txtkode1.SetFocus
End Sub

Private Sub cmdsearch_Click()
    If Option1.Value = True Then
        carisql1 = "select kodecust, namacust, alamatcust, kota from am_customer"
        namatabel = "Customer"
    Else
        carisql1 = "select kode, nama from am_area"
        namatabel = "Area"
    End If
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    txtKode = hasil
    cariunit
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdsearch1_Click()
    If Option1.Value = True Then
        carisql1 = "select kodecust, namacust, alamatcust, kota from am_customer"
        namatabel = "Customer"
    Else
        carisql1 = "select kode, nama from am_area"
        namatabel = "Area"
    End If
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtkode1 = hasil
    cariunit1
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub
