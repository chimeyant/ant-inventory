VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Begin VB.Form frmbalance 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7020
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmbalance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option5 
      Caption         =   "Comparative Current Month With Previous Month With Variance"
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
      Left            =   1680
      TabIndex        =   17
      Top             =   2640
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Comparative Current Month With Previous Month"
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
      Left            =   1680
      TabIndex        =   4
      Top             =   2160
      Width           =   3975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Current Month"
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
      Left            =   1680
      TabIndex        =   3
      Top             =   1920
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Other Month"
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
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   2910
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      Format          =   107151361
      CurrentDate     =   38063
   End
   Begin VB.TextBox txtcom2 
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
      Left            =   4320
      MaxLength       =   4
      TabIndex        =   2
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox txtcom1 
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
      Left            =   1680
      MaxLength       =   4
      TabIndex        =   1
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox txtarea1 
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
      Left            =   1680
      MaxLength       =   4
      TabIndex        =   0
      Top             =   1200
      Width           =   735
   End
   Begin Crystal.CrystalReport crystal 
      Left            =   120
      Top             =   2880
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
   Begin TDBNumber6Ctl.TDBNumber txtpanjang 
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   1935
      _Version        =   65536
      _ExtentX        =   3413
      _ExtentY        =   503
      Calculator      =   "frmbalance.frx":2372
      Caption         =   "frmbalance.frx":2392
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmbalance.frx":23E7
      Keys            =   "frmbalance.frx":2405
      Spin            =   "frmbalance.frx":243F
      AlignHorizontal =   2
      AlignVertical   =   2
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "####0;;Null"
      EditMode        =   0
      Enabled         =   0
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "####0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   12
      MinValue        =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   2085486597
      Value           =   1
      MaxValueVT      =   1937178629
      MinValueVT      =   1397948421
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Comparative YTD Actual With Budget"
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
      Left            =   1680
      TabIndex        =   5
      Top             =   2400
      Width           =   3015
   End
   Begin Chameleon.chameleonButton cmdsearch1 
      Height          =   285
      Left            =   360
      TabIndex        =   11
      Top             =   1200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Report Code"
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
      MICON           =   "frmbalance.frx":2467
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch2 
      Height          =   285
      Left            =   360
      TabIndex        =   12
      Top             =   1560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "From Company"
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
      MICON           =   "frmbalance.frx":2781
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdview 
      Height          =   375
      Left            =   4800
      TabIndex        =   8
      Top             =   2760
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
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
      MICON           =   "frmbalance.frx":2A9B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   5760
      TabIndex        =   9
      Top             =   2760
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmbalance.frx":2DB5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch3 
      Height          =   285
      Left            =   3000
      TabIndex        =   16
      Top             =   1560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "To Company"
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
      MICON           =   "frmbalance.frx":30CF
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Balance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Sheet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   10335
   End
End
Attribute VB_Name = "frmbalance"
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

Dim obj2 As New ADODB.Connection
Dim rst2 As New ADODB.Recordset
Dim sql2 As String

Private Sub cariarea1()
    If txtarea1 = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from gl_rforms where form_no = '" & txtarea1 & "' and report_type = '1'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Report Code " & txtarea1 & " Not Found.", vbExclamation, "Warning"
        txtarea1 = ""
        txtarea1.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdview_Click()
    If txtarea1 = "" Or txtcom1 = "" Or txtcom2 = "" Then
        MsgBox "Data entry not complete.", vbInformation, "Information"
        Exit Sub
    End If
    
    If txtcom1 > txtcom2 Then
        MsgBox "Invalid Company Range.", vbInformation, "Information"
        Exit Sub
    End If
    
    If Option4.Value = True Then
        If Val(bulan) <= txtpanjang And tahun = tahunawal Then
            Option1.Value = True
            
            Exit Sub
        Else
            OBJ.Open dsn
            SQL = "select * from gl_company where kdcomp = '" & txtcom1 & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                If Val(RST!awalbulan) >= txtpanjang And tahun = tahunawal Then
                    Option1.Value = True
                    
                    OBJ.Close
                    Exit Sub
                End If
            End If
            OBJ.Close
        End If
    End If
    
    crystal.Reset
    crystal.WindowState = crptMaximized
    crystal.WindowShowCloseBtn = True
    crystal.WindowShowPrintSetupBtn = True
    crystal.WindowShowSearchBtn = True
    crystal.WindowShowRefreshBtn = True
    If Option1.Value = True Then
        crystal.Connect = dsnreport
        crystal.ReportFileName = AppPath & "\reports\gl\report\inc_bal1.rpt"
        crystal.DataFiles(0) = "Proc(gl_inc_bal1)"
        crystal.ParameterFields(0) = "@form;" + txtarea1 + ";true"
        crystal.ParameterFields(1) = "@com1;" + txtcom1 + ";true"
        crystal.ParameterFields(2) = "@com2;" + txtcom2 + ";true"
        crystal.ParameterFields(3) = "@periode;" + period + ";true"
        crystal.ParameterFields(4) = "@namauser;" + nmuser + ";true"
        crystal.ParameterFields(5) = "@namacomp;" + nmcomp + ";true"
    ElseIf Option2.Value = True Then
        crystal.Connect = dsnreport
        crystal.ReportFileName = AppPath & "\reports\gl\report\inc_bal2.rpt"
        crystal.DataFiles(0) = "Proc(gl_inc_bal2)"
        crystal.ParameterFields(0) = "@form;" + txtarea1 + ";true"
        crystal.ParameterFields(1) = "@com1;" + txtcom1 + ";true"
        crystal.ParameterFields(2) = "@com2;" + txtcom2 + ";true"
        crystal.ParameterFields(3) = "@periode;" + period + ";true"
        crystal.ParameterFields(4) = "@namauser;" + nmuser + ";true"
        crystal.ParameterFields(5) = "@namacomp;" + nmcomp + ";true"
        crystal.ParameterFields(6) = "@bulan;" + bulan + ";true"
    ElseIf Option3.Value = True Then
        crystal.Connect = dsnreport
        crystal.ReportFileName = AppPath & "\report\gl\report\inc_bal4.rpt"
        crystal.DataFiles(0) = "Proc(gl_inc_bal4)"
        crystal.ParameterFields(0) = "@form;" + txtarea1 + ";true"
        crystal.ParameterFields(1) = "@com1;" + txtcom1 + ";true"
        crystal.ParameterFields(2) = "@com2;" + txtcom2 + ";true"
        crystal.ParameterFields(3) = "@periode;" + period + ";true"
        crystal.ParameterFields(4) = "@namauser;" + nmuser + ";true"
        crystal.ParameterFields(5) = "@namacomp;" + nmcomp + ";true"
        crystal.ParameterFields(6) = "@bulan;" + bulan + ";true"
        crystal.ParameterFields(7) = "@pilih;" + "balance" + ";true"
    'ElseIf Option4.Value = True Then
    '    crystal.Connect = dsnreport
    '    crystal.ReportFileName = App.Path & "\report\inc_bal6.rpt"
    '    crystal.DataFiles(0) = "Proc(gl_inc_bal6)"
    '    crystal.ParameterFields(0) = "@form;" + txtarea1 + ";true"
    '    crystal.ParameterFields(1) = "@com1;" + txtcom1 + ";true"
    '    crystal.ParameterFields(2) = "@com2;" + txtcom2 + ";true"
    '    crystal.ParameterFields(3) = "@periode;" + period1 + ";true"
    '    crystal.ParameterFields(4) = "@namauser;" + nmuser + ";true"
    '    crystal.ParameterFields(5) = "@namacomp;" + nmcomp + ";true"
    '    crystal.ParameterFields(6) = "@bulan;" + Str(txtpanjang) + ";true"
    'ElseIf Option5.Value = True Then
    '    crystal.Connect = dsnreport
    '    crystal.ReportFileName = App.Path & "\report\inc_bal7.rpt"
    '    crystal.DataFiles(0) = "Proc(gl_inc_bal2)"
    '    crystal.ParameterFields(0) = "@form;" + txtarea1 + ";true"
    '    crystal.ParameterFields(1) = "@com1;" + txtcom1 + ";true"
    '    crystal.ParameterFields(2) = "@com2;" + txtcom2 + ";true"
    '    crystal.ParameterFields(3) = "@periode;" + period + ";true"
    '    crystal.ParameterFields(4) = "@namauser;" + nmuser + ";true"
    '    crystal.ParameterFields(5) = "@namacomp;" + nmcomp + ";true"
    '    crystal.ParameterFields(6) = "@bulan;" + bulan + ";true"
    End If
    
    crystal.RetrieveDataFiles
    crystal.Action = 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Option1_Click()
    txtpanjang.Enabled = False
End Sub

Private Sub Option2_Click()
    txtpanjang.Enabled = False
End Sub

Private Sub Option3_Click()
    txtpanjang.Enabled = False
End Sub

Private Sub Option4_Click()
    txtpanjang.Enabled = True
End Sub

Private Sub txtarea1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmdview.SetFocus
End Sub

Private Sub txtarea1_LostFocus()
    cariarea1
End Sub

Private Sub cmdsearch1_Click()
    carisql1 = "select Form_no, description from gl_rforms"
    namatabel = "Balance Sheet"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtarea1 = hasil
    hasil = ""
    hasil1 = ""
    txtcom1.SetFocus
End Sub

Private Sub caricom1()
    If txtcom1 = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from gl_company where kdcomp = '" & txtcom1 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Company " & txtcom1 & " Not Found.", vbExclamation, "Warning"
        txtcom1 = ""
        txtcom1.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub caricom2()
    If txtcom2 = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from gl_company where kdcomp = '" & txtcom2 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Company " & txtcom2 & " Not Found.", vbExclamation, "Warning"
        txtcom2 = ""
        txtcom2.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub txtcom1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtcom2.SetFocus
End Sub

Private Sub txtcom1_LostFocus()
    caricom1
End Sub

Private Sub txtcom2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmdview.SetFocus
End Sub

Private Sub txtcom2_LostFocus()
    caricom2
End Sub

Private Sub cmdsearch2_Click()
    carisql1 = "select kdcomp, nmcompscr from gl_company"
    namatabel = "Company"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch2_GotFocus()
    If hasil = "" Then Exit Sub
    txtcom1 = hasil
    hasil = ""
    hasil1 = ""
    txtcom2.SetFocus
End Sub

Private Sub cmdsearch3_Click()
    carisql1 = "select kdcomp, nmcompscr from gl_company"
    namatabel = "Company"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch3_GotFocus()
    If hasil = "" Then Exit Sub
    txtcom2 = hasil
    hasil = ""
    hasil1 = ""
    cmdview.SetFocus
End Sub

Function nmcomp()
    nmcomp = " Konsolidasi "
    If txtcom1 = txtcom2 Then
        OBJ.Open dsn
        SQL = "select * from gl_company where kdcomp = '" & txtcom1 & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then nmcomp = RST!nmcompprn
        OBJ.Close
    End If
End Function

Function period()
    OBJ.Open dsn
    SQL = "select * from gl_company where kdcomp = '" & txtcom1 & "'"
    Set RST = OBJ.Execute(SQL)
    period = "Per " & Format(RST!tglakhir, "dd MMMM yyyy")
    OBJ.Close
End Function

Function bulan()
    OBJ.Open dsn
    SQL = "select * from gl_company where kdcomp = '" & txtcom1 & "'"
    Set RST = OBJ.Execute(SQL)
    bulan = RST!periode
    OBJ.Close
End Function

Function period1()
    OBJ.Open dsn
    SQL = "select * from gl_company where kdcomp = '" & txtcom1 & "'"
    Set RST = OBJ.Execute(SQL)
    date1 = RST!tglakhir
    OBJ.Close
    
    date1.Day = 28
    date1.Month = txtpanjang + 1
    date1.Day = 1
    date1 = date1 - 1
    
    period1 = "Per " & Format(date1, "dd MMMM yyyy")
End Function

Function tahun()
    OBJ1.Open dsn
    SQL1 = "select * from gl_company where kdcomp = '" & txtcom1 & "'"
    Set RST1 = OBJ1.Execute(SQL1)
    tahun = Format(RST1!tglawal, "yyyy")
    OBJ1.Close
End Function

Function tahunawal()
    obj2.Open dsn
    sql2 = "select * from gl_company where kdcomp = '" & txtcom1 & "'"
    Set rst2 = obj2.Execute(sql2)
    tahunawal = rst2!awaltahun
    obj2.Close
End Function

