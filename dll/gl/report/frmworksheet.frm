VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Begin VB.Form frmworksheet 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5775
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
   Icon            =   "frmworksheet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkakum 
      Caption         =   "Pendapatan dan Biaya Di akumulasi"
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
      TabIndex        =   14
      Top             =   2160
      Width           =   2895
   End
   Begin VB.CheckBox chksaldo 
      Caption         =   "Tampilkan Account Saldo 0"
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
      Width           =   2535
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      Format          =   113115137
      CurrentDate     =   37728
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
   Begin VB.TextBox txtarea2 
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
      Left            =   4200
      MaxLength       =   4
      TabIndex        =   1
      Top             =   1200
      Width           =   735
   End
   Begin Crystal.CrystalReport crystal 
      Left            =   360
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
      TabIndex        =   2
      Top             =   1560
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   503
      Calculator      =   "frmworksheet.frx":2372
      Caption         =   "frmworksheet.frx":2392
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmworksheet.frx":23E7
      Keys            =   "frmworksheet.frx":2405
      Spin            =   "frmworksheet.frx":2447
      AlignHorizontal =   2
      AlignVertical   =   2
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "####0;;0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "####0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   10
      MinValue        =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   126877701
      Value           =   10
      MaxValueVT      =   1937178629
      MinValueVT      =   1397948421
   End
   Begin Chameleon.chameleonButton cmdsearch1 
      Height          =   285
      Left            =   360
      TabIndex        =   8
      Top             =   1200
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
      MICON           =   "frmworksheet.frx":246F
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
      Left            =   2880
      TabIndex        =   9
      Top             =   1200
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
      MICON           =   "frmworksheet.frx":2789
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
      Left            =   3600
      TabIndex        =   4
      Top             =   2520
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
      MICON           =   "frmworksheet.frx":2AA3
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
      Left            =   4560
      TabIndex        =   5
      Top             =   2520
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
      MICON           =   "frmworksheet.frx":2DBD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label4 
      Caption         =   "Panjang Account"
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
      TabIndex        =   13
      Top             =   1590
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Work"
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
      TabIndex        =   11
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
      TabIndex        =   10
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label Label3 
      Caption         =   "max 10 (Full Acount)"
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
      Left            =   2640
      TabIndex        =   7
      Top             =   1590
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   10335
   End
End
Attribute VB_Name = "frmworksheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim str1, str2, str3, str4, str5 As String

Private Sub cariarea1()
    If txtarea1 = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from gl_company where kdcomp = '" & txtarea1 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Company " & txtarea1 & " Not Found.", vbExclamation, "Warning"
        txtarea1 = ""
        txtarea1.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub cariarea2()
    If txtarea2 = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from gl_company where kdcomp = '" & txtarea2 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Company " & txtarea2 & " Not Found.", vbExclamation, "Warning"
        txtarea2 = ""
        txtarea2.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdview_Click()
    If txtarea1 = "" Or txtarea2 = "" Then
        MsgBox "Data entry not complite.", vbInformation, "Information"
        Exit Sub
    End If
    
    If txtarea2 < txtarea1 Then
        MsgBox "To Kode Can Not Smaller Then From Kode.", vbExclamation, "Warning"
        txtarea2 = ""
        txtarea2.SetFocus
        Exit Sub
    End If
    
    If txtarea1 <> txtarea2 Then
        str1 = 1
        OBJ.Open dsn
        SQL = "select periode from gl_company where kdcomp >= '" & txtarea1 & "' and kdcomp <= '" & txtarea2 & "' group by periode"
        Set RST = OBJ.Execute(SQL)
        Do While Not RST.EOF
            str1 = str1 + 1
            RST.MoveNext
        Loop
        
        SQL = "select tglakhir from gl_company where kdcomp = '" & txtarea1 & "'"
        Set RST = OBJ.Execute(SQL)
        date1 = RST!tglakhir
        OBJ.Close
        
        If str1 > 2 Then
            MsgBox "Periode Not Same.", vbExclamation, "Warning"
            Exit Sub
        End If
    End If
    
    If chksaldo.Value = 0 Then
        str3 = "saldo"
    Else
        str3 = "semua"
    End If
    
    If chkakum.Value = 0 Then
        str5 = "tidak"
    Else
        str5 = "ya"
    End If
    
    OBJ.Open dsn
    SQL = "select periode from gl_company where kdcomp = '" & txtarea1 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then str4 = RST!periode
    OBJ.Close
    
    crystal.Reset
    crystal.WindowState = crptMaximized
    crystal.WindowShowCloseBtn = True
    crystal.WindowShowPrintSetupBtn = True
    crystal.WindowShowSearchBtn = True
    crystal.WindowShowRefreshBtn = True
    crystal.Connect = dsnreport
    If txtarea1 = txtarea2 And txtpanjang <> 10 Then
        str2 = Str(txtpanjang)
        crystal.DataFiles(0) = "Proc(gl_worksheet1)"
        crystal.ReportFileName = AppPath & "\reports\gl\report\worksheet1.rpt"
        crystal.ParameterFields(0) = "@area1;" + txtarea1 + ";true"
        crystal.ParameterFields(1) = "@panjang;" + str2 + ";true"
        crystal.ParameterFields(2) = "@namauser;" + nmuser + ";true"
        crystal.ParameterFields(3) = "@pilih;" + str3 + ";true"
        crystal.ParameterFields(4) = "@pilih4;" + str4 + ";true"
        crystal.ParameterFields(5) = "@pilih3;" + str5 + ";true"
    ElseIf txtarea1 = txtarea2 And txtpanjang = 10 Then
        crystal.DataFiles(0) = "Proc(gl_worksheet2)"
        crystal.ReportFileName = AppPath & "\reports\gl\report\worksheet2.rpt"
        crystal.ParameterFields(0) = "@area1;" + txtarea1 + ";true"
        crystal.ParameterFields(1) = "@namauser;" + nmuser + ";true"
        crystal.ParameterFields(2) = "@pilih;" + str3 + ";true"
        crystal.ParameterFields(3) = "@pilih4;" + str4 + ";true"
        crystal.ParameterFields(4) = "@pilih3;" + str5 + ";true"
    ElseIf txtarea1 <> txtarea2 And txtpanjang <> 10 Then
        str2 = Str(txtpanjang)
        crystal.DataFiles(0) = "Proc(gl_worksheet3)"
        crystal.ReportFileName = AppPath & "\reports\gl\report\worksheet3.rpt"
        crystal.ParameterFields(0) = "@area1;" + txtarea1 + ";true"
        crystal.ParameterFields(1) = "@area2;" + txtarea2 + ";true"
        crystal.ParameterFields(2) = "@namauser;" + nmuser + ";true"
        crystal.ParameterFields(3) = "@panjang;" + str2 + ";true"
        crystal.ParameterFields(4) = "@tanggal;" + Format(date1, "yyyyMMdd") + ";true"
        crystal.ParameterFields(5) = "@pilih;" + str3 + ";true"
        crystal.ParameterFields(6) = "@pilih4;" + str4 + ";true"
        crystal.ParameterFields(7) = "@pilih3;" + str5 + ";true"
    ElseIf txtarea1 <> txtarea2 And txtpanjang = 10 Then
        crystal.DataFiles(0) = "Proc(gl_worksheet4)"
        crystal.ReportFileName = AppPath & "\reports\gl\report\worksheet4.rpt"
        crystal.ParameterFields(0) = "@area1;" + txtarea1 + ";true"
        crystal.ParameterFields(1) = "@area2;" + txtarea2 + ";true"
        crystal.ParameterFields(2) = "@namauser;" + nmuser + ";true"
        crystal.ParameterFields(3) = "@tanggal;" + Format(date1, "yyyyMMdd") + ";true"
        crystal.ParameterFields(4) = "@pilih;" + str3 + ";true"
        crystal.ParameterFields(5) = "@pilih4;" + str4 + ";true"
        crystal.ParameterFields(6) = "@pilih3;" + str5 + ";true"
    End If
    crystal.RetrieveDataFiles
    crystal.Action = 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtarea1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtarea2.SetFocus
End Sub

Private Sub txtarea1_LostFocus()
    cariarea1
End Sub

Private Sub txtarea2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmdview.SetFocus
End Sub

Private Sub txtarea2_LostFocus()
    cariarea2
End Sub

Private Sub cmdsearch1_Click()
    carisql1 = "select kdcomp, nmcompscr from gl_company"
    namatabel = "Company"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtarea1 = hasil
    hasil = ""
    hasil1 = ""
    txtarea2.SetFocus
End Sub

Private Sub cmdsearch2_Click()
    carisql1 = "select kdcomp, nmcompscr from gl_company"
    namatabel = "Company"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch2_GotFocus()
    If hasil = "" Then Exit Sub
    txtarea2 = hasil
    hasil = ""
    hasil1 = ""
    cmdview.SetFocus
End Sub

Function nmcomp()
    OBJ.Open dsn
    SQL = "select * from gl_company where kdcomp = '" & txtarea1 & "'"
    Set RST = OBJ.Execute(SQL)
    nmcomp = RST!nmcompprn
    OBJ.Close
End Function
