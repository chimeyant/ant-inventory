VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0A1435CB-EB1C-11D4-89B0-204C4F4F5020}#3.0#0"; "akprogressbar.ocx"
Begin VB.Form frmdefineaccbank 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Define Account Bank/Kas"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10695
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
   Icon            =   "frmdefineaccbank.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   9600
      TabIndex        =   6
      Top             =   5880
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
      MICON           =   "frmdefineaccbank.frx":2372
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
      Left            =   8640
      TabIndex        =   5
      Top             =   5880
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
      MICON           =   "frmdefineaccbank.frx":268C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBText6Ctl.TDBText txtkodecomp 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   503
      Caption         =   "frmdefineaccbank.frx":29A6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmdefineaccbank.frx":2A12
      Key             =   "frmdefineaccbank.frx":2A30
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   0
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   0
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   0
      ErrorBeep       =   0
      MaxLength       =   4
      LengthAsByte    =   0
      Text            =   ""
      Furigana        =   0
      HighlightText   =   -1
      IMEMode         =   0
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin Chameleon.chameleonButton cmdsearch 
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Kode Company"
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
      MICON           =   "frmdefineaccbank.frx":2A6C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   3495
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   6165
      _Version        =   393216
      Cols            =   4
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
      _Band(0).Cols   =   4
   End
   Begin akProgress.akProgressBar ak 
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   5880
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   661
      BackColour      =   -2147483633
      FontColour      =   8421504
      BarColour       =   16761024
      Horizontal      =   -1  'True
      ReverseGradient =   0   'False
      Max             =   100
      Min             =   0
      GapWidth        =   0
      LineWidth       =   1
      Caption         =   0
      BorderStyle     =   0
      Margin          =   2
      Gradient        =   0
      Alignment       =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid1 
      Height          =   1695
      Left            =   0
      TabIndex        =   2
      Top             =   4080
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   2990
      _Version        =   393216
      Cols            =   4
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
      _Band(0).Cols   =   4
   End
   Begin Chameleon.chameleonButton cmdlist1 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   5880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "List Bank"
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
      MICON           =   "frmdefineaccbank.frx":2D86
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdlist2 
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   5880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "List Currency"
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
      MICON           =   "frmdefineaccbank.frx":30A0
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
      Left            =   10320
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lblnamacomp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2520
      TabIndex        =   8
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmdefineaccbank"
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

Dim posrow, posrow1 As String

Private Sub cmdadd_Click()
    If txtkodecomp = "" Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If

    If grid.Rows = 2 Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If grid1.Rows = 2 Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
        
    OBJ.Open dsn
    SQL = "delete from am_autoaccbank where kodecomp = '" & txtkodecomp & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    ak.CaptionType = CaptionNone
    ak.Max = (grid.Rows - 2) + (grid1.Rows - 2)
    grid.Row = 1
    Do While True
        If grid.Rows = grid.Row + 1 Then Exit Do
        
        OBJ.Open dsn
        SQL = "insert into am_autoaccbank ("
        SQL = SQL + "kodecomp, "
        SQL = SQL + "noacc, "
        SQL = SQL + "kodebank)"
    
        SQL = SQL + " values('" & txtkodecomp & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 2) & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 1) & "')"
        Set RST = OBJ.Execute(SQL)
        OBJ.Close
        
        ak.Value = ak.Value + 1
        grid.Row = grid.Row + 1
    Loop
    
    grid1.Row = 1
    Do While True
        If grid1.Rows = grid1.Row + 1 Then Exit Do
        
        OBJ.Open dsn
        SQL = "insert into am_autoaccbank ("
        SQL = SQL + "kodecomp, "
        SQL = SQL + "noacc, "
        SQL = SQL + "kodebank)"
    
        SQL = SQL + " values('" & txtkodecomp & "',"
        SQL = SQL + "'" & grid1.TextMatrix(grid1.Row, 2) & "',"
        SQL = SQL + "'" & grid1.TextMatrix(grid1.Row, 1) & "')"
        Set RST = OBJ.Execute(SQL)
        OBJ.Close
        
        ak.Value = ak.Value + 1
        grid1.Row = grid1.Row + 1
    Loop
    
    ak.Value = 0
    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
    Unload Me
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdlist1_Click()
    Crystal.Reset
    Crystal.WindowState = crptMaximized
    Crystal.WindowShowCloseBtn = True
    Crystal.WindowShowPrintSetupBtn = True
    Crystal.WindowShowSearchBtn = True
    Crystal.WindowShowRefreshBtn = True
    Crystal.Connect = dsnreport1
    Crystal.DataFiles(0) = "Proc(am_listbank)"
    Crystal.ReportFileName = AppPath & "\reports\finance\sale\listbank.rpt"
    Crystal.ParameterFields(0) = "@kode1;0;true"
    Crystal.ParameterFields(1) = "@kode2;z;true"
    Crystal.ParameterFields(2) = "@namauser;" + nmuser + ";true"
    Crystal.RetrieveDataFiles
    Crystal.Action = 1
End Sub

Private Sub cmdlist2_Click()
    Crystal.Reset
    Crystal.WindowState = crptMaximized
    Crystal.WindowShowCloseBtn = True
    Crystal.WindowShowPrintSetupBtn = True
    Crystal.WindowShowSearchBtn = True
    Crystal.Connect = dsnreport1
    Crystal.DataFiles(0) = "Proc(gl_curlist)"
    Crystal.ReportFileName = AppPath & "\reports\finance\sale\listcurr.rpt"
    Crystal.ParameterFields(0) = "@area1;0;true"
    Crystal.ParameterFields(1) = "@area2;z;true"
    Crystal.ParameterFields(2) = "@namauser;" + nmuser + ";true"
    Crystal.RetrieveDataFiles
    Crystal.Action = 1
End Sub

Private Sub cmdsearch_Click()
    carisql1 = "select kdcomp, nmcompscr from gl_company"
    namatabel = "Company"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    txtkodecomp = hasil
    lblnamacomp = hasil1
    hasil = ""
    hasil1 = ""
    cariautojurnal
End Sub

Private Sub Form_Activate()
    OBJ.Open dsn
    SQL = "select * from am_period"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "The period is empty !!" & vbCrLf & _
        "Please define Period on proces, Starting period date and Ending period date.", vbCritical, "Critical"
        
        OBJ.Close
        Unload Me
        Exit Sub
    End If
    OBJ.Close
    
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='464' and b.kodeuser = '1" & kuser & "'"
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
   
    grid.TextMatrix(0, 0) = "Bank"
    grid.TextMatrix(0, 1) = "KodeBank"
    grid.TextMatrix(0, 2) = "NoAccount"
    grid.TextMatrix(0, 3) = "Description"
    
    grid.ColWidth(0) = 4000
    grid.ColWidth(1) = 1000
    grid.ColWidth(2) = 1200
    grid.ColWidth(3) = 4000
    
    grid.RowHeightMin = 300
    
    grid1.TextMatrix(0, 0) = "Cash"
    grid1.TextMatrix(0, 1) = "Currency"
    grid1.TextMatrix(0, 2) = "NoAccount"
    grid1.TextMatrix(0, 3) = "Description"
    
    grid1.ColWidth(0) = 4000
    grid1.ColWidth(1) = 1000
    grid1.ColWidth(2) = 1200
    grid1.ColWidth(3) = 4000
    
    grid1.RowHeightMin = 300
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    
    posrow = grid.Row
    Select Case grid.Col
        Case 2
            If grid.TextMatrix(grid.Row, 0) = "" Or txtkodecomp = "" Then Exit Sub
            
            If grid.TextMatrix(grid.Row, 2) <> "" Then
                If MsgBox("Cancel this account ?", vbYesNo + vbQuestion, "Question") = vbYes Then
                    grid.TextMatrix(grid.Row, 2) = ""
                    grid.TextMatrix(grid.Row, 3) = ""
                    Exit Sub
                End If
            End If
                        
            carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "'"
            namatabel = "Company Account"

            frmsearch.Show vbModal
    End Select
End Sub

Private Sub grid_GotFocus()
    If hasil = "" Then Exit Sub
    
    Select Case grid.Col
        Case 2
            grid.Row = posrow
            grid.Col = 2
            grid.CellAlignment = 1
            grid.TextMatrix(grid.Row, 2) = hasil
            hasil = ""
            hasil1 = ""
            hasil2 = ""
                        
            OBJ.Open dsn
            SQL = "select * from gl_masterac where noac = '" & grid.TextMatrix(grid.Row, 2) & "'"
            Set RST = OBJ.Execute(SQL)
            If RST!flag = 1 Then
                MsgBox "Account Header tidak bisa dipakai untuk jurnal.", vbExclamation, "Warning"
                
                OBJ.Close
                Exit Sub
            End If
            grid.TextMatrix(grid.Row, 3) = RST!nmac
            OBJ.Close
    End Select
End Sub

Private Sub grid1_Click()
    If grid1.MouseRow = 0 Then Exit Sub
    
    posrow1 = grid1.Row
    Select Case grid1.Col
        Case 2
            If grid1.TextMatrix(grid1.Row, 0) = "" Or txtkodecomp = "" Then Exit Sub
            
            If grid1.TextMatrix(grid1.Row, 2) <> "" Then
                If MsgBox("Cancel this account ?", vbYesNo + vbQuestion, "Question") = vbYes Then
                    grid1.TextMatrix(grid1.Row, 2) = ""
                    grid1.TextMatrix(grid1.Row, 3) = ""
                    Exit Sub
                End If
            End If
                        
            carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtkodecomp & "'"
            namatabel = "Company Account"

            frmsearch.Show vbModal
    End Select
End Sub

Private Sub grid1_GotFocus()
    If hasil = "" Then Exit Sub
    
    Select Case grid1.Col
        Case 2
            grid1.Row = posrow1
            grid1.Col = 2
            grid1.CellAlignment = 1
            grid1.TextMatrix(grid1.Row, 2) = hasil
            hasil = ""
            hasil1 = ""
            hasil2 = ""
                        
            OBJ.Open dsn
            SQL = "select * from gl_masterac where noac = '" & grid1.TextMatrix(grid1.Row, 2) & "'"
            Set RST = OBJ.Execute(SQL)
            If RST!flag = 1 Then
                MsgBox "Account Header tidak bisa dipakai untuk jurnal.", vbExclamation, "Warning"
                
                OBJ.Close
                Exit Sub
            End If
            grid1.TextMatrix(grid1.Row, 3) = RST!nmac
            OBJ.Close
    End Select
End Sub

Private Sub txtkodecomp_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtkodecomp_LostFocus
End Sub

Private Sub txtkodecomp_LostFocus()
    If txtkodecomp = "" Then Exit Sub
    If txtkodecomp.SelLength <> 0 Then Exit Sub
        
    OBJ.Open dsn
    SQL = "select * from gl_company where kdcomp = '" & txtkodecomp & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblnamacomp = RST!nmcompscr
    Else
        MsgBox "Company " & txtkodecomp & " Not Found.", vbInformation, "Information"
        txtkodecomp = ""
        txtkodecomp.SetFocus
    End If
    OBJ.Close
    
    cariautojurnal
End Sub

Private Sub cariautojurnal()
    If txtkodecomp = "" Then Exit Sub
    
    grid.Rows = grid.Rows + 1
    grid.Rows = grid.Rows - 1
    grid.Clear
    grid.Rows = 2
    grid.TextMatrix(0, 0) = "Bank"
    grid.TextMatrix(0, 1) = "KodeBank"
    grid.TextMatrix(0, 2) = "NoAccount"
    grid.TextMatrix(0, 3) = "Description"
    
    grid.ColWidth(0) = 4000
    grid.ColWidth(1) = 1000
    grid.ColWidth(2) = 1200
    grid.ColWidth(3) = 4000
    
    grid.RowHeightMin = 300
    grid.Row = 1
    
    grid1.Rows = grid1.Rows + 1
    grid1.Rows = grid1.Rows - 1
    grid1.Clear
    grid1.Rows = 2
    grid1.TextMatrix(0, 0) = "Cash"
    grid1.TextMatrix(0, 1) = "Currency"
    grid1.TextMatrix(0, 2) = "NoAccount"
    grid1.TextMatrix(0, 3) = "Description"
    
    grid1.ColWidth(0) = 4000
    grid1.ColWidth(1) = 1000
    grid1.ColWidth(2) = 1200
    grid1.ColWidth(3) = 4000
    
    grid1.RowHeightMin = 300
    grid1.Row = 1
    
    ak.CaptionType = CaptionNone
    OBJ.Open dsn
    SQL = "select count(description)'tot' from am_bank"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then ak.Max = RST!tot
    
    SQL = "select description,kode from am_bank order by description"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        grid.Col = 0
        grid.CellAlignment = 1
        grid.TextMatrix(grid.Row, 0) = RST!Description
        grid.Col = 1
        grid.CellAlignment = 1
        grid.TextMatrix(grid.Row, 1) = RST!kode

        OBJ1.Open dsn
        SQL1 = "SELECT noacc FROM am_autoaccbank where kodecomp = '" & txtkodecomp & "' and kodebank = '" & grid.TextMatrix(grid.Row, 1) & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then grid.TextMatrix(grid.Row, 2) = RST1!noacc Else grid.TextMatrix(grid.Row, 2) = ""
        
        SQL1 = "SELECT nmac FROM gl_masterac where noac = '" & grid.TextMatrix(grid.Row, 2) & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then grid.TextMatrix(grid.Row, 3) = RST1!nmac Else grid.TextMatrix(grid.Row, 3) = ""
        OBJ1.Close
        
        ak.Value = ak.Value + 1
        grid.Rows = grid.Rows + 1
        grid.Row = grid.Row + 1
        RST.MoveNext
    Loop
    
    SQL = "select kdkurs from gl_kurs order by kdkurs"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        grid1.Col = 0
        grid1.CellAlignment = 1
        grid1.TextMatrix(grid1.Row, 0) = "Kas Tunai"
        grid1.Col = 1
        grid1.CellAlignment = 1
        grid1.TextMatrix(grid1.Row, 1) = RST!kdkurs
    
        OBJ1.Open dsn
        SQL1 = "SELECT noacc FROM am_autoaccbank where kodecomp = '" & txtkodecomp & "' and kodebank = '" & grid1.TextMatrix(grid1.Row, 1) & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then grid1.TextMatrix(grid1.Row, 2) = RST1!noacc Else grid1.TextMatrix(grid1.Row, 2) = ""
        
        SQL1 = "SELECT nmac FROM gl_masterac where noac = '" & grid1.TextMatrix(grid1.Row, 2) & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then grid1.TextMatrix(grid1.Row, 3) = RST1!nmac Else grid1.TextMatrix(grid1.Row, 3) = ""
        OBJ1.Close
        
        ak.Value = ak.Value + 1
        grid1.Rows = grid1.Rows + 1
        grid1.Row = grid1.Row + 1
        RST.MoveNext
    Loop
    OBJ.Close
    
    ak.Value = 0
End Sub
