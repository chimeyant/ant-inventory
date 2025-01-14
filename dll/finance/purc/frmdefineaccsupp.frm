VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~3.OCX"
Begin VB.Form frmdefineaccsupp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Define Account Supplier"
   ClientHeight    =   5745
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
   Icon            =   "frmdefineaccsupp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   360
      Left            =   1050
      TabIndex        =   9
      Top             =   2955
      Visible         =   0   'False
      Width           =   8115
      Begin XtremeSuiteControls.ProgressBar ak 
         Height          =   285
         Left            =   -45
         TabIndex        =   10
         Top             =   75
         Width           =   8235
         _Version        =   851970
         _ExtentX        =   14526
         _ExtentY        =   503
         _StockProps     =   93
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   9600
      TabIndex        =   5
      Top             =   5280
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
      MICON           =   "frmdefineaccsupp.frx":2372
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
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   503
      Caption         =   "frmdefineaccsupp.frx":268C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmdefineaccsupp.frx":26F8
      Key             =   "frmdefineaccsupp.frx":2716
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
      MICON           =   "frmdefineaccsupp.frx":2752
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
      Height          =   4575
      Left            =   45
      TabIndex        =   6
      Top             =   630
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   8070
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
   Begin Chameleon.chameleonButton chameleonButton2 
      Height          =   375
      Left            =   7320
      TabIndex        =   1
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Show Defined"
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
      MICON           =   "frmdefineaccsupp.frx":2A6C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton chameleonButton3 
      Height          =   375
      Left            =   8880
      TabIndex        =   2
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Show Undefined"
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
      MICON           =   "frmdefineaccsupp.frx":2D86
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton chameleonButton4 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   5280
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Automatically Define Account Supplier"
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
      MICON           =   "frmdefineaccsupp.frx":30A0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton chameleonButton5 
      Height          =   375
      Left            =   6360
      TabIndex        =   3
      Top             =   5280
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Save manualy Defined Account Supplier"
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
      MICON           =   "frmdefineaccsupp.frx":33BA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblnamacomp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmdefineaccsupp"
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

Dim posrow, strsupp, strtype, strid As String

Private Sub chameleonButton2_Click()
    If txtkodecomp = "" Then Exit Sub
    
    grid.Rows = grid.Rows + 1
    grid.Rows = grid.Rows - 1
    grid.Clear
    grid.Rows = 2
    grid.TextMatrix(0, 0) = "Supplier"
    grid.TextMatrix(0, 1) = "KodeSupp"
    grid.TextMatrix(0, 2) = "NoAccount"
    grid.TextMatrix(0, 3) = "Description"
    
    grid.ColWidth(0) = 4000
    grid.ColWidth(1) = 1000
    grid.ColWidth(2) = 1200
    grid.ColWidth(3) = 4000
    
    grid.RowHeightMin = 300
    grid.Row = 1
    
    
    OBJ.Open dsn
    SQL = "select count(namasupp)'tot' from am_supplier"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then ak.Max = RST!tot
    OBJ.Close
    Frame1.Visible = True
    
    OBJ.Open dsn
    SQL = "select namasupp,kodesupp from am_supplier order by namasupp"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
    
        OBJ1.Open dsn
        SQL1 = "SELECT noacc FROM am_autoaccsupp where kodecomp = '" & txtkodecomp & "' and kodesupp = '" & RST!kodesupp & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then grid.TextMatrix(grid.Row, 2) = RST1!noacc Else grid.TextMatrix(grid.Row, 2) = ""
        
        If grid.TextMatrix(grid.Row, 2) <> "" Then
            SQL1 = "SELECT nmac FROM gl_masterac where noac = '" & grid.TextMatrix(grid.Row, 2) & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then grid.TextMatrix(grid.Row, 3) = RST1!nmac Else grid.TextMatrix(grid.Row, 3) = ""
            
            grid.Col = 0
            grid.CellAlignment = 1
            grid.TextMatrix(grid.Row, 0) = RST!namasupp
            grid.Col = 1
            grid.CellAlignment = 1
            grid.TextMatrix(grid.Row, 1) = RST!kodesupp
            
            grid.Rows = grid.Rows + 1
            grid.Row = grid.Row + 1
        End If
        OBJ1.Close
        
        ak.Value = ak.Value + 1
        RST.MoveNext
    Loop
    OBJ.Close
    
    ak.Value = 0
    Frame1.Visible = False
End Sub

Private Sub chameleonButton3_Click()
    If txtkodecomp = "" Then Exit Sub
    
    grid.Rows = grid.Rows + 1
    grid.Rows = grid.Rows - 1
    grid.Clear
    grid.Rows = 2
    grid.TextMatrix(0, 0) = "Supplier"
    grid.TextMatrix(0, 1) = "KodeSupp"
    grid.TextMatrix(0, 2) = "NoAccount"
    grid.TextMatrix(0, 3) = "Description"
    
    grid.ColWidth(0) = 4000
    grid.ColWidth(1) = 1000
    grid.ColWidth(2) = 1200
    grid.ColWidth(3) = 4000
    
    grid.RowHeightMin = 300
    grid.Row = 1
    
  
    OBJ.Open dsn
    SQL = "select count(namasupp)'tot' from am_supplier"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then ak.Max = RST!tot
    OBJ.Close
    Frame1.Visible = True
    
    OBJ.Open dsn
    SQL = "select namasupp,kodesupp from am_supplier order by namasupp"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
    
        OBJ1.Open dsn
        SQL1 = "SELECT noacc FROM am_autoaccsupp where kodecomp = '" & txtkodecomp & "' and kodesupp = '" & RST!kodesupp & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If RST1.EOF Then
            grid.Col = 0
            grid.CellAlignment = 1
            grid.TextMatrix(grid.Row, 0) = RST!namasupp
            grid.Col = 1
            grid.CellAlignment = 1
            grid.TextMatrix(grid.Row, 1) = RST!kodesupp
            
            grid.Rows = grid.Rows + 1
            grid.Row = grid.Row + 1
        Else
            If RST1!noacc = "" Then
                grid.Col = 0
                grid.CellAlignment = 1
                grid.TextMatrix(grid.Row, 0) = RST!namasupp
                grid.Col = 1
                grid.CellAlignment = 1
                grid.TextMatrix(grid.Row, 1) = RST!kodesupp
                
                grid.Rows = grid.Rows + 1
                grid.Row = grid.Row + 1
            End If
        End If
        OBJ1.Close
        
        ak.Value = ak.Value + 1
        RST.MoveNext
    Loop
    OBJ.Close
    
    ak.Value = 0
    Frame1.Visible = False
End Sub

Private Sub chameleonButton4_Click()
    Exit Sub
    If txtkodecomp = "" Then Exit Sub
    If MsgBox("Account Gl akan otomatis bertambah sesuai supplier yang belum terdefinisikan." & vbCrLf & _
    "Lanjutkan proses ?", vbYesNo + vbQuestion, "Account Supplier") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from am_option"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        strsupp = RST!ac_supp
        strtype = RST!c_type
        strid = RST!c_id
    
        SQL = "select top 1 noac from gl_masterac where noac like '" & strsupp & "%' order by noac desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then strsupp = RST!noac
    
        strsupp = strsupp + 1
        
        SQL = "select namasupp,kodesupp from am_supplier order by namasupp"
        Set RST = OBJ.Execute(SQL)
        Do While Not RST.EOF
            OBJ1.Open dsn
            SQL1 = "SELECT noacc FROM am_autoaccsupp where kodecomp = '" & txtkodecomp & "' and kodesupp = '" & RST!kodesupp & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If RST1.EOF Then
                SQL1 = "insert into am_autoaccsupp"
                SQL1 = SQL1 + "(kodecomp"
                SQL1 = SQL1 + ",noacc"
                SQL1 = SQL1 + ",kodesupp)"
                
                SQL1 = SQL1 + "VALUES"
                SQL1 = SQL1 + "('" & txtkodecomp & "'"
                SQL1 = SQL1 + ", '" & strsupp & "'"
                SQL1 = SQL1 + ", '" & RST!kodesupp & "')"
                Set RST1 = OBJ1.Execute(SQL1)
                
                SQL1 = "insert into gl_masterac"
                SQL1 = SQL1 + "(noac"
                SQL1 = SQL1 + ",nmac"
                SQL1 = SQL1 + ",typeac"
                SQL1 = SQL1 + ",jenisac1"
                SQL1 = SQL1 + ",jenisac2"
                SQL1 = SQL1 + ",jenisac3"
                SQL1 = SQL1 + ",jenisac4"
                SQL1 = SQL1 + ",jenisac5"
                SQL1 = SQL1 + ",jenisac6"
                SQL1 = SQL1 + ",jenisac7"
                SQL1 = SQL1 + ",jenisac8"
                SQL1 = SQL1 + ",jenisac9"
                SQL1 = SQL1 + ",jenisac10"
                SQL1 = SQL1 + ",flag"
                SQL1 = SQL1 + ",idupdate"
                SQL1 = SQL1 + ",dateupdate"
                SQL1 = SQL1 + ",identry"
                SQL1 = SQL1 + ",Dateentry)"
                
                SQL1 = SQL1 + "VALUES"
                SQL1 = SQL1 + "('" & strsupp & "'"
                SQL1 = SQL1 + ", 'Hutang " & Mid(RST!namasupp, 1, 40) & "'"
                SQL1 = SQL1 + ", 'LI'"
                SQL1 = SQL1 + ", '" & strtype & "'"
                SQL1 = SQL1 + ", ''"
                SQL1 = SQL1 + ", ''"
                SQL1 = SQL1 + ", ''"
                SQL1 = SQL1 + ", ''"
                SQL1 = SQL1 + ", ''"
                SQL1 = SQL1 + ", ''"
                SQL1 = SQL1 + ", ''"
                SQL1 = SQL1 + ", ''"
                SQL1 = SQL1 + ", '" & RST!kodesupp & "'"
                SQL1 = SQL1 + ", '0'"
                SQL1 = SQL1 + ", ''"
                SQL1 = SQL1 + ", convert(datetime,'" & tanggalsekarang & "')"
                SQL1 = SQL1 + ", '" & kuser & "'"
                SQL1 = SQL1 + ", convert(datetime,'" & tanggalsekarang & "'))"
                Set RST1 = OBJ1.Execute(SQL1)
                
                SQL1 = "insert into gl_chacct"
                SQL1 = SQL1 + "(kdcomp"
                SQL1 = SQL1 + ",noac"
                SQL1 = SQL1 + ",typeac"
                SQL1 = SQL1 + ",balancedb"
                SQL1 = SQL1 + ",balancecr"
                SQL1 = SQL1 + ",begindb"
                SQL1 = SQL1 + ",begincr"
                SQL1 = SQL1 + ",periode01"
                SQL1 = SQL1 + ",periode02"
                SQL1 = SQL1 + ",periode03"
                SQL1 = SQL1 + ",periode04"
                SQL1 = SQL1 + ",periode05"
                SQL1 = SQL1 + ",periode06"
                SQL1 = SQL1 + ",periode07"
                SQL1 = SQL1 + ",periode08"
                SQL1 = SQL1 + ",periode09"
                SQL1 = SQL1 + ",periode10"
                SQL1 = SQL1 + ",periode11"
                SQL1 = SQL1 + ",periode12"
                SQL1 = SQL1 + ",periode13"
                SQL1 = SQL1 + ",last01"
                SQL1 = SQL1 + ",last02"
                SQL1 = SQL1 + ",last03"
                SQL1 = SQL1 + ",last04"
                SQL1 = SQL1 + ",last05"
                SQL1 = SQL1 + ",last06"
                SQL1 = SQL1 + ",last07"
                SQL1 = SQL1 + ",last08"
                SQL1 = SQL1 + ",last09"
                SQL1 = SQL1 + ",last10"
                SQL1 = SQL1 + ",last11"
                SQL1 = SQL1 + ",last12"
                SQL1 = SQL1 + ",last13"
                SQL1 = SQL1 + ",temp01"
                SQL1 = SQL1 + ",temp02"
                SQL1 = SQL1 + ",temp03"
                SQL1 = SQL1 + ",temp04"
                SQL1 = SQL1 + ",temp05"
                SQL1 = SQL1 + ",temp06"
                SQL1 = SQL1 + ",temp07"
                SQL1 = SQL1 + ",temp08"
                SQL1 = SQL1 + ",temp09"
                SQL1 = SQL1 + ",temp10"
                SQL1 = SQL1 + ",temp11"
                SQL1 = SQL1 + ",temp12"
                SQL1 = SQL1 + ",temp13"
                SQL1 = SQL1 + ",budget01"
                SQL1 = SQL1 + ",budget02"
                SQL1 = SQL1 + ",budget03"
                SQL1 = SQL1 + ",budget04"
                SQL1 = SQL1 + ",budget05"
                SQL1 = SQL1 + ",budget06"
                SQL1 = SQL1 + ",budget07"
                SQL1 = SQL1 + ",budget08"
                SQL1 = SQL1 + ",budget09"
                SQL1 = SQL1 + ",budget10"
                SQL1 = SQL1 + ",budget11"
                SQL1 = SQL1 + ",budget12"
                SQL1 = SQL1 + ",budget13)"
                
                SQL1 = SQL1 + "VALUES"
                SQL1 = SQL1 + "('" & strid & "'"
                SQL1 = SQL1 + ", '" & strsupp & "'"
                SQL1 = SQL1 + ", 'LI'"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0')"
                SQL1 = SQL1 + ", convert(money,'0'))"
                Set RST1 = OBJ1.Execute(SQL1)
                
                strsupp = strsupp + 1
            End If
            OBJ1.Close
            
            RST.MoveNext
        Loop
    End If
    OBJ.Close
    
    MsgBox "Define account Supplier complete, click ok to continue  ...", vbInformation, "Information"
    Unload Me
End Sub

Private Sub chameleonButton5_Click()
    If txtkodecomp = "" Then Exit Sub
    If MsgBox("Account Gl akan bertambah sesuai supplier yang ada di grid." & vbCrLf & _
    "Lanjutkan proses ?", vbYesNo + vbQuestion, "Account Supplier") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from am_option"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        strsupp = RST!ac_supp
        strtype = RST!c_type
        strid = RST!c_id
    End If
    OBJ.Close
    
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 0) = "" Then Exit Do
        
        If grid.TextMatrix(grid.Row, 2) <> "" Then
            OBJ1.Open dsn
            SQL1 = "SELECT noacc FROM am_autoaccsupp where kodecomp = '" & txtkodecomp & "' and kodesupp = '" & grid.TextMatrix(grid.Row, 1) & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If RST1.EOF Then
                SQL1 = "insert into am_autoaccsupp"
                SQL1 = SQL1 + "(kodecomp"
                SQL1 = SQL1 + ",noacc"
                SQL1 = SQL1 + ",kodesupp)"
                
                SQL1 = SQL1 + "VALUES"
                SQL1 = SQL1 + "('" & txtkodecomp & "'"
                SQL1 = SQL1 + ", '" & grid.TextMatrix(grid.Row, 2) & "'"
                SQL1 = SQL1 + ", '" & grid.TextMatrix(grid.Row, 1) & "')"
                Set RST1 = OBJ1.Execute(SQL1)
            Else
                If RST1!noacc = "" Then
                    SQL1 = "update am_autoaccsupp set "
                    SQL1 = SQL1 + "noacc='" & grid.TextMatrix(grid.Row, 2) & "'"
                    SQL1 = SQL1 + " where kodecomp='" & txtkodecomp & "' and kodesupp='" & grid.TextMatrix(grid.Row, 1) & "'"
                    Set RST1 = OBJ1.Execute(SQL1)
                End If
            End If
            OBJ1.Close
        End If
        grid.Row = grid.Row + 1
    Loop
    
    MsgBox "Define account Supplier complete, click ok to continue  ...", vbInformation, "Information"
    Unload Me
End Sub

Private Sub cmdclose_Click()
    Unload Me
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
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='274' and b.kodeuser = '2" & kuser & "'"
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
    grid.TextMatrix(0, 0) = "Supplier"
    grid.TextMatrix(0, 1) = "KodeSupp"
    grid.TextMatrix(0, 2) = "NoAccount"
    grid.TextMatrix(0, 3) = "Description"
    
    grid.ColWidth(0) = 4000
    grid.ColWidth(1) = 1000
    grid.ColWidth(2) = 1200
    grid.ColWidth(3) = 4000
    
    grid.RowHeightMin = 300
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    If Not chameleonButton5.Enabled Then Exit Sub
    
    posrow = grid.Row
            
    Select Case grid.Col
        Case 2
            If txtkodecomp = "" Then Exit Sub
            If grid.TextMatrix(grid.Row, 0) = "" Then Exit Sub
            
            If grid.TextMatrix(grid.Row, 2) <> "" Then
                If MsgBox("Cancel this account ?", vbYesNo + vbQuestion, "Question") = vbYes Then
                    grid.TextMatrix(grid.Row, 2) = ""
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
    If Not chameleonButton5.Enabled Then Exit Sub
    
    Select Case grid.Col
        Case 2
            grid.Row = posrow
            grid.Col = 2
            grid.CellAlignment = 1
            grid.TextMatrix(grid.Row, 2) = hasil
            grid.TextMatrix(grid.Row, 3) = hasil1
            
            hasil = ""
            hasil1 = ""
            hasil2 = ""
    End Select
End Sub

Private Sub txtkodecomp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chameleonButton2.SetFocus
    KeyAscii = 0
End Sub

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function
