VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmlaporanpo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Purchase Order"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Chameleon.chameleonButton cmdnotprint 
      Height          =   495
      Left            =   5640
      TabIndex        =   3
      Top             =   1080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Belum  Print"
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
      MICON           =   "frmlaporanpo.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Preview"
      Height          =   255
      Left            =   5640
      TabIndex        =   4
      Top             =   1800
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.ListBox List2 
      Height          =   2760
      ItemData        =   "frmlaporanpo.frx":031A
      Left            =   120
      List            =   "frmlaporanpo.frx":031C
      Style           =   1  'Checkbox
      TabIndex        =   7
      Top             =   480
      Width           =   5415
   End
   Begin Crystal.CrystalReport Crystal 
      Left            =   6675
      Top             =   90
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Top             =   2880
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
      MICON           =   "frmlaporanpo.frx":031E
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
      TabIndex        =   5
      Top             =   2160
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
      MICON           =   "frmlaporanpo.frx":0638
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
      TabIndex        =   0
      Top             =   120
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
      Format          =   196870147
      CurrentDate     =   37426
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   4560
      TabIndex        =   1
      Top             =   120
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
      Format          =   196870147
      CurrentDate     =   37426
   End
   Begin Chameleon.chameleonButton cmdsubmit 
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Top             =   480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Sudah Print"
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
      MICON           =   "frmlaporanpo.frx":0952
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
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   450
      Calculator      =   "frmlaporanpo.frx":0C6C
      Caption         =   "frmlaporanpo.frx":0C8C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmlaporanpo.frx":0CF8
      Keys            =   "frmlaporanpo.frx":0D16
      Spin            =   "frmlaporanpo.frx":0D58
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
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   450
      Calculator      =   "frmlaporanpo.frx":0D80
      Caption         =   "frmlaporanpo.frx":0DA0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmlaporanpo.frx":0E0C
      Keys            =   "frmlaporanpo.frx":0E2A
      Spin            =   "frmlaporanpo.frx":0E6C
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
   Begin VB.Label Label1 
      Caption         =   "Display Purchase Order from                                             to"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   165
      Width           =   4335
   End
End
Attribute VB_Name = "frmlaporanpo"
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
    SQL = "select nopo from am_pohdr where ref = 'B' and tglpo >= '" & tanggal1 & "' and tglpo <= '" & tanggal2 & "' order by nopo"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        List2.AddItem RST!nopo
        
        RST.MoveNext
    Loop
    OBJ.Close
End Sub

Private Sub cmdprint_Click()
    OBJ.Open dsn
    SQL = "delete from am_potemp"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    For i = 0 To List2.ListCount - 1
        If List2.Selected(i) = True Then
            OBJ.Open dsn
            SQL = "SELECT (cast(cast((a.qty) as real) as nvarchar) + ' ' + e.namasatuan)'quantity'"
            SQL = SQL + ",a.price,(a.price*a.qty)'harga',a.lineitem,"
            SQL = SQL + "b.nopo,b.tglpo,b.ket3,b.ket4,"
            SQL = SQL + "c.namabarang,"
            SQL = SQL + "d.namasupp,d.contactperson,d.telpsupp,d.faxsupp,"
            SQL = SQL + "g.symkurs"
            SQL = SQL + " FROM am_polin a left join am_pohdr b"
            SQL = SQL + " ON a.nopo=b.nopo left join am_apitemmst c"
            SQL = SQL + " ON a.kodebarang=c.kodebarang left join am_supplier d"
            SQL = SQL + " ON b.kodesupp=d.kodesupp left join am_apunit e"
            SQL = SQL + " ON a.kodesatuan=e.kodesatuan left join gl_kurs g"
            SQL = SQL + " ON b.kodecur=g.kdkurs"
            SQL = SQL + " WHERE b.nopo='" & List2.List(i) & "'"
            Set RST = OBJ.Execute(SQL)
            Do While Not RST.EOF
                txtnil1 = RST!Price
                txtnil2 = RST!harga
                
                OBJ1.Open dsn
                SQL1 = "insert into am_potemp"
                SQL1 = SQL1 + " (quantity"
                SQL1 = SQL1 + ",price"
                SQL1 = SQL1 + ",harga"
                SQL1 = SQL1 + ",lineitem"
                SQL1 = SQL1 + ",nopo"
                SQL1 = SQL1 + ",tglpo"
                SQL1 = SQL1 + ",ket3"
                SQL1 = SQL1 + ",ket4"
                SQL1 = SQL1 + ",namabarang"
                SQL1 = SQL1 + ",namasupp"
                SQL1 = SQL1 + ",contactperson"
                SQL1 = SQL1 + ",telpsupp"
                SQL1 = SQL1 + ",faxsupp"
                SQL1 = SQL1 + ",symkurs)"
                
                SQL1 = SQL1 + "VALUES"
                SQL1 = SQL1 + "('" & RST!quantity & "'"
                SQL1 = SQL1 + ", convert(money,'" & txtnil1 & "')"
                SQL1 = SQL1 + ", convert(money,'" & txtnil2 & "')"
                SQL1 = SQL1 + ", convert(numeric,'" & RST!lineitem & "')"
                SQL1 = SQL1 + ", '" & RST!nopo & "'"
                SQL1 = SQL1 + ", convert(datetime,'" & Month(RST!tglpo) & "/" & Day(RST!tglpo) & "/" & Year(RST!tglpo) & "')"
                SQL1 = SQL1 + ", '" & RST!ket3 & "'"
                SQL1 = SQL1 + ", '" & RST!ket4 & "'"
                SQL1 = SQL1 + ", '" & RST!namabarang & "'"
                SQL1 = SQL1 + ", '" & RST!namasupp & "'"
                SQL1 = SQL1 + ", '" & RST!contactperson & "'"
                SQL1 = SQL1 + ", '" & RST!telpsupp & "'"
                SQL1 = SQL1 + ", '" & RST!faxsupp & "'"
                SQL1 = SQL1 + ", '" & RST!symkurs & "')"
                Set RST1 = OBJ1.Execute(SQL1)
                OBJ1.Close
                
                RST.MoveNext
            Loop
            
            'If Check2.Value = 0 Then
                SQL = "update am_pohdr set ref = 'P' where nopo = '" & List2.List(i) & "'"
                Set RST = OBJ.Execute(SQL)
            'End If
            OBJ.Close
        End If
    Next i
    
    Crystal.Reset
    If Check2.Value = 1 Then
        Crystal.WindowState = crptMaximized
        Crystal.WindowShowCloseBtn = True
        Crystal.WindowShowPrintBtn = True
        Crystal.WindowShowSearchBtn = True
        Crystal.WindowShowPrintSetupBtn = True
        Crystal.Destination = crptToWindow
    Else
        Crystal.Destination = crptToPrinter
    End If
    
    Crystal.Connect = dsnreport
    Crystal.DataFiles(0) = "Proc(am_pomany)"
    Crystal.ReportFileName = AppPath & "\reports\purchasing\purc\purchaseordermany.rpt"
    Crystal.ParameterFields(0) = "@namauser;" + nmuser + ";true"
    Crystal.RetrieveDataFiles
    Crystal.Action = 1
End Sub

Private Sub cmdsubmit_Click()
    List2.Clear
    
    If date1 > date2 Then
        MsgBox "Invalid date range, search abort.", vbExclamation, "Error"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select nopo from am_pohdr where ref = 'P' and tglpo >= '" & tanggal1 & "' and tglpo <= '" & tanggal2 & "' order by nopo"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        List2.AddItem RST!nopo
        
        RST.MoveNext
    Loop
    OBJ.Close
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
