VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmanalisajual 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Analisa Penjualan"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
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
   Icon            =   "frmanalisajual.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdview 
      Caption         =   "View"
      Height          =   360
      Left            =   2715
      TabIndex        =   13
      Top             =   4410
      Width           =   840
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   360
      Left            =   3585
      TabIndex        =   12
      Top             =   4410
      Width           =   840
   End
   Begin Crystal.CrystalReport crystal 
      Left            =   120
      Top             =   4800
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
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   480
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
      Format          =   110559235
      CurrentDate     =   37464
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   840
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
      Format          =   110559235
      CurrentDate     =   37464
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   1815
      Left            =   0
      TabIndex        =   5
      Top             =   2400
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3201
      _Version        =   393216
      Rows            =   0
      Cols            =   3
      FixedRows       =   0
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
      _Band(0).Cols   =   3
   End
   Begin TDBNumber6Ctl.TDBNumber txtbaris 
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Top             =   2040
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   503
      Calculator      =   "frmanalisajual.frx":2372
      Caption         =   "frmanalisajual.frx":2392
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmanalisajual.frx":23FE
      Keys            =   "frmanalisajual.frx":241C
      Spin            =   "frmanalisajual.frx":2466
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
      EditMode        =   1
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "####0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   5
      MinValue        =   2
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   1245189
      Value           =   2
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.Label Label6 
      Caption         =   "Compare"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Data :"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Vertikal"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1710
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Horisontal"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1350
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "To Date"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   870
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "From Date"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   510
      Width           =   855
   End
   Begin MSForms.ComboBox cmbtypey 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   1680
      Width           =   1935
      VariousPropertyBits=   746608667
      MaxLength       =   2
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3413;503"
      ListRows        =   11
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbtypex 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
      VariousPropertyBits=   746608667
      MaxLength       =   2
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3413;503"
      ListRows        =   11
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmanalisajual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim str1, str2 As String
Dim i, j As Integer

Private Sub cmbtypex_Change()
    grid.Clear
    grid.Rows = txtbaris
    grid.RowHeightMin = 300
    grid.ColWidth(0) = 1200
    grid.ColWidth(1) = 1000
    grid.ColWidth(2) = 2000

    For i = 1 To txtbaris
        grid.TextMatrix(i - 1, 0) = cmbtypex & " " & i
    Next i
End Sub

Private Sub cmbtypex_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub cmbtypex_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
End Sub

Private Sub cmbtypey_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub cmbtypey_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
   
    
    date1 = Date
    date2 = Date
    
    cmbtypex.Clear
    cmbtypex.ColumnCount = 1
    cmbtypex.ListWidth = "3 cm"
    
    cmbtypey.Clear
    cmbtypey.ColumnCount = 1
    cmbtypey.ListWidth = "3 cm"

    cmbtypey.AddItem "Quantity"
    
    cmbtypex.AddItem "Salesman"
    cmbtypex.AddItem "Item"
    cmbtypex.AddItem "Area"
End Sub

Private Sub cmdview_Click()
    If date2 < date1 Then
        MsgBox "To Date Can Not Smaller Then From Date.", vbExclamation, "Warning"
        date2.SetFocus
        Exit Sub
    End If
    
    For i = 1 To txtbaris
        If grid.TextMatrix(i - 1, 1) = "" Or grid.TextMatrix(i - 1, 2) = "" Then
            MsgBox "User harus mengisi ", vbExclamation, "Warning"
            Exit Sub
        End If
    Next i
    
    crystal.Reset
    crystal.WindowState = crptMaximized
    crystal.WindowShowCloseBtn = True
    crystal.WindowShowPrintSetupBtn = True
    crystal.WindowShowSearchBtn = True
    crystal.WindowShowRefreshBtn = True
    crystal.Connect = dsnreport
    
    Select Case cmbtypex
        Case "Salesman"
            str1 = "sales"
        Case "Item"
            str1 = "item"
        Case "Area"
            str1 = "area"
    End Select
    
    Select Case txtbaris
        Case 2
            crystal.DataFiles(0) = "Proc(am_analisajual1)"
            crystal.ReportFileName = AppPath & "\reports\sale\inv\grafik1.rpt"
            crystal.ParameterFields(0) = "@tanggal1;" & Format(date1, "yyyyMMdd") & ";true"
            crystal.ParameterFields(1) = "@tanggal2;" & Format(date2, "yyyyMMdd") & ";true"
            crystal.ParameterFields(2) = "@pilih ;" + str1 + ";true"
            crystal.ParameterFields(3) = "@comp1 ;" + grid.TextMatrix(0, 1) + ";true"
            crystal.ParameterFields(4) = "@comp2 ;" + grid.TextMatrix(1, 1) + ";true"
        Case 3
            crystal.DataFiles(0) = "Proc(am_analisajual2)"
            crystal.ReportFileName = AppPath & "\reports\sale\inv\grafik2.rpt"
            crystal.ParameterFields(0) = "@tanggal1;" & Format(date1, "yyyyMMdd") & ";true"
            crystal.ParameterFields(1) = "@tanggal2;" & Format(date2, "yyyyMMdd") & ";true"
            crystal.ParameterFields(2) = "@pilih ;" + str1 + ";true"
            crystal.ParameterFields(3) = "@comp1 ;" + grid.TextMatrix(0, 1) + ";true"
            crystal.ParameterFields(4) = "@comp2 ;" + grid.TextMatrix(1, 1) + ";true"
            crystal.ParameterFields(5) = "@comp3 ;" + grid.TextMatrix(2, 1) + ";true"
        Case 4
            crystal.DataFiles(0) = "Proc(am_analisajual3)"
            crystal.ReportFileName = AppPath & "\reports\sale\inv\grafik3.rpt"
            crystal.ParameterFields(0) = "@tanggal1;" & Format(date1, "yyyyMMdd") & ";true"
            crystal.ParameterFields(1) = "@tanggal2;" & Format(date2, "yyyyMMdd") & ";true"
            crystal.ParameterFields(2) = "@pilih ;" + str1 + ";true"
            crystal.ParameterFields(3) = "@comp1 ;" + grid.TextMatrix(0, 1) + ";true"
            crystal.ParameterFields(4) = "@comp2 ;" + grid.TextMatrix(1, 1) + ";true"
            crystal.ParameterFields(5) = "@comp3 ;" + grid.TextMatrix(2, 1) + ";true"
            crystal.ParameterFields(5) = "@comp4 ;" + grid.TextMatrix(3, 1) + ";true"
        Case 5
            crystal.DataFiles(0) = "Proc(am_analisajual4)"
            crystal.ReportFileName = AppPath & "\reports\sale\inv\grafik4.rpt"
            crystal.ParameterFields(0) = "@tanggal1;" & Format(date1, "yyyyMMdd") & ";true"
            crystal.ParameterFields(1) = "@tanggal2;" & Format(date2, "yyyyMMdd") & ";true"
            crystal.ParameterFields(2) = "@pilih ;" + str1 + ";true"
            crystal.ParameterFields(3) = "@comp1 ;" + grid.TextMatrix(0, 1) + ";true"
            crystal.ParameterFields(4) = "@comp2 ;" + grid.TextMatrix(1, 1) + ";true"
            crystal.ParameterFields(5) = "@comp3 ;" + grid.TextMatrix(2, 1) + ";true"
            crystal.ParameterFields(5) = "@comp4 ;" + grid.TextMatrix(3, 1) + ";true"
            crystal.ParameterFields(6) = "@comp5 ;" + grid.TextMatrix(4, 1) + ";true"
    End Select
    
    crystal.RetrieveDataFiles
    crystal.Action = 1
End Sub

Private Sub grid_Click()
    If cmbtypex = "" Then Exit Sub
    j = grid.Row
    If grid.MouseCol = 0 Then
        Select Case cmbtypex
            Case "Salesman"
                carisql1 = "select kodesales, namasales from AM_salesman"
                namatabel = "Sales"
            Case "Item"
                carisql1 = "select kodebarang, namabarang from am_itemmst"
                namatabel = "Item"
            Case "Area"
                carisql1 = "select kode, nama from am_area"
                namatabel = "Area"
        End Select
        frmsearch.Show vbModal
    End If
End Sub

Private Sub grid_GotFocus()
    If hasil = "" Then Exit Sub
    grid.TextMatrix(j, 1) = hasil
    OBJ.Open dsn
    Select Case cmbtypex
        Case "Salesman"
            SQL = "select (namasales)'nama' from AM_salesman where kodesales = '" & hasil & "'"
        Case "Item"
            SQL = "select (namabarang)'nama' from am_itemmst where kodebarang = '" & hasil & "'"
        Case "Area"
            SQL = "select nama from am_area where kode = '" & hasil & "'"
    End Select
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then grid.TextMatrix(j, 2) = RST!nama
    OBJ.Close
    carisales
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub txtbaris_Change()
    grid.Clear
    grid.Rows = txtbaris
    grid.RowHeightMin = 300
    grid.ColWidth(0) = 1200
    grid.ColWidth(1) = 1000
    grid.ColWidth(2) = 2000
    
    For i = 1 To txtbaris
        grid.TextMatrix(i - 1, 0) = cmbtypex & " " & i
    Next i
End Sub

Private Sub carisales()
    'If txtsales = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "SELECT * FROM AM_Salesman WHERE KodeSales = '" & grid.TextMatrix(j, 1) & "'"
    Set RST = OBJ.Execute(SQL)
'-------------------- 0 = sales non aktif -------------------
    If RST!idupdate = "0" Then
        MsgBox "Salesman " & grid.TextMatrix(j, 2) & " is not active !", vbExclamation, "Warning"
        grid.TextMatrix(j, 1) = ""
        grid.TextMatrix(j, 2) = ""
    End If
    OBJ.Close
End Sub
