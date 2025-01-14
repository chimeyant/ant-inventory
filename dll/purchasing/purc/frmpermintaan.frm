VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmpermintaan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Permintaan Barang"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10770
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.CheckBox cekPO 
      Height          =   375
      Left            =   1320
      TabIndex        =   23
      Top             =   1920
      Width           =   1215
      _Version        =   851970
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Dengan PO"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin VB.TextBox txtbrg 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Left            =   4920
      MaxLength       =   100
      TabIndex        =   21
      Top             =   2760
      Visible         =   0   'False
      Width           =   1815
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   4920
      TabIndex        =   17
      Top             =   2520
      Visible         =   0   'False
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   450
      Calculator      =   "frmpermintaan.frx":0000
      Caption         =   "frmpermintaan.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpermintaan.frx":008C
      Keys            =   "frmpermintaan.frx":00AA
      Spin            =   "frmpermintaan.frx":00EC
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##,###,##0.00;(##,###,##0.00);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##,###,##0.00;(##,###,##0.00)"
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
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.PictureBox blank 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5040
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   20
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox check 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5520
      Picture         =   "frmpermintaan.frx":0114
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   19
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox uncheck 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      Picture         =   "frmpermintaan.frx":03F6
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   18
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtpemesan 
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
      MaxLength       =   30
      TabIndex        =   0
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox txtket 
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
      Height          =   765
      Left            =   7680
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox txtbagian 
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
      MaxLength       =   30
      TabIndex        =   1
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox txtnobukti 
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
      Locked          =   -1  'True
      MaxLength       =   17
      TabIndex        =   8
      Top             =   120
      Width           =   1815
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   9600
      TabIndex        =   4
      Top             =   4440
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
      MICON           =   "frmpermintaan.frx":0744
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdclear 
      Height          =   375
      Left            =   8640
      TabIndex        =   5
      Top             =   4440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Clear"
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
      MICON           =   "frmpermintaan.frx":0A5E
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
      Left            =   7680
      TabIndex        =   6
      Top             =   4440
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
      MICON           =   "frmpermintaan.frx":0D78
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
      TabIndex        =   9
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
      Format          =   143261699
      CurrentDate     =   37426
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   8640
      TabIndex        =   2
      Top             =   1200
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
      Format          =   143261699
      CurrentDate     =   37426
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   1815
      Left            =   0
      TabIndex        =   12
      Top             =   2400
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   3201
      _Version        =   393216
      Cols            =   7
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
      _Band(0).Cols   =   7
   End
   Begin XtremeSuiteControls.CheckBox CheckBox1 
      Height          =   255
      Left            =   1080
      TabIndex        =   25
      Top             =   4290
      Width           =   1215
      _Version        =   851970
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Dengan PO"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
      Value           =   1
   End
   Begin MSComCtl2.DTPicker dtpdefault 
      Height          =   285
      Left            =   8640
      TabIndex        =   26
      Top             =   600
      Visible         =   0   'False
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
      Format          =   143261699
      CurrentDate     =   2
   End
   Begin Crystal.CrystalReport Crystal 
      Left            =   9960
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "* Checklist                             jika permintaan butuh dibuatkan PO        di bagian Purchasing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   240
      TabIndex        =   24
      Top             =   4320
      Width           =   4815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Pemesan"
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
      Left            =   120
      TabIndex        =   16
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan"
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
      Left            =   6480
      TabIndex        =   15
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Selesai yang diminta"
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
      Left            =   6480
      TabIndex        =   14
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bagian"
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
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
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
      Left            =   120
      TabIndex        =   11
      Top             =   510
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "No Permintaan"
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
      Left            =   120
      TabIndex        =   10
      Top             =   150
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   7560
      Shape           =   4  'Rounded Rectangle
      Top             =   4320
      Width           =   3015
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   0
      TabIndex        =   22
      Top             =   960
      Width           =   10815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   7
      Top             =   4200
      Width           =   10815
   End
End
Attribute VB_Name = "frmpermintaan"
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
Dim posrow As String


Private Sub cmdadd_Click()
    Dim strformat As String
    Dim strno As String
    strformat = Format(Date, "yymm")
    If txtnobukti = "" Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        txtpemesan.SetFocus
        Exit Sub
    End If
    
    If grid.Rows = 2 Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    grid.Row = 1
    Do While True
        If grid.Rows = grid.Row + 1 Then Exit Do
        
        If grid.TextMatrix(grid.Row, 2) = "0.00" Or grid.TextMatrix(grid.Row, 2) = "" Then
            MsgBox "Data entry not complete, on row " & grid.Row, vbExclamation, "Warning"
            Exit Sub
        End If
        If grid.TextMatrix(grid.Row, 5) = "" Then
            MsgBox "Data entry not complete, on row " & grid.Row, vbExclamation, "Warning"
            Exit Sub
        End If
        If grid.TextMatrix(grid.Row, 6) = "" Then
            MsgBox "Data entry not complete, on row " & grid.Row, vbExclamation, "Warning"
            Exit Sub
        End If
        
        grid.Row = grid.Row + 1
   Loop
   'ambil nomor bukti
   OBJ.Open dsn
   SQL = "select top 1 nobkt from am_perminhdr where nobkt like 'NP-' + '" + strformat + "%' order by nobkt desc"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        strno = Right(RST!nobkt, 3)
    Else
        strno = 0
    End If
    
    strno = strno + 1

    If Len(strno) = 1 Then txtnobukti = "NP-" & strformat & "00" & strno
    If Len(strno) = 2 Then txtnobukti = "NP-" & strformat & "0" & strno
    If Len(strno) = 3 Then txtnobukti = "NP-" & strformat & strno
   
    'SIMPAN KE TABLE PERMINHDR

    SQL = "insert into am_perminhdr ("
    SQL = SQL + "nobkt, "
    SQL = SQL + "tglbkt, "
    SQL = SQL + "divisi, "
    SQL = SQL + "pemesan, "
    SQL = SQL + "tglselesai, "
    SQL = SQL + "keterangan, "
    SQL = SQL + "flag, "
    SQL = SQL + "flagpo)"

    SQL = SQL + " values ('" & txtnobukti & "',"
    SQL = SQL + "convert(datetime,'" & tanggal1 & "'),"
    SQL = SQL + "'" & txtbagian & "',"
    SQL = SQL + "'" & txtpemesan & "',"
    SQL = SQL + "convert(datetime,'" & tanggal2 & "'),"
    SQL = SQL + "'" & txtket & "',"
    SQL = SQL + "'0',"
    If cekPO.Value = xtpChecked Then
        SQL = SQL + "'1')"
    Else
        SQL = SQL + "'0')"
    End If
    Set RST = OBJ.Execute(SQL)
    
        'SIMPAN KE TABLE PERMININ
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        SQL = "insert into am_perminin ("
        SQL = SQL + "nobkt, "
        SQL = SQL + "nmbrg, "
        SQL = SQL + "qty, "
        SQL = SQL + "pekerja, "
        SQL = SQL + "keperluan, "
        SQL = SQL + "lineitem, "
        SQL = SQL + "status, "
        SQL = SQL + "nopo, "
        SQL = SQL + "nopo2, "
        SQL = SQL + "tglpo, "
        SQL = SQL + "kdsatuan)"

        SQL = SQL + " values ('" & txtnobukti & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 1) & "',"
        SQL = SQL + "convert(money,'" & Format(grid.TextMatrix(grid.Row, 2), "general number") & "'),"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 5) & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 6) & "',"
        SQL = SQL + "convert(numeric,'" & grid.Row & "'),'0','','',"
        SQL = SQL + "convert(datetime,'" & tanggaldefault & "'),"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 3) & "')"
        Set RST = OBJ.Execute(SQL)
        grid.Row = grid.Row + 1
        DoEvents
    Loop
    
    SQL = "Insert into am_perminapp ("
    SQL = SQL + "nobkt, "
    SQL = SQL + "nopo, "
    SQL = SQL + "kodesupp, "
    SQL = SQL + "status, "
    SQL = SQL + "tglbkt, "
    SQL = SQL + "tglpo, "
    SQL = SQL + "tgl)"
    SQL = SQL + " values('" & txtnobukti & "',"
    SQL = SQL + "'',"
    SQL = SQL + "'','0',"
    SQL = SQL + "convert(datetime,'" & tanggal1 & "'),"
    SQL = SQL + "convert(datetime,'" & tanggal1 & "'),"
    SQL = SQL + "convert(datetime,'" & tanggal1 & "'))"
    Set RST = OBJ.Execute(SQL)
    
    
    OBJ.Close
    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
    Cetak_Bukti
    cmdclear_Click
End Sub

Function tanggal1()
    tanggal1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function
Function tanggal2()
    tanggal2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function

Function tanggaldefault()
    tanggaldefault = Month(dtpdefault) & "/" & Day(dtpdefault) & "/" & Year(dtpdefault)
End Function

Private Sub cmdclear_Click()
    Dim strformat As String
    Dim strno As String
    strformat = Format(Date, "yymm")
    
    hapusgrid
    date1.Value = Date
    date2.Value = Date
    txtpemesan = ""
    txtbagian = ""
    txtket = ""
    
    OBJ.Open dsn
    SQL = "select top 1 nobkt from am_perminhdr where nobkt like 'NP-' + '" + strformat + "%' order by nobkt desc"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        strno = Right(RST!nobkt, 3)
    Else
        strno = 0
    End If
    OBJ.Close
    
    strno = strno + 1

    If Len(strno) = 1 Then txtnobukti = "NP-" & strformat & "00" & strno
    If Len(strno) = 2 Then txtnobukti = "NP-" & strformat & "0" & strno
    If Len(strno) = 3 Then txtnobukti = "NP-" & strformat & strno
    
End Sub

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Dim strformat As String
    Dim strno As String
    strformat = Format(Date, "yymm")
    
    grid.TextMatrix(0, 1) = "Nama Barang"
    grid.TextMatrix(0, 2) = "Qty"
    grid.TextMatrix(0, 3) = "Kode"
    grid.TextMatrix(0, 4) = "Satuan"
    grid.TextMatrix(0, 5) = "Dikerjakan Oleh"
    grid.TextMatrix(0, 6) = "Keperluan"
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 2500
    grid.ColWidth(2) = 1000
    grid.ColWidth(3) = 0
    grid.ColWidth(4) = 1000
    grid.ColWidth(5) = 1500
    grid.ColWidth(6) = 4000
    grid.RowHeightMin = 300
    date1.Value = Date
    date2.Value = Date
    
    OBJ.Open dsn
    SQL = "select top 1 nobkt from am_perminhdr where nobkt like 'NP-' + '" + strformat + "%' order by nobkt desc"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        strno = Right(RST!nobkt, 3)
    Else
        strno = 0
    End If
    OBJ.Close
    
    strno = strno + 1

    If Len(strno) = 1 Then txtnobukti = "NP-" & strformat & "00" & strno
    If Len(strno) = 2 Then txtnobukti = "NP-" & strformat & "0" & strno
    If Len(strno) = 3 Then txtnobukti = "NP-" & strformat & strno

End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    If txtnobukti = "" Then Exit Sub
    
    posrow = grid.Row
    Select Case grid.Col
        Case 0:
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            If grid.CellPicture = uncheck Then
                Set grid.CellPicture = check
                If MsgBox("Delete Row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                    Set grid.CellPicture = uncheck
                    hapusrow
                    Exit Sub
                End If
                Set grid.CellPicture = uncheck
            End If
        Case 1:
            txtbrg.Width = grid.ColWidth(grid.Col) - 40
            txtbrg = grid.TextMatrix(grid.Row, grid.Col)
            txtbrg.Left = grid.Left + grid.CellLeft
            txtbrg.Top = grid.Top + grid.CellTop + 20
            txtbrg.Visible = True
            txtbrg.SetFocus
        Case 2:
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            
            txtnilai.Width = grid.ColWidth(grid.Col) - 40
            txtnilai = grid.TextMatrix(grid.Row, grid.Col)
            txtnilai.Left = grid.Left + grid.CellLeft
            txtnilai.Top = grid.Top + grid.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
        Case 4:
            carisql1 = "select kodesatuan,namasatuan from am_apunit"
            namatabel = "Satuan."
            frmsearch.Show vbModal
        Case 5, 6:
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            
            txtbrg.Width = grid.ColWidth(grid.Col) - 40
            txtbrg = grid.TextMatrix(grid.Row, grid.Col)
            txtbrg.Left = grid.Left + grid.CellLeft
            txtbrg.Top = grid.Top + grid.CellTop + 20
            txtbrg.Visible = True
            txtbrg.SetFocus
    End Select
End Sub

Private Sub SetRow(ByVal idx As Integer, ByVal hapus As String)
    grid.Row = idx
    grid.Col = 0
    If hapus Then Set grid.CellPicture = uncheck.Picture
    grid.Col = 1
End Sub

Private Sub grid_EnterCell()
    If grid.MouseRow = 0 Then Exit Sub
    If txtnobukti = "" Then Exit Sub
    
    Select Case grid.Col
    Case 1
        posrow = grid.Row
        
        txtbrg.Width = grid.ColWidth(grid.Col) - 40
        txtbrg = grid.TextMatrix(grid.Row, grid.Col)
        txtbrg.Left = grid.Left + grid.CellLeft
        txtbrg.Top = grid.Top + grid.CellTop + 20
        txtbrg.Visible = True
        txtbrg.SetFocus
    Case 2
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            
        posrow = grid.Row
        
        txtnilai.Width = grid.ColWidth(grid.Col) - 40
        txtnilai = grid.TextMatrix(grid.Row, grid.Col)
        txtnilai.Left = grid.Left + grid.CellLeft
        txtnilai.Top = grid.Top + grid.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
    Case 4:
        'carisql1 = "select kodesatuan,namasatuan from am_apunit"
        'namatabel = "Satuan."
        'frmsearch.Show vbModal

    Case 5, 6
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
        posrow = grid.Row
        
        txtbrg.Width = grid.ColWidth(grid.Col) - 40
        txtbrg = grid.TextMatrix(grid.Row, grid.Col)
        txtbrg.Left = grid.Left + grid.CellLeft
        txtbrg.Top = grid.Top + grid.CellTop + 20
        txtbrg.Visible = True
        txtbrg.SetFocus
    End Select
End Sub

Private Sub grid_GotFocus()
    Select Case grid.Col
    Case 4
        If hasil = "" Then Exit Sub
        grid.TextMatrix(grid.Row, 3) = hasil
        grid.TextMatrix(grid.Row, 4) = hasil1
        hasil = ""
        hasil1 = ""
    End Select
End Sub

Private Sub txtbrg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    posrow = grid.Row
    Select Case grid.Col
        Case 1:
            If txtbrg = "" Then
                txtbrg.Visible = False
                grid.SetFocus
                grid.Row = posrow
            Exit Sub
            End If
            If Len(Trim(txtbrg)) > 30 Then
                MsgBox "Max.30 karakter", vbExclamation, AppName
                Exit Sub
            End If
            grid.TextMatrix(grid.Row, 1) = txtbrg.text
            SetRow grid.Row, True
            If grid.Row = (grid.Rows - 1) Then grid.Rows = grid.Rows + 1
            grid.SetFocus
        Case 5:
            If Len(Trim(txtbrg)) > 30 Then
                MsgBox "Max.30 karakter", vbExclamation, AppName
                Exit Sub
            End If
            grid.TextMatrix(grid.Row, 5) = txtbrg.text
        Case 6:
            grid.TextMatrix(grid.Row, 6) = txtbrg.text
    End Select
        txtbrg = ""
        txtbrg.Visible = False
        grid.SetFocus
        grid.Row = posrow
    End If
End Sub

Private Sub txtbrg_LostFocus()
    txtbrg.Visible = False
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grid.TextMatrix(grid.Row, 2) = Format(txtnilai, "#,##0.00")
        txtnilai = 0
        txtnilai.Visible = False
        grid.SetFocus
        grid.Row = posrow
    End If
End Sub


Private Sub hapusrow()
    grid.TextMatrix(grid.Row, 1) = ""
    grid.TextMatrix(grid.Row, 2) = ""
    grid.TextMatrix(grid.Row, 3) = ""
    grid.TextMatrix(grid.Row, 4) = ""
    grid.TextMatrix(grid.Row, 5) = ""
    grid.TextMatrix(grid.Row, 6) = ""
    Do While True
        If grid.TextMatrix(grid.Row + 1, 1) = "" Then
            grid.TextMatrix(grid.Row, 1) = ""
            grid.TextMatrix(grid.Row, 2) = ""
            grid.TextMatrix(grid.Row, 3) = ""
            grid.TextMatrix(grid.Row, 4) = ""
            grid.TextMatrix(grid.Row, 5) = ""
            grid.TextMatrix(grid.Row, 6) = ""
            Exit Do
        End If
        grid.TextMatrix(grid.Row, 1) = grid.TextMatrix(grid.Row + 1, 1)
        grid.TextMatrix(grid.Row, 2) = grid.TextMatrix(grid.Row + 1, 2)
        grid.TextMatrix(grid.Row, 3) = grid.TextMatrix(grid.Row + 1, 3)
        grid.TextMatrix(grid.Row, 4) = grid.TextMatrix(grid.Row + 1, 4)
        grid.TextMatrix(grid.Row, 5) = grid.TextMatrix(grid.Row + 1, 5)
        grid.TextMatrix(grid.Row, 6) = grid.TextMatrix(grid.Row + 1, 6)
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = grid.Rows - 1
    grid.Col = 0
    Set grid.CellPicture = blank
End Sub

Private Sub hapusgrid()
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        grid.TextMatrix(grid.Row, 1) = ""
        grid.TextMatrix(grid.Row, 2) = ""
        grid.TextMatrix(grid.Row, 3) = ""
        grid.TextMatrix(grid.Row, 4) = ""
        grid.TextMatrix(grid.Row, 5) = ""
        grid.TextMatrix(grid.Row, 6) = ""
        grid.Col = 0
        Set grid.CellPicture = blank
        grid.Row = grid.Row + 1
    Loop
    
    grid.Rows = 2
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 2500
    grid.ColWidth(2) = 1000
    grid.ColWidth(3) = 0
    grid.ColWidth(4) = 1000
    grid.ColWidth(5) = 1500
    grid.ColWidth(6) = 4000
End Sub

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
End Sub

Private Sub Cetak_Bukti()
    Crystal.Reset
    Crystal.WindowState = crptMaximized
    Crystal.WindowShowCloseBtn = True
    Crystal.WindowShowPrintSetupBtn = True
    Crystal.WindowShowSearchBtn = True
    'Crystal.Destination = crptToPrinter
    Crystal.Connect = dsnreport
    Crystal.DataFiles(0) = "Proc(am_nota)"
    Crystal.ReportFileName = AppPath & "\reports\purchasing\purc\nota.rpt"
    Crystal.ParameterFields(0) = "@nobkt;" + txtnobukti + ";true"
    Crystal.RetrieveDataFiles
    Crystal.Action = 1
    
    Exit Sub
    
    With rptpermintaan
         SQL1 = "Exec am_printpermin '" & txtnobukti & "'"
        .DataControl1.Source = SQL1
        .DataControl1.ConnectionString = dsn
        .Show
    End With
End Sub


