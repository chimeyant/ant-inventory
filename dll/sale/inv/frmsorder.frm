VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmsorder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Sales Order"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
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
   Icon            =   "frmsorder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtlimit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5295
      TabIndex        =   28
      Text            =   "0"
      Top             =   435
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.TextBox txttotal_piutang 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7290
      TabIndex        =   27
      Text            =   "0"
      Top             =   435
      Visible         =   0   'False
      Width           =   1950
   End
   Begin Chameleon.chameleonButton cmdtes 
      Height          =   375
      Left            =   8430
      TabIndex        =   26
      Top             =   75
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Tes"
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
      MICON           =   "frmsorder.frx":2372
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtketerangan 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      MaxLength       =   200
      TabIndex        =   5
      Top             =   1920
      Width           =   7575
   End
   Begin TDBText6Ctl.TDBText txtket 
      Height          =   255
      Left            =   7320
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   450
      Caption         =   "frmsorder.frx":268C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmsorder.frx":26F8
      Key             =   "frmsorder.frx":2716
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
      BorderStyle     =   0
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   50
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
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   7320
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   450
      Calculator      =   "frmsorder.frx":2752
      Caption         =   "frmsorder.frx":2772
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmsorder.frx":27DE
      Keys            =   "frmsorder.frx":27FC
      Spin            =   "frmsorder.frx":283E
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
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.TextBox txtsales 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtkodecust 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtnobukti 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtapply 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1560
      Width           =   2055
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
      Left            =   480
      Picture         =   "frmsorder.frx":2866
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   14
      Top             =   1200
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
      Left            =   720
      Picture         =   "frmsorder.frx":2BB4
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   13
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
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
      Left            =   240
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
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
      Format          =   120913923
      CurrentDate     =   37426
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   2175
      Left            =   0
      TabIndex        =   8
      Top             =   2280
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   3836
      _Version        =   393216
      Cols            =   6
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
      _Band(0).Cols   =   6
   End
   Begin Chameleon.chameleonButton cmdsearch5 
      Height          =   285
      Left            =   3720
      TabIndex        =   21
      Top             =   1560
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Salesman"
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
      MICON           =   "frmsorder.frx":2E96
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch1 
      Height          =   285
      Left            =   240
      TabIndex        =   22
      Top             =   840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Customer"
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
      MICON           =   "frmsorder.frx":31B0
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
      Left            =   8280
      TabIndex        =   11
      Top             =   4560
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
      MICON           =   "frmsorder.frx":34CA
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
      Left            =   7320
      TabIndex        =   10
      Top             =   4560
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
      MICON           =   "frmsorder.frx":37E4
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
      Left            =   6360
      TabIndex        =   9
      Top             =   4560
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
      MICON           =   "frmsorder.frx":3AFE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      Caption         =   "Jasa Angkutan"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   1950
      Width           =   1095
   End
   Begin VB.Label lblalamatcust 
      BackColor       =   &H80000014&
      Height          =   255
      Left            =   1560
      TabIndex        =   24
      Top             =   1200
      Width           =   7575
   End
   Begin VB.Label lblsat 
      Caption         =   "    Nama Satuan :"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   4770
      Width           =   5385
   End
   Begin VB.Label lblsales 
      BackColor       =   &H80000014&
      Height          =   255
      Left            =   6120
      TabIndex        =   20
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "No Sales Order"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   150
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "No PO Customer"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   1590
      Width           =   1455
   End
   Begin VB.Label lblnamacust 
      BackColor       =   &H80000014&
      Height          =   255
      Left            =   2880
      TabIndex        =   17
      Top             =   840
      Width           =   6255
   End
   Begin VB.Label Label13 
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   510
      Width           =   1455
   End
   Begin VB.Label lblitem 
      Caption         =   "    Nama Barang :"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   4530
      Width           =   5385
   End
End
Attribute VB_Name = "frmsorder"
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
Dim SP As New ADODB.Command

Dim posrow, poscol, str99, str21, str1 As String
Dim int3 As Integer
Dim boo1 As Boolean
Dim str_empid, tgl As String

Private Sub cmdadd_Click()
    If Len(Trim(txtnobukti)) = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        txtnobukti.SetFocus
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_period where tanggal1 <= '" & tanggalinv & "' and tanggal2 >= '" & tanggalinv & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        OBJ.Close
        MsgBox "Can not add, out of date range.", vbExclamation, "Warning"
        Exit Sub
    End If
    OBJ.Close

    txtnobukti = Trim(txtnobukti)

    If txtnobukti = "" Or txtsales = "" Or txtkodecust = "" Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If

    If txtapply = "" Then
        If MsgBox("Continue with blank PO number ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    End If

    If grid.Rows = 2 Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If

    grid.Row = 1
    Do While True
        If grid.Rows = grid.Row + 1 Then Exit Do
        If grid.TextMatrix(grid.Row, 3) = "0.00" Or grid.TextMatrix(grid.Row, 4) = "" Or Val(Format(grid.TextMatrix(grid.Row, 3), "general number")) < Val(Format(grid.TextMatrix(grid.Row, 5), "general number")) Then
            MsgBox "Data Entry Not Complete, On Row " & grid.Row, vbExclamation, "Warning"
            Exit Sub
        End If
        
        boo1 = False
        OBJ.Open dsn
        SQL = "select a.noso from am_solin a left join am_sohdr b on a.noso=b.noso and a.tglso=b.tglso "
        SQL = SQL + "where a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and b.kodecust = '" & txtkodecust & "'"
        SQL = SQL + " and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "' and b.tglso = '" & tanggalinv & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then boo1 = True
        
        SQL = "select noso from am_soapp "
        SQL = SQL + "where kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and kodecust = '" & txtkodecust & "'"
        SQL = SQL + " and kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "' and tglso = '" & tanggalinv & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then boo1 = True
        OBJ.Close
        
        If boo1 Then
            If MsgBox("Sales Order untuk tanggal, customer, item, dan satuan yang sama sudah ada." & vbCrLf & _
            "Program sudah memperingatkan user, lanjutkan input Sales Order ?", vbYesNo + vbExclamation, "Warning") = vbNo Then Exit Sub
        End If
        
        grid.Row = grid.Row + 1
    Loop
    
    int3 = 0
    OBJ.Open dsn
    SQL = "select noso from am_sohdr where noso = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        
        If Asc(Mid(txtnobukti, 1, 1)) = 80 Then
            OBJ.Open dsn
            SQL = "select top 1 noso from am_sohdr where noso like 'P-" & str1 & "%' order by noso desc"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                str99 = Right(RST!noso, 5)
            Else
                str99 = 0
            End If
            OBJ.Close
            
            str99 = str99 + 1
        
            If Len(str99) = 1 Then txtnobukti = "P-" & str1 & "0000" & str99
            If Len(str99) = 2 Then txtnobukti = "P-" & str1 & "000" & str99
            If Len(str99) = 3 Then txtnobukti = "P-" & str1 & "00" & str99
            If Len(str99) = 4 Then txtnobukti = "P-" & str1 & "0" & str99
            If Len(str99) = 5 Then txtnobukti = "P-" & str1 & str99
        ElseIf Asc(Mid(txtnobukti, 1, 1)) = 76 Then
            OBJ.Open dsn
            SQL = "select top 1 noso from am_sohdr where noso like 'L-" & str1 & "%' order by noso desc"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                str99 = Right(RST!noso, 5)
            Else
                str99 = 0
            End If
            OBJ.Close
            
            str99 = str99 + 1
        
            If Len(str99) = 1 Then txtnobukti = "L-" & str1 & "0000" & str99
            If Len(str99) = 2 Then txtnobukti = "L-" & str1 & "000" & str99
            If Len(str99) = 3 Then txtnobukti = "L-" & str1 & "00" & str99
            If Len(str99) = 4 Then txtnobukti = "L-" & str1 & "0" & str99
            If Len(str99) = 5 Then txtnobukti = "L-" & str1 & str99
        End If
        
        If Left(txtnobukti, 2) = "PP" Then txtnobukti = Mid(txtnobukti, 2, 8)
        If Left(txtnobukti, 2) = "LL" Then txtnobukti = Mid(txtnobukti, 2, 8)
        int3 = 1
    Else
        OBJ.Close
    End If

    OBJ.Open dsn
    SQL = "select noso from am_sohdr where noso = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        
        If Asc(Mid(txtnobukti, 1, 1)) = 80 Then
            OBJ.Open dsn
            SQL = "select top 1 noso from am_sohdr where noso like 'P-" & str1 & "%' order by noso desc"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                str99 = Right(RST!noso, 5)
            Else
                str99 = 0
            End If
            OBJ.Close
            
            str99 = str99 + 1
        
            If Len(str99) = 1 Then txtnobukti = "P-" & str1 & "0000" & str99
            If Len(str99) = 2 Then txtnobukti = "P-" & str1 & "000" & str99
            If Len(str99) = 3 Then txtnobukti = "P-" & str1 & "00" & str99
            If Len(str99) = 4 Then txtnobukti = "P-" & str1 & "0" & str99
            If Len(str99) = 5 Then txtnobukti = "P-" & str1 & str99
        ElseIf Asc(Mid(txtnobukti, 1, 1)) = 76 Then
            OBJ.Open dsn
            SQL = "select top 1 noso from am_sohdr where noso like 'L-" & str1 & "%' order by noso desc"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                str99 = Right(RST!noso, 5)
            Else
                str99 = 0
            End If
            OBJ.Close
            
            str99 = str99 + 1
        
            If Len(str99) = 1 Then txtnobukti = "L-" & str1 & "0000" & str99
            If Len(str99) = 2 Then txtnobukti = "L-" & str1 & "000" & str99
            If Len(str99) = 3 Then txtnobukti = "L-" & str1 & "00" & str99
            If Len(str99) = 4 Then txtnobukti = "L-" & str1 & "0" & str99
            If Len(str99) = 5 Then txtnobukti = "L-" & str1 & str99
        End If
        If Left(txtnobukti, 2) = "PP" Then txtnobukti = Mid(txtnobukti, 2, 8)
        If Left(txtnobukti, 2) = "LL" Then txtnobukti = Mid(txtnobukti, 2, 8)
        int3 = 1
    Else
        OBJ.Close
    End If

    OBJ.Open dsn
    SQL = "select * from am_sohdr where noso = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close

        MsgBox "Data Already Exist, save aborted, Click OK To Continue ...", vbInformation, "Information"
        cmdclear_Click

        Exit Sub
    End If
    OBJ.Close

    OBJ.Open dsn
    SQL = "insert into am_sohdr ("
    SQL = SQL + "noso,"
    SQL = SQL + "tglso,"
    SQL = SQL + "kodecust,"
    SQL = SQL + "kodesales,"
    SQL = SQL + "nopo,"
    SQL = SQL + "jasa,"
    SQL = SQL + "flag,"
    SQL = SQL + "identry,"
    SQL = SQL + "dateentry,"
    SQL = SQL + "idupdate,"
    SQL = SQL + "dateupdate)"

    SQL = SQL + " values("
    SQL = SQL + "'" & txtnobukti & "',"
    SQL = SQL + "convert(datetime,'" & tanggalinv & "'),"
    SQL = SQL + "'" & txtkodecust & "',"
    SQL = SQL + "'" & txtsales & "',"
    SQL = SQL + "'" & txtapply & "',"
    SQL = SQL + "'" & txtketerangan & "',"
    If par3 = "1" Then SQL = SQL + "'1'," Else SQL = SQL + "'0',"
    SQL = SQL + "'" & kuser & "',"
    SQL = SQL + "convert(datetime,'" & tanggalsekarang & "'),"
    SQL = SQL + "'',"
    SQL = SQL + "convert(datetime,'" & tanggalsekarang & "'))"
    Set RST = OBJ.Execute(SQL)

    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        SQL = "insert into am_solin ("
        SQL = SQL + "noso,"
        SQL = SQL + "tglso,"
        SQL = SQL + "kodebarang,"
        SQL = SQL + "qty,"
        SQL = SQL + "keterangan,"
        SQL = SQL + "lineitem,"
        SQL = SQL + "kodesatuan,"
        SQL = SQL + "BN)"

        SQL = SQL + " values("
        SQL = SQL + "'" & txtnobukti & "',"
        SQL = SQL + "convert(datetime,'" & tanggalinv & "'),"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 1) & "',"
        SQL = SQL + "convert(money,'" & Format(grid.TextMatrix(grid.Row, 3), "general number") & "'),"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 2) & "',"
        SQL = SQL + "convert(numeric,'" & grid.Row & "'),"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 4) & "',"
        SQL = SQL + "convert(money,'" & Format(grid.TextMatrix(grid.Row, 5), "general number") & "'))"
        Set RST = OBJ.Execute(SQL)

        grid.Row = grid.Row + 1
    Loop
    OBJ.Close
    
    If par3 = "1" Then 'par3 kalo 1 langsung masuk soapp, jadi tidak perlu di cek apakah uda export/belum
        OBJ.Open dsn
        SQL = "SELECT b.noso,b.tglso,b.kodecust,b.kodesales,b.nopo,b.jasa,a.kodebarang,a.qty,a.keterangan,a.kodesatuan,a.lineitem,a.bn FROM am_solin a left join AM_sohdr b on a.noso=b.noso WHERE b.noso = '" & txtnobukti & "'"
        Set RST = OBJ.Execute(SQL)
        Do While Not RST.EOF
            OBJ1.Open dsn
            SQL1 = "INSERT INTO AM_soapp"
            SQL1 = SQL1 + " (noso"
            SQL1 = SQL1 + ", Tglso"
            SQL1 = SQL1 + ", kodecust"
            SQL1 = SQL1 + ", kodesales"
            SQL1 = SQL1 + ", nopo"
            SQL1 = SQL1 + ", jasa"
            SQL1 = SQL1 + ", Kodebarang"
            SQL1 = SQL1 + ", qty"
            SQL1 = SQL1 + ", keterangan"
            SQL1 = SQL1 + ", kodesatuan"
            SQL1 = SQL1 + ", lineitem"
            SQL1 = SQL1 + ", bn"
            SQL1 = SQL1 + ", flag1"
            SQL1 = SQL1 + ", flag2)"
        
            SQL1 = SQL1 + " VALUES"
            SQL1 = SQL1 + " ('" & RST!noso & "'"
            SQL1 = SQL1 + ",Convert(dateTime, '" & Month(RST!tglso) & "/" & Day(RST!tglso) & "/" & Year(RST!tglso) & "')"
            SQL1 = SQL1 + ", '" & RST!kodecust & "'"
            SQL1 = SQL1 + ", '" & RST!kodesales & "'"
            SQL1 = SQL1 + ", '" & RST!nopo & "'"
            SQL1 = SQL1 + ", '" & RST!jasa & "'"
            SQL1 = SQL1 + ", '" & RST!kodebarang & "'"
            SQL1 = SQL1 + ",Convert (Money, '" & RST!qty & "')"
            SQL1 = SQL1 + ", '" & RST!keterangan & "'"
            SQL1 = SQL1 + ", '" & RST!kodesatuan & "'"
            SQL1 = SQL1 + ",Convert (numeric, '" & RST!lineitem & "')"
            SQL1 = SQL1 + ",Convert (Money, '" & RST!bn & "')"
            SQL1 = SQL1 + ", '1'"
            SQL1 = SQL1 + ", '0')"
            Set RST1 = OBJ1.Execute(SQL1)
            OBJ1.Close
        
            RST.MoveNext
        Loop
        OBJ.Close
    End If

    If int3 = 1 Then
        MsgBox "Data already exist, data was saved with next number " & txtnobukti & vbCrLf & _
        "Click OK To Continue ...", vbExclamation, "Warning"
    Else
        MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
    End If
    cmdclear_Click
End Sub

Private Sub cmdclear_Click()
    hapusemua
    
    txtnobukti = ""
    txtnobukti.SetFocus
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdsearch1_Click()
    carisql1 = "select kodecust, namacust, alamatcust, kota from am_customer"
    namatabel = "Customer"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtkodecust = hasil
    caricustomer
    
    'Blokir Customer dengan piutang limit
    'OBJ.Open dsn
    'SQL = "Select * From am_customer Where kodecust='" & hasil & "'"
    'Set RST = OBJ.Execute(SQL)
    
    'If RST!status = "1" Then
        'MsgBox "Customer terblokir, silahkan konfirmasi kepusat untuk membuka blokir", vbCritical, AppName
        'OBJ.Close
        'cmdclear_Click
        'Exit Sub
    'End If
    'OBJ.Close
    
    cmdtes_Click 'Limit Piutang
    txtapply.SetFocus
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdsearch5_Click()
    carisql1 = "select kodesales, namasales,idupdate from AM_salesman"
    namatabel = "Sales"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch5_GotFocus()
    If hasil = "" Then Exit Sub
    txtsales = hasil
    carisales
    hasil = ""
    hasil1 = ""
    txtketerangan.SetFocus
End Sub

Private Sub cmdtes_Click()
Dim limit As Boolean
    limit = False
    str_empid = txtkodecust.text
    Set SP = New ADODB.Command
    SP.ActiveConnection = dsn
    SP.CommandType = adCmdStoredProc
    SP.CommandText = "am_piutang_total"

    SP.Parameters.Append SP.CreateParameter("empid", adVarChar, adParamInput, 10, str_empid)
    Set RST = SP.Execute
        
    If Not RST.EOF Then
        txttotal_piutang.text = RST.Fields(2)
        txttotal_piutang = Format(RST.Fields(2), "#,##0.00")
    End If

    Set SP.ActiveConnection = Nothing
    
    OBJ.Open dsn
    SQL = "Select limit From am_customer Where KodeCust='" & txtkodecust & "' and flaglimit='1'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtlimit = Format(RST!limit, "#,##0.00")
        limit = True
    End If
    OBJ.Close
    
    If limit = True Then
        If txttotal_piutang > txtlimit Then
            frmwarning.lblnamacust = lblnamacust
            frmwarning.Show vbModal
            cmdclear_Click
        End If
    End If
    
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
    
    'VALIDASI USER
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='141' and b.kodeuser = '1" & kuser & "'"
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
     
    grid.TextMatrix(0, 1) = "Kode Barang"
    grid.TextMatrix(0, 2) = "Keterangan"
    grid.TextMatrix(0, 3) = "Quantity"
    grid.TextMatrix(0, 4) = "Satuan"
    grid.TextMatrix(0, 5) = "BN"
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1500
    grid.ColWidth(2) = 4000
    grid.ColWidth(3) = 1500
    grid.ColWidth(4) = 1000
    grid.ColWidth(5) = 500
    grid.RowHeightMin = 300
    
    date1.Value = Date
    
    OBJ.Open dsn
    SQL = "select id3, kode3 from am_branch"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If RST!id3 = "1" Then str1 = RST!kode3 Else str1 = "0"
    Else
        str1 = "0"
    End If
    OBJ.Close
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    If grid.TextMatrix(grid.Row, 1) <> "" Then
        OBJ.Open dsn
        SQL = "SELECT b.kodesatuan,b.namasatuan,a.namabarang FROM AM_ITEMDTL a left join am_unit b ON a.kodesatuan=b.kodesatuan WHERE a.KodeBarang = '" & grid.TextMatrix(grid.Row, 1) & "' and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            lblitem = "    Nama Barang : " & RST!namabarang
            lblsat = "    Nama Satuan : " & RST!namasatuan
        Else
            lblitem = "    Nama Barang : "
            lblsat = "    Nama Satuan : "
        End If
        OBJ.Close
    End If
    
    If txtnobukti = "" Or txtkodecust = "" Or txtsales = "" Then Exit Sub
    posrow = grid.Row
    poscol = grid.Col
    Select Case grid.Col
        Case 0
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            If grid.CellPicture = uncheck Then
                Set grid.CellPicture = check
                If MsgBox("Delete That Row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                    Set grid.CellPicture = uncheck
                    hapusrow
                    Exit Sub
                End If
                Set grid.CellPicture = uncheck
            End If
        Case 1
            If grid.TextMatrix(grid.Row, 1) <> "" Then Exit Sub
            If grid.Rows - 1 = 50 Then
                MsgBox "Maximum line 50", vbExclamation, "Warning"
                Exit Sub
            End If

            If txtket.Visible = True Then Exit Sub

            txtket.Width = grid.ColWidth(grid.Col) - 40
            txtket = grid.TextMatrix(grid.Row, grid.Col)
            txtket.Left = grid.Left + grid.CellLeft
            txtket.Top = grid.Top + grid.CellTop + 20
            txtket.Visible = True
            txtket.SetFocus
        Case 2, 4
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub

            If txtket.Visible = True Then Exit Sub

            txtket.Width = grid.ColWidth(grid.Col) - 40
            txtket = grid.TextMatrix(grid.Row, grid.Col)
            txtket.Left = grid.Left + grid.CellLeft
            txtket.Top = grid.Top + grid.CellTop + 20
            txtket.Visible = True
            txtket.SetFocus
        Case 3
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub

            txtnilai.Width = grid.ColWidth(grid.Col) - 40
            txtnilai = grid.TextMatrix(grid.Row, grid.Col)
            txtnilai.Left = grid.Left + grid.CellLeft
            txtnilai.Top = grid.Top + grid.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
        Case 5
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            
            OBJ.Open dsn
            SQL = "SELECT kodeproduk FROM AM_ITEMmst WHERE KodeBarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                If RST!kodeproduk = "C999" Then
                    OBJ.Close
                    Exit Sub
                Else
                    OBJ.Close
                End If
            Else
                OBJ.Close
                Exit Sub
            End If

            txtnilai.Width = grid.ColWidth(grid.Col) - 40
            txtnilai = grid.TextMatrix(grid.Row, grid.Col)
            txtnilai.Left = grid.Left + grid.CellLeft
            txtnilai.Top = grid.Top + grid.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
    End Select
End Sub

Private Sub grid_EnterCell()
    If txtnobukti = "" Or txtkodecust = "" Or txtsales = "" Then Exit Sub
    Select Case grid.Col
    Case 1
        If grid.TextMatrix(grid.Row, 1) <> "" Then Exit Sub
        If txtket.Visible = True Then Exit Sub

        posrow = grid.Row
        poscol = grid.Col
        txtket.Width = grid.ColWidth(grid.Col) - 40
        txtket = grid.TextMatrix(grid.Row, grid.Col)
        txtket.Left = grid.Left + grid.CellLeft
        txtket.Top = grid.Top + grid.CellTop + 20
        txtket.Visible = True
        txtket.SetFocus
    Case 2, 4
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
        If txtket.Visible = True Then Exit Sub

        posrow = grid.Row
        poscol = grid.Col
        txtket.Width = grid.ColWidth(grid.Col) - 40
        txtket = grid.TextMatrix(grid.Row, grid.Col)
        txtket.Left = grid.Left + grid.CellLeft
        txtket.Top = grid.Top + grid.CellTop + 20
        txtket.Visible = True
        txtket.SetFocus
    Case 3
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub

        posrow = grid.Row
        poscol = grid.Col
        txtnilai.Width = grid.ColWidth(grid.Col) - 40
        txtnilai = grid.TextMatrix(grid.Row, grid.Col)
        txtnilai.Left = grid.Left + grid.CellLeft
        txtnilai.Top = grid.Top + grid.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
    Case 5
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
        
        OBJ.Open dsn
        SQL = "SELECT kodeproduk FROM AM_ITEMmst WHERE KodeBarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            If RST!kodeproduk = "C999" Then
                OBJ.Close
                Exit Sub
            Else
                OBJ.Close
            End If
        Else
            OBJ.Close
            Exit Sub
        End If

        posrow = grid.Row
        poscol = grid.Col
        txtnilai.Width = grid.ColWidth(grid.Col) - 40
        txtnilai = grid.TextMatrix(grid.Row, grid.Col)
        txtnilai.Left = grid.Left + grid.CellLeft
        txtnilai.Top = grid.Top + grid.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
    End Select
End Sub

Private Sub grid_GotFocus()
    If hasil = "" Then Exit Sub
    
    If grid.Col = 4 Then
        grid.Row = 1
        Do While True
            If grid.Rows = 2 Or grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
            
            If grid.TextMatrix(grid.Row, 4) = hasil And grid.TextMatrix(grid.Row, 1) = grid.TextMatrix(posrow, 1) And posrow <> grid.Row Then
                MsgBox "Kode Barang already exist.", vbInformation, "Information"
                hasil = ""
                grid.Row = posrow
                grid.Col = 4
                grid.SetFocus
                Exit Sub
            End If
            grid.Row = grid.Row + 1
        Loop
    End If
    
    grid.Row = posrow
    grid.Col = poscol
    grid.CellAlignment = 1
    grid.TextMatrix(grid.Row, grid.Col) = hasil
    hasil = ""
    hasil1 = ""
    hasil2 = ""

    If grid.Col = 1 Then
        OBJ.Open dsn
        SQL = "select * from am_itemmst where kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            grid.TextMatrix(grid.Row, 3) = "0.00"
            grid.TextMatrix(grid.Row, 5) = "0.00"
            
            lblitem = "    Nama Barang : " & RST!namabarang

            SetRow grid.Row, True
            grid.SetFocus
            grid.Col = 2
            If grid.Row = (grid.Rows - 1) Then grid.Rows = grid.Rows + 1
        Else
            MsgBox "Item Not Found", vbExclamation, "Warning"
            grid.TextMatrix(grid.Row, 1) = ""
        End If
        OBJ.Close
    ElseIf grid.Col = 4 Then
        OBJ.Open dsn
        SQL = "SELECT b.kodesatuan,b.namasatuan,a.namabarang FROM AM_ITEMDTL a left join am_unit b ON a.kodesatuan=b.kodesatuan WHERE a.KodeBarang = '" & grid.TextMatrix(grid.Row, 1) & "' and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            lblsat = "    Nama Satuan : " & RST!namasatuan
            lblitem = "    Nama Barang : " & RST!namabarang

            grid.SetFocus
            grid.Col = 5
        Else
            lblsat = "    Nama Satuan :"
            lblitem = "    Nama Barang :"
            
            grid.TextMatrix(grid.Row, 4) = ""
        End If
        OBJ.Close
    End If
End Sub

Private Sub grid_Scroll()
    txtket.Visible = False
    txtnilai.Visible = False
End Sub

Private Sub txtapply_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtsales.SetFocus
End Sub

Private Sub txtket_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0

    If KeyAscii = 13 Then
        Select Case grid.Col
            Case 1
                grid.Row = posrow

                grid.SetFocus
                grid.Col = 1
                grid.CellAlignment = 1
                grid.TextMatrix(grid.Row, 1) = txtket
                txtket = ""
                txtket.Visible = False

                OBJ.Open dsn
                SQL = "select * from am_itemmst where kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "' and len(kodebarang)=8"
                Set RST = OBJ.Execute(SQL)
                If Not RST.EOF Then
                    grid.TextMatrix(grid.Row, 3) = "0.00"
                    grid.TextMatrix(grid.Row, 5) = "0.00"
                    
                    lblitem = "    Nama Barang : " & RST!namabarang
                    OBJ.Close
                    grid.Col = 0
                    Set grid.CellPicture = uncheck.Picture
                    If grid.Row = (grid.Rows - 1) Then grid.Rows = grid.Rows + 1
                Else
                    OBJ.Close
                    grid.TextMatrix(posrow, 1) = ""
                    txtket = ""

                    carisql1 = "select kodebarang, namabarang from am_itemmst"
                    namatabel = "Item"

                    frmsearch.Show vbModal
                End If
                grid.Col = 2
            Case 4
                grid.Row = 1
                Do While True
                    If grid.Rows = 2 Or grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
                    
                    If grid.TextMatrix(grid.Row, 4) = txtket And grid.TextMatrix(grid.Row, 1) = grid.TextMatrix(posrow, 1) And posrow <> grid.Row Then
                        MsgBox "Kode Barang already exist.", vbInformation, "Information"
                        txtket = ""
                        grid.Row = posrow
                        grid.Col = 4
                        grid.SetFocus
                        Exit Sub
                    End If
                    grid.Row = grid.Row + 1
                Loop

                grid.Row = posrow

                grid.SetFocus
                grid.Col = 4
                grid.CellAlignment = 1
                grid.TextMatrix(grid.Row, 4) = txtket
                txtket = ""
                txtket.Visible = False

                OBJ.Open dsn
                SQL = "SELECT b.kodesatuan,b.namasatuan,a.namabarang FROM AM_ITEMDTL a left join am_unit b ON a.kodesatuan=b.kodesatuan WHERE a.KodeBarang = '" & grid.TextMatrix(grid.Row, 1) & "' and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
                Set RST = OBJ.Execute(SQL)
                If RST.EOF Then
                    grid.TextMatrix(posrow, 4) = ""
                    lblsat = "    Nama Satuan :"
                    lblitem = "    Nama Barang :"

                    txtket = ""

                    carisql1 = "SELECT b.kodesatuan,b.namasatuan FROM AM_ITEMDTL a left join am_unit b on a.kodesatuan = b.kodesatuan WHERE a.KodeBarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
                    namatabel = "Satuan "

                    frmsearch.Show vbModal
                Else
                    lblsat = "    Nama Satuan : " & RST!namasatuan
                    lblitem = "    Nama Barang : " & RST!namabarang
                End If
                OBJ.Close
                grid.Col = 5
            Case 2
                grid.SetFocus
                grid.Col = 2
                grid.CellAlignment = 1
                grid.TextMatrix(grid.Row, 2) = txtket
                txtket = ""
                txtket.Visible = False
        End Select
    ElseIf KeyAscii = 27 Then
        txtket.Visible = False
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtket_LostFocus()
    txtket.Visible = False
End Sub

Private Sub txtkodecust_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtapply.SetFocus
End Sub

Private Sub txtkodecust_LostFocus()
    caricustomer
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grid.TextMatrix(grid.Row, grid.Col) = Format(txtnilai, "###,###,##0.00")
        txtnilai = 0

        txtnilai.Visible = False
        grid.SetFocus
        grid.Row = posrow
        grid.Col = poscol
    ElseIf KeyAscii = 27 Then
        txtnilai.Visible = False
    End If
End Sub

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
End Sub

Private Sub txtnobukti_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then txtnobukti = ""
End Sub

Private Sub txtnobukti_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then date1.SetFocus
    If Len(txtnobukti) > 1 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 80 Then
        OBJ.Open dsn
        SQL = "select top 1 noso from am_sohdr where noso like 'P-" & str1 & "%' order by noso desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            str99 = Right(RST!noso, 5)
        Else
            str99 = 0
        End If
        OBJ.Close
        
        str99 = str99 + 1
    
        If Len(str99) = 1 Then txtnobukti = "P-" & str1 & "0000" & str99
        If Len(str99) = 2 Then txtnobukti = "P-" & str1 & "000" & str99
        If Len(str99) = 3 Then txtnobukti = "P-" & str1 & "00" & str99
        If Len(str99) = 4 Then txtnobukti = "P-" & str1 & "0" & str99
        If Len(str99) = 5 Then txtnobukti = "P-" & str1 & str99
        
    ElseIf KeyAscii = 76 Then 'NON PAJAK
        OBJ.Open dsn
        SQL = "select top 1 noso from am_sohdr where noso like 'L-" & str1 & "%' order by noso desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            str99 = Right(RST!noso, 5)
        Else
            str99 = 0
        End If
        OBJ.Close
       
        str99 = str99 + 1
    
        If Len(str99) = 1 Then txtnobukti = "L-" & str1 & "0000" & str99
        If Len(str99) = 2 Then txtnobukti = "L-" & str1 & "000" & str99
        If Len(str99) = 3 Then txtnobukti = "L-" & str1 & "00" & str99
        If Len(str99) = 4 Then txtnobukti = "L-" & str1 & "0" & str99
        If Len(str99) = 5 Then txtnobukti = "L-" & str1 & str99
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtnobukti_KeyUp(KeyCode As Integer, Shift As Integer)
    If Left(txtnobukti, 2) = "PP" Then txtnobukti = Mid(txtnobukti, 2, 8)
    If Left(txtnobukti, 2) = "LL" Then txtnobukti = Mid(txtnobukti, 2, 8)
End Sub

Private Sub txtsales_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtketerangan.SetFocus
End Sub

Private Sub txtsales_LostFocus()
    carisales
End Sub

Private Sub SetRow(ByVal idx As Integer, ByVal hapus As String)
    grid.Row = idx
    grid.Col = 0
    If hapus Then Set grid.CellPicture = uncheck.Picture
    grid.Col = 1
End Sub

Private Sub hapusemua()
    If Date > date1.MaxDate Then
        date1 = date1.MaxDate
    ElseIf Date < date1.MinDate Then
        date1 = date1.MinDate
    Else
        date1 = Date
    End If
    txtkodecust = ""
    lblnamacust = ""
    lblalamatcust = ""
    txtsales = ""
    lblsales = ""
    txtapply = ""
    txtketerangan = ""
    txtlimit = "0"
    txttotal_piutang = "0"

    hapusgrid

    lblitem = "    Nama Barang : "
    lblsat = "    Nama Satuan : "
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
        grid.Col = 0
        Set grid.CellPicture = blank
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1500
    grid.ColWidth(2) = 4000
    grid.ColWidth(3) = 1500
    grid.ColWidth(4) = 1000
    grid.ColWidth(5) = 500
End Sub

Private Sub hapusrow()
    grid.TextMatrix(grid.Row, 1) = ""
    grid.TextMatrix(grid.Row, 2) = ""
    grid.TextMatrix(grid.Row, 3) = ""
    grid.TextMatrix(grid.Row, 4) = ""
    grid.TextMatrix(grid.Row, 5) = ""
    Do While True
        If grid.TextMatrix(grid.Row + 1, 1) = "" Then
            grid.TextMatrix(grid.Row, 1) = ""
            grid.TextMatrix(grid.Row, 2) = ""
            grid.TextMatrix(grid.Row, 3) = ""
            grid.TextMatrix(grid.Row, 4) = ""
            grid.TextMatrix(grid.Row, 5) = ""
            Exit Do
        End If
        grid.TextMatrix(grid.Row, 1) = grid.TextMatrix(grid.Row + 1, 1)
        grid.TextMatrix(grid.Row, 2) = grid.TextMatrix(grid.Row + 1, 2)
        grid.TextMatrix(grid.Row, 3) = grid.TextMatrix(grid.Row + 1, 3)
        grid.TextMatrix(grid.Row, 4) = grid.TextMatrix(grid.Row + 1, 4)
        grid.TextMatrix(grid.Row, 5) = grid.TextMatrix(grid.Row + 1, 5)
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = grid.Rows - 1
    grid.Col = 0
    Set grid.CellPicture = blank
    lblitem = "    Nama Barang : "
    lblsat = "    Nama Satuan : "
End Sub

Private Sub caricustomer()
    If txtkodecust = "" Then Exit Sub

    OBJ.Open dsn
    SQL = "select * from am_customer where kodecust = '" & txtkodecust & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblnamacust = RST!namacust
        lblalamatcust = RST!alamatcust
        
        SQL = "select top 1 kodesales from am_sohdr where kodecust = '" & txtkodecust & "' order by tglso desc"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then txtsales = RST!kodesales
        
        SQL = "SELECT * FROM AM_Salesman WHERE KodeSales = '" & txtsales & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            lblsales = RST!namasales
'-------------------- 0 = sales non aktif -------------------
            If RST!idupdate = "0" Then
                MsgBox "Salesman " & lblsales & " is not active !", vbExclamation, "Warning"
                lblsales = ""
                txtsales = ""
                txtsales.SetFocus
            End If
        Else
        End If
'------------------------------------------------------------
    Else
        MsgBox "Customer " & txtkodecust & " Not found.", vbExclamation, "Warning"
        txtkodecust = ""
        lblnamacust = ""
        lblalamatcust = ""
        txtkodecust.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub carisales()
    If txtsales = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "SELECT * FROM AM_Salesman WHERE KodeSales = '" & txtsales & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblsales = RST!namasales
    Else
        MsgBox "Salesman " & txtsales & " Not found.", vbExclamation, "Warning"
        txtsales = ""
        txtsales.SetFocus
    End If
'-------------------- 0 = sales non aktif -------------------
    If RST!idupdate = "0" Then
        MsgBox "Salesman " & lblsales & " is not active !", vbExclamation, "Warning"
        lblsales = ""
        txtsales = ""
        txtsales.SetFocus
    End If
'------------------------------------------------------------
    OBJ.Close
End Sub

Function tanggalinv()
    tanggalinv = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function
