VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmpenerimaan_appretur 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Confirm Retur Penerimaan Barang"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10095
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
   Icon            =   "frmpenerimaan_appretur.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtref2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7920
      MaxLength       =   20
      TabIndex        =   4
      Top             =   960
      Width           =   1815
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai1 
      Height          =   255
      Left            =   8880
      TabIndex        =   7
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   450
      Calculator      =   "frmpenerimaan_appretur.frx":2372
      Caption         =   "frmpenerimaan_appretur.frx":2392
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpenerimaan_appretur.frx":23FE
      Keys            =   "frmpenerimaan_appretur.frx":241C
      Spin            =   "frmpenerimaan_appretur.frx":245E
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
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   7920
      TabIndex        =   17
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
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
      Format          =   122355715
      CurrentDate     =   37426
   End
   Begin VB.TextBox txtnoretur 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      MaxLength       =   17
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox txtnobeli 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      MaxLength       =   20
      TabIndex        =   3
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtcurr1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      MaxLength       =   20
      TabIndex        =   5
      Top             =   1320
      Width           =   615
   End
   Begin VB.ListBox List2 
      Height          =   3570
      ItemData        =   "frmpenerimaan_appretur.frx":2486
      Left            =   120
      List            =   "frmpenerimaan_appretur.frx":2488
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   2535
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
      Left            =   7800
      Picture         =   "frmpenerimaan_appretur.frx":248A
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   14
      Top             =   120
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
      Left            =   8040
      Picture         =   "frmpenerimaan_appretur.frx":27D8
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   13
      Top             =   120
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
      Left            =   7560
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   9120
      TabIndex        =   11
      Top             =   3720
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
      MICON           =   "frmpenerimaan_appretur.frx":2ABA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid1 
      Height          =   1935
      Left            =   2760
      TabIndex        =   8
      Top             =   1680
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   3413
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
   Begin Chameleon.chameleonButton cmdpost1 
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      ToolTipText     =   "Submit"
      Top             =   600
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "4"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
         Weight          =   500
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
      MICON           =   "frmpenerimaan_appretur.frx":2DD4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilaikurs1 
      Height          =   285
      Left            =   5160
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   503
      Calculator      =   "frmpenerimaan_appretur.frx":30EE
      Caption         =   "frmpenerimaan_appretur.frx":310E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpenerimaan_appretur.frx":317A
      Keys            =   "frmpenerimaan_appretur.frx":3198
      Spin            =   "frmpenerimaan_appretur.frx":31DA
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,##0.00;;0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,##0.00"
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
   Begin Chameleon.chameleonButton cmdclear1 
      Height          =   375
      Left            =   8160
      TabIndex        =   10
      Top             =   3720
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
      MICON           =   "frmpenerimaan_appretur.frx":3202
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdadd1 
      Height          =   375
      Left            =   7200
      TabIndex        =   9
      Top             =   3720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Confirm"
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
      MICON           =   "frmpenerimaan_appretur.frx":351C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtppn 
      Height          =   285
      Left            =   7080
      TabIndex        =   26
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   503
      Calculator      =   "frmpenerimaan_appretur.frx":3836
      Caption         =   "frmpenerimaan_appretur.frx":3856
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpenerimaan_appretur.frx":38C2
      Keys            =   "frmpenerimaan_appretur.frx":38E0
      Spin            =   "frmpenerimaan_appretur.frx":3922
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,##0.00;;0"
      EditMode        =   1
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,##0.00"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   99
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   1999699973
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "No Invoice"
      Height          =   255
      Left            =   6480
      TabIndex        =   25
      Top             =   990
      Width           =   1335
   End
   Begin VB.Label Label15 
      Height          =   255
      Left            =   7920
      TabIndex        =   24
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lblket1a 
      Height          =   255
      Left            =   2760
      TabIndex        =   23
      Top             =   3630
      Width           =   4335
   End
   Begin VB.Label lblket2a 
      Height          =   255
      Left            =   2760
      TabIndex        =   22
      Top             =   3870
      Width           =   4335
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Tanggal Retur"
      Height          =   255
      Left            =   6720
      TabIndex        =   21
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "No LPB"
      Height          =   255
      Left            =   3480
      TabIndex        =   20
      Top             =   990
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "No Retur"
      Height          =   255
      Left            =   3480
      TabIndex        =   19
      Top             =   630
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Currency"
      Height          =   255
      Left            =   3480
      TabIndex        =   18
      Top             =   1350
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Confirm"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   16
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Unconfirm"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmpenerimaan_appretur"
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

Dim str2, str3 As String
Dim posrow, posrow1 As String
Dim i, j As Integer

Private Sub cmdadd1_Click()
    If Len(Trim(txtnoretur)) = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        txtnoretur.SetFocus
        Exit Sub
    End If
    
    If txtnoretur = "" Or txtnobeli = "" Or txtcurr1 = "" Or txtnilaikurs1 = 0 Or txtref2 = "" Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If Grid1.Rows = 2 Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
        
    If MsgBox("Are you sure want to confirm (Retur)?", vbQuestion + vbYesNo, "Question") = vbNo Then
        cmdclear1_Click
        Exit Sub
    End If
        
    OBJ1.Open dsn
    SQL1 = "select * from am_apopnfil where nobeli = '" & txtnoretur & "'"
    Set RST1 = OBJ1.Execute(SQL1)
    If Not RST1.EOF Then
        OBJ1.Close
        MsgBox "Can not confirm, please check Retur.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    Grid1.Row = 1
    Do While True
        If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Do

        SQL1 = "select * from am_beliapp where nobeli = '" & txtnobeli & "' and flag2 >= '1' and kodebarang = '" & Grid1.TextMatrix(Grid1.Row, 1) & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If RST1.EOF Then
            MsgBox "Can not confirm, LPB not found / LPB not confirm." & vbCrLf & _
            "Please make sure LPB exist and confirm.", vbExclamation, "Warning"
            OBJ1.Close
            Exit Sub
        End If
        
        Grid1.Row = Grid1.Row + 1
    Loop
    
    SQL1 = "delete from am_apopnfil where nobeli = '" & txtnoretur & "'"
    Set RST1 = OBJ1.Execute(SQL1)
    OBJ1.Close
    
    OBJ1.Open dsn
    Grid1.Row = 1
    Do While True
        If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Do

        SQL1 = "update AM_beliretur set "
        SQL1 = SQL1 + "nilaikurs = Convert(Money, '" & txtnilaikurs1 & "'), "
        SQL1 = SQL1 + "nopo = '" & txtref2 & "', "
        SQL1 = SQL1 + "flag2 = '1' where noretur = '" & txtnoretur & "' and kodebarang = '" & Grid1.TextMatrix(Grid1.Row, 1) & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        
        Grid1.Row = Grid1.Row + 1
    Loop
    OBJ1.Close
    
    OBJ1.Open dsn
    SQL1 = "select b.noretur,b.kodesupp,b.kodecur,b.nilaikurs,sum((b.qty/(a.qtyuse/a.qty))*b.price)'amount'"
    SQL1 = SQL1 + " from am_beliretur b left join am_belireturtemp a on a.noretur=b.noretur and a.kodebarang=b.kodebarang"
    SQL1 = SQL1 + " where b.noretur = '" & txtnoretur & "' group by b.noretur,b.kodesupp,b.kodecur,b.nilaikurs"
    Set RST1 = OBJ1.Execute(SQL1)
    If Not RST1.EOF Then
        OBJ.Open dsn
        SQL = "insert into am_apopnfil ("
        SQL = SQL + "kodesupp, "
        SQL = SQL + "nobeli, "
        SQL = SQL + "tglbeli, "
        SQL = SQL + "noapply, "
        SQL = SQL + "transtype, "
        SQL = SQL + "keterangan, "
        SQL = SQL + "amount, "
        SQL = SQL + "potongan, "
        SQL = SQL + "ppn, "
        SQL = SQL + "selisih, "
        SQL = SQL + "kodecur, "
        SQL = SQL + "nilaikurs)"
        
        SQL = SQL + " values ('" & RST1!kodesupp & "',"
        SQL = SQL + "'" & txtnoretur & "',"
        SQL = SQL + "convert(datetime,'" & tanggal2 & "'),"
        SQL = SQL + "'" & txtref2 & "',"
        SQL = SQL + "'CI',"
        SQL = SQL + "'No Retur " & txtnoretur & "',"
        SQL = SQL + "convert(money,'" & RST1!amount * -1 & "'),"
        SQL = SQL + "convert(money,'0'),"
        If txtppn = 0 Then SQL = SQL + "convert(money,'0')," Else SQL = SQL + "convert(money,'" & RST1!amount * 0.1 * -1 & "'),"
        SQL = SQL + "convert(money,'0'),"
        SQL = SQL + "'" & RST1!kodecur & "',"
        SQL = SQL + "convert(money,'" & RST1!nilaikurs & "'))"
        Set RST = OBJ.Execute(SQL)
        OBJ.Close
    End If
    OBJ1.Close
    
    List2.RemoveItem (j)
    
    MsgBox "Data confirm, Click OK To Continue ...", vbInformation, "Information"
    cmdclear1_Click
End Sub

Private Sub cmdclear1_Click()
    hapusgrid1
    
    txtnoretur = ""
    date2.Value = Date
    txtnobeli = ""
    txtcurr1 = ""
    txtref2 = ""
    txtnilaikurs1 = 0
    txtppn = 0
    Label15 = ""
    
    lblket1a = "Nama Barang : "
    lblket2a = "Nama Satuan : "
    
    txtnoretur.SetFocus
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdpost1_Click()
    If List2.text = "" Then Exit Sub

    hapusgrid1
    date2 = Date
    txtnobeli = ""
    txtcurr1 = ""
    txtref2 = ""
    Label15 = ""
    txtppn = 0
    txtnilaikurs1 = 0
    txtnoretur = List2.text
    j = List2.ListIndex
    
    OBJ.Open dsn
    SQL = "select nobeli from am_beliretur where noretur = '" & txtnoretur & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ1.Open dsn
        SQL1 = "select * from am_beliapp where nobeli = '" & RST!nobeli & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If RST1.EOF Then
            MsgBox "Can not confirm, LPB not found / LPB not confirm." & vbCrLf & _
            "Please make sure LPB exist and confirm.", vbExclamation, "Warning"
            OBJ1.Close
            OBJ.Close
            txtnoretur = ""
            Exit Sub
        End If
        OBJ1.Close
    End If
    OBJ.Close

    OBJ.Open dsn
    SQL = "select distinct a.tglretur,a.nobeli,a.kodecur,a.kodesupp,b.nilaikurs,b.ref2,b.ppn from am_beliretur a left join am_beliapp b on a.nobeli=b.nobeli where a.noretur = '" & txtnoretur & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        date2 = RST!tglretur
        Label15 = Format(RST!tglretur, "dd MMMM yyyy")
        txtnobeli = RST!nobeli
        str3 = RST!kodesupp
        txtcurr1 = RST!kodecur
        txtnilaikurs1 = RST!nilaikurs
        txtref2 = RST!ref2
        txtppn = RST!ppn
        
        Grid1.Row = 1
        SQL = "select a.kodebarang,a.kodesatuan,a.qty,b.price from am_beliretur a left join am_beliapp b on a.nobeli=b.nobeli and a.kodebarang=b.kodebarang where a.noretur = '" & txtnoretur & "' order by a.lineitem asc"
        Set RST = OBJ.Execute(SQL)
        Do Until RST.EOF
            Grid1.Col = 1
            Grid1.CellAlignment = 1
            Grid1.TextMatrix(Grid1.Row, 1) = RST!kodebarang
            Grid1.Col = 2
            Grid1.CellAlignment = 1
            Grid1.TextMatrix(Grid1.Row, 2) = RST!kodesatuan
            Grid1.TextMatrix(Grid1.Row, 3) = Format(RST!qty, "###,###,##0.00")
            Grid1.TextMatrix(Grid1.Row, 4) = Format(RST!Price, "###,###,##0.00")
            
            Grid1.Col = 0
            Set Grid1.CellPicture = uncheck.Picture

            Grid1.Rows = Grid1.Rows + 1
            Grid1.Row = Grid1.Row + 1
            RST.MoveNext
        Loop
    Else
        MsgBox "Data not found.", vbExclamation, "Warning"
        txtnoretur = ""
        txtnoretur.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='104' and b.kodeuser = '2" & kuser & "'"
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
   
    Grid1.TextMatrix(0, 1) = "Kode Barang"
    Grid1.TextMatrix(0, 2) = "K/Sat."
    Grid1.TextMatrix(0, 3) = "Retur"
    Grid1.TextMatrix(0, 4) = "Price"
    Grid1.ColWidth(0) = 250
    Grid1.ColWidth(1) = 1500
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 1500
    Grid1.ColWidth(5) = 0
    
    Grid1.RowHeightMin = 300
    
    date2.Value = Date
    
    List2.Clear
    
    OBJ.Open dsn
    SQL = "SELECT distinct noretur FROM AM_beliretur WHERE flag2 = '0'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        List2.AddItem RST!noretur
    
        RST.MoveNext
    Loop
    OBJ.Close
    
    lblket1a = "Nama Barang : "
    lblket2a = "Nama Satuan : "
End Sub

Private Sub grid1_Click()
    If Grid1.MouseRow = 0 Then Exit Sub
    If txtnoretur = "" Or txtnobeli = "" Then Exit Sub
    
    OBJ1.Open dsn
    SQL1 = "SELECT * FROM am_apitemmst WHERE KodeBarang = '" & Grid1.TextMatrix(Grid1.Row, 1) & "'"
    Set RST1 = OBJ1.Execute(SQL1)
    If Not RST1.EOF Then lblket1a = "Nama Barang : " & RST1!namabarang Else lblket1a = "Nama Barang : "

    SQL1 = "SELECT * FROM am_apunit WHERE Kodesatuan = '" & Grid1.TextMatrix(Grid1.Row, 2) & "'"
    Set RST1 = OBJ1.Execute(SQL1)
    If Not RST1.EOF Then lblket2a = "Nama Satuan : " & RST1!namasatuan Else lblket2a = "Nama Satuan : "
    OBJ1.Close
End Sub

Private Sub txtcurr1_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtcurr1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtnilaikurs1.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtnilaikurs1_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtnilaikurs1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Grid1.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtnobeli_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtnobeli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtref2.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtnoretur_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtnoretur_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtnobeli.SetFocus
    KeyAscii = 0
End Sub

Private Sub hapusgrid1()
    Grid1.Row = 1
    Do While True
        If Grid1.TextMatrix(Grid1.Row, 1) = "" Then Exit Do
        Grid1.TextMatrix(Grid1.Row, 1) = ""
        Grid1.TextMatrix(Grid1.Row, 2) = ""
        Grid1.TextMatrix(Grid1.Row, 3) = ""
        Grid1.TextMatrix(Grid1.Row, 4) = ""
        Grid1.TextMatrix(Grid1.Row, 5) = ""
        Grid1.Col = 0
        Set Grid1.CellPicture = blank
        Grid1.Row = Grid1.Row + 1
    Loop
    Grid1.Rows = 2
    Grid1.ColWidth(0) = 250
    Grid1.ColWidth(1) = 1500
    Grid1.ColWidth(2) = 1000
    Grid1.ColWidth(3) = 1000
    Grid1.ColWidth(4) = 1500
    Grid1.ColWidth(5) = 0
End Sub

Function tanggal2()
      tanggal2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function

Private Sub txtref2_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtref2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtcurr1.SetFocus
    KeyAscii = 0
End Sub
