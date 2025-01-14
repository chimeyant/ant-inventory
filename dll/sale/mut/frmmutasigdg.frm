VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmmutasigdg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mutasi Gudang"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9600
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   9600
   StartUpPosition =   1  'CenterOwner
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   7560
      TabIndex        =   15
      Top             =   2400
      Visible         =   0   'False
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   450
      Calculator      =   "frmmutasigdg.frx":0000
      Caption         =   "frmmutasigdg.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmmutasigdg.frx":008C
      Keys            =   "frmmutasigdg.frx":00AA
      Spin            =   "frmmutasigdg.frx":00EC
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
   Begin TDBText6Ctl.TDBText txtket 
      Height          =   255
      Left            =   7680
      TabIndex        =   19
      Top             =   2040
      Visible         =   0   'False
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   450
      Caption         =   "frmmutasigdg.frx":0114
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmmutasigdg.frx":0180
      Key             =   "frmmutasigdg.frx":019E
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
      Left            =   8640
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   18
      Top             =   2760
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
      Left            =   9120
      Picture         =   "frmmutasigdg.frx":01DA
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   17
      Top             =   2760
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
      Left            =   8880
      Picture         =   "frmmutasigdg.frx":04BC
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   16
      Top             =   2760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtdesc 
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
      MaxLength       =   100
      TabIndex        =   12
      Top             =   1560
      Width           =   5775
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
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtgudang 
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
      MaxLength       =   10
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtgudang2 
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
      MaxLength       =   10
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   8640
      TabIndex        =   0
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
      MICON           =   "frmmutasigdg.frx":080A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4260
      _Version        =   393216
      Cols            =   10
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
      _Band(0).Cols   =   10
   End
   Begin Chameleon.chameleonButton cmdsearch1 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "dari Gudang"
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
      MICON           =   "frmmutasigdg.frx":0B24
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
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "ke Gudang"
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
      MICON           =   "frmmutasigdg.frx":0E3E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1680
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
      Format          =   134742019
      CurrentDate     =   37426
   End
   Begin Chameleon.chameleonButton cmdsearch3 
      Height          =   285
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "No Bukti"
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
      MICON           =   "frmmutasigdg.frx":1158
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdclear 
      Height          =   375
      Left            =   7680
      TabIndex        =   14
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
      MICON           =   "frmmutasigdg.frx":1472
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsave 
      Height          =   375
      Left            =   6720
      TabIndex        =   21
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
      MICON           =   "frmmutasigdg.frx":178C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lbltotal 
      Alignment       =   1  'Right Justify
      Caption         =   "    Total Barang : 0"
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
      Left            =   7080
      TabIndex        =   20
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label8 
      Caption         =   "Desc/Reference"
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
      Left            =   240
      TabIndex        =   13
      Top             =   1590
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Tanggal Bukti"
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
      TabIndex        =   11
      Top             =   510
      Width           =   1095
   End
   Begin VB.Label lblgudang 
      BackColor       =   &H80000014&
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
      Left            =   3480
      TabIndex        =   7
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label lblgudang2 
      BackColor       =   &H80000014&
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
      Left            =   3480
      TabIndex        =   6
      Top             =   1215
      Width           =   2295
   End
End
Attribute VB_Name = "frmmutasigdg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim posrow, poscol As String
Dim str99 As String

Private Sub cmdclear_Click()
    Dim strformat As String
    strformat = Format(Date, "yymm")
    
    txtgudang = ""
    txtgudang2 = ""
    lblgudang = ""
    lblgudang2 = ""
    txtdesc = ""
    Date1.Value = Date
    hapusgrid
    
    OBJ.Open dsn
    SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'PG0-' + '" + strformat + "%' order by nobpb desc"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        str99 = Right(RST!nobpb, 4)
    Else
        str99 = 0
    End If
    OBJ.Close

    str99 = str99 + 1
    
    If Len(str99) = 1 Then txtnobukti = "PG0-" & strformat & "000" & str99
    If Len(str99) = 2 Then txtnobukti = "PG0-" & strformat & "00" & str99
    If Len(str99) = 3 Then txtnobukti = "PG0-" & strformat & "0" & str99
    If Len(str99) = 4 Then txtnobukti = "PG0-" & strformat & str99
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim strformat As String
    strformat = Format(Date, "yymm")
    
    If Len(Trim(txtnobukti)) = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        txtnobukti.SetFocus
        Exit Sub
    End If

    If txtnobukti = "" Or txtgudang = "" Or txtgudang2 = "" Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If txtgudang2 = txtgudang Then
        MsgBox "Data Entry Not Complete" & vbCrLf & "Form/To Gudang cannot be the same", vbExclamation, "Warning"
        Exit Sub
    End If
        
    If grid.Rows = 2 Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
       
    grid.Row = 1
    Do While True
        If grid.Rows = grid.Row + 1 Then Exit Do
        If grid.TextMatrix(grid.Row, 3) = "0.00" Or grid.TextMatrix(grid.Row, 4) = "" Then
            MsgBox "Data Entry Not Complete, On Row " & grid.Row, vbExclamation, "Warning"
            Exit Sub
        End If
        
        If grid.TextMatrix(grid.Row, 6) = "" Then
            MsgBox "There is no number Lot, On Row " & grid.Row, vbExclamation, "Warning"
            Exit Sub
        End If
        grid.Row = grid.Row + 1
    Loop

    OBJ.Open dsn
    SQL = "select top 1 nobpb from am_bpbhdr where nobpb like 'PG0-' + '" + strformat + "%' order by nobpb desc"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        str99 = Right(RST!nobpb, 4)
    Else
        str99 = 0
    End If
    OBJ.Close
            
    str99 = str99 + 1
        
    If Len(str99) = 1 Then txtnobukti = "PG0-" & strformat & "000" & str99
    If Len(str99) = 2 Then txtnobukti = "PG0-" & strformat & "00" & str99
    If Len(str99) = 3 Then txtnobukti = "PG0-" & strformat & "0" & str99
    If Len(str99) = 4 Then txtnobukti = "PG0-" & strformat & str99
    
    OBJ.Open dsn
    SQL = "insert into am_bpbhdr ("
    SQL = SQL + "type,"
    SQL = SQL + "nobpb,"
    SQL = SQL + "tglbpb,"
    SQL = SQL + "kodegudang,"
    SQL = SQL + "keterangan,"
    SQL = SQL + "noref,"
    SQL = SQL + "identry,"
    SQL = SQL + "dateentry,"
    SQL = SQL + "idupdate,"
    SQL = SQL + "dateupdate)"
    
    SQL = SQL + " values("
    SQL = SQL + "'99',"
    SQL = SQL + "'" & txtnobukti & "',"
    SQL = SQL + "convert(datetime,'" & tanggalinv & "'),"
    SQL = SQL + "'" & txtgudang & "',"
    SQL = SQL + "'" & txtdesc & "',"
    SQL = SQL + "'Mutasi Gudang',"
    SQL = SQL + "'" & kuser & "',"
    SQL = SQL + "convert(datetime,'" & tanggalsekarang & "'),"
    SQL = SQL + "'',"
    SQL = SQL + "convert(datetime,'" & tanggalsekarang & "'))"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "insert into am_bpbhdr ("
    SQL = SQL + "type,"
    SQL = SQL + "nobpb,"
    SQL = SQL + "tglbpb,"
    SQL = SQL + "kodegudang,"
    SQL = SQL + "keterangan,"
    SQL = SQL + "noref,"
    SQL = SQL + "identry,"
    SQL = SQL + "dateentry,"
    SQL = SQL + "idupdate,"
    SQL = SQL + "dateupdate)"
    
    SQL = SQL + " values("
    SQL = SQL + "'88',"
    SQL = SQL + "'" & txtnobukti & "',"
    SQL = SQL + "convert(datetime,'" & tanggalinv & "'),"
    SQL = SQL + "'" & txtgudang2 & "',"
    SQL = SQL + "'" & txtdesc & "',"
    SQL = SQL + "'Mutasi Gudang',"
    SQL = SQL + "'" & kuser & "',"
    SQL = SQL + "convert(datetime,'" & tanggalsekarang & "'),"
    SQL = SQL + "'',"
    SQL = SQL + "convert(datetime,'" & tanggalsekarang & "'))"
    Set RST = OBJ.Execute(SQL)
    
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        
        SQL = "insert into am_bpblin ("
        SQL = SQL + "type,"
        SQL = SQL + "nobpb,"
        SQL = SQL + "tglbpb,"
        SQL = SQL + "kodebarang,"
        SQL = SQL + "qty,"
        SQL = SQL + "keterangan,"
        SQL = SQL + "lineitem,"
        SQL = SQL + "kodesatuan)"
        
        SQL = SQL + " values("
        SQL = SQL + "'99',"
        SQL = SQL + "'" & txtnobukti & "',"
        SQL = SQL + "convert(datetime,'" & tanggalinv & "'),"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 1) & "',"
        SQL = SQL + "convert(money,'" & Format(grid.TextMatrix(grid.Row, 3), "general number") * -1 & "'),"
        SQL = SQL + "'Mutasi Gudang',"
        SQL = SQL + "convert(numeric,'" & grid.Row & "'),"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 4) & "')"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "insert into am_bpblin ("
        SQL = SQL + "type,"
        SQL = SQL + "nobpb,"
        SQL = SQL + "tglbpb,"
        SQL = SQL + "kodebarang,"
        SQL = SQL + "qty,"
        SQL = SQL + "keterangan,"
        SQL = SQL + "lineitem,"
        SQL = SQL + "kodesatuan)"
        
        SQL = SQL + " values("
        SQL = SQL + "'88',"
        SQL = SQL + "'" & txtnobukti & "',"
        SQL = SQL + "convert(datetime,'" & tanggalinv & "'),"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 1) & "',"
        SQL = SQL + "convert(money,'" & Format(grid.TextMatrix(grid.Row, 3), "general number") & "'),"
        SQL = SQL + "'Mutasi Gudang',"
        SQL = SQL + "convert(numeric,'" & grid.Row & "'),"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 4) & "')"
        Set RST = OBJ.Execute(SQL)
        
        grid.Row = grid.Row + 1
    Loop
    
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        'Tabel stok gudang out
        SQL = "insert into am_stokgudang ("
        SQL = SQL + "nolot,"
        SQL = SQL + "palet,"
        SQL = SQL + "tanggal,"
        SQL = SQL + "ref,"
        SQL = SQL + "keterangan,"
        SQL = SQL + "kodebarang,"
        SQL = SQL + "namabarang,"
        SQL = SQL + "kg,"
        SQL = SQL + "kgperpalet,"
        SQL = SQL + "hppperkg,"
        SQL = SQL + "qin,"
        SQL = SQL + "qout,"
        SQL = SQL + "kdsatuan,"
        SQL = SQL + "satuan,"
        SQL = SQL + "gudang,"
        SQL = SQL + "username,"
        SQL = SQL + "flag)"
        
        SQL = SQL + " values("
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 6) & "',"
        SQL = SQL + "'01' + '" & grid.TextMatrix(grid.Row, 6) & "',"
        SQL = SQL + "convert(datetime,'" & tanggalinv & "'),"
        SQL = SQL + "'" & txtnobukti & "',"
        SQL = SQL + "'Mutasi Gudang',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 1) & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 2) & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 7) & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 8) & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 9) & "',"
        SQL = SQL + "'0.00',"
        SQL = SQL + "convert(money,'" & Format(grid.TextMatrix(grid.Row, 3), "general number") & "'),"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 4) & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 5) & "',"
        SQL = SQL + "'" & txtgudang & "',"
        SQL = SQL + "'" & nmuser & "',"
        SQL = SQL + "'1')"
        Set RST = OBJ.Execute(SQL)
        
        'Tabel stok gudang in
        SQL = "insert into am_stokgudang ("
        SQL = SQL + "nolot,"
        SQL = SQL + "palet,"
        SQL = SQL + "tanggal,"
        SQL = SQL + "ref,"
        SQL = SQL + "keterangan,"
        SQL = SQL + "kodebarang,"
        SQL = SQL + "namabarang,"
        SQL = SQL + "kg,"
        SQL = SQL + "kgperpalet,"
        SQL = SQL + "hppperkg,"
        SQL = SQL + "qin,"
        SQL = SQL + "qout,"
        SQL = SQL + "kdsatuan,"
        SQL = SQL + "satuan,"
        SQL = SQL + "gudang,"
        SQL = SQL + "username,"
        SQL = SQL + "flag)"
        
        SQL = SQL + " values("
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 1) & "',"
        SQL = SQL + "'01' + '" & grid.TextMatrix(grid.Row, 6) & "',"
        SQL = SQL + "convert(datetime,'" & tanggalinv & "'),"
        SQL = SQL + "'" & txtnobukti & "',"
        SQL = SQL + "'Mutasi Gudang',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 1) & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 2) & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 7) & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 8) & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 9) & "',"
        SQL = SQL + "convert(money,'" & Format(grid.TextMatrix(grid.Row, 3), "general number") & "'),"
        SQL = SQL + "'0.00',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 4) & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 5) & "',"
        SQL = SQL + "'" & txtgudang2 & "',"
        SQL = SQL + "'" & nmuser & "',"
        SQL = SQL + "'0')"
        Set RST = OBJ.Execute(SQL)
        grid.Row = grid.Row + 1
    Loop
    OBJ.Close
    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
    frmpindahshow.txtnobkt = txtnobukti
    frmpindahshow.Show vbModal
    cmdclear_Click
End Sub

Private Sub cmdsearch1_Click()
    carisql1 = "select kodegudang, namagudang from am_gudang"
    namatabel = "Gudang"
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtgudang = hasil
    lblgudang = hasil1
    hasil = ""
    hasil1 = ""
    txtgudang2.SetFocus
End Sub

Private Sub cmdsearch2_Click()
    carisql1 = "select kodegudang, namagudang from am_gudang"
    namatabel = "Gudang"
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch2_GotFocus()
    If hasil = "" Then Exit Sub
    If hasil = "G3" Then
        MsgBox "Mutasi ke WIP saat ini belum bisa dilakukan", vbCritical, AppName
        Exit Sub
    End If
    txtgudang2 = hasil
    lblgudang2 = hasil1
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    txtdesc.SetFocus
End Sub

Private Sub Form_Load()
    cmdclear_Click
    grid.TextMatrix(0, 1) = "Kode Barang"
    grid.TextMatrix(0, 2) = "Nama Barang"
    grid.TextMatrix(0, 3) = "Quantity"
    grid.TextMatrix(0, 4) = "K/Satuan"
    grid.TextMatrix(0, 5) = "N/Satuan"
    grid.TextMatrix(0, 6) = "No Lot"
    grid.TextMatrix(0, 7) = "kg"
    grid.TextMatrix(0, 8) = "kg/palet"
    grid.TextMatrix(0, 9) = "hpp/kg"
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1200
    grid.ColWidth(2) = 2500
    grid.ColWidth(3) = 1000
    grid.ColWidth(4) = 1000
    grid.ColWidth(5) = 1000
    grid.ColWidth(6) = 1500
    grid.ColWidth(7) = 0
    grid.ColWidth(8) = 0
    grid.ColWidth(9) = 0
    grid.RowHeightMin = 300
    
End Sub

Private Sub grid_Click()
On Error Resume Next
    If grid.MouseRow = 0 Then Exit Sub
    If txtnobukti = "" Or txtgudang = "" Or txtgudang2 = "" Then Exit Sub

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
            If grid.Rows - 1 = 200 Then
                MsgBox "Maximum line 200", vbExclamation, "Warning"
                Exit Sub
            End If
            
            If txtket.Visible = True Then Exit Sub
            
            txtket.Width = grid.ColWidth(grid.Col) - 40
            txtket = grid.TextMatrix(grid.Row, grid.Col)
            txtket.Left = grid.Left + grid.CellLeft
            txtket.Top = grid.Top + grid.CellTop + 20
            txtket.Visible = True
            txtket.SetFocus
        Case 4
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            If grid.TextMatrix(grid.Row, 3) = "0.00" Or grid.TextMatrix(grid.Row, 3) = "0" Then
                MsgBox "Qty cannot be null", vbCritical, AppName
                Exit Sub
            End If
            
            If txtket.Visible = True Then Exit Sub
            
            txtket.Width = grid.ColWidth(grid.Col) - 40
            txtket = grid.TextMatrix(grid.Row, grid.Col)
            txtket.Left = grid.Left + grid.CellLeft
            txtket.Top = grid.Top + grid.CellTop + 20
            txtket.Visible = True
            txtket.SetFocus
        Case 3
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            
            If txtnilai.Visible = True Then Exit Sub
            
            txtnilai.Width = grid.ColWidth(grid.Col) - 40
            txtnilai = grid.TextMatrix(grid.Row, grid.Col)
            txtnilai.Left = grid.Left + grid.CellLeft
            txtnilai.Top = grid.Top + grid.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
        Case 6
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            If grid.TextMatrix(grid.Row, 3) = "0.00" Or grid.TextMatrix(grid.Row, 3) = "0" Then
                MsgBox "Qty cannot be null", vbCritical, AppName
                Exit Sub
            End If
            'periksa nomor lot kalau tidak ditemukan bisa input manual
            OBJ.Open dsn
            SQL = "Select distinct nolot,tanggal,kg,kgperpalet,hppperkg from am_stokgudang"
            SQL = SQL + " Where kodebarang='" & grid.TextMatrix(grid.Row, 1) & "' and flag = '0'"
            Set RST = OBJ.Execute(SQL)
            
            If RST.EOF Then
                OBJ.Close
                txtket.Width = grid.ColWidth(grid.Col) - 40
                txtket = grid.TextMatrix(grid.Row, grid.Col)
                txtket.Left = grid.Left + grid.CellLeft
                txtket.Top = grid.Top + grid.CellTop + 20
                txtket.Visible = True
                txtket.SetFocus
                Exit Sub
            End If
            OBJ.Close
            
            hasil8 = grid.TextMatrix(grid.Row, 1)
            carisql1 = "Select distinct nolot,tanggal,kg,kgperpalet,hppperkg from am_stokgudang"
            namatabel = "Daftar Lot"
            frmsearch.Show vbModal
    End Select
End Sub

Private Sub grid_EnterCell()
    If txtnobukti = "" Or txtgudang = "" Or txtgudang2 = "" Then Exit Sub
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
    Case 4
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
        If txtket.Visible = True Then Exit Sub
        If grid.TextMatrix(grid.Row, 3) = "0.00" Or grid.TextMatrix(grid.Row, 3) = "0" Then
            MsgBox "Qty cannot be null", vbCritical, AppName
            Exit Sub
        End If
            
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
        
        If txtnilai.Visible = True Then Exit Sub
            
        posrow = grid.Row
        poscol = grid.Col
        txtnilai.Width = grid.ColWidth(grid.Col) - 40
        txtnilai = grid.TextMatrix(grid.Row, grid.Col)
        txtnilai.Left = grid.Left + grid.CellLeft
        txtnilai.Top = grid.Top + grid.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
    Case 6
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
        If grid.TextMatrix(grid.Row, 3) = "0.00" Or grid.TextMatrix(grid.Row, 3) = "0" Then
            MsgBox "Qty cannot be null", vbCritical, AppName
            Exit Sub
        End If
        posrow = grid.Row
        poscol = grid.Col
        hasil8 = grid.TextMatrix(grid.Row, 1)
        carisql1 = "Select distinct nolot,tanggal,kg,kgperpalet,hppperkg from am_stokgudang"
        namatabel = "Daftar Lot"
        frmsearch.Show vbModal
    End Select
End Sub

Private Sub grid_GotFocus()
    If hasil = "" Then Exit Sub
    
    If grid.Col = 6 Then
        'periksa lot pada item yang sama sudah ada atau belum
        grid.Row = 1
        Do While True
            If grid.Rows = 2 Or grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
            If grid.TextMatrix(grid.Row, 6) = hasil And grid.TextMatrix(grid.Row, 1) = grid.TextMatrix(posrow, 1) And posrow <> grid.Row Then
            
                MsgBox "Nomor Lot is already exist.", vbInformation, "Information"
                hasil = ""
                grid.Row = posrow
                grid.Col = 6
                grid.SetFocus
                Exit Sub
            End If
            grid.Row = grid.Row + 1
        Loop
        If hasil = "" Then Exit Sub
        
        grid.Row = posrow
        grid.Col = poscol
        grid.TextMatrix(grid.Row, 6) = hasil
        grid.TextMatrix(grid.Row, 7) = hasil2
        grid.TextMatrix(grid.Row, 8) = hasil3
        grid.TextMatrix(grid.Row, 9) = hasil4
        hasil = "": hasil1 = "": hasil2 = "": hasil3 = "": hasil4 = "": hasil8 = ""
        Exit Sub
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
            grid.TextMatrix(grid.Row, 2) = RST!NamaBarang
            grid.TextMatrix(grid.Row, 3) = "0.00"
            
            SetRow grid.Row, True
            lbltotal.Caption = "    Total Barang : " & grid.Rows - 1
            grid.SetFocus
            grid.Col = 2
            If grid.Row = (grid.Rows - 1) Then grid.Rows = grid.Rows + 1
        Else
            MsgBox "Item Not Found", vbExclamation, "Warning"
            grid.TextMatrix(grid.Row, 1) = ""
        End If
        OBJ.Close
    End If

    If grid.Col = 4 Then
        OBJ.Open dsn
        SQL = "SELECT a.namabarang,b.kodesatuan,b.namasatuan FROM AM_ITEMDTL a left join am_unit b ON a.kodesatuan=b.kodesatuan WHERE a.KodeBarang = '" & grid.TextMatrix(grid.Row, 1) & "' and a.kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            grid.TextMatrix(grid.Row, 2) = RST!NamaBarang
            grid.TextMatrix(grid.Row, 5) = RST!namasatuan
            grid.SetFocus
            grid.Col = 5
            
        Else
            grid.TextMatrix(grid.Row, 5) = ""
            grid.TextMatrix(grid.Row, 4) = ""
        End If
        If OBJ.State = 1 Then OBJ.Close
    End If
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
        grid.TextMatrix(grid.Row, 7) = ""
        grid.TextMatrix(grid.Row, 8) = ""
        grid.TextMatrix(grid.Row, 9) = ""
        grid.Col = 0
        Set grid.CellPicture = blank
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1200
    grid.ColWidth(2) = 2500
    grid.ColWidth(3) = 1000
    grid.ColWidth(4) = 1000
    grid.ColWidth(5) = 1000
    grid.ColWidth(6) = 1500
    grid.ColWidth(7) = 0
    grid.ColWidth(8) = 0
    grid.ColWidth(9) = 0
End Sub

Private Sub hapusrow()
    grid.TextMatrix(grid.Row, 1) = ""
    grid.TextMatrix(grid.Row, 2) = ""
    grid.TextMatrix(grid.Row, 3) = ""
    grid.TextMatrix(grid.Row, 4) = ""
    grid.TextMatrix(grid.Row, 5) = ""
    grid.TextMatrix(grid.Row, 6) = ""
    grid.TextMatrix(grid.Row, 7) = ""
    grid.TextMatrix(grid.Row, 8) = ""
    grid.TextMatrix(grid.Row, 9) = ""
    
    Do While True
        If grid.TextMatrix(grid.Row + 1, 1) = "" Then
            grid.TextMatrix(grid.Row, 1) = ""
            grid.TextMatrix(grid.Row, 2) = ""
            grid.TextMatrix(grid.Row, 3) = ""
            grid.TextMatrix(grid.Row, 4) = ""
            grid.TextMatrix(grid.Row, 5) = ""
            grid.TextMatrix(grid.Row, 6) = ""
            grid.TextMatrix(grid.Row, 7) = ""
            grid.TextMatrix(grid.Row, 8) = ""
            grid.TextMatrix(grid.Row, 9) = ""
            Exit Do
        End If
        grid.TextMatrix(grid.Row, 1) = grid.TextMatrix(grid.Row + 1, 1)
        grid.TextMatrix(grid.Row, 2) = grid.TextMatrix(grid.Row + 1, 2)
        grid.TextMatrix(grid.Row, 3) = grid.TextMatrix(grid.Row + 1, 3)
        grid.TextMatrix(grid.Row, 4) = grid.TextMatrix(grid.Row + 1, 4)
        grid.TextMatrix(grid.Row, 5) = grid.TextMatrix(grid.Row + 1, 5)
        grid.TextMatrix(grid.Row, 6) = grid.TextMatrix(grid.Row + 1, 6)
        grid.TextMatrix(grid.Row, 7) = grid.TextMatrix(grid.Row + 1, 7)
        grid.TextMatrix(grid.Row, 8) = grid.TextMatrix(grid.Row + 1, 8)
        grid.TextMatrix(grid.Row, 9) = grid.TextMatrix(grid.Row + 1, 9)
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = grid.Rows - 1
    grid.Col = 0
    Set grid.CellPicture = blank
    
    If grid.Rows = 2 Then
        lbltotal.Caption = "    Total Barang : 0"
    Else
        lbltotal.Caption = "    Total Barang : " & grid.Rows - 2
    End If
End Sub
Private Sub SetRow(ByVal idx As Integer, ByVal hapus As String)
    grid.Row = idx
    grid.Col = 0
    If hapus Then
        Set grid.CellPicture = uncheck.Picture
    End If
    grid.Col = 1
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
                    grid.TextMatrix(grid.Row, 2) = RST!NamaBarang
                    grid.TextMatrix(grid.Row, 3) = "0.00"
                    grid.Col = 0
                    Set grid.CellPicture = uncheck.Picture
                    OBJ.Close
                    If grid.Row = (grid.Rows - 1) Then grid.Rows = grid.Rows + 1
                Else
                    OBJ.Close
                    grid.TextMatrix(posrow, 1) = ""
                    txtket = ""
                    
                    If txtgudang = "G3" Then
                        carisql1 = "Select distinct a.kodebarang,b.namabarang from am_bpblin a inner join am_itemmst b "
                        carisql1 = carisql1 + "on a.kodebarang = b.KodeBarang inner join am_bpbhdr c on a.nobpb =c.nobpb "
                        carisql1 = carisql1 + "Where c.kodegudang = 'G3'" ' and c.noref like '" + lblnolot + "%'"
                        namatabel = "Item Gudang WIF"
                    Else
                
                    carisql1 = "select kodebarang, namabarang from am_itemmst"
                    namatabel = "Item"
                    End If
                    frmsearch.Show vbModal
                End If
                grid.Col = 1
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
                SQL = "SELECT namabarang,kodesatuan FROM AM_ITEMDTL WHERE KodeBarang = '" & grid.TextMatrix(grid.Row, 1) & "' and kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
                Set RST = OBJ.Execute(SQL)
                If RST.EOF Then
                    grid.TextMatrix(posrow, 4) = ""
                    txtket = ""
                    carisql1 = "SELECT b.kodesatuan,b.namasatuan FROM AM_ITEMDTL a left join am_unit b on a.kodesatuan = b.kodesatuan WHERE a.KodeBarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
                    namatabel = "Satuan "
                    frmsearch.Show vbModal
                Else
                    grid.TextMatrix(grid.Row, 2) = RST!NamaBarang
                    SQL = "SELECT namasatuan FROM AM_unit WHERE Kodesatuan = '" & grid.TextMatrix(grid.Row, 4) & "'"
                    Set RST = OBJ.Execute(SQL)
                    If Not RST.EOF Then grid.TextMatrix(grid.Row, 2) = RST!namasatuan
                    
                End If
                OBJ.Close
            Case 6
                grid.TextMatrix(grid.Row, 6) = txtket.text
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

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        txtnilai_LostFocus
        Exit Sub
    End If
    
    If KeyAscii = 13 Then
        'OBJ.Open dsn
        'SQL = "Select SUM(a.qty)as qty From am_bpblin a inner join am_bpbhdr b "
        'SQL = SQL + "on a.nobpb = b.nobpb and (a.type = '01' or a.type = '99') "
        'SQL = SQL + "Where a.kodebarang = '" + grid.TextMatrix(grid.Row, 1) + "' and b.kodegudang = 'G3' "
        'SQL = SQL + "and b.noref like '" + lblnolot + "%'"
        'Set RST = OBJ.Execute(SQL)
        
        grid.TextMatrix(grid.Row, grid.Col) = Format(txtnilai, "###,###,##0.00")
        'OBJ.Close
        txtnilai = 0
        txtnilai.Visible = False
        grid.SetFocus
        grid.Row = posrow
        grid.Col = poscol
    End If
End Sub

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
End Sub

Function tanggalinv()
    tanggalinv = Month(Date1) & "/" & Day(Date1) & "/" & Year(Date1)
End Function

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function
