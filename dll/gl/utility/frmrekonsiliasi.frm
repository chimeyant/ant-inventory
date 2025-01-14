VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Begin VB.Form frmrekonsiliasi 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   15315
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   Icon            =   "frmrekonsiliasi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   15315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtacc1 
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
      Left            =   1200
      MaxLength       =   15
      TabIndex        =   21
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtacc2 
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
      MaxLength       =   15
      TabIndex        =   20
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mark"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   14280
      TabIndex        =   15
      Top             =   705
      Width           =   975
      Begin VB.CheckBox cbK 
         BackColor       =   &H00808080&
         Caption         =   "K"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   615
         Width           =   495
      End
      Begin VB.CheckBox cbD 
         BackColor       =   &H00C0C0C0&
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         FillColor       =   &H00808080&
         Height          =   255
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   735
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         FillColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   210
         Width           =   735
      End
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   8520
      TabIndex        =   13
      Top             =   840
      Visible         =   0   'False
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   450
      Calculator      =   "frmrekonsiliasi.frx":000C
      Caption         =   "frmrekonsiliasi.frx":002C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmrekonsiliasi.frx":0098
      Keys            =   "frmrekonsiliasi.frx":00B6
      Spin            =   "frmrekonsiliasi.frx":00F8
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   12648447
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,###,##0.00;(###,###,###,##0.00)"
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
      Left            =   10440
      TabIndex        =   14
      Top             =   840
      Visible         =   0   'False
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
      _ExtentY        =   450
      Caption         =   "frmrekonsiliasi.frx":0120
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmrekonsiliasi.frx":018C
      Key             =   "frmrekonsiliasi.frx":01AA
      BackColor       =   12648447
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
      MaxLength       =   60
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
   Begin XtremeSuiteControls.ProgressBar Pg 
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   7695
      _Version        =   851970
      _ExtentX        =   13573
      _ExtentY        =   661
      _StockProps     =   93
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      UseVisualStyle  =   0   'False
      TextAlignment   =   2
   End
   Begin VB.PictureBox blank 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7440
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox check 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8160
      Picture         =   "frmrekonsiliasi.frx":01E6
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox uncheck 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   7800
      Picture         =   "frmrekonsiliasi.frx":04C8
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2280
      Top             =   6720
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   14737632
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   8705
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      BackColorFixed  =   -2147483637
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
      _Band(0).Cols   =   11
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1200
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
      Format          =   106823683
      CurrentDate     =   37749
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   4200
      TabIndex        =   3
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
      Format          =   106823683
      CurrentDate     =   37749
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   14280
      TabIndex        =   10
      Top             =   6840
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
      MICON           =   "frmrekonsiliasi.frx":0816
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton btnview 
      Height          =   375
      Left            =   6360
      TabIndex        =   11
      Top             =   1200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "View"
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
      MICON           =   "frmrekonsiliasi.frx":0B30
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdupdate 
      Height          =   375
      Left            =   13320
      TabIndex        =   19
      Top             =   6840
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Update"
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
      MICON           =   "frmrekonsiliasi.frx":0E4A
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
      Left            =   0
      TabIndex        =   22
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "From Account"
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
      MICON           =   "frmrekonsiliasi.frx":1164
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
      Left            =   3120
      TabIndex        =   23
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "To Account"
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
      MICON           =   "frmrekonsiliasi.frx":147E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblwait 
      Caption         =   "Please wait..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   7440
      TabIndex        =   18
      Top             =   1320
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblrecord 
      Appearance      =   0  'Flat
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   "To Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "From Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   870
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "  Rekonsiliasi"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   615
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   15375
   End
End
Attribute VB_Name = "frmrekonsiliasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String


Private Sub btnview_Click()
    If txtacc1 = "" And txtacc2 = "" Then Exit Sub
    If btnview.Caption = "View" Then
        If date2 < date1 Then
            MsgBox "To Date Can Not Smaller Then From Date.", vbExclamation, "Warning"
            Exit Sub
        End If
        lblwait.Visible = True
        Screen.MousePointer = vbHourglass
        DoEvents
        btnview.Caption = "Clear"
        opendata
        Screen.MousePointer = vbDefault
        If grid.Rows = 1 Then
            grid.Rows = 2
            Adodc1.Refresh
            grid.Refresh
        End If
        lblwait.Visible = False
    ElseIf btnview.Caption = "Clear" Then
        Screen.MousePointer = vbHourglass
        DoEvents
        hapusgrid
        btnview.Caption = "View"
        Screen.MousePointer = vbDefault
        cbD.Value = Unchecked
        cbK.Value = Unchecked
    End If
End Sub

Private Sub cbD_Click()
Dim i, j As Integer
    If cbD.Value = Checked Then
        If grid.TextMatrix(grid.Row, 1) = "" Then
            MsgBox "There is no data can be marked!.", vbExclamation, AppName
            cbD.Value = Unchecked
            Exit Sub
        End If

        grid.Row = 1
        Do While True
            grid.Col = 0
            If grid.CellPicture = check Then GoTo lewati:
            If grid.TextMatrix(grid.Row, 7) = "D" Then
                For i = 0 To grid.Cols - 1
                grid.Col = i
                grid.CellBackColor = &HE0E0E0
                Next
lewati:
            End If
            If grid.Row = grid.Rows - 1 Then Exit Do
            grid.Row = grid.Row + 1
         Loop
    Else
        grid.Row = 1
        Do While True
            grid.Col = 0
            If grid.CellPicture = check Then GoTo lewati2:
            If grid.TextMatrix(grid.Row, 7) = "D" Then
                For i = 0 To grid.Cols - 1
                grid.Col = i
                grid.CellBackColor = &HFFFFFF
                Next
lewati2:
            End If
            If grid.Row = grid.Rows - 1 Then Exit Do
            grid.Row = grid.Row + 1
         Loop
    End If
End Sub

Private Sub cbK_Click()
Dim i, j As Integer
    If cbK.Value = Checked Then
        If grid.TextMatrix(grid.Row, 1) = "" Then
            MsgBox "There is no data can be marked!.", vbExclamation, AppName
            cbK.Value = Unchecked
            Exit Sub
        End If
        
        grid.Row = 1
        Do While True
            grid.Col = 0
            If grid.CellPicture = check Then GoTo lewati:
            If grid.TextMatrix(grid.Row, 7) = "K" Then
                For i = 0 To grid.Cols - 1
                grid.Col = i
                grid.CellBackColor = &HC0C0C0
                Next
lewati:
            End If
            If grid.Row = grid.Rows - 1 Then Exit Do
            grid.Row = grid.Row + 1
         Loop
    Else
        grid.Row = 1
        Do While True
            grid.Col = 0
            If grid.CellPicture = check Then GoTo lewati2:
            If grid.TextMatrix(grid.Row, 7) = "K" Then
                For i = 0 To grid.Cols - 1
                grid.Col = i
                grid.CellBackColor = &HFFFFFF
                Next
lewati2:
            End If
            If grid.Row = grid.Rows - 1 Then Exit Do
            grid.Row = grid.Row + 1
         Loop
    End If
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdsearch3_Click()

End Sub

Private Sub cmdsearch1_Click()
    carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp >= '01' and a.kdcomp <= '01'"
    namatabel = "Company Account  "
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc1 = hasil
    hasil = ""
    hasil1 = ""
    txtacc2.SetFocus
End Sub

Private Sub cmdsearch2_Click()
    carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp >= '01' and a.kdcomp <= '01'"
    namatabel = "Company Account  "
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch2_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc2 = hasil
    hasil = ""
    hasil1 = ""
    btnview.SetFocus
End Sub

Private Sub cmdupdate_Click()
Dim jmldata As Integer
    lblwait.Visible = True
    cmdupdate.Enabled = False
    jmldata = 0
    grid.Row = 1
    Do While True
        grid.Col = 0
        If grid.CellPicture = check Then
            OBJ.Open dsn
            SQL = "Update gl_transaksi set idupdate = '1' Where notrx = '" & grid.TextMatrix(grid.Row, 3) & "' "
            SQL = SQL + "and dbkrtrx = '" & grid.TextMatrix(grid.Row, 7) & "' and desctrx = '" & grid.TextMatrix(grid.Row, 6) & "'"
            Set RST = OBJ.Execute(SQL)
            OBJ.Close
            jmldata = jmldata + 1
        End If
        If grid.Row = grid.Rows - 1 Then Exit Do
        grid.Row = grid.Row + 1
    Loop
        If jmldata = 0 Then
            MsgBox "There's no data to update..!", vbExclamation, AppName
            lblwait.Visible = False
            cmdupdate.Enabled = True
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        DoEvents
        hapusgrid
        cbD.Value = Unchecked
        cbK.Value = Unchecked

        btnview.Caption = "Clear"
        opendata
        Screen.MousePointer = vbDefault
        If grid.Rows = 1 Then
            grid.Rows = 2
            Adodc1.Refresh
            grid.Refresh
        End If
        MsgBox jmldata & " data updated !", vbInformation, AppName
        lblwait.Visible = False
        cmdupdate.Enabled = True
End Sub

Private Sub Form_Load()
    date1 = Date
    date2 = Date
    setgrid
End Sub

Private Sub opendata()
    Adodc1.ConnectionString = dsn
    SQL = "Select tgltrx,kdtrx,notrx,kurs,noactrx,desctrx,dbkrtrx,amounttrx,currtrx,cekbg From gl_transaksi "
    SQL = SQL + "Where idupdate <> '1' and tgltrx >= '" & tanggal1 & "' and tgltrx <= '" & tanggal2 & "' "
    SQL = SQL + "and noactrx >= '" & txtacc1 & "' and noactrx <='" & txtacc2 & "'"
    SQL = SQL + "Order By notrx DESC"
    Adodc1.RecordSource = SQL
    Set grid.DataSource = Adodc1
    Adodc1.Refresh
    Adodc1.Recordset.Requery -1
    Pg.Visible = True
    setdata
    grid.Refresh
End Sub

Private Sub setgrid()
    With grid
        .Cols = 11
        .TextMatrix(0, 0) = "No"
        .TextMatrix(0, 1) = "Tgl Trx"
        .TextMatrix(0, 2) = "Kd.Trx"
        .TextMatrix(0, 3) = "No.Trx"
        .TextMatrix(0, 4) = "Kurs"
        .TextMatrix(0, 5) = "No. Acc"
        .TextMatrix(0, 6) = "Description"
        .TextMatrix(0, 7) = "D/K"
        .TextMatrix(0, 8) = "Amount"
        .TextMatrix(0, 9) = "Currency"
        .TextMatrix(0, 10) = "Cek/Giro"
        .ColAlignmentFixed(0) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .ColAlignmentFixed(8) = flexAlignCenterCenter
        .ColAlignmentFixed(10) = flexAlignCenterCenter
        .ColAlignment(0) = flexAlignRightCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignRightCenter
        .ColAlignment(7) = flexAlignCenterCenter
        .ColAlignment(8) = flexAlignRightCenter
        .ColAlignment(9) = flexAlignCenterCenter
        .ColWidth(0) = 800
        .ColWidth(1) = 1000
        .ColWidth(2) = 650
        .ColWidth(3) = 1100
        .ColWidth(4) = 650
        .ColWidth(5) = 1000
        .ColWidth(6) = 5000
        .ColWidth(7) = 400
        .ColWidth(8) = 2000
        .ColWidth(9) = 900
        .ColWidth(10) = 1200
    End With
End Sub
Private Sub setdata()
On Error Resume Next
Dim jml As String
    setgrid
    jml = Adodc1.Recordset.RecordCount
    Pg.Min = 0
    Pg.Max = jml
    Pg.Value = 0
    Pg.text = Str(Int(((Pg.Value / Pg.Max) * 100))) & " % Complete"
    With Adodc1.Recordset
        .MoveFirst
        Do While Not .EOF
            With grid
            grid.Col = 0
            Set grid.CellPicture = uncheck.Picture
            .TextMatrix(.Row, 0) = grid.Row
            .TextMatrix(.Row, 1) = Adodc1.Recordset!tgltrx
            .TextMatrix(.Row, 2) = Adodc1.Recordset!kdtrx
            .TextMatrix(.Row, 3) = Adodc1.Recordset!notrx
            .TextMatrix(.Row, 4) = Adodc1.Recordset!kurs
            .TextMatrix(.Row, 5) = Adodc1.Recordset!noactrx
            .TextMatrix(.Row, 6) = Adodc1.Recordset!desctrx
            .TextMatrix(.Row, 7) = Adodc1.Recordset!dbkrtrx
            .TextMatrix(.Row, 8) = Format(Adodc1.Recordset!amounttrx, "#,##0.00")
            .TextMatrix(.Row, 9) = Adodc1.Recordset!currtrx
            .TextMatrix(.Row, 10) = Adodc1.Recordset!cekbg
            'SetAlternatingGrid grid.Row
            Pg.Value = Pg.Value + 1
            Pg.text = Str(Int(((Pg.Value / Pg.Max) * 100))) & " % Complete"
            
            If grid.Row = jml Then Exit Do
            .Row = .Row + 1
            End With

            lblrecord = grid.Row & " Record"
            Adodc1.Recordset.MoveNext
        Loop
    End With
    Pg.Value = 0
    Pg.Visible = False
End Sub
Private Sub hapusgrid()
On Error Resume Next
    Dim jml As Integer
    jml = grid.Rows
    grid.Row = 1
    Do While True
        grid.Col = 0
        Set grid.CellPicture = blank
        grid.TextMatrix(grid.Row, 0) = ""
        grid.TextMatrix(grid.Row, 1) = ""
        grid.TextMatrix(grid.Row, 2) = ""
        grid.TextMatrix(grid.Row, 3) = ""
        grid.TextMatrix(grid.Row, 4) = ""
        grid.TextMatrix(grid.Row, 5) = ""
        grid.TextMatrix(grid.Row, 6) = ""
        grid.TextMatrix(grid.Row, 7) = ""
        grid.TextMatrix(grid.Row, 8) = ""
        grid.TextMatrix(grid.Row, 9) = ""
        grid.TextMatrix(grid.Row, 10) = ""
        If grid.Row = jml - 1 Then Exit Do
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
End Sub

Private Sub hapusrow()
    grid.TextMatrix(grid.Row, 0) = ""
    grid.TextMatrix(grid.Row, 1) = ""
    grid.TextMatrix(grid.Row, 2) = ""
    grid.TextMatrix(grid.Row, 3) = ""
    grid.TextMatrix(grid.Row, 4) = ""
    grid.TextMatrix(grid.Row, 5) = ""
    grid.TextMatrix(grid.Row, 6) = ""
    grid.TextMatrix(grid.Row, 7) = ""
    grid.TextMatrix(grid.Row, 8) = ""
    grid.TextMatrix(grid.Row, 9) = ""
    grid.TextMatrix(grid.Row, 10) = ""
    Do While True
        grid.TextMatrix(grid.Row, 0) = grid.TextMatrix(grid.Row + 1, 0)
        grid.TextMatrix(grid.Row, 1) = grid.TextMatrix(grid.Row + 1, 1)
        grid.TextMatrix(grid.Row, 2) = grid.TextMatrix(grid.Row + 1, 2)
        grid.TextMatrix(grid.Row, 3) = grid.TextMatrix(grid.Row + 1, 3)
        grid.TextMatrix(grid.Row, 4) = grid.TextMatrix(grid.Row + 1, 4)
        grid.TextMatrix(grid.Row, 5) = grid.TextMatrix(grid.Row + 1, 5)
        grid.TextMatrix(grid.Row, 6) = grid.TextMatrix(grid.Row + 1, 6)
        grid.TextMatrix(grid.Row, 7) = grid.TextMatrix(grid.Row + 1, 7)
        grid.TextMatrix(grid.Row, 8) = grid.TextMatrix(grid.Row + 1, 8)
        grid.TextMatrix(grid.Row, 9) = grid.TextMatrix(grid.Row + 1, 9)
        grid.TextMatrix(grid.Row, 10) = grid.TextMatrix(grid.Row + 1, 10)
        If grid.Row = grid.Rows - 2 Then Exit Do
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = grid.Rows - 1
    grid.Col = 0
    Set grid.CellPicture = blank
End Sub

Function tanggal1()
    tanggal1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggal2()
    tanggal2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function

Private Function SetAlternatingGrid(ByVal i As Integer)
    Dim j, k As Integer
    j = 0
    k = 0
    For k = 1 To grid.Rows - 1
    If grid.TextMatrix(i, 7) = "D" Then
        For j = 0 To grid.Cols - 1
        grid.Col = j
        grid.CellBackColor = &HE0E0E0
        Next
    Else
        grid.CellBackColor = &HFFFFFF
    End If
    Next k
End Function

Private Sub Form_Resize()
    Pg.Move (Me.Width - Pg.Width) / 2, (Me.Height - Pg.Height) / 2
    Me.Top = 2250
    Me.Height = Screen.Height - 3150
    grid.Height = Me.Height - grid.Top - 900
    cmdclose.Top = Me.Height - 800
    cmdupdate.Top = Me.Height - 800
    lblrecord.Top = Me.Height - 800
End Sub

Private Sub grid_Click()
    Dim j As Integer
    j = 0
    If grid.MouseRow = 0 Then Exit Sub
    Select Case grid.Col
        Case 0:
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            If grid.CellPicture = uncheck Then
                Set grid.CellPicture = check
                For j = 0 To grid.Cols - 1
                    grid.Col = j
                    grid.CellBackColor = &HC0FFFF
                Next
            ElseIf grid.CellPicture = check Then
                Set grid.CellPicture = uncheck
                If grid.TextMatrix(grid.Row, 7) = "D" And cbD.Value = Checked Then
                    For j = 0 To grid.Cols - 1
                        grid.Col = j
                        grid.CellBackColor = &HE0E0E0
                    Next
                ElseIf grid.TextMatrix(grid.Row, 7) = "K" And cbK.Value = Checked Then
                    For j = 0 To grid.Cols - 1
                        grid.Col = j
                        grid.CellBackColor = &HC0C0C0
                    Next
                Else
                    For j = 0 To grid.Cols - 1
                        grid.Col = j
                        grid.CellBackColor = &HFFFFFF
                    Next
                End If
            End If
        Case 5:
            carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '01'"
            namatabel = "Company Account"
            frmsearch.Show vbModal
        Case 6:
            If grid.Rows = 2 And grid.TextMatrix(grid.Row, 2) = "" Then Exit Sub
            If txtket.Visible = True Then Exit Sub
            
            txtket.Width = grid.ColWidth(grid.Col) - 40
            txtket = grid.TextMatrix(grid.Row, grid.Col)
            txtket.Left = grid.Left + grid.CellLeft
            txtket.Top = grid.Top + grid.CellTop - 20
            txtket.Height = grid.CellHeight
            txtket.Visible = True
            txtket.SetFocus
        Case 8:
            If grid.Rows = 2 And grid.TextMatrix(grid.Row, 8) = "" Then Exit Sub
            If txtnilai.Visible = True Then Exit Sub
        
            txtnilai.Width = grid.ColWidth(grid.Col) - 40
            txtnilai = grid.TextMatrix(grid.Row, grid.Col)
            txtnilai.Left = grid.Left + grid.CellLeft
            txtnilai.Top = grid.Top + grid.CellTop - 20
            txtnilai.Height = grid.CellHeight
            txtnilai.Visible = True
            txtnilai.SetFocus
        Case 10:
            If grid.Rows = 2 And grid.TextMatrix(grid.Row, 2) = "" Then Exit Sub
            If txtket.Visible = True Then Exit Sub
            
            txtket.Width = grid.ColWidth(grid.Col) - 40
            txtket = grid.TextMatrix(grid.Row, grid.Col)
            txtket.Left = grid.Left + grid.CellLeft
            txtket.Top = grid.Top + grid.CellTop - 20
            txtket.Height = grid.CellHeight
            txtket.Visible = True
            txtket.SetFocus
    End Select
End Sub

Private Sub grid_EnterCell()
    Dim posrow As Integer
    posrow = grid.Row
    Select Case grid.Col
        Case 6:
                If txtket.Visible = True Then Exit Sub
                posrow = grid.Row
                txtket.Width = grid.ColWidth(grid.Col) - 40
                txtket = grid.TextMatrix(grid.Row, grid.Col)
                txtket.Left = grid.Left + grid.CellLeft
                txtket.Top = grid.Top + grid.CellTop - 20
                txtket.Height = grid.CellHeight
                txtket.Visible = True
                txtket.SetFocus
        Case 8:
                If txtnilai.Visible = True Then Exit Sub
                posrow = grid.Row
                txtnilai.Width = grid.ColWidth(grid.Col) - 40
                txtnilai = grid.TextMatrix(grid.Row, grid.Col)
                txtnilai.Left = grid.Left + grid.CellLeft
                txtnilai.Top = grid.Top + grid.CellTop - 20
                txtnilai.Height = grid.CellHeight
                txtnilai.Visible = True
                txtnilai.SetFocus
        Case 10:
                If txtket.Visible = True Then Exit Sub
                posrow = grid.Row
                txtket.Width = grid.ColWidth(grid.Col) - 40
                txtket = grid.TextMatrix(grid.Row, grid.Col)
                txtket.Left = grid.Left + grid.CellLeft
                txtket.Top = grid.Top + grid.CellTop - 20
                txtket.Height = grid.CellHeight
                txtket.Visible = True
                txtket.SetFocus
    End Select
End Sub

Private Sub grid_GotFocus()
    If grid.MouseRow = 0 Then Exit Sub
    Select Case grid.Col
        Case 5:
            If hasil = "" Then Exit Sub
            If MsgBox("Are you sure you want to change this Account..?", vbQuestion + vbYesNo, AppName) = vbYes Then
                grid.SetFocus
                grid.Col = 5
                grid.TextMatrix(grid.Row, 5) = hasil
                OBJ.Open dsn
                SQL = "Update gl_transaksi Set noactrx = '" & hasil & "' Where kdtrx = '" & grid.TextMatrix(grid.Row, 2) & "' "
                SQL = SQL + "and dbkrtrx = '" & grid.TextMatrix(grid.Row, 7) & "' and desctrx = '" & grid.TextMatrix(grid.Row, 6) & "'"
                Set RST = OBJ.Execute(SQL)
                OBJ.Close
                hasil = ""
                hasil1 = ""
            Else
                hasil = ""
                hasil1 = ""
            End If
    End Select
End Sub

Private Sub txtket_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Select Case grid.Col
            Case 6:
            If MsgBox("Are you sure you want to change this data..?", vbQuestion + vbYesNo, AppName) = vbYes Then
                OBJ.Open dsn
                SQL = "Update gl_transaksi Set desctrx = '" & txtket.text & "' Where notrx = '" & grid.TextMatrix(grid.Row, 3) & "' "
                SQL = SQL + "and dbkrtrx = '" & grid.TextMatrix(grid.Row, 7) & "' and desctrx = '" & grid.TextMatrix(grid.Row, 6) & "'"
                Set RST = OBJ.Execute(SQL)
                OBJ.Close
                grid.SetFocus
                grid.Col = 6
                grid.TextMatrix(grid.Row, 6) = txtket.text
            Else
                grid.SetFocus
                grid.Col = 6
            End If
                txtket = ""
                txtket.Visible = False
            Case 10:
            If MsgBox("Are you sure you want to change this data..?", vbQuestion + vbYesNo, AppName) = vbYes Then
                OBJ.Open dsn
                SQL = "Update gl_transaksi Set cekbg = '" & txtket.text & "' Where notrx = '" & grid.TextMatrix(grid.Row, 3) & "' "
                SQL = SQL + "and dbkrtrx = '" & grid.TextMatrix(grid.Row, 7) & "' and desctrx = '" & grid.TextMatrix(grid.Row, 6) & "'"
                Set RST = OBJ.Execute(SQL)
                OBJ.Close
                grid.SetFocus
                grid.Col = 10
                grid.TextMatrix(grid.Row, 10) = txtket.text
            Else
                grid.SetFocus
                grid.Col = 6
            End If
                txtket = ""
                txtket.Visible = False
        End Select
    End If
End Sub

Private Sub txtket_LostFocus()
    txtket.Visible = False
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If MsgBox("Are you sure you want to change this amount..?", vbQuestion + vbYesNo, AppName) = vbYes Then
        OBJ.Open dsn
        SQL = "Update gl_transaksi Set amounttrx = '" & txtnilai.Value & "' Where notrx = '" & grid.TextMatrix(grid.Row, 3) & "' "
        SQL = SQL + "and dbkrtrx = '" & grid.TextMatrix(grid.Row, 7) & "' and desctrx = '" & grid.TextMatrix(grid.Row, 6) & "'"
        Set RST = OBJ.Execute(SQL)
        OBJ.Close
        grid.SetFocus
        grid.Col = 8
        grid.TextMatrix(grid.Row, 8) = txtnilai.Value
        grid.TextMatrix(grid.Row, 8) = Format(grid.TextMatrix(grid.Row, 8), "#,##0.00")
    Else
        grid.SetFocus
        grid.Col = 8
    End If
        txtnilai = ""
        txtnilai.Visible = False
    End If
End Sub

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
End Sub
