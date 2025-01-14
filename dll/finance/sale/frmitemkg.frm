VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmitemkg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Define Kilogram for Base Unit"
   ClientHeight    =   7815
   ClientLeft      =   3615
   ClientTop       =   3105
   ClientWidth     =   14655
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
   Icon            =   "frmitemkg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   14655
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option4 
      Caption         =   "8 digit item"
      Height          =   255
      Left            =   12120
      TabIndex        =   23
      Top             =   120
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton Option3 
      Caption         =   "5 digit item"
      Height          =   255
      Left            =   13320
      TabIndex        =   22
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtkode1 
      Appearance      =   0  'Flat
      DataField       =   "KodeArea"
      Height          =   285
      Left            =   6720
      MaxLength       =   10
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtKode 
      Appearance      =   0  'Flat
      DataField       =   "KodeArea"
      Height          =   285
      Left            =   4080
      MaxLength       =   10
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "All Items"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      ToolTipText     =   "120"
      Top             =   150
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   10
      Top             =   5640
      Visible         =   0   'False
      Width           =   3255
      Begin VB.OptionButton Option5 
         Caption         =   "From this year to last year"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   3015
      End
      Begin VB.OptionButton Option2 
         Caption         =   "From December this year to January next year"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   3015
      End
      Begin Chameleon.chameleonButton cmdclosecopy 
         Height          =   375
         Left            =   2280
         TabIndex        =   17
         Top             =   1560
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
         MICON           =   "frmitemkg.frx":27A2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdcopy 
         Height          =   375
         Left            =   1320
         TabIndex        =   16
         Top             =   1560
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Copy"
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
         MICON           =   "frmitemkg.frx":2ABC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.OptionButton Option1 
         Caption         =   "On this year"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin MSForms.ComboBox cmb1 
         Height          =   285
         Left            =   360
         TabIndex        =   12
         Top             =   490
         Width           =   855
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "1508;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmb2 
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Top             =   490
         Width           =   855
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "1508;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "to"
         Height          =   255
         Left            =   1320
         TabIndex        =   19
         Top             =   520
         Width           =   255
      End
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   6480
      Visible         =   0   'False
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   450
      Calculator      =   "frmitemkg.frx":2DD6
      Caption         =   "frmitemkg.frx":2DF6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmitemkg.frx":2E62
      Keys            =   "frmitemkg.frx":2E80
      Spin            =   "frmitemkg.frx":2EC2
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
   Begin TDBNumber6Ctl.TDBNumber txtahun 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   503
      Calculator      =   "frmitemkg.frx":2EEA
      Caption         =   "frmitemkg.frx":2F0A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmitemkg.frx":2F6F
      Keys            =   "frmitemkg.frx":2F8D
      Spin            =   "frmitemkg.frx":2FD7
      AlignHorizontal =   2
      AlignVertical   =   0
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
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   2100
      MinValue        =   2005
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   1245189
      Value           =   2005
      MaxValueVT      =   1330839557
      MinValueVT      =   1431175173
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   13560
      TabIndex        =   8
      Top             =   7320
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
      MICON           =   "frmitemkg.frx":2FFF
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsubmit 
      Height          =   285
      Left            =   8520
      TabIndex        =   4
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "Submit"
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
      MICON           =   "frmitemkg.frx":3319
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
      Height          =   6735
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   11880
      _Version        =   393216
      Cols            =   16
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
      _Band(0).Cols   =   16
   End
   Begin Chameleon.chameleonButton cmdsave 
      Height          =   375
      Left            =   12600
      TabIndex        =   7
      Top             =   7320
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
      MICON           =   "frmitemkg.frx":3633
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdshowcopy 
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   7320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Show Copy"
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
      MICON           =   "frmitemkg.frx":394D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch 
      Height          =   285
      Left            =   3120
      TabIndex        =   20
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "From Item"
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmitemkg.frx":3C67
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch1 
      Height          =   285
      Left            =   5760
      TabIndex        =   21
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "To Item"
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmitemkg.frx":3F81
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      Caption         =   "Tahun"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   150
      Width           =   615
   End
End
Attribute VB_Name = "frmitemkg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim posrow, poscol As String
Dim g, h, i As Integer

Private Sub Check1_Click()
    txtKode = ""
    txtkode1 = ""
End Sub

Private Sub cmb1_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub cmb1_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
End Sub

Private Sub cmb2_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub cmb2_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
End Sub

Private Sub cmdclosecopy_Click()
    cmb1 = ""
    cmb2 = ""
    Frame1.Visible = False
End Sub

Private Sub cmdcopy_Click()
    If MsgBox("Continue copy ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    If grid.Rows = 1 Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    If (cmb1 = "" Or cmb2 = "") And Option1.Value = True Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    If cmb1 = cmb2 And Option1.Value = True Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If Option1.Value = True Then
        g = cmb1.ListIndex + 4
        h = cmb2.ListIndex + 4
        
        For i = 1 To grid.Rows - 1
            grid.TextMatrix(i, h) = grid.TextMatrix(i, g)
        Next i
    ElseIf Option2.Value = True Then
        grid.Row = 1
        Do While True
            OBJ.Open dsn
            SQL = "select * from am_itemkg where tahun = '" & txtahun + 1 & "'"
            SQL = SQL + " and kodebarang = '" & Trim(grid.TextMatrix(grid.Row, 0)) & "'"
            SQL = SQL + " and kodesatuan = '" & Trim(grid.TextMatrix(grid.Row, 2)) & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                SQL = "delete from am_itemkg where tahun = '" & txtahun + 1 & "'"
                SQL = SQL + " and kodebarang = '" & Trim(grid.TextMatrix(grid.Row, 0)) & "'"
                SQL = SQL + " and kodesatuan = '" & Trim(grid.TextMatrix(grid.Row, 2)) & "'"
                Set RST = OBJ.Execute(SQL)
            End If
            OBJ.Close
            
            OBJ.Open dsn
            SQL = "insert into am_itemkg ("
            SQL = SQL + "kodebarang, "
            SQL = SQL + "kodesatuan, "
            SQL = SQL + "kg1, "
            SQL = SQL + "kg2, "
            SQL = SQL + "kg3, "
            SQL = SQL + "kg4, "
            SQL = SQL + "kg5, "
            SQL = SQL + "kg6, "
            SQL = SQL + "kg7, "
            SQL = SQL + "kg8, "
            SQL = SQL + "kg9, "
            SQL = SQL + "kg10, "
            SQL = SQL + "kg11, "
            SQL = SQL + "kg12, "
            SQL = SQL + "tahun)"
        
            SQL = SQL + " values('" & Trim(grid.TextMatrix(grid.Row, 0)) & "',"
            SQL = SQL + "'" & Trim(grid.TextMatrix(grid.Row, 2)) & "',"
            SQL = SQL + "convert(money,'" & Val(Format(grid.TextMatrix(grid.Row, 15), "general number")) & "'),"
            SQL = SQL + "convert(money,'0'),"
            SQL = SQL + "convert(money,'0'),"
            SQL = SQL + "convert(money,'0'),"
            SQL = SQL + "convert(money,'0'),"
            SQL = SQL + "convert(money,'0'),"
            SQL = SQL + "convert(money,'0'),"
            SQL = SQL + "convert(money,'0'),"
            SQL = SQL + "convert(money,'0'),"
            SQL = SQL + "convert(money,'0'),"
            SQL = SQL + "convert(money,'0'),"
            SQL = SQL + "convert(money,'0'),"
            SQL = SQL + "'" & txtahun + 1 & "')"
            Set RST = OBJ.Execute(SQL)
            OBJ.Close
            
            If grid.Rows - 1 = grid.Row Then Exit Do
            
            grid.Row = grid.Row + 1
        Loop
    ElseIf Option5.Value = True Then
        If MsgBox("Are you sure wanto continue this action ?" & vbCrLf & "This action will change last year value.", vbOKCancel + vbQuestion, "Warning") = vbCancel Then Exit Sub
        
        grid.Row = 1
        Do While True
            OBJ.Open dsn
            SQL = "select * from am_itemkg where tahun = '" & txtahun - 1 & "'"
            SQL = SQL + " and kodebarang = '" & Trim(grid.TextMatrix(grid.Row, 0)) & "'"
            SQL = SQL + " and kodesatuan = '" & Trim(grid.TextMatrix(grid.Row, 2)) & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                SQL = "delete from am_itemkg where tahun = '" & txtahun - 1 & "'"
                SQL = SQL + " and kodebarang = '" & Trim(grid.TextMatrix(grid.Row, 0)) & "'"
                SQL = SQL + " and kodesatuan = '" & Trim(grid.TextMatrix(grid.Row, 2)) & "'"
                Set RST = OBJ.Execute(SQL)
            End If
            OBJ.Close
            
            OBJ.Open dsn
            SQL = "insert into am_itemkg ("
            SQL = SQL + "kodebarang, "
            SQL = SQL + "kodesatuan, "
            SQL = SQL + "kg1, "
            SQL = SQL + "kg2, "
            SQL = SQL + "kg3, "
            SQL = SQL + "kg4, "
            SQL = SQL + "kg5, "
            SQL = SQL + "kg6, "
            SQL = SQL + "kg7, "
            SQL = SQL + "kg8, "
            SQL = SQL + "kg9, "
            SQL = SQL + "kg10, "
            SQL = SQL + "kg11, "
            SQL = SQL + "kg12, "
            SQL = SQL + "tahun)"
        
            SQL = SQL + " values('" & Trim(grid.TextMatrix(grid.Row, 0)) & "',"
            SQL = SQL + "'" & Trim(grid.TextMatrix(grid.Row, 2)) & "',"
            SQL = SQL + "convert(money,'" & Val(Format(grid.TextMatrix(grid.Row, 4), "general number")) & "'),"
            SQL = SQL + "convert(money,'" & Val(Format(grid.TextMatrix(grid.Row, 5), "general number")) & "'),"
            SQL = SQL + "convert(money,'" & Val(Format(grid.TextMatrix(grid.Row, 6), "general number")) & "'),"
            SQL = SQL + "convert(money,'" & Val(Format(grid.TextMatrix(grid.Row, 7), "general number")) & "'),"
            SQL = SQL + "convert(money,'" & Val(Format(grid.TextMatrix(grid.Row, 8), "general number")) & "'),"
            SQL = SQL + "convert(money,'" & Val(Format(grid.TextMatrix(grid.Row, 9), "general number")) & "'),"
            SQL = SQL + "convert(money,'" & Val(Format(grid.TextMatrix(grid.Row, 10), "general number")) & "'),"
            SQL = SQL + "convert(money,'" & Val(Format(grid.TextMatrix(grid.Row, 11), "general number")) & "'),"
            SQL = SQL + "convert(money,'" & Val(Format(grid.TextMatrix(grid.Row, 12), "general number")) & "'),"
            SQL = SQL + "convert(money,'" & Val(Format(grid.TextMatrix(grid.Row, 13), "general number")) & "'),"
            SQL = SQL + "convert(money,'" & Val(Format(grid.TextMatrix(grid.Row, 14), "general number")) & "'),"
            SQL = SQL + "convert(money,'" & Val(Format(grid.TextMatrix(grid.Row, 15), "general number")) & "'),"
            SQL = SQL + "'" & txtahun - 1 & "')"
            Set RST = OBJ.Execute(SQL)
            OBJ.Close
            
            If grid.Rows - 1 = grid.Row Then Exit Do
            
            grid.Row = grid.Row + 1
        Loop
    End If
    
    MsgBox "Copy complete, Click OK To Continue ...", vbInformation, "Information"
End Sub

Private Sub cmdsave_Click()
    If MsgBox("Save this change ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    If grid.Rows = 1 Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    grid.Row = 1
    Do While True
        OBJ.Open dsn
        SQL = "select * from am_itemkg where tahun = '" & txtahun & "'"
        SQL = SQL + " and kodebarang = '" & Trim(grid.TextMatrix(grid.Row, 0)) & "'"
        SQL = SQL + " and kodesatuan = '" & Trim(grid.TextMatrix(grid.Row, 2)) & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            SQL = "delete from am_itemkg where tahun = '" & txtahun & "'"
            SQL = SQL + " and kodebarang = '" & Trim(grid.TextMatrix(grid.Row, 0)) & "'"
            SQL = SQL + " and kodesatuan = '" & Trim(grid.TextMatrix(grid.Row, 2)) & "'"
            Set RST = OBJ.Execute(SQL)
        End If
        OBJ.Close
        
        OBJ.Open dsn
        SQL = "insert into am_itemkg ("
        SQL = SQL + "kodebarang, "
        SQL = SQL + "kodesatuan, "
        SQL = SQL + "kg1, "
        SQL = SQL + "kg2, "
        SQL = SQL + "kg3, "
        SQL = SQL + "kg4, "
        SQL = SQL + "kg5, "
        SQL = SQL + "kg6, "
        SQL = SQL + "kg7, "
        SQL = SQL + "kg8, "
        SQL = SQL + "kg9, "
        SQL = SQL + "kg10, "
        SQL = SQL + "kg11, "
        SQL = SQL + "kg12, "
        SQL = SQL + "tahun)"
    
        SQL = SQL + " values('" & Trim(grid.TextMatrix(grid.Row, 0)) & "',"
        SQL = SQL + "'" & Trim(grid.TextMatrix(grid.Row, 2)) & "',"
        SQL = SQL + "convert(money,'" & Val(Format(grid.TextMatrix(grid.Row, 4), "general number")) & "'),"
        SQL = SQL + "convert(money,'" & Val(Format(grid.TextMatrix(grid.Row, 5), "general number")) & "'),"
        SQL = SQL + "convert(money,'" & Val(Format(grid.TextMatrix(grid.Row, 6), "general number")) & "'),"
        SQL = SQL + "convert(money,'" & Val(Format(grid.TextMatrix(grid.Row, 7), "general number")) & "'),"
        SQL = SQL + "convert(money,'" & Val(Format(grid.TextMatrix(grid.Row, 8), "general number")) & "'),"
        SQL = SQL + "convert(money,'" & Val(Format(grid.TextMatrix(grid.Row, 9), "general number")) & "'),"
        SQL = SQL + "convert(money,'" & Val(Format(grid.TextMatrix(grid.Row, 10), "general number")) & "'),"
        SQL = SQL + "convert(money,'" & Val(Format(grid.TextMatrix(grid.Row, 11), "general number")) & "'),"
        SQL = SQL + "convert(money,'" & Val(Format(grid.TextMatrix(grid.Row, 12), "general number")) & "'),"
        SQL = SQL + "convert(money,'" & Val(Format(grid.TextMatrix(grid.Row, 13), "general number")) & "'),"
        SQL = SQL + "convert(money,'" & Val(Format(grid.TextMatrix(grid.Row, 14), "general number")) & "'),"
        SQL = SQL + "convert(money,'" & Val(Format(grid.TextMatrix(grid.Row, 15), "general number")) & "'),"
        SQL = SQL + "'" & txtahun & "')"
        Set RST = OBJ.Execute(SQL)
        OBJ.Close
        
        If grid.Rows - 1 = grid.Row Then Exit Do
        
        grid.Row = grid.Row + 1
    Loop
    
    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
End Sub

Private Sub cmdsearch_Click()
    If Check1.Value = 1 Then Exit Sub
    
    If Option3.Value = True Then carisql1 = "select kodebarang, namabarang from am_itemmst where len(kodebarang)=5"
    If Option4.Value = True Then carisql1 = "select kodebarang, namabarang from am_itemmst where len(kodebarang)=8"
    namatabel = "Barang "
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    txtKode = hasil
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdsearch1_Click()
    If Check1.Value = 1 Then Exit Sub

    If Option3.Value = True Then carisql1 = "select kodebarang, namabarang from am_itemmst where len(kodebarang)=5"
    If Option4.Value = True Then carisql1 = "select kodebarang, namabarang from am_itemmst where len(kodebarang)=8"
    namatabel = "Barang "
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtkode1 = hasil
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdshowcopy_Click()
    Frame1.Visible = True
End Sub

Private Sub cmdsubmit_Click()
    If Len(txtahun) <> 4 Then
        MsgBox "Data Entry Not Complete. (Year format YYYY)", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If Not (txtahun >= 2005 And txtahun <= 2100) Then
        MsgBox "Data Entry Not Complete. (min year 2005 and max year 2100)", vbExclamation, "Warning"
        Exit Sub
    End If
    
    grid.FixedCols = 0
    If Check1.Value = 1 Then
        OBJ.Open dsn
        SQL = "select d.kodebarang,d.namabarang,d.kodesatuan,e.namasatuan,"
        SQL = SQL + "isnull(f.kg1,0)'Jan',"
        SQL = SQL + "isnull(f.kg2,0)'Feb',"
        SQL = SQL + "isnull(f.kg3,0)'Mar',"
        SQL = SQL + "isnull(f.kg4,0)'Apr',"
        SQL = SQL + "isnull(f.kg5,0)'May',"
        SQL = SQL + "isnull(f.kg6,0)'Jun',"
        SQL = SQL + "isnull(f.kg7,0)'Jul',"
        SQL = SQL + "isnull(f.kg8,0)'Aug',"
        SQL = SQL + "isnull(f.kg9,0)'Sep',"
        SQL = SQL + "isnull(f.kg10,0)'Okt',"
        SQL = SQL + "isnull(f.kg11,0)'Nov',"
        SQL = SQL + "isnull(f.kg12,0)'Des'"
        SQL = SQL + " from am_itemdtl d left join am_unit e on d.kodesatuan=e.kodesatuan "
        SQL = SQL + "left join am_itemkg f on d.kodebarang=f.kodebarang and d.kodesatuan=f.kodesatuan and f.tahun = '" & txtahun & "'"
        If Option4.Value = True Then SQL = SQL + " where len(d.kodebarang)=8 order by d.kodebarang asc,d.kodesatuan desc"
        If Option3.Value = True Then SQL = SQL + " where len(d.kodebarang)=5 order by d.kodebarang asc,d.kodesatuan desc"
        Set RST = OBJ.Execute(SQL)
        Set grid.DataSource = RST
        OBJ.Close
    ElseIf Check1.Value = 0 Then
        If txtKode = "" Or txtkode1 = "" Then
            MsgBox "Data entry not Complete.", vbExclamation, "Warning"
            Exit Sub
        End If
        
        If txtkode1 < txtKode Then
            MsgBox "To... can not Smaller Then From...", vbExclamation, "Warning"
            txtkode1 = ""
            txtkode1.SetFocus
            Exit Sub
        End If
        
        OBJ.Open dsn
        SQL = "select d.kodebarang,d.namabarang,d.kodesatuan,e.namasatuan,"
        SQL = SQL + "isnull(f.kg1,0)'Jan',"
        SQL = SQL + "isnull(f.kg2,0)'Feb',"
        SQL = SQL + "isnull(f.kg3,0)'Mar',"
        SQL = SQL + "isnull(f.kg4,0)'Apr',"
        SQL = SQL + "isnull(f.kg5,0)'May',"
        SQL = SQL + "isnull(f.kg6,0)'Jun',"
        SQL = SQL + "isnull(f.kg7,0)'Jul',"
        SQL = SQL + "isnull(f.kg8,0)'Aug',"
        SQL = SQL + "isnull(f.kg9,0)'Sep',"
        SQL = SQL + "isnull(f.kg10,0)'Okt',"
        SQL = SQL + "isnull(f.kg11,0)'Nov',"
        SQL = SQL + "isnull(f.kg12,0)'Des'"
        SQL = SQL + " from am_itemdtl d left join am_unit e on d.kodesatuan=e.kodesatuan "
        SQL = SQL + "left join am_itemkg f on d.kodebarang=f.kodebarang and d.kodesatuan=f.kodesatuan and f.tahun = '" & txtahun & "'"
        If Option4.Value = True Then SQL = SQL + " where len(d.kodebarang)=8 and d.kodebarang>='" & txtKode & "' and d.kodebarang<='" & txtkode1 & "' order by d.kodebarang asc,d.kodesatuan desc"
        If Option3.Value = True Then SQL = SQL + " where len(d.kodebarang)=5 and d.kodebarang>='" & txtKode & "' and d.kodebarang<='" & txtkode1 & "' order by d.kodebarang asc,d.kodesatuan desc"
        Set RST = OBJ.Execute(SQL)
        Set grid.DataSource = RST
        OBJ.Close
    End If
    
    grid.ColWidth(0) = 1000
    grid.ColWidth(1) = 1000
    grid.ColWidth(2) = 1000
    grid.ColWidth(3) = 1000
    grid.ColWidth(4) = 1000
    grid.ColWidth(5) = 1000
    grid.ColWidth(6) = 1000
    grid.ColWidth(7) = 1000
    grid.ColWidth(8) = 1000
    grid.ColWidth(9) = 1000
    grid.ColWidth(10) = 1000
    grid.ColWidth(11) = 1000
    grid.ColWidth(12) = 1000
    grid.ColWidth(13) = 1000
    grid.ColWidth(14) = 1000
    grid.ColWidth(15) = 1000
    grid.ColWidth(16) = 1000
    
    grid.RowHeightMin = 300
    
    grid.MergeCells = 2
    grid.MergeCol(0) = True
    grid.MergeCol(1) = True
    
    grid.FixedCols = 4
End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='384' and b.kodeuser = '1" & kuser & "'"
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

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    txtahun = Year(Date)
    
    cmb1.AddItem "Jan"
    cmb1.AddItem "Feb"
    cmb1.AddItem "Mar"
    cmb1.AddItem "Apr"
    cmb1.AddItem "May"
    cmb1.AddItem "Jun"
    cmb1.AddItem "Jul"
    cmb1.AddItem "Aug"
    cmb1.AddItem "Sep"
    cmb1.AddItem "Oct"
    cmb1.AddItem "Nov"
    cmb1.AddItem "Des"
    
    cmb2.AddItem "Jan"
    cmb2.AddItem "Feb"
    cmb2.AddItem "Mar"
    cmb2.AddItem "Apr"
    cmb2.AddItem "May"
    cmb2.AddItem "Jun"
    cmb2.AddItem "Jul"
    cmb2.AddItem "Aug"
    cmb2.AddItem "Sep"
    cmb2.AddItem "Oct"
    cmb2.AddItem "Nov"
    cmb2.AddItem "Des"
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    
    posrow = grid.Row
    poscol = grid.Col
    Select Case grid.Col
        Case 4 To 16
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            
            txtnilai.Width = grid.ColWidth(grid.Col) - 40
            txtnilai = grid.TextMatrix(grid.Row, grid.Col)
            txtnilai.Left = grid.Left + grid.CellLeft
            txtnilai.Top = grid.Top + grid.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
    End Select
End Sub

Private Sub grid_EnterCell()
    Select Case grid.Col
    Case 4 To 16
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
        If grid.Row = 0 Then Exit Sub
    
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

Private Sub grid_Scroll()
    txtnilai.Visible = False
End Sub

Private Sub Option1_Click()
    cmb1.Enabled = True
    cmb2.Enabled = True
End Sub

Private Sub Option2_Click()
    cmb1 = ""
    cmb2 = ""
    cmb1.Enabled = False
    cmb2.Enabled = False
End Sub

Private Sub txtkode_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtkode_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtKode1_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtKode1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtnilai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then
        If grid.Row > 1 Then grid.Row = grid.Row - 1
        
        grid_EnterCell
    ElseIf KeyCode = 40 Then
        If grid.Rows - 1 <> grid.Row Then grid.Row = grid.Row + 1
        
        grid_EnterCell
    End If
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grid.TextMatrix(grid.Row, grid.Col) = Format(txtnilai, "###,###,##0.00")
        
        txtnilai = 0
        txtnilai.Visible = False
        grid.SetFocus
        grid.Row = posrow
    ElseIf KeyAscii = 27 Then
        txtnilai = 0
        txtnilai.Visible = False
    End If
End Sub

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
End Sub
