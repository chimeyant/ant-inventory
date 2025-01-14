VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Begin VB.Form frmaging 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Define Aging and Header"
   ClientHeight    =   2295
   ClientLeft      =   5100
   ClientTop       =   5385
   ClientWidth     =   6015
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtKet4 
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
      Left            =   3600
      MaxLength       =   15
      TabIndex        =   6
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox txtKet3 
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
      Left            =   3600
      MaxLength       =   15
      TabIndex        =   5
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtKet2 
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
      Left            =   3600
      MaxLength       =   15
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox txtKet1 
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
      Left            =   3600
      MaxLength       =   15
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin TDBNumber6Ctl.TDBNumber txtaging2 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   503
      Calculator      =   "frmaging.frx":0000
      Caption         =   "frmaging.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmaging.frx":008C
      Keys            =   "frmaging.frx":00AA
      Spin            =   "frmaging.frx":00F4
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##0;;0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   -1
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   846397445
      MinValueVT      =   1325400069
   End
   Begin TDBNumber6Ctl.TDBNumber txtaging3 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   960
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   503
      Calculator      =   "frmaging.frx":011C
      Caption         =   "frmaging.frx":013C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmaging.frx":01A8
      Keys            =   "frmaging.frx":01C6
      Spin            =   "frmaging.frx":0210
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##0;;0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   -1
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   846397445
      MinValueVT      =   1325400069
   End
   Begin TDBNumber6Ctl.TDBNumber txtaging1 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   495
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   503
      Calculator      =   "frmaging.frx":0238
      Caption         =   "frmaging.frx":0258
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmaging.frx":02C4
      Keys            =   "frmaging.frx":02E2
      Spin            =   "frmaging.frx":032C
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##0;;0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   -1
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   846397445
      MinValueVT      =   1325400069
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   1800
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
      MICON           =   "frmaging.frx":0354
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdadd 
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   1800
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
      MICON           =   "frmaging.frx":066E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      Caption         =   "Batasan 4 s/d               ~  hari         Kolom 4"
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
      TabIndex        =   12
      Top             =   1350
      Width           =   3975
   End
   Begin VB.Label Label7 
      Caption         =   "Batasan 3 s/d                    hari         Kolom 3"
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
      TabIndex        =   11
      Top             =   990
      Width           =   3975
   End
   Begin VB.Label Label5 
      Caption         =   "Batasan 2 s/d                    hari         Kolom 2"
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
      TabIndex        =   10
      Top             =   630
      Width           =   4095
   End
   Begin VB.Label Label3 
      Caption         =   "Batasan 1 s/d                    hari         Kolom 1"
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
      TabIndex        =   9
      Top             =   270
      Width           =   4335
   End
End
Attribute VB_Name = "frmaging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private Sub cmdadd_Click()
    If txtaging1 = 0 Or txtKet1 = "" Or txtaging2 = 0 Or txtKet2 = "" Or txtaging3 = 0 Or txtKet3 = "" Or txtKet4 = "" Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If Val(txtaging1) > Val(txtaging2) Then
        txtaging2 = 0
        txtaging2.SetFocus
        Exit Sub
    End If
    
    If Val(txtaging2) > Val(txtaging3) Then
        txtaging3 = 0
        txtaging3.SetFocus
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "delete FROM AM_Aging"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "INSERT INTO AM_Aging"
    SQL = SQL + "(kolom1"
    SQL = SQL + ",kolom2"
    SQL = SQL + ",kolom3"
    SQL = SQL + ",kolom4"
    SQL = SQL + ",desc1"
    SQL = SQL + ",desc2"
    SQL = SQL + ",desc3"
    SQL = SQL + ",desc4"
    SQL = SQL + ",desc5)"

    SQL = SQL + "VALUES"
    SQL = SQL + "(convert(money,'" & txtaging1 & "')"
    SQL = SQL + ",convert(money,'" & txtaging2 & "')"
    SQL = SQL + ",convert(money,'" & txtaging3 & "')"
    SQL = SQL + ",convert(money,'0')"
    SQL = SQL + ",'" & txtKet1 & "'"
    SQL = SQL + ",'" & txtKet2 & "'"
    SQL = SQL + ",'" & txtKet3 & "'"
    SQL = SQL + ",'" & txtKet4 & "'"
    SQL = SQL + ",'-')"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='364' and b.kodeuser = '1" & kuser & "'"
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
    
    OBJ.Open dsn
    SQL = "select * from am_aging"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtaging1 = RST!kolom1
        txtaging2 = RST!kolom2
        txtaging3 = RST!kolom3
        
        txtKet1 = RST!desc1
        txtKet2 = RST!desc2
        txtKet3 = RST!desc3
        txtKet4 = RST!desc4
    End If
    OBJ.Close
End Sub

Private Sub txtaging1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtaging2.SetFocus
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtaging1_LostFocus()
    If txtaging1 = 0 Then Exit Sub
    txtKet1 = "1 - " & txtaging1 & " Hari"
End Sub

Private Sub txtaging2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtaging3.SetFocus
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtaging2_LostFocus()
    If txtaging2 = 0 Then Exit Sub
    If Val(txtaging1) > Val(txtaging2) Then
        txtaging2 = 0
        txtaging2.SetFocus
        Exit Sub
    End If
    txtKet2 = (Val(txtaging1) + 1) & " - " & txtaging2 & " Hari"
End Sub

Private Sub txtAging3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdadd.SetFocus
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txtAging3_LostFocus()
    If txtaging3 = 0 Then Exit Sub
    If Val(txtaging2) > Val(txtaging3) Then
        txtaging3 = 0
        txtaging3.SetFocus
        Exit Sub
    End If
    txtKet3 = (Val(txtaging2) + 1) & " - " & txtaging3 & " Hari"
    txtKet4 = "> " & txtaging3 & " Hari"
End Sub
