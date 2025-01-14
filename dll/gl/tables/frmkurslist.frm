VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmkurslist 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Currency"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6630
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmkurslist.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtarea1 
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
      MaxLength       =   4
      TabIndex        =   0
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtarea2 
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
      MaxLength       =   4
      TabIndex        =   1
      Top             =   1560
      Width           =   615
   End
   Begin Crystal.CrystalReport crystal 
      Left            =   240
      Top             =   2280
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
   Begin Chameleon.chameleonButton cmdsearch1 
      Height          =   285
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "From"
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
      MICON           =   "frmkurslist.frx":2372
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
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   503
      BTYPE           =   8
      TX              =   "To"
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
      MICON           =   "frmkurslist.frx":268C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdview 
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   2040
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Preview"
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
      MICON           =   "frmkurslist.frx":29A6
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
      Left            =   5400
      TabIndex        =   7
      Top             =   2040
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
      MICON           =   "frmkurslist.frx":2CC0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Currency"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label lblnamarea2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   1560
      Width           =   4335
   End
   Begin VB.Label lblnamarea1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10335
   End
End
Attribute VB_Name = "frmkurslist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private Sub cariarea1()
    If txtarea1 = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from gl_kurs where kdkurs = '" & txtarea1 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Type " & txtarea1 & " Not Found.", vbExclamation, "Warning"
        txtarea1 = ""
        lblnamarea1 = ""
        txtarea1.SetFocus
    Else
        lblnamarea1 = RST!nmkurs
    End If
    OBJ.Close
End Sub

Private Sub cariarea2()
    If txtarea2 = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from gl_kurs where kdkurs = '" & txtarea2 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Type " & txtarea2 & " Not Found.", vbExclamation, "Warning"
        txtarea2 = ""
        lblnamarea2 = ""
        txtarea2.SetFocus
    Else
        lblnamarea2 = RST!nmkurs
    End If
    OBJ.Close
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdview_Click()
    If txtarea1 = "" Then txtarea1 = "0"
    If txtarea2 = "" Then txtarea2 = "z"
        
    If txtarea2 < txtarea1 Then
        MsgBox "To Currency Can Not Smaller Then From Currency.", vbExclamation, "Warning"
        txtarea2 = ""
        lblnamarea2 = ""
        txtarea2.SetFocus
        Exit Sub
    End If
    
    crystal.Reset
    crystal.WindowState = crptMaximized
    crystal.WindowShowCloseBtn = True
    crystal.WindowShowPrintSetupBtn = True
    crystal.WindowShowSearchBtn = True
    crystal.Connect = dsnreport
    crystal.DataFiles(0) = "Proc(gl_curlist)"
    crystal.ReportFileName = AppPath & "\reports\gl\tables\curlist.rpt"
    crystal.ParameterFields(0) = "@area1;" + txtarea1 + ";true"
    crystal.ParameterFields(1) = "@area2;" + txtarea2 + ";true"
    crystal.ParameterFields(2) = "@namauser;" + nmuser + ";true"
    crystal.RetrieveDataFiles
    crystal.Action = 1
    
    If txtarea1 = "0" Then txtarea1 = ""
    If txtarea2 = "z" Then txtarea2 = ""
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtarea1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtarea2.SetFocus
End Sub

Private Sub txtarea1_LostFocus()
    cariarea1
End Sub

Private Sub txtarea2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmdview.SetFocus
End Sub

Private Sub txtarea2_LostFocus()
    cariarea2
End Sub

Private Sub cmdsearch1_Click()
    carisql1 = "select kdkurs, nmkurs from gl_kurs"
    namatabel = "Currency"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtarea1 = hasil
    lblnamarea1 = hasil1
    hasil = ""
    hasil1 = ""
    txtarea2.SetFocus
End Sub

Private Sub cmdsearch2_Click()
    carisql1 = "select kdkurs, nmkurs from gl_kurs"
    namatabel = "Currency"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch2_GotFocus()
    If hasil = "" Then Exit Sub
    txtarea2 = hasil
    lblnamarea2 = hasil1
    hasil = ""
    hasil1 = ""
    cmdview.SetFocus
End Sub

Function nmuser()
    nmuser = "-no user-"
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select * from am_user where kodeuser = '" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    nmuser = RST!namauser
    '    OBJ.Close
    'End If
End Function
