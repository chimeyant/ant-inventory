VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Begin VB.Form frmbank 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Bank"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   610
      Left            =   1320
      TabIndex        =   9
      Top             =   1230
      Visible         =   0   'False
      Width           =   3495
      Begin VB.ListBox List1 
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
         Height          =   615
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   3495
      End
   End
   Begin VB.TextBox txtacc 
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
      MaxLength       =   15
      TabIndex        =   2
      Top             =   960
      Width           =   3495
   End
   Begin VB.TextBox txtkode 
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
      MaxLength       =   5
      TabIndex        =   0
      Top             =   240
      Width           =   735
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
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   1
      Top             =   600
      Width           =   3495
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   1440
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
      MICON           =   "frmbank.frx":0000
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
      Left            =   3000
      TabIndex        =   4
      Top             =   1440
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
      MICON           =   "frmbank.frx":031A
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
      Left            =   2040
      TabIndex        =   3
      Top             =   1440
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
      MICON           =   "frmbank.frx":0634
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
      Caption         =   "Acc Number"
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
      TabIndex        =   8
      Top             =   990
      Width           =   975
   End
   Begin VB.Label lblcatcode 
      Caption         =   "Kode Bank"
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
      TabIndex        =   7
      Top             =   270
      Width           =   975
   End
   Begin VB.Label lbldesc 
      Caption         =   "Description"
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
      TabIndex        =   6
      Top             =   630
      Width           =   975
   End
End
Attribute VB_Name = "frmbank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim int1 As Integer

Private Sub cmdadd_Click()
    If Len(Trim(txtkode)) = 0 Then
        MsgBox "Data entry not Complete.", vbExclamation, "Warning"
        txtkode.SetFocus
        Exit Sub
    End If
    
    If txtkode = "" Or txtdesc = "" Or txtacc = "" Then
        MsgBox "Data entry not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    txtkode = Trim(txtkode)
    
    OBJ.Open dsn
    SQL = "SELECT * FROM am_bank WHERE Kode = '" & txtkode & " '"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Data already exist.", vbInformation, "Information"
        cmdclear_Click
        OBJ.Close
        Exit Sub
    End If
    
    txtacc = Trim(txtacc)
    SQL = "SELECT * FROM am_bank WHERE acc = '" & txtacc & " '"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Account number already exist.", vbInformation, "Information"
        txtacc.SetFocus
        OBJ.Close
        Exit Sub
    End If
    
    SQL = "INSERT INTO am_bank"
    SQL = SQL + "(Kode"
    SQL = SQL + ",description"
    SQL = SQL + ",acc"
    SQL = SQL + ",flag)"
    
    SQL = SQL + "VALUES"
    SQL = SQL + "('" & txtkode & "'"
    SQL = SQL + ", '" & txtdesc & "'"
    SQL = SQL + ", '" & txtacc & "'"
    SQL = SQL + ", 'x')"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Data saved, click ok to continue  ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdclear_Click()
    txtkode = ""
    txtdesc = ""
    txtacc = ""
    txtkode.SetFocus
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='71' and b.kodeuser = '2" & kuser & "'"
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
    'Me.Top = (frmainmenu.Height - Me.Height) / 2
    'Me.Left = (frmainmenu.Width - Me.Width) / 2
    
    txtkode.ToolTipText = "max length = " & txtkode.MaxLength
    txtdesc.ToolTipText = "max length = " & txtdesc.MaxLength
End Sub

Private Sub txtacc_Change()
    txtacc = Trim(txtacc)
    If txtacc = "" Then
        Frame1.Visible = False
        Exit Sub
    End If
    
    Frame1.Visible = True
    List1.Clear
    int1 = 0
    
    OBJ.Open dsn
    SQL = "select acc from am_bank where acc like '%" & txtacc & "%'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        List1.AddItem RST!acc, int1
        
        int1 = int1 + 1
        RST.MoveNext
    Loop
    OBJ.Close
End Sub

Private Sub txtacc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Frame1.Visible = False
    If KeyAscii = 13 Then
        Frame1.Visible = False
        cmdadd.SetFocus
    End If
End Sub

Private Sub txtdesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtacc.SetFocus
End Sub

Private Sub txtKode_GotFocus()
   ' Call Blok(txtkode)
End Sub

Private Sub txtKode_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtdesc.SetFocus
End Sub

Private Sub txtKode_LostFocus()
    If txtkode = "" Then Exit Sub
    If txtkode.SelLength <> 0 Then Exit Sub
    
    OBJ.Open dsn
    SQL = "SELECT * FROM am_bank WHERE Kode = '" & txtkode & " '"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtdesc = RST!Description
        txtacc = RST!acc
        
        MsgBox "Data already exist.", vbInformation, "Information"
        txtkode.SetFocus
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    txtdesc = ""
    txtacc = ""
    txtdesc.SetFocus
End Sub
