VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Begin VB.Form frmcategory 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Category"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtcatcode 
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
      Left            =   1560
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
      Left            =   1560
      MaxLength       =   30
      TabIndex        =   1
      Top             =   600
      Width           =   3615
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   1080
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
      MICON           =   "frmcategory.frx":0000
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
      Left            =   3360
      TabIndex        =   3
      Top             =   1080
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
      MICON           =   "frmcategory.frx":031A
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
      Left            =   2400
      TabIndex        =   2
      Top             =   1080
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
      MICON           =   "frmcategory.frx":0634
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblcatcode 
      Caption         =   "Kode Category"
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
      Top             =   270
      Width           =   1215
   End
   Begin VB.Label lbldesc 
      Caption         =   "Nama Category"
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
      TabIndex        =   5
      Top             =   630
      Width           =   1335
   End
End
Attribute VB_Name = "frmcategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private Sub cmdadd_Click()
    If Len(Trim(txtcatcode)) = 0 Then
        MsgBox "Data entry not Complete.", vbExclamation, "Warning"
        txtcatcode.SetFocus
        Exit Sub
    End If
    
    If txtcatcode = "" Or txtdesc = "" Then
        MsgBox "Data entry not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If txtcatcode = "C999" Then
        MsgBox "C999 is Bonus Category, user cannot manualy add this category.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    txtcatcode = Trim(txtcatcode)
    
    OBJ.Open dsn
    SQL = "SELECT * FROM AM_PRODUK WHERE KodeProduk = '" & txtcatcode & " '"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Data already exist.", vbInformation, "Information"
        cmdclear_Click
        OBJ.Close
        Exit Sub
    End If
    
    SQL = "INSERT INTO AM_PRODUK"
    SQL = SQL + "(KodeProduk"
    SQL = SQL + ",NamaProduk"
    SQL = SQL + ",idupdate"
    SQL = SQL + ",dateupdate"
    SQL = SQL + ",identry"
    SQL = SQL + ",Dateentry)"
    
    SQL = SQL + "VALUES"
    SQL = SQL + "('" & txtcatcode & "'"
    SQL = SQL + ", '" & txtdesc & "'"
    SQL = SQL + ", ' '"
    SQL = SQL + ", convert(datetime,'" & tanggalsekarang & "')"
    SQL = SQL + ", '" & kuser & "'"
    SQL = SQL + ", convert(datetime,'" & tanggalsekarang & "'))"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Data saved, click ok to continue  ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function

Private Sub cmdclear_Click()
    txtcatcode = ""
    txtdesc = ""
    txtcatcode.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='21' and b.kodeuser = '1" & kuser & "'"
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

Private Sub txtcatcode_GotFocus()
    Call Blok(txtcatcode)
End Sub

Private Sub txtcatcode_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtdesc.SetFocus
End Sub

Private Sub txtcatcode_LostFocus()
    If txtcatcode = "" Then Exit Sub
    If txtcatcode.SelLength <> 0 Then Exit Sub
    
    OBJ.Open dsn
    SQL = "SELECT * FROM AM_PRODUK WHERE KodeProduk = '" & txtcatcode & " '"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtdesc = RST!namaproduk
        MsgBox "Data already exist.", vbInformation, "Information"
        txtcatcode.SetFocus
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    txtdesc = ""
    txtdesc.SetFocus
End Sub
