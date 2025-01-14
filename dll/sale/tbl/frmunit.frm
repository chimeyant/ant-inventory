VERSION 5.00
Begin VB.Form frmunit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Satuan"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   405
      Left            =   3255
      TabIndex        =   8
      Top             =   1305
      Width           =   720
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      Height          =   405
      Left            =   2535
      TabIndex        =   7
      Top             =   1305
      Width           =   720
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "Save"
      Height          =   405
      Left            =   1755
      TabIndex        =   6
      Top             =   1305
      Width           =   780
   End
   Begin VB.TextBox txtinit 
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
      MaxLength       =   20
      TabIndex        =   2
      Top             =   960
      Width           =   2655
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
      MaxLength       =   20
      TabIndex        =   1
      Top             =   600
      Width           =   2655
   End
   Begin VB.TextBox txtunitcode 
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
      MaxLength       =   3
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Initial"
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
      TabIndex        =   3
      Top             =   990
      Width           =   855
   End
   Begin VB.Label lblsatcode 
      Caption         =   "Kode Satuan"
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
      TabIndex        =   4
      Top             =   270
      Width           =   975
   End
   Begin VB.Label lbldesc 
      Caption         =   "Nama Satuan"
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
      Width           =   975
   End
End
Attribute VB_Name = "frmunit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private Sub cmdclear_Click()
    txtunitcode = ""
    txtdesc = ""
    txtinit = ""
    txtunitcode.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdadd_Click()
    If Len(Trim(txtunitcode)) = 0 Then
        MsgBox "Data entry not Complete.", vbExclamation, "Warning"
        txtunitcode.SetFocus
        Exit Sub
    End If
    
    If txtdesc = "" Or txtunitcode = "" Or txtinit = "" Then
        MsgBox "Data entry not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    txtunitcode = Trim(txtunitcode)
    
    OBJ.Open dsn
    SQL = "SELECT * FROM AM_UNIT WHERE KodeSatuan = '" & txtunitcode & " '"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Data already exist.", vbInformation, "Information"
        cmdclear_Click
        OBJ.Close
        Exit Sub
    End If
    
    SQL = "INSERT INTO AM_UNIT"
    SQL = SQL + "(KodeSatuan"
    SQL = SQL + ",NamaSatuan"
    SQL = SQL + ",init"
    SQL = SQL + ",idupdate"
    SQL = SQL + ",dateupdate"
    SQL = SQL + ",identry"
    SQL = SQL + ",Dateentry)"
    
    SQL = SQL + "VALUES"
    SQL = SQL + " ('" & txtunitcode & "'"
    SQL = SQL + ", '" & txtdesc & "'"
    SQL = SQL + ", '" & txtinit & "'"
    SQL = SQL + ", ''"
    SQL = SQL + ", convert(datetime,'" & tanggalsekarang & "')"
    SQL = SQL + ", '" & kuser & "'"
    SQL = SQL + ", convert(datetime,'" & tanggalsekarang & "'))"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
        
    MsgBox "Data saved, click ok to continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='11' and b.kodeuser = '1" & kuser & "'"
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

Private Sub txtdesc_Change()
    txtinit = txtdesc
End Sub

Private Sub txtunitcode_GotFocus()
    Call Blok(txtunitcode)
End Sub

Private Sub txtunitcode_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtdesc.SetFocus
End Sub

Private Sub txtunitcode_LostFocus()
    If txtunitcode = "" Then Exit Sub
    If txtunitcode.SelLength <> 0 Then Exit Sub
    
    OBJ.Open dsn
    SQL = "SELECT * FROM AM_UNIT WHERE KodeSatuan = '" & txtunitcode & " '"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtdesc = RST!namasatuan
        txtinit = RST!init
                    
        MsgBox "Data already exist.", vbInformation, "Information"
        txtunitcode.SetFocus
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    txtdesc = ""
    txtinit = ""
    txtdesc.SetFocus
End Sub

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function
