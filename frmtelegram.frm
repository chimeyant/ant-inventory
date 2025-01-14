VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmtelegram 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Id Telegram"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4245
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   4245
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Edit Bot"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox txtbot 
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
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   3855
   End
   Begin VB.TextBox txtname 
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
      Left            =   960
      TabIndex        =   1
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox txtid 
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
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin XtremeSuiteControls.PushButton cmdclose 
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   3000
      Width           =   945
      _Version        =   851970
      _ExtentX        =   1667
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Close"
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
   End
   Begin XtremeSuiteControls.PushButton cmdsave 
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   960
      Width           =   945
      _Version        =   851970
      _ExtentX        =   1667
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Save ID"
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
   End
   Begin XtremeSuiteControls.PushButton btnsave 
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   2400
      Width           =   945
      _Version        =   851970
      _ExtentX        =   1667
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Save BOT"
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
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "BOT Telegram"
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
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
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
      Top             =   240
      Width           =   615
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   4215
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   4215
   End
End
Attribute VB_Name = "frmtelegram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private Sub btnsave_Click()
    If txtbot = "" Then
        MsgBox "Column BOT is Empty", vbCritical, AppName
        Exit Sub
    End If
    OBJ.Open dsn
    SQL = "Update am_telid set bot= '" & txtbot & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "BOT is updated", vbInformation, AppName
    
End Sub

Private Sub Check1_Click()
    If Check1.Value = Unchecked Then
        txtbot.Enabled = False
        btnsave.Enabled = False
    Else
        txtbot.Enabled = True
        btnsave.Enabled = True
    End If
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdsave_Click()
    If txtid = "" Then
        MsgBox "Column Id is empty", vbCritical, AppName
        Exit Sub
    ElseIf txtname = "" Then
        MsgBox "Column name is empty", vbCritical, AppName
        Exit Sub
    ElseIf txtbot = "" Then
        MsgBox "Column BOT is empty", vbCritical, AppName
        Exit Sub
    End If
    
    'Simpan ke tabel am_telid
    OBJ.Open dsn
    SQL = "Insert into am_telid(id,name,bot) Values('" & txtid & "','" & txtname & "','" & txtbot & "')"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Data is Saved", vbInformation, AppName
    txtid = ""
    txtname = ""
End Sub

Private Sub Form_Load()

OBJ.Open dsn
SQL = "Select bot from am_telid"
Set RST = OBJ.Execute(SQL)

If RST.EOF Then
    txtbot.Enabled = True
Else
    txtbot = RST!Bot
    txtbot.Enabled = False
    btnsave.Enabled = False
End If
OBJ.Close
End Sub
