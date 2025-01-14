VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LOGIN SYSTEM PT.SPARTA PRIMA"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4065
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   4065
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkremote 
      BackColor       =   &H00808080&
      Caption         =   "Remote Server"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   210
      TabIndex        =   8
      Top             =   1410
      Width           =   1410
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   5835
      Top             =   990
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   6555
      Top             =   765
   End
   Begin XtremeSuiteControls.PushButton cmdexit 
      Height          =   375
      Left            =   2955
      TabIndex        =   4
      Top             =   1380
      Width           =   945
      _Version        =   851970
      _ExtentX        =   1667
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Exit"
      BackColor       =   -2147483633
      Appearance      =   6
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   990
      PasswordChar    =   "*"
      TabIndex        =   2
      ToolTipText     =   "Silahkan Masukan Password / Kata Sandi Anda"
      Top             =   645
      Width           =   2850
   End
   Begin VB.TextBox txtUser 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1005
      TabIndex        =   1
      ToolTipText     =   "Silahkan Masukan Username Anda"
      Top             =   195
      Width           =   2835
   End
   Begin XtremeSuiteControls.PushButton cmdlogin 
      Height          =   375
      Left            =   1950
      TabIndex        =   5
      Top             =   1365
      Width           =   945
      _Version        =   851970
      _ExtentX        =   1667
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Login"
      BackColor       =   -2147483633
      Appearance      =   6
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBar1 
      Height          =   90
      Left            =   -15
      TabIndex        =   6
      Top             =   1140
      Visible         =   0   'False
      Width           =   5475
      _Version        =   851970
      _ExtentX        =   9657
      _ExtentY        =   159
      _StockProps     =   93
      BackColor       =   12632256
      Scrolling       =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
      FlatStyle       =   -1  'True
      BarColor        =   255
      MarqueeDelay    =   5
   End
   Begin VB.Label lbljam 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3120
      TabIndex        =   7
      Top             =   1845
      Width           =   840
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxxx"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   210
      TabIndex        =   9
      Top             =   1785
      Width           =   5790
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   420
      Left            =   945
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   2970
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   420
      Left            =   945
      Shape           =   4  'Rounded Rectangle
      Top             =   150
      Width           =   2970
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   885
      Left            =   -15
      Top             =   1200
      Width           =   4185
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password    :"
      Height          =   225
      Left            =   165
      TabIndex        =   3
      Top             =   705
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name  :"
      Height          =   225
      Left            =   90
      TabIndex        =   0
      Top             =   270
      Width           =   1095
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program Name   : Exclusive Inventory Technology System
'Alias          : EI-Tech System
'Copyright      : 2012
'Company        : SPARTA PRIMA
'Programmer     : U. Selamat Raharja & Chandra Kirana

Public Event Login(ByVal UserName, Password As String)

Private Sub cmdexit_Click()
    End
End Sub

Private Sub cmdlogin_Click()
    If chkremote = 1 Then
        DoEvents
        Label3 = "Please wait connecting to server......"
        
        DoEvents
        dbRemoteServer = "36.64.1.231"
        remoteserver = True
        LoadDatabaseProperty
        ProgressBar1.Visible = False
        Label3 = ""
    End If
    RaiseEvent Login(txtUser, txtPassword)
End Sub

Private Sub Form_Load()
    Label3 = "Waiting Response...!"
    
    DoEvents
End Sub

Private Sub Timer1_Timer()
    If Me.Caption = "ENTER YOUR USER AND PASSWORD..!" Then
        Me.Caption = "LOGIN SYSTEM PT.SPARTA PRIMA"
    Else
        Me.Caption = "ENTER YOUR USER AND PASSWORD..!"
    End If
    DoEvents
End Sub

Private Sub Timer2_Timer()
    lbljam = Time
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtPassword = "" Then
            MsgBox "Please Enter Your Password...!", vbInformation, AppName
            txtPassword.SetFocus
            Exit Sub
        End If
    cmdlogin_Click
    End If
End Sub
