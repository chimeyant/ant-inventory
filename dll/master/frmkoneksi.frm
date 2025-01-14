VERSION 5.00
Begin VB.Form frmkoneksi 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " SQL Server Koneksi"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4725
   Icon            =   "frmkoneksi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Server Properties"
      Height          =   2085
      Left            =   45
      TabIndex        =   0
      Top             =   150
      Width           =   4635
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test"
         Height          =   345
         Left            =   1575
         TabIndex        =   11
         Top             =   1650
         Width           =   960
      End
      Begin VB.CommandButton cmdBuild 
         Caption         =   "Build"
         Height          =   345
         Left            =   2550
         TabIndex        =   10
         Top             =   1650
         Width           =   960
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1230
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   1275
         Width           =   3255
      End
      Begin VB.TextBox txtUser 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1230
         TabIndex        =   8
         Top             =   930
         Width           =   3255
      End
      Begin VB.TextBox txtDatabase 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1230
         TabIndex        =   7
         Top             =   585
         Width           =   3255
      End
      Begin VB.TextBox txtServer 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1230
         TabIndex        =   3
         Top             =   240
         Width           =   3255
      End
      Begin VB.CommandButton cmdSelesai 
         Caption         =   "Selesai"
         Height          =   345
         Left            =   3540
         TabIndex        =   1
         Top             =   1650
         Width           =   960
      End
      Begin VB.Label Label4 
         Caption         =   "Database"
         Height          =   225
         Left            =   120
         TabIndex        =   6
         Top             =   660
         Width           =   1275
      End
      Begin VB.Label Label3 
         Caption         =   "Password"
         Height          =   225
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   1275
      End
      Begin VB.Label Label2 
         Caption         =   "User"
         Height          =   225
         Left            =   135
         TabIndex        =   4
         Top             =   975
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Server"
         Height          =   225
         Left            =   135
         TabIndex        =   2
         Top             =   315
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmkoneksi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuild_Click()
    Dim StrDbProperty As String
    Dim i As Integer
    
    For i = 1 To 4 Step 1
        If i = 1 Then
            StrDbProperty = Cheap_Encrypt(txtServer.text)
        End If
        If i = 2 Then
            StrDbProperty = StrDbProperty + "|" + Cheap_Encrypt(txtDatabase.text)
        End If
        If i = 3 Then
            StrDbProperty = StrDbProperty + "|" + Cheap_Encrypt(txtUser.text)
        End If
        If i = 4 Then
            StrDbProperty = StrDbProperty + "|" + Cheap_Encrypt(txtPassword.text)
        End If
    Next i
    
    Open AppPath + "\sqlserver.dll" For Output As #1
        Print #1, StrDbProperty
        MsgBox "Konfigurasi telah berhasil dibuat..!", vbInformation, AppName
    Close
End Sub

Private Sub cmdSelesai_Click()
    MsgBox "Silahkan Anda Restart Program Terlebih Dahulu...!", vbInformation, AppName
    Unload Me
End Sub

Private Sub DoLoadFileSetting()
    Dim StrKoneksi As String
    Dim ArrKoneksi() As String
    Dim i As Integer
    
    Open AppPath + "\sqlserver.dll" For Input As #1
        Line Input #1, StrKoneksi
    Close
    
    ArrKoneksi = Split(StrKoneksi, "|")
    
    For i = 1 To 4 Step 1
        If i = 1 Then
            txtServer.text = Cheap_Decrypt(ArrKoneksi(0))
        End If
        If i = 2 Then
            txtDatabase.text = Cheap_Decrypt(ArrKoneksi(1))
        End If
        If i = 3 Then
            txtUser.text = Cheap_Decrypt(ArrKoneksi(2))
        End If
        If i = 4 Then
            txtPassword.text = Cheap_Decrypt(ArrKoneksi(3))
        End If
    Next i
End Sub


Private Sub cmdTest_Click()
    OpenSQLDB txtServer.text, txtDatabase.text, txtUser.text, txtPassword.text
    If ConSQL.State = 1 Then
        MsgBox "Koneksi berhasil...!!", vbInformation, AppName
        CloseSQLDB
    End If
End Sub

Private Sub Form_Load()
    DoLoadFileSetting
End Sub
