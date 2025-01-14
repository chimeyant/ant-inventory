VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmbackup 
   Caption         =   "BACKUP DATA"
   ClientHeight    =   2025
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5010
   ControlBox      =   0   'False
   Icon            =   "frmbackup.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2025
   ScaleWidth      =   5010
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdbrowse2 
      Caption         =   "..."
      Height          =   360
      Left            =   4230
      TabIndex        =   9
      Top             =   510
      Width           =   600
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   930
      TabIndex        =   8
      Top             =   510
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Keterangan"
      Height          =   765
      Left            =   75
      TabIndex        =   6
      Top             =   1185
      Width           =   3465
      Begin VB.Label Label2 
         Caption         =   "Pastikan tidak ada komputer yang sedang melakukan transaksi  "
         Height          =   510
         Left            =   135
         TabIndex        =   7
         Top             =   210
         Width           =   2835
      End
   End
   Begin VB.CommandButton cmdproses 
      Caption         =   " Process"
      Height          =   330
      Left            =   3990
      TabIndex        =   5
      Top             =   1200
      Width           =   840
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      Height          =   330
      Left            =   3990
      TabIndex        =   4
      Top             =   1560
      Width           =   840
   End
   Begin VB.CommandButton cmdbrowse 
      Caption         =   "..."
      Height          =   360
      Left            =   4230
      TabIndex        =   3
      Top             =   135
      Width           =   600
   End
   Begin MSComDlg.CommonDialog cmndlg 
      Left            =   2940
      Top             =   1365
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBar1 
      Height          =   225
      Left            =   75
      TabIndex        =   2
      Top             =   915
      Width           =   4740
      _Version        =   851970
      _ExtentX        =   8361
      _ExtentY        =   397
      _StockProps     =   93
   End
   Begin VB.TextBox txtfile 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   930
      TabIndex        =   1
      Top             =   150
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Log Data"
      Height          =   225
      Left            =   60
      TabIndex        =   10
      Top             =   585
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Main Data"
      Height          =   225
      Left            =   60
      TabIndex        =   0
      Top             =   210
      Width           =   795
   End
End
Attribute VB_Name = "frmbackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private OBJ As New ADODB.Connection
Private OBJ1 As New ADODB.Connection
Private RS As ADODB.Recordset
Private SQL As String
Private fso As FileSystemObject
Private destination_maindata As String
Private source_maindata As String
Private destination_logdata As String
Private source_logdata As String

Private pesan As String

Private Sub cmdbrowse_Click()
    With cmndlg
        .CancelError = False
        .DialogTitle = "File Main Backup"
        .Filter = "SQL Server Database (*.mdf)|*.mdf"
        .ShowSave
        If .FileName <> "" Then
            destination_maindata = .FileName
        Else
            Exit Sub
        End If
    End With
End Sub

Private Sub cmdbrowse2_Click()
    With cmndlg
        .CancelError = False
        .DialogTitle = "File Main Backup"
        .Filter = "SQL Server Database (*.ldf)|*.ldf"
        .ShowSave
        If .FileName <> "" Then
            source_logdata = .FileName
        Else

            Exit Sub
        End If
    End With
End Sub

Private Sub cmdclose_Click()
     Unload Me
End Sub

Private Sub proseshapuspenjualan()
    On Error GoTo err_handler:
    OBJ.Open dsn
    SQL = "EXEC am_backup"
    OBJ.Execute (SQL)
    OBJ.Close
    pesan = "Proses backup penjualan...Ok"
    Exit Sub
err_handler:
    If OBJ.State = 1 Then OBJ.Close
    MsgBox Err.Description, vbCritical, AppName
End Sub

Private Sub proses_detachdata()
    On Error GoTo err_handler:
    Dim strConn As String
    strConn = "provider=SQLOLEDB.1;Password=" & dbPass & ";User ID=" & dbUser & ";Initial Catalog=master;Data Source=" & dbServer
    OBJ.Open strConn
    SQL = "EXEC sp_detach_db 'eitdb','true'"
    OBJ.Execute SQL
    OBJ.Close
    pesan = "Proses Detach....ok"
    Exit Sub
err_handler:
    If OBJ.State = 1 Then OBJ.Close
    MsgBox Err.Description, vbCritical, AppName
End Sub

Private Sub proseshapuspembelian()
    On Error GoTo err_handler:
    OBJ.Open dsn
    SQL = "EXEC am_resetpenjualan"
    OBJ.Execute SQL
    OBJ.Close
    Exit Sub
    pesan = "Proses backup pembelian...ok"
err_handler:
    If OBJ.State = 1 Then OBJ.Close
    MsgBox Err.Description, vbCritical, AppName
End Sub

Private Sub attach_database()
    On Error GoTo err_handler:
    Dim strConn As String
    strConn = "provider=SQLOLEDB.1;Password=" & dbPass & ";User ID=" & dbUser & ";Initial Catalog=master;Data Source=" & dbServer
    OBJ.Open strConn
    SQL = "CREATE DATABASE eitdb on (FILENAME='" & source_maindata & "'),(FILENAM='" & source_logdata & "') for ATTACH"
    OBJ.Execute SQL
    OBJ.Close
    pesan = "Proses Attach Database...Ok"
    Exit Sub
err_handler:
    If OBJ.State = 1 Then OBJ.Close
End Sub

Private Sub backup_data()
    Set fso = New FileSystemObject
    'MAIN DATA BACKUP
    fso.CopyFile source_maindata, destination_logdata, True
    'LOG DATA BACKUP
    fso.CopyFile source_logdata, destination_logdata, True
    pesan = "Proses Backup Selesai...!"
End Sub

