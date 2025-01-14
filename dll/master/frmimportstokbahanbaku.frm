VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmimportstokbahanbaku 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Stok Bahan Bau"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5205
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cmndlg 
      Left            =   105
      Top             =   765
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeSuiteControls.PushButton cmdclose 
      Height          =   390
      Left            =   4245
      TabIndex        =   2
      Top             =   990
      Width           =   930
      _Version        =   851970
      _ExtentX        =   1640
      _ExtentY        =   688
      _StockProps     =   79
      Caption         =   "CLOSE"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ProgressBar pgr 
      Height          =   330
      Left            =   90
      TabIndex        =   0
      Top             =   390
      Width           =   5070
      _Version        =   851970
      _ExtentX        =   8943
      _ExtentY        =   582
      _StockProps     =   93
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton cmdimport 
      Height          =   390
      Left            =   3285
      TabIndex        =   3
      Top             =   990
      Width           =   930
      _Version        =   851970
      _ExtentX        =   1640
      _ExtentY        =   688
      _StockProps     =   79
      Caption         =   "IMPORT"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdbrowse 
      Height          =   390
      Left            =   2340
      TabIndex        =   4
      Top             =   990
      Width           =   930
      _Version        =   851970
      _ExtentX        =   1640
      _ExtentY        =   688
      _StockProps     =   79
      Caption         =   "BROWSE"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Import Progress"
      Height          =   210
      Left            =   105
      TabIndex        =   1
      Top             =   120
      Width           =   1530
   End
End
Attribute VB_Name = "frmimportstokbahanbaku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private OBJ1 As New ADODB.Connection
Private RST1 As ADODB.Recordset
Private OBJ2 As New ADODB.Connection
Private RST2 As ADODB.Recordset

Private flname As String

Private Sub cmdbrowse_Click()
    On Error GoTo err_msg
    
    
    With cmndlg
        .CancelError = False
        .DialogTitle = "File Import Bahan Baku"
        .Filter = "MS Execel 2003 (*.xls)|*.xls"
        .ShowOpen
        If .FileName <> "" Then
            flname = .FileName
        Else
            lblstatus = "Data Tidak Ditemukan"
            Exit Sub
        End If
    End With
    Exit Sub
err_msg:
    MsgBox "Gagal melakukan proses import...!" & Err.Description & ""
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdimport_Click()
    Dim dsnexcel As String
    Dim SQL1 As String
    Dim SQL2 As String
    Dim j As Integer
    
    dsnexcel = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & flname & "; Extended Properties=""Excel 8.0;HDR=YES;"""
    
    SQL1 = "select * from [STOKBB$]"
    OBJ1.Open dsnexcel
    Set RST1 = OBJ1.Execute(SQL1)
    Do While Not RST1.EOF
        j = j + 1
        RST1.MoveNext
        DoEvents
    Loop
    RST1.MoveFirst
    pgr.Max = j
    
    OpenDB
    SQL2 = "DELETE FROM am_invloc"
    ConSQL.Execute (SQL2)
    
    Do While Not RST1.EOF
        'IMPORT DATA KE TABLE STOK AWAL
        SQL2 = "Insert Into am_invloc values('"
        SQL2 = SQL2 + RST1!kode + "',"
        SQL2 = SQL2 & "convert(money,'" & RST1!opname & "'),"
        SQL2 = SQL2 & "convert(datetime,'1/1/1900'),'0')"
        
        ConSQL.Execute SQL2
        
        pgr.Value = pgr.Value + 1
        RST1.MoveNext
        DoEvents
    Loop
    CloseSQLDB
    OBJ1.Close
End Sub


