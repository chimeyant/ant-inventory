VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmaddreaktor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Reaktor"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   105
      TabIndex        =   0
      Top             =   135
      Width           =   5055
      Begin XtremeSuiteControls.PushButton btnClose 
         Height          =   360
         Left            =   3735
         TabIndex        =   9
         Top             =   1995
         Width           =   1050
         _Version        =   851970
         _ExtentX        =   1852
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Close"
         Appearance      =   6
      End
      Begin VB.TextBox txtketerangan 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   1230
         TabIndex        =   8
         Top             =   1215
         Width           =   3555
      End
      Begin VB.TextBox txtkapasitas 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1245
         TabIndex        =   6
         Top             =   870
         Width           =   1125
      End
      Begin VB.TextBox txtnmreaktor 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1245
         TabIndex        =   4
         Top             =   540
         Width           =   3555
      End
      Begin VB.TextBox txtkdreaktor 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1245
         TabIndex        =   2
         Top             =   210
         Width           =   1125
      End
      Begin XtremeSuiteControls.PushButton btnSimpan 
         Height          =   360
         Left            =   1455
         TabIndex        =   10
         Top             =   2010
         Width           =   1050
         _Version        =   851970
         _ExtentX        =   1852
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Simpan"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton btnclear 
         Height          =   360
         Left            =   2595
         TabIndex        =   11
         Top             =   2010
         Width           =   1050
         _Version        =   851970
         _ExtentX        =   1852
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Clear"
         Appearance      =   6
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         Height          =   225
         Left            =   105
         TabIndex        =   7
         Top             =   1245
         Width           =   1185
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Kapasitas"
         Height          =   225
         Left            =   90
         TabIndex        =   5
         Top             =   900
         Width           =   1185
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Reaktor"
         Height          =   225
         Left            =   120
         TabIndex        =   3
         Top             =   570
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Reaktor"
         Height          =   225
         Left            =   135
         TabIndex        =   1
         Top             =   240
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmaddreaktor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RST As ADODB.Recordset
Private SQL As String

Private Sub Open_MySQLDB()
   OpenSQLDB dbServer, dbName, dbUser, dbPass
End Sub

Private Sub btnclear_Click()
    txtkdreaktor = ""
    txtnmreaktor = ""
    txtkapasitas = ""
    txtketerangan = ""
    txtkdreaktor.SetFocus
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub


Private Sub btnSimpan_Click()
    If txtkdreaktor = "" Then
        txtkdreaktor.SetFocus
        Exit Sub
    End If
    
    SQL = "select kdreaktor from am_reaktor where kdreaktor='" + txtkdreaktor + "'"
    
    Open_MySQLDB
    Set RST = ConSQL.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Reaktor Telah ada...!", vbInformation, AppName
        CloseSQLDB
        btnclear_Click
        Exit Sub
    End If
    
    SQL = "Insert Into am_reaktor values('"
    SQL = SQL + txtkdreaktor + "','"
    SQL = SQL + txtnmreaktor + "',"
    SQL = SQL + txtkapasitas + ",'"
    SQL = SQL + txtketerangan + "')"
    
    ConSQL.Execute SQL
    btnclear_Click
End Sub


