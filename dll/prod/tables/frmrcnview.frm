VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmrcnview 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View | Print | Delete"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3810
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   3810
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton cmdrcn 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   975
      _Version        =   851970
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "No. Rcn"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnview 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
      _Version        =   851970
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "View"
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
   Begin XtremeSuiteControls.PushButton Close 
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
      _Version        =   851970
      _ExtentX        =   1931
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
   Begin Crystal.CrystalReport crystal 
      Left            =   0
      Top             =   1680
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
   Begin XtremeSuiteControls.PushButton btnDelete 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   1095
      _Version        =   851970
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Delete"
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
   Begin VB.Label lbltgl1 
      Alignment       =   1  'Right Justify
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
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Kg"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3240
      TabIndex        =   5
      Top             =   1080
      Width           =   555
   End
   Begin VB.Label lblkg 
      Alignment       =   1  'Right Justify
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
      Left            =   2040
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblnorcn 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmrcnview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RS As ADODB.Recordset
Private OBJ As New ADODB.Connection
Private SQL As String

Private Sub btnDelete_Click()
    If lblnorcn = "" Then Exit Sub
    If MsgBox("Are you sure you want to delete this production plan ?", vbYesNo + vbQuestion, "Confirmation") = vbNo Then Exit Sub
    OBJ.Open dsn
    SQL = "DELETE From am_rcnprod Where KD_RCN='" & lblnorcn & "'"
    Set RS = OBJ.Execute(SQL)
    
    SQL = "DELETE From am_rcnbb Where Kd_RCN='" & lblnorcn & "'"
    Set RS = OBJ.Execute(SQL)
    
    SQL = "DELETE From am_rcnpack Where KD_RCN='" & lblnorcn & "'"
    Set RS = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Production Plan " & lblnorcn & " is successfuly delete", vbInformation, AppName
    lblnorcn = ""
    lbltgl1 = ""
    lblkg = ""
End Sub

Private Sub btnview_Click()
    crystal.Reset
    crystal.WindowState = crptMaximized
    crystal.WindowShowCloseBtn = True
    crystal.WindowShowPrintSetupBtn = True
    crystal.WindowShowSearchBtn = True
    crystal.Connect = dsnreport
    crystal.DataFiles(0) = "Proc(am_prodplan)"
    crystal.ReportFileName = AppPath & "\reports\produksi\tables\prod_plan.rpt"
    crystal.ParameterFields(0) = "@kode ;" + lblnorcn + ";true"
    crystal.ParameterFields(1) = "@namauser ;" + nmuser + ";true"
    crystal.RetrieveDataFiles
    crystal.Action = 1
End Sub

Private Sub Close_Click()
    Unload Me
End Sub

Private Sub cmdrcn_Click()
    carisql1 = "Select KD_RCN,TGL1,TGL2,SUM(TOTALKG)'Kg' From am_rcnprod"
    namatabel = "Rencana Produksi"
    frmsearch.Show vbModal
End Sub

Private Sub cmdrcn_GotFocus()
    lblnorcn = hasil
    lbltgl1 = "From : " & Format(hasil1, "dd-MM-YYYY") & "  To : " & Format(hasil2, "dd/MM/YYYY")
    lblkg = Format(hasil3, "##,###,###,##0.00")
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    hasil3 = ""
End Sub
