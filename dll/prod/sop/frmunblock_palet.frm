VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmunblock_palet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Key Palet"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5280
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtpalet 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1245
      TabIndex        =   1
      Top             =   3420
      Width           =   3975
   End
   Begin XtremeSuiteControls.PushButton btnClose 
      Height          =   450
      Left            =   90
      TabIndex        =   0
      Top             =   4395
      Width           =   5100
      _Version        =   851970
      _ExtentX        =   8996
      _ExtentY        =   794
      _StockProps     =   79
      Caption         =   "CLOSE"
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
   Begin XtremeSuiteControls.PushButton btnbuka_blokscan 
      Height          =   435
      Left            =   3870
      TabIndex        =   2
      Top             =   3945
      Width           =   1320
      _Version        =   851970
      _ExtentX        =   2328
      _ExtentY        =   767
      _StockProps     =   79
      Caption         =   "UnLock Palet"
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
   Begin XtremeSuiteControls.PushButton btnlock_palet 
      Height          =   435
      Left            =   90
      TabIndex        =   3
      Top             =   3945
      Width           =   1335
      _Version        =   851970
      _ExtentX        =   2355
      _ExtentY        =   767
      _StockProps     =   79
      Caption         =   "Lock Palet"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   2595
      Left            =   0
      TabIndex        =   4
      Top             =   345
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   4577
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   -2147483631
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.Label lblrecords 
      Alignment       =   1  'Right Justify
      Caption         =   "Records."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   75
      TabIndex        =   10
      Top             =   3000
      Width           =   5145
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   30
      X2              =   5235
      Y1              =   3315
      Y2              =   3330
   End
   Begin VB.Label Label5 
      Caption         =   "PALET :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   270
      TabIndex        =   9
      Top             =   3480
      Width           =   810
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LOCKED"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   3975
      TabIndex        =   8
      Top             =   60
      Width           =   1275
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SCAN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   2865
      TabIndex        =   7
      Top             =   60
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TGL. SCAN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   1920
      TabIndex        =   6
      Top             =   60
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PALET"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   0
      TabIndex        =   5
      Top             =   60
      Width           =   1905
   End
End
Attribute VB_Name = "frmunblock_palet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private OBJ As New ADODB.Connection
Private RST As ADODB.Recordset
Private SQL As String

Private Sub btnbuka_blokscan_Click()
    OBJ.Open dsn
    SQL = "Update list_mutasi_produksi_header set ref3='0' Where kode_palet ='" & txtpalet & "'"
    Set RST = OBJ.Execute(SQL)
    
    MsgBox "Palet is successfuly unlocked ", vbInformation, AppName
    txtpalet = ""
    
    OBJ.Close
    openlockpalet
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnlock_palet_Click()
    OBJ.Open dsn
    SQL = "Update list_mutasi_produksi_header set ref3='1' Where kode_palet ='" & txtpalet & "'"
    Set RST = OBJ.Execute(SQL)
    
    MsgBox "Palet is successfuly locked ", vbInformation, AppName
    txtpalet = ""
    
    OBJ.Close
    openlockpalet
End Sub

Sub openlockpalet()
    OBJ.Open dsn
    SQL = "Select kode_palet,tanggal,ref1,ref2 From list_mutasi_produksi_header Where ref3='1'"
    Set RST = OBJ.Execute(SQL)
    Set grid.DataSource = RST
    OBJ.Close
    setGrid
    hitungrow
End Sub

Private Sub Form_Load()
    openlockpalet
End Sub
Private Sub setGrid()
    With grid
        .Cols = 4
        .ColWidth(0) = 1900
        .ColWidth(1) = 950
        .ColWidth(2) = 1100
        .ColWidth(3) = 1100

        .ColAlignmentFixed(0) = flexAlignCenterCenter
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .Refresh
    End With
End Sub
Sub hitungrow()
    If grid.Rows = 0 Then lblrecords = "0 Records.": Exit Sub
    grid.Row = 0
    Do While True
        lblrecords = grid.Rows & " Records."
        'MsgBox grid.Rows & " " & grid.Row
        If grid.Rows = grid.Row + 1 Or grid.Rows = 0 Then Exit Do
        grid.Row = grid.Row + 1
    Loop
End Sub
Private Sub grid_Click()
    txtpalet = grid.TextMatrix(grid.Row, 0)
End Sub
