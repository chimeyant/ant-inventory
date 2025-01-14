VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmRfID_opt 
   Caption         =   "Operator"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtopt 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1050
      TabIndex        =   4
      Top             =   60
      Width           =   2430
   End
   Begin VB.PictureBox blank 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   615
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox check 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   255
      Picture         =   "frmRfID_opt.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox uncheck 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      Picture         =   "frmRfID_opt.frx":02E2
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   2865
      Left            =   45
      TabIndex        =   0
      Top             =   540
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   5054
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   8421504
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin XtremeSuiteControls.PushButton btnsave 
      Height          =   435
      Left            =   3615
      TabIndex        =   6
      Top             =   30
      Width           =   1020
      _Version        =   851970
      _ExtentX        =   1799
      _ExtentY        =   767
      _StockProps     =   79
      Caption         =   "Save"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   5
   End
   Begin XtremeSuiteControls.PushButton btnClose 
      Height          =   435
      Left            =   45
      TabIndex        =   7
      Top             =   3420
      Width           =   4590
      _Version        =   851970
      _ExtentX        =   8096
      _ExtentY        =   767
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
      Appearance      =   5
   End
   Begin VB.Label Label1 
      Caption         =   "+ Operator"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   135
      TabIndex        =   5
      Top             =   135
      Width           =   915
   End
End
Attribute VB_Name = "frmRfID_opt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private OBJ As New ADODB.Connection
Private RST As ADODB.Recordset
Private SQL As String

Private Sub btnclose_Click()
    Unload Me
End Sub

Private Sub btnsave_Click()
    If txtopt = "" Then Exit Sub
    
    'CEK OPERATOR
    OBJ.Open dsn
    SQL = "Select * From produksi_opt Where operator='" & txtopt & "'"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    
    If Not RST.EOF Then
        MsgBox "Operator is already exist.", vbCritical, "WARNING"
        txtopt = ""
        Exit Sub
    End If
    'SIMPAN OPERATOR
    RST.AddNew
    RST!operator = txtopt
    RST!Status = "O"
    RST.Update
    MsgBox "Data is successfully saved", vbInformation, AppName
    txtopt = ""
    OBJ.Close
    hapusgrid
    initGrid
    opendata
End Sub

Private Sub Form_Load()
    initGrid
    opendata
End Sub

Private Sub opendata()
    OBJ.Open dsn
    SQL = "Select * From produksi_opt order by operator asc"
    Set RST = OBJ.Execute(SQL)
    
    grid.Row = 1
    Do While Not RST.EOF
        grid.TextMatrix(grid.Row, 1) = RST!operator
        grid.Col = 0
        Set grid.CellPicture = uncheck
        grid.Rows = grid.Rows + 1
        grid.Row = grid.Row + 1
        RST.MoveNext
    Loop
    OBJ.Close
End Sub

Private Sub initGrid()
    With grid
        .Cols = 2
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Operator"
        
        .ColWidth(0) = 500
        .ColWidth(1) = 3500
    End With
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    Select Case grid.Col
        Case 0:
            Set grid.CellPicture = check
            If MsgBox("Delete that row..", vbQuestion + vbYesNo, "Konfimasi") = vbYes Then
                OBJ.Open dsn
                SQL = "Delete From produksi_opt Where operator = '" & grid.TextMatrix(grid.Row, 1) & "'"
                OBJ.Execute SQL
                Set grid.CellPicture = uncheck
                hapusrow
                MsgBox "Data successfully deleted.", vbInformation, AppName
                OBJ.Close
            Else
                Set grid.CellPicture = uncheck
            End If
            
        Case 1:
            frmRfID.txtopt = Me.grid.TextMatrix(grid.Row, 1)
            Unload Me
    End Select
End Sub

Private Sub hapusrow()
    grid.TextMatrix(grid.Row, 1) = ""
    Do While True
        If grid.TextMatrix(grid.Row + 1, 1) = "" Then
            grid.TextMatrix(grid.Row, 1) = ""
            Exit Do
        End If
        grid.TextMatrix(grid.Row, 1) = grid.TextMatrix(grid.Row + 1, 1)
        grid.Row = grid.Row + 1
    Loop
    grid.Col = 0
    Set grid.CellPicture = blank
    grid.Rows = grid.Rows - 1
    
End Sub

Private Sub hapusgrid()
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        grid.TextMatrix(grid.Row, 0) = ""
        grid.TextMatrix(grid.Row, 1) = ""

        grid.Col = 0
        Set grid.CellPicture = blank
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
End Sub
