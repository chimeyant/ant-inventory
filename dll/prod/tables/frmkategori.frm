VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmkategori 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kategori"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4755
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   4755
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtinisial 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2040
      TabIndex        =   7
      Top             =   3960
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.TextBox txtktgori 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   3495
   End
   Begin XtremeSuiteControls.PushButton btnClose 
      Height          =   465
      Left            =   3720
      TabIndex        =   4
      Top             =   4560
      Width           =   930
      _Version        =   851970
      _ExtentX        =   1640
      _ExtentY        =   820
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
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnSave 
      Height          =   465
      Left            =   2760
      TabIndex        =   5
      Top             =   4560
      Width           =   930
      _Version        =   851970
      _ExtentX        =   1640
      _ExtentY        =   820
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
      UseVisualStyle  =   -1  'True
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   3420
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   6033
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   -2147483642
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      Caption         =   "Kategori"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "Kode"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   795
   End
   Begin VB.Label lblkode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   1080
   End
End
Attribute VB_Name = "frmkategori"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private OBJ As New ADODB.Connection
Private RST As ADODB.Recordset
Private SQL As String


Function getkode() As String    '001'
    On Error GoTo Err_handler:
    Dim SQL As String
    Dim tempkode As String
    Dim kode As Long
    
    Set OBJ = New ADODB.Connection
    OBJ.Open dsn
    SQL = "select max(kdkategori)as ktgori from am_kategori"
    Set RST = OBJ.Execute(SQL)

    If IsNull(RST!ktgori) = True Or RST!ktgori = "" Then
        getkode = "001"
    Else
        kode = CLng(Mid(RST!ktgori, 1, 3)) + 1

        If (Len(Trim(Str(kode))) = 1) Then
            tempkode = "00" + Trim(Str(kode))
        End If
        If (Len(Trim(Str(kode))) = 2) Then
            tempkode = "0" + Trim(Str(kode))
        End If
        If (Len(Trim(Str(kode))) = 3) Then
            tempkode = Trim(Str(kode))
        End If
        getkode = tempkode
    End If
    OBJ.Close
    Exit Function
Err_handler:
    getkode = strnumber + "001"
End Function

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnSave_Click()
    If lblkode = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "Select * From am_kategori Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    With RST
        .AddNew
        !kdkategori = lblkode
        !kategori = txtktgori
        .Update
    End With
    OBJ.Close
    opendata
    lblkode = getkode
    txtktgori = ""
End Sub

Private Sub Form_Load()
    lblkode = getkode
    setGrid
    opendata
End Sub

Private Sub opendata()
    hapusgrid
    OBJ.Open dsn
    SQL = "Select * From am_kategori order by kdkategori asc"
    Set RST = OBJ.Execute(SQL)
    grid.Row = 1
    Do While Not RST.EOF
        grid.TextMatrix(grid.Row, 0) = RST!kdkategori
        grid.TextMatrix(grid.Row, 1) = RST!kategori
        grid.Rows = grid.Rows + 1
        grid.Row = grid.Row + 1
        RST.MoveNext
    Loop
    OBJ.Close
End Sub
Private Sub hapusgrid()
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 0) = "" Then Exit Do
        grid.TextMatrix(grid.Row, 0) = ""
        grid.TextMatrix(grid.Row, 1) = ""
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
    setGrid
End Sub

Private Sub setGrid()
    With grid
        .ColWidth(0) = 400
        .ColWidth(1) = 3500
    End With
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    Select Case grid.Col
        Case 1:
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            txtinisial.Width = grid.ColWidth(grid.Col) - 40
            txtinisial = grid.TextMatrix(grid.Row, grid.Col)
            txtinisial.Left = grid.Left + grid.CellLeft
            txtinisial.Top = grid.Top + grid.CellTop
            txtinisial.Visible = True
            txtinisial.SetFocus
    End Select
End Sub
Private Sub txtinisial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtinisial = "" Then
            MsgBox "Kategori tidak boleh kosong", vbCritical, AppName
        Exit Sub
        End If
        'Update Kategori
        If MsgBox("Anda yakin ingin mengubah kategori ini", vbQuestion + vbYesNo, AppName) = vbYes Then
            OBJ.Open dsn
            SQL = "Update am_kategori set kategori='" & txtinisial & "' Where kdkategori='" & grid.TextMatrix(grid.Row, 0) & "'"
            Set RST = OBJ.Execute(SQL)
            OBJ.Close
            grid.TextMatrix(grid.Row, 1) = txtinisial
        End If
        grid.SetFocus
    End If
End Sub

Private Sub txtinisial_LostFocus()
    txtinisial = ""
    txtinisial.Visible = False
End Sub
