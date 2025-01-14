VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmgroup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filler And Packaging Group"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Left            =   3960
      Picture         =   "frmgroup.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   16
      Top             =   780
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
      Left            =   4215
      Picture         =   "frmgroup.frx":034E
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   15
      Top             =   780
      Visible         =   0   'False
      Width           =   255
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
      Left            =   4485
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   14
      Top             =   765
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtperson 
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
      Left            =   4875
      TabIndex        =   12
      Top             =   210
      Width           =   2250
   End
   Begin VB.ComboBox cmbgroup2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmgroup.frx":0630
      Left            =   5505
      List            =   "frmgroup.frx":0652
      TabIndex        =   8
      Text            =   "--Select Group--"
      Top             =   1455
      Width           =   1605
   End
   Begin VB.ComboBox cmbgroup1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmgroup.frx":06B1
      Left            =   1770
      List            =   "frmgroup.frx":06D3
      TabIndex        =   5
      Text            =   "--Select Group--"
      Top             =   1455
      Width           =   1605
   End
   Begin VB.TextBox txtlot 
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
      Left            =   1125
      TabIndex        =   2
      Top             =   285
      Width           =   2430
   End
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
      Height          =   360
      Left            =   1140
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   2430
   End
   Begin XtremeSuiteControls.PushButton btnClose 
      Height          =   465
      Left            =   6210
      TabIndex        =   0
      Top             =   5040
      Width           =   1050
      _Version        =   851970
      _ExtentX        =   1852
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   3090
      Left            =   90
      TabIndex        =   9
      Top             =   1830
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   5450
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid1 
      Height          =   3090
      Left            =   3810
      TabIndex        =   10
      Top             =   1830
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   5450
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
   Begin XtremeSuiteControls.PushButton btnsimpan 
      Height          =   375
      Left            =   4890
      TabIndex        =   13
      Top             =   660
      Width           =   1050
      _Version        =   851970
      _ExtentX        =   1852
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Save Person"
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
      Left            =   5100
      TabIndex        =   17
      Top             =   5040
      Width           =   1050
      _Version        =   851970
      _ExtentX        =   1852
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
   Begin XtremeSuiteControls.PushButton btnhapus 
      Height          =   375
      Left            =   5970
      TabIndex        =   18
      Top             =   660
      Width           =   1170
      _Version        =   851970
      _ExtentX        =   2064
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Delete Person"
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
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "+ Person"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3915
      TabIndex        =   11
      Top             =   255
      Width           =   915
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "PENGEMASAN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   300
      Left            =   4005
      TabIndex        =   7
      Top             =   1485
      Width           =   1140
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PENGISIAN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   300
      Left            =   300
      TabIndex        =   6
      Top             =   1500
      Width           =   1020
   End
   Begin VB.Label Label1 
      Caption         =   "Nomor Lot"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   195
      TabIndex        =   4
      Top             =   315
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Kode Palet"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   975
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000006&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   450
      Left            =   90
      Top             =   1395
      Width           =   3450
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000006&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   450
      Left            =   3810
      Top             =   1395
      Width           =   3450
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0FFFF&
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   1050
      Left            =   3855
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   3435
   End
End
Attribute VB_Name = "frmgroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private OBJ As New ADODB.Connection
Private RST As ADODB.Recordset
Private SQL As String

Private Sub btnclose_Click()
    Unload Me
End Sub

Private Sub btnhapus_Click()
    If txtperson = "" Then Exit Sub
    
    'CEK OPERATOR
    OBJ.Open dsn
    SQL = "Select * From produksi_opt Where operator= '" & txtperson & "' and status='P'"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    
    If RST.EOF Then
        OBJ.Close
        MsgBox "Person Name is not found", vbCritical, AppName
        Exit Sub
    End If
    
    SQL = "Delete from produksi_opt Where operator= '" & txtperson & "' and status='P'"
    Set RST = OBJ.Execute(SQL)
    
    MsgBox "Data is successfully deleted..", vbInformation, AppName
    OBJ.Close
    txtperson = ""
    hapusgrid
    initGrid
    opendata
End Sub

Private Sub btnsave_Click()
    If txtlot = "" Then Exit Sub
    If cmbgroup1.text = "--Select Group--" Or cmbgroup2.text = "--Select Group--" Then
        MsgBox "Pease select group first !", vbCritical, AppName
        Exit Sub
    End If
    'CEK GROUP PERSON
    OBJ.Open dsn
    SQL = "Select * From produksi_group Where nolot= '" & txtlot & "'"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    
    If Not RST.EOF Then
        OBJ.Close
        MsgBox "Nolot is Already exist.", vbCritical, "WARNING"
        txtlot = ""
        Exit Sub
    End If
    
    'SIMPAN GROUP PERSON
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        grid.Col = 0
        If grid.CellPicture = uncheck Then GoTo berikutnya:
        RST.AddNew
        RST!nolot = txtlot
        RST!palet = txtpalet
        RST!Group = cmbgroup1.text
        RST!personil = grid.TextMatrix(grid.Row, 1)
        RST!fillerdate = Format(Date, "yyyy/MM/dd")
        RST.Update
berikutnya:
        grid.Row = grid.Row + 1
    Loop

    SQL = "Select * From produksi_group Where nolot= '" & txtlot & "'"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    
    grid1.Row = 1
    Do While True
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
        grid1.Col = 0
        If grid1.CellPicture = uncheck Then GoTo berikutnya2:
        RST.AddNew
        RST!nolot = txtlot
        RST!palet = txtpalet
        RST!Group = cmbgroup2.text
        RST!personil = grid1.TextMatrix(grid1.Row, 1)
        RST!packdate = Format(Date, "yyyy/MM/dd")
        RST.Update
berikutnya2:
        grid1.Row = grid1.Row + 1
    Loop
    
    OBJ.Close
    MsgBox "Data is successfuly saved.", vbInformation, AppName
    txtpalet = ""
    txtlot = ""
    cmbgroup1.text = "--Select Group--"
    cmbgroup2.text = "--Select Group--"
    hapusgrid
    initGrid
    opendata
End Sub

Private Sub btnsimpan_Click()
    If txtperson = "" Then Exit Sub
    
    'CEK OPERATOR
    OBJ.Open dsn
    SQL = "Select * From produksi_opt Where operator='" & txtperson & "' and status='P'"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    
    If Not RST.EOF Then
        OBJ.Close
        MsgBox "Person Name is already exist.", vbCritical, "WARNING"
        txtopt = ""
        Exit Sub
    End If
    'SIMPAN OPERATOR
    RST.AddNew
    RST!operator = txtperson
    RST!Status = "P"
    RST.Update
    MsgBox "Data is successfully saved", vbInformation, AppName
    txtperson = ""
    OBJ.Close
    hapusgrid
    initGrid
    opendata
End Sub

Private Sub cmbgroup1_LostFocus()
    If cmbgroup1.text = "" Then cmbgroup1.text = "--Select Group--"
End Sub

Private Sub cmbgroup2_LostFocus()
    If cmbgroup2.text = "" Then cmbgroup2.text = "--Select Group--"
End Sub

Private Sub Form_Load()
    initGrid
    opendata
End Sub

Private Sub opendata()
    OBJ.Open dsn
    SQL = "Select * From produksi_opt Where status='P' order by operator asc"
    Set RST = OBJ.Execute(SQL)
    grid.Row = 1
    Do While Not RST.EOF
        grid.TextMatrix(grid.Row, 1) = RST!operator
        grid.Col = 0
        Set grid.CellPicture = uncheck
        setAlternatingGrid grid.Row
        grid.Rows = grid.Rows + 1
        grid.Row = grid.Row + 1
        RST.MoveNext
    Loop
    
    SQL = "Select * From produksi_opt Where status='P' order by operator asc"
    Set RST = OBJ.Execute(SQL)
    grid1.Row = 1
    Do While Not RST.EOF
        grid1.TextMatrix(grid1.Row, 1) = RST!operator
        grid1.Col = 0
        Set grid1.CellPicture = uncheck
        setAlternatingGrid1 grid1.Row
        grid1.Rows = grid1.Rows + 1
        grid1.Row = grid1.Row + 1
        RST.MoveNext
    Loop
    OBJ.Close
End Sub

Private Sub initGrid()
    With grid
        .Cols = 2
        .TextMatrix(0, 0) = "PILIH"
        .TextMatrix(0, 1) = "PERSONIL"
        
        .ColWidth(0) = 500
        .ColWidth(1) = 1500
    End With
    With grid1
        .Cols = 2
        .TextMatrix(0, 0) = "PILIH"
        .TextMatrix(0, 1) = "PERSONIL"
        
        .ColWidth(0) = 500
        .ColWidth(1) = 1500
    End With
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
    
    grid1.Row = 1
    Do While True
        If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Do
        grid1.TextMatrix(grid1.Row, 0) = ""
        grid1.TextMatrix(grid1.Row, 1) = ""

        grid1.Col = 0
        Set grid1.CellPicture = blank
        grid1.Row = grid1.Row + 1
    Loop
    grid1.Rows = 2
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    Select Case grid.Col
        Case 0:
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            If cmbgroup1.text = "--Select Group--" Then
                MsgBox "Pease select the filler group first !", vbExclamation, "PERHATIAN"
                Exit Sub
            End If
            If grid.CellPicture = uncheck Then
                Set grid.CellPicture = check
                setAlternatingGrid grid.Row
            Else
                Set grid.CellPicture = uncheck
                setAlternatingGrid grid.Row
            End If
    End Select
End Sub

Private Sub grid1_Click()
    If grid1.MouseRow = 0 Then Exit Sub
    Select Case grid1.Col
        Case 0:
            If cmbgroup1.text = "--Select Group--" Then
                MsgBox "Pease select the packaging group first !", vbExclamation, "PERHATIAN"
                Exit Sub
            End If
            If grid1.TextMatrix(grid1.Row, 1) = "" Then Exit Sub
            If grid1.CellPicture = uncheck Then
                Set grid1.CellPicture = check
                setAlternatingGrid1 grid1.Row
            Else
                Set grid1.CellPicture = uncheck
                setAlternatingGrid1 grid1.Row
            End If
    End Select
End Sub

Private Function setAlternatingGrid1(ByVal i As Integer)
    Dim j As Integer
    j = 0
    
    If grid1.CellPicture = check Then
        For j = 0 To grid1.Cols - 1
        grid1.Col = j
        grid1.CellBackColor = &HE0E0E0
        Next
    Else
        For j = 0 To grid1.Cols - 1
        grid1.Col = j
        grid1.CellBackColor = &HFFFFFF
        Next
    End If
End Function
Private Function setAlternatingGrid(ByVal i As Integer)
    Dim j As Integer
    j = 0
    
    If grid.CellPicture = check Then
        For j = 0 To grid.Cols - 1
        grid.Col = j
        grid.CellBackColor = &HE0E0E0
        Next
    Else
        For j = 0 To grid.Cols - 1
        grid.Col = j
        grid.CellBackColor = &HFFFFFF
        Next
    End If
End Function

Private Sub txtpalet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtlot = Mid(txtpalet, 3, 15)
    End If
End Sub
