VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmshow 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12735
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   12735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox blank 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5400
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   8
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
      Left            =   4920
      Picture         =   "frmshow.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   7
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
      Left            =   5160
      Picture         =   "frmshow.frx":03B6
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   255
      Left            =   9120
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   133169155
      CurrentDate     =   38767
   End
   Begin VB.TextBox txtstring 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Left            =   11280
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox cmbDK 
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
      Left            =   11760
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   11280
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   450
      Calculator      =   "frmshow.frx":076C
      Caption         =   "frmshow.frx":078C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmshow.frx":07F8
      Keys            =   "frmshow.frx":0816
      Spin            =   "frmshow.frx":0858
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,###,##0.00;(###,###,###,##0.00)"
      HighlightText   =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999999999
      MinValue        =   -999999999999999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   -65531
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   2778
      _Version        =   393216
      Cols            =   22
      FixedCols       =   0
      BackColorBkg    =   -2147483632
      AllowUserResizing=   1
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
      _Band(0).Cols   =   22
   End
   Begin XtremeSuiteControls.PushButton cmdclose 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   12495
      _Version        =   851970
      _ExtentX        =   22040
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Cancel"
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
End
Attribute VB_Name = "frmshow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim posrow As String

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cmbDK.additem "K"
    cmbDK.additem "D"
    
    grid.TextMatrix(0, 1) = "Tanggal"
    grid.TextMatrix(0, 2) = "kode"
    grid.TextMatrix(0, 3) = "No Transaksi"
    grid.TextMatrix(0, 5) = "Account"
    grid.TextMatrix(0, 6) = "Keterangan"
    grid.TextMatrix(0, 7) = "D/K"
    grid.TextMatrix(0, 8) = "Amount"
    grid.TextMatrix(0, 11) = "Status"
    
    grid.Rows = 2
    grid.ColWidth(0) = 0 '900
    grid.ColWidth(1) = 1300
    grid.ColWidth(2) = 600
    grid.ColWidth(3) = 1200
    grid.ColWidth(4) = 0 '900
    grid.ColWidth(5) = 1300
    grid.ColWidth(6) = 5100
    grid.ColWidth(7) = 500
    grid.ColWidth(8) = 1300
    grid.ColWidth(9) = 0
    grid.ColWidth(10) = 0 '900
    grid.ColWidth(11) = 800
    grid.ColWidth(12) = 0
    grid.ColWidth(13) = 0
    grid.ColWidth(14) = 0
    grid.ColWidth(15) = 0
    grid.ColWidth(16) = 0
    grid.ColWidth(17) = 0
    grid.ColWidth(18) = 0
    grid.ColWidth(19) = 0
    grid.ColWidth(20) = 0
    grid.ColWidth(21) = 300
    
    grid.ColAlignment(1) = flexAlignLeftCenter
    grid.ColAlignment(5) = flexAlignRightCenter
    grid.ColAlignment(8) = flexAlignRightCenter
End Sub

Private Sub grid_Click()
    posrow = grid.Row
    Select Case grid.Col
        Case 1
            If date1.Visible = True Then Exit Sub
            
            date1.Width = grid.ColWidth(grid.Col) - 20
            date1.Height = 290
            'If grid.TextMatrix(grid.Row, grid.Col) <> "" Then date1 = grid.TextMatrix(grid.Row, 3)
            date1.Left = grid.Left + grid.CellLeft - 10
            date1.Top = grid.Top + grid.CellTop - 20
            date1.Visible = True
            date1 = Date
            date1.SetFocus
        Case 5
            carisql1 = "Select noac,nmac From gl_masterac"
            namatabel = "Account"
            frmsearch.Show vbModal
        Case 6
            If txtstring.Visible = True Then Exit Sub
            
            txtstring.Width = grid.ColWidth(grid.Col) - 40
            txtstring = grid.TextMatrix(grid.Row, grid.Col)
            txtstring.Left = grid.Left + grid.CellLeft
            txtstring.Top = grid.Top + grid.CellTop - 30
            txtstring.Visible = True
            txtstring.SetFocus
            If grid.Col = 6 Then txtstring = frmFArepair.grid.TextMatrix(1, 6)
        Case 7
            If cmbDK.Visible = True Then Exit Sub
            
            cmbDK.Width = grid.ColWidth(grid.Col) - 40
            cmbDK = grid.TextMatrix(grid.Row, grid.Col)
            cmbDK.Left = grid.Left + grid.CellLeft
            cmbDK.Top = grid.Top + grid.CellTop - 30
            cmbDK.Visible = True
            cmbDK.SetFocus
        Case 8
            txtnilai.Width = grid.ColWidth(grid.Col) - 40
            txtnilai = grid.TextMatrix(grid.Row, grid.Col)
            txtnilai.Left = grid.Left + grid.CellLeft
            txtnilai.Top = grid.Top + grid.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
        Case 21
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            If grid.CellPicture = uncheck Then
                Set grid.CellPicture = check
                If MsgBox("Save That Row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                    Set grid.CellPicture = blank
                    hapusgrid
                    Exit Sub
                End If
                Set grid.CellPicture = uncheck
            End If
    End Select
End Sub

Private Sub date1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then date1.Visible = False
    
    If KeyCode = 13 Then
        grid.TextMatrix(posrow, 1) = Format(date1, "dd/MM/yyyy")
        grid.TextMatrix(posrow, 0) = "01"
        grid.TextMatrix(posrow, 2) = "JJ"
        grid.TextMatrix(posrow, 3) = frmFArepair.txtkdaktiva
        grid.TextMatrix(posrow, 4) = frmFArepair.grid.TextMatrix(1, 4)
        grid.TextMatrix(posrow, 10) = frmFArepair.grid.TextMatrix(1, 10)
        grid.TextMatrix(posrow, 11) = frmFArepair.grid.TextMatrix(1, 11)
        grid.TextMatrix(posrow, 12) = frmFArepair.grid.TextMatrix(1, 12)
        grid.TextMatrix(posrow, 13) = frmFArepair.grid.TextMatrix(1, 13)
        grid.TextMatrix(posrow, 14) = grid.Row
        grid.TextMatrix(posrow, 15) = frmFArepair.grid.TextMatrix(1, 15)
        grid.TextMatrix(posrow, 16) = frmFArepair.grid.TextMatrix(1, 16)
        grid.TextMatrix(posrow, 17) = frmFArepair.grid.TextMatrix(1, 17)
        grid.TextMatrix(posrow, 18) = frmFArepair.grid.TextMatrix(1, 18)
        grid.TextMatrix(posrow, 19) = frmFArepair.grid.TextMatrix(1, 19)
        grid.TextMatrix(posrow, 20) = frmFArepair.grid.TextMatrix(1, 20)
        grid.Col = 21
        Set grid.CellPicture = uncheck
        
        grid.SetFocus
        grid.Row = posrow
        date1.Visible = False
        grid.Rows = grid.Rows + 1
        grid.Row = grid.Row + 1
    End If
End Sub

Private Sub grid_GotFocus()
    posrow = grid.Row
    Select Case grid.Col
        Case 5
            If hasil = "" Then Exit Sub
            grid.TextMatrix(grid.Row, 5) = hasil
            hasil = ""
            hasil1 = ""
    End Select
End Sub

Private Sub txtstring_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txtstring.Visible = False
    
    If KeyCode = 13 Then
        Select Case grid.Col
            Case 5: grid.TextMatrix(posrow, 5) = txtstring
            Case 6: grid.TextMatrix(posrow, 6) = txtstring
        End Select
        grid.SetFocus
        grid.Row = posrow
        txtstring.Visible = False
    End If
End Sub

Private Sub cmbDK_Click()
    grid.Row = posrow
    
    grid.SetFocus
    grid.TextMatrix(grid.Row, 7) = cmbDK
    cmbDK.Visible = False
End Sub

Private Sub cmbDK_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then cmbDK.Visible = False
End Sub

Private Sub txtnilai_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then txtnilai.Visible = False
    
    If KeyCode = 13 Then
        grid.TextMatrix(posrow, 8) = Format(txtnilai, "#,##0.00")
        grid.TextMatrix(posrow, 9) = Format(txtnilai, "#,##0.00")
        
        grid.SetFocus
        grid.Row = posrow
        txtnilai.Visible = False
    End If
End Sub
Private Sub hapusgrid()
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        grid.TextMatrix(grid.Row, 0) = ""
        grid.TextMatrix(grid.Row, 1) = ""
        grid.TextMatrix(grid.Row, 2) = ""
        grid.TextMatrix(grid.Row, 3) = ""
        grid.TextMatrix(grid.Row, 4) = ""
        grid.TextMatrix(grid.Row, 5) = ""
        grid.TextMatrix(grid.Row, 6) = ""
        grid.TextMatrix(grid.Row, 7) = ""
        grid.TextMatrix(grid.Row, 8) = ""
        grid.TextMatrix(grid.Row, 9) = ""
        grid.TextMatrix(grid.Row, 10) = ""
        grid.TextMatrix(grid.Row, 11) = ""
        grid.TextMatrix(grid.Row, 12) = ""
        grid.TextMatrix(grid.Row, 13) = ""
        grid.TextMatrix(grid.Row, 14) = ""
        grid.TextMatrix(grid.Row, 15) = ""
        grid.TextMatrix(grid.Row, 16) = ""
        grid.TextMatrix(grid.Row, 17) = ""
        grid.TextMatrix(grid.Row, 18) = ""
        grid.TextMatrix(grid.Row, 19) = ""
        grid.TextMatrix(grid.Row, 20) = ""
        grid.TextMatrix(grid.Row, 21) = ""
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
End Sub
