VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmseri 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Define No Seri Faktur Pajak"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
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
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   89325571
      CurrentDate     =   37426
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
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
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   89325571
      CurrentDate     =   37426
   End
   Begin TDBText6Ctl.TDBText txtket 
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   4680
      Visible         =   0   'False
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   397
      Caption         =   "frmseri.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmseri.frx":006C
      Key             =   "frmseri.frx":008A
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   0
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   0
      BorderStyle     =   0
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   -1
      Format          =   ""
      FormatMode      =   1
      AutoConvert     =   -1
      ErrorBeep       =   0
      MaxLength       =   50
      LengthAsByte    =   0
      Text            =   ""
      Furigana        =   0
      HighlightText   =   -1
      IMEMode         =   0
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   3735
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   6588
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
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
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   6480
      TabIndex        =   5
      Top             =   4680
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmseri.frx":00C6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton chameleonButton1 
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   4680
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Submit"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmseri.frx":03E0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lbltype 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "From                                                  to"
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
      Left            =   360
      TabIndex        =   6
      Top             =   270
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   -120
      TabIndex        =   8
      Top             =   0
      Width           =   9615
   End
End
Attribute VB_Name = "frmseri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim str2, str3, str4, posrow As String
Dim boo1 As Boolean

Private Sub caristock()
    If date1 > date2 Then
        MsgBox "Invalid date range, search abort.", vbExclamation, "Error"
        Exit Sub
    End If
    
    If Year(date1) <> Year(date2) Then
        MsgBox "Pencarian gagal, range tanggal harus pada tahun yang sama.", vbExclamation, "Error"
        Exit Sub
    End If
    
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        grid.TextMatrix(grid.Row, 0) = ""
        grid.TextMatrix(grid.Row, 1) = ""
        grid.TextMatrix(grid.Row, 2) = ""
        grid.TextMatrix(grid.Row, 3) = ""
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
    grid.ColWidth(0) = 1500
    grid.ColWidth(1) = 1500
    grid.ColWidth(2) = 3400
    grid.ColWidth(3) = 750
    
    OBJ.Open dsn
    SQL = "select a.nobkt,a.tglbkt,a.kodecust,a.noseri,b.namacust from am_invhdr a left join am_customer b on a.kodecust=b.kodecust where a.type = '" & lbltype & "' and b.nonpwp <> '' and a.ppn <> 0 and a.tglbkt >= '" & batas1 & "' and a.tglbkt <= '" & batas2 & "' order by a.tglbkt,a.nobkt"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        Do While Not RST.EOF
            grid.Col = 1
            grid.CellAlignment = 1
            grid.TextMatrix(grid.Row, 0) = RST!nobkt
            grid.TextMatrix(grid.Row, 1) = Format(RST!tglbkt, "MMM dd yyyy")
            grid.TextMatrix(grid.Row, 2) = RST!kodecust & " - " & RST!namacust
            grid.TextMatrix(grid.Row, 3) = RST!noseri
            
            grid.Rows = grid.Rows + 1
            grid.Row = grid.Row + 1
            
            RST.MoveNext
        Loop
    Else
        MsgBox "No item found.", vbInformation, "Information"
    End If
    OBJ.Close
End Sub

Function batas1()
    batas1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function batas2()
    batas2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function

Private Sub chameleonButton1_Click()
    caristock
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='184' and b.kodeuser = '1" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If RST.EOF Then
    '        MsgBox "User Rights Denied !!" & vbCrLf & _
    '        "Please contact your Administrator.", vbCritical, "User Rights"
    '        OBJ.Close
    '        Unload Me
    '        Exit Sub
    '    End If
    '    OBJ.Close
    'End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
  
    grid.TextMatrix(0, 0) = "No. Bukti"
    grid.TextMatrix(0, 1) = "Tanggal"
    grid.TextMatrix(0, 2) = "Customer"
    grid.TextMatrix(0, 3) = "No. Seri"
    grid.ColWidth(0) = 1500
    grid.ColWidth(1) = 1500
    grid.ColWidth(2) = 3400
    grid.ColWidth(3) = 750
    
    grid.RowHeightMin = 300
    
    date1 = Date
    date2 = Date
    
    lbltype = "I"
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    posrow = grid.Row
    Select Case grid.Col
        Case 0, 1, 2
            txtket.Visible = False
        Case 3
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            If grid.Rows - 1 = 50 Then
                MsgBox "Maximum line 50", vbExclamation, "Warning"
                Exit Sub
            End If
            
            If txtket.Visible = True Then Exit Sub
            
            txtket.Width = grid.ColWidth(grid.Col) - 40
            txtket = grid.TextMatrix(grid.Row, grid.Col)
            txtket.Left = grid.Left + grid.CellLeft
            txtket.Top = grid.Top + grid.CellTop + 20
            txtket.Visible = True
            txtket.SetFocus
    End Select
End Sub

Private Sub grid_Scroll()
    txtket = ""
    txtket.Visible = False
End Sub

Private Sub txtket_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(txtket)) <> 8 And txtket <> "" Then
            MsgBox "No.Seri length must 8 digit.", vbCritical, "Information"
            txtket = ""
            txtket.Visible = False
            Exit Sub
        End If
        Select Case grid.Col
            Case 3
                grid.Row = posrow
                boo1 = True
                
                OBJ.Open dsn
                SQL = "select noseri,nobkt from am_invhdr where noseri = '" & txtket & "' and year(tglbkt) = '" & Format(date1, "yyyy") & "'"
                Set RST = OBJ.Execute(SQL)
                If Not RST.EOF Then
                    If txtket = "" Then
                        If MsgBox("Empty no seri ?", vbQuestion + vbYesNo, "Question") = vbNo Then
                            txtket = ""
                            txtket.Visible = False
                            OBJ.Close
                            Exit Sub
                        End If
                        
                        SQL = "update am_invhdr set noseri = '' where type = '" & lbltype & "' and nobkt = '" & grid.TextMatrix(grid.Row, 0) & "'"
                        Set RST = OBJ.Execute(SQL)
                    Else
                        If RST!nobkt <> grid.TextMatrix(grid.Row, 0) Then boo1 = False
                    End If
                Else
                    SQL = "select type from am_invhdr where type = '" & lbltype & "' and nobkt = '" & grid.TextMatrix(grid.Row, 0) & "' and idupdate = 'print'"
                    Set RST = OBJ.Execute(SQL)
                    If RST.EOF Then
                        SQL = "update am_invhdr set noseri = '" & txtket & "' where type = '" & lbltype & "' and nobkt = '" & grid.TextMatrix(grid.Row, 0) & "'"
                        Set RST = OBJ.Execute(SQL)
                    Else
                        boo1 = False
                    End If
                End If
                OBJ.Close
                
                If boo1 = False Then
                    MsgBox "Can not update cause already printed or No.Seri already exist.", vbCritical, "Information"
                    
                    txtket = ""
                    txtket.Visible = False
                    
                    Exit Sub
                End If
                
                grid.SetFocus
                grid.Col = 3
                grid.CellAlignment = 1
                grid.TextMatrix(grid.Row, 3) = txtket
                txtket = ""
                txtket.Visible = False
                
                grid.Col = 3
        End Select
    ElseIf KeyAscii = 27 Then
        txtket.Visible = False
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtket_LostFocus()
    txtket = ""
    txtket.Visible = False
End Sub
