VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0A1435CB-EB1C-11D4-89B0-204C4F4F5020}#3.0#0"; "akprogressbar.ocx"
Begin VB.Form frmunconfirm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Unconfirm"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11415
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   5760
      Picture         =   "frmunconfirm.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   13
      Top             =   720
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
      Left            =   5520
      Picture         =   "frmunconfirm.frx":034E
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   12
      Top             =   720
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
      Left            =   6000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   11
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   10320
      TabIndex        =   7
      Top             =   5760
      Width           =   855
      _ExtentX        =   1508
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
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmunconfirm.frx":0630
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdclear 
      Height          =   375
      Left            =   9360
      TabIndex        =   6
      Top             =   5760
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Clear"
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
      MICON           =   "frmunconfirm.frx":094A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdadd 
      Height          =   375
      Left            =   8400
      TabIndex        =   5
      Top             =   5760
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "UnConfirm"
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
      MICON           =   "frmunconfirm.frx":0C64
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   480
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
      Format          =   135004163
      CurrentDate     =   37694
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   2880
      TabIndex        =   2
      Top             =   480
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
      Format          =   135004163
      CurrentDate     =   37694
   End
   Begin Chameleon.chameleonButton cmdswitch 
      Height          =   285
      Left            =   4800
      TabIndex        =   3
      ToolTipText     =   "Search by ..."
      Top             =   480
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
      BTYPE           =   4
      TX              =   ""
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmunconfirm.frx":0F7E
      PICN            =   "frmunconfirm.frx":1298
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   4815
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   8493
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
      _Band(0).Cols   =   4
   End
   Begin akProgress.akProgressBar ak 
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   5760
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   661
      BackColour      =   -2147483633
      FontColour      =   16512
      BarColour       =   8421631
      Horizontal      =   -1  'True
      ReverseGradient =   0   'False
      Max             =   100
      Min             =   0
      GapWidth        =   0
      LineWidth       =   1
      Caption         =   0
      BorderStyle     =   0
      Margin          =   2
      Gradient        =   0
      Alignment       =   2
   End
   Begin MSForms.ComboBox cmbdivisi 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "2566;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      DropButtonStyle =   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      Caption         =   "Sub Divisi"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   150
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "From                                            To"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   510
      Width           =   2535
   End
End
Attribute VB_Name = "frmunconfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Obj As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim i, j As Integer
Dim num01 As Long

Private Sub cmbdivisi_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    'KeyCode = 0
End Sub

Private Sub cmbdivisi_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then date1.SetFocus
    'KeyAscii = 0
End Sub

Private Sub cmdadd_Click()
    If cmbdivisi = "" Or date2 < date1 Then
        MsgBox "Data entry not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    j = 0
    For i = 1 To grid.Rows - 1
        grid.Col = 0
        grid.Row = i
        If grid.CellPicture = check Then j = j + 1
    Next i
    
    If j = 0 Then
        MsgBox "Save aborted, there is no data to save.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    For i = 1 To grid.Rows - 1
        grid.Col = 0
        grid.Row = i
        
        If grid.CellPicture = check Then
            Obj.Open dsn
            SQL = "SELECT flag2,ref2 FROM am_beliapp WHERE nobeli = '" & grid.TextMatrix(grid.Row, 1) & "'"
            Set RST = Obj.Execute(SQL)
            If Not RST.EOF Then
                If RST!flag2 = "2" Then
                    Obj.Close
                    MsgBox "Data already Posting, Un Confirm aborted ...", vbCritical, "Information"
                    Exit Sub
                End If
                
               ' MsgBox RST!ref2
                'MsgBox RST!kodesupp
                
                If Trim(RST!ref2) = "" Then
                    MsgBox "There is no reference on this confirm Data.", vbExclamation, "Warning"
                Else
                    SQL = "SELECT * FROM am_apopnfil WHERE noapply = '" & RST!ref2 & "' and transtype <> 'I'"
                    Set RST = Obj.Execute(SQL)
                    If Not RST.EOF Then
                        MsgBox "There is apply on this confirm Data, user must remove those apply (CN/DN/PM).", vbExclamation, "Warning"
                    Else
                        SQL = "delete FROM am_apopnfil WHERE nobeli = '" & grid.TextMatrix(grid.Row, 1) & "'"
                        Set RST = Obj.Execute(SQL)
                        
                        Dim novou As String
                        SQL = "Select ref1 From am_beliapp Where nobeli = '" & grid.TextMatrix(grid.Row, 1) & "'"
                        Set RST = Obj.Execute(SQL)
                        novou = RST!ref1
                        
                        SQL = "Delete From am_voucherhdr Where novoucher = '" & novou & "'"
                        Set RST = Obj.Execute(SQL)
                        
                        SQL = "Delete From am_voucherin Where novoucher = '" & novou & "'"
                        Set RST = Obj.Execute(SQL)
                        
                        SQL = "update am_beliapp set flag2='0',ref1='',ref2='' WHERE nobeli = '" & grid.TextMatrix(grid.Row, 1) & "'"
                        Set RST = Obj.Execute(SQL)
                    End If
                End If
            End If
            Obj.Close
        End If
    Next i
    
    MsgBox "Un Confirm done ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdclear_Click()
    cmbdivisi = ""
    date1 = Date
    date2 = Date
    
    grid.Clear
    grid.Rows = 2
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdswitch_Click()
    If cmbdivisi = "" Or date2 < date1 Then
        MsgBox "Data entry not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    grid.Clear
    grid.Rows = 2
    grid.TextMatrix(0, 1) = "No Beli"
    grid.TextMatrix(0, 2) = "No PO"
    grid.TextMatrix(0, 3) = "Supplier"
    grid.ColWidth(0) = 300
    grid.ColWidth(1) = 1750
    grid.ColWidth(2) = 1750
    grid.ColWidth(3) = 4000
    grid.RowHeightMin = 300
        
    grid.Row = 1
    Obj.Open dsn
    SQL = "SELECT count(c.nobeli)'tot' FROM am_beliapp c WHERE c.nobeli like '%" & cmbdivisi & "' and c.flag2='1' and c.tglbeli>=convert(datetime,'" & tanggal1 & "') and c.tglbeli<=convert(datetime,'" & tanggal2 & "')"
    Set RST = Obj.Execute(SQL)
    If Not RST.EOF Then num01 = RST!tot
    If num01 <> 0 Then
        ak.Max = RST!tot
        ak.Value = 0
        ak.Value = 1
        ak.CaptionType = CaptionPercent
    End If
    
    SQL = "SELECT c.nobeli,c.nopo,a.namasupp FROM am_beliapp c left join am_supplier a on a.kodesupp=c.kodesupp WHERE c.nobeli like '%" & cmbdivisi & "' and c.flag2='1' and c.tglbeli>=convert(datetime,'" & tanggal1 & "') and c.tglbeli<=convert(datetime,'" & tanggal2 & "') order by c.nobeli"
    Set RST = Obj.Execute(SQL)
    Do While Not RST.EOF
        grid.Col = 1
        grid.CellAlignment = 1
        grid.TextMatrix(grid.Row, 1) = RST!NoBeli
        grid.Col = 2
        grid.CellAlignment = 1
        grid.TextMatrix(grid.Row, 2) = RST!NOPO
        grid.TextMatrix(grid.Row, 3) = RST!namasupp
                
        grid.Col = 0
        Set grid.CellPicture = uncheck.Picture
        ak.Value = ak.Value + 1
        grid.Rows = grid.Rows + 1
        grid.Row = grid.Row + 1
        RST.MoveNext
    Loop
    Obj.Close
    If num01 <> 0 Then
        ak.Value = 0
        ak.CaptionType = CaptionNone
    End If
End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='71' and b.kodeuser = '2" & kuser & "'"
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
    
    date1 = Date
    date2 = Date
    
    Obj.Open dsn
    SQL = "select * from am_kode"
    Set RST = Obj.Execute(SQL)
    If Not RST.EOF Then
        Do While Not RST.EOF
            cmbdivisi.AddItem RST!kode3
            
            RST.MoveNext
        Loop
    End If
    Obj.Close
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    
    Select Case grid.Col
        Case 0
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            If grid.CellPicture = uncheck Then
                If MsgBox("UnConfirm this PO ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                    Set grid.CellPicture = check
                End If
            ElseIf grid.CellPicture = check Then
                If MsgBox("Undo this flag ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                    Set grid.CellPicture = uncheck
                End If
            End If
    End Select
End Sub

Function tanggal1()
    tanggal1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggal2()
    tanggal2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function
