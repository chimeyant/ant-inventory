VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmcekgiro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Maintain Giro"
   ClientHeight    =   5985
   ClientLeft      =   3615
   ClientTop       =   3105
   ClientWidth     =   11055
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
   Icon            =   "frmcekgiro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   9120
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
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
      CustomFormat    =   "dd MMM yyyy"
      Format          =   89784323
      CurrentDate     =   37421
   End
   Begin VB.PictureBox black 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4800
      Picture         =   "frmcekgiro.frx":2372
      ScaleHeight     =   255
      ScaleWidth      =   375
      TabIndex        =   9
      Top             =   6240
      Width           =   375
   End
   Begin VB.PictureBox blank 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   6000
      ScaleHeight     =   255
      ScaleWidth      =   375
      TabIndex        =   8
      Top             =   6240
      Width           =   375
   End
   Begin VB.PictureBox dot 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5400
      Picture         =   "frmcekgiro.frx":38D8
      ScaleHeight     =   255
      ScaleWidth      =   375
      TabIndex        =   7
      Top             =   6240
      Width           =   375
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid1 
      Height          =   4935
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   8705
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   120
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
      Format          =   89784323
      CurrentDate     =   37421
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   9960
      TabIndex        =   5
      Top             =   5520
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
      MICON           =   "frmcekgiro.frx":3D5E
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
      Left            =   9000
      TabIndex        =   4
      Top             =   5520
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
      MICON           =   "frmcekgiro.frx":4078
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
      Left            =   8040
      TabIndex        =   3
      Top             =   5520
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Save"
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
      MICON           =   "frmcekgiro.frx":4392
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai1 
      Height          =   255
      Left            =   7200
      TabIndex        =   10
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
      _Version        =   65536
      _ExtentX        =   1296
      _ExtentY        =   450
      Calculator      =   "frmcekgiro.frx":46AC
      Caption         =   "frmcekgiro.frx":46CC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcekgiro.frx":4738
      Keys            =   "frmcekgiro.frx":4756
      Spin            =   "frmcekgiro.frx":4798
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
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.Label Label6 
      Caption         =   "Tanggal J. Tempo"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   150
      Width           =   1335
   End
End
Attribute VB_Name = "frmcekgiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim OBJ1 As New ADODB.Connection
Dim RST1 As New ADODB.Recordset
Dim SQL1 As String

Dim posrow As String

Private Sub cmdadd_Click()
    If MsgBox("Save this change ?", vbYesNo + vbQuestion, "Question") = vbNo Then Exit Sub
    
    Grid1.Row = 1
    Do While True
        If Grid1.Rows = Grid1.Row + 1 Then Exit Do
        
        Grid1.Col = 3
        If Grid1.CellPicture = dot Then
            If Grid1.TextMatrix(Grid1.Row, 2) = "" Then GoTo capekdeh
        End If
        Grid1.Col = 4
        If Grid1.CellPicture = dot Then
            If Grid1.TextMatrix(Grid1.Row, 2) = "" Then GoTo capekdeh
        End If
        
        Grid1.Col = 3
        If Grid1.CellPicture = dot Then
            OBJ.Open dsn
            SQL = "update am_cashsub set tglcair=Convert(dateTime, '" & tanggalgrid & "') " & _
            "where nogiro='" & Grid1.TextMatrix(Grid1.Row, 0) & "' and " & _
            "nobkt='" & Grid1.TextMatrix(Grid1.Row, 6) & "'"
            Set RST = OBJ.Execute(SQL)
            OBJ.Close
            
            GoTo capekdeh
        End If
        
        Grid1.Col = 4
        If Grid1.CellPicture = dot Then
            OBJ.Open dsn
            SQL = "update am_cashsub set tgltolak=Convert(dateTime, '" & tanggalgrid & "') " & _
            "where nogiro='" & Grid1.TextMatrix(Grid1.Row, 0) & "' and " & _
            "nobkt='" & Grid1.TextMatrix(Grid1.Row, 6) & "'"
            Set RST = OBJ.Execute(SQL)
            OBJ.Close
        End If
        
capekdeh:
        Grid1.Row = Grid1.Row + 1
    Loop
    
    MsgBox "Data saved, click ok to continue  ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdclear_Click()
    Grid1.Clear
    Grid1.Rows = 2
    
    Grid1.TextMatrix(0, 0) = "No Giro"
    Grid1.TextMatrix(0, 1) = "Bank"
    Grid1.TextMatrix(0, 2) = "Tgl Cair/Tolak"
    Grid1.TextMatrix(0, 3) = "C"
    Grid1.TextMatrix(0, 4) = "T"
    Grid1.TextMatrix(0, 5) = "Nilai Giro"
    Grid1.TextMatrix(0, 6) = "No Bayar"
    Grid1.TextMatrix(0, 7) = "Customer"
    
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 1500
    Grid1.ColWidth(2) = 1500
    Grid1.ColWidth(3) = 300
    Grid1.ColWidth(4) = 300
    Grid1.ColWidth(5) = 1500
    Grid1.ColWidth(6) = 1000
    Grid1.ColWidth(7) = 3300
    
    Grid1.RowHeightMin = 300
    
    date1 = Date
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub date1_Change()
    Grid1.Clear
    Grid1.Rows = 2
    
    Grid1.TextMatrix(0, 0) = "No Giro"
    Grid1.TextMatrix(0, 1) = "Bank"
    Grid1.TextMatrix(0, 2) = "Tgl Cair/Tolak"
    Grid1.TextMatrix(0, 3) = "C"
    Grid1.TextMatrix(0, 4) = "T"
    Grid1.TextMatrix(0, 5) = "Nilai Giro"
    Grid1.TextMatrix(0, 6) = "No Bayar"
    Grid1.TextMatrix(0, 7) = "Customer"
    
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 1500
    Grid1.ColWidth(2) = 1500
    Grid1.ColWidth(3) = 300
    Grid1.ColWidth(4) = 300
    Grid1.ColWidth(5) = 1500
    Grid1.ColWidth(6) = 1000
    Grid1.ColWidth(7) = 3300
    
    Grid1.RowHeightMin = 300
    
    Grid1.Row = 1
    
    OBJ.Open dsn
    SQL = "select * from am_cashsub where typebayar='G' and tgljt = convert(datetime, '" & tanggal1 & "') and year(tgltolak)=1900 and year(tglcair)=1900 order by nogiro"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        Grid1.TextMatrix(Grid1.Row, 0) = RST!nogiro
        Grid1.TextMatrix(Grid1.Row, 1) = RST!bank
        'grid1.TextMatrix(grid1.Row, 2) = RST!acbank
        Grid1.TextMatrix(Grid1.Row, 5) = Format(RST!jumlah, "###,###,###,##0.00")
        Grid1.TextMatrix(Grid1.Row, 6) = RST!nobkt
        
        OBJ1.Open dsn
        SQL1 = "select namacust from am_customer where kodecust = '" & RST!kodecust & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then Grid1.TextMatrix(Grid1.Row, 7) = RST1!namacust
        OBJ1.Close
        
        RST.MoveNext
        Grid1.Rows = Grid1.Rows + 1
        Grid1.Row = Grid1.Row + 1
    Loop
    OBJ.Close
End Sub

Private Sub date2_CloseUp()
    Grid1.TextMatrix(posrow, 2) = Format(date2, "dd MMM yyyy")

    Grid1.SetFocus
    Grid1.Row = posrow
    date2.Visible = False
End Sub

Private Sub date2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then date2.Visible = False
    If KeyCode = 13 Then
        Grid1.TextMatrix(posrow, 2) = Format(date2, "dd MMM yyyy")

        Grid1.SetFocus
        Grid1.Row = posrow
        date2.Visible = False
    End If
End Sub

Private Sub date2_LostFocus()
    date2.Visible = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    
    
    date1 = Date
    
    Grid1.TextMatrix(0, 0) = "No Giro"
    Grid1.TextMatrix(0, 1) = "Bank"
    Grid1.TextMatrix(0, 2) = "Tgl Cair/Tolak"
    Grid1.TextMatrix(0, 3) = "C"
    Grid1.TextMatrix(0, 4) = "T"
    Grid1.TextMatrix(0, 5) = "Nilai Giro"
    Grid1.TextMatrix(0, 6) = "No Bayar"
    Grid1.TextMatrix(0, 7) = "Customer"
    
    Grid1.ColWidth(0) = 1000
    Grid1.ColWidth(1) = 1500
    Grid1.ColWidth(2) = 1500
    Grid1.ColWidth(3) = 300
    Grid1.ColWidth(4) = 300
    Grid1.ColWidth(5) = 1500
    Grid1.ColWidth(6) = 1000
    Grid1.ColWidth(7) = 3300
    
    Grid1.RowHeightMin = 300
End Sub

Private Sub grid1_Click()
    If Grid1.MouseRow = 0 Then Exit Sub
    If Grid1.TextMatrix(Grid1.Row, 0) = "" Then Exit Sub
    posrow = Grid1.Row
    
    Select Case Grid1.Col
        Case 2
            If date2.Visible = True Then Exit Sub
            
            date2.Width = Grid1.ColWidth(2) - 20
            date2.Height = 290
            If Grid1.TextMatrix(Grid1.Row, 2) <> "" Then date2 = Grid1.TextMatrix(Grid1.Row, 2) Else date2 = date1
            date2.Left = Grid1.Left + Grid1.CellLeft - 10
            date2.Top = Grid1.Top + Grid1.CellTop - 20
            date2.Visible = True
            date2.SetFocus
        Case 3
            Grid1.Col = 4
            If Grid1.CellPicture = dot Then Exit Sub
            Grid1.Col = 3
            If Grid1.CellPicture = blank Then
                Set Grid1.CellPicture = dot
            Else
                Set Grid1.CellPicture = blank
            End If
        Case 4
            Grid1.Col = 3
            If Grid1.CellPicture = dot Then Exit Sub
            Grid1.Col = 4
            If Grid1.CellPicture = blank Then
                Set Grid1.CellPicture = dot
            Else
                Set Grid1.CellPicture = blank
            End If
    End Select
End Sub

Private Sub grid1_EnterCell()
    If Grid1.TextMatrix(Grid1.Row, 0) = "" Then Exit Sub
    posrow = Grid1.Row
    
    Select Case Grid1.Col
        Case 2
            If date2.Visible = True Then Exit Sub

            date2.Width = Grid1.ColWidth(2) - 20
            date2.Height = 290
            If Grid1.TextMatrix(Grid1.Row, 2) <> "" Then date2 = Grid1.TextMatrix(Grid1.Row, 2) Else date2 = date1
            date2.Left = Grid1.Left + Grid1.CellLeft - 10
            date2.Top = Grid1.Top + Grid1.CellTop - 20
            date2.Visible = True
            date2.SetFocus
    End Select
End Sub

Private Sub Grid1_Scroll()
    date2.Visible = False
End Sub

Function tanggal1()
    tanggal1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function

Function tanggalgrid()
    tanggalgrid = Month(Grid1.TextMatrix(Grid1.Row, 2)) & "/" & Day(Grid1.TextMatrix(Grid1.Row, 2)) & "/" & Year(Grid1.TextMatrix(Grid1.Row, 2))
End Function
