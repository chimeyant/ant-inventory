VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmrekonsil 
   Caption         =   "Form1"
   ClientHeight    =   6645
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13350
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   13350
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox uncheck 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   7680
      Picture         =   "frmrekonsil.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox check 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8040
      Picture         =   "frmrekonsil.frx":034E
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox blank 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7320
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin XtremeSuiteControls.ProgressBar Pg 
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   2400
      Visible         =   0   'False
      Width           =   7695
      _Version        =   851970
      _ExtentX        =   13573
      _ExtentY        =   661
      _StockProps     =   93
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
      UseVisualStyle  =   0   'False
      TextAlignment   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   8705
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      BackColorFixed  =   -2147483637
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
      _Band(0).Cols   =   11
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2160
      Top             =   6120
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   14737632
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1080
      TabIndex        =   5
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
      Format          =   110493699
      CurrentDate     =   37749
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   4080
      TabIndex        =   6
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
      Format          =   110493699
      CurrentDate     =   37749
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "From Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   270
      Width           =   1335
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   "To Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   240
      Width           =   735
   End
   Begin VB.Label lblrecord 
      Appearance      =   0  'Flat
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   6240
      Width           =   1335
   End
End
Attribute VB_Name = "frmrekonsil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private Sub Form_Load()
    date1 = Date
    date2 = Date
    setgrid
End Sub

Private Sub opendata()
    Adodc1.ConnectionString = dsn
    SQL = "Select tgltrx,kdtrx,notrx,kurs,noactrx,desctrx,dbkrtrx,amounttrx,currtrx,cekbg From gl_transaksi "
    SQL = SQL + "Where idupdate <> '1' and tgltrx >= '" & tanggal1 & "' and tgltrx <= '" & tanggal2 & "' "
    SQL = SQL + "and noactrx >= '" & txtacc1 & "' and noactrx <='" & txtacc2 & "'"
    SQL = SQL + "Order By notrx DESC"
    Adodc1.RecordSource = SQL
    Set grid.DataSource = Adodc1
    Adodc1.Refresh
    Adodc1.Recordset.Requery -1
    Pg.Visible = True
    setdata
    grid.Refresh
End Sub

Private Sub setgrid()
    With grid
        .Cols = 11
        .TextMatrix(0, 0) = "No"
        .TextMatrix(0, 1) = "Tgl Trx"
        .TextMatrix(0, 2) = "Kd.Trx"
        .TextMatrix(0, 3) = "No.Trx"
        .TextMatrix(0, 4) = "Kurs"
        .TextMatrix(0, 5) = "No. Acc"
        .TextMatrix(0, 6) = "Description"
        .TextMatrix(0, 7) = "D/K"
        .TextMatrix(0, 8) = "Amount"
        .TextMatrix(0, 9) = "Currency"
        .TextMatrix(0, 10) = "Cek/Giro"
        .ColAlignmentFixed(0) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .ColAlignmentFixed(8) = flexAlignCenterCenter
        .ColAlignmentFixed(10) = flexAlignCenterCenter
        .ColAlignment(0) = flexAlignRightCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignRightCenter
        .ColAlignment(7) = flexAlignCenterCenter
        .ColAlignment(8) = flexAlignRightCenter
        .ColAlignment(9) = flexAlignCenterCenter
        .ColWidth(0) = 800
        .ColWidth(1) = 1000
        .ColWidth(2) = 650
        .ColWidth(3) = 1100
        .ColWidth(4) = 650
        .ColWidth(5) = 1000
        .ColWidth(6) = 5000
        .ColWidth(7) = 400
        .ColWidth(8) = 2000
        .ColWidth(9) = 900
        .ColWidth(10) = 1200
    End With
End Sub
Private Sub setdata()
On Error Resume Next
Dim jml As String
    setgrid
    jml = Adodc1.Recordset.RecordCount
    Pg.Min = 0
    Pg.Max = jml
    Pg.Value = 0
    Pg.text = Str(Int(((Pg.Value / Pg.Max) * 100))) & " % Complete"
    With Adodc1.Recordset
        .MoveFirst
        Do While Not .EOF
            With grid
            grid.Col = 0
            Set grid.CellPicture = uncheck.Picture
            .TextMatrix(.Row, 0) = grid.Row
            .TextMatrix(.Row, 1) = Adodc1.Recordset!tgltrx
            .TextMatrix(.Row, 2) = Adodc1.Recordset!kdtrx
            .TextMatrix(.Row, 3) = Adodc1.Recordset!notrx
            .TextMatrix(.Row, 4) = Adodc1.Recordset!kurs
            .TextMatrix(.Row, 5) = Adodc1.Recordset!noactrx
            .TextMatrix(.Row, 6) = Adodc1.Recordset!desctrx
            .TextMatrix(.Row, 7) = Adodc1.Recordset!dbkrtrx
            .TextMatrix(.Row, 8) = Format(Adodc1.Recordset!amounttrx, "#,##0.00")
            .TextMatrix(.Row, 9) = Adodc1.Recordset!currtrx
            .TextMatrix(.Row, 10) = Adodc1.Recordset!cekbg
            'SetAlternatingGrid grid.Row
            Pg.Value = Pg.Value + 1
            Pg.text = Str(Int(((Pg.Value / Pg.Max) * 100))) & " % Complete"
            
            If grid.Row = jml Then Exit Do
            .Row = .Row + 1
            End With

            lblrecord = grid.Row & " Record"
            Adodc1.Recordset.MoveNext
        Loop
    End With
    Pg.Value = 0
    Pg.Visible = False
End Sub
Private Sub hapusgrid()
On Error Resume Next
    Dim jml As Integer
    jml = grid.Rows
    grid.Row = 1
    Do While True
        grid.Col = 0
        Set grid.CellPicture = blank
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
        If grid.Row = jml - 1 Then Exit Do
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = 2
End Sub

Private Sub hapusrow()
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
    Do While True
        grid.TextMatrix(grid.Row, 0) = grid.TextMatrix(grid.Row + 1, 0)
        grid.TextMatrix(grid.Row, 1) = grid.TextMatrix(grid.Row + 1, 1)
        grid.TextMatrix(grid.Row, 2) = grid.TextMatrix(grid.Row + 1, 2)
        grid.TextMatrix(grid.Row, 3) = grid.TextMatrix(grid.Row + 1, 3)
        grid.TextMatrix(grid.Row, 4) = grid.TextMatrix(grid.Row + 1, 4)
        grid.TextMatrix(grid.Row, 5) = grid.TextMatrix(grid.Row + 1, 5)
        grid.TextMatrix(grid.Row, 6) = grid.TextMatrix(grid.Row + 1, 6)
        grid.TextMatrix(grid.Row, 7) = grid.TextMatrix(grid.Row + 1, 7)
        grid.TextMatrix(grid.Row, 8) = grid.TextMatrix(grid.Row + 1, 8)
        grid.TextMatrix(grid.Row, 9) = grid.TextMatrix(grid.Row + 1, 9)
        grid.TextMatrix(grid.Row, 10) = grid.TextMatrix(grid.Row + 1, 10)
        If grid.Row = grid.Rows - 2 Then Exit Do
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = grid.Rows - 1
    grid.Col = 0
    Set grid.CellPicture = blank
End Sub

Function tanggal1()
    tanggal1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggal2()
    tanggal2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function

