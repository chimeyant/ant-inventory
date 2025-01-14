VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0A1435CB-EB1C-11D4-89B0-204C4F4F5020}#3.0#0"; "akprogressbar.ocx"
Begin VB.Form frmdaftartagihsub 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pilih Faktur yang akan ditagih"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
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
   Icon            =   "frmdaftartagihsub.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin akProgress.akProgressBar pb1 
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   4080
      Visible         =   0   'False
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   503
      BackColour      =   -2147483633
      FontColour      =   4210752
      BarColour       =   16776960
      Horizontal      =   -1  'True
      ReverseGradient =   0   'False
      Max             =   100
      Min             =   0
      GapWidth        =   0
      LineWidth       =   1
      Caption         =   3
      BorderStyle     =   0
      Margin          =   2
      Gradient        =   3
      Alignment       =   2
   End
   Begin VB.TextBox txtnobukti 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      MaxLength       =   15
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   3615
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   6376
      _Version        =   393216
      Cols            =   6
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
      _Band(0).Cols   =   6
   End
   Begin VB.Label Label1 
      Caption         =   "press ENTER to search   -   double click on grid below to choose   -   press ESC to cancel"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   150
      Width           =   6495
   End
   Begin VB.Label Label2 
      Caption         =   "Faktur"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   150
      Width           =   855
   End
End
Attribute VB_Name = "frmdaftartagihsub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim RST1 As New ADODB.Recordset
Dim SQL1 As String

Private Sub cari()
    grid.Clear
    grid.Rows = 2
    
    grid.TextMatrix(0, 0) = "Faktur"
    grid.TextMatrix(0, 1) = "Tanggal"
    grid.TextMatrix(0, 2) = "Kode Cust"
    grid.TextMatrix(0, 3) = "Nama Cust"
    grid.TextMatrix(0, 4) = "Alamat"
    grid.TextMatrix(0, 5) = "Nilai"
    
    grid.ColWidth(0) = 1000
    grid.ColWidth(1) = 1000
    grid.ColWidth(2) = 1000
    grid.ColWidth(3) = 2000
    grid.ColWidth(4) = 2000
    grid.ColWidth(5) = 1500
    
    grid.RowHeightMin = 300
    
    If txtnobukti = "" Then Exit Sub
    PB1.Visible = True
    PB1.Value = 0
    PB1.Caption = "wait..."
    grid.Row = 1
    
    txtnobukti.Enabled = False
    OBJ.Open dsn
    SQL = "select noapply,isnull(sum(amount+potongan+ppn+selisih),0)'nilai' from am_aropnfil where noapply like '" & txtnobukti & "%' group by noapply "
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        If RST!nilai > 0 Then
            SQL1 = "select nobkt,tglbkt,kodecust from am_aropnfil where nobkt = '" & RST!noapply & "'"
            Set RST1 = OBJ.Execute(SQL1)
            If Not RST1.EOF Then
                grid.TextMatrix(grid.Row, 0) = RST1!nobkt
                grid.TextMatrix(grid.Row, 1) = Format(RST1!tglbkt, "dd/MM/yyyy")
                grid.TextMatrix(grid.Row, 2) = RST1!kodecust
                grid.TextMatrix(grid.Row, 5) = Format(RST!nilai, "###,###,##0.00")
    
                SQL1 = "select namacust,alamatcust from am_customer where kodecust = '" & grid.TextMatrix(grid.Row, 2) & "'"
                Set RST1 = OBJ.Execute(SQL1)
                grid.TextMatrix(grid.Row, 3) = RST1!namacust
                grid.TextMatrix(grid.Row, 4) = RST1!alamatcust
            End If
     
            grid.Rows = grid.Rows + 1
            grid.Row = grid.Row + 1
        End If
        
        RST.MoveNext
        If PB1.Value = 100 Then PB1.Value = 0
        PB1.Value = PB1.Value + 1
    Loop
    OBJ.Close
    PB1.Value = 0
    PB1.Caption = "Done"
    'pb1.Visible = False
    txtnobukti.Enabled = True
    txtnobukti.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub grid_DblClick()
    If grid.MouseRow = 0 Then Exit Sub
    If grid.TextMatrix(grid.Row, 0) = "" Then Exit Sub
    
    frmdaftartagih.grid.TextMatrix(setup1, 1) = grid.TextMatrix(grid.Row, 0)
    frmdaftartagih.grid.TextMatrix(setup1, 2) = grid.TextMatrix(grid.Row, 1)
    frmdaftartagih.grid.TextMatrix(setup1, 3) = grid.TextMatrix(grid.Row, 2)
    frmdaftartagih.grid.TextMatrix(setup1, 4) = grid.TextMatrix(grid.Row, 3)
    frmdaftartagih.grid.TextMatrix(setup1, 5) = grid.TextMatrix(grid.Row, 4)
    frmdaftartagih.grid.TextMatrix(setup1, 6) = grid.TextMatrix(grid.Row, 5)
    
    Unload Me
End Sub

Private Sub txtnobukti_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then cari
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
       
    grid.Clear
    grid.Rows = 2
    
    grid.TextMatrix(0, 0) = "Faktur"
    grid.TextMatrix(0, 1) = "Tanggal"
    grid.TextMatrix(0, 2) = "Kode Cust"
    grid.TextMatrix(0, 3) = "Nama Cust"
    grid.TextMatrix(0, 4) = "Alamat"
    grid.TextMatrix(0, 5) = "Nilai"
    
    grid.ColWidth(0) = 1000
    grid.ColWidth(1) = 1000
    grid.ColWidth(2) = 1000
    grid.ColWidth(3) = 2000
    grid.ColWidth(4) = 2000
    grid.ColWidth(5) = 1500
    
    grid.RowHeightMin = 300
End Sub
