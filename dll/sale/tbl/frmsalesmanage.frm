VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmsalesmanage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manage Sales"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   3615
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   4335
      _Version        =   851970
      _ExtentX        =   7646
      _ExtentY        =   6376
      _StockProps     =   68
      Appearance      =   3
      Color           =   4
      Begin VB.Frame Frame1 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   840
         TabIndex        =   10
         Top             =   1680
         Width           =   1695
         Begin XtremeSuiteControls.RadioButton opton 
            Height          =   375
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Width           =   855
            _Version        =   851970
            _ExtentX        =   1508
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Active"
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
         Begin XtremeSuiteControls.RadioButton optoff 
            Height          =   375
            Left            =   240
            TabIndex        =   12
            Top             =   600
            Width           =   1215
            _Version        =   851970
            _ExtentX        =   2143
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Not Active"
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
      End
      Begin Chameleon.chameleonButton cmdupdate 
         Height          =   375
         Left            =   3360
         TabIndex        =   6
         Top             =   3120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Update"
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
         MICON           =   "frmsalesmanage.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdcancel 
         Height          =   375
         Left            =   2400
         TabIndex        =   9
         Top             =   3120
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Cancel"
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
         MICON           =   "frmsalesmanage.frx":031A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblstatus 
         Caption         =   "Label4"
         Height          =   375
         Left            =   3240
         TabIndex        =   13
         Top             =   1800
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lblsales 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   375
         Left            =   840
         TabIndex        =   8
         Top             =   1200
         Width           =   3375
      End
      Begin VB.Label lblkode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   615
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   3255
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   5741
      _Version        =   393216
      Rows            =   1
      FixedRows       =   0
      FixedCols       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   0
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   4335
      _ExtentX        =   7646
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
      MICON           =   "frmsalesmanage.frx":0634
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "STATUS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SALES"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " KODE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "frmsalesmanage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String
Dim i As Integer

Private Sub cmdcancel_Click()
    lblkode = ""
    lblsales = ""
    opton.Value = False
    optoff.Value = False
    TabControl1.Visible = False
End Sub

Private Sub cmdupdate_click()
    If MsgBox("Are You Sure Want To Update ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    OBJ.Open dsn
    SQL = "UPDATE am_salesman SET "
    SQL = SQL + "IdUpdate = '" & lblstatus & "' "
    SQL = SQL + "WHERE KodeSales = '" & lblkode & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    MsgBox "Data updated, click ok to continue ...", vbInformation, "Information"
    cmdcancel_Click
    opendata
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub opendata()
    OBJ.Open dsn
    SQL = "Select KodeSales,NamaSales,IdUpdate From Am_salesman order by KodeSales asc"
    Set RST = OBJ.Execute(SQL)
    Set grid.DataSource = RST
    SetAlternatingGrid (grid.Row)
    OBJ.Close
End Sub
Private Sub Form_Load()
    Call opendata
End Sub
Private Function SetAlternatingGrid(ByVal i As Integer)
    With grid
        .ColWidth(0) = 800
        .ColWidth(1) = 2200
        .ColWidth(2) = 1000
    End With
    grid.Row = 0
    Do While True
        grid.Col = 0
        If grid.TextMatrix(grid.Row, 2) = "0" Then
            grid.TextMatrix(grid.Row, 2) = "Not Active"
            For i = 0 To grid.Cols - 1
            grid.Col = i
            grid.CellBackColor = &HE0E0E0
            Next
        Else
            grid.TextMatrix(grid.Row, 2) = "Active"
        End If
        If grid.Row = grid.Rows - 1 Then Exit Do
        grid.Row = grid.Row + 1
    Loop
End Function

Private Sub grid_Click()
    TabControl1.Visible = True
    lblkode = grid.TextMatrix(grid.Row, 0)
    lblsales = grid.TextMatrix(grid.Row, 1)
    If grid.TextMatrix(grid.Row, 2) = "Active" Then
        opton.Value = True
        optoff.Value = False
        lblstatus = ""
    ElseIf grid.TextMatrix(grid.Row, 2) = "Not Active" Then
        opton.Value = False
        optoff.Value = True
        lblstatus = "0"
    End If
End Sub

Private Sub optoff_Click()
    If optoff.Value = True Then
        lblstatus = "0"
    Else
        lblstatus = ""
    End If
End Sub

Private Sub opton_Click()
    If opton.Value = True Then
        lblstatus = ""
    Else
        lblstatus = "0"
    End If
End Sub
