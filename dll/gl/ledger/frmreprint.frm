VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmreprint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reprint Cash/Bank Out"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6600
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   6600
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List2 
      Height          =   3660
      ItemData        =   "frmreprint.frx":0000
      Left            =   120
      List            =   "frmreprint.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   480
      Width           =   5415
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   495
      Left            =   5640
      TabIndex        =   0
      Top             =   3600
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
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
      MICON           =   "frmreprint.frx":0004
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
      Left            =   2280
      TabIndex        =   2
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
      Format          =   134938627
      CurrentDate     =   37426
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   4440
      TabIndex        =   3
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
      Format          =   134938627
      CurrentDate     =   37426
   End
   Begin Chameleon.chameleonButton cmdsubmit 
      Height          =   495
      Left            =   5640
      TabIndex        =   5
      Top             =   600
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Show"
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
      MICON           =   "frmreprint.frx":031E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdprint 
      Height          =   495
      Left            =   5640
      TabIndex        =   6
      Top             =   1200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Preview Print"
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
      MICON           =   "frmreprint.frx":0638
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Crystal.CrystalReport Crystal 
      Left            =   6000
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label1 
      Caption         =   "Display Cash Bank Out from                                              to"
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
      Left            =   120
      TabIndex        =   4
      Top             =   165
      Width           =   4335
   End
End
Attribute VB_Name = "frmreprint"
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

Dim i As Integer

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdprint_Click()
CetakBukti
End Sub

Private Sub cmdsubmit_Click()
    List2.Clear
    If date1 > date2 Then
        MsgBox "Invalid date range, search abort.", vbExclamation, "Error"
        Exit Sub
    End If

    OBJ.Open dsn
    SQL = "Select nobkt,jenis,tgl,keterangan From gl_pengeluaran"
    SQL = SQL + " Where tgl >= '" & tanggal1 & "' and tgl <= '" & tanggal2 & "' order by nobkt asc"
            
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        List2.AddItem RST!nobkt
        RST.MoveNext
    Loop
    OBJ.Close
End Sub

Private Sub CetakBukti()
    For i = 0 To List2.ListCount - 1
        If List2.Selected(i) = True Then
            Crystal.Reset
            Crystal.WindowState = crptMaximized
            Crystal.WindowShowCloseBtn = True
            Crystal.WindowShowPrintSetupBtn = True
            Crystal.WindowShowSearchBtn = True
            Crystal.WindowShowRefreshBtn = True
            Crystal.Connect = dsnreport
            Crystal.DataFiles(0) = "Proc(gl_bukti_pengeluaran)"
            Crystal.ReportFileName = AppPath & "\reports\gl\ledger\bukti_pengeluaran.rpt"
            Crystal.ParameterFields(0) = "@nobukti;" + List2.List(i) + ";true"
            Crystal.RetrieveDataFiles
            Crystal.Action = 1
        End If
    Next i
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    date1 = Date
    date2 = Date
End Sub

Function tanggal1()
    tanggal1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggal2()
    tanggal2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function
