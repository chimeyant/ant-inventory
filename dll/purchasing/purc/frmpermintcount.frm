VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmpermintcount 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Days Count (Permintaan Barang)"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   3300
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.RadioButton optoutstanding 
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   1695
      _Version        =   851970
      _ExtentX        =   2990
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Outstanding"
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
      Value           =   -1  'True
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   2640
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
      MICON           =   "frmpermintcount.frx":0000
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
      Left            =   1320
      TabIndex        =   2
      Top             =   2640
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Preview"
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
      MICON           =   "frmpermintcount.frx":031A
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
      Height          =   300
      Left            =   1320
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
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
      Format          =   143917059
      CurrentDate     =   38679
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   300
      Left            =   1320
      TabIndex        =   4
      Top             =   1920
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
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
      Format          =   143917059
      CurrentDate     =   38679
   End
   Begin Crystal.CrystalReport Crystal 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin XtremeSuiteControls.RadioButton optclose 
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   600
      Width           =   1695
      _Version        =   851970
      _ExtentX        =   2990
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Selesai"
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
   Begin Chameleon.chameleonButton cmdtes 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Tes"
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
      MICON           =   "frmpermintcount.frx":0634
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin XtremeSuiteControls.RadioButton opt7day 
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   960
      Width           =   1815
      _Version        =   851970
      _ExtentX        =   3201
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Permintaan > 7 Hari"
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
   Begin XtremeSuiteControls.PopupControl popupstatus 
      Left            =   600
      Top             =   0
      _Version        =   851970
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
      Animation       =   2
      ShowDelay       =   20000
   End
   Begin VB.Label Label6 
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
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label5 
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
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1950
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   2520
      Width           =   3495
   End
End
Attribute VB_Name = "frmpermintcount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclear_Click()
    If (date1 > date2) Then
        MsgBox "From Date Greather To Date.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    Crystal.Reset
    Crystal.WindowState = crptMaximized
    Crystal.WindowShowCloseBtn = True
    Crystal.WindowShowPrintSetupBtn = True
    Crystal.WindowShowSearchBtn = True
    Crystal.WindowShowRefreshBtn = True
    Crystal.Connect = dsnreport
    If opt7day.Value = True Then
        Crystal.DataFiles(0) = "Proc(am_perminwarning)"
        Crystal.ReportFileName = AppPath & "\reports\purchasing\purc\nota_up7day.rpt"
    Else
        Crystal.DataFiles(0) = "Proc(am_perminpolate)"
        Crystal.ReportFileName = AppPath & "\reports\purchasing\purc\nota_late.rpt"
        Crystal.ParameterFields(0) = "@kode1;" + Format(date1, "yyyyMMdd") + ";true"
        Crystal.ParameterFields(1) = "@kode2;" + Format(date2, "yyyyMMdd") + ";true"
        If optoutstanding.Value = True Then
            Crystal.ParameterFields(2) = "@status;" + "0" + ";true"
        ElseIf optclose.Value = True Then
            Crystal.ParameterFields(2) = "@status;" + "1" + ";true"
        End If
    End If
    Crystal.RetrieveDataFiles
    Crystal.Action = 1
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdtes_Click()
Dim OBJ As ADODB.Connection
Dim RST As ADODB.Recordset
Dim cmd As ADODB.Command
Dim SQL As String
Dim result As Variant
Dim jml As Long

    Set OBJ = New ADODB.Connection
    OBJ.Open dsn
    
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = OBJ
    cmd.CommandText = "am_perminwarning"
    cmd.CommandType = adCmdStoredProc
    
    Set RST = New ADODB.Recordset
    RST.CursorType = adOpenForwardOnly
    RST.LockType = adLockReadOnly
    RST.Open cmd
    
    jml = 0
    Do While Not RST.EOF
        jml = jml + 1
        RST.MoveNext
    Loop
    RST.MoveFirst
    
    If jml > 0 Then
        SetPopupInfo popupstatus, "PURCHASING", "Jumlah Outstanding Permintaan barang Lebih dari 3 Hari: " _
        & jml & " Item, Mohon untuk segera diproses. Terima kasih"
        popupstatus.Show
        'MsgBox "Jumlah outstanding Permintaan barang Lebih dari 3 Hari: " & jml & _
        '" Item, Mohon untuk segera di proses" & vbCrLf & "Terima Kasih.", vbExclamation, AppName
    End If
    
    RST.Close
    Set RST = Nothing
    OBJ.Close
    Set OBJ = Nothing
    
End Sub

Private Sub SetPopupInfo(popup As XtremeSuiteControls.PopupControl, ByVal msgTitle As String, ByVal msgStatus As String)
    Dim Item As PopupControlItem
    
    popup.RemoveAllItems
    popup.Icons.RemoveAll
    
    Set Item = popup.AddItem(5, 6, 170, 19, AppName)
    Item.Hyperlink = True 'False
    
    Set Item = popup.AddItem(5, 27, 160, 25, msgTitle)
    Item.TextAlignment = DT_LEFT
    Item.CalculateHeight
    Item.CalculateWidth
    
    Set Item = popup.AddItem(5, 50, 170, 200, msgStatus)
    Item.TextAlignment = DT_LEFT Or DT_WORDBREAK
    Item.CalculateHeight
    
    popup.VisualTheme = xtpPopupThemeMSN
    popup.SetSize 200, 130
End Sub

Private Sub opt7day_Click()
    date1.Enabled = False
    date2.Enabled = False
End Sub

Private Sub optclose_Click()
    date1.Enabled = True
    date2.Enabled = True
End Sub

Private Sub optoutstanding_Click()
    date1.Enabled = True
    date2.Enabled = True
End Sub

Private Sub popupstatus_ItemClick(ByVal Item As XtremeSuiteControls.IPopupControlItem)
    MsgBox "ok"
End Sub
Private Sub Form_Load()
    date1.Value = Date
    date2.Value = Date
    If nmuser = "Creator" Then cmdtes.Visible = True
End Sub

Function tanggal1()
    tanggal1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function
Function tanggal2()
    tanggal2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function



