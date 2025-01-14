VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmmutwip 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan Stock"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3195
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   3195
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.RadioButton optlot 
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   1080
      Width           =   1575
      _Version        =   851970
      _ExtentX        =   2778
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "By Lot"
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
   Begin VB.CheckBox ChDetail 
      Caption         =   "Detail"
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
      Left            =   2160
      TabIndex        =   7
      Top             =   600
      Width           =   975
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   2160
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
      MICON           =   "frmmutwip.frx":0000
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
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
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
      Format          =   142934017
      CurrentDate     =   42039
   End
   Begin Chameleon.chameleonButton cmdview 
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   2160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "View"
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
      MICON           =   "frmmutwip.frx":031A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Crystal.CrystalReport crystal 
      Left            =   3480
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin XtremeSuiteControls.RadioButton optitem 
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   1440
      Width           =   1575
      _Version        =   851970
      _ExtentX        =   2778
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "By Item"
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
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Gudang"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   855
   End
   Begin MSForms.ComboBox cmbtype 
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Top             =   600
      Width           =   735
      VariousPropertyBits=   746608667
      MaxLength       =   2
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "1296;503"
      ListRows        =   15
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
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
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmmutwip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String
Dim akses As Boolean

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdview_Click()
    If cmbtype = "" Then Exit Sub
    Crystal.Reset
    Crystal.WindowState = crptMaximized
    Crystal.WindowShowCloseBtn = True
    Crystal.WindowShowPrintSetupBtn = True
    Crystal.WindowShowSearchBtn = True
    Crystal.Connect = dsnreport
    Crystal.DataFiles(0) = "Proc(am_mutwip_bylot)"
    If ChDetail.Value = Unchecked Then
        If optitem.Value = True Then
            Crystal.ReportFileName = AppPath & "\reports\sale\mut\stockbyitem.rpt"
        ElseIf optlot.Value = True Then
            Crystal.ReportFileName = AppPath & "\reports\sale\mut\stockbylot.rpt"
        End If
    Else
        If akses = True Then
            If optitem.Value = True Then
                Crystal.ReportFileName = AppPath & "\reports\sale\mut\stockbyitemhpp_detail.rpt"
            ElseIf optlot.Value = True Then
                Crystal.ReportFileName = AppPath & "\reports\sale\mut\stockbylothpp_detail.rpt"
            End If
        Else
            If optitem.Value = True Then
                Crystal.ReportFileName = AppPath & "\reports\sale\mut\stockbyitem_detail.rpt"
            ElseIf optlot.Value = True Then
                Crystal.ReportFileName = AppPath & "\reports\sale\mut\stockbylot_detail.rpt"
            End If
        End If
    End If
    Crystal.ParameterFields(0) = "@tgl1;" & Format(Date1, "yyyy/MM/dd") & ";true"
    Crystal.ParameterFields(1) = "@gudang;" & cmbtype & ";true"
    If cmbtype = "G1" Then
        Crystal.ParameterFields(2) = "@status;" & "1" & ";true"
    ElseIf cmbtype = "G2" Then
        Crystal.ParameterFields(2) = "@status;" & "2" & ";true"
    ElseIf cmbtype = "G3" Then
        Crystal.ParameterFields(2) = "@status;" & "0" & ";true"
    ElseIf cmbtype = "G4" Then
        Crystal.ParameterFields(2) = "@status;" & "3" & ";true"
    ElseIf cmbtype = "G5" Then
        Crystal.ParameterFields(2) = "@status;" & "4" & ";true"
    End If
    Crystal.ParameterFields(3) = "@user;" & nmuser & ";True"

    Crystal.RetrieveDataFiles
    Crystal.Action = 1
End Sub

Private Sub Form_Load()
    Dim i As Integer
    i = 0
    cmbtype.Clear
    cmbtype.ColumnCount = 2
    cmbtype.ListWidth = "4 cm"
    cmbtype.ColumnWidths = "1 cm; 3 cm"
    
    OBJ.Open dsn
    SQL = "select * from am_gudang"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        Do While Not RST.EOF
            cmbtype.AddItem RST!kodegudang
            cmbtype.List(i, 1) = RST!namagudang
            i = i + 1
            RST.MoveNext
        Loop
    End If
    'Periksa hak akses hpp

    SQL = "Select * From LIST_USERS Where username = '" & nmuser & "'"
    Set RST = OBJ.Execute(SQL)
    
    If Not RST.EOF Then
        If RST!gl = "1" Then
            akses = True
        Else
            akses = False
        End If
    Else
        If nmuser = "Creator" Then akses = True
    End If
    OBJ.Close
    
    Date1 = Date
End Sub
