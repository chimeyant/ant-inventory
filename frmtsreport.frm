VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmtsreport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan Troubleshoot"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3030
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   3030
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbstatus 
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
      ItemData        =   "frmtsreport.frx":0000
      Left            =   960
      List            =   "frmtsreport.frx":0002
      TabIndex        =   5
      Top             =   1320
      Width           =   1815
   End
   Begin XtremeSuiteControls.PushButton cmdclose 
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   2160
      Width           =   945
      _Version        =   851970
      _ExtentX        =   1667
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Close"
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
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
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
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   134348803
      CurrentDate     =   38767
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
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
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   134348803
      CurrentDate     =   38767
   End
   Begin XtremeSuiteControls.PushButton cmdview 
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   2160
      Width           =   945
      _Version        =   851970
      _ExtentX        =   1667
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "View"
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
   End
   Begin Crystal.CrystalReport crystal 
      Left            =   240
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Label Label3 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   6
      Top             =   1335
      Width           =   780
   End
   Begin VB.Label Label1 
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
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label4 
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
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "frmtsreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str1 As String
Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdview_Click()
    If date1 > date2 Then
        MsgBox "Invalid date range, search abort.", vbExclamation, "Error"
        Exit Sub
    End If
    If cmbstatus = "" Then
        MsgBox "Select the report status first", vbExclamation, AppName
        Exit Sub
    End If
    If cmbstatus = "Outstanding" Then str1 = "0"
    If cmbstatus = "Proses" Then str1 = "1"
    If cmbstatus = "Cancel" Then str1 = "2"
    If cmbstatus = "Close" Then str1 = "3"
    If cmbstatus = "All" Then str1 = ""
    
    crystal.reset
    crystal.WindowState = crptMaximized
    crystal.WindowShowCloseBtn = True
    crystal.WindowShowPrintBtn = True
    crystal.WindowShowPrintSetupBtn = True
    crystal.WindowShowSearchBtn = True
    crystal.WindowShowRefreshBtn = True
    crystal.Connect = dsnreport
    crystal.DataFiles(0) = "Proc(am_daftaritg)"
    If str1 = "" Then
        crystal.ReportFileName = App.Path & "\reports\Forms\itgall.rpt"
    Else
        crystal.ReportFileName = App.Path & "\reports\Forms\itg.rpt"
    End If
    crystal.ParameterFields(0) = "@tgl1;" & Format(date1, "yyyy/MM/dd") & ";true"
    crystal.ParameterFields(1) = "@tgl2;" & Format(date2, "yyyy/MM/dd") & ";true"
    crystal.ParameterFields(2) = "@status;" + str1 + ";true"
    crystal.RetrieveDataFiles
    crystal.Action = 1
End Sub

Private Sub Form_Load()
    cmbstatus.additem "Outstanding"
    cmbstatus.additem "Proses"
    cmbstatus.additem "Cancel"
    cmbstatus.additem "Close"
    cmbstatus.additem "All"
    date1 = Date
    date2 = Date
    dsnreport
End Sub
Function tanggal1()
    tanggal1 = Year(date1) & "/" & Month(date1) & "/" & Day(date1)
End Function

Function tanggal2()
    tanggal2 = Year(date2) & "/" & Month(date2) & "/" & Day(date2)
End Function
