VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmabout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Us"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5175
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2100
      Left            =   90
      TabIndex        =   2
      Top             =   480
      Width           =   4995
      Begin VB.Label lblserialnumber 
         Caption         =   "Label8"
         Height          =   225
         Left            =   1830
         TabIndex        =   14
         Top             =   1695
         Width           =   3060
      End
      Begin VB.Label lblconfirmnumber 
         Caption         =   "Label8"
         Height          =   225
         Left            =   1830
         TabIndex        =   13
         Top             =   1350
         Width           =   3060
      End
      Begin VB.Label lbllisensi 
         Caption         =   "Label8"
         Height          =   225
         Left            =   1830
         TabIndex        =   12
         Top             =   1035
         Width           =   3060
      End
      Begin VB.Label lblcompany 
         Caption         =   "Label8"
         Height          =   225
         Left            =   1830
         TabIndex        =   11
         Top             =   735
         Width           =   3060
      End
      Begin VB.Label lblappver 
         Caption         =   "Label8"
         Height          =   225
         Left            =   1830
         TabIndex        =   10
         Top             =   480
         Width           =   3060
      End
      Begin VB.Label lblappname 
         Caption         =   "Label8"
         Height          =   225
         Left            =   1830
         TabIndex        =   9
         Top             =   195
         Width           =   3060
      End
      Begin VB.Label Label7 
         Caption         =   "Liciented To"
         Height          =   225
         Left            =   135
         TabIndex        =   8
         Top             =   1035
         Width           =   1365
      End
      Begin VB.Label Label6 
         Caption         =   "Company "
         Height          =   225
         Left            =   120
         TabIndex        =   7
         Top             =   735
         Width           =   1365
      End
      Begin VB.Label Label5 
         Caption         =   "Serial Number "
         Height          =   225
         Left            =   120
         TabIndex        =   6
         Top             =   1665
         Width           =   1365
      End
      Begin VB.Label Label4 
         Caption         =   "Confirm Number "
         Height          =   225
         Left            =   105
         TabIndex        =   5
         Top             =   1365
         Width           =   1365
      End
      Begin VB.Label Label3 
         Caption         =   "Aplication Ver."
         Height          =   225
         Left            =   120
         TabIndex        =   4
         Top             =   465
         Width           =   1365
      End
      Begin VB.Label Label2 
         Caption         =   "Aplication Name "
         Height          =   225
         Left            =   135
         TabIndex        =   3
         Top             =   195
         Width           =   1365
      End
   End
   Begin XtremeSuiteControls.PushButton btnclose 
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   2670
      Width           =   795
      _Version        =   851970
      _ExtentX        =   1402
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "CLOSE"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ANT INVENTORY SYSTEM"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   105
      TabIndex        =   1
      Top             =   75
      Width           =   4995
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnclose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblappname = ": " & AppName
    lblappver = ": " & AppVer
    lblcompany = ": " & AppComp & AppCopyright
    lbllisensi = ": " & "PT. SPARTA PRIMA"
    lblconfirmnumber = ": " & "0000-0000-0000-0000"
    lblserialnumber = ": " & "1111-1111-1111-1111"
End Sub
