VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmhitung 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1335
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   4095
      _Version        =   851970
      _ExtentX        =   7223
      _ExtentY        =   2355
      _StockProps     =   79
      Caption         =   "PRODUK"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Select"
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
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "All"
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
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtkdbarang 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   720
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.TextBox txtbarang 
         Appearance      =   0  'Flat
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
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   2640
      End
   End
   Begin XtremeSuiteControls.PushButton cmdclose 
      Height          =   315
      Left            =   3480
      TabIndex        =   0
      Top             =   3000
      Width           =   870
      _Version        =   851970
      _ExtentX        =   1535
      _ExtentY        =   556
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
      Appearance      =   5
   End
   Begin XtremeSuiteControls.PushButton cmdview 
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      Top             =   3000
      Width           =   870
      _Version        =   851970
      _ExtentX        =   1535
      _ExtentY        =   556
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
      Appearance      =   5
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   1335
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   4095
      _Version        =   851970
      _ExtentX        =   7223
      _ExtentY        =   2355
      _StockProps     =   79
      Caption         =   "BAHAN BAKU"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
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
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   360
         Width           =   2640
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   720
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "All"
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
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Select"
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
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmhitung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    If Check1.Value = True Then
        Check2.Value = False
        namatabel = "produk"
        carisql1 = "select kode_produk,nama_produk from list_produk_master"
        frmsearch.Show vbModal
    End If
End Sub

Private Sub Check1_GotFocus()
    If hasil = "" Then Exit Sub
    txtkdbarang = hasil
    txtbarang = hasil1
    hasil = ""
    hasil1 = ""
    namatabel = ""
    carisql1 = ""
End Sub

Private Sub Check2_Click()
    If Check2.Value = True Then
        Check1.Value = False
    End If
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

