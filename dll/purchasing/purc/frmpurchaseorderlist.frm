VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmpurchaseorderlist 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Laporan Purchase Order"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   120
      TabIndex        =   31
      Top             =   3120
      Visible         =   0   'False
      Width           =   3495
      Begin VB.TextBox txtkode7 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   33
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtkode6 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   32
         Top             =   120
         Width           =   1575
      End
      Begin Chameleon.chameleonButton cmdsearch7 
         Height          =   285
         Left            =   0
         TabIndex        =   34
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         BTYPE           =   9
         TX              =   "To Kode"
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
         MICON           =   "frmpurchaseorderlist.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdsearch6 
         Height          =   285
         Left            =   0
         TabIndex        =   35
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         BTYPE           =   9
         TX              =   "From Kode"
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
         MICON           =   "frmpurchaseorderlist.frx":031A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker date3 
         Height          =   300
         Left            =   1080
         TabIndex        =   36
         Top             =   840
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   134742019
         CurrentDate     =   38679
      End
      Begin MSComCtl2.DTPicker date4 
         Height          =   300
         Left            =   1080
         TabIndex        =   37
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   134742019
         CurrentDate     =   38679
      End
      Begin VB.Label Label6 
         Caption         =   "From Date"
         Height          =   255
         Left            =   0
         TabIndex        =   39
         Top             =   870
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "To Date"
         Height          =   255
         Left            =   0
         TabIndex        =   38
         Top             =   1230
         Width           =   975
      End
   End
   Begin VB.OptionButton Option9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Daftar Purchase Order by Produk+Date"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   3375
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   120
      TabIndex        =   29
      Top             =   3120
      Visible         =   0   'False
      Width           =   3375
      Begin VB.TextBox txtkode5 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   16
         Top             =   120
         Width           =   1575
      End
      Begin Chameleon.chameleonButton cmdsearch5 
         Height          =   285
         Left            =   120
         TabIndex        =   30
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         BTYPE           =   9
         TX              =   "Supplier"
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
         MICON           =   "frmpurchaseorderlist.frx":0634
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.OptionButton Option8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Daftar Purchase Order by Supplier+Produk"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   3495
   End
   Begin VB.OptionButton Option7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Daftar Purchase Order"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   2655
   End
   Begin VB.OptionButton Option6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Daftar Purchase Order by Supplier"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   2775
   End
   Begin VB.OptionButton Option5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Daftar Purchase Order by Produk"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   120
      TabIndex        =   26
      Top             =   3120
      Visible         =   0   'False
      Width           =   3375
      Begin VB.TextBox txtkode4 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   15
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtkode3 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   14
         Top             =   120
         Width           =   1575
      End
      Begin Chameleon.chameleonButton cmdsearch3 
         Height          =   285
         Left            =   120
         TabIndex        =   27
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         BTYPE           =   9
         TX              =   "Supplier"
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
         MICON           =   "frmpurchaseorderlist.frx":094E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdsearch4 
         Height          =   285
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         BTYPE           =   9
         TX              =   "Produk"
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
         MICON           =   "frmpurchaseorderlist.frx":0C68
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   120
      TabIndex        =   23
      Top             =   3120
      Visible         =   0   'False
      Width           =   3375
      Begin VB.TextBox txtkode1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   12
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox txtkode2 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   13
         Top             =   480
         Width           =   1575
      End
      Begin Chameleon.chameleonButton cmdsearch2 
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         BTYPE           =   9
         TX              =   "To Kode"
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
         MICON           =   "frmpurchaseorderlist.frx":0F82
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdsearch1 
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         BTYPE           =   9
         TX              =   "From Kode"
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
         MICON           =   "frmpurchaseorderlist.frx":129C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancel Purchase Order"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   2055
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Close Purchase Order"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   1935
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Out Standing"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Daftar Purchase Order by Date"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.ComboBox cmbkode 
      Height          =   315
      Left            =   1320
      TabIndex        =   9
      Top             =   2760
      Width           =   1575
   End
   Begin Crystal.CrystalReport Crystal 
      Left            =   120
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   2760
      TabIndex        =   18
      Top             =   4800
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
      MICON           =   "frmpurchaseorderlist.frx":15B6
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
      Left            =   1800
      TabIndex        =   17
      Top             =   4800
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
      MICON           =   "frmpurchaseorderlist.frx":18D0
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
      TabIndex        =   10
      Top             =   3240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   134742019
      CurrentDate     =   38679
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   300
      Left            =   1320
      TabIndex        =   11
      Top             =   3600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   134742019
      CurrentDate     =   38679
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "To Date"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   3630
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "From Date"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   3270
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Sub Divisi"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   2790
      Width           =   1095
   End
End
Attribute VB_Name = "frmpurchaseorderlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private Sub cmbkode_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub cmbkode_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdclear_Click()
    If cmbkode = "" Then
        MsgBox "Data entry not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If Option5.Value = True Then
        If txtkode2 < txtkode1 Then
            MsgBox "To kode can not smaller then From kode.", vbExclamation, "Warning"
            date2.SetFocus
            Exit Sub
        End If
    ElseIf Option9.Value = True Then
        If txtkode7 < txtkode6 Then
            MsgBox "To kode can not smaller then From kode.", vbExclamation, "Warning"
            date2.SetFocus
            Exit Sub
        End If
        If date4 < date3 Then
            MsgBox "To Date can not smaller then From Date.", vbExclamation, "Warning"
            date2.SetFocus
            Exit Sub
        End If
    ElseIf Option6.Value = True Or Option8.Value = True Then
        
    Else
        If date2 < date1 Then
            MsgBox "To Date can not smaller then From Date.", vbExclamation, "Warning"
            date2.SetFocus
            Exit Sub
        End If
    End If
        
    Crystal.Reset
    Crystal.WindowState = crptMaximized
    Crystal.WindowShowCloseBtn = True
    Crystal.WindowShowPrintSetupBtn = True
    Crystal.WindowShowSearchBtn = True
    Crystal.WindowShowRefreshBtn = True
    Crystal.Connect = dsnreport
    If Option7.Value = True Then
        Crystal.DataFiles(0) = "Proc(am_outstandingpo)"
        Crystal.ReportFileName = AppPath & "\reports\purchasing\purc\outstandingpo1.rpt"
    ElseIf Option1.Value = True Then
        Crystal.DataFiles(0) = "Proc(am_polist)"
        Crystal.ReportFileName = AppPath & "\reports\purchasing\purc\purchaseorderlist.rpt"
    ElseIf Option5.Value = True Then
        Crystal.DataFiles(0) = "Proc(am_polist1)"
        Crystal.ReportFileName = AppPath & "\reports\purchasing\purc\purchaseorderlist1.rpt"
    ElseIf Option6.Value = True Then
        Crystal.DataFiles(0) = "Proc(am_polist3)"
        Crystal.ReportFileName = AppPath & "\reports\purchasing\purc\purchaseorderlist3.rpt"
    ElseIf Option8.Value = True Then
        Crystal.DataFiles(0) = "Proc(am_polist2)"
        Crystal.ReportFileName = AppPath & "\reports\purchasing\purc\purchaseorderlist2.rpt"
    ElseIf Option2.Value = True Then
        Crystal.DataFiles(0) = "Proc(am_outstandingpo)"
        Crystal.ReportFileName = AppPath & "\reports\purchasing\purc\outstandingpo.rpt"
    ElseIf Option9.Value = True Then
        Crystal.DataFiles(0) = "Proc(am_polist4)"
        Crystal.ReportFileName = AppPath & "\reports\purchasing\purc\purchaseorderlist4.rpt"
    ElseIf Option3.Value = True Then
        Crystal.DataFiles(0) = "Proc(am_poclose)"
        Crystal.ReportFileName = AppPath & "\reports\purchasing\purc\purchaseorderclose.rpt"
    ElseIf Option4.Value = True Then
        Crystal.DataFiles(0) = "Proc(am_pocancel)"
        Crystal.ReportFileName = AppPath & "\reports\purchasing\purc\purchaseordercancel.rpt"
    End If
    Crystal.ParameterFields(0) = "@kode1;" + cmbkode + ";true"
    Crystal.ParameterFields(1) = "@namauser;" + nmuser + ";true"
    If Option5.Value = True Then
        Crystal.ParameterFields(2) = "@kode2;" + txtkode1 + ";true"
        Crystal.ParameterFields(3) = "@kode3;" + txtkode2 + ";true"
    ElseIf Option6.Value = True Then
        Crystal.ParameterFields(2) = "@kode2;" + txtkode5 + ";true"
        Crystal.ParameterFields(3) = "@kode3;" + txtkode5 + ";true"
    ElseIf Option8.Value = True Then
        Crystal.ParameterFields(2) = "@kode2;" + txtkode3 + ";true"
        Crystal.ParameterFields(3) = "@kode3;" + txtkode4 + ";true"
    ElseIf Option9.Value = True Then
        Crystal.ParameterFields(2) = "@kode2;" + txtkode6 + ";true"
        Crystal.ParameterFields(3) = "@kode3;" + txtkode7 + ";true"
        Crystal.ParameterFields(4) = "@kode4;" + Format(date3, "yyyyMMdd") + ";true"
        Crystal.ParameterFields(5) = "@kode5;" + Format(date4, "yyyyMMdd") + ";true"
    Else
        Crystal.ParameterFields(2) = "@kode2;" + Format(date1, "yyyyMMdd") + ";true"
        Crystal.ParameterFields(3) = "@kode3;" + Format(date2, "yyyyMMdd") + ";true"
    End If
    Crystal.RetrieveDataFiles
    Crystal.Action = 1
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdsearch1_Click()
    carisql1 = "select kodebarang, namabarang from am_apitemmst"
    namatabel = "Bahan Baku"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    
    txtkode1 = hasil
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdsearch2_Click()
    carisql1 = "select kodebarang, namabarang from am_apitemmst"
    namatabel = "Bahan Baku"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch2_GotFocus()
    If hasil = "" Then Exit Sub
    
    txtkode2 = hasil
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdsearch3_Click()
    carisql1 = "select namasupp, AlamatSupp1,kodesupp from am_supplier"
    namatabel = "Supplier"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch3_GotFocus()
    If hasil = "" Then Exit Sub
    
    txtkode3 = hasil2
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdsearch4_Click()
    carisql1 = "select kodebarang, namabarang from am_apitemmst"
    namatabel = "Bahan Baku"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch4_GotFocus()
    If hasil = "" Then Exit Sub
    
    txtkode4 = hasil
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdsearch5_Click()
    carisql1 = "select namasupp, AlamatSupp1,kodesupp from am_supplier"
    namatabel = "Supplier"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch5_GotFocus()
    If hasil = "" Then Exit Sub
    
    txtkode5 = hasil2
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdsearch6_Click()
    carisql1 = "select kodebarang, namabarang from am_apitemmst"
    namatabel = "Bahan Baku"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch6_GotFocus()
    If hasil = "" Then Exit Sub
    
    txtkode6 = hasil
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdsearch7_Click()
    carisql1 = "select kodebarang, namabarang from am_apitemmst"
    namatabel = "Bahan Baku"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch7_GotFocus()
    If hasil = "" Then Exit Sub
    
    txtkode7 = hasil
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '   SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='115' and b.kodeuser = '2" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If RST.EOF Then
    '        MsgBox "User Rights Denied !!" & vbCrLf & _
    '        "Please contact your Administrator.", vbCritical, "User Rights"
    '        OBJ.Close
    '        Unload Me
    '        Exit Sub
    '    End If
    '    OBJ.Close
    'End If
End Sub

Private Sub Form_Load()
        
    date1 = Date
    date2 = Date
    date3 = Date
    date4 = Date
    
    OBJ.Open dsn
    SQL = "select * from am_kode"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        Do While Not RST.EOF
            cmbkode.AddItem RST!kode3
            
            RST.MoveNext
        Loop
    End If
    OBJ.Close
End Sub

Private Sub Option1_Click()
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = False
    Frame4.Visible = False
End Sub

Private Sub Option2_Click()
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = False
    Frame4.Visible = False
End Sub

Private Sub Option3_Click()
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = False
    Frame4.Visible = False
End Sub

Private Sub Option4_Click()
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = False
    Frame4.Visible = False
End Sub

Private Sub Option5_Click()
    Frame1.Visible = True
    Frame2.Visible = False
    Frame3.Visible = False
    Frame4.Visible = False
End Sub

Private Sub Option6_Click()
    Frame2.Visible = False
    Frame1.Visible = False
    Frame3.Visible = True
    Frame4.Visible = False
End Sub

Private Sub Option7_Click()
    Frame1.Visible = False
    Frame2.Visible = False
    Frame3.Visible = False
    Frame4.Visible = False
End Sub

Private Sub Option8_Click()
    Frame1.Visible = False
    Frame2.Visible = True
    Frame3.Visible = False
    Frame4.Visible = False
End Sub

Private Sub Option9_Click()
    Frame4.Visible = True
End Sub
