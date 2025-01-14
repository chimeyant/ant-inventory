VERSION 5.00
Object = "{E2D000D0-2DA1-11D2-B358-00104B59D73D}#1.0#0"; "titext6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Begin VB.Form frmreportgroup 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6165
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmreportgroup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbdk 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmreportgroup.frx":2372
      Left            =   1800
      List            =   "frmreportgroup.frx":237C
      TabIndex        =   7
      Top             =   3600
      Width           =   3375
   End
   Begin VB.TextBox txtdesc2 
      Appearance      =   0  'Flat
      DataField       =   "NamaArea"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      MaxLength       =   100
      TabIndex        =   2
      Top             =   1800
      Width           =   4215
   End
   Begin VB.ComboBox cmbtype 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmreportgroup.frx":2386
      Left            =   1800
      List            =   "frmreportgroup.frx":2388
      TabIndex        =   1
      Top             =   1440
      Width           =   3375
   End
   Begin VB.ComboBox cmbspace 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmreportgroup.frx":238A
      Left            =   1800
      List            =   "frmreportgroup.frx":23A0
      TabIndex        =   3
      Top             =   2160
      Width           =   3375
   End
   Begin VB.ComboBox cmbmode 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmreportgroup.frx":23EE
      Left            =   1800
      List            =   "frmreportgroup.frx":2404
      TabIndex        =   5
      Top             =   2880
      Width           =   3375
   End
   Begin VB.ComboBox cmbcolom 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmreportgroup.frx":246F
      Left            =   1800
      List            =   "frmreportgroup.frx":248E
      TabIndex        =   4
      Top             =   2520
      Width           =   3375
   End
   Begin VB.TextBox txtsign 
      Appearance      =   0  'Flat
      DataField       =   "NamaArea"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   6
      Top             =   3240
      Width           =   375
   End
   Begin TDBText6Ctl.TDBText txtgroup 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   1080
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   556
      Caption         =   "frmreportgroup.frx":24FF
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmreportgroup.frx":256B
      Key             =   "frmreportgroup.frx":2589
      BackColor       =   -2147483643
      EditMode        =   0
      ForeColor       =   -2147483640
      ReadOnly        =   0
      ShowContextMenu =   0
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MarginBottom    =   1
      Enabled         =   -1
      MousePointer    =   0
      Appearance      =   0
      BorderStyle     =   1
      AlignHorizontal =   0
      AlignVertical   =   0
      MultiLine       =   0
      ScrollBars      =   0
      PasswordChar    =   ""
      AllowSpace      =   0
      Format          =   "9"
      FormatMode      =   0
      AutoConvert     =   0
      ErrorBeep       =   0
      MaxLength       =   6
      LengthAsByte    =   0
      Text            =   ""
      Furigana        =   0
      HighlightText   =   -1
      IMEMode         =   0
      IMEStatus       =   0
      DropWndWidth    =   0
      DropWndHeight   =   0
      ScrollBarMode   =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
   End
   Begin Chameleon.chameleonButton cmdsave 
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   4080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Save"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmreportgroup.frx":25C5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdelete 
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   4080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Delete"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmreportgroup.frx":28DF
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdclear 
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   4080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Clear"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmreportgroup.frx":2BF9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   4920
      TabIndex        =   11
      Top             =   4080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmreportgroup.frx":2F13
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      Caption         =   "Debet / Kredit (Cash Flow Only)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   360
      TabIndex        =   30
      Top             =   3630
      Width           =   1335
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   28
      Top             =   390
      Width           =   855
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Type Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   27
      Top             =   630
      Width           =   1035
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Report Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   26
      Top             =   150
      Width           =   975
   End
   Begin VB.Label lbldesc1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1440
      TabIndex        =   25
      Top             =   390
      Width           =   4215
   End
   Begin VB.Label lblcode1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1440
      TabIndex        =   24
      Top             =   150
      Width           =   735
   End
   Begin VB.Label lbltype1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1440
      TabIndex        =   23
      Top             =   630
      Width           =   4215
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      Caption         =   "Group No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   1110
      Width           =   1335
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      Caption         =   "Type Account"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   1470
      Width           =   1335
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   1830
      Width           =   1335
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      Caption         =   "Space After"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   2190
      Width           =   1335
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      Caption         =   "Print Coloum"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   2550
      Width           =   1335
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      Caption         =   "Print Mode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   2910
      Width           =   1335
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      Caption         =   "Sign (*)                              (+ to - / - to +, For Display Only)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   3270
      Width           =   4455
   End
   Begin VB.Label lblspace 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5280
      TabIndex        =   15
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblcolom 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5280
      TabIndex        =   14
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblmode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5280
      TabIndex        =   13
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lbltype 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5280
      TabIndex        =   12
      Top             =   1440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   -120
      TabIndex        =   29
      Top             =   0
      Width           =   7815
   End
End
Attribute VB_Name = "frmreportgroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbcolom_Click()
    lblcolom = cmbcolom.ListIndex
End Sub

Private Sub cmbcolom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmbmode.SetFocus
    KeyAscii = 0
End Sub

Private Sub cmbdk_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmbmode_Click()
    lblmode = cmbmode.ListIndex
End Sub

Private Sub cmbmode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtsign.SetFocus
    KeyAscii = 0
End Sub

Private Sub cmbspace_Click()
    lblspace = cmbspace.ListIndex
End Sub

Private Sub cmbspace_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And cmbcolom.Enabled = True Then cmbcolom.SetFocus
    If KeyAscii = 13 And cmbcolom.Enabled = False Then cmbmode.SetFocus
    KeyAscii = 0
End Sub

Private Sub cmbtype_Click()
    lbltype = cmbtype.ListIndex
    If lbltype <> 0 Then
        cmbcolom.Enabled = True
        cmbcolom = ""
        lblcolom = ""
    Else
        cmbcolom = ""
        cmbcolom.Enabled = False
        lblcolom = 9
    End If
End Sub

Private Sub cmbtype_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtdesc2.SetFocus
    KeyAscii = 0
End Sub

Private Sub cmdclear_Click()
    If txtgroup.Enabled = False Then Exit Sub
    txtgroup.Enabled = True
    cmbtype.Enabled = True
    cmbcolom.Enabled = True
    txtgroup = ""
    hapusgroup
    txtgroup.SetFocus
End Sub

Private Sub hapusgroup()
    cmbtype = ""
    txtdesc2 = ""
    cmbspace = ""
    cmbcolom = ""
    cmbmode = ""
    txtsign = ""
    lbltype = ""
    lblspace = ""
    lblcolom = ""
    lblmode = ""
    cmbcolom.Enabled = False
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdelete_Click()
    If txtgroup = "" Or txtdesc2 = "" Or cmbtype = "" Or cmbspace = "" Or cmbmode = "" Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If setup3 = "" Then
        MsgBox "There Is No Data To Delete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    frmreport.grid1.Row = setup3
    frmreport.grid1.TextMatrix(frmreport.grid1.Row, 0) = ""
    frmreport.grid1.TextMatrix(frmreport.grid1.Row, 1) = ""
    frmreport.grid1.TextMatrix(frmreport.grid1.Row, 2) = ""
    frmreport.grid1.TextMatrix(frmreport.grid1.Row, 3) = ""
    frmreport.grid1.TextMatrix(frmreport.grid1.Row, 4) = ""
    frmreport.grid1.TextMatrix(frmreport.grid1.Row, 5) = ""
    frmreport.grid1.TextMatrix(frmreport.grid1.Row, 6) = ""
    frmreport.grid1.TextMatrix(frmreport.grid1.Row, 7) = ""
    frmreport.grid1.TextMatrix(frmreport.grid1.Row, 8) = ""
    frmreport.grid1.TextMatrix(frmreport.grid1.Row, 9) = ""
    frmreport.grid1.TextMatrix(frmreport.grid1.Row, 10) = ""
    frmreport.grid1.TextMatrix(frmreport.grid1.Row, 11) = ""
    frmreport.grid1.TextMatrix(frmreport.grid1.Row, 12) = ""
    frmreport.grid1.TextMatrix(frmreport.grid1.Row, 13) = ""
    Do While True
        If frmreport.grid1.TextMatrix(frmreport.grid1.Row + 1, 1) = "" Then
            frmreport.grid1.TextMatrix(frmreport.grid1.Row, 0) = ""
            frmreport.grid1.TextMatrix(frmreport.grid1.Row, 1) = ""
            frmreport.grid1.TextMatrix(frmreport.grid1.Row, 2) = ""
            frmreport.grid1.TextMatrix(frmreport.grid1.Row, 3) = ""
            frmreport.grid1.TextMatrix(frmreport.grid1.Row, 4) = ""
            frmreport.grid1.TextMatrix(frmreport.grid1.Row, 5) = ""
            frmreport.grid1.TextMatrix(frmreport.grid1.Row, 6) = ""
            frmreport.grid1.TextMatrix(frmreport.grid1.Row, 7) = ""
            frmreport.grid1.TextMatrix(frmreport.grid1.Row, 8) = ""
            frmreport.grid1.TextMatrix(frmreport.grid1.Row, 9) = ""
            frmreport.grid1.TextMatrix(frmreport.grid1.Row, 10) = ""
            frmreport.grid1.TextMatrix(frmreport.grid1.Row, 11) = ""
            frmreport.grid1.TextMatrix(frmreport.grid1.Row, 12) = ""
            frmreport.grid1.TextMatrix(frmreport.grid1.Row, 13) = ""
            Exit Do
        End If
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 0) = frmreport.grid1.TextMatrix(frmreport.grid1.Row + 1, 0)
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 1) = frmreport.grid1.TextMatrix(frmreport.grid1.Row + 1, 1)
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 2) = frmreport.grid1.TextMatrix(frmreport.grid1.Row + 1, 2)
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 3) = frmreport.grid1.TextMatrix(frmreport.grid1.Row + 1, 3)
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 4) = frmreport.grid1.TextMatrix(frmreport.grid1.Row + 1, 4)
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 5) = frmreport.grid1.TextMatrix(frmreport.grid1.Row + 1, 5)
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 6) = frmreport.grid1.TextMatrix(frmreport.grid1.Row + 1, 6)
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 7) = frmreport.grid1.TextMatrix(frmreport.grid1.Row + 1, 7)
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 8) = frmreport.grid1.TextMatrix(frmreport.grid1.Row + 1, 8)
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 9) = frmreport.grid1.TextMatrix(frmreport.grid1.Row + 1, 9)
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 10) = frmreport.grid1.TextMatrix(frmreport.grid1.Row + 1, 10)
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 11) = frmreport.grid1.TextMatrix(frmreport.grid1.Row + 1, 11)
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 12) = frmreport.grid1.TextMatrix(frmreport.grid1.Row + 1, 12)
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 13) = frmreport.grid1.TextMatrix(frmreport.grid1.Row + 1, 13)
        frmreport.grid1.Row = frmreport.grid1.Row + 1
    Loop
    frmreport.grid1.Rows = frmreport.grid1.Rows - 1
    frmreport.lbltotgroup = "Total Group : " & frmreport.grid1.Rows - 2
    
    For x = setup3 To 99
        For z = 0 To 100
            For y = 0 To 10
                myarray(x, y, z) = myarray(x + 1, y, z)
            Next y
        Next z
    Next x
    
    MsgBox "Data Group Report Is Deleted, Click OK To Continue ...", vbInformation, "Information"
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If cmbdk = "" Or txtgroup = "" Or txtdesc2 = "" Or cmbtype = "" Or cmbspace = "" Or cmbmode = "" Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If cmbcolom.Enabled = True And cmbcolom = "" Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If Len(Trim(txtgroup)) = 0 Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
        
    frmreport.grid1.Row = 1
    Do While txtgroup.Enabled = True
        If frmreport.grid1.TextMatrix(frmreport.grid1.Row, 0) = "" Then Exit Do
        
        If frmreport.grid1.TextMatrix(frmreport.grid1.Row, 0) = txtgroup Then
            MsgBox "Can't Add, Group Report " & txtgroup & " Already Exsist.", vbInformation, "Information"
            cmdclear_Click
            Exit Sub
        End If
        frmreport.grid1.Row = frmreport.grid1.Row + 1
    Loop
    
    If txtgroup.Enabled = True Then
        frmreport.grid1.Row = 1
        Do While True
            If frmreport.grid1.TextMatrix(frmreport.grid1.Row, 0) = "" Then
                frmreport.grid1.TextMatrix(frmreport.grid1.Row, 0) = txtgroup
                frmreport.grid1.TextMatrix(frmreport.grid1.Row, 1) = lbltype
                frmreport.grid1.TextMatrix(frmreport.grid1.Row, 2) = txtdesc2
                frmreport.grid1.TextMatrix(frmreport.grid1.Row, 3) = lblspace
                frmreport.grid1.TextMatrix(frmreport.grid1.Row, 4) = lblcolom
                frmreport.grid1.TextMatrix(frmreport.grid1.Row, 5) = lblmode
                frmreport.grid1.TextMatrix(frmreport.grid1.Row, 6) = txtsign
                frmreport.grid1.TextMatrix(frmreport.grid1.Row, 13) = cmbdk
                
                convword1
            
                frmreport.grid1.Rows = frmreport.grid1.Rows + 1
                Exit Do
            End If
            frmreport.grid1.Row = frmreport.grid1.Row + 1
        Loop
        frmreport.lbltotgroup = "Total Group : " & frmreport.grid1.Rows - 2
        MsgBox "Group Report Is Added, Click OK To Continue ...", vbInformation, "Information"
        
        txtgroup.Enabled = True
        cmbtype.Enabled = True
        cmbcolom.Enabled = True
        hapusgroup
        txtgroup.SetFocus
    Else
        frmreport.grid1.TextMatrix(setup3, 1) = lbltype
        frmreport.grid1.TextMatrix(setup3, 2) = txtdesc2
        frmreport.grid1.TextMatrix(setup3, 3) = lblspace
        frmreport.grid1.TextMatrix(setup3, 4) = lblcolom
        frmreport.grid1.TextMatrix(setup3, 5) = lblmode
        frmreport.grid1.TextMatrix(setup3, 6) = txtsign
        frmreport.grid1.TextMatrix(setup3, 13) = cmbdk
        
        convword2
        
        MsgBox "Group Report Is Updated, Click OK To Continue ...", vbInformation, "Information"
        Unload Me
    End If
End Sub

Private Sub convword2()
    If frmreport.lblapor = 1 And frmreport.grid1.TextMatrix(setup3, 1) = "0" Then
        frmreport.grid1.TextMatrix(setup3, 7) = "Header Only"
    ElseIf frmreport.lblapor = 1 And frmreport.grid1.TextMatrix(setup3, 1) = "1" Then
        frmreport.grid1.TextMatrix(setup3, 7) = "Assets"
    ElseIf frmreport.lblapor = 1 And frmreport.grid1.TextMatrix(setup3, 1) = "2" Then
        frmreport.grid1.TextMatrix(setup3, 7) = "Liability"
    ElseIf frmreport.lblapor = 1 And frmreport.grid1.TextMatrix(setup3, 1) = "3" Then
        frmreport.grid1.TextMatrix(setup3, 7) = "Capital"
    ElseIf frmreport.lblapor = 1 And frmreport.grid1.TextMatrix(setup3, 1) = "4" Then
        frmreport.grid1.TextMatrix(setup3, 7) = "Income Summary"
    ElseIf frmreport.lblapor = 2 And frmreport.grid1.TextMatrix(setup3, 1) = "0" Then
        frmreport.grid1.TextMatrix(setup3, 7) = "Header Only"
    ElseIf frmreport.lblapor = 2 And frmreport.grid1.TextMatrix(setup3, 1) = "1" Then
        frmreport.grid1.TextMatrix(setup3, 7) = "Income"
    ElseIf frmreport.lblapor = 2 And frmreport.grid1.TextMatrix(setup3, 1) = "2" Then
        frmreport.grid1.TextMatrix(setup3, 7) = "Expenses"
    ElseIf frmreport.lblapor = 3 And frmreport.grid1.TextMatrix(setup3, 1) = "0" Then
        frmreport.grid1.TextMatrix(setup3, 7) = "Header Only"
    ElseIf frmreport.lblapor = 3 And frmreport.grid1.TextMatrix(setup3, 1) = "1" Then
        frmreport.grid1.TextMatrix(setup3, 7) = "Assets"
    ElseIf frmreport.lblapor = 3 And frmreport.grid1.TextMatrix(setup3, 1) = "2" Then
        frmreport.grid1.TextMatrix(setup3, 7) = "Liability"
    ElseIf frmreport.lblapor = 3 And frmreport.grid1.TextMatrix(setup3, 1) = "3" Then
        frmreport.grid1.TextMatrix(setup3, 7) = "Capital"
    ElseIf frmreport.lblapor = 3 And frmreport.grid1.TextMatrix(setup3, 1) = "4" Then
        frmreport.grid1.TextMatrix(setup3, 7) = "Income"
    ElseIf frmreport.lblapor = 3 And frmreport.grid1.TextMatrix(setup3, 1) = "5" Then
        frmreport.grid1.TextMatrix(setup3, 7) = "Expenses"
    End If
        
    frmreport.grid1.TextMatrix(setup3, 8) = frmreport.grid1.TextMatrix(setup3, 2)
        
    If frmreport.grid1.TextMatrix(setup3, 3) = "0" Then
        frmreport.grid1.TextMatrix(setup3, 9) = "Title"
    ElseIf frmreport.grid1.TextMatrix(setup3, 3) = "1" Then
        frmreport.grid1.TextMatrix(setup3, 9) = "Space After 1"
    ElseIf frmreport.grid1.TextMatrix(setup3, 3) = "2" Then
        frmreport.grid1.TextMatrix(setup3, 9) = "Space After 2"
    ElseIf frmreport.grid1.TextMatrix(setup3, 3) = "3" Then
        frmreport.grid1.TextMatrix(setup3, 9) = "Space After 3"
    ElseIf frmreport.grid1.TextMatrix(setup3, 3) = "4" Then
        frmreport.grid1.TextMatrix(setup3, 9) = "Space After 4"
    ElseIf frmreport.grid1.TextMatrix(setup3, 3) = "5" Then
        frmreport.grid1.TextMatrix(setup3, 9) = "Eject"
    End If
        
    If frmreport.grid1.TextMatrix(setup3, 4) = "0" Then
        frmreport.grid1.TextMatrix(setup3, 10) = "Header"
    ElseIf frmreport.grid1.TextMatrix(setup3, 4) = "1" Then
        frmreport.grid1.TextMatrix(setup3, 10) = "Detail"
    ElseIf frmreport.grid1.TextMatrix(setup3, 4) = "2" Then
        frmreport.grid1.TextMatrix(setup3, 10) = "Sub Total"
    ElseIf frmreport.grid1.TextMatrix(setup3, 4) = "3" Then
        frmreport.grid1.TextMatrix(setup3, 10) = "Total"
    ElseIf frmreport.grid1.TextMatrix(setup3, 4) = "4" Then
        frmreport.grid1.TextMatrix(setup3, 10) = "Grand Total 1"
    ElseIf frmreport.grid1.TextMatrix(setup3, 4) = "5" Then
        frmreport.grid1.TextMatrix(setup3, 10) = "Grand Total 2"
    ElseIf frmreport.grid1.TextMatrix(setup3, 4) = "6" Then
        frmreport.grid1.TextMatrix(setup3, 10) = "Grand Total 3"
    ElseIf frmreport.grid1.TextMatrix(setup3, 4) = "7" Then
        frmreport.grid1.TextMatrix(setup3, 10) = "Grand Total 4"
    ElseIf frmreport.grid1.TextMatrix(setup3, 4) = "8" Then
        frmreport.grid1.TextMatrix(setup3, 10) = "Grand Total 5"
    ElseIf frmreport.grid1.TextMatrix(setup3, 4) = "9" Then
        frmreport.grid1.TextMatrix(setup3, 10) = ""
    End If
        
    If frmreport.grid1.TextMatrix(setup3, 5) = "0" Then
        frmreport.grid1.TextMatrix(setup3, 11) = "Normal"
    ElseIf frmreport.grid1.TextMatrix(setup3, 5) = "1" Then
        frmreport.grid1.TextMatrix(setup3, 11) = "Tebal"
    ElseIf frmreport.grid1.TextMatrix(setup3, 5) = "2" Then
        frmreport.grid1.TextMatrix(setup3, 11) = "Tebal & Garis Bawah"
    ElseIf frmreport.grid1.TextMatrix(setup3, 5) = "3" Then
        frmreport.grid1.TextMatrix(setup3, 11) = "Blok Text & Angka"
    ElseIf frmreport.grid1.TextMatrix(setup3, 5) = "4" Then
        frmreport.grid1.TextMatrix(setup3, 11) = "Blok Text, Angka Tebal"
    ElseIf frmreport.grid1.TextMatrix(setup3, 5) = "5" Then
        frmreport.grid1.TextMatrix(setup3, 11) = "Text Tebal, Blok Angka"
    End If
        
    frmreport.grid1.TextMatrix(setup3, 12) = frmreport.grid1.TextMatrix(setup3, 6)
End Sub

Private Sub convword1()
    If frmreport.lblapor = 1 And frmreport.grid1.TextMatrix(frmreport.grid1.Row, 1) = "0" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 7) = "Header Only"
    ElseIf frmreport.lblapor = 1 And frmreport.grid1.TextMatrix(frmreport.grid1.Row, 1) = "1" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 7) = "Assets"
    ElseIf frmreport.lblapor = 1 And frmreport.grid1.TextMatrix(frmreport.grid1.Row, 1) = "2" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 7) = "Liability"
    ElseIf frmreport.lblapor = 1 And frmreport.grid1.TextMatrix(frmreport.grid1.Row, 1) = "3" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 7) = "Capital"
    ElseIf frmreport.lblapor = 1 And frmreport.grid1.TextMatrix(frmreport.grid1.Row, 1) = "4" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 7) = "Income Summary"
    ElseIf frmreport.lblapor = 2 And frmreport.grid1.TextMatrix(frmreport.grid1.Row, 1) = "0" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 7) = "Header Only"
    ElseIf frmreport.lblapor = 2 And frmreport.grid1.TextMatrix(frmreport.grid1.Row, 1) = "1" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 7) = "Income"
    ElseIf frmreport.lblapor = 2 And frmreport.grid1.TextMatrix(frmreport.grid1.Row, 1) = "2" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 7) = "Expenses"
        
    ElseIf frmreport.lblapor = 3 And frmreport.grid1.TextMatrix(frmreport.grid1.Row, 1) = "0" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 7) = "Header Only"
    ElseIf frmreport.lblapor = 3 And frmreport.grid1.TextMatrix(frmreport.grid1.Row, 1) = "1" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 7) = "Assets"
    ElseIf frmreport.lblapor = 3 And frmreport.grid1.TextMatrix(frmreport.grid1.Row, 1) = "2" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 7) = "Liability"
    ElseIf frmreport.lblapor = 3 And frmreport.grid1.TextMatrix(frmreport.grid1.Row, 1) = "3" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 7) = "Capital"
    ElseIf frmreport.lblapor = 3 And frmreport.grid1.TextMatrix(frmreport.grid1.Row, 1) = "4" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 7) = "Income"
    ElseIf frmreport.lblapor = 3 And frmreport.grid1.TextMatrix(frmreport.grid1.Row, 1) = "5" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 7) = "Expenses"
    End If
        
    frmreport.grid1.TextMatrix(frmreport.grid1.Row, 8) = frmreport.grid1.TextMatrix(frmreport.grid1.Row, 2)
        
    If frmreport.grid1.TextMatrix(frmreport.grid1.Row, 3) = "0" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 9) = "Title"
    ElseIf frmreport.grid1.TextMatrix(frmreport.grid1.Row, 3) = "1" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 9) = "Space After 1"
    ElseIf frmreport.grid1.TextMatrix(frmreport.grid1.Row, 3) = "2" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 9) = "Space After 2"
    ElseIf frmreport.grid1.TextMatrix(frmreport.grid1.Row, 3) = "3" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 9) = "Space After 3"
    ElseIf frmreport.grid1.TextMatrix(frmreport.grid1.Row, 3) = "4" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 9) = "Space After 4"
    ElseIf frmreport.grid1.TextMatrix(frmreport.grid1.Row, 3) = "5" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 9) = "Eject"
    End If
        
    If frmreport.grid1.TextMatrix(frmreport.grid1.Row, 4) = "0" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 10) = "Header"
    ElseIf frmreport.grid1.TextMatrix(frmreport.grid1.Row, 4) = "1" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 10) = "Detail"
    ElseIf frmreport.grid1.TextMatrix(frmreport.grid1.Row, 4) = "2" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 10) = "Sub Total"
    ElseIf frmreport.grid1.TextMatrix(frmreport.grid1.Row, 4) = "3" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 10) = "Total"
    ElseIf frmreport.grid1.TextMatrix(frmreport.grid1.Row, 4) = "4" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 10) = "Grand Total 1"
    ElseIf frmreport.grid1.TextMatrix(frmreport.grid1.Row, 4) = "5" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 10) = "Grand Total 2"
    ElseIf frmreport.grid1.TextMatrix(frmreport.grid1.Row, 4) = "6" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 10) = "Grand Total 3"
    ElseIf frmreport.grid1.TextMatrix(frmreport.grid1.Row, 4) = "7" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 10) = "Grand Total 4"
    ElseIf frmreport.grid1.TextMatrix(frmreport.grid1.Row, 4) = "8" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 10) = "Grand Total 5"
    ElseIf frmreport.grid1.TextMatrix(frmreport.grid1.Row, 4) = "9" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 10) = ""
    End If
        
    If frmreport.grid1.TextMatrix(frmreport.grid1.Row, 5) = "0" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 11) = "Normal"
    ElseIf frmreport.grid1.TextMatrix(frmreport.grid1.Row, 5) = "1" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 11) = "Tebal"
    ElseIf frmreport.grid1.TextMatrix(frmreport.grid1.Row, 5) = "2" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 11) = "Tebal & Garis Bawah"
    ElseIf frmreport.grid1.TextMatrix(frmreport.grid1.Row, 5) = "3" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 11) = "Blok Text & Angka"
    ElseIf frmreport.grid1.TextMatrix(frmreport.grid1.Row, 5) = "4" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 11) = "Blok Text, Angka Tebal"
    ElseIf frmreport.grid1.TextMatrix(frmreport.grid1.Row, 5) = "5" Then
        frmreport.grid1.TextMatrix(frmreport.grid1.Row, 11) = "Text Tebal, Blok Angka"
    End If
        
    frmreport.grid1.TextMatrix(frmreport.grid1.Row, 12) = frmreport.grid1.TextMatrix(frmreport.grid1.Row, 6)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    cmbtype.Clear
    If setup2 = 1 Then
        cmbtype.AddItem "Header Only"
        cmbtype.AddItem "Assets"
        cmbtype.AddItem "Liability"
        cmbtype.AddItem "Capital"
        cmbtype.AddItem "Income Summary"
    ElseIf setup2 = 3 Then
        cmbtype.AddItem "Header Only"
        cmbtype.AddItem "Assets"
        cmbtype.AddItem "Liability"
        cmbtype.AddItem "Capital"
        cmbtype.AddItem "Income"
        cmbtype.AddItem "Expenses"
    Else
        cmbtype.AddItem "Header Only"
        cmbtype.AddItem "Income"
        cmbtype.AddItem "Expenses"
    End If
    lblcode1 = frmreport.txtreportcode
    lbldesc1 = frmreport.txtdesc1
    If setup2 = 1 Then
        lbltype1 = "Balance Sheet"
        cmbdk = "0"
        cmbdk.Enabled = False
    ElseIf setup2 = 2 Then
        lbltype1 = "Income Statement"
        cmbdk = "0"
        cmbdk.Enabled = False
    ElseIf setup2 = 3 Then
        lbltype1 = "Cash Flow"
    End If
    If setup3 <> "" Then
        txtgroup = frmreport.grid1.TextMatrix(setup3, 0)
        lbltype = frmreport.grid1.TextMatrix(setup3, 1)
        cmbtype.ListIndex = lbltype
        txtdesc2 = frmreport.grid1.TextMatrix(setup3, 2)
        lblspace = frmreport.grid1.TextMatrix(setup3, 3)
        cmbspace.ListIndex = lblspace
        lblcolom = frmreport.grid1.TextMatrix(setup3, 4)
        If lblcolom <> 9 Then cmbcolom.ListIndex = lblcolom
        lblmode = frmreport.grid1.TextMatrix(setup3, 5)
        cmbmode.ListIndex = lblmode
        txtsign = frmreport.grid1.TextMatrix(setup3, 6)
        If setup2 = 1 Then
            cmbdk = "0"
            cmbdk.Enabled = False
        ElseIf setup2 = 2 Then
            cmbdk = "0"
            cmbdk.Enabled = False
        ElseIf setup2 = 3 Then
            cmbdk = frmreport.grid1.TextMatrix(setup3, 13)
            cmbdk.Enabled = True
        End If
        txtgroup.Enabled = False
        
        If myarray(setup3, 0, 1) <> "" Then
            cmbtype.Enabled = False
            cmbcolom.Enabled = False
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    setup3 = ""
    frmreport.SSTab1.Tab = 1
End Sub

Private Sub txtdesc2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmbspace.SetFocus
End Sub

Private Sub txtgroup_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmbtype.SetFocus
End Sub

Private Sub txtsign_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And cmbdk.Enabled = True Then cmbdk.SetFocus
    If KeyAscii <> 42 And KeyAscii <> 8 Then KeyAscii = 0
End Sub
