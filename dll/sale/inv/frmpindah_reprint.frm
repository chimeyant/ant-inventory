VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmpindah_reprint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Pindah Gudang"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5415
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport crystal 
      Left            =   1695
      Top             =   4365
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
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   4410
      TabIndex        =   0
      Top             =   4380
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
      MICON           =   "frmpindah_reprint.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdview 
      Height          =   375
      Left            =   3450
      TabIndex        =   1
      Top             =   4380
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Print"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmpindah_reprint.frx":031A
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
      Left            =   1080
      TabIndex        =   2
      Top             =   105
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
      Format          =   194576387
      CurrentDate     =   37464
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   465
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
      Format          =   194576387
      CurrentDate     =   37464
   End
   Begin Chameleon.chameleonButton cmddown 
      Height          =   255
      Left            =   45
      TabIndex        =   4
      Top             =   825
      Width           =   5310
      _ExtentX        =   9366
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "66666"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Marlett"
         Size            =   9
         Charset         =   2
         Weight          =   500
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
      MICON           =   "frmpindah_reprint.frx":0634
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label8 
      Caption         =   "To :"
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
      Left            =   4650
      TabIndex        =   17
      Top             =   435
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label7 
      Caption         =   "From :"
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
      Left            =   3645
      TabIndex        =   16
      Top             =   435
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblto 
      Alignment       =   1  'Right Justify
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
      Left            =   4875
      TabIndex        =   15
      Top             =   435
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblfrom 
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
      Left            =   4245
      TabIndex        =   14
      Top             =   435
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblnobkt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3675
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000C&
      Height          =   3705
      Left            =   45
      Top             =   1110
      Width           =   5310
   End
   Begin VB.Label lblrow 
      Caption         =   "0 Baris."
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
      Left            =   225
      TabIndex        =   12
      Top             =   4410
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TO"
      BeginProperty Font 
         Name            =   "News706 BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4020
      TabIndex        =   11
      Top             =   1200
      Width           =   1260
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FROM"
      BeginProperty Font 
         Name            =   "News706 BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2895
      TabIndex        =   10
      Top             =   1200
      Width           =   1110
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TGL BUKTI"
      BeginProperty Font 
         Name            =   "News706 BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1470
      TabIndex        =   9
      Top             =   1200
      Width           =   1410
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NO BUKTI"
      BeginProperty Font 
         Name            =   "News706 BT"
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
      TabIndex        =   8
      Top             =   1200
      Width           =   1335
   End
   Begin MSForms.ListBox ListBox 
      Height          =   2880
      Left            =   120
      TabIndex        =   7
      Top             =   1470
      Width           =   5175
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "9128;5080"
      ColumnCount     =   4
      MultiSelect     =   1
      SpecialEffect   =   3
      FontName        =   "Tahoma"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
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
      Left            =   120
      TabIndex        =   6
      Top             =   495
      Width           =   1215
   End
   Begin VB.Label Label2 
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
      Left            =   120
      TabIndex        =   5
      Top             =   135
      Width           =   1215
   End
End
Attribute VB_Name = "frmpindah_reprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim i As Integer

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmddown_Click()
    ListBox.Clear
    ListBox.ColumnCount = 5
    ListBox.ColumnWidths = "2.5 cm; 3 cm; 1 cm; 1 cm; 1 cm"
    
    If date1 > Date Then Exit Sub
    i = 0
    OBJ.Open dsn
    SQL = "Select distinct a.nobpb,a.tglbpb,b.Dari,c.Ke "
    SQL = SQL + "From am_bpbhdr a inner join (select b.nobpb,b.kodegudang 'Dari' from am_bpbhdr b Where type = '99')b "
    SQL = SQL + "on a.nobpb = b.nobpb inner join (select c.nobpb,c.kodegudang 'Ke' from am_bpbhdr c Where type = '88')c "
    SQL = SQL + "on a.nobpb = c.nobpb Where a.tglbpb >= '" & batas1 & "' and a.tglbpb <= '" & batas2 & "' "
    SQL = SQL + "and a.nobpb like 'PG0%' Order By a.nobpb"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        ListBox.AddItem RST!nobpb
        ListBox.List(i, 1) = RST!tglbpb
        ListBox.List(i, 2) = RST!Dari
        ListBox.List(i, 3) = "->"
        ListBox.List(i, 4) = RST!Ke
        i = i + 1
        RST.MoveNext
    Loop
    OBJ.Close
    lblrow = ListBox.ListCount & " Baris."
End Sub

Private Sub cmdview_Click()
    If ListBox.ListCount = 0 Then Exit Sub
                    
    For i = 0 To ListBox.ListCount - 1
        If ListBox.Selected(i) = True Then
            ListBox.Selected(i) = False
            frmpindahshow.txtnobkt = ListBox.List(i, 0)
            frmpindahshow.Show vbModal
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
Function batas1()
    batas1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function batas2()
    batas2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function

'Private Sub ListBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If ListBox.ListCount = 0 Then Exit Sub
                    
'    For i = 0 To ListBox.ListCount - 1
'        If ListBox.Selected(i) = True Then
'            ListBox.Selected(i) = False
'            lblnobkt = ListBox.List(i, 0)
'            lblfrom = ListBox.List(i, 2)
'            lblto = ListBox.List(i, 4)
'        End If
'    Next i
'End Sub
