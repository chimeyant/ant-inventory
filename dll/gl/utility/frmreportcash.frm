VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Begin VB.Form frmreportcash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6540
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmreportcash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtacc1 
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
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox txtacc2 
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
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox txtacc3 
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
      Left            =   4800
      MaxLength       =   15
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox txtacc4 
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
      Left            =   4800
      MaxLength       =   15
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox txtacc5 
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
      Left            =   4800
      MaxLength       =   15
      TabIndex        =   5
      Top             =   2040
      Width           =   1455
   End
   Begin VB.ComboBox cmblineno 
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
      ItemData        =   "frmreportcash.frx":2372
      Left            =   1440
      List            =   "frmreportcash.frx":2385
      TabIndex        =   0
      Top             =   1200
      Width           =   615
   End
   Begin Chameleon.chameleonButton cmdsave 
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   2520
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
      MICON           =   "frmreportcash.frx":2398
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
      Left            =   3480
      TabIndex        =   7
      Top             =   2520
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
      MICON           =   "frmreportcash.frx":26B2
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
      Left            =   4440
      TabIndex        =   8
      Top             =   2520
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
      MICON           =   "frmreportcash.frx":29CC
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
      Left            =   5400
      TabIndex        =   9
      Top             =   2520
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
      MICON           =   "frmreportcash.frx":2CE6
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
      Left            =   240
      TabIndex        =   18
      Top             =   1680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Account #1"
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
      MICON           =   "frmreportcash.frx":3000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch2 
      Height          =   285
      Left            =   240
      TabIndex        =   19
      Top             =   2040
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Account #2"
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
      MICON           =   "frmreportcash.frx":331A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch3 
      Height          =   285
      Left            =   3600
      TabIndex        =   20
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Account #3"
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
      MICON           =   "frmreportcash.frx":3634
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
      Left            =   3600
      TabIndex        =   21
      Top             =   1680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Account #4"
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
      MICON           =   "frmreportcash.frx":394E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch5 
      Height          =   285
      Left            =   3600
      TabIndex        =   22
      Top             =   2040
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Account #5"
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
      MICON           =   "frmreportcash.frx":3C68
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
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
      TabIndex        =   16
      Top             =   630
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
      TabIndex        =   15
      Top             =   150
      Width           =   735
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
      TabIndex        =   14
      Top             =   390
      Width           =   4215
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      Caption         =   "Column #"
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
      TabIndex        =   13
      Top             =   1230
      Width           =   1095
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
      TabIndex        =   12
      Top             =   150
      Width           =   975
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
      TabIndex        =   11
      Top             =   630
      Width           =   1035
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
      TabIndex        =   10
      Top             =   390
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   -120
      TabIndex        =   17
      Top             =   0
      Width           =   7815
   End
End
Attribute VB_Name = "frmreportcash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private Sub cariline()
    hapusdetail
    
    OBJ.Open dsn
    SQL = "select * from gl_cforms where form_no = '" & lblcode1 & "' and line_no = '" & cmblineno & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtacc1 = original(RST!acc_no1)
        txtacc2 = original(RST!acc_no2)
        txtacc3 = original(RST!acc_no3)
        txtacc4 = original(RST!acc_no4)
        txtacc5 = original(RST!acc_no5)
    End If
    OBJ.Close
End Sub

Private Sub cmblineno_Click()
    If cmblineno = "" Then Exit Sub
    cariline
End Sub

Private Sub cmdclear_Click()
    If cmblineno.Enabled = False Then Exit Sub
    cmblineno.Enabled = True
    cmblineno = ""
    cmblineno.SetFocus
    hapusdetail
End Sub

Private Sub hapusdetail()
    txtacc1 = ""
    txtacc2 = ""
    txtacc3 = ""
    txtacc4 = ""
    txtacc5 = ""
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdelete_Click()
    If cmblineno = "" Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If MsgBox("Are You Sure Want To Delete ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "delete from gl_cforms where form_no = '" & lblcode1 & "' and line_no = '" & cmblineno & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Column Cash Flow Is Deleted, Click OK To Continue ...", vbInformation, "Information"
    
    cmdclear_Click
    Unload Me
End Sub

Private Sub cmdsearch1_Click()
    If cmblineno = "" Then Exit Sub
    'namatabel = "Cash/Bank"
    'setup1 = "0"
    'setup2 = "z"
    cektype
    carisql1 = "select noac, nmac from gl_masterac"
    namatabel = "Account "
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc1 = hasil
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdsearch2_Click()
    If cmblineno = "" Then Exit Sub
    'namatabel = "Cash/Bank"
    'setup1 = "0"
    'setup2 = "z"
    cektype
    carisql1 = "select noac, nmac from gl_masterac"
    namatabel = "Account "
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch2_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc2 = hasil
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdsearch3_Click()
    If cmblineno = "" Then Exit Sub
    cektype
    carisql1 = "select noac, nmac from gl_masterac"
    namatabel = "Account "
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch3_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc3 = hasil
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdsearch4_Click()
    If cmblineno = "" Then Exit Sub
    cektype
    carisql1 = "select noac, nmac from gl_masterac"
    namatabel = "Account "
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch4_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc4 = hasil
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdsearch5_Click()
    If cmblineno = "" Then Exit Sub
    cektype
    carisql1 = "select noac, nmac from gl_masterac"
    namatabel = "Account "
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch5_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc5 = hasil
    hasil = ""
    hasil1 = ""
End Sub

Private Sub cmdSave_Click()
    If cmblineno = "" Or (txtacc1 = "" And txtacc2 = "" And txtacc3 = "" And txtacc4 = "" And txtacc5 = "") Then
        MsgBox "Data Entry Not Complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "delete from gl_cforms where form_no = '" & lblcode1 & "' and line_no = '" & cmblineno & "'"
    Set RST = OBJ.Execute(SQL)
                                    
    SQL = "insert into gl_cforms"
    SQL = SQL + "(form_no"
    SQL = SQL + ",line_no"
    SQL = SQL + ",acc_no1"
    SQL = SQL + ",acc_no2"
    SQL = SQL + ",acc_no3"
    SQL = SQL + ",acc_no4"
    SQL = SQL + ",acc_no5)"

    SQL = SQL + "VALUES"
    SQL = SQL + "('" & lblcode1 & "'"
    SQL = SQL + ", '" & cmblineno & "'"
    SQL = SQL + ", '" & x_original(txtacc1) & "'"
    SQL = SQL + ", '" & x_original(txtacc2) & "'"
    SQL = SQL + ", '" & x_original(txtacc3) & "'"
    SQL = SQL + ", '" & x_original(txtacc4) & "'"
    SQL = SQL + ", '" & x_original(txtacc5) & "')"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Column Cash Flow Is Updated, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    lblcode1 = frmreport.txtreportcode
    lbldesc1 = frmreport.txtdesc1
    lbltype1 = "Cash Flow"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    setup1 = ""
    frmreport.SSTab1.Tab = 0
    frmreport.SSTab1.Tab = 1
End Sub

Private Sub txtacc1_KeyPress(KeyAscii As Integer)
    If cmblineno = "" Then
        KeyAscii = 0
        Exit Sub
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And txtacc2.Enabled = True Then txtacc2.SetFocus
    If KeyAscii = 13 And txtacc2.Enabled = False Then cmdsave.SetFocus
End Sub

Private Sub txtacc1_LostFocus()
    If txtacc1 = "" Or cmblineno = "" Then Exit Sub
    OBJ.Open dsn
    cektype
    If cmblineno >= -9 And cmblineno <= -1 Then
        SQL = "select * from gl_masterac where noac like '" & x_original(txtacc1) & "%' and (typeac = '" & setup5 & "')"
    Else
        SQL = "select * from gl_masterac where noac = '" & x_original(txtacc1) & "' and (typeac = '" & setup5 & "')"
    End If
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Account " & txtacc1 & " Not Found.", vbInformation, "Information"
        txtacc1 = ""
        txtacc1.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub txtacc2_KeyPress(KeyAscii As Integer)
    If cmblineno = "" Then
        KeyAscii = 0
        Exit Sub
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtacc3.SetFocus
End Sub

Private Sub txtacc2_LostFocus()
    If txtacc2 = "" Then Exit Sub
    cektype
    OBJ.Open dsn
    SQL = "select * from gl_masterac where noac = '" & x_original(txtacc2) & "' and (typeac = '" & setup5 & "')"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Account " & txtacc2 & " Not Found.", vbInformation, "Information"
        txtacc2 = ""
        txtacc2.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub txtacc3_KeyPress(KeyAscii As Integer)
    If cmblineno = "" Then
        KeyAscii = 0
        Exit Sub
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtacc4.SetFocus
End Sub

Private Sub txtacc3_LostFocus()
    If txtacc3 = "" Then Exit Sub
    cektype
    OBJ.Open dsn
    SQL = "select * from gl_masterac where noac = '" & x_original(txtacc3) & "' and (typeac = '" & setup5 & "')"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Account " & txtacc3 & " Not Found.", vbInformation, "Information"
        txtacc3 = ""
        txtacc3.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub txtacc4_KeyPress(KeyAscii As Integer)
    If cmblineno = "" Then
        KeyAscii = 0
        Exit Sub
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtacc5.SetFocus
End Sub

Private Sub txtacc4_LostFocus()
    If txtacc4 = "" Then Exit Sub
    cektype
    OBJ.Open dsn
    SQL = "select * from gl_masterac where noac = '" & x_original(txtacc4) & "' and (typeac = '" & setup5 & "')"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Account " & txtacc4 & " Not Found.", vbInformation, "Information"
        txtacc4 = ""
        txtacc4.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub txtacc5_KeyPress(KeyAscii As Integer)
    If cmblineno = "" Then
        KeyAscii = 0
        Exit Sub
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmdsave.SetFocus
End Sub

Private Sub txtacc5_LostFocus()
    If txtacc5 = "" Then Exit Sub
    cektype
    OBJ.Open dsn
    SQL = "select * from gl_masterac where noac = '" & x_original(txtacc5) & "' and (typeac = '" & setup5 & "')"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Account " & txtacc5 & " Not Found.", vbInformation, "Information"
        txtacc5 = ""
        txtacc5.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub cektype()
    If setup2 = 1 Then
        If setup4 = "1" Then
            setup5 = "AS"
        ElseIf setup4 = "2" Then
            setup5 = "LI"
        ElseIf setup4 = "3" Then
            setup5 = "CA' or typeac = 'IS"
        ElseIf setup4 = "4" Then
            setup5 = "IS' or typeac = 'CA"
        End If
    ElseIf setup2 = 3 Then
        If setup4 = "1" Then
            setup5 = "AS"
        ElseIf setup4 = "2" Then
            setup5 = "LI"
        ElseIf setup4 = "3" Then
            setup5 = "CA"
        ElseIf setup4 = "4" Then
            setup5 = "IN"
        ElseIf setup4 = "5" Then
            setup5 = "EX"
        End If
    Else
        If setup4 = "1" Then
            setup5 = "IN"
        ElseIf setup4 = "2" Then
            setup5 = "EX"
        End If
    End If
End Sub
