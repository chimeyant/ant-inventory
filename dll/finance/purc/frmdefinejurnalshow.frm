VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmdefinejurnalshow 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Define Journal"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2895
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
   Icon            =   "frmdefinejurnalshow.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   2655
      Begin MSForms.ComboBox cmbtype 
         Height          =   300
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   2175
         VariousPropertyBits=   612386843
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3836;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         DropButtonStyle =   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   2655
      Begin VB.OptionButton Option9 
         Caption         =   "D/K"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Kredit"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Debet"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   2655
      Begin VB.OptionButton Option3 
         Caption         =   "Jurnal C (Hutang Dagang)"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Jurnal B (PPn Masukan)"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Jurnal A (Persediaan)"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Cancel"
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
      MICON           =   "frmdefinejurnalshow.frx":2372
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmdefinejurnalshow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private Sub cmbtype_Click()
    hasil = cmbtype
    Unload Me
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 27 Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    OBJ.Open dsn
    SQL = "select distinct substring(kodebarang,1,3)'a' from am_beliapp order by a"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        cmbtype.AddItem RST!a
        
        RST.MoveNext
    Loop
    OBJ.Close
    
    cmbtype.AddItem "Supp"
    cmbtype.AddItem "one"
    cmbtype.AddItem "Bank"
End Sub

Private Sub Option1_Click()
    hasil = "Jurnal a"
    Unload Me
End Sub

Private Sub Option2_Click()
    hasil = "Jurnal b"
    Unload Me
End Sub

Private Sub Option3_Click()
    hasil = "Jurnal c"
    Unload Me
End Sub

Private Sub Option6_Click()
    hasil = "D"
    Unload Me
End Sub

Private Sub Option7_Click()
    hasil = "K"
    Unload Me
End Sub

Private Sub Option9_Click()
    hasil = "D/K"
    Unload Me
End Sub
