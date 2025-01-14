VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Begin VB.Form frmsjalandelete 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete Surat Jalan"
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2295
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmsjalandelete.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   2295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TDBNumber6Ctl.TDBNumber txtbaris 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   615
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   503
      Calculator      =   "frmsjalandelete.frx":2372
      Caption         =   "frmsjalandelete.frx":2392
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmsjalandelete.frx":23FE
      Keys            =   "frmsjalandelete.frx":241C
      Spin            =   "frmsjalandelete.frx":2466
      AlignHorizontal =   2
      AlignVertical   =   2
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "####0;;0"
      EditMode        =   1
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "####0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   10
      MinValue        =   1
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   1245189
      Value           =   1
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.Label Label2 
      Caption         =   "Baris ke -"
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
      TabIndex        =   1
      Top             =   150
      Width           =   855
   End
End
Attribute VB_Name = "frmsjalandelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdview_Click()
    frmsjalanedit.txtbaris = txtbaris
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Me.Caption = "Baris 1 s/d " & frmsjalanedit.Grid.Rows - 2
End Sub

Private Sub txtbaris_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
