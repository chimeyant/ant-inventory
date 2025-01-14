VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Begin VB.Form frminvoicedesc2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete Description"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frminvoicedesc2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtnodo1 
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
      Height          =   765
      Left            =   1080
      MaxLength       =   500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   840
      Width           =   4815
   End
   Begin Chameleon.chameleonButton cmdview 
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   1800
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "OK"
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
      MICON           =   "frminvoicedesc2.frx":2372
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   1800
      Width           =   855
      _ExtentX        =   1508
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
      MICON           =   "frminvoicedesc2.frx":268C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblupdateby 
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
      Left            =   1080
      TabIndex        =   1
      Top             =   510
      Width           =   3615
   End
   Begin VB.Label lblnoinv 
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
      Left            =   1080
      TabIndex        =   2
      Top             =   150
      Width           =   3615
   End
   Begin VB.Label Label3 
      Caption         =   "Deleted by"
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
      Top             =   510
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "No Invoice"
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
      TabIndex        =   4
      Top             =   150
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Comment"
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
      TabIndex        =   3
      Top             =   870
      Width           =   1095
   End
End
Attribute VB_Name = "frminvoicedesc2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim OBJ1 As New ADODB.Connection
Dim RST1 As New ADODB.Recordset
Dim SQL1 As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdview_Click()
    If txtnodo1 = "" Then
        MsgBox "User must supply description.", vbInformation, "information"
        
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_aropnfil where nobkt = '" & lblnoinv & "' and transtype = 'I'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        OBJ1.Open dsn
        SQL1 = "insert into am_invdelete ("
        SQL1 = SQL1 + "kodecust, "
        SQL1 = SQL1 + "nobkt, "
        SQL1 = SQL1 + "tglbkt, "
        SQL1 = SQL1 + "transtype, "
        SQL1 = SQL1 + "keterangan, "
        SQL1 = SQL1 + "amount, "
        SQL1 = SQL1 + "potongan, "
        SQL1 = SQL1 + "ppn, "
        SQL1 = SQL1 + "selisih, "
        SQL1 = SQL1 + "kodecur, "
        SQL1 = SQL1 + "nilaikurs, "
        SQL1 = SQL1 + "iddelete, "
        SQL1 = SQL1 + "datedelete, "
        SQL1 = SQL1 + "reason)"
    
        SQL1 = SQL1 + " values ('" & RST!kodecust & "',"
        SQL1 = SQL1 + "'" & lblnoinv & "',"
        SQL1 = SQL1 + "convert(datetime,'" & Format(RST!tglbkt, "MM/dd/yyyy") & "'),"
        SQL1 = SQL1 + "'I',"
        SQL1 = SQL1 + "'" & RST!keterangan & "',"
        SQL1 = SQL1 + "convert(money,'" & RST!amount & "'),"
        SQL1 = SQL1 + "convert(money,'" & RST!potongan & "'),"
        SQL1 = SQL1 + "convert(money,'" & RST!ppn & "'),"
        SQL1 = SQL1 + "convert(money,'" & RST!selisih & "'),"
        SQL1 = SQL1 + "'" & RST!kodecur & "',"
        SQL1 = SQL1 + "convert(money,'" & RST!nilaikurs & "'),"
        SQL1 = SQL1 + "'" & kuser & "',"
        SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
        SQL1 = SQL1 + "'" & txtnodo1 & "')"
        Set RST1 = OBJ1.Execute(SQL1)
        OBJ1.Close
        
        RST.MoveNext
    Loop
    OBJ.Close
    
    ops_tf = True
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    lblnoinv = frminvoicedit.txtnobukti
    lblupdateby = nmuser
End Sub

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function
