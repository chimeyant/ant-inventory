VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Begin VB.Form frmpackaging 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Packaging Bahan Baku"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4920
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4920
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtkdpackaging 
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
      Height          =   285
      Left            =   2085
      MaxLength       =   4
      TabIndex        =   2
      Top             =   135
      Width           =   615
   End
   Begin VB.TextBox txtqty 
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
      Height          =   285
      Left            =   2085
      MaxLength       =   20
      TabIndex        =   1
      Top             =   840
      Width           =   630
   End
   Begin VB.TextBox txtnama 
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
      Height          =   285
      Left            =   2085
      MaxLength       =   20
      TabIndex        =   0
      Top             =   495
      Width           =   2655
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   3870
      TabIndex        =   3
      Top             =   1305
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
      MICON           =   "frmpackaging.frx":0000
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
      Left            =   2910
      TabIndex        =   4
      Top             =   1305
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Clear"
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
      MICON           =   "frmpackaging.frx":031A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdadd 
      Height          =   375
      Left            =   1950
      TabIndex        =   5
      Top             =   1305
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Save"
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
      MICON           =   "frmpackaging.frx":0634
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lbldesc 
      Caption         =   "Nama Packaging"
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
      Left            =   45
      TabIndex        =   8
      Top             =   525
      Width           =   1935
   End
   Begin VB.Label lblsatcode 
      Caption         =   "Kode Packaging Bahan Baku"
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
      Left            =   45
      TabIndex        =   7
      Top             =   180
      Width           =   2010
   End
   Begin VB.Label Label2 
      Caption         =   "Qty / Packaging"
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
      Left            =   30
      TabIndex        =   6
      Top             =   855
      Width           =   1935
   End
End
Attribute VB_Name = "frmpackaging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private Sub cmdclear_Click()
    txtkdpackaging = ""
    txtqty = ""
    txtnama = ""
    txtkdpackaging.SetFocus
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdadd_Click()
    If Len(Trim(txtkdpackaging)) = 0 Then
        MsgBox "Data entry not Complete.", vbExclamation, "Warning"
        txtkdpackaging.SetFocus
        Exit Sub
    End If
    
    If txtqty = "" Or txtkdpackaging = "" Or txtnama = "" Then
        MsgBox "Data entry not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    txtkdpackaging = Trim(txtkdpackaging)
    
    OBJ.Open dsn
    SQL = "SELECT * FROM am_appackaging WHERE kdpckg = '" & txtkdpackaging & " '"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Data already exist.", vbInformation, "Information"
        cmdclear_Click
        OBJ.Close
        Exit Sub
    End If
    SQL = "INSERT INTO am_appackaging"
    SQL = SQL + "(kdpckg"
    SQL = SQL + ",nmpckg"
    SQL = SQL + ",qty)"
    
    SQL = SQL + "VALUES"
    SQL = SQL + " ('" & txtkdpackaging & "'"
    SQL = SQL + ", '" & txtnama & "'"
    SQL = SQL + ", convert(money,'" & txtqty & "'))"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Data saved, click ok to continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    txtkdpackaging.ToolTipText = "max length = " & txtkdpackaging.MaxLength
    txtnama.ToolTipText = "max length = " & txtnama.MaxLength
    txtqty.ToolTipText = "max length = " & txtqty.MaxLength
End Sub

Private Sub txtNama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtqty.SetFocus
End Sub

Private Sub txtkdpackaging_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtnama.SetFocus
End Sub

Private Sub txtkdpackaging_LostFocus()
    If txtkdpackaging = "" Then Exit Sub
    If txtkdpackaging.SelLength <> 0 Then Exit Sub
    OBJ.Open dsn
    SQL = "SELECT * FROM am_appackaging WHERE kdpckg = '" & txtkdpackaging & " '"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtqty = RST!qty
        txtnama = RST!nmpckg
            
        MsgBox "Data already exist.", vbInformation, "Information"
        txtkdpackaging.SetFocus
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    txtqty = ""
    txtnama = ""
    txtnama.SetFocus
End Sub

