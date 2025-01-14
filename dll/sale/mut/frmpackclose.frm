VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Begin VB.Form frmpackclose 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Open /Close Lot permintaan packaging"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6510
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   6510
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtdeletelot 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4560
      Locked          =   -1  'True
      MaxLength       =   17
      TabIndex        =   11
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox txtnolotpending 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      Locked          =   -1  'True
      MaxLength       =   17
      TabIndex        =   6
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox txtnolot 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   17
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   2040
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
      MICON           =   "frmpackclose.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdlot 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Search Close Lot"
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
      MICON           =   "frmpackclose.frx":031A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdopen 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Open Lot"
      ENAB            =   0   'False
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
      MICON           =   "frmpackclose.frx":0634
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
      Left            =   120
      TabIndex        =   4
      Top             =   2040
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
      MICON           =   "frmpackclose.frx":094E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdpending 
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Search Pending Lot"
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
      MICON           =   "frmpackclose.frx":0C68
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdcloselot 
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   840
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Manual Close Lot"
      ENAB            =   0   'False
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
      MICON           =   "frmpackclose.frx":0F82
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmddellot 
      Height          =   375
      Left            =   4560
      TabIndex        =   10
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Search Lot"
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
      MICON           =   "frmpackclose.frx":129C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmddeletelot 
      Height          =   375
      Left            =   4560
      TabIndex        =   12
      Top             =   840
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Delete Lot"
      ENAB            =   0   'False
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
      MICON           =   "frmpackclose.frx":15B6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Menghapus Lot maka Cetakan 1,2,3 ..dst.. akan terhapus"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   4560
      TabIndex        =   13
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      Height          =   1935
      Left            =   4320
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Close pending lot permintaan packaging"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2400
      TabIndex        =   9
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Buka lot permintaan packaging"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   1935
      Left            =   2160
      Top             =   0
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   1935
      Left            =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmpackclose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private Sub cmdclear_Click()
    txtnolot = ""
    txtnolotpending = ""
    txtdeletelot = ""
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdcloselot_Click()
    If txtnolotpending = "" Then Exit Sub
    If MsgBox("Apakah Lot permintaan barang sudah sesuai dengan SOP", vbQuestion + vbYesNo, "Konfirmasi Close Manual") = vbNo Then Exit Sub
    OBJ.Open dsn
    SQL = "UPDATE am_gudang_permintaan SET flag ='2' Where nolot = '" & txtnolotpending & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    MsgBox "lot was closed manually.", vbInformation, AppName
End Sub

Private Sub cmddeletelot_Click()
    If txtdeletelot = "" Then Exit Sub
    If MsgBox("Yakin ingin menghapus lot ini ?", vbQuestion + vbYesNo, AppName) = vbNo Then Exit Sub
    OBJ.Open dsn
    SQL = "Delete From am_gudang_permintaan Where nolot = '" & txtdeletelot & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    MsgBox "Lot has been deleted", vbInformation, AppName
    cmdclear_Click
End Sub

Private Sub cmddellot_Click()
    carisql1 = "Select distinct nolot From am_gudang_permintaan"
    namatabel = "Lot Delete"
    frmsearch.Show vbModal
End Sub

Private Sub cmddellot_GotFocus()
    If hasil = "" Then Exit Sub
    txtdeletelot = hasil
    hasil = ""
    carisql1 = ""
    namatabel = ""
End Sub

Private Sub cmdlot_Click()
    carisql1 = "Select distinct nolot From am_gudang_permintaan"
    namatabel = "Lot Packaging"
    frmsearch.Show vbModal
End Sub

Private Sub cmdlot_GotFocus()
    If hasil = "" Then Exit Sub
    txtnolot = hasil
    hasil = ""
    carisql1 = ""
    namatabel = ""
End Sub

Private Sub cmdopen_Click()
    If txtnolot = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "UPDATE am_gudang_permintaan SET flag ='1' Where nolot = '" & txtnolot & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    MsgBox "Packaging Lot has opened", vbInformation, AppName
    cmdclear_Click
End Sub

Private Sub cmdpending_Click()
    carisql1 = "Select distinct nolot From am_gudang_permintaan"
    namatabel = "Lot Packpending"
    frmsearch.Show vbModal
End Sub

Private Sub cmdpending_GotFocus()
    If hasil = "" Then Exit Sub
    txtnolotpending = hasil
    hasil = ""
    carisql1 = ""
    namatabel = ""
End Sub

Private Sub txtdeletelot_Change()
    If txtdeletelot <> "" Then
        cmddeletelot.Enabled = True
    Else
        cmddeletelot.Enabled = False
    End If
End Sub

Private Sub txtnolot_Change()
    If txtnolot <> "" Then
        cmdopen.Enabled = True
    Else
        cmdopen.Enabled = False
    End If
End Sub

Private Sub txtnolotpending_Change()
    If txtnolotpending <> "" Then
        cmdcloselot.Enabled = True
    Else
        cmdcloselot.Enabled = False
    End If
End Sub
