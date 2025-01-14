VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmunposthutang 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Unposting Pembayaran Hutang"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtnobkt 
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
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
   Begin XtremeSuiteControls.PushButton cmdclose 
      Height          =   345
      Left            =   3240
      TabIndex        =   0
      Top             =   480
      Width           =   975
      _Version        =   851970
      _ExtentX        =   1720
      _ExtentY        =   609
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
   Begin Chameleon.chameleonButton cmdnobkt 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "No.Bukti"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmunposthutang.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin XtremeSuiteControls.PushButton cmdunpost 
      Height          =   345
      Left            =   2160
      TabIndex        =   3
      Top             =   480
      Width           =   975
      _Version        =   851970
      _ExtentX        =   1720
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "Unposting"
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
   Begin VB.Label lblapply 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   4335
   End
End
Attribute VB_Name = "frmunposthutang"
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

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdnobkt_Click()
    carisql1 = "Select a.NoBkt,c.NoApply,b.NamaSupp From am_apcashhdr a"
    carisql1 = carisql1 + " inner join am_supplier b on a.Kodesupp = b.KodeSupp"
    carisql1 = carisql1 + " left join am_apopnfil c on a.NoBkt = c.NoBeli"
    carisql1 = carisql1 + " Where a.Posted = '1'"
    
    namatabel = "Unposting Pembayaran Hutang"
    frmsearch.Show vbModal
End Sub

Private Sub cmdnobkt_GotFocus()
    If hasil = "" Then Exit Sub
    txtnobkt = hasil
    lblapply = "No.Apply : " & hasil1 & vbCrLf & "Supplier   : " & hasil2
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdunpost_Click()
    If MsgBox("Nomor Bukti akan di Unposting" & vbCrLf & "Klik tombol OK untuk melanjutkan", vbQuestion + vbOKCancel, "Konfirmasi") = vbCancel Then Clearform: Exit Sub
        OBJ.Open dsn
        SQL = "Update am_apcashhdr set Posted = '0' Where NoBkt = '" & txtnobkt & "'"
        OBJ.Execute SQL

        SQL = "Delete From gl_transaksi Where notrx = '" & txtnobkt & "'"
        OBJ.Execute SQL
        OBJ.Close
        
        MsgBox "No Bukti: " & txtnobkt & "Berhasil di unposting", vbInformation, AppName
        Call Clearform
End Sub
Private Sub Clearform()
    txtnobkt = ""
    lblapply = ""
End Sub

