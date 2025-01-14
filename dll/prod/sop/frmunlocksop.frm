VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Begin VB.Form frmunlocksop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unlock SOP"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4560
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtketerangan 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   720
      Width           =   3255
   End
   Begin VB.TextBox txtlot 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   3255
   End
   Begin Chameleon.chameleonButton cmdunlock 
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   2160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "UnLock"
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
      MICON           =   "frmunlocksop.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsop 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "No Lot"
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
      MICON           =   "frmunlocksop.frx":031A
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
      Left            =   3600
      TabIndex        =   5
      Top             =   2160
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
      MICON           =   "frmunlocksop.frx":0634
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      Caption         =   "Reason :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   780
      Width           =   975
   End
End
Attribute VB_Name = "frmunlocksop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private OBJ As New ADODB.Connection
Private RST As ADODB.Recordset
Private SQL As String

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdsop_Click()
    namatabel = "nolot"
    carisql1 = "Select a.kode_produk,a.nama_produk,b.nolot from list_produk_master a "
    carisql1 = carisql1 + "inner join list_produksi_master b on a.kode_produk=b.kode_produk"
    frmsearch.Show vbModal
End Sub

Private Sub cmdsop_GotFocus()
    txtlot = hasil2
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    carisql1 = ""
    'txtketerangan.SetFocus
End Sub

Private Sub cmdunlock_Click()
    If txtlot = "" Then Exit Sub
    If txtketerangan = "" Then
        MsgBox "Mohon kolom keterangan diisi dengan alasan edit sop", vbCritical, AppName
        Exit Sub
    End If
    OBJ.Open dsn
    SQL = "select * from list_masterkeyLot where noso='" & txtlot & "'"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    If RST.EOF Then
        'start tidak ada data input (sementara)
        With RST
            .AddNew
            !noso = txtlot
            !tgl = Date
            !UserName = nmuser
            !otoritas = "1"
            !cetaksop = "0"
            !otorisasi = nmuser
            !keterangan = txtketerangan
            .Update
        End With
    Else
        SQL = "Update list_masterkeyLot set otoritas='1',otorisasi='" & nmuser & "',keterangan = '" & txtketerangan & "'"
        SQL = SQL + " Where noso ='" & txtlot & "'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "Update list_produksi_master set flagprint = '3' Where nolot='" & txtlot & "'"
        Set RST = OBJ.Execute(SQL)
        
    End If
    OBJ.Close
    MsgBox "SOP is successfully unlocked ", vbInformation, AppName
    txtlot = ""
    txtketerangan = ""
End Sub
