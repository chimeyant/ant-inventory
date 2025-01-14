VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Troubleshooting"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   8865
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picrespon 
      Height          =   1815
      Left            =   960
      ScaleHeight     =   1755
      ScaleWidth      =   7755
      TabIndex        =   18
      Top             =   2640
      Visible         =   0   'False
      Width           =   7815
      Begin XtremeSuiteControls.RadioButton optproses 
         Height          =   255
         Left            =   4920
         TabIndex        =   22
         Top             =   120
         Width           =   855
         _Version        =   851970
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Proses"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.TextBox txtnote 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   600
         Width           =   7560
      End
      Begin VB.TextBox txtkode 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3000
         TabIndex        =   20
         Top             =   120
         Width           =   1725
      End
      Begin XtremeSuiteControls.PushButton cmdcari 
         Height          =   375
         Left            =   1560
         TabIndex        =   19
         Top             =   120
         Width           =   1305
         _Version        =   851970
         _ExtentX        =   2302
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Find ITG form"
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
      Begin XtremeSuiteControls.RadioButton optcancel 
         Height          =   255
         Left            =   5880
         TabIndex        =   23
         Top             =   120
         Width           =   855
         _Version        =   851970
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cancel"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optclose 
         Height          =   255
         Left            =   6840
         TabIndex        =   24
         Top             =   120
         Width           =   855
         _Version        =   851970
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Close"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdupdate 
         Height          =   390
         Left            =   120
         TabIndex        =   26
         Top             =   1320
         Width           =   7545
         _Version        =   851970
         _ExtentX        =   13309
         _ExtentY        =   688
         _StockProps     =   79
         Caption         =   "Update"
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
      Begin VB.Label Label8 
         Caption         =   "Catatan :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   900
      End
   End
   Begin VB.ComboBox cmbdepartemen 
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
      ItemData        =   "frmts.frx":0000
      Left            =   1200
      List            =   "frmts.frx":0002
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtnama 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1185
      TabIndex        =   0
      Top             =   480
      Width           =   2565
   End
   Begin VB.TextBox txthal 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5385
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   810
      Width           =   3405
   End
   Begin VB.TextBox txtket 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1560
      Width           =   8640
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   5400
      TabIndex        =   2
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
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
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   134348803
      CurrentDate     =   38767
   End
   Begin XtremeSuiteControls.PushButton cmdclose 
      Height          =   375
      Left            =   7800
      TabIndex        =   6
      Top             =   4800
      Width           =   945
      _Version        =   851970
      _ExtentX        =   1667
      _ExtentY        =   661
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
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton cmdsave 
      Height          =   375
      Left            =   6840
      TabIndex        =   5
      Top             =   4800
      Width           =   945
      _Version        =   851970
      _ExtentX        =   1667
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Save"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton cmdclear 
      Height          =   375
      Left            =   5880
      TabIndex        =   16
      Top             =   4800
      Width           =   945
      _Version        =   851970
      _ExtentX        =   1667
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Clear"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton cmdrespon 
      Height          =   375
      Left            =   3960
      TabIndex        =   17
      Top             =   4800
      Visible         =   0   'False
      Width           =   1785
      _Version        =   851970
      _ExtentX        =   3149
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Respon"
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
   Begin VB.Label lblnourut 
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
      Left            =   1200
      TabIndex        =   15
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label7 
      Caption         =   "No. Urut"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   14
      Top             =   120
      Width           =   900
   End
   Begin VB.Label lbluser 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   960
      TabIndex        =   13
      Top             =   4800
      Width           =   1860
   End
   Begin VB.Label Label6 
      Caption         =   "Pemohon :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   12
      Top             =   4800
      Width           =   900
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "   Permasalahan :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   0
      TabIndex        =   11
      Top             =   1320
      Width           =   8895
   End
   Begin VB.Label Label4 
      Caption         =   "Perihal"
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
      Left            =   4440
      TabIndex        =   10
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Pemohon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   255
      TabIndex        =   9
      Top             =   510
      Width           =   900
   End
   Begin VB.Label Label3 
      Caption         =   "Departemen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   8
      Top             =   975
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "Tanggal"
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
      Left            =   4440
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "frmts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private Sub cmdcari_Click()
    carisql1 = "select pemohon,keterangan,kdts,tanggal from am_troubleshoot where status='0'"
    namatabel = "Formulir ITG"
    frmsearch.Show vbModal
End Sub

Private Sub cmdcari_GotFocus()
    If hasil = "" Then Exit Sub
    txtnama = hasil
    txtket = hasil1
    txtkode = hasil2
    
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdclear_Click()
    Dim strformat As String
    Dim strno As String
    strformat = Format(Date, "yymm")
    txtnama = ""
    txthal = ""
    txtket = ""
    txtkode = ""
    txtnote = ""
    optproses.Value = False: optcancel.Value = False: optclose.Value = False
    date1 = Date
    cmbdepartemen = ""
    picrespon.Visible = False
    OBJ.Open dsn
    SQL = "select top 1 kdts from am_troubleshoot where kdts like 'ITG-' + '" + strformat + "%' order by kdts desc"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        strno = Right(RST!kdts, 4)
    Else
        strno = 0
    End If
    OBJ.Close
    
    strno = strno + 1

    If Len(strno) = 1 Then lblnourut = "ITG-" & strformat & "000" & strno
    If Len(strno) = 2 Then lblnourut = "ITG-" & strformat & "00" & strno
    If Len(strno) = 3 Then lblnourut = "ITG-" & strformat & "0" & strno
    If Len(strno) = 4 Then lblnourut = "ITG-" & strformat & strno
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdrespon_Click()
    picrespon.Visible = True
End Sub

Private Sub cmdsave_Click()
    Dim strformat As String
    Dim strno As String
    strformat = Format(Date, "yymm")
    Dim nourut As String
    
    If txtnama = "" Or txthal = "" Or txtket = "" Or cmbdepartemen.text = "" Then
        MsgBox "Mohon lengkapi data pada kolom isian", vbCritical, AppName
        Exit Sub
    End If
        
    If MsgBox("Ajukan form troubleshooting ?", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select top 1 kdts from am_troubleshoot where kdts like 'ITG-' + '" + strformat + "%' order by kdts desc"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        strno = Right(RST!kdts, 4)
    Else
        strno = 0
    End If
    
    strno = strno + 1

    If Len(strno) = 1 Then nourut = "ITG-" & strformat & "000" & strno
    If Len(strno) = 2 Then nourut = "ITG-" & strformat & "00" & strno
    If Len(strno) = 3 Then nourut = "ITG-" & strformat & "0" & strno
    If Len(strno) = 4 Then nourut = "ITG-" & strformat & strno
    
    SQL = "Select * From am_troubleshoot Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    
    RST.AddNew
    RST!kdts = nourut
    RST!pemohon = txtnama
    RST!departement = cmbdepartemen
    RST!hal = txthal
    RST!tanggal = date1
    RST!keterangan = txtket
    RST!status = "0"
    RST!catatan = ""
    RST.Update
    
    OBJ.Close
    MsgBox "Permohonan berhasil diajukan", vbInformation, AppName
    cmdclear_Click
End Sub

Private Sub cmdupdate_Click()
    If optproses.Value = False And optcancel.Value = False And optclose.Value = False Then
        MsgBox "Status belum dipilih", vbExclamation, AppName
        Exit Sub
    End If
    
    OBJ.Open dsn
    If optproses.Value = True Then
        SQL = "Update am_troubleshoot set status='1',catatan='" & txtnote & "' Where kdts='" & txtkode & "'"
    ElseIf optcancel.Value = True Then
        SQL = "Update am_troubleshoot set status='2',catatan='" & txtnote & "' Where kdts='" & txtkode & "'"
    ElseIf optclose.Value = True Then
        SQL = "Update am_troubleshoot set status='3',catatan='" & txtnote & "' Where kdts='" & txtkode & "'"
    End If
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    MsgBox "Berhasil disimpan", vbInformation, AppName
    cmdclear_Click
End Sub

Private Sub Form_Load()
    Dim strformat As String
    Dim strno As String
    strformat = Format(Date, "yymm")
    date1 = Date
    
    cmbdepartemen.additem "Pembelian"
    cmbdepartemen.additem "Marketing"
    cmbdepartemen.additem "Finance & Accounting"
    cmbdepartemen.additem "HRD"
    cmbdepartemen.additem "Produksi Lem"
    cmbdepartemen.additem "Produksi Karet"
    cmbdepartemen.additem "Maintenance"
    cmbdepartemen.additem "Laboratorium & QC"
    cmbdepartemen.additem "Gudang"
    cmbdepartemen.additem "Sekretariat ISO"
    lbluser = UserOnline
    If UserOnline = "Creator" Then cmdrespon.Visible = True
    
    OBJ.Open dsn
    SQL = "select top 1 kdts from am_troubleshoot where kdts like 'ITG-' + '" + strformat + "%' order by kdts desc"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        strno = Right(RST!kdts, 4)
    Else
        strno = 0
    End If
    OBJ.Close
    
    strno = strno + 1

    If Len(strno) = 1 Then lblnourut = "ITG-" & strformat & "000" & strno
    If Len(strno) = 2 Then lblnourut = "ITG-" & strformat & "00" & strno
    If Len(strno) = 3 Then lblnourut = "ITG-" & strformat & "0" & strno
    If Len(strno) = 4 Then lblnourut = "ITG-" & strformat & strno
    
    
End Sub

