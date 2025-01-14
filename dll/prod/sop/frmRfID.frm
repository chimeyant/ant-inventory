VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmRfID 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set ID Card"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8790
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   2220
      Left            =   90
      TabIndex        =   7
      Top             =   930
      Width           =   4020
      _Version        =   851970
      _ExtentX        =   7091
      _ExtentY        =   3916
      _StockProps     =   79
      Caption         =   "Input"
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
      Begin VB.TextBox txtkocekan 
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
         Left            =   1200
         TabIndex        =   10
         Top             =   1245
         Width           =   885
      End
      Begin VB.TextBox txtopt 
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
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1695
         Width           =   2370
      End
      Begin VB.TextBox txtnolot 
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
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   810
         Width           =   2370
      End
      Begin VB.TextBox txtidcard 
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
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   375
         Width           =   2370
      End
      Begin XtremeSuiteControls.PushButton btnGetData 
         Height          =   315
         Left            =   270
         TabIndex        =   12
         Top             =   360
         Width           =   870
         _Version        =   851970
         _ExtentX        =   1535
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "ID Card"
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
      Begin XtremeSuiteControls.PushButton btnopt 
         Height          =   315
         Left            =   285
         TabIndex        =   14
         ToolTipText     =   "Tambah Operator"
         Top             =   1695
         Width           =   870
         _Version        =   851970
         _ExtentX        =   1535
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Operator"
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
      Begin XtremeSuiteControls.PushButton cmdcari 
         Height          =   315
         Left            =   285
         TabIndex        =   15
         ToolTipText     =   "Tambah Operator"
         Top             =   810
         Width           =   870
         _Version        =   851970
         _ExtentX        =   1535
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "No. Lot"
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Kocekan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   16
         Top             =   1260
         Width           =   870
      End
   End
   Begin XtremeSuiteControls.PushButton btnInit 
      Height          =   435
      Left            =   975
      TabIndex        =   3
      Top             =   495
      Width           =   1020
      _Version        =   851970
      _ExtentX        =   1799
      _ExtentY        =   767
      _StockProps     =   79
      Caption         =   "Initialize"
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
   Begin VB.ComboBox cbReader 
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
      ItemData        =   "frmRfID.frx":0000
      Left            =   960
      List            =   "frmRfID.frx":0002
      Style           =   1  'Simple Combo
      TabIndex        =   1
      Text            =   "cbReader"
      Top             =   120
      Width           =   3135
   End
   Begin RichTextLib.RichTextBox rbOutput 
      Height          =   3105
      Left            =   4305
      TabIndex        =   2
      Top             =   45
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   5477
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmRfID.frx":0004
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.PushButton btnConnect 
      Height          =   435
      Left            =   3060
      TabIndex        =   4
      Top             =   495
      Width           =   1020
      _Version        =   851970
      _ExtentX        =   1799
      _ExtentY        =   767
      _StockProps     =   79
      Caption         =   "Connect"
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
   Begin XtremeSuiteControls.PushButton btnReset 
      Height          =   435
      Left            =   1965
      TabIndex        =   5
      Top             =   3195
      Width           =   1020
      _Version        =   851970
      _ExtentX        =   1799
      _ExtentY        =   767
      _StockProps     =   79
      Caption         =   "Reset"
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
   Begin XtremeSuiteControls.PushButton btnQuit 
      Height          =   435
      Left            =   7695
      TabIndex        =   6
      Top             =   3195
      Width           =   1020
      _Version        =   851970
      _ExtentX        =   1799
      _ExtentY        =   767
      _StockProps     =   79
      Caption         =   "Quit"
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
   Begin XtremeSuiteControls.PushButton btnsave 
      Height          =   435
      Left            =   3075
      TabIndex        =   13
      Top             =   3195
      Width           =   1020
      _Version        =   851970
      _ExtentX        =   1799
      _ExtentY        =   767
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
      Appearance      =   5
   End
   Begin VB.Label Label1 
      Caption         =   "Reader"
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
      Left            =   210
      TabIndex        =   0
      Top             =   150
      Width           =   720
   End
End
Attribute VB_Name = "frmRfID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private OBJ As New ADODB.Connection
Private RST As ADODB.Recordset
Private SQL As String

Dim retCode, Protocol, hContext, hCard, ReaderCount As Long
Dim sReaderList As String * 256
Dim sReaderGroup As String
Dim ioRequest As SCARD_IO_REQUEST
Dim SendLen, RecvLen As Long
Dim SendBuff(0 To 255) As Byte
Dim RecvBuff(0 To 255) As Byte
Dim validATS As Boolean

Private Sub btnConnect_Click()
    'Connect to reader using a shared connection
    retCode = SCardConnect(hContext, _
                           cbReader.text, _
                           SCARD_SHARE_SHARED, _
                           SCARD_PROTOCOL_T0 Or SCARD_PROTOCOL_T1, _
                           hCard, _
                           Protocol)
                           
    If retCode <> SCARD_S_SUCCESS Then
        Call DisplayOut(GetScardErrMsg(retCode), 2)
        Exit Sub
    Else
        Call DisplayOut("Successful connection to " & cbReader.text, 1)
    End If
    
    btnConnect.Enabled = False
    'Ambil IDCard
    btnGetData_Click
End Sub

Private Sub btnGetData_Click()
    Dim index As Integer
    Dim tempstr As String
    
    validATS = False
    Call ClearBuffers
    SendBuff(0) = &HFF
    SendBuff(1) = &HCA
    
    
    SendBuff(3) = &H0
    SendBuff(4) = &H0
    
    SendLen = 5
    RecvLen = &HFF
    
    retCode = SendAPDU(2)
    If retCode <> SCARD_S_SUCCESS Then
        Exit Sub
    End If
    
    'Interpret and display return values
    If validATS Then
        For index = 0 To RecvLen - 3
            'tempstr = tempstr & Right$("00" & Hex(RecvBuff(index)), 2) & " "
            'tempstr = tempstr & Right$("00" & RecvBuff(index), 2)
            tempstr = tempstr & RecvBuff(index)
        Next index
        
        Call DisplayOut(Trim(tempstr), 4)
        txtidcard = tempstr
    End If
    
End Sub

Private Sub btnInit_Click()
    Dim index As Integer
    
    For index = 0 To 255
        sReaderList = sReaderList & vbNullChar
    Next index
    
    ReaderCount = 255
    
    'Establish context
    retCode = SCardEstablishContext(SCARD_SCOPE_USER, 0, 0, hContext)
    
    If retCode <> SCARD_S_SUCCESS Then
        Call DisplayOut(GetScardErrMsg(retCode), 2)
        Exit Sub
    End If
    
    'List readers
    Call cbReader.Clear
    
    retCode = SCardListReaders(hContext, _
                               sReaderGroup, _
                               sReaderList, _
                               ReaderCount)
                               
    If retCode <> SCARD_S_SUCCESS Then
        Call DisplayOut(GetScardErrMsg(retCode), 2)
        Exit Sub
    End If
    
    'Add readers to combobox control
    Call LoadListToControl(cbReader, sReaderList)
    cbReader.ListIndex = 0
    
    'Set default reader to ACR122 NFC Reader
    For index = 0 To cbReader.ListIndex - 1
        cbReader.ListIndex = index
        
        If InStr(cbReader.text, "ACR122U") > 0 Then
            Exit For
        End If
    Next index
    
    btnConnect.Enabled = True
    btnInit.Enabled = False
End Sub

Public Sub DisplayOut(ByVal out As String, ByVal mode As Integer)

    Select Case mode
        Case 1
            rbOutput.SelColor = vbBlue
            
        Case 2
            rbOutput.SelColor = vbRed
            
        Case 3
            rbOutput.SelColor = vbBlack
            out = "<< " & out
            
        Case 4
            rbOutput.SelColor = vbBlack
            out = ">> " & out
    End Select
    
    rbOutput.SelText = out & vbCrLf
    rbOutput.SelStart = Len(rbOutput.text)
    rbOutput.SelColor = vbBlack
End Sub

Private Sub btnopt_Click()
    frmRfID_opt.Show vbModal
End Sub

Private Sub btnQuit_Click()
    retCode = SCardDisconnect(hCard, SCARD_UNPOWER_CARD)
    retCode = SCardReleaseContext(hContext)
    Unload Me
End Sub

Private Sub btnReset_Click()
    rbOutput.text = ""
    txtidcard = ""
    txtopt = ""
    txtnolot = ""
    txtkocekan = ""
    
    retCode = SCardDisconnect(hCard, SCARD_UNPOWER_CARD)
    retCode = SCardReleaseContext(hContext)
    Call Initialize
End Sub

Private Sub btnsave_Click()
'Proses Simpan Ke tabel produksi
    If txtidcard = "" Then
        MsgBox "Kolom RFID tidak boleh kosong", vbCritical, AppName
        Exit Sub
    End If
    If txtnolot = "" Then
        MsgBox "Kolom NoLot tidak boleh kosong", vbCritical, AppName
        Exit Sub
    End If
    If txtopt = "" Then
        MsgBox "Kolom Operator tidak boleh kosong", vbCritical, AppName
        Exit Sub
    End If
    OBJ.Open dsn
    SQL = "Select * From produksi Where 0=1"
    Set RST = New ADODB.Recordset
    RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    With RST
        .AddNew
        !nolot = txtnolot
        !kocekan = txtkocekan
        !noid = txtidcard
        !operator = txtopt
        !date_entry = Format(Date, "yyyy/MM/dd")
        !Status = "1" 'Aktif
        .Update
    End With
    OBJ.Close
    MsgBox "Data berhasil disimpan", vbInformation, AppName
    btnReset_Click
    btnInit_Click
End Sub

Private Sub cmdcari_Click()
namatabel = "nolot"
    carisql1 = "Select a.kode_produk,a.nama_produk,b.nolot from list_produk_master a "
    carisql1 = carisql1 + "inner join list_produksi_master b on a.kode_produk=b.kode_produk"
    carisql1 = carisql1 + " where b.flagprint <> '4'"
    frmsearch.Show vbModal
End Sub

Private Sub cmdcari_GotFocus()
    If hasil = "" Then Exit Sub
    txtnolot = hasil2
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub Form_Load()
    Call Initialize
End Sub

Public Sub Initialize()
    cbReader.text = ""
    btnConnect.Enabled = False

    btnInit.Enabled = True

    Call DisplayOut("Program ready", 1)
End Sub
Public Sub ClearBuffers()
    Dim index As Integer
    
    For index = 0 To 255
        RecvBuff(index) = &H0
        SendBuff(index) = &H0
    Next index
End Sub
Public Function SendAPDU(ByVal mode As Integer)
    Dim index As Integer
    Dim tempstr As String
    
    ioRequest.dwProtocol = Protocol
    ioRequest.cbPciLength = Len(ioRequest)
    
    tempstr = ""
    
    For index = 0 To SendLen - 1
        tempstr = tempstr & Right$("00" & Hex(SendBuff(index)), 2) & " "
    Next index
    
    Call DisplayOut(tempstr, 3)
    
    retCode = SCardTransmit(hCard, _
                            ioRequest, _
                            SendBuff(0), _
                            SendLen, _
                            ioRequest, _
                            RecvBuff(0), _
                            RecvLen)
                            
    If retCode <> SCARD_S_SUCCESS Then
        Call DisplayOut(GetScardErrMsg(retCode), 2)
        SendAPDU = retCode
        Exit Function
    End If
    
    tempstr = ""
    
    Select Case mode
        Case 1
            For index = 0 To RecvLen - 1
                tempstr = tempstr & Right$("00" & Hex(RecvBuff(index)), 2) & " "
            Next index
              
        Case 2
            For index = RecvLen - 2 To RecvLen - 1
                If InStr(Hex(RecvBuff(index)), "A") = 2 Then
                    tempstr = tempstr & Hex(RecvBuff(index))
                Else
                    tempstr = tempstr & Right$("00" & Hex(RecvBuff(index)), 2)
                End If
            Next index
            
            If tempstr = "6A81" Then
                Call DisplayOut("This function is not supported", 2)
                SendAPDU = retCode
                Exit Function
            End If
            
            validATS = True
    End Select
    
    Call DisplayOut(tempstr, 4)
    SendAPDU = retCode
    
End Function

Private Sub txtidcard_Change()
On Error GoTo Err_handler:
    If txtidcard <> "" Then
        OBJ.Open dsn
        'CEK KARTU YANG APAKAH MASIH DIGUNAKAN (PROSES)
        SQL = "Select * From produksi Where noid = '" & txtidcard & "' and status = '1'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Kartu masih digunakan dalam proses produksi" & _
            vbCrLf & "No Lot : " & RST!nolot & _
            vbCrLf & "Operator : " & RST!operator, vbExclamation, "WARNING"
            OBJ.Close

            btnReset_Click
            btnInit_Click
            Exit Sub
        End If
        OBJ.Close
    End If
    Exit Sub
Err_handler:
    If OBJ.State = 1 Then OBJ.Close
    MsgBox Err.Description, vbCritical, AppName
End Sub
