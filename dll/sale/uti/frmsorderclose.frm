VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmsorderclose 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Closing Sales Order"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
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
   Icon            =   "frmsorderclose.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtket 
      Appearance      =   0  'Flat
      Height          =   525
      Left            =   1560
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1920
      Width           =   7575
   End
   Begin VB.TextBox txtsales 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7920
      MaxLength       =   10
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtkodecust 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7920
      MaxLength       =   10
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtnobukti 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
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
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   91291651
      CurrentDate     =   37426
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   8280
      TabIndex        =   7
      Top             =   2520
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
      MICON           =   "frmsorderclose.frx":2372
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
      Left            =   7320
      TabIndex        =   6
      Top             =   2520
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
      MICON           =   "frmsorderclose.frx":268C
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
      Left            =   6120
      TabIndex        =   5
      Top             =   2520
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Closing SO"
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
      MICON           =   "frmsorderclose.frx":29A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch 
      Height          =   285
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "No Sales Order"
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
      MICON           =   "frmsorderclose.frx":2CC0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label3 
      Caption         =   "Keterangan Closing"
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Salesman"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Customer"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblalamatcust 
      BackColor       =   &H80000014&
      Height          =   255
      Left            =   1560
      TabIndex        =   12
      Top             =   1200
      Width           =   7575
   End
   Begin VB.Label lblsales 
      BackColor       =   &H80000014&
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   1560
      Width           =   7575
   End
   Begin VB.Label lblnamacust 
      BackColor       =   &H80000014&
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   840
      Width           =   7575
   End
   Begin VB.Label Label13 
      Caption         =   "Tanggal"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   510
      Width           =   1455
   End
End
Attribute VB_Name = "frmsorderclose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private Sub cmdadd_Click()
    If txtnobukti = "" Or txtsales = "" Or txtkodecust = "" Or txtket = "" Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select flag2 from am_soapp where noso = '" & txtnobukti & "' and flag2 = '9'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        MsgBox "Sales Order tidak bisa diClosing, Sales Order sudah di Cancel.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    SQL = "select * from am_soclose where noso = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        MsgBox "Closing aborted, SO already closing.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    SQL = "select * from am_period where tanggal1 <= '" & tanggalinv & "' and tanggal2 >= '" & tanggalinv & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        OBJ.Close
        MsgBox "Can not closing, out of date range.", vbExclamation, "Warning"
        Exit Sub
    End If
    OBJ.Close
    
    If MsgBox("Are you sure want to closing Sales Order ?", vbQuestion + vbYesNo, "Question") = vbNo Then
        cmdclear_Click
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "insert into am_soclose ("
    SQL = SQL + "noso,"
    SQL = SQL + "keterangan,"
    SQL = SQL + "identry,"
    SQL = SQL + "dateentry)"
    
    SQL = SQL + " values("
    SQL = SQL + "'" & txtnobukti & "',"
    SQL = SQL + "'" & txtket & "',"
    SQL = SQL + "'" & kuser & "',"
    SQL = SQL + "convert(datetime,'" & tanggalsekarang & "'))"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Data Is Updated, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdclear_Click()
    hapusemua
    
    txtnobukti.Enabled = True
    cmdsearch.Enabled = True
    date1.Enabled = True
    txtnobukti = ""
    txtnobukti.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdsearch_Click()
    If v_fastsearching = True Then
        If v_fstgl1 > v_fstgl2 Then
            MsgBox "Invalid date range, search abort.", vbExclamation, "Error"
            Exit Sub
        End If
        carisql1 = "select distinct noso, convert(char(11),tglso )'tglso' from am_soapp where tglso >= '" & batas1 & "' and tglso <= '" & batas2 & "'"
    Else
        carisql1 = "select distinct noso, convert(char(11),tglso )'tglso' from am_soapp"
    End If
    namatabel = "Sales Order"

    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    txtnobukti = hasil
    carinvoice
    hasil = ""
    hasil1 = ""
    txtket.SetFocus
End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='154' and b.kodeuser = '1" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If RST.EOF Then cmdadd.Enabled = False
    '    OBJ.Close
        
    '    If cmdadd.Enabled = False Then
    '        MsgBox "User Rights Denied !!" & vbCrLf & _
    '        "Please contact your Administrator.", vbCritical, "User Rights"
        
    '        Unload Me
    '    End If
    'End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    
    
    date1.Value = Date
End Sub

Private Sub txtnobukti_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then date1.SetFocus
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtnobukti_LostFocus()
    carinvoice
End Sub

Function tanggalinv()
      tanggalinv = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggalsekarang()
      tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function

Function batas1()
    batas1 = Month(v_fstgl1) & "/" & Day(v_fstgl1) & "/" & Year(v_fstgl1)
End Function

Function batas2()
    batas2 = Month(v_fstgl2) & "/" & Day(v_fstgl2) & "/" & Year(v_fstgl2)
End Function

Private Sub hapusemua()
    date1 = Date
    txtkodecust = ""
    lblnamacust = ""
    lblalamatcust = ""
    txtsales = ""
    lblsales = ""
    txtket = ""
End Sub

Private Sub carinvoice()
    If txtnobukti = "" Then Exit Sub
    If txtnobukti.SelLength <> 0 Then Exit Sub
    
    hapusemua

    OBJ.Open dsn
    SQL = "select * from am_soapp where noso = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        date1 = RST!tglso
        txtkodecust = RST!kodecust
        txtsales = RST!kodesales

        SQL = "select * from am_customer where kodecust = '" & txtkodecust & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            lblnamacust = RST!namacust
            lblalamatcust = RST!alamatcust
        End If

        SQL = "select * from am_salesman where kodesales = '" & txtsales & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then lblsales = RST!namasales
        
        SQL = "select * from am_soclose where noso = '" & txtnobukti & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then txtket = RST!keterangan
        
        txtnobukti.Enabled = False
        cmdsearch.Enabled = False
        date1.Enabled = False
    Else
        MsgBox "Data not found.", vbExclamation, "Warning"
        txtnobukti = ""
        txtnobukti.SetFocus
    End If
    OBJ.Close
End Sub
