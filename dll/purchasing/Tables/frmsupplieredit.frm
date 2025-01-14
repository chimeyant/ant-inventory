VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Begin VB.Form frmsupplieredit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Supplier"
   ClientHeight    =   3330
   ClientLeft      =   5715
   ClientTop       =   5520
   ClientWidth     =   5895
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
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
      Height          =   2760
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CheckBox chk1 
      Caption         =   "Ya"
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
      Left            =   1680
      TabIndex        =   9
      Top             =   2490
      Width           =   615
   End
   Begin VB.OptionButton ops2 
      Caption         =   "Umum"
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
      Left            =   3000
      TabIndex        =   8
      Top             =   2160
      Width           =   855
   End
   Begin VB.OptionButton ops1 
      Caption         =   "Bahan Baku"
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
      Left            =   1680
      TabIndex        =   7
      Top             =   2160
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox txtkode 
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
      Left            =   120
      MaxLength       =   10
      TabIndex        =   19
      Top             =   3000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtalamat1 
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
      Left            =   1680
      MaxLength       =   40
      TabIndex        =   3
      Top             =   750
      Width           =   3975
   End
   Begin VB.TextBox txtelp 
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
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   5
      Top             =   1440
      Width           =   3975
   End
   Begin VB.TextBox txtfax 
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
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   6
      Top             =   1800
      Width           =   3975
   End
   Begin VB.TextBox txtkontak 
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
      Left            =   1680
      MaxLength       =   30
      TabIndex        =   4
      Top             =   1080
      Width           =   3975
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
      Left            =   1680
      MaxLength       =   40
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
   Begin VB.TextBox txtalamat 
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
      Left            =   1680
      MaxLength       =   40
      TabIndex        =   2
      Top             =   480
      Width           =   3975
   End
   Begin Chameleon.chameleonButton cmdsearch0 
      Height          =   285
      Left            =   240
      TabIndex        =   14
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Supplier"
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
      MICON           =   "frmsupplieredit.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   4800
      TabIndex        =   13
      Top             =   2850
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
      MICON           =   "frmsupplieredit.frx":031A
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
      Left            =   3840
      TabIndex        =   12
      Top             =   2850
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
      MICON           =   "frmsupplieredit.frx":0634
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmddel 
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   2850
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Delete"
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
      MICON           =   "frmsupplieredit.frx":094E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdupdate 
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   2850
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Update"
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
      MICON           =   "frmsupplieredit.frx":0C68
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label4 
      Caption         =   "Pengusaha kena Pajak"
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
      Left            =   240
      TabIndex        =   21
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Category Supplier"
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
      TabIndex        =   20
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Telephone"
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
      TabIndex        =   18
      Top             =   1470
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Faxsimile"
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
      TabIndex        =   17
      Top             =   1830
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Contact Person"
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
      TabIndex        =   16
      Top             =   1110
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Alamat Supplier"
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
      TabIndex        =   15
      Top             =   510
      Width           =   1215
   End
End
Attribute VB_Name = "frmsupplieredit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private Sub cmdclear_Click()
    cmdsearch0.Enabled = True
    txtkode = ""
    txtnama = ""
    txtalamat = ""
    txtalamat1 = ""
    txtelp = ""
    txtfax = ""
    txtkontak = ""
    chk1.Value = 0
    ops1.Value = True
    txtnama.SetFocus
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmddel_Click()
    OBJ.Open dsn
    SQL = "select * from am_pohdr"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        OBJ.Close
        MsgBox "User can not delete supplier.", vbExclamation, "Warning"
        Exit Sub
    End If
    OBJ.Close
    
    If txtkode = "" Or txtnama = "" Or txtalamat = "" Then
        MsgBox "Data entry not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    txtnama = Trim(txtnama)
    
    OBJ.Open dsn
    SQL = "select namasupp from am_supplier where namasupp = '" & txtnama & "' and kodesupp <> '" & txtkode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Data already exist.", vbInformation, "Information"
        txtnama.SetFocus
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    If MsgBox("Are you sure want to delete ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from AM_pohdr WHERE Kodesupp = '" & txtkode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then GoTo jump1
    
    SQL = "select * from AM_apopnfil WHERE Kodesupp = '" & txtkode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then GoTo jump1
    
    SQL = "DELETE FROM AM_Supplier WHERE Kodesupp = '" & txtkode & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Data deleted, click ok to continue ...", vbInformation, "Information"
    cmdclear_Click
    
    Exit Sub
    
jump1:
    OBJ.Close
    MsgBox "Can not delete, data in use.", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdsearch0_Click()
    carisql1 = "select namasupp, AlamatSupp1, kodesupp from am_supplier"
    namatabel = "Supplier"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch0_GotFocus()
    If hasil = "" Then Exit Sub
    txtkode = hasil2
    carisupplier
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    txtnama.SetFocus
End Sub

Private Sub cmdupdate_click()
    If txtkode = "" Or txtnama = "" Or txtalamat = "" Then
        MsgBox "Data entry not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    txtnama = Trim(txtnama)
    
    If MsgBox("Are You Sure Want To Update ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "UPDATE AM_Supplier SET "
    SQL = SQL + "NamaSupp = '" & txtnama & "',"
    SQL = SQL + "AlamatSupp1 = '" & txtalamat & "',"
    SQL = SQL + "AlamatSupp2 = '" & txtalamat1 & "',"
    SQL = SQL + "telpsupp = '" & txtelp & "',"
    SQL = SQL + "faxsupp = '" & txtfax & "',"
    If ops1.Value = True Then SQL = SQL + "category = '1'," Else SQL = SQL + "category = '2',"
    SQL = SQL + "wp = '" & chk1.Value & "',"
    SQL = SQL + "contactperson = '" & txtkontak & "'"
    SQL = SQL + " WHERE KodeSupp = '" & txtkode & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Data updated, click ok to continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='12' and b.kodeuser = '2" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If RST.EOF Then cmdupdate.Enabled = False
        
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='13' and b.kodeuser = '2" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If RST.EOF Then cmddel.Enabled = False
    '    OBJ.Close
    'End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 27 Then
        List1.Visible = False
        If txtkode = "" Then Exit Sub
        
        OBJ.Open dsn
        SQL = "select * from am_supplier where kodesupp = '" & txtkode & "'"
        Set RST = OBJ.Execute(SQL)
        hasil = "q"
        If Not RST.EOF Then txtnama = RST!namasupp
        OBJ.Close
        hasil = ""
    End If
End Sub

Private Sub Form_Load()
   
    
    txtnama.ToolTipText = "max length = " & txtnama.MaxLength
    txtalamat.ToolTipText = "max length = " & txtalamat.MaxLength
    txtalamat1.ToolTipText = "max length = " & txtalamat1.MaxLength
    txtkontak.ToolTipText = "max length = " & txtkontak.MaxLength
    txtelp.ToolTipText = "max length = " & txtelp.MaxLength
    txtfax.ToolTipText = "max length = " & txtfax.MaxLength
End Sub

Private Sub List1_DblClick()
    txtnama = List1.text
    txtnama = Trim(txtnama)
    OBJ.Open dsn
    SQL = "select * from am_supplier where namasupp = '" & txtnama & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtkode = RST!kodesupp
        txtalamat = RST!alamatsupp1
        txtalamat1 = RST!alamatsupp2
        txtelp = RST!telpsupp
        txtfax = RST!faxsupp
        txtkontak = RST!contactperson
        chk1.Value = RST!wp
        If RST!Category = "1" Then ops1.Value = True Else ops2.Value = True
    End If
    OBJ.Close
    List1.Visible = False
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtnama = List1.text
        txtnama = Trim(txtnama)
        OBJ.Open dsn
        SQL = "select * from am_supplier where namasupp = '" & txtnama & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            txtkode = RST!kodesupp
            txtalamat = RST!alamatsupp1
            txtalamat1 = RST!alamatsupp2
            txtelp = RST!telpsupp
            txtfax = RST!faxsupp
            txtkontak = RST!contactperson
            chk1.Value = RST!wp
            If RST!Category = "1" Then ops1.Value = True Else ops2.Value = True
        End If
        OBJ.Close
        List1.Visible = False
    End If
End Sub

Private Sub txtAlamat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtalamat1.SetFocus
End Sub

Private Sub txtAlamat1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtkontak.SetFocus
End Sub

Private Sub carisupplier()
    If txtkode = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from am_supplier where kodesupp = '" & txtkode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtnama = RST!namasupp
        txtalamat = RST!alamatsupp1
        txtalamat1 = RST!alamatsupp2
        txtelp = RST!telpsupp
        txtfax = RST!faxsupp
        txtkontak = RST!contactperson
        chk1.Value = RST!wp
        If RST!Category = "1" Then ops1.Value = True Else ops2.Value = True
        
        cmdsearch0.Enabled = False
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    MsgBox "Data not found.", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub txtelp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtfax.SetFocus
End Sub

Private Sub txtFax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then ops1.SetFocus
End Sub

Private Sub txtKontak_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtelp.SetFocus
End Sub

Private Sub txtnama_Change()
    If hasil = "" Then cari
End Sub

Private Sub txtnama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then List1.Visible = False
    If KeyAscii = 13 Then
        List1.Visible = False
        txtalamat.SetFocus
        
        OBJ.Open dsn
        txtnama = Trim(txtnama)
        SQL = "select * from am_supplier where namasupp = '" & txtnama & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            txtkode = RST!kodesupp
            txtalamat = RST!alamatsupp1
            txtalamat1 = RST!alamatsupp2
            txtelp = RST!telpsupp
            txtfax = RST!faxsupp
            txtkontak = RST!contactperson
            chk1.Value = RST!wp
            If RST!Category = "1" Then ops1.Value = True Else ops2.Value = True
        Else
            txtkode = ""
            txtalamat = ""
            txtalamat1 = ""
            txtelp = ""
            txtfax = ""
            txtkontak = ""
            chk1.Value = 0
            ops1.Value = True
        End If
        OBJ.Close
    End If
End Sub

Private Sub cari()
    If txtnama = "" Then
        List1.Visible = False
        Exit Sub
    End If
    List1.Clear
    
    OBJ.Open dsn
    SQL = "select namasupp from am_supplier where namasupp like '" & txtnama & "%' order by namasupp"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        Do While Not RST.EOF
            List1.AddItem RST!namasupp
            RST.MoveNext
        Loop
        List1.Visible = True
    Else
        List1.Visible = False
    End If
    OBJ.Close
End Sub
