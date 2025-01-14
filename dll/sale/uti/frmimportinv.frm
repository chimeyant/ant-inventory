VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frmimportinv 
   Caption         =   "Import Invoice"
   ClientHeight    =   2175
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2175
   ScaleWidth      =   5400
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5175
      _Version        =   851970
      _ExtentX        =   9128
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   "Progress"
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
      Begin XtremeSuiteControls.ProgressBar Pg 
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   4935
         _Version        =   851970
         _ExtentX        =   8705
         _ExtentY        =   450
         _StockProps     =   93
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
         UseVisualStyle  =   0   'False
         TextAlignment   =   2
      End
      Begin VB.Label lblstatus 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   615
         Width           =   4695
      End
   End
   Begin MSComDlg.CommonDialog cmndlg 
      Left            =   5040
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   1740
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
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
      MICON           =   "frmimportinv.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdImport 
      Height          =   375
      Left            =   3495
      TabIndex        =   3
      Top             =   1740
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Import"
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
      MICON           =   "frmimportinv.frx":031A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdgetfile 
      Height          =   375
      Left            =   105
      TabIndex        =   4
      Top             =   1740
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Get File (*.TRS)"
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
      MICON           =   "frmimportinv.frx":0634
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton chameleonButton1 
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Import BPB"
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
      MICON           =   "frmimportinv.frx":094E
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
      Left            =   255
      TabIndex        =   7
      Top             =   1395
      Width           =   4830
   End
   Begin VB.Label lblfile 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Filename"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "frmimportinv"
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

Dim OBJ2 As New ADODB.Connection
Dim RST2 As New ADODB.Recordset
Dim SQL2 As String

Private RSExcel As DAO.Recordset
Private DBExcel As DAO.Database
Private SQLExcel As String

Dim SP As New ADODB.Command
Dim vsp(2) As Variant
Dim fso As FileSystemObject
Dim j As Integer


Private flname As String
Private source_file As String

Private Sub chameleonButton1_Click()
    Pg.Visible = True
    cmdgetfile.Enabled = False
    cmdclose.Enabled = False
    importBPB
End Sub
Private Sub importBPB()
    Dim jmlrec As Integer
    OBJ2.Open dsn
    SQL2 = "SELECT count(nobpb)'totalbaris' FROM OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0', 'Data Source=" & Label1 & ";Extended Properties=Excel 8.0')...[Sheet2$]"
    Set RST2 = OBJ2.Execute(SQL2)
    If Not RST2.EOF Then j = RST2!totalbaris Else j = 0
    jmlrec = j
    OBJ2.Close

    OBJ.Open dsn
    SQL = "SELECT * FROM OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0', 'Data Source=" & Label1 & ";Extended Properties=Excel 8.0')...[Sheet2$]"
    Set RST = OBJ.Execute(SQL)
    lblfile.Caption = "Import data is on progress..."
    If RST.State = 1 Then
        Pg.Visible = True
        Pg.Min = 0
        Pg.Max = jmlrec
        Pg.Value = 0
    
        OBJ1.Open dsn
        Do While Not RST.EOF
            If RST!nobpb = "End" Then Exit Do 'di akhir baris bpb ada string "End"
                'SQL1 = "DELETE am_beliapp Where nobeli='" & RST!newbpb & "'"
                'OBJ1.Execute SQL1
                'MsgBox SQL1, vbInformation
                
                'SQL1 = "DELETE am_belihdr Where nobeli='" & RST!newbpb & "'"
                'OBJ1.Execute SQL1
                'MsgBox SQL1, vbInformation
                
                'SQL1 = "DELETE am_belilin Where nobeli='" & RST!newbpb & "'"
                'OBJ1.Execute SQL1
                'MsgBox SQL1, vbInformation
                
                'SQL1 = "Update am_beliapp set nobeli= '" & RST!newbpb & "', tglbeli='" & Format(RST!tglbpb, "yyyy-MM-dd") & "'"
                'SQL1 = SQL1 + " Where nobeli='" & RST!nobpb & "'"
                'OBJ1.Execute SQL1
                'MsgBox SQL1, vbInformation
                
                'SQL1 = "Update am_belihdr set nobeli= '" & RST!newbpb & "', tglbeli='" & Format(RST!tglbpb, "yyyy-MM-dd") & "'"
        'Update field database
                SQL1 = "Update am_belihdr set nosj= '" & RST!newsj & "'"
                SQL1 = SQL1 + " Where nobeli='" & RST!nobpb & "'"
                OBJ1.Execute SQL1
                'MsgBox SQL1, vbInformation
                
                'SQL1 = "Update am_belilin set nobeli= '" & RST!newbpb & "'"
                'SQL1 = SQL1 + " Where nobeli='" & RST!nobpb & "'"
                'OBJ1.Execute SQL1
                'MsgBox SQL1, vbInformation
                
                
                Pg.Value = Pg.Value + 1
                Pg.text = Str(Int(((Pg.Value / Pg.Max) * 100))) & " % Complete"
            RST.MoveNext
        Loop
        OBJ1.Close
        Pg.Visible = True
        cmdgetfile.Enabled = True
        cmdclose.Enabled = True
        MsgBox "Berhasil", vbInformation
    End If
    OBJ.Close
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdgetfile_Click()
    With cmndlg
        .CancelError = False
        .DialogTitle = "File Data"
        .Filter = "Transfer File (*.trs)|*.trs"
        .ShowOpen
        If .FileName <> "" Then
        flname = .FileName
            Set fso = New FileSystemObject
'            fso.CopyFile flname, "\\10.201.0.2\bagi-bagi\SERVER\INV\" & .FileTitle, True
            OBJ.Open dsn
            SQL = "EXEC xp_cmdshell 'net use Z: \\10.201.0.2\bagi-bagi\SERVER\INV /User:bagi_bagi kilometer13 /persistent:no'"
            OBJ.Execute (SQL)
            SQL = "EXEC xp_cmdshell 'copy Z:\" & .FileTitle & " C:\DATA'"
            OBJ.Execute (SQL)
            SQL = "EXEC xp_cmdshell 'net use Z: /delete /y'"
            OBJ.Execute (SQL)
            OBJ.Close
            
            Label1 = "C:\DATA\" & .FileTitle
            
            lblfile = .FileTitle
            lblfile = " Ready to import...!"
            Exit Sub
        Else
            lblfile = "Data Tidak Ditemukan"
            Exit Sub
        End If
    End With
End Sub

Private Sub cmdimport_Click()
    Pg.Visible = True
    cmdgetfile.Enabled = False
    cmdclose.Enabled = False
    importdata2
End Sub
Private Sub importdata2()
On Error GoTo Err_handler:
Dim jmlrec, jmlpost As Long

jmlrec = 0
    
    If flname = "" Then
        MsgBox "Silahkan pilih terlebih dahulu File(*.trs) yang akan diimport", vbCritical, AppName
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "SELECT count(kodecust)'totalbaris' FROM OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0', 'Data Source=" & Label1 & ";Extended Properties=Excel 8.0')...[customer$]"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then j = RST!totalbaris Else j = 0
    jmlrec = j
    
    SQL = "SELECT count(nobkt)'totalbaris' FROM OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0', 'Data Source=" & Label1 & ";Extended Properties=Excel 8.0')...[invlin$]"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then j = RST!totalbaris Else j = 0
    jmlrec = jmlrec + j
    
    SQL = "SELECT count(nobkt)'totalbaris' FROM OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0', 'Data Source=" & Label1 & ";Extended Properties=Excel 8.0')...[invhdr$]"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then j = RST!totalbaris Else j = 0
    jmlrec = jmlrec + j
    jmlpost = j
    
    OBJ.Close
   
    Pg.Min = 0
    Pg.Max = jmlrec
    Pg.Value = 0
    cmdImport.Enabled = False
    Me.MousePointer = vbHourglass
    'customer
    OBJ.Open dsn
    SQL = "SELECT * FROM OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0', 'Data Source=" & Label1 & ";Extended Properties=Excel 8.0')...[customer$]"
    Set RST = OBJ.Execute(SQL)
    lblfile.Caption = "Import data is on progress..."
    If RST.State = 1 Then
        Do While Not RST.EOF
            OBJ1.Open dsn
            SQL1 = "select * from AM_customer where kodecust = '" & RST!kodecust & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If RST1.EOF Then
                SQL1 = "INSERT INTO AM_Customer"
                SQL1 = SQL1 + "(KodeCust"
                SQL1 = SQL1 + ",NamaCust"
                SQL1 = SQL1 + ",AlamatCust"
                SQL1 = SQL1 + ",kota"
                SQL1 = SQL1 + ",TelpCust"
                SQL1 = SQL1 + ",FaxCust"
                SQL1 = SQL1 + ",contactPerson"
                SQL1 = SQL1 + ",KodePos"
                SQL1 = SQL1 + ",termcust"
                SQL1 = SQL1 + ",limit"
                SQL1 = SQL1 + ",Namanpwp"
                SQL1 = SQL1 + ",alamatnpwp"
                SQL1 = SQL1 + ",Nonpwp"
                SQL1 = SQL1 + ",kodearea"
                SQL1 = SQL1 + ",kodeacgl"
                SQL1 = SQL1 + ",identry"
                SQL1 = SQL1 + ",dateupdate"
                SQL1 = SQL1 + ",idupdate"
                SQL1 = SQL1 + ",Dateentry)"
            
                SQL1 = SQL1 + "VALUES"
                SQL1 = SQL1 + " ('" & RST!kodecust & "'"
                SQL1 = SQL1 + ", '" & RST!namacust & "'"
                SQL1 = SQL1 + ", '" & RST!alamatcust & "'"
                SQL1 = SQL1 + ", '" & RST!kota & "'"
                SQL1 = SQL1 + ", '" & RST!telpcust & "'"
                SQL1 = SQL1 + ", '" & RST!faxcust & "'"
                SQL1 = SQL1 + ", '" & RST!contactperson & "'"
                SQL1 = SQL1 + ", '" & RST!kodepos & "'"
                SQL1 = SQL1 + ", convert(money,'" & RST!termcust & "')"
                SQL1 = SQL1 + ", convert(money,'" & RST!limit & "')"
                SQL1 = SQL1 + ", '" & RST!namanpwp & "'"
                SQL1 = SQL1 + ", '" & RST!alamatnpwp & "'"
                SQL1 = SQL1 + ", '" & RST!nonpwp & "'"
                SQL1 = SQL1 + ", '" & RST!kodearea & "'"
                SQL1 = SQL1 + ", '" & RST!kodeacgl & "'"
                SQL1 = SQL1 + ", '" & RST!identry & "'"
                SQL1 = SQL1 + ",Convert(dateTime, '" & Month(RST!dateupdate) & "/" & Day(RST!dateupdate) & "/" & Year(RST!dateupdate) & "')"
                SQL1 = SQL1 + ", '" & RST!idupdate & "'"
                SQL1 = SQL1 + ",Convert(dateTime, '" & Month(RST!dateentry) & "/" & Day(RST!dateentry) & "/" & Year(RST!dateentry) & "'))"
                Set RST1 = OBJ1.Execute(SQL1)
            Else
                SQL1 = "Update am_customer set"
                SQL1 = SQL1 + " namacust= '" & RST!namacust & "'"
                SQL1 = SQL1 + ", alamatcust= '" & RST!alamatcust & "'"
                SQL1 = SQL1 + ", kota= '" & RST!kota & "'"
                SQL1 = SQL1 + ", telpcust= '" & RST!telpcust & "'"
                SQL1 = SQL1 + ", faxcust= '" & RST!faxcust & "'"
                SQL1 = SQL1 + ", contactperson= '" & RST!contactperson & "'"
                SQL1 = SQL1 + ", kodepos= '" & RST!kodepos & "'"
                SQL1 = SQL1 + ", termcust= convert(money,'" & RST!termcust & "')"
                SQL1 = SQL1 + ", limit= convert(money,'" & RST!limit & "')"
                SQL1 = SQL1 + ", namanpwp= '" & RST!namanpwp & "'"
                SQL1 = SQL1 + ", alamatnpwp= '" & RST!alamatnpwp & "'"
                SQL1 = SQL1 + ", nonpwp= '" & RST!nonpwp & "'"
                SQL1 = SQL1 + ", kodearea= '" & RST!kodearea & "'"
                SQL1 = SQL1 + ", kodeacgl= '" & RST!kodeacgl & "'"
                SQL1 = SQL1 + ", identry= '" & RST!identry & "'"
                SQL1 = SQL1 + ", dateupdate= convert(datetime,'" & Month(RST!dateupdate) & "/" & Day(RST!dateupdate) & "/" & Year(RST!dateupdate) & "')"
                SQL1 = SQL1 + ", idupdate= '" & RST!idupdate & "'"
                SQL1 = SQL1 + ", dateentry= convert(datetime,'" & Month(RST!dateentry) & "/" & Day(RST!dateentry) & "/" & Year(RST!dateentry) & "')"
                SQL1 = SQL1 + " Where kodecust= '" & RST!kodecust & "'"
                Set RST1 = OBJ1.Execute(SQL1)
            End If
            Pg.Value = Pg.Value + 1
            lblstatus.Caption = RST!namacust
            Pg.text = Str(Int(((Pg.Value / Pg.Max) * 100))) & " % Complete"
            OBJ1.Close
            DoEvents
            RST.MoveNext
        Loop
    End If
    OBJ.Close
    'invlin
    OBJ.Open dsn
    SQL = "SELECT * FROM OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0', 'Data Source=" & Label1 & ";Extended Properties=Excel 8.0')...[invlin$]"
    Set RST = OBJ.Execute(SQL)
    If RST.State = 1 Then
        Do While Not RST.EOF
            OBJ1.Open dsn
            SQL1 = "select * from am_invlin where nobkt = '" & RST!nobkt & "' and kodebarang = '" & RST!kodebarang & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If RST1.EOF Then
                SQL1 = "INSERT INTO am_invlin"
                SQL1 = SQL1 + "(type"
                SQL1 = SQL1 + ",nobkt"
                SQL1 = SQL1 + ",kodebarang"
                SQL1 = SQL1 + ",qty"
                SQL1 = SQL1 + ",price"
                SQL1 = SQL1 + ",lineitem"
                SQL1 = SQL1 + ",kodesatuan"
                SQL1 = SQL1 + ",bn"
                SQL1 = SQL1 + ",discline)"
                
                SQL1 = SQL1 + " values"
                SQL1 = SQL1 + " ('" & RST!Type & "'"
                SQL1 = SQL1 + ", '" & RST!nobkt & "'"
                SQL1 = SQL1 + ", '" & RST!kodebarang & "'"
                SQL1 = SQL1 + ", convert(money,'" & RST!qty & "')"
                SQL1 = SQL1 + ", convert(money,'" & RST!Price & "')"
                SQL1 = SQL1 + ", convert(numeric,'" & RST!lineitem & "')"
                SQL1 = SQL1 + ", '" & RST!kodesatuan & "'"
                SQL1 = SQL1 + ", convert(money,'" & RST!bn & "')"
                SQL1 = SQL1 + ", convert(money,'" & RST!discline & "'))"
                Set RST1 = OBJ1.Execute(SQL1)
                Pg.Value = Pg.Value + 1
                lblstatus.Caption = "No. Bukti " & RST!nobkt & ", Kode Barang " & RST!kodebarang
                Pg.text = Str(Int(((Pg.Value / Pg.Max) * 100))) & " % Complete"
            Else
                SQL1 = "Update am_invlin set"
                SQL1 = SQL1 + " type= '" & RST!Type & "'"
                SQL1 = SQL1 + ", kodebarang= '" & RST!kodebarang & "'"
                SQL1 = SQL1 + ", qty= convert(money,'" & RST!qty & "')"
                SQL1 = SQL1 + ", price= convert(money,'" & RST!Price & "')"
                SQL1 = SQL1 + ", lineitem= convert(numeric,'" & RST!lineitem & "')"
                SQL1 = SQL1 + ", kodesatuan= '" & RST!kodesatuan & "'"
                SQL1 = SQL1 + ", bn= convert(money,'" & RST!bn & "')"
                SQL1 = SQL1 + ", discline= convert(money,'" & RST!discline & "')"
                SQL1 = SQL1 + " Where nobkt= '" & RST!nobkt & "' and kodebarang= '" & RST!kodebarang & "'"
                Set RST1 = OBJ1.Execute(SQL1)
                Pg.Value = Pg.Value + 1
                lblstatus.Caption = "Update No. Bukti " & RST!nobkt & ", Kode Barang " & RST!kodebarang
                Pg.text = Str(Int(((Pg.Value / Pg.Max) * 100))) & " % Complete"
            End If
            OBJ1.Close
            DoEvents
            RST.MoveNext
        Loop
    End If
    OBJ.Close
    'invhdr
    OBJ.Open dsn
    SQL = "SELECT * FROM OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0', 'Data Source=" & Label1 & ";Extended Properties=Excel 8.0')...[invhdr$]"
    Set RST = OBJ.Execute(SQL)
    If RST.State = 1 Then
        Do While Not RST.EOF
            OBJ1.Open dsn
            SQL1 = "select * from am_invhdr where nobkt = '" & RST!nobkt & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If RST1.EOF Then
                SQL1 = "insert into am_invhdr"
                SQL1 = SQL1 + "(nosj"
                SQL1 = SQL1 + ",noapply"
                SQL1 = SQL1 + ",type"
                SQL1 = SQL1 + ",nobkt"
                SQL1 = SQL1 + ",tglbkt"
                SQL1 = SQL1 + ",kodecust"
                SQL1 = SQL1 + ",namacust"
                SQL1 = SQL1 + ",alamatcust"
                SQL1 = SQL1 + ",kodesales"
                SQL1 = SQL1 + ",discprc"
                SQL1 = SQL1 + ",discamt"
                SQL1 = SQL1 + ",ppn"
                SQL1 = SQL1 + ",ppnbm"
                SQL1 = SQL1 + ",termpay"
                SQL1 = SQL1 + ",identry"
                SQL1 = SQL1 + ",dateentry"
                SQL1 = SQL1 + ",idupdate"
                SQL1 = SQL1 + ",dateupdate"
                SQL1 = SQL1 + ",Posted"
                SQL1 = SQL1 + ",kodecur"
                SQL1 = SQL1 + ",nilaikurs"
                SQL1 = SQL1 + ",noseri)"
            
                SQL1 = SQL1 + " values"
                SQL1 = SQL1 + " ('" & RST!nosj & "'"
                SQL1 = SQL1 + ", '" & RST!noapply & "'"
                SQL1 = SQL1 + ", '" & RST!Type & "'"
                SQL1 = SQL1 + ", '" & RST!nobkt & "'"
                SQL1 = SQL1 + ", Convert(dateTime, '" & Month(RST!tglbkt) & "/" & Day(RST!tglbkt) & "/" & Year(RST!tglbkt) & "')"
                SQL1 = SQL1 + ", '" & RST!kodecust & "'"
                SQL1 = SQL1 + ", '" & RST!namacust & "'"
                SQL1 = SQL1 + ", '" & RST!alamatcust & "'"
                SQL1 = SQL1 + ", '" & RST!kodesales & "'"
                SQL1 = SQL1 + ", convert(money,'" & RST!discprc & "')"
                SQL1 = SQL1 + ", convert(money,'" & RST!discamt & "')"
                SQL1 = SQL1 + ", convert(money,'" & RST!ppn & "')"
                SQL1 = SQL1 + ", convert(money,'" & RST!ppnbm & "')"
                SQL1 = SQL1 + ", '" & RST!termpay & "'"
                SQL1 = SQL1 + ", '" & RST!identry & "'"
                SQL1 = SQL1 + ", Convert(dateTime, '" & Month(RST!dateentry) & "/" & Day(RST!dateentry) & "/" & Year(RST!dateentry) & "')"
                SQL1 = SQL1 + ", '" & RST!idupdate & "'"
                SQL1 = SQL1 + ", Convert(dateTime, '" & Month(RST!dateupdate) & "/" & Day(RST!dateupdate) & "/" & Year(RST!dateupdate) & "')"
                SQL1 = SQL1 + ", '" & RST!posted & "'"
                SQL1 = SQL1 + ", '" & RST!kodecur & "'"
                SQL1 = SQL1 + ", convert(money,'" & RST!nilaikurs & "')"
                SQL1 = SQL1 + ", '" & RST!noseri & "')"
                Set RST1 = OBJ1.Execute(SQL1)
                Pg.Value = Pg.Value + 1
                lblstatus.Caption = "No. Bukti: " & RST!nobkt & ", No. Apply: " & RST!noapply
                Pg.text = Str(Int(((Pg.Value / Pg.Max) * 100))) & " % Complete"
            Else
                SQL1 = "Update am_invhdr set"
                SQL1 = SQL1 + " nosj= '" & RST!nosj & "'"
                SQL1 = SQL1 + ", noapply= '" & RST!noapply & "'"
                SQL1 = SQL1 + ", type= '" & RST!Type & "'"
                SQL1 = SQL1 + ", tglbkt= convert(datetime,'" & Month(RST!tglbkt) & "/" & Day(RST!tglbkt) & "/" & Year(RST!tglbkt) & "')"
                SQL1 = SQL1 + ", kodecust= '" & RST!kodecust & "'"
                SQL1 = SQL1 + ", namacust= '" & RST!namacust & "'"
                SQL1 = SQL1 + ", alamatcust= '" & RST!alamatcust & "'"
                SQL1 = SQL1 + ", kodesales= '" & RST!kodesales & "'"
                SQL1 = SQL1 + ", discprc= convert(money,'" & RST!discprc & "')"
                SQL1 = SQL1 + ", discamt= convert(money,'" & RST!discamt & "')"
                SQL1 = SQL1 + ", ppn= convert(money,'" & RST!ppn & "')"
                SQL1 = SQL1 + ", ppnbm= convert(money,'" & RST!ppnbm & "')"
                SQL1 = SQL1 + ", termpay= '" & RST!termpay & "'"
                SQL1 = SQL1 + ", identry= '" & RST!identry & "'"
                SQL1 = SQL1 + ", dateentry= convert(DateTime, '" & Month(RST!dateentry) & "/" & Day(RST!dateentry) & "/" & Year(RST!dateentry) & "')"
                SQL1 = SQL1 + ", idupdate= '" & RST!idupdate & "'"
                SQL1 = SQL1 + ", dateupdate= convert(DateTime, '" & Month(RST!dateupdate) & "/" & Day(RST!dateupdate) & "/" & Year(RST!dateupdate) & "')"
                SQL1 = SQL1 + ", posted= '" & RST!posted & "'"
                SQL1 = SQL1 + ", kodecur= '" & RST!kodecur & "'"
                SQL1 = SQL1 + ", nilaikurs= convert(money,'" & RST!nilaikurs & "')"
                SQL1 = SQL1 + ", noseri= '" & RST!noseri & "'"
                SQL1 = SQL1 + " Where nobkt= '" & RST!nobkt & "'"
                Set RST1 = OBJ1.Execute(SQL1)
                Pg.Value = Pg.Value + 1
                lblstatus.Caption = "Update No. Bukti: " & RST!nobkt & ", No. Apply: " & RST!noapply
                Pg.text = Str(Int(((Pg.Value / Pg.Max) * 100))) & " % Complete"
            End If
            OBJ1.Close
            DoEvents
            RST.MoveNext
        Loop
    End If
    OBJ.Close


    Pg.Min = 0
    Pg.Max = jmlpost
    Pg.Value = 0
    lblfile.Caption = "Please Wait..."
    'aropnfil by invhdr
    OBJ.Open dsn
    SQL = "SELECT * FROM OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0', 'Data Source=" & Label1 & ";Extended Properties=Excel 8.0')...[invhdr$]"
    Set RST = OBJ.Execute(SQL)
    If RST.State = 1 Then
        Do While Not RST.EOF
            If RST!nobkt = "" Or IsNull(RST!nobkt) Then Exit Do
            DoEvents
            
            OBJ1.Open dsn
            SQL1 = "select * from am_invhdr where nobkt = '" & RST!nobkt & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                SQL2 = "Select * From am_aropnfil Where nobkt = '" & RST!nobkt & "'"
                Set RST2 = OBJ1.Execute(SQL2)
                If Not RST2.EOF Then
                    SQL2 = "Delete am_aropnfil Where nobkt = '" & RST!nobkt & "'"
                    Set RST2 = OBJ1.Execute(SQL2)
                End If
            End If
            
            SP.ActiveConnection = dsn
            SP.CommandType = adCmdStoredProc
            SP.CommandText = "am_postinginv"
            vsp(0) = RST1!nobkt
            vsp(1) = Format(RST1!tglbkt, "yyyyMMdd")
            vsp(2) = "sj"
            SP.Execute , vsp
            Set SP = Nothing
            OBJ1.Close
            DoEvents
            
            Pg.Value = Pg.Value + 1
            lblstatus.Caption = "No. Bukti: " & RST!nobkt & ", Tgl. Bukti: " & RST!tglbkt
            Pg.text = Str(Int(((Pg.Value / Pg.Max) * 100))) & " % Complete"
            RST.MoveNext
        Loop
    End If
    Me.MousePointer = vbDefault
    OBJ.Close
            
    lblfile.Caption = "import success"
    flname = ""
    Pg.Value = 0
    Kill Label1
    MsgBox "Data Berhasil diimport", vbInformation, AppName
    Unload Me
    Exit Sub
Err_handler:
    If OBJ.State = 1 Then OBJ.Close
    cmdclose.Enabled = True
    MsgBox "Import gagal." + Chr(13) + "Mohon tutup Program, kemudian Import data kembali" + Chr(13) + "Error : " + Err.Description, vbCritical, AppName
End Sub
Private Sub importdata()
'On Error GoTo Err_handler:
Dim jmlrec As Long

jmlrec = 0
    If flname = "" Then
        MsgBox "Silahkan pilih terlebih dahulu File(*.trs) yang akan diimport", vbCritical, AppName
        Exit Sub
    End If

'total record dataexcel
    Set DBExcel = OpenDatabase(flname, False, True, "Excel 8.0")
    Set RSExcel = DBExcel.OpenRecordset("customer$")
    Do While Not RSExcel.EOF
        If RSExcel!kodecust = "" Or IsNull(RSExcel!kodecust) Then Exit Do
        jmlrec = jmlrec + 1
        RSExcel.MoveNext
    Loop
    
    Set DBExcel = OpenDatabase(flname, False, True, "Excel 8.0")
    Set RSExcel = DBExcel.OpenRecordset("invlin$")
    Do While Not RSExcel.EOF
        If RSExcel!nobkt = "" Or IsNull(RSExcel!nobkt) Then Exit Do
        jmlrec = jmlrec + 1
        RSExcel.MoveNext
    Loop
    
    Set DBExcel = OpenDatabase(flname, False, True, "Excel 8.0")
    Set RSExcel = DBExcel.OpenRecordset("invhdr$")
    Do While Not RSExcel.EOF
        If RSExcel!nobkt = "" Or IsNull(RSExcel!nobkt) Then Exit Do
        jmlrec = jmlrec + 1
        RSExcel.MoveNext
    Loop
    
    Pg.Min = 0
    Pg.Max = jmlrec
    Pg.Value = 0
cmdImport.Enabled = False
lblfile.Caption = "Please Wait..."
    
    Set DBExcel = OpenDatabase(flname, False, True, "Excel 8.0")
    Set RSExcel = DBExcel.OpenRecordset("customer$")
    RSExcel.MoveFirst
'SIMPAN KE TABEL AM_CUSTOMER
    Do While Not RSExcel.EOF
        If RSExcel!kodecust = "" Or IsNull(RSExcel!kodecust) Then Exit Do
        DoEvents
        OBJ.Open dsn
        SQL = "Select * From am_customer Where kodecust = '" + RSExcel!kodecust + "'"
        Set RST = New ADODB.Recordset
        RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
        If Not RST.EOF Then
            With RST
                !kodecust = RSExcel!kodecust
                !namacust = RSExcel!namacust
                If IsNull(RSExcel!alamatcust) Then
                    !alamatcust = ""
                Else
                    !alamatcust = RSExcel!alamatcust
                End If
                If IsNull(RSExcel!telpcust) Then
                    !telpcust = ""
                Else
                    !telpcust = RSExcel!telpcust
                End If
                If IsNull(RSExcel!faxcust) Then
                    !faxcust = ""
                Else
                    !faxcust = RSExcel!faxcust
                End If
                If IsNull(RSExcel!contactperson) Then
                    !contactperson = ""
                Else
                    !contactperson = RSExcel!contactperson
                End If
                If IsNull(RSExcel!kodepos) Then
                    !kodepos = ""
                Else
                    !kodepos = RSExcel!kodepos
                End If
                If IsNull(RSExcel!identry) Then
                    !identry = ""
                Else
                    !identry = RSExcel!identry
                End If
                If IsNull(RSExcel!dateentry) Then
                    !dateentry = Date
                Else
                    !dateentry = RSExcel!dateentry
                End If
                If IsNull(RSExcel!idupdate) Then
                    !idupdate = ""
                Else
                    !idupdate = RSExcel!idupdate
                End If
                If IsNull(RSExcel!dateupdate) Then
                    !dateupdate = Date
                Else
                    !dateupdate = RSExcel!dateupdate
                End If
                If IsNull(RSExcel!kota) Then
                    !kota = ""
                Else
                    !kota = RSExcel!kota
                End If
                If IsNull(RSExcel!termcust) Then
                    !termcust = ""
                Else
                    !termcust = RSExcel!termcust
                End If
                If IsNull(RSExcel!limit) Then
                    !limit = ""
                Else
                    !limit = RSExcel!limit
                End If
                If IsNull(RSExcel!namanpwp) Then
                    !namanpwp = ""
                Else
                    !namanpwp = RSExcel!namanpwp
                End If
                If IsNull(RSExcel!alamatnpwp) Then
                    !alamatnpwp = ""
                Else
                    !alamatnpwp = RSExcel!alamatnpwp
                End If
                If IsNull(RSExcel!nonpwp) Then
                    !nonpwp = ""
                Else
                    !nonpwp = RSExcel!nonpwp
                End If
                If IsNull(RSExcel!kodearea) Then
                    !kodearea = ""
                Else
                    !kodearea = RSExcel!kodearea
                End If
                If IsNull(RSExcel!kodeacgl) Then
                    !kodeacgl = ""
                Else
                    !kodeacgl = RSExcel!kodeacgl
                End If
                .Update
            End With
            Pg.Value = Pg.Value + 1
            lblstatus.Caption = RST!namacust
            Pg.text = Str(Int(((Pg.Value / Pg.Max) * 100))) & " % Complete"
            OBJ.Close
        Else
            With RST
                .AddNew
                !kodecust = RSExcel!kodecust
                !namacust = RSExcel!namacust
                If IsNull(RSExcel!alamatcust) Then
                    !alamatcust = ""
                Else
                    !alamatcust = RSExcel!alamatcust
                End If
                If IsNull(RSExcel!telpcust) Then
                    !telpcust = ""
                Else
                    !telpcust = RSExcel!telpcust
                End If
                If IsNull(RSExcel!faxcust) Then
                    !faxcust = ""
                Else
                    !faxcust = RSExcel!faxcust
                End If
                If IsNull(RSExcel!contactperson) Then
                    !contactperson = ""
                Else
                    !contactperson = RSExcel!contactperson
                End If
                If IsNull(RSExcel!kodepos) Then
                    !kodepos = ""
                Else
                    !kodepos = RSExcel!kodepos
                End If
                If IsNull(RSExcel!identry) Then
                    !identry = ""
                Else
                    !identry = RSExcel!identry
                End If
                If IsNull(RSExcel!dateentry) Then
                    !dateentry = Date
                Else
                    !dateentry = RSExcel!dateentry
                End If
                If IsNull(RSExcel!idupdate) Then
                    !idupdate = ""
                Else
                    !idupdate = RSExcel!idupdate
                End If
                If IsNull(RSExcel!dateupdate) Then
                    !dateupdate = Date
                Else
                    !dateupdate = RSExcel!dateupdate
                End If
                If IsNull(RSExcel!kota) Then
                    !kota = ""
                Else
                    !kota = RSExcel!kota
                End If
                If IsNull(RSExcel!termcust) Then
                    !termcust = ""
                Else
                    !termcust = RSExcel!termcust
                End If
                If IsNull(RSExcel!limit) Then
                    !limit = ""
                Else
                    !limit = RSExcel!limit
                End If
                If IsNull(RSExcel!namanpwp) Then
                    !namanpwp = ""
                Else
                    !namanpwp = RSExcel!namanpwp
                End If
                If IsNull(RSExcel!alamatnpwp) Then
                    !alamatnpwp = ""
                Else
                    !alamatnpwp = RSExcel!alamatnpwp
                End If
                If IsNull(RSExcel!nonpwp) Then
                    !nonpwp = ""
                Else
                    !nonpwp = RSExcel!nonpwp
                End If
                If IsNull(RSExcel!kodearea) Then
                    !kodearea = ""
                Else
                    !kodearea = RSExcel!kodearea
                End If
                If IsNull(RSExcel!kodeacgl) Then
                    !kodeacgl = ""
                Else
                    !kodeacgl = RSExcel!kodeacgl
                End If
                .Update
            End With
            Pg.Value = Pg.Value + 1
            lblstatus.Caption = RST!namacust
            Pg.text = Str(Int(((Pg.Value / Pg.Max) * 100))) & " % Complete"
            OBJ.Close
        End If
        RSExcel.MoveNext
    Loop
    
'============================
'SIMPAN KE TABEL AM_INVLIN
    lblfile.Caption = "Periksa data barang..."
    Set DBExcel = OpenDatabase(flname, False, True, "Excel 8.0")
    Set RSExcel = DBExcel.OpenRecordset("invlin$")
    RSExcel.MoveFirst

    Do While Not RSExcel.EOF
    'UPDATE = DELETE OLD RECORD
        If RSExcel!nobkt = "" Or IsNull(RSExcel!nobkt) Then Exit Do
        DoEvents
        OBJ.Open dsn
        SQL = "Select * From am_invlin Where nobkt = '" & RSExcel!nobkt & "' and kodebarang = '" & RSExcel!kodebarang & "'"
        Set RST = New ADODB.Recordset
        RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
        If Not RST.EOF Then
            OBJ.Close
            
            OBJ.Open dsn
            SQL = "Delete From am_invlin Where nobkt = '" & RSExcel!nobkt & "'"
            Set RST = OBJ.Execute(SQL)
            lblfile.Caption = "Update data barang..."
            lblstatus.Caption = "No. Bukti " & RSExcel!nobkt & ", Kode Barang " & RSExcel!kodebarang
        End If
        OBJ.Close
        RSExcel.MoveNext
    Loop
    
    Set DBExcel = OpenDatabase(flname, False, True, "Excel 8.0")
    Set RSExcel = DBExcel.OpenRecordset("invlin$")
    RSExcel.MoveFirst
    
    Do While Not RSExcel.EOF
    'INSERT NEW RECORD (ADD & UPDATE)
        If RSExcel!nobkt = "" Or IsNull(RSExcel!nobkt) Then Exit Do
        DoEvents
        OBJ.Open dsn
        SQL = "Select * From am_invlin Where nobkt = '" & RSExcel!nobkt & "' and kodebarang = '" & RSExcel!kodebarang & "'"
        Set RST = New ADODB.Recordset
        RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
            With RST
                .AddNew
                !Type = RSExcel!Type
                !nobkt = RSExcel!nobkt
                !kodebarang = RSExcel!kodebarang
                !qty = RSExcel!qty
                !Price = RSExcel!Price
                !lineitem = RSExcel!lineitem
                !kodesatuan = RSExcel!kodesatuan
                !bn = RSExcel!bn
                !discline = RSExcel!discline
                .Update
            End With
            Pg.Value = Pg.Value + 1
            lblstatus.Caption = "No. Bukti " & RSExcel!nobkt & ", Kode Barang " & RSExcel!kodebarang
            Pg.text = Str(Int(((Pg.Value / Pg.Max) * 100))) & " % Complete"
            OBJ.Close
            RSExcel.MoveNext
    Loop

'SIMPAN KE TABEL AM_INVHDR
    lblfile.Caption = "Periksa nomor bukti..."
    Set DBExcel = OpenDatabase(flname, False, True, "Excel 8.0")
    Set RSExcel = DBExcel.OpenRecordset("invhdr$")
    RSExcel.MoveFirst
    
    Do While Not RSExcel.EOF
        If RSExcel!nobkt = "" Or IsNull(RSExcel!nobkt) Then Exit Do
        DoEvents
        
        OBJ.Open dsn
        SQL = "Select * From am_invhdr Where nobkt = '" + RSExcel!nobkt + "'"
        Set RST = New ADODB.Recordset
        RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
        If Not RST.EOF Then
            'If RST!noseri = "" Then
                With RST
                    !nobkt = RSExcel!nobkt
                    !noapply = RSExcel!noapply
                    !tglbkt = RSExcel!tglbkt
                    !kodecust = RSExcel!kodecust
                    !namacust = RSExcel!namacust
                    !alamatcust = RSExcel!alamatcust
                    !kodesales = RSExcel!kodesales
                    !discprc = RSExcel!discprc
                    !discamt = RSExcel!discamt
                    !ppn = RSExcel!ppn
                    !ppnbm = RSExcel!ppnbm
                    !Type = RSExcel!Type
                    !termpay = RSExcel!termpay
                    If IsNull(RSExcel!identry) Then
                        !identry = ""
                    Else
                        !identry = RSExcel!identry
                    End If
                    !dateentry = RSExcel!dateentry
                    If IsNull(RSExcel!idupdate) Then
                        !idupdate = ""
                    Else
                        !idupdate = RSExcel!idupdate
                    End If
                    !dateupdate = RSExcel!dateupdate
                    !posted = RSExcel!posted
                    !kodecur = RSExcel!kodecur
                    !nilaikurs = RSExcel!nilaikurs
                    If IsNull(RSExcel!nosj) Then
                        !nosj = ""
                    Else
                        !nosj = RSExcel!nosj
                    End If
                    If RST!noseri = "" Then
                        If IsNull(RSExcel!noseri) Then
                            !noseri = ""
                        Else
                            !noseri = RSExcel!noseri
                        End If
                    End If
                    .Update
                End With
                Pg.Value = Pg.Value + 1
                lblstatus.Caption = "Update No. Bukti " & RSExcel!nobkt & ", No. Apply " & RSExcel!noapply
                Pg.text = Str(Int(((Pg.Value / Pg.Max) * 100))) & " % Complete"
        Else
            With RST
                .AddNew
                !nobkt = RSExcel!nobkt
                !noapply = RSExcel!noapply
                !tglbkt = RSExcel!tglbkt
                !kodecust = RSExcel!kodecust
                !namacust = RSExcel!namacust
                !alamatcust = RSExcel!alamatcust
                !kodesales = RSExcel!kodesales
                !discprc = RSExcel!discprc
                !discamt = RSExcel!discamt
                !ppn = RSExcel!ppn
                !ppnbm = RSExcel!ppnbm
                !Type = RSExcel!Type
                !termpay = RSExcel!termpay
                If IsNull(RSExcel!identry) Then
                    !identry = ""
                Else
                    !identry = RSExcel!identry
                End If
                !dateentry = RSExcel!dateentry
                If IsNull(RSExcel!idupdate) Then
                    !idupdate = ""
                Else
                    !idupdate = RSExcel!idupdate
                End If
                !dateupdate = RSExcel!dateupdate
                !posted = RSExcel!posted
                !kodecur = RSExcel!kodecur
                !nilaikurs = RSExcel!nilaikurs
                If IsNull(RSExcel!nosj) Then
                    !nosj = ""
                Else
                    !nosj = RSExcel!nosj
                End If
                If IsNull(RSExcel!noseri) Then
                    !noseri = ""
                Else
                    !noseri = RSExcel!noseri
                End If
                .Update
            End With
            Pg.Value = Pg.Value + 1
            lblstatus.Caption = "No. Bukti " & RSExcel!nobkt & ", No. Apply " & RSExcel!noapply
            Pg.text = Str(Int(((Pg.Value / Pg.Max) * 100))) & " % Complete"
        End If
        OBJ.Close
        RSExcel.MoveNext
    Loop
'==========================================
    jmlrec = 0
    Set DBExcel = OpenDatabase(flname, False, True, "Excel 8.0")
    Set RSExcel = DBExcel.OpenRecordset("invhdr$")
    Do While Not RSExcel.EOF
        If RSExcel!nobkt = "" Or IsNull(RSExcel!nobkt) Then Exit Do
        jmlrec = jmlrec + 1
        RSExcel.MoveNext
    Loop
    
    Pg.Min = 0
    Pg.Max = jmlrec
    Pg.Value = 0
'============================================
        lblfile.Caption = "Mohon tunggu hingga proses impor selesai..."
    Set DBExcel = OpenDatabase(flname, False, True, "Excel 8.0")
    Set RSExcel = DBExcel.OpenRecordset("invhdr$")
    RSExcel.MoveFirst
    
    Do While Not RSExcel.EOF
        If RSExcel!nobkt = "" Or IsNull(RSExcel!nobkt) Then Exit Do
        DoEvents
        
        OBJ.Open dsn
        SQL = "Select * From am_invhdr Where nobkt = '" + RSExcel!nobkt + "'"
        Set RST = New ADODB.Recordset
        RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
        If Not RST.EOF Then
            SQL = "Select * From am_aropnfil Where nobkt = '" + RSExcel!nobkt + "'"
            Set RST = New ADODB.Recordset
            RST.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
            If Not RST.EOF Then
                lblstatus.Caption = "Update No. Bukti " & RSExcel!nobkt & ", Tgl. Bukti " & RSExcel!tglbkt
                SQL = "Delete am_aropnfil Where nobkt = '" + RSExcel!nobkt + "'"
                OBJ.Execute SQL
            End If
            
        End If
            
            SP.ActiveConnection = dsn
            SP.CommandType = adCmdStoredProc
            SP.CommandText = "am_postinginv"
            vsp(0) = RST!nobkt
            vsp(1) = Format(RST!tglbkt, "yyyyMMdd")
            vsp(2) = "sj"
            SP.Execute , vsp
            Set SP = Nothing
            OBJ.Close
            
            Pg.Value = Pg.Value + 1
            lblstatus.Caption = "No. Bukti " & RSExcel!nobkt & ", Tgl. Bukti " & RSExcel!tglbkt
            Pg.text = Str(Int(((Pg.Value / Pg.Max) * 100))) & " % Complete"
            
        RSExcel.MoveNext
    Loop

    
    DBExcel.Close
    Set DBExcel = Nothing
    lblfile.Caption = "import success"
    MsgBox "Data Berhasil diimport", vbInformation, AppName
    Kill source_file
    flname = ""
    Pg.Value = 0
    lblstatus = "0 %"
    Unload Me
    Exit Sub
Err_handler:
    If OBJ.State = 1 Then OBJ.Close
    cmdclose.Enabled = True
    MsgBox "Import gagal." + Chr(13) + "Mohon tutup Program, kemudian Import data kembali" + Chr(13) + "Error : " + Err.Description, vbCritical, AppName
End Sub
