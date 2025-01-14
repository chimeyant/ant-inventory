VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmsoimport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Sales Order"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cmndlg 
      Left            =   180
      Top             =   330
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker date3 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
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
      Format          =   146341891
      CurrentDate     =   37426
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   2520
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
      MICON           =   "frmsoimport.frx":0000
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
      Left            =   240
      TabIndex        =   0
      Top             =   2505
      Width           =   1575
      _ExtentX        =   2778
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
      MICON           =   "frmsoimport.frx":031A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdimport 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   2520
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
      MICON           =   "frmsoimport.frx":0634
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
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Ready to Import !!!"
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
      Height          =   1815
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   2415
      Left            =   -120
      TabIndex        =   5
      Top             =   0
      Width           =   4095
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   3360
      Width           =   1215
   End
End
Attribute VB_Name = "frmsoimport"
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

Dim fso As FileSystemObject

Dim j As Integer

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function

Private Sub cmdclear_Click()
    Dim flname As String
    'With cmndlg
         'With cmndlg
        '.CancelError = False
        '.DialogTitle = "File Data"
        '.Filter = "Transfer File (*.trs)|*.trs"
        '.ShowOpen
        'If .FileName <> "" Then
            'flname = .FileName
            'Label2 = "C:\DATA\" & .FileTitle
            'Set fso = New FileSystemObject
            'fso.CopyFile flname, "\\" & dbServer & "\Data\" & .FileTitle, True
            'Label3 = .FileTitle
            'Label3 = Label3 + vbCrLf + "Data Ditemukan dan telah siap untuk di import...!"
            'Exit Sub
        'Else
            'Label3 = "Data Tidak Ditemukan"
            'Exit Sub
        'End If
        'End With
    'End With
    'cmdclear.Enabled = False
    Label2 = ""
    With cmndlg
         With cmndlg
        .CancelError = False
        .DialogTitle = "File Data"
        .Filter = "Transfer File (*.trs)|*.trs"
        .ShowOpen
        If .FileName <> "" Then
            flname = .FileName
            Label2 = "C:\DATA\" & .FileTitle
            Set fso = New FileSystemObject
            fso.CopyFile flname, "\\10.201.0.2\bagi-bagi\SERVER\SO\" & .FileTitle, True
            OBJ.Open dsn
            SQL = "EXEC xp_cmdshell 'net use Z: \\10.201.0.2\bagi-bagi\SERVER\SO /User:bagi_bagi kilometer13 /persistent:no'"
            OBJ.Execute (SQL)
            SQL = "EXEC xp_cmdshell 'copy Z:\" & .FileTitle & " C:\DATA'"
            OBJ.Execute (SQL)
            SQL = "EXEC xp_cmdshell 'net use Z: /delete /y'"
            OBJ.Execute (SQL)
            OBJ.Close
            
            Label3 = .FileTitle
            Label3 = Label3 + vbCrLf + "Data Ditemukan dan telah siap untuk di import...!"
            Exit Sub
        Else
            Label3 = "Data Tidak Ditemukan"
            Exit Sub
        End If
        End With
    End With
    
    cmdclear.Enabled = False
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdimport_Click()
On Error GoTo Err_handler
    If Label2 = "" Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    Label3 = "  please wait a moment ..."
    
    If MsgBox("Please make sure file exsist and valid." & vbCrLf & "Are you sure want to continue import file " & Label2 & " ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    Label3 = "Result :"
    'so
    OBJ.Open dsn
    SQL = "SELECT * FROM OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0', 'Data Source=" & Label2 & ";Extended Properties=Excel 8.0')...[so$]"
    Set RST = OBJ.Execute(SQL)
    If RST.State = 1 Then
        OBJ1.Open dsn
        SQL1 = "update AM_soapp set flag2='2' where flag2='1'"
        Set RST1 = OBJ1.Execute(SQL1)
        OBJ1.Close
        
        Do While Not RST.EOF
            OBJ1.Open dsn
            SQL1 = "select * from AM_soapp where noso = '" & RST!noso & "' and kodebarang = '" & RST!kodebarang & "' and kodesatuan = '" & RST!kodesatuan & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If RST1.EOF Then
                SQL1 = "INSERT INTO AM_soapp"
                SQL1 = SQL1 + " (noso"
                SQL1 = SQL1 + ", Tglso"
                SQL1 = SQL1 + ", kodecust"
                SQL1 = SQL1 + ", kodesales"
                SQL1 = SQL1 + ", nopo"
                SQL1 = SQL1 + ", jasa"
                SQL1 = SQL1 + ", Kodebarang"
                SQL1 = SQL1 + ", qty"
                SQL1 = SQL1 + ", keterangan"
                SQL1 = SQL1 + ", kodesatuan"
                SQL1 = SQL1 + ", lineitem"
                SQL1 = SQL1 + ", bn"
                SQL1 = SQL1 + ", flag1"
                SQL1 = SQL1 + ", flag2)"
            
                SQL1 = SQL1 + "VALUES"
                SQL1 = SQL1 + " ('" & RST!noso & "'"
                SQL1 = SQL1 + ",Convert(dateTime, '" & Month(RST!tglso) & "/" & Day(RST!tglso) & "/" & Year(RST!tglso) & "')"
                SQL1 = SQL1 + ", '" & RST!kodecust & "'"
                SQL1 = SQL1 + ", '" & RST!kodesales & "'"
                SQL1 = SQL1 + ", '" & RST!nopo & "'"
                SQL1 = SQL1 + ", '" & RST!jasa & "'"
                SQL1 = SQL1 + ", '" & RST!kodebarang & "'"
                SQL1 = SQL1 + ",Convert (Money, '" & RST!qty & "')"
                SQL1 = SQL1 + ", '" & RST!keterangan & "'"
                SQL1 = SQL1 + ", '" & RST!kodesatuan & "'"
                SQL1 = SQL1 + ",Convert (numeric, '" & RST!lineitem & "')"
                SQL1 = SQL1 + ",Convert (Money, '" & RST!bn & "')"
                SQL1 = SQL1 + ", '1'"
                SQL1 = SQL1 + ", '1')"
                Set RST1 = OBJ1.Execute(SQL1)
            End If
            OBJ1.Close
            
            RST.MoveNext
        Loop
        
        SQL = "SELECT count(noso)'totalbaris' FROM OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0', 'Data Source=" & Label2 & ";Extended Properties=Excel 8.0')...[so$]"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then j = RST!totalbaris Else j = 0
        Label3 = Label3 + vbCrLf + "   Import Sales Order, " & j & " rows affected." 'customer
        
    End If
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "SELECT * FROM OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0', 'Data Source=" & Label2 & ";Extended Properties=Excel 8.0')...[customer$]"
    Set RST = OBJ.Execute(SQL)
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
                SQL1 = SQL1 + ", '" & kuser & "'"
                SQL1 = SQL1 + ",Convert(dateTime, '" & Month(RST!dateupdate) & "/" & Day(RST!dateupdate) & "/" & Year(RST!dateupdate) & "')"
                SQL1 = SQL1 + ", '" & kuser & "'"
                SQL1 = SQL1 + ",Convert(dateTime, '" & Month(RST!dateentry) & "/" & Day(RST!dateentry) & "/" & Year(RST!dateentry) & "'))"
                Set RST1 = OBJ1.Execute(SQL1)
            End If
            OBJ1.Close
            
            RST.MoveNext
        Loop
    End If
    OBJ.Close
    'customer
    OBJ.Open dsn
    SQL = "SELECT * FROM OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0', 'Data Source=" & Label2 & ";Extended Properties=Excel 8.0')...[customer$]"
    Set RST = OBJ.Execute(SQL)
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
                SQL1 = SQL1 + ", '" & kuser & "'"
                SQL1 = SQL1 + ",Convert(dateTime, '" & Month(RST!dateupdate) & "/" & Day(RST!dateupdate) & "/" & Year(RST!dateupdate) & "')"
                SQL1 = SQL1 + ", '" & kuser & "'"
                SQL1 = SQL1 + ",Convert(dateTime, '" & Month(RST!dateentry) & "/" & Day(RST!dateentry) & "/" & Year(RST!dateentry) & "'))"
                Set RST1 = OBJ1.Execute(SQL1)
            End If
            OBJ1.Close
            
            RST.MoveNext
        Loop
        
        SQL = "SELECT count(kodecust)'totalbaris' FROM OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0', 'Data Source=" & Label2 & ";Extended Properties=Excel 8.0')...[customer$]"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then j = RST!totalbaris Else j = 0
        Label3 = Label3 + vbCrLf + "   Import Customer, " & j & " rows affected."
    End If
    OBJ.Close
    'sales
    OBJ.Open dsn
    SQL = "SELECT * FROM OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0', 'Data Source=" & Label2 & ";Extended Properties=Excel 8.0')...[sales$]"
    Set RST = OBJ.Execute(SQL)
    If RST.State = 1 Then
        Do While Not RST.EOF
            OBJ1.Open dsn
            SQL1 = "select * from AM_salesman where kodesales = '" & RST!kodesales & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If RST1.EOF Then
                SQL1 = "INSERT INTO AM_salesman"
                SQL1 = SQL1 + " (KodeSales"
                SQL1 = SQL1 + ", NamaSales"
                SQL1 = SQL1 + ", TargetSales"
                SQL1 = SQL1 + ", InsentiveSales"
                SQL1 = SQL1 + ", IdEntry"
                SQL1 = SQL1 + ", DateEntry"
                SQL1 = SQL1 + ", IdUpdate"
                SQL1 = SQL1 + ", DateUpdate)"
                
                SQL1 = SQL1 + "VALUES"
                SQL1 = SQL1 + " ('" & RST!kodesales & "'"
                SQL1 = SQL1 + ", '" & RST!namasales & "'"
                SQL1 = SQL1 + ", convert(money,'" & RST!targetsales & "')"
                SQL1 = SQL1 + ", convert(money,'" & RST!insentivesales & "')"
                SQL1 = SQL1 + ", '" & kuser & "'"
                SQL1 = SQL1 + ",Convert(dateTime, '" & Month(RST!dateentry) & "/" & Day(RST!dateentry) & "/" & Year(RST!dateentry) & "')"
                SQL1 = SQL1 + ", '" & kuser & "'"
                SQL1 = SQL1 + ",Convert(dateTime, '" & Month(RST!dateupdate) & "/" & Day(RST!dateupdate) & "/" & Year(RST!dateupdate) & "'))"
                Set RST1 = OBJ1.Execute(SQL1)
            End If
            OBJ1.Close
            
            RST.MoveNext
        Loop
        
        SQL = "SELECT count(kodesales)'totalbaris' FROM OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0', 'Data Source=" & Label2 & ";Extended Properties=Excel 8.0')...[sales$]"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then j = RST!totalbaris Else j = 0
        Label3 = Label3 + vbCrLf + "   Import Salesman, " & j & " rows affected."
    End If
    OBJ.Close
    'area
    OBJ.Open dsn
    SQL = "SELECT * FROM OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0', 'Data Source=" & Label2 & ";Extended Properties=Excel 8.0')...[area$]"
    Set RST = OBJ.Execute(SQL)
    If RST.State = 1 Then
        Do While Not RST.EOF
            OBJ1.Open dsn
            SQL1 = "select * from AM_area where kode = '" & RST!kode & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If RST1.EOF Then
                SQL1 = "INSERT INTO AM_area"
                SQL1 = SQL1 + "(Kode"
                SQL1 = SQL1 + ",Nama"
                SQL1 = SQL1 + ",identry"
                SQL1 = SQL1 + ",dateupdate"
                SQL1 = SQL1 + ",idupdate"
                SQL1 = SQL1 + ",Dateentry)"
            
                SQL1 = SQL1 + "VALUES"
                SQL1 = SQL1 + " ('" & RST!kode & "'"
                SQL1 = SQL1 + ", '" & RST!nama & "'"
                SQL1 = SQL1 + ", '" & kuser & "'"
                SQL1 = SQL1 + ",Convert(dateTime, '" & Month(RST!dateupdate) & "/" & Day(RST!dateupdate) & "/" & Year(RST!dateupdate) & "')"
                SQL1 = SQL1 + ", '" & kuser & "'"
                SQL1 = SQL1 + ",Convert(dateTime, '" & Month(RST!dateentry) & "/" & Day(RST!dateentry) & "/" & Year(RST!dateentry) & "'))"
                Set RST1 = OBJ1.Execute(SQL1)
            End If
            OBJ1.Close
            
            RST.MoveNext
        Loop
        
        SQL = "SELECT count(kode)'totalbaris' FROM OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0', 'Data Source=" & Label2 & ";Extended Properties=Excel 8.0')...[area$]"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then j = RST!totalbaris Else j = 0
        Label3 = Label3 + vbCrLf + "   Import Area, " & j & " rows affected."
    End If
    OBJ.Close
    'unit
    OBJ.Open dsn
    SQL = "SELECT * FROM OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0', 'Data Source=" & Label2 & ";Extended Properties=Excel 8.0')...[unit$]"
    Set RST = OBJ.Execute(SQL)
    If RST.State = 1 Then
        Do While Not RST.EOF
            OBJ1.Open dsn
            SQL1 = "select * from AM_unit where kodesatuan = '" & RST!kodesatuan & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If RST1.EOF Then
                SQL1 = "INSERT INTO AM_unit"
                SQL1 = SQL1 + " (KodeSatuan"
                SQL1 = SQL1 + ", NamaSatuan"
                SQL1 = SQL1 + ", init"
                SQL1 = SQL1 + ", IdEntry"
                SQL1 = SQL1 + ", DateEntry"
                SQL1 = SQL1 + ", IdUpdate"
                SQL1 = SQL1 + ", DateUpdate)"
            
                SQL1 = SQL1 + "VALUES"
                SQL1 = SQL1 + " ('" & RST!kodesatuan & "'"
                SQL1 = SQL1 + ", '" & RST!namasatuan & "'"
                SQL1 = SQL1 + ", '" & RST!init & "'"
                SQL1 = SQL1 + ", '" & kuser & "'"
                SQL1 = SQL1 + ",Convert(dateTime, '" & Month(RST!dateentry) & "/" & Day(RST!dateentry) & "/" & Year(RST!dateentry) & "')"
                SQL1 = SQL1 + ", '" & kuser & "'"
                SQL1 = SQL1 + ",Convert(dateTime, '" & Month(RST!dateupdate) & "/" & Day(RST!dateupdate) & "/" & Year(RST!dateupdate) & "'))"
                Set RST1 = OBJ1.Execute(SQL1)
            End If
            OBJ1.Close
            
            RST.MoveNext
        Loop
        
        SQL = "SELECT count(kodesatuan)'totalbaris' FROM OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0', 'Data Source=" & Label2 & ";Extended Properties=Excel 8.0')...[unit$]"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then j = RST!totalbaris Else j = 0
        Label3 = Label3 + vbCrLf + "   Import Satuan, " & j & " rows affected."
    End If
    OBJ.Close
    'produk
    OBJ.Open dsn
    SQL = "SELECT * FROM OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0', 'Data Source=" & Label2 & ";Extended Properties=Excel 8.0')...[produk$]"
    Set RST = OBJ.Execute(SQL)
    If RST.State = 1 Then
        Do While Not RST.EOF
            OBJ1.Open dsn
            SQL1 = "select * from AM_produk where kodeproduk = '" & RST!kodeproduk & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If RST1.EOF Then
                SQL1 = "INSERT INTO AM_produk"
                SQL1 = SQL1 + "(Kodeproduk"
                SQL1 = SQL1 + ",Namaproduk"
                SQL1 = SQL1 + ",identry"
                SQL1 = SQL1 + ",dateupdate"
                SQL1 = SQL1 + ",idupdate"
                SQL1 = SQL1 + ",Dateentry)"
            
                SQL1 = SQL1 + "VALUES"
                SQL1 = SQL1 + " ('" & RST!kodeproduk & "'"
                SQL1 = SQL1 + ", '" & RST!namaproduk & "'"
                SQL1 = SQL1 + ", '" & kuser & "'"
                SQL1 = SQL1 + ",Convert(dateTime, '" & Month(RST!dateupdate) & "/" & Day(RST!dateupdate) & "/" & Year(RST!dateupdate) & "')"
                SQL1 = SQL1 + ", '" & kuser & "'"
                SQL1 = SQL1 + ",Convert(dateTime, '" & Month(RST!dateentry) & "/" & Day(RST!dateentry) & "/" & Year(RST!dateentry) & "'))"
                Set RST1 = OBJ1.Execute(SQL1)
            End If
            OBJ1.Close
            
            RST.MoveNext
        Loop
        
        SQL = "SELECT count(kodeproduk)'totalbaris' FROM OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0', 'Data Source=" & Label2 & ";Extended Properties=Excel 8.0')...[produk$]"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then j = RST!totalbaris Else j = 0
        Label3 = Label3 + vbCrLf + "   Import Category, " & j & " rows affected."
    End If
    OBJ.Close
    'item
    OBJ.Open dsn
    SQL = "SELECT * FROM OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0', 'Data Source=" & Label2 & ";Extended Properties=Excel 8.0')...[item$]"
    Set RST = OBJ.Execute(SQL)
    If RST.State = 1 Then
        Do While Not RST.EOF
            OBJ1.Open dsn
            SQL1 = "select * from AM_itemmst where kodebarang = '" & RST!kodebarang & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If RST1.EOF Then
                SQL1 = "INSERT INTO AM_itemmst"
                SQL1 = SQL1 + "(Kodebarang"
                SQL1 = SQL1 + ",Namabarang"
                SQL1 = SQL1 + ",jenisbarang"
                SQL1 = SQL1 + ",kodeproduk"
                SQL1 = SQL1 + ",identry"
                SQL1 = SQL1 + ",dateupdate"
                SQL1 = SQL1 + ",idupdate"
                SQL1 = SQL1 + ",Dateentry)"
            
                SQL1 = SQL1 + "VALUES"
                SQL1 = SQL1 + " ('" & RST!kodebarang & "'"
                SQL1 = SQL1 + ", '" & RST!namabarang & "'"
                SQL1 = SQL1 + ", '" & RST!jenisbarang & "'"
                SQL1 = SQL1 + ", '" & RST!kodeproduk & "'"
                SQL1 = SQL1 + ", '" & kuser & "'"
                SQL1 = SQL1 + ",Convert(dateTime, '" & Month(RST!dateupdate) & "/" & Day(RST!dateupdate) & "/" & Year(RST!dateupdate) & "')"
                SQL1 = SQL1 + ", '" & kuser & "'"
                SQL1 = SQL1 + ",Convert(dateTime, '" & Month(RST!dateentry) & "/" & Day(RST!dateentry) & "/" & Year(RST!dateentry) & "'))"
                Set RST1 = OBJ1.Execute(SQL1)
            End If
            OBJ1.Close
            
            RST.MoveNext
        Loop
        SQL = "SELECT count(kodebarang)'totalbaris' FROM OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0', 'Data Source=" & Label2 & ";Extended Properties=Excel 8.0')...[item$]"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then j = RST!totalbaris Else j = 0
        Label3 = Label3 + vbCrLf + "   Import Item, " & j & " rows affected."
    End If
    OBJ.Close
    'itemdtl
    OBJ.Open dsn
    SQL = "SELECT * FROM OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0', 'Data Source=" & Label2 & ";Extended Properties=Excel 8.0')...[item_$]"
    Set RST = OBJ.Execute(SQL)
    If RST.State = 1 Then
        Do While Not RST.EOF
            OBJ1.Open dsn
            SQL1 = "select * from AM_itemdtl where kodebarang = '" & RST!kodebarang & "' and kodesatuan = '" & RST!kodesatuan & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If RST1.EOF Then
                SQL1 = "INSERT INTO AM_itemdtl"
                SQL1 = SQL1 + "(Kodebarang"
                SQL1 = SQL1 + ",Namabarang"
                SQL1 = SQL1 + ",level_"
                SQL1 = SQL1 + ",kodesatuan"
                SQL1 = SQL1 + ",pricesale"
                SQL1 = SQL1 + ",hpprata2"
                SQL1 = SQL1 + ",konversi)"
            
                SQL1 = SQL1 + "VALUES"
                SQL1 = SQL1 + " ('" & RST!kodebarang & "'"
                SQL1 = SQL1 + ", '" & RST!namabarang & "'"
                SQL1 = SQL1 + ", convert(money,'" & RST!level_ & "')"
                SQL1 = SQL1 + ", '" & RST!kodesatuan & "'"
                SQL1 = SQL1 + ", convert(money,'" & RST!pricesale & "')"
                SQL1 = SQL1 + ", convert(money,'" & RST!hpprata2 & "')"
                SQL1 = SQL1 + ", convert(money,'" & RST!konversi & "'))"
                Set RST1 = OBJ1.Execute(SQL1)
            End If
            OBJ1.Close
            
            RST.MoveNext
        Loop
    End If
    OBJ.Close
    'rule
    OBJ.Open dsn
    SQL = "SELECT * FROM OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0', 'Data Source=" & Label2 & ";Extended Properties=Excel 8.0')...[rule$]"
    Set RST = OBJ.Execute(SQL)
    If RST.State = 1 Then
        Do While Not RST.EOF
            OBJ1.Open dsn
            SQL1 = "select * from AM_itemcode where lev = '" & RST!lev & "' and kode = '" & RST!kode & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If RST1.EOF Then
                SQL1 = "INSERT INTO AM_itemcode"
                SQL1 = SQL1 + "(lev"
                SQL1 = SQL1 + ",kode"
                SQL1 = SQL1 + ",ket)"
            
                SQL1 = SQL1 + "VALUES"
                SQL1 = SQL1 + " ('" & RST!lev & "'"
                SQL1 = SQL1 + ", '" & RST!kode & "'"
                SQL1 = SQL1 + ", '" & RST!ket & "')"
                Set RST1 = OBJ1.Execute(SQL1)
            End If
            OBJ1.Close
            
            RST.MoveNext
        Loop
        
        SQL = "SELECT count(lev)'totalbaris' FROM OPENDATASOURCE('Microsoft.ACE.OLEDB.12.0', 'Data Source=" & Label2 & ";Extended Properties=Excel 8.0')...[rule$]"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then j = RST!totalbaris Else j = 0
        Label3 = Label3 + vbCrLf + "   Import Rule, " & j & " rows affected."
    End If
    OBJ.Close
    
    Kill Label2
    
    MsgBox "Import Complete.", vbInformation, "Information"
    
    cmdimport.Enabled = False
    Exit Sub
Err_handler:
    MsgBox Err.Description, vbCritical, AppName
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

