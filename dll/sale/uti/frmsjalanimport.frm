VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmsjalanimport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Surat Jalan"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Format          =   91553795
      CurrentDate     =   37426
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1320
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
      MICON           =   "frmsjalanimport.frx":0000
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
      TabIndex        =   0
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
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
      MICON           =   "frmsjalanimport.frx":031A
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
      Left            =   1680
      TabIndex        =   1
      Top             =   1320
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
      MICON           =   "frmsjalanimport.frx":0634
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
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Import Surat Jalan harus dijalankan di Server atau Workstation yang ada SQL Servernya."
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
      Height          =   405
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   3375
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
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Left            =   -120
      TabIndex        =   5
      Top             =   0
      Width           =   3975
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
Attribute VB_Name = "frmsjalanimport"
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

Dim i As Integer

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function

Private Sub cmdclear_Click()
    Label2 = Dir$("c:\*_sj.trs")
    If Label2 <> "" Then
        Label2 = "c:\" & Label2
        Label3 = "File Found." & vbCrLf & Label2
    Else
        Label3 = "File Not Found."
    End If
    cmdclear.Enabled = False
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdimport_Click()
On Error Resume Next
    If Label2 = "" Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    Label3 = "  please wait a moment ..."
    
    If MsgBox("Please make sure file exsist and valid." & vbCrLf & "Are you sure want to continue import file " & Label2 & " ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
        
    OBJ.Open dsn
    SQL = "SELECT * FROM OPENDATASOURCE('Microsoft.Jet.OLEDB.4.0', 'Data Source=" & Label2 & ";Extended Properties=Excel 8.0')...[sj$]"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ1.Open dsn
        SQL1 = "update AM_sjapp set flag2 = '2' where flag2 = '1'"
        Set RST1 = OBJ1.Execute(SQL1)
        OBJ1.Close
        
        Do While Not RST.EOF
            OBJ1.Open dsn
            SQL1 = "select * from AM_sjapp where nosj = '" & RST!nosj & "' and kodebarang = '" & RST!kodebarang & "' and kodesatuan = '" & RST!kodesatuan & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If RST1.EOF Then
                SQL1 = "select * from AM_sjappdelete where nosj = '" & RST!nosj & "' and kodebarang = '" & RST!kodebarang & "' and kodesatuan = '" & RST!kodesatuan & "'"
                Set RST1 = OBJ1.Execute(SQL1)
                If RST1.EOF Then
                    SQL1 = "INSERT INTO AM_sjapp"
                    SQL1 = SQL1 + " (nosj"
                    SQL1 = SQL1 + ", Tglsj"
                    SQL1 = SQL1 + ", kodecust"
                    SQL1 = SQL1 + ", kodesales"
                    SQL1 = SQL1 + ", nopo"
                    SQL1 = SQL1 + ", noso"
                    SQL1 = SQL1 + ", Kodegudang"
                    SQL1 = SQL1 + ", tglkirim"
                    SQL1 = SQL1 + ", via"
                    SQL1 = SQL1 + ", Kodebarang"
                    SQL1 = SQL1 + ", qty"
                    SQL1 = SQL1 + ", keterangan"
                    SQL1 = SQL1 + ", kodesatuan"
                    SQL1 = SQL1 + ", lineitem"
                    SQL1 = SQL1 + ", bn"
                    SQL1 = SQL1 + ", flag1"
                    SQL1 = SQL1 + ", flag2)"
                
                    SQL1 = SQL1 + "VALUES"
                    SQL1 = SQL1 + " ('" & RST!nosj & "'"
                    SQL1 = SQL1 + ",Convert(dateTime, '" & Month(RST!TglSJ) & "/" & Day(RST!TglSJ) & "/" & Year(RST!TglSJ) & "')"
                    SQL1 = SQL1 + ", '" & RST!kodecust & "'"
                    SQL1 = SQL1 + ", '" & RST!kodesales & "'"
                    SQL1 = SQL1 + ", '" & RST!nopo & "'"
                    SQL1 = SQL1 + ", '" & RST!noso & "'"
                    SQL1 = SQL1 + ", '" & RST!kodegudang & "'"
                    SQL1 = SQL1 + ",Convert(dateTime, '" & Month(RST!tglkirim) & "/" & Day(RST!tglkirim) & "/" & Year(RST!tglkirim) & "')"
                    SQL1 = SQL1 + ", '" & RST!via & "'"
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
            End If
            OBJ1.Close
            
            RST.MoveNext
        Loop
    End If
    OBJ.Close
    
    'customer
    OBJ.Open dsn
    SQL = "SELECT * FROM OPENDATASOURCE('Microsoft.Jet.OLEDB.4.0', 'Data Source=" & Label2 & ";Extended Properties=Excel 8.0')...[customer$]"
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
        
    Kill Label2
    
    OBJ.Open dsn
    SQL = "select count(nosj)'totalbaris' from AM_sjapp where flag2 = '1'"
    Set RST = OBJ.Execute(SQL)
    i = RST!totalbaris
    OBJ.Close
    
    MsgBox "Import Complete.", vbInformation, "Information"
    Label3 = "  " & i & " rows affected. (Surat Jalan)"
    cmdimport.Enabled = False
End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='424' and b.kodeuser = '1" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If RST.EOF Then
    '        MsgBox "User Rights Denied !!" & vbCrLf & _
    '        "Please contact your Administrator.", vbCritical, "User Rights"
    '        OBJ.Close
    '        Unload Me
    '        Exit Sub
    '    End If
    '    OBJ.Close
    'End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

