VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmsjalanexport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export Surat Jalan"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3255
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   1920
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
      MICON           =   "frmsjalanexport.frx":0000
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
      Left            =   1200
      TabIndex        =   1
      Top             =   1920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Export"
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
      MICON           =   "frmsjalanexport.frx":031A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   1080
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
      Format          =   143327235
      CurrentDate     =   37426
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   1440
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
      Format          =   143327235
      CurrentDate     =   37426
   End
   Begin VB.Label Label3 
      Caption         =   "To Date"
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
      TabIndex        =   7
      Top             =   1470
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "From Date"
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
      TabIndex        =   6
      Top             =   1110
      Width           =   1095
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Ready to Export !!"
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
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   -240
      TabIndex        =   4
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "frmsjalanexport"
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

Dim objapp As Excel.Application
Dim objbook As Excel.Workbook
Dim objsheet1, objsheet2 As Excel.Worksheet

Dim i, j As Integer

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function

Function tanggal1()
    tanggal1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggal2()
    tanggal2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function

Private Sub cmdclear_Click()
    If nmuser <> "santo" Or nmuser <> "Creator" Then
        MsgBox "Access denied", vbCritical, "Hak Akses Pengguna"
        Exit Sub
    End If
    If (date1 > date2) Then
        MsgBox "From Date Greather Then To Date.", vbExclamation, "Warning"
        Exit Sub
    End If

    If MsgBox("Please make sure parameters are correct." & vbCrLf & "Are you sure want to continue ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "SELECT count(nosj)'baris' FROM AM_sjhdr WHERE tglsj >= '" & tanggal1 & "' and tglsj <= '" & tanggal2 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST!baris = 0 Then
        MsgBox "There is no record to export.", vbExclamation, "Warning"
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "SELECT b.nosj,b.tglsj,b.kodecust,b.kodesales,b.nopo,b.noso,b.kodegudang,b.tglkirim,b.via,a.kodebarang,a.qty,a.keterangan,a.kodesatuan,a.lineitem,a.bn FROM am_sjlin a left join AM_sjhdr b on a.nosj=b.nosj WHERE b.tglsj >= '" & tanggal1 & "' and b.tglsj <= '" & tanggal2 & "'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        OBJ1.Open dsn
        SQL1 = "select * from AM_sjapp where nosj = '" & RST!nosj & "' and kodebarang = '" & RST!kodebarang & "' and kodesatuan = '" & RST!kodesatuan & "'"
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
        
            SQL1 = SQL1 + " VALUES"
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
            SQL1 = SQL1 + ", '0')"
            Set RST1 = OBJ1.Execute(SQL1)
        End If
        OBJ1.Close
    
        RST.MoveNext
    Loop
    OBJ.Close
    
    OBJ.Open dsn
    Label2 = "  please wait a moment ..."
    
    SQL = "select count(nosj)'totalbaris' from AM_sjapp where flag2 <> '9' and tglsj >= '" & tanggal1 & "' and tglsj <= '" & tanggal2 & "'"
    Set RST = OBJ.Execute(SQL)
    j = RST!totalbaris
        
    SQL = "select * from AM_sjapp where flag2 <> '9' and tglsj >= '" & tanggal1 & "' and tglsj <= '" & tanggal2 & "'"
    Set RST = OBJ.Execute(SQL)
    
    Set objapp = New Excel.Application
    objapp.Visible = False
    Set objbook = objapp.Workbooks.Add
    Set objsheet1 = objapp.Worksheets.Add
    
    objsheet1.Name = "sj"
    objsheet1.Columns.AutoFit
    
    For i = 0 To RST.Fields.Count - 1 Step 1
        objsheet1.Cells(1, i + 1).Value = RST.Fields(i).Name
    Next i
    
    objsheet1.Range("A2").CopyFromRecordset RST
    OBJ.Close
    
    'customer mulai
    OBJ.Open dsn
    SQL = "select top 20 * from am_customer where kodecust<'C-20000' order by kodecust desc"
    Set RST = OBJ.Execute(SQL)
        
    Set objsheet2 = objapp.Worksheets.Add
    
    objsheet2.Name = "customer"
    objsheet2.Columns.AutoFit
    
    For i = 0 To RST.Fields.Count - 1 Step 1
        objsheet2.Cells(1, i + 1).Value = RST.Fields(i).Name
    Next i
    
    objsheet2.Range("A2").CopyFromRecordset RST
    OBJ.Close
    
    objbook.SaveAs "c:\" & Year(Date) & Month(Date) & Day(Date) & "_" & Hour(Time) & Minute(Time) & "_sj.trs"
    objbook.Close False
    Set objbook = Nothing
    Set objsheet1 = Nothing
    Set objsheet2 = Nothing
    Set objapp = Nothing
    
    OBJ.Open dsn
    SQL = "update AM_sjhdr set via2 = '2' where tglsj >= '" & tanggal1 & "' and tglsj <= '" & tanggal2 & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Export Complete.", vbInformation, "Information"
    Label2 = "  " & j & " rows affected. (Surat Jalan)" & vbCrLf & _
    "   " & "File saved on Local Disk (C:\)"
    cmdclear.Enabled = False
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    'vallidasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='414' and b.kodeuser = '1" & kuser & "'"
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

Private Sub Form_Load()
    
    
    date1 = Date
    date2 = Date
End Sub
