VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmsoexport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export Sales Order"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3255
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   2160
      TabIndex        =   3
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
      MICON           =   "frmsoexport.frx":0000
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
      TabIndex        =   2
      Top             =   2520
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
      MICON           =   "frmsoexport.frx":031A
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
      Top             =   1680
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
      Format          =   91750403
      CurrentDate     =   37426
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   2040
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
      Format          =   91750403
      CurrentDate     =   37426
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Export Sales Order hanya di jalankan di cabang Sparta yang tidak memiliki stock termasuk Tamansari"
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
      Height          =   675
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   2775
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
      TabIndex        =   7
      Top             =   1710
      Width           =   1095
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
      TabIndex        =   6
      Top             =   2070
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
      TabIndex        =   4
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Left            =   -120
      TabIndex        =   5
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "frmsoexport"
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
Dim objsheet1, objsheet2, objsheet3, objsheet4, objsheet6, objsheet7, objsheet8, objsheet9, objsheet10 As Excel.Worksheet

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
    If (date1 > date2) Then
        MsgBox "From Date Greather Then To Date.", vbExclamation, "Warning"
        Exit Sub
    End If
     
    If MsgBox("Please make sure parameters are correct." & vbCrLf & "Are you sure want to continue ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "SELECT count(noso)'baris' FROM AM_sohdr WHERE tglso >= '" & tanggal1 & "' and tglso <= '" & tanggal2 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST!baris = 0 Then
        MsgBox "There is no record to export.", vbExclamation, "Warning"
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    Label2 = "  please wait a moment ..."
    'so mulai
    OBJ.Open dsn
    SQL = "SELECT b.noso,b.tglso,b.kodecust,b.kodesales,b.nopo,b.jasa,a.kodebarang,a.qty,a.keterangan,a.kodesatuan,a.lineitem,a.bn FROM am_solin a left join AM_sohdr b on a.noso=b.noso WHERE b.tglso >= '" & tanggal1 & "' and b.tglso <= '" & tanggal2 & "'"
    Set RST = OBJ.Execute(SQL)
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
        
            SQL1 = SQL1 + " VALUES"
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
            SQL1 = SQL1 + ", '0')"
            Set RST1 = OBJ1.Execute(SQL1)
        End If
        OBJ1.Close
    
        RST.MoveNext
    Loop
    OBJ.Close
        
    OBJ.Open dsn
    SQL = "select count(noso)'totalbaris' from AM_soapp where tglso >= '" & tanggal1 & "' and tglso <= '" & tanggal2 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then j = RST!totalbaris Else j = 0
    
    SQL = "select * from AM_soapp where tglso >= '" & tanggal1 & "' and tglso <= '" & tanggal2 & "'"
    Set RST = OBJ.Execute(SQL)
    
    Set objapp = New Excel.Application
    objapp.Visible = False
    Set objbook = objapp.Workbooks.Add
    Set objsheet1 = objapp.Worksheets.Add
    
    objsheet1.Name = "so"
    objsheet1.Columns.AutoFit
    
    For i = 0 To RST.Fields.Count - 1 Step 1
        objsheet1.Cells(1, i + 1).Value = RST.Fields(i).Name
    Next i
    
    objsheet1.Range("A2").CopyFromRecordset RST
    Label2 = "   Sales Order, " & j & " rows affected."
    OBJ.Close
    
    'customer mulai
    OBJ1.Open dsn
    SQL1 = "select distinct b.* from am_sohdr a left join am_customer b on a.kodecust=b.kodecust where a.tglso>='" & tanggal1 & "' and a.tglso <= '" & tanggal2 & "'"
    Set RST1 = OBJ1.Execute(SQL1)
        
    Set objsheet2 = objapp.Worksheets.Add
    
    objsheet2.Name = "customer"
    objsheet2.Columns.AutoFit
    
    For i = 0 To RST1.Fields.Count - 1 Step 1
        objsheet2.Cells(1, i + 1).Value = RST1.Fields(i).Name
    Next i
    
    objsheet2.Range("A2").CopyFromRecordset RST1
    
    'sales mulai
    SQL1 = "select * from am_salesman"
    Set RST1 = OBJ1.Execute(SQL1)
        
    Set objsheet3 = objapp.Worksheets.Add
    
    objsheet3.Name = "sales"
    objsheet3.Columns.AutoFit
    
    For i = 0 To RST1.Fields.Count - 1 Step 1
        objsheet3.Cells(1, i + 1).Value = RST1.Fields(i).Name
    Next i
    i = 2
    Do While Not RST1.EOF
        objsheet3.Cells(i, 1).Value = "'" & RST1.Fields(0).Value
        objsheet3.Cells(i, 2).Value = RST1.Fields(1).Value
        objsheet3.Cells(i, 3).Value = RST1.Fields(2).Value
        objsheet3.Cells(i, 4).Value = RST1.Fields(3).Value
        objsheet3.Cells(i, 5).Value = RST1.Fields(4).Value
        objsheet3.Cells(i, 6).Value = RST1.Fields(5).Value
        objsheet3.Cells(i, 7).Value = RST1.Fields(6).Value
        objsheet3.Cells(i, 8).Value = RST1.Fields(7).Value
        
        RST1.MoveNext
        i = i + 1
    Loop
    OBJ1.Close
    
    'area mulai
    OBJ1.Open dsn
    SQL1 = "select * from am_area"
    Set RST1 = OBJ1.Execute(SQL1)
        
    Set objsheet4 = objapp.Worksheets.Add
    
    objsheet4.Name = "area"
    objsheet4.Columns.AutoFit
    
    For i = 0 To RST1.Fields.Count - 1 Step 1
        objsheet4.Cells(1, i + 1).Value = RST1.Fields(i).Name
    Next i
    i = 2
    Do While Not RST1.EOF
        objsheet4.Cells(i, 1).Value = "'" & RST1.Fields(0).Value
        objsheet4.Cells(i, 2).Value = RST1.Fields(1).Value
        objsheet4.Cells(i, 3).Value = RST1.Fields(2).Value
        objsheet4.Cells(i, 4).Value = RST1.Fields(3).Value
        objsheet4.Cells(i, 5).Value = RST1.Fields(4).Value
        objsheet4.Cells(i, 6).Value = RST1.Fields(5).Value
        
        RST1.MoveNext
        i = i + 1
    Loop
    OBJ1.Close
    
    'unit mulai
    OBJ1.Open dsn
    SQL1 = "select * from am_unit"
    Set RST1 = OBJ1.Execute(SQL1)
        
    Set objsheet6 = objapp.Worksheets.Add
    
    objsheet6.Name = "unit"
    objsheet6.Columns.AutoFit
    
    For i = 0 To RST1.Fields.Count - 1 Step 1
        objsheet6.Cells(1, i + 1).Value = RST1.Fields(i).Name
    Next i
    
    objsheet6.Range("A2").CopyFromRecordset RST1
    
    'produk mulai
    SQL1 = "select * from am_produk"
    Set RST1 = OBJ1.Execute(SQL1)
        
    Set objsheet7 = objapp.Worksheets.Add
    
    objsheet7.Name = "produk"
    objsheet7.Columns.AutoFit
    
    For i = 0 To RST1.Fields.Count - 1 Step 1
        objsheet7.Cells(1, i + 1).Value = RST1.Fields(i).Name
    Next i
    
    objsheet7.Range("A2").CopyFromRecordset RST1
    OBJ1.Close
    
    'item master mulai
    OBJ1.Open dsn
    SQL1 = "select * from am_itemmst"
    Set RST1 = OBJ1.Execute(SQL1)
        
    Set objsheet8 = objapp.Worksheets.Add
    
    objsheet8.Name = "item"
    objsheet8.Columns.AutoFit
    
    For i = 0 To RST1.Fields.Count - 1 Step 1
        objsheet8.Cells(1, i + 1).Value = RST1.Fields(i).Name
    Next i
    
    objsheet8.Range("A2").CopyFromRecordset RST1
    
    'item detail mulai
    SQL1 = "select * from AM_itemdtl"
    Set RST1 = OBJ1.Execute(SQL1)
    
    Set objsheet9 = objapp.Worksheets.Add
    
    objsheet9.Name = "item_"
    objsheet9.Columns.AutoFit
    
    For i = 0 To RST1.Fields.Count - 1 Step 1
        objsheet9.Cells(1, i + 1).Value = RST1.Fields(i).Name
    Next i
    
    objsheet9.Range("A2").CopyFromRecordset RST1
    
    'rule mulai
    SQL1 = "select * from am_itemcode"
    Set RST1 = OBJ1.Execute(SQL1)
        
    Set objsheet10 = objapp.Worksheets.Add
    
    objsheet10.Name = "rule"
    objsheet10.Columns.AutoFit
    
    For i = 0 To RST1.Fields.Count - 1 Step 1
        objsheet10.Cells(1, i + 1).Value = RST1.Fields(i).Name
    Next i
    i = 2
    Do While Not RST1.EOF
        objsheet10.Cells(i, 1).Value = "'" & RST1.Fields(0).Value
        objsheet10.Cells(i, 2).Value = "'" & RST1.Fields(1).Value
        objsheet10.Cells(i, 3).Value = RST1.Fields(2).Value
        
        RST1.MoveNext
        i = i + 1
    Loop
    OBJ1.Close
    '==============
    objbook.SaveAs "c:\" & Year(Date) & Month(Date) & Day(Date) & "_" & Hour(Time) & Minute(Time) & "_so.trs"
    objbook.Close False
    Set objbook = Nothing
    Set objsheet1 = Nothing
    Set objsheet2 = Nothing
    Set objsheet3 = Nothing
    Set objsheet4 = Nothing
    Set objsheet6 = Nothing
    Set objsheet7 = Nothing
    Set objsheet8 = Nothing
    Set objsheet9 = Nothing
    Set objsheet10 = Nothing
    Set objapp = Nothing
    
    Label2 = Label2 + vbCrLf + "   File saved on Local Disk (C:\)"
    
    OBJ.Open dsn
    SQL = "update AM_sohdr set flag = '1' where tglso >= '" & tanggal1 & "' and tglso <= '" & tanggal2 & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Export Complete.", vbInformation, "Information"
    
    cmdclear.Enabled = False
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='394' and b.kodeuser = '1" & kuser & "'"
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
