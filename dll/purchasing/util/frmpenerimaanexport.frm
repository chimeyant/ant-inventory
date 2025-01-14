VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmpenerimaanexport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export Data"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   2040
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
      MICON           =   "frmpenerimaanexport.frx":0000
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
      TabIndex        =   2
      Top             =   2025
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Export Penerimaan Barang"
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
      MICON           =   "frmpenerimaanexport.frx":031A
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
      Top             =   1200
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
      Format          =   106823683
      CurrentDate     =   37426
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   1560
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
      Format          =   106823683
      CurrentDate     =   37426
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
      Height          =   855
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   -240
      TabIndex        =   7
      Top             =   0
      Width           =   5655
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
      TabIndex        =   5
      Top             =   1590
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
      TabIndex        =   4
      Top             =   1230
      Width           =   1095
   End
End
Attribute VB_Name = "frmpenerimaanexport"
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
Dim objsheet1, objsheet2, objsheet3, objsheet4, objsheet5, objsheet6, objsheet7, objsheet8, objsheet9 As Excel.Worksheet

Dim i, j, k As Integer

Function tanggal1()
    tanggal1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggal2()
    tanggal2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function

Private Sub cmdclear_Click()
    If (date1 > date2) Then
        MsgBox "From Date Greather Then To Date.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If MsgBox("Please make sure parameters are correct." & vbCrLf & "Are you sure want to continue ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "SELECT count(nobeli)'baris' FROM AM_belihdr WHERE tglbeli >= '" & tanggal1 & "' and tglbeli <= '" & tanggal2 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST!baris = 0 Then
        MsgBox "There is no record to export.", vbExclamation, "Warning"
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "SELECT b.nobeli,b.tglbeli,b.nopo,c.kodecur,c.nilaikurs,a.kodebarang,a.qty,d.price,a.kodesatuan,a.lineitem,c.kodesupp,b.driver,c.ket1,c.ket2,c.ket3 FROM am_belilin a left join AM_belihdr b on a.nobeli=b.nobeli left join am_pohdr c on b.nopo=c.nopo left join am_polin d on a.kodebarang=d.kodebarang and d.nopo=b.nopo WHERE b.tglbeli >= '" & tanggal1 & "' and b.tglbeli <= '" & tanggal2 & "'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        OBJ1.Open dsn
        SQL1 = "select * from AM_beliapp where nobeli = '" & RST!nobeli & "' and kodebarang = '" & RST!kodebarang & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        If RST1.EOF Then
            SQL1 = "INSERT INTO AM_beliapp"
            SQL1 = SQL1 + " (noBeli"
            SQL1 = SQL1 + ", TglBeli"
            SQL1 = SQL1 + ", nopo"
            SQL1 = SQL1 + ", ref1"
            SQL1 = SQL1 + ", ref2"
            SQL1 = SQL1 + ", kodesupp"
            SQL1 = SQL1 + ", kodecur"
            SQL1 = SQL1 + ", nilaikurs"
            SQL1 = SQL1 + ", Kodebarang"
            SQL1 = SQL1 + ", qty"
            SQL1 = SQL1 + ", Price"
            SQL1 = SQL1 + ", kodesatuan"
            SQL1 = SQL1 + ", keterangan"
            SQL1 = SQL1 + ", keterangan2"
            SQL1 = SQL1 + ", keterangan3"
            SQL1 = SQL1 + ", keterangan4"
            SQL1 = SQL1 + ", ppn"
            SQL1 = SQL1 + ", lineitem"
            SQL1 = SQL1 + ", flag1"
            SQL1 = SQL1 + ", flag2)"
        
            SQL1 = SQL1 + " VALUES"
            SQL1 = SQL1 + " ('" & RST!nobeli & "'"
            SQL1 = SQL1 + ",Convert(dateTime, '" & Month(RST!tglbeli) & "/" & Day(RST!tglbeli) & "/" & Year(RST!tglbeli) & "')"
            SQL1 = SQL1 + ", '" & RST!nopo & "'"
            SQL1 = SQL1 + ", ''"
            SQL1 = SQL1 + ", ''"
            SQL1 = SQL1 + ", '" & RST!kodesupp & "'"
            SQL1 = SQL1 + ", '" & RST!kodecur & "'"
            SQL1 = SQL1 + ",Convert (Money, '" & RST!nilaikurs & "')"
            SQL1 = SQL1 + ", '" & RST!kodebarang & "'"
            SQL1 = SQL1 + ",Convert (Money, '" & RST!qty & "')"
            SQL1 = SQL1 + ",Convert (Money, '" & RST!Price & "')"
            SQL1 = SQL1 + ", '" & RST!kodesatuan & "'"
            SQL1 = SQL1 + ", '" & RST!driver & "'"
            SQL1 = SQL1 + ", '" & RST!ket1 & "'"
            SQL1 = SQL1 + ", '" & RST!ket2 & "'"
            SQL1 = SQL1 + ", '" & RST!ket3 & "'"
            SQL1 = SQL1 + ",Convert (Money, '0')"
            SQL1 = SQL1 + ",Convert (numeric, '" & RST!lineitem & "')"
            SQL1 = SQL1 + ", '0'"
            SQL1 = SQL1 + ", '0')"
            Set RST1 = OBJ1.Execute(SQL1)
        End If
        OBJ1.Close
    
        RST.MoveNext
    Loop
    OBJ.Close
    
    Label2 = "  please wait a moment ..."
    
    OBJ.Open dsn
    SQL = "select count(nobeli)'totalbaris' from AM_beliapp where tglbeli >= '" & tanggal1 & "' and tglbeli <= '" & tanggal2 & "'"
    Set RST = OBJ.Execute(SQL)
    j = RST!totalbaris
    
    SQL = "select count(noretur)'totalbaris' from AM_beliretur where tglretur >= '" & tanggal1 & "' and tglretur <= '" & tanggal2 & "'"
    Set RST = OBJ.Execute(SQL)
    k = RST!totalbaris
    '===peneriman
    SQL = "select * from AM_beliapp where tglbeli >= '" & tanggal1 & "' and tglbeli <= '" & tanggal2 & "'"
    Set RST = OBJ.Execute(SQL)
    
    Set objapp = New Excel.Application
    objapp.Visible = False
    Set objbook = objapp.Workbooks.Add
    Set objsheet1 = objapp.Worksheets.Add
    Set objsheet2 = objapp.Worksheets.Add
    Set objsheet3 = objapp.Worksheets.Add
    Set objsheet4 = objapp.Worksheets.Add
    Set objsheet5 = objapp.Worksheets.Add
    Set objsheet6 = objapp.Worksheets.Add
    Set objsheet7 = objapp.Worksheets.Add
    Set objsheet8 = objapp.Worksheets.Add
    Set objsheet9 = objapp.Worksheets.Add
    
    objsheet1.Name = "terima"
    objsheet1.Columns.AutoFit
    
    For i = 0 To RST.Fields.Count - 1 Step 1
        objsheet1.Cells(1, i + 1).Value = RST.Fields(i).Name
    Next i
    
    objsheet1.Range("A2").CopyFromRecordset RST
    Label2 = "  " & j & " records PENERIMAAN affected."
    '===supplier
    SQL = "select kodesupp,namasupp,alamatsupp1,alamatsupp2,telpsupp,faxsupp,contactperson,Category,Wp from AM_supplier"
    Set RST = OBJ.Execute(SQL)
    
    objsheet2.Name = "supplier"
    objsheet2.Columns.AutoFit
    
    For i = 0 To RST.Fields.Count - 1 Step 1
        objsheet2.Cells(1, i + 1).Value = RST.Fields(i).Name
    Next i
    
    objsheet2.Range("A2").CopyFromRecordset RST
    '===retur
    SQL = "select * from AM_beliretur where tglretur >= '" & tanggal1 & "' and tglretur <= '" & tanggal2 & "'"
    Set RST = OBJ.Execute(SQL)
    
    objsheet3.Name = "retur"
    objsheet3.Columns.AutoFit
    
    For i = 0 To RST.Fields.Count - 1 Step 1
        objsheet3.Cells(1, i + 1).Value = RST.Fields(i).Name
    Next i
    
    objsheet3.Range("A2").CopyFromRecordset RST
    Label2 = Label2 + vbCrLf + "  " & k & " records RETUR affected."
    '===retur temporary
    SQL = "select a.noretur,b.nobeli,b.kodebarang,b.qty,b.qtyuse from AM_beliretur a left join am_belilin b on a.nobeli=b.nobeli and a.kodebarang=b.kodebarang where a.tglretur >= '" & tanggal1 & "' and a.tglretur <= '" & tanggal2 & "'"
    Set RST = OBJ.Execute(SQL)
    
    objsheet4.Name = "returtemp"
    objsheet4.Columns.AutoFit
    
    For i = 0 To RST.Fields.Count - 1 Step 1
        objsheet4.Cells(1, i + 1).Value = RST.Fields(i).Name
    Next i
    
    objsheet4.Range("A2").CopyFromRecordset RST
    OBJ.Close
    '===barang
    OBJ1.Open dsn
    SQL1 = "select KodeBarang,NamaBarang,KodeSatuan,KodeProduk,KodeSatuanMutasi from AM_apitemmst"
    Set RST1 = OBJ1.Execute(SQL1)

    objsheet5.Name = "item"
    objsheet5.Columns.AutoFit

    For i = 0 To RST1.Fields.Count - 1 Step 1
        objsheet5.Cells(1, i + 1).Value = RST1.Fields(i).Name
    Next i

    objsheet5.Range("A2").CopyFromRecordset RST1
    '===satuan
    SQL1 = "select KodeSatuan,NamaSatuan,Initial from AM_apunit"
    Set RST1 = OBJ1.Execute(SQL1)
    
    objsheet6.Name = "unit"
    objsheet6.Columns.AutoFit
    
    For i = 0 To RST1.Fields.Count - 1 Step 1
        objsheet6.Cells(1, i + 1).Value = RST1.Fields(i).Name
    Next i
    i = 2
    Do While Not RST1.EOF
        objsheet6.Cells(i, 1).Value = "'" & RST1.Fields(0).Value
        objsheet6.Cells(i, 2).Value = "'" & RST1.Fields(1).Value
        objsheet6.Cells(i, 3).Value = "'" & RST1.Fields(2).Value
        
        RST1.MoveNext
        i = i + 1
    Loop
    '===revisi update/delete penerimaan
    SQL1 = "select * from AM_belirev where flag2='0'"
    Set RST1 = OBJ1.Execute(SQL1)
    
    objsheet7.Name = "revisi"
    objsheet7.Columns.AutoFit
    
    For i = 0 To RST1.Fields.Count - 1 Step 1
        objsheet7.Cells(1, i + 1).Value = RST1.Fields(i).Name
    Next i
    
    objsheet7.Range("A2").CopyFromRecordset RST1
    '===item code
    SQL1 = "select lev,kode,ket from AM_apitemcode"
    Set RST1 = OBJ1.Execute(SQL1)
    
    objsheet8.Name = "rule"
    objsheet8.Columns.AutoFit
    
    For i = 0 To RST1.Fields.Count - 1 Step 1
        objsheet8.Cells(1, i + 1).Value = RST1.Fields(i).Name
    Next i
    i = 2
    Do While Not RST1.EOF
        objsheet8.Cells(i, 1).Value = "'" & RST1.Fields(0).Value
        objsheet8.Cells(i, 2).Value = "'" & RST1.Fields(1).Value
        objsheet8.Cells(i, 3).Value = RST1.Fields(2).Value
        
        RST1.MoveNext
        i = i + 1
    Loop
    '===kode PO
    SQL1 = "select kode1,kode2,kode3 from AM_kode"
    Set RST1 = OBJ1.Execute(SQL1)
    
    objsheet9.Name = "divisi"
    objsheet9.Columns.AutoFit
    
    For i = 0 To RST1.Fields.Count - 1 Step 1
        objsheet9.Cells(1, i + 1).Value = RST1.Fields(i).Name
    Next i
    i = 2
    Do While Not RST1.EOF
        objsheet9.Cells(i, 1).Value = "'" & RST1.Fields(0).Value
        objsheet9.Cells(i, 2).Value = "'" & RST1.Fields(1).Value
        objsheet9.Cells(i, 3).Value = "'" & RST1.Fields(2).Value
        
        RST1.MoveNext
        i = i + 1
    Loop
    OBJ1.Close
    '===
    
    objbook.SaveAs "c:\" & Year(Date) & Month(Date) & Day(Date) & "_" & Hour(Time) & Minute(Time) & "_terima.trs"
    objbook.Close False
    Set objbook = Nothing
    Set objsheet1 = Nothing
    Set objsheet2 = Nothing
    Set objsheet3 = Nothing
    Set objsheet4 = Nothing
    Set objsheet5 = Nothing
    Set objsheet6 = Nothing
    Set objsheet7 = Nothing
    Set objsheet8 = Nothing
    Set objsheet9 = Nothing
    Set objapp = Nothing
    
    Label2 = Label2 + vbCrLf + "   File saved on Local Disk (C:\)"
    
    OBJ.Open dsn
    SQL = "update AM_beliapp set flag1 = '1' where flag1 = '0' and tglbeli >= '" & tanggal1 & "' and tglbeli <= '" & tanggal2 & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "update AM_beliretur set flag1 = '1' where flag1 = '0' and tglretur >= '" & tanggal1 & "' and tglretur <= '" & tanggal2 & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "update AM_belirev set flag2 = '1' where flag2 = '0'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Export Complete.", vbInformation, "Information"
    cmdclear.Enabled = False
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    SQL = "select su  from userlist where username = '" & nmuser & "'"
    OBJ.Execute SQL
 
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='84' and b.kodeuser = '2" & kuser & "'"
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
