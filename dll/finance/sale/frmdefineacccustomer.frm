VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~3.OCX"
Begin VB.Form frmdefineacccustomer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Define Account Customer"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5145
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   5145
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton cmdclose 
      Height          =   480
      Left            =   4050
      TabIndex        =   2
      Top             =   3510
      Width           =   990
      _Version        =   851970
      _ExtentX        =   1746
      _ExtentY        =   847
      _StockProps     =   79
      Caption         =   "Close"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ProgressBar PB1 
      Height          =   270
      Left            =   105
      TabIndex        =   1
      Top             =   2970
      Width           =   4950
      _Version        =   851970
      _ExtentX        =   8731
      _ExtentY        =   476
      _StockProps     =   93
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2730
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   4815
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin XtremeSuiteControls.PushButton cmdverifikasi 
      Height          =   450
      Left            =   90
      TabIndex        =   3
      Top             =   3525
      Width           =   990
      _Version        =   851970
      _ExtentX        =   1746
      _ExtentY        =   794
      _StockProps     =   79
      Caption         =   "Verifikasi"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdcreateacount 
      Height          =   465
      Left            =   1170
      TabIndex        =   4
      Top             =   3525
      Width           =   1005
      _Version        =   851970
      _ExtentX        =   1773
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "Create Account"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label lblstatus 
      BackStyle       =   0  'Transparent
      Height          =   240
      Left            =   120
      TabIndex        =   5
      Top             =   3255
      Width           =   4935
   End
End
Attribute VB_Name = "frmdefineacccustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RS As ADODB.Recordset
Private RST As ADODB.Recordset
Private RST1 As ADODB.Recordset
Private OBJ As New ADODB.Connection
Private OBJ1 As New ADODB.Connection
Private SQL As String
Private SQL1 As String

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdcreateacount_Click()
    'On Error GoTo err_msg
    
    Dim strcust, strtype, strid As String
    Dim int2, jml As Integer
    
    OBJ.Open dsn
    SQL = "select c_type,c_id,ac_cust from am_options"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        strcust = RST!ac_cust
        strtype = RST!c_type
        strid = RST!c_id
        
        OBJ1.Open dsn
        SQL1 = "select top 1 noac from gl_masterac where noac like '" & strcust & "%' order by noac desc"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then strcust = RST1!noac
        strcust = strcust + 1
    End If
    OBJ1.Close
    OBJ.Close
    
    'proses create account baru
    SQL = "am_verfcust"
    OBJ.Open dsn
    
    Set RS = OBJ.Execute(SQL)
    Do While Not RS.EOF
        jml = jml + 1
        RS.MoveNext
        DoEvents
    Loop
    
    PB1.Max = jml
    RS.MoveFirst
    
     OBJ1.Open dsn1
    Do While Not RS.EOF
        'simpan ke table am_autoaccust
        lblstatus = RS!kodecust & " : " & strcust
        
        SQL1 = "insert into am_autoaccust ("
        SQL1 = SQL1 + "kodecomp, "
        SQL1 = SQL1 + "noacc, "
        SQL1 = SQL1 + "kodecust)"
        
        SQL1 = SQL1 + " values('" & strid & "',"
        SQL1 = SQL1 + "'" & strcust & "',"
        SQL1 = SQL1 + "'" & RS!kodecust & "')"
        
        OBJ.Execute SQL1
        
        'simpan ke table gl_masterac di eitdb
        SQL1 = "insert into gl_masterac"
        SQL1 = SQL1 + "(noac"
        SQL1 = SQL1 + ",nmac"
        SQL1 = SQL1 + ",typeac"
        SQL1 = SQL1 + ",jenisac1"
        SQL1 = SQL1 + ",jenisac2"
        SQL1 = SQL1 + ",jenisac3"
        SQL1 = SQL1 + ",jenisac4"
        SQL1 = SQL1 + ",jenisac5"
        SQL1 = SQL1 + ",jenisac6"
        SQL1 = SQL1 + ",jenisac7"
        SQL1 = SQL1 + ",jenisac8"
        SQL1 = SQL1 + ",jenisac9"
        SQL1 = SQL1 + ",jenisac10"
        SQL1 = SQL1 + ",flag"
        SQL1 = SQL1 + ",idupdate"
        SQL1 = SQL1 + ",dateupdate"
        SQL1 = SQL1 + ",identry"
        SQL1 = SQL1 + ",Dateentry)"
            
        SQL1 = SQL1 + "VALUES"
        SQL1 = SQL1 + "('" & strcust & "'"
        SQL1 = SQL1 + ", '" & Mid(RS!namacust + " (" + RS!kodecust + ")", 1, 40) & "'"
        SQL1 = SQL1 + ", 'AS'"
        SQL1 = SQL1 + ", '" & strtype & "'"
        SQL1 = SQL1 + ", ''"
        SQL1 = SQL1 + ", ''"
        SQL1 = SQL1 + ", ''"
        SQL1 = SQL1 + ", ''"
        SQL1 = SQL1 + ", ''"
        SQL1 = SQL1 + ", ''"
        SQL1 = SQL1 + ", ''"
        SQL1 = SQL1 + ", ''"
        SQL1 = SQL1 + ", '" & RS!kodecust & "'"
        SQL1 = SQL1 + ", '0'"
        SQL1 = SQL1 + ", ''"
        SQL1 = SQL1 + ", convert(datetime,'" & tanggalsekarang & "')"
        SQL1 = SQL1 + ", '" & nmuser & "'"
        SQL1 = SQL1 + ", convert(datetime,'" & tanggalsekarang & "'))"
        
        OBJ.Execute SQL1
        
        
        SQL1 = "insert into gl_chacct"
        SQL1 = SQL1 + "(kdcomp"
        SQL1 = SQL1 + ",noac"
        SQL1 = SQL1 + ",typeac"
        SQL1 = SQL1 + ",balancedb"
        SQL1 = SQL1 + ",balancecr"
        SQL1 = SQL1 + ",begindb"
        SQL1 = SQL1 + ",begincr"
        SQL1 = SQL1 + ",periode01"
        SQL1 = SQL1 + ",periode02"
        SQL1 = SQL1 + ",periode03"
        SQL1 = SQL1 + ",periode04"
        SQL1 = SQL1 + ",periode05"
        SQL1 = SQL1 + ",periode06"
        SQL1 = SQL1 + ",periode07"
        SQL1 = SQL1 + ",periode08"
        SQL1 = SQL1 + ",periode09"
        SQL1 = SQL1 + ",periode10"
        SQL1 = SQL1 + ",periode11"
        SQL1 = SQL1 + ",periode12"
        SQL1 = SQL1 + ",periode13"
        SQL1 = SQL1 + ",last01"
        SQL1 = SQL1 + ",last02"
        SQL1 = SQL1 + ",last03"
        SQL1 = SQL1 + ",last04"
        SQL1 = SQL1 + ",last05"
        SQL1 = SQL1 + ",last06"
        SQL1 = SQL1 + ",last07"
        SQL1 = SQL1 + ",last08"
        SQL1 = SQL1 + ",last09"
        SQL1 = SQL1 + ",last10"
        SQL1 = SQL1 + ",last11"
        SQL1 = SQL1 + ",last12"
        SQL1 = SQL1 + ",last13"
        SQL1 = SQL1 + ",temp01"
        SQL1 = SQL1 + ",temp02"
        SQL1 = SQL1 + ",temp03"
        SQL1 = SQL1 + ",temp04"
        SQL1 = SQL1 + ",temp05"
        SQL1 = SQL1 + ",temp06"
        SQL1 = SQL1 + ",temp07"
        SQL1 = SQL1 + ",temp08"
        SQL1 = SQL1 + ",temp09"
        SQL1 = SQL1 + ",temp10"
        SQL1 = SQL1 + ",temp11"
        SQL1 = SQL1 + ",temp12"
        SQL1 = SQL1 + ",temp13"
        SQL1 = SQL1 + ",budget01"
        SQL1 = SQL1 + ",budget02"
        SQL1 = SQL1 + ",budget03"
        SQL1 = SQL1 + ",budget04"
        SQL1 = SQL1 + ",budget05"
        SQL1 = SQL1 + ",budget06"
        SQL1 = SQL1 + ",budget07"
        SQL1 = SQL1 + ",budget08"
        SQL1 = SQL1 + ",budget09"
        SQL1 = SQL1 + ",budget10"
        SQL1 = SQL1 + ",budget11"
        SQL1 = SQL1 + ",budget12"
        SQL1 = SQL1 + ",budget13)"
            
        SQL1 = SQL1 + "VALUES"
        SQL1 = SQL1 + "('" & strid & "'"
        SQL1 = SQL1 + ", '" & strcust & "'"
        SQL1 = SQL1 + ", 'AS'"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0'))"
        
        OBJ.Execute SQL1
        
        'simpan ke table gl_masterac di pusatbeli
       
        SQL1 = "insert into gl_masterac"
        SQL1 = SQL1 + "(noac"
        SQL1 = SQL1 + ",nmac"
        SQL1 = SQL1 + ",typeac"
        SQL1 = SQL1 + ",jenisac1"
        SQL1 = SQL1 + ",jenisac2"
        SQL1 = SQL1 + ",jenisac3"
        SQL1 = SQL1 + ",jenisac4"
        SQL1 = SQL1 + ",jenisac5"
        SQL1 = SQL1 + ",jenisac6"
        SQL1 = SQL1 + ",jenisac7"
        SQL1 = SQL1 + ",jenisac8"
        SQL1 = SQL1 + ",jenisac9"
        SQL1 = SQL1 + ",jenisac10"
        SQL1 = SQL1 + ",flag"
        SQL1 = SQL1 + ",idupdate"
        SQL1 = SQL1 + ",dateupdate"
        SQL1 = SQL1 + ",identry"
        SQL1 = SQL1 + ",Dateentry)"
            
        SQL1 = SQL1 + "VALUES"
        SQL1 = SQL1 + "('" & strcust & "'"
        SQL1 = SQL1 + ", '" & Mid(RS!namacust + " (" + RS!kodecust + ")", 1, 40) & "'"
        SQL1 = SQL1 + ", 'AS'"
        SQL1 = SQL1 + ", '" & strtype & "'"
        SQL1 = SQL1 + ", ''"
        SQL1 = SQL1 + ", ''"
        SQL1 = SQL1 + ", ''"
        SQL1 = SQL1 + ", ''"
        SQL1 = SQL1 + ", ''"
        SQL1 = SQL1 + ", ''"
        SQL1 = SQL1 + ", ''"
        SQL1 = SQL1 + ", ''"
        SQL1 = SQL1 + ", '" & RS!kodecust & "'"
        SQL1 = SQL1 + ", '0'"
        SQL1 = SQL1 + ", ''"
        SQL1 = SQL1 + ", convert(datetime,'" & tanggalsekarang & "')"
        SQL1 = SQL1 + ", '" & nmuser & "'"
        SQL1 = SQL1 + ", convert(datetime,'" & tanggalsekarang & "'))"
        
        OBJ1.Execute SQL1
        
        SQL1 = "insert into gl_chacct"
        SQL1 = SQL1 + "(kdcomp"
        SQL1 = SQL1 + ",noac"
        SQL1 = SQL1 + ",typeac"
        SQL1 = SQL1 + ",balancedb"
        SQL1 = SQL1 + ",balancecr"
        SQL1 = SQL1 + ",begindb"
        SQL1 = SQL1 + ",begincr"
        SQL1 = SQL1 + ",periode01"
        SQL1 = SQL1 + ",periode02"
        SQL1 = SQL1 + ",periode03"
        SQL1 = SQL1 + ",periode04"
        SQL1 = SQL1 + ",periode05"
        SQL1 = SQL1 + ",periode06"
        SQL1 = SQL1 + ",periode07"
        SQL1 = SQL1 + ",periode08"
        SQL1 = SQL1 + ",periode09"
        SQL1 = SQL1 + ",periode10"
        SQL1 = SQL1 + ",periode11"
        SQL1 = SQL1 + ",periode12"
        SQL1 = SQL1 + ",periode13"
        SQL1 = SQL1 + ",last01"
        SQL1 = SQL1 + ",last02"
        SQL1 = SQL1 + ",last03"
        SQL1 = SQL1 + ",last04"
        SQL1 = SQL1 + ",last05"
        SQL1 = SQL1 + ",last06"
        SQL1 = SQL1 + ",last07"
        SQL1 = SQL1 + ",last08"
        SQL1 = SQL1 + ",last09"
        SQL1 = SQL1 + ",last10"
        SQL1 = SQL1 + ",last11"
        SQL1 = SQL1 + ",last12"
        SQL1 = SQL1 + ",last13"
        SQL1 = SQL1 + ",temp01"
        SQL1 = SQL1 + ",temp02"
        SQL1 = SQL1 + ",temp03"
        SQL1 = SQL1 + ",temp04"
        SQL1 = SQL1 + ",temp05"
        SQL1 = SQL1 + ",temp06"
        SQL1 = SQL1 + ",temp07"
        SQL1 = SQL1 + ",temp08"
        SQL1 = SQL1 + ",temp09"
        SQL1 = SQL1 + ",temp10"
        SQL1 = SQL1 + ",temp11"
        SQL1 = SQL1 + ",temp12"
        SQL1 = SQL1 + ",temp13"
        SQL1 = SQL1 + ",budget01"
        SQL1 = SQL1 + ",budget02"
        SQL1 = SQL1 + ",budget03"
        SQL1 = SQL1 + ",budget04"
        SQL1 = SQL1 + ",budget05"
        SQL1 = SQL1 + ",budget06"
        SQL1 = SQL1 + ",budget07"
        SQL1 = SQL1 + ",budget08"
        SQL1 = SQL1 + ",budget09"
        SQL1 = SQL1 + ",budget10"
        SQL1 = SQL1 + ",budget11"
        SQL1 = SQL1 + ",budget12"
        SQL1 = SQL1 + ",budget13)"
            
        SQL1 = SQL1 + "VALUES"
        SQL1 = SQL1 + "('" & strid & "'"
        SQL1 = SQL1 + ", '" & strcust & "'"
        SQL1 = SQL1 + ", 'AS'"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0')"
        SQL1 = SQL1 + ", convert(money,'0'))"
        
        OBJ1.Execute SQL1
        
        strcust = strcust + 1
        
        PB1.Value = PB1.Value + 1
        RS.MoveNext
        DoEvents
    Loop
    OBJ1.Close
    OBJ.Close
    MsgBox "Process Completed...!", vbInformation, AppName
    lblstatus = ""
    Exit Sub
err_msg:
    MsgBox Err.Description
    
End Sub

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function


Private Sub cmdverifikasi_Click()
    On Error GoTo err_msg
    
    Dim i As Integer
    Dim li As ListItem
    
    SQL = "am_verfcust"
    OBJ.Open dsn
    Set RS = OBJ.Execute(SQL)
       
    If RS.EOF Then
        MsgBox "No Customer for define customer account"
        OBJ.Close
        Exit Sub
    End If
       
    i = 0
    Do While Not RS.EOF
        i = i + 1
        RS.MoveNext
    Loop
    
    PB1.Max = i
    
    ListView1.ColumnHeaders.Add , , "No", 500
    ListView1.ColumnHeaders.Add , , "Kode Cust", 1000
    ListView1.ColumnHeaders.Add , , "Nama Cust", 3000
    ListView1.View = lvwReport
    ListView1.BorderStyle = ccFixedSingle
    
    ListView1.ListItems.Clear
    i = 1
    RS.MoveFirst
    Do While Not RS.EOF
        Set li = ListView1.ListItems.Add(, , i)
        li.SubItems(1) = RS!kodecust
        li.SubItems(2) = RS!namacust
        i = i + 1
        PB1.Value = PB1.Value + 1
        RS.MoveNext
    Loop
    PB1.Value = 0
    OBJ.Close
    Exit Sub
err_msg:
    MsgBox "Fatal error : " & Err.Description, vbCritical, AppName
End Sub
