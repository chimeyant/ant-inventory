VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmsjalanprint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Surat Jalan"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5370
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmsjalanprint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport crystal 
      Left            =   0
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   4800
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
      MICON           =   "frmsjalanprint.frx":2372
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdview 
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   4800
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Print"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmsjalanprint.frx":268C
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
      Format          =   141164547
      CurrentDate     =   37464
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   840
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
      Format          =   141164547
      CurrentDate     =   37464
   End
   Begin Chameleon.chameleonButton cmddown 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "66666"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Marlett"
         Size            =   9
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmsjalanprint.frx":29A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSForms.ComboBox cmbstatus 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3201;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      DropButtonStyle =   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
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
      TabIndex        =   8
      Top             =   510
      Width           =   1215
   End
   Begin VB.Label Label4 
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
      Top             =   870
      Width           =   1215
   End
   Begin MSForms.ListBox ListBox 
      Height          =   3135
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   5175
      ScrollBars      =   3
      DisplayStyle    =   2
      Size            =   "9128;5530"
      ColumnCount     =   4
      MultiSelect     =   1
      SpecialEffect   =   3
      FontName        =   "Tahoma"
      FontHeight      =   135
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      Caption         =   "Status"
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
      TabIndex        =   9
      Top             =   150
      Width           =   1215
   End
End
Attribute VB_Name = "frmsjalanprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim i As Integer
Dim boo1 As Boolean

Private Sub cmbstatus_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub cmbstatus_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmddown_Click()
    ListBox.Clear
    ListBox.ColumnCount = 4
    ListBox.ColumnWidths = "2 cm; 1 cm; 4 cm; 1 cm"
    
    If date1 > Date Then Exit Sub
    If cmbstatus = "" Then Exit Sub
        
    i = 0
    OBJ.Open dsn
    If cmbstatus = "Belum Print" Then SQL = "select a.nosj,a.KodeGudang,b.namacust,a.noso from am_sjhdr a left join am_customer b on a.kodecust=b.kodecust where a.via2='0' and a.tglsj >= '" & batas1 & "' and a.tglsj <= '" & batas2 & "' order by a.nosj"
    If cmbstatus = "Sudah Print" Then SQL = "select a.nosj,a.KodeGudang,b.namacust,a.noso from am_sjhdr a left join am_customer b on a.kodecust=b.kodecust where a.via2='1' and a.tglsj >= '" & batas1 & "' and a.tglsj <= '" & batas2 & "' order by a.nosj"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        ListBox.AddItem RST!nosj
        ListBox.List(i, 1) = RST!kodegudang
        ListBox.List(i, 2) = RST!namacust
        ListBox.List(i, 3) = RST!noso
        i = i + 1
        RST.MoveNext
    Loop
    OBJ.Close
End Sub

Private Sub cmdview_Click()
    
    If ListBox.ListCount = 0 Then Exit Sub
    
    OBJ.Open dsn
    SQL = "DELETE FROM am_sjtemp"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
                    
    For i = 0 To ListBox.ListCount - 1
        If ListBox.Selected(i) = True Then
            ListBox.Selected(i) = False
            
            OBJ.Open dsn
            SQL = "INSERT INTO am_sjtemp (nosj,tglsj,kodecust,nopo,tglkirim,via,"
            SQL = SQL + "kodebarang,qty,lineitem,keterangan,bn,namacust,alamatcust,"
            SQL = SQL + "kota,area,satuan,namabarang,sales,totalrec,hal,noso,tglso)"
            
            SQL = SQL + " SELECT a.nosj,a.tglsj,a.kodecust,a.nopo,a.tglkirim,a.via,"
            SQL = SQL + "b.kodebarang,b.qty,b.lineitem,b.keterangan,b.bn,e.namacust,e.alamatcust,"
            SQL = SQL + "e.kota,(h.nama)'area',(g.init)'satuan',c.namabarang,f.namasales,"
            SQL = SQL + "isnull((SELECT count(z.nosj)FROM am_sjlin z WHERE z.nosj=a.nosj),0)'totalrec',"
            SQL = SQL + "(case when b.lineitem >=1 and b.lineitem<=10 then 1 else 2 end)'hal',a.noso,i.tglso"
            SQL = SQL + " FROM am_sjlin b left join am_sjhdr a"
            SQL = SQL + " ON a.nosj=b.nosj left join am_customer e"
            SQL = SQL + " ON a.kodecust=e.kodecust left join am_area h"
            SQL = SQL + " ON e.kodearea=h.kode left join am_salesman f"
            SQL = SQL + " ON a.kodesales=f.kodesales left join am_itemdtl c"
            SQL = SQL + " ON b.kodebarang=c.kodebarang and b.kodesatuan=c.kodesatuan left join am_unit g"
            SQL = SQL + " ON b.kodesatuan=g.kodesatuan left join am_soapp i"
            SQL = SQL + " ON a.noso=i.noso and b.kodebarang=i.kodebarang and b.kodesatuan=i.kodesatuan WHERE a.nosj = '" & ListBox.List(i, 0) & "'"
            Set RST = OBJ.Execute(SQL)

            SQL = "update am_sjhdr set via2='1' where nosj = '" & ListBox.List(i, 0) & "'"
            Set RST = OBJ.Execute(SQL)
            OBJ.Close
        End If
    Next i
    
    crystal.Reset
    'sementara di tampilkan
    '----------------------
    'crystal.WindowState = crptMaximized
    'crystal.WindowShowCloseBtn = True
    'crystal.WindowShowPrintSetupBtn = True
    'crystal.WindowShowSearchBtn = True
    '----------------------
    crystal.Destination = crptToPrinter    'langsung print
    crystal.Connect = dsnreport
    crystal.DataFiles(0) = "Proc(am_sjalan)"
    crystal.ReportFileName = AppPath & "\reports\sale\inv\sjalan.rpt"
    crystal.ParameterFields(0) = "@namauser ;" + nmuser + ";true"
    crystal.RetrieveDataFiles
    crystal.Action = 1
    
    cmddown_Click
End Sub

Private Sub Form_Activate()
   ' If kuser <> "q" Then
   '     OBJ.Open dsn
   '     SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='165' and b.kodeuser = '1" & kuser & "'"
   '     Set RST = OBJ.Execute(SQL)
   '     If RST.EOF Then
   '         MsgBox "User Rights Denied !!" & vbCrLf & _
   '         "Please contact your Administrator.", vbCritical, "User Rights"
   '         OBJ.Close
    '        Unload Me
    '        Exit Sub
    '    End If
     '   OBJ.Close
    'End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    date1 = Date
    date2 = Date
    
    cmbstatus.Clear
    
    cmbstatus.AddItem "Sudah Print"
    cmbstatus.AddItem "Belum Print"
    
    OBJ.Open dsn
    SQL = "select * FROM am_options where para5 = '1'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        cmddown.Enabled = False
        cmdview.Enabled = False
        boo1 = False
    Else
        SQL = "update am_options set para5 = '1'"
        Set RST = OBJ.Execute(SQL)
        boo1 = True
    End If
    OBJ.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If boo1 Then
        OBJ.Open dsn
        SQL = "DELETE FROM am_sjtemp"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "update am_options set para5 = '0'"
        Set RST = OBJ.Execute(SQL)
        OBJ.Close
    End If
End Sub

Function batas1()
    batas1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function batas2()
    batas2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function
