VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmsjalanshowagain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Surat Jalan"
   ClientHeight    =   615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2775
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmsjalanshowagain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtnodo1 
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
      Left            =   960
      MaxLength       =   15
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin Crystal.CrystalReport crystal 
      Left            =   120
      Top             =   1440
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
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Skip"
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
      MICON           =   "frmsjalanshowagain.frx":2372
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
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Preview"
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
      MICON           =   "frmsjalanshowagain.frx":268C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdprint 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
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
      MICON           =   "frmsjalanshowagain.frx":29A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmsjalanshowagain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim str1 As String
Dim boo1 As Boolean

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdprint_Click()
    OBJ.Open dsn
    SQL = "DELETE FROM am_sjtemp"
    Set RST = OBJ.Execute(SQL)
    
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
    SQL = SQL + " ON a.noso=i.noso and b.kodebarang=i.kodebarang WHERE a.nosj = '" & setup1 & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    crystal.Reset
    crystal.Destination = crptToPrinter
    crystal.Connect = dsnreport
    crystal.DataFiles(0) = "Proc(am_sjalan)"
    crystal.ReportFileName = AppPath & "\reports\gl\inv\sjalan.rpt"
    crystal.ParameterFields(0) = "@namauser ;" + nmuser + ";true"
    crystal.RetrieveDataFiles
    crystal.Action = 1
    
    cmdclose.Caption = "Continue"
    cmdview.Enabled = False
    cmdprint.Enabled = False
End Sub

Private Sub cmdview_Click()
    OBJ.Open dsn
    SQL = "DELETE FROM am_sjtemp"
    Set RST = OBJ.Execute(SQL)
        
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
    SQL = SQL + " ON a.noso=i.noso and b.kodebarang=i.kodebarang WHERE a.nosj = '" & setup1 & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    crystal.Reset
    crystal.WindowState = crptMaximized
    crystal.WindowShowCloseBtn = True
    crystal.WindowShowPrintSetupBtn = False
    crystal.WindowShowSearchBtn = True
    crystal.Connect = dsnreport
    crystal.DataFiles(0) = "Proc(am_sjalan)"
    crystal.ReportFileName = AppPath & "\reports\sale\inv\sjalan.rpt"
    crystal.ParameterFields(0) = "@namauser ;" + nmuser + ";true"
    crystal.RetrieveDataFiles
    crystal.Action = 1
    
    cmdclose.Caption = "Continue"
End Sub

Private Sub Form_Load()
    OBJ.Open dsn
    SQL = "select * FROM am_options where para5 = '1'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        cmdprint.Enabled = False
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
