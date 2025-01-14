VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmsoh 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Quantitiy On Hand + Stock Awal"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
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
   ScaleHeight     =   4575
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8070
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Quantity On Hand"
      TabPicture(0)   =   "frmsoh.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdrefresh"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "grid"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtsearch"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Stock Awal"
      TabPicture(1)   =   "frmsoh.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(2)=   "Label7"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "."
      TabPicture(2)   =   "frmsoh.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   ".."
      TabPicture(3)   =   "frmsoh.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "..."
      TabPicture(4)   =   "frmsoh.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      Begin VB.Frame Frame2 
         Caption         =   "Saldo Awal Bahan Baku"
         Height          =   2415
         Left            =   -74880
         TabIndex        =   18
         Top             =   120
         Width           =   6735
         Begin VB.TextBox txtdesc 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   3240
            MaxLength       =   40
            TabIndex        =   19
            Top             =   360
            Width           =   3255
         End
         Begin VB.TextBox txtkode 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   4
            Top             =   360
            Width           =   1335
         End
         Begin Chameleon.chameleonButton cmdsearch0 
            Height          =   285
            Left            =   240
            TabIndex        =   20
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            BTYPE           =   9
            TX              =   "Bahan Baku"
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
            MICON           =   "frmsoh.frx":008C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin TDBNumber6Ctl.TDBNumber txtnilai 
            Height          =   285
            Left            =   1800
            TabIndex        =   5
            Top             =   1440
            Width           =   1575
            _Version        =   65536
            _ExtentX        =   2778
            _ExtentY        =   503
            Calculator      =   "frmsoh.frx":03A6
            Caption         =   "frmsoh.frx":03C6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmsoh.frx":0432
            Keys            =   "frmsoh.frx":0450
            Spin            =   "frmsoh.frx":0492
            AlignHorizontal =   1
            AlignVertical   =   0
            Appearance      =   0
            BackColor       =   -2147483643
            BorderStyle     =   1
            BtnPositioning  =   0
            ClipMode        =   0
            ClearAction     =   0
            DecimalPoint    =   "."
            DisplayFormat   =   "###,###,##0.00;;0"
            EditMode        =   1
            Enabled         =   -1
            ErrorBeep       =   0
            ForeColor       =   -2147483640
            Format          =   "###,###,##0.00"
            HighlightText   =   1
            MarginBottom    =   1
            MarginLeft      =   1
            MarginRight     =   1
            MarginTop       =   1
            MaxValue        =   999999999999999
            MinValue        =   0
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   ","
            ShowContextMenu =   0
            ValueVT         =   1245189
            Value           =   0
            MaxValueVT      =   5
            MinValueVT      =   5
         End
         Begin Chameleon.chameleonButton cmdsave 
            Height          =   420
            Left            =   4440
            TabIndex        =   6
            Top             =   1800
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   741
            BTYPE           =   4
            TX              =   "Save"
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
            MICON           =   "frmsoh.frx":04BA
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
            Height          =   420
            Left            =   5520
            TabIndex        =   7
            Top             =   1800
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   741
            BTYPE           =   4
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
            MICON           =   "frmsoh.frx":07D4
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   3
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label6 
            Caption         =   "Sub Divisi"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   1110
            Width           =   1095
         End
         Begin VB.Label lblunitcode 
            Caption         =   "Satuan Bahan Baku"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lblsatuan 
            BackColor       =   &H80000005&
            Height          =   255
            Left            =   1800
            TabIndex        =   23
            Top             =   720
            Width           =   4695
         End
         Begin VB.Label Label28 
            Caption         =   "Quantity Awal"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   1470
            Width           =   1455
         End
         Begin VB.Label lblsub 
            BackColor       =   &H80000014&
            Height          =   255
            Left            =   1800
            TabIndex        =   21
            Top             =   1080
            Width           =   4695
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tanggal Saldo Awal"
         Height          =   2415
         Left            =   -68040
         TabIndex        =   16
         Top             =   120
         Width           =   2295
         Begin Chameleon.chameleonButton cmdset 
            Height          =   420
            Left            =   120
            TabIndex        =   9
            Top             =   1800
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   741
            BTYPE           =   4
            TX              =   "Set Date"
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
            MICON           =   "frmsoh.frx":0AEE
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   3
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Chameleon.chameleonButton cmdlock 
            Height          =   420
            Left            =   1200
            TabIndex        =   10
            Top             =   1800
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   741
            BTYPE           =   4
            TX              =   "Lock"
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
            MICON           =   "frmsoh.frx":0E08
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
            Left            =   120
            TabIndex        =   8
            Top             =   600
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
            Format          =   107216899
            CurrentDate     =   37426
         End
         Begin VB.Label Label13 
            Caption         =   "Tanggal"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.TextBox txtsearch 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   0
         Top             =   120
         Width           =   2295
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
         Height          =   3135
         Left            =   0
         TabIndex        =   1
         Top             =   480
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   5530
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         GridLines       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   9
      End
      Begin Chameleon.chameleonButton cmdrefresh 
         Height          =   420
         Left            =   8160
         TabIndex        =   2
         Top             =   3720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   741
         BTYPE           =   4
         TX              =   "Refresh"
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
         MICON           =   "frmsoh.frx":1122
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "press esc to exit                              press esc to exit"
         Height          =   255
         Left            =   -74760
         TabIndex        =   17
         Top             =   2520
         Width           =   8895
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "press esc to exit"
         Height          =   255
         Left            =   4800
         TabIndex        =   15
         Top             =   3900
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Find"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   150
         Width           =   495
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3720
         Width           =   6615
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6960
         TabIndex        =   12
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7560
         TabIndex        =   11
         Top             =   3720
         Visible         =   0   'False
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmsoh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private m_SortColumn As Integer
Private m_SortAscending As Integer

Private Sub cmdclear_Click()
    txtkode = ""
    txtdesc = ""
    lblsatuan = ""
    lblsub = ""
    txtnilai = 0
    txtkode.SetFocus
End Sub

Private Sub cmdlock_Click()
    If MsgBox("Lock this beginning date ?", vbYesNo + vbQuestion, "Question") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "update am_invloc set lock = '1'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Date lock, click ok to continue ...", vbInformation, "Information"
End Sub

Private Sub cmdrefresh_Click()
    m_SortColumn = -1
    m_SortAscending = -1
    
    showtabel
    txtsearch = ""
End Sub

Private Sub cmdsave_Click()
    If Len(Trim(txtkode)) = 0 Then
        MsgBox "Data entry not Complete.", vbExclamation, "Warning"
        txtkode.SetFocus
        Exit Sub
    End If
    
    If txtkode = "" Or txtdesc = "" Or txtnilai = 0 Then
       MsgBox "Data entry not Complete.", vbExclamation, "Warning"
       Exit Sub
    End If
    
    txtkode = Trim(txtkode)
    
    OBJ.Open dsn
    SQL = "select * from am_invloc where lock = '1'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Record lock.", vbExclamation, "Warning"
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "select * from am_invloc where kodebarang = '" & txtkode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If Year(RST!tglupdate) = 1900 Then
            SQL = "delete from am_invloc where kodebarang = '" & txtkode & "'"
            Set RST = OBJ.Execute(SQL)
        Else
            MsgBox "Data already close.", vbInformation, "Information"
            OBJ.Close
            cmdclear_Click
            Exit Sub
        End If
    End If
    
    SQL = "INSERT INTO am_invloc"
    SQL = SQL + "(KodeBarang"
    SQL = SQL + ",lock"
    SQL = SQL + ",qtyawal"
    SQL = SQL + ",tglupdate)"
    
    SQL = SQL + "VALUES"
    SQL = SQL + " ('" & txtkode & "'"
    SQL = SQL + ", '0'"
    SQL = SQL + ", convert(money,'" & txtnilai & "')"
    SQL = SQL + ", convert(datetime,'01/01/1900'))"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
        
    MsgBox "Data saved, click ok to continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdsearch0_Click()
    carisql1 = "select kodebarang, namabarang from am_apitemmst"
    namatabel = "Bahan Baku"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch0_GotFocus()
    If hasil = "" Then Exit Sub
    txtkode = hasil
    caritem
    hasil = ""
    hasil1 = ""
    hasil2 = ""
    txtdesc.SetFocus
End Sub

Private Sub caritem()
    If txtkode = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from am_apitemmst where kodebarang = '" & txtkode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtdesc = RST!namabarang
        lblsatuan = RST!kodesatuanmutasi
        lblsub = RST!kodeproduk
        
        SQL = "select * from am_apunit where kodesatuan = '" & lblsatuan & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then lblsatuan = RST!kodesatuan & " - " & RST!namasatuan
        
        OBJ.Close
        Exit Sub
    Else
        txtdesc = ""
        lblsatuan = ""
        lblsub = ""
    End If
    OBJ.Close
End Sub

Private Sub cmdset_Click()
    OBJ.Open dsn
    SQL = "select * from am_invloc where lock = '1'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Record lock.", vbExclamation, "Warning"
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    If MsgBox("Accept this beginning date ?", vbYesNo + vbQuestion, "Question") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "update am_invloc set tglupdate = convert(datetime,'" & Month(date1) & "/" & Day(date1) & "/" & Year(date1) & "')"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Date set, click ok to continue ...", vbInformation, "Information"
End Sub

Private Sub Form_Activate()
   ' If kuser <> "q" Then
   '     OBJ.Open dsn
   '     SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='24' and b.kodeuser = '2" & kuser & "'"
   '     Set RST = OBJ.Execute(SQL)
   '     If RST.EOF Then SSTab1.TabEnabled(1) = False
   '     OBJ.Close
   ' End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
  
    cmdrefresh_Click
End Sub

Private Sub grid_Click()
    If grid.TextMatrix(0, 0) = "" Then Exit Sub
                
    Label3 = grid.MouseCol
    If grid.MouseRow > 0 Then Exit Sub
    
    If grid.MouseCol <> m_SortColumn Then
        If m_SortColumn >= 0 Then
            grid.TextMatrix(0, m_SortColumn) = _
                Mid$(grid.TextMatrix(0, m_SortColumn), 3)
        End If
        m_SortColumn = grid.MouseCol
        
        m_SortAscending = True
        grid.TextMatrix(0, m_SortColumn) = _
            "> " & grid.TextMatrix(0, m_SortColumn)
    Else
        m_SortAscending = Not m_SortAscending
        
        If m_SortAscending Then
            grid.TextMatrix(0, m_SortColumn) = _
                "> " & Mid$(grid.TextMatrix(0, m_SortColumn), 3)
        Else
            grid.TextMatrix(0, m_SortColumn) = _
                "< " & Mid$(grid.TextMatrix(0, m_SortColumn), 3)
        End If
    End If
    
    Label4 = Mid$(grid.TextMatrix(0, Label3), 3)
    grid.Row = 1
    grid.RowSel = grid.Rows - 1
    grid.Col = m_SortColumn
    txtsearch = ""
    If m_SortAscending Then
        grid.Sort = flexSortStringAscending
    Else
        grid.Sort = flexSortStringDescending
    End If
    
    If txtsearch.Visible = True Then txtsearch.SetFocus
End Sub

Private Sub txtdesc_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtdesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtnilai.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtkode_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtKode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtdesc.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtsearch_Change()
    OBJ.Open dsn
    If Label4 = "namabarang" Then
        SQL = "select a.namabarang,b.namasatuan,"
        SQL = SQL + "(isnull((select sum(k.qtyawal) from am_invloc k where a.kodebarang=k.kodebarang),0)+"
        SQL = SQL + "isnull((select sum(c.qtyuse) from am_belilin c where a.kodebarang=c.kodebarang),0))'beli',"
        SQL = SQL + "isnull((select sum(d.qty) from am_uselin d where a.kodebarang=d.kodebarang),0)'pakai',"
        SQL = SQL + "isnull((select sum(e.qty) from am_usesisa e where a.kodebarang=e.kodebarang),0)'sisapakai',"
        SQL = SQL + "isnull((select sum(f.qty) from am_beliretur f where a.kodebarang=f.kodebarang),0)'returbeli',"
        SQL = SQL + "isnull((select sum(m.qty) from am_mutlin m where m.type='01' and a.kodebarang=m.kodebarang),0)'mutasiin',"
        SQL = SQL + "isnull((select sum(n.qty) from am_mutlin n where (n.type='02' or n.type='03') and a.kodebarang=n.kodebarang),0)'mutasiout',"
        SQL = SQL + "(isnull((select sum(l.qtyawal) from am_invloc l where a.kodebarang=l.kodebarang),0)+"
        SQL = SQL + "isnull((select sum(g.qtyuse) from am_belilin g where a.kodebarang=g.kodebarang),0)-"
        SQL = SQL + "isnull((select sum(h.qty) from am_uselin h where a.kodebarang=h.kodebarang),0)+"
        SQL = SQL + "isnull((select sum(i.qty) from am_usesisa i where a.kodebarang=i.kodebarang),0)-"
        SQL = SQL + "isnull((select sum(j.qty) from am_beliretur j where a.kodebarang=j.kodebarang),0)+"
        SQL = SQL + "isnull((select sum(o.qty) from am_mutlin o where o.type='01' and a.kodebarang=o.kodebarang),0)-"
        SQL = SQL + "isnull((select sum(p.qty) from am_mutlin p where (p.type='02' or p.type='03') and a.kodebarang=p.kodebarang),0))'onhand'"
        SQL = SQL + "from am_apitemmst a left join am_apunit b on a.kodesatuanmutasi=b.kodesatuan"
        SQL = SQL + " where a.namabarang like '" + txtsearch + "%'"
    ElseIf Label4 = "namasatuan" Then
        SQL = "select a.namabarang,b.namasatuan,"
        SQL = SQL + "(isnull((select sum(k.qtyawal) from am_invloc k where a.kodebarang=k.kodebarang),0)+"
        SQL = SQL + "isnull((select sum(c.qtyuse) from am_belilin c where a.kodebarang=c.kodebarang),0))'beli',"
        SQL = SQL + "isnull((select sum(d.qty) from am_uselin d where a.kodebarang=d.kodebarang),0)'pakai',"
        SQL = SQL + "isnull((select sum(e.qty) from am_usesisa e where a.kodebarang=e.kodebarang),0)'sisapakai',"
        SQL = SQL + "isnull((select sum(f.qty) from am_beliretur f where a.kodebarang=f.kodebarang),0)'returbeli',"
        SQL = SQL + "isnull((select sum(m.qty) from am_mutlin m where m.type='01' and a.kodebarang=m.kodebarang),0)'mutasiin',"
        SQL = SQL + "isnull((select sum(n.qty) from am_mutlin n where (n.type='02' or n.type='03') and a.kodebarang=n.kodebarang),0)'mutasiout',"
        SQL = SQL + "(isnull((select sum(l.qtyawal) from am_invloc l where a.kodebarang=l.kodebarang),0)+"
        SQL = SQL + "isnull((select sum(g.qtyuse) from am_belilin g where a.kodebarang=g.kodebarang),0)-"
        SQL = SQL + "isnull((select sum(h.qty) from am_uselin h where a.kodebarang=h.kodebarang),0)+"
        SQL = SQL + "isnull((select sum(i.qty) from am_usesisa i where a.kodebarang=i.kodebarang),0)-"
        SQL = SQL + "isnull((select sum(j.qty) from am_beliretur j where a.kodebarang=j.kodebarang),0)+"
        SQL = SQL + "isnull((select sum(o.qty) from am_mutlin o where o.type='01' and a.kodebarang=o.kodebarang),0)-"
        SQL = SQL + "isnull((select sum(p.qty) from am_mutlin p where (p.type='02' or p.type='03') and a.kodebarang=p.kodebarang),0))'onhand'"
        SQL = SQL + "from am_apitemmst a left join am_apunit b on a.kodesatuanmutasi=b.kodesatuan"
        SQL = SQL + " where b.namasatuan like '" + txtsearch + "%'"
    End If
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        grid.Clear
        grid.Rows = 2
        OBJ.Close
        Label2 = ""
        
        grid.ColWidth(0) = 3000
        grid.ColWidth(1) = 1000
        grid.ColWidth(2) = 700
        grid.ColWidth(3) = 700
        grid.ColWidth(4) = 700
        grid.ColWidth(5) = 700
        grid.ColWidth(6) = 700
        grid.ColWidth(7) = 700
        grid.ColWidth(8) = 700
        
        Exit Sub
    End If
    Set RST = OBJ.Execute(SQL)
    Set grid.DataSource = RST
    Label2 = grid.Rows - 1 & " Items"
    OBJ.Close
    grid.TextMatrix(0, Label3) = _
            "> " & grid.TextMatrix(0, Label3)
    grid.Sort = flexSortStringAscending
    
    grid.ColWidth(0) = 3000
    grid.ColWidth(1) = 1000
    grid.ColWidth(2) = 700
    grid.ColWidth(3) = 700
    grid.ColWidth(4) = 700
    grid.ColWidth(5) = 700
    grid.ColWidth(6) = 700
    grid.ColWidth(7) = 700
    grid.ColWidth(8) = 700
End Sub

Private Sub txtsearch_KeyPress(KeyAscii As Integer)
    If Label4 = "beli" Or Label4 = "pakai" Or Label4 = "sisapakai" Or Label4 = "returbeli" Or Label4 = "onhand" Or Label4 = "mutasiin" Or Label4 = "mutasiout" Then KeyAscii = 0
End Sub

Private Sub showtabel()
    OBJ.Open dsn
    SQL = "select a.namabarang,b.namasatuan,"
    SQL = SQL + "(isnull((select sum(k.qtyawal) from am_invloc k where a.kodebarang=k.kodebarang),0)+"
    SQL = SQL + "isnull((select sum(c.qtyuse) from am_belilin c where a.kodebarang=c.kodebarang),0))'beli',"
    SQL = SQL + "isnull((select sum(d.qty) from am_uselin d where a.kodebarang=d.kodebarang),0)'pakai',"
    SQL = SQL + "isnull((select sum(e.qty) from am_usesisa e where a.kodebarang=e.kodebarang),0)'sisapakai',"
    SQL = SQL + "isnull((select sum(f.qty) from am_beliretur f where a.kodebarang=f.kodebarang),0)'returbeli',"
    SQL = SQL + "isnull((select sum(m.qty) from am_mutlin m where m.type='01' and a.kodebarang=m.kodebarang),0)'mutasiin',"
    SQL = SQL + "isnull((select sum(n.qty) from am_mutlin n where (n.type='02' or n.type='03') and a.kodebarang=n.kodebarang),0)'mutasiout',"
    SQL = SQL + "(isnull((select sum(l.qtyawal) from am_invloc l where a.kodebarang=l.kodebarang),0)+"
    SQL = SQL + "isnull((select sum(g.qtyuse) from am_belilin g where a.kodebarang=g.kodebarang),0)-"
    SQL = SQL + "isnull((select sum(h.qty) from am_uselin h where a.kodebarang=h.kodebarang),0)+"
    SQL = SQL + "isnull((select sum(i.qty) from am_usesisa i where a.kodebarang=i.kodebarang),0)-"
    SQL = SQL + "isnull((select sum(j.qty) from am_beliretur j where a.kodebarang=j.kodebarang),0)+"
    SQL = SQL + "isnull((select sum(o.qty) from am_mutlin o where o.type='01' and a.kodebarang=o.kodebarang),0)-"
    SQL = SQL + "isnull((select sum(p.qty) from am_mutlin p where (p.type='02' or p.type='03') and a.kodebarang=p.kodebarang),0))'onhand'"
    SQL = SQL + "from am_apitemmst a left join am_apunit b on a.kodesatuanmutasi=b.kodesatuan"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        grid.Clear
        grid.Rows = 2
        OBJ.Close
        
        grid.ColWidth(0) = 3000
        grid.ColWidth(1) = 1000
        grid.ColWidth(2) = 700
        grid.ColWidth(3) = 700
        grid.ColWidth(4) = 700
        grid.ColWidth(5) = 700
        grid.ColWidth(6) = 700
        grid.ColWidth(7) = 700
        grid.ColWidth(8) = 700
        
        Exit Sub
    End If
    OBJ.Close
    
    OBJ.Open dsn
    Set RST = OBJ.Execute(SQL)
    Set grid.DataSource = RST
    OBJ.Close
    
    Label2 = grid.Rows - 1 & " Items"
    grid.TextMatrix(0, 0) = "> " & grid.TextMatrix(0, 0)
    Label4 = Mid$(grid.TextMatrix(0, 0), 3)
    m_SortColumn = 0
    Label3 = 0
    grid.Col = 0
    grid.Sort = flexSortStringAscending
    
    grid.ColWidth(0) = 3000
    grid.ColWidth(1) = 1000
    grid.ColWidth(2) = 700
    grid.ColWidth(3) = 700
    grid.ColWidth(4) = 700
    grid.ColWidth(5) = 700
    grid.ColWidth(6) = 700
    grid.ColWidth(7) = 700
    grid.ColWidth(8) = 700
End Sub
