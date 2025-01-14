VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmcekgiro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Maintain Giro"
   ClientHeight    =   4020
   ClientLeft      =   3615
   ClientTop       =   3105
   ClientWidth     =   8895
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
   Icon            =   "frmcekgiro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "manual"
      Height          =   3855
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Visible         =   0   'False
      Width           =   8655
      Begin MSComCtl2.DTPicker date3 
         Height          =   285
         Left            =   240
         TabIndex        =   30
         Top             =   360
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
         Format          =   88997891
         CurrentDate     =   37421
      End
      Begin Chameleon.chameleonButton cmdpro1 
         Height          =   375
         Left            =   720
         TabIndex        =   31
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "proses Tolak"
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
         MICON           =   "frmcekgiro.frx":2372
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdpro2 
         Height          =   375
         Left            =   720
         TabIndex        =   32
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "proses Cair"
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
         MICON           =   "frmcekgiro.frx":268C
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
   Begin VB.CheckBox Check1 
      Caption         =   "Sudah Cair/Tolak"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   2880
      TabIndex        =   28
      Top             =   2550
      Width           =   5775
   End
   Begin VB.TextBox txtcari 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      MaxLength       =   15
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   3180
      ItemData        =   "frmcekgiro.frx":29A6
      Left            =   0
      List            =   "frmcekgiro.frx":29A8
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   2655
   End
   Begin TDBNumber6Ctl.TDBNumber nilaikurs 
      Height          =   255
      Left            =   7200
      TabIndex        =   21
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   450
      Calculator      =   "frmcekgiro.frx":29AA
      Caption         =   "frmcekgiro.frx":29CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcekgiro.frx":2A2F
      Keys            =   "frmcekgiro.frx":2A4D
      Spin            =   "frmcekgiro.frx":2A97
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "####0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999999999
      MinValue        =   -999999999999999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   ","
      ShowContextMenu =   -1
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   6840
      TabIndex        =   10
      Top             =   120
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
      Format          =   88997891
      CurrentDate     =   37421
   End
   Begin VB.TextBox txtapply 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4200
      MaxLength       =   15
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   7800
      TabIndex        =   9
      Top             =   3570
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
      MICON           =   "frmcekgiro.frx":2ABF
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
      Left            =   6840
      TabIndex        =   8
      Top             =   3570
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
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
      MICON           =   "frmcekgiro.frx":2DD9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdadd 
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   3570
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Save Cair/Tolak"
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
      MICON           =   "frmcekgiro.frx":30F3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtneto 
      Height          =   255
      Left            =   4200
      TabIndex        =   15
      Top             =   1920
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   450
      Calculator      =   "frmcekgiro.frx":340D
      Caption         =   "frmcekgiro.frx":342D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmcekgiro.frx":3499
      Keys            =   "frmcekgiro.frx":34B7
      Spin            =   "frmcekgiro.frx":34F9
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483628
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,###,##0.00;(###,###,###,##0.00)"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999999999
      MinValue        =   -999999999999999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   1
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   775290885
      MinValueVT      =   1701576709
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   4200
      TabIndex        =   5
      Top             =   3120
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
      Format          =   88997891
      CurrentDate     =   37421
   End
   Begin Chameleon.chameleonButton cmdelete 
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   3570
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Batal Cair/Tolak"
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
      MICON           =   "frmcekgiro.frx":3521
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label11 
      Caption         =   "Nomor Giro"
      Height          =   255
      Left            =   2880
      TabIndex        =   27
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Tgl Jatuh Tempo"
      Height          =   255
      Left            =   2880
      TabIndex        =   26
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4200
      TabIndex        =   25
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label Label5 
      Caption         =   "Search Giro"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   510
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Cair/Tolak"
      Height          =   255
      Left            =   2880
      TabIndex        =   23
      Top             =   2790
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Tgl Cair/Tolak"
      Height          =   255
      Left            =   2880
      TabIndex        =   22
      Top             =   3150
      Width           =   1335
   End
   Begin MSForms.ComboBox cmbct 
      Height          =   285
      Left            =   4200
      TabIndex        =   4
      Top             =   2760
      Width           =   1575
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "2778;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      DropButtonStyle =   3
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label3 
      Caption         =   "No. Bukti"
      Height          =   255
      Left            =   2880
      TabIndex        =   20
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblkode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4200
      TabIndex        =   19
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label lblnoac 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4200
      TabIndex        =   18
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label lblbank 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4200
      TabIndex        =   17
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Label Label28 
      Caption         =   "Supplier"
      Height          =   255
      Left            =   2880
      TabIndex        =   16
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Nilai"
      Height          =   255
      Left            =   2880
      TabIndex        =   14
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "B a n k"
      Height          =   255
      Left            =   2880
      TabIndex        =   13
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "No. Account"
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblsup 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4200
      TabIndex        =   11
      Top             =   2280
      Width           =   4455
   End
End
Attribute VB_Name = "frmcekgiro"
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

Dim strcust, strkurs, str1, str2, str3 As String

Private Sub Check1_Click()
    cmdclear_Click
End Sub

Private Sub cmbct_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub cmbct_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
End Sub

Private Sub cmdadd_Click()
    If lblkode = "" Or txtapply = "" Or strcust = "" Or cmbct = "" Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If date1 > date2 Then
        MsgBox "Tanggal tolak/cair lebih kecil dari jatuh tempo.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If MsgBox("Are you sure want to Save this change ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from am_apcashsub where nobukti = '" & lblkode & "' And type = 'G' and nogiro = '" & txtapply & "' and year(tanggalcair) <> '1900'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Giro sudah cair.", vbExclamation, "Warning"
        OBJ.Close
        Exit Sub
    End If
    
    SQL = "select * from am_apcashsub where nobukti = '" & lblkode & "' And type = 'G' and nogiro = '" & txtapply & "' and year(tanggaltolak) <> '1900'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Giro sudah tolak.", vbExclamation, "Warning"
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    If cmbct = "Tolak" Then
        OBJ.Open dsn
        SQL = "select * from am_apcashlin where kodebayar = 'GT' and noapply = '" & txtapply & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "There is already apply for this giro.", vbExclamation, "Warning"
            OBJ.Close
            Exit Sub
        End If
        OBJ.Close
    End If
    
    If cmbct = "Tolak" Then
        OBJ.Open dsn
        SQL = "update am_apcashsub set tanggalcair = convert(datetime,'01/01/1900'),tanggaltolak = convert(datetime,'" & tanggal2 & "') from am_apcashsub where nobukti = '" & lblkode & "' And type = 'G' and nogiro = '" & txtapply & "'"
        Set RST = OBJ.Execute(SQL)
        OBJ.Close
        '==========
        OBJ.Open dsn
        SQL = "select Type,Bank,Byadmin,Jumlah from am_apcashsub where nobukti = '" & lblkode & "' and nogiro = '" & txtapply & "' and type='G'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            OBJ1.Open dsn
            SQL1 = "select noacc from am_auto where kodecomp = '" & str1 & "' and jurnal_ = 'Jurnal h'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then str3 = RST1!noacc Else str3 = ""
            OBJ1.Close
                
            OBJ1.Open dsn
            SQL1 = "insert into gl_transaksi "
            SQL1 = SQL1 + "(kdcomp, "
            SQL1 = SQL1 + "tgltrx, "
            SQL1 = SQL1 + "kdtrx, "
            SQL1 = SQL1 + "notrx, "
            SQL1 = SQL1 + "kurs, "
            SQL1 = SQL1 + "noactrx, "
            SQL1 = SQL1 + "desctrx, "
            SQL1 = SQL1 + "dbkrtrx, "
            SQL1 = SQL1 + "amounttrx, "
            SQL1 = SQL1 + "nilaitrx, "
            SQL1 = SQL1 + "currtrx, "
            SQL1 = SQL1 + "flag, "
            SQL1 = SQL1 + "flagprint, "
            SQL1 = SQL1 + "flagadjustment, "
            SQL1 = SQL1 + "cekbg, "
            SQL1 = SQL1 + "identry, "
            SQL1 = SQL1 + "idupdate, "
            SQL1 = SQL1 + "dateentry, "
            SQL1 = SQL1 + "dateupdate, "
            SQL1 = SQL1 + "lineitem)"
            
            SQL1 = SQL1 + " values"
            SQL1 = SQL1 + "('" & str1 & "',"
            SQL1 = SQL1 + "convert(datetime,'" & tanggal2 & "'),"
            SQL1 = SQL1 + "'MG',"
            SQL1 = SQL1 + "'" & lblkode & "',"
            SQL1 = SQL1 + "convert(money,'1'),"
            SQL1 = SQL1 + "'" & str3 & "',"
            SQL1 = SQL1 + "'Giro Tolak',"
            SQL1 = SQL1 + "'D',"
            SQL1 = SQL1 + "convert(money,'" & RST!jumlah & "'),"
            SQL1 = SQL1 + "convert(money,'" & RST!jumlah & "'),"
            SQL1 = SQL1 + "'" & str2 & "',"
            SQL1 = SQL1 + "'B',"
            SQL1 = SQL1 + "'J',"
            SQL1 = SQL1 + "'0',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "'auto',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(numeric,'1'))"
            Set RST1 = OBJ1.Execute(SQL1)
            OBJ1.Close
            
            OBJ1.Open dsn
            SQL1 = "select noacc from am_auto where kodecomp = '" & str1 & "' and jurnal_ = 'Jurnal i'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then str3 = RST1!noacc Else str3 = ""
            OBJ1.Close
            
            OBJ1.Open dsn
            SQL1 = "insert into gl_transaksi "
            SQL1 = SQL1 + "(kdcomp, "
            SQL1 = SQL1 + "tgltrx, "
            SQL1 = SQL1 + "kdtrx, "
            SQL1 = SQL1 + "notrx, "
            SQL1 = SQL1 + "kurs, "
            SQL1 = SQL1 + "noactrx, "
            SQL1 = SQL1 + "desctrx, "
            SQL1 = SQL1 + "dbkrtrx, "
            SQL1 = SQL1 + "amounttrx, "
            SQL1 = SQL1 + "nilaitrx, "
            SQL1 = SQL1 + "currtrx, "
            SQL1 = SQL1 + "flag, "
            SQL1 = SQL1 + "flagprint, "
            SQL1 = SQL1 + "flagadjustment, "
            SQL1 = SQL1 + "cekbg, "
            SQL1 = SQL1 + "identry, "
            SQL1 = SQL1 + "idupdate, "
            SQL1 = SQL1 + "dateentry, "
            SQL1 = SQL1 + "dateupdate, "
            SQL1 = SQL1 + "lineitem)"
            
            SQL1 = SQL1 + " values"
            SQL1 = SQL1 + "('" & str1 & "',"
            SQL1 = SQL1 + "convert(datetime,'" & tanggal2 & "'),"
            SQL1 = SQL1 + "'MG',"
            SQL1 = SQL1 + "'" & lblkode & "',"
            SQL1 = SQL1 + "convert(money,'1'),"
            SQL1 = SQL1 + "'" & str3 & "',"
            SQL1 = SQL1 + "'Giro Tolak',"
            SQL1 = SQL1 + "'K',"
            SQL1 = SQL1 + "convert(money,'" & RST!jumlah & "'),"
            SQL1 = SQL1 + "convert(money,'" & RST!jumlah & "'),"
            SQL1 = SQL1 + "'" & str2 & "',"
            SQL1 = SQL1 + "'B',"
            SQL1 = SQL1 + "'J',"
            SQL1 = SQL1 + "'0',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "'auto',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(numeric,'2'))"
            Set RST1 = OBJ1.Execute(SQL1)
            OBJ1.Close
        End If
        OBJ.Close
        '=============
    ElseIf cmbct = "Cair" Then
        OBJ.Open dsn
        SQL = "update am_apcashsub set tanggalcair = convert(datetime,'" & tanggal2 & "'),tanggaltolak = convert(datetime,'01/01/1900') from am_apcashsub where nobukti = '" & lblkode & "' And type = 'G' and nogiro = '" & txtapply & "'"
        Set RST = OBJ.Execute(SQL)
        OBJ.Close
        '==========
        OBJ.Open dsn
        SQL = "select Type,Bank,Byadmin,Jumlah from am_apcashsub where nobukti = '" & lblkode & "' and nogiro = '" & txtapply & "' and type='G'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            OBJ1.Open dsn
            SQL1 = "select noacc from am_auto where kodecomp = '" & str1 & "' and jurnal_ = 'Jurnal h'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then str3 = RST1!noacc Else str3 = ""
            OBJ1.Close
                
            OBJ1.Open dsn
            SQL1 = "insert into gl_transaksi "
            SQL1 = SQL1 + "(kdcomp, "
            SQL1 = SQL1 + "tgltrx, "
            SQL1 = SQL1 + "kdtrx, "
            SQL1 = SQL1 + "notrx, "
            SQL1 = SQL1 + "kurs, "
            SQL1 = SQL1 + "noactrx, "
            SQL1 = SQL1 + "desctrx, "
            SQL1 = SQL1 + "dbkrtrx, "
            SQL1 = SQL1 + "amounttrx, "
            SQL1 = SQL1 + "nilaitrx, "
            SQL1 = SQL1 + "currtrx, "
            SQL1 = SQL1 + "flag, "
            SQL1 = SQL1 + "flagprint, "
            SQL1 = SQL1 + "flagadjustment, "
            SQL1 = SQL1 + "cekbg, "
            SQL1 = SQL1 + "identry, "
            SQL1 = SQL1 + "idupdate, "
            SQL1 = SQL1 + "dateentry, "
            SQL1 = SQL1 + "dateupdate, "
            SQL1 = SQL1 + "lineitem)"
            
            SQL1 = SQL1 + " values"
            SQL1 = SQL1 + "('" & str1 & "',"
            SQL1 = SQL1 + "convert(datetime,'" & tanggal2 & "'),"
            SQL1 = SQL1 + "'MG',"
            SQL1 = SQL1 + "'" & lblkode & "',"
            SQL1 = SQL1 + "convert(money,'1'),"
            SQL1 = SQL1 + "'" & str3 & "',"
            SQL1 = SQL1 + "'Giro Cair',"
            SQL1 = SQL1 + "'D',"
            SQL1 = SQL1 + "convert(money,'" & RST!jumlah & "'),"
            SQL1 = SQL1 + "convert(money,'" & RST!jumlah & "'),"
            SQL1 = SQL1 + "'" & str2 & "',"
            SQL1 = SQL1 + "'B',"
            SQL1 = SQL1 + "'J',"
            SQL1 = SQL1 + "'0',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "'auto',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(numeric,'1'))"
            Set RST1 = OBJ1.Execute(SQL1)
            OBJ1.Close
            
            OBJ1.Open dsn
            SQL1 = "select noacc from am_autoaccbank where kodecomp = '" & str1 & "' and kodebank = '" & lblbank & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then str3 = RST1!noacc Else str3 = ""
            OBJ1.Close
            
            OBJ1.Open dsn
            SQL1 = "insert into gl_transaksi "
            SQL1 = SQL1 + "(kdcomp, "
            SQL1 = SQL1 + "tgltrx, "
            SQL1 = SQL1 + "kdtrx, "
            SQL1 = SQL1 + "notrx, "
            SQL1 = SQL1 + "kurs, "
            SQL1 = SQL1 + "noactrx, "
            SQL1 = SQL1 + "desctrx, "
            SQL1 = SQL1 + "dbkrtrx, "
            SQL1 = SQL1 + "amounttrx, "
            SQL1 = SQL1 + "nilaitrx, "
            SQL1 = SQL1 + "currtrx, "
            SQL1 = SQL1 + "flag, "
            SQL1 = SQL1 + "flagprint, "
            SQL1 = SQL1 + "flagadjustment, "
            SQL1 = SQL1 + "cekbg, "
            SQL1 = SQL1 + "identry, "
            SQL1 = SQL1 + "idupdate, "
            SQL1 = SQL1 + "dateentry, "
            SQL1 = SQL1 + "dateupdate, "
            SQL1 = SQL1 + "lineitem)"
            
            SQL1 = SQL1 + " values"
            SQL1 = SQL1 + "('" & str1 & "',"
            SQL1 = SQL1 + "convert(datetime,'" & tanggal2 & "'),"
            SQL1 = SQL1 + "'MG',"
            SQL1 = SQL1 + "'" & lblkode & "',"
            SQL1 = SQL1 + "convert(money,'1'),"
            SQL1 = SQL1 + "'" & str3 & "',"
            SQL1 = SQL1 + "'Giro Cair',"
            SQL1 = SQL1 + "'K',"
            SQL1 = SQL1 + "convert(money,'" & RST!jumlah & "'),"
            SQL1 = SQL1 + "convert(money,'" & RST!jumlah & "'),"
            SQL1 = SQL1 + "'" & str2 & "',"
            SQL1 = SQL1 + "'B',"
            SQL1 = SQL1 + "'J',"
            SQL1 = SQL1 + "'0',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "'auto',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(numeric,'2'))"
            Set RST1 = OBJ1.Execute(SQL1)
            OBJ1.Close
        End If
        OBJ.Close
        '=======
    End If
    
    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdclear_Click()
    date1 = Date
    date2 = Date
    lblkode = ""
    txtapply = ""
    lblbank = ""
    lblnoac = ""
    lblsup = ""
    strkurs = ""
    strcust = ""
    nilaikurs = 0
    txtneto = 0
    cmbct = ""
    Label6 = ""
    txtcari = ""
    List1.Clear
    
    txtapply.SetFocus
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdelete_Click()
    If lblkode = "" Or txtapply = "" Or cmbct = "" Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If MsgBox("Are you sure want to Cancel ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from am_apcashsub where nobukti = '" & lblkode & "' And type = 'G' and nogiro = '" & txtapply & "' and year(tanggalcair) <> '1900'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        SQL = "select * from am_apcashsub where nobukti = '" & lblkode & "' And type = 'G' and nogiro = '" & txtapply & "' and year(tanggaltolak) <> '1900'"
        Set RST = OBJ.Execute(SQL)
        If RST.EOF Then
            MsgBox "There is no data to cancel.", vbExclamation, "Warning"
            OBJ.Close
            Exit Sub
        End If
    End If
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "select * from gl_transaksi where kdcomp='" & str1 & "' and notrx = '" & lblkode & "' And kdtrx = 'MG'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        MsgBox "There is already ledger transaction, abort proces to cancel.", vbExclamation, "Warning"
        
        Exit Sub
    End If
    OBJ.Close
    
    If cmbct = "Tolak" Then
        OBJ.Open dsn
        SQL = "select * from am_apcashlin where kodebayar = 'GT' and noapply = '" & txtapply & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "There is already apply for this data.", vbExclamation, "Warning"
            OBJ.Close
            Exit Sub
        End If
        OBJ.Close
    End If
    
    OBJ.Open dsn
    SQL = "update am_apcashsub set tanggalcair = convert(datetime,'01/01/1900'),tanggaltolak = convert(datetime,'01/01/1900') from am_apcashsub where nobukti = '" & lblkode & "' And type = 'G' and nogiro = '" & txtapply & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Data Is Deleted, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdpro1_Click()
    If MsgBox("lanjutkan proses tolak ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
        
    'If cmbct = "Tolak" Then
    '    OBJ.Open dsn
    '    SQL = "select * from am_apcashlin where kodebayar = 'GT' and noapply = '" & txtapply & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If Not RST.EOF Then
    '        MsgBox "There is already apply for this giro.", vbExclamation, "Warning"
    '        OBJ.Close
    '        Exit Sub
    '    End If
    '    OBJ.Close
    'End If
    
    'tolak
    
        OBJ.Open dsn
        SQL = "select nobukti,tanggaltolak,Type,Bank,Byadmin,Jumlah from am_apcashsub where year(tanggaltolak) = '" & Year(date3) & "' and type='G' order by tanggaltolak"
        Set RST = OBJ.Execute(SQL)
        Do While Not RST.EOF
        'If Not RST.EOF Then
            OBJ1.Open dsn
            SQL1 = "select noacc from am_auto where kodecomp = '" & str1 & "' and jurnal_ = 'Jurnal h'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then str3 = RST1!noacc Else str3 = ""
            OBJ1.Close
                
            OBJ1.Open dsn
            SQL1 = "insert into gl_transaksi "
            SQL1 = SQL1 + "(kdcomp, "
            SQL1 = SQL1 + "tgltrx, "
            SQL1 = SQL1 + "kdtrx, "
            SQL1 = SQL1 + "notrx, "
            SQL1 = SQL1 + "kurs, "
            SQL1 = SQL1 + "noactrx, "
            SQL1 = SQL1 + "desctrx, "
            SQL1 = SQL1 + "dbkrtrx, "
            SQL1 = SQL1 + "amounttrx, "
            SQL1 = SQL1 + "nilaitrx, "
            SQL1 = SQL1 + "currtrx, "
            SQL1 = SQL1 + "flag, "
            SQL1 = SQL1 + "flagprint, "
            SQL1 = SQL1 + "flagadjustment, "
            SQL1 = SQL1 + "cekbg, "
            SQL1 = SQL1 + "identry, "
            SQL1 = SQL1 + "idupdate, "
            SQL1 = SQL1 + "dateentry, "
            SQL1 = SQL1 + "dateupdate, "
            SQL1 = SQL1 + "lineitem)"
            
            SQL1 = SQL1 + " values"
            SQL1 = SQL1 + "('" & str1 & "',"
            SQL1 = SQL1 + "convert(datetime,'" & Month(RST!tanggaltolak) & "/" & Day(RST!tanggaltolak) & "/" & Year(RST!tanggaltolak) & "'),"
            SQL1 = SQL1 + "'MG',"
            SQL1 = SQL1 + "'" & RST!nobukti & "',"
            SQL1 = SQL1 + "convert(money,'1'),"
            SQL1 = SQL1 + "'" & str3 & "',"
            SQL1 = SQL1 + "'Giro Tolak',"
            SQL1 = SQL1 + "'D',"
            SQL1 = SQL1 + "convert(money,'" & RST!jumlah & "'),"
            SQL1 = SQL1 + "convert(money,'" & RST!jumlah & "'),"
            SQL1 = SQL1 + "'" & str2 & "',"
            SQL1 = SQL1 + "'B',"
            SQL1 = SQL1 + "'J',"
            SQL1 = SQL1 + "'0',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "'auto',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(numeric,'1'))"
            Set RST1 = OBJ1.Execute(SQL1)
            OBJ1.Close
            
            OBJ1.Open dsn
            SQL1 = "select noacc from am_auto where kodecomp = '" & str1 & "' and jurnal_ = 'Jurnal i'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then str3 = RST1!noacc Else str3 = ""
            OBJ1.Close
            
            OBJ1.Open dsn
            SQL1 = "insert into gl_transaksi "
            SQL1 = SQL1 + "(kdcomp, "
            SQL1 = SQL1 + "tgltrx, "
            SQL1 = SQL1 + "kdtrx, "
            SQL1 = SQL1 + "notrx, "
            SQL1 = SQL1 + "kurs, "
            SQL1 = SQL1 + "noactrx, "
            SQL1 = SQL1 + "desctrx, "
            SQL1 = SQL1 + "dbkrtrx, "
            SQL1 = SQL1 + "amounttrx, "
            SQL1 = SQL1 + "nilaitrx, "
            SQL1 = SQL1 + "currtrx, "
            SQL1 = SQL1 + "flag, "
            SQL1 = SQL1 + "flagprint, "
            SQL1 = SQL1 + "flagadjustment, "
            SQL1 = SQL1 + "cekbg, "
            SQL1 = SQL1 + "identry, "
            SQL1 = SQL1 + "idupdate, "
            SQL1 = SQL1 + "dateentry, "
            SQL1 = SQL1 + "dateupdate, "
            SQL1 = SQL1 + "lineitem)"
            
            SQL1 = SQL1 + " values"
            SQL1 = SQL1 + "('" & str1 & "',"
            SQL1 = SQL1 + "convert(datetime,'" & Month(RST!tanggaltolak) & "/" & Day(RST!tanggaltolak) & "/" & Year(RST!tanggaltolak) & "'),"
            SQL1 = SQL1 + "'MG',"
            SQL1 = SQL1 + "'" & RST!nobukti & "',"
            SQL1 = SQL1 + "convert(money,'1'),"
            SQL1 = SQL1 + "'" & str3 & "',"
            SQL1 = SQL1 + "'Giro Tolak',"
            SQL1 = SQL1 + "'K',"
            SQL1 = SQL1 + "convert(money,'" & RST!jumlah & "'),"
            SQL1 = SQL1 + "convert(money,'" & RST!jumlah & "'),"
            SQL1 = SQL1 + "'" & str2 & "',"
            SQL1 = SQL1 + "'B',"
            SQL1 = SQL1 + "'J',"
            SQL1 = SQL1 + "'0',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "'auto',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(numeric,'2'))"
            Set RST1 = OBJ1.Execute(SQL1)
            OBJ1.Close
            
            RST.MoveNext
        Loop
        OBJ.Close
    
    MsgBox "proses tolak selesai.", vbInformation, "Information"
End Sub

Private Sub cmdpro2_Click()
    If MsgBox("lanjutkan proses cair ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    'Cair
    
        OBJ.Open dsn
        SQL = "select nobukti,tanggalcair,Type,Bank,Byadmin,Jumlah from am_apcashsub where year(tanggalcair) = '" & Year(date3) & "' and type='G' order by tanggalcair"
        Set RST = OBJ.Execute(SQL)
        Do While Not RST.EOF
            OBJ1.Open dsn
            SQL1 = "select noacc from am_auto where kodecomp = '" & str1 & "' and jurnal_ = 'Jurnal h'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then str3 = RST1!noacc Else str3 = ""
            OBJ1.Close
                
            OBJ1.Open dsn
            SQL1 = "insert into gl_transaksi "
            SQL1 = SQL1 + "(kdcomp, "
            SQL1 = SQL1 + "tgltrx, "
            SQL1 = SQL1 + "kdtrx, "
            SQL1 = SQL1 + "notrx, "
            SQL1 = SQL1 + "kurs, "
            SQL1 = SQL1 + "noactrx, "
            SQL1 = SQL1 + "desctrx, "
            SQL1 = SQL1 + "dbkrtrx, "
            SQL1 = SQL1 + "amounttrx, "
            SQL1 = SQL1 + "nilaitrx, "
            SQL1 = SQL1 + "currtrx, "
            SQL1 = SQL1 + "flag, "
            SQL1 = SQL1 + "flagprint, "
            SQL1 = SQL1 + "flagadjustment, "
            SQL1 = SQL1 + "cekbg, "
            SQL1 = SQL1 + "identry, "
            SQL1 = SQL1 + "idupdate, "
            SQL1 = SQL1 + "dateentry, "
            SQL1 = SQL1 + "dateupdate, "
            SQL1 = SQL1 + "lineitem)"
            
            SQL1 = SQL1 + " values"
            SQL1 = SQL1 + "('" & str1 & "',"
            SQL1 = SQL1 + "convert(datetime,'" & Month(RST!tanggalcair) & "/" & Day(RST!tanggalcair) & "/" & Year(RST!tanggalcair) & "'),"
            SQL1 = SQL1 + "'MG',"
            SQL1 = SQL1 + "'" & RST!nobukti & "',"
            SQL1 = SQL1 + "convert(money,'1'),"
            SQL1 = SQL1 + "'" & str3 & "',"
            SQL1 = SQL1 + "'Giro Cair',"
            SQL1 = SQL1 + "'D',"
            SQL1 = SQL1 + "convert(money,'" & RST!jumlah & "'),"
            SQL1 = SQL1 + "convert(money,'" & RST!jumlah & "'),"
            SQL1 = SQL1 + "'" & str2 & "',"
            SQL1 = SQL1 + "'B',"
            SQL1 = SQL1 + "'J',"
            SQL1 = SQL1 + "'0',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "'auto',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(numeric,'1'))"
            Set RST1 = OBJ1.Execute(SQL1)
            OBJ1.Close
            
            OBJ1.Open dsn
            SQL1 = "select noacc from am_autoaccbank where kodecomp = '" & str1 & "' and kodebank = '" & RST!bank & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then str3 = RST1!noacc Else str3 = ""
            OBJ1.Close
            
            OBJ1.Open dsn
            SQL1 = "insert into gl_transaksi "
            SQL1 = SQL1 + "(kdcomp, "
            SQL1 = SQL1 + "tgltrx, "
            SQL1 = SQL1 + "kdtrx, "
            SQL1 = SQL1 + "notrx, "
            SQL1 = SQL1 + "kurs, "
            SQL1 = SQL1 + "noactrx, "
            SQL1 = SQL1 + "desctrx, "
            SQL1 = SQL1 + "dbkrtrx, "
            SQL1 = SQL1 + "amounttrx, "
            SQL1 = SQL1 + "nilaitrx, "
            SQL1 = SQL1 + "currtrx, "
            SQL1 = SQL1 + "flag, "
            SQL1 = SQL1 + "flagprint, "
            SQL1 = SQL1 + "flagadjustment, "
            SQL1 = SQL1 + "cekbg, "
            SQL1 = SQL1 + "identry, "
            SQL1 = SQL1 + "idupdate, "
            SQL1 = SQL1 + "dateentry, "
            SQL1 = SQL1 + "dateupdate, "
            SQL1 = SQL1 + "lineitem)"
            
            SQL1 = SQL1 + " values"
            SQL1 = SQL1 + "('" & str1 & "',"
            SQL1 = SQL1 + "convert(datetime,'" & Month(RST!tanggalcair) & "/" & Day(RST!tanggalcair) & "/" & Year(RST!tanggalcair) & "'),"
            SQL1 = SQL1 + "'MG',"
            SQL1 = SQL1 + "'" & RST!nobukti & "',"
            SQL1 = SQL1 + "convert(money,'1'),"
            SQL1 = SQL1 + "'" & str3 & "',"
            SQL1 = SQL1 + "'Giro Cair',"
            SQL1 = SQL1 + "'K',"
            SQL1 = SQL1 + "convert(money,'" & RST!jumlah & "'),"
            SQL1 = SQL1 + "convert(money,'" & RST!jumlah & "'),"
            SQL1 = SQL1 + "'" & str2 & "',"
            SQL1 = SQL1 + "'B',"
            SQL1 = SQL1 + "'J',"
            SQL1 = SQL1 + "'0',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "'auto',"
            SQL1 = SQL1 + "'',"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
            SQL1 = SQL1 + "convert(numeric,'2'))"
            Set RST1 = OBJ1.Execute(SQL1)
            OBJ1.Close
            
            RST.MoveNext
        Loop
        OBJ.Close
        
    MsgBox "proses cair selesai.", vbInformation, "Information"
End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='264' and b.kodeuser = '1" & kuser & "'"
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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 119 Then
        date3 = Date
        Frame2.Visible = True
    End If
    If KeyCode = 120 Then Frame2.Visible = False
End Sub

Private Sub Form_Load()
    
    date1 = Date
    date2 = Date
    
    cmbct.AddItem "Cair"
    cmbct.AddItem "Tolak"
    
    OBJ.Open dsn
    SQL = "select kdcomp from gl_company"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then str1 = RST!kdcomp
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "select kdkurs from gl_kurs where base='1'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then str2 = RST!kdkurs
    OBJ.Close
End Sub

Private Sub List1_DblClick()
    lblbank = ""
    lblnoac = ""
    txtneto = 0
    lblsup = ""
    lblkode = ""
    txtapply = ""
    strcust = ""
    strkurs = ""
    nilaikurs = 0
    date2 = Date
    cmbct = ""
    date1 = Date
    Label6 = ""
    
    txtapply = List1.text
        
    OBJ.Open dsn
    SQL = "select * from am_apcashsub where nogiro = '" & txtapply & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblbank = RST!bank
        lblnoac = RST!acbank
        txtneto = RST!jumlah
        strcust = RST!kodesupp
        lblkode = RST!nobukti
        date1 = RST!tgljt
        Label6 = Format(date1, "dd MMMM yyyy")
        
        If Year(RST!tanggalcair) = 1900 And Year(RST!tanggaltolak) <> 1900 Then
            cmbct = "Tolak"
            date2 = RST!tanggaltolak
        ElseIf Year(RST!tanggalcair) <> 1900 And Year(RST!tanggaltolak) = 1900 Then
            cmbct = "Cair"
            date2 = RST!tanggalcair
        ElseIf Year(RST!tanggalcair) = 1900 And Year(RST!tanggaltolak) = 1900 Then
            cmbct = ""
            date2 = Date
        End If
        
        SQL = "select kodecur,nilaikurs from am_apcashhdr where kodesupp = '" & strcust & "' and nobkt = '" & lblkode & "'"
        Set RST = OBJ.Execute(SQL)
        strkurs = RST!kodecur
        nilaikurs = RST!nilaikurs
        
        SQL = "select kodesupp,namasupp from am_supplier where kodesupp = '" & strcust & "'"
        Set RST = OBJ.Execute(SQL)
        lblsup = RST!kodesupp & " - " & RST!namasupp
    End If
    OBJ.Close
End Sub

Private Sub txtapply_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtapply_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Function tanggal1()
    tanggal1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggal2()
    tanggal2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function

Private Sub txtcari_Change()
    List1.Clear
    
    If txtcari = "" Then Exit Sub
        
    OBJ.Open dsn
    If Check1.Value = 0 Then SQL = "select nogiro from am_apcashsub where type='g' and nogiro like '" & txtcari & "%' and year(tanggalcair) = '1900' and year(tanggaltolak) = '1900' order by nogiro"
    If Check1.Value = 1 Then SQL = "select nogiro from am_apcashsub where type='g' and nogiro like '" & txtcari & "%' and (year(tanggalcair) <> '1900' or year(tanggaltolak) <> '1900') order by nogiro"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        List1.AddItem RST!nogiro
    
        RST.MoveNext
    Loop
    OBJ.Close
End Sub

Private Sub txtcari_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
