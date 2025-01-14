VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmpermintaanedit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Permintaan Barang"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9885
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtnobukti 
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
      Left            =   1320
      MaxLength       =   17
      TabIndex        =   8
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtbagian 
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
      Left            =   1320
      MaxLength       =   30
      TabIndex        =   7
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox txtket 
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
      Height          =   765
      Left            =   6720
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox txtpemesan 
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
      Left            =   1320
      MaxLength       =   30
      TabIndex        =   5
      Top             =   1200
      Width           =   2655
   End
   Begin VB.PictureBox uncheck 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      Picture         =   "frmpermintaanedit.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox check 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5520
      Picture         =   "frmpermintaanedit.frx":034E
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox blank 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5040
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtbrg 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Left            =   4920
      MaxLength       =   100
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   1815
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   4920
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   450
      Calculator      =   "frmpermintaanedit.frx":0630
      Caption         =   "frmpermintaanedit.frx":0650
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpermintaanedit.frx":06BC
      Keys            =   "frmpermintaanedit.frx":06DA
      Spin            =   "frmpermintaanedit.frx":071C
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "##,###,##0.00;(##,###,##0.00);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "##,###,##0.00;(##,###,##0.00)"
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
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   8760
      TabIndex        =   9
      Top             =   4440
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
      MICON           =   "frmpermintaanedit.frx":0744
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
      Left            =   7800
      TabIndex        =   10
      Top             =   4440
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
      MICON           =   "frmpermintaanedit.frx":0A5E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdedit 
      Height          =   375
      Left            =   6840
      TabIndex        =   11
      Top             =   4440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Update"
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
      MICON           =   "frmpermintaanedit.frx":0D78
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
      Left            =   1320
      TabIndex        =   12
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
      Format          =   134742019
      CurrentDate     =   37426
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   7680
      TabIndex        =   13
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
      Format          =   134742019
      CurrentDate     =   37426
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   1815
      Left            =   0
      TabIndex        =   14
      Top             =   2400
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   3201
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin Chameleon.chameleonButton cmddel 
      Height          =   375
      Left            =   5880
      TabIndex        =   23
      Top             =   4440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Delete"
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
      MICON           =   "frmpermintaanedit.frx":1092
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch 
      Height          =   285
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "No Permintaan"
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
      MICON           =   "frmpermintaanedit.frx":13AC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   5520
      Shape           =   4  'Rounded Rectangle
      Top             =   4320
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "No Permintaan"
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
      Left            =   120
      TabIndex        =   20
      Top             =   150
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
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
      Left            =   120
      TabIndex        =   19
      Top             =   510
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bagian"
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
      Left            =   120
      TabIndex        =   18
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Selesai yang diminta"
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
      Left            =   5520
      TabIndex        =   17
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan"
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
      Left            =   5520
      TabIndex        =   16
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Pemesan"
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
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   0
      TabIndex        =   22
      Top             =   960
      Width           =   9975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   21
      Top             =   4200
      Width           =   9975
   End
End
Attribute VB_Name = "frmpermintaanedit"
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
Dim posrow As String

Private Sub cmdclear_Click()
    hapusgrid
    date1.Value = Date
    date2.Value = Date
    txtnobukti = ""
    txtpemesan = ""
    txtbagian = ""
    txtket = ""
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmddel_Click()
    If txtnobukti = "" Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        txtpemesan.SetFocus
        Exit Sub
    End If
    'cek po sudah terbit atau belum
    OBJ.Open dsn
    SQL = "Select * From am_perminapp where nobkt = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    
    If Not RST.EOF Then
        If RST!Status <> "0" Then
            MsgBox "Maaf, nomor permintaan " & txtnobukti & " sudah diotorisasi & tidak bisa dihapus", vbExclamation, AppName
            cmdclear_Click
            Exit Sub
        End If
    End If
    OBJ.Close
    
    If MsgBox("Are you sure want to delete ?", vbQuestion + vbYesNo, "Question") = vbNo Then
        cmdclear_Click
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "delete from am_perminhdr where nobkt = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)

    SQL = "delete from am_perminin where nobkt = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "delete from am_perminapp where nobkt = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close

    MsgBox "Data Is Deleted, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdedit_Click()
    If txtnobukti = "" Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        txtpemesan.SetFocus
        Exit Sub
    End If
    
    If grid.Rows = 2 Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    grid.Row = 1
    Do While True
        If grid.Rows = grid.Row + 1 Then Exit Do
        
        If grid.TextMatrix(grid.Row, 2) = "0.00" Or grid.TextMatrix(grid.Row, 2) = "" Then
            MsgBox "Data entry not complete, on row " & grid.Row, vbExclamation, "Warning"
            Exit Sub
        End If
        If grid.TextMatrix(grid.Row, 3) = "" Then
            MsgBox "Data entry not complete, on row " & grid.Row, vbExclamation, "Warning"
            Exit Sub
        End If
        If grid.TextMatrix(grid.Row, 4) = "" Then
            MsgBox "Data entry not complete, on row " & grid.Row, vbExclamation, "Warning"
            Exit Sub
        End If
        
        grid.Row = grid.Row + 1
    Loop
   
   'cek status NP jika sudah close tidak bisa edit
    OBJ.Open dsn
    SQL = "Select * From am_perminhdr Where nobkt='" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If RST!flag = "1" Then
            MsgBox "Sorry, Nota cannot be changed, because it has been processed", vbExclamation, AppName
            OBJ.Close
            cmdclear_Click
        End If
    End If
    OBJ.Close
    
    If MsgBox("Are you sure want to update ?", vbQuestion + vbYesNo, "Question") = vbNo Then
        cmdclear_Click
        Exit Sub
    End If
   
        OBJ.Open dsn
        SQL = "delete from am_perminhdr where nobkt = '" & txtnobukti & "'"
        Set RST = OBJ.Execute(SQL)

        SQL = "delete from am_perminin where nobkt = '" & txtnobukti & "'"
        Set RST = OBJ.Execute(SQL)

        SQL = "delete from am_perminapp where nobkt = '" & txtnobukti & "'"
        Set RST = OBJ.Execute(SQL)
        'UPDATE KE TABLE PERMINHDR

        SQL = "insert into am_perminhdr ("
        SQL = SQL + "nobkt, "
        SQL = SQL + "tglbkt, "
        SQL = SQL + "divisi, "
        SQL = SQL + "pemesan, "
        SQL = SQL + "tglselesai, "
        SQL = SQL + "keterangan, "
        SQL = SQL + "flag)"
    
        SQL = SQL + " values ('" & txtnobukti & "',"
        SQL = SQL + "convert(datetime,'" & tanggal1 & "'),"
        SQL = SQL + "'" & txtbagian & "',"
        SQL = SQL + "'" & txtpemesan & "',"
        SQL = SQL + "convert(datetime,'" & tanggal2 & "'),"
        SQL = SQL + "'" & txtket & "',"
        SQL = SQL + "'0')"
        Set RST = OBJ.Execute(SQL)
    
        'UPDATE KE TABLE PERMININ
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        SQL = "insert into am_perminin ("
        SQL = SQL + "nobkt, "
        SQL = SQL + "nmbrg, "
        SQL = SQL + "qty, "
        SQL = SQL + "pekerja, "
        SQL = SQL + "keperluan, "
        SQL = SQL + "lineitem, "
        SQL = SQL + "status, "
        SQL = SQL + "nopo, "
        SQL = SQL + "nopo2)"

        SQL = SQL + " values ('" & txtnobukti & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 1) & "',"
        SQL = SQL + "convert(money,'" & Format(grid.TextMatrix(grid.Row, 2), "general number") & "'),"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 3) & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 4) & "',"
        SQL = SQL + "convert(numeric,'" & grid.Row & "'),'0','','')"
        Set RST = OBJ.Execute(SQL)
        grid.Row = grid.Row + 1
        DoEvents
    Loop
    
    SQL = "Insert into am_perminapp ("
    SQL = SQL + "nobkt, "
    SQL = SQL + "nopo, "
    SQL = SQL + "kodesupp, "
    SQL = SQL + "status, "
    SQL = SQL + "tglbkt, "
    SQL = SQL + "tglpo, "
    SQL = SQL + "tgl)"
    SQL = SQL + " values('" & txtnobukti & "',"
    SQL = SQL + "'',"
    SQL = SQL + "'','0',"
    SQL = SQL + "convert(datetime,'" & tanggal1 & "'),"
    SQL = SQL + "convert(datetime,'" & tanggal1 & "'),"
    SQL = SQL + "convert(datetime,'" & tanggal1 & "'))"
    Set RST = OBJ.Execute(SQL)
    
    OBJ.Close
    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub
Function tanggal1()
    tanggal1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function
Function tanggal2()
    tanggal2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function
Private Sub cmdsearch_Click()
    carisql1 = "Select nobkt,pemesan From am_perminhdr"
    namatabel = "Permintaan"
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    txtnobukti = hasil
    hasil = ""
    hasil1 = ""
    Call History
End Sub
Private Sub txtbrg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    posrow = grid.Row
    Select Case grid.Col
        Case 1:
            If txtbrg = "" Then
                txtbrg.Visible = False
                grid.SetFocus
                grid.Row = posrow
            Exit Sub
            End If
            If Len(Trim(txtbrg)) > 30 Then
                MsgBox "Max.30 karakter", vbExclamation, AppName
                Exit Sub
            End If
            grid.TextMatrix(grid.Row, 1) = txtbrg.text
            SetRow grid.Row, True
            If grid.Row = (grid.Rows - 1) Then grid.Rows = grid.Rows + 1
            grid.SetFocus
        Case 3:
            If Len(Trim(txtbrg)) > 30 Then
                MsgBox "Max.30 karakter", vbExclamation, AppName
                Exit Sub
            End If
            grid.TextMatrix(grid.Row, 3) = txtbrg.text
        Case 4:
            grid.TextMatrix(grid.Row, 4) = txtbrg.text
    End Select
        txtbrg = ""
        txtbrg.Visible = False
        grid.SetFocus
        grid.Row = posrow
    End If
End Sub

Private Sub txtbrg_LostFocus()
    txtbrg.Visible = False
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grid.TextMatrix(grid.Row, 2) = Format(txtnilai, "#,##0.00")
        txtnilai = 0
        txtnilai.Visible = False
        grid.SetFocus
        grid.Row = posrow
    End If
End Sub
Private Sub hapusrow()
    grid.TextMatrix(grid.Row, 1) = ""
    grid.TextMatrix(grid.Row, 2) = ""
    grid.TextMatrix(grid.Row, 3) = ""
    grid.TextMatrix(grid.Row, 4) = ""
    Do While True
        If grid.TextMatrix(grid.Row + 1, 1) = "" Then
            grid.TextMatrix(grid.Row, 1) = ""
            grid.TextMatrix(grid.Row, 2) = ""
            grid.TextMatrix(grid.Row, 3) = ""
            grid.TextMatrix(grid.Row, 4) = ""
            Exit Do
        End If
        grid.TextMatrix(grid.Row, 1) = grid.TextMatrix(grid.Row + 1, 1)
        grid.TextMatrix(grid.Row, 2) = grid.TextMatrix(grid.Row + 1, 2)
        grid.TextMatrix(grid.Row, 3) = grid.TextMatrix(grid.Row + 1, 3)
        grid.TextMatrix(grid.Row, 4) = grid.TextMatrix(grid.Row + 1, 4)
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = grid.Rows - 1
    grid.Col = 0
    Set grid.CellPicture = blank
End Sub

Private Sub hapusgrid()
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        grid.TextMatrix(grid.Row, 1) = ""
        grid.TextMatrix(grid.Row, 2) = ""
        grid.TextMatrix(grid.Row, 3) = ""
        grid.TextMatrix(grid.Row, 4) = ""
        grid.Col = 0
        Set grid.CellPicture = blank
        grid.Row = grid.Row + 1
    Loop
    
    grid.Rows = 2
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 2500
    grid.ColWidth(2) = 1000
    grid.ColWidth(3) = 1500
    grid.ColWidth(4) = 4000
End Sub

Private Sub History()
    'cek po sudah terbit atau belum
    OBJ.Open dsn
    SQL = "Select * From am_perminapp where nobkt = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    
    If Not RST.EOF Then
        If RST!Status <> "0" Then
            MsgBox "Maaf, nomor permintaan " & txtnobukti & " sudah diotorisasi & tidak bisa diedit", vbExclamation, AppName
            cmdclear_Click
            Exit Sub
        End If
    End If
    OBJ.Close
    
    Call hapusgrid
    OBJ.Open dsn
    SQL = "Select * From am_perminhdr Where nobkt = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    
    If Not RST.EOF Then
        txtpemesan = RST!pemesan
        txtbagian = RST!divisi
        txtket = RST!keterangan
        date1.Value = RST!tglbkt
        date2.Value = RST!tglselesai
        
        grid.Row = 1
        SQL = "Select * From am_perminin Where nobkt = '" & txtnobukti & "' order by lineitem asc"
        Set RST = OBJ.Execute(SQL)
        Do Until RST.EOF
            grid.Col = 1
            grid.CellAlignment = 1
            grid.TextMatrix(grid.Row, 1) = RST!nmbrg
            grid.TextMatrix(grid.Row, 2) = RST!qty
            grid.TextMatrix(grid.Row, 3) = RST!pekerja
            grid.TextMatrix(grid.Row, 4) = RST!keperluan
            
            SetRow grid.Row, True
            grid.Rows = grid.Rows + 1
            grid.Row = grid.Row + 1
            RST.MoveNext
        Loop
    End If
    OBJ.Close
End Sub

Private Sub SetRow(ByVal idx As Integer, ByVal hapus As String)
    grid.Row = idx
    grid.Col = 0
    If hapus Then
        Set grid.CellPicture = uncheck.Picture
    End If
    grid.Col = 1
End Sub

Private Sub Form_Load()
    grid.TextMatrix(0, 1) = "Nama Barang"
    grid.TextMatrix(0, 2) = "Qty"
    grid.TextMatrix(0, 3) = "Dikerjakan Oleh"
    grid.TextMatrix(0, 4) = "Keperluan"
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 2500
    grid.ColWidth(2) = 1000
    grid.ColWidth(3) = 1500
    grid.ColWidth(4) = 4000
    grid.RowHeightMin = 300
    date1.Value = Date
    date2.Value = Date
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    If txtnobukti = "" Then Exit Sub
    
    posrow = grid.Row
    Select Case grid.Col
        Case 0:
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            If grid.CellPicture = uncheck Then
                Set grid.CellPicture = check
                If MsgBox("Delete Row ?", vbQuestion + vbYesNo, "Question") = vbYes Then
                    Set grid.CellPicture = uncheck
                    hapusrow
                    Exit Sub
                End If
                Set grid.CellPicture = uncheck
            End If
        Case 1:
            txtbrg.Width = grid.ColWidth(grid.Col) - 40
            txtbrg = grid.TextMatrix(grid.Row, grid.Col)
            txtbrg.Left = grid.Left + grid.CellLeft
            txtbrg.Top = grid.Top + grid.CellTop + 20
            txtbrg.Visible = True
            txtbrg.SetFocus
        Case 2:
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            
            txtnilai.Width = grid.ColWidth(grid.Col) - 40
            txtnilai = grid.TextMatrix(grid.Row, grid.Col)
            txtnilai.Left = grid.Left + grid.CellLeft
            txtnilai.Top = grid.Top + grid.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
        Case 3, 4:
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            
            txtbrg.Width = grid.ColWidth(grid.Col) - 40
            txtbrg = grid.TextMatrix(grid.Row, grid.Col)
            txtbrg.Left = grid.Left + grid.CellLeft
            txtbrg.Top = grid.Top + grid.CellTop + 20
            txtbrg.Visible = True
            txtbrg.SetFocus
    End Select
End Sub

Private Sub grid_EnterCell()
    If grid.MouseRow = 0 Then Exit Sub
    If txtnobukti = "" Then Exit Sub
    
    Select Case grid.Col
    Case 1
        posrow = grid.Row
        
        txtbrg.Width = grid.ColWidth(grid.Col) - 40
        txtbrg = grid.TextMatrix(grid.Row, grid.Col)
        txtbrg.Left = grid.Left + grid.CellLeft
        txtbrg.Top = grid.Top + grid.CellTop + 20
        txtbrg.Visible = True
        txtbrg.SetFocus
    Case 2
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            
        posrow = grid.Row
        
        txtnilai.Width = grid.ColWidth(grid.Col) - 40
        txtnilai = grid.TextMatrix(grid.Row, grid.Col)
        txtnilai.Left = grid.Left + grid.CellLeft
        txtnilai.Top = grid.Top + grid.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
    Case 3, 4
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
        posrow = grid.Row
        
        txtbrg.Width = grid.ColWidth(grid.Col) - 40
        txtbrg = grid.TextMatrix(grid.Row, grid.Col)
        txtbrg.Left = grid.Left + grid.CellLeft
        txtbrg.Top = grid.Top + grid.CellTop + 20
        txtbrg.Visible = True
        txtbrg.SetFocus
    End Select
End Sub

Private Sub txtnobukti_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then date1.SetFocus
End Sub

Private Sub txtnobukti_LostFocus()
    Call History
End Sub

