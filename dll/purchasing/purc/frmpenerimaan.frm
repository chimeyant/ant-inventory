VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmpenerimaan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Penerimaan Barang"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9975
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
   Icon            =   "frmpenerimaan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3645
      Top             =   870
   End
   Begin VB.CheckBox chkmanual 
      Caption         =   "OTOMATIS"
      Height          =   300
      Left            =   3345
      TabIndex        =   27
      Top             =   135
      Value           =   1  'Checked
      Width           =   1710
   End
   Begin VB.Frame Frame1 
      Caption         =   "Note"
      Height          =   2175
      Left            =   6960
      TabIndex        =   23
      Top             =   120
      Width           =   2895
      Begin VB.Label Label7 
         Caption         =   "Jika Qty (PO) dengan Qty (Use) adalah sama maka Satuan (PO) dan Satuan (Use) diharuskan sama juga."
         Height          =   615
         Left            =   120
         TabIndex        =   26
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label6 
         Caption         =   "Qty (USE) : Quantity diisi sesuai dengan quantity yang nantinya dipakai di modul Pemakaian."
         Height          =   615
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "Qty (PO) : Quantity diisi sesuai dengan quantity yang di PO."
         Height          =   495
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.TextBox txtsj 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtdriver 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   40
      TabIndex        =   5
      Top             =   2040
      Width           =   5415
   End
   Begin VB.TextBox txtkend 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   40
      TabIndex        =   4
      Top             =   1680
      Width           =   5415
   End
   Begin TDBNumber6Ctl.TDBNumber txtnilai 
      Height          =   255
      Left            =   5160
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   450
      Calculator      =   "frmpenerimaan.frx":2372
      Caption         =   "frmpenerimaan.frx":2392
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpenerimaan.frx":23FE
      Keys            =   "frmpenerimaan.frx":241C
      Spin            =   "frmpenerimaan.frx":245E
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
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.TextBox txtpo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtnobukti 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      MaxLength       =   17
      TabIndex        =   0
      Top             =   120
      Width           =   1815
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
      Left            =   5400
      Picture         =   "frmpenerimaan.frx":2486
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   13
      Top             =   480
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
      Left            =   5640
      Picture         =   "frmpenerimaan.frx":27D4
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   12
      Top             =   480
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
      Left            =   5160
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1440
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
      Format          =   135200771
      CurrentDate     =   37426
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   2055
      Left            =   0
      TabIndex        =   7
      Top             =   2400
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   3625
      _Version        =   393216
      Cols            =   10
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
      _Band(0).Cols   =   10
   End
   Begin Chameleon.chameleonButton cmdsearch1 
      Height          =   285
      Left            =   240
      TabIndex        =   16
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "No P.O."
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
      MICON           =   "frmpenerimaan.frx":2AB6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   8760
      TabIndex        =   10
      Top             =   4680
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
      MICON           =   "frmpenerimaan.frx":2DD0
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
      TabIndex        =   9
      Top             =   4680
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
      MICON           =   "frmpenerimaan.frx":30EA
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
      Left            =   6840
      TabIndex        =   8
      Top             =   4680
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   3
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
      MICON           =   "frmpenerimaan.frx":3404
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber txtnil1 
      Height          =   225
      Left            =   4200
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmpenerimaan.frx":371E
      Caption         =   "frmpenerimaan.frx":373E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpenerimaan.frx":37AA
      Keys            =   "frmpenerimaan.frx":37C8
      Spin            =   "frmpenerimaan.frx":380A
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
      HighlightText   =   0
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
      ReadOnly        =   1
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin TDBNumber6Ctl.TDBNumber txtnil2 
      Height          =   225
      Left            =   4200
      TabIndex        =   21
      Top             =   360
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   397
      Calculator      =   "frmpenerimaan.frx":3832
      Caption         =   "frmpenerimaan.frx":3852
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmpenerimaan.frx":38BE
      Keys            =   "frmpenerimaan.frx":38DC
      Spin            =   "frmpenerimaan.frx":391E
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
      HighlightText   =   0
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
      ReadOnly        =   1
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin Chameleon.chameleonButton chameleonButton1 
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   4680
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Preview no save"
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
      MICON           =   "frmpenerimaan.frx":3946
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   6720
      Shape           =   4  'Rounded Rectangle
      Top             =   4560
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "Keterangan"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   2070
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "No Surat Jalan"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   1350
      Width           =   1455
   End
   Begin VB.Label Label31 
      Caption         =   "Driver / No.Kendaraan"
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   1620
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "No LPB"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   150
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Tanggal LPB"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   510
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   22
      Top             =   4440
      Width           =   9975
   End
End
Attribute VB_Name = "frmpenerimaan"
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

Dim SP As New ADODB.Command

Dim posrow As String
Dim bo1 As Boolean
Dim i1, i2 As Integer

Private Sub chameleonButton1_Click()
Cetak_Bukti
End Sub


Private Sub cmdadd_Click()
    On Error GoTo err_handler:
        
    If Len(Trim(txtnobukti)) = 0 Then
        MsgBox "Data Entry Not Complete.", vbExclamation, "Warning"
        txtnobukti.SetFocus
        Exit Sub
    End If
    
    If txtnobukti = "" Or txtpo = "" Or grid.Rows = 2 Then
        MsgBox "Data Entry Not Complete", vbExclamation, "Warning"
        Exit Sub
    End If
    
    i1 = 1
    Do While txtpo <> "" And i1 <> Len(txtpo)
        If Asc(Mid(txtpo, i1, 1)) = 47 Then Exit Do
        i1 = i1 + 1
    Loop
    
    i2 = 1
    Do While txtnobukti <> "" And i2 <> Len(txtnobukti)
        If Asc(Mid(txtnobukti, i2, 1)) = 47 Then Exit Do
        i2 = i2 + 1
    Loop
    
    If Mid(txtpo, i1 + 1, Len(txtpo) - i1 + 1) <> Mid(txtnobukti, i2 + 1, Len(txtnobukti) - i2 + 1) Then
        MsgBox "Invalid division between LPB and PO.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If Mid(txtnobukti, 10, 1) <> "/" Then
        MsgBox "Mohon periksa nobukti, ada kesalahan pada format nobukti.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    bo1 = True
    OBJ1.Open dsn
    SQL1 = "select *,len(kode)'lebar' from am_nomax"
    Set RST1 = OBJ1.Execute(SQL1)
    Do While Not RST1.EOF
        If (RST1!lebar = 4 And Right(Trim(txtpo), 4) = RST1!kode) Or _
        (RST1!lebar = 5 And Right(Trim(txtpo), 5) = RST1!kode) Or _
        (RST1!lebar = 6 And Right(Trim(txtpo), 6) = RST1!kode) Then
            bo1 = False
            Exit Do
        End If
        RST1.MoveNext
    Loop
    OBJ1.Close
    
    grid.Row = 1
    Do While True
        If grid.Rows = grid.Row + 1 Then Exit Do
        
        If grid.TextMatrix(grid.Row, 3) = "0.00" Or grid.TextMatrix(grid.Row, 6) = "0.00" Then
            MsgBox "Data entry not complete, on row " & grid.Row, vbExclamation, "Warning"
            Exit Sub
        End If
        
        If (grid.TextMatrix(grid.Row, 4) = grid.TextMatrix(grid.Row, 7)) And (Val(Format(grid.TextMatrix(grid.Row, 3), "general number")) <> Val(Format(grid.TextMatrix(grid.Row, 6), "general number"))) Then
            MsgBox "Qty <> QtyUse , on row " & grid.Row, vbExclamation, "Warning"
            Exit Sub
        End If
                
        If bo1 Then
            OBJ1.Open dsn
            SQL1 = "select qty from am_polin where nopo = '" & txtpo & "' and kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil1 = RST1!qty
            Else
                txtnil1 = 0
            End If
            
            SQL1 = "select isnull(sum(a.qty),0)'qty' from am_belilin a left join am_belihdr b on a.nobeli=b.nobeli where b.nopo = '" & txtpo & "' and kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
            Set RST1 = OBJ1.Execute(SQL1)
            If Not RST1.EOF Then
                txtnil2 = RST1!qty
            Else
                txtnil2 = 0
            End If
            OBJ1.Close
            
            If Val(Format(grid.TextMatrix(grid.Row, 3), "general number")) > (txtnil1 - txtnil2) Then
                MsgBox "Purchase Order required, Qty max = " & (txtnil1 - txtnil2), vbExclamation, "Information"
                Exit Sub
            End If
        End If
        
        grid.Row = grid.Row + 1
    Loop
        
    OBJ.Open dsn
    SQL = "select * from am_belihdr where nobeli = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        
        MsgBox "Data Already Exist, Click OK To Continue ...", vbInformation, "Information"
        cmdclear_Click
        Exit Sub
    End If
    OBJ.Close
    
    'SIMPAN KE TABLE BELI HEADER
    OBJ.Open dsn
    SQL = "insert into am_belihdr ("
    SQL = SQL + "nobeli, "
    SQL = SQL + "tglbeli, "
    SQL = SQL + "nopo, "
    SQL = SQL + "nosj, "
    SQL = SQL + "nokend, "
    SQL = SQL + "driver, "
    SQL = SQL + "terima)"

    SQL = SQL + " values ('" & txtnobukti & "',"
    SQL = SQL + "convert(datetime,'" & tanggalpo & "'),"
    SQL = SQL + "'" & txtpo & "',"
    SQL = SQL + "'" & txtsj & "',"
    SQL = SQL + "'" & txtkend & "',"
    SQL = SQL + "'" & txtdriver & "',"
    SQL = SQL + "'1')"
    Set RST = OBJ.Execute(SQL)

    
    'SIMPAN KE TABLE DETAIL BELI
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do

        SQL = "insert into am_belilin ("
        SQL = SQL + "nobeli, "
        SQL = SQL + "kodebarang, "
        SQL = SQL + "qty, "
        SQL = SQL + "kodesatuan, "
        SQL = SQL + "qtyUse, "
        SQL = SQL + "kodesatuanuse, "
        SQL = SQL + "lineitem)"

        SQL = SQL + " values ('" & txtnobukti & "',"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 1) & "',"
        SQL = SQL + "convert(money,'" & Format(grid.TextMatrix(grid.Row, 3), "general number") & "'),"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 4) & "',"
        SQL = SQL + "convert(money,'" & Format(grid.TextMatrix(grid.Row, 6), "general number") & "'),"
        SQL = SQL + "'" & grid.TextMatrix(grid.Row, 7) & "',"
        SQL = SQL + "convert(numeric,'" & grid.Row & "'))"
        Set RST = OBJ.Execute(SQL)
        grid.Row = grid.Row + 1
        DoEvents
    Loop
    
    'SIMPAN KE TABLE BELI APPLY UNTUK DI KONFIRM
    SQL = "SELECT b.nobeli,b.tglbeli,b.nopo,c.kodecur,c.nilaikurs,a.kodebarang,a.qty,d.price,a.kodesatuan,a.lineitem,c.kodesupp,b.driver,c.ket1,c.ket2,c.ket3 FROM am_belilin a left join AM_belihdr b on a.nobeli=b.nobeli left join am_pohdr c on b.nopo=c.nopo left join am_polin d on a.kodebarang=d.kodebarang and d.nopo=b.nopo WHERE b.nobeli = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        OBJ1.Open dsn
        SQL1 = "select * from AM_beliapp where nobeli = '" & RST!NoBeli & "' and kodebarang = '" & RST!kodebarang & "'"
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
            SQL1 = SQL1 + " ('" & RST!NoBeli & "'"
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
            SQL1 = SQL1 + ", '1'"
            SQL1 = SQL1 + ", '0')"

            Set RST1 = OBJ1.Execute(SQL1)
        End If
        OBJ1.Close
        RST.MoveNext
        DoEvents
    Loop
    
    OBJ.Close
    
    Cetak_Bukti

    MsgBox "Data Is Saved, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
    Exit Sub
err_handler:
    If OBJ.State = 1 Then OBJ.Close
    MsgBox Err.Description, vbCritical, AppName
End Sub

Private Sub cmdclear_Click()
    hapusgrid
    
    txtnobukti = ""
    If Date > date1.MaxDate Then
        date1 = date1.MaxDate
    ElseIf Date < date1.MinDate Then
        date1 = date1.MinDate
    Else
        date1.Value = Date
    End If
    txtpo = ""
    txtsj = ""
    txtkend = ""
    txtdriver = ""
    
    txtnobukti.SetFocus
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdsearch1_Click()
    SP.ActiveConnection = dsn
    SP.CommandType = adCmdStoredProc
    SP.CommandText = "am_caripo"
    SP.Execute
    Set SP = Nothing
    
    If v_fastsearching = True Then
        If v_fstgl1 > v_fstgl2 Then
            MsgBox "Invalid date range, search abort.", vbExclamation, "Error"
            Exit Sub
        End If
        carisql1 = "select nopo, convert(char(11),tglpo)'tglpo' from am_findpo where tglpo >= '" & batas1 & "' and tglpo <= '" & batas2 & "' and tglpo <= '" & tanggalpo & "'"
    Else
        carisql1 = "select nopo, convert(char(11),tglpo)'tglpo' from am_findpo where tglpo <= '" & tanggalpo & "'"
    End If
    namatabel = "Purchase Order "
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtpo = hasil
    If chkmanual.Value = 1 Then
        History
    End If
   
    hapusgrid
    hasil = ""
    hasil1 = ""
    txtsj.SetFocus
End Sub

Private Sub date1_Change()
    If chkmanual.Value = 0 Then Exit Sub
    timer1.Enabled = True
End Sub

Private Sub Form_Activate()
    OBJ.Open dsn
    SQL = "select * from am_period"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "The period is empty !!" & vbCrLf & _
        "Please define Period on proces, Starting period date and Ending period date.", vbCritical, "Critical"
        
        OBJ.Close
        Unload Me
        Exit Sub
    End If
    OBJ.Close
    
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    
    
    grid.TextMatrix(0, 1) = "Kode Barang"
    grid.TextMatrix(0, 2) = "Nama Barang"
    grid.TextMatrix(0, 3) = "Qty (PO)"
    grid.TextMatrix(0, 4) = "K/Sat."
    grid.TextMatrix(0, 5) = "Satuan"
    grid.TextMatrix(0, 6) = "Qty (USE)"
    grid.TextMatrix(0, 7) = "K/Sat."
    grid.TextMatrix(0, 8) = "Satuan"
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1200
    grid.ColWidth(2) = 2500
    grid.ColWidth(3) = 1000
    grid.ColWidth(4) = 800
    grid.ColWidth(5) = 1000
    grid.ColWidth(6) = 1000
    grid.ColWidth(7) = 800
    grid.ColWidth(8) = 1000
    grid.ColWidth(9) = 0
    
    grid.RowHeightMin = 300
    
    date1.Value = Date
    
    OBJ.Open dsn
    SQL = "select * from am_period"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        date1.MinDate = RST!tanggal1
        date1.MaxDate = RST!tanggal2
    End If
    OBJ.Close
End Sub

Private Sub grid_Click()
    If grid.MouseRow = 0 Then Exit Sub
    If txtnobukti = "" Or txtpo = "" Then Exit Sub
    
    posrow = grid.Row
    Select Case grid.Col
        Case 0
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
        Case 1
            If grid.TextMatrix(grid.Row, 1) <> "" Then Exit Sub
            If grid.Rows - 1 = 50 Then
                MsgBox "Maximum line 50", vbExclamation, "Warning"
                Exit Sub
            End If
            If grid.Row <> 1 And grid.TextMatrix(grid.Row - 1, 1) = "" Then Exit Sub
            
            carisql1 = "select a.kodebarang, a.kodesatuan, b.namabarang from am_polin a left join am_apitemmst b on a.kodebarang=b.kodebarang where a.nopo = '" & txtpo & "'"
            namatabel = "Item on PO"
            
            frmsearch.Show vbModal
        Case 3, 6
            If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            
            txtnilai.Width = grid.ColWidth(grid.Col) - 40
            txtnilai = grid.TextMatrix(grid.Row, grid.Col)
            txtnilai.Left = grid.Left + grid.CellLeft
            txtnilai.Top = grid.Top + grid.CellTop + 20
            txtnilai.Visible = True
            txtnilai.SetFocus
    End Select
End Sub

Private Sub grid_EnterCell()
    If grid.MouseRow = 0 Then Exit Sub
    If txtnobukti = "" Or txtpo = "" Then Exit Sub
    
    Select Case grid.Col
    Case 3, 6
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Sub
            
        posrow = grid.Row
        
        txtnilai.Width = grid.ColWidth(grid.Col) - 40
        txtnilai = grid.TextMatrix(grid.Row, grid.Col)
        txtnilai.Left = grid.Left + grid.CellLeft
        txtnilai.Top = grid.Top + grid.CellTop + 20
        txtnilai.Visible = True
        txtnilai.SetFocus
    End Select
End Sub

Private Sub grid_GotFocus()
    If hasil = "" Then Exit Sub
    
    Select Case grid.Col
        Case 1
            grid.Row = 1
            Do While True
                If grid.Rows = 2 Or grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
                If grid.TextMatrix(grid.Row, 1) = hasil And posrow <> grid.Row Then
                
                    MsgBox "Item already exist.", vbInformation, "Information"
                    hasil = ""
                    hasil1 = ""
                    grid.Row = posrow
                    grid.Col = 1
                    grid.SetFocus
                    Exit Sub
                End If
                grid.Row = grid.Row + 1
            Loop
            
            grid.Row = posrow
            grid.Col = 1
            grid.CellAlignment = 1
            grid.TextMatrix(grid.Row, 1) = hasil
            grid.Col = 4
            grid.CellAlignment = 1
            grid.TextMatrix(grid.Row, 4) = hasil1

            hasil = ""
            hasil1 = ""
            hasil2 = ""
            
            bo1 = True
            OBJ1.Open dsn
            SQL1 = "select *,len(kode)'lebar' from am_nomax"
            Set RST1 = OBJ1.Execute(SQL1)
            Do While Not RST1.EOF
                If (RST1!lebar = 4 And Right(Trim(txtpo), 4) = RST1!kode) Or _
                (RST1!lebar = 5 And Right(Trim(txtpo), 5) = RST1!kode) Or _
                (RST1!lebar = 6 And Right(Trim(txtpo), 6) = RST1!kode) Then
                    bo1 = False
                    Exit Do
                End If
                RST1.MoveNext
            Loop
            OBJ1.Close
            
            If bo1 Then
                OBJ1.Open dsn
                SQL1 = "select qty from am_polin where nopo = '" & txtpo & "' and kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then
                    txtnil1 = RST1!qty
                Else
                    txtnil1 = 0
                End If
                
                SQL1 = "select isnull(sum(a.qty),0)'qty' from am_belilin a left join am_belihdr b on a.nobeli=b.nobeli where b.nopo = '" & txtpo & "' and kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
                Set RST1 = OBJ1.Execute(SQL1)
                If Not RST1.EOF Then
                    txtnil2 = RST1!qty
                Else
                    txtnil2 = 0
                End If
                OBJ1.Close
                
                If txtnil1 - txtnil2 = 0 Then
                    MsgBox "Purchase Order required is complete", vbExclamation, "Information"
                    
                    grid.TextMatrix(grid.Row, 1) = ""
                    grid.TextMatrix(grid.Row, 4) = ""
                    grid.TextMatrix(grid.Row, 7) = ""
                
                    Exit Sub
                End If
            
                grid.TextMatrix(grid.Row, 9) = Format(txtnil1 - txtnil2, "###,##0.00")
            Else
                grid.TextMatrix(grid.Row, 9) = "0.00"
            End If
            
            OBJ.Open dsn
            SQL = "select a.namabarang,a.kodesatuanmutasi,b.namasatuan,(c.namasatuan)'satmutasi' from am_apitemmst a left join am_apunit b on a.kodesatuan=b.kodesatuan left join am_apunit c on a.kodesatuanmutasi=c.kodesatuan where a.kodebarang = '" & grid.TextMatrix(grid.Row, 1) & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                grid.TextMatrix(grid.Row, 2) = RST!namabarang
                grid.TextMatrix(grid.Row, 3) = "0.00"
                grid.TextMatrix(grid.Row, 5) = RST!namasatuan
                grid.TextMatrix(grid.Row, 6) = "0.00"
                grid.TextMatrix(grid.Row, 7) = RST!kodesatuanmutasi
                grid.TextMatrix(grid.Row, 8) = RST!satmutasi
                
                SetRow grid.Row, True
                If grid.Row = (grid.Rows - 1) Then grid.Rows = grid.Rows + 1
                grid.SetFocus
                grid.Col = 2
            Else
                MsgBox "Item Not Found", vbExclamation, "Warning"
                
                grid.TextMatrix(grid.Row, 1) = ""
                grid.TextMatrix(grid.Row, 2) = ""
                grid.TextMatrix(grid.Row, 3) = ""
                grid.TextMatrix(grid.Row, 4) = ""
                grid.TextMatrix(grid.Row, 5) = ""
                grid.TextMatrix(grid.Row, 6) = ""
                grid.TextMatrix(grid.Row, 7) = ""
                grid.TextMatrix(grid.Row, 8) = ""
                grid.TextMatrix(grid.Row, 9) = ""
            End If
            OBJ.Close
    End Select
End Sub

Private Sub grid_Scroll()
    txtnilai.Visible = False
End Sub

Private Sub timer1_Timer()
    timer1.Enabled = False
    History
End Sub

Private Sub txtdriver_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then grid.SetFocus
End Sub

Private Sub txtkend_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtdriver.SetFocus
End Sub

Private Sub txtnilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        bo1 = True
        
        OBJ1.Open dsn
        SQL1 = "select *,len(kode)'lebar' from am_nomax"
        Set RST1 = OBJ1.Execute(SQL1)
        Do While Not RST1.EOF
            If (RST1!lebar = 4 And Right(Trim(txtpo), 4) = RST1!kode) Or _
            (RST1!lebar = 5 And Right(Trim(txtpo), 5) = RST1!kode) Or _
            (RST1!lebar = 6 And Right(Trim(txtpo), 6) = RST1!kode) Then
                bo1 = False
                Exit Do
            End If
            RST1.MoveNext
        Loop
        OBJ1.Close
        
        If bo1 Then
            If grid.Col = 3 Then
                If Val(Format(txtnilai, "general number")) > Val(Format(grid.TextMatrix(grid.Row, 9), "general number")) Then
                    MsgBox "Purchase Order required, Qty max = " & grid.TextMatrix(grid.Row, 9), vbExclamation, "Information"
                Else
                    grid.TextMatrix(grid.Row, grid.Col) = Format(txtnilai, "###,###,##0.00")
                End If
            Else
                grid.TextMatrix(grid.Row, grid.Col) = Format(txtnilai, "###,###,##0.00")
            End If
        Else
            grid.TextMatrix(grid.Row, grid.Col) = Format(txtnilai, "###,###,##0.00")
        End If
        
        txtnilai = 0
        txtnilai.Visible = False
        grid.SetFocus
        grid.Row = posrow
    ElseIf KeyAscii = 27 Then
        txtnilai = 0
        txtnilai_LostFocus
    End If
End Sub

Private Sub txtnilai_LostFocus()
    txtnilai.Visible = False
End Sub

Private Sub txtnobukti_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then date1.SetFocus
End Sub

Private Sub txtnobukti_LostFocus()
    carinvoice
End Sub

Private Sub txtpo_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtpo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtsj.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtsj_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtkend.SetFocus
End Sub

Function tanggalpo()
    tanggalpo = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Private Sub hapusgrid()
    grid.Row = 1
    Do While True
        If grid.TextMatrix(grid.Row, 1) = "" Then Exit Do
        grid.TextMatrix(grid.Row, 1) = ""
        grid.TextMatrix(grid.Row, 2) = ""
        grid.TextMatrix(grid.Row, 3) = ""
        grid.TextMatrix(grid.Row, 4) = ""
        grid.TextMatrix(grid.Row, 5) = ""
        grid.TextMatrix(grid.Row, 6) = ""
        grid.TextMatrix(grid.Row, 7) = ""
        grid.TextMatrix(grid.Row, 8) = ""
        grid.TextMatrix(grid.Row, 9) = ""
        grid.Col = 0
        Set grid.CellPicture = blank
        grid.Row = grid.Row + 1
    Loop
    
    grid.Rows = 2
    grid.ColWidth(0) = 250
    grid.ColWidth(1) = 1200
    grid.ColWidth(2) = 2500
    grid.ColWidth(3) = 1000
    grid.ColWidth(4) = 800
    grid.ColWidth(5) = 1000
    grid.ColWidth(6) = 1000
    grid.ColWidth(7) = 800
    grid.ColWidth(8) = 1000
    grid.ColWidth(9) = 0
End Sub

Private Sub hapusrow()
    grid.TextMatrix(grid.Row, 1) = ""
    grid.TextMatrix(grid.Row, 2) = ""
    grid.TextMatrix(grid.Row, 3) = ""
    grid.TextMatrix(grid.Row, 4) = ""
    grid.TextMatrix(grid.Row, 5) = ""
    grid.TextMatrix(grid.Row, 6) = ""
    grid.TextMatrix(grid.Row, 7) = ""
    grid.TextMatrix(grid.Row, 8) = ""
    grid.TextMatrix(grid.Row, 9) = ""
    Do While True
        If grid.TextMatrix(grid.Row + 1, 1) = "" Then
            grid.TextMatrix(grid.Row, 1) = ""
            grid.TextMatrix(grid.Row, 2) = ""
            grid.TextMatrix(grid.Row, 3) = ""
            grid.TextMatrix(grid.Row, 4) = ""
            grid.TextMatrix(grid.Row, 5) = ""
            grid.TextMatrix(grid.Row, 6) = ""
            grid.TextMatrix(grid.Row, 7) = ""
            grid.TextMatrix(grid.Row, 8) = ""
            grid.TextMatrix(grid.Row, 9) = ""
            Exit Do
        End If
        grid.TextMatrix(grid.Row, 1) = grid.TextMatrix(grid.Row + 1, 1)
        grid.TextMatrix(grid.Row, 2) = grid.TextMatrix(grid.Row + 1, 2)
        grid.TextMatrix(grid.Row, 3) = grid.TextMatrix(grid.Row + 1, 3)
        grid.TextMatrix(grid.Row, 4) = grid.TextMatrix(grid.Row + 1, 4)
        grid.TextMatrix(grid.Row, 5) = grid.TextMatrix(grid.Row + 1, 5)
        grid.TextMatrix(grid.Row, 6) = grid.TextMatrix(grid.Row + 1, 6)
        grid.TextMatrix(grid.Row, 7) = grid.TextMatrix(grid.Row + 1, 7)
        grid.TextMatrix(grid.Row, 8) = grid.TextMatrix(grid.Row + 1, 8)
        grid.TextMatrix(grid.Row, 9) = grid.TextMatrix(grid.Row + 1, 9)
        grid.Row = grid.Row + 1
    Loop
    grid.Rows = grid.Rows - 1
    grid.Col = 0
    Set grid.CellPicture = blank
End Sub

Private Sub SetRow(ByVal idx As Integer, ByVal hapus As String)
    grid.Row = idx
    grid.Col = 0
    If hapus Then Set grid.CellPicture = uncheck.Picture
    grid.Col = 1
End Sub

Private Sub carinvoice()
    If txtnobukti = "" Then Exit Sub
    If txtnobukti.SelLength <> 0 Then Exit Sub
    
    hapusgrid
    
    If Date > date1.MaxDate Then
        date1 = date1.MaxDate
    ElseIf Date < date1.MinDate Then
        date1 = date1.MinDate
    Else
        date1.Value = Date
    End If
    txtpo = ""
    txtsj = ""
    txtkend = ""
    txtdriver = ""

    OBJ.Open dsn
    SQL = "select * from am_belihdr where nobeli = '" & txtnobukti & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Data already exist.", vbInformation, "Information"
        txtnobukti = ""
        txtnobukti.SetFocus
    End If
    OBJ.Close
End Sub

Function batas1()
    batas1 = Month(v_fstgl1) & "/" & Day(v_fstgl1) & "/" & Year(v_fstgl1)
End Function

Function batas2()
    batas2 = Month(v_fstgl2) & "/" & Day(v_fstgl2) & "/" & Year(v_fstgl2)
End Function

Private Sub History()
    Dim s_divisi As String
    Dim s_format As String
    Dim s_format2 As String
    Dim s_kdbpb As String
    
    s_divisi = Mid(txtpo, 11, 6)
    s_format = "YY.MM."
    s_format2 = "/"
    
    OBJ1.Open dsn
    SQL1 = "select top 1 nobeli from am_belihdr where nobeli like '" & Format(date1, s_format) & "%" & s_divisi & "' order by nobeli desc"
    
    Set RST1 = OBJ1.Execute(SQL1)
    If Not RST1.EOF Then
        s_kdbpb = Mid(RST1!NoBeli, Len(s_format) + 1, 3)
    Else
        s_kdbpb = 0
    End If
    
    s_kdbpb = s_kdbpb + 1
    
    If Len(s_kdbpb) = 1 Then txtnobukti = Format(date1, s_format) & "00" & s_kdbpb & s_format2 & s_divisi
    If Len(s_kdbpb) = 2 Then txtnobukti = Format(date1, s_format) & "0" & s_kdbpb & s_format2 & s_divisi
    If Len(s_kdbpb) = 3 Then txtnobukti = Format(date1, s_format) & s_kdbpb & s_format2 & s_divisi
        
    OBJ1.Close
End Sub

Private Sub Cetak_Bukti()
    With rptbpb
         SQL1 = "Exec am_printbpb '" & txtnobukti & "'"
        .DataControl1.Source = SQL1
        .DataControl1.ConnectionString = dsn
        .Show vbModal
    End With
End Sub

