VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmdaftarjual 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan Penjualan"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3930
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
   Icon            =   "frmdaftarjual.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   3930
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      Caption         =   "Enabled Date"
      Height          =   420
      Left            =   3030
      TabIndex        =   31
      Top             =   4185
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.TextBox txtinv3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   15
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox txtinv4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   16
      Top             =   3720
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1200
      TabIndex        =   17
      Top             =   4080
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
      Format          =   135069699
      CurrentDate     =   37464
   End
   Begin Crystal.CrystalReport crystal 
      Left            =   480
      Top             =   4920
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
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   1200
      TabIndex        =   18
      Top             =   4440
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
      Format          =   135069699
      CurrentDate     =   37464
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   2760
      TabIndex        =   20
      Top             =   4920
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
      MICON           =   "frmdaftarjual.frx":2372
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
      Left            =   1800
      TabIndex        =   19
      Top             =   4920
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
      MICON           =   "frmdaftarjual.frx":268C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch3 
      Height          =   285
      Left            =   240
      TabIndex        =   25
      Top             =   3360
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "From"
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmdaftarjual.frx":29A6
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch4 
      Height          =   285
      Left            =   240
      TabIndex        =   26
      Top             =   3720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "To"
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmdaftarjual.frx":2CC0
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   360
      TabIndex        =   24
      Top             =   120
      Width           =   3255
      Begin VB.OptionButton Option15 
         BackColor       =   &H00FFFFFF&
         Caption         =   "(monthly) per Sales"
         Height          =   255
         Left            =   0
         TabIndex        =   30
         Top             =   1920
         Width           =   2415
      End
      Begin VB.OptionButton Option14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "per Customer (WP)"
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton Option13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "(monthly) per Area Customer"
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   1680
         Width           =   2415
      End
      Begin VB.OptionButton Option12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "(monthly) per Customer"
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   29
         Top             =   0
         Width           =   1095
         Begin VB.OptionButton Option8 
            BackColor       =   &H00FFFFFF&
            Caption         =   "5 Digit"
            Height          =   255
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   975
         End
         Begin VB.OptionButton Option9 
            BackColor       =   &H00FFFFFF&
            Caption         =   "8 Digit"
            Height          =   255
            Left            =   0
            TabIndex        =   9
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   28
         Top             =   720
         Width           =   1095
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Detail"
            Height          =   255
            Left            =   0
            TabIndex        =   11
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Summary"
            Height          =   255
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.OptionButton Option11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kilogram per Area"
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   2520
         Width           =   1695
      End
      Begin VB.OptionButton Option10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kilogram per Barang"
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   2280
         Width           =   1815
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "per Faktur"
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "per Area Customer"
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "per Salesman"
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "per Customer"
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "per Barang"
         Height          =   255
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Value           =   -1  'True
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker date3 
         Height          =   285
         Left            =   960
         TabIndex        =   14
         Top             =   2790
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMMM yyyy"
         Format          =   135069699
         CurrentDate     =   37464
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Periode"
         ForeColor       =   &H80000011&
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   2820
         Width           =   615
      End
   End
   Begin VB.Label Label1 
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
      Height          =   3255
      Left            =   -360
      TabIndex        =   23
      Top             =   0
      Width           =   6135
   End
   Begin VB.Label Label4 
      Caption         =   "To Date"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   4470
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "From Date"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   4110
      Width           =   855
   End
End
Attribute VB_Name = "frmdaftarjual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim str1, str2 As String

Function batas1()
    batas1 = Format(v_fstgl1, "MM/dd/yyyy")
End Function

Function batas2()
    batas2 = Format(v_fstgl2, "MM/dd/yyyy")
End Function

Private Sub Check1_Click()
    If Check1.Value = Unchecked Then
        date1.Enabled = False
        date2.Enabled = False
        date3.Enabled = True
    ElseIf Check1.Value = Checked Then
        date1.Enabled = True
        date2.Enabled = True
        date3.Enabled = False
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdsearch3_Click()
    If Option2.Value = True Or Option12.Value = True Then
        carisql1 = "select kodecust, namacust, alamatcust, kota from am_customer"
        namatabel = "Customer"
    ElseIf Option1.Value = True Or Option10.Value = True Then
        If Option8.Value = True Then carisql1 = "select kodebarang, namabarang from am_itemmst where len(kodebarang)=5"
        If Option9.Value = True Then carisql1 = "select kodebarang, namabarang from am_itemmst where len(kodebarang)=8"
        namatabel = "Barang "
    ElseIf Option5.Value = True Or Option15.Value = True Then
        carisql1 = "select kodesales, namasales,idupdate from AM_salesman"
        namatabel = "Sales"
    ElseIf Option6.Value = True Or Option11.Value = True Or Option13.Value = True Then
        carisql1 = "select kode, nama from am_area"
        namatabel = "Area"
    ElseIf Option7.Value = True Then
        If v_fastsearching = True Then
            If v_fstgl1 > v_fstgl2 Then
                MsgBox "Invalid date range, search abort.", vbExclamation, "Error"
                Exit Sub
            End If
            carisql1 = "select nobkt, convert(char(11),tglbkt )'tglbkt', type from am_invhdr where type = 'I' and tglbkt >= '" & batas1 & "' and tglbkt <= '" & batas2 & "'"
        Else
            'Diberi batas dari tgl 01-01-2015 biar tidak lambat
            carisql1 = "select nobkt, convert(char(11),tglbkt )'tglbkt', type from am_invhdr where type = 'I' and TglBkt >='2015-01-01'"
        End If
        namatabel = "Faktur Penjualan"
    End If
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch4_Click()
    If Option2.Value = True Or Option12.Value = True Then
        carisql1 = "select kodecust, namacust, alamatcust, kota from am_customer"
        namatabel = "Customer"
    ElseIf Option1.Value = True Or Option10.Value = True Then
        If Option8.Value = True Then carisql1 = "select kodebarang, namabarang from am_itemmst where len(kodebarang)=5"
        If Option9.Value = True Then carisql1 = "select kodebarang, namabarang from am_itemmst where len(kodebarang)=8"
        namatabel = "Barang "
    ElseIf Option5.Value = True Or Option15.Value = True Then
        carisql1 = "select kodesales, namasales, idupdate from AM_salesman"
        namatabel = "Sales"
    ElseIf Option6.Value = True Or Option13.Value = True Or Option11.Value = True Then
        carisql1 = "select kode, nama from am_area"
        namatabel = "Area"
    ElseIf Option7.Value = True Then
        If v_fastsearching = True Then
            If v_fstgl1 > v_fstgl2 Then
                MsgBox "Invalid date range, search abort.", vbExclamation, "Error"
                Exit Sub
            End If
            carisql1 = "select nobkt, convert(char(11),tglbkt )'tglbkt', type from am_invhdr where type = 'I' and tglbkt >= '" & batas1 & "' and tglbkt <= '" & batas2 & "'"
        Else
            carisql1 = "select nobkt, convert(char(11),tglbkt )'tglbkt', type from am_invhdr where type = 'I'"
        End If
        namatabel = "Faktur Penjualan"
    End If
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch3_GotFocus()
    If hasil = "" Then Exit Sub
    txtinv3 = hasil
    If Option5.Value = True Then carisales
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdsearch4_GotFocus()
    If hasil = "" Then Exit Sub
    txtinv4 = hasil
    If Option5.Value = True Then carisales
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub date3_Change()
    date1 = date3
    date2 = date3
    
    date1.Day = 1
    date2.Day = 25
    date2 = date2 + 10
    date2.Day = 1
    date2 = date2 - 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Option1_Click()
    txtinv3 = ""
    txtinv4 = ""
    date1 = Date
    date2 = Date
    date3 = Date
    date3.Enabled = False
    date2.Enabled = True
    date1.Enabled = True
    txtinv3.Enabled = True
    txtinv4.Enabled = True
    cmdsearch3.Enabled = True
    cmdsearch4.Enabled = True
    Label3.ForeColor = &H80000011
    Check1.Visible = False
    
    enebelin
End Sub

Private Sub Option10_Click()
    txtinv3 = ""
    txtinv4 = ""
    date1 = Date
    date2 = Date
    date3 = Date
    date3.Enabled = True
    date2.Enabled = False
    date1.Enabled = False
    txtinv3.Enabled = True
    txtinv4.Enabled = True
    cmdsearch3.Enabled = True
    cmdsearch4.Enabled = True
    Label3.ForeColor = &H80000012
    Check1.Visible = True
    Check1.Value = Unchecked
    
    date1.Day = 1
    date2.Day = 25
    date2 = date2 + 10
    date2.Day = 1
    date2 = date2 - 1
    
    Option9.Value = True
    Option8.Enabled = False
    Option9.Enabled = False
    
    Option4.Value = True
    Option4.Enabled = False
    Option3.Enabled = False
End Sub

Private Sub Option11_Click()
    txtinv3 = ""
    txtinv4 = ""
    date1 = Date
    date2 = Date
    date3 = Date
    date3.Enabled = True
    date2.Enabled = False
    date1.Enabled = False
    txtinv3.Enabled = True
    txtinv4.Enabled = True
    cmdsearch3.Enabled = True
    cmdsearch4.Enabled = True
    Label3.ForeColor = &H80000012
    Check1.Visible = True
    Check1.Value = Unchecked
    
    date1.Day = 1
    date2.Day = 25
    date2 = date2 + 10
    date2.Day = 1
    date2 = date2 - 1
    
    Option9.Value = True
    Option8.Enabled = False
    Option9.Enabled = False
    
    Option4.Value = True
    Option4.Enabled = False
    Option3.Enabled = False
End Sub

Private Sub Option12_Click()
    txtinv3 = ""
    txtinv4 = ""
    date1 = Date
    date2 = Date
    date3 = Date
    date3.Enabled = False
    date2.Enabled = False
    date1.Enabled = False
    txtinv3.Enabled = True
    txtinv4.Enabled = True
    cmdsearch3.Enabled = True
    cmdsearch4.Enabled = True
    Label3.ForeColor = &H80000011
    Check1.Visible = False
    
    Option9.Value = True
    Option8.Enabled = False
    Option9.Enabled = False
    
    Option4.Value = True
    Option4.Enabled = False
    Option3.Enabled = False
    
    date1.Month = 1
    date1.Day = 1
    date2.Month = 12
    date2.Day = 31
End Sub

Private Sub Option13_Click()
    txtinv3 = ""
    txtinv4 = ""
    date1 = Date
    date2 = Date
    date3 = Date
    date3.Enabled = False
    date2.Enabled = False
    date1.Enabled = False
    txtinv3.Enabled = True
    txtinv4.Enabled = True
    cmdsearch3.Enabled = True
    cmdsearch4.Enabled = True
    Label3.ForeColor = &H80000011
    Check1.Visible = False
    
    Option9.Value = True
    Option8.Enabled = False
    Option9.Enabled = False
    
    Option4.Value = True
    Option4.Enabled = False
    Option3.Enabled = False
    
    date1.Month = 1
    date1.Day = 1
    date2.Month = 12
    date2.Day = 31
End Sub

Private Sub Option14_Click()
    txtinv3 = ""
    txtinv4 = ""
    date1 = Date
    date2 = Date
    date3 = Date
    date3.Enabled = False
    date2.Enabled = True
    date1.Enabled = True
    txtinv3.Enabled = False
    txtinv4.Enabled = False
    cmdsearch3.Enabled = False
    cmdsearch4.Enabled = False
    Label3.ForeColor = &H80000011
    Check1.Visible = False
    
    enebelin
End Sub

Private Sub Option15_Click()
    txtinv3 = ""
    txtinv4 = ""
    date1 = Date
    date2 = Date
    date3 = Date
    date3.Enabled = False
    date2.Enabled = False
    date1.Enabled = False
    txtinv3.Enabled = True
    txtinv4.Enabled = True
    cmdsearch3.Enabled = True
    cmdsearch4.Enabled = True
    Label3.ForeColor = &H80000011
    Check1.Visible = False
    
    Option9.Value = True
    Option8.Enabled = False
    Option9.Enabled = False
    
    Option4.Value = True
    Option4.Enabled = False
    Option3.Enabled = False
    
    date1.Month = 1
    date1.Day = 1
    date2.Month = 12
    date2.Day = 31
End Sub

Private Sub Option2_Click()
    txtinv3 = ""
    txtinv4 = ""
    date1 = Date
    date2 = Date
    date3 = Date
    date3.Enabled = False
    date2.Enabled = True
    date1.Enabled = True
    txtinv3.Enabled = True
    txtinv4.Enabled = True
    cmdsearch3.Enabled = True
    cmdsearch4.Enabled = True
    Label3.ForeColor = &H80000011
    Check1.Visible = False
    
    enebelin
End Sub

Private Sub Option5_Click()
    txtinv3 = ""
    txtinv4 = ""
    date1 = Date
    date2 = Date
    date3 = Date
    date3.Enabled = False
    date2.Enabled = True
    date1.Enabled = True
    txtinv3.Enabled = True
    txtinv4.Enabled = True
    cmdsearch3.Enabled = True
    cmdsearch4.Enabled = True
    Label3.ForeColor = &H80000011
    Check1.Visible = False
    
    enebelin
End Sub

Private Sub Option6_Click()
    txtinv3 = ""
    txtinv4 = ""
    date1 = Date
    date2 = Date
    date3 = Date
    date3.Enabled = False
    date2.Enabled = True
    date1.Enabled = True
    txtinv3.Enabled = True
    txtinv4.Enabled = True
    cmdsearch3.Enabled = True
    cmdsearch4.Enabled = True
    Label3.ForeColor = &H80000011
    Check1.Visible = False
    
    enebelin
End Sub

Private Sub Option7_Click()
    txtinv3 = ""
    txtinv4 = ""
    date1 = Date
    date2 = Date
    date3 = Date
    date3.Enabled = False
    date2.Enabled = True
    date1.Enabled = True
    txtinv3.Enabled = True
    txtinv4.Enabled = True
    cmdsearch3.Enabled = True
    cmdsearch4.Enabled = True
    Label3.ForeColor = &H80000011
    Check1.Visible = False
    
    enebelin
End Sub

Private Sub txtinv3_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtinv4.SetFocus
End Sub

Private Sub txtinv4_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 And date1.Enabled Then date1.SetFocus Else cmdview.SetFocus
End Sub

Private Sub txtinv3_LostFocus()
    cariinv3
End Sub

Private Sub txtinv4_LostFocus()
    cariinv4
End Sub

Private Sub cmdview_Click()
    If (txtinv3 = "" Or txtinv4 = "") And Option14.Value = False Then Exit Sub
    If txtinv4 < txtinv3 Then
        MsgBox "To statement Can Not Smaller Then From statement.", vbOKOnly + vbExclamation, "Warning"
        txtinv4 = ""
        txtinv4.SetFocus
        Exit Sub
    End If
    
    If date2 < date1 Then
        MsgBox "To Date Can Not Smaller Then From Date.", vbExclamation, "Warning"
        date2.SetFocus
        Exit Sub
    End If
    
    If Option1.Value = True Then str1 = "barang"
    If Option2.Value = True Then str1 = "customer"
    If Option5.Value = True Then str1 = "sales"
    If Option6.Value = True Then str1 = "area"
    If Option7.Value = True Then str1 = "bukti"
    If Option10.Value = True Then str1 = "kilo"
    If Option11.Value = True Then str1 = "kiloarea"
    If Option12.Value = True Then str1 = "mcust"
    If Option13.Value = True Then str1 = "marea"
    If Option14.Value = True Then str1 = "wp"
    If Option15.Value = True Then str1 = "msales"
    
    If Option8.Value = True Then str2 = "lima"
    If Option9.Value = True Then str2 = "dela"
        
    crystal.Reset
    crystal.WindowState = crptMaximized
    crystal.WindowShowCloseBtn = True
    crystal.WindowShowPrintSetupBtn = True
    crystal.WindowShowSearchBtn = True
    crystal.WindowShowRefreshBtn = True
    
    crystal.Connect = dsnreport
    If Option10.Value = True Then
        crystal.DataFiles(0) = "Proc(am_daftarjualkilo)"
    ElseIf Option11.Value = True Then
        crystal.DataFiles(0) = "Proc(am_daftarjualkiloarea)"
    ElseIf Option12.Value = True Or Option13.Value = True Then
        crystal.DataFiles(0) = "Proc(am_daftarjualbulan)"
    ElseIf Option14.Value = True Then
        crystal.DataFiles(0) = "Proc(am_daftarjualwp)"
    ElseIf Option15.Value = True Then
        crystal.DataFiles(0) = "Proc(am_daftarjualsales)"
    ElseIf Option5.Value = True Then
        If Option3.Value = True Then
        crystal.DataFiles(0) = "Proc(am_daftarjual)"
        End If
    Else
        crystal.DataFiles(0) = "Proc(am_daftarjual)"
    End If
    If Option1.Value = True Then
        If Option3.Value = True Then crystal.ReportFileName = AppPath & "\reports\sale\inv\jbarang.rpt"
        If Option4.Value = True Then crystal.ReportFileName = AppPath & "\reports\sale\inv\jbarang_sum.rpt"
    ElseIf Option2.Value = True Then
        If Option3.Value = True Then crystal.ReportFileName = AppPath & "\reports\sale\inv\jcustomer.rpt"
        If Option4.Value = True Then crystal.ReportFileName = AppPath & "\reports\sale\inv\jcustomer_sum.rpt"
    ElseIf Option5.Value = True Then
        If Option3.Value = True Then crystal.ReportFileName = AppPath & "\reports\sale\inv\jsales1.rpt"
        If Option4.Value = True Then crystal.ReportFileName = AppPath & "\reports\sale\inv\jsales_sum.rpt"
    ElseIf Option6.Value = True Then
        If Option3.Value = True Then crystal.ReportFileName = AppPath & "\reports\sale\inv\jarea.rpt"
        If Option4.Value = True Then crystal.ReportFileName = AppPath & "\reports\sale\inv\jarea_sum.rpt"
    ElseIf Option7.Value = True Then
        If Option3.Value = True Then crystal.ReportFileName = AppPath & "\reports\sale\inv\jbukti.rpt"
        If Option4.Value = True Then crystal.ReportFileName = AppPath & "\reports\sale\inv\jbukti_sum.rpt"
    ElseIf Option10.Value = True Then
        crystal.ReportFileName = AppPath & "\reports\sale\inv\jkilo.rpt"
    ElseIf Option11.Value = True Then
        crystal.ReportFileName = AppPath & "\reports\sale\inv\jkiloarea.rpt"
    ElseIf Option12.Value = True Then
        crystal.ReportFileName = AppPath & "\reports\sale\inv\jcustomerbulan.rpt"
    ElseIf Option13.Value = True Then
        crystal.ReportFileName = AppPath & "\reports\sale\inv\jareabulan.rpt"
    ElseIf Option14.Value = True Then
        crystal.ReportFileName = AppPath & "\reports\sale\inv\jcustomerwp.rpt"
    ElseIf Option15.Value = True Then
        crystal.ReportFileName = AppPath & "\reports\sale\inv\jsalesbulan.rpt"
    End If
    If Option12.Value = True Or Option13.Value = True Then
        crystal.ParameterFields(0) = "@tanggal1;" & Format(date1, "yyyyMMdd") & ";true"
        crystal.ParameterFields(1) = "@tanggal2;" & Format(date2, "yyyyMMdd") & ";true"
        crystal.ParameterFields(2) = "@batas1 ;" + txtinv3 + ";true"
        crystal.ParameterFields(3) = "@batas2 ;" + txtinv4 + ";true"
        crystal.ParameterFields(4) = "@pilih ;" + str1 + ";true"
        crystal.ParameterFields(5) = "@namauser ;" + nmuser + ";true"
    ElseIf Option5.Value = True Then
        If Option3.Value = True Then
            crystal.ParameterFields(0) = "@tanggal1;" & Format(date1, "yyyyMMdd") & ";true"
            crystal.ParameterFields(1) = "@tanggal2;" & Format(date2, "yyyyMMdd") & ";true"
            crystal.ParameterFields(2) = "@batas1 ;" + txtinv3 + ";true"
            crystal.ParameterFields(3) = "@batas2 ;" + txtinv4 + ";true"
            crystal.ParameterFields(4) = "@pilih ;" + str1 + ";true"
            crystal.ParameterFields(5) = "@pilih2 ;" + str2 + ";true"
            crystal.ParameterFields(6) = "@namauser ;" + nmuser + ";true"
        Else
            crystal.ParameterFields(0) = "@tanggal1;" & Format(date1, "yyyyMMdd") & ";true"
            crystal.ParameterFields(1) = "@tanggal2;" & Format(date2, "yyyyMMdd") & ";true"
            crystal.ParameterFields(2) = "@batas1 ;" + txtinv3 + ";true"
            crystal.ParameterFields(3) = "@batas2 ;" + txtinv4 + ";true"
            crystal.ParameterFields(4) = "@pilih ;" + str1 + ";true"
            crystal.ParameterFields(5) = "@pilih2 ;" + str2 + ";true"
            crystal.ParameterFields(6) = "@namauser ;" + nmuser + ";true"
        End If
    Else
        crystal.ParameterFields(0) = "@tanggal1;" & Format(date1, "yyyyMMdd") & ";true"
        crystal.ParameterFields(1) = "@tanggal2;" & Format(date2, "yyyyMMdd") & ";true"
        crystal.ParameterFields(2) = "@batas1 ;" + txtinv3 + ";true"
        crystal.ParameterFields(3) = "@batas2 ;" + txtinv4 + ";true"
        crystal.ParameterFields(4) = "@pilih ;" + str1 + ";true"
        crystal.ParameterFields(5) = "@pilih2 ;" + str2 + ";true"
        crystal.ParameterFields(6) = "@namauser ;" + nmuser + ";true"
    End If
    crystal.RetrieveDataFiles
    crystal.Action = 1
End Sub

Private Sub Form_Load()
    date1 = Date
    date2 = Date
    date3 = Date
End Sub

Private Sub cariinv3()
    If txtinv3 = "" Then Exit Sub
    
    OBJ.Open dsn
    If Option2.Value = True Or Option12.Value = True Then
        SQL = "select * from am_customer where kodecust = '" & txtinv3 & "'"
    ElseIf Option1.Value = True Or Option10.Value = True Then
        SQL = "select * from am_itemmst where kodebarang = '" & txtinv3 & "'"
    ElseIf Option5.Value = True Then
        SQL = "select * from am_salesman where kodesales = '" & txtinv3 & "'"
    ElseIf Option6.Value = True Or Option11.Value = True Or Option13.Value = True Then
        SQL = "select * from am_area where kode = '" & txtinv3 & "'"
    ElseIf Option7.Value = True Then
        SQL = "select nobkt from am_invhdr where nobkt = '" & txtinv3 & "'"
    End If
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Data not found.", vbExclamation, "Warning"
        txtinv3 = ""
        txtinv3.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub cariinv4()
    If txtinv4 = "" Then Exit Sub
    
    OBJ.Open dsn
    If Option2.Value = True Or Option12.Value = True Then
        SQL = "select * from am_customer where kodecust = '" & txtinv4 & "'"
    ElseIf Option1.Value = True Or Option10.Value = True Then
        SQL = "select * from am_itemmst where kodebarang = '" & txtinv4 & "'"
    ElseIf Option5.Value = True Then
        SQL = "select * from am_salesman where kodesales = '" & txtinv4 & "'"
    ElseIf Option6.Value = True Or Option11.Value = True Or Option13.Value = True Then
        SQL = "select * from am_area where kode = '" & txtinv4 & "'"
    ElseIf Option7.Value = True Then
        SQL = "select nobkt from am_invhdr where nobkt = '" & txtinv4 & "'"
    End If
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Data not found.", vbExclamation, "Warning"
        txtinv4 = ""
        txtinv4.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub enebelin()
    Option8.Enabled = True
    Option9.Enabled = True
    Option9.Value = True
    
    Option4.Enabled = True
    Option3.Enabled = True
    Option4.Value = True
End Sub
Private Sub carisales()
    If txtinv3 = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "SELECT * FROM AM_Salesman WHERE KodeSales = '" & txtinv3 & "'"
    Set RST = OBJ.Execute(SQL)
'-------------------- 0 = sales non aktif -------------------
    If RST!idupdate = "0" Then
        MsgBox "Salesman " & RST!namasales & " is not active !", vbExclamation, "Warning"
        txtinv3 = ""
    End If
    OBJ.Close
End Sub
