VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Begin VB.Form frmSPKL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Surat Perintah Kerja Lembur"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9645
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   9645
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtkode 
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
      Height          =   300
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   36
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtkabag 
      Alignment       =   2  'Center
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
      Height          =   300
      Left            =   7440
      MaxLength       =   15
      TabIndex        =   34
      Top             =   5880
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.TextBox txthrd 
      Alignment       =   2  'Center
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
      Height          =   300
      Left            =   5160
      MaxLength       =   15
      TabIndex        =   33
      Top             =   5930
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000A&
      FillStyle       =   0  'Solid
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2355
      ScaleWidth      =   2475
      TabIndex        =   28
      Top             =   3960
      Visible         =   0   'False
      Width           =   2535
      Begin VB.TextBox txtnamattd 
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
         Height          =   300
         Left            =   600
         MaxLength       =   50
         TabIndex        =   29
         Top             =   0
         Width           =   1875
      End
      Begin VB.TextBox txtid 
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
         Height          =   300
         Left            =   0
         MaxLength       =   3
         TabIndex        =   32
         Top             =   0
         Width           =   555
      End
      Begin Chameleon.chameleonButton cmdsavettd 
         Height          =   375
         Left            =   1560
         TabIndex        =   30
         Top             =   1920
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Save TTD"
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
         MICON           =   "frmSPKL.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdupload 
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Upload Image TTD"
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
         MICON           =   "frmSPKL.frx":031A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Image Imgttd 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Left            =   120
         Stretch         =   -1  'True
         Top             =   360
         Width           =   2235
      End
   End
   Begin VB.ComboBox cmbhari 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmSPKL.frx":0634
      Left            =   1920
      List            =   "frmSPKL.frx":0636
      TabIndex        =   4
      Text            =   "Minggu"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check1"
      Height          =   255
      Left            =   6960
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   11
      Top             =   4620
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check1"
      Height          =   255
      Left            =   9240
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   10
      Top             =   4620
      Width           =   255
   End
   Begin VB.ComboBox cmbstatus 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmSPKL.frx":0638
      Left            =   5760
      List            =   "frmSPKL.frx":063A
      TabIndex        =   6
      Text            =   "Hari Libur"
      Top             =   2520
      Width           =   1815
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
      Height          =   540
      Left            =   1920
      TabIndex        =   9
      Top             =   3360
      Width           =   7575
   End
   Begin VB.TextBox txtnik 
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
      Height          =   660
      Left            =   7080
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox txtunit 
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
      Height          =   300
      Left            =   7080
      MaxLength       =   15
      TabIndex        =   3
      Top             =   2160
      Width           =   2415
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
      Height          =   300
      Left            =   1920
      TabIndex        =   2
      Top             =   2160
      Width           =   4335
   End
   Begin VB.TextBox txtnama 
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
      Height          =   660
      Left            =   1920
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1440
      Width           =   4335
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   8640
      TabIndex        =   14
      Top             =   6600
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
      MICON           =   "frmSPKL.frx":063C
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
      Height          =   330
      Left            =   4080
      TabIndex        =   5
      Top             =   2520
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   582
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
      Format          =   134217729
      CurrentDate     =   41743
   End
   Begin MSComCtl2.DTPicker dtpjam 
      Height          =   330
      Left            =   1920
      TabIndex        =   7
      Top             =   3000
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   582
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
      Format          =   134217730
      CurrentDate     =   41743
   End
   Begin MSComCtl2.DTPicker dtpselesai 
      Height          =   330
      Left            =   4920
      TabIndex        =   8
      Top             =   3000
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   582
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
      Format          =   134217730
      CurrentDate     =   41743
   End
   Begin Chameleon.chameleonButton cmdsave 
      Height          =   375
      Left            =   7680
      TabIndex        =   12
      Top             =   6600
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
      MICON           =   "frmSPKL.frx":0956
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComDlg.CommonDialog cmndlg 
      Left            =   2760
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblid 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   8400
      TabIndex        =   37
      Top             =   4200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SPP/HRD-FRM/IV/008"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6840
      TabIndex        =   35
      Top             =   120
      Width           =   2595
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Disetujui Pimpinan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   2760
      TabIndex        =   27
      Top             =   4560
      Width           =   2235
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   2235
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mengetahui HRD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   5040
      TabIndex        =   26
      Top             =   4560
      Width           =   2235
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   2235
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kepala Bagian"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   7320
      TabIndex        =   25
      Top             =   4560
      Width           =   2235
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Left            =   7320
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   2235
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Nik"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6600
      TabIndex        =   24
      Top             =   1440
      Width           =   555
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6600
      TabIndex        =   23
      Top             =   2160
      Width           =   555
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis Pekerjaan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   22
      Top             =   3360
      Width           =   1995
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PT. SPARTA PRIMA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   195
      TabIndex        =   21
      Top             =   120
      Width           =   2595
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SURAT PERINTAH KERJA LEMBUR"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   90
      TabIndex        =   20
      Top             =   555
      Width           =   9675
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Lembur pada Hari                         Tanggal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   19
      Top             =   2520
      Width           =   6375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Di perintahkan kepada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   18
      Top             =   1140
      Width           =   2460
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "N a m a "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   17
      Top             =   1485
      Width           =   1275
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Bagian"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   16
      Top             =   2250
      Width           =   1275
   End
   Begin VB.Label lblterbilang 
      BackStyle       =   0  'Transparent
      Caption         =   "s/d selesai jam"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3480
      TabIndex        =   15
      Top             =   3000
      Width           =   1470
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Mulai jam"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   13
      Top             =   2940
      Width           =   1275
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000A&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   585
      Left            =   0
      Top             =   6495
      Width           =   10485
   End
End
Attribute VB_Name = "frmSPKL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private RS As ADODB.Recordset
Private OBJ As New ADODB.Connection
Private RSStream As ADODB.Stream
Private SQL As String
Private flname As String

Private Sub Check1_Click()
'cek user kabag
    If Check1.Value = Checked Then
        If UserOnline = "bina" Or UserOnline = "gyo" Then
            OBJ.Open dsn
            SQL = "Select * From list_ttd Where idttd = '003'"
            Set RS = OBJ.Execute(SQL)
            If RS.EOF Then
                OBJ.Close
                Exit Sub
            End If
            txtkabag.Visible = True
            txtkabag = RS!nama
            lblid = "003"
            flname = App.Path & "\temp\" & RS!idttd & ".jpg"
            If Not IsNull(RS!idttd) Then
            Set RSStream = New ADODB.Stream
                RSStream.Type = adTypeBinary
                RSStream.Open
                RSStream.Write RS!ttd
                RSStream.SaveToFile flname, adSaveCreateOverWrite
                RSStream.Close
                Set RSStream = Nothing
                Image1.Picture = LoadPicture(flname)
            End If
            OBJ.Close
        ElseIf UserOnline = "hariyanto" Then
            OBJ.Open dsn
            SQL = "Select * From list_ttd Where idttd = '004'"
            Set RS = OBJ.Execute(SQL)
            If RS.EOF Then
                OBJ.Close
                Exit Sub
            End If
            txtkabag.Visible = True
            txtkabag = RS!nama
            lblid = "004"
            flname = App.Path & "\temp\" & RS!idttd & ".jpg"
            If Not IsNull(RS!idttd) Then
            Set RSStream = New ADODB.Stream
                RSStream.Type = adTypeBinary
                RSStream.Open
                RSStream.Write RS!ttd
                RSStream.SaveToFile flname, adSaveCreateOverWrite
                RSStream.Close
                Set RSStream = Nothing
                Image1.Picture = LoadPicture(flname)
            End If
            OBJ.Close
        ElseIf UserOnline = "alibahari" Then
            OBJ.Open dsn
            SQL = "Select * From list_ttd Where idttd = '005'"
            Set RS = OBJ.Execute(SQL)
            If RS.EOF Then
                OBJ.Close
                Exit Sub
            End If
            txtkabag.Visible = True
            txtkabag = RS!nama
            lblid = "005"
            flname = App.Path & "\temp\" & RS!idttd & ".jpg"
            If Not IsNull(RS!idttd) Then
            Set RSStream = New ADODB.Stream
                RSStream.Type = adTypeBinary
                RSStream.Open
                RSStream.Write RS!ttd
                RSStream.SaveToFile flname, adSaveCreateOverWrite
                RSStream.Close
                Set RSStream = Nothing
                Image1.Picture = LoadPicture(flname)
            End If
            OBJ.Close
        End If
    Else
        flname = ""
        Image1.Picture = LoadPicture(flname)
        txtkabag.Visible = False
    End If
End Sub

Private Sub Check2_Click()
'cek user hrd
    If Check2.Value = Checked Then
        If UserOnline = "Wakidi" Or UserOnline = "Creator" Then
            OBJ.Open dsn
            SQL = "select * from list_ttd where idttd='002'"
            Set RS = OBJ.Execute(SQL)
            If RS.EOF Then
                OBJ.Close
                Exit Sub
            End If
            txthrd.Visible = True
            txthrd = RS!nama
            flname = App.Path & "\temp\" & RS!idttd & ".jpg"
            If Not IsNull(RS!ttd) Then
            Set RSStream = New ADODB.Stream
                RSStream.Type = adTypeBinary
                RSStream.Open
                RSStream.Write RS!ttd
                RSStream.SaveToFile flname, adSaveCreateOverWrite
                RSStream.Close
                Set RSStream = Nothing
                Image2.Picture = LoadPicture(flname)
            End If
            OBJ.Close
        Else
            
        End If
    ElseIf Check2.Value = Unchecked Then
        flname = ""
        Image2.Picture = LoadPicture(flname)
        txthrd.Visible = False
    End If
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdsave_Click()
'LIST_SPKL
    OBJ.Open dsn
    SQL = "Select * From list_spkl Where 0=1"
    Set RS = New ADODB.Recordset
    RS.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    
    RS.AddNew
    RS!idspkl = txtkode
    RS!nama = txtnama
    RS!nik = txtnik
    RS!bagian = txtbagian
    RS!unit = txtunit
    RS!hari = cmbhari.text
    RS!tgllembur = Format(date1, "yyyy/MM/dd")
    RS!statushari = cmbstatus
    RS!jamon = Format(dtpjam, "HH:mm:ss")
    RS!jamoff = Format(dtpselesai, "HH:mm:ss")
    RS!keterangan = txtket
    RS!pimpinan = "0"
    RS!hrd = "0"
    RS!kabag = lblid
    RS!Flag = "0"
    RS.Update
    
    OBJ.Close
    MsgBox "Berhasil disimpan", vbInformation, AppName
End Sub

Private Sub cmdsavettd_Click()
    OBJ.Open dsn
    SQL = "Select * From LIST_TTD Where idttd = '" & txtid & "'"
    Set RS = OBJ.Execute(SQL)
    If Not RS.EOF Then
        MsgBox "Maaf nomor id telah digunakan", vbCritical, AppName
        OBJ.Close
        Exit Sub
    End If
    
    SQL = "select * from LIST_TTD where 0=1"
    Set RS = New ADODB.Recordset
    RS.Open SQL, OBJ, adOpenDynamic, adLockOptimistic
    With RS
        .AddNew
        !idttd = txtid
        !nama = txtnamattd
        If flname <> "" Then
            Set RSStream = New ADODB.Stream
            RSStream.Type = adTypeBinary
            RSStream.Open
            RSStream.LoadFromFile flname
            !ttd = RSStream.Read
            RSStream.Close
        End If
        .Update
        End With
    OBJ.Close
    MsgBox "Data berhasil disimpan", vbInformation, AppName
    txtnamattd = ""
    txtid = GetNumber
    flname = ""
    Imgttd.Picture = LoadPicture(flname)
End Sub

Private Sub cmdupload_Click()
    On Error GoTo err_msg
    With cmndlg
        .CancelError = False
        .DialogTitle = "Gambar SOP"
        .Filter = "File Gambar (*.jpg)|*.jpg"
        .ShowOpen
        If .FileName <> "" Then
            flname = .FileName
            Imgttd.Picture = LoadPicture(flname)
        Else
            Imgttd.Picture = LoadPicture(flname)
            Exit Sub
        End If
    End With
    Exit Sub
err_msg:
End Sub

Private Sub Form_Load()
    cmbstatus.additem "Hari Biasa"
    cmbstatus.additem "Hari Libur"
    
    cmbhari.additem "Senin"
    cmbhari.additem "Selasa"
    cmbhari.additem "Rabu"
    cmbhari.additem "Kamis"
    cmbhari.additem "Jumat"
    cmbhari.additem "Sabtu"
    cmbhari.additem "Minggu"
    date1 = Date
    dtpjam = Format(Now, "HH:mm:ss")
    dtpselesai = Format(Now, "HH:mm:ss")
    txtid = GetNumber
    txtkode = getnospkl
    If UserOnline = "Creator" Then Picture1.Visible = True
End Sub

Private Function GetNumber() As String
    On Error GoTo Err_handler
    Dim tempnumber As Long
    Dim nobkt As String
    Dim lengthnumber As Integer
    
    OBJ.Open dsn
    SQL = "select max(idttd) nottd from LIST_TTD"
    Set RS = OBJ.Execute(SQL)
    tempnumber = CLng(RS!nottd) + 1
    lengthnumber = Len(Trim(Str(tempnumber)))

    txtid = GetNumber
    Select Case lengthnumber
        Case 1: nobkt = "00" + Trim(Str(tempnumber))
        Case 2: nobkt = "0" + Trim(Str(tempnumber))
        Case 3: nobkt = Trim(Str(tempnumber))
    End Select
    
    GetNumber = nobkt

    OBJ.Close
    Exit Function
Err_handler:
    GetNumber = "001"
    If OBJ.State = 1 Then OBJ.Close
End Function

Function getnospkl() As String    '23060001'
    On Error GoTo Err_handler:
    Dim SQL As String
    Dim strnumber As String
    Dim tempkode As String
    Dim kode As Long
    
    strnumber = Format(Date, "yymm")
    
    Set OBJ = New ADODB.Connection
    OBJ.Open dsn
    SQL = "select max(idspkl)as spkl from list_spkl where idspkl like '" + strnumber + "%'"
    Set RS = OBJ.Execute(SQL)

    If IsNull(RS!spkl) = True Or RS!spkl = "" Then
        getnospkl = strnumber + "0001"
    Else
        kode = CLng(Mid(RS!spkl, 5, 4)) + 1
        
        If (Len(Trim(Str(kode))) = 1) Then
            tempkode = strnumber + "000" + Trim(Str(kode))
        End If
        If (Len(Trim(Str(kode))) = 2) Then
            tempkode = strnumber + "00" + Trim(Str(kode))
        End If
        If (Len(Trim(Str(kode))) = 3) Then
            tempkode = strnumber + "0" + Trim(Str(kode))
        End If
        If (Len(Trim(Str(kode))) = 4) Then
            tempkode = strnumber + Trim(Str(kode))
        End If
        getnospkl = tempkode
    End If
    OBJ.Close
    Exit Function
Err_handler:
    getnospkl = strnumber + "0001"
    If OBJ.State = 1 Then OBJ.Close
End Function

Private Sub txtid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        OBJ.Open dsn
        SQL = "select * from list_ttd where idttd='" & txtid & "'"
        Set RS = OBJ.Execute(SQL)
        If RS.EOF Then
            OBJ.Close
            Exit Sub
        End If
        
        txtnamattd = RS!nama
        flname = App.Path & "\temp\" & RS!idttd & ".jpg"
        If Not IsNull(RS!ttd) Then
        Set RSStream = New ADODB.Stream
            RSStream.Type = adTypeBinary
            RSStream.Open
            RSStream.Write RS!ttd
            RSStream.SaveToFile flname, adSaveCreateOverWrite
            RSStream.Close
            Set RSStream = Nothing
            Imgttd.Picture = LoadPicture(flname)
        End If
        OBJ.Close
    End If
End Sub
