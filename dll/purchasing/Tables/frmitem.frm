VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmitem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bahan Baku"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Bahan Baku"
      Height          =   2175
      Left            =   120
      TabIndex        =   35
      Top             =   2160
      Width           =   6615
      Begin VB.TextBox txtdesc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   17
         Top             =   600
         Width           =   4575
      End
      Begin VB.TextBox txtkode 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtsatuan 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   18
         Top             =   960
         Width           =   495
      End
      Begin VB.ComboBox cmbkode 
         Height          =   315
         Left            =   1800
         TabIndex        =   20
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtsatuanmutasi 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   19
         Top             =   1320
         Width           =   495
      End
      Begin Chameleon.chameleonButton cmdsearch2 
         Height          =   285
         Left            =   240
         TabIndex        =   36
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         BTYPE           =   9
         TX              =   "Satuan P.Order"
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
         MICON           =   "frmitem.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdsearch1 
         Height          =   285
         Left            =   240
         TabIndex        =   37
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         BTYPE           =   9
         TX              =   "Satuan Mutasi"
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
         MICON           =   "frmitem.frx":031A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblitemcode 
         Caption         =   "Kode Bahan Baku"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label lblitemname 
         Caption         =   "Nama Bahan Baku"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   630
         Width           =   1335
      End
      Begin VB.Label lblsatuan 
         BackColor       =   &H80000005&
         Height          =   255
         Left            =   2400
         TabIndex        =   40
         Top             =   960
         Width           =   3975
      End
      Begin VB.Label Label4 
         Caption         =   "Sub Divisi"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   1710
         Width           =   1095
      End
      Begin VB.Label lblsatuanmutasi 
         BackColor       =   &H80000005&
         Height          =   255
         Left            =   2400
         TabIndex        =   38
         Top             =   1320
         Width           =   3975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Code Rules"
      Height          =   1935
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txt1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3120
         MaxLength       =   200
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txt2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3120
         MaxLength       =   200
         TabIndex        =   5
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txt3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3120
         MaxLength       =   200
         TabIndex        =   9
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txt4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3120
         MaxLength       =   200
         TabIndex        =   13
         Top             =   1320
         Width           =   2535
      End
      Begin Chameleon.chameleonButton cmd1 
         Height          =   285
         Left            =   5760
         TabIndex        =   2
         Top             =   240
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "+"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
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
         MICON           =   "frmitem.frx":0634
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmd2 
         Height          =   285
         Left            =   5760
         TabIndex        =   6
         Top             =   600
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "+"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
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
         MICON           =   "frmitem.frx":094E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmd3 
         Height          =   285
         Left            =   5760
         TabIndex        =   10
         Top             =   960
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "+"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
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
         MICON           =   "frmitem.frx":0C68
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmd4 
         Height          =   285
         Left            =   5760
         TabIndex        =   14
         Top             =   1320
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "+"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
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
         MICON           =   "frmitem.frx":0F82
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdel1 
         Height          =   285
         Left            =   6120
         TabIndex        =   3
         Top             =   240
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "-"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
         MICON           =   "frmitem.frx":129C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdel2 
         Height          =   285
         Left            =   6120
         TabIndex        =   7
         Top             =   600
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "-"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
         MICON           =   "frmitem.frx":15B6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdel3 
         Height          =   285
         Left            =   6120
         TabIndex        =   11
         Top             =   960
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "-"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
         MICON           =   "frmitem.frx":18D0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdel4 
         Height          =   285
         Left            =   6120
         TabIndex        =   15
         Top             =   1320
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "-"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
         MICON           =   "frmitem.frx":1BEA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Kode bahan baku = Level 1 + Level 2 + . + Level3 + . + 2 character"
         Height          =   225
         Left            =   240
         TabIndex        =   34
         Top             =   1680
         Width           =   6135
      End
      Begin MSForms.ComboBox l3 
         Height          =   285
         Left            =   840
         TabIndex        =   8
         Top             =   960
         Width           =   855
         VariousPropertyBits=   746608667
         MaxLength       =   4
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "1508;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         DropButtonStyle =   3
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox l2 
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Top             =   600
         Width           =   855
         VariousPropertyBits=   746608667
         MaxLength       =   3
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "1508;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         DropButtonStyle =   3
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox l1 
         Height          =   285
         Left            =   840
         TabIndex        =   0
         Top             =   240
         Width           =   855
         VariousPropertyBits=   746608667
         MaxLength       =   1
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "1508;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         DropButtonStyle =   3
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox l4 
         Height          =   285
         Left            =   840
         TabIndex        =   12
         Top             =   1320
         Width           =   855
         VariousPropertyBits=   746608667
         MaxLength       =   2
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "1508;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         DropButtonStyle =   3
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label11 
         Caption         =   "Level 4"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   1350
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Level 3"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   990
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Level 2"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   630
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Level 1"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "= 1 character"
         Height          =   255
         Left            =   1800
         TabIndex        =   29
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "= 2 character + ."
         Height          =   255
         Left            =   1800
         TabIndex        =   28
         Top             =   630
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "= 3 charcter + ."
         Height          =   255
         Left            =   1800
         TabIndex        =   27
         Top             =   990
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "= 2 character"
         Height          =   255
         Left            =   1800
         TabIndex        =   26
         Top             =   1350
         Width           =   1335
      End
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   5760
      TabIndex        =   24
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
      MICON           =   "frmitem.frx":1F04
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
      Left            =   4800
      TabIndex        =   23
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
      MICON           =   "frmitem.frx":221E
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
      Left            =   1200
      TabIndex        =   21
      Top             =   4440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Add Bahan Baku"
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
      MICON           =   "frmitem.frx":2538
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdupdate 
      Height          =   375
      Left            =   3000
      TabIndex        =   22
      Top             =   4440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Update Bahan Baku"
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
      MICON           =   "frmitem.frx":2852
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
Attribute VB_Name = "frmitem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim i As Integer

Private Sub cmbkode_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub cmbkode_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmd1_Click()
    If MsgBox("Save rule Level 1 ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    If l1 = "" Or txt1 = "" Then
        MsgBox "Data entry not complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If Len(Trim(l1)) < 1 Then
        MsgBox "Data entry not complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_apitemcode where lev = '1' and kode = '" & l1 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If MsgBox("Rule Level 1 already exist, this action will update descripiton." & vbCrLf & _
        "Continue update description ?", vbQuestion + vbYesNo, "Question") = vbNo Then
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_polin where kodebarang like '" & l1 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Purchase Order already use Rule Level 1." & vbCrLf & _
            "Update ABORTED !", vbExclamation, "Information"
                
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_belilin where kodebarang like '" & l1 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Penerimaan already use Rule Level 1." & vbCrLf & _
            "Update ABORTED !", vbExclamation, "Information"
                
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_uselin where kodebarang like '" & l1 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Pemakaian already use Rule Level 1." & vbCrLf & _
            "Update ABORTED !", vbExclamation, "Information"
                
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_mutlin where kodebarang like '" & l1 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Mutasi already use Rule Level 1." & vbCrLf & _
            "Update ABORTED !", vbExclamation, "Information"
                
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_invloc where kodebarang like '" & l1 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Beginning Stock already use Rule Level 1." & vbCrLf & _
            "Update ABORTED !", vbExclamation, "Information"
                
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_price where kodebarang like '" & l1 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Price List already use Rule Level 1." & vbCrLf & _
            "Update ABORTED !", vbExclamation, "Information"
                
            OBJ.Close
            Exit Sub
        End If
        
        If MsgBox("This action will DELETE Rule Level 2, 3, and 4." & vbCrLf & _
        "Where those rule level are related with Level 1, Continue DELETE those level ?", vbQuestion + vbYesNo, "Question") = vbNo Then
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "delete from am_apitemcode where lev = '2' and kode like '" & l1 & "%'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "delete from am_apitemcode where lev = '3' and kode like '" & l1 & "%'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "delete from am_apitemcode where lev = '4' and kode like '" & l1 & "%'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "delete from am_apitemmst where kodebarang like '" & l1 & "%'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "update am_apitemcode set ket = '" & txt1 & "' where lev = '1' and kode = '" & l1 & "'"
        Set RST = OBJ.Execute(SQL)
    Else
        SQL = "insert into am_apitemcode"
        SQL = SQL + " (lev,"
        SQL = SQL + " kode,"
        SQL = SQL + " ket)"
        
        SQL = SQL + " values"
        SQL = SQL + " ('1'"
        SQL = SQL + " ,'" & l1 & "'"
        SQL = SQL + " ,'" & txt1 & "')"
        Set RST = OBJ.Execute(SQL)
    End If
    OBJ.Close
    
    MsgBox "Item Code Rules Is Saved, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmd2_Click()
    If MsgBox("Save rule Level 2 ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    If l1 = "" Or txt1 = "" Or l2 = "" Or txt2 = "" Then
        MsgBox "Data entry not complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If Len(Trim(l1)) < 1 Or Len(Trim(l2)) < 3 Then
        MsgBox "Data entry not complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If Right(l2, 1) <> "." Then
        MsgBox "The last character of level 2 must be (.)", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_apitemcode where lev = '1' and kode = '" & l1 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 1, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "select * from am_apitemcode where lev = '2' and kode = '" & l1 & l2 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If MsgBox("Rule Level 2 already exist, this action will update descripiton." & vbCrLf & _
        "Continue update description ?", vbQuestion + vbYesNo, "Question") = vbNo Then
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_polin where kodebarang like '" & l1 & l2 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Purchase Order already use Rule Level 1 and Rule Level 2." & vbCrLf & _
            "Update ABORTED !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_belilin where kodebarang like '" & l1 & l2 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Penerimaan already use Rule Level 1 and Rule Level 2." & vbCrLf & _
            "Update ABORTED !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_uselin where kodebarang like '" & l1 & l2 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Pemakaian already use Rule Level 1 and Rule Level 2." & vbCrLf & _
            "Update ABORTED !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_mutlin where kodebarang like '" & l1 & l2 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Mutasi already use Rule Level 1 and Rule Level 2." & vbCrLf & _
            "Update ABORTED !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_invloc where kodebarang like '" & l1 & l2 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Beginning Stock already use Rule Level 1 and Rule Level 2." & vbCrLf & _
            "Update ABORTED !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_price where kodebarang like '" & l1 & l2 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Price List already use Rule Level 1 and Rule Level 2." & vbCrLf & _
            "Update ABORTED !", vbExclamation, "Information"
                
            OBJ.Close
            Exit Sub
        End If
        
        If MsgBox("This action will DELETE Rule Level 3 and 4." & vbCrLf & _
        "Where those rule level are related with Level 2, Continue DELETE those level ?", vbQuestion + vbYesNo, "Question") = vbNo Then
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "delete from am_apitemcode where lev = '3' and kode like '" & l1 & l2 & "%'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "delete from am_apitemcode where lev = '4' and kode like '" & l1 & l2 & "%'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "delete from am_apitemmst where kodebarang like '" & l1 & l2 & "%'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "update am_apitemcode set ket = '" & txt2 & "' where lev = '2' and kode = '" & l1 & l2 & "'"
        Set RST = OBJ.Execute(SQL)
    Else
        SQL = "insert into am_apitemcode"
        SQL = SQL + " (lev,"
        SQL = SQL + " kode,"
        SQL = SQL + " ket)"
        
        SQL = SQL + " values"
        SQL = SQL + " ('2'"
        SQL = SQL + " ,'" & l1 & l2 & "'"
        SQL = SQL + " ,'" & txt2 & "')"
        Set RST = OBJ.Execute(SQL)
    End If
    OBJ.Close
    
    MsgBox "Item Code Rules Is Saved, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmd3_Click()
    If MsgBox("Save rule Level 3 ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    If l1 = "" Or txt1 = "" Or l2 = "" Or txt2 = "" Or l3 = "" Or txt3 = "" Then
        MsgBox "Data entry not complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If Len(Trim(l1)) < 1 Or Len(Trim(l2)) < 3 Or Len(Trim(l3)) < 4 Then
        MsgBox "Data entry not complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If Right(l3, 1) <> "." Then
        MsgBox "The last character of level 3 must be (.)", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_apitemcode where lev = '1' and kode = '" & l1 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 1, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    
    SQL = "select * from am_apitemcode where lev = '2' and kode = '" & l1 & l2 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 2, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "select * from am_apitemcode where lev = '3' and kode = '" & l1 & l2 & l3 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If MsgBox("Rule Level 3 already exist, this action will update descripiton." & vbCrLf & _
        "Continue update description ?", vbQuestion + vbYesNo, "Question") = vbNo Then
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_polin where kodebarang like '" & l1 & l2 & l3 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Purchase Order already use Rule Level 1, 2, and Rule Level 3." & vbCrLf & _
            "Update ABORTED !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_belilin where kodebarang like '" & l1 & l2 & l3 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Penerimaan already use Rule Level 1, 2, and Rule Level 3." & vbCrLf & _
            "Update ABORTED !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_uselin where kodebarang like '" & l1 & l2 & l3 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Pemakaian already use Rule Level 1, 2, and Rule Level 3." & vbCrLf & _
            "Update ABORTED !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_mutlin where kodebarang like '" & l1 & l2 & l3 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Mutasi already use Rule Level 1, 2, and Rule Level 3." & vbCrLf & _
            "Update ABORTED !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_invloc where kodebarang like '" & l1 & l2 & l3 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Beginning Stock already use Rule Level 1, 2, and Rule Level 3." & vbCrLf & _
            "Update ABORTED !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_price where kodebarang like '" & l1 & l2 & l3 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Price List already use Rule Level 1, 2, and Rule Level 3." & vbCrLf & _
            "Update ABORTED !", vbExclamation, "Information"
                
            OBJ.Close
            Exit Sub
        End If
        
        If MsgBox("This action will DELETE Rule Level 4." & vbCrLf & _
        "Where those rule level are related with Level 3, Continue DELETE those level ?", vbQuestion + vbYesNo, "Question") = vbNo Then
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "delete from am_apitemcode where lev = '4' and kode like '" & l1 & l2 & l3 & "%'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "delete from am_apitemmst where kodebarang like '" & l1 & l2 & l3 & "%'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "update am_apitemcode set ket = '" & txt3 & "' where lev = '3' and kode = '" & l1 & l2 & l3 & "'"
        Set RST = OBJ.Execute(SQL)
    Else
        SQL = "insert into am_apitemcode"
        SQL = SQL + " (lev,"
        SQL = SQL + " kode,"
        SQL = SQL + " ket)"
        
        SQL = SQL + " values"
        SQL = SQL + " ('3'"
        SQL = SQL + " ,'" & l1 & l2 & l3 & "'"
        SQL = SQL + " ,'" & txt3 & "')"
        Set RST = OBJ.Execute(SQL)
    End If
    OBJ.Close
    
    MsgBox "Item Code Rules Is Saved, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmd4_Click()
    If MsgBox("Save rule Level 4 ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    If l1 = "" Or txt1 = "" Or l2 = "" Or txt2 = "" Or l3 = "" Or txt3 = "" Or l4 = "" Or txt4 = "" Then
        MsgBox "Data entry not complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If Len(Trim(l1)) < 1 Or Len(Trim(l2)) < 3 Or Len(Trim(l3)) < 4 Or Len(Trim(l4)) < 2 Then
        MsgBox "Data entry not complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_apitemcode where lev = '1' and kode = '" & l1 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 1, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    
    SQL = "select * from am_apitemcode where lev = '2' and kode = '" & l1 & l2 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 2, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    
    SQL = "select * from am_apitemcode where lev = '3' and kode = '" & l1 & l2 & l3 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 3, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "select * from am_apitemcode where lev = '4' and kode = '" & l1 & l2 & l3 & l4 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If MsgBox("Rule Level 4 already exist, this action will update descripiton." & vbCrLf & _
        "Continue update description ?", vbQuestion + vbYesNo, "Question") = vbNo Then
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_polin where kodebarang like '" & l1 & l2 & l3 & l4 & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Purchase Order already use Rule Level 1, 2, 3, and Rule Level 4." & vbCrLf & _
            "Update ABORTED !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_belilin where kodebarang like '" & l1 & l2 & l3 & l4 & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Penerimaan already use Rule Level 1, 2, 3, and Rule Level 4." & vbCrLf & _
            "Update ABORTED !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_uselin where kodebarang like '" & l1 & l2 & l3 & l4 & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Pemakaian already use Rule Level 1, 2, 3, and Rule Level 4." & vbCrLf & _
            "Update ABORTED !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_mutlin where kodebarang like '" & l1 & l2 & l3 & l4 & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Mutasi already use Rule Level 1, 2, 3, and Rule Level 4." & vbCrLf & _
            "Update ABORTED !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_invloc where kodebarang like '" & l1 & l2 & l3 & l4 & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Beginning Stock already use Rule Level 1, 2, 3, and Rule Level 4." & vbCrLf & _
            "Update ABORTED !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_price where kodebarang like '" & l1 & l2 & l3 & l4 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Price List already use Rule Level 1, 2, 3, and Rule Level 4." & vbCrLf & _
            "Update ABORTED !", vbExclamation, "Information"
                
            OBJ.Close
            Exit Sub
        End If
        
        If MsgBox("Proses ini akan MENGHAPUS data Item Master Bahan Baku." & vbCrLf & _
        "User harus memasukan data bahan baku yang baru, lanjutkan ?", vbQuestion + vbYesNo, "Question") = vbNo Then
            OBJ.Close
            Exit Sub
        End If
                
        SQL = "delete from am_apitemmst where kodebarang = '" & l1 & l2 & l3 & l4 & "'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "update am_apitemcode set ket = '" & txt4 & "' where lev = '4' and kode = '" & l1 & l2 & l3 & l4 & "'"
        Set RST = OBJ.Execute(SQL)
        
        txtdesc = ""
        txtsatuan = ""
        txtsatuanmutasi = ""
        lblsatuan = ""
        lblsatuanmutasi = ""
        cmbkode = ""
    Else
        SQL = "insert into am_apitemcode"
        SQL = SQL + " (lev,"
        SQL = SQL + " kode,"
        SQL = SQL + " ket)"
        
        SQL = SQL + " values"
        SQL = SQL + " ('4'"
        SQL = SQL + " ,'" & l1 & l2 & l3 & l4 & "'"
        SQL = SQL + " ,'" & txt4 & "')"
        Set RST = OBJ.Execute(SQL)
    End If
    OBJ.Close
    
    MsgBox "Item Code Rules Is Saved, Click OK To Continue ...", vbInformation, "Information"
End Sub

Private Sub cmdadd_Click()
    If Len(Trim(txtkode)) = 0 Then
        MsgBox "Data entry not Complete.", vbExclamation, "Warning"
        txtkode.SetFocus
        Exit Sub
    End If
    
    If txtkode = "" Or txtdesc = "" Or txtsatuanmutasi = "" Or txtsatuan = "" Or cmbkode = "" Or l1 = "" Or l2 = "" Or l3 = "" Or l4 = "" Then
       MsgBox "Data entry not Complete.", vbExclamation, "Warning"
       Exit Sub
    End If
    
    txtkode = Trim(txtkode)
    
    OBJ.Open dsn
    SQL = "select * from am_apitemcode where lev = '1' and kode = '" & l1 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 1, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    
    SQL = "select * from am_apitemcode where lev = '2' and kode = '" & l1 & l2 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 2, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    
    SQL = "select * from am_apitemcode where lev = '3' and kode = '" & l1 & l2 & l3 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 3, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    
    SQL = "select * from am_apitemcode where lev = '4' and kode = '" & l1 & l2 & l3 & l4 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 4, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "select * from am_apitemmst where kodebarang = '" & txtkode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        MsgBox "Data already exist.", vbInformation, "Information"
        OBJ.Close
        cmdclear_Click
        Exit Sub
    End If
    
    SQL = "INSERT INTO am_apitemmst"
    SQL = SQL + "(KodeBarang"
    SQL = SQL + ",NamaBarang"
    SQL = SQL + ",KodeSatuan"
    SQL = SQL + ",KodeSatuanmutasi"
    SQL = SQL + ",Kodeproduk)"
    
    SQL = SQL + "VALUES"
    SQL = SQL + " ('" & txtkode & "'"
    SQL = SQL + ", '" & txtdesc & "'"
    SQL = SQL + ", '" & txtsatuan & "'"
    SQL = SQL + ", '" & txtsatuanmutasi & "'"
    SQL = SQL + ", '" & cmbkode & "')"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
        
    MsgBox "Data saved, click ok to continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdclear_Click()
    txtkode = ""
    txtdesc = ""
    txtsatuan = ""
    cmbkode = ""
    lblsatuan = ""
    txtsatuanmutasi = ""
    lblsatuanmutasi = ""
    
    l1.Clear
    l1.ColumnCount = 2
    l1.ListWidth = "6 cm"
    l1.ColumnWidths = "2 cm; 4 cm"
    i = 0
    
    OBJ.Open dsn
    SQL = "select distinct substring(b.kode,1,1)'kode',b.ket from am_apitemcode b where b.lev='1' order by substring(b.kode,1,1)"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        l1.AddItem RST!kode
        l1.List(i, 1) = RST!ket
        i = i + 1
        RST.MoveNext
    Loop
    OBJ.Close
    
    l2.Clear
    l3.Clear
    l4.Clear
    l1 = ""
    l2 = ""
    l3 = ""
    l4 = ""
    txt1 = ""
    txt2 = ""
    txt3 = ""
    txt4 = ""
    
    txtkode.SetFocus
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdel1_Click()
    If MsgBox("Delete rule Level 1 ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    If l1 = "" Or txt1 = "" Then
        MsgBox "Data entry not complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If Len(Trim(l1)) < 1 Then
        MsgBox "Data entry not complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_apitemcode where lev = '1' and kode = '" & l1 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If MsgBox("Continue DELETE Rule Level 1 ?", vbQuestion + vbYesNo, "Question") = vbNo Then
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_polin where kodebarang like '" & l1 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Purchase Order already use Rule Level 1." & vbCrLf & _
            "Delete ABORTED !", vbExclamation, "Information"
                
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_belilin where kodebarang like '" & l1 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Penerimaan already use Rule Level 1." & vbCrLf & _
            "Delete ABORTED !", vbExclamation, "Information"
                
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_uselin where kodebarang like '" & l1 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Pemakaian already use Rule Level 1." & vbCrLf & _
            "Delete ABORTED !", vbExclamation, "Information"
                
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_mutlin where kodebarang like '" & l1 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Mutasi already use Rule Level 1." & vbCrLf & _
            "Delete ABORTED !", vbExclamation, "Information"
                
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_invloc where kodebarang like '" & l1 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Beginning Stock already use Rule Level 1." & vbCrLf & _
            "Delete ABORTED !", vbExclamation, "Information"
                
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_price where kodebarang like '" & l1 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Price List already use Rule Level 1." & vbCrLf & _
            "Delete ABORTED !", vbExclamation, "Information"
                
            OBJ.Close
            Exit Sub
        End If
        
        If MsgBox("This action will DELETE Rule Level 1, 2, 3, and 4." & vbCrLf & _
        "Where those rule level are related with Level 1, Continue DELETE those level ?", vbQuestion + vbYesNo, "Question") = vbNo Then
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "delete from am_apitemcode where lev = '1' and kode like '" & l1 & "%'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "delete from am_apitemcode where lev = '2' and kode like '" & l1 & "%'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "delete from am_apitemcode where lev = '3' and kode like '" & l1 & "%'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "delete from am_apitemcode where lev = '4' and kode like '" & l1 & "%'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "delete from am_apitemmst where kodebarang like '" & l1 & "%'"
        Set RST = OBJ.Execute(SQL)
    Else
        MsgBox "Data not found, Delete ABORTED !", vbExclamation, "Information"
            
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    MsgBox "Item Code Rules Is Deleted, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdel2_Click()
    If MsgBox("Delete rule Level 2 ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    If l1 = "" Or txt1 = "" Or l2 = "" Or txt2 = "" Then
        MsgBox "Data entry not complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If Len(Trim(l1)) < 1 Or Len(Trim(l2)) < 3 Then
        MsgBox "Data entry not complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_apitemcode where lev = '1' and kode = '" & l1 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 1, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "select * from am_apitemcode where lev = '2' and kode = '" & l1 & l2 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If MsgBox("Continue DELETE Rule Level 2 ?", vbQuestion + vbYesNo, "Question") = vbNo Then
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_polin where kodebarang like '" & l1 & l2 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Purchase Order already use Rule Level 1 and Rule Level 2." & vbCrLf & _
            "Delete ABORTED !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_belilin where kodebarang like '" & l1 & l2 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Penerimaan already use Rule Level 1 and Rule Level 2." & vbCrLf & _
            "Delete ABORTED !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_uselin where kodebarang like '" & l1 & l2 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Pemakaian already use Rule Level 1 and Rule Level 2." & vbCrLf & _
            "Delete ABORTED !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_mutlin where kodebarang like '" & l1 & l2 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Mutasi already use Rule Level 1 and Rule Level 2." & vbCrLf & _
            "Delete ABORTED !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_invloc where kodebarang like '" & l1 & l2 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Beginning Stock already use Rule Level 1 and Rule Level 2." & vbCrLf & _
            "Delete ABORTED !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_price where kodebarang like '" & l1 & l2 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Price List already use Rule Level 1 and Rule Level 2." & vbCrLf & _
            "Delete ABORTED !", vbExclamation, "Information"
                
            OBJ.Close
            Exit Sub
        End If
        
        If MsgBox("This action will DELETE Rule Level 2, 3 and 4." & vbCrLf & _
        "Where those rule level are related with Level 2, Continue DELETE those level ?", vbQuestion + vbYesNo, "Question") = vbNo Then
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "delete from am_apitemcode where lev = '2' and kode like '" & l1 & l2 & "%'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "delete from am_apitemcode where lev = '3' and kode like '" & l1 & l2 & "%'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "delete from am_apitemcode where lev = '4' and kode like '" & l1 & l2 & "%'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "delete from am_apitemmst where kodebarang like '" & l1 & l2 & "%'"
        Set RST = OBJ.Execute(SQL)
    Else
        MsgBox "Data not found, Delete ABORTED !", vbExclamation, "Information"
            
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    MsgBox "Item Code Rules Is Deleted, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdel3_Click()
    If MsgBox("Delete rule Level 3 ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    If l1 = "" Or txt1 = "" Or l2 = "" Or txt2 = "" Or l3 = "" Or txt3 = "" Then
        MsgBox "Data entry not complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If Len(Trim(l1)) < 1 Or Len(Trim(l2)) < 3 Or Len(Trim(l3)) < 4 Then
        MsgBox "Data entry not complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_apitemcode where lev = '1' and kode = '" & l1 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 1, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    
    SQL = "select * from am_apitemcode where lev = '2' and kode = '" & l1 & l2 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 2, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "select * from am_apitemcode where lev = '3' and kode = '" & l1 & l2 & l3 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If MsgBox("Continue DELETE Rule Level 3 ?", vbQuestion + vbYesNo, "Question") = vbNo Then
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_polin where kodebarang like '" & l1 & l2 & l3 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Purchase Order already use Rule Level 1, 2, and Rule Level 3." & vbCrLf & _
            "Delete ABORTED !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_belilin where kodebarang like '" & l1 & l2 & l3 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Penerimaan already use Rule Level 1, 2, and Rule Level 3." & vbCrLf & _
            "Delete ABORTED !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_uselin where kodebarang like '" & l1 & l2 & l3 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Pemakaian already use Rule Level 1, 2, and Rule Level 3." & vbCrLf & _
            "Delete ABORTED !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_mutlin where kodebarang like '" & l1 & l2 & l3 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Mutasi already use Rule Level 1, 2, and Rule Level 3." & vbCrLf & _
            "Delete ABORTED !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_invloc where kodebarang like '" & l1 & l2 & l3 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Beginning Stock already use Rule Level 1, 2, and Rule Level 3." & vbCrLf & _
            "Delete ABORTED !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_price where kodebarang like '" & l1 & l2 & l3 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Price List already use Rule Level 1, 2, and Rule Level 3." & vbCrLf & _
            "Delete ABORTED !", vbExclamation, "Information"
                
            OBJ.Close
            Exit Sub
        End If
        
        If MsgBox("This action will DELETE Rule Level 3 and 4." & vbCrLf & _
        "Where those rule level are related with Level 3, Continue DELETE those level ?", vbQuestion + vbYesNo, "Question") = vbNo Then
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "delete from am_apitemcode where lev = '3' and kode like '" & l1 & l2 & l3 & "%'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "delete from am_apitemcode where lev = '4' and kode like '" & l1 & l2 & l3 & "%'"
        Set RST = OBJ.Execute(SQL)
        
        SQL = "delete from am_apitemmst where kodebarang like '" & l1 & l2 & l3 & "%'"
        Set RST = OBJ.Execute(SQL)
    Else
        MsgBox "Data not found, Delete ABORTED !", vbExclamation, "Information"
            
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    MsgBox "Item Code Rules Is Deleted, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdel4_Click()
    If MsgBox("Delete rule Level 4 ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    If l1 = "" Or txt1 = "" Or l2 = "" Or txt2 = "" Or l3 = "" Or txt3 = "" Or l4 = "" Or txt4 = "" Then
        MsgBox "Data entry not complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    If Len(Trim(l1)) < 1 Or Len(Trim(l2)) < 3 Or Len(Trim(l3)) < 4 Or Len(Trim(l4)) < 2 Then
        MsgBox "Data entry not complite.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_apitemcode where lev = '1' and kode = '" & l1 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 1, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    
    SQL = "select * from am_apitemcode where lev = '2' and kode = '" & l1 & l2 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 2, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    
    SQL = "select * from am_apitemcode where lev = '3' and kode = '" & l1 & l2 & l3 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Unrecognized Level 3, action aborted,", vbExclamation, "Unrecognized"
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "select * from am_apitemcode where lev = '4' and kode = '" & l1 & l2 & l3 & l4 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If MsgBox("Continue DELETE Rule Level 4 ?", vbQuestion + vbYesNo, "Question") = vbNo Then
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_polin where kodebarang like '" & l1 & l2 & l3 & l4 & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Purchase Order already use Rule Level 1, 2, 3, and Rule Level 4." & vbCrLf & _
            "Delete ABORTED !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_belilin where kodebarang like '" & l1 & l2 & l3 & l4 & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Penerimaan already use Rule Level 1, 2, 3, and Rule Level 4." & vbCrLf & _
            "Delete ABORTED !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_uselin where kodebarang like '" & l1 & l2 & l3 & l4 & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Pemakaian already use Rule Level 1, 2, 3, and Rule Level 4." & vbCrLf & _
            "Delete ABORTED !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_mutlin where kodebarang like '" & l1 & l2 & l3 & l4 & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Mutasi already use Rule Level 1, 2, 3, and Rule Level 4." & vbCrLf & _
            "Delete ABORTED !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_invloc where kodebarang like '" & l1 & l2 & l3 & l4 & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Beginning Stock already use Rule Level 1, 2, 3, and Rule Level 4." & vbCrLf & _
            "Delete ABORTED !", vbExclamation, "Information"
            
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "select * from am_price where kodebarang like '" & l1 & l2 & l3 & l4 & "%'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            MsgBox "Price List already use Rule Level 1, 2, 3, and Rule Level 4." & vbCrLf & _
            "Delete ABORTED !", vbExclamation, "Information"
                
            OBJ.Close
            Exit Sub
        End If
        
        If MsgBox("Proses ini akan MENGHAPUS data Item Master Bahan Baku." & vbCrLf & _
        "User harus memasukan data bahan baku yang baru, lanjutkan ?", vbQuestion + vbYesNo, "Question") = vbNo Then
            OBJ.Close
            Exit Sub
        End If
        
        If MsgBox("This action will DELETE Rule Level 4, Continue DELETE ?", vbQuestion + vbYesNo, "Question") = vbNo Then
            OBJ.Close
            Exit Sub
        End If
        
        SQL = "delete from am_apitemcode where lev = '4' and kode = '" & l1 & l2 & l3 & l4 & "'"
        Set RST = OBJ.Execute(SQL)
                
        SQL = "delete from am_apitemmst where kodebarang = '" & l1 & l2 & l3 & l4 & "'"
        Set RST = OBJ.Execute(SQL)
    Else
        MsgBox "Data not found, Delete ABORTED !", vbExclamation, "Information"
            
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    MsgBox "Item Code Rules Is Deleted, Click OK To Continue ...", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub cmdsearch1_Click()
    carisql1 = "select kodesatuan, namasatuan from am_apunit"
    namatabel = "Satuan Bahan Baku"
    
    frmsearch.Show
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtsatuanmutasi = hasil
    carisatuanmutasi
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdsearch2_Click()
    carisql1 = "select kodesatuan, namasatuan from am_apunit"
    namatabel = "Satuan Bahan Baku"
    
    frmsearch.Show
End Sub

Private Sub cmdsearch2_GotFocus()
    If hasil = "" Then Exit Sub
    txtsatuan = hasil
    carisatuan
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdupdate_click()
    If txtkode = "" Or txtdesc = "" Or txtsatuanmutasi = "" Or txtsatuan = "" Or cmbkode = "" Or l1 = "" Or l2 = "" Or l3 = "" Or l4 = "" Then
       MsgBox "Data entry not Complete.", vbExclamation, "Warning"
       Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_apitemmst where kodebarang = '" & txtkode & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Data not found, update aborted.", vbExclamation, "Information"
        OBJ.Close
        
        Exit Sub
    End If
    OBJ.Close
    
    If MsgBox("Are You Sure Want To Update ?", vbQuestion + vbYesNo, "Question") = vbNo Then Exit Sub
    
    OBJ.Open dsn
    SQL = "UPDATE am_apitemmst SET "
    SQL = SQL + "NamaBarang = '" & txtdesc & "'"
    SQL = SQL + "WHERE KodeBarang =  '" & txtkode & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "UPDATE am_apitemcode SET "
    SQL = SQL + "ket = '" & txtdesc & "'"
    SQL = SQL + "WHERE kode = '" & txtkode & "'"
    Set RST = OBJ.Execute(SQL)
    
    SQL = "select * from AM_polin WHERE Kodebarang = '" & txtkode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then GoTo jump1
    
    SQL = "select * from AM_uselin WHERE Kodebarang = '" & txtkode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then GoTo jump1
    
    SQL = "select * from AM_mutlin WHERE Kodebarang = '" & txtkode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then GoTo jump1
    
    SQL = "select * from AM_belilin WHERE Kodebarang = '" & txtkode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then GoTo jump1
    
    SQL = "select * from AM_price WHERE Kodebarang = '" & txtkode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then GoTo jump1
    
    SQL = "UPDATE am_apitemmst SET "
    SQL = SQL + "kodeproduk = '" & cmbkode & "'"
    SQL = SQL + "WHERE Kodebarang =  '" & txtkode & "'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Data updated, click ok to continue ...", vbInformation, "Information"
    cmdclear_Click
    
    Exit Sub
    
jump1:
    OBJ.Close
    MsgBox "Name updated, but can not update Sub Divisi, data in use.", vbInformation, "Information"
    cmdclear_Click
End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='61' and b.kodeuser = '2" & kuser & "'"
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
    txtkode.ToolTipText = "max length = " & txtkode.MaxLength
    txtdesc.ToolTipText = "max length = " & txtdesc.MaxLength
    txtsatuan.ToolTipText = "max length = " & txtsatuan.MaxLength
    
    OBJ.Open dsn
    SQL = "select * from am_kode"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        cmbkode.AddItem RST!kode3
            
        RST.MoveNext
    Loop
    OBJ.Close
    
    l1.Clear
    l1.ColumnCount = 2
    l1.ListWidth = "6 cm"
    l1.ColumnWidths = "2 cm; 4 cm"
    i = 0
    
    OBJ.Open dsn
    SQL = "select distinct substring(b.kode,1,1)'kode',b.ket from am_apitemcode b where b.lev='1' order by substring(b.kode,1,1)"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        l1.AddItem RST!kode
        l1.List(i, 1) = RST!ket
        i = i + 1
        RST.MoveNext
    Loop
    OBJ.Close
End Sub

Private Sub l1_Change()
    l2.Clear
    l2.ColumnCount = 2
    l2.ListWidth = "6 cm"
    l2.ColumnWidths = "2 cm; 4 cm"
    l3.Clear
    l4.Clear
    txt2 = ""
    txt3 = ""
    txt4 = ""
    
    i = 0
    OBJ.Open dsn
    SQL = "select distinct substring(b.kode,2,3)'kode',b.ket from am_apitemcode b where b.lev='2' and substring(b.kode,1,1)='" & l1 & "' order by substring(b.kode,2,3)"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        l2.AddItem RST!kode
        l2.List(i, 1) = RST!ket
        i = i + 1
        RST.MoveNext
    Loop
    OBJ.Close
    
    txtkode = l1 & l2 & l3 & l4
End Sub

Private Sub l1_DropButtonClick()
    OBJ.Open dsn
    SQL = "select b.ket from am_apitemcode b where b.lev='1' and substring(b.kode,1,1)='" & l1 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txt1 = RST!ket
    Else
        txt1 = ""
    End If
    OBJ.Close
End Sub

Private Sub l1_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub l2_Change()
    l3.Clear
    l3.ColumnCount = 2
    l3.ListWidth = "6 cm"
    l3.ColumnWidths = "2 cm; 4 cm"
    l4.Clear
    txt3 = ""
    txt4 = ""
    
    i = 0
    OBJ.Open dsn
    SQL = "select distinct substring(b.kode,5,4)'kode',b.ket from am_apitemcode b where b.lev='3' and substring(b.kode,1,4) = '" & l1 & l2 & "' order by substring(b.kode,5,4)"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        l3.AddItem RST!kode
        l3.List(i, 1) = RST!ket
        i = i + 1
        RST.MoveNext
    Loop
    OBJ.Close
    
    txtkode = l1 & l2 & l3 & l4
End Sub

Private Sub l2_DropButtonClick()
    OBJ.Open dsn
    SQL = "select b.ket from am_apitemcode b where b.lev='2' and b.kode = '" & l1 & l2 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txt2 = RST!ket
    Else
        txt2 = ""
    End If
    OBJ.Close
End Sub

Private Sub l2_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub l3_Change()
    l4.Clear
    l4.ColumnCount = 2
    l4.ListWidth = "6 cm"
    l4.ColumnWidths = "2 cm; 4 cm"
    txt4 = ""
    
    i = 0
    OBJ.Open dsn
    SQL = "select distinct substring(b.kode,9,2)'kode',b.ket from am_apitemcode b where b.lev='4' and substring(b.kode,1,8) = '" & l1 & l2 & l3 & "' order by substring(b.kode,9,2)"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        l4.AddItem RST!kode
        l4.List(i, 1) = RST!ket
        i = i + 1
        RST.MoveNext
    Loop
    OBJ.Close
    
    txtkode = l1 & l2 & l3 & l4
End Sub

Private Sub l3_DropButtonClick()
    OBJ.Open dsn
    SQL = "select b.ket from am_apitemcode b where b.lev='3' and b.kode = '" & l1 & l2 & l3 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txt3 = RST!ket
    Else
        txt3 = ""
    End If
    OBJ.Close
End Sub

Private Sub l3_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub l4_Change()
    txtkode = l1 & l2 & l3 & l4
End Sub

Private Sub l4_DropButtonClick()
    OBJ.Open dsn
    SQL = "select b.ket from am_apitemcode b where b.lev='4' and b.kode = '" & l1 & l2 & l3 & l4 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txt4 = RST!ket
        txtdesc = RST!ket
    Else
        txt4 = ""
    End If
    OBJ.Close
End Sub

Private Sub l4_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtdesc_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtdesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtsatuan.SetFocus
End Sub

Private Sub txtkode_Change()
    txtdesc = ""
    txtsatuan = ""
    lblsatuan = ""
    txtsatuanmutasi = ""
    lblsatuanmutasi = ""
    cmbkode = ""
    
    OBJ.Open dsn
    SQL = "select * from am_apitemmst where kodebarang = '" & txtkode & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txtdesc = RST!NamaBarang
        txtsatuan = RST!kodesatuan
        txtsatuanmutasi = RST!kodesatuanmutasi
        cmbkode = RST!kodeproduk
            
        SQL = "select * from am_apunit where kodesatuan = '" & txtsatuan & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then lblsatuan = RST!namasatuan
            
        SQL = "select * from am_apunit where kodesatuan = '" & txtsatuanmutasi & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then lblsatuanmutasi = RST!namasatuan
    End If
    OBJ.Close
End Sub

Private Sub txtkode_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub txtKode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtdesc.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtsatuan_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtsatuanmutasi.SetFocus
End Sub

Private Sub txtsatuan_LostFocus()
    carisatuan
End Sub

Private Sub txtsatuanmutasi_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmbkode.SetFocus
End Sub

Private Sub txtsatuanmutasi_LostFocus()
    carisatuanmutasi
End Sub

Private Sub carisatuan()
    If txtsatuan = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from am_apunit where kodesatuan = '" & txtsatuan & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblsatuan = RST!namasatuan
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    MsgBox "Data not found.", vbInformation, "Infomation"
    txtsatuan = ""
    lblsatuan = ""
    txtsatuan.SetFocus
End Sub

Private Sub carisatuanmutasi()
    If txtsatuanmutasi = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from am_apunit where kodesatuan = '" & txtsatuanmutasi & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblsatuanmutasi = RST!namasatuan
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    MsgBox "Data not found.", vbInformation, "Infomation"
    txtsatuanmutasi = ""
    lblsatuanmutasi = ""
    txtsatuanmutasi.SetFocus
End Sub

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function
