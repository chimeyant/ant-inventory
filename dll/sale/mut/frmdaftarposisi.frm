VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmdaftarposisi 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Laporan Stock"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   7223
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Mutasi dan Posisi"
      TabPicture(0)   =   "frmdaftarposisi.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdsearch4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdsearch3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdsearch2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdsearch1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "date2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "date1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Option2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Option1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtinv2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtinv1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtinv3"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtinv4"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Stock Akhir"
      TabPicture(1)   =   "frmdaftarposisi.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label7"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label11"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label9"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label8"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "l1"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "l2"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "l4"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "l5"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "l6"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "l7"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "l8"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label10"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "cmdsearch5"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "date3"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txt8"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "txt7"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txt6"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "txt5"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "txt1"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "txt2"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "txt4"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "txtgudang"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).ControlCount=   25
      Begin VB.TextBox txtinv4 
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
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   8
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtinv3 
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
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   7
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtinv1 
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
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   5
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtinv2 
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
         Left            =   4080
         MaxLength       =   15
         TabIndex        =   6
         Top             =   1560
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Posisi Stock"
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
         TabIndex        =   1
         Top             =   480
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Mutasi Stock"
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
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtgudang 
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
         Left            =   -73200
         MaxLength       =   10
         TabIndex        =   24
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox txt4 
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
         Left            =   -73200
         MaxLength       =   200
         TabIndex        =   14
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox txt2 
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
         Left            =   -73200
         MaxLength       =   200
         TabIndex        =   12
         Top             =   840
         Width           =   3735
      End
      Begin VB.TextBox txt1 
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
         Left            =   -73200
         MaxLength       =   200
         TabIndex        =   10
         Top             =   480
         Width           =   3735
      End
      Begin VB.TextBox txt5 
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
         Left            =   -73200
         MaxLength       =   200
         TabIndex        =   16
         Top             =   1680
         Width           =   3735
      End
      Begin VB.TextBox txt6 
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
         Left            =   -73200
         MaxLength       =   200
         TabIndex        =   18
         Top             =   2040
         Width           =   3735
      End
      Begin VB.TextBox txt7 
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
         Left            =   -73200
         MaxLength       =   200
         TabIndex        =   20
         Top             =   2400
         Width           =   3735
      End
      Begin VB.TextBox txt8 
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
         Left            =   -73200
         MaxLength       =   200
         TabIndex        =   22
         Top             =   2760
         Width           =   3735
      End
      Begin MSComCtl2.DTPicker date3 
         Height          =   285
         Left            =   -73200
         TabIndex        =   23
         Top             =   3240
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
         Format          =   143982595
         CurrentDate     =   37845
      End
      Begin Chameleon.chameleonButton cmdsearch5 
         Height          =   285
         Left            =   -74400
         TabIndex        =   34
         Top             =   3600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BTYPE           =   9
         TX              =   "Gudang"
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
         MICON           =   "frmdaftarposisi.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
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
         Format          =   143982595
         CurrentDate     =   37845
      End
      Begin MSComCtl2.DTPicker date2 
         Height          =   285
         Left            =   3600
         TabIndex        =   4
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
         Format          =   143982595
         CurrentDate     =   37845
      End
      Begin Chameleon.chameleonButton cmdsearch1 
         Height          =   285
         Left            =   240
         TabIndex        =   38
         Top             =   1560
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BTYPE           =   9
         TX              =   "Dari Barang"
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
         MICON           =   "frmdaftarposisi.frx":0352
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdsearch2 
         Height          =   285
         Left            =   2880
         TabIndex        =   39
         Top             =   1560
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BTYPE           =   9
         TX              =   "s/d Barang"
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
         MICON           =   "frmdaftarposisi.frx":066C
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Chameleon.chameleonButton cmdsearch3 
         Height          =   285
         Left            =   240
         TabIndex        =   40
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BTYPE           =   9
         TX              =   "Dari Gudang"
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
         MICON           =   "frmdaftarposisi.frx":0986
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
         Left            =   2880
         TabIndex        =   41
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BTYPE           =   9
         TX              =   "s/d Gudang"
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
         MICON           =   "frmdaftarposisi.frx":0CA0
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Dari Tanggal"
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
         TabIndex        =   37
         Top             =   1230
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "s/d"
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
         Left            =   3240
         TabIndex        =   36
         Top             =   1230
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "per Tanggal"
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
         Left            =   -74400
         TabIndex        =   35
         Top             =   3270
         Width           =   1335
      End
      Begin MSForms.ComboBox l8 
         Height          =   285
         Left            =   -74160
         TabIndex        =   21
         Top             =   2760
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
      Begin MSForms.ComboBox l7 
         Height          =   285
         Left            =   -74160
         TabIndex        =   19
         Top             =   2400
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
      Begin MSForms.ComboBox l6 
         Height          =   285
         Left            =   -74160
         TabIndex        =   17
         Top             =   2040
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
      Begin MSForms.ComboBox l5 
         Height          =   285
         Left            =   -74160
         TabIndex        =   15
         Top             =   1680
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
      Begin MSForms.ComboBox l4 
         Height          =   285
         Left            =   -74160
         TabIndex        =   13
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
      Begin MSForms.ComboBox l2 
         Height          =   285
         Left            =   -74160
         TabIndex        =   11
         Top             =   840
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
      Begin MSForms.ComboBox l1 
         Height          =   285
         Left            =   -74160
         TabIndex        =   9
         Top             =   480
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
      Begin VB.Label Label8 
         Caption         =   "Level 1"
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
         Left            =   -74760
         TabIndex        =   33
         Top             =   510
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Level 2 "
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
         Left            =   -74760
         TabIndex        =   32
         Top             =   870
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Kolom1"
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
         Left            =   -74760
         TabIndex        =   31
         Top             =   1350
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Kolom2"
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
         Left            =   -74760
         TabIndex        =   30
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Kolom3"
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
         Left            =   -74760
         TabIndex        =   29
         Top             =   2070
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Kolom4"
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
         Left            =   -74760
         TabIndex        =   28
         Top             =   2430
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Kolom5"
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
         Left            =   -74760
         TabIndex        =   27
         Top             =   2790
         Width           =   735
      End
   End
   Begin Crystal.CrystalReport Crystal 
      Left            =   240
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   4680
      TabIndex        =   26
      Top             =   4200
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
      MICON           =   "frmdaftarposisi.frx":0FBA
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
      Left            =   3720
      TabIndex        =   25
      Top             =   4200
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
      MICON           =   "frmdaftarposisi.frx":12D4
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
Attribute VB_Name = "frmdaftarposisi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim str1 As String
Dim i As Integer

Private Sub cariinv1()
    If txtinv1 = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from am_itemmst where kodebarang = '" & txtinv1 & "' and len(kodebarang)=8"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Data not found.", vbExclamation, "Warning"
        txtinv1 = ""
        txtinv1.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub cariinv2()
    If txtinv2 = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from am_itemmst where kodebarang = '" & txtinv2 & "' and len(kodebarang)=8"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Data not found.", vbExclamation, "Warning"
        txtinv2 = ""
        txtinv2.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub cariinv3()
    If txtinv3 = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from am_gudang where kodegudang = '" & txtinv3 & "'"
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
    SQL = "select * from am_gudang where kodegudang = '" & txtinv4 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Data not found.", vbExclamation, "Warning"
        txtinv4 = ""
        txtinv4.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdsearch5_Click()
    carisql1 = "select kodegudang, namagudang from am_gudang"
    namatabel = "Gudang"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch5_GotFocus()
    If hasil = "" Then Exit Sub
    txtgudang = hasil
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdview_Click()
    If par5 = "0" Then str1 = "sj" Else str1 = "ki"
    
    If SSTab1.Tab = 0 Then
        If txtinv1 = "" Or txtinv2 = "" Then Exit Sub
        If txtinv3 = "" Or txtinv4 = "" Then Exit Sub
        If txtinv4 < txtinv3 Then
            MsgBox "To gudang Can Not Smaller Then From gudang.", vbOKOnly + vbExclamation, "Warning"
            txtinv4 = ""
            txtinv4.SetFocus
            Exit Sub
        End If
        
        If txtinv2 < txtinv1 Then
            MsgBox "To barang Can Not Smaller Then From barang.", vbOKOnly + vbExclamation, "Warning"
            txtinv4 = ""
            txtinv4.SetFocus
            Exit Sub
        End If
        
        If date2 < Date1 Then
            MsgBox "To Date Can Not Smaller Then From Date.", vbExclamation, "Warning"
            date2.SetFocus
            Exit Sub
        End If
        
        If Option2.Value = True Then txtinv4 = txtinv3

        Crystal.Reset
        Crystal.WindowState = crptMaximized
        Crystal.WindowShowCloseBtn = True
        Crystal.WindowShowPrintSetupBtn = True
        Crystal.WindowShowSearchBtn = True
        Crystal.WindowShowRefreshBtn = True
        Crystal.Connect = dsnreport
        If Option1.Value = True Then
            Crystal.ReportFileName = AppPath & "\reports\sale\mut\posisistock.rpt"
            Crystal.DataFiles(0) = "Proc(am_posisistock)"
        ElseIf Option2.Value = True Then
            Crystal.ReportFileName = AppPath & "\reports\sale\mut\mutasistock.rpt"
            Crystal.DataFiles(0) = "Proc(am_mutasistock)"
        End If
        Crystal.ParameterFields(0) = "@namauser;" + nmuser + ";true"
        Crystal.ParameterFields(1) = "@tanggal1 ;" + Format(Date1, "yyyymmdd") + ";true"
        Crystal.ParameterFields(2) = "@tanggal2 ;" + Format(date2, "yyyymmdd") + ";true"
        Crystal.ParameterFields(3) = "@kode1;" & txtinv1 & ";true"
        Crystal.ParameterFields(4) = "@kode2;" & txtinv2 & ";true"
        Crystal.ParameterFields(5) = "@kode3;" & txtinv3 & ";true"
        Crystal.ParameterFields(6) = "@kode4;" & txtinv4 & ";true"
        Crystal.ParameterFields(7) = "@kode5;" & str1 & ";true"
    Else
        If txtgudang = "" Then Exit Sub
        If l1 = "" Or l2 = "" Then
            MsgBox "Invalid Level1 and 2.", vbOKOnly + vbExclamation, "Warning"
            Exit Sub
        End If
        
        If l4 = "" And l5 = "" And l6 = "" And l7 = "" And l8 = "" Then
            MsgBox "Invalid Define Coloumn.", vbOKOnly + vbExclamation, "Warning"
            Exit Sub
        End If
        
        Crystal.Reset
        Crystal.WindowState = crptMaximized
        Crystal.WindowShowCloseBtn = True
        Crystal.WindowShowPrintSetupBtn = True
        Crystal.WindowShowSearchBtn = True
        Crystal.WindowShowRefreshBtn = True
        Crystal.Connect = dsnreport
        Crystal.ReportFileName = AppPath & "\reports\sale\mut\stockakhir.rpt"
        Crystal.DataFiles(0) = "Proc(am_stockakhir)"
        Crystal.ParameterFields(0) = "@namauser;" + nmuser + ";true"
        Crystal.ParameterFields(1) = "@tanggal ;" + Format(date3, "yyyymmdd") + ";true"
        Crystal.ParameterFields(2) = "@kode;" & l1 & l2 & ";true"
        Crystal.ParameterFields(3) = "@kode1;" & l4 & ";true"
        Crystal.ParameterFields(4) = "@kode2;" & l5 & ";true"
        Crystal.ParameterFields(5) = "@kode3;" & l6 & ";true"
        Crystal.ParameterFields(6) = "@kode4;" & l7 & ";true"
        Crystal.ParameterFields(7) = "@kode5;" & l8 & ";true"
        Crystal.ParameterFields(8) = "@kode6;" & txtgudang & ";true"
        Crystal.ParameterFields(9) = "@kode7;" & l1 & ";true"
        Crystal.ParameterFields(10) = "@kode8;" & str1 & ";true"
    End If
    Crystal.RetrieveDataFiles
    Crystal.Action = 1
End Sub

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='124' and b.kodeuser = '1" & kuser & "'"
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
    
    Date1 = Date
    date2 = Date
    date3 = Date
    
    l1.Clear
    l1.ColumnCount = 2
    l1.ListWidth = "6 cm"
    l1.ColumnWidths = "2 cm; 4 cm"
    i = 0
    
    OBJ.Open dsn
    SQL = "select distinct substring(b.kode,1,1)'kode',b.ket from am_itemcode b where b.lev='1' order by substring(b.kode,1,1)"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        l1.AddItem RST!kode
        l1.List(i, 1) = RST!ket
        i = i + 1
        RST.MoveNext
    Loop
    OBJ.Close
End Sub

Function nmuser()
    'nmuser = "-no user-"
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select * from am_user where kodeuser = '" & kuser & "'"
    '    Set RST = OBJ.Execute(SQL)
    '    If Not RST.EOF Then nmuser = RST!namauser
    '    OBJ.Close
    'End If
End Function

Private Sub l1_Change()
    l2.Clear
    l2.ColumnCount = 2
    l2.ListWidth = "6 cm"
    l2.ColumnWidths = "2 cm; 4 cm"
    l4.Clear
    l5.Clear
    l6.Clear
    l7.Clear
    l8.Clear
    txt2 = ""
    txt4 = ""
    txt5 = ""
    txt6 = ""
    txt7 = ""
    txt8 = ""
    
    i = 0
    OBJ.Open dsn
    SQL = "select distinct substring(b.kode,2,2)'kode',b.ket from am_itemcode b where b.lev='2' and substring(b.kode,1,1)='" & l1 & "' order by substring(b.kode,2,2)"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        l2.AddItem RST!kode
        l2.List(i, 1) = RST!ket
        i = i + 1
        RST.MoveNext
    Loop
    OBJ.Close
End Sub

Private Sub l1_DropButtonClick()
    OBJ.Open dsn
    SQL = "select b.ket from am_itemcode b where b.lev='1' and substring(b.kode,1,1)='" & l1 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txt1 = RST!ket
    Else
        txt1 = ""
    End If
    OBJ.Close
End Sub

Private Sub l1_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub l1_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
End Sub

Private Sub l2_Change()
    l4.Clear
    l4.ColumnCount = 2
    l4.ListWidth = "6 cm"
    l4.ColumnWidths = "2 cm; 4 cm"
    l5.Clear
    l5.ColumnCount = 2
    l5.ListWidth = "6 cm"
    l5.ColumnWidths = "2 cm; 4 cm"
    l6.Clear
    l6.ColumnCount = 2
    l6.ListWidth = "6 cm"
    l6.ColumnWidths = "2 cm; 4 cm"
    l7.Clear
    l7.ColumnCount = 2
    l7.ListWidth = "6 cm"
    l7.ColumnWidths = "2 cm; 4 cm"
    l8.Clear
    l8.ColumnCount = 2
    l8.ListWidth = "6 cm"
    l8.ColumnWidths = "2 cm; 4 cm"
    txt4 = ""
    txt5 = ""
    txt6 = ""
    txt7 = ""
    txt8 = ""
    
    If l1 <> "" Then
        i = 0
        OBJ.Open dsn
        If l1 = "L" Then SQL = "select distinct substring(b.kode,2,2)'kode',b.ket from am_itemcode b where b.lev='4' and substring(b.kode,1,1)='L' order by substring(b.kode,2,2)"
        If l1 = "K" Then SQL = "select distinct substring(b.kode,4,2)'kode',b.ket from am_itemcode b where b.lev='4' and substring(b.kode,1,1)='K' and substring(b.kode,2,2) = '" & l2 & "' order by substring(b.kode,4,2)"
        If l1 = "W" Then SQL = "select distinct substring(b.kode,4,2)'kode',b.ket from am_itemcode b where b.lev='4' and substring(b.kode,1,1)='W' and substring(b.kode,2,2) = '" & l2 & "' order by substring(b.kode,4,2)"
        If l1 = "R" Then SQL = "select distinct substring(b.kode,4,2)'kode',b.ket from am_itemcode b where b.lev='4' and substring(b.kode,1,1)='R' and substring(b.kode,2,2) = '" & l2 & "' order by substring(b.kode,4,2)"
        Set RST = OBJ.Execute(SQL)
        Do While Not RST.EOF
            l4.AddItem RST!kode
            l4.List(i, 1) = RST!ket
            l5.AddItem RST!kode
            l5.List(i, 1) = RST!ket
            l6.AddItem RST!kode
            l6.List(i, 1) = RST!ket
            l7.AddItem RST!kode
            l7.List(i, 1) = RST!ket
            l8.AddItem RST!kode
            l8.List(i, 1) = RST!ket
            
            i = i + 1
            RST.MoveNext
        Loop
        OBJ.Close
    End If
End Sub

Private Sub l2_DropButtonClick()
    OBJ.Open dsn
    SQL = "select b.ket from am_itemcode b where b.lev='2' and substring(b.kode,1,3)='" & l1 & l2 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        txt2 = RST!ket
    Else
        txt2 = ""
    End If
    OBJ.Close
End Sub

Private Sub l2_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
End Sub

Private Sub l4_DropButtonClick()
    If l1 <> "" Then
        OBJ.Open dsn
        If l1 = "L" Then SQL = "select b.ket from am_itemcode b where b.lev='4' and substring(b.kode,1,1)='L' and substring(b.kode,2,2) = '" & l4 & "'"
        If l1 = "K" Then SQL = "select b.ket from am_itemcode b where b.lev='4' and substring(b.kode,1,1)='K' and substring(b.kode,2,2) = '" & l2 & "' and substring(b.kode,4,2) = '" & l4 & "'"
        If l1 = "W" Then SQL = "select b.ket from am_itemcode b where b.lev='4' and substring(b.kode,1,1)='W' and substring(b.kode,2,2) = '" & l2 & "' and substring(b.kode,4,2) = '" & l4 & "'"
        If l1 = "R" Then SQL = "select b.ket from am_itemcode b where b.lev='4' and substring(b.kode,1,1)='R' and substring(b.kode,2,2) = '" & l2 & "' and substring(b.kode,4,2) = '" & l4 & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            txt4 = RST!ket
        Else
            txt4 = ""
        End If
        OBJ.Close
    End If
End Sub

Private Sub l4_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
End Sub

Private Sub l5_DropButtonClick()
    If l1 <> "" Then
        OBJ.Open dsn
        If l1 = "L" Then SQL = "select b.ket from am_itemcode b where b.lev='4' and substring(b.kode,1,1)='L' and substring(b.kode,2,2) = '" & l5 & "'"
        If l1 = "K" Then SQL = "select b.ket from am_itemcode b where b.lev='4' and substring(b.kode,1,1)='K' and substring(b.kode,2,2) = '" & l2 & "' and substring(b.kode,4,2) = '" & l5 & "'"
        If l1 = "W" Then SQL = "select b.ket from am_itemcode b where b.lev='4' and substring(b.kode,1,1)='W' and substring(b.kode,2,2) = '" & l2 & "' and substring(b.kode,4,2) = '" & l5 & "'"
        If l1 = "R" Then SQL = "select b.ket from am_itemcode b where b.lev='4' and substring(b.kode,1,1)='R' and substring(b.kode,2,2) = '" & l2 & "' and substring(b.kode,4,2) = '" & l5 & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            txt5 = RST!ket
        Else
            txt5 = ""
        End If
        OBJ.Close
    End If
End Sub

Private Sub l5_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
End Sub

Private Sub l6_DropButtonClick()
    If l1 <> "" Then
        OBJ.Open dsn
        If l1 = "L" Then SQL = "select b.ket from am_itemcode b where b.lev='4' and substring(b.kode,1,1)='L' and substring(b.kode,2,2) = '" & l6 & "'"
        If l1 = "K" Then SQL = "select b.ket from am_itemcode b where b.lev='4' and substring(b.kode,1,1)='K' and substring(b.kode,2,2) = '" & l2 & "' and substring(b.kode,4,2) = '" & l6 & "'"
        If l1 = "W" Then SQL = "select b.ket from am_itemcode b where b.lev='4' and substring(b.kode,1,1)='W' and substring(b.kode,2,2) = '" & l2 & "' and substring(b.kode,4,2) = '" & l6 & "'"
        If l1 = "R" Then SQL = "select b.ket from am_itemcode b where b.lev='4' and substring(b.kode,1,1)='R' and substring(b.kode,2,2) = '" & l2 & "' and substring(b.kode,4,2) = '" & l6 & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            txt6 = RST!ket
        Else
            txt6 = ""
        End If
        OBJ.Close
    End If
End Sub

Private Sub l6_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
End Sub

Private Sub l7_DropButtonClick()
    If l1 <> "" Then
        OBJ.Open dsn
        If l1 = "L" Then SQL = "select b.ket from am_itemcode b where b.lev='4' and substring(b.kode,1,1)='L' and substring(b.kode,2,2) = '" & l7 & "'"
        If l1 = "K" Then SQL = "select b.ket from am_itemcode b where b.lev='4' and substring(b.kode,1,1)='K' and substring(b.kode,2,2) = '" & l2 & "' and substring(b.kode,4,2) = '" & l7 & "'"
        If l1 = "W" Then SQL = "select b.ket from am_itemcode b where b.lev='4' and substring(b.kode,1,1)='W' and substring(b.kode,2,2) = '" & l2 & "' and substring(b.kode,4,2) = '" & l7 & "'"
        If l1 = "R" Then SQL = "select b.ket from am_itemcode b where b.lev='4' and substring(b.kode,1,1)='R' and substring(b.kode,2,2) = '" & l2 & "' and substring(b.kode,4,2) = '" & l7 & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            txt7 = RST!ket
        Else
            txt7 = ""
        End If
        OBJ.Close
    End If
End Sub

Private Sub l7_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
End Sub

Private Sub l8_DropButtonClick()
    If l1 <> "" Then
        OBJ.Open dsn
        If l1 = "L" Then SQL = "select b.ket from am_itemcode b where b.lev='4' and substring(b.kode,1,1)='L' and substring(b.kode,2,2) = '" & l8 & "'"
        If l1 = "K" Then SQL = "select b.ket from am_itemcode b where b.lev='4' and substring(b.kode,1,1)='K' and substring(b.kode,2,2) = '" & l2 & "' and substring(b.kode,4,2) = '" & l8 & "'"
        If l1 = "W" Then SQL = "select b.ket from am_itemcode b where b.lev='4' and substring(b.kode,1,1)='W' and substring(b.kode,2,2) = '" & l2 & "' and substring(b.kode,4,2) = '" & l8 & "'"
        If l1 = "R" Then SQL = "select b.ket from am_itemcode b where b.lev='4' and substring(b.kode,1,1)='R' and substring(b.kode,2,2) = '" & l2 & "' and substring(b.kode,4,2) = '" & l8 & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then
            txt8 = RST!ket
        Else
            txt8 = ""
        End If
        OBJ.Close
    End If
End Sub

Private Sub l8_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
End Sub

Private Sub txtgudang_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmdview.SetFocus
End Sub

Private Sub txtgudang_LostFocus()
    If txtgudang = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from am_gudang where kodegudang = '" & txtgudang & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Data not found.", vbExclamation, "Warning"
        txtgudang = ""
        txtgudang.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub txtinv1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtinv2.SetFocus
End Sub

Private Sub txtinv2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtinv3.SetFocus
End Sub

Private Sub txtinv3_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtinv4.SetFocus
End Sub

Private Sub txtinv4_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmdview.SetFocus
End Sub

Private Sub cmdsearch1_Click()
    carisql1 = "select kodebarang, namabarang from am_itemmst"
    namatabel = "Item"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch2_Click()
    carisql1 = "select kodebarang, namabarang from am_itemmst"
    namatabel = "Item"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch3_Click()
    carisql1 = "select kodegudang, namagudang from am_gudang"
    namatabel = "Gudang"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch4_Click()
    carisql1 = "select kodegudang, namagudang from am_gudang"
    namatabel = "Gudang"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtinv1 = hasil
    txtinv2 = hasil
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdsearch2_GotFocus()
    If hasil = "" Then Exit Sub
    txtinv2 = hasil
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdsearch3_GotFocus()
    If hasil = "" Then Exit Sub
    txtinv3 = hasil
    txtinv4 = hasil
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdsearch4_GotFocus()
    If hasil = "" Then Exit Sub
    txtinv4 = hasil
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub txtinv1_LostFocus()
    cariinv1
End Sub

Private Sub txtinv2_LostFocus()
    cariinv2
End Sub

Private Sub txtinv3_LostFocus()
    cariinv3
End Sub

Private Sub txtinv4_LostFocus()
    cariinv4
End Sub
