VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmseribrowse 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Browse No Seri Faktur Pajak"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Konfiguras"
      Height          =   975
      Left            =   360
      TabIndex        =   20
      Top             =   1170
      Width           =   5310
      Begin VB.TextBox txtkodepajak2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1395
         TabIndex        =   24
         Top             =   555
         Width           =   1800
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "Save"
         Height          =   285
         Left            =   4290
         TabIndex        =   22
         Top             =   555
         Width           =   900
      End
      Begin VB.TextBox txtkodepajak1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1395
         TabIndex        =   21
         Top             =   255
         Width           =   1800
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kode Pajak 2"
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
         Left            =   75
         TabIndex        =   25
         Top             =   570
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kode Pajak 1"
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
         Left            =   75
         TabIndex        =   23
         Top             =   270
         Width           =   1215
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   0
      TabIndex        =   5
      Top             =   2295
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7223
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
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
      TabCaption(0)   =   "Reset Flag"
      TabPicture(0)   =   "frmseribrowse.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdreset"
      Tab(0).Control(1)=   "List1"
      Tab(0).Control(2)=   "txtinv3"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Print Faktur Pajak Standar"
      TabPicture(1)   =   "frmseribrowse.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lbltype"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdview1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "date3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Check2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "List2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Crystal"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin Crystal.CrystalReport Crystal 
         Left            =   0
         Top             =   3600
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2760
         ItemData        =   "frmseribrowse.frx":0038
         Left            =   120
         List            =   "frmseribrowse.frx":003A
         Style           =   1  'Checkbox
         TabIndex        =   17
         Top             =   480
         Width           =   5775
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Preview"
         Height          =   255
         Left            =   1920
         TabIndex        =   8
         Top             =   3720
         Value           =   1  'Checked
         Width           =   975
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
         Left            =   -71880
         MaxLength       =   15
         TabIndex        =   15
         Top             =   3600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2985
         ItemData        =   "frmseribrowse.frx":003C
         Left            =   -74880
         List            =   "frmseribrowse.frx":003E
         Style           =   1  'Checkbox
         TabIndex        =   14
         Top             =   480
         Width           =   5775
      End
      Begin Chameleon.chameleonButton cmdreset 
         Height          =   375
         Left            =   -74880
         TabIndex        =   6
         Top             =   3600
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Reset Flag"
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
         MICON           =   "frmseribrowse.frx":0040
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker date3 
         Height          =   285
         Left            =   1920
         TabIndex        =   7
         Top             =   3360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   619642883
         CurrentDate     =   37510
      End
      Begin Chameleon.chameleonButton cmdview1 
         Height          =   375
         Left            =   4560
         TabIndex        =   9
         Top             =   3600
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Preview/Print"
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
         MICON           =   "frmseribrowse.frx":035A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   3
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lbltype 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4080
         TabIndex        =   18
         Top             =   3600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "Tanggal Faktur Pajak"
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
         TabIndex        =   16
         Top             =   3390
         Width           =   1575
      End
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
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
      Format          =   619642883
      CurrentDate     =   37426
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   3840
      TabIndex        =   4
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
      Format          =   619642883
      CurrentDate     =   37426
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "All"
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
      Left            =   360
      TabIndex        =   1
      Top             =   630
      Value           =   -1  'True
      Width           =   495
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "With Range"
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
      Left            =   360
      TabIndex        =   2
      Top             =   870
      Width           =   1215
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   4560
      TabIndex        =   11
      Top             =   6525
      Width           =   1335
      _ExtentX        =   2355
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmseribrowse.frx":0674
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton chameleonButton1 
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      Top             =   6525
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Submit"
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
      MICON           =   "frmseribrowse.frx":098E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker date0 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   975
      _ExtentX        =   1720
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
      CustomFormat    =   "yyyy"
      Format          =   619642883
      UpDown          =   -1  'True
      CurrentDate     =   37426
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Define Tahun"
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
      Left            =   360
      TabIndex        =   19
      Top             =   270
      Width           =   1215
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "to"
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
      Left            =   3480
      TabIndex        =   12
      Top             =   870
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   2205
      Left            =   -120
      TabIndex        =   13
      Top             =   0
      Width           =   9615
   End
End
Attribute VB_Name = "frmseribrowse"
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

Dim str1, str2, str3, str4, posrow As String
Dim i As Integer

Private Sub caristock()
    List1.Clear
    List2.Clear
    
    i = 0
    If Option2.Value = True Then
        If date1 > date2 Then
            MsgBox "Invalid date range, search abort.", vbExclamation, "Error"
            Exit Sub
        End If
        
        If Year(date1) <> Year(date2) Then
            MsgBox "Pencarian gagal, range tanggal harus pada tahun yang sama.", vbExclamation, "Error"
            Exit Sub
        End If
        
        If Year(date0) <> Year(date1) Then
            MsgBox "Pencarian gagal, tahun pada range tanggal harus sama dengan define tahun.", vbExclamation, "Error"
            Exit Sub
        End If
        
        OBJ.Open dsn
        SQL = "select a.* from am_invhdr a left join am_customer b on a.kodecust=b.kodecust where b.nonpwp <> '' and a.ppn <> 0 and a.tglbkt >= '" & batas1 & "' and a.tglbkt <= '" & batas2 & "' order by a.noseri"
        Set RST = OBJ.Execute(SQL)
        Do While Not RST.EOF
            List1.AddItem RST!noseri & " - (No.Bukti = " & RST!nobkt & "; Date = " & Format(RST!tglbkt, "dd-MM-yyyy") & ")"
                        
            If RST!idupdate <> "print" Then
                List1.Selected(i) = False
                List2.AddItem RST!noseri & " - (No.Bukti = " & RST!nobkt & "; Date = " & Format(RST!tglbkt, "dd-MM-yyyy") & ")"
            Else
                List1.Selected(i) = True
            End If
            
            RST.MoveNext
        Loop
        OBJ.Close
    Else
        OBJ.Open dsn
        SQL = "select a.* from am_invhdr a left join am_customer b on a.kodecust=b.kodecust where b.nonpwp <> '' and a.ppn <> 0 and a.noseri <> '' and year(a.tglbkt) = '" & Format(date0, "yyyy") & "' order by a.noseri"
        Set RST = OBJ.Execute(SQL)
        Do While Not RST.EOF
            List1.AddItem RST!noseri & " - (No.Bukti = " & RST!nobkt & "; Date = " & Format(RST!tglbkt, "dd MMM yyyy") & ")"
            
            If RST!idupdate <> "print" Then
                List1.Selected(i) = False
                List2.AddItem RST!noseri & " - (No.Bukti = " & RST!nobkt & "; Date = " & Format(RST!tglbkt, "dd-MM-yyyy") & ")"
            Else
                List1.Selected(i) = True
            End If
            
            RST.MoveNext
            i = i + 1
        Loop
        OBJ.Close
    End If
End Sub

Function batas1()
    batas1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function batas2()
    batas2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function

Private Sub chameleonButton1_Click()
    caristock
    MsgBox "Refresh Complete.", vbInformation, "Information"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdreset_Click()
    txtinv3 = Mid(List1.text, 1, 8)
    List1.text = ""
    
    If txtinv3 = "" Then
        MsgBox "Data entry not complete.", vbInformation, "Information"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from am_invhdr where noseri = '" & txtinv3 & "' and year(tglbkt) = '" & Format(date0, "yyyy") & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If MsgBox("Reset " & txtinv3 & " ?", vbYesNo + vbQuestion, "Question") = vbNo Then
            OBJ.Close
        
            txtinv3 = ""
            Exit Sub
        End If
        SQL = "update am_invhdr set idupdate = ' ' where noseri = '" & txtinv3 & "' and year(tglbkt) = '" & Format(date0, "yyyy") & "'"
        Set RST = OBJ.Execute(SQL)
    Else
        OBJ.Close
        
        MsgBox "No seri not found.", vbInformation, "Information"
        txtinv3 = ""
        Exit Sub
    End If
    OBJ.Close
    MsgBox "Reset No seri complete.", vbInformation, "Information"
    txtinv3 = ""
    caristock
End Sub

Private Sub cmdsave_Click()
    Open AppPath + "\setting.txt" For Output As #1
        Print #1, txtkodepajak1
        Print #1, txtkodepajak2
        MsgBox "Konfigurasi telah berhasil dibuat..!", vbInformation, AppName
    Close
End Sub

Private Sub LoadKodePajak()
    Dim kodepajak1, kodepajak2 As String
    Open AppPath + "\setting.txt" For Input As #1
        Line Input #1, kodepajak1
        Line Input #1, kodepajak2
    Close
    txtkodepajak1 = kodepajak1
    txtkodepajak2 = kodepajak2
End Sub

Private Sub cmdview1_Click()
    For i = 0 To List2.ListCount - 1
        If List2.Selected(i) = True Then
            Crystal.Reset
            
            If Check2.Value = 1 Then
                Crystal.WindowState = crptMaximized
                Crystal.WindowShowCloseBtn = True
                Crystal.WindowShowPrintBtn = False
                Crystal.WindowShowSearchBtn = True
                Crystal.Destination = crptToWindow
            Else
                Crystal.Destination = crptToPrinter
            End If
            
            List2.ListIndex = i
            OBJ.Open dsn
            SQL = "select kodecur,type from am_invhdr where noseri = '" & Mid(List2.text, 1, 8) & "' and year(tglbkt) = '" & Format(date0, "yyyy") & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then
                str1 = RST!kodecur
                lbltype = RST!Type
            Else
                MsgBox "No seri mismatch, make sure the length is 8 digit.", vbInformation, "Information"
                OBJ.Close
                Exit Sub
            End If
            
            SQL = "select base from gl_kurs where kdkurs = '" & str1 & "'"
            Set RST = OBJ.Execute(SQL)
            If Not RST.EOF Then str1 = RST!base
            OBJ.Close
            
            Crystal.Connect = dsnreport
            Crystal.DataFiles(0) = "Proc(am_fakturpajak)"
            If str1 = "1" Then Crystal.ReportFileName = AppPath & "\reports\sale\inv\fakturstandar1.rpt"
            If str1 = "0" Then Crystal.ReportFileName = AppPath & "\reports\sale\inv\fakturstandar.rpt"
            Crystal.ParameterFields(0) = "@tanggal1;" & Format(date3, "yyyymmdd") & ";true"
            Crystal.ParameterFields(1) = "@noinv1;" & Mid(List2.text, 1, 8) & ";true"
            Crystal.ParameterFields(2) = "@pilih;" & "standar" & ";true"
            Crystal.ParameterFields(3) = "@pilih1;" & lbltype & ";true"
            Crystal.ParameterFields(4) = "@nourut;" & Mid(List2.text, 1, 8) & ";true"
            Crystal.ParameterFields(5) = "@kodepajak1;" & txtkodepajak1 & ";true"
            Crystal.ParameterFields(6) = "@kodepajak2;" & txtkodepajak2 & ";true"
            Crystal.RetrieveDataFiles
            Crystal.Action = 1
            
            If Check2.Value = 0 Then
                OBJ.Open dsn
                SQL = "update am_invhdr set idupdate = 'print' WHERE noseri = '" & Mid(List2.text, 1, 8) & "' and year(tglbkt) = '" & Format(date0, "yyyy") & "'"
                Set RST = OBJ.Execute(SQL)
                OBJ.Close
                
                txtinv3 = ""
                date3.Value = Date
            End If
        End If
    Next i
    If Check2.Value = 0 Then caristock
End Sub

Private Sub Form_Activate()
    LoadKodePajak
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    date0 = Date
    date1 = Date
    date2 = Date
    date3 = Date
    date1.Enabled = False
    date2.Enabled = False
End Sub

Private Sub Option1_Click()
    date1.Enabled = False
    date2.Enabled = False
End Sub

Private Sub Option2_Click()
    date1.Enabled = True
    date2.Enabled = True
End Sub
