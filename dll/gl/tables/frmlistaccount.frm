VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmlistaccount 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Company Account"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7695
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmlistaccount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Chameleon.chameleonButton cmdsearch1 
      Height          =   285
      Left            =   600
      TabIndex        =   18
      Top             =   1800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "From Account"
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
      MICON           =   "frmlistaccount.frx":2372
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Income Summary"
      Enabled         =   0   'False
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
      Left            =   1800
      TabIndex        =   8
      Top             =   3240
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.TextBox txtcomp 
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
      Left            =   1920
      MaxLength       =   4
      TabIndex        =   0
      Top             =   1200
      Width           =   615
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Expenses"
      Enabled         =   0   'False
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
      Left            =   1800
      TabIndex        =   7
      Top             =   3000
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Income"
      Enabled         =   0   'False
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
      Left            =   1800
      TabIndex        =   6
      Top             =   2760
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Capital"
      Enabled         =   0   'False
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
      Left            =   720
      TabIndex        =   5
      Top             =   3240
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Liability"
      Enabled         =   0   'False
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
      Left            =   720
      TabIndex        =   4
      Top             =   3000
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Assets"
      Enabled         =   0   'False
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
      Left            =   720
      TabIndex        =   3
      Top             =   2760
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.OptionButton opskode 
      Caption         =   "No. Account"
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
      TabIndex        =   17
      Top             =   1560
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.OptionButton opstype 
      Caption         =   "Type Account"
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
      TabIndex        =   16
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox txtarea1 
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
      Left            =   1920
      MaxLength       =   15
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtarea2 
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
      Left            =   1920
      MaxLength       =   15
      TabIndex        =   2
      Top             =   2160
      Width           =   1335
   End
   Begin Crystal.CrystalReport crystal 
      Left            =   120
      Top             =   4560
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
   Begin Chameleon.chameleonButton cmdview 
      Height          =   375
      Left            =   5520
      TabIndex        =   9
      Top             =   3240
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
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
      MICON           =   "frmlistaccount.frx":268C
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
      Left            =   6480
      TabIndex        =   10
      Top             =   3240
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
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
      MICON           =   "frmlistaccount.frx":29A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch2 
      Height          =   285
      Left            =   600
      TabIndex        =   19
      Top             =   2160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "To Account"
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
      MICON           =   "frmlistaccount.frx":2CC0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch 
      Height          =   285
      Left            =   480
      TabIndex        =   20
      Top             =   1200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Company Code"
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
      MICON           =   "frmlistaccount.frx":2FDA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblcomp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2640
      TabIndex        =   21
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Company Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   0
      TabIndex        =   13
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label lblnamarea2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3360
      TabIndex        =   12
      Top             =   2160
      Width           =   3975
   End
   Begin VB.Label lblnamarea1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3360
      TabIndex        =   11
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   10335
   End
End
Attribute VB_Name = "frmlistaccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim str1, str2, str3, str4, str5, str6, str7 As String

Private Sub cariarea1()
    If txtarea1 = "" Or txtcomp = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select b.noac, b.nmac, b.typeac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtcomp & "' and a.noac = '" & x_original(txtarea1) & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Account " & txtarea1 & " Not Found.", vbExclamation, "Warning"
        txtarea1 = ""
        lblnamarea1 = ""
        txtarea1.SetFocus
    Else
        lblnamarea1 = RST!nmac
    End If
    OBJ.Close
End Sub

Private Sub cariarea2()
    If txtarea2 = "" Or txtcomp = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select b.noac, b.nmac, b.typeac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtcomp & "' and a.noac = '" & x_original(txtarea2) & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Account " & txtarea2 & " Not Found.", vbExclamation, "Warning"
        txtarea2 = ""
        lblnamarea2 = ""
        txtarea2.SetFocus
    Else
        lblnamarea2 = RST!nmac
    End If
    OBJ.Close
End Sub

Private Sub cariarea3()
    If txtcomp = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from gl_company where kdcomp = '" & txtcomp & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Company " & txtcomp & " Not Found.", vbExclamation, "Warning"
        txtcomp = ""
        lblcomp = ""
        txtcomp.SetFocus
    Else
        lblcomp = RST!nmcompscr
    End If
    OBJ.Close
    
    If opskode.Value = True Then
        txtarea1 = ""
        txtarea2 = ""
        lblnamarea1 = ""
        lblnamarea2 = ""
    Else
        Check1.Value = 1
        Check2.Value = 1
        Check3.Value = 1
        Check4.Value = 1
        Check5.Value = 1
        Check6.Value = 1
    End If
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdsearch_Click()
    carisql1 = "select kdcomp, nmcompscr from gl_company"
    namatabel = "Company"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    txtcomp = hasil
    txtcomp_LostFocus
    hasil = ""
End Sub

Private Sub cmdview_Click()
    If txtcomp = "" Then Exit Sub
    If opskode.Value = True Then
        If txtarea1 = "" Then txtarea1 = "0"
        If txtarea2 = "" Then txtarea2 = "z"
        If txtarea1 <> "" Then txtarea1 = x_original(txtarea1)
        If txtarea2 <> "" Then txtarea2 = x_original(txtarea2)
        
        If txtarea2 < txtarea1 Then
            MsgBox "To Acc Can Not Smaller Then From Acc.", vbExclamation, "Warning"
            txtarea2 = ""
            lblnamarea2 = ""
            txtarea2.SetFocus
            Exit Sub
        End If
        str1 = "no"
    Else
        If Check1.Value = 0 And Check2.Value = 0 And Check3.Value = 0 And Check4.Value = 0 And Check5.Value = 0 And Check6.Value = 0 Then Exit Sub
        str1 = "type"
    End If
    
    If Check1.Value = 1 Then
        str2 = "1"
    Else
        str2 = "0"
    End If
    If Check2.Value = 1 Then
        str3 = "1"
    Else
        str3 = "0"
    End If
    If Check3.Value = 1 Then
        str4 = "1"
    Else
        str4 = "0"
    End If
    If Check4.Value = 1 Then
        str5 = "1"
    Else
        str5 = "0"
    End If
    If Check5.Value = 1 Then
        str6 = "1"
    Else
        str6 = "0"
    End If
    If Check6.Value = 1 Then
        str7 = "1"
    Else
        str7 = "0"
    End If
    
    crystal.Reset
    crystal.WindowState = crptMaximized
    crystal.WindowShowCloseBtn = True
    crystal.WindowShowPrintSetupBtn = True
    crystal.WindowShowSearchBtn = True
    crystal.WindowShowRefreshBtn = True
    crystal.Connect = dsnreport
    crystal.DataFiles(0) = "Proc(gl_accountlist)"
    If opskode.Value = True Then crystal.ReportFileName = AppPath & "\reports\gl\tables\accountlist1.rpt"
    If opstype.Value = True Then crystal.ReportFileName = AppPath & "\reports\gl\tables\accountlist.rpt"
    crystal.ParameterFields(0) = "@comp;" + txtcomp + ";true"
    crystal.ParameterFields(1) = "@flag;" + str1 + ";true"
    crystal.ParameterFields(2) = "@type1;" + str2 + ";true"
    crystal.ParameterFields(3) = "@type2;" + str3 + ";true"
    crystal.ParameterFields(4) = "@type3;" + str4 + ";true"
    crystal.ParameterFields(5) = "@type4;" + str5 + ";true"
    crystal.ParameterFields(6) = "@type5;" + str6 + ";true"
    crystal.ParameterFields(7) = "@type6;" + str7 + ";true"
    crystal.ParameterFields(8) = "@area1;" + txtarea1 + ";true"
    crystal.ParameterFields(9) = "@area2;" + txtarea2 + ";true"
    crystal.ParameterFields(10) = "@namauser;" + nmuser + ";true"
    crystal.RetrieveDataFiles
    crystal.Action = 1
    
    If txtarea1 = "0" Then txtarea1 = ""
    If txtarea2 = "z" Then txtarea2 = ""
    If txtarea1 <> "" Then txtarea1 = original(txtarea1)
    If txtarea2 <> "" Then txtarea2 = original(txtarea2)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub opskode_Click()
    Check1.Enabled = False
    Check2.Enabled = False
    Check3.Enabled = False
    Check4.Enabled = False
    Check5.Enabled = False
    Check6.Enabled = False
    
    cmdsearch1.Enabled = True
    cmdsearch2.Enabled = True
    txtarea1.Enabled = True
    txtarea2.Enabled = True
    txtarea1 = ""
    txtarea2 = ""
    lblnamarea1 = ""
    lblnamarea2 = ""
End Sub

Private Sub opstype_Click()
    Check1.Enabled = True
    Check2.Enabled = True
    Check3.Enabled = True
    Check4.Enabled = True
    Check5.Enabled = True
    Check6.Enabled = True
    
    cmdsearch1.Enabled = False
    cmdsearch2.Enabled = False
    txtarea1.Enabled = False
    txtarea2.Enabled = False
    txtarea1 = ""
    txtarea2 = ""
    lblnamarea1 = ""
    lblnamarea2 = ""
End Sub

Private Sub txtarea1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtarea2.SetFocus
End Sub

Private Sub txtarea1_LostFocus()
    cariarea1
End Sub

Private Sub txtarea2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmdview.SetFocus
End Sub

Private Sub txtarea2_LostFocus()
    cariarea2
End Sub

Private Sub cmdsearch1_Click()
    carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtcomp & "'"
    namatabel = "Company Account"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtarea1 = hasil
    lblnamarea1 = hasil1
    hasil = ""
    hasil1 = ""
    txtarea2.SetFocus
End Sub

Private Sub cmdsearch2_Click()
    carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp = '" & txtcomp & "'"
    namatabel = "Company Account"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch2_GotFocus()
    If hasil = "" Then Exit Sub
    txtarea2 = hasil
    lblnamarea2 = hasil1
    hasil = ""
    hasil1 = ""
    cmdview.SetFocus
End Sub

Private Sub txtcomp_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtarea1.SetFocus
End Sub

Private Sub txtcomp_LostFocus()
    cariarea3
End Sub
