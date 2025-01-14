VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmbalancetrial 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6375
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
   Icon            =   "frmbalancetrial.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtacc1 
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
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox txtacc2 
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
      Left            =   4560
      MaxLength       =   15
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CheckBox chkxls 
      Caption         =   "Eksport to Excel"
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
      Left            =   1680
      TabIndex        =   6
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CheckBox chksaldo 
      Caption         =   "Tampilkan Account Saldo 0"
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
      Left            =   1680
      TabIndex        =   4
      Top             =   2040
      Width           =   2535
   End
   Begin VB.CheckBox chkakum 
      Caption         =   "Pendapatan dan Biaya Di akumulasi"
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
      Left            =   1680
      TabIndex        =   5
      Top             =   2280
      Width           =   3135
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
      Left            =   1680
      MaxLength       =   4
      TabIndex        =   0
      Top             =   1200
      Width           =   735
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
      Left            =   4560
      MaxLength       =   4
      TabIndex        =   1
      Top             =   1200
      Width           =   735
   End
   Begin Crystal.CrystalReport crystal 
      Left            =   720
      Top             =   3840
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
      Left            =   4200
      TabIndex        =   7
      Top             =   2880
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
      MICON           =   "frmbalancetrial.frx":2372
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
      Left            =   5160
      TabIndex        =   8
      Top             =   2880
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
      MICON           =   "frmbalancetrial.frx":268C
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
      Left            =   360
      TabIndex        =   12
      Top             =   1200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "From Company"
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
      MICON           =   "frmbalancetrial.frx":29A6
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
      Left            =   3240
      TabIndex        =   13
      Top             =   1200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "To Company"
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
      MICON           =   "frmbalancetrial.frx":2CC0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch3 
      Height          =   285
      Left            =   360
      TabIndex        =   14
      Top             =   1560
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
      MICON           =   "frmbalancetrial.frx":2FDA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch4 
      Height          =   285
      Left            =   3240
      TabIndex        =   15
      Top             =   1560
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
      MICON           =   "frmbalancetrial.frx":32F4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Balance"
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
      TabIndex        =   10
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Trial"
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
      TabIndex        =   9
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   10335
   End
End
Attribute VB_Name = "frmbalancetrial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim SP As New ADODB.Command
Dim vsp(9) As Variant

Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet

Dim str1, str2, str3, str4, str5, str6, str7 As String
Dim i As Integer

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdsearch3_Click()
    If txtarea1 = "" Or txtarea2 = "" Then Exit Sub
    
    carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp >= '" & txtarea1 & "' and a.kdcomp <= '" & txtarea2 & "'"
    If txtarea1 = txtarea2 Then
        namatabel = "Company Account "
    Else
        namatabel = "Company Account  "
    End If
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch3_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc1 = hasil
    hasil = ""
    hasil1 = ""
    txtacc2.SetFocus
End Sub

Private Sub cmdsearch4_Click()
    If txtarea1 = "" Or txtarea2 = "" Then Exit Sub
    
    carisql1 = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp >= '" & txtarea1 & "' and a.kdcomp <= '" & txtarea2 & "'"
    If txtarea1 = txtarea2 Then
        namatabel = "Company Account "
    Else
        namatabel = "Company Account  "
    End If
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch4_GotFocus()
    If hasil = "" Then Exit Sub
    txtacc2 = hasil
    hasil = ""
    hasil1 = ""
    cmdview.SetFocus
End Sub

Private Sub cmdview_Click()
    If txtarea1 = "" Or txtarea2 = "" Then
        MsgBox "Data entry not complite.", vbInformation, "Information"
        Exit Sub
    End If
    If txtacc1 = "" Then txtacc1 = "0"
    If txtacc2 = "" Then txtacc2 = "z"
    If txtacc1 <> "" Then txtacc1 = x_original(txtacc1)
    If txtacc2 <> "" Then txtacc2 = x_original(txtacc2)
    
    If txtarea2 < txtarea1 Then
        MsgBox "To Company Can Not Smaller Then From Company.", vbExclamation, "Warning"
        txtarea2 = ""
        txtarea2.SetFocus
        Exit Sub
    End If
    If txtacc2 < txtacc1 Then
        MsgBox "To Account Can Not Smaller Then From Account.", vbExclamation, "Warning"
        txtacc2 = ""
        txtacc2.SetFocus
        Exit Sub
    End If
    
    If txtarea1 = txtarea2 Then
        str1 = "noall"
    Else
        str1 = "all"
    End If
    
    If chksaldo.Value = 0 Then
        str2 = "saldo"
    Else
        str2 = "semua"
    End If
    
    If chkakum.Value = 0 Then
        str3 = "tidak"
    Else
        str3 = "ya"
    End If
    
    OBJ.Open dsn
    SQL = "select * from gl_company where kdcomp = '" & txtarea1 & "'"
    Set RST = OBJ.Execute(SQL)
    str4 = RST!periode
    
    SQL = "select * from gl_accrl"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        str5 = RST!rl_ptd
    Else
        MsgBox "Set Account Laba/Rugi Is Empty.", vbInformation, "Information"
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    If chkxls.Value = 0 Then
        crystal.Reset
        crystal.WindowState = crptMaximized
        crystal.WindowShowCloseBtn = True
        crystal.WindowShowPrintSetupBtn = True
        crystal.WindowShowSearchBtn = True
        crystal.WindowShowRefreshBtn = True
        crystal.Connect = dsnreport
        crystal.DataFiles(0) = "Proc(gl_trialbalance)"
        crystal.ReportFileName = AppPath & "\reports\gl\report\trialbalance.rpt"
        crystal.ParameterFields(0) = "@com1;" + txtarea1 + ";true"
        crystal.ParameterFields(1) = "@com2;" + txtarea2 + ";true"
        crystal.ParameterFields(2) = "@acc1;" + txtacc1 + ";true"
        crystal.ParameterFields(3) = "@acc2;" + txtacc2 + ";true"
        crystal.ParameterFields(4) = "@pilih1;" + str1 + ";true"
        crystal.ParameterFields(5) = "@pilih2;" + str2 + ";true"
        crystal.ParameterFields(6) = "@pilih3;" + str3 + ";true"
        crystal.ParameterFields(7) = "@pilih4;" + str4 + ";true"
        crystal.ParameterFields(8) = "@pilih5;" + str5 + ";true"
        crystal.ParameterFields(9) = "@namauser;" + nmuser + ";true"
        crystal.RetrieveDataFiles
        crystal.Action = 1
        
        If txtarea1 = "0" Then txtarea1 = ""
        If txtarea2 = "z" Then txtarea2 = ""
    Else
        SP.ActiveConnection = dsn
        SP.CommandType = adCmdStoredProc
        SP.CommandText = "gl_trialbalancexcel"
        vsp(0) = txtarea1
        vsp(1) = txtarea2
        vsp(2) = txtacc1
        vsp(3) = txtacc2
        vsp(4) = str1
        vsp(5) = str2
        vsp(6) = str3
        vsp(7) = str4
        vsp(8) = str5
        vsp(9) = nmuser
        SP.Execute , vsp
        Set SP = Nothing
        
        OBJ.Open dsn
        If str2 = "saldo" Then
            SQL = "select * from gl_tbalance where s_awal<>0 or debet<>0 or credit<>0 order by no_acc"
        Else
            SQL = "select * from gl_tbalance order by no_acc"
        End If
        Set RST = OBJ.Execute(SQL)
            
        Set xlApp = CreateObject("Excel.Application")
        xlApp.Visible = True
        Set xlBook = xlApp.Workbooks.Add
        Set xlSheet = xlBook.ActiveSheet
        
        xlSheet.Columns.AutoFit
        
        xlSheet.Cells(1, 1).Value = "Trial Balance"
        If txtarea1 = txtarea2 Then xlSheet.Cells(2, 1).Value = "Company : " & txtarea1 & " - " & str6
        If txtarea1 <> txtarea2 Then xlSheet.Cells(2, 1).Value = "Company : " & txtarea1 & " S/D " & txtarea2 & " (Konsolidasi)"
        xlSheet.Cells(3, 1).Value = "Periode " & str7
        
        For i = 0 To RST.Fields.Count - 1 Step 1
            xlSheet.Cells(5, i + 1).Value = RST.Fields(i).Name
        Next i
         
        xlSheet.Range("A6").CopyFromRecordset RST
                
        SQL = "delete from gl_tbalance"
        Set RST = OBJ.Execute(SQL)
        OBJ.Close
        
        xlBook.SaveAs "c:\x_trial.xls"
        
        xlBook.Close False
        xlApp.Quit
        Set xlSheet = Nothing
        Set xlBook = Nothing
        Set xlApp = Nothing
        
        Unload Me
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtacc1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtacc2.SetFocus
End Sub

Private Sub txtacc1_LostFocus()
    cariacc1
End Sub

Private Sub txtacc2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmdview.SetFocus
End Sub

Private Sub txtacc2_LostFocus()
    cariacc2
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
    If KeyAscii = 13 Then txtacc1.SetFocus
End Sub

Private Sub txtarea2_LostFocus()
    cariarea2
End Sub

Private Sub cmdsearch1_Click()
    carisql1 = "select kdcomp, nmcompscr from gl_company"
    namatabel = "Company"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtarea1 = hasil
    str6 = hasil1
    hasil = ""
    hasil1 = ""
    txtarea2.SetFocus
    
    OBJ.Open dsn
    SQL = "select * from gl_company where kdcomp = '" & txtarea1 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        str7 = Format(RST!tglakhir, "MMMM") & " " & Format(RST!tglakhir, "yyyy")
    End If
    OBJ.Close
End Sub

Private Sub cmdsearch2_Click()
    carisql1 = "select kdcomp, nmcompscr from gl_company"
    namatabel = "Company"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch2_GotFocus()
    If hasil = "" Then Exit Sub
    txtarea2 = hasil
    hasil = ""
    hasil1 = ""
    cmdview.SetFocus
End Sub

Private Sub cariacc1()
    If txtacc1 = "" Then Exit Sub
    If txtarea1 = "" Or txtarea2 = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp >= '" & txtarea1 & "' and a.kdcomp <= '" & txtarea2 & "' and b.noac = '" & x_original(txtacc1) & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Account " & txtacc1 & " Not Found.", vbExclamation, "Warning"
        txtacc1 = ""
        txtacc1.SetFocus
    Else
        If txtarea1 = txtarea2 Then txtacc1 = original(RST!noac)
    End If
    OBJ.Close
End Sub

Private Sub cariacc2()
    If txtacc2 = "" Then Exit Sub
    If txtarea1 = "" Or txtarea2 = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select b.noac, b.nmac from gl_chacct a left join gl_masterac b on a.noac = b.noac where a.kdcomp >= '" & txtarea1 & "' and a.kdcomp <= '" & txtarea2 & "' and b.noac = '" & x_original(txtacc2) & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Account " & txtacc2 & " Not Found.", vbExclamation, "Warning"
        txtacc2 = ""
        txtacc2.SetFocus
    Else
        If txtarea1 = txtarea2 Then txtacc2 = original(RST!noac)
    End If
    OBJ.Close
End Sub

Private Sub cariarea1()
    If txtarea1 = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from gl_company where kdcomp = '" & txtarea1 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Company " & txtarea1 & " Not Found.", vbExclamation, "Warning"
        txtarea1 = ""
        txtarea1.SetFocus
    Else
        str6 = RST!nmcompscr
        str7 = Format(RST!tglakhir, "MMMM") & " " & Format(RST!tglakhir, "yyyy")
    End If
    OBJ.Close
End Sub

Private Sub cariarea2()
    If txtarea2 = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from gl_company where kdcomp = '" & txtarea2 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Company " & txtarea2 & " Not Found.", vbExclamation, "Warning"
        txtarea2 = ""
        txtarea2.SetFocus
    End If
    OBJ.Close
End Sub
