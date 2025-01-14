VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Begin VB.Form frmclose 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5040
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6735
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
   Icon            =   "frmclose.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1320
      MaxLength       =   4
      TabIndex        =   14
      Top             =   1200
      Width           =   735
   End
   Begin MSComCtl2.DTPicker date3 
      Height          =   255
      Left            =   960
      TabIndex        =   13
      Top             =   5880
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Format          =   106233857
      CurrentDate     =   37767
   End
   Begin MSComCtl2.DTPicker date4 
      Height          =   255
      Left            =   2640
      TabIndex        =   12
      Top             =   5880
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Format          =   106233857
      CurrentDate     =   37767
   End
   Begin VB.Frame Frame3 
      Caption         =   "Periode After Closing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   360
      TabIndex        =   7
      Top             =   2880
      Width           =   6015
      Begin MSComCtl2.DTPicker date2 
         Height          =   285
         Left            =   2040
         TabIndex        =   2
         Top             =   1080
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
         Format          =   106233859
         CurrentDate     =   37694
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   285
         Left            =   2040
         TabIndex        =   1
         Top             =   720
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
         Format          =   106233859
         CurrentDate     =   37694
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   2400
         TabIndex        =   8
         Top             =   360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtperiode"
         BuddyDispid     =   196617
         OrigLeft        =   2655
         OrigTop         =   360
         OrigRight       =   2895
         OrigBottom      =   645
         Max             =   13
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   "Tanggal Akhir Proses"
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
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label txtperiode 
         Alignment       =   2  'Center
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
         Left            =   2040
         TabIndex        =   0
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         Caption         =   "Tanggal Awal Proses"
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
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   750
         Width           =   1815
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         Caption         =   "Periode On Proses"
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
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   390
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Periode To Closing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   3
      Top             =   1680
      Width           =   6015
      Begin VB.Label lblakhir 
         Caption         =   "Tanggal Akhir Proces:"
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
         TabIndex        =   6
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label lblawal 
         Caption         =   "Tanggal Awal Proces :"
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
         TabIndex        =   5
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label lblperiode 
         Caption         =   "Periode On Proces :"
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
         TabIndex        =   4
         Top             =   240
         Width           =   3735
      End
   End
   Begin Chameleon.chameleonButton cmdsubmit 
      Height          =   375
      Left            =   4320
      TabIndex        =   16
      Top             =   4560
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmclose.frx":2372
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
      Left            =   5280
      TabIndex        =   17
      Top             =   4560
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
      MICON           =   "frmclose.frx":268C
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
      TabIndex        =   21
      Top             =   1200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Company"
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
      MICON           =   "frmclose.frx":29A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Closing"
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
      TabIndex        =   19
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction"
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
      TabIndex        =   18
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label lblnama 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2160
      TabIndex        =   15
      Top             =   1200
      Width           =   4215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   10335
   End
End
Attribute VB_Name = "frmclose"
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
Dim vsp(7) As Variant

Dim str1, str2, str3, str5 As String

Private Sub cariarea1()
    If txtarea1 = "" Then Exit Sub
    lblnama = ""
    lblperiode = "Periode On Proces : "
    lblawal = "Tanggal Awal Proces : "
    lblakhir = "Tanggal Akhir Proces : "
    txtperiode = 1
    date1 = Date
    date2 = Date
    
    OBJ.Open dsn
    SQL = "select * from gl_company where kdcomp = '" & txtarea1 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        lblnama = RST!nmcompscr
        lblperiode = "Periode On Proces : " & RST!periode
        lblawal = "Tanggal Awal Proces : " & Format(RST!tglawal, "dd MMMM yyyy")
        lblakhir = "Tanggal Akhir Proces : " & Format(RST!tglakhir, "dd MMMM yyyy")
        str1 = Format(RST!tglakhir, "MM/dd/yyyy")
        str2 = RST!periode
        date3 = RST!tglawal
        date4 = RST!tglakhir
        
        If (Val(str2) >= "1" And Val(str2) <= "12") Then
            If date4.Month = 12 Then
                If date4.Day = 31 Then
                    txtperiode = 1
                Else
                    txtperiode = Val(str2) + 1
                End If
            Else
                txtperiode = Val(str2) + 1
            End If
        ElseIf (date4.Month = 12 And date4.Day = 31) Or Val(str2) = "13" Then
            txtperiode = 1
        End If
        date1 = date4 + 1
        date2.Day = 28
        date2.Month = date1.Month
        date2.Year = date1.Year
        date2 = date2 + 4
        date2.Day = 1
        date2 = date2 - 1
    Else
        MsgBox "Company " & txtarea1 & " Not Found.", vbExclamation, "Warning"
        txtarea1 = ""
        txtarea1.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    date1 = Date
    date2 = Date
    txtperiode = 1
    
    OBJ.Open dsn
    SQL = "insert into toogle"
    SQL = SQL + "(comp_id"
    SQL = SQL + ",task)"

    SQL = SQL + "VALUES"
    SQL = SQL + "('" & GetTheComputerName & "'"
    SQL = SQL + ",'Closing')"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
End Sub

Private Sub CmdSubmit_Click()
    If txtarea1 = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from toogle"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        If RST!comp_id <> GetTheComputerName Then
            MsgBox "Access denied" & vbCrLf & _
            "Computer name : " & RST!comp_id & vbCrLf & _
            "Task : " & RST!task, vbExclamation, "Error"
            OBJ.Close
            Unload Me
            Exit Sub
        End If
        
        RST.MoveNext
    Loop
    OBJ.Close
    
    If MsgBox("Make Sure The Transaction Is Posted." & vbCrLf & _
    "Continue Closing ?", vbYesNo + vbQuestion, "Question") = vbNo Then Exit Sub
    
    If MsgBox("Accept This Periode After Closing ?", vbYesNo + vbQuestion, "Question") = vbNo Then Exit Sub
    
    If date1 <= date4 Then
        MsgBox "Periode After Process Can't Smaller Then Periode On Process.", vbInformation, "Information"
        Exit Sub
    End If
    
    If date2 <= date1 Then
        MsgBox "Periode Akhir After Process Can't Smaller Then Periode Awal After Process.", vbInformation, "Information"
        Exit Sub
    End If
    
    If Val(txtperiode) <= Val(str2) And date1.Year = date4.Year Then
        MsgBox "Periode After Process Can't Smaller Then Periode On Process.", vbInformation, "Information"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select * from gl_accrl"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Set Account Laba/Rugi Is Empty.", vbInformation, "Information"
        OBJ.Close
        Exit Sub
    Else
        If RST!rl_ptd = "" Or RST!rl_ytd = "" Then
            MsgBox "Set Account Laba/Rugi Is Empty.", vbInformation, "Information"
            OBJ.Close
            Exit Sub
        End If
    End If
    
    SQL = "select * from gl_chacct where typeac = 'IS'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Income Summary Not Found.", vbInformation, "Information"
        OBJ.Close
        Exit Sub
    End If
    
    SQL = "select * from gl_accrl"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        If MsgBox(RST!rl_ptd & " as Period to Date" & vbCrLf & _
        RST!rl_ytd & " as Year to Date" & vbCrLf & _
        "Confirm ?", vbQuestion + vbYesNo, "Confirmation") = vbNo Then
            OBJ.Close
            Exit Sub
        End If
    End If
    
    SQL = "select * from gl_company where kdcomp = '" & txtarea1 & "'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ1.Open dsn
        SQL1 = "select * from gl_transaksi where tgltrx >= '" & Month(RST!tglawal) & "/" & Day(RST!tglawal) & "/" & Year(RST!tglawal) & "' and tgltrx <= '" & Month(RST!tglakhir) & "/" & Day(RST!tglakhir) & "/" & Year(RST!tglakhir) & "' and kdcomp= '" & txtarea1 & "' and flag <> 'P'"
        Set RST1 = OBJ1.Execute(SQL1)
        If Not RST1.EOF Then
            MsgBox "Closing Will Not Continue Until The Transaction Are Posted", vbInformation, "Information"
            OBJ.Close
            OBJ1.Close
            Exit Sub
        End If
        OBJ1.Close
    End If
        
    If (Val(str2) >= "1" And Val(str2) <= "12") Then
        If date4.Month = 12 Then
            If date4.Day = 31 Then
                If MsgBox("Continue End Year Proces ?", vbYesNo + vbQuestion, "Question") = vbNo Then
                    OBJ.Close
                    Exit Sub
                End If
                SP.ActiveConnection = dsn
                SP.CommandType = adCmdStoredProc
                SP.CommandText = "gl_endyear"
                vsp(0) = txtarea1
                vsp(1) = str2
                vsp(2) = txtperiode
                vsp(3) = str1
                vsp(4) = tanggal1
                vsp(5) = tanggal2
                vsp(6) = kuser
                vsp(7) = date4.Year
                SP.Execute , vsp
                Set SP = Nothing
            Else
                If MsgBox("Continue End Month Proces ?", vbYesNo + vbQuestion, "Question") = vbNo Then
                    OBJ.Close
                    Exit Sub
                End If
                SP.ActiveConnection = dsn
                SP.CommandType = adCmdStoredProc
                SP.CommandText = "gl_endmonth"
                vsp(0) = txtarea1
                vsp(1) = str2
                vsp(2) = txtperiode
                vsp(3) = str1
                vsp(4) = tanggal1
                vsp(5) = tanggal2
                vsp(6) = kuser
                vsp(7) = date4.Year
                SP.Execute , vsp
                Set SP = Nothing
            End If
        Else
            If MsgBox("Continue End Month Proces ?", vbYesNo + vbQuestion, "Question") = vbNo Then
                OBJ.Close
                Exit Sub
            End If
            SP.ActiveConnection = dsn
            SP.CommandType = adCmdStoredProc
            SP.CommandText = "gl_endmonth"
            vsp(0) = txtarea1
            vsp(1) = str2
            vsp(2) = txtperiode
            vsp(3) = str1
            vsp(4) = tanggal1
            vsp(5) = tanggal2
            vsp(6) = kuser
            vsp(7) = date4.Year
            SP.Execute , vsp
            Set SP = Nothing
        End If
    ElseIf Val(str2) = "13" Then
        If MsgBox("Continue End Year Proces ?", vbYesNo + vbQuestion, "Question") = vbNo Then
            OBJ.Close
            Exit Sub
        End If
        SP.ActiveConnection = dsn
        SP.CommandType = adCmdStoredProc
        SP.CommandText = "gl_endyear"
        vsp(0) = txtarea1
        vsp(1) = str2
        vsp(2) = txtperiode
        vsp(3) = str1
        vsp(4) = tanggal1
        vsp(5) = tanggal2
        vsp(6) = kuser
        vsp(7) = date4.Year
        SP.Execute , vsp
        Set SP = Nothing
    End If
    OBJ.Close
    MsgBox "Closing Complete", vbInformation, "Information"
    Unload Me
End Sub

Function tanggal1()
    tanggal1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function

Function tanggal2()
    tanggal2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function

Private Sub Form_Unload(Cancel As Integer)
    OBJ.Open dsn
    SQL = "delete from toogle where comp_id = '" & GetTheComputerName & "' and task = 'Closing'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
End Sub

Private Sub txtarea1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmdsubmit.SetFocus
End Sub

Private Sub txtarea1_LostFocus()
    cariarea1
End Sub

Private Sub cmdsearch1_Click()
    carisql1 = "select kdcomp, nmcompscr from gl_company"
    namatabel = "Company"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtarea1 = hasil
    hasil = ""
    hasil1 = ""
    cariarea1
End Sub

