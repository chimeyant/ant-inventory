VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmpostingjfa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Posting Penjualan Fixed Assets"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7815
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmpostingjfa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker date3 
      Height          =   255
      Left            =   6480
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Format          =   193724417
      CurrentDate     =   37749
   End
   Begin VB.TextBox txtkodefa2 
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
      MaxLength       =   8
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox txtkodefa1 
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
      MaxLength       =   8
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox txtcom1 
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
      TabIndex        =   0
      Top             =   1200
      Width           =   855
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   2520
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
      Format          =   193724419
      CurrentDate     =   37694
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   5040
      TabIndex        =   4
      Top             =   2520
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
      Format          =   193724419
      CurrentDate     =   37694
   End
   Begin Chameleon.chameleonButton cmdsearch1 
      Height          =   285
      Left            =   360
      TabIndex        =   16
      Top             =   1200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Kode Company"
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
      MICON           =   "frmpostingjfa.frx":2372
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
      Left            =   360
      TabIndex        =   17
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "From Aktiva"
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
      MICON           =   "frmpostingjfa.frx":268C
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
      TabIndex        =   18
      Top             =   2040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "To Aktiva"
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
      MICON           =   "frmpostingjfa.frx":29A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsubmit 
      Height          =   375
      Left            =   5655
      TabIndex        =   5
      Top             =   3000
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
      MICON           =   "frmpostingjfa.frx":2CC0
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
      Left            =   6600
      TabIndex        =   6
      Top             =   3000
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
      MICON           =   "frmpostingjfa.frx":2FDA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Posting"
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
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Penjualan Fixed Assets"
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
      Width           =   6375
   End
   Begin VB.Label lblkodefa2 
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
      Left            =   3240
      TabIndex        =   11
      Top             =   2040
      Width           =   4215
   End
   Begin VB.Label lblkodefa1 
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
      Left            =   3240
      TabIndex        =   10
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Label lblcom1 
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
      TabIndex        =   9
      Top             =   1200
      Width           =   4815
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      Caption         =   "From Buy Date"
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
      TabIndex        =   8
      Top             =   2550
      Width           =   1455
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      Caption         =   "To Buy Date"
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
      Left            =   3720
      TabIndex        =   7
      Top             =   2550
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   10335
   End
End
Attribute VB_Name = "frmpostingjfa"
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

Dim str1 As String
Private DISPOSAL As Boolean

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdsearch1_Click()
    carisql1 = "select kdcomp, nmcompscr from gl_company"
    namatabel = "Company"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtcom1 = hasil
    caricom1
    hasil = ""
End Sub

Private Sub cmdsearch2_Click()
    setup6 = txtcom1
    carisql1 = "select kdaktiva, nmaktiva from gl_aktiva"
    namatabel = "Posting Penjualan F/A"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch2_GotFocus()
    If hasil = "" Then Exit Sub
    txtkodefa1 = hasil
    txtkodefa1_LostFocus
    hasil = ""
End Sub

Private Sub cmdsearch3_Click()
    setup6 = txtcom1
    carisql1 = "select kdaktiva, nmaktiva from gl_aktiva"
    namatabel = "Posting Penjualan F/A"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch3_GotFocus()
    If hasil = "" Then Exit Sub
    txtkodefa2 = hasil
    txtkodefa2_LostFocus
    hasil = ""
End Sub

Private Sub CmdSubmit_Click()
    On Error GoTo err_handler
    
    OBJ.Open dsn
    SQL = "select * from toogle"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        If RST!comp_id <> GetTheComputerName Then
            MsgBox "Access denied" & vbCrLf & _
            "Computer name : " & RST!comp_id & " Username : " & UserOnline & vbCrLf & _
            "Task : " & RST!task, vbExclamation, "Error"
            OBJ.Close
            Unload Me
            Exit Sub
        End If
        
        RST.MoveNext
    Loop
    OBJ.Close
    
    If txtcom1 = "" Then Exit Sub
    
    If txtkodefa1 = "" Then txtkodefa1 = "0"
    If txtkodefa2 = "" Then txtkodefa2 = "z"
    
    If date2 < date1 Then
        MsgBox "To Buy Date Can Not Smaller Then From Buy Date.", vbExclamation, "Warning"
        date2.SetFocus
        Exit Sub
    End If
    
    If txtkodefa2 < txtkodefa1 Then
        MsgBox "To Aktiva Can Not Smaller Then From Aktiva.", vbExclamation, "Warning"
        txtkodefa2 = ""
        txtkodefa2.SetFocus
        Exit Sub
    End If
    
    
        
    OBJ.Open dsn
    SQL = "select * from gl_aktiva where kdcomp = '" & txtcom1 & "' and kdaktiva >= '" & txtkodefa1 & "' and kdaktiva <= '" & txtkodefa2 & "' and tgljual >= '" & tanggal1 & "' and tgljual <= '" & tanggal2 & "' and flag = 'P'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "There Is No Data To Posting", vbInformation, "Information"
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "insert into toogle"
    SQL = SQL + "(comp_id"
    SQL = SQL + ",task)"

    SQL = SQL + "VALUES"
    SQL = SQL + "('" & GetTheComputerName & "'"
    SQL = SQL + ",'Posting Penjualan Fixed Assets')"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "Select * From gl_transaksi Where kdtrx = 'JJ' and notrx >= '" & txtkodefa1 & "' and notrx <= '" & txtkodefa2 & "' and  LEFT(desctrx,8)='Disposal'"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        DISPOSAL = True
    Else
        DISPOSAL = False
    End If
    OBJ.Close
    
    OBJ.Open dsn
    SQL = "select * from gl_aktiva where kdcomp = '" & txtcom1 & "' and kdaktiva >= '" & txtkodefa1 & "' and kdaktiva <= '" & txtkodefa2 & "' and tgljual >= '" & tanggal1 & "' and tgljual <= '" & tanggal2 & "' and flag = 'P' and curr1 <> ''"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        date3 = RST!tgljual
        str1 = RST!kdaktiva
        
        OBJ1.Open dsn
        SQL1 = "insert into gl_transaksi "
        SQL1 = SQL1 + "(kdcomp, "
        SQL1 = SQL1 + "tgltrx, "
        SQL1 = SQL1 + "kdtrx, "
        SQL1 = SQL1 + "notrx, "
        SQL1 = SQL1 + "kurs, "
        SQL1 = SQL1 + "noactrx, "
        SQL1 = SQL1 + "desctrx, "
        SQL1 = SQL1 + "dbkrtrx, "
        SQL1 = SQL1 + "amounttrx, "
        SQL1 = SQL1 + "nilaitrx, "
        SQL1 = SQL1 + "currtrx, "
        SQL1 = SQL1 + "flag, "
        SQL1 = SQL1 + "flagprint, "
        SQL1 = SQL1 + "flagadjustment, "
        SQL1 = SQL1 + "cekbg, "
        SQL1 = SQL1 + "identry, "
        SQL1 = SQL1 + "idupdate, "
        SQL1 = SQL1 + "dateentry, "
        SQL1 = SQL1 + "dateupdate, "
        SQL1 = SQL1 + "reconsil, "
        SQL1 = SQL1 + "lineitem)"
            
        SQL1 = SQL1 + " values"
        SQL1 = SQL1 + "('" & txtcom1 & "',"
        SQL1 = SQL1 + "convert(datetime,'" & tanggal3 & "'),"
        SQL1 = SQL1 + "'JJ',"
        SQL1 = SQL1 + "'" & RST!kdaktiva & "',"
        SQL1 = SQL1 + "convert(money,'" & RST!kurs1 & "'),"
        SQL1 = SQL1 + "'" & RST!ac_aktiva & "',"
        If DISPOSAL = True Then
            SQL1 = SQL1 + "'Disposal Aktiva " & RST!kdaktiva & "',"
        Else
            SQL1 = SQL1 + "'Penjualan Aktiva " & RST!kdaktiva & "',"
        End If
        SQL1 = SQL1 + "'K',"
        SQL1 = SQL1 + "convert(money,'" & RST!nilaibeli & "'),"
        SQL1 = SQL1 + "convert(money,'" & RST!hargabeli & "'),"
        SQL1 = SQL1 + "'" & RST!curr1 & "',"
        SQL1 = SQL1 + "'B',"
        SQL1 = SQL1 + "'J',"
        SQL1 = SQL1 + "'0',"
        SQL1 = SQL1 + "' ',"
        SQL1 = SQL1 + "'" & kuser & "',"
        SQL1 = SQL1 + "' ',"
        SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
        SQL1 = SQL1 + "convert(datetime,' '),"
        SQL1 = SQL1 + "' ',"
        SQL1 = SQL1 + "convert(numeric,'2'))"
        Set RST1 = OBJ1.Execute(SQL1)
        
        SQL1 = "insert into gl_transaksi "
        SQL1 = SQL1 + "(kdcomp, "
        SQL1 = SQL1 + "tgltrx, "
        SQL1 = SQL1 + "kdtrx, "
        SQL1 = SQL1 + "notrx, "
        SQL1 = SQL1 + "kurs, "
        SQL1 = SQL1 + "noactrx, "
        SQL1 = SQL1 + "desctrx, "
        SQL1 = SQL1 + "dbkrtrx, "
        SQL1 = SQL1 + "amounttrx, "
        SQL1 = SQL1 + "nilaitrx, "
        SQL1 = SQL1 + "currtrx, "
        SQL1 = SQL1 + "flag, "
        SQL1 = SQL1 + "flagprint, "
        SQL1 = SQL1 + "flagadjustment, "
        SQL1 = SQL1 + "cekbg, "
        SQL1 = SQL1 + "identry, "
        SQL1 = SQL1 + "idupdate, "
        SQL1 = SQL1 + "dateentry, "
        SQL1 = SQL1 + "dateupdate, "
        SQL1 = SQL1 + "reconsil, "
        SQL1 = SQL1 + "lineitem)"
            
        SQL1 = SQL1 + " values"
        SQL1 = SQL1 + "('" & txtcom1 & "',"
        SQL1 = SQL1 + "convert(datetime,'" & tanggal3 & "'),"
        SQL1 = SQL1 + "'JJ',"
        SQL1 = SQL1 + "'" & RST!kdaktiva & "',"
        SQL1 = SQL1 + "convert(money,'" & RST!kurs1 & "'),"
        SQL1 = SQL1 + "'" & RST!ac_lawan & "',"
        If DISPOSAL = True Then
            SQL1 = SQL1 + "'Disposal Aktiva " & RST!kdaktiva & "',"
        Else
            SQL1 = SQL1 + "'Penjualan Aktiva " & RST!kdaktiva & "',"
        End If
        SQL1 = SQL1 + "'D',"
        SQL1 = SQL1 + "convert(money,'" & RST!nilaibeli & "'),"
        SQL1 = SQL1 + "convert(money,'" & RST!hargabeli & "'),"
        SQL1 = SQL1 + "'" & RST!curr1 & "',"
        SQL1 = SQL1 + "'B',"
        SQL1 = SQL1 + "'J',"
        SQL1 = SQL1 + "'0',"
        SQL1 = SQL1 + "' ',"
        SQL1 = SQL1 + "'" & kuser & "',"
        SQL1 = SQL1 + "' ',"
        SQL1 = SQL1 + "convert(datetime,'" & tanggalsekarang & "'),"
        SQL1 = SQL1 + "convert(datetime,' '),"
        SQL1 = SQL1 + "' ',"
        SQL1 = SQL1 + "convert(numeric,'1'))"
        Set RST1 = OBJ1.Execute(SQL1)
        
        SQL1 = "update gl_transaksi set flag = 'J' where kdcomp = '" & txtcom1 & "' and (kdtrx = 'JS' or kdtrx = 'JB') and notrx >= '" & txtkodefa1 & "' and notrx <= '" & txtkodefa2 & "' and flag = 'B' and tgltrx > '" & tanggal3 & "'"
        Set RST1 = OBJ1.Execute(SQL1)
        
        SQL1 = "update gl_aktiva set flag = 'J' where kdcomp = '" & txtcom1 & "' and kdaktiva = '" & str1 & "' and flag = 'P'"
        Set RST1 = OBJ1.Execute(SQL1)
        OBJ1.Close
        
        RST.MoveNext
    Loop
    OBJ.Close
    
    If txtkodefa1 = "0" Then txtkodefa1 = ""
    If txtkodefa2 = "z" Then txtkodefa2 = ""
    
    OBJ.Open dsn
    SQL = "delete from toogle where comp_id = '" & GetTheComputerName & "' and task = 'Posting Penjualan Fixed Assets'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
    
    MsgBox "Posting Complete", vbInformation, "Information"
    
    cmdclose_Click
    Exit Sub
err_handler:
    OBJ.Open dsn
    SQL = "delete from toogle where comp_id = '" & GetTheComputerName & "' and task = 'Posting Penjualan Fixed Assets'"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
End Sub

Function tanggalsekarang()
    tanggalsekarang = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
End Function

Function tanggal1()
    tanggal1 = Month(date1) & "/" & Day(date1) & "/" & Year(date1)
End Function

Function tanggal2()
    tanggal2 = Month(date2) & "/" & Day(date2) & "/" & Year(date2)
End Function

Function tanggal3()
    tanggal3 = Month(date3) & "/" & Day(date3) & "/" & Year(date3)
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    date1 = Date
    date2 = Date
    
    
End Sub

Private Sub txtcom1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtkodefa1.SetFocus
End Sub

Private Sub txtcom1_LostFocus()
    caricom1
End Sub

Private Sub caricom1()
    If txtcom1 = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from gl_company where kdcomp = '" & txtcom1 & "'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Company " & txtcom1 & " Not Found.", vbExclamation, "Warning"
        txtcom1 = ""
        lblcom1 = ""
        txtcom1.SetFocus
    Else
        lblcom1 = RST!nmcompscr
        date1 = RST!tglawal
        date2 = RST!tglakhir
        txtkodefa1.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub txtkodefa1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtkodefa2.SetFocus
End Sub

Private Sub txtkodefa1_LostFocus()
    carifa1
End Sub

Private Sub carifa1()
    If txtkodefa1 = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from gl_aktiva where kdcomp = '" & txtcom1 & "' and kdaktiva = '" & txtkodefa1 & "' and flag = 'P' and curr1 <> ''"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Aktiva " & txtkodefa1 & " Not Found.", vbExclamation, "Warning"
        txtkodefa1 = ""
        lblkodefa1 = ""
        txtkodefa1.SetFocus
    Else
        lblkodefa1 = RST!nmaktiva
        txtkodefa2.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub txtkodefa2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then date1.SetFocus
End Sub

Private Sub txtkodefa2_LostFocus()
    carifa2
End Sub

Private Sub carifa2()
    If txtkodefa2 = "" Then Exit Sub
    OBJ.Open dsn
    SQL = "select * from gl_aktiva where kdcomp = '" & txtcom1 & "' and kdaktiva = '" & txtkodefa2 & "' and flag = 'P' and curr1 <> ''"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Aktiva " & txtkodefa2 & " Not Found.", vbExclamation, "Warning"
        txtkodefa2 = ""
        lblkodefa2 = ""
        txtkodefa2.SetFocus
    Else
        lblkodefa2 = RST!nmaktiva
        date1.SetFocus
    End If
    OBJ.Close
End Sub
