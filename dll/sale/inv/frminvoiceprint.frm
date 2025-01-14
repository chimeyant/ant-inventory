VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D20B3A1A-4F4C-11D9-9ED5-00112F04C2B8}#20.0#0"; "Terbilang.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "CODEJO~2.OCX"
Begin VB.Form frminvoiceprint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Faktur Penjualan"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frminvoiceprint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.CheckBox cb 
      Height          =   255
      Left            =   1320
      TabIndex        =   11
      Top             =   1800
      Width           =   1215
      _Version        =   851970
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "> 6 Ror"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin TDBNumber6Ctl.TDBNumber txtvalue 
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   450
      Calculator      =   "frminvoiceprint.frx":2372
      Caption         =   "frminvoiceprint.frx":2392
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frminvoiceprint.frx":23FE
      Keys            =   "frminvoiceprint.frx":241C
      Spin            =   "frminvoiceprint.frx":2466
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   "."
      DisplayFormat   =   "###,###,###,##0;(###,###,###,##0);0"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "###,###,###,##0;(###,###,###,##0)"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   999999999999999
      MinValue        =   -999999999999999
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   1
      Separator       =   ","
      ShowContextMenu =   0
      ValueVT         =   -65531
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin Sukedi.Terbilang Terbilang1 
      Left            =   480
      Top             =   2760
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   285
      Left            =   1320
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
      Format          =   143392771
      CurrentDate     =   38773
   End
   Begin VB.TextBox txtnodo2 
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
      MaxLength       =   15
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox txtnodo1 
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
      MaxLength       =   15
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin Crystal.CrystalReport crystal 
      Left            =   0
      Top             =   2760
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
   Begin Chameleon.chameleonButton cmdsearch 
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "Dr No Bukti"
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
      MICON           =   "frminvoiceprint.frx":248E
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
      Left            =   2280
      TabIndex        =   5
      Top             =   2160
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
      MICON           =   "frminvoiceprint.frx":27A8
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
      Left            =   1320
      TabIndex        =   4
      Top             =   2160
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
      MICON           =   "frminvoiceprint.frx":2AC2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdsearch1 
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      BTYPE           =   9
      TX              =   "s/d No Bukti"
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
      MICON           =   "frminvoiceprint.frx":2DDC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker date2 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   1440
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
      Format          =   143392771
      CurrentDate     =   38773
   End
   Begin VB.Label Label3 
      Caption         =   "s/d Tanggal"
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
      TabIndex        =   9
      Top             =   1470
      Width           =   1215
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
      TabIndex        =   8
      Top             =   1110
      Width           =   1215
   End
End
Attribute VB_Name = "frminvoiceprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim RST1 As New ADODB.Recordset
Dim SQL1 As String

Dim OBJ2 As New ADODB.Connection
Dim RST2 As New ADODB.Recordset
Dim SQL2 As String

Dim SP As New ADODB.Command
Dim vsp(4) As Variant

Dim str1 As String

Private Sub cariinv1()
    If txtnodo1 = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from am_invhdr where nobkt = '" & txtnodo1 & "' and type = 'I'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Invoice " & txtnodo1 & " Not Found.", vbExclamation, "Warning"
        txtnodo1 = ""
        txtnodo1.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub cariinv2()
    If txtnodo2 = "" Then Exit Sub
    
    OBJ.Open dsn
    SQL = "select * from am_invhdr where nobkt = '" & txtnodo2 & "' and type = 'I'"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        MsgBox "Invoice " & txtnodo2 & " Not Found.", vbExclamation, "Warning"
        txtnodo2 = ""
        txtnodo2.SetFocus
    End If
    OBJ.Close
End Sub

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdsearch_Click()
    If v_fastsearching = True Then
        If v_fstgl1 > v_fstgl2 Then
            MsgBox "Invalid date range, search abort.", vbExclamation, "Error"
            Exit Sub
        End If
        carisql1 = "select nobkt, convert(char(11),tglbkt )'tglbkt', type from am_invhdr where type = 'I' and tglbkt >= '" & batas1 & "' and tglbkt <= '" & batas2 & "'"
    Else
        carisql1 = "select nobkt, convert(char(11),tglbkt )'tglbkt', type from am_invhdr where type = 'I'"
    End If
    namatabel = "Penjualan"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch_GotFocus()
    If hasil = "" Then Exit Sub
    txtnodo1 = hasil
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdsearch1_Click()
    If v_fastsearching = True Then
        If v_fstgl1 > v_fstgl2 Then
            MsgBox "Invalid date range, search abort.", vbExclamation, "Error"
            Exit Sub
        End If
        carisql1 = "select nobkt, convert(char(11),tglbkt )'tglbkt', type from am_invhdr where type = 'I' and tglbkt >= '" & batas1 & "' and tglbkt <= '" & batas2 & "'"
    Else
        carisql1 = "select nobkt, convert(char(11),tglbkt )'tglbkt', type from am_invhdr where type = 'I'"
    End If
    namatabel = "Penjualan"
    
    frmsearch.Show vbModal
End Sub

Private Sub cmdsearch1_GotFocus()
    If hasil = "" Then Exit Sub
    txtnodo2 = hasil
    hasil = ""
    hasil1 = ""
    hasil2 = ""
End Sub

Private Sub cmdview_Click()
    If txtnodo1 = "" Then Exit Sub
    If txtnodo2 = "" Then Exit Sub
    If txtnodo1 > txtnodo2 Then
        MsgBox "Error on No bukti.", vbInformation, "information"
        
        Exit Sub
    End If
    
    If date1 > date2 Then
        MsgBox "Error on Date.", vbInformation, "information"
        Exit Sub
    End If
    
    OBJ.Open dsn
    SQL = "select nobkt from am_terbilang"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        OBJ.Close
        OBJ2.Open dsn
        SQL2 = "Delete from am_terbilang"
        Set RST2 = OBJ2.Execute(SQL2)
        OBJ2.Close
        MsgBox "Preview aborted, please try again.", vbExclamation, "Warning"
        Exit Sub
    End If
    OBJ.Close
    
    SP.ActiveConnection = dsn
    SP.CommandType = adCmdStoredProc
    SP.CommandText = "am_invoicesay"
    vsp(0) = txtnodo1
    vsp(1) = txtnodo2
    vsp(2) = Format(date1, "yyyyMMdd")
    vsp(3) = Format(date2, "yyyyMMdd")
    vsp(4) = str1
    SP.Execute , vsp
    Set SP = Nothing
    
    OBJ.Open dsn
    SQL = "select nobkt,nilai,namakurs from am_terbilang"
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        If MsgBox("Translate value aborted, continue with blank converted word.", vbQuestion + vbYesNo, "Question") = vbNo Then
            OBJ.Close
            Exit Sub
        End If
    End If
    
    Do While Not RST.EOF
        SQL1 = "select kodecur from am_invhdr where nobkt = '" & RST!nobkt & "'"
        Set RST1 = OBJ.Execute(SQL1)
        If Not RST1.EOF Then
            OBJ2.Open dsn
            SQL2 = "select base from gl_kurs where kdkurs = '" & RST1!kodecur & "'"
            Set RST2 = OBJ2.Execute(SQL2)
            If RST2!base = "1" Then
                txtvalue.Format = "###,###,###,##0;(###,###,###,##0)"
                txtvalue.DisplayFormat = "###,###,###,##0;(###,###,###,##0);0"
            Else
                txtvalue.Format = "###,###,###,##0.00;(###,###,###,##0.00)"
                txtvalue.DisplayFormat = "###,###,###,##0.00;(###,###,###,##0.00);0.00"
            End If
            OBJ2.Close
        End If
        
        txtvalue = RST!nilai
        Terbilang1.Result1 = txtvalue
        Terbilang1.Currency1 = LCase(RST!namakurs)
        SQL1 = "update am_terbilang set terbilang = '" & Terbilang1.Terbilang & "' where nobkt = '" & RST!nobkt & "'"
        Set RST1 = OBJ.Execute(SQL1)
        
        RST.MoveNext
    Loop
    OBJ.Close
    'MsgBox txtnodo1 & vbCrLf & txtnodo2 & vbCrLf & Format(date1, "yyyyMMdd") & vbCrLf & Format(date2, "yyyyMMdd") & vbCrLf & str1 & vbCrLf & nmuser
    crystal.Reset
    crystal.WindowState = crptMaximized
    crystal.WindowShowCloseBtn = True
    crystal.WindowShowPrintSetupBtn = True
    crystal.WindowShowSearchBtn = True
    crystal.Connect = dsnreport
    crystal.DataFiles(0) = "Proc(am_invoice)"
    If cb.Value = xtpChecked Then
        crystal.ReportFileName = AppPath & "\reports\sale\inv\invoicenx.rpt"
    Else
        crystal.ReportFileName = AppPath & "\reports\sale\inv\invoice.rpt"
    End If
    crystal.ParameterFields(0) = "@noinv1;" & txtnodo1 & ";true"
    crystal.ParameterFields(1) = "@noinv2;" & txtnodo2 & ";true"
    crystal.ParameterFields(2) = "@tanggal1;" & Format(date1, "yyyyMMdd") & ";true"
    crystal.ParameterFields(3) = "@tanggal2;" & Format(date2, "yyyyMMdd") & ";true"
    crystal.ParameterFields(4) = "@type1;" & str1 & ";true"
    crystal.ParameterFields(5) = "@namauser ;" + nmuser + ";true"
    crystal.RetrieveDataFiles
    crystal.Action = 1
    
    OBJ.Open dsn
    SQL = "delete from am_terbilang"
    Set RST = OBJ.Execute(SQL)
    OBJ.Close
End Sub

Function batas1()
    batas1 = Month(v_fstgl1) & "/" & Day(v_fstgl1) & "/" & Year(v_fstgl1)
End Function

Function batas2()
    batas2 = Month(v_fstgl2) & "/" & Day(v_fstgl2) & "/" & Year(v_fstgl2)
End Function

Private Sub Form_Activate()
    'validasi user
    'If kuser <> "q" Then
    '    OBJ.Open dsn
    '    SQL = "select a.* from am_level a left join am_user b on a.kode=b.kodelevel where a.program='175' and b.kodeuser = '1" & kuser & "'"
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
   
    
    str1 = "I"
    
    date1 = Date
    date2 = Date
End Sub

Private Sub txtnodo1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtnodo2.SetFocus
End Sub

Private Sub txtnodo1_LostFocus()
    cariinv1
End Sub

Private Sub txtnodo2_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then date1.SetFocus
End Sub

Private Sub txtnodo2_LostFocus()
    cariinv2
End Sub
