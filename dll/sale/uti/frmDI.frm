VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "chameleon.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmDI 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proses Invoice"
   ClientHeight    =   1005
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4035
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.ProgressBar Pg 
      Height          =   285
      Left            =   90
      TabIndex        =   2
      Top             =   120
      Width           =   3825
      _Version        =   851970
      _ExtentX        =   6747
      _ExtentY        =   503
      _StockProps     =   93
      TextAlignment   =   2
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   375
      Left            =   3060
      TabIndex        =   0
      Top             =   525
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
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmDI.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdhapusL 
      Height          =   375
      Left            =   2190
      TabIndex        =   1
      Top             =   525
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Proses"
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
      MICON           =   "frmDI.frx":081A
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
Attribute VB_Name = "frmDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdhapusL_Click()
    If nmuser <> "Creator" Then
        MsgBox "Akses ditolak", vbCritical, AppName
        Unload Me
    End If
    If MsgBox("Invoice akan segera diproses" + vbCrLf + "Klik tombol OK untuk melanjutkan..", vbInformation + vbOKCancel, AppName) = vbCancel Then Exit Sub
    
'    Exit Sub
    Pg.Min = 0
    Pg.Max = 100
    Pg.Value = 0
    cmdhapusL.Enabled = False
    Me.MousePointer = vbHourglass
    
    OBJ.Open dsn
    'Hapus SJ
    SQL = "Delete from am_sjapp Where nosj like 'LL%'"
    Set RST = OBJ.Execute(SQL)
        Pg.Value = Pg.Value + 5
        Pg.text = "5 % Complete"
    
    SQL = "Delete from am_sjhdr Where nosj like 'LL%'"
    Set RST = OBJ.Execute(SQL)
        Pg.Value = Pg.Value + 5
        Pg.text = "10 % Complete"
    
    SQL = "Delete from am_sjlin Where nosj like 'LL%'"
    Set RST = OBJ.Execute(SQL)
        Pg.Value = Pg.Value + 5
        Pg.text = "15 % Complete"
    
    SQL = "Delete from am_sjdrop Where nosj like 'LL%'"
    Set RST = OBJ.Execute(SQL)
        Pg.Value = Pg.Value + 5
        Pg.text = "20 % Complete"
    
    SQL = "Delete from am_sjdesc Where nosj like 'LL%'"
    Set RST = OBJ.Execute(SQL)
        Pg.Value = Pg.Value + 5
        Pg.text = "25 % Complete"
    
    'Hapus SO
    SQL = "Delete From am_sjlist Where noso like 'L%'"
    Set RST = OBJ.Execute(SQL)
        Pg.Value = Pg.Value + 5
        Pg.text = "30 % Complete"
    
    SQL = "Delete From am_soapp Where noso like 'L%'"
    Set RST = OBJ.Execute(SQL)
        Pg.Value = Pg.Value + 5
        Pg.text = "35 % Complete"
    
    SQL = "Delete From am_sohdr Where noso like 'L%'"
    Set RST = OBJ.Execute(SQL)
        Pg.Value = Pg.Value + 5
        Pg.text = "40 % Complete"
    
    SQL = "Delete From am_solin Where noso like 'L%'"
    Set RST = OBJ.Execute(SQL)
        Pg.Value = Pg.Value + 5
        Pg.text = "45 % Complete"
    
    'Hapus Inv
    SQL = "Delete From am_invhdr Where nobkt like 'L%'"
    Set RST = OBJ.Execute(SQL)
        Pg.Value = Pg.Value + 5
        Pg.text = "50 % Complete"
    
    SQL = "Delete From am_invlin Where nobkt like 'L%'"
    Set RST = OBJ.Execute(SQL)
        Pg.Value = Pg.Value + 5
        Pg.text = "55 % Complete"
    
    SQL = "Delete From am_invdelete Where nobkt like 'L%'"
    Set RST = OBJ.Execute(SQL)
        Pg.Value = Pg.Value + 5
        Pg.text = "60 % Complete"
        
    SQL = "Delete From am_invdesc Where noinv like 'L%'"
    Set RST = OBJ.Execute(SQL)
        Pg.Value = Pg.Value + 5
        Pg.text = "65 % Complete"
        
    'Hapus Faktur
    SQL = "Delete from am_aropnfil Where noapply like 'L%'"
    Set RST = OBJ.Execute(SQL)
        Pg.Value = Pg.Value + 5
        Pg.text = "70 % Complete"
    
    SQL = "Delete from am_aropnfil1 Where noapply like 'L%'"
    Set RST = OBJ.Execute(SQL)
        Pg.Value = Pg.Value + 5
        Pg.text = "75 % Complete"
        
    SQL = "Delete From am_cashhdr Where noapply like 'L%'"
    Set RST = OBJ.Execute(SQL)
        Pg.Value = Pg.Value + 5
        Pg.text = "80 % Complete"
        
    SQL = "Delete From am_cashlin Where noapply like 'L%'"
    Set RST = OBJ.Execute(SQL)
        Pg.Value = Pg.Value + 5
        Pg.text = "85 % Complete"
        
    SQL = "Delete From am_cashsub Where nobkt like 'L%'"
    Set RST = OBJ.Execute(SQL)
        Pg.Value = Pg.Value + 5
        Pg.text = "90 % Complete"
        
    SQL = "Delete From am_aropninv Where faktur like 'L%'"
    Set RST = OBJ.Execute(SQL)
        Pg.Value = Pg.Value + 5
        Pg.text = "95 % Complete"
        
    SQL = "Delete From am_returjual Where noapply like 'L%'"
    Set RST = OBJ.Execute(SQL)
        Pg.Value = Pg.Value + 5
        Pg.text = "100 % Complete"
        
    SQL = "Delete From am_temp_pembayaranbyjenis Where noapply like 'L%'"
    Set RST = OBJ.Execute(SQL)
        Pg.Value = Pg.Value + 5
        Pg.text = "100 % Complete"
        
    
    SQL = "delete a from  am_beliapp a inner join am_voucherhdr b on a.Ref1 = b.novoucher Where b.ispajak = '0'"
    Set RST = OBJ.Execute(SQL)
        Pg.Value = Pg.Value + 5
        Pg.text = "100 % Complete"
    
    SQL = "delete a From am_apopnfil a inner join am_voucherhdr b on a.NoApply = b.novoucher Where b.ispajak = '0'"
    Set RST = OBJ.Execute(SQL)
        Pg.Value = Pg.Value + 5
        Pg.text = "100 % Complete"
        
    
    SQL = "DELETE a From am_apcashhdr a inner join no_bank_payment b on a.NoBkt = b.no_payment "
    SQL = SQL + " inner join am_voucherhdr c on b.no_voucher = c.novoucher Where c.ispajak = '0'"
    Set RST = OBJ.Execute(SQL)
        Pg.Value = Pg.Value + 5
        Pg.text = "100 % Complete"
        
    SQL = "Delete a From am_apcashsub a inner join no_bank_payment b on a.nobukti = b.no_payment "
    SQL = SQL + " inner join am_voucherhdr c on b.no_voucher = c.novoucher Where c.ispajak = '0'"
    Set RST = OBJ.Execute(SQL)
        Pg.Value = Pg.Value + 5
        Pg.text = "100 % Complete"
    
    SQL = "Delete a From am_voucherin a inner join am_voucherhdr b on a.novoucher = b.novoucher where b.ispajak = '0' "
    Set RST = OBJ.Execute(SQL)
        Pg.Value = Pg.Value + 5
        Pg.text = "100 % Complete"
        
    SQL = "Delete  From am_voucherhdr Where ispajak = '0' "
    Set RST = OBJ.Execute(SQL)
        Pg.Value = Pg.Value + 5
        Pg.text = "100 % Complete"

    OBJ.Close
    MsgBox "Data  berhasil diproses.", vbInformation, AppName
    Pg.Value = 0
    cmdhapusL.Enabled = True
    Me.MousePointer = vbDefault
    
'#SJ
'Select * From am_sjapp Where nosj like 'LL%'   lost
'Select * From am_sjapp Where nosj like 'PP%'
'Select * From am_sjhdr Where nosj like 'LL%'   lost
'Select * From am_sjlin Where nosj like 'LL%'   lost
'Select * From am_sjdrop Where nosj like 'LL%'  lost
'Select * From am_sjdesc Where nosj like 'LL%'  lost

'#SO
'Select * From am_sjlist Where noso like 'L%'    lost
'Select * From am_soapp Where noso like 'L%'    lost
'Select * From am_sohdr Where noso like 'L%'    lost
'Select * From am_solin Where noso like 'L%'    lost

'#INV
'Select * From am_invhdr Where NoBkt like 'L%'  lost
'Select * From am_invlin Where NoBkt like 'L%'  lost (nobkt = 012345)
'Select * From am_invdelete Where NoBkt like 'L%'   lost
'Select * From am_invdesc Where noinv like 'L%'     lost

'#FAKTUR
'Select * From am_aropnfil Where NoApply like 'L%'  lost
'Select * From am_cashhdr Where NoApply like 'L%'   lost
'Select * From am_cashlin Where NoApply like 'L%'   lost
'Select * From am_cashsub Where NoBkt like 'L%'     lost
'Select * From am_aropnfil1 Where NoApply like 'L%'
'Select * From am_aropninv Where faktur like 'L%'
'Select * From am_returjual Where noapply like 'L%'     lost
'Select * From am_temp_pembayaranbyjenis Where noapply like 'L%'
'==================================================================
'#PEMBELIAN
'Select a.* From am_beliapp a inner join am_voucherhdr b on a.Ref1 = b.novoucher Where b.ispajak = '0'
'Select a.* From am_apopnfil a inner join am_voucherhdr b on a.NoApply = b.novoucher Where b.ispajak = '0'

'Select * From am_apcashhdr a inner join no_bank_payment b on a.NoBkt = b.no_payment
'inner join am_voucherhdr c on b.no_voucher = c.novoucher Where c.ispajak = '0'

'Select * From am_apcashsub a inner join no_bank_payment b on a.nobukti = b.no_payment
'inner join am_voucherhdr c on b.no_voucher = c.novoucher Where c.ispajak = '0'

'#VOUCHER HDR-IN
'Select a.*,b.* From am_voucherhdr a inner join am_voucherin b on a.novoucher =b.novoucher Where a.ispajak = '0'
'#VOUCHERIN
'Select a.* From am_voucherin a inner join am_voucherhdr b on a.novoucher = b.novoucher where b.ispajak = '0'
'#VOUCHERHDR
'Select * From am_voucherhdr Where ispajak = '0'
End Sub

Private Sub Form_Load()
    'If nmuser <> "Creator" Then
    '    MsgBox "Akses ditolak", vbCritical, AppName
    '    Unload Me
    'End If
End Sub
