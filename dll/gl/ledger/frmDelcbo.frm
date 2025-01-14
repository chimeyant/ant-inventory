VERSION 5.00
Begin VB.Form frmDelcbo 
   Caption         =   "HAPUS PAYMENT"
   ClientHeight    =   4425
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   ControlBox      =   0   'False
   DrawMode        =   15  'Merge Pen Not
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   4560
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmddel 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   15
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3360
      Width           =   2895
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   11
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox txtdesc 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2880
      Width           =   2895
   End
   Begin VB.TextBox txtkpd 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2400
      Width           =   2895
   End
   Begin VB.TextBox txtgiro 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1920
      Width           =   2895
   End
   Begin VB.TextBox txtnovou 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox txtnotrans 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox txtnopayment 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enter No. Payment Here"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label6 
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Desc"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Di Bayar Kepada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Giro/Check"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "No. Voucher"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "No. Transaksi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "frmDelcbo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private Sub cmdclose_Click()
Unload Me
End Sub

Private Sub cmddel_Click()
On Error GoTo Err_handler
    If MsgBox("Data yang sudah dihapus tidak bisa dikembalikan." + Chr(13) + _
        "Anda yakin ingin menghapus data ini ?", vbYesNo + vbQuestion, "Konfirmasi") = vbYes Then
                
        OBJ.Open dsn
            SQL = "Delete From no_bank_payment Where notrx='" & txtnotrans & "' and no_payment = '" & txtnopayment & "'"
            OBJ.Execute SQL
            
            SQL = "Delete From gl_transaksi Where notrx='" & txtnotrans & "'"
            OBJ.Execute SQL
        OBJ.Close
        MsgBox "Data berhasil dihapus.", vbInformation, AppName
        Call Clear
    End If
    Exit Sub
Err_handler:
    If OBJ.State = 1 Then OBJ.Close
    MsgBox Err.Description, vbCritical, "ERROR"
End Sub

Private Sub Form_Load()
    Call Clear
End Sub

Private Sub txtnopayment_KeyPress(KeyAscii As Integer)
On Error GoTo Err_handler
If KeyAscii = 13 Then
    OBJ.Open dsn
    SQL = "Select a.notrx,a.no_voucher,a.kpd,a.tgljt,b.kdtrx,b.desctrx,b.amounttrx,b.cekbg From no_bank_payment a "
    SQL = SQL + "inner join gl_transaksi b on a.notrx=b.notrx "
    SQL = SQL + "Where a.no_payment='" & txtnopayment & "'"
    Set RST = OBJ.Execute(SQL)
    
    If RST.EOF Then
        OBJ.Close
        Call Clear
        MsgBox "Data tidak ditemukan.", vbCritical, AppName
        Exit Sub
    End If
    txtnotrans = RST!notrx
    txtnovou = RST!no_voucher
    txtgiro = RST!cekbg
    txtkpd = RST!kpd
    txtdesc = Mid(RST!desctrx, 6, 60)
    txtamount = Format(RST!amounttrx, "###,###,##0.00")
    
    OBJ.Close
End If
Exit Sub
Err_handler:
    If OBJ.State = 1 Then OBJ.Close
    MsgBox Err.Description, vbCritical, "ERROR"
End Sub


Private Sub Clear()
    txtnopayment = ""
    txtnotrans = ""
    txtnovou = ""
    txtkpd = ""
    txtgiro = ""
    txtdesc = ""
    txtamount = "0.00"
End Sub
