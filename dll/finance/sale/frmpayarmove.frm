VERSION 5.00
Begin VB.Form frmpayarmove 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Move No Bukti from faktur to faktur"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6720
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox txtnofak2 
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
      Left            =   4920
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtnofak1 
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
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtnobkt 
      Appearance      =   0  'Flat
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
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Width           =   6735
   End
   Begin VB.Label Label3 
      Caption         =   "Move to No. Faktur"
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
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "No. Faktur"
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
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "No. Bukti"
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
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmpayarmove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Dim OBJ2 As New ADODB.Connection
Dim RST2 As New ADODB.Recordset
Dim SQL2 As String

Private Sub txtnofak1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        OBJ.Open dsn
        SQL = "Select * From am_cashlin Where noapply = '" & txtnofak1 & "'"
        Set RST = OBJ.Execute(SQL)
        If Not RST.EOF Then txtnobkt = RST!nobkt
        OBJ.Close
    End If
End Sub

'Update am_aropnfil set noapply = '' Where noapply = '' and nobkt = '' and transtype = 'PM'
'Update am_cashlin set NoApply = 'P1604940' Where NoBkt = 'P34499' and NoApply = 'P1603584'
