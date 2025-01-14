VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Begin VB.Form frmpermintaanclose 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Close Permintaan Barang"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6975
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List3 
      Height          =   1815
      ItemData        =   "frmpermintaanclose.frx":0000
      Left            =   3960
      List            =   "frmpermintaanclose.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   2880
      Width           =   2895
   End
   Begin VB.ListBox List2 
      Height          =   1815
      ItemData        =   "frmpermintaanclose.frx":0004
      Left            =   3960
      List            =   "frmpermintaanclose.frx":0006
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   2895
   End
   Begin Chameleon.chameleonButton cmdclose 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   873
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
      MICON           =   "frmpermintaanclose.frx":0008
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Chameleon.chameleonButton cmdpost 
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      ToolTipText     =   "Submit"
      Top             =   1320
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "4"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
         Weight          =   700
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
      MICON           =   "frmpermintaanclose.frx":0322
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox List1 
      Height          =   4110
      ItemData        =   "frmpermintaanclose.frx":063C
      Left            =   120
      List            =   "frmpermintaanclose.frx":063E
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   600
      Width           =   2895
   End
   Begin Chameleon.chameleonButton cmdpost1 
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      ToolTipText     =   "Submit"
      Top             =   3480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "4"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
         Weight          =   700
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
      MICON           =   "frmpermintaanclose.frx":0640
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   3
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cancel Permintaan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Close Permintaan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Permintaan Barang"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "frmpermintaanclose"
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

Dim i, j As Integer

Private Sub cmdclose_Click()
    Unload Me
End Sub

Private Sub cmdpost_Click()
    If List1.ListCount = 0 Then Exit Sub
    j = 0
    For i = 1 To List1.ListCount
        If List1.Selected(i - 1) = True Then j = j + 1
    Next i
    
    If j = 0 Then
        MsgBox "To close Permintaan Barang " + Chr(13) + "user must select/check at least one Permintaan Barang.", vbExclamation, "Information"
        Exit Sub
    End If
    
    For i = 1 To List1.ListCount
        If List1.Selected(i - 1) = True Then
            OBJ.Open dsn
            SQL = "update am_perminhdr set flag = '1' where nobkt = '" & List1.List(i - 1) & "'"
            Set RST = OBJ.Execute(SQL)
            
            SQL = "update am_perminin set status = '1' Where nobkt = '" & List1.List(i - 1) & "'"
            Set RST = OBJ.Execute(SQL)
            
            SQL = "update am_perminapp set status = '1' where nobkt = '" & List1.List(i - 1) & "'"
            Set RST = OBJ.Execute(SQL)
            OBJ.Close
            
            List2.AddItem List1.List(i - 1)
        End If
    Next i
    
    List1.Clear
    OBJ.Open dsn
    SQL = "SELECT nobkt FROM AM_perminhdr WHERE flag = '0' order by nobkt asc"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        List1.AddItem RST!nobkt
        
        RST.MoveNext
    Loop
    OBJ.Close
    
    MsgBox "Closing Complete.", vbInformation, "Information"
    
End Sub

Private Sub cmdpost1_Click()
    If List1.ListCount = 0 Then Exit Sub
    j = 0
    For i = 1 To List1.ListCount
        If List1.Selected(i - 1) = True Then j = j + 1
    Next i
    
    If j = 0 Then
        MsgBox "To cancel Permintaan Barang " + Chr(13) + "user must select/check at least one Permintaan Barang.", vbExclamation, "Information"
        Exit Sub
    End If
    
    For i = 1 To List1.ListCount
        If List1.Selected(i - 1) = True Then
            OBJ.Open dsn
            SQL = "update am_perminhdr set flag = '2' where nobkt = '" & List1.List(i - 1) & "'"
            Set RST = OBJ.Execute(SQL)
            
            SQL = "update am_perminapp set status = '2' where nobkt = '" & List1.List(i - 1) & "'"
            Set RST = OBJ.Execute(SQL)
            OBJ.Close
            
            List3.AddItem List1.List(i - 1)
        End If
    Next i
    
    List1.Clear
    OBJ.Open dsn
    SQL = "SELECT nobkt FROM AM_perminhdr WHERE flag = '0' order by nobkt asc"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        List1.AddItem RST!nobkt
        
        RST.MoveNext
    Loop
    OBJ.Close
    
    MsgBox "Canceling Complete.", vbInformation, "Information"
End Sub

Private Sub Form_Load()
    OBJ.Open dsn
    SQL = "SELECT nobkt FROM am_perminhdr WHERE flag = '0' order by nobkt asc"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        List1.AddItem RST!nobkt
    
        RST.MoveNext
    Loop
    
    SQL = "SELECT nobkt FROM am_perminhdr WHERE flag = '1' order by nobkt asc"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        List2.AddItem RST!nobkt
    
        RST.MoveNext
    Loop
    
    SQL = "SELECT nobkt FROM am_perminhdr WHERE flag = '2' order by nobkt asc"
    Set RST = OBJ.Execute(SQL)
    Do While Not RST.EOF
        List3.AddItem RST!nobkt
    
        RST.MoveNext
    Loop
    OBJ.Close
End Sub
