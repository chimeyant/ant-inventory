VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmpayapeditsearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search by ..."
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmpayapeditsearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtsearch 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Top             =   720
      Width           =   3135
   End
   Begin VB.OptionButton Option2 
      Caption         =   "by Cek/Giro"
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
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "by Faktur"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   3375
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Single Click To Choose"
      Top             =   1080
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   5953
      _Version        =   393216
      FixedCols       =   0
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      Caption         =   "Find"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   750
      Width           =   735
   End
End
Attribute VB_Name = "frmpayapeditsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 27 Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    loading_
End Sub

Private Sub grid_Click()
    If grid.TextMatrix(0, 0) = "" Then Exit Sub
    
    If grid.MouseRow > 0 Then
        hasil3 = grid.TextMatrix(grid.Row, 0)
        
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub grid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
        Exit Sub
    End If
    If grid.TextMatrix(0, 0) = "" Then Exit Sub
    If KeyAscii = 13 Then
        hasil3 = grid.TextMatrix(grid.Row, 0)
        Unload Me
    End If
End Sub

Private Sub Option1_Click()
    loading_
End Sub

Private Sub Option2_Click()
    loading_
End Sub

Private Sub txtsearch_Change()
    OBJ.Open dsn
    If Option1.Value = True Then
        SQL = "select nobkt,noapply from AM_apcashlin where kodebayar = 'PM' and noapply like '" + txtsearch + "%'"
    Else
        SQL = "select nobukti,nogiro from AM_apcashsub where nogiro <> '' and nogiro like '" + txtsearch + "%'"
    End If
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        grid.Clear
        grid.Rows = 2
        grid.ColWidth(0) = 1750
        grid.ColWidth(1) = 1750
        OBJ.Close
        Exit Sub
    End If
    Set RST = OBJ.Execute(SQL)
    Set grid.DataSource = RST
    OBJ.Close
    
    grid.ColWidth(0) = 1750
    grid.ColWidth(1) = 1750
End Sub

Private Sub txtsearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
        Exit Sub
    End If
    If KeyAscii = 13 And grid.Rows = 2 Then
        hasil3 = grid.TextMatrix(1, 0)
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub loading_()
    OBJ.Open dsn
    If Option1.Value = True Then
        SQL = "select nobkt,noapply from AM_apcashlin where kodebayar = 'PM'"
    Else
        SQL = "select nobukti,nogiro from AM_apcashsub where nogiro <> ''"
    End If
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        grid.Clear
        grid.Rows = 2
        grid.ColWidth(0) = 1750
        grid.ColWidth(1) = 1750
        
        OBJ.Close
        Exit Sub
    End If
    OBJ.Close
    
    OBJ.Open dsn
    Set RST = OBJ.Execute(SQL)
    Set grid.DataSource = RST
    OBJ.Close
    
    grid.Col = 0
    grid.Sort = flexSortStringAscending
    
    grid.ColWidth(0) = 1750
    grid.ColWidth(1) = 1750
End Sub
