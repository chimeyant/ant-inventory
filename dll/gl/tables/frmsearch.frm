VERSION 5.00
Object = "{AF7F3CA9-4499-4F24-9A04-4D8E6DC36378}#2.0#0"; "Chameleon.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmsearch 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4215
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option2 
      Caption         =   "Bank"
      BeginProperty Font 
         Name            =   "Arial"
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
      Top             =   3840
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Cash"
      BeginProperty Font 
         Name            =   "Arial"
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
      Top             =   3600
      Value           =   -1  'True
      Width           =   735
   End
   Begin Chameleon.chameleonButton cmdcancel 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      BTYPE           =   4
      TX              =   "Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      MPTR            =   1
      MICON           =   "frmsearch.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Single Click To Choose"
      Top             =   840
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3413
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.TextBox txtsearch 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label3 
      Height          =   210
      Left            =   3780
      TabIndex        =   5
      Top             =   2850
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2850
      Width           =   1935
   End
   Begin VB.Label lbltabel 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   510
      Width           =   735
   End
   Begin VB.Menu mnupopup 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu mnusort1 
         Caption         =   "Sort Coloumn 1"
      End
      Begin VB.Menu mnusort2 
         Caption         =   "Sort Coloumn 2"
      End
   End
End
Attribute VB_Name = "frmsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OBJ As New ADODB.Connection
Dim RST As New ADODB.Recordset
Dim SQL As String

Private m_SortColumn As Integer
Private m_SortAscending As Integer

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 34 Or KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 27 Then cmdCancel_Click
End Sub

Private Sub grid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        cmdCancel_Click
        Exit Sub
    End If
    If grid.TextMatrix(0, 0) = "" Then Exit Sub
    If KeyAscii = 13 Then
        If namatabel = "Transaction" Then
            hasil = grid.TextMatrix(grid.Row, 0)
            hasil1 = grid.TextMatrix(grid.Row, 1)
            hasil2 = grid.TextMatrix(grid.Row, 2)
        ElseIf namatabel = "Transaction  " Then
            hasil = grid.TextMatrix(grid.Row, 0)
            hasil1 = ""
            hasil2 = ""
        Else
            hasil = grid.TextMatrix(grid.Row, 0)
            hasil1 = grid.TextMatrix(grid.Row, 1)
            hasil2 = ""
        End If
        Unload Me
    End If
End Sub

Private Sub showtran()
    OBJ.Open dsn
    SQL = carisql1
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        grid.Clear
        grid.Rows = 2
        If namatabel = "Transaction" Then
            grid.ColWidth(0) = 1325
            grid.ColWidth(1) = 1325
            grid.ColWidth(2) = 1325
            grid.ColWidth(3) = 0
        Else
            grid.ColWidth(0) = 1325
            grid.ColWidth(1) = 0
            grid.ColWidth(2) = 0
            grid.ColWidth(3) = 0
        End If
        OBJ.Close
        Exit Sub
    End If
    Set RST = OBJ.Execute(SQL)
    Set grid.DataSource = RST
    OBJ.Close
    
    Label2 = grid.Rows - 1 & " Records"
    grid.TextMatrix(0, 0) = "> " & grid.TextMatrix(0, 0)
    Label4 = Mid$(grid.TextMatrix(0, 0), 3)
    m_SortColumn = 0
    Label3 = 0
    grid.Col = 0
    grid.Sort = flexSortStringAscending
    
    If namatabel = "Transaction" Then
        grid.ColWidth(0) = 1325
        grid.ColWidth(1) = 1325
        grid.ColWidth(2) = 1325
        grid.ColWidth(3) = 0
    Else
        grid.ColWidth(0) = 1325
        grid.ColWidth(1) = 0
        grid.ColWidth(2) = 0
        grid.ColWidth(3) = 0
    End If
End Sub

Private Sub showtabel()
    OBJ.Open dsn
    If namatabel = "Account" Then
        SQL = carisql1 + " and b.typeac = '" & setup5 & "'"
    ElseIf namatabel = "Account " Then
        SQL = carisql1 + " where typeac = '" & setup5 & "'"
    ElseIf namatabel = "Account  " Then
        SQL = carisql1 + " and (b.typeac = 'AS' or b.typeac = 'LI')"
    ElseIf namatabel = "Company Account " Then
        SQL = carisql1 + " group by b.noac,b.nmac"
    ElseIf namatabel = "Company Account  " Then
        SQL = carisql1 + " group by b.noac,b.nmac"
    ElseIf namatabel = "Balance Sheet" Then
        SQL = carisql1 + " where report_type = '1'"
    ElseIf namatabel = "Income Statement" Then
        SQL = carisql1 + " where report_type = '2'"
    ElseIf namatabel = "Cash Flow" Then
        SQL = carisql1 + " where report_type = '3'"
    ElseIf namatabel = "Buku Besar" Then
        SQL = carisql1 + " where report_type = '4'"
    ElseIf namatabel = "Fixed Assets" Then
        SQL = carisql1 + " where kdcomp = '" & setup6 & "'"
    ElseIf namatabel = "Fixed Assets " Then
        SQL = carisql1 + " where kdcomp = '" & setup6 & "' and flag <> 'N'"
    ElseIf namatabel = "Unposting Pembelian F/A" Then
        SQL = carisql1 + " where kdcomp = '" & setup6 & "' and flag = 'P'"
    ElseIf namatabel = "Posting Penjualan F/A" Then
        SQL = carisql1 + " where kdcomp = '" & setup6 & "' and flag = 'P' and curr1 <> ''"
    ElseIf namatabel = "Posting Pembelian F/A" Then
        SQL = carisql1 + " where kdcomp = '" & setup6 & "' and flag = 'N'"
    ElseIf namatabel = "Unposting Penjualan F/A" Then
        SQL = carisql1 + " where kdcomp = '" & setup6 & "' and flag = 'J'"
    ElseIf namatabel = " Fixed  Assets" Then
        SQL = carisql1 + " where kdcomp = '" & setup6 & "' and (flag = 'P' or flag = 'J')"
    ElseIf namatabel = "Group Code" Then
        SQL = carisql1 + " where form_no = '" & setup5 & "'"
    ElseIf namatabel = "Company Type " Then
        SQL = carisql1 + " group by b.kdtype,b.nmtype"
    ElseIf namatabel = "User Level" Then
        SQL = carisql1 + " group by kode,keterangan"
    
    'ElseIf namatabel = "Cash/Bank" And Option1.Value = True Then
        'sql = "select a.noac,b.nmac from gl_cash a left join gl_masterac b on a.noac=b.noac"
        
        'SQL = "select b.noac,b.nmac from gl_cash a left join gl_masterac b on a.noac=b.noac left join gl_chacct c on a.noac=c.noac where kdcomp >= '" & setup1 & "' and kdcomp <= '" & setup2 & "' group by b.noac,b.nmac"
    'ElseIf namatabel = "Cash/Bank" And Option2.Value = True Then
        'sql = "select a.noac,b.nmac from gl_bank a left join gl_masterac b on a.noac=b.noac"
        'SQL = "select b.noac,b.nmac from gl_bank a left join gl_masterac b on a.noac=b.noac left join gl_chacct c on a.noac=c.noac where kdcomp >= '" & setup1 & "' and kdcomp <= '" & setup2 & "' group by b.noac,b.nmac"
    'ElseIf namatabel = "Cash/Bank " And Option1.Value = True Then
        'sql = "select a.noac,b.nmac from gl_cash a left join gl_masterac b on a.noac=b.noac"
        
        'SQL = "select b.noac,b.nmac from gl_cash a left join gl_masterac b on a.noac=b.noac left join gl_chacct c on a.noac=c.noac where kdcomp >= '" & setup1 & "' and kdcomp <= '" & setup2 & "' group by b.noac,b.nmac"
    'ElseIf namatabel = "Cash/Bank " And Option2.Value = True Then
        'sql = "select a.noac,b.nmac from gl_bank a left join gl_masterac b on a.noac=b.noac"
        'SQL = "select b.noac,b.nmac from gl_bank a left join gl_masterac b on a.noac=b.noac left join gl_chacct c on a.noac=c.noac where kdcomp >= '" & setup1 & "' and kdcomp <= '" & setup2 & "' group by b.noac,b.nmac"
    Else
        SQL = carisql1
    End If
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        grid.Clear
        grid.Rows = 2
        grid.ColWidth(1) = 2940
        grid.ColWidth(2) = 0
        OBJ.Close
        Exit Sub
    End If
    Set RST = OBJ.Execute(SQL)
    Set grid.DataSource = RST
    OBJ.Close
    
    Label2 = grid.Rows - 1 & " Records"
    grid.TextMatrix(0, 0) = "> " & grid.TextMatrix(0, 0)
    Label4 = Mid$(grid.TextMatrix(0, 0), 3)
    m_SortColumn = 0
    Label3 = 0
    grid.Col = 0
    grid.Sort = flexSortStringAscending
    
    If namatabel = "Account" Or _
    namatabel = "Company Account" Or _
    namatabel = "Company Account " Or _
    namatabel = "Account  " Or _
    namatabel = "Account " Then
        grid.Row = 1
        Do While True
            grid.TextMatrix(grid.Row, 0) = original(grid.TextMatrix(grid.Row, 0))
            
            If (grid.Rows - 1) = grid.Row Then Exit Do
            grid.Row = grid.Row + 1
        Loop
    End If
    
    grid.ColWidth(1) = 2940
    grid.ColWidth(2) = 0
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    
    m_SortColumn = -1
    m_SortAscending = -1
    
    If namatabel = "Transaction" Or namatabel = "Transaction  " Then
        Label1.Visible = False
        txtsearch.Visible = False
        grid.Top = 480
        grid.Height = 2295
        lbltabel = "Searching Tabel " & namatabel
        showtran
        Exit Sub
    End If
    lbltabel = "Searching Tabel " & namatabel
    
    'If namatabel = "Cash/Bank" Then Me.Height = 4260
    showtabel
End Sub

Private Sub grid_Click()
    If grid.TextMatrix(0, 0) = "" Then Exit Sub
    Label3 = grid.MouseCol
    If grid.MouseRow > 0 Then
        If namatabel = "Transaction" Then
            hasil = grid.TextMatrix(grid.Row, 0)
            hasil1 = grid.TextMatrix(grid.Row, 1)
            hasil2 = grid.TextMatrix(grid.Row, 2)
        ElseIf namatabel = "Transaction  " Then
            hasil = grid.TextMatrix(grid.Row, 0)
            hasil1 = ""
            hasil2 = ""
        Else
            hasil = grid.TextMatrix(grid.Row, 0)
            hasil1 = grid.TextMatrix(grid.Row, 1)
            hasil2 = ""
        End If
        Unload Me
        Exit Sub
    End If
    If grid.MouseCol <> m_SortColumn Then
        If m_SortColumn >= 0 Then
            grid.TextMatrix(0, m_SortColumn) = _
                Mid$(grid.TextMatrix(0, m_SortColumn), 3)
        End If
        m_SortColumn = grid.MouseCol
        
        m_SortAscending = True
        grid.TextMatrix(0, m_SortColumn) = _
            "> " & grid.TextMatrix(0, m_SortColumn)
    Else
        m_SortAscending = Not m_SortAscending
        
        If m_SortAscending Then
            grid.TextMatrix(0, m_SortColumn) = _
                "> " & Mid$(grid.TextMatrix(0, m_SortColumn), 3)
        Else
            grid.TextMatrix(0, m_SortColumn) = _
                "< " & Mid$(grid.TextMatrix(0, m_SortColumn), 3)
        End If
    End If
    
    Label4 = Mid$(grid.TextMatrix(0, Label3), 3)
    grid.Row = 1
    grid.RowSel = grid.Rows - 1
    grid.Col = m_SortColumn

    If m_SortAscending Then
        grid.Sort = flexSortStringAscending
    Else
        grid.Sort = flexSortStringDescending
    End If
    
    If txtsearch.Visible = True Then txtsearch.SetFocus
End Sub

Private Sub Option1_Click()
    showtabel
End Sub

Private Sub Option2_Click()
    showtabel
End Sub

Private Sub txtsearch_Change()
    OBJ.Open dsn
    If namatabel = "Account" Then
        SQL = carisql1 + " and b.typeac = '" & setup5 & "' and b." + Label4 + " like '" + x_original(txtsearch) + "%'"
    ElseIf namatabel = "Account " Then
        SQL = carisql1 + " where typeac = '" & setup5 & "' and " + Label4 + " like '" + x_original(txtsearch) + "%'"
    ElseIf namatabel = "Account  " Then
        SQL = carisql1 + " and (b.typeac = 'AS' or b.typeac = 'LI') and b." + Label4 + " like '" + x_original(txtsearch) + "%'"
    ElseIf namatabel = "Company Account" Then
        SQL = carisql1 + " and b." + Label4 + " like '" + x_original(txtsearch) + "%'"
    ElseIf namatabel = "Company Account " Then
        SQL = carisql1 + " and b." + Label4 + " like '" + x_original(txtsearch) + "%' group by b.noac,b.nmac"
    ElseIf namatabel = "Company Account  " Then
        SQL = carisql1 + " and b." + Label4 + " like '" + x_original(txtsearch) + "%' group by b.noac,b.nmac"
    ElseIf namatabel = "Balance Sheet" Then
        SQL = carisql1 + " where report_type = '1' and " + Label4 + " like '" + txtsearch + "%'"
    ElseIf namatabel = "Income Statement" Then
        SQL = carisql1 + " where report_type = '2' and " + Label4 + " like '" + txtsearch + "%'"
    ElseIf namatabel = "Cash Flow" Then
        SQL = carisql1 + " where report_type = '3' and " + Label4 + " like '" + txtsearch + "%'"
    ElseIf namatabel = "Buku Besar" Then
        SQL = carisql1 + " where report_type = '4' and " + Label4 + " like '" + txtsearch + "%'"
    ElseIf namatabel = "Fixed Assets" Then
        SQL = carisql1 + " where kdcomp = '" & setup6 & "' and " + Label4 + " like '" + txtsearch + "%'"
    ElseIf namatabel = "Fixed Assets " Then
        SQL = carisql1 + " where kdcomp = '" & setup6 & "' and " + Label4 + " like '" + txtsearch + "%' and flag <> 'N'"
    ElseIf namatabel = "Unposting Pembelian F/A" Then
        SQL = carisql1 + " where kdcomp = '" & setup6 & "' and flag = 'P' and " + Label4 + " like '" + txtsearch + "%'"
    ElseIf namatabel = "Posting Penjualan F/A" Then
        SQL = carisql1 + " where kdcomp = '" & setup6 & "' and flag = 'P' and " + Label4 + " like '" + txtsearch + "%' and curr1 <> ''"
    ElseIf namatabel = "Posting Pembelian F/A" Then
        SQL = carisql1 + " where kdcomp = '" & setup6 & "' and flag = 'N' and " + Label4 + " like '" + txtsearch + "%'"
    ElseIf namatabel = "Unposting Penjualan F/A" Then
        SQL = carisql1 + " where kdcomp = '" & setup6 & "' and flag = 'J' and " + Label4 + " like '" + txtsearch + "%'"
    ElseIf namatabel = " Fixed  Assets" Then
        SQL = carisql1 + " where kdcomp = '" & setup6 & "' and (flag = 'P' or flag = 'J') and " + Label4 + " like '" + txtsearch + "%'"
    ElseIf namatabel = "Group Code" Then
        SQL = carisql1 + " where form_no = '" & setup5 & "' and " + Label4 + " like '" + txtsearch + "%'"
    ElseIf namatabel = "Company Type " Then
        SQL = carisql1 + " where b." + Label4 + " like '" + txtsearch + "%' group by b.kdtype,b.nmtype"
    ElseIf namatabel = "User Level" Then
        SQL = carisql1 + " and " + Label4 + " like '" + txtsearch + "%' group by kode,keterangan"
    ElseIf namatabel = "User" Then
        SQL = carisql1 + " and " + Label4 + " like '" + txtsearch + "%'"
    
    'ElseIf namatabel = "Cash/Bank" And Option1.Value = True Then
        'sql = "select noac from gl_cash where noac like '" + x_original(txtsearch) + "%'"
        'sql = "select a.noac,b.nmac from gl_cash a left join gl_masterac b on a.noac=b.noac where a.noac like '" + x_original(txtsearch) + "%'"
        
        'SQL = "select b.noac,b.nmac from gl_cash a left join gl_masterac b on a.noac=b.noac left join gl_chacct c on a.noac=c.noac where kdcomp >= '" & setup1 & "' and kdcomp <= '" & setup2 & "' and b." + Label4 + " like '" + x_original(txtsearch) + "%' group by b.noac,b.nmac"
    'ElseIf namatabel = "Cash/Bank" And Option2.Value = True Then
        'sql = "select noac from gl_bank where noac like '" + x_original(txtsearch) + "%'"
        'sql = "select a.noac,b.nmac from gl_bank a left join gl_masterac b on a.noac=b.noac where a.noac like '" + x_original(txtsearch) + "%'"
        
        'SQL = "select b.noac,b.nmac from gl_bank a left join gl_masterac b on a.noac=b.noac left join gl_chacct c on a.noac=c.noac where kdcomp >= '" & setup1 & "' and kdcomp <= '" & setup2 & "' and b." + Label4 + " like '" + x_original(txtsearch) + "%' group by b.noac,b.nmac"
    'ElseIf namatabel = "Cash/Bank " And Option1.Value = True Then
        'sql = "select noac from gl_cash where noac like '" + x_original(txtsearch) + "%'"
        'sql = "select a.noac,b.nmac from gl_cash a left join gl_masterac b on a.noac=b.noac where a.noac like '" + x_original(txtsearch) + "%'"
        
        'SQL = "select b.noac,b.nmac from gl_cash a left join gl_masterac b on a.noac=b.noac left join gl_chacct c on a.noac=c.noac where kdcomp >= '" & setup1 & "' and kdcomp <= '" & setup2 & "' and b." + Label4 + " like '" + x_original(txtsearch) + "%' group by b.noac,b.nmac"
    'ElseIf namatabel = "Cash/Bank " And Option2.Value = True Then
        'sql = "select noac from gl_bank where noac like '" + x_original(txtsearch) + "%'"
        'sql = "select a.noac,b.nmac from gl_bank a left join gl_masterac b on a.noac=b.noac where a.noac like '" + x_original(txtsearch) + "%'"
        
        'SQL = "select b.noac,b.nmac from gl_bank a left join gl_masterac b on a.noac=b.noac left join gl_chacct c on a.noac=c.noac where kdcomp >= '" & setup1 & "' and kdcomp <= '" & setup2 & "' and b." + Label4 + " like '" + x_original(txtsearch) + "%' group by b.noac,b.nmac"
    Else
        SQL = carisql1 + " where " + Label4 + " like '" + txtsearch + "%'"
    End If
    Set RST = OBJ.Execute(SQL)
    If RST.EOF Then
        grid.Clear
        grid.Rows = 2
        grid.ColWidth(1) = 2940
        grid.ColWidth(2) = 0
        OBJ.Close
        Label2 = ""
        Exit Sub
    End If
    Set RST = OBJ.Execute(SQL)
    Set grid.DataSource = RST
    Label2 = grid.Rows - 1 & " Records"
    OBJ.Close
    grid.TextMatrix(0, Label3) = _
            "> " & grid.TextMatrix(0, Label3)
    grid.Sort = flexSortStringAscending
    
    If namatabel = "Account" Or _
    namatabel = "Company Account" Or _
    namatabel = "Company Account " Or _
    namatabel = "Account  " Or _
    namatabel = "Account " Then
        grid.Row = 1
        Do While True
            grid.TextMatrix(grid.Row, 0) = original(grid.TextMatrix(grid.Row, 0))
            
            If (grid.Rows - 1) = grid.Row Then Exit Do
            grid.Row = grid.Row + 1
        Loop
    End If
    
    grid.ColWidth(1) = 2940
    grid.ColWidth(2) = 0
End Sub


Private Sub txtsearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        cmdCancel_Click
        Exit Sub
    End If
    If KeyAscii = 13 And grid.Rows = 2 Then
        If namatabel = "Transaction" Then
            hasil = grid.TextMatrix(1, 0)
            hasil1 = grid.TextMatrix(1, 1)
            hasil2 = grid.TextMatrix(1, 2)
        ElseIf namatabel = "Transaction  " Then
            hasil = grid.TextMatrix(1, 0)
            hasil1 = ""
            hasil2 = ""
        Else
            hasil = grid.TextMatrix(1, 0)
            hasil1 = grid.TextMatrix(1, 1)
            hasil2 = ""
        End If
        Unload Me
        Exit Sub
    End If
    If Label3 = "" Then KeyAscii = 0
End Sub


