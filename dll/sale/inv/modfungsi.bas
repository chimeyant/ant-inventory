Attribute VB_Name = "modfungsi"
 Private OBJ As New ADODB.Connection
 Private RST As ADODB.Recordset
 
 Sub subme()
    OBJ.Open dsn
    SQL = "select * from am_options"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        par1 = RST!para1
        par2 = RST!para2
        par3 = RST!para3
        par4 = RST!para4
        par5 = RST!para6
    Else
        par1 = "0"
        par2 = "0"
        par3 = "0"
        par4 = "0"
        par5 = "0"
    End If
    
    SQL = "select * from am_branch"
    Set RST = OBJ.Execute(SQL)
    If Not RST.EOF Then
        kar_1 = RST!kode1
        kar_2 = RST!kode2
        kar_3 = RST!kode3
        kar_4 = RST!kode4
        
        
    Else
        kar_1 = ""
        kar_2 = ""
        kar_3 = ""
        kar_4 = ""
    End If
    
    'SQL = "select * from am_user where kodeuser like '1%'"
    'Set RST = OBJ.Execute(SQL)
    'If Not RST.EOF Then
    '    flag = False
        
    '    Me.Height = 3360
    'Else
    '    Me.Height = 2280
    '    flag = True
    '    kuser = "q"
    '    nmuser = "-no user-"
    'End If
    OBJ.Close
End Sub

