VERSION 5.00
Begin VB.Form create_frm 
   Caption         =   "新建运动会"
   ClientHeight    =   5910
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   9195
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton ydhcreate 
      Caption         =   "创建运动会"
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox txtydhmc 
      Height          =   375
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   5280
      Width           =   3015
   End
   Begin VB.CommandButton ydhopen 
      Caption         =   "打开"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton ydhdel 
      Caption         =   "删除"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   4080
      Width           =   1215
   End
   Begin VB.ListBox ydh_list 
      Height          =   3120
      ItemData        =   "create_frm.frx":0000
      Left            =   0
      List            =   "create_frm.frx":0002
      TabIndex        =   0
      Top             =   600
      Width           =   4575
   End
   Begin VB.Label ydh_name 
      Caption         =   "请选择运动会名称"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "create_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    ydh_refresh
    'cn.Close
    'Main.out.Enabled = False
End Sub
Private Sub ydhcreate_Click()
    ydhmc = Trim(txtydhmc.Text)
    If ydhmc = "" Then
        MsgBox "名称不能为空！"
        txtydhmc.Text = ""
        txtydhmc.SetFocus
        Exit Sub
    End If
    txtsql = "select * from ydh where is_open = true"
    rs.Open txtsql, cn, 2, 3
    rs.MoveFirst
    Do While Not rs.EOF
        rs.Fields("is_open") = False
        rs.MoveNext
    Loop
    rs.AddNew
    rs.Fields("ydh_name") = ydhmc
    rs.Fields("is_open") = True
    rs.Update
    rs.Close
    If Dir(App.Path & "\" & ydhmc) = "" Then
        MkDir (App.Path & "\" & ydhmc)
    End If
    create_data ydhmc
    ydh_refresh
End Sub
Sub ydh_refresh()
    ydh_list.Clear
    txtsql = "select * from ydh"
    rs.Open txtsql, cn, 1, 1
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            n = n + 1
            ydh_list.AddItem rs.Fields("ydh_name")
            If rs.Fields("is_open") = True Then
                ydh_list.Selected(n - 1) = True
            End If
            rs.MoveNext
        Loop
        
    Else
        ydh_list.AddItem "没有记录"
    End If
    rs.Close
    'ydh_list.Selected(2) = True
End Sub

Private Sub ydhdel_Click()
    m = MsgBox("您是否真的要删除这届运动会吗？", 17, "删除提示")
    If m = "vbyes" Then
        
    End If
End Sub

Sub create_data(ydhmc)
    Dim cat As New ADOX.Catalog
    Dim conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim pstr As String
    'Set cat = New ADOX.Catalog
    pstr = "Provider=Microsoft.Jet.OLEDB.4.0;"
    pstr = pstr & "Data Source=" & App.Path & "\" & ydhmc & "\" & "sdata.mdb;"
    
    
    'pstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path & "\edit.mdb" + ";"
    
    '创建数据库
    cat.create pstr
    
    
    Dim tbl As New Table
    cat.ActiveConnection = pstr
    tbl.Name = "MyTable" '表的名称
    tbl.Columns.Append "编号", adInteger '表的第一个字段
    tbl.Columns.Append "姓名", adVarWChar, 8 '表的第二个字段
    tbl.Columns.Append "住址", adVarWChar, 50 '表的第三个字段
    cat.Tables.Append tbl '建立数据表
    
    conn.Open pstr
    rs.CursorLocation = adUseClient
    rs.Open "MyTable", conn, adOpenKeyset, adLockPessimistic
    rs.AddNew '往表中添加新记录
    rs.Fields(0).Value = 9801
    rs.Fields(1).Value = "孙悟空"
    rs.Fields(2).Value = "广州市花果山"
    rs.Update
    conn.Close
End Sub

Private Sub ydhopen_Click()
    Dim conn As New DBcls
    conn.ydhmc = ydh_list.Text
    txtsql = "select * from MyTable"
    conn.rs.Open txtsql, conn.OpenConn, 1, 1
    MsgBox rs("姓名")
    
End Sub
