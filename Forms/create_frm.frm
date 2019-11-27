VERSION 5.00
Begin VB.Form create_frm 
   Caption         =   "新建运动会"
   ClientHeight    =   4935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   4185
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   3975
      Begin VB.CommandButton ydhopen 
         Caption         =   "打开"
         Height          =   375
         Left            =   2640
         TabIndex        =   7
         Top             =   160
         Width           =   975
      End
      Begin VB.CommandButton ydhdel 
         Caption         =   "删除"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   160
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "请选择运动会名称"
      Height          =   3375
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   3975
      Begin VB.ListBox ydh_list 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3000
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "新建运动会"
      Height          =   700
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.CommandButton ydhcreate 
         Caption         =   "新建"
         Height          =   375
         Left            =   3120
         TabIndex        =   2
         Top             =   200
         Width           =   735
      End
      Begin VB.TextBox txtydhmc 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "create_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DBtool As New DBcls
Dim rs As ADODB.Recordset
    
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
    DBtool.SetConnToFile App.Path & "\bpdata.mdb"
    Dim res As Long
    Dim txtsql As String
    txtsql = "INSERT INTO ydh ([ydh_name],[is_open]) VALUES ('" & ydhmc & "',true)"
    res = DBtool.ExecNonQuery(txtsql)
    txtsql = "UPDATE ydh SET is_open = 0 WHERE ydh_name <> '" & ydhmc & "'"
    res = DBtool.ExecNonQuery(txtsql)
    
    
'    Do While Not rs.EOF
'        rs.Fields("is_open") = False
'        rs.MoveNext
'    Loop
'    rs.AddNew
'    rs.Fields("ydh_name") = ydhmc
'    rs.Fields("is_open") = True
'    rs.Update
'    rs.Close
    If Dir(App.Path & "\" & ydhmc) = "" Then
        MkDir (App.Path & "\" & ydhmc)
    End If
    create_data ydhmc
    ydh_refresh
End Sub
Sub ydh_refresh()
    ydh_list.Clear
    Dim n As Integer
    DBtool.SetConnToFile App.Path & "\bpdata.mdb"
    Set rs = DBtool.ExecQuery("select * from ydh")
    
    If rs.RecordCount = 0 Then
        DBtool.ReleaseRecordset rs
        Exit Sub
    End If
    Do While Not rs.EOF
        n = n + 1
        ydh_list.AddItem rs.Fields("ydh_name")
        If rs.Fields("is_open") = True Then
            ydh_list.Selected(n - 1) = True
        End If
        rs.MoveNext
    Loop
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

