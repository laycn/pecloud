VERSION 5.00
Begin VB.Form create_frm 
   Caption         =   "�½��˶���"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   4305
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   4560
      Width           =   3975
      Begin VB.CommandButton ydhopen 
         Caption         =   "��"
         Height          =   375
         Left            =   2640
         TabIndex        =   7
         Top             =   160
         Width           =   975
      End
      Begin VB.CommandButton ydhdel 
         Caption         =   "ɾ��"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   160
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "��ѡ���˶�������"
      Height          =   3615
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   3975
      Begin VB.ListBox ydh_list 
         Height          =   3120
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "�½��˶���"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.CommandButton ydhcreate 
         Caption         =   "�½�"
         Height          =   375
         Left            =   3120
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtydhmc 
         Height          =   375
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
        MsgBox "���Ʋ���Ϊ�գ�"
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
    m = MsgBox("���Ƿ����Ҫɾ������˶�����", 17, "ɾ����ʾ")
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
    
    '�������ݿ�
    cat.create pstr
    
    
    Dim tbl As New Table
    cat.ActiveConnection = pstr
    tbl.Name = "MyTable" '�������
    tbl.Columns.Append "���", adInteger '��ĵ�һ���ֶ�
    tbl.Columns.Append "����", adVarWChar, 8 '��ĵڶ����ֶ�
    tbl.Columns.Append "סַ", adVarWChar, 50 '��ĵ������ֶ�
    cat.Tables.Append tbl '�������ݱ�
    
    conn.Open pstr
    rs.CursorLocation = adUseClient
    rs.Open "MyTable", conn, adOpenKeyset, adLockPessimistic
    rs.AddNew '����������¼�¼
    rs.Fields(0).Value = 9801
    rs.Fields(1).Value = "�����"
    rs.Fields(2).Value = "�����л���ɽ"
    rs.Update
    conn.Close
End Sub

Private Sub ydhopen_Click()
    Dim conn As New DBcls
    conn.ydhmc = ydh_list.Text
    txtsql = "select * from MyTable"
    conn.rs.Open txtsql, conn.openConn, 1, 1
    MsgBox rs("����")
    
End Sub
