VERSION 5.00
Begin VB.Form create_frm 
   Caption         =   "�½��˶���"
   ClientHeight    =   5910
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   9195
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton ydhcreate 
      Caption         =   "�����˶���"
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
      Caption         =   "��"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton ydhdel 
      Caption         =   "ɾ��"
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
      Caption         =   "��ѡ���˶�������"
      BeginProperty Font 
         Name            =   "����"
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
        MsgBox "���Ʋ���Ϊ�գ�"
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
        ydh_list.AddItem "û�м�¼"
    End If
    rs.Close
    'ydh_list.Selected(2) = True
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
    conn.rs.Open txtsql, conn.OpenConn, 1, 1
    MsgBox rs("����")
    
End Sub
