VERSION 5.00
Begin VB.Form create_frm 
   Caption         =   "�½��˶���"
   ClientHeight    =   4935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   4185
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   4200
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
      Height          =   3375
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   3975
      Begin VB.ListBox ydh_list 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3000
         ItemData        =   "create_frm.frx":0000
         Left            =   120
         List            =   "create_frm.frx":0002
         TabIndex        =   4
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "�½��˶���"
      Height          =   700
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.CommandButton ydhcreate 
         Caption         =   "�½�"
         Height          =   375
         Left            =   3120
         TabIndex        =   2
         Top             =   200
         Width           =   735
      End
      Begin VB.TextBox txtydhmc 
         BeginProperty Font 
            Name            =   "����"
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
Option Explicit

Dim DBtool As New DBcls
'Dim Rs As ADODB.Recordset
    
Private Sub Form_Load()
    ydh_refresh
    'cn.Close
    'Main.out.Enabled = False
End Sub

Private Sub ydhcreate_Click()
    Dim txtmc As String
    txtmc = Trim(txtydhmc.Text)
    If txtmc = "" Then
        MsgBox "���Ʋ���Ϊ�գ�"
        txtydhmc.Text = ""
        txtydhmc.SetFocus
        Exit Sub
    End If
    res.MoveFirst
    Do While Not res.EOF
        res("is_open") = False
        res.MoveNext
    Loop
    res.AddNew
    res("ydh_name") = txtmc
    res("is_open") = True
    res.Update
    If Dir(App.Path & "\" & txtmc) = "" Then
        MkDir (App.Path & "\" & txtmc)
    End If
    create_data txtmc
    ydh_refresh
End Sub
Sub ydh_refresh()
    ydh_list.Clear
    Dim i As Integer
    res.MoveFirst
    For i = 0 To res.RecordCount - 1
        ydh_list.AddItem res("ydh_name")
        If res("is_open") = True Then
            ydh_list.Selected(i) = True
        End If
        res.MoveNext
    Next i
End Sub

Private Sub ydhdel_Click()
    Dim m As String
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
    Dim rs As ADODB.Recordset
    ydhmc = ydh_list.Text
    Set rs = ExeSQL("select * from MyTable", ydhmc)
    'MsgBox rs.RecordCount
    
    Main.Caption = "�ﾶ�˶�����������ϵͳ" & "  ��ǰ�˶��᣺" & ydhmc
    Unload Me
    If res.RecordCount > 0 Then
        res.MoveFirst
        Do While Not res.EOF
            res("is_open") = IIf(res("ydh_name") = ydhmc, True, False)
            res.MoveNext
        Loop
    End If
End Sub
