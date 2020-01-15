VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form sign_unit 
   Caption         =   "报名单位设置"
   ClientHeight    =   6795
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   10635
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame5 
      Height          =   855
      Left            =   7320
      TabIndex        =   20
      Top             =   4620
      Width           =   3135
      Begin VB.CommandButton del_cmd 
         Caption         =   "删除全部"
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   22
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton del_cmd 
         Caption         =   "删除"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   21
         Top             =   280
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   7320
      TabIndex        =   17
      Top             =   5520
      Width           =   3135
      Begin VB.CommandButton save_ok 
         Caption         =   "确定"
         Height          =   375
         Index           =   1
         Left            =   2100
         TabIndex        =   19
         Top             =   300
         Width           =   855
      End
      Begin VB.CommandButton save_ok 
         Caption         =   "保存"
         Height          =   375
         Index           =   0
         Left            =   200
         TabIndex        =   18
         Top             =   300
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "增加一批单位"
      Height          =   2295
      Left            =   7320
      TabIndex        =   3
      Top             =   2160
      Width           =   3135
      Begin VB.CommandButton add_more 
         Caption         =   "批量增加"
         Height          =   375
         Left            =   1800
         TabIndex        =   16
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   15
         Text            =   "10"
         Top             =   840
         Width           =   500
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Index           =   1
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   400
         Width           =   1960
      End
      Begin VB.Label Label3 
         Caption         =   "添加完成后可以在左边表格里修改数据"
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   120
         TabIndex        =   23
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "增加数量"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "组别选择"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "增加一个单位"
      Height          =   1935
      Left            =   7320
      TabIndex        =   2
      Top             =   120
      Width           =   3135
      Begin VB.CommandButton add_single 
         Caption         =   "添加"
         Height          =   375
         Left            =   1920
         TabIndex        =   11
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   1
         Left            =   1080
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   0
         Left            =   1080
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   720
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Index           =   0
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   320
         Width           =   1960
      End
      Begin VB.Label Label1 
         Caption         =   "单位全称"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "单位简称"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "组别选择"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6375
      Left            =   40
      TabIndex        =   0
      Top             =   40
      Width           =   7095
      Begin VB.TextBox txtQty 
         Height          =   270
         Left            =   3000
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   4680
         Width           =   735
      End
      Begin MSComctlLib.ListView unit_list 
         Height          =   6195
         Left            =   40
         TabIndex        =   1
         Top             =   120
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   10927
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "sign_unit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
'Dim d As Object
Dim unit_code As Object
Dim group_arr()

Private x As New clslist
'Private Px As Single, Py As Single

Private Sub add_more_Click()
    '组织数据
    Dim zb As String, unit_num As Integer
    zb = Combo1(1).Text
    unit_num = Val(Trim(Text2.Text))
    If unit_num = 0 Then
        MsgBox "数量不能为空"
        Exit Sub
    End If
    
    Dim i As Integer
    For i = 1 To unit_num
        x.additem 0, unit_code(zb) + i, zb, "新单位" & i, "新单位" & i
    Next i
    unit_code(zb) = unit_code(zb) + unit_num
End Sub

Private Sub add_single_Click()
    '组织数据
    Dim jc, qc, zb As String
    jc = Text1(0).Text
    qc = Text1(1).Text
    zb = Trim(Combo1(0).Text)
    
    Set rs = ExeSQL("select * from sign_unit", ydhmc)
    rs.AddNew
    rs("unit_code") = unit_code(zb) + 1
    rs("short_name") = jc
    rs("unit_name") = qc
    rs("unit_group") = zb
    rs.Update
    rs.Close
    'MsgBox "添加成功！"
    
    unit_code(zb) = unit_code(zb) + 1
    unit_refresh
End Sub

Private Sub del_cmd_Click(index As Integer)
    If index = 1 Then
        unit_list.ListItems.Clear
    ElseIf index = 0 Then
        If unit_list.ListItems.Count = 0 Then
            MsgBox "记录为空，不能删除"
            Exit Sub
        End If
        
        If unit_list.SelectedItem.index <> 0 Then
            Dim txtsql As String
            txtsql = "select * from sign_unit where id = " & Val(unit_list.ListItems(unit_list.SelectedItem.index).Text)
            Set rs = ExeSQL(txtsql, ydhmc)
            rs.Delete
            rs.Update
            rs.Close
        End If
    
        Set x.list = unit_list
        Set x.textbox = txtQty
        x.Ismvartext
        x.removeitem unit_list.SelectedItem.index
    End If
End Sub

Private Sub Form_Load()
    '窗体居中
    With Screen
        Me.Left = (.Width - Me.Width) / 2
        Me.Top = (.Height - Me.Height) / 2
    End With
    
    '初始化控件
    Text1(0).Text = ""
    Text1(1).Text = ""
    
    '加载数据
    '组别信息存入字典
    '判断组别字典对象是否存在
    If d Is Nothing Then
        group_info
    End If
    
    If d.Count > 0 Then
        group_arr = d.Keys
        ReDim group_code_arr(d.Count - 1)
        Dim i As Integer
        For i = 0 To UBound(group_arr)
            Dim txtsql As String
            Dim rs1 As ADODB.Recordset
            txtsql = "select id,unit_code,unit_group from sign_unit where unit_group ='" & group_arr(i) & "' order by id DESC"
            Set rs1 = ExeSQL(txtsql, ydhmc)
            Set unit_code = CreateObject("Scripting.Dictionary")
            If rs1.RecordCount >= 1 Then
                rs1.MoveFirst
                unit_code(group_arr(i)) = rs1("unit_code").Value
                rs1.Close
            Else
                unit_code(group_arr(i)) = d(group_arr(i)) * 1000
            End If
        Next i
        
        '加载combo组别控件
        Dim vkey As Variant
        For Each vkey In d
            Combo1(0).additem vkey
            Combo1(1).additem vkey
        Next
        Combo1(0).Text = Combo1(0).list(0)
        Combo1(1).Text = Combo1(1).list(0)
    Else
        For i = 0 To 1
            Combo1(i).Enabled = False
            Combo1(i).additem "没有数据"
            Combo1(i).Text = Combo1(i).list(0)
        Next i
    End If
    
    
    '加载listview标题
    Set x = Nothing
    Set x.list = unit_list
    Set x.textbox = txtQty
    
    x.addcolumn "编号", "id", 0, False, False
    x.addcolumn "单位编号", "unit_code", 1000, False, False
    x.addcolumn "竞赛组", "jgroup", 1200, False, False
    x.addcolumn "单位简称", "short_name", 1200, False, True
    x.addcolumn "单位全称", "unit_name", 2000, False, True
    
    '加载参赛队伍数据
    unit_refresh

    
    
End Sub
Sub unit_refresh()
    
    unit_list.ListItems.Clear
    
    Set rs = ExeSQL("select * from sign_unit", ydhmc)
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            x.additem rs("id"), rs("unit_code"), rs("unit_group"), rs("short_name"), rs("unit_name")
            rs.MoveNext
        Loop
        rs.Close
    End If
End Sub

Private Sub Form_Unload(cancel As Integer)
    Set rs = Nothing
End Sub

Private Sub save_ok_Click(index As Integer)
    If index = 0 Then
        '获取listview里全部数据
        Dim rs2 As ADODB.Recordset
        Set rs2 = ExeSQL("select * from sign_unit", ydhmc)
        Dim i As Integer
        
        If unit_list.ListItems.Count > 0 Then
            For i = 1 To unit_list.ListItems.Count
                If unit_list.ListItems(i).Text > 0 Then
                    'Rs2("short_name") = unit_list.ListItems(i).Text
                    rs2("short_name") = unit_list.ListItems(i).SubItems(3)
                    rs2("unit_name") = unit_list.ListItems(i).SubItems(4)
                    rs2.Update
                    rs2.MoveNext
                Else
                    rs2.AddNew
                    rs2("unit_code") = unit_list.ListItems(i).SubItems(1)
                    rs2("unit_group") = unit_list.ListItems(i).SubItems(2)
                    rs2("short_name") = unit_list.ListItems(i).SubItems(3)
                    rs2("unit_name") = unit_list.ListItems(i).SubItems(4)
                    rs2.Update
                    rs2.MoveNext
                
                End If
            Next i
        Else
            Do While Not rs2.EOF
                rs2.Delete
                rs2.Update
                rs2.MoveNext
            Loop
        End If
        rs2.Close
        MsgBox "保存成功！"
    ElseIf index = 1 Then
        Unload Me
    End If
    
End Sub

Private Sub unit_list_GotFocus()
    'MsgBox unit_list.SelectedItem.SubItems(4)
    'MsgBox unit_list.ListItems(unit_list.SelectedItem.index).Text
End Sub
