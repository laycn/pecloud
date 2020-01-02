VERSION 5.00
Begin VB.Form custom_xm 
   Caption         =   "比赛项目管理及自定义项目设置"
   ClientHeight    =   5010
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   5370
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   2400
      TabIndex        =   14
      Top             =   4080
      Width           =   2895
      Begin VB.CommandButton Command3 
         Caption         =   "保存退出"
         Height          =   495
         Left            =   1440
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   2400
      TabIndex        =   12
      Top             =   3240
      Width           =   2895
      Begin VB.CommandButton Command2 
         Caption         =   "删除选定项目"
         Height          =   495
         Left            =   200
         TabIndex        =   13
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3255
      Left            =   2400
      TabIndex        =   3
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton Command1 
         Caption         =   "添加自定义项目"
         Height          =   495
         Left            =   200
         TabIndex        =   11
         Top             =   2640
         Width           =   2535
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Index           =   1
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2040
         Width           =   2415
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Index           =   0
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1120
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   1200
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   620
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "详细分类："
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "新项目类别："
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "新项目名称："
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   700
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "定义新项目"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      Begin VB.ListBox watch_xm_list 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4050
         Left            =   60
         TabIndex        =   1
         Top             =   600
         Width           =   2150
      End
      Begin VB.Label Label1 
         Caption         =   "项目列表"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   220
         Width           =   975
      End
   End
End
Attribute VB_Name = "custom_xm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private colType As New Collection
Private Sub Combo1_Click(Index As Integer)
    If Index = 0 Then
        Combo1(1).Clear
        If Combo1(0).Text = "径赛" Then
            Combo1(1).additem "直道"
            Combo1(1).additem "弯道"
            Combo1(1).additem "长跑"
            Combo1(1).additem "接力"
        ElseIf Combo1(0).Text = "田赛" Then
            Combo1(1).additem "高度"
            Combo1(1).additem "远度"
            Combo1(1).additem "投掷"
        ElseIf Combo1(0).Text = "计数" Then
            Combo1(1).additem "跳绳"
        End If
        Combo1(1).Text = Combo1(1).list(0)
    End If
End Sub

Private Sub Command1_Click()
    Dim rs As ADODB.Recordset
    Dim code_num, i As Integer
    Dim code_type() As String
    Dim txtsql As String
    txtsql = "select * from match_xm where is_system = false"
    Set rs = ExeSQL(txtsql, ydhmc)
    
    If rs.RecordCount = 0 Then
        code_num = 100
    Else
        rs.MoveLast
        code_num = rs("xm_code").Value + 1
    End If
    
    rs.AddNew
    rs("xm_code") = code_num
    rs("xm_name") = Trim(Text1.Text)
    
    ReDim code_type(Len(colType(Combo1(1).Text)))
    For i = 1 To Len(colType(Combo1(1).Text))
        code_type(i) = Mid(colType(Combo1(1).Text), i, 1)
    Next
    rs("xm_type") = Val(code_type(1))
    rs("xm_type_xx") = Val(code_type(2))
    rs.Update
    rs.Close
    '显示项目列表
    watch_refresh
    MsgBox "添加成功"
End Sub

Private Sub Form_Load()
    '窗体居中显示
    With Screen
        Me.Left = (.Width - Me.Width) / 2
        Me.Top = (.Height - Me.Height) / 2
    End With
    
    '初始化文本框
    Text1.Text = ""
    
    '显示项目列表
    watch_refresh
    
    '加载列表信息
    combo1_load
    
    '加载项目类型代码
    matchType
    
End Sub

Sub combo1_load()
    Combo1(0).additem "径赛"
    Combo1(0).additem "田赛"
    Combo1(0).additem "计数"
    Combo1(0).Text = Combo1(0).list(0)
End Sub

Sub matchType()
    With colType
    .add "11", "直道"
    .add "12", "弯道"
    .add "13", "长跑"
    .add "13", "接力"
    .add "21", "高度"
    .add "22", "远度"
    .add "23", "投掷"
    .add "31", "跳绳"
    End With
End Sub

Sub watch_refresh()
    Dim rs As ADODB.Recordset
    Set rs = ExeSQL("select xm_name from match_xm order by id", ydhmc)
    
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            watch_xm_list.Clear
            watch_xm_list.additem rs("xm_name")
            rs.MoveNext
        Loop
        rs.Close
    End If
End Sub
