VERSION 5.00
Begin VB.Form xm_con 
   Caption         =   "各组项目设置"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   6465
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame7 
      Height          =   735
      Left            =   120
      TabIndex        =   19
      Top             =   4440
      Width           =   2175
      Begin VB.TextBox xm_sum 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   1400
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "0"
         Top             =   210
         Width           =   500
      End
      Begin VB.Label Label4 
         Caption         =   "比赛项目个数："
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   320
         Width           =   1575
      End
   End
   Begin VB.Frame Frame6 
      Height          =   735
      Left            =   2400
      TabIndex        =   16
      Top             =   5160
      Width           =   3975
      Begin VB.CommandButton custom_xm 
         Caption         =   "自定义项目设置"
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   160
         Width           =   2055
      End
      Begin VB.CommandButton cmd_ok 
         Caption         =   "确定"
         Height          =   495
         Left            =   2880
         TabIndex        =   17
         Top             =   160
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      Height          =   5055
      Left            =   2400
      TabIndex        =   12
      Top             =   120
      Width           =   1095
      Begin VB.CommandButton Command2 
         Caption         =   "清空"
         Height          =   495
         Index           =   2
         Left            =   180
         TabIndex        =   15
         Top             =   4450
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "删除"
         Height          =   495
         Index           =   1
         Left            =   180
         TabIndex        =   14
         Top             =   3000
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "添加"
         Height          =   495
         Index           =   0
         Left            =   180
         TabIndex        =   13
         Top             =   2280
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   5160
      Width           =   2175
      Begin VB.CommandButton cp_ps 
         Caption         =   "粘贴"
         Height          =   495
         Index           =   1
         Left            =   1200
         TabIndex        =   11
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton cp_ps 
         Caption         =   "复制"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   150
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5055
      Left            =   3600
      TabIndex        =   6
      Top             =   120
      Width           =   2775
      Begin VB.ListBox all_xm 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4470
         Left            =   75
         TabIndex        =   8
         Top             =   530
         Width           =   2610
      End
      Begin VB.Label Label3 
         Caption         =   "总项目列表"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3495
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   2175
      Begin VB.ListBox zb_xm 
         Height          =   2940
         Left            =   75
         TabIndex        =   5
         Top             =   480
         Width           =   1995
      End
      Begin VB.Label Label2 
         Caption         =   "所选组比赛项目列表"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   460
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "竞赛组别选择"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   150
         Width           =   1215
      End
   End
End
Attribute VB_Name = "xm_con"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset
Dim group_dm As String
Dim zb_xm_arr() As Variant
Dim cp_arr() As Variant
Private dic_group As Object
Private dic_watch_xm As Object


Private Sub cmd_ok_Click()
    Dim i As Integer
    For i = 0 To UBound(zb_xm_arr)
        MsgBox zb_xm_arr(i)
    Next i
End Sub

Private Sub Combo1_Click()
    '加载组别项目列表
    zb_xm_load Combo1.Text
    
'    MsgBox d(group_dm)
'    group_dm = Combo1.Text

End Sub

Sub zb_xm_load(data)
    '清除zb_xm列表
    zb_xm.Clear
    Dim txtsql As String
    txtsql = "select xm_name from group_xm where group_xm_code = '" & dic_group(data) & "' order by id"
    Set rs = ExeSQL(txtsql, ydhmc)
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            zb_xm.additem rs("xm_name")
            rs.MoveNext
        Loop
        rs.Close
    End If
    '加载选定组的项目数
    xm_sum.Text = zb_xm.ListCount
    
End Sub



Private Sub Combo1_GotFocus()
    
    Dim tempStr As String
    tempStr = Join(zb_xm_arr, ",")
    If LenB(tempStr) > 0 Then
        Dim txtsql As String, i As Integer
        For i = 0 To UBound(zb_xm_arr) - 1
            txtsql = "select * from group_xm where group_xm_code ='" & dic_group(Combo1.Text) & "' and xm_name ='" & zb_xm_arr(i) & "'"
            Set rs = ExeSQL(txtsql, ydhmc)
            If rs.RecordCount = 0 Then
                rs.AddNew
                rs("xm_code") = dic_watch_xm(zb_xm_arr(i)) '数据从比赛项目字典中获得
                rs("xm_name") = zb_xm_arr(i)   '组别项目数据中获得
                rs("group_xm_code") = dic_group(Combo1.Text)   '组别字典中获得
                'MsgBox zb_xm_arr(i)
                rs.Update
            End If
            
'            If rs("xm_name") <> zb_xm_arr(i) Then
'                rs.AddNew
'                rs("xm_code") = dic_watch_xm(zb_xm_arr(i)) '数据从比赛项目字典中获得
'                rs("xm_name") = zb_xm_arr(i)   '组别项目数据中获得
'                rs("group_xm_code") = dic_group(Combo1.Text)   '组别字典中获得
'                'MsgBox zb_xm_arr(i)
'                rs.Update
'                rs.MoveNext
'            Else
'                rs("xm_code") = dic_watch_xm(zb_xm_arr(i)) '数据从比赛项目字典中获得
'                rs("xm_name") = zb_xm_arr(i)   '组别项目数据中获得
'                'MsgBox zb_xm_arr(i)
'                rs.Update
'                rs.MoveNext
'            End If
        Next i
        Erase zb_xm_arr
        rs.Close
    End If
End Sub

Private Sub Command2_Click(Index As Integer)
    Dim i As Integer
    If Index = 0 And all_xm.ListIndex <> -1 And CheckBox(zb_xm, all_xm.list(all_xm.ListIndex)) Then
        zb_xm.additem all_xm.list(all_xm.ListIndex)
        ReDim zb_xm_arr(zb_xm.ListCount)
        For i = 0 To zb_xm.ListCount - 1
            zb_xm_arr(i) = zb_xm.list(i)
        Next i
    ElseIf Index = 1 And zb_xm.ListIndex <> -1 Then
        zb_xm.removeitem zb_xm.ListIndex
        ReDim zb_xm_arr(zb_xm.ListCount)
        For i = 0 To zb_xm.ListCount - 1
            zb_xm_arr(i) = zb_xm.list(i)
        Next i
    ElseIf Index = 2 Then
        zb_xm.Clear
    End If
    
    
    If zb_xm.ListCount > 0 Then
        zb_xm.Selected(zb_xm.ListCount - 1) = True
    End If
    xm_sum.Text = zb_xm.ListCount
End Sub

'检测列表里是否有重复值
Function CheckBox(vData As Object, str As String) As Boolean
    Dim i As Integer
    For i = 0 To vData.ListCount - 1
        If str = vData.list(i) Then
            CheckBox = False '有重复不添加
            Exit Function
        End If
    Next i
    CheckBox = True '无重复可以添加
End Function

Private Sub cp_ps_Click(Index As Integer)
    Dim i As Integer
    
    If Index = 0 Then
        ReDim cp_arr(zb_xm.ListCount)
        For i = 0 To zb_xm.ListCount - 1
            cp_arr(i) = zb_xm.list(i)
        Next i
    ElseIf Index = 1 Then
        zb_xm.Clear
        For i = 0 To UBound(cp_arr) - 1
            zb_xm.additem cp_arr(i)
        Next i
        zb_xm_arr = cp_arr
    End If
End Sub

Private Sub Form_Load()
'    '判断组别字典对象是否存在
'    If d Is Nothing Then
'        group_info
'    End If
'
'    '添加组别信息
'    If d.Count = 0 Then
'        Combo1.Enabled = False
'        Combo1.additem "没有数据"
'        Combo1.Text = Combo1.list(0)
'    Else
'        Dim vkey As Variant
'        For Each vkey In d
'            Combo1.additem vkey
'        Next
'        Combo1.Text = Combo1.list(0)
'    End If

    '加载组别信息
    Set dic_group = CreateObject("Scripting.Dictionary")
    Dim txtsql As String
    txtsql = "select * from sign_group order by id"
    Set rs = ExeSQL(txtsql, ydhmc)
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            dic_group(rs("group_type").Value) = rs("group_code").Value & rs("group_sex").Value
            Combo1.additem rs("group_type").Value
            rs.MoveNext
        Loop
        rs.Close
    Else
        Combo1.Enabled = False
        Combo1.additem "没有数据"
        Combo1.Text = Combo1.list(0)
    End If
    Combo1.Text = Combo1.list(0)
    
    
    '加载项目总表
    xm_refresh
    
    group_dm = Combo1.Text
End Sub

Sub xm_refresh()
    Set dic_watch_xm = CreateObject("Scripting.Dictionary")
    Set rs = ExeSQL("select xm_code,xm_name from match_xm order by id", ydhmc)
    
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            dic_watch_xm(rs("xm_name").Value) = rs("xm_code").Value
            all_xm.additem rs("xm_name")
            rs.MoveNext
        Loop
        rs.Close
    End If
End Sub

