VERSION 5.00
Begin VB.Form xm_con 
   Caption         =   "各组项目设置"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
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
      Begin VB.CommandButton custom_xm_cmd 
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
         Top             =   4440
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
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2790
         Left            =   75
         TabIndex        =   5
         Top             =   540
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



Private dic_group As Object
Private dic_watch_xm As Object
Private zb_xm_arr() As Variant
Private cp_arr() As Variant
Private old_arr() As Variant

Private Sub all_xm_DblClick()
    add_dbEvent
    '更新显示项目数量
    update_xm_sum
End Sub

Private Sub add_dbEvent()
    If CheckBox(zb_xm, all_xm.list(all_xm.ListIndex)) Then
        If IsNotEmpty(zb_xm_arr) = False Then
            ReDim zb_xm_arr(4, 0)
            zb_xm_arr(0, 0) = dic_watch_xm(all_xm.list(all_xm.ListIndex))
            zb_xm_arr(1, 0) = all_xm.list(all_xm.ListIndex)
            zb_xm_arr(2, 0) = dic_group(Combo1.Text)
            zb_xm_arr(3, 0) = Mid(dic_group(Combo1.Text), 2, 1)
            zb_xm_arr(4, 0) = True
        End If
        Dim flag_xm, i As Integer
        flag_xm = 0
        For i = 0 To UBound(zb_xm_arr, 2)
            If zb_xm_arr(1, i) = all_xm.list(all_xm.ListIndex) And zb_xm_arr(2, i) = dic_group(Combo1.Text) Then
                If flag_xm = 0 Then
                    zb_xm_arr(4, i) = True
                    flag_xm = flag_xm + 1
                Else
                    zb_xm_arr(4, i) = False
                End If
            End If
        Next i
        If flag_xm = 0 Then
            Dim u_d As Integer
            u_d = UBound(zb_xm_arr, 2) + 1
            ReDim Preserve zb_xm_arr(4, u_d)
            zb_xm_arr(0, u_d) = dic_watch_xm(all_xm.list(all_xm.ListIndex))
            zb_xm_arr(1, u_d) = all_xm.list(all_xm.ListIndex)
            zb_xm_arr(2, u_d) = dic_group(Combo1.Text)
            zb_xm_arr(3, u_d) = Mid(dic_group(Combo1.Text), 2, 1)
            zb_xm_arr(4, u_d) = True
        End If
                
        'zb_xm列表真实添加
        zb_xm.additem all_xm.list(all_xm.ListIndex)
    End If
    
    
    If zb_xm.ListCount > 0 Then
        zb_xm.Selected(zb_xm.ListCount - 1) = True
    End If
End Sub

Public Sub custom_xm_cmd_Click()
    custom_xm.show 1
End Sub

Private Sub zb_xm_DblClick()
    del_dbEvent
    '更新显示项目数量
    update_xm_sum
End Sub

Private Sub del_dbEvent()
    Dim i As Integer
    For i = 0 To UBound(zb_xm_arr, 2)
        If zb_xm_arr(1, i) = zb_xm.list(zb_xm.ListIndex) And zb_xm_arr(2, i) = dic_group(Combo1.Text) Then
            zb_xm_arr(4, i) = False
        End If
    Next i
    
    'zb_xm列表真实删除
    zb_xm.removeitem zb_xm.ListIndex
    
    If zb_xm.ListCount > 0 Then
        zb_xm.Selected(zb_xm.ListCount - 1) = True
    End If
End Sub

Private Sub cmd_ok_Click()
    Dim txtsql As String, i, flag_x As Integer
    txtsql = "select * from group_xm"
    Set rs = ExeSQL(txtsql, ydhmc)
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            rs.Delete
            rs.Update
            rs.MoveNext
        Loop
        For i = 0 To UBound(zb_xm_arr, 2)
            If zb_xm_arr(4, i) = True Then
                rs.AddNew
                rs("xm_code") = zb_xm_arr(0, i)
                rs("xm_name") = zb_xm_arr(1, i)
                rs("group_xm_code") = zb_xm_arr(2, i)
                rs("group_sex") = zb_xm_arr(3, i)
                rs.Update
            End If
        Next i
    ElseIf rs.RecordCount = 0 Then
        For i = 0 To UBound(zb_xm_arr, 2)
            If zb_xm_arr(4, i) = True Then
                rs.AddNew
                rs("xm_code") = zb_xm_arr(0, i)
                rs("xm_name") = zb_xm_arr(1, i)
                rs("group_xm_code") = zb_xm_arr(2, i)
                rs("group_sex") = zb_xm_arr(3, i)
                rs.Update
            End If
        Next i
    End If
    rs.Close
    MsgBox "保存成功，自动关闭进入下一项！"
    Unload Me
End Sub

Private Sub Combo1_Click()
    '加载组别项目列表
    zb_xm_load Combo1.Text, zb_xm_arr
End Sub

Sub zb_xm_load(data, vData)
    '清除zb_xm列表
    zb_xm.Clear

    '从数组中加载数据
    If IsNotEmpty(vData) = True Then
        Dim i As Integer
        For i = 0 To UBound(vData, 2)
            If vData(2, i) = dic_group(data) And vData(4, i) = True Then
                zb_xm.additem vData(1, i)
            End If
        Next i
    End If

    '加载选定组的项目数
    '更新显示项目数量
    update_xm_sum
    
End Sub
Function IsNotEmpty(ByVal sArray As Variant) As Boolean '判断数组是否为空
    Dim i As Long
    IsNotEmpty = True
    On Error GoTo lerr:
    i = UBound(sArray)
    Exit Function
lerr:
    IsNotEmpty = False
End Function

Private Sub Command2_Click(index As Integer)
    If index = 0 And all_xm.ListIndex <> -1 And CheckBox(zb_xm, all_xm.list(all_xm.ListIndex)) Then
        add_dbEvent
    ElseIf index = 1 And all_xm.ListIndex <> -1 Then
        del_dbEvent
    ElseIf index = 2 Then
        Dim i, j As Integer
        For j = 0 To zb_xm.ListCount
            For i = 0 To UBound(zb_xm_arr, 2)
                If zb_xm_arr(1, i) = zb_xm.list(j) And zb_xm_arr(2, i) = dic_group(Combo1.Text) Then
                    zb_xm_arr(4, i) = False
                End If
            Next i
        Next j
        
        zb_xm.Clear
    End If
        
    
    '更新显示项目数量
    update_xm_sum
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

Private Sub cp_ps_Click(index As Integer)
    Dim i, j As Integer
    If index = 0 Then
        If zb_xm.ListCount > 0 Then
            ReDim cp_arr(zb_xm.ListCount - 1)
            For i = 0 To zb_xm.ListCount - 1
                cp_arr(i) = zb_xm.list(i)
            Next i
            cp_ps(1).Enabled = True
        Else
            MsgBox "没有数据，不能复制"
        End If
    ElseIf index = 1 Then
        If IsNotEmpty(cp_arr) Then
            If zb_xm.ListCount > 0 Then
                '记录旧列别数组
                ReDim old_arr(zb_xm.ListCount - 1)
                For i = 0 To zb_xm.ListCount - 1
                    old_arr(i) = zb_xm.list(i)
                Next i
                '在总表里软删除旧列表数据
                For j = 0 To UBound(old_arr)
                    For i = 0 To UBound(zb_xm_arr, 2)
                        If zb_xm_arr(1, i) = zb_xm.list(j) And zb_xm_arr(2, i) = dic_group(Combo1.Text) Then
                            zb_xm_arr(4, i) = False
                            zb_xm.removeitem (j)
                        End If
                    Next i
                Next j
            End If
            
            '在总表里新增复制数据
            Dim flag_xm As Integer
            For j = 0 To UBound(cp_arr)
                flag_xm = 0
                For i = 0 To UBound(zb_xm_arr, 2)
                    If zb_xm_arr(1, i) = cp_arr(j) And zb_xm_arr(2, i) = dic_group(Combo1.Text) Then
                        If flag_xm = 0 Then
                            zb_xm_arr(4, i) = True
                            zb_xm.additem cp_arr(j)
                            flag_xm = flag_xm + 1
                        Else
                            zb_xm_arr(4, i) = False
                        End If
                    End If
                Next i
                If flag_xm = 0 Then
                    Dim u_d As Integer
                    u_d = UBound(zb_xm_arr, 2) + 1
                    ReDim Preserve zb_xm_arr(4, u_d)
                    zb_xm_arr(0, u_d) = dic_watch_xm(cp_arr(j))
                    zb_xm_arr(1, u_d) = cp_arr(j)
                    zb_xm_arr(2, u_d) = dic_group(Combo1.Text)
                    zb_xm_arr(3, u_d) = Mid(dic_group(Combo1.Text), 2, 1)
                    zb_xm_arr(4, u_d) = True
                    zb_xm.additem cp_arr(j)
                End If
            Next j
        Else
            MsgBox "没有复制，不能粘贴"
        End If
        
        '更新显示项目数量
        update_xm_sum
    End If
End Sub

Private Sub update_xm_sum()
    '更新显示项目数量
    xm_sum.Text = zb_xm.ListCount
End Sub

Private Sub Form_Load()
    '窗体居中显示
    With Screen
        Me.Left = (.Width - Me.Width) / 2
        Me.Top = (.Height - Me.Height) / 2
    End With

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
    
    
    
    '加载项目总表
    xm_refresh
    
    '加载组别项目到数组
    zb_xm_lb
    
    
    '加载组别默认选项，放到最后
    Combo1.Text = Combo1.list(0)
    cp_ps(1).Enabled = False
End Sub

Sub zb_xm_lb()
    Dim txtsql As String, i As Integer
    txtsql = "select * from group_xm"
    Set rs = ExeSQL(txtsql, ydhmc)
    If rs.RecordCount > 0 Then
        ReDim zb_xm_arr(4, rs.RecordCount - 1)
        For i = 0 To UBound(zb_xm_arr, 2)
            zb_xm_arr(0, i) = rs("xm_code")
            zb_xm_arr(1, i) = rs("xm_name")
            zb_xm_arr(2, i) = rs("group_xm_code")
            zb_xm_arr(3, i) = rs("group_sex")
            zb_xm_arr(4, i) = True
            rs.MoveNext
        Next i
        rs.Close
'    Else
'        ReDim zb_xm_arr(4, 0)
    End If
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
