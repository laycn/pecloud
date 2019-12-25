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
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   2040
         Width           =   2415
      End
      Begin VB.ComboBox Combo1 
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
         Index           =   0
         Left            =   1200
         TabIndex        =   8
         Text            =   "Combo1"
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

Private Sub Form_Load()
    '窗体居中
    With Screen
        Me.Left = (.Width - Me.Width) / 2
        Me.Top = (.Height - Me.Height) / 2
    End With
    
    '初始化文本框
    Text1.Text = ""
    
    '显示项目列表
    watch_refresh
    
End Sub

Sub watch_refresh()
    Dim rs As ADODB.Recordset
    Set rs = ExeSQL("select xm_name from match_xm order by id", ydhmc)
    
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            watch_xm_list.additem rs("xm_name")
            rs.MoveNext
        Loop
        rs.Close
    End If
End Sub
