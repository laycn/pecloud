VERSION 5.00
Begin VB.Form config_orientation 
   BackColor       =   &H8000000E&
   Caption         =   "竞赛规程设置向导"
   ClientHeight    =   6735
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10800
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   6735
   ScaleWidth      =   10800
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   20
      TabIndex        =   0
      Top             =   -120
      Width           =   10770
      Begin VB.Label Label1 
         BackColor       =   &H80000014&
         Caption         =   "第一步：运动会类型选择"
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
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   8
         Top             =   285
         Visible         =   0   'False
         Width           =   7000
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000014&
         Caption         =   "欢迎使用竞赛规程向导"
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
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   1
         Top             =   280
         Width           =   7000
      End
   End
   Begin VB.Frame mian_con 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5535
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   10770
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   3495
         Left            =   2880
         TabIndex        =   11
         Top             =   960
         Width           =   6615
         Begin VB.OptionButton Option1 
            Caption         =   "其他运动会"
            Height          =   735
            Index           =   1
            Left            =   1440
            TabIndex        =   13
            Top             =   1200
            Width           =   4335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "学校运动会"
            Height          =   285
            Index           =   0
            Left            =   1440
            TabIndex        =   12
            Top             =   840
            Width           =   4215
         End
      End
   End
   Begin VB.Frame mian_con 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5535
      Index           =   0
      Left            =   20
      TabIndex        =   10
      Top             =   560
      Width           =   10770
      Begin VB.Label Label2 
         Caption         =   "软件使用介绍"
         Height          =   4455
         Left            =   2160
         TabIndex        =   9
         Top             =   600
         Width           =   7335
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   20
      TabIndex        =   3
      Top             =   6030
      Width           =   10770
      Begin VB.CommandButton Command4 
         Caption         =   "完成"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9720
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "取消"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8640
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "下一步"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7560
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "上一步"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "config_orientation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command2_Click()
    Label1(0).Visible = False
    Label1(1).Visible = True
    mian_con(0).Visible = False
    mian_con(1).Visible = True
    
End Sub

Private Sub Form_Load()
    Label1(1).Visible = False
    mian_con(1).Visible = False
End Sub
