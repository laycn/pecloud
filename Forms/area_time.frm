VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form area_time 
   Caption         =   "Form1"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   9240
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame2 
      Caption         =   "比赛时间设置"
      Height          =   5055
      Left            =   3600
      TabIndex        =   13
      Top             =   120
      Width           =   4215
   End
   Begin VB.Frame Frame1 
      Caption         =   "场地设置"
      Height          =   2655
      Left            =   80
      TabIndex        =   0
      Top             =   80
      Width           =   3375
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   300
         Index           =   0
         Left            =   2800
         TabIndex        =   2
         Top             =   400
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   365
         Index           =   0
         Left            =   2465
         TabIndex        =   1
         Text            =   "8"
         Top             =   370
         Width           =   615
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   300
         Index           =   1
         Left            =   2800
         TabIndex        =   6
         Top             =   820
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   365
         Index           =   1
         Left            =   2465
         TabIndex        =   4
         Text            =   "8"
         Top             =   800
         Width           =   615
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   300
         Index           =   2
         Left            =   2805
         TabIndex        =   7
         Top             =   1470
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   365
         Index           =   2
         Left            =   2465
         TabIndex        =   8
         Text            =   "8"
         Top             =   1440
         Width           =   615
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   300
         Index           =   3
         Left            =   2805
         TabIndex        =   9
         Top             =   2125
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   365
         Index           =   3
         Left            =   2465
         TabIndex        =   10
         Text            =   "12"
         Top             =   2100
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "中长跑每组最多人数："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   12
         Top             =   2170
         Width           =   2130
      End
      Begin VB.Label Label1 
         Caption         =   "跨栏跑道数："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   11
         Top             =   1500
         Width           =   1290
      End
      Begin VB.Label Label1 
         Caption         =   "弯道道数："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1260
         TabIndex        =   5
         Top             =   850
         Width           =   1050
      End
      Begin VB.Label Label1 
         Caption         =   "直道道数："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1260
         TabIndex        =   3
         Top             =   440
         Width           =   1050
      End
   End
End
Attribute VB_Name = "area_time"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    '窗体居中显示
    With Screen
        Me.Left = (.Width - Me.Width) / 2
        Me.Top = (.Height - Me.Height) / 2
    End With
    
End Sub
