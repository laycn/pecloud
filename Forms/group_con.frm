VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form group_con 
   Caption         =   "���ñ������"
   ClientHeight    =   4905
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8520
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   4905
   ScaleWidth      =   8520
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   5160
      TabIndex        =   10
      Top             =   4080
      Width           =   3220
      Begin VB.CommandButton cancel 
         Caption         =   "ȡ��"
         Height          =   375
         Left            =   1920
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton savecmd 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2895
      Left            =   5160
      TabIndex        =   9
      Top             =   1200
      Width           =   3220
      Begin VB.Label Label1 
         Caption         =   "˵����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   5160
      TabIndex        =   4
      Top             =   0
      Width           =   3220
      Begin VB.CommandButton delall 
         Caption         =   "ɾ��ȫ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1100
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton add 
         Caption         =   "��ӷ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2130
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton del 
         Caption         =   "ɾ������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   80
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox add_group 
         Height          =   375
         Left            =   80
         TabIndex        =   5
         Top             =   240
         Width           =   3010
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   5050
      Begin VB.CommandButton Command1 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Top             =   3840
         Width           =   855
      End
      Begin VB.TextBox txtQty 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1200
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   2880
         Width           =   735
      End
      Begin MSComctlLib.ListView group_list 
         Height          =   4455
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   7858
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "group_con"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private x As New clslist
Private Px As Single, Py As Single

Private Sub add_Click()
    Set x.list = group_list
    Dim txtstr As String
    txtstr = Trim(add_group.Text)
    If txtstr = "" Then
        MsgBox "������ݲ���Ϊ��"
        Exit Sub
    End If
    x.additem "", txtstr, "����", txtstr & "����"
    x.additem "", txtstr, "Ů��", txtstr & "Ů��"
End Sub

Private Sub Command1_Click()
    MsgBox group_list.ListItems(1).SubItems(3)
    'MsgBox group_list.ListItems(group_list.SelectedItem.index).Text
End Sub

Private Sub del_Click()
    If group_list.ListItems.Count = 0 Then
        MsgBox "��¼Ϊ�գ�����ɾ��"
        Exit Sub
    End If
    Set x.list = group_list
    x.removeitem group_list.SelectedItem.index
End Sub

Private Sub delall_Click()
    group_list.ListItems.Clear
    
    'txtQty.Visible = False
End Sub

Private Sub Form_Load()
    Set x.list = group_list
    Set x.textbox = txtQty
    
    x.addcolumn "", "id", 0, False, False
    x.addcolumn "������", "jgroup", 1200, False, True
    x.addcolumn "�Ա�", "xb", 800, False, False
    x.addcolumn "����������", "jsmc", 1800, False, True
    
    
    x.additem "", "Сѧ��", "����", "Сѧ������"
    x.additem "", "Сѧ��", "Ů��", "Сѧ��Ů��"
    x.additem "", "��ѧ��", "����", "��ѧ������"
    x.additem "", "��ѧ��", "Ů��", "��ѧ��Ů��"
    
    x.Resize
    
End Sub

Private Sub group_list_Click()
    'MsgBox group_list.SelectedItem.index
End Sub

Private Sub group_list_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'MsgBox group_list.SelectedItem.index
    'MsgBox Item
End Sub

Private Sub savecmd_Click()
    Dim d As Object '�����ֵ����
    Set d = CreateObject("Scripting.Dictionary")
'    d.add "a", 200
'    d.add "b", 300
'    d.add "c", 400
'    d.add "a", 500
'    d("a") = 200
'    d("a") = 500
'    MsgBox d.Count
    MsgBox group_list.ListItems.Count
    Dim i As Integer
    Dim temp_arr(1)
    For i = 1 To group_list.ListItems.Count
        d(group_list.ListItems(i).SubItems(1)) = group_list.ListItems(i).SubItems(1)
    Next i
    MsgBox d.Count
End Sub
