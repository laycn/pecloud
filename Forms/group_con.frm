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
      TabIndex        =   9
      Top             =   4080
      Width           =   3220
      Begin VB.CommandButton cancel 
         Caption         =   "ȡ��"
         Height          =   375
         Left            =   1920
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2895
      Left            =   5160
      TabIndex        =   8
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
         TabIndex        =   12
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   5160
      TabIndex        =   3
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox add_group 
         Height          =   375
         Left            =   80
         TabIndex        =   4
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
Dim Rs As ADODB.Recordset

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

Private Sub cancel_Click()
    Unload Me
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
End Sub

Private Sub Form_Load()
    Set x = Nothing
    Set x.list = group_list
    Set x.textbox = txtQty
    
    x.addcolumn "", "id", 0, False, False
    x.addcolumn "������", "jgroup", 1200, False, True
    x.addcolumn "�Ա�", "xb", 800, False, False
    x.addcolumn "����������", "jsmc", 1800, False, True
    
    Set Rs = ExeSQL("select * from sign_group", ydhmc)
    
    If Rs.RecordCount > 0 Then
        Do While Not Rs.EOF
            x.additem "", Rs("group_name"), Rs("group_sex_str"), Rs("group_type")
            Rs.MoveNext
        Loop
        Rs.Close
    End If
End Sub

Private Sub group_list_Click()
    'MsgBox group_list.SelectedItem.index
End Sub

Private Sub group_list_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'MsgBox group_list.SelectedItem.index
    'MsgBox Item
End Sub

Private Sub savecmd_Click()
    If group_list.ListItems.Count = 0 Then
        MsgBox "û�����ݲ��ܱ���"
        Exit Sub
    End If
    Dim d As Object '�����ֵ����
    Set d = CreateObject("Scripting.Dictionary")
    Dim i As Integer, k As Integer
    ReDim temp_arr(4, group_list.ListItems.Count - 1)
    k = 0   '���ó�ʼֵ
    For i = 1 To group_list.ListItems.Count
        If Not d.Exists(group_list.ListItems(i).SubItems(1)) Then
            k = k + 1
            d(group_list.ListItems(i).SubItems(1)) = k
        End If
        '�ؼ�����װ�䵽����
        temp_arr(0, i - 1) = k
        If group_list.ListItems(i).SubItems(2) = "����" Then
            temp_arr(1, i - 1) = 1
        Else
            temp_arr(1, i - 1) = 2
        End If
        temp_arr(2, i - 1) = group_list.ListItems(i).SubItems(1)
        temp_arr(3, i - 1) = group_list.ListItems(i).SubItems(2)
        temp_arr(4, i - 1) = group_list.ListItems(i).SubItems(3)
    Next i
    
    '�������ݴ������ݿ�
    If Rs.RecordCount <> 0 Then
        Do While Not Rs.EOF
            Rs.Delete
            Rs.MoveNext
        Loop
    End If
    For i = 1 To UBound(temp_arr, 2) + 1
        Rs.AddNew
        Rs("group_code") = temp_arr(0, i - 1)
        Rs("group_sex") = temp_arr(1, i - 1)
        Rs("group_name") = temp_arr(2, i - 1)
        Rs("group_sex_str") = temp_arr(3, i - 1)
        Rs("group_type") = temp_arr(4, i - 1)
    Next i
    Rs.Update
    Rs.Close
    MsgBox "����ɹ�"
    Unload Me
End Sub
