VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form sign_unit 
   Caption         =   "������λ����"
   ClientHeight    =   6795
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   10635
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame5 
      Height          =   855
      Left            =   7320
      TabIndex        =   20
      Top             =   4620
      Width           =   3135
      Begin VB.CommandButton del_cmd 
         Caption         =   "ɾ��ȫ��"
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   22
         Top             =   280
         Width           =   1455
      End
      Begin VB.CommandButton del_cmd 
         Caption         =   "ɾ��"
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
         Caption         =   "ȷ��"
         Height          =   375
         Index           =   1
         Left            =   2100
         TabIndex        =   19
         Top             =   300
         Width           =   855
      End
      Begin VB.CommandButton save_ok 
         Caption         =   "����"
         Height          =   375
         Index           =   0
         Left            =   200
         TabIndex        =   18
         Top             =   300
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "����һ����λ"
      Height          =   2295
      Left            =   7320
      TabIndex        =   3
      Top             =   2160
      Width           =   3135
      Begin VB.CommandButton add_all 
         Caption         =   "��������"
         Height          =   375
         Left            =   1800
         TabIndex        =   16
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "�����ɺ��������߱�����޸�����"
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   120
         TabIndex        =   23
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "��������"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "���ѡ��"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "����һ����λ"
      Height          =   1935
      Left            =   7320
      TabIndex        =   2
      Top             =   120
      Width           =   3135
      Begin VB.CommandButton add_single 
         Caption         =   "���"
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
         Caption         =   "��λȫ��"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "��λ���"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "���ѡ��"
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
Dim Rs As ADODB.Recordset
Dim d As Object

Private x As New clslist
Private Px As Single, Py As Single

Private Sub add_single_Click()
    '��֯����
    Dim jc, qc, zb, unit_code As String
    jc = Text1(0).Text
    qc = Text1(1).Text
    zb = Trim(Combo1(0).Text)
    Dim rs1 As ADODB.Recordset
    Dim strsql As String
    strsql = "select id,unit_code,unit_group from sign_unit where unit_group ='" & zb & "' order by id DESC"
    Set rs1 = ExeSQL(strsql, ydhmc)
    If rs1.RecordCount >= 1 Then
        rs1.MoveFirst
        unit_code = rs1("unit_code") + 1
    Else
        unit_code = d(Trim(Combo1(0).Text)) * 1000 + 1
    End If
    
    Set Rs = ExeSQL("select * from sign_unit", ydhmc)
    Rs.AddNew
    Rs("unit_code") = unit_code
    Rs("short_name") = jc
    Rs("unit_name") = qc
    Rs("unit_group") = zb
    Rs.Update
    MsgBox "��ӳɹ���"
    
    unit_refresh
End Sub

Private Sub Form_Load()
    
    '��ʼ���ؼ�
    Text1(0).Text = ""
    Text1(1).Text = ""
    
    '����listview����
    Set x = Nothing
    Set x.list = unit_list
    Set x.textbox = txtQty
    
    x.addcolumn "���", "id", 0, False, False
    x.addcolumn "��λ���", "unit_code", 1000, False, False
    x.addcolumn "������", "jgroup", 1200, False, False
    x.addcolumn "��λ���", "short_name", 1200, False, True
    x.addcolumn "��λȫ��", "unit_name", 2000, False, True
    
    '���ز�����������
    unit_refresh

    
    '����combo���ؼ�
    Dim rs1 As ADODB.Recordset
    Set d = CreateObject("Scripting.Dictionary")
    Set rs1 = ExeSQL("select id,group_code, group_name from sign_group order by id", ydhmc)
    If rs1.RecordCount > 0 Then
        Do While Not rs1.EOF
            d(rs1("group_name").Value) = rs1("group_code").Value
            rs1.MoveNext
        Loop
        rs1.Close
    End If
    Dim vkey As Variant
    For Each vkey In d
        Combo1(0).additem vkey
        Combo1(1).additem vkey
    Next
    Combo1(0).Text = Combo1(0).list(0)
    Combo1(1).Text = Combo1(1).list(0)
End Sub
Sub unit_refresh()
    
    unit_list.ListItems.Clear
    
    Set Rs = ExeSQL("select * from sign_unit", ydhmc)
    If Rs.RecordCount > 0 Then
        Do While Not Rs.EOF
            x.additem Rs("id"), Rs("unit_code"), Rs("unit_group"), Rs("short_name"), Rs("unit_name")
            Rs.MoveNext
        Loop
        Rs.Close
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Rs = Nothing
End Sub

Private Sub unit_list_GotFocus()
    'MsgBox unit_list.SelectedItem.SubItems(4)
    'MsgBox unit_list.ListItems(unit_list.SelectedItem.index).Text
End Sub
