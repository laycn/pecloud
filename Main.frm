VERSION 5.00
Begin VB.Form Main 
   Caption         =   "�ﾶ�˶�����������ϵͳ"
   ClientHeight    =   6825
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   11055
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton setup 
      Caption         =   "������"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton ydh_system 
      Caption         =   "�˶������"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Menu system 
      Caption         =   "�����˶���"
      Index           =   0
      Begin VB.Menu create 
         Caption         =   "�½����"
      End
      Begin VB.Menu hr 
         Caption         =   "-"
      End
      Begin VB.Menu out 
         Caption         =   "�˳�"
      End
   End
   Begin VB.Menu config 
      Caption         =   "�����������"
      Index           =   0
      Begin VB.Menu group 
         Caption         =   "������������"
      End
      Begin VB.Menu unit 
         Caption         =   "������λ����"
      End
      Begin VB.Menu Watchxm 
         Caption         =   "����������Ŀ����"
      End
   End
   Begin VB.Menu bianpai 
      Caption         =   "��������"
      Index           =   0
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function AnimateWindow Lib "user32" (ByVal hwnd As Long, ByVal dwTime As Long, ByVal dwFlags As Long) As Long

Private Sub bianpai_Click(index As Integer)
    'Form2.show 1
End Sub

Private Sub create_Click()
    create_frm.show 1
End Sub

Private Sub Form_Load()
    With Screen
        Me.Left = (.Width - Me.Width) / 2
        Me.Top = (.Height - Me.Height) / 2
    End With
    AnimateWindow hwnd, 300, &H10&
    Me.Refresh
'    Dim DBtool As New DBcls
'    Dim Rs As ADODB.Recordset
'    DBtool.SetConnToFile App.Path & "\bpdata.mdb"
'    Set Rs = DBtool.ExecQuery("select * from ydh")
'    If Rs.RecordCount > 0 Then
'        Do While Not Rs.EOF
'            If Rs("is_open") = True Then
'                ydhmc = Rs("ydh_name")
'                Me.Caption = "�ﾶ�˶�����������ϵͳ" & "  ��ǰ�˶��᣺" & ydhmc
'            End If
'            Rs.MoveNext
'        Loop
'    Else
'        '�ͷ���Դ
'        DBtool.ReleaseRecordset Rs
'    End If
    Dim txtsql As String
    txtsql = "select * from ydh"
    Set res = ExeSQL(txtsql)
    If res.RecordCount > 0 Then
        Do While Not res.EOF
            If res("is_open") = True Then ydhmc = res("ydh_name")
            res.MoveNext
        Loop
    End If
    Me.Caption = "�ﾶ�˶�����������ϵͳ" & "  ��ǰ�˶��᣺" & ydhmc

End Sub

Private Sub group_Click()
    group_con.show 1
End Sub

Private Sub out_Click()
    End
End Sub

Private Sub setup_Click()
    config_orientation.show 1
End Sub

Private Sub unit_Click()
    sign_unit.show 1
End Sub

Private Sub Watchxm_Click()
    xm_con.show 1
End Sub

Private Sub ydh_system_Click()
    create_frm.show 1
End Sub
