VERSION 5.00
Begin VB.Form Main 
   Caption         =   "�ﾶ�˶�����������ϵͳ"
   ClientHeight    =   6885
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   10770
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Menu system 
      Caption         =   "ϵͳ����"
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
   Begin VB.Menu before 
      Caption         =   "ǰ�ڹ���"
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub create_Click()
    create_frm.Show 1
End Sub

Private Sub Form_Load()
    Dim DBtool As New DBcls
    Dim rs As ADODB.Recordset
    DBtool.SetConnToFile App.Path & "\bpdata.mdb"
    Set rs = DBtool.ExecQuery("select * from ydh")
    If rs.RecordCount > 0 Then
        Do While Not rs.EOF
            If rs("is_open") = True Then
                ydhmc = rs("ydh_name")
                Me.Caption = "�ﾶ�˶�����������ϵͳ" & "  ��ǰ�˶��᣺" & ydhmc
            End If
            rs.MoveNext
        Loop
    Else
        '�ͷ���Դ
        DBtool.ReleaseRecordset rs
    End If
    MsgBox ydhmc
End Sub

Private Sub out_Click()
    End
End Sub
