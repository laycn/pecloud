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

Private Sub out_Click()
    End
End Sub
