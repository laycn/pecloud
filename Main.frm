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
   Begin VB.Menu wj 
      Caption         =   "�ļ�"
      Begin VB.Menu create 
         Caption         =   "�½�"
      End
      Begin VB.Menu out 
         Caption         =   "�˳�"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub create_Click()
    create_frm.Show 1
End Sub

Private Sub Form_Load()
    OpenConn
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CloseConn
End Sub

Private Sub out_Click()
    End
End Sub
