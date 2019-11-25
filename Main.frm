VERSION 5.00
Begin VB.Form Main 
   Caption         =   "田径运动会编排与管理系统"
   ClientHeight    =   6885
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   10770
   StartUpPosition =   3  '窗口缺省
   Begin VB.Menu wj 
      Caption         =   "文件"
      Begin VB.Menu create 
         Caption         =   "新建"
      End
      Begin VB.Menu out 
         Caption         =   "退出"
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
