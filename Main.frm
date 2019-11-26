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
   Begin VB.Menu system 
      Caption         =   "系统功能"
      Begin VB.Menu create 
         Caption         =   "新建与打开"
      End
      Begin VB.Menu hr 
         Caption         =   "-"
      End
      Begin VB.Menu out 
         Caption         =   "退出"
      End
   End
   Begin VB.Menu before 
      Caption         =   "前期工作"
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
                Me.Caption = "田径运动会编排与管理系统" & "  当前运动会：" & ydhmc
            End If
            rs.MoveNext
        Loop
    Else
        '释放资源
        DBtool.ReleaseRecordset rs
    End If
    MsgBox ydhmc
End Sub

Private Sub out_Click()
    End
End Sub
