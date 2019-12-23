Attribute VB_Name = "Public"
Option Explicit
'公共变量存放
Public ydhmc As String
Public res As ADODB.Recordset

Public d As Object


Sub group_info()
    Set d = CreateObject("Scripting.Dictionary")
    Dim rs1 As ADODB.Recordset
    Set rs1 = ExeSQL("select id,group_code, group_name from sign_group order by id", ydhmc)
    If rs1.RecordCount > 0 Then
        Do While Not rs1.EOF
            d(rs1("group_name").Value) = rs1("group_code").Value
            rs1.MoveNext
        Loop
        rs1.Close
    End If
End Sub


