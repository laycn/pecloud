Attribute VB_Name = "Public"
Option Explicit
'�����������
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

'����б����Ƿ����ظ�ֵ
Function CheckBox(vData As Object, str As String) As Boolean
    Dim i As Integer
    For i = 0 To vData.ListCount - 1
        If str = vData.list(i) Then
            CheckBox = False '���ظ������
            Exit Function
        End If
    Next i
    CheckBox = True '���ظ��������
End Function

'�ж������Ƿ�Ϊ��
Function IsNotEmpty(ByVal sArray As Variant) As Boolean
    Dim i As Long
    IsNotEmpty = True
    On Error GoTo lerr:
    i = UBound(sArray)
    Exit Function
lerr:
    IsNotEmpty = False
End Function
