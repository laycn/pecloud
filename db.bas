Attribute VB_Name = "db"
Public cn As New ADODB.Connection
'������¼������
Public rs As New ADODB.Recordset
Public rs1 As New ADODB.Recordset
Public rs2 As New ADODB.Recordset
Public txtsql As String
Public txtsql1 As String
Public txtsql2 As String
Public Sub OpenConn()
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    cn.CursorLocation = adUseClient
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\bpdata.mdb;Persist Security Info=False;"
End Sub

'�ر����ݿ�
Public Sub CloseConn()
    If rs.State = True Then
        rs.Close
        Set rs = Nothing
    End If
    cn.Close
    Set cn = Nothing
End Sub
