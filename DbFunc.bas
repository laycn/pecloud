Attribute VB_Name = "DbFunc"
Option Explicit
'���ݿ�ѡ��:ADODB.Recordset
Public Function ExeSQL(ByVal Sql As String, Optional ydhmc As String) As ADODB.Recordset
    On Error GoTo ErrHandler:
    Dim Connstr As String
    Dim CN As ADODB.Connection
    Dim Rs As ADODB.Recordset
    Dim strArray() As String
    'Dim DataPath As String
    
    Set CN = New ADODB.Connection
    Set Rs = New ADODB.Recordset
    If ydhmc = "" Then
        Connstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\bpdata.mdb"
    'DataPath = App.Path & "\data\data.mdb"
    'Connstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + DataPath + ";Persist Security Info=False" + ";Jet OLEDB:Database Password=blackye"
    Else
        Connstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & ydhmc & "\sdata.mdb"
    End If
    strArray = Split(Sql)
    CN.Open Connstr
    
    If StrComp(UCase$(strArray(0)), "select", vbTextCompare) = 0 Then
           Rs.Open Trim$(Sql), CN, adOpenKeyset, adLockOptimistic
        Set ExeSQL = Rs
    Else
        CN.Execute Sql
    End If

ExeSQl_Exit:
    Set Rs = Nothing
    Set CN = Nothing
    Exit Function
    
ErrHandler:
    '��ʾ������Ϣ
    MsgBox "�����:" & Err.Number & " ������Ϣ��" & Err.Description, vbExclamation
    Resume ExeSQl_Exit
End Function
