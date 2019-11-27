VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBcls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'��ȨЭ�飺MIT
'�����ˣ�����
Private m_Res As ADODB.Recordset
Private m_Conn As ADODB.Connection
Private m_Command As ADODB.Command
Private m_ConnString As String
Private m_FilePath As String
Private m_Params As New Collection
 
'�涨��ÿ������ִ����ɺ󣬶��������m_Command�����Ҵ���һ���յĶ���
'�����ʼ�����ͷš�
Private Sub Class_Initialize()
 
End Sub
 
Private Sub Class_Terminate()
    Set m_Res = Nothing
    Set m_Conn = Nothing
    Set m_Command = Nothing
End Sub
 
'�������ԡ�
'���ݿ������ַ���
Public Property Get ConnectionString() As String
    ConnectionString = m_ConnString
End Property
 
Public Property Let ConnectionString(ByVal vNewValue As String)
    m_ConnString = vNewValue
End Property
 
'���෽����
Public Function ExecQuery(ByVal SqlStr As String) As Recordset
 
    Dim tempRes As New Recordset
 
    Set m_Command = New ADODB.Command
    Call openConn
    m_Command.ActiveConnection = m_Conn
    m_Command.CommandText = SqlStr
    Set tempRes = m_Command.Execute()
    tempRes.ActiveConnection = Nothing
    Call closeConn
    Set ExecQuery = tempRes
    Set m_Command = Nothing
     
End Function
 
Public Function ExecParamQuery(ByVal SqlStr As String, ParamArray Params())
 
    Dim tempRes As New Recordset
 
    Dim i       As Long
 
    Set m_Command = New ADODB.Command
    '������
    Call openConn
    m_Command.ActiveConnection = m_Conn
    m_Command.CommandText = SqlStr
    m_Command.CommandType = adCmdText
 
    '���ò���
    With m_Command
 
        For Each param In Params
 
            Dim Para As ADODB.Parameter
 
            Set Para = .CreateParameter(CStr(i), GetVarType(param), adParamInput, LenB(param))
            Para.Value = param
            .Parameters.Append Para
        Next
 
    End With
 
    '��ȡִ�к��¼��
    Set tempRes = m_Command.Execute()
    '�����ݿ������ѹ�
    tempRes.ActiveConnection = Nothing
    '�ر����ݿ�����
    Call closeConn
    '�������ݼ���������
    Set ExecParamQuery = tempRes
    '����������
    Set m_Command = Nothing
     
End Function
 
Public Function ExecNonQuery(ByVal SqlStr As String) As Long
 
    '����Ӱ����������
    Dim affectedRows As Long
 
    '������Command����
    Set m_Command = New ADODB.Command
    '������
    Call openConn
    '��Command�����ݿ�����
    m_Command.ActiveConnection = m_Conn
    '����SQL���
    m_Command.CommandText = SqlStr
    '����SQL����
    m_Command.CommandType = adCmdText
    '��ȡִ�к�Ӱ������
    m_Command.Execute affectedRows
    '�ر����ݿ�����
    Call closeConn
    '����������
    Set m_Command = Nothing
    '����Ӱ������
    ExecNonQuery = affectedRows
End Function
 
Public Function ExecParamNonQuery(ByVal SqlStr As String, ParamArray Params()) As Long
 
    Dim i            As Long
 
    Dim affectedRows As Long
 
    Set m_Command = New ADODB.Command
    '������
    Call openConn
    m_Command.ActiveConnection = m_Conn
    m_Command.CommandText = SqlStr
    m_Command.CommandType = adCmdText
 
    '���ò���
    With m_Command
 
        For Each param In Params
 
            Dim Para As ADODB.Parameter
 
            Set Para = .CreateParameter(CStr(i), GetVarType(param), adParamInput, LenB(param))
            Para.Value = param
            .Parameters.Append Para
        Next
 
    End With
 
    '��ȡִ�к��¼��
    m_Command.Execute affectedRows
    '�ر����ݿ�����
    Call closeConn
    '����������
    Set m_Command = Nothing
    '����Ӱ������
    ExecParamNonQuery = affectedRows
End Function
 
Public Sub SetConnToFile(ByVal FilePath As String)
    m_ConnString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & FilePath & ";"
End Sub
 
Public Sub ReleaseRecordset(ByRef dbRes As ADODB.Recordset)
    Set dbRes = Nothing
End Sub
 
'���ڲ�������
Private Sub openConn()
    Set m_Conn = New ADODB.Connection
    m_Conn.CursorLocation = adUseClient
    m_Conn.Open ConnectionString
End Sub
 
Private Sub closeConn()
    m_Conn.Close
    Set m_Conn = Nothing
End Sub
 
Public Function GetVarType(ByRef Value As Variant) As DataTypeEnum
 
    Dim k As New ADODB.Command
 
    Select Case VarType(Value)
 
        Case VbVarType.vbString
            GetVarType = DataTypeEnum.adVarChar
 
        Case VbVarType.vbInteger
            GetVarType = DataTypeEnum.adSmallInt
 
        Case VbVarType.vbBoolean
            GetVarType = DataTypeEnum.adBoolean
 
        Case VbVarType.vbCurrency
            GetVarType = DataTypeEnum.adCurrency
 
        Case VbVarType.vbDate
            GetVarType = DataTypeEnum.adDate
 
        Case Else
            GetVarType = DataTypeEnum.adVariant
    End Select
 
End Function
