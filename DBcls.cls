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
'授权协议：MIT
'创建人：孙瑞
Private m_Res As ADODB.Recordset
Private m_Conn As ADODB.Connection
Private m_Command As ADODB.Command
Private m_ConnString As String
Private m_FilePath As String
Private m_Params As New Collection
 
'规定：每个函数执行完成后，都必须清空m_Command，并且创建一个空的对象
'【类初始化、释放】
Private Sub Class_Initialize()
 
End Sub
 
Private Sub Class_Terminate()
    Set m_Res = Nothing
    Set m_Conn = Nothing
    Set m_Command = Nothing
End Sub
 
'【类属性】
'数据库连接字符串
Public Property Get ConnectionString() As String
    ConnectionString = m_ConnString
End Property
 
Public Property Let ConnectionString(ByVal vNewValue As String)
    m_ConnString = vNewValue
End Property
 
'【类方法】
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
    '打开连接
    Call openConn
    m_Command.ActiveConnection = m_Conn
    m_Command.CommandText = SqlStr
    m_Command.CommandType = adCmdText
 
    '设置参数
    With m_Command
 
        For Each param In Params
 
            Dim Para As ADODB.Parameter
 
            Set Para = .CreateParameter(CStr(i), GetVarType(param), adParamInput, LenB(param))
            Para.Value = param
            .Parameters.Append Para
        Next
 
    End With
 
    '获取执行后记录集
    Set tempRes = m_Command.Execute()
    '与数据库连接脱钩
    tempRes.ActiveConnection = Nothing
    '关闭数据库连接
    Call closeConn
    '返回数据集对象引用
    Set ExecParamQuery = tempRes
    '清空命令对象
    Set m_Command = Nothing
     
End Function
 
Public Function ExecNonQuery(ByVal SqlStr As String) As Long
 
    '定义影响行数变量
    Dim affectedRows As Long
 
    '创建绑定Command对象
    Set m_Command = New ADODB.Command
    '打开连接
    Call openConn
    '绑定Command到数据库连接
    m_Command.ActiveConnection = m_Conn
    '设置SQL语句
    m_Command.CommandText = SqlStr
    '设置SQL类型
    m_Command.CommandType = adCmdText
    '获取执行后影响行数
    m_Command.Execute affectedRows
    '关闭数据库连接
    Call closeConn
    '清空命令对象
    Set m_Command = Nothing
    '返回影响行数
    ExecNonQuery = affectedRows
End Function
 
Public Function ExecParamNonQuery(ByVal SqlStr As String, ParamArray Params()) As Long
 
    Dim i            As Long
 
    Dim affectedRows As Long
 
    Set m_Command = New ADODB.Command
    '打开连接
    Call openConn
    m_Command.ActiveConnection = m_Conn
    m_Command.CommandText = SqlStr
    m_Command.CommandType = adCmdText
 
    '设置参数
    With m_Command
 
        For Each param In Params
 
            Dim Para As ADODB.Parameter
 
            Set Para = .CreateParameter(CStr(i), GetVarType(param), adParamInput, LenB(param))
            Para.Value = param
            .Parameters.Append Para
        Next
 
    End With
 
    '获取执行后记录集
    m_Command.Execute affectedRows
    '关闭数据库连接
    Call closeConn
    '清空命令对象
    Set m_Command = Nothing
    '返回影响行数
    ExecParamNonQuery = affectedRows
End Function
 
Public Sub SetConnToFile(ByVal FilePath As String)
    m_ConnString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & FilePath & ";"
End Sub
 
Public Sub ReleaseRecordset(ByRef dbRes As ADODB.Recordset)
    Set dbRes = Nothing
End Sub
 
'【内部方法】
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

