# vb 学习笔记

## 控件用法

### listview

常用操作：

```vb
'获取当前行数和列数:
MsgBox "行数：" & ListView1.ListItems.Count & "列数：" & ListView1.ColumnHeaders.Count

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ListView1.ToolTipText = "行数：" & ListView1.ListItems.Count & "列数：" & ListView1.ColumnHeaders.Count
End Sub

'设置宽度：
ListView1.ColumnHeaders.Add , , "备注", 1500

'当前选中行:
MsgBox ListView1.SelectedItem.Index

'获取复选框的值:
MsgBox ListView1.ListItems(1).Checked

'删除指定行:
ListView1.ListItems.Remove 1

'去掉鼠标左键点击标签编辑:labeledit属性改为1

'读取某行第一列内容:
ListView1.ListItems(i).Text

'读取某行第几列内容 (不包括第一列):
ListView1.ListItems(Num).SubItems (4)

'读取当前选中行第一列内容:
ListView1.ListItems(ListView1.SelectedItem.Index).Text


'循环查找读取项目:
Dim i As Integer
For i = 1 To ListView1.ListItems.Count
	If ListView1.ListItems(i).Text = 4 Then MsgBox listView1.ListItems(i).Text '第一列
    If ListView1.ListItems(i).SubItems(1) = 4 Then MsgBox ListView1.ListItems(i).SubItems(1) '第二列
Next i

'清空列表头:
ListView1.ColumnHeaders.Clear

'清空列表:
ListView1.ListItems.Clear
    
'右键菜单:
Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then PopupMenu CommandLst '按下鼠标右键 显示菜单
End Sub

'当前选中判断:
Dim N
If ListView1.ListItems.Count <> 0 Then N = ListView1.SelectedItem.Index Else MsgBox "当前没有任何主机在线!", vbInformation, "警告:": Exit Sub
If N < 1 Then MsgBox "你没有选中任何主机!", vbInformation, "警告:": Exit Sub

'VB设置某行为选中/非选中状态:
ListView.ListItems(i).Selected = True '选中第i行
ListView.ListItems(i).Selected = False '选中第i行

'设置ListView  item项颜色
ListView1.ListItems(i).ForeColor = vbRed
ListView1.ListItems(i).ListSubItems.Item(1).ForeColor = vbRed
ListView1.ListItems(i).ListSubItems.Item(2).ForeColor = vbRed
```

```vb
Private Sub Form_Load()
    ListView1.ListItems.Clear               '清空列表
    ListView1.ColumnHeaders.Clear           '清空列表头
    ListView1.View = lvwReport              '设置列表显示方式
    ListView1.GridLines = True              '显示网络线
    ListView1.LabelEdit = lvwManual         '禁止标签编辑
    ListView1.FullRowSelect = True          '选择整行
  
    ListView1.ColumnHeaders.Add , , "ID", 500 '给列表中添加列名
    ListView1.ColumnHeaders.Add , , "本地 IP", 1500
    ListView1.ColumnHeaders.Add , , "本地端口", 1200
    ListView1.ColumnHeaders.Add , , "协议", 550
    ListView1.ColumnHeaders.Add , , "远程 IP", 1500
    ListView1.ColumnHeaders.Add , , "远程端口", 900
    ListView1.ColumnHeaders.Add , , "当前状态", 900
    ListView1.ColumnHeaders.Add , , "连接时间", 900
    '-------------------------------------------------------
    Dim x
    x = ListView1.ListItems.Count + 1
    ListView1.ListItems.Add , , x
    ListView1.ListItems(x).SubItems(1) = "00:00:00"
    ListView1.ListItems(x).SubItems(2) = "2008-01-01"
    ListView1.ListItems(x).SubItems(3) = "(无)"
    '-------------------------------------------------------
    ListView1.ListItems.Clear               '清空列表
    ListView1.ListItems.Add , , "1"
    'ListView1.ListItems.Add , , "1", , 1   '添加图标 后面那个1是ImageList1控件中的图标索引号
    ListView1.ListItems(1).SubItems(1) = "00:00:00"
    ListView1.ListItems(1).SubItems(2) = "2008-01-01"
    ListView1.ListItems(1).SubItems(3) = "(无)"
  
    ListView1.ListItems.Add , , "2"
    ListView1.ListItems(2).SubItems(1) = "00:00:01"
    ListView1.ListItems(2).SubItems(2) = "2008-01-01"
    ListView1.ListItems(2).SubItems(3) = "(无)"
    '-------------------------------------------------------
    '下列的属性也可以 单击控件右键->属性 进行设置。
    ListView1.View = lvwReport              '设置显示方式为列表
    ListView1.AllowColumnReorder = True     '对行进行程序排列，用鼠标进行排列
    ListView1.Arrange = lvwAutoLeft         '图标横排列
    ListView1.Arrange = lvwAutoTop          '图标竖排列
    ListView1.FlatScrollBar = False         '显示滚动条
    ListView1.FlatScrollBar = True          '隐藏滚动条
    ListView1.FullRowSelect = True          '选择整行
    ListView1.LabelEdit = lvwManual         '禁止标签编辑
    ListView1.GridLines = True              '显示网络线
    ListView1.LabelWrap = True              '图标可以换行
    ListView1.MultiSelect = True            '可以选择多个项目
    ListView1.PictureAlignment = lvwTopLeft '图片对齐方式是左顶部，其他有右顶部(1)、左底部(2)、右底部(3)、居中(4)、平铺(5)
    ListView1.Checkboxes = True             '显示复选框
    'ListView1.DropHighlight = ListView1.ListItems.Item(2)   '显示系统颜色
End Sub
```

更新多条记录到数据库

```vb
rs.Open SQL, conn, adOpenStatic, adLockOptimistic
For i = 1 To ListView1.ListItems.Count
    rs.AddNew
    rs.Fields("数据库的列名") = ListView1.ListItems(i).SubItems(1)
    '注释:ListItems(i).SubItems(1) 当i=1时,取第一行第2列的值.
    rs.Fields("数据库的列名") = ListView1.ListItems(i).SubItems(2)
    rs.Fields("数据库的列名") = ListView1.ListItems(i).SubItems(3)
    '**********(要获取多少数据这中间自己加)******
    rs.Fields("数据库的列名") = ListView1.ListItems(i).SubItems(9)
    rs.Fields("数据库的列名") = ListView1.ListItems(i).SubItems(10)
    rs.Update
Next i
```



## vb 与数据库

### ADOX 新建 access 数据库

> 准备工作：需要在 工程 -> 引用中选择对象库“Microsoft ADOExt 2.1. For DDL Security”，简称为 ADOX 。其库文件名是：Msadox.dll。路径地址：“C:\Program Files (x86)\Common Files\System\ado”。
>
> ADOX 常用方法有：Append（包括 Columns、Groups、Indexex、Keys、Procedures、Tables、Users、Views、Create(创建新的目录)、delete(删除集合中的对象)、Refresh(更新集合中的对象)等等。

1. 创建新的数据库

~~~vb
Dim cat As New ADOX.Catalog
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim pstr As String
'Set cat = New ADOX.Catalog
pstr = "Provider=Microsoft.Jet.OLEDB.4.0;"
pstr = pstr & "Data Source=" & App.Path & "\" & "edit.mdb;"
'创建数据库
cat.Create pstr

~~~
2. 创建表

~~~vb
Dim tbl As New Table
cat.ActiveConnection = pstr
tbl.Name = "MyTable" '表的名称
tbl.Columns.Append "编号", adInteger '表的第一个字段
tbl.Columns.Append "姓名", adVarWChar, 8 '表的第二个字段
tbl.Columns.Append "住址", adVarWChar, 50 '表的第三个字段
cat.Tables.Append tbl '建立数据表

conn.Open pstr
rs.CursorLocation = adUseClient
rs.Open "MyTable", conn, adOpenKeyset, adLockPessimistic
rs.AddNew '往表中添加新记录
rs.Fields(0).Value = 9801
rs.Fields(1).Value = "孙悟空"
rs.Fields(2).Value = "广州市花果山"
rs.Update
conn.Close
~~~

   

### ADO 操作数据库

> DataGrid 部件使用，部件中加载 Microsoft DataGrid control 6.0 (sp6) 控件。

~~~vb
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim command As New ADODB.command

'数据库的连接放到 load 中加载
Private Sub form_load()
    Dim str As String
    str = App.Path
    If Right(str, 1) <> "\" Then
        str = str + "\"
    End If
    pstr = "Provider=Microsoft.Jet.OLEDB.4.0;"
    pstr = pstr & "Persist Security Info=False;"
    pstr = pstr & "Data Source=" & App.Path & "\" & "gongzi.mdb"
    conn.Open pstr
    rs.CursorLocation = adUseClient
    rs.Open "工资表", conn, adOpenKeyset, adLockPessimistic
    Set DataGrid1.DataSource = rs '查看命令
End Sub
   
'查询并创建新的表
Private Sub Command1_Click()
    Dim bm As String
    Dim sql As String
    If Text1.Text <> "" Then
        bm = Trim(Text1.Text)
        sql = "select 编号,姓名,实发工资 Into " & bm & " From 工资表 where 实发工资 > 2000"
        Set command.ActiveConnection = conn
        command.CommandText = sql
        command.Execute
    Else
        MsgBox "你必须输入一个名字"
    End If
End Sub
               
'修改好数据可通过这条命令更新        
Private Sub Command2_Click()
    rs.UpdateBatch
End Sub
~~~

使用 ADO 编程模型需添加 ADO 对象库的“引用”——"Microsoft ActiveX Data Objects 2.x Library "

一、连接数据库

1. 连接到 SQL 数据库：

    ~~~vb
    dim cnn as new ADODB.Connection '创建 Connection 对象 cnn,关键字 new 用于创建新对象
    cnn.ConnectionString = "Provider=SQLOLEDB.1;Password=密码;User ID=用户名;Initial Catalog=SQL数据库文件;Data Source=localhost;" '指定提供者，设置数据源
    cnn.open  '打开到数据库的连接
    '……一系列操作
    cnn.Close

    '或者
    dim cnn as new ADODB.Connection '创建 Connection 对象 cnn,关键字 new 用于创建新对象
    cnn.Open "Provider=SQLOLEDB.1;Password=密码;User ID=用户名;Initial Catalog=SQL数据库文件;Data Source=localhost;"    '打开到数据库的连接
    '……..
    cnn.Close
    ~~~

2. 连接到ACCESS数据库

    ~~~VB
    Dim cnn as New ADODB.Connection '创建Connection对象cnn,关键字new用于创建新对象
    cnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ACCESS数据库文件.mdb"   '指定提供者，设置数据源
    cnn.Open      '打开到数据库的连接
    '……
    cnn.Close

    '或者第二种
    Dim cnn as New ADODB.Connection '创建Connection对象cnn,关键字new用于创建新对象
    cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ACCESS数据库文件.mdb"   '打开到数据库的连接
    '……
    cnn.Close
    ~~~

二、读数据库操作

  读数据库操作一般可通过 recordset 对象实现。

~~~vb
dim cnn as new ADODB.Connection 
cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=ACCESS数据库文件.mdb"

dim rs as new Recordset    '声明一个记录集对象
rs.open [记录源, cnn, 游标类型, 锁定类型, 命令类型]      '也可先设置记录集相关属性

text1.text=rs("字段名称1或索引")  '假设读取出来的字段信息显示在文本框中，三种读取方法等价
text2.text=rs.fields("字段名称2或索引")
text3.text=rs!字段名称3

....

rs.Close
cnn.Close
set rs=nothing  '释放rs对象实例
set cnn=nothing  '释放connection对象实例
~~~

记录源一般为一条SQL查询语句，以实现查询目的

recordset对象还保持查询返回的记录的位置，它返回第一个检索到的记录，并允许你一次一项逐步扫描其他结果。

recordset对象的部分属性及方法如下

> Rs(i) : 读取第 i 个字段的数据，i从0开始
>
> Rs(字段名)：读取指定字段的数据
>
> Rs.EOF：记录指针指到记录的尾部
>
> Rs.BOF：记录指针指到记录的头部
>
> Rs.RecordCount：游标中的数据记录总数
>
> Rs.PageSize：当对象设有分页时，用于指定逻辑页中的记录个数
>
> Rs.pageCount：返回记录集中的逻辑分页数
>
> Rs.MoveNext：将记录指针移动到下一个记录
>
> Rs.MovePrev：将记录指针移动到上一个记录
>
> Rs.MoveFirst：将记录指针移动到第一个记录
>
> Rs.MoveLast：将记录指针移动到最后一个记录
>
> Rs.delete：将当前记录删除
>
> Rs.addNew：添加一个新记录（行）

三、写数据库操作

方法1：

> dim rs as new recordset
>
> rs.Open [记录源, cnn, 游标类型, 锁定类型, 命令类型]
>
> rs.addnew   '告诉rs我们要添加一行
>
> rs("字段名称1或索引") =  值1  '给要添加行的一个字段赋值，三种方法等价
>
> rs.fields("字段名称2或索引") =  值2 
>
> rs!字段名称3=  值3
>
> ……
>
> rs.update
>
> rs.close
>
> cnn.close
>
> set rs = nothing
>
> set cnn = nothing

addnew使用方法后，如果要放弃添加的结果，应调用记录集的CancelUpdate方法放弃

方法2：

> strSQL = "Insert Into 数据表(字段1,字段2……) Values(值1, 值2……)"  '拼写Insert插入语句
>
> cnn.Execute strSQL   '执行Insert语句实现添加

四、修改数据库操作

与插入数据相似，方法2运用SQL语句更新：

> strSQL= "Update 数据表 Set 字段1 = 新值1, 字段2 = 新值2……"

五、删除数据操作

> rs.Delete
>
> .....
>
> rs.Update
>
> 也可用SQL语句来实现
>
> strSQL = "delete from 学生情况表 where 学号=‘07001’"

分页介绍----使用记录集

用到几个记录集属性

> rs.pagesize：定义一页显示记录的条数
>
> rs.recordcount：统计数据库记录总数
>
> rs.pagecount：统计总页数
>
> rs.absolutepage：将数据库指针移动到当前要显示的数据记录的第一条记录。如果分5页显示，将absolutepage属性设为2，则当前记录指针移至第2页第一条记录。