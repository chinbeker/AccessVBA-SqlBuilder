# Access VBA SQL Builder 使用文档

## 📖 简介

这是一个流畅的 SQL 构建器，用于在 Access VBA 中便捷地构建和执行 SQL 语句。支持参数化查询。

**包含两个部分：**
- `SqlBuilder` - 类模块，用于构建 SQL 语句
- `DbSql` - 标准模块，提供数据库操作方法

---

## 🔧 可用方法速查表

### SqlBuilder 类方法

#### 查询设置
| 方法 | 说明 |
|------|------|
| `SelectAll` | 选择全部字段 |
| `Top value` | 指定返回记录数 |
| `Distinct` | 去除重复值 |
| `Field name, [value], [alias]` | 添加字段 |
| `From table, [alias]` | 设置数据源 |
| `Derived alias` | 设置为派生表 |

#### 表连接
| 方法 | 说明 |
|------|------|
| `Join table, sourceTable, joinField, [sourceField]` | INNER JOIN |
| `LeftJoin table, sourceTable, joinField, [sourceField]` | LEFT JOIN |

#### 条件筛选
| 方法 | 说明 |
|------|------|
| `Where condition, [value]` | 筛选条件 |
| `Group field` | 分组 |
| `Having condition, [value]` | 分组筛选 |
| `Order field, [reverse]` | 排序 |

#### 参数化查询
| 方法 | 说明 |
|------|------|
| `Param name, value, [type]` | 给参数赋值 |
---

<br>

### DbSql 模块方法

#### 连接与基础操作
| 方法 | 说明 | 返回值 |
|------|------|--------|
| `SelectDynaset(sqlString)` | 运行查询SQL，返回动态集 | DAO.Recordset |
| `SelectSnapshot(sqlString)` | 运行查询SQL，返回快照 | DAO.Recordset |
| `Execute(sqlString)` | 运行非查询SQL | Long |
| `TableDef(name)` | 获取表定义 | DAO.TableDef |

#### 打开整张表
| 方法 | 说明 | 返回值 |
|------|------|--------|
| `OpenTable(name)` | 打开本地表 | DAO.Recordset |
| `TableDynaset(name)` | 以动态集打开整张表 | DAO.Recordset |
| `TableSnapshot(name)` | 以快照打开整张表 | DAO.Recordset |

#### 增删改操作
| 方法 | 说明 | 返回值 |
|------|------|--------|
| `Insert(table, sql)` | 插入数据 | Long |
| `Update(table, sql)` | 更新数据 | Long |
| `Delete(table, sql)` | 删除数据 | Long |
| `Clear(table)` | 清空表数据 | Long |

#### 查询记录集
| 方法 | 说明 | 返回值 |
|------|------|--------|
| `Find(sql)` | 返回记录集（快照） | DAO.Recordset |
| `Record(sql)` | 返回记录集（动态集） | DAO.Recordset |
| `First(sql)` | 返回第一条记录（快照） | DAO.Recordset |
| `FirstRecord(sql)` | 返回第一条记录（动态集） | DAO.Recordset |

#### 表格首尾记录
| 方法 | 说明 | 返回值 |
|------|------|--------|
| `TableFirst(table, [orderBy])` | 表格第一条记录（快照） | DAO.Recordset |
| `TableFirstRecord(table, [orderBy])` | 表格第一条记录（动态集） | DAO.Recordset |
| `TableLast(table, [orderBy])` | 表格最后一条记录（快照） | DAO.Recordset |
| `TableLastRecord(table, [orderBy])` | 表格最后一条记录（动态集） | DAO.Recordset |

#### 统计数量
| 方法 | 说明 | 返回值 |
|------|------|--------|
| `Count(sql)` | 统计记录数 | Long |
| `TableCount(table)` | 统计整张表记录数 | Long |

#### 获取单个值
| 方法 | 说明 | 返回值 |
|------|------|--------|
| `GetValue(sql)` | 获取第一个字段的值 | Variant |
| `GetValueFromSql(sqlString)` | 从SQL获取第一个字段的值 | Variant |
| `Lookup(table, field, condition)` | 查询字段值 | Variant |
| `FirstValue(table, field, [condition], [orderBy])` | 第一条记录的字段值 | Variant |
| `LastValue(table, field, [condition], [orderBy])` | 最后一条记录的字段值 | Variant |

#### 更新单个值
| 方法 | 说明 | 返回值 |
|------|------|--------|
| `SetValue(table, field, value, condition)` | 设置字段值 | Boolean |

#### 联合查询
| 方法 | 说明 | 返回值 |
|------|------|--------|
| `Union(sql1, sql2, ...)` | 联合查询（去重） | DAO.Recordset |
| `UnionAll(sql1, sql2, ...)` | 联合查询（含重复） | DAO.Recordset |

#### 辅助方法
| 方法 | 说明 | 返回值 |
|------|------|--------|
| `Parameter(name)` | 引用参数（不被转义） | String |
| `Field(name)` | 引用字段（不被转义） | String |
| `Expression(expr)` | 引用表达式（不被转义） | String |


<br>
<br>

## 🚀 快速开始

### 1. 基本查询

```vba
Public Sub BasicQueryExample()
    Dim sql As New SqlBuilder
    Dim rs As DAO.Recordset

    ' 构建查询：SELECT * FROM Users WHERE Age > 18 ORDER BY Name
    sql.SelectAll
    sql.From "Users"
    sql.Where "Age >", 18
    sql.Order "Name"

    ' 执行查询（DbSql 是模块，直接调用）
    Set rs = DbSql.Find(sql)

    ' 处理结果
    Do While Not rs.EOF
        Debug.Print rs("Name") & " - " & rs("Age")
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
End Sub
```

### 2. 参数化查询

```vba
Public Sub ParameterizedQueryExample()
    Dim sql As New SqlBuilder
    Dim rs As DAO.Recordset

    ' 使用参数，防止 SQL 注入
    sql.SelectAll
    sql.From "Employees"
    sql.Where "Department =", DbSql.Parameter("Dept")
    sql.Where "Salary >", DbSql.Parameter("MinSalary")

    ' 参数赋值
    sql.Param "Dept", "Sales"
    sql.Param "MinSalary", 50000

    Set rs = DbSql.Find(sql)
    ' 处理结果...
End Sub
```

### 3. 指定字段查询

```vba
Public Sub SpecificFieldsExample()
    Dim sql As New SqlBuilder
    Dim rs As DAO.Recordset

    ' SELECT UserID, Name, Email FROM Users WHERE Status = 'Active'
    sql.Field "UserID"
    sql.Field "Name"
    sql.Field "Email"
    sql.From "Users"
    sql.Where "Status =", "Active"

    Set rs = DbSql.Find(sql)
    ' 处理结果...
End Sub
```

### 4. 字段别名

```vba
Public Sub AliasExample()
    Dim sql As New SqlBuilder
    Dim rs As DAO.Recordset

    ' SELECT UserID AS ID, Name AS FullName FROM Users
    sql.Field "UserID", , "ID"
    sql.Field "Name", , "FullName"
    sql.From "Users"

    Set rs = DbSql.Find(sql)
    Debug.Print rs("ID") & ": " & rs("FullName")
End Sub
```

### 5. 插入数据

```vba
Public Sub InsertExample()
    Dim sql As New SqlBuilder
    Dim AffectedRows As Long

    ' INSERT INTO Users (Name, Age, Email) VALUES ('张三', 25, 'zhang@example.com')
    sql.Field "Name", "张三"
    sql.Field "Age", 25
    sql.Field "Email", "zhang@example.com"

    AffectedRows = DbSql.Insert("Users", sql)

    Debug.Print "插入了 " & AffectedRows & " 行"
End Sub
```

### 6. 更新数据

```vba
Public Sub UpdateExample()
    Dim sql As New SqlBuilder
    Dim AffectedRows As Long

    ' UPDATE Users SET Age = 26, Status = 'Active' WHERE Name = '张三'
    sql.Field "Age", 26
    sql.Field "Status", "Active"
    sql.Where "Name =", "张三"

    AffectedRows = DbSql.Update("Users", sql)

    Debug.Print "更新了 " & AffectedRows & " 行"
End Sub
```

### 7. 删除数据

```vba
Public Sub DeleteExample()
    Dim sql As New SqlBuilder
    Dim AffectedRows As Long

    ' DELETE FROM Users WHERE Status = 'Inactive'
    sql.Where "Status =", "Inactive"

    AffectedRows = DbSql.Delete("Users", sql)

    Debug.Print "删除了 " & AffectedRows & " 行"
End Sub
```

### 8. 表连接（INNER JOIN）

```vba
Public Sub JoinExample()
    Dim sql As New SqlBuilder
    Dim rs As DAO.Recordset

    ' SELECT Orders.OrderID, Customers.Name
    ' FROM Orders
    ' INNER JOIN Customers ON Orders.CustomerID = Customers.CustomerID

    sql.Field "Orders.OrderID"
    sql.Field "Customers.Name"
    sql.From "Orders"
    sql.Join "Customers", "Orders", "CustomerID"

    Set rs = DbSql.Find(sql)
End Sub
```

### 9. 左连接（LEFT JOIN）

```vba
Public Sub LeftJoinExample()
    Dim sql As New SqlBuilder
    Dim rs As DAO.Recordset

    ' SELECT Customers.Name, Orders.OrderID
    ' FROM Orders
    ' LEFT JOIN Customers ON Orders.CustomerID = Customers.CustomerID
    ' WHERE Orders.OrderID IS NULL

    sql.Field "Customers.Name"
    sql.Field "Orders.OrderID"
    sql.From "Orders"
    sql.LeftJoin "Customers", "Orders", "CustomerID"
    sql.Where "Orders.OrderID IS NULL"  ' 查找没有订单的客户

    Set rs = DbSql.Find(sql)
End Sub
```

### 10. 分组与聚合

```vba
Public Sub GroupByExample()
    Dim sql As New SqlBuilder
    Dim rs As DAO.Recordset

    ' SELECT Department, COUNT(*) AS EmpCount, AVG(Salary) AS AvgSalary
    ' FROM Employees
    ' GROUP BY Department
    ' HAVING COUNT(*) > 5

    sql.Field "Department"
    sql.Field "COUNT(*)", , "EmpCount"
    sql.Field "AVG(Salary)", , "AvgSalary"
    sql.From "Employees"
    sql.Group "Department"
    sql.Having "COUNT(*) >", 5

    Set rs = DbSql.Find(sql)
End Sub
```

### 11. 子查询（派生表）

```vba
' 示例1：FROM 子查询（将子查询作为数据源）
Public Sub SubQueryExample()
    Dim subSql As New SqlBuilder
    Dim sql As New SqlBuilder
    Dim rs As DAO.Recordset

    ' 子查询：筛选出成年用户
    subSql.Field "UserID"
    subSql.Field "Name"
    subSql.Field "Age"
    subSql.From "Users"
    subSql.Where "Age >= 18"

    ' 主查询：从子查询结果中继续查询
    sql.Field "u.Name"
    sql.Field "u.Age"
    sql.From subSql, "u"         ' 关键：直接使用子查询作为数据源，并同时设置别名
    sql.Where "u.Age < 30"
    sql.Order "u.Age"
    Set rs = DbSql.Find(sql)
End Sub


' 示例2：JOIN 子查询（将子查询作为连接表）
Public Sub SubQueryExample()
    Dim subSql As New SqlBuilder
    Dim sql As New SqlBuilder
    Dim rs As DAO.Recordset

    ' 子查询：获取每个部门的平均工资
    subSql.Field "DepartmentID"
    subSql.Field "AVG(Salary)", , "AvgSalary"
    subSql.From "Employees"
    subSql.Group "DepartmentID"
    subSql.Derived "DeptAvg"         ' 关键：设置别名

    ' 主查询：连接子查询
    sql.Field "Employees.Name"
    sql.Field "Employees.Salary"
    sql.Field "DeptAvg.AvgSalary"
    sql.From "Employees"
    sql.Join subSql, "Employees", "DepartmentID", "DepartmentID"
    sql.Where "Employees.Salary > DeptAvg.AvgSalary"

    Set rs = DbSql.Find(sql)
End Sub
```

### 12. TOP N 查询

```vba
Public Sub TopExample()
    Dim sql As New SqlBuilder
    Dim rs As DAO.Recordset

    ' 获取工资最高的前10名员工
    sql.SelectAll
    sql.Top 10
    sql.From "Employees"

    ' True = DESC (倒序)
    sql.Order "Salary", True

    Set rs = DbSql.Find(sql)
End Sub
```

### 13. DISTINCT 去重查询

```vba
Public Sub DistinctExample()
    Dim sql As New SqlBuilder
    Dim rs As DAO.Recordset

    ' SELECT DISTINCT Department FROM Employees

    sql.Distinct
    sql.Field "Department"
    sql.From "Employees"

    Set rs = DbSql.Find(sql)
End Sub
```

### 14. 统计数量

```vba
Public Sub CountExample()
    Dim sql As New SqlBuilder
    Dim TotalCount As Long

    ' 统计活跃用户数量
    sql.From "Users"
    sql.Where "Status =", "Active"

    TotalCount = DbSql.Count(sql)
    Debug.Print "活跃用户数：" & TotalCount
End Sub
```

### 15. 快捷查询方法

```vba
Public Sub QuickQueryExample()
    Dim Value As Variant

    ' 获取第一个值（第4个参数是排序字段）
    Value = DbSql.FirstValue("Users", "Name", "Age > 18", "Age")
    Debug.Print "最年轻的成年用户：" & Value

    ' 获取最后一个值（第4个参数是排序字段）
    Value = DbSql.LastValue("Users", "Name", , "Age")
    Debug.Print "最年长的用户：" & Value

    ' 查询指定值（没有排序）
    Value = DbSql.Lookup("Users", "Email", "Name = '张三'")
    Debug.Print "张三的邮箱：" & Value

    ' 设置值
    Dim Success As Boolean
    Success = DbSql.SetValue("Users", "Status", "Inactive", "LastLogin < #2024-01-01#")
    If Success Then Debug.Print "更新成功"
End Sub
```

### 16. 联合查询（UNION）

```vba
Public Sub UnionExample()
    Dim sql1 As New SqlBuilder
    Dim sql2 As New SqlBuilder
    Dim sql3 As New SqlBuilder
    Dim rs As DAO.Recordset

    ' 查询所有客户和供应商的名称
    sql1.Field "CompanyName", , "Name"
    sql1.From "Customers"

    sql2.Field "CompanyName", , "Name"
    sql2.From "Suppliers"

    sql3.Field "ContactName", , "Name"
    sql3.From "Employees"

    ' UNION（去重）
    Set rs = DbSql.Union(sql1, sql2, sql3)

    ' UNION ALL（包含重复）
    Set rs = DbSql.UnionAll(sql1, sql2, sql3)
End Sub
```

### 17. 获取单个值

```vba
Public Sub GetSingleValueExample()
    Dim sql As New SqlBuilder
    Dim UserName As Variant

    ' 用于复杂的SQL查询，简单的请用快速查询

    ' 获取用户ID为1001的用户名
    sql.Field "Name"
    sql.From "Users"
    sql.Where "UserID =", 1001

    UserName = DbSql.GetValue(sql)

    If Not IsNull(UserName) Then
        Debug.Print "用户名：" & UserName
    Else
        Debug.Print "用户不存在"
    End If
End Sub
```

### 18. 引用字段和表达式

```vba
Public Sub FieldValueReferenceExample()
    Dim sql As New SqlBuilder
    Dim rs As DAO.Recordset

    ' 使用 DbSql.Field 引用字段（不会被转义）
    ' DbSql.Field() 的作用是告诉 SQL 构建器：这是一个字段名，不要加引号转义

    sql.Field "UserId", DbSql.Field("Users.UserId")          '值是另一个表的字段
    sql.From "Users"
    DbSql.Insert "Employees", Sql


    Set sql = New SqlBuilder

    ' 使用 DbSql.Expression 包裹表达式（不会被转义）
    sql.Field "Name", DbSql.Expression("Trim([Name])")       '值是个表达式，要去除首位空格

    DbSql.Update "Employees", Sql

End Sub
```

### 19. 清空表数据

```vba
Public Sub ClearTableExample()
    Dim AffectedRows As Long

    ' 删除表中的所有数据
    AffectedRows = DbSql.Clear("TempTable")

    Debug.Print "删除了 " & AffectedRows & " 行"
End Sub
```

### 20. 获取表记录数

```vba
Public Sub TableCountExample()
    Dim TotalCount As Long

    ' 统计表中的总记录数
    TotalCount = DbSql.TableCount("Users")

    Debug.Print "用户表共有 " & TotalCount & " 条记录"
End Sub
```
