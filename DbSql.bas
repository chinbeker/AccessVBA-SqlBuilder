Attribute VB_Name = "DbSql"

'@Lang VBA

' ============================================================================
' Access VBA SQL Builder
' Repository: https://github.com/chinbeker/AccessVBA-SqlBuilder
' License: MIT
' Author: huachen
' Version: 1.0.0
' Created: 2026
' Description: Fluent SQL builder and database helper for Access VBA
' ============================================================================


Option Compare Database
Option Explicit

'数据库对象
Private Database As DAO.Database

' 判断空字符串
Private Function IsEmptyString(ByVal str As Variant) As Boolean
    If VBA.VarType(str) = VBA.vbString Then
        IsEmptyString = (VBA.Len(str) = 0)
    Else
        IsEmptyString = True
    End If
End Function
' 错误消息
Private Sub ShowError(ByRef ErrorObject As Object)
    MsgBox "发生错误：                         " & vbCrLf & vbCrLf & ErrorObject.Description, vbCritical + vbOKOnly, "系统错误"
End Sub

' 建立数据库连接
Private Sub CreateConnection()
    On Error GoTo ErrorHandler
    If Database Is Nothing Then Set Database = CurrentDb()
    Exit Sub

ErrorHandler:
    Call ShowError(Err)
    Exit Sub
End Sub

'检查表是否存在
Public Function TableExists(ByVal TableName As String) As Boolean
    On Error Resume Next
    If Not StringBase.IsNullOrEmpty(TableName) Then
        Dim tdf As DAO.TableDef
        Call CreateConnection
        Set tdf = Database.TableDefs(TableName)
        If Not tdf Is Nothing Then
            TableExists = (Left(tdf.Name, 4) <> "MSys")
            Set tdf = Nothing
        End If
    End If
    Err.Clear
    On Error GoTo 0
End Function

' 引用字段（拼接SQL字符串时不会加引号）
Public Function Field(ByVal Name As String) As String
    If Not IsEmptyString(Name) Then Field = "[$$]" & Name
End Function

' 引用表达式（拼接SQL字符串时不会加引号）
Public Function Expression(ByVal expr As String) As String
    Expression = Field(expr)
End Function
' 引用参数（拼接SQL字符串时不会加引号）
Public Function Parameter(ByVal Name As String) As String
    If Not IsEmptyString(Name) Then Parameter = "[$$][Param_" & Name & "]"
End Function


' 运行任意查询类SQL语句（动态集）
Public Function SelectDynaset(ByVal SqlString As String) As DAO.Recordset
    On Error GoTo ErrorHandler
    If IsEmptyString(SqlString) Then Exit Function
    Call CreateConnection
    Set SelectDynaset = Database.OpenRecordset(SqlString, dbOpenDynaset, dbSeeChanges)
    Exit Function
ErrorHandler:
    Call ShowError(Err)
    Exit Function
End Function

' 运行任意查询类SQL语句（快照）
Public Function SelectSnapshot(ByVal SqlString As String) As DAO.Recordset
    On Error GoTo ErrorHandler
    If IsEmptyString(SqlString) Then Exit Function
    Call CreateConnection
    Set SelectSnapshot = Database.OpenRecordset(SqlString, dbOpenSnapshot)
    Exit Function
ErrorHandler:
    Call ShowError(Err)
    Exit Function
End Function

' 运行任意非查询SQL语句
Public Function Execute(ByVal SqlString As String) As Long
    On Error GoTo ErrorHandler
    If IsEmptyString(SqlString) Then Exit Function
    Call CreateConnection
    Database.Execute SqlString
    Execute = Database.RecordsAffected
    Exit Function
ErrorHandler:
    Call ShowError(Err)
    Exit Function
End Function
' 获取表定义
Public Function TableDef(ByVal Name As String) As DAO.TableDef
    On Error GoTo ErrorHandler
    If IsEmptyString(Name) Then Exit Function
    Call CreateConnection
    Set TableDef = Database.TableDefs(Name)
    Exit Function
ErrorHandler:
    Call ShowError(Err)
    Exit Function
End Function

' 获取整张表（本地表）
Public Function OpenTable(ByVal Name As String) As DAO.Recordset
    On Error GoTo ErrorHandler
    If IsEmptyString(Name) Then Exit Function
    Call CreateConnection
    Set OpenTable = Database.OpenRecordset(Name, dbOpenTable)
    Exit Function
ErrorHandler:
    Call ShowError(Err)
    Exit Function
End Function

' 获取整张表（动态集）
Public Function TableDynaset(ByVal Name As String) As DAO.Recordset
    On Error GoTo ErrorHandler
    If IsEmptyString(Name) Then Exit Function
    Call CreateConnection
    Set TableDynaset = Database.OpenRecordset(Name, dbOpenDynaset, dbSeeChanges)
    Exit Function
ErrorHandler:
    Call ShowError(Err)
    Exit Function
End Function

' 获取整张表（快照）
Public Function TableSnapshot(ByVal Name As String) As DAO.Recordset
    On Error GoTo ErrorHandler
    If IsEmptyString(Name) Then Exit Function
    Call CreateConnection
    Set TableSnapshot = Database.OpenRecordset(Name, dbOpenSnapshot)
    Exit Function

ErrorHandler:
    Call ShowError(Err)
    Exit Function
End Function

' 插入数据
Public Function Insert(ByVal TableName As String, ByRef Sql As SqlBuilder) As Long
    On Error GoTo ErrorHandler
    If IsEmptyString(TableName) Then Exit Function
    Sql.Into TableName
    Dim SqlString As String
    SqlString = Sql.ToSqlString(4)
    Call CreateConnection
    If Sql.HasParam Then
        Dim Def As DAO.QueryDef
        Set Def = Database.CreateQueryDef("", SqlString)
        Sql.SetQueryDef Def
        Set Sql = Nothing
        Def.Execute dbFailOnError
        Insert = Def.RecordsAffected
        Def.Close
        Set Def = Nothing
    Else
        Set Sql = Nothing
        Database.Execute SqlString
        Insert = Database.RecordsAffected
    End If
    Exit Function

ErrorHandler:
    Call ShowError(Err)
    Exit Function
End Function

' 更新数据
Public Function Update(ByVal TableName As String, ByRef Sql As SqlBuilder) As Long
    On Error GoTo ErrorHandler
    If IsEmptyString(TableName) Then Exit Function
    Sql.From TableName
    Dim SqlString As String
    SqlString = Sql.ToSqlString(3)
    Call CreateConnection
    If Sql.HasParam Then
        Dim Def As DAO.QueryDef
        Set Def = Database.CreateQueryDef("", SqlString)
        Sql.SetQueryDef Def
        Set Sql = Nothing
        Def.Execute dbFailOnError
        Update = Def.RecordsAffected
        Def.Close
        Set Def = Nothing
    Else
        Set Sql = Nothing
        Database.Execute SqlString
        Update = Database.RecordsAffected
    End If
    Exit Function

ErrorHandler:
    Call ShowError(Err)
    Exit Function
End Function

' 删除数据
Public Function Delete(ByVal TableName As String, ByRef Sql As SqlBuilder) As Long
    On Error GoTo ErrorHandler
    If IsEmptyString(TableName) Then Exit Function
    Sql.From TableName
    Dim SqlString As String
    SqlString = Sql.ToSqlString(2)
    Call CreateConnection
    If Sql.HasParam Then
        Dim Def As DAO.QueryDef
        Set Def = Database.CreateQueryDef("", SqlString)
        Sql.SetQueryDef Def
        Set Sql = Nothing
        Def.Execute dbFailOnError
        Delete = Def.RecordsAffected
        Def.Close
        Set Def = Nothing
    Else
        Set Sql = Nothing
        Database.Execute SqlString
        Delete = Database.RecordsAffected
    End If
    Exit Function

ErrorHandler:
    Call ShowError(Err)
    Exit Function
End Function

' 清空数据
Public Function Clear(ByVal TableName As String) As Long
    On Error GoTo ErrorHandler
    If IsEmptyString(TableName) Then Exit Function
    Call CreateConnection
    Database.Execute "DELETE FROM " & TableName
    Clear = Database.RecordsAffected
    Exit Function
ErrorHandler:
    Call ShowError(Err)
    Exit Function
End Function

' 统计数量
Public Function Count(ByRef Sql As SqlBuilder) As Long
    On Error GoTo ErrorHandler
    Dim SqlString As String
    SqlString = Sql.ToSqlString(1)
    Call CreateConnection
    Dim rs As DAO.Recordset
    If Sql.HasParam Then
        Dim Def As DAO.QueryDef
        Set Def = Database.CreateQueryDef("", SqlString)
        Sql.SetQueryDef Def
        Set rs = Def.OpenRecordset(dbOpenSnapshot)
        Def.Close
        Set Def = Nothing
    Else
        Set rs = Database.OpenRecordset(SqlString, dbOpenSnapshot)
    End If
    Set Sql = Nothing
    Count = rs(0)
    rs.Close
    Set rs = Nothing
    Exit Function

ErrorHandler:
    Call ShowError(Err)
    Exit Function
End Function

' 统计整张表格数量
Public Function TableCount(ByVal TableName As String) As Long
    On Error GoTo ErrorHandler
    If IsEmptyString(TableName) Then Exit Function
    TableCount = Application.DCount("*", TableName)
    Exit Function

ErrorHandler:
    Call ShowError(Err)
    Exit Function
End Function

' 返回记录（快照）
Public Function Find(ByRef Sql As SqlBuilder) As DAO.Recordset
    On Error GoTo ErrorHandler
    Dim SqlString As String
    SqlString = Sql.ToSqlString(0)
    Call CreateConnection
    If Sql.HasParam Then
        Dim Def As DAO.QueryDef
        Set Def = Database.CreateQueryDef("", SqlString)
        Sql.SetQueryDef Def
        Set Sql = Nothing
        Set Find = Def.OpenRecordset(dbOpenSnapshot)
        Def.Close
        Set Def = Nothing
    Else
        Set Sql = Nothing
        Set Find = Database.OpenRecordset(SqlString, dbOpenSnapshot)
    End If
    Exit Function

ErrorHandler:
    Call ShowError(Err)
    Exit Function
End Function

' 返回记录（动态集）
Public Function Record(ByRef Sql As SqlBuilder) As DAO.Recordset
    On Error GoTo ErrorHandler
    Dim SqlString As String
    SqlString = Sql.ToSqlString(0)
    Call CreateConnection
    If Sql.HasParam Then
        Dim Def As DAO.QueryDef
        Set Def = Database.CreateQueryDef("", SqlString)
        Sql.SetQueryDef Def
        Set Sql = Nothing
        Set Record = Def.OpenRecordset(dbOpenDynaset, dbSeeChanges)
        Def.Close
        Set Def = Nothing
    Else
        Set Sql = Nothing
        Set Record = Database.OpenRecordset(SqlString, dbOpenDynaset, dbSeeChanges)
    End If
    Exit Function

ErrorHandler:
    Call ShowError(Err)
    Exit Function
End Function

' 第一条记录（快照）
Public Function First(ByRef Sql As SqlBuilder) As DAO.Recordset
    On Error GoTo ErrorHandler
    Sql.Top 1
    Dim SqlString As String
    SqlString = Sql.ToSqlString(0)
    Call CreateConnection
    If Sql.HasParam Then
        Dim Def As DAO.QueryDef
        Set Def = Database.CreateQueryDef("", SqlString)
        Sql.SetQueryDef Def
        Set Sql = Nothing
        Set First = Def.OpenRecordset(dbOpenSnapshot)
        Def.Close
        Set Def = Nothing
    Else
        Set Sql = Nothing
        Set First = Database.OpenRecordset(SqlString, dbOpenSnapshot)
    End If
    Exit Function

ErrorHandler:
    Call ShowError(Err)
    Exit Function
End Function

' 第一条记录（动态集）
Public Function FirstRecord(ByRef Sql As SqlBuilder) As DAO.Recordset
    On Error GoTo ErrorHandler
    Sql.Top 1
    Dim SqlString As String
    SqlString = Sql.ToSqlString(0)
    Call CreateConnection
    If Sql.HasParam Then
        Dim Def As DAO.QueryDef
        Set Def = Database.CreateQueryDef("", SqlString)
        Sql.SetQueryDef Def
        Set Sql = Nothing
        Set FirstRecord = Def.OpenRecordset(dbOpenDynaset, dbSeeChanges)
        Def.Close
        Set Def = Nothing
    Else
        Set Sql = Nothing
        Set FirstRecord = Database.OpenRecordset(SqlString, dbOpenDynaset, dbSeeChanges)
    End If
    Exit Function

ErrorHandler:
    Call ShowError(Err)
    Exit Function
End Function


' 表格第一条记录（快照）
Public Function TableFirst(ByVal TableName As String, Optional ByVal OrderByField As String) As DAO.Recordset
    On Error GoTo ErrorHandler
    If IsEmptyString(TableName) Then Exit Function
    Dim Sql As New SqlBuilder
    Sql.SelectAll
    Sql.From TableName
    If Not IsMissing(OrderByField) And Not IsEmptyString(OrderByField) Then Sql.Order OrderByField
    Set TableFirst = First(Sql)
    Set Sql = Nothing
    Exit Function

ErrorHandler:
    Call ShowError(Err)
    Exit Function
End Function

' 表格第一条记录（动态集）
Public Function TableFirstRecord(ByVal TableName As String, Optional ByVal OrderByField As String) As DAO.Recordset
    On Error GoTo ErrorHandler
    If IsEmptyString(TableName) Then Exit Function
    Dim Sql As New SqlBuilder
    Sql.SelectAll
    Sql.From TableName
    If Not IsMissing(OrderByField) And Not IsEmptyString(OrderByField) Then Sql.Order OrderByField
    Set TableFirstRecord = FirstRecord(Sql)
    Set Sql = Nothing
    Exit Function

ErrorHandler:
    Call ShowError(Err)
    Exit Function
End Function

' 表格最后一条记录（快照）
Public Function TableLast(ByVal TableName As String, Optional ByVal OrderByField As String) As DAO.Recordset
    On Error GoTo ErrorHandler
    If IsEmptyString(TableName) Then Exit Function
    Dim Sql As New SqlBuilder
    Sql.SelectAll
    Sql.From TableName
    If Not IsMissing(OrderByField) And Not IsEmptyString(OrderByField) Then Sql.Order OrderByField, True
    Set TableLast = First(Sql)
    Set Sql = Nothing
    Exit Function

ErrorHandler:
    Call ShowError(Err)
    Exit Function
End Function

' 表格最后一条记录（动态集）
Public Function TableLastRecord(ByVal TableName As String, Optional ByVal OrderByField As String) As DAO.Recordset
    On Error GoTo ErrorHandler
    If IsEmptyString(TableName) Then Exit Function
    Dim Sql As New SqlBuilder
    Sql.SelectAll
    Sql.From TableName
    If Not IsMissing(OrderByField) And Not IsEmptyString(OrderByField) Then Sql.Order OrderByField, True
    Set TableLastRecord = FirstRecord(Sql)
    Set Sql = Nothing
    Exit Function

ErrorHandler:
    Call ShowError(Err)
    Exit Function
End Function


' 获取第一条记录的指定字段值
Public Function GetValue(ByRef Sql As SqlBuilder) As Variant
    On Error GoTo ErrorHandler
    Sql.Top 1
    Dim SqlString As String
    SqlString = Sql.ToSqlString(0)
    Call CreateConnection
    Dim rs As DAO.Recordset
    If Sql.HasParam Then
        Dim Def As DAO.QueryDef
        Set Def = Database.CreateQueryDef("", SqlString)
        Sql.SetQueryDef Def
        Set Sql = Nothing
        Set rs = Def.OpenRecordset(dbOpenSnapshot)
        Def.Close
        Set Def = Nothing
    Else
        Set Sql = Nothing
        Set rs = Database.OpenRecordset(SqlString, dbOpenSnapshot)
    End If
    If Not rs.EOF Then
        GetValue = rs(0)
    Else
        GetValue = Null
    End If
    rs.Close
    Set rs = Nothing
    Exit Function

ErrorHandler:
    Call ShowError(Err)
    Exit Function
End Function

' 获取第一条记录的指定字段值
Public Function GetValueFromSql(ByVal SqlString As String) As Variant
    On Error GoTo ErrorHandler
    If IsEmptyString(SqlString) Then Exit Function

    Call CreateConnection
    Dim rs As DAO.Recordset
    Set rs = Database.OpenRecordset(SqlString, dbOpenSnapshot)

    If Not rs.EOF Then
        GetValueFromSql = rs(0)
    Else
        GetValueFromSql = Null
    End If

    rs.Close
    Set rs = Nothing
    Exit Function

ErrorHandler:
    Call ShowError(Err)
    Exit Function
End Function


' 快速查询记录指定字段值
Public Function Lookup(ByVal TableName As String, ByVal Field As String, ByVal Condition As String) As Variant
    On Error GoTo ErrorHandler
    If IsEmptyString(TableName) Then Exit Function
    If IsEmptyString(Field) Then Exit Function
    If IsEmptyString(Condition) Then Exit Function
    Lookup = Application.DLookup(Field, TableName, Condition)
    Exit Function

ErrorHandler:
    Call ShowError(Err)
    Exit Function
End Function

' 快速查询第一条记录指定字段值
Public Function FirstValue(ByVal TableName As String, ByVal Field As String, Optional ByVal Condition As String, Optional ByVal OrderByField As String) As Variant
    On Error GoTo ErrorHandler
    If IsEmptyString(TableName) Then Exit Function
    If IsEmptyString(Field) Then Exit Function
    If IsMissing(OrderByField) Or IsEmptyString(OrderByField) Then
        If IsMissing(Condition) Or IsEmptyString(Condition) Then
            FirstValue = Application.DFirst(Field, TableName)
        Else
            FirstValue = Application.DFirst(Field, TableName, Condition)
        End If
    Else
        Dim Sql As New SqlBuilder
        Sql.Top 1
        Sql.Field Field
        Sql.From TableName
        Sql.Where Condition
        Sql.Order OrderByField
        FirstValue = GetValue(Sql)
        Set Sql = Nothing
    End If
    Exit Function

ErrorHandler:
    Call ShowError(Err)
    Exit Function
End Function


' 快速查询最后一条记录指定字段值
Public Function LastValue(ByVal TableName As String, ByVal Field As String, Optional ByVal Condition As String, Optional ByVal OrderByField As String) As Variant
    On Error GoTo ErrorHandler
    If IsEmptyString(TableName) Then Exit Function
    If IsEmptyString(Field) Then Exit Function
    If IsMissing(OrderByField) Or IsEmptyString(OrderByField) Then
        If IsMissing(Condition) Or IsEmptyString(Condition) Then
            LastValue = Application.DLast(Field, TableName)
        Else
            LastValue = Application.DLast(Field, TableName, Condition)
        End If
    Else
        Dim Sql As New SqlBuilder
        Sql.Field Field
        Sql.From TableName
        Sql.Where Condition
        Sql.Order OrderByField, True
        LastValue = GetValue(Sql)
        Set Sql = Nothing
    End If
    Exit Function

ErrorHandler:
    Call ShowError(Err)
    Exit Function
End Function


' 设置第一条记录的指定字段值
Public Function SetValue(ByVal TableName As String, ByVal Field As String, ByVal value As Variant, ByVal Condition As String) As Boolean
    On Error GoTo ErrorHandler
    If IsEmptyString(TableName) Then Exit Function
    If IsEmptyString(Field) Then Exit Function
    If IsEmptyString(Condition) Then Exit Function
    Dim Sql As New SqlBuilder
    Sql.From TableName
    Sql.Field Field, value
    Sql.Where Condition
    Dim SqlString As String
    SqlString = Sql.ToSqlString(3)
    Call CreateConnection
    Dim Affected As Long
    If Sql.HasParam Then
        Dim Def As DAO.QueryDef
        Set Def = Database.CreateQueryDef("", SqlString)
        Sql.SetQueryDef Def
        Set Sql = Nothing
        Def.Execute dbFailOnError
        Affected = Def.RecordsAffected
        Def.Close
        Set Def = Nothing
    Else
        Set Sql = Nothing
        Database.Execute SqlString
        Affected = Database.RecordsAffected
    End If
    If Affected > 0 Then
        SetValue = True
    Else
        SetValue = False
    End If
    Exit Function

ErrorHandler:
    Call ShowError(Err)
    Exit Function
End Function

' 联合查询（去重）
Public Function Union(ParamArray SqlBuilders() As Variant) As DAO.Recordset
    On Error GoTo ErrorHandler
    Dim Length As Long
    Length = UBound(SqlBuilders) - LBound(SqlBuilders) + 1

    If Length < 2 Then
        MsgBox "至少需要两个查询进行 Union", vbCritical + vbOKOnly, "系统错误"
        Exit Function
    End If

    Dim SqlString As String
    Dim i As Long
    Dim TempSql As String

    TempSql = SqlBuilders(LBound(SqlBuilders)).ToSqlString(0, False)

    SqlString = TempSql
    Length = UBound(SqlBuilders)
    For i = (LBound(SqlBuilders) + 1) To Length
        If TypeOf SqlBuilders(i) Is SqlBuilder Then
            TempSql = SqlBuilders(i).ToSqlString(0, False)
            SqlString = SqlString & " UNION " & TempSql
        Else
            MsgBox "参数必须是 SqlBuilder 对象类型", vbCritical + vbOKOnly, "系统错误"
            Exit Function
        End If
    Next i

    If VBA.Len(SqlString) > 0 Then
        SqlString = SqlString & ";"
        Call CreateConnection
        Set Union = Database.OpenRecordset(SqlString, dbOpenSnapshot)
    Else
        MsgBox "SQL 语法错误", vbCritical + vbOKOnly, "系统错误"
    End If
    Exit Function

ErrorHandler:
    Call ShowError(Err)
    Exit Function
End Function

' 联合查询（有重复）
Public Function UnionAll(ParamArray SqlBuilders() As Variant) As DAO.Recordset
    On Error GoTo ErrorHandler
    Dim Length As Long
    Length = UBound(SqlBuilders) - LBound(SqlBuilders) + 1

    If Length < 2 Then
        MsgBox "至少需要两个查询进行 Union", vbCritical + vbOKOnly, "系统错误"
        Exit Function
    End If

    Dim SqlString As String
    Dim i As Long
    Dim TempSql As String

    TempSql = SqlBuilders(LBound(SqlBuilders)).ToSqlString(0, False)

    SqlString = TempSql
    Length = UBound(SqlBuilders)
    For i = (LBound(SqlBuilders) + 1) To Length
        If TypeOf SqlBuilders(i) Is SqlBuilder Then
            TempSql = SqlBuilders(i).ToSqlString(0, False)
            SqlString = SqlString & " UNION All " & TempSql
        Else
            MsgBox "参数必须是 SqlBuilder 对象类型", vbCritical + vbOKOnly, "系统错误"
            Exit Function
        End If
    Next i

    If VBA.Len(SqlString) > 0 Then
        SqlString = SqlString & ";"
        Call CreateConnection
        Set UnionAll = Database.OpenRecordset(SqlString, dbOpenSnapshot)
    Else
        MsgBox "SQL 语法错误", vbCritical + vbOKOnly, "系统错误"
    End If
    Exit Function

ErrorHandler:
    Call ShowError(Err)
    Exit Function
End Function

' 创建临时表（克隆目标表的结构）（只能使用ADO，DAO不支持内存临时表）
Public Function CreateTempTable(ByVal CloneTable As String) As ADODB.Recordset
    On Error GoTo ErrorHandler
    Dim source As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim fieldDef As ADODB.Field

    If Not DbSql.TableExists(CloneTable) Then
        MsgBox "数据表不存在", vbCritical + vbOKOnly, "系统错误"
        Set CreateTempTable = Nothing
        Exit Function
    End If

    Set source = New ADODB.Recordset
    source.Open CloneTable, Application.CurrentProject.AccessConnection, adOpenForwardOnly, adLockReadOnly

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

    For Each fieldDef In source.Fields
        rs.Fields.Append fieldDef.Name, fieldDef.Type, fieldDef.DefinedSize, 106
        rs.Fields(fieldDef.Name).Precision = fieldDef.Precision
        rs.Fields(fieldDef.Name).NumericScale = fieldDef.NumericScale
    Next fieldDef

    rs.Open
    Set rs.ActiveConnection = Nothing
    source.Close
    Set source = Nothing
    Set CreateTempTable = rs
    Exit Function

ErrorHandler:
    MsgBox "创建临时表失败", vbCritical + vbOKOnly, "系统错误"
    If Not source Is Nothing Then
        If source.State = adStateOpen Then source.Close
        Set source = Nothing
    End If
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    Set CreateTempTable = Nothing
    Exit Function
End Function
