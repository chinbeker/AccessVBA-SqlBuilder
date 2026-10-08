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
Private CurrentDatabase As DAO.Database

' 判断空字符串
Private Function IsEmptyString(ByVal str As Variant) As Boolean
    If VBA.VarType(str) = VBA.vbString Then
        IsEmptyString = (VBA.Len(str) = 0)
    Else
        IsEmptyString = True
    End If
End Function

'检查表是否存在
Private Function TableExists(ByVal TableName As String) As Boolean
    On Error Resume Next
    If Not IsEmptyString(TableName) Then
        Dim tdf As DAO.TableDef
        Call ConnectDatabase
        Set tdf = CurrentDatabase.TableDefs(TableName)
        TableExists = (Not tdf Is Nothing)
        Set tdf = Nothing
    End If
    On Error GoTo 0
End Function

' 建立数据库连接
Private Sub ConnectDatabase()
    If CurrentDatabase Is Nothing Then Set CurrentDatabase = Application.CurrentDb
End Sub

' 使用指定数据库
Public Sub UseCustomDatabase(ByRef DB As DAO.Database)
    If Not DB Is Nothing Then Set CurrentDatabase = DB
End Sub

' 使用默认数据库
Public Sub UseDefaultDatabase()
    Set CurrentDatabase = Application.CurrentDb
End Sub

' 获取默认数据库
Public Function GetCurrentDatabase() As DAO.Database
    Call ConnectDatabase
    Set GetCurrentDatabase = CurrentDatabase
End Function

' 引用字段（拼接SQL字符串时不会加引号）
Public Function Field(ByVal FieldName As String) As String
    If IsEmptyString(FieldName) Then
        Err.Raise 449, "DbSql.Field", "字段名称不能为空"
        Exit Function
    End If
    Field = "[$$]" & FieldName
End Function

' 引用表达式（拼接SQL字符串时不会加引号）
Public Function Expression(ByVal expr As String) As String
    If IsEmptyString(expr) Then
        Err.Raise 449, "DbSql.Expression", "表达式不能为空"
        Exit Function
    End If
    Expression = "[$$]" & expr
End Function

' 引用参数（拼接SQL字符串时不会加引号）
Public Function Parameter(ByVal ParamName As String) As String
    If IsEmptyString(ParamName) Then
        Err.Raise 449, "DbSql.Expression", "参数名称不能为空"
        Exit Function
    End If
    ParamName = VBA.Replace(ParamName, " ", "_")
    ParamName = VBA.Replace(ParamName, ".", "_")
    ParamName = VBA.Replace(ParamName, "!", "_")
    Parameter = "[$$][Param_" & ParamName & "]"
End Function


' 运行任意查询类SQL语句（动态集）
Public Function SelectDynaset(ByVal SqlString As String) As DAO.Recordset
    On Error GoTo ErrorHandler
    If IsEmptyString(SqlString) Then Exit Function
    Call ConnectDatabase
    Set SelectDynaset = CurrentDatabase.OpenRecordset(SqlString, dbOpenDynaset, dbSeeChanges)
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "DbSql.SelectDynaset" & vbCrLf & Err.Source, Err.Description
    Exit Function
End Function

' 运行任意查询类SQL语句（快照）
Public Function SelectSnapshot(ByVal SqlString As String) As DAO.Recordset
    On Error GoTo ErrorHandler
    If IsEmptyString(SqlString) Then Exit Function
    Call ConnectDatabase
    Set SelectSnapshot = CurrentDatabase.OpenRecordset(SqlString, dbOpenSnapshot)
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "DbSql.SelectSnapshot" & vbCrLf & Err.Source, Err.Description
    Exit Function
End Function

' 运行任意非查询SQL语句
Public Function Execute(ByVal SqlString As String) As Long
    On Error GoTo ErrorHandler
    If IsEmptyString(SqlString) Then Exit Function
    Call ConnectDatabase
    CurrentDatabase.Execute SqlString, dbFailOnError
    Execute = CurrentDatabase.RecordsAffected
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "DbSql.Execute" & vbCrLf & Err.Source, Err.Description
    Exit Function
End Function


' 获取表定义
Public Function TableDef(ByVal TableName As String) As DAO.TableDef
    On Error GoTo ErrorHandler
    If IsEmptyString(TableName) Then Exit Function
    Call ConnectDatabase
    Set TableDef = CurrentDatabase.TableDefs(TableName)
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "DbSql.TableDef" & vbCrLf & Err.Source, Err.Description
    Exit Function
End Function

' 获取整张表
Public Function OpenTable(ByVal TableName As String, Optional ByVal ReadOnly As Boolean = False) As DAO.Recordset
    On Error GoTo ErrorHandler
    If IsEmptyString(TableName) Then Exit Function
    Call ConnectDatabase
    If ReadOnly Then
        Set OpenTable = CurrentDatabase.OpenRecordset(TableName, dbOpenSnapshot)
    Else
        Set OpenTable = CurrentDatabase.OpenRecordset(TableName, dbOpenDynaset, dbSeeChanges)
    End If
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "DbSql.OpenTable" & vbCrLf & Err.Source, Err.Description
    Exit Function
End Function

' 指定表格查找记录
Public Function TableFind(ByVal TableName As String, ByVal Condition As String, Optional ByVal ReadOnly As Boolean = False) As DAO.Recordset
    On Error GoTo ErrorHandler
    If IsEmptyString(TableName) Then Exit Function
    If IsEmptyString(Condition) Then
        Err.Raise 449, "DbSql.TableFind", "查询条件不能为空"
        Exit Function
    End If
    Dim SqlString As String
    SqlString = "SELECT " & TableName & ".* FROM " & TableName & " WHERE (" & Condition & ")"
    Call ConnectDatabase
    If ReadOnly Then
        Set TableFind = CurrentDatabase.OpenRecordset(SqlString, dbOpenSnapshot)
    Else
        Set TableFind = CurrentDatabase.OpenRecordset(SqlString, dbOpenDynaset, dbSeeChanges)
    End If
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "DbSql.TableFind" & vbCrLf & Err.Source, Err.Description
    Exit Function
End Function


' 指定表格查找第一条记录
Public Function TableFindFirst(ByVal TableName As String, ByVal Condition As String, Optional ByVal OrderBy As String, Optional ByVal ReadOnly As Boolean = False) As DAO.Recordset
    On Error GoTo ErrorHandler
    If IsEmptyString(TableName) Then Exit Function
    If IsEmptyString(Condition) Then
        Err.Raise 449, "DbSql.TableFindFirst", "查询条件不能为空"
        Exit Function
    End If
    Dim SqlString As String
    SqlString = "SELECT TOP 1 " & TableName & ".* FROM " & TableName & " WHERE (" & Condition & ")"
    If Not VBA.IsMissing(OrderBy) And Not IsEmptyString(OrderBy) Then SqlString = SqlString & " ORDER BY " & OrderBy
    Call ConnectDatabase
    If ReadOnly Then
        Set TableFindFirst = CurrentDatabase.OpenRecordset(SqlString, dbOpenSnapshot)
    Else
        Set TableFindFirst = CurrentDatabase.OpenRecordset(SqlString, dbOpenDynaset, dbSeeChanges)
    End If
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "DbSql.TableFindFirst" & vbCrLf & Err.Source, Err.Description
    Exit Function
End Function

' 表格第一条记录
Public Function TableFirst(ByVal TableName As String, Optional ByVal OrderBy As String, Optional ByVal ReadOnly As Boolean = False) As DAO.Recordset
    On Error GoTo ErrorHandler
    If IsEmptyString(TableName) Then Exit Function
    Dim SqlString As String
    SqlString = "SELECT TOP 1 " & TableName & ".* FROM " & TableName
    If Not VBA.IsMissing(OrderBy) And Not IsEmptyString(OrderBy) Then SqlString = SqlString & " ORDER BY " & OrderBy
    Call ConnectDatabase
    If ReadOnly Then
        Set TableFirst = CurrentDatabase.OpenRecordset(SqlString, dbOpenSnapshot)
    Else
        Set TableFirst = CurrentDatabase.OpenRecordset(SqlString, dbOpenDynaset, dbSeeChanges)
    End If
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "DbSql.TableFirst" & vbCrLf & Err.Source, Err.Description
    Exit Function
End Function

' 表格最后一条记录
Public Function TableLast(ByVal TableName As String, ByVal OrderByField As String, Optional ByVal ReadOnly As Boolean = False) As DAO.Recordset
    On Error GoTo ErrorHandler
    If IsEmptyString(TableName) Then Exit Function
    If IsEmptyString(OrderByField) Then
        Err.Raise 449, "DbSql.TableLast", "排序字段不能为空"
        Exit Function
    End If
    Dim SqlString As String
    SqlString = "SELECT TOP 1 " & TableName & ".* FROM " & TableName & " ORDER BY " & OrderByField & " DESC"
    Call ConnectDatabase
    If ReadOnly Then
        Set TableLast = CurrentDatabase.OpenRecordset(SqlString, dbOpenSnapshot)
    Else
        Set TableLast = CurrentDatabase.OpenRecordset(SqlString, dbOpenDynaset, dbSeeChanges)
    End If
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "DbSql.TableLast" & vbCrLf & Err.Source, Err.Description
    Exit Function
End Function


' 插入数据
Public Function Insert(ByVal TableName As String, ByRef Sql As SqlBuilder) As Long
    On Error GoTo ErrorHandler
    If IsEmptyString(TableName) Then Exit Function
    Sql.Into TableName
    Dim SqlString As String
    SqlString = Sql.ToSqlString(4)
    Call ConnectDatabase
    If Sql.HasParam Then
        Dim Def As DAO.QueryDef
        Set Def = CurrentDatabase.CreateQueryDef("", SqlString)
        Sql.SetQueryDef Def
        Def.Execute dbFailOnError
        Insert = Def.RecordsAffected
        Def.Close
        Set Def = Nothing
    Else
        CurrentDatabase.Execute SqlString, dbFailOnError
        Insert = CurrentDatabase.RecordsAffected
    End If
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "DbSql.Insert" & vbCrLf & Err.Source, Err.Description
    Exit Function
End Function

' 更新数据
Public Function Update(ByVal TableName As String, ByRef Sql As SqlBuilder) As Long
    On Error GoTo ErrorHandler
    If IsEmptyString(TableName) Then Exit Function
    Sql.From TableName
    Dim SqlString As String
    SqlString = Sql.ToSqlString(3)
    Call ConnectDatabase
    If Sql.HasParam Then
        Dim Def As DAO.QueryDef
        Set Def = CurrentDatabase.CreateQueryDef("", SqlString)
        Sql.SetQueryDef Def
        Def.Execute dbFailOnError
        Update = Def.RecordsAffected
        Def.Close
        Set Def = Nothing
    Else
        CurrentDatabase.Execute SqlString, dbFailOnError
        Update = CurrentDatabase.RecordsAffected
    End If
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "DbSql.Update" & vbCrLf & Err.Source, Err.Description
    Exit Function
End Function

' 删除数据(快速)
Public Function Delete(ByVal TableName As String, ByVal Condition As String) As Long
    On Error GoTo ErrorHandler
    If IsEmptyString(TableName) Then Exit Function
    If IsEmptyString(Condition) Then
        Err.Raise 449, "DbSql.Delete", "DELETE 必须定义 WHERE 语句"
        Exit Function
    End If
    Dim SqlString As String
    SqlString = "DELETE" & " FROM " & TableName & " WHERE (" & Condition & ")"
    Call ConnectDatabase
    CurrentDatabase.Execute SqlString, dbFailOnError
    Delete = CurrentDatabase.RecordsAffected
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "DbSql.Delete" & vbCrLf & Err.Source, Err.Description
    Exit Function
End Function

' 删除数据(复杂条件)
Public Function DeleteFind(ByVal TableName As String, ByRef Sql As SqlBuilder) As Long
    On Error GoTo ErrorHandler
    If IsEmptyString(TableName) Then Exit Function
    Sql.From TableName
    Dim SqlString As String
    SqlString = Sql.ToSqlString(2)
    Call ConnectDatabase
    If Sql.HasParam Then
        Dim Def As DAO.QueryDef
        Set Def = CurrentDatabase.CreateQueryDef("", SqlString)
        Sql.SetQueryDef Def
        Def.Execute dbFailOnError
        DeleteFind = Def.RecordsAffected
        Def.Close
        Set Def = Nothing
    Else
        CurrentDatabase.Execute SqlString, dbFailOnError
        DeleteFind = CurrentDatabase.RecordsAffected
    End If
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "DbSql.DeleteFind" & vbCrLf & Err.Source, Err.Description
    Exit Function
End Function


' 清空数据
Public Function Clear(ByVal TableName As String) As Long
    On Error GoTo ErrorHandler
    If IsEmptyString(TableName) Then Exit Function
    Call ConnectDatabase
    CurrentDatabase.Execute "DELETE FROM " & TableName, dbFailOnError
    Clear = CurrentDatabase.RecordsAffected
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "DbSql.Clear" & vbCrLf & Err.Source, Err.Description
    Exit Function
End Function

' 统计数量
Public Function Count(ByVal TableName As String, ByVal Condition As String) As Long
    On Error GoTo ErrorHandler
    If IsEmptyString(TableName) Then Exit Function
    If IsEmptyString(Condition) Then
        Err.Raise 449, "DbSql.Count", "查询条件不能为空"
        Exit Function
    End If
    Count = Application.DCount("*", TableName, Condition)
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "DbSql.Count" & vbCrLf & Err.Source, Err.Description
    Exit Function
End Function


' 统计整张表格数量
Public Function TableCount(ByVal TableName As String) As Long
    On Error GoTo ErrorHandler
    If IsEmptyString(TableName) Then Exit Function
    TableCount = Application.DCount("*", TableName)
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "DbSql.TableCount" & vbCrLf & Err.Source, Err.Description
    Exit Function
End Function

' 统计数量
Public Function FindCount(ByRef Sql As SqlBuilder) As Long
    On Error GoTo ErrorHandler
    Dim SqlString As String
    SqlString = Sql.ToSqlString(1)
    Call ConnectDatabase
    Dim rs As DAO.Recordset
    If Sql.HasParam Then
        Dim Def As DAO.QueryDef
        Set Def = CurrentDatabase.CreateQueryDef("", SqlString)
        Sql.SetQueryDef Def
        Set rs = Def.OpenRecordset(dbOpenSnapshot)
        Def.Close
        Set Def = Nothing
    Else
        Set rs = CurrentDatabase.OpenRecordset(SqlString, dbOpenSnapshot)
    End If
    FindCount = rs(0)
    rs.Close
    Set rs = Nothing
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "DbSql.FindCount" & vbCrLf & Err.Source, Err.Description
    Exit Function
End Function

' 返回记录（快照）
Public Function Find(ByRef Sql As SqlBuilder) As DAO.Recordset
    On Error GoTo ErrorHandler
    Dim SqlString As String
    SqlString = Sql.ToSqlString(0)
    Call ConnectDatabase
    If Sql.HasParam Then
        Dim Def As DAO.QueryDef
        Set Def = CurrentDatabase.CreateQueryDef("", SqlString)
        Sql.SetQueryDef Def
        If Sql.ReadOnly Then
            Set Find = Def.OpenRecordset(dbOpenSnapshot)
        Else
            Set Find = Def.OpenRecordset(dbOpenDynaset, dbSeeChanges)
        End If
        Def.Close
        Set Def = Nothing
    Else
        If Sql.ReadOnly Then
            Set Find = CurrentDatabase.OpenRecordset(SqlString, dbOpenSnapshot)
        Else
            Set Find = CurrentDatabase.OpenRecordset(SqlString, dbOpenDynaset, dbSeeChanges)
        End If
    End If
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "DbSql.Find" & vbCrLf & Err.Source, Err.Description
    Exit Function
End Function

' 第一条记录（快照）
Public Function First(ByRef Sql As SqlBuilder) As DAO.Recordset
    On Error GoTo ErrorHandler
    Dim SqlString As String
    Sql.Top 1
    SqlString = Sql.ToSqlString(0)
    Call ConnectDatabase
    If Sql.HasParam Then
        Dim Def As DAO.QueryDef
        Set Def = CurrentDatabase.CreateQueryDef("", SqlString)
        Sql.SetQueryDef Def
        If Sql.ReadOnly Then
            Set First = Def.OpenRecordset(dbOpenSnapshot)
        Else
            Set First = Def.OpenRecordset(dbOpenDynaset, dbSeeChanges)
        End If
        Def.Close
        Set Def = Nothing
    Else
        If Sql.ReadOnly Then
            Set First = CurrentDatabase.OpenRecordset(SqlString, dbOpenSnapshot)
        Else
            Set First = CurrentDatabase.OpenRecordset(SqlString, dbOpenDynaset, dbSeeChanges)
        End If
    End If
    Sql.Top 0
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "DbSql.First" & vbCrLf & Err.Source, Err.Description
    Exit Function
End Function


' 获取第一条记录的指定字段值
Public Function GetValue(ByRef Sql As SqlBuilder) As Variant
    On Error GoTo ErrorHandler
    Dim rs As DAO.Recordset
    Sql.ReadOnly
    Set rs = DbSql.First(Sql)
    If Not rs.EOF Then
        GetValue = rs(0)
    Else
        GetValue = Null
    End If
    rs.Close
    Set rs = Nothing
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "DbSql.GetValue" & vbCrLf & Err.Source, Err.Description
    Exit Function
End Function

' 获取第一条记录的指定字段值
Public Function GetValueFromSql(ByVal SqlString As String) As Variant
    On Error GoTo ErrorHandler
    GetValueFromSql = Null
    If IsEmptyString(SqlString) Then Exit Function
    Dim rs As DAO.Recordset
    Call ConnectDatabase
    Set rs = CurrentDatabase.OpenRecordset(SqlString, dbOpenSnapshot)
    If Not rs.EOF Then GetValueFromSql = rs(0)
    rs.Close
    Set rs = Nothing
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "DbSql.GetValueFromSql" & vbCrLf & Err.Source, Err.Description
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
    Err.Raise Err.Number, "DbSql.Lookup" & vbCrLf & Err.Source, Err.Description
    Exit Function
End Function

' 快速查询第一条记录指定字段值
Public Function FirstValue(ByVal TableName As String, ByVal Field As String, Optional ByVal Condition As String, Optional ByVal OrderByField As String) As Variant
    On Error GoTo ErrorHandler
    If IsEmptyString(TableName) Then Exit Function
    If IsEmptyString(Field) Then Exit Function
    If VBA.IsMissing(OrderByField) Or IsEmptyString(OrderByField) Then
        If VBA.IsMissing(Condition) Or IsEmptyString(Condition) Then
            FirstValue = Application.DFirst(Field, TableName)
        Else
            FirstValue = Application.DFirst(Field, TableName, Condition)
        End If
    Else
        Dim SqlString As String
        SqlString = "SELECT TOP 1 " & TableName & "." & Field & " FROM " & TableName
        If Not IsEmptyString(Condition) Then SqlString = SqlString & " WHERE (" & Condition & ")"
        SqlString = SqlString & " ORDER BY " & OrderByField
        FirstValue = DbSql.GetValueFromSql(SqlString)
    End If
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "DbSql.FirstValue" & vbCrLf & Err.Source, Err.Description
    Exit Function
End Function


' 快速查询最后一条记录指定字段值
Public Function LastValue(ByVal TableName As String, ByVal Field As String, Optional ByVal Condition As String, Optional ByVal OrderByField As String) As Variant
    On Error GoTo ErrorHandler
    If IsEmptyString(TableName) Then Exit Function
    If IsEmptyString(Field) Then Exit Function
    If VBA.IsMissing(OrderByField) Or IsEmptyString(OrderByField) Then
        If VBA.IsMissing(Condition) Or IsEmptyString(Condition) Then
            LastValue = Application.DLast(Field, TableName)
        Else
            LastValue = Application.DLast(Field, TableName, Condition)
        End If
    Else
        Dim SqlString As String
        SqlString = "SELECT TOP 1 " & TableName & "." & Field & " FROM " & TableName
        If Not IsEmptyString(Condition) Then SqlString = SqlString & " WHERE (" & Condition & ")"
        SqlString = SqlString & " ORDER BY " & OrderByField & " DESC"
        LastValue = DbSql.GetValueFromSql(SqlString)
    End If
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "DbSql.LastValue" & vbCrLf & Err.Source, Err.Description
    Exit Function
End Function


' 设置第一条记录的指定字段值
Public Function SetValue(ByVal TableName As String, ByVal Field As String, ByVal Value As Variant, ByVal Condition As String) As Boolean
    On Error GoTo ErrorHandler
    If IsEmptyString(TableName) Then Exit Function
    If IsEmptyString(Field) Then Exit Function
    If IsEmptyString(Condition) Then Exit Function
    Dim Affected As Long
    Dim Sql As New SqlBuilder
    Sql.Field Field, Value
    Sql.Where Condition
    Affected = DbSql.Update(TableName, Sql)
    Set Sql = Nothing
    If Affected > 0 Then
        SetValue = True
    Else
        SetValue = False
    End If
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "DbSql.SetValue" & vbCrLf & Err.Source, Err.Description
    Exit Function
End Function



' 联合查询（唯一值）
Public Function Union(ParamArray SqlBuilders() As Variant) As DAO.Recordset
    Dim Length As Long
    Length = UBound(SqlBuilders) - LBound(SqlBuilders) + 1

    If Length < 2 Then
        Err.Raise 3075, "DbSql.Union", "至少需要两个查询进行 Union"
        Exit Function
    End If

    Dim SqlString As String
    Dim i As Long
    Dim TempSql As String

    TempSql = SqlBuilders(LBound(SqlBuilders)).ToSqlString(0, False)

    SqlString = TempSql
    Length = UBound(SqlBuilders)
    For i = (LBound(SqlBuilders) + 1) To Length
        If VBA.IsObject(SqlBuilders(i)) And TypeOf SqlBuilders(i) Is SqlBuilder Then
            TempSql = SqlBuilders(i).ToSqlString(0, False)
            SqlString = SqlString & " UNION " & TempSql
        Else
            Err.Raise 3001, "DbSql.Union", "参数必须是 SqlBuilder 对象类型"
            Exit Function
        End If
    Next i

    If VBA.Len(SqlString) > 0 Then
        SqlString = SqlString & ";"
        Call ConnectDatabase
        Set Union = CurrentDatabase.OpenRecordset(SqlString, dbOpenSnapshot)
    Else
        Err.Raise 3075, "DbSql.Union", "UNION 语句语法错误"
    End If
End Function

' 联合查询（有重复）
Public Function UnionAll(ParamArray SqlBuilders() As Variant) As DAO.Recordset
    Dim Length As Long
    Length = UBound(SqlBuilders) - LBound(SqlBuilders) + 1

    If Length < 2 Then
        Err.Raise 3075, "DbSql.UnionAll", "至少需要两个查询进行 Union"
        Exit Function
    End If

    Dim SqlString As String
    Dim i As Long
    Dim TempSql As String

    TempSql = SqlBuilders(LBound(SqlBuilders)).ToSqlString(0, False)

    SqlString = TempSql
    Length = UBound(SqlBuilders)
    For i = (LBound(SqlBuilders) + 1) To Length
        If VBA.IsObject(SqlBuilders(i)) And TypeOf SqlBuilders(i) Is SqlBuilder Then
            TempSql = SqlBuilders(i).ToSqlString(0, False)
            SqlString = SqlString & " UNION All " & TempSql
        Else
            Err.Raise 3001, "DbSql.UnionAll", "参数必须是 SqlBuilder 对象类型"
            Exit Function
        End If
    Next i

    If VBA.Len(SqlString) > 0 Then
        SqlString = SqlString & ";"
        Call ConnectDatabase
        Set UnionAll = CurrentDatabase.OpenRecordset(SqlString, dbOpenSnapshot)
    Else
        Err.Raise 3075, "DbSql.Union", "UNION All 语句语法错误"
    End If
End Function
