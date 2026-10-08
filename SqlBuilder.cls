VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Lang VBA

'============================================================================
' Access VBA SQL Builder
' Repository: https://github.com/chinbeker/AccessVBA-SqlBuilder
' License: MIT
' Author: huachen
' Version: 1.0.0
' Created: 2026
' Description: Fluent SQL builder and database helper for Access VBA
'============================================================================

'============================================================================
'                         SQL DDL 数据类型关键字
'----------------------------------------------------------------------------
' TEXT | CHAR                 短文本（长度）              TEXT(50)
' MEMO | LONGTEXT             长文本                      MEMO
' BYTE                        字节型（0-255）             BYTE
' INTEGER | SHORT             整型（-32768~32767）        INTEGER
' LONG | INT                  长整型                      LONG
' SINGLE | REAL               单精度浮点数                SINGLE
' DOUBLE | FLOAT              双精度浮点数                DOUBLE
' DECIMAL | NUMERIC           十进制数（精确小数）        DECIMAL(10,2)
' CURRENCY | MONEY            货币                        CURRENCY
' DATETIME | DATE             日期/时间                   DATETIME
' YESNO | BIT | BOOLEAN       是/否                       YESNO
' BINARY                      二进制（固定长度）          BINARY(50)
' OLEOBJECT | LONGBINARY      OLE 对象（图片/文件）       OLEOBJECT
' GUID                        全局唯一标识符              GUID
'============================================================================

'============================================================================
'                          Access SQL 标准写法
'----------------------------------------------------------------------------
' TEXT(n)                     短文本                      TEXT(50)
' MEMO                        长文本                      MEMO
' BYTE                        字节型                      BYTE
' INTEGER                     整型                        INTEGER
' LONG                        长整型                      LONG
' SINGLE                      单精度                      SINGLE
' DOUBLE                      双精度                      DOUBLE
' DECIMAL(p,s)                十进制                      DECIMAL(10,2)
' CURRENCY                    货币                        CURRENCY
' DATETIME                    日期/时间                   DATETIME
' YESNO                       是/否                       YESNO
' BINARY(n)                   二进制                      BINARY(50)
' OLEOBJECT                   OLE 对象                    OLEOBJECT
' GUID                        GUID                        GUID
' AUTOINCREMENT               自动编号                    AUTOINCREMENT
'============================================================================

Option Compare Database
Option Explicit

Private TopSqlString As String
Private SelectAllField As Boolean
Private IsDerivedTable As Boolean
Private IsOpenSnapshot As Boolean
Private DistinctSqlString As String
Private FromTableName As String
Private IntoTableName As String
Private DerivedTableName As String

Private FieldCollection As Collection
Private ValueCollection As Collection
Private JoinCollection As Collection
Private WhereCollection As Collection
Private GroupCollection As Collection
Private HavingCollection As Collection
Private OrderCollection As Collection
Private ParamDefCollection As Collection
Private ParamTypeCollection As Collection
Private ParamCollection As Collection



' 判断空字符串
Private Function IsEmptyString(ByVal str As Variant) As Boolean
    If VBA.VarType(str) = VBA.vbString Then
        IsEmptyString = (VBA.Len(str) = 0)
    Else
        IsEmptyString = True
    End If
End Function




'判断验证
Private Function HasField() As Boolean
    If SelectAllField Then
        HasField = True
    ElseIf FieldCollection Is Nothing Then
        HasField = False
    Else
        HasField = (FieldCollection.Count > 0)
    End If
End Function

Private Function Validation() As Boolean
    If FieldCollection Is Nothing Or ValueCollection Is Nothing Then
        Validation = False
    Else
        Validation = (ValueCollection.Count > 0 And FieldCollection.Count = ValueCollection.Count)
    End If
End Function

Private Function HasInto() As Boolean
    HasInto = Not IsEmptyString(IntoTableName)
End Function

Private Function HasSource() As Boolean
    HasSource = Not IsEmptyString(FromTableName)
End Function

Private Function HasWhere() As Boolean
    If WhereCollection Is Nothing Then
        HasWhere = False
    Else
        HasWhere = (WhereCollection.Count > 0)
    End If
End Function



'--------------------------属性--------------------------------
'是否使用参数
Public Property Get HasParam() As Boolean
    If ParamDefCollection Is Nothing Or ParamTypeCollection Is Nothing Or ParamCollection Is Nothing Then
        HasParam = False
    Else
        HasParam = (ParamDefCollection.Count > 0 And ParamDefCollection.Count = ParamTypeCollection.Count And ParamDefCollection.Count = ParamCollection.Count)
    End If
End Property

'是否作为派生表（子查询表）
Public Property Get HasDerived() As Boolean
    HasDerived = IsDerivedTable
End Property





' --------------------------公共方法--------------------------------
'指定记录数
Public Sub Top(ByVal Val As Long)
    If Val > 0 Then
        TopSqlString = "TOP " & Val & " "
    Else
        TopSqlString = ""
    End If
End Sub

'唯一值
Public Sub Distinct(Optional ByVal Val As Boolean = True)
    If Val = True Then
        DistinctSqlString = "DISTINCT "
    Else
        DistinctSqlString = ""
    End If
End Sub

'选择所有字段
Public Sub SelectAll(Optional ByVal Val As Boolean = True)
    SelectAllField = Val
End Sub

'选择所有字段
Public Function ReadOnly(Optional ByVal Val As Boolean = True) As Boolean
    If Not VBA.IsMissing(Val) Then IsOpenSnapshot = Val
    ReadOnly = IsOpenSnapshot
End Function




' 数据源
Public Sub From(ByRef FromTable As Variant, Optional ByVal Alias As String = "")
    If VBA.VarType(FromTable) = VBA.vbString And Not IsEmptyString(FromTable) Then
        If Not IsEmptyString(Alias) Then
            FromTableName = FromTable & " AS " & Alias
        Else
            FromTableName = FromTable
        End If
    ElseIf VBA.IsObject(FromTable) And TypeOf FromTable Is SqlBuilder Then
        If Not IsEmptyString(Alias) Then FromTable.Derived Alias
        If FromTable.HasDerived Then
            FromTableName = FromTable.ToSqlString
        Else
            Err.Raise 3075, "SqlBuilder.From", "子查询必须添加别名"
        End If
    Else
        Err.Raise 449, "SqlBuilder.From", "参数1无效，仅支持 String 和 SqlBuilder 类型"
    End If
End Sub

Public Sub Into(ByVal TableName As String)
    If Not IsEmptyString(TableName) Then IntoTableName = TableName
End Sub

Public Function Alias() As String
    Alias = DerivedTableName
End Function

'开启子查询的派生表，设置别名
Public Sub Derived(ByVal Alias As String)
    If Not IsEmptyString(Alias) Then
        IsDerivedTable = True
        DerivedTableName = Alias
    Else
        IsDerivedTable = False
        DerivedTableName = ""
    End If
End Sub


' 表连接
Public Sub Join(ByRef JoinTable As Variant, ByVal SourceTable As String, ByVal JoinField As String, Optional ByVal SourceField As String = "")
    If IsEmptyString(SourceTable) Then Exit Sub
    If IsEmptyString(JoinField) Then Exit Sub
    If IsEmptyString(SourceField) Then SourceField = JoinField
    If JoinCollection Is Nothing Then Set JoinCollection = New Collection
    If VBA.VarType(JoinTable) = VBA.vbString And Not IsEmptyString(JoinTable) Then
        JoinCollection.Add " INNER JOIN " & JoinTable & " ON " & SourceTable & "." & SourceField & "=" & JoinTable & "." & JoinField
    ElseIf VBA.IsObject(JoinTable) And TypeOf JoinTable Is SqlBuilder Then
        If JoinTable.HasDerived Then
            JoinCollection.Add " INNER JOIN " & JoinTable.ToSqlString & " ON " & SourceTable & "." & SourceField & "=" & JoinTable.Alias & "." & JoinField
        Else
            Err.Raise 3075, "SqlBuilder.Join", "子查询必须添加别名"
        End If
    Else
        Err.Raise 449, "SqlBuilder.Join", "参数1无效，仅支持 String 和 SqlBuilder 类型"
    End If
End Sub

' 左连接
Public Sub LeftJoin(ByRef JoinTable As Variant, ByVal SourceTable As String, ByVal JoinField As String, Optional ByVal SourceField As String = "")
    If IsEmptyString(SourceTable) Then Exit Sub
    If IsEmptyString(JoinField) Then Exit Sub
    If IsEmptyString(SourceField) Then SourceField = JoinField
    If JoinCollection Is Nothing Then Set JoinCollection = New Collection
    If VBA.VarType(JoinTable) = VBA.vbString And Not IsEmptyString(JoinTable) Then
        JoinCollection.Add " LEFT JOIN " & JoinTable & " ON " & SourceTable & "." & SourceField & "=" & JoinTable & "." & JoinField
    ElseIf VBA.IsObject(JoinTable) And TypeOf JoinTable Is SqlBuilder Then
        If JoinTable.HasDerived Then
            JoinCollection.Add " LEFT JOIN " & JoinTable.ToSqlString & " ON " & SourceTable & "." & SourceField & "=" & JoinTable.Alias & "." & JoinField
        Else
            Err.Raise 3075, "SqlBuilder.LeftJoin", "子查询必须添加别名"
        End If
    Else
        Err.Raise 449, "SqlBuilder.LeftJoin", "参数1无效，仅支持 String 和 SqlBuilder 类型"
    End If
End Sub

Private Function ToSourceSqlString(Optional ByVal IsUpdate As Boolean = False) As String
    If HasSource Then
        Dim SourceSqlString As String
        If Not JoinCollection Is Nothing Then
            Dim Length As Long
            Length = JoinCollection.Count
            If Length > 0 Then
                If Length = 1 Then
                    SourceSqlString = FromTableName & JoinCollection(1)
                Else
                    SourceSqlString = "(" & FromTableName & JoinCollection(1) & ")"
                    Dim i As Long
                    For i = 2 To Length
                        If i < Length Then
                            SourceSqlString = "(" & SourceSqlString & JoinCollection(i) & ")"
                        Else
                            SourceSqlString = SourceSqlString & JoinCollection(i)
                        End If
                    Next i
                End If
            Else
                SourceSqlString = FromTableName
            End If
        Else
            SourceSqlString = FromTableName
        End If
        If IsUpdate = False Then SourceSqlString = " FROM " & SourceSqlString
        ToSourceSqlString = SourceSqlString
    Else
        If IsUpdate Then
            Err.Raise 3144, "SqlBuilder.ToSourceSqlString", "UPDATE 语句语法错误"
        Else
            Err.Raise 3075, "SqlBuilder.ToSourceSqlString", "FROM 语句错误"
        End If
    End If
End Function


'字段
Public Sub Field(ByVal FieldName As String, Optional ByVal Value As Variant, Optional ByVal Alias As String)
    If IsEmptyString(FieldName) Then Exit Sub
    If FieldCollection Is Nothing Then Set FieldCollection = New Collection
    If Not VBA.IsMissing(Alias) And Not IsEmptyString(Alias) Then
        FieldCollection.Add FieldName & " AS " & Alias
    Else
        FieldCollection.Add FieldName
    End If
    If Not VBA.IsMissing(Value) Then
        If ValueCollection Is Nothing Then Set ValueCollection = New Collection
        Select Case VBA.VarType(Value)
            Case VBA.vbString
                If VBA.Left(Value, 4) = "[$$]" Then
                    ValueCollection.Add VBA.Mid(Value, 5)
                Else
                    ValueCollection.Add "'" & VBA.Replace(Value, "'", "''") & "'"
                End If
            Case VBA.vbDate
                ValueCollection.Add "#" & VBA.Format(Value, "yyyy-mm-dd") & "#"
            Case VBA.vbNull
                ValueCollection.Add "NULL"
            Case Else
                ValueCollection.Add Value
        End Select
    End If
End Sub

'常量列
Public Sub Constant(ByVal Alias As String, ByVal Value As Variant)
    If IsEmptyString(Alias) Then Exit Sub
    If VBA.IsEmpty(Value) Then Exit Sub
    If FieldCollection Is Nothing Then Set FieldCollection = New Collection
    Dim FieldName As String
    Select Case VBA.VarType(Value)
        Case VBA.vbString
            FieldName = "'" & VBA.Replace(Value, "'", "''") & "'"
        Case VBA.vbDate
            FieldName = "#" & VBA.Format(Value, "yyyy-mm-dd") & "#"
        Case VBA.vbNull
            FieldName = "NULL"
        Case Else
            FieldName = VBA.CStr(Value)
    End Select
    FieldCollection.Add FieldName & " AS " & Alias
End Sub

'计算列
Public Sub Computed(ByVal Alias As String, ByVal Expression As String)
    If IsEmptyString(Alias) Then Exit Sub
    If IsEmptyString(Expression) Then Exit Sub
    If FieldCollection Is Nothing Then Set FieldCollection = New Collection
    FieldCollection.Add Expression & " AS " & Alias
End Sub

Private Function ToFieldSqlString(Optional ByVal IsUpdate As Boolean = False) As String
    If HasField Then
        Dim FieldSqlString As String
        If SelectAllField And HasSource Then
            FieldSqlString = FromTableName & ".*"
        ElseIf Not FieldCollection Is Nothing Then
            If FieldCollection.Count > 0 Then
                Dim i As Long
                Dim Length As Long
                Length = FieldCollection.Count
                For i = 1 To Length
                    If i < Length Then
                        If IsUpdate = True And Validation Then
                            FieldSqlString = FieldSqlString & FieldCollection(i) & "=" & ValueCollection(i) & ", "
                        Else
                            FieldSqlString = FieldSqlString & FieldCollection(i) & ", "
                        End If
                    ElseIf IsUpdate = True And Validation Then
                        FieldSqlString = FieldSqlString & FieldCollection(i) & "=" & ValueCollection(i)
                    Else
                        FieldSqlString = FieldSqlString & FieldCollection(i)
                    End If
                Next i
            End If
        End If
        ToFieldSqlString = FieldSqlString
    Else
        Err.Raise 3085, "SqlBuilder.ToFieldSqlString", "查询字段未定义"
    End If
End Function

Private Function ToValueSqlString(Optional ByVal IsUpdate As Boolean = False) As String
    If Not ValueCollection Is Nothing And ValueCollection.Count > 0 Then
        Dim i As Long
        Dim Length As Long
        Dim ValueSqlString As String
        Length = ValueCollection.Count
        For i = 1 To Length
            If i < Length Then
                ValueSqlString = ValueSqlString & ValueCollection(i) & ", "
            Else
                ValueSqlString = ValueSqlString & ValueCollection(i)
            End If
        Next i
        ToValueSqlString = ValueSqlString
    Else
        Err.Raise 3085, "SqlBuilder.ToValueSqlString", "字段值未定义"
    End If
End Function

' 筛选
Public Sub Where(ByVal Condition As String, Optional ByVal Value As Variant)
    If IsEmptyString(Condition) Then Exit Sub
    If WhereCollection Is Nothing Then Set WhereCollection = New Collection
    If Not VBA.IsMissing(Value) Then
        Select Case VBA.VarType(Value)
            Case VBA.vbString
                If Left(Value, 4) = "[$$]" Then
                    WhereCollection.Add Condition & " " & Mid(Value, 5)
                Else
                    WhereCollection.Add Condition & " '" & Replace(Value, "'", "''") & "'"
                End If
            Case VBA.vbDate
                WhereCollection.Add Condition & " #" & Format(Value, "yyyy-mm-dd") & "#"
            Case VBA.vbNull
                WhereCollection.Add Condition & " IS NULL"
            Case Else
                WhereCollection.Add Condition & " " & Value
        End Select
    Else
        WhereCollection.Add Condition
    End If
End Sub

Private Function ToWhereSqlString() As String
    If Not WhereCollection Is Nothing Then
        If WhereCollection.Count > 0 Then
            Dim i As Long
            Dim Length As Long
            Dim WhereSqlString As String
            Length = WhereCollection.Count
            For i = 1 To Length
                If i < Length Then
                    WhereSqlString = WhereSqlString & "(" & WhereCollection(i) & ") And "
                Else
                    WhereSqlString = WhereSqlString & "(" & WhereCollection(i) & ")"
                End If
            Next i
            If Length > 1 Then WhereSqlString = "(" & WhereSqlString & ")"
            If Not IsEmptyString(WhereSqlString) Then WhereSqlString = " WHERE " & WhereSqlString
            ToWhereSqlString = WhereSqlString
        End If
    End If
End Function


' 参数
Public Sub Param(ByVal ParamName As String, ByVal Val As Variant, Optional ByVal ParamType As String)
    If IsEmptyString(ParamName) Then Exit Sub
    If ParamDefCollection Is Nothing Then Set ParamDefCollection = New Collection
    If ParamCollection Is Nothing Then Set ParamCollection = New Collection
    If ParamTypeCollection Is Nothing Then Set ParamTypeCollection = New Collection
    ParamName = VBA.Replace(ParamName, " ", "_")
    ParamName = VBA.Replace(ParamName, ".", "_")
    ParamName = VBA.Replace(ParamName, "!", "_")
    ParamDefCollection.Add "Param_" & ParamName
    ParamCollection.Add Val
    If VBA.IsMissing(ParamType) Or IsEmptyString(ParamType) Then
        Select Case VBA.VarType(Val)
            Case VBA.vbByte: ParamTypeCollection.Add "BYTE"
            Case VBA.vbInteger: ParamTypeCollection.Add "INTEGER"
            Case VBA.vbLong: ParamTypeCollection.Add "LONG"
            Case VBA.vbSingle: ParamTypeCollection.Add "SINGLE"
            Case VBA.vbDouble: ParamTypeCollection.Add "DOUBLE"
            Case VBA.vbDecimal: ParamTypeCollection.Add "DECIMAL"
            Case VBA.vbCurrency: ParamTypeCollection.Add "CURRENCY"
            Case VBA.vbDate: ParamTypeCollection.Add "DATETIME"
            Case VBA.vbBoolean: ParamTypeCollection.Add "YESNO"
            Case VBA.vbLongLong: ParamTypeCollection.Add "BIGINT"
            Case VBA.vbString
                If VBA.Len(Value) > 255 Then
                    ParamTypeCollection.Add "MEMO"
                Else
                    ParamTypeCollection.Add "Text(255)"
                End If
            Case Else: ParamTypeCollection.Add "Text(255)"
        End Select
    Else
        ParamTypeCollection.Add ParamType
    End If
End Sub

Public Sub SetQueryDef(ByRef Def As DAO.QueryDef)
    If Me.HasParam Then
        Dim i As Long
        Dim Length As Long
        Length = ParamDefCollection.Count
        For i = 1 To Length
            Def.Parameters(ParamDefCollection(i)) = ParamCollection(i)
        Next i
    End If
End Sub

Private Function ToParamDefSqlString() As String
    If Me.HasParam Then
        If Not ParamDefCollection Is Nothing Then
            If ParamDefCollection.Count > 0 Then
                Dim i As Long
                Dim Length As Long
                Dim ParamDefSqlString As String
                Length = ParamDefCollection.Count
                For i = 1 To Length
                    If i < Length Then
                        ParamDefSqlString = ParamDefSqlString & "[" & ParamDefCollection(i) & "] " & ParamTypeCollection(i) & ", "
                    Else
                        ParamDefSqlString = ParamDefSqlString & "[" & ParamDefCollection(i) & "] " & ParamTypeCollection(i)
                    End If
                Next i
                If Not IsEmptyString(ParamDefSqlString) Then ParamDefSqlString = "PARAMETERS " & ParamDefSqlString & "; "
                ToParamDefSqlString = ParamDefSqlString
            End If
        End If
    End If
End Function


' 分组
Public Sub Group(ByVal FieldName As String)
    If Not IsEmptyString(FieldName) Then
        If GroupCollection Is Nothing Then Set GroupCollection = New Collection
        GroupCollection.Add FieldName
    End If
End Sub

Private Function ToGroupSqlString() As String
    If Not GroupCollection Is Nothing Then
        If GroupCollection.Count > 0 Then
            Dim i As Long
            Dim Length As Long
            Dim GroupSqlString As String
            Length = GroupCollection.Count
            For i = 1 To Length
                If i < Length Then
                    GroupSqlString = GroupSqlString & GroupCollection(i) & ", "
                Else
                    GroupSqlString = GroupSqlString & GroupCollection(i)
                End If
            Next i
            If Not IsEmptyString(GroupSqlString) Then GroupSqlString = " GROUP BY " & GroupSqlString
            ToGroupSqlString = GroupSqlString
        End If
    End If
End Function

' 筛选
Public Sub Having(ByVal Condition As String, Optional ByVal Value As Variant)
    If IsEmptyString(Condition) Then Exit Sub
    If HavingCollection Is Nothing Then Set HavingCollection = New Collection
    If Not VBA.IsMissing(Value) Then
        Select Case VBA.VarType(Value)
            Case VBA.vbString
                If Left(Value, 4) = "[$$]" Then
                    HavingCollection.Add Condition & " " & Mid(Value, 5)
                Else
                    HavingCollection.Add Condition & " '" & Replace(Value, "'", "''") & "'"
                End If
            Case VBA.vbDate
                HavingCollection.Add Condition & " #" & Format(Value, "yyyy-mm-dd") & "#"
            Case VBA.vbNull
                HavingCollection.Add Condition & " IS NULL"
            Case Else
                HavingCollection.Add Condition & " " & Value
        End Select
    Else
        HavingCollection.Add Condition
    End If
End Sub

Private Function ToHavingSqlString() As String
    If Not HavingCollection Is Nothing Then
        If HavingCollection.Count > 0 Then
            Dim i As Long
            Dim Length As Long
            Dim HavingSqlString As String
            Length = HavingCollection.Count
            For i = 1 To Length
                If i < Length Then
                    HavingSqlString = HavingSqlString & "(" & HavingCollection(i) & ") And "
                Else
                    HavingSqlString = HavingSqlString & "(" & HavingCollection(i) & ")"
                End If
            Next i
            If Length > 1 Then HavingSqlString = "(" & HavingSqlString & ")"
            If Not IsEmptyString(HavingSqlString) Then HavingSqlString = " HAVING " & HavingSqlString
            ToHavingSqlString = HavingSqlString
        End If
    End If
End Function

'排序
Public Sub Order(ByVal FieldName As String, Optional ByVal Reverse As Boolean = False)
    If IsEmptyString(FieldName) Then Exit Sub
    If OrderCollection Is Nothing Then Set OrderCollection = New Collection
    If Reverse = True Then
        OrderCollection.Add FieldName & " DESC"
    Else
        OrderCollection.Add FieldName
    End If
End Sub

Private Function ToOrderSqlString() As String
    If Not OrderCollection Is Nothing Then
        If OrderCollection.Count > 0 Then
            Dim i As Long
            Dim Length As Long
            Dim OrderSqlString As String
            Length = OrderCollection.Count
            For i = 1 To Length
                If i < Length Then
                    OrderSqlString = OrderSqlString & OrderCollection(i) & ", "
                Else
                    OrderSqlString = OrderSqlString & OrderCollection(i)
                End If
            Next i
            If Not IsEmptyString(OrderSqlString) Then OrderSqlString = " ORDER BY " & OrderSqlString
            ToOrderSqlString = OrderSqlString
        End If
    End If
End Function


' 拼接sql字符串
Public Function ToSqlString(Optional ByVal Operation As Byte = 0, Optional ByVal Semicolon As Boolean = True) As String
    Dim EndSemicolon As String
    If Semicolon Then EndSemicolon = ";"
    Dim SqlString As String
    If IsDerivedTable Then
        If HasField And HasSource Then
            SqlString = "(" & ToParamDefSqlString & "SELECT " & DistinctSqlString & TopSqlString & ToFieldSqlString & ToSourceSqlString & ToWhereSqlString & ToGroupSqlString & ToHavingSqlString & ToOrderSqlString & ") AS " & DerivedTableName
        Else
            Err.Raise 3075, "SqlBuilder.ToSqlString", "SQL 语法错误"
            Exit Function
        End If
    Else
        Select Case Operation
            '新增
            Case 4
                If IsEmptyString(IntoTableName) Then
                    Err.Raise 3155, "SqlBuilder.ToSqlString", "INSERT INTO 语句语法错误"
                    Exit Function
                End If
                If Validation Then
                    If HasSource Then
                        SqlString = ToParamDefSqlString & "INSERT INTO " & IntoTableName & " (" & ToFieldSqlString & ") SELECT " & ToValueSqlString & ToSourceSqlString & ToWhereSqlString & EndSemicolon
                    Else
                        SqlString = ToParamDefSqlString & "INSERT INTO " & IntoTableName & " (" & ToFieldSqlString & ") VALUES (" & ToValueSqlString & ")" & EndSemicolon
                    End If
                Else
                    Err.Raise 3155, "SqlBuilder.ToSqlString", "INSERT 语句语法错误"
                    Exit Function
                End If
            '更改
            Case 3
                If HasSource Then
                    If Validation Then
                        SqlString = ToParamDefSqlString & "UPDATE " & ToSourceSqlString(True) & " SET " & ToFieldSqlString(True) & ToWhereSqlString & EndSemicolon
                    Else
                        Err.Raise 3144, "SqlBuilder.ToSqlString", "UPDATE 字段未定义"
                        Exit Function
                    End If
                Else
                    Err.Raise 3144, "SqlBuilder.ToSqlString", "UPDATE 语句语法错误"
                    Exit Function
                End If
            '删除
            Case 2
                If HasSource Then
                    If HasWhere Then
                        SqlString = ToParamDefSqlString & "DELETE" & ToSourceSqlString & ToWhereSqlString & EndSemicolon
                    Else
                        Err.Raise 3075, "SqlBuilder.ToSqlString", "DELETE 必须定义 WHERE 语句"
                        Exit Function
                    End If
                Else
                    Err.Raise 3075, "SqlBuilder.ToSqlString", "FROM 语句错误"
                    Exit Function
                End If
            '计数
            Case 1
                If HasSource Then
                    SqlString = ToParamDefSqlString & "SELECT COUNT(*)" & ToSourceSqlString & ToWhereSqlString & EndSemicolon
                Else
                    Err.Raise 3075, "SqlBuilder.ToSqlString", "FROM 语句错误"
                    Exit Function
                End If
            '查询
            Case 0
                If HasField And HasSource Then
                    SqlString = ToParamDefSqlString & "SELECT " & DistinctSqlString & TopSqlString & ToFieldSqlString & ToSourceSqlString & ToWhereSqlString & ToGroupSqlString & ToHavingSqlString & ToOrderSqlString & EndSemicolon
                Else
                    Err.Raise 3075, "SqlBuilder.ToSqlString", "SQL 语句语法错误"
                    Exit Function
                End If
        End Select
    End If

    'MsgBox SqlString
    If VBA.Len(SqlString) > 0 Then
        ToSqlString = SqlString
    Else
        Err.Raise 3075, "SqlBuilder.ToSqlString", "SQL 语句语法错误"
    End If
End Function

'复用重置对象
Public Sub Reset()
    TopSqlString = ""
    SelectAllField = False
    IsDerivedTable = False
    IsOpenSnapshot = False
    DistinctSqlString = ""
    FromTableName = ""
    IntoTableName = ""
    DerivedTableName = ""

    Set FieldCollection = Nothing
    Set ValueCollection = Nothing
    Set JoinCollection = Nothing
    Set WhereCollection = Nothing
    Set GroupCollection = Nothing
    Set HavingCollection = Nothing
    Set OrderCollection = Nothing
    Set ParamDefCollection = Nothing
    Set ParamTypeCollection = Nothing
    Set ParamCollection = Nothing
End Sub

Private Sub Class_Terminate()
    Call Reset
End Sub
