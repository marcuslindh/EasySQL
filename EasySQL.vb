Imports System.Data.SqlClient
Imports System.Text

Public Class SQLBuilder

    Private Property _Type As QueryType
    Private Property _Columns As New List(Of SQLBuilderColumn)
    Private Property _TableName As String = ""
    Private Property _Database As SqlConnection
    Private Property _Command As SqlCommand
    Private Property _OrderByColumns As String = ""
    Private Property _OrderBy_Order As OrderByOrder = OrderByOrder.ascending


    Private Property SecureParameterNames As String() = "A B C D E F G H I J K L M N O P Q R S T U V W X Y Z AA BB CC DD EE FF GG HH II JJ KK LL MM NN OO PP QQ RR SS TT UU VV WW XX YY ZZ AB AC AD AE AF AG AH AI AJ AK AL AM AN AO AP AQ AR AS AT AU AV AW AX AY AZ".Split(" ")
    Private Property SecureParameterNamesLocation As Integer = 0

    Private Property _Test As Boolean = False

    Public Enum QueryType
        [Select] = 1
        Insert = 2
        Delete = 3
        Update = 4
    End Enum
    Public Enum QueryOperation
        Equal = 1
        GreaterThan = 2
        GreaterThanOrEqualTo = 3
        LessThan = 4
        LessThanOrEqualTo = 5
        NotEqualTo = 6
        [Like] = 7
        [IN] = 8
    End Enum
    Public Enum QueryLogicalOperator
        [And] = 1
        [OR] = 2
        GroupBy = 3
    End Enum
    Public Enum OrderByOrder
        ascending = 1
        descending = 2
    End Enum

    Private Sub InheritParameters(ByRef sql As SQLBuilder)
        sql._Type = Me._Type
        sql._Columns = Me._Columns
        sql._TableName = Me._TableName
        sql._Database = Me._Database
        sql._Command = Me._Command
        sql._Test = Me._Test
        sql._OrderBy_Order = Me._OrderBy_Order
        sql._OrderByColumns = Me._OrderByColumns
    End Sub


    Public Function [Select]() As SQLBuilder
        Dim sql As New SQLBuilder
        InheritParameters(sql)
        sql._Type = QueryType.Select
        'sql._Database = Me._Database
        'sql._Command = Me._Command

        Return sql
    End Function
    Public Function [Select](Columns As String) As SQLBuilder
        Dim sql As New SQLBuilder
        InheritParameters(sql)
        sql._Type = QueryType.Select
        'sql._Database = Me._Database
        'sql._Command = Me._Command

        If Columns.Contains(",") Then
            Dim s() As String = Columns.Split(",")
            For Each item In s
                sql._Columns.Add(New SQLBuilderColumn With {.ColumnName = item.Trim, .NoType = True, .NoValue = True})
            Next
        Else
            sql._Columns.Add(New SQLBuilderColumn With {.ColumnName = Columns, .NoType = True, .NoValue = True})
        End If

        Return sql
    End Function
    Public Function Insert() As SQLBuilder
        Dim sql As New SQLBuilder
        InheritParameters(sql)
        sql._Type = QueryType.Insert
        'sql._Database = Me._Database
        'sql._Command = Me._Command

        Return sql
    End Function
    Public Function Delete() As SQLBuilder
        Dim sql As New SQLBuilder
        InheritParameters(sql)
        sql._Type = QueryType.Delete
        'sql._Database = Me._Database
        'sql._Command = Me._Command


        Return sql
    End Function
    Public Function Update() As SQLBuilder
        Dim sql As New SQLBuilder
        InheritParameters(sql)
        sql._Type = QueryType.Update
        'sql._Database = Me._Database
        'sql._Command = Me._Command


        Return sql
    End Function

    Public Function Testing() As SQLBuilder
        Dim sql As New SQLBuilder
        InheritParameters(sql)
        'sql._Type = Me._Type
        'sql._Columns = Me._Columns
        'sql._TableName = Me._TableName
        'sql._Database = Me._Database
        'sql._Command = Me._Command
        sql._Test = True
        Return sql
    End Function

    Public Function From(Name As String) As SQLBuilder
        Dim sql As New SQLBuilder
        InheritParameters(sql)
        'sql._Type = Me._Type
        'sql._Columns = Me._Columns
        sql._TableName = Name
        'sql._Database = Me._Database
        'sql._Command = Me._Command
        'sql._Test = Me._Test

        Return sql
    End Function

    Public Function Where(ColumnName As String, Type As SqlDbType, Operation As QueryOperation, Value As Object, LogicalOperator As QueryLogicalOperator) As SQLBuilder
        Dim sql As New SQLBuilder
        InheritParameters(sql)
        'sql._Type = Me._Type
        'sql._Columns = Me._Columns
        'sql._TableName = Me._TableName
        'sql._Database = Me._Database
        'sql._Command = Me._Command
        'sql._Test = Me._Test

        sql._Columns.Add(New SQLBuilderColumn With {.ColumnName = ColumnName, .Type = Type, .Value = Value, .isWhere = True, .Operation = Operation, .LogicalOperator = LogicalOperator})

        Return sql
    End Function
    Public Function Where(ColumnName As String, Type As SqlDbType, Operation As QueryOperation, Value As Object) As SQLBuilder
        Dim sql As New SQLBuilder
        InheritParameters(sql)
        'sql._Type = Me._Type
        'sql._Columns = Me._Columns
        'sql._TableName = Me._TableName
        'sql._Database = Me._Database
        'sql._Command = Me._Command
        'sql._Test = Me._Test

        sql._Columns.Add(New SQLBuilderColumn With {.ColumnName = ColumnName, .Type = Type, .Value = Value, .isWhere = True, .Operation = Operation})

        Return sql
    End Function

    Public Function Column(ColumnName As String, Type As SqlDbType, Value As Object) As SQLBuilder
        Dim sql As New SQLBuilder
        InheritParameters(sql)
        'sql._Type = Me._Type
        'sql._Columns = Me._Columns
        'sql._TableName = Me._TableName
        'sql._Database = Me._Database
        'sql._Command = Me._Command
        'sql._Test = Me._Test

        sql._Columns.Add(New SQLBuilderColumn With {.ColumnName = ColumnName, .Type = Type, .Value = Value})

        Return sql
    End Function
    Public Function Column(ColumnName As String, Value As Object) As SQLBuilder
        Dim sql As New SQLBuilder
        InheritParameters(sql)
        'sql._Type = Me._Type
        'sql._Columns = Me._Columns
        'sql._TableName = Me._TableName
        'sql._Database = Me._Database
        'sql._Command = Me._Command
        'sql._Test = Me._Test

        sql._Columns.Add(New SQLBuilderColumn With {.ColumnName = ColumnName, .NoValue = True})

        Return sql
    End Function
    Public Function Column(ColumnName As String) As SQLBuilder
        Dim sql As New SQLBuilder
        InheritParameters(sql)
        'sql._Type = Me._Type
        'sql._Columns = Me._Columns
        'sql._TableName = Me._TableName
        'sql._Database = Me._Database
        'sql._Command = Me._Command
        'sql._Test = Me._Test
        sql._Columns.Add(New SQLBuilderColumn With {.ColumnName = ColumnName, .NoType = True, .NoValue = True})

        Return sql
    End Function

    Public Function GroupBy(ColumnName As String) As SQLBuilder
        Dim sql As New SQLBuilder
        InheritParameters(sql)
        'sql._Type = Me._Type
        'sql._Columns = Me._Columns
        'sql._TableName = Me._TableName
        'sql._Database = Me._Database
        'sql._Command = Me._Command
        'sql._Test = Me._Test

        sql._Columns.Add(New SQLBuilderColumn With {.ColumnName = ColumnName, .NoType = True, .NoValue = True, .LogicalOperator = QueryLogicalOperator.GroupBy})

        Return sql
    End Function

    Public Function OrderBy(Collumns As String, Optional ByVal Order As OrderByOrder = OrderByOrder.ascending) As SQLBuilder
        Dim sql As New SQLBuilder
        InheritParameters(sql)
        sql._OrderByColumns = Collumns
        sql._OrderBy_Order = Order

        Return sql
    End Function

    Public Function Database(db As SqlConnection) As SQLBuilder
        Dim sql As New SQLBuilder
        InheritParameters(sql)
        'sql._Type = Me._Type
        'sql._Columns = Me._Columns
        'sql._TableName = Me._TableName
        'sql._Database = db
        'sql._Command = Me._Command
        'sql._Test = Me._Test

        Return sql
    End Function

    Private Function GetSecureName() As String
        SecureParameterNamesLocation += 1
        Return SecureParameterNames(SecureParameterNamesLocation - 1)
    End Function

    Private Function CheckForSyntaxErrors() As String

        If _TableName = "" Then
            Return "No table added (.From(""tabelname""))"
        End If

        If _Test = False Then
            Try
                If Not _Database.State = ConnectionState.Open Then
                    Return "No Database added (.Database(DB))"
                End If
            Catch ex As Exception
                Return "No Database added (.Database(DB))"
            End Try
        End If

        Return ""
    End Function

    Private Sub Generate_GetSQL_Where(ByRef sql As StringBuilder)
        Dim isFirst = True
        For Each col In _Columns
            If col.isWhere Then
                If isFirst Then
                    isFirst = False
                    sql.Append(" WHERE ")
                    sql.Append(col.ColumnName)

                    Select Case col.Operation
                        Case QueryOperation.Equal
                            sql.Append(" = ")
                        Case QueryOperation.GreaterThan
                            sql.Append(" > ")
                        Case QueryOperation.GreaterThanOrEqualTo
                            sql.Append(" >= ")
                        Case QueryOperation.LessThan
                            sql.Append(" < ")
                        Case QueryOperation.LessThanOrEqualTo
                            sql.Append(" <= ")
                        Case QueryOperation.Like
                            sql.Append(" LIKE ")
                        Case QueryOperation.IN
                            sql.Append(" IN (")
                    End Select

                    If col.Operation = QueryOperation.IN Then
                        Dim _First As Boolean = True
                        For Each perm In col.ParameterNames
                            If _First Then
                                sql.Append("@" & perm)
                                _First = False
                            Else
                                sql.Append(", @" & perm)
                            End If
                        Next
                        sql.Append(")")

                    Else
                        sql.Append("@" & col.ParameterName)
                    End If



                Else
                    Select Case col.LogicalOperator
                        Case QueryLogicalOperator.And
                            sql.Append(" AND ")
                        Case QueryLogicalOperator.OR
                            sql.Append(" OR ")
                        Case Else
                            sql.Append(" AND ")
                    End Select


                    sql.Append(col.ColumnName)

                    Select Case col.Operation
                        Case QueryOperation.Equal
                            sql.Append(" = ")
                        Case QueryOperation.GreaterThan
                            sql.Append(" > ")
                        Case QueryOperation.GreaterThanOrEqualTo
                            sql.Append(" >= ")
                        Case QueryOperation.LessThan
                            sql.Append(" < ")
                        Case QueryOperation.LessThanOrEqualTo
                            sql.Append(" <= ")
                        Case QueryOperation.Like
                            sql.Append(" LIKE ")
                        Case QueryOperation.IN
                            sql.Append(" IN (")
                    End Select

                    If col.Operation = QueryOperation.IN Then
                        Dim _First As Boolean = True
                        For Each perm In col.ParameterNames
                            If _First Then
                                sql.Append("@" & perm)
                                _First = False
                            Else
                                sql.Append(", @" & perm)
                            End If
                        Next
                        sql.Append(")")

                    Else
                        sql.Append("@" & col.ParameterName)
                    End If
                End If
            End If
        Next
    End Sub

    Private Sub Generate_GetSQL_SELECT(ByRef sql As StringBuilder)
        sql.Append("SELECT ")

        Dim isFirst As Boolean = True
        For Each col In _Columns
            If col.isWhere = False Then
                If isFirst Then
                    sql.Append(col.ColumnName)
                    isFirst = False
                Else
                    If Not col.LogicalOperator = QueryLogicalOperator.GroupBy Then
                        sql.Append(", " & col.ColumnName)
                    End If
                End If

            End If
        Next

        sql.Append(" FROM ")
        sql.Append(_TableName)


        Generate_GetSQL_Where(sql)

        isFirst = True

        For Each col In _Columns
            If col.LogicalOperator = QueryLogicalOperator.GroupBy Then
                If isFirst Then
                    isFirst = False
                    sql.Append(" GROUP BY ")
                End If

                sql.Append(col.ColumnName)
            End If
        Next

        If Not _OrderByColumns = "" Then
            sql.Append(" ORDER BY ")
            sql.Append(_OrderByColumns)

            Select Case _OrderBy_Order
                Case OrderByOrder.ascending
                    sql.Append(" ASC")
                Case OrderByOrder.descending
                    sql.Append(" DESC")
            End Select

        End If
    End Sub

    Private Sub Generate_GetSQL_INSERT(ByRef sql As StringBuilder)
        sql.Append("INSERT INTO ")
        sql.Append(_TableName & " (")

        Dim isFirst As Boolean = True
        For Each col In _Columns
            If col.isWhere = False Then
                If isFirst Then
                    sql.Append(col.ColumnName)
                    isFirst = False
                Else
                    sql.Append(", " & col.ColumnName)
                End If
            End If
        Next

        sql.Append(") VALUES (")

        isFirst = True
        For Each col In _Columns
            If col.isWhere = False Then
                If isFirst Then
                    sql.Append("@" & col.ParameterName)
                    isFirst = False
                Else
                    sql.Append(", @" & col.ParameterName)
                End If
            End If
        Next

        sql.Append(")")

    End Sub
    Private Sub Generate_GetSQL_DELETE(ByRef sql As StringBuilder)
        sql.Append("DELETE ")

        sql.Append(" FROM ")
        sql.Append(_TableName)

        Generate_GetSQL_Where(sql)
    End Sub

    Private Sub Generate_GetSQL_UPDATE(ByRef sql As StringBuilder)
        sql.Append("UPDATE ")
        sql.Append(_TableName)
        sql.Append(" SET ")

        Dim isFirst As Boolean = True
        For Each col In _Columns
            If col.isWhere = False Then
                If isFirst Then
                    sql.Append(col.ColumnName)
                    sql.Append(" = ")
                    sql.Append("@" & col.ParameterName)
                    isFirst = False
                Else
                    sql.Append(", ")
                    sql.Append(col.ColumnName)
                    sql.Append(" = ")
                    sql.Append("@" & col.ParameterName)
                End If
            End If
        Next

        Generate_GetSQL_Where(sql)
    End Sub



    Public ReadOnly Property GetSQL As String
        Get
            Dim sql As New StringBuilder

            Dim SyntaxErrors As String = CheckForSyntaxErrors()
            If Not SyntaxErrors = "" Then
                Throw New Exception(SyntaxErrors)
            End If


            For Each col In _Columns
                If Not col.NoType = True Then

                    If col.Operation = QueryOperation.IN Then
                        Dim items
                        Try
                            items = CType(col.Value, List(Of Object))


                            For i As Integer = 0 To items.Count - 1
                                col.ParameterNames.Add(GetSecureName())
                            Next
                        Catch ex As Exception
                            Try
                                items = CType(col.Value, Integer())

                                For i As Integer = 0 To items.Length - 1
                                    col.ParameterNames.Add(GetSecureName())
                                Next
                            Catch ex2 As Exception
                                Try
                                    items = CType(col.Value, String())

                                    For i As Integer = 0 To items.Length - 1
                                        col.ParameterNames.Add(GetSecureName())
                                    Next
                                Catch ex3 As Exception
                                    Throw New Exception("EasySQL: you can only use types List(Of Object), Integer(), String() (QueryOperation.IN)")
                                End Try
                            End Try
                        End Try


                    Else
                        col.ParameterName = GetSecureName()
                    End If



                End If
            Next

            Select Case _Type
                Case QueryType.Select
                    Generate_GetSQL_SELECT(sql)
                Case QueryType.Insert
                    Generate_GetSQL_INSERT(sql)
                Case QueryType.Delete
                    Generate_GetSQL_DELETE(sql)
                Case QueryType.Update
                    Generate_GetSQL_UPDATE(sql)
            End Select

            Return sql.ToString
        End Get
    End Property

    Public Function GetSqlCommand() As SqlCommand

        _Command = New SqlCommand(GetSQL, _Database)

        For Each col In _Columns
            If col.NoValue = False Then

                _Command.Parameters.Add(New SqlParameter With {.ParameterName = col.ParameterName, .SqlDbType = col.Type, .Value = col.Value})

            End If
        Next



        Return _Command
    End Function


End Class

Public Class SQLBuilderColumn
    Public Property ColumnName As String = ""
    Public Property Type As SqlDbType
    Public Property Value As Object
    Public Property NoType As Boolean = False
    Public Property NoValue As Boolean = False
    Public Property ParameterName As String = ""
    Public Property ParameterNames As New List(Of String)
    Public Property isWhere As Boolean = False
    Public Property Operation As SQLBuilder.QueryOperation
    Public Property LogicalOperator As SQLBuilder.QueryLogicalOperator


End Class