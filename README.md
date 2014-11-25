EasySQL
=======

Create sql querys Easy in .net for System.Data.SqlClient

##Examples

###Select

```vb.net
Dim db As SqlConnection()
dim sql as new SQLBuilder

sql.Select("*").From("tableName").Database(db).GetSQL()
```
will return

```sql
SELECT * FROM tableName
```

###Select Where

```vb.net
Dim db As SqlConnection()
dim sql as new SQLBuilder

sql.Select("*").From("tableName").Where("ID", SqlDbType.Int, SQLBuilder.QueryOperation.Equal, 1).Database(db).GetSQL()
```
will return and use SqlParameter for fields

```sql
SELECT * FROM tableName WHERE ID = @A
```

###Get SqlCommand
```vb.net
Dim db As SqlConnection()
Dim sql As New SQLBuilder
    
Using com = sql.Select("*").From("tableName"). _
    Where("ID", SqlDbType.Int, SQLBuilder.QueryOperation.Equal, 1). _
    Database(db).GetSqlCommand
        
    Using dr As SqlDataReader = com.ExecuteReader
        While dr.Read
            Trace.WriteLine(dr("Name"))
            Trace.WriteLine(dr("Lastname"))
        End While
    End Using
End Using
```