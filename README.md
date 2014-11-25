EasySQL
=======

Create sql querys Easy in .net for System.Data.SqlClient

##Examples

###Select

```vb.net
dim sql as new SQLBuilder

sql.Select("*").From("tableName").Database(db).GetSQL()
```
will return

```sql
SELECT * FROM tableName
```

###Select Where

```vb.net
dim sql as new SQLBuilder

sql.Select("*").From("tableName").Where("ID", SqlDbType.Int, SQLBuilder.QueryOperation.Equal, 1).Database(db).GetSQL()
```
will return and use SqlParameter for fields

```sql
SELECT * FROM tableName WHERE ID = @A
```