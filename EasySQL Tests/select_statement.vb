Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports EasySQL


<TestClass()> Public Class select_statement

    'Test Table (Name: TestTable)
    '| ID | Name | LastName | Year |



    <TestMethod()> Public Sub select_statement()
        Dim sql As New SQLBuilder
        Assert.AreEqual(sql.Select("ID, Name, LastName, Year").From("TestTable").Testing().GetSQL, "SELECT ID, Name, LastName, Year FROM TestTable")
    End Sub

    <TestMethod()> Public Sub select_statement_AllColumns()
        Dim sql As New SQLBuilder
        Assert.AreEqual(sql.Select("*").From("TestTable").Testing().GetSQL, "SELECT * FROM TestTable")
    End Sub

    <TestMethod()> Public Sub select_statement_Columns()
        Dim sql As New SQLBuilder
        Assert.AreEqual(sql.Select.Column("ID").Column("Name").Column("LastName").Column("Year").From("TestTable").Testing().GetSQL, "SELECT ID, Name, LastName, Year FROM TestTable")
    End Sub

    <TestMethod()> Public Sub select_statement_Columns_AllColumns()
        Dim sql As New SQLBuilder
        Assert.AreEqual(sql.Select.Column("*").From("TestTable").Testing().GetSQL, "SELECT * FROM TestTable")
    End Sub

    <TestMethod()> Public Sub select_statement_Where_Equal()
        Dim sql As New SQLBuilder
        Assert.AreEqual(sql.Select.Column("*").From("TestTable").Where("ID", SqlDbType.Int, SQLBuilder.QueryOperation.Equal, 1).Testing().GetSQL, "SELECT * FROM TestTable WHERE ID = @A")
    End Sub

    <TestMethod()> Public Sub select_statement_Where_GreaterThan()
        Dim sql As New SQLBuilder
        Assert.AreEqual(sql.Select.Column("*").From("TestTable").Where("ID", SqlDbType.Int, SQLBuilder.QueryOperation.GreaterThan, 1).Testing().GetSQL, "SELECT * FROM TestTable WHERE ID > @A")
    End Sub

    <TestMethod()> Public Sub select_statement_Where_GreaterThanOrEqualTo()
        Dim sql As New SQLBuilder
        Assert.AreEqual(sql.Select.Column("*").From("TestTable").Where("ID", SqlDbType.Int, SQLBuilder.QueryOperation.GreaterThanOrEqualTo, 1).Testing().GetSQL, "SELECT * FROM TestTable WHERE ID >= @A")
    End Sub

    <TestMethod()> Public Sub select_statement_Where_LessThan()
        Dim sql As New SQLBuilder
        Assert.AreEqual(sql.Select.Column("*").From("TestTable").Where("ID", SqlDbType.Int, SQLBuilder.QueryOperation.LessThan, 1).Testing().GetSQL, "SELECT * FROM TestTable WHERE ID < @A")
    End Sub

    <TestMethod()> Public Sub select_statement_Where_LessThanOrEqualTo()
        Dim sql As New SQLBuilder
        Assert.AreEqual(sql.Select.Column("*").From("TestTable").Where("ID", SqlDbType.Int, SQLBuilder.QueryOperation.LessThanOrEqualTo, 1).Testing().GetSQL, "SELECT * FROM TestTable WHERE ID <= @A")
    End Sub

    <TestMethod()> Public Sub select_statement_Where_Like()
        Dim sql As New SQLBuilder
        Assert.AreEqual(sql.Select.Column("*").From("TestTable").Where("Name", SqlDbType.NVarChar, SQLBuilder.QueryOperation.Like, "Ma%").Testing().GetSQL, "SELECT * FROM TestTable WHERE Name LIKE @A")
    End Sub

    <TestMethod()> Public Sub select_statement_Where_IN()
        Dim sql As New SQLBuilder
        Dim ids As Integer() = {1, 2, 3, 4}
        Assert.AreEqual(sql.Select.Column("*").From("TestTable").Where("ID", SqlDbType.NVarChar, SQLBuilder.QueryOperation.IN, ids).Testing().GetSQL, "SELECT * FROM TestTable WHERE ID IN (@A, @B, @C, @D)")
    End Sub

    <TestMethod()> Public Sub select_statement_Where_IN2()
        Dim sql As New SQLBuilder
        'Dim ids As Integer() = {1, 2, 3, 4}

        Dim names As String() = {"test1", "test2", "test3"}

        Assert.AreEqual(sql.Select.Column("*").From("TestTable").Where("ID", SqlDbType.Int, SQLBuilder.QueryOperation.Equal, 1).Where("Name", SqlDbType.NVarChar, SQLBuilder.QueryOperation.IN, names).Testing().GetSQL, "SELECT * FROM TestTable WHERE ID = @A AND Name IN (@B, @C, @D)")
        Dim com = sql.Select.Column("*").From("TestTable").Where("ID", SqlDbType.Int, SQLBuilder.QueryOperation.Equal, 1).Where("Name", SqlDbType.NVarChar, SQLBuilder.QueryOperation.IN, names).Testing().GetSqlCommand

        Assert.AreEqual(com.Parameters(0).ParameterName, "A")
        Assert.AreEqual(com.Parameters(1).ParameterName, "B")
        Assert.AreEqual(com.Parameters(2).ParameterName, "C")
        Assert.AreEqual(com.Parameters(3).ParameterName, "D")

    End Sub


    <TestMethod()> Public Sub select_statement_GroupBy()
        Dim sql As New SQLBuilder
        Assert.AreEqual(sql.Select.Column("*").From("TestTable").GroupBy("Name").Testing().GetSQL, "SELECT * FROM TestTable GROUP BY Name")
    End Sub

    <TestMethod()> Public Sub select_statement_Where_GroupBy()
        Dim sql As New SQLBuilder
        Assert.AreEqual(sql.Select.Column("*").From("TestTable").Where("ID", SqlDbType.NVarChar, SQLBuilder.QueryOperation.GreaterThan, 1).GroupBy("Name").Testing().GetSQL, "SELECT * FROM TestTable WHERE ID > @A GROUP BY Name")
    End Sub

    <TestMethod()> Public Sub select_statement_COUNT_Where_GroupBy()
        Dim sql As New SQLBuilder
        Assert.AreEqual(sql.Select.Column("Year").Column("COUNT(*) TotalCount").From("TestTable").Where("ID", SqlDbType.Int, SQLBuilder.QueryOperation.GreaterThan, 1).GroupBy("Year").Testing().GetSQL, "SELECT Year, COUNT(*) TotalCount FROM TestTable WHERE ID > @A GROUP BY Year")
    End Sub

    <TestMethod()> Public Sub select_statement_OrderBy_Default()
        Dim sql As New SQLBuilder
        Assert.AreEqual(sql.Select("*").From("TestTable").OrderBy("Name").Testing.GetSQL, "SELECT * FROM TestTable ORDER BY Name ASC")
    End Sub

    <TestMethod()> Public Sub select_statement_OrderBy_descending()
        Dim sql As New SQLBuilder
        Assert.AreEqual(sql.Select("*").From("TestTable").OrderBy("Name", SQLBuilder.OrderByOrder.descending).Testing.GetSQL, "SELECT * FROM TestTable ORDER BY Name DESC")
    End Sub


    <TestMethod()> Public Sub Update_statement_Where_Dub_ColumnName()
        Dim sql As New SQLBuilder
        Assert.AreEqual(sql.Update.Column("PageID", SqlDbType.Int, 1).From("Views").Where("PageID", SqlDbType.Int, SQLBuilder.QueryOperation.Equal, 1).Testing().GetSQL, "UPDATE Views SET PageID = @A WHERE PageID = @B")
    End Sub


End Class