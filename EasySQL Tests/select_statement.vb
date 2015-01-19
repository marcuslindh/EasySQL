Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports EasySQL


<TestClass()> Public Class select_statement

    'Test Table (Name: TestTable)
    '| ID | Name | LastName | Year |



    <TestMethod()> Public Sub select_statement()
        Dim sql As New SQLBuilder
        Assert.AreEqual(sql.Select("ID, Name, LastName, Year").From("TestTable").GetSQL, "SELECT ID, Name, LastName, Year FROM TestTable")
    End Sub

    <TestMethod()> Public Sub select_statement_AllColumns()
        Dim sql As New SQLBuilder
        Assert.AreEqual(sql.Select("*").From("TestTable").GetSQL, "SELECT * FROM TestTable")
    End Sub

    <TestMethod()> Public Sub select_statement_Columns()
        Dim sql As New SQLBuilder
        Assert.AreEqual(sql.Select.Column("ID").Column("Name").Column("LastName").Column("Year").From("TestTable").GetSQL, "SELECT ID, Name, LastName, Year FROM TestTable")
    End Sub

    <TestMethod()> Public Sub select_statement_Columns_AllColumns()
        Dim sql As New SQLBuilder
        Assert.AreEqual(sql.Select.Column("*").From("TestTable").GetSQL, "SELECT * FROM TestTable")
    End Sub


    <TestMethod()> Public Sub select_statement_Where_Equal()
        Dim sql As New SQLBuilder
        Assert.AreEqual(sql.Select.Column("*").From("TestTable").Where("ID", SqlDbType.Int, SQLBuilder.QueryOperation.Equal, 1).GetSQL, "SELECT * FROM TestTable WHERE ID = @A")
    End Sub

    <TestMethod()> Public Sub select_statement_Where_GreaterThan()
        Dim sql As New SQLBuilder
        Assert.AreEqual(sql.Select.Column("*").From("TestTable").Where("ID", SqlDbType.Int, SQLBuilder.QueryOperation.GreaterThan, 1).GetSQL, "SELECT * FROM TestTable WHERE ID > @A")
    End Sub

    <TestMethod()> Public Sub select_statement_Where_GreaterThanOrEqualTo()
        Dim sql As New SQLBuilder
        Assert.AreEqual(sql.Select.Column("*").From("TestTable").Where("ID", SqlDbType.Int, SQLBuilder.QueryOperation.GreaterThanOrEqualTo, 1).GetSQL, "SELECT * FROM TestTable WHERE ID >= @A")
    End Sub

    <TestMethod()> Public Sub select_statement_Where_LessThan()
        Dim sql As New SQLBuilder
        Assert.AreEqual(sql.Select.Column("*").From("TestTable").Where("ID", SqlDbType.Int, SQLBuilder.QueryOperation.LessThan, 1).GetSQL, "SELECT * FROM TestTable WHERE ID < @A")
    End Sub

    <TestMethod()> Public Sub select_statement_Where_LessThanOrEqualTo()
        Dim sql As New SQLBuilder
        Assert.AreEqual(sql.Select.Column("*").From("TestTable").Where("ID", SqlDbType.Int, SQLBuilder.QueryOperation.LessThanOrEqualTo, 1).GetSQL, "SELECT * FROM TestTable WHERE ID <= @A")
    End Sub

    <TestMethod()> Public Sub select_statement_Where_Like()
        Dim sql As New SQLBuilder
        Assert.AreEqual(sql.Select.Column("*").From("TestTable").Where("Name", SqlDbType.NVarChar, SQLBuilder.QueryOperation.Like, "Ma%").GetSQL, "SELECT * FROM TestTable WHERE Name LIKE @A")
    End Sub


    <TestMethod()> Public Sub select_statement_GroupBy()
        Dim sql As New SQLBuilder
        Assert.AreEqual(sql.Select.Column("*").From("TestTable").GroupBy("Name").GetSQL, "SELECT * FROM TestTable GROUP BY Name")
    End Sub

    <TestMethod()> Public Sub select_statement_Where_GroupBy()
        Dim sql As New SQLBuilder
        Assert.AreEqual(sql.Select.Column("*").From("TestTable").Where("ID", SqlDbType.NVarChar, SQLBuilder.QueryOperation.GreaterThan, 1).GroupBy("Name").GetSQL, "SELECT * FROM TestTable WHERE ID > @A GROUP BY Name")
    End Sub

    <TestMethod()> Public Sub select_statement_COUNT_Where_GroupBy()
        Dim sql As New SQLBuilder
        Assert.AreEqual(sql.Select.Column("Year").Column("COUNT(*) TotalCount").From("TestTable").Where("ID", SqlDbType.Int, SQLBuilder.QueryOperation.GreaterThan, 1).GroupBy("Year").GetSQL, "SELECT Year, COUNT(*) TotalCount FROM TestTable WHERE ID > @A GROUP BY Year")
    End Sub



End Class