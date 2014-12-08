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

End Class