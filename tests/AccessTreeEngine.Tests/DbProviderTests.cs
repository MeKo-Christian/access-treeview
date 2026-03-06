using NUnit.Framework;
using MeKo.TreeEngine;
using System;

namespace TreeEngine64.Tests;

[TestFixture]
public class DbProviderTests
{
    [Test]
    public void Constructor_Stores_ConnectionString()
    {
        var provider = new DbProvider(
            connectionString: "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=test.accdb",
            providerName: "System.Data.OleDb",
            tableName: "tblTree",
            idColumn: "NodeID",
            parentIdColumn: "ParentID",
            captionColumn: "NodeText"
        );

        Assert.That(provider.ConnectionString,
            Is.EqualTo("Provider=Microsoft.ACE.OLEDB.16.0;Data Source=test.accdb"));
    }

    [Test]
    public void Constructor_Stores_ProviderName()
    {
        var provider = new DbProvider(
            "conn", "System.Data.OleDb", "tbl", "id", "pid", "cap");

        Assert.That(provider.ProviderName, Is.EqualTo("System.Data.OleDb"));
    }

    [Test]
    public void Constructor_Stores_TableName()
    {
        var provider = new DbProvider(
            "conn", "System.Data.OleDb", "tblTree", "id", "pid", "cap");

        Assert.That(provider.TableName, Is.EqualTo("tblTree"));
    }

    [Test]
    public void Constructor_Stores_ColumnNames()
    {
        var provider = new DbProvider(
            "conn", "prov", "tbl", "NodeID", "ParentID", "NodeText", "Icon");

        Assert.That(provider.IdColumn, Is.EqualTo("NodeID"));
        Assert.That(provider.ParentIdColumn, Is.EqualTo("ParentID"));
        Assert.That(provider.CaptionColumn, Is.EqualTo("NodeText"));
        Assert.That(provider.IconKeyColumn, Is.EqualTo("Icon"));
    }

    [Test]
    public void Constructor_IconKeyColumn_Defaults_To_Null()
    {
        var provider = new DbProvider(
            "conn", "prov", "tbl", "id", "pid", "cap");

        Assert.That(provider.IconKeyColumn, Is.Null);
    }

    [Test]
    public void Constructor_Throws_On_Null_ConnectionString()
    {
        Assert.Throws<ArgumentNullException>(() => new DbProvider(
            null, "System.Data.OleDb", "tbl", "id", "pid", "cap"));
    }

    [Test]
    public void Constructor_Throws_On_Null_ProviderName()
    {
        Assert.Throws<ArgumentNullException>(() => new DbProvider(
            "conn", null, "tbl", "id", "pid", "cap"));
    }

    [Test]
    public void Constructor_Throws_On_Null_TableName()
    {
        Assert.Throws<ArgumentNullException>(() => new DbProvider(
            "conn", "System.Data.OleDb", null, "id", "pid", "cap"));
    }

    [Test]
    public void Constructor_Throws_On_Null_IdColumn()
    {
        Assert.Throws<ArgumentNullException>(() => new DbProvider(
            "conn", "prov", "tbl", null, "pid", "cap"));
    }

    [Test]
    public void Constructor_Throws_On_Null_ParentIdColumn()
    {
        Assert.Throws<ArgumentNullException>(() => new DbProvider(
            "conn", "prov", "tbl", "id", null, "cap"));
    }

    [Test]
    public void Constructor_Throws_On_Null_CaptionColumn()
    {
        Assert.Throws<ArgumentNullException>(() => new DbProvider(
            "conn", "prov", "tbl", "id", "pid", null));
    }

    [Test]
    public void Implements_ITreeDataProvider()
    {
        var provider = new DbProvider(
            "conn", "prov", "tbl", "id", "pid", "cap");

        Assert.That(provider, Is.InstanceOf<ITreeDataProvider>());
    }
}

[TestFixture]
public class ParseConnectionStringTests
{
    [Test]
    public void Extracts_Custom_Keys_From_ConnectionString()
    {
        var input = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=test.accdb;Table=tblTree;IdCol=NID;ParentCol=PID;CaptionCol=Cap";
        var (dbConn, config) = TreeEngine.ParseConnectionString(input);

        Assert.That(dbConn, Is.EqualTo("Provider=Microsoft.ACE.OLEDB.16.0;Data Source=test.accdb"));
        Assert.That(config["Table"], Is.EqualTo("tblTree"));
        Assert.That(config["IdCol"], Is.EqualTo("NID"));
        Assert.That(config["ParentCol"], Is.EqualTo("PID"));
        Assert.That(config["CaptionCol"], Is.EqualTo("Cap"));
    }

    [Test]
    public void Extracts_DbProvider_Key()
    {
        var input = "Data Source=test.accdb;DbProvider=System.Data.SqlClient;Table=tbl";
        var (dbConn, config) = TreeEngine.ParseConnectionString(input);

        Assert.That(dbConn, Is.EqualTo("Data Source=test.accdb"));
        Assert.That(config["DbProvider"], Is.EqualTo("System.Data.SqlClient"));
    }

    [Test]
    public void Extracts_IconCol_Key()
    {
        var input = "Data Source=test.accdb;IconCol=IconKey";
        var (dbConn, config) = TreeEngine.ParseConnectionString(input);

        Assert.That(dbConn, Is.EqualTo("Data Source=test.accdb"));
        Assert.That(config["IconCol"], Is.EqualTo("IconKey"));
    }

    [Test]
    public void Returns_Full_String_When_No_Custom_Keys()
    {
        var input = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=test.accdb";
        var (dbConn, config) = TreeEngine.ParseConnectionString(input);

        Assert.That(dbConn, Is.EqualTo("Provider=Microsoft.ACE.OLEDB.16.0;Data Source=test.accdb"));
        Assert.That(config, Is.Empty);
    }

    [Test]
    public void Custom_Keys_Are_Case_Insensitive()
    {
        var input = "Data Source=x;table=tblFoo;IDCOL=MyId";
        var (_, config) = TreeEngine.ParseConnectionString(input);

        Assert.That(config["Table"], Is.EqualTo("tblFoo"));
        Assert.That(config["IdCol"], Is.EqualTo("MyId"));
    }
}
