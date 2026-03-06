using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;

namespace Access.TreeEngine;

/// <summary>
/// Data provider that loads tree nodes from a database via ADO.NET.
/// Works with any DbProviderFactory (OleDb for Access, SqlClient for SQL Server, etc.)
/// </summary>
public class DbProvider : ITreeDataProvider
{
    private readonly string _connectionString;
    private readonly string _providerName;
    private readonly string _tableName;
    private readonly string _idCol;
    private readonly string _parentIdCol;
    private readonly string _captionCol;
    private readonly string _iconKeyCol;

    public DbProvider(
        string connectionString,
        string providerName,
        string tableName,
        string idColumn,
        string parentIdColumn,
        string captionColumn,
        string iconKeyColumn = null)
    {
        _connectionString = connectionString ?? throw new ArgumentNullException(nameof(connectionString));
        _providerName = providerName ?? throw new ArgumentNullException(nameof(providerName));
        _tableName = tableName ?? throw new ArgumentNullException(nameof(tableName));
        _idCol = idColumn ?? throw new ArgumentNullException(nameof(idColumn));
        _parentIdCol = parentIdColumn ?? throw new ArgumentNullException(nameof(parentIdColumn));
        _captionCol = captionColumn ?? throw new ArgumentNullException(nameof(captionColumn));
        _iconKeyCol = iconKeyColumn;
    }

    public string ConnectionString => _connectionString;
    public string ProviderName => _providerName;
    public string TableName => _tableName;
    public string IdColumn => _idCol;
    public string ParentIdColumn => _parentIdCol;
    public string CaptionColumn => _captionCol;
    public string IconKeyColumn => _iconKeyCol;

    public List<TreeNode> GetRootNodes()
    {
        return QueryNodes(
            $"SELECT * FROM [{_tableName}] WHERE [{_parentIdCol}] IS NULL OR [{_parentIdCol}] = ''",
            Array.Empty<(string, object)>());
    }

    public List<TreeNode> GetChildren(string parentId)
    {
        return QueryNodes(
            $"SELECT * FROM [{_tableName}] WHERE [{_parentIdCol}] = @parentId",
            new[] { ("@parentId", (object)parentId) });
    }

    public bool HasChildren(string nodeId)
    {
        var factory = DbProviderFactories.GetFactory(_providerName);
        using var conn = factory.CreateConnection();
        conn.ConnectionString = _connectionString;
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"SELECT COUNT(*) FROM [{_tableName}] WHERE [{_parentIdCol}] = @nodeId";
        var param = cmd.CreateParameter();
        param.ParameterName = "@nodeId";
        param.Value = nodeId;
        cmd.Parameters.Add(param);
        return Convert.ToInt32(cmd.ExecuteScalar()) > 0;
    }

    public TreeNode GetNode(string nodeId)
    {
        var nodes = QueryNodes(
            $"SELECT * FROM [{_tableName}] WHERE [{_idCol}] = @nodeId",
            new[] { ("@nodeId", (object)nodeId) });
        return nodes.Count > 0 ? nodes[0] : null;
    }

    public List<TreeNode> Find(string text, int maxResults)
    {
        // Note: TOP is Access/SQL Server syntax. May need adjustment for other DBs.
        return QueryNodes(
            $"SELECT TOP {maxResults} * FROM [{_tableName}] WHERE [{_captionCol}] LIKE @text",
            new[] { ("@text", (object)$"%{text}%") });
    }

    private List<TreeNode> QueryNodes(string sql, (string name, object value)[] parameters)
    {
        var result = new List<TreeNode>();
        var factory = DbProviderFactories.GetFactory(_providerName);
        using var conn = factory.CreateConnection();
        conn.ConnectionString = _connectionString;
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = sql;
        foreach (var (name, value) in parameters)
        {
            var param = cmd.CreateParameter();
            param.ParameterName = name;
            param.Value = value;
            cmd.Parameters.Add(param);
        }
        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            var node = new TreeNode(
                id: reader[_idCol]?.ToString() ?? "",
                parentId: reader[_parentIdCol]?.ToString() ?? "",
                caption: reader[_captionCol]?.ToString() ?? ""
            );
            if (_iconKeyCol != null && reader[_iconKeyCol] != DBNull.Value)
                node.IconKey = reader[_iconKeyCol].ToString();
            result.Add(node);
        }
        return result;
    }
}
