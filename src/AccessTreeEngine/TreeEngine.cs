using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace Access.TreeEngine;

[ComVisible(true)]
[Guid("A1B2C3D4-3333-3333-3333-000000000001")]
[ProgId("Access.TreeEngine")]
[ClassInterface(ClassInterfaceType.None)]
public class TreeEngine : ITreeEngine
{
    private ITreeDataProvider _provider;

    public void SetProvider(ITreeDataProvider provider)
    {
        _provider = provider;
    }

    public void Initialize(string connectionString, object context = null)
    {
        var (dbConnectionString, config) = ParseConnectionString(connectionString);

        var dbProvider = new DbProvider(
            connectionString: dbConnectionString,
            providerName: config.GetValueOrDefault("DbProvider", "System.Data.OleDb"),
            tableName: config.GetValueOrDefault("Table", "tblTreeNodes"),
            idColumn: config.GetValueOrDefault("IdCol", "NodeID"),
            parentIdColumn: config.GetValueOrDefault("ParentCol", "ParentID"),
            captionColumn: config.GetValueOrDefault("CaptionCol", "NodeText"),
            iconKeyColumn: config.GetValueOrDefault("IconCol", null)
        );

        _provider = new CachingProviderDecorator(dbProvider);
    }

    public static (string DbConnectionString, Dictionary<string, string> Config) ParseConnectionString(string connectionString)
    {
        var customKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Table", "IdCol", "ParentCol", "CaptionCol", "IconCol", "DbProvider"
        };

        var config = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var dbParts = new List<string>();

        // Split on semicolons, preserving key=value pairs
        var parts = connectionString.Split(';', StringSplitOptions.RemoveEmptyEntries);
        foreach (var part in parts)
        {
            var trimmed = part.Trim();
            var eqIndex = trimmed.IndexOf('=');
            if (eqIndex > 0)
            {
                var key = trimmed.Substring(0, eqIndex).Trim();
                var value = trimmed.Substring(eqIndex + 1).Trim();
                if (customKeys.Contains(key))
                {
                    config[key] = value;
                    continue;
                }
            }
            dbParts.Add(trimmed);
        }

        return (string.Join(";", dbParts), config);
    }

    public ITreeNodeCollection GetRootNodes()
    {
        EnsureProvider();
        return new TreeNodeCollection(_provider.GetRootNodes());
    }

    public ITreeNodeCollection GetChildren(string nodeId)
    {
        EnsureProvider();
        return new TreeNodeCollection(_provider.GetChildren(nodeId));
    }

    public bool HasChildren(string nodeId)
    {
        EnsureProvider();
        return _provider.HasChildren(nodeId);
    }

    public ITreeNodeCollection Find(string text, int maxResults = 100)
    {
        EnsureProvider();
        return new TreeNodeCollection(_provider.Find(text, maxResults));
    }

    public ITreeNode GetNode(string nodeId)
    {
        EnsureProvider();
        return _provider.GetNode(nodeId);
    }

    public void Invalidate(string nodeId)
    {
        if (_provider is CachingProviderDecorator caching)
            caching.Invalidate(nodeId);
    }

    public void Reload()
    {
        if (_provider is CachingProviderDecorator caching)
            caching.InvalidateAll();
    }

    private void EnsureProvider()
    {
        if (_provider == null)
            throw new InvalidOperationException(
                "TreeEngine not initialized. Call SetProvider() or Initialize() first.");
    }
}
