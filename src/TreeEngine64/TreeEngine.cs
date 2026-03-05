using System;
using System.Runtime.InteropServices;

namespace MeKo.TreeEngine;

[ComVisible(true)]
[Guid("A1B2C3D4-3333-3333-3333-000000000001")]
[ProgId("MeKo.TreeEngine")]
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
        // Will be wired to OleDbProvider in Task 1.6
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
        // Will be meaningful with caching provider
    }

    public void Reload()
    {
        // Will be meaningful with DB provider
    }

    private void EnsureProvider()
    {
        if (_provider == null)
            throw new InvalidOperationException(
                "TreeEngine not initialized. Call SetProvider() or Initialize() first.");
    }
}
