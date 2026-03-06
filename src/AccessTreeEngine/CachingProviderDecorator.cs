using System.Collections.Concurrent;
using System.Collections.Generic;

namespace MeKo.TreeEngine;

public class CachingProviderDecorator : ITreeDataProvider
{
    private readonly ITreeDataProvider _inner;
    private readonly ConcurrentDictionary<string, List<TreeNode>> _childrenCache = new();

    public CachingProviderDecorator(ITreeDataProvider inner)
    {
        _inner = inner;
    }

    public List<TreeNode> GetRootNodes() => _inner.GetRootNodes();

    public List<TreeNode> GetChildren(string parentId)
    {
        return _childrenCache.GetOrAdd(parentId, id => _inner.GetChildren(id));
    }

    public bool HasChildren(string nodeId) => _inner.HasChildren(nodeId);
    public TreeNode GetNode(string nodeId) => _inner.GetNode(nodeId);
    public List<TreeNode> Find(string text, int maxResults) => _inner.Find(text, maxResults);

    public void Invalidate(string nodeId)
    {
        _childrenCache.TryRemove(nodeId, out _);
    }

    public void InvalidateAll()
    {
        _childrenCache.Clear();
    }
}
