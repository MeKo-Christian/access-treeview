using NUnit.Framework;
using MeKo.TreeEngine;
using System.Collections.Generic;

namespace TreeEngine64.Tests;

[TestFixture]
public class CachingProviderTests
{
    private int _getChildrenCallCount;
    private CachingProviderDecorator _caching;

    [SetUp]
    public void Setup()
    {
        _getChildrenCallCount = 0;
        var inner = new CallCountingProvider(
            new List<TreeNode>
            {
                new TreeNode("1", "", "Root"),
                new TreeNode("1.1", "1", "Child")
            },
            () => _getChildrenCallCount++
        );
        _caching = new CachingProviderDecorator(inner);
    }

    [Test]
    public void GetChildren_Caches_Results()
    {
        _caching.GetChildren("1");
        _caching.GetChildren("1");

        Assert.That(_getChildrenCallCount, Is.EqualTo(1), "Second call should use cache");
    }

    [Test]
    public void GetChildren_Different_Ids_Are_Cached_Separately()
    {
        _caching.GetChildren("1");
        _caching.GetChildren("2");

        Assert.That(_getChildrenCallCount, Is.EqualTo(2));
    }

    [Test]
    public void Invalidate_Clears_Cache_For_Node()
    {
        _caching.GetChildren("1");
        _caching.Invalidate("1");
        _caching.GetChildren("1");

        Assert.That(_getChildrenCallCount, Is.EqualTo(2), "Should re-fetch after invalidation");
    }

    [Test]
    public void InvalidateAll_Clears_Entire_Cache()
    {
        _caching.GetChildren("1");
        _caching.GetChildren("2");
        _caching.InvalidateAll();
        _caching.GetChildren("1");

        Assert.That(_getChildrenCallCount, Is.EqualTo(3));
    }

    [Test]
    public void Other_Methods_Pass_Through()
    {
        var roots = _caching.GetRootNodes();
        Assert.That(roots.Count, Is.EqualTo(1));
        Assert.That(roots[0].Caption, Is.EqualTo("Root"));
    }
}

/// <summary>
/// Test helper: wraps InMemoryProvider and counts GetChildren calls.
/// </summary>
public class CallCountingProvider : ITreeDataProvider
{
    private readonly InMemoryProvider _inner;
    private readonly System.Action _onGetChildren;

    public CallCountingProvider(List<TreeNode> nodes, System.Action onGetChildren)
    {
        _inner = new InMemoryProvider(nodes);
        _onGetChildren = onGetChildren;
    }

    public List<TreeNode> GetRootNodes() => _inner.GetRootNodes();
    public List<TreeNode> GetChildren(string parentId)
    {
        _onGetChildren();
        return _inner.GetChildren(parentId);
    }
    public bool HasChildren(string nodeId) => _inner.HasChildren(nodeId);
    public TreeNode GetNode(string nodeId) => _inner.GetNode(nodeId);
    public List<TreeNode> Find(string text, int maxResults) => _inner.Find(text, maxResults);
}
