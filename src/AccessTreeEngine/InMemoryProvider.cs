using System;
using System.Collections.Generic;
using System.Linq;

namespace Access.TreeEngine;

public class InMemoryProvider : ITreeDataProvider
{
    private readonly List<TreeNode> _allNodes;

    public InMemoryProvider(List<TreeNode> nodes)
    {
        _allNodes = nodes ?? new List<TreeNode>();
    }

    public List<TreeNode> GetRootNodes()
        => _allNodes.Where(n => string.IsNullOrEmpty(n.ParentId)).ToList();

    public List<TreeNode> GetChildren(string parentId)
        => _allNodes.Where(n => n.ParentId == parentId).ToList();

    public bool HasChildren(string nodeId)
        => _allNodes.Any(n => n.ParentId == nodeId);

    public TreeNode GetNode(string nodeId)
        => _allNodes.FirstOrDefault(n => n.Id == nodeId);

    public List<TreeNode> Find(string text, int maxResults)
        => _allNodes
            .Where(n => n.Caption.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0)
            .Take(maxResults)
            .ToList();
}
