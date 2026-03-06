using System.Collections.Generic;

namespace Access.TreeEngine;

public interface ITreeDataProvider
{
    List<TreeNode> GetRootNodes();
    List<TreeNode> GetChildren(string parentId);
    bool HasChildren(string nodeId);
    TreeNode GetNode(string nodeId);
    List<TreeNode> Find(string text, int maxResults);
}
