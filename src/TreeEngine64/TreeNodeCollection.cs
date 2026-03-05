using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace MeKo.TreeEngine;

[ComVisible(true)]
[Guid("A1B2C3D4-2222-2222-2222-000000000002")]
[ClassInterface(ClassInterfaceType.None)]
public class TreeNodeCollection : ITreeNodeCollection
{
    private readonly List<TreeNode> _nodes;

    public TreeNodeCollection(List<TreeNode> nodes)
    {
        _nodes = nodes ?? new List<TreeNode>();
    }

    public int Count => _nodes.Count;

    public ITreeNode this[int index] => _nodes[index - 1]; // 1-based for VBA

    [DispId(-4)]
    public IEnumerator GetEnumerator() => _nodes.GetEnumerator();
}
