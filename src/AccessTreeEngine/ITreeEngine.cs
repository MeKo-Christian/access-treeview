using System;
using System.Runtime.InteropServices;

namespace Access.TreeEngine;

[ComVisible(true)]
[Guid("A1B2C3D4-1111-1111-1111-000000000003")]
[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
public interface ITreeEngine
{
    void Initialize(string connectionString, object context = null);
    ITreeNodeCollection GetRootNodes();
    ITreeNodeCollection GetChildren(string nodeId);
    bool HasChildren(string nodeId);
    ITreeNodeCollection Find(string text, int maxResults = 100);
    ITreeNode GetNode(string nodeId);
    void Invalidate(string nodeId);
    void Reload();
}
