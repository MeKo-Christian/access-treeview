using System;
using System.Runtime.InteropServices;

namespace MeKo.TreeViewHost;

[ComVisible(true)]
[Guid("B1B2C3D4-1111-1111-1111-000000000001")]
[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
public interface ITreeViewHost
{
    object Engine { get; set; }
    string SelectedNodeId { get; }
    bool CheckBoxes { get; set; }

    void Initialize(object engine);
    void Reload();
    void ExpandNode(string nodeId);
    void CollapseNode(string nodeId);
    void SelectNode(string nodeId);
    bool FindAndSelect(string text);
}
